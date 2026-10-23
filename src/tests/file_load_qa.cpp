// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.

// QA harness for the Find/Replace in Files loader (FileLoad::load, through HiddenSciGuard::loadTextFile)
// on real files, with the expected text worked out independently here:
//   U. UTF-8 is taken as it is, a BOM dropped
//   W. UTF-16 LE/BE with BOM and LE without BOM are decoded; an odd length is read like
//      N++ in Find and refused in Replace (the dangling byte cannot be written back); without
//      BOM it is binary within N++'s first block of 128 KB + 4 bytes, as N++ takes no UTF-16 there
//   A. ANSI is decoded through the system codepage
//   B. binary: skipped, or taken as raw bytes with the binary skip off
//   S. sizes around the 8 KB header and a large file arrive whole; a NUL after the
//      header does not make a file binary; the size limit applies
//   R. missing, folder, locked by another program: unreadable
//   M. one string reused for file after file: exact content, capacity kept
//   T. text over 2 GB is refused before the full read (needs a sparse file, else SKIP)
//   C. every skip is counted once
//   O. a worker's options: read-only files left out before opening, the file's state stamped
//   K. changed since loaded: another size, a later write time, a new file in its place (atomic
//      write) or a file that is gone; a hard link keeps the old file, a link to a file follows
//   N. not writable: held open by another program that shares it for reading only, or its folder
//      is gone; the file stays as it was, no temp file is left, the skip is counted once
//   P. a worker's pace: it gets the size before the full read and may hold it back, a stop ends it
//   D. kind of drive: the folder of the test files is on a local drive, a drive letter that is not
//      there counts as fast
// Windows only. The file_search_qa replica covers the same decisions portably.
//
// Build (from src/tests):
//   cl /std:c++20 /EHsc /I.. file_load_qa.cpp ..\FileLoad.cpp ..\Encoding.cpp ..\StringUtils.cpp shlwapi.lib advapi32.lib /Fe:file_load_qa.exe
//   x86_64-w64-mingw32-g++ -std=c++20 -Wall -Wextra -static -I.. -o file_load_qa.exe file_load_qa.cpp ../FileLoad.cpp ../Encoding.cpp ../StringUtils.cpp -lshlwapi
#include <windows.h>
#include "../PluginInterface.h"
#include "../HiddenSciGuard.h"

#include <atomic>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <string>

NppData nppData{};

namespace fs = std::filesystem;
using Skip = HiddenSciGuard::SkipReason;
using Kind = HiddenSciGuard::LoadKind;

static int failures = 0;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }

static fs::path g_dir;
static fs::path write(const std::string& name, const std::string& bytes)
{
    const fs::path p = g_dir / name;
    std::ofstream(p, std::ios::binary).write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    return p;
}

// Expected text, converted here without the Encoding unit
static std::string utf8Of(const std::wstring& w)
{
    if (w.empty()) return {};
    const int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
    std::string out(static_cast<size_t>(n), '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), out.data(), n, nullptr, nullptr);
    return out;
}
static std::string ansiToUtf8(const std::string& a)
{
    const int n = MultiByteToWideChar(CP_ACP, 0, a.data(), static_cast<int>(a.size()), nullptr, 0);
    std::wstring w(static_cast<size_t>(n), L'\0');
    MultiByteToWideChar(CP_ACP, 0, a.data(), static_cast<int>(a.size()), w.data(), n);
    return utf8Of(w);
}
static std::string utf16(const std::wstring& w, bool bigEndian, bool bom)
{
    std::string out;
    if (bom) out += bigEndian ? std::string("\xFE\xFF", 2) : std::string("\xFF\xFE", 2);
    for (wchar_t c : w) {
        const char lo = static_cast<char>(c & 0xFF), hi = static_cast<char>((c >> 8) & 0xFF);
        out += bigEndian ? hi : lo;
        out += bigEndian ? lo : hi;
    }
    return out;
}

struct Loaded {
    Skip reason = Skip::None;
    Kind kind = Kind::Text;
    Encoding::EncodingInfo enc;
    std::string content;
};

static Loaded load(HiddenSciGuard& g, const fs::path& p)
{
    Loaded r;
    r.reason = g.loadTextFile(p, r.content, r.enc, r.kind);
    return r;
}

int main()
{
    g_dir = fs::temp_directory_path() / ("mr_file_load_qa_" + std::to_string(GetCurrentProcessId()));
    std::error_code ec;
    fs::remove_all(g_dir, ec);
    fs::create_directories(g_dir);
    std::printf("folder %s, ANSI codepage %u\n", g_dir.string().c_str(), GetACP());

    const std::wstring textW = L"Gr\u00FC\u00DFe Fi \u20AC\r\nsecond line Fi\n";
    const std::string text8 = utf8Of(textW);
    HiddenSciGuard find;           // Find in Files
    HiddenSciGuard replace;        // Replace in Files
    replace.setVerifyRoundtrip(true);

    std::printf("\n=== U) UTF-8 ===\n");
    {
        const Loaded a = load(find, write("u8.txt", text8));
        CHECK("U1 without BOM: Text/UTF-8, content is the file", a.reason == Skip::None && a.kind == Kind::Text
            && a.enc.kind == Encoding::Kind::UTF8 && !a.enc.withBOM && a.content == text8);
        const Loaded b = load(find, write("u8bom.txt", "\xEF\xBB\xBF" + text8));
        CHECK("U2 with BOM: Text/UTF-8 with BOM, content without it", b.reason == Skip::None
            && b.enc.kind == Encoding::Kind::UTF8 && b.enc.withBOM && b.enc.bomBytes == 3 && b.content == text8);
        const Loaded c = load(find, write("u8bomonly.txt", "\xEF\xBB\xBF"));
        CHECK("U3 BOM alone: empty text", c.reason == Skip::None && c.enc.withBOM && c.content.empty());
        const Loaded d = load(replace, write("u8r.txt", "\xEF\xBB\xBF" + text8));
        CHECK("U4 Replace takes UTF-8 as well (it writes back unchanged)", d.reason == Skip::None && d.content == text8);
    }

    std::printf("\n=== W) UTF-16 ===\n");
    {
        const Loaded le = load(find, write("le.txt", utf16(textW, false, true)));
        CHECK("W1 LE with BOM decoded", le.reason == Skip::None && le.enc.kind == Encoding::Kind::UTF16LE && le.content == text8);
        const Loaded be = load(find, write("be.txt", utf16(textW, true, true)));
        CHECK("W2 BE with BOM decoded", be.reason == Skip::None && be.enc.kind == Encoding::Kind::UTF16BE && be.content == text8);
        const std::wstring ascii = L"First Fight\r\nsecond line Fi\r\nthird Fi line\r\n";
        const Loaded nb = load(find, write("lenobom.txt", utf16(ascii, false, false)));
        CHECK("W3 LE without BOM decoded", nb.reason == Skip::None && nb.enc.kind == Encoding::Kind::UTF16LE && nb.content == utf8Of(ascii));
        const fs::path odd = write("odd.txt", utf16(textW, false, true) + "X");
        const Loaded of = load(find, odd);
        CHECK("W4 odd length in Find: read like N++, the dangling byte dropped", of.reason == Skip::None && of.content == text8);
        const Loaded orp = load(replace, odd);
        CHECK("W5 odd length in Replace: not decodable, nothing loaded", orp.reason == Skip::Undecodable && orp.content.empty());
        const Loaded er = load(replace, write("ler.txt", utf16(textW, false, true)));
        CHECK("W6 even length in Replace: decoded", er.reason == Skip::None && er.content == text8);
        auto asciiUnits = [&](size_t units) { std::wstring w; while (w.size() < units) w += ascii; w.resize(units); return w; };
        const Loaded os = load(find, write("oddsmall.txt", utf16(asciiUnits(100), false, false) + "X"));
        CHECK("W7 no BOM, odd length within the 8 KB header: binary", os.reason == Skip::Binary);
        const Loaded om = load(find, write("oddmid.txt", utf16(asciiUnits(20000), false, false) + "X"));
        CHECK("W8 no BOM, odd length within N++'s first block: binary, not UTF-8 text", om.reason == Skip::Binary);
        const std::wstring bigW = asciiUnits(70000);
        const fs::path big = write("oddbig.txt", utf16(bigW, false, false) + "X");
        const Loaded ol = load(find, big);
        CHECK("W9 no BOM, odd length beyond N++'s first block: UTF-16 LE, the dangling byte dropped", ol.reason == Skip::None
            && ol.enc.kind == Encoding::Kind::UTF16LE && ol.content == utf8Of(bigW));
        const Loaded olr = load(replace, big);
        CHECK("W10 the same in Replace: not decodable, nothing loaded", olr.reason == Skip::Undecodable && olr.content.empty());
    }

    std::printf("\n=== A) ANSI ===\n");
    {
        const std::string ansi = "Gr\xFC\xDF" "e Fi\r\nline two\r\n";
        const Loaded a = load(find, write("ansi.txt", ansi));
        CHECK("A1 decoded through the system codepage", a.reason == Skip::None && a.enc.kind == Encoding::Kind::ANSI
            && a.enc.codepage == GetACP() && a.content == ansiToUtf8(ansi));
        const Loaded r = load(replace, write("ansir.txt", ansi));
        CHECK("A2 Replace: round-trips, decoded", r.reason == Skip::None && r.content == ansiToUtf8(ansi));
    }

    std::printf("\n=== B) binary ===\n");
    {
        std::string bin = std::string("MZ\x90\x00\x03\x00", 6);
        for (int i = 0; i < 20000; ++i) bin += static_cast<char>((i * 7919) & 0xFF);
        const fs::path p = write("bin.exe", bin);
        const Loaded on = load(find, p);
        CHECK("B1 skip on: Binary, nothing loaded", on.reason == Skip::Binary && on.content.empty());
        HiddenSciGuard raw;
        raw.setSkipBinaryEnabled(false);
        const Loaded off = load(raw, p);
        CHECK("B2 skip off: raw bytes, all of them", off.reason == Skip::None && off.kind == Kind::RawBytes
            && off.content == bin && off.enc.kind == Encoding::Kind::ANSI);
        const std::string ctl = std::string("ctl\x00" "Fi \xC3\xA4\n", 10);
        const Loaded u = load(raw, write("ctl.bin", ctl));
        CHECK("B3 raw bytes that are valid UTF-8 are marked UTF-8", u.reason == Skip::None && u.kind == Kind::RawBytes
            && u.enc.kind == Encoding::Kind::UTF8 && u.content == ctl);
    }

    std::printf("\n=== S) sizes ===\n");
    {
        bool whole = true;
        for (size_t size : { size_t{ 0 }, size_t{ 1 }, size_t{ 8191 }, size_t{ 8192 }, size_t{ 8193 }, size_t{ 3 } << 20 }) {
            std::string s(size, 'a');
            for (size_t i = 0; i < size; i += 61) s[i] = '\n';
            const Loaded l = load(find, write("size" + std::to_string(size) + ".txt", s));
            if (l.reason != Skip::None || l.content != s) {
                whole = false;
                std::printf("  size %zu: reason %d, %zu bytes loaded\n", size, static_cast<int>(l.reason), l.content.size());
            }
        }
        CHECK("S1 0, 1, 8191, 8192, 8193 bytes and 3 MB arrive whole", whole);
        std::string late(9000, 'A');
        late += std::string("\0tail Fi", 8);
        const Loaded n = load(find, write("latenul.txt", late));
        CHECK("S2 a NUL after the 8 KB header: text, loaded whole", n.reason == Skip::None && n.kind == Kind::Text && n.content == late);
        HiddenSciGuard limited;
        limited.setFileSizeLimitEnabled(true);
        limited.setMaxFileSizeMB(1);
        const Loaded big = load(limited, g_dir / "size3145728.txt");
        const Loaded small = load(limited, g_dir / "size8193.txt");
        CHECK("S3 size limit 1 MB: 3 MB too large, 8 KB loaded", big.reason == Skip::TooLarge && small.reason == Skip::None);
    }

    std::printf("\n=== R) unreadable ===\n");
    {
        HiddenSciGuard g;
        const Loaded missing = load(g, g_dir / "does_not_exist.txt");
        const Loaded folder = load(g, g_dir);
        const fs::path locked = write("locked.txt", text8);
        const HANDLE h = CreateFileW(locked.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        const Loaded l = load(g, locked);
        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
        CHECK("R1 missing file: Unreadable", missing.reason == Skip::Unreadable);
        CHECK("R2 a folder: Unreadable", folder.reason == Skip::Unreadable);
        CHECK("R3 held exclusively by another program: Unreadable", h != INVALID_HANDLE_VALUE && l.reason == Skip::Unreadable);
        CHECK("C1 every skip counted once", g.getSkippedUnreadableCount() == 3 && g.getSkippedTotalCount() == 3);
    }

    std::printf("\n=== M) one string for file after file ===\n");
    {
        HiddenSciGuard g;
        std::string content;
        Encoding::EncodingInfo enc;
        Kind kind;
        const bool bigOk = g.loadTextFile(g_dir / "size3145728.txt", content, enc, kind) == Skip::None && content.size() == (size_t{ 3 } << 20);
        const size_t capacity = content.capacity();
        const bool smallOk = g.loadTextFile(g_dir / "u8bom.txt", content, enc, kind) == Skip::None && content == text8;
        const bool kept = content.capacity() >= capacity;
        const bool leOk = g.loadTextFile(g_dir / "le.txt", content, enc, kind) == Skip::None && content == text8;
        const bool skipEmpty = g.loadTextFile(g_dir / "bin.exe", content, enc, kind) == Skip::Binary && content.empty();
        const bool again = g.loadTextFile(g_dir / "u8.txt", content, enc, kind) == Skip::None && content == text8;
        CHECK("M1 a small file after a large one: exact content, capacity kept", bigOk && smallOk && kept);
        CHECK("M2 decoded, skipped and plain files in turn: exact content, empty on a skip", leOk && skipEmpty && again);
    }

    std::printf("\n=== T) text over 2 GB ===\n");
    {
        const fs::path p = g_dir / "huge.txt";
        const HANDLE h = CreateFileW(p.c_str(), GENERIC_READ | GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        DWORD ret = 0;
        const bool sparse = h != INVALID_HANDLE_VALUE && DeviceIoControl(h, FSCTL_SET_SPARSE, nullptr, 0, nullptr, 0, &ret, nullptr);
        if (sparse) {
            const std::string head(8192, 'x');
            DWORD written = 0;
            WriteFile(h, head.data(), static_cast<DWORD>(head.size()), &written, nullptr);
            LARGE_INTEGER end;
            end.QuadPart = (LONGLONG{ 5 } << 29);   // 2.5 GB, no space taken
            SetFilePointerEx(h, end, nullptr, FILE_BEGIN);
            SetEndOfFile(h);
        }
        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
        if (!sparse) {
            std::printf("SKIP T (no sparse files here)\n");
        }
        else {
            HiddenSciGuard g;
            const ULONGLONG t0 = GetTickCount64();
            const Loaded l = load(g, p);
            const ULONGLONG ms = GetTickCount64() - t0;
            CHECK("T1 TooLarge before the full read (" + std::to_string(ms) + " ms)", l.reason == Skip::TooLarge && l.content.empty() && ms < 2000);
        }
    }

    std::printf("\n=== O) options of a worker ===\n");
    {
        FileLoad::Options readOnlySkip;
        readOnlySkip.skipReadOnly = true;
        const fs::path ro = write("readonly.txt", text8);
        SetFileAttributesW(ro.c_str(), FILE_ATTRIBUTE_READONLY);
        std::string content;
        const FileLoad::Result a = FileLoad::load(ro, readOnlySkip, content);
        const FileLoad::Result b = FileLoad::load(ro, FileLoad::Options{}, content);
        SetFileAttributesW(ro.c_str(), FILE_ATTRIBUTE_NORMAL);
        CHECK("O1 read-only: left out when asked, loaded otherwise", a.skip == Skip::ReadOnly && b.skip == Skip::None && content == text8);

        FileLoad::Options stamped;
        stamped.stamp = true;
        const fs::path u8 = g_dir / "u8.txt";
        const FileLoad::Result s = FileLoad::load(u8, stamped, content);
        const FileLoad::Result n = FileLoad::load(u8, FileLoad::Options{}, content);
        CHECK("O2 stamped: valid, the file's size, unchanged since", s.skip == Skip::None && s.stamp.valid
            && s.stamp.size == fs::file_size(u8) && FileLoad::unchangedSince(u8, s.stamp));
        CHECK("O3 stamped only when asked", !n.stamp.valid && !FileLoad::unchangedSince(u8, n.stamp));
    }

    std::printf("\n=== K) changed since loaded ===\n");
    {
        FileLoad::Options stamped;
        stamped.stamp = true;
        std::string content;
        HiddenSciGuard g;   // writeFile: temp file and replace, as Replace in Files writes

        auto lastWrite = [](const fs::path& p) {
            FILETIME t{};
            const HANDLE h = CreateFileW(p.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr);
            if (h != INVALID_HANDLE_VALUE) { GetFileTime(h, nullptr, nullptr, &t); CloseHandle(h); }
            return t;
        };
        auto setLastWrite = [](const fs::path& p, FILETIME t) {
            const HANDLE h = CreateFileW(p.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, 0, nullptr);
            if (h == INVALID_HANDLE_VALUE) return false;
            const bool ok = SetFileTime(h, nullptr, nullptr, &t) != FALSE;
            CloseHandle(h);
            return ok;
        };

        const fs::path a = write("k_size.txt", text8);
        const FileLoad::Stamp s1 = FileLoad::load(a, stamped, content).stamp;
        write("k_size.txt", text8 + "x");
        CHECK("K1 written again with another size: changed", s1.valid && !FileLoad::unchangedSince(a, s1));

        const fs::path b = write("k_time.txt", text8);
        const FileLoad::Stamp s2 = FileLoad::load(b, stamped, content).stamp;
        {
            std::fstream f(b, std::ios::binary | std::ios::in | std::ios::out);
            f.seekp(0);
            f.put('X');
        }
        FILETIME later = lastWrite(b);
        ULARGE_INTEGER t;
        t.LowPart = later.dwLowDateTime;
        t.HighPart = later.dwHighDateTime;
        t.QuadPart += 10000000;   // one second on, whatever the file system's resolution
        later.dwLowDateTime = t.LowPart;
        later.dwHighDateTime = t.HighPart;
        const bool timeSet = setLastWrite(b, later);
        CHECK("K2 same size, written in place later: changed", timeSet && !FileLoad::unchangedSince(b, s2));

        const fs::path c = write("k_id.txt", text8);
        const FILETIME original = lastWrite(c);
        const FileLoad::Stamp s3 = FileLoad::load(c, stamped, content).stamp;
        const bool replaced = g.writeFile(c, text8) && setLastWrite(c, original);
        const FileLoad::Stamp now3 = FileLoad::load(c, stamped, content).stamp;
        CHECK("K3 an atomic write with the same bytes and time: a new file in its place, changed",
            replaced && now3.size == s3.size && now3.lastWrite == s3.lastWrite && now3.fileId != s3.fileId
            && !FileLoad::unchangedSince(c, s3));

        const fs::path d = write("k_hard_a.txt", text8);
        const fs::path e = g_dir / "k_hard_b.txt";
        if (CreateHardLinkW(e.c_str(), d.c_str(), nullptr)) {
            const FileLoad::Stamp s4 = FileLoad::load(e, stamped, content).stamp;
            g.writeFile(d, text8 + "changed");
            std::string other;
            FileLoad::load(e, FileLoad::Options{}, other);
            CHECK("K4 hard link: an atomic write of one name leaves the other with the old file, unchanged",
                FileLoad::unchangedSince(e, s4) && other == text8);
        }
        else {
            std::printf("SKIP K4 (no hard links here)\n");
        }

        const fs::path target = write("k_sym_target.txt", text8);
        const fs::path link = g_dir / "k_sym_link.txt";
        const bool linked = (CreateSymbolicLinkW(link.c_str(), target.c_str(), 0x2 /* allow unprivileged */)
            || CreateSymbolicLinkW(link.c_str(), target.c_str(), 0)) && GetFileAttributesW(link.c_str()) != INVALID_FILE_ATTRIBUTES;
        if (linked) {
            const FileLoad::Stamp s5 = FileLoad::load(link, stamped, content).stamp;
            g.writeFile(target, text8 + "changed");
            CHECK("K5 link to a file: after an atomic write of the target it shows a changed file", s5.valid && !FileLoad::unchangedSince(link, s5));
        }
        else {
            std::printf("SKIP K5 (no symbolic links without privilege here)\n");
        }

        CHECK("K6 a file that is gone: changed", !FileLoad::unchangedSince(g_dir / "does_not_exist.txt", s1));
    }

    std::printf("\n=== N) not writable ===\n");
    {
        HiddenSciGuard g;
        const fs::path held = write("n_held.txt", text8);
        const HANDLE h = CreateFileW(held.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        const Loaded l = load(g, held);
        const bool written = g.writeFile(held, text8 + "changed");
        if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
        HiddenSciGuard probe;
        CHECK("N1 shared for reading only by another program: loads", h != INVALID_HANDLE_VALUE && l.reason == Skip::None);
        CHECK("N2 but is not written and stays as it was", !written && load(probe, held).content == text8);
        CHECK("N3 no temp file left", !fs::exists(held.wstring() + L".mr_tmp"));
        CHECK("N4 folder gone: not written", !g.writeFile(g_dir / "gone" / "n.txt", text8));
        g.noteSkip(Skip::Unwritable);
        CHECK("N5 counted once", g.getSkippedUnwritableCount() == 1 && g.getSkippedTotalCount() == 1);
    }

    std::printf("\n=== P) a worker's pace ===\n");
    {
        std::string content;
        const fs::path big = g_dir / "size3145728.txt";
        uint64_t seen = 0;
        FileLoad::Pace refuse;
        refuse.beforeRead = [&seen](uint64_t size) { seen = size; return false; };
        const FileLoad::Result a = FileLoad::load(big, FileLoad::Options{}, content, &refuse);
        CHECK("P1 the size before the full read; held back: unreadable, nothing loaded",
            seen == fs::file_size(big) && a.skip == Skip::Unreadable && content.empty());

        seen = 0;
        const FileLoad::Result bin = FileLoad::load(g_dir / "bin.exe", FileLoad::Options{}, content, &refuse);
        CHECK("P2 a binary file is skipped before that point, nothing to hold back", bin.skip == Skip::Binary && seen == 0);

        std::atomic<bool> stop{ true };
        FileLoad::Pace stopped;
        stopped.stop = &stop;
        const FileLoad::Result b = FileLoad::load(big, FileLoad::Options{}, content, &stopped);
        const FileLoad::Result small = FileLoad::load(g_dir / "u8.txt", FileLoad::Options{}, content, &stopped);
        CHECK("P3 a stop ends the full read: unreadable; a file within the header still loads",
            b.skip == Skip::Unreadable && small.skip == Skip::None && content == text8);
    }

    std::printf("\n=== D) kind of drive ===\n");
    {
        const FileLoad::DriveKind local = FileLoad::driveKind(g_dir);
        const char* name = local == FileLoad::DriveKind::Fast ? "fast" : local == FileLoad::DriveKind::Slow ? "slow" : "network";
        CHECK(std::string("D1 the folder of the test files is on a local drive (") + name + ")", local != FileLoad::DriveKind::Network);

        const DWORD used = GetLogicalDrives();
        int unused = -1;
        for (int i = 25; i >= 2 && unused < 0; --i)
            if (!(used & (DWORD{ 1 } << i))) unused = i;
        if (unused < 0) {
            std::printf("SKIP D2 (every drive letter in use)\n");
        }
        else {
            const std::wstring missing = std::wstring(1, static_cast<wchar_t>(L'A' + unused)) + L":\\folder";
            CHECK("D2 a path on a drive letter that is not there: fast", FileLoad::driveKind(missing) == FileLoad::DriveKind::Fast);
        }
    }

    fs::remove_all(g_dir, ec);
    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
