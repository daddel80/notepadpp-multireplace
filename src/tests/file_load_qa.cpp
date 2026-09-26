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

// QA harness for the Find/Replace in Files loader (HiddenSciGuard::loadTextFile) on real
// files, with the expected text worked out independently here:
//   U. UTF-8 is taken as it is, a BOM dropped
//   W. UTF-16 LE/BE with BOM and LE without BOM are decoded; an odd length is read like
//      N++ in Find and refused in Replace (the dangling byte cannot be written back)
//   A. ANSI is decoded through the system codepage
//   B. binary: skipped, or taken as raw bytes with the binary skip off
//   S. sizes around the 8 KB header and a large file arrive whole; a NUL after the
//      header does not make a file binary; the size limit applies
//   R. missing, folder, locked by another program: unreadable
//   M. one string reused for file after file: exact content, capacity kept
//   T. text over 2 GB is refused before the full read (needs a sparse file, else SKIP)
//   C. every skip is counted once
// Windows only. The file_search_qa replica covers the same decisions portably.
//
// Build (from src/tests):
//   cl /std:c++20 /EHsc /I.. file_load_qa.cpp ..\Encoding.cpp ..\StringUtils.cpp shlwapi.lib advapi32.lib /Fe:file_load_qa.exe
//   x86_64-w64-mingw32-g++ -std=c++20 -Wall -Wextra -static -I.. -o file_load_qa.exe file_load_qa.cpp ../Encoding.cpp ../StringUtils.cpp -lshlwapi
#include <windows.h>
#include "../PluginInterface.h"
#include "../HiddenSciGuard.h"

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

    fs::remove_all(g_dir, ec);
    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
