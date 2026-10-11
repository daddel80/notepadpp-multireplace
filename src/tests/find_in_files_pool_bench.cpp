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

// Thread test for Find in Files: MultiReplace's work per file on real folders, in two ways:
//   serial  as handleFindInFiles does it with LoadThreads=0: load, prefilter, fill and search, file after file
//   pool    FileLoadPool loads, decodes and prefilters ahead on worker threads; the consumer takes the files
//           strictly in list order and fills and searches them, as the UI thread does
// Loading is the plugin's (FileLoad, FileLoadPool, LiteralPrefilter), the file list DirectoryWalk::collect.
// Fill and search run on a Scintilla document with the Notepad++ Boost regex bridge (Notepad++ 697f46b),
// every hit with its line, no dock. All modes must find the same: hits, filled files and skips are compared.
// --device puts a simulated USB stick in front of every load. It serves one request after another, as USB
// mass storage without UASP does; a file costs an access time and its bytes at the stick's read speed:
//   usb2  2 ms per file, 30 MB/s        usb3  0.5 ms per file, 150 MB/s
// The stick counts as a removable drive. The summary shows its busy time (no mode can be faster) and how
// often it had to jump back to a file for its rest; in the model a jump costs nothing.
// --cpus N keeps the program on the first N hardware threads, as on a smaller machine. MinGW's
// hardware_concurrency counts only those, so LoadThreads=auto is reported as for such a machine.
// --cold measures the first read: one --dir per mode, all with the same files, and one query. Every mode
// reads its own folder once, then all read theirs again, warm, as a reference. --budgets gives each mode its
// byte budget in MB (0 = the pool's 64 MB): a file above it is read only when it is the next, alone.
// --drives lists every drive with what Windows reports about it (type, bus, seek penalty, product) and what
// LoadThreads=auto takes there; without --dir that is all. The summary names the drives of the folders.
// Build: find_in_files_pool_bench.sh (MinGW-w64).
// Usage: find_in_files_pool_bench [--dir PATH]... [--rounds N] [--threads 0,1,2,...] [--queries absent,fi,in]
//                                 [--device none|usb2|usb3] [--cpus N] [--cold [--budgets 0,0,1,...]] [--drives]
//                                 [--out CSV]
//   Defaults: the folders source-text and source next to the program; 8 rounds, the first warms the file
//   cache and is left out; 0 (serial), 1, 2, 4, 8 and hardware threads - 1; all queries; no device;
//   thread_bench.csv next to the program. The summary is appended to <CSV name>_summary.txt.

#include <windows.h>
#include <winioctl.h>
#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <filesystem>
#include <forward_list>
#include <functional>
#include <map>
#include <memory>
#include <mutex>
#include <numeric>
#include <optional>
#include <random>
#include <set>
#include <stdexcept>
#include <string>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#include "PluginInterface.h"
#include "HiddenSciGuard.h"
#include "FileLoadPool.h"
#include "LiteralPrefilter.h"
#include "DirectoryWalk.h"
#include "DiskLoad.h"   // FileLoad.cpp under another name, made by find_in_files_pool_bench.sh

#undef FindText   // <windows.h> maps it to FindTextW, Document::FindText would not link

#include "ScintillaTypes.h"
#include "ILoader.h"
#include "ILexer.h"
#include "Debugging.h"
#include "Geometry.h"
#include "Platform.h"
#include "CharacterType.h"
#include "CharacterCategoryMap.h"
#include "Position.h"
#include "SplitVector.h"
#include "Partitioning.h"
#include "RunStyles.h"
#include "CellBuffer.h"
#include "PerLine.h"
#include "CharClassify.h"
#include "Decoration.h"
#include "CaseFolder.h"
#include "Document.h"

NppData nppData{};

using namespace Scintilla;
using namespace Scintilla::Internal;

namespace Scintilla::Internal {
void Platform::Assert(const char* c, const char* file, int line) noexcept { std::fprintf(stderr, "Assertion [%s] failed at %s %d\n", c, file, line); std::abort(); }
void Platform::DebugPrintf(const char*, ...) noexcept {}
void Platform::DebugDisplay(const char*) noexcept {}
bool Platform::ShowAssertionPopUps(bool) noexcept { return false; }
ColourRGBA Platform::Chrome() { return ColourRGBA(0xe0, 0xe0, 0xe0); }
ColourRGBA Platform::ChromeHighlight() { return ColourRGBA(0xff, 0xff, 0xff); }
const char* Platform::DefaultFont() { return "Consolas"; }
int Platform::DefaultFontSize() { return 10; }
unsigned int Platform::DoubleClickTime() { return 500; }
}

namespace fs = std::filesystem;

using Clock = std::chrono::steady_clock;
static double msSince(Clock::time_point t) { return std::chrono::duration<double, std::milli>(Clock::now() - t).count(); }

// ---------------------------------------------------------------------------------------------
// The simulated USB stick and the plugin's loader behind it
// ---------------------------------------------------------------------------------------------

// One request after another; the caller waits as long as the stick needs
class Stick {
public:
    bool set(const std::wstring& name)
    {
        if (name == L"none") { _on = false; return true; }
        if (name == L"usb2") { _msPerFile = 2.0; _mbPerSecond = 30; _on = true; return true; }
        if (name == L"usb3") { _msPerFile = 0.5; _mbPerSecond = 150; _on = true; return true; }
        return false;
    }
    bool on() const { return _on; }
    void open(const void* file, uint64_t headerBytes) { serve(file, false, _msPerFile + ms(headerBytes)); }
    void read(const void* file, uint64_t restBytes) { serve(file, true, ms(restBytes)); }

    // Since the last reset: how long the stick was busy (no run can be faster), and how often it had to
    // jump back to a file for its rest because it served another file in between. A jump costs nothing here;
    // a real stick may lose its read-ahead there, a hard disk seeks.
    struct Busy {
        double ms = 0;
        size_t jumps = 0;
    };
    Busy busy()
    {
        std::lock_guard<std::mutex> lock(_m);
        return _busy;
    }
    void resetBusy()
    {
        std::lock_guard<std::mutex> lock(_m);
        _busy = Busy{};
        _last = nullptr;
    }

private:
    double ms(uint64_t bytes) const { return static_cast<double>(bytes) / (_mbPerSecond * 1000.0); }

    // Busy as the stick is busy: Sleep cannot wait fractions of a millisecond
    void serve(const void* file, bool rest, double ms)
    {
        if (!_on) return;
        std::lock_guard<std::mutex> lock(_m);
        if (rest && file != _last) ++_busy.jumps;
        _last = file;
        const auto until = Clock::now() + std::chrono::duration_cast<Clock::duration>(std::chrono::duration<double, std::milli>(ms));
        while (Clock::now() < until) {}
        _busy.ms += ms;
    }

    std::mutex _m;
    bool _on = false;
    double _msPerFile = 0;
    double _mbPerSecond = 1;
    Busy _busy;
    const void* _last = nullptr;   // the file of the last request
};

static Stick g_stick;
static std::map<std::filesystem::path::string_type, uint64_t> g_sizes;   // file sizes for the stick, taken before the runs
static std::atomic<long long> g_loadUs{ 0 };   // time spent in FileLoad::load, all threads together

// FileLoad::load as the plugin and FileLoadPool call it: the stick first, then the plugin's code (DiskLoad)
namespace FileLoad {

Result load(const std::filesystem::path& fp, const Options& options, std::string& content, const Pace* pace)
{
    const auto t0 = Clock::now();
    if (g_stick.on()) {
        const auto it = g_sizes.find(fp.native());
        g_stick.open(&fp, it == g_sizes.end() ? BINARY_CHECK_SIZE : (std::min)(it->second, uint64_t{ BINARY_CHECK_SIZE }));
    }
    DiskLoad::Options disk;
    disk.maxFileSize = options.maxFileSize;
    disk.skipBinary = options.skipBinary;
    disk.verifyRoundtrip = options.verifyRoundtrip;
    disk.skipReadOnly = options.skipReadOnly;
    disk.stamp = options.stamp;
    DiskLoad::Pace diskPace;
    diskPace.stop = pace ? pace->stop : nullptr;
    diskPace.beforeRead = [pace, &fp](uint64_t fileSize) {
        if (pace && pace->beforeRead && !pace->beforeRead(fileSize)) return false;
        if (g_stick.on() && fileSize > BINARY_CHECK_SIZE) g_stick.read(&fp, fileSize - BINARY_CHECK_SIZE);
        return true;
    };
    const DiskLoad::Result loaded = DiskLoad::load(fp, disk, content, &diskPace);

    Result result;
    result.skip = static_cast<SkipReason>(loaded.skip);
    result.kind = static_cast<LoadKind>(loaded.kind);
    result.enc = loaded.enc;
    result.stamp.volume = loaded.stamp.volume;
    result.stamp.fileId = loaded.stamp.fileId;
    result.stamp.size = loaded.stamp.size;
    result.stamp.lastWrite = loaded.stamp.lastWrite;
    result.stamp.valid = loaded.stamp.valid;
    g_loadUs += std::chrono::duration_cast<std::chrono::microseconds>(Clock::now() - t0).count();
    return result;
}

// The simulated stick is a removable drive
DriveKind driveKind(const std::filesystem::path& folder)
{
    return g_stick.on() ? DriveKind::Fast : static_cast<DriveKind>(DiskLoad::driveKind(folder));
}

}

// ---------------------------------------------------------------------------------------------
// Fill and search as the UI thread does them
// ---------------------------------------------------------------------------------------------

// The hidden buffer's document: SC_DOCUMENTOPTION_STYLES_NONE | SC_DOCUMENTOPTION_TEXT_LARGE, no undo
struct SciDoc {
    Document* d = nullptr;
    int cp = -1;
    bool folder = false;

    SciDoc() {
        d = new Document(static_cast<DocumentOption>(0x1 | 0x100));
        d->AddRef();
        d->SetUndoCollection(false);
        setCodePage(SC_CP_UTF8);
    }
    ~SciDoc() { d->Release(); }
    SciDoc(const SciDoc&) = delete;
    SciDoc& operator=(const SciDoc&) = delete;

    void setCodePage(int c) {                    // SCI_SETCODEPAGE: a change drops the case folder
        if (c == cp) return;
        cp = c;
        d->SetDBCSCodePage(c);
        folder = false;
    }
    void ensureFolder() {                        // Editor::SearchInTarget -> CaseFolderForEncoding
        if (folder) return;
        if (cp == SC_CP_UTF8) d->SetCaseFolder(std::make_unique<CaseFolderUnicode>());
        else {
            auto f = std::make_unique<CaseFolderTable>();
            f->StandardASCII();
            d->SetCaseFolder(std::move(f));
        }
        folder = true;
    }
    // HiddenSciGuard::setText: SCI_DELETERANGE, SCI_SETCODEPAGE, SCI_APPENDTEXT
    void setText(const std::string& s, int codepage) {
        if (d->Length()) d->DeleteChars(0, d->Length());
        setCodePage(codepage);
        if (!s.empty()) d->InsertString(0, s.data(), static_cast<Sci::Position>(s.size()));
    }
    Sci::Position length() const { return d->Length(); }
    Sci::Line lineFromPos(Sci::Position p) const { return d->SciLineFromPosition(p); }
    Sci::Position positionAfter(Sci::Position p) const { return d->NextPosition(p, 1); }
    // SCI_SETTARGETRANGE + SCI_SEARCHINTARGET (+ SCI_GETTARGETEND)
    Sci::Position search(Sci::Position from, Sci::Position to, const std::string& s, int flags, Sci::Position& end) {
        ensureFolder();
        Sci::Position len = static_cast<Sci::Position>(s.size());
        const Sci::Position p = d->FindText(from, to, s.c_str(), static_cast<FindOption>(flags), &len);
        if (p >= 0) end = p + len;
        return p;
    }
};

// MultiReplacePanel.h
constexpr int SKIPCRLFASONE = 0x08000000;
constexpr int EMPTYMATCH_ALL = 0x40000000;

struct Query {
    const char* key;    // for --queries
    const char* name;
    std::string text;
    int flags;          // MultiReplace::buildSearchFlags(wholeWord false, matchCase true, regex, false, false)
};

static std::vector<Query> queries()
{
    const int regex = SCFIND_MATCHCASE | SCFIND_REGEXP | SCFIND_POSIX | SKIPCRLFASONE | EMPTYMATCH_ALL;
    return {
        { "absent", "absent", "zqxjvqk", SCFIND_MATCHCASE },
        { "fi", "Fi regex", "Fi", regex },
        { "in", "in", "in", SCFIND_MATCHCASE },
    };
}

// What one run found; equal in every mode
struct Found {
    size_t files = 0, filled = 0, skipped = 0, binary = 0;
    uint64_t hits = 0;
    bool operator==(const Found& o) const {
        return files == o.files && filled == o.filled && skipped == o.skipped && binary == o.binary && hits == o.hits;
    }
};

// load: time in FileLoad::load summed over all threads, waits for the stick included; stick, jumps: Stick::Busy
struct Times {
    double total = 0, wait = 0, load = 0, fill = 0, search = 0, stick = 0;
    size_t jumps = 0;
};

// One file as the consumer gets it, from its own load or from the pool
struct Item {
    FileLoad::SkipReason skip = FileLoad::SkipReason::None;
    int codepage = 0;
    bool mayMatch = false;
    const std::string* text = nullptr;
};

// The loop body of handleFindInFiles from the prefilter on
struct Consumer {
    SciDoc doc;
    struct RawHit { Sci::Position pos, len; Sci::Line line; };
    std::vector<RawHit> raw;
    Found found;
    Times times;

    void take(const Item& f, const Query& q) {
        ++found.files;
        if (f.skip != FileLoad::SkipReason::None) {
            ++found.skipped;
            if (f.skip == FileLoad::SkipReason::Binary) ++found.binary;
            return;
        }
        if (!f.mayMatch) return;
        auto t0 = Clock::now();
        doc.setText(*f.text, f.codepage);
        times.fill += msSince(t0);
        ++found.filled;

        t0 = Clock::now();
        const Sci::Position length = doc.length();
        Sci::Position pos = 0;
        while (pos <= length) {
            Sci::Position end = 0;
            const Sci::Position at = doc.search(pos, length, q.text, q.flags, end);
            if (at < 0) break;
            raw.push_back({ at, end - at, doc.lineFromPos(at) });
            pos = (end > at) ? end : doc.positionAfter(at);   // advanceAfterMatch
            if (end <= at && pos <= at) break;
        }
        found.hits += raw.size();
        raw.clear();
        times.search += msSince(t0);
    }
};

// threads 0: serial as in the plugin with LoadThreads=0; otherwise FileLoadPool with that many workers
static void runOnce(const std::vector<fs::path>& files, const Query& q, unsigned threads, uint64_t budgetMB, Found& found, Times& times)
{
    Consumer consumer;
    HiddenSciGuard guard;
    g_loadUs = 0;
    g_stick.resetBusy();
    const auto t0 = Clock::now();
    if (threads == 0) {
        std::string text;   // file after file loads into the same string
        for (const auto& path : files) {
            Item item;
            Encoding::EncodingInfo enc;
            HiddenSciGuard::LoadKind kind = HiddenSciGuard::LoadKind::Text;
            item.skip = guard.loadTextFile(path, text, enc, kind);
            if (item.skip == FileLoad::SkipReason::None) {
                item.codepage = FileLoad::codepageFor(kind, enc);
                LiteralPrefilter prefilter(text);
                item.mayMatch = prefilter.mayMatch(q.text, q.flags);
                item.text = &text;
            }
            consumer.take(item, q);
        }
    }
    else {
        FileLoadPool::Limits limits;
        if (budgetMB > 0) limits.budget = budgetMB << 20;
        FileLoadPool pool(files, guard.loadOptions(), { { q.text, q.text, q.flags } }, threads, limits);
        FileLoadPool::File f;
        for (;;) {
            const auto w0 = Clock::now();
            const FileLoadPool::Next got = pool.next(f);
            consumer.times.wait += msSince(w0);
            if (got == FileLoadPool::Next::Pending) continue;
            if (got == FileLoadPool::Next::Done) break;
            Item item;
            item.skip = f.result.skip;
            item.codepage = FileLoad::codepageFor(f.result.kind, f.result.enc);
            item.mayMatch = f.handOver;
            item.text = &f.text;
            consumer.take(item, q);
            pool.recycle(std::move(f.text));
        }
    }
    consumer.times.total = msSince(t0);
    consumer.times.load = static_cast<double>(g_loadUs.load()) / 1000.0;
    const Stick::Busy busy = g_stick.busy();
    consumer.times.stick = busy.ms;
    consumer.times.jumps = busy.jumps;
    found = consumer.found;
    times = consumer.times;
}

static std::vector<fs::path> listFiles(const fs::path& root)
{
    // collectScanFiles with filter *.*, subfolders, hidden folders left out
    DirectoryWalk::Options options;
    options.root = { true, true };
    options.subfolder = [](const fs::path&, bool hidden) {
        return hidden ? DirectoryWalk::Take{ false, false } : DirectoryWalk::Take{ true, true };
    };
    DirectoryWalk::Result walk = DirectoryWalk::collect(root, options);
    return walk.status == DirectoryWalk::Status::Done ? std::move(walk.files) : std::vector<fs::path>{};
}

static double median(std::vector<double> v)
{
    if (v.empty()) return 0;
    std::sort(v.begin(), v.end());
    const size_t n = v.size();
    return n % 2 ? v[n / 2] : (v[n / 2 - 1] + v[n / 2]) / 2;
}

static std::string utf8(const std::wstring& w)
{
    if (w.empty()) return {};
    const int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
    std::string s(static_cast<size_t>(n > 0 ? n : 0), '\0');
    if (n > 0) WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), s.data(), n, nullptr, nullptr);
    return s;
}

static fs::path programFolder()
{
    wchar_t buf[MAX_PATH] = {};
    const DWORD n = GetModuleFileNameW(nullptr, buf, MAX_PATH);
    return (n > 0 && n < MAX_PATH) ? fs::path(buf).parent_path() : fs::current_path();
}

static std::vector<std::wstring> splitList(const std::wstring& list)
{
    std::vector<std::wstring> parts;
    for (size_t p = 0; p <= list.size();) {
        const size_t q = list.find(L',', p);
        const std::wstring part = list.substr(p, q == std::wstring::npos ? std::wstring::npos : q - p);
        if (!part.empty()) parts.push_back(part);
        if (q == std::wstring::npos) break;
        p = q + 1;
    }
    return parts;
}

static std::string busName(STORAGE_BUS_TYPE bus)
{
    switch (bus) {
    case BusTypeUsb: return "USB";
    case BusTypeSata: return "SATA";
    case BusTypeNvme: return "NVMe";
    case BusTypeSd: return "SD";
    case BusTypeMmc: return "MMC";
    case BusType1394: return "FireWire";
    case BusTypeScsi: return "SCSI";
    case BusTypeAta: return "ATA";
    case BusTypeAtapi: return "ATAPI";
    case BusTypeSas: return "SAS";
    case BusTypeRAID: return "RAID";
    case BusTypeiScsi: return "iSCSI";
    case BusTypeVirtual: return "virtual";
    case BusTypeFileBackedVirtual: return "virtual disk file";
    case BusTypeSpaces: return "Storage Spaces";
    default: return "bus " + std::to_string(static_cast<int>(bus));
    }
}

// What Windows reports about the drive of a folder: the facts LoadThreads=auto decides on
static std::string describeDrive(const fs::path& folder)
{
    wchar_t volumePath[MAX_PATH + 1] = {};
    if (!GetVolumePathNameW(folder.c_str(), volumePath, MAX_PATH + 1)) return utf8(folder.wstring()) + " unknown";
    std::string text = utf8(volumePath);
    switch (GetDriveTypeW(volumePath)) {
    case DRIVE_REMOVABLE: text += " removable"; break;
    case DRIVE_FIXED: text += " fixed"; break;
    case DRIVE_CDROM: text += " optical"; break;
    case DRIVE_RAMDISK: text += " RAM disk"; break;
    case DRIVE_REMOTE: return text + " network";
    default: return text + " unknown type";
    }

    wchar_t volumeName[MAX_PATH + 1] = {};
    if (!GetVolumeNameForVolumeMountPointW(volumePath, volumeName, MAX_PATH + 1)) return text + ", no volume";
    std::wstring device = volumeName;
    if (!device.empty() && device.back() == L'\\') device.pop_back();
    const HANDLE h = CreateFileW(device.c_str(), 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
    if (h == INVALID_HANDLE_VALUE) return text + ", volume not opened";

    STORAGE_PROPERTY_QUERY query{};
    query.QueryType = PropertyStandardQuery;
    query.PropertyId = StorageDeviceProperty;
    alignas(8) unsigned char buffer[1024] = {};
    DWORD bytes = 0;
    if (DeviceIoControl(h, IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof query, buffer, sizeof buffer, &bytes, nullptr)
        && bytes >= sizeof(STORAGE_DEVICE_DESCRIPTOR)) {
        const auto* d = reinterpret_cast<const STORAGE_DEVICE_DESCRIPTOR*>(buffer);
        auto field = [&](DWORD offset) {
            if (offset == 0 || offset >= bytes) return std::string();
            const char* s = reinterpret_cast<const char*>(buffer + offset);
            std::string v(s, strnlen(s, bytes - offset));
            const size_t a = v.find_first_not_of(' ');
            return a == std::string::npos ? std::string() : v.substr(a, v.find_last_not_of(' ') - a + 1);
        };
        const std::string vendor = field(d->VendorIdOffset), product = field(d->ProductIdOffset);
        text += ", " + busName(d->BusType);
        if (!vendor.empty() || !product.empty()) text += ", " + vendor + (vendor.empty() || product.empty() ? "" : " ") + product;
    }
    else text += ", bus not reported";

    DEVICE_SEEK_PENALTY_DESCRIPTOR penalty{};
    query.PropertyId = StorageDeviceSeekPenaltyProperty;
    const bool known = DeviceIoControl(h, IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof query, &penalty, sizeof penalty, &bytes, nullptr)
        && bytes >= sizeof penalty;
    CloseHandle(h);
    return text + (known ? (penalty.IncursSeekPenalty ? ", seek penalty yes" : ", seek penalty no") : ", seek penalty not reported");
}

static void listDrives()
{
    const UINT oldMode = SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);   // no "insert a disk" box
    const DWORD mask = GetLogicalDrives();
    for (int i = 0; i < 26; ++i) {
        if (!(mask & (DWORD{ 1 } << i))) continue;
        const wchar_t root[] = { static_cast<wchar_t>(L'A' + i), L':', L'\\', 0 };
        const unsigned workers = FileLoadPool::workersFor(-1, root);
        std::printf("%s; LoadThreads=auto takes %u worker%s\n", describeDrive(root).c_str(), workers, workers == 1 ? "" : "s");
    }
    SetErrorMode(oldMode);
}

// The first n hardware threads the process may use
static bool limitCpus(unsigned n)
{
    DWORD_PTR process = 0, system = 0;
    if (n == 0 || !GetProcessAffinityMask(GetCurrentProcess(), &process, &system)) return false;
    DWORD_PTR mask = 0;
    unsigned taken = 0;
    for (unsigned bit = 0; bit < sizeof(DWORD_PTR) * 8 && taken < n; ++bit) {
        const DWORD_PTR cpu = DWORD_PTR{ 1 } << bit;
        if (process & cpu) {
            mask |= cpu;
            ++taken;
        }
    }
    return taken == n && SetProcessAffinityMask(GetCurrentProcess(), mask);
}

static std::string modeName(unsigned threads, uint64_t budgetMB = 0)
{
    const std::string name = threads == 0 ? "serial" : std::to_string(threads) + (threads == 1 ? " worker" : " workers");
    return (threads > 0 && budgetMB > 0) ? name + ", " + std::to_string(budgetMB) + " MB" : name;
}

static void writeRow(FILE* csv, const std::string& dir, const char* query, unsigned threads, int round, const Found& found,
                     const Times& times, unsigned hw, const std::string& device)
{
    SYSTEMTIME st{};
    GetLocalTime(&st);
    std::fprintf(csv, "%04u-%02u-%02u %02u:%02u:%02u,%s,%s,%u,%d,%zu,%zu,%zu,%zu,%llu,%.3f,%.3f,%.3f,%.3f,%.3f,%.3f,%zu,%u,%s\n",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond,
        dir.c_str(), query, threads, round, found.files, found.filled, found.skipped, found.binary,
        static_cast<unsigned long long>(found.hits), times.total, times.wait, times.load, times.fill, times.search,
        times.stick, times.jumps, hw, device.c_str());
}

// --cold: every mode reads its own folder for the first time, as the first search after the file cache has
// emptied. The mode with the most workers goes first, so a slow start costs it, not serial. Round 0 in the
// CSV is the first read, rounds 1 to 3 read the same folders again, warm.
static bool runCold(const std::vector<std::string>& names, const std::vector<std::vector<fs::path>>& fileLists,
                    const Query& q, const std::vector<unsigned>& threads, const std::vector<uint64_t>& budgets, FILE* csv,
                    unsigned hw, const std::string& device, std::string& summary)
{
    constexpr int WARM_ROUNDS = 3;
    std::vector<size_t> order(threads.size());
    std::iota(order.begin(), order.end(), size_t{ 0 });
    std::stable_sort(order.begin(), order.end(), [&](size_t a, size_t b) { return threads[a] > threads[b]; });

    std::vector<Times> first(threads.size());
    std::vector<std::vector<double>> warm(threads.size());
    std::vector<Found> found(threads.size());
    bool equal = true;
    for (int round = 0; round <= WARM_ROUNDS; ++round) {
        for (size_t i : order) {
            Found f;
            Times t;
            runOnce(fileLists[i], q, threads[i], budgets[i], f, t);
            if (round == 0) {
                first[i] = t;
                found[i] = f;
                std::printf("first read of %s, %s: %.0f ms\n", names[i].c_str(), modeName(threads[i], budgets[i]).c_str(), t.total);
            }
            else {
                warm[i].push_back(t.total);
                equal = equal && f == found[i];
            }
            writeRow(csv, names[i], q.name, threads[i], round, f, t, hw, device);
        }
        std::fflush(csv);
        if (round > 0) std::printf("warm round %d of %d done\n", round, WARM_ROUNDS);
    }
    for (const Found& f : found) equal = equal && f == found.front();

    const size_t serial = static_cast<size_t>(std::find(threads.begin(), threads.end(), 0u) - threads.begin());
    char line[512];
    std::snprintf(line, sizeof line, "\n%s: %zu files in each folder, %zu filled, %zu skipped, %llu hits\n",
        q.name, found.front().files, found.front().filled, found.front().skipped, static_cast<unsigned long long>(found.front().hits));
    summary += line;
    summary += "  mode              folder    first ms   ms/file  / serial   warm ms  first / warm\n";

    // A folder whose first read took less than twice as long as a warm one was still cached
    std::vector<size_t> rows(threads.size());
    std::iota(rows.begin(), rows.end(), size_t{ 0 });
    std::stable_sort(rows.begin(), rows.end(), [&](size_t a, size_t b) { return threads[a] < threads[b]; });
    std::string cached;
    double coldMin = 0, coldMax = 0;
    for (size_t i : rows) {
        const double files = static_cast<double>((std::max)(found[i].files, size_t{ 1 }));
        const double warmMs = median(warm[i]);
        const double factor = warmMs > 0 ? first[i].total / warmMs : 0.0;
        std::snprintf(line, sizeof line, "  %-16s  %-7s %9.1f  %8.2f  %7.2f  %8.1f  %12.1f%s\n",
            modeName(threads[i], budgets[i]).c_str(), names[i].c_str(), first[i].total, first[i].total / files,
            first[i].total / first[serial].total, warmMs, factor, factor < 2 ? "  cached" : "");
        summary += line;
        if (factor < 2) {
            cached += (cached.empty() ? "" : ", ") + names[i] + " (" + modeName(threads[i], budgets[i]) + ")";
            continue;
        }
        coldMin = (coldMin == 0) ? factor : (std::min)(coldMin, factor);
        coldMax = (std::max)(coldMax, factor);
    }
    if (cached.empty())
        std::snprintf(line, sizeof line, "\nThe first read was cold in every folder: %.0f to %.0f times as long as warm.\n", coldMin, coldMax);
    else
        std::snprintf(line, sizeof line, "\nStill cached, so these first reads do not count: %s. Run it again after a restart, or plug a USB stick in again.\n", cached.c_str());
    summary += line;
    return equal;
}

struct Cell {
    std::string dir;
    unsigned autoWorkers = 0;   // what LoadThreads=auto picks for the folder
    const Query* query = nullptr;
    std::map<unsigned, std::vector<Times>> runs;   // per mode, measured rounds in order
    std::vector<Found> found;
};

int wmain(int argc, wchar_t** argv)
{
    std::vector<fs::path> dirs;
    int rounds = 8;
    std::vector<unsigned> threads;
    std::vector<uint64_t> budgets;   // MB per mode, --cold only
    std::vector<std::wstring> keys;
    std::wstring device = L"none";
    unsigned cpus = 0;
    bool cold = false;
    bool drives = false;
    fs::path out = programFolder() / L"thread_bench.csv";

    for (int i = 1; i < argc; ++i) {
        const std::wstring a = argv[i];
        auto value = [&]() -> std::wstring { return (i + 1 < argc) ? argv[++i] : L""; };
        if (a == L"--dir") dirs.emplace_back(value());
        else if (a == L"--rounds") rounds = (std::max)(2, _wtoi(value().c_str()));
        else if (a == L"--out") out = value();
        else if (a == L"--device") device = value();
        else if (a == L"--queries") keys = splitList(value());
        else if (a == L"--cpus") cpus = static_cast<unsigned>((std::max)(0, _wtoi(value().c_str())));
        else if (a == L"--cold") cold = true;
        else if (a == L"--drives") drives = true;
        else if (a == L"--budgets")
            for (const std::wstring& b : splitList(value())) budgets.push_back(static_cast<uint64_t>((std::max)(0, _wtoi(b.c_str()))));
        else if (a == L"--threads") {
            threads.clear();
            for (const std::wstring& t : splitList(value())) threads.push_back(static_cast<unsigned>(_wtoi(t.c_str())));
        }
        else {
            std::fprintf(stderr, "usage: find_in_files_pool_bench [--dir PATH]... [--rounds N] [--threads 0,1,2,...]"
                " [--queries absent,fi,in] [--device none|usb2|usb3] [--cpus N] [--cold [--budgets 0,0,1,...]] [--drives] [--out CSV]\n");
            return 2;
        }
    }
    if (!g_stick.set(device)) {
        std::fprintf(stderr, "unknown device %s, use none, usb2 or usb3\n", utf8(device).c_str());
        return 2;
    }
    if (drives) {
        listDrives();
        if (dirs.empty()) return 0;
        std::printf("\n");
    }
    const unsigned allHw = std::thread::hardware_concurrency();
    if (cpus > 0 && !limitCpus(cpus)) {
        std::fprintf(stderr, "cannot keep the program on %u of %u hardware threads\n", cpus, allHw);
        return 2;
    }
    const unsigned hw = std::thread::hardware_concurrency();
    if (threads.empty()) {
        threads = { 0, 1, 2, 4, 8 };
        if (hw > 9) threads.push_back(hw - 1);
    }
    if (dirs.empty())
        for (const wchar_t* name : { L"source-text", L"source" })
            if (fs::is_directory(programFolder() / name)) dirs.push_back(programFolder() / name);
    if (dirs.empty() || std::find(threads.begin(), threads.end(), 0u) == threads.end()) {
        std::fprintf(stderr, "no folder found, or serial mode (0) missing in --threads\n");
        return 2;
    }

    std::vector<Query> qs;
    for (const Query& q : queries()) {
        const std::wstring key(q.key, q.key + std::strlen(q.key));
        if (keys.empty() || std::find(keys.begin(), keys.end(), key) != keys.end()) qs.push_back(q);
    }
    if (qs.empty()) {
        std::fprintf(stderr, "no query selected, use absent, fi or in\n");
        return 2;
    }
    if (cold && (dirs.size() != threads.size() || qs.size() != 1)) {
        std::fprintf(stderr, "--cold needs one --dir per mode in --threads and one query\n");
        return 2;
    }
    if (budgets.empty()) budgets.assign(threads.size(), 0);
    if (!cold && std::any_of(budgets.begin(), budgets.end(), [](uint64_t b) { return b > 0; })) {
        std::fprintf(stderr, "--budgets only with --cold\n");
        return 2;
    }
    if (budgets.size() != threads.size()) {
        std::fprintf(stderr, "--budgets needs one value per mode in --threads\n");
        return 2;
    }

    std::vector<Cell> cells;
    std::vector<std::vector<fs::path>> fileLists;
    std::string driveLines;   // for the summary, one per drive
    std::set<std::string> drivesSeen;
    for (const auto& dir : dirs) {
        const auto t0 = Clock::now();
        fileLists.push_back(listFiles(dir));
        const double listMs = msSince(t0);
        if (g_stick.on())
            for (const fs::path& file : fileLists.back()) {
                std::error_code ec;
                const uint64_t size = fs::file_size(file, ec);
                if (!ec) g_sizes[file.native()] = size;
            }
        const unsigned autoWorkers = FileLoadPool::workersFor(-1, dir);
        std::printf("%s: %zu files, listed in %.0f ms, LoadThreads=auto takes %u worker%s\n", utf8(dir.wstring()).c_str(),
            fileLists.back().size(), listMs, autoWorkers, autoWorkers == 1 ? "" : "s");
        const std::string drive = describeDrive(dir);
        if (drivesSeen.insert(drive).second) {
            std::printf("  drive %s\n", drive.c_str());
            driveLines += "drive " + drive + "\n";
        }
        for (const Query& q : qs) cells.push_back({ utf8(dir.filename().wstring()), autoWorkers, &q, {}, {} });
    }

    FILE* csv = _wfopen(out.c_str(), L"ab");
    if (!csv) { std::fprintf(stderr, "cannot write %s\n", utf8(out.wstring()).c_str()); return 2; }
    std::fseek(csv, 0, SEEK_END);
    if (_ftelli64(csv) == 0)
        std::fputs("time,folder,query,threads,round,files,filled,skipped,binary,hits,total_ms,wait_ms,load_ms,fill_ms,search_ms,stick_ms,"
            "jumps,hardware_threads,device\n", csv);

    const fs::path summaryPath = out.parent_path() / (out.stem().wstring() + L"_summary.txt");
    SYSTEMTIME now{};
    if (cold) {
        std::vector<std::string> names;
        for (const auto& dir : dirs) names.push_back(utf8(dir.filename().wstring()));
        GetLocalTime(&now);
        char head[256];
        std::snprintf(head, sizeof head, "\n== %04u-%02u-%02u %02u:%02u, first read, every mode on its own folder, %u of %u hardware threads\n",
            now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, hw, allHw);
        std::string summary = head + driveLines;
        const bool equal = runCold(names, fileLists, qs.front(), threads, budgets, csv, hw, utf8(device), summary);
        std::fclose(csv);
        summary += equal ? "All modes found the same.\n" : "MISMATCH: the modes did not find the same, see the CSV.\n";
        std::fputs(summary.c_str(), stdout);
        FILE* sum = _wfopen(summaryPath.c_str(), L"ab");
        if (sum) { std::fputs(summary.c_str(), sum); std::fclose(sum); }
        return equal ? 0 : 1;
    }

    std::mt19937 rng(20261010);
    std::printf("%d rounds, device %s, modes:", rounds, utf8(device).c_str());
    for (unsigned t : threads) std::printf(" %u", t);
    std::printf(" (0 = serial), %u of %u hardware threads\n", hw, allHw);

    for (int round = 0; round < rounds; ++round) {
        for (size_t c = 0; c < cells.size(); ++c) {
            Cell& cell = cells[c];
            const auto& files = fileLists[c / qs.size()];
            std::vector<unsigned> order = threads;
            std::shuffle(order.begin(), order.end(), rng);
            for (unsigned t : order) {
                Found found;
                Times times;
                runOnce(files, *cell.query, t, 0, found, times);
                cell.found.push_back(found);
                if (round > 0) cell.runs[t].push_back(times);
                writeRow(csv, cell.dir, cell.query->name, t, round, found, times, hw, utf8(device));
            }
        }
        std::fflush(csv);
        std::printf("round %d of %d done\n", round + 1, rounds);
    }
    std::fclose(csv);

    // Summary: median per mode, and the median of the per-round ratios to serial in the same round
    bool equal = true;
    GetLocalTime(&now);
    char line[512];
    std::snprintf(line, sizeof line, "\n== %04u-%02u-%02u %02u:%02u, device %s, %d rounds (the first left out), %u of %u hardware threads\n",
        now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, utf8(device).c_str(), rounds, hw, allHw);
    std::string summary = line + driveLines;
    for (const Cell& cell : cells) {
        for (const Found& f : cell.found) equal = equal && f == cell.found.front();
        const Found& f = cell.found.front();
        std::snprintf(line, sizeof line, "\n%s, %s: %zu files, %zu filled, %zu skipped, %llu hits; auto takes %u worker%s\n",
            cell.dir.c_str(), cell.query->name, f.files, f.filled, f.skipped, static_cast<unsigned long long>(f.hits),
            cell.autoWorkers, cell.autoWorkers == 1 ? "" : "s");
        summary += line;
        summary += g_stick.on() ? "  mode       total ms  / serial   wait ms   load ms   fill ms   search ms   stick ms   jumps\n"
                                : "  mode       total ms  / serial   wait ms   load ms   fill ms   search ms\n";
        const std::vector<Times>& serial = cell.runs.at(0);
        for (const auto& [t, runs] : cell.runs) {
            std::vector<double> total, ratio, wait, load, fill, search, stick, jumps;
            for (size_t r = 0; r < runs.size(); ++r) {
                total.push_back(runs[r].total);
                ratio.push_back(runs[r].total / serial[r].total);
                wait.push_back(runs[r].wait);
                load.push_back(runs[r].load);
                fill.push_back(runs[r].fill);
                search.push_back(runs[r].search);
                stick.push_back(runs[r].stick);
                jumps.push_back(static_cast<double>(runs[r].jumps));
            }
            std::snprintf(line, sizeof line, "  %-9s %9.1f  %7.2f  %8.1f  %8.1f  %8.1f  %10.1f",
                modeName(t).c_str(),
                median(total), median(ratio), median(wait), median(load), median(fill), median(search));
            summary += line;
            if (g_stick.on()) {
                std::snprintf(line, sizeof line, "  %9.1f  %6.0f", median(stick), median(jumps));
                summary += line;
            }
            summary += "\n";
        }
    }
    summary += equal ? "\nAll modes found the same.\n" : "\nMISMATCH: the modes did not find the same, see the CSV.\n";
    std::fputs(summary.c_str(), stdout);
    FILE* sum = _wfopen(summaryPath.c_str(), L"ab");
    if (sum) { std::fputs(summary.c_str(), sum); std::fclose(sum); }
    return equal ? 0 : 1;
}
