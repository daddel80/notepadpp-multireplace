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

// QA for loading ahead on worker threads (FileLoadPool), where Find/Replace in Files take their files from.
//   Linux: FileLoadPool.cpp against a simulated loader (FileLoad::load replaced at link time), under
//     ThreadSanitizer and under AddressSanitizer with UndefinedBehaviorSanitizer:
//     L1 every file once and in list order, with its skip, kind, text and prefilter answers
//     L2 a file over the budget is read only once it is the next one, and the budget is kept
//     L3 next() reports Pending while the next file still loads, nothing is lost meanwhile
//     L4 cancel at random points: next() ends, the workers stop, a large read stops between blocks
//     L5 an exception in a load becomes a skip (bad_alloc: too large, else unreadable)
//     L6 buffers are reused, but none over 1 MB
//     L7 the number of workers: the setting as given (at most 8), else by the drive: 8 on a network
//        drive, 4 on a fast one, 1 on a slow one
//   Windows (also under Wine): the pool against loading on the UI thread as handleFindInFiles and
//     handleReplaceInFiles do it, file by file on a generated corpus and on the folders given with
//     --dir: skip, kind, encoding, text and prefilter answers equal, with 1, 2, 4 and 8 workers.
//     Every folder is read once beforehand, so all times shown per mode come from the file cache.
// Build and run: file_load_pool_qa.sh

#ifndef _WIN32

#include <atomic>
#include <chrono>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <memory>
#include <mutex>
#include <new>
#include <random>
#include <stdexcept>
#include <string>
#include <thread>
#include <vector>

#include "FileLoadPool.h"
#include "LiteralPrefilter.h"

namespace {

    // What the simulated loader does with one file
    struct Spec {
        size_t size = 0;            // text bytes, also the size reported before the full read
        FileLoad::SkipReason skip = FileLoad::SkipReason::None;
        bool raw = false;           // raw bytes in codepage 0, else UTF-8 text
        int delayUs = 0;
        int throws = 0;             // after the gate: 1 bad_alloc, 2 another exception
        unsigned words = 0;         // which of the words the text holds
    };

    const char* const WORDS_UTF8[] = { "alpha", "Beta", "gr\xC3\xBC" "n" };
    const char* const WORDS_ANSI[] = { "alpha", "Beta", "gr\xFC" "n" };

    struct Run {
        std::vector<Spec> specs;
        const std::vector<FileLoadPool::Pattern>* patterns = nullptr;
        uint64_t budget = 0;
        std::atomic<size_t> taken{ 0 };          // files the consumer has taken
        std::atomic<uint64_t> outstanding{ 0 };  // bytes past the gate, kept for the consumer and not taken yet
        std::atomic<uint64_t> nextBytes{ 0 };    // the same for files that passed as the next one, budget or not
        std::unique_ptr<std::atomic<int>[]> counted;   // 1 in outstanding, 2 in nextBytes
        std::mutex booking;                            // a worker's check and the consumer's booking, one at a time
        std::atomic<int> gateViolations{ 0 };
        std::atomic<int> bigBuffers{ 0 };        // loads that were handed a buffer over 1 MB
        std::atomic<int> reusedBuffers{ 0 };
    };

    Run* g_run = nullptr;
    FileLoad::DriveKind g_drive = FileLoad::DriveKind::Fast;

    // The pool hands a text over only if no patterns are given or one of them may match
    bool handedOver(const Run& run, const Spec& spec, const std::string& text)
    {
        if (run.patterns->empty()) return true;
        LiteralPrefilter prefilter(text);
        for (const auto& p : *run.patterns)
            if (prefilter.mayMatch(spec.raw ? p.ansi : p.utf8, p.searchFlags)) return true;
        return false;
    }

    void fillText(size_t index, const Spec& spec, std::string& out)
    {
        out.assign(spec.size, static_cast<char>('c' + index % 20));
        size_t pos = 0;
        for (unsigned w = 0; w < 3; ++w) {
            if (!(spec.words & (1u << w))) continue;
            const char* word = spec.raw ? WORDS_ANSI[w] : WORDS_UTF8[w];
            const size_t len = std::strlen(word);
            if (pos + len > out.size()) break;
            out.replace(pos, len, word);
            pos += len + 7;
        }
    }
}

// The simulated loader, in place of FileLoad.cpp
namespace FileLoad {

    Result load(const std::filesystem::path& fp, const Options&, std::string& content, const Pace* pace)
    {
        Run& run = *g_run;
        const size_t index = static_cast<size_t>(std::stoul(fp.string()));
        const Spec& spec = run.specs[index];
        if (content.capacity() > (size_t{ 1 } << 20)) ++run.bigBuffers;
        if (content.capacity() >= 4096) ++run.reusedBuffers;
        content.clear();

        Result r;
        if (spec.delayUs) std::this_thread::sleep_for(std::chrono::microseconds(spec.delayUs));
        if (spec.skip != SkipReason::None) {
            r.skip = spec.skip;
            return r;
        }
        if (pace && pace->beforeRead) {
            if (!pace->beforeRead(spec.size)) {
                r.skip = SkipReason::Unreadable;
                return r;
            }
            // Under the consumer's lock: the file it may be taking right now is still counted here, not in the
            // pool any more. A file passing as the next one may pass over the budget; it is counted apart.
            std::lock_guard<std::mutex> lock(run.booking);
            const size_t taken = run.taken.load();
            const uint64_t window = (taken < run.specs.size() && run.counted[taken].load() == 1) ? run.specs[taken].size : 0;
            const bool next = index <= taken + 1;
            if (next) {
                run.nextBytes += spec.size;
                run.counted[index].store(2);
            }
            else {
                const uint64_t before = run.outstanding.fetch_add(spec.size);
                run.counted[index].store(1);
                if (spec.size > run.budget || before + spec.size > run.budget + window)
                    ++run.gateViolations;
            }
        }
        // Like the pool, nothing stays counted for a file that is not handed over
        auto release = [&] {
            const int counted = run.counted[index].exchange(0);
            if (counted == 1) run.outstanding -= spec.size;
            if (counted == 2) run.nextBytes -= spec.size;
        };
        for (int block = 0; spec.size > run.budget && block < 4; ++block) {
            if (pace && pace->stop && pace->stop->load()) {
                release();
                r.skip = SkipReason::Unreadable;
                return r;
            }
            std::this_thread::sleep_for(std::chrono::microseconds(200));
        }
        if (spec.throws) release();
        if (spec.throws == 1) throw std::bad_alloc();
        if (spec.throws == 2) throw std::runtime_error("simulated");
        fillText(index, spec, content);
        if (!handedOver(run, spec, content)) release();
        r.kind = spec.raw ? LoadKind::RawBytes : LoadKind::Text;
        r.enc.kind = spec.raw ? Encoding::Kind::ANSI : Encoding::Kind::UTF8;
        return r;
    }

    DriveKind driveKind(const std::filesystem::path&) { return g_drive; }
}

static int failures = 0;
static void CHECK(const std::string& what, bool ok)
{
    std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str());
    if (!ok) ++failures;
}

// The file as the pool has to hand it over
static bool expected(const Run& run, size_t index, const std::vector<FileLoadPool::Pattern>& patterns, const FileLoadPool::File& f)
{
    const Spec& spec = run.specs[index];
    FileLoad::SkipReason skip = spec.skip;
    if (skip == FileLoad::SkipReason::None && spec.throws)
        skip = (spec.throws == 1) ? FileLoad::SkipReason::TooLarge : FileLoad::SkipReason::Unreadable;
    if (f.result.skip != skip) return false;
    if (skip != FileLoad::SkipReason::None) return !f.handOver && f.text.empty();

    std::string text;
    fillText(index, spec, text);
    if (f.result.kind != (spec.raw ? FileLoad::LoadKind::RawBytes : FileLoad::LoadKind::Text)) return false;

    LiteralPrefilter prefilter(text);
    std::vector<char> mayMatch;
    bool handOver = patterns.empty();
    for (const auto& p : patterns) {
        const bool m = prefilter.mayMatch(spec.raw ? p.ansi : p.utf8, p.searchFlags);
        mayMatch.push_back(m ? 1 : 0);
        handOver = handOver || m;
    }
    return f.mayMatch == mayMatch && f.handOver == handOver && (handOver ? f.text == text : f.text.empty());
}

int main()
{
    // A pool that stops handing out files fails the test instead of hanging it
    std::atomic<int> progress{ 0 };
    std::thread([&progress] {
        for (int last = -1;;) {
            std::this_thread::sleep_for(std::chrono::seconds(20));
            const int now = progress.load();
            if (now == last) { std::printf("FAIL no progress in run %d (deadlock)\n", now); std::fflush(stdout); std::_Exit(1); }
            last = now;
        }
    }).detach();

    const std::vector<FileLoadPool::Pattern> words = {
        { "alpha", "alpha", SCFIND_MATCHCASE },
        { "beta", "beta", 0 },
        { "gr\xFC" "n", "gr\xC3\xBC" "n", SCFIND_MATCHCASE },
        { "a.p", "a.p", SCFIND_REGEXP },
    };

    std::mt19937 rng(20261010);
    bool ordered = true, budgetKept = true, noBigBuffer = true, completed = true;
    size_t files = 0, cancels = 0, skips = 0, thrown = 0, large = 0, pendings = 0, reused = 0;
    const int RUNS = 300;
    for (int r = 0; r < RUNS; ++r) {
        progress = r;
        Run run;
        const size_t count = rng() % 300;
        const unsigned workers = 1 + rng() % FileLoadPool::MAX_WORKERS;
        const size_t ahead = 1 + rng() % 40;
        run.budget = 1 + rng() % 40000;
        run.counted.reset(new std::atomic<int>[count + 1]());
        for (size_t i = 0; i < count; ++i) {
            Spec s;
            const unsigned roll = rng() % 1000;
            if (roll < 5) s.size = (size_t{ 1 } << 20) + rng() % 100000;      // over 1 MB: never reused
            else if (roll < 35) s.size = static_cast<size_t>(run.budget) + 1 + rng() % 20000;
            else s.size = rng() % 6000;
            if (rng() % 100 < 5) s.skip = static_cast<FileLoad::SkipReason>(1 + rng() % 5);
            s.raw = rng() % 5 == 0;
            s.delayUs = (rng() % 1000 < 2) ? 70000 : static_cast<int>(rng() % 150);
            s.throws = (rng() % 100 == 0) ? 1 + static_cast<int>(rng() % 2) : 0;
            s.words = rng() % 8;
            large += s.size > run.budget;
            skips += s.skip != FileLoad::SkipReason::None;
            thrown += s.throws != 0;
            run.specs.push_back(s);
        }
        const size_t cancelAt = (rng() % 4 == 0) ? rng() % (count + 1) : SIZE_MAX;
        const std::vector<FileLoadPool::Pattern> patterns = (rng() % 2) ? words : std::vector<FileLoadPool::Pattern>{};
        run.patterns = &patterns;

        std::vector<std::filesystem::path> paths;
        for (size_t i = 0; i < count; ++i) paths.emplace_back(std::to_string(i));
        g_run = &run;
        {
            FileLoadPool pool(paths, FileLoad::Options{}, patterns, workers, FileLoadPool::Limits{ ahead, run.budget });
            size_t expect = 0;
            FileLoadPool::File f;
            for (;;) {
                const FileLoadPool::Next got = pool.next(f);
                if (got == FileLoadPool::Next::Pending) { ++pendings; continue; }
                if (got == FileLoadPool::Next::Done) break;
                if (!expected(run, expect, patterns, f)) {
                    if (ordered) std::printf("  run %d: file %zu not as expected\n", r, expect);
                    ordered = false;
                }
                // Booked as one step: a worker checking meanwhile sees the file either still counted or taken
                {
                    std::lock_guard<std::mutex> lock(run.booking);
                    const int counted = run.counted[expect].load();
                    if (counted == 1) run.outstanding -= run.specs[expect].size;
                    if (counted == 2) run.nextBytes -= run.specs[expect].size;
                    run.counted[expect].store(0);
                    run.taken = expect + 1;
                }
                ++expect;
                pool.recycle(std::move(f.text));
                if (expect == cancelAt) { pool.cancel(); ++cancels; }
            }
            const bool complete = (cancelAt >= 1 && cancelAt <= count) ? expect == cancelAt : expect == count;
            if (!complete) std::printf("  run %d: %zu of %zu files, cancel at %zu\n", r, expect, count, cancelAt);
            completed = completed && complete;
            files += expect;
        }
        budgetKept = budgetKept && run.gateViolations == 0;
        noBigBuffer = noBigBuffer && run.bigBuffers == 0;
        reused += static_cast<size_t>(run.reusedBuffers.load());
    }
    std::printf("%d runs: %zu files taken, %zu over the budget, %zu skipped, %zu threw, %zu canceled runs, %zu Pending answers\n",
        RUNS, files, large, skips, thrown, cancels, pendings);
    CHECK("L1 every file once and in list order, with its skip, kind, text and prefilter answers", ordered && completed);
    CHECK("L2 a file over the budget only once it is the next one, the budget kept", budgetKept && large > 0);
    CHECK("L3 Pending while the next file still loads", pendings > 0);
    CHECK("L4 cancel at random points ends next() and the workers", cancels > 0 && completed);
    CHECK("L5 exceptions in a load arrive as skips", thrown > 0 && ordered);
    CHECK("L6 buffers reused, none over 1 MB", reused > 0 && noBigBuffer);

    {
        Run run;
        const std::vector<FileLoadPool::Pattern> none;
        run.patterns = &none;
        g_run = &run;
        FileLoadPool empty({}, FileLoad::Options{}, {}, 2);
        FileLoadPool::File f;
        const bool done = empty.next(f) == FileLoadPool::Next::Done;
        FileLoadPool idle({ "0" }, FileLoad::Options{}, {}, 0);
        CHECK("L1 no files: Done at once; no workers asked for: none running", done && empty.workers() == 2 && idle.workers() == 0);

        run.specs.assign(200, Spec{ 100000, FileLoad::SkipReason::None, false, 0, 0, 0 });
        run.budget = 50000;
        run.counted.reset(new std::atomic<int>[201]());
        std::vector<std::filesystem::path> paths;
        for (size_t i = 0; i < 200; ++i) paths.emplace_back(std::to_string(i));
        const auto t0 = std::chrono::steady_clock::now();
        { FileLoadPool untouched(paths, FileLoad::Options{}, {}, 4, FileLoadPool::Limits{ 8, run.budget }); }
        const double ms = std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - t0).count();
        CHECK("L4 destroyed without taking a file: the workers stop within a block (" + std::to_string(static_cast<int>(ms)) + " ms)", ms < 1000);
    }

    g_drive = FileLoad::DriveKind::Fast;
    const bool given = FileLoadPool::workersFor(0, ".") == 0 && FileLoadPool::workersFor(2, ".") == 2
        && FileLoadPool::workersFor(8, ".") == 8 && FileLoadPool::workersFor(9, ".") == 8;
    const bool fast = FileLoadPool::workersFor(-1, ".") == 4;
    g_drive = FileLoad::DriveKind::Network;
    const bool network = FileLoadPool::workersFor(-1, ".") == 8 && FileLoadPool::workersFor(1, ".") == 1;
    g_drive = FileLoad::DriveKind::Slow;
    const bool slow = FileLoadPool::workersFor(-1, ".") == 1 && FileLoadPool::workersFor(3, ".") == 3;
    CHECK("L7 workers: the setting as given (at most 8), else 8 on a network drive, 4 on a fast one, 1 on a slow one",
        given && fast && network && slow);

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}

#else

// ---------------------------------------------------------------------------------------------
// Windows: the pool against loading on the UI thread
// ---------------------------------------------------------------------------------------------
#include <windows.h>
#include <chrono>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <random>
#include <string>
#include <thread>
#include <vector>

#include "PluginInterface.h"
#include "HiddenSciGuard.h"
#include "FileLoadPool.h"
#include "LiteralPrefilter.h"
#include "DirectoryWalk.h"

NppData nppData{};

namespace fs = std::filesystem;

namespace {

    int failures = 0;
    void CHECK(const std::string& what, bool ok)
    {
        std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str());
        if (!ok) ++failures;
    }

    std::string utf8(const std::wstring& w)
    {
        if (w.empty()) return {};
        const int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
        std::string s(static_cast<size_t>(n > 0 ? n : 0), '\0');
        if (n > 0) WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), s.data(), n, nullptr, nullptr);
        return s;
    }

    std::string bytesIn(const std::wstring& w, UINT codepage)
    {
        if (w.empty()) return {};
        const int n = WideCharToMultiByte(codepage, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
        std::string s(static_cast<size_t>(n > 0 ? n : 0), '\0');
        if (n > 0) WideCharToMultiByte(codepage, 0, w.data(), static_cast<int>(w.size()), s.data(), n, nullptr, nullptr);
        return s;
    }

    // The find texts, as the panel hands them to the pool: bytes in codepage 0 and in UTF-8
    std::vector<FileLoadPool::Pattern> patternsFor(bool literalOnly)
    {
        struct Find { const wchar_t* text; int flags; };
        const Find finds[] = {
            { L"Fi", SCFIND_MATCHCASE },
            { L"second", 0 },
            { L"gr\u00FCn", SCFIND_MATCHCASE },
            { L"F.i", SCFIND_MATCHCASE | SCFIND_REGEXP },
        };
        std::vector<FileLoadPool::Pattern> patterns;
        for (const Find& f : finds) {
            if (literalOnly && (f.flags & SCFIND_REGEXP)) continue;
            patterns.push_back({ bytesIn(f.text, 0), bytesIn(f.text, CP_UTF8), f.flags });
        }
        return patterns;
    }

    // What one file came to
    struct Outcome {
        FileLoad::SkipReason skip = FileLoad::SkipReason::None;
        FileLoad::LoadKind kind = FileLoad::LoadKind::Text;
        Encoding::EncodingInfo enc{};
        std::vector<char> mayMatch;
        bool handOver = false;
        std::string text;
        FileLoad::Stamp stamp{};
    };

    bool same(const Outcome& a, const Outcome& b)
    {
        if (a.skip != b.skip || a.handOver != b.handOver || a.text != b.text) return false;
        if (a.skip != FileLoad::SkipReason::None) return true;
        return a.kind == b.kind && a.enc.kind == b.enc.kind && a.enc.codepage == b.enc.codepage
            && a.enc.withBOM == b.enc.withBOM && a.enc.bomBytes == b.enc.bomBytes && a.mayMatch == b.mayMatch;
    }

    struct Mode {
        const char* name;
        bool replace;          // Replace in Files: roundtrip check, read-only skipped, stamped
        bool skipBinary;
        bool withPatterns;
        FileLoadPool::Limits limits;
    };

    // handleFindInFiles / handleReplaceInFiles on the UI thread: read-only check (Replace), load, prefilter
    Outcome serial(HiddenSciGuard& guard, const fs::path& p, const Mode& mode, const std::vector<FileLoadPool::Pattern>& patterns)
    {
        Outcome o;
        if (mode.replace) {
            const DWORD attrs = GetFileAttributesW(p.c_str());
            if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) {
                o.skip = FileLoad::SkipReason::ReadOnly;
                return o;
            }
        }
        o.skip = guard.loadTextFile(p, o.text, o.enc, o.kind);
        if (o.skip != FileLoad::SkipReason::None) {
            o.text.clear();
            return o;
        }
        const bool utf8Doc = FileLoad::codepageFor(o.kind, o.enc) == SC_CP_UTF8;
        LiteralPrefilter prefilter(o.text);
        o.handOver = patterns.empty();
        for (const auto& pattern : patterns) {
            const bool m = prefilter.mayMatch(utf8Doc ? pattern.utf8 : pattern.ansi, pattern.searchFlags);
            o.mayMatch.push_back(m ? 1 : 0);
            o.handOver = o.handOver || m;
        }
        if (!o.handOver) o.text.clear();
        return o;
    }

    // Generated corpus: every encoding and skip the loader knows, a file over a small budget, many small ones
    fs::path makeCorpus()
    {
        const fs::path dir = fs::temp_directory_path() / ("mr_file_load_pool_qa_" + std::to_string(GetCurrentProcessId()));
        std::error_code ec;
        fs::remove_all(dir, ec);
        fs::create_directories(dir / "many");
        auto write = [](const fs::path& p, const std::string& bytes) {
            std::ofstream(p, std::ios::binary).write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
        };
        auto utf16 = [](const std::wstring& w, bool bigEndian, bool bom) {
            std::string out;
            if (bom) out += bigEndian ? std::string("\xFE\xFF", 2) : std::string("\xFF\xFE", 2);
            for (wchar_t c : w) {
                const char lo = static_cast<char>(c & 0xFF), hi = static_cast<char>((c >> 8) & 0xFF);
                out += bigEndian ? hi : lo;
                out += bigEndian ? lo : hi;
            }
            return out;
        };
        const std::wstring textW = L"First Fight gr\u00FCn\r\nsecond line Fi\n";
        const std::wstring asciiW = L"First Fight\r\nsecond line Fi\r\nthird Fi line\r\n";
        write(dir / "u8.txt", utf8(textW));
        write(dir / "u8bom.txt", "\xEF\xBB\xBF" + utf8(textW));
        write(dir / "le.txt", utf16(textW, false, true));
        write(dir / "be.txt", utf16(textW, true, true));
        write(dir / "lenobom.txt", utf16(asciiW, false, false));
        write(dir / "odd.txt", utf16(textW, false, true) + "X");
        write(dir / "ansi.txt", bytesIn(textW, CP_ACP));
        write(dir / "empty.txt", "");
        write(dir / "ctl.bin", std::string("ctl\x00" "Fi \xC3\xA4\n", 10));
        std::string bin = std::string("MZ\x90\x00\x03\x00", 6);
        for (int i = 0; i < 20000; ++i) bin += static_cast<char>((i * 7919) & 0xFF);
        write(dir / "bin.exe", bin);
        write(dir / "readonly.txt", utf8(textW));
        SetFileAttributesW((dir / "readonly.txt").c_str(), FILE_ATTRIBUTE_READONLY);
        std::string big(3 << 20, 'x');
        for (size_t i = 0; i < big.size(); i += 97) big[i] = '\n';
        big += "last line Fi\n";
        write(dir / "big.txt", big);

        std::mt19937 rng(20261010);
        for (int i = 0; i < 400; ++i) {
            std::string body;
            const int lines = static_cast<int>(rng() % 200);
            for (int l = 0; l < lines; ++l) body += (rng() % 7 == 0) ? "a line with Fi in it\n" : "plain words on a line\n";
            std::string bytes;
            switch (rng() % 6) {
            case 0: bytes = "\xEF\xBB\xBF" + body; break;
            case 1: bytes = utf16(std::wstring(body.begin(), body.end()), false, true); break;
            case 2: bytes = body + bytesIn(L"gr\u00FCn\n", CP_ACP); break;
            case 3: bytes = body + std::string("\0\x01\x02", 3); break;
            default: bytes = body + utf8(L"gr\u00FCn\n"); break;
            }
            write(dir / "many" / ("f" + std::to_string(i) + ".txt"), bytes);
        }
        return dir;
    }

    std::vector<fs::path> listFiles(const fs::path& root)
    {
        DirectoryWalk::Options options;
        options.root = { true, true };
        DirectoryWalk::Result walk = DirectoryWalk::collect(root, options);
        return walk.status == DirectoryWalk::Status::Done ? std::move(walk.files) : std::vector<fs::path>{};
    }

    using Clock = std::chrono::steady_clock;
    double msSince(Clock::time_point t) { return std::chrono::duration<double, std::milli>(Clock::now() - t).count(); }

    // The first read of a folder, often from disk: not part of the comparison
    void readBeforehand(const fs::path& folder, const std::vector<fs::path>& files)
    {
        FileLoad::Options options;
        options.skipBinary = false;
        std::string content;
        const auto t0 = Clock::now();
        for (const auto& p : files) FileLoad::load(p, options, content);
        std::printf("  %s: %zu files read once beforehand in %.0f ms\n", utf8(folder.filename().wstring()).c_str(), files.size(), msSince(t0));
    }

    // Every file of the folder both ways; false on the first difference
    bool compare(const fs::path& folder, const std::vector<fs::path>& files, const Mode& mode)
    {
        const std::vector<FileLoadPool::Pattern> patterns = mode.withPatterns ? patternsFor(mode.replace) : std::vector<FileLoadPool::Pattern>{};
        HiddenSciGuard guard;
        guard.setSkipBinaryEnabled(mode.skipBinary);
        guard.setVerifyRoundtrip(mode.replace);

        auto t0 = Clock::now();
        std::vector<Outcome> reference;
        for (const auto& p : files) reference.push_back(serial(guard, p, mode, patterns));
        const double serialMs = msSince(t0);

        FileLoad::Options options = guard.loadOptions();
        options.skipReadOnly = mode.replace;
        options.stamp = mode.replace;
        bool equal = true;
        std::string times;
        for (unsigned workers : { 1u, 2u, 4u, 8u }) {
            std::vector<Outcome> pooled;
            pooled.reserve(files.size());
            t0 = Clock::now();
            {
                FileLoadPool pool(files, options, patterns, workers, mode.limits);
                FileLoadPool::File f;
                for (;;) {
                    const FileLoadPool::Next got = pool.next(f);
                    if (got == FileLoadPool::Next::Pending) continue;
                    if (got == FileLoadPool::Next::Done) break;
                    Outcome o;
                    o.skip = f.result.skip;
                    o.kind = f.result.kind;
                    o.enc = f.result.enc;
                    o.mayMatch = std::move(f.mayMatch);
                    o.handOver = f.handOver;
                    o.text = std::move(f.text);
                    o.stamp = f.result.stamp;
                    pooled.push_back(std::move(o));
                }
            }
            char line[64];
            std::snprintf(line, sizeof line, ", %u: %.0f ms", workers, msSince(t0));
            times += line;

            bool sameHere = pooled.size() == files.size();
            for (size_t i = 0; sameHere && i < files.size(); ++i) {
                const Outcome& o = pooled[i];
                const bool stampOk = !mode.replace || o.skip != FileLoad::SkipReason::None
                    || (o.stamp.valid && FileLoad::unchangedSince(files[i], o.stamp));
                if (!same(reference[i], o) || !stampOk) {
                    std::printf("  %s, %u workers: %s differs (skip %d/%d, hand over %d/%d, %zu/%zu bytes, stamp %d)\n",
                        mode.name, workers, utf8(files[i].wstring()).c_str(), static_cast<int>(reference[i].skip),
                        static_cast<int>(o.skip), reference[i].handOver, o.handOver, reference[i].text.size(), o.text.size(), stampOk);
                    sameHere = false;
                }
            }
            equal = equal && sameHere;
        }
        std::printf("  %s, %s: %zu files, serial %.0f ms%s\n", utf8(folder.filename().wstring()).c_str(), mode.name, files.size(), serialMs, times.c_str());
        return equal;
    }
}

int wmain(int argc, wchar_t** argv)
{
    std::vector<fs::path> dirs;
    for (int i = 1; i < argc; ++i) {
        const std::wstring a = argv[i];
        if (a == L"--dir" && i + 1 < argc) dirs.emplace_back(argv[++i]);
        else { std::fprintf(stderr, "usage: file_load_pool_qa [--dir PATH]...\n"); return 2; }
    }

    const fs::path corpus = makeCorpus();
    std::printf("corpus %s, ANSI codepage %u, %u hardware threads\n", utf8(corpus.wstring()).c_str(), GetACP(), std::thread::hardware_concurrency());

    const Mode modes[] = {
        { "Find", false, true, true, {} },
        { "Find, binary as raw bytes", false, false, true, {} },
        { "Find, no prefilter", false, true, false, {} },
        { "Replace", true, true, true, {} },
        { "Find, 4 ahead, 1 MB budget", false, true, true, { 4, uint64_t{ 1 } << 20 } },
        { "Replace, 4 ahead, 1 MB budget", true, true, true, { 4, uint64_t{ 1 } << 20 } },
    };

    bool corpusEqual = true;
    const std::vector<fs::path> corpusFiles = listFiles(corpus);
    readBeforehand(corpus, corpusFiles);
    for (const Mode& mode : modes) corpusEqual = compare(corpus, corpusFiles, mode) && corpusEqual;
    CHECK("W1 generated corpus (" + std::to_string(corpusFiles.size()) + " files): the pool loads every file as the UI thread does, 1, 2, 4 and 8 workers", corpusEqual);

    for (const fs::path& dir : dirs) {
        const std::vector<fs::path> files = listFiles(dir);
        readBeforehand(dir, files);
        bool equal = !files.empty();
        for (const Mode& mode : modes) equal = compare(dir, files, mode) && equal;
        CHECK("W2 " + utf8(dir.wstring()) + " (" + std::to_string(files.size()) + " files): the same", equal);
    }

    SetFileAttributesW((corpus / "readonly.txt").c_str(), FILE_ATTRIBUTE_NORMAL);
    std::error_code ec;
    fs::remove_all(corpus, ec);
    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}

#endif
