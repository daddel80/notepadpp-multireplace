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

// QA harness for the Find/Replace in Files folder walk (DirectoryWalk). Runs the
// real unit on real folder trees next to the walk it replaces, collectScanFiles'
// recursive_directory_iterator loop (reference(), a name rule standing in for the
// hidden attribute):
//   A. random trees: same files, same folder order, same entry count, with and
//      without the hidden-folder rule and without subfolders
//   B. a folder that disappears between listing and opening (the full-drive abort
//      guy038 reported): the old walk throws and loses the scan, the new one
//      counts the folder and goes on
//   C. an unreadable folder (POSIX, not as root): counted instead of dropped silently
//   D. root errors are reported, never thrown
//   E. cancel stops at the requested entry
//   F. links: a folder link or a link loop is not entered, a file link is kept,
//      a broken link is ignored
//   G. deep nesting
// Links need symlink rights (Windows: developer mode), C needs POSIX permissions;
// both are reported as SKIP where not available. A listing that fails midway
// (readdir/FindNextFile error) needs fault injection and is not covered here.
//
// Build (from src/tests):
//   g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. -o directory_walk_qa directory_walk_qa.cpp ../DirectoryWalk.cpp
//   cl /std:c++20 /EHsc /I.. directory_walk_qa.cpp ..\DirectoryWalk.cpp /Fe:directory_walk_qa.exe
// Usage: directory_walk_qa [seed] [trees]
#include "../DirectoryWalk.h"

#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <random>
#include <string>
#include <system_error>
#include <vector>
#ifndef _WIN32
#include <unistd.h>
#endif

namespace fs = std::filesystem;

static int failures = 0;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }

static bool g_links = false;   // symlinks can be created here
static void touch(const fs::path& p) { std::ofstream(p) << "x"; }
static bool hiddenRule(const fs::path& p) { return p.filename().string().rfind('h', 0) == 0; }   // stands in for the hidden attribute
static std::vector<fs::path> sorted(std::vector<fs::path> v) { std::sort(v.begin(), v.end()); return v; }

struct Trace {
    std::vector<fs::path> files, folders;   // kept files; real folders reached (candidates to enter)
    std::vector<size_t> fileAt;             // entry number of every kept file
    size_t seen = 0;
};

// collectScanFiles before the change, recording what it saw
static Trace reference(const fs::path& root, bool recurse, bool skipHidden)
{
    Trace t;
    if (recurse) {
        auto it = fs::recursive_directory_iterator(root, fs::directory_options::skip_permission_denied);
        for (auto end = fs::recursive_directory_iterator(); it != end; ++it) {
            ++t.seen;
            if (it->symlink_status().type() == fs::file_type::directory) t.folders.push_back(it->path());
            if (it->is_directory() && skipHidden && hiddenRule(it->path())) {
                it.disable_recursion_pending();
                continue;
            }
            if (it->is_regular_file()) { t.files.push_back(it->path()); t.fileAt.push_back(t.seen); }
        }
    }
    else {
        for (auto& e : fs::directory_iterator(root, fs::directory_options::skip_permission_denied)) {
            ++t.seen;
            if (e.is_regular_file()) { t.files.push_back(e.path()); t.fileAt.push_back(t.seen); }
        }
    }
    return t;
}

static Trace walk(const fs::path& root, bool recurse, bool skipHidden, DirectoryWalk::Result* out = nullptr, size_t cancelAt = 0)
{
    Trace t;
    DirectoryWalk::Options o;
    o.recurse = recurse;
    o.enterFolder = [&](const fs::path& p) { t.folders.push_back(p); return !(skipHidden && hiddenRule(p)); };
    o.keepFile = [&](const fs::path&) { t.fileAt.push_back(t.seen); return true; };
    o.keepGoing = [&](size_t n) { t.seen = n; return n != cancelAt; };
    DirectoryWalk::Result r = DirectoryWalk::collect(root, o);
    t.files = r.files;
    if (out) *out = std::move(r);
    return t;
}

static void grow(std::mt19937& rng, const fs::path& dir, int depth)
{
    const int files = static_cast<int>(rng() % 5), dirs = depth > 0 ? static_cast<int>(rng() % 4) : 0;
    for (int i = 0; i < files; ++i) touch(dir / ("f" + std::to_string(i) + (rng() % 2 ? ".txt" : ".bin")));
    for (int i = 0; i < dirs; ++i) {
        const fs::path sub = dir / ((rng() % 3 == 0 ? "h" : "d") + std::to_string(i));
        fs::create_directory(sub);
        grow(rng, sub, depth - 1);
    }
    if (!g_links) return;
    std::error_code ec;
    if (rng() % 4 == 0) fs::create_directory_symlink(dir, dir / "loop", ec);
    if (rng() % 4 == 0) fs::create_symlink(dir / "gone.txt", dir / "broken", ec);
    if (rng() % 3 == 0) { touch(dir / "target.txt"); fs::create_symlink(dir / "target.txt", dir / "link.txt", ec); }
    if (depth > 0 && rng() % 3 == 0) {
        fs::create_directory(dir / "real");
        touch(dir / "real" / "r.txt");
        fs::create_directory_symlink(dir / "real", dir / "dirlink", ec);
    }
}

static bool same(const Trace& a, const Trace& b)
{
    return a.files == b.files && a.folders == b.folders && a.fileAt == b.fileAt && a.seen == b.seen;
}

int main(int argc, char** argv)
{
    const unsigned seed = argc > 1 ? static_cast<unsigned>(std::strtoul(argv[1], nullptr, 10)) : 20260924u;
    const int trees = argc > 2 ? std::atoi(argv[2]) : 200;
    std::mt19937 rng(seed);
    const fs::path base = fs::temp_directory_path() / ("mr_directory_walk_qa_" + std::to_string(seed) + "_" + std::to_string(rng() % 100000));
    fs::remove_all(base);
    fs::create_directories(base);
    {
        std::error_code ec;
        fs::create_directory_symlink(base, base / "probe", ec);
        g_links = !ec;
        fs::remove(base / "probe", ec);
    }
    std::printf("seed %u, %d trees, links %s, base %s\n", seed, trees, g_links ? "on" : "off", base.string().c_str());

    std::printf("\n=== A) random trees: new walk == old walk ===\n");
    {
        int bad[3] = { 0, 0, 0 };
        size_t entries = 0, counted = 0;
        for (int n = 0; n < trees; ++n) {
            const fs::path root = base / ("a" + std::to_string(n));
            fs::create_directory(root);
            grow(rng, root, 1 + static_cast<int>(rng() % 4));
            const bool modes[3][2] = { { true, false }, { true, true }, { false, false } };
            for (int m = 0; m < 3; ++m) {
                DirectoryWalk::Result r;
                const Trace ref = reference(root, modes[m][0], modes[m][1]);
                const Trace now = walk(root, modes[m][0], modes[m][1], &r);
                if (m == 0) entries += ref.seen;
                counted += r.unreadableFolders;
                if (!same(ref, now) || r.status != DirectoryWalk::Status::Done) {
                    if (++bad[m] <= 3)
                        std::printf("  MISMATCH tree %d mode %d: files %zu/%zu folders %zu/%zu seen %zu/%zu\n", n, m,
                            ref.files.size(), now.files.size(), ref.folders.size(), now.folders.size(), ref.seen, now.seen);
                }
            }
        }
        std::printf("  %zu entries walked\n", entries);
        CHECK("A1 all subfolders: files, folder order, entry numbers, mismatches: " + std::to_string(bad[0]), bad[0] == 0);
        CHECK("A2 hidden-folder rule: mismatches: " + std::to_string(bad[1]), bad[1] == 0);
        CHECK("A3 without subfolders: mismatches: " + std::to_string(bad[2]), bad[2] == 0);
        CHECK("A4 healthy trees report no unreadable folder", counted == 0);
    }

    std::printf("\n=== B) folder disappears between listing and opening ===\n");
    {
        auto build = [](const fs::path& root) {
            fs::create_directories(root / "a");
            fs::create_directories(root / "b" / "sub");
            fs::create_directories(root / "c");
            touch(root / "a" / "1.txt");
            touch(root / "b" / "2.txt");
            touch(root / "b" / "sub" / "3.txt");
            touch(root / "c" / "4.txt");
            touch(root / "5.txt");
        };
        const fs::path oldRoot = base / "b_old", newRoot = base / "b_new";
        build(oldRoot);
        build(newRoot);

        bool threw = false;
        try {
            auto it = fs::recursive_directory_iterator(oldRoot, fs::directory_options::skip_permission_denied);
            for (auto end = fs::recursive_directory_iterator(); it != end; ++it)
                if (it->path().filename() == "b") fs::remove_all(it->path());
        }
        catch (const fs::filesystem_error& e) {
            threw = true;
            std::printf("  old walk: %s\n", e.what());
        }
        CHECK("B1 old walk: recursive_directory_iterator throws, the whole scan is lost", threw);

        DirectoryWalk::Options o;
        o.enterFolder = [](const fs::path& p) { if (p.filename() == "b") fs::remove_all(p); return true; };
        DirectoryWalk::Result r;
        bool newThrew = false;
        try { r = DirectoryWalk::collect(newRoot, o); }
        catch (...) { newThrew = true; }
        const std::vector<fs::path> want = { newRoot / "5.txt", newRoot / "a" / "1.txt", newRoot / "c" / "4.txt" };
        CHECK("B2 new walk does not throw and finishes", !newThrew && r.status == DirectoryWalk::Status::Done);
        CHECK("B3 the vanished folder is counted once", r.unreadableFolders == 1);
        CHECK("B4 every file outside it is found", sorted(r.files) == sorted(want));
    }

    std::printf("\n=== C) unreadable folder ===\n");
#ifndef _WIN32
    if (::geteuid() == 0) {
        std::printf("SKIP C (running as root: every folder is readable)\n");
    }
    else {
        const fs::path root = base / "c";
        fs::create_directories(root / "open");
        fs::create_directories(root / "locked" / "deeper");
        touch(root / "open" / "1.txt");
        touch(root / "locked" / "2.txt");
        touch(root / "locked" / "deeper" / "3.txt");
        touch(root / "4.txt");
        fs::permissions(root / "locked", fs::perms::none);

        const Trace ref = reference(root, true, false);
        DirectoryWalk::Result r;
        const Trace now = walk(root, true, false, &r);
        const std::vector<fs::path> want = { root / "4.txt", root / "open" / "1.txt" };
        CHECK("C1 old walk skipped it without a trace", sorted(ref.files) == sorted(want));
        CHECK("C2 new walk finds the same files", sorted(now.files) == sorted(want) && r.status == DirectoryWalk::Status::Done);
        CHECK("C3 and counts the folder it could not read", r.unreadableFolders == 1);
        fs::permissions(root / "locked", fs::perms::owner_all);
    }
#else
    std::printf("SKIP C (POSIX permissions only)\n");
#endif

    std::printf("\n=== D) root errors ===\n");
    {
        DirectoryWalk::Result missing, file;
        bool threw = false;
        touch(base / "plain.txt");
        try {
            missing = DirectoryWalk::collect(base / "does_not_exist", {});
            file = DirectoryWalk::collect(base / "plain.txt", {});
        }
        catch (...) { threw = true; }
        CHECK("D1 no exception", !threw);
        CHECK("D2 missing root: RootUnreadable, no such file or directory",
            missing.status == DirectoryWalk::Status::RootUnreadable && missing.rootError == std::errc::no_such_file_or_directory);
        CHECK("D3 root is a file: RootUnreadable with an error", file.status == DirectoryWalk::Status::RootUnreadable && file.rootError);
        CHECK("D4 no files, no folder count", missing.files.empty() && file.files.empty() && missing.unreadableFolders == 0);
    }

    std::printf("\n=== E) cancel ===\n");
    {
        const fs::path root = base / "e";
        fs::create_directories(root / "sub");
        touch(root / "x.txt");
        std::mt19937 r2(seed + 1);
        grow(r2, root / "sub", 4);
        const Trace ref = reference(root, true, false);
        const size_t k = ref.seen / 2 + 1;
        DirectoryWalk::Result r;
        const Trace now = walk(root, true, false, &r, k);
        const size_t wantFiles = static_cast<size_t>(std::count_if(ref.fileAt.begin(), ref.fileAt.end(), [&](size_t at) { return at < k; }));
        std::printf("  %zu entries, cancel at %zu\n", ref.seen, k);
        CHECK("E1 status Canceled", r.status == DirectoryWalk::Status::Canceled);
        CHECK("E2 no entry after the canceling one", now.seen == k);
        CHECK("E3 files are exactly those before it", now.files.size() == wantFiles
            && std::equal(now.files.begin(), now.files.end(), ref.files.begin()));
    }

    std::printf("\n=== F) links ===\n");
    if (!g_links) {
        std::printf("SKIP F (no symlink rights)\n");
    }
    else {
        const fs::path root = base / "f";
        fs::create_directories(root / "real");
        touch(root / "real" / "a.txt");
        std::error_code ec;
        fs::create_directory_symlink(root / "real", root / "dirlink", ec);
        fs::create_symlink(root / "real" / "a.txt", root / "filelink.txt", ec);
        fs::create_symlink(root / "nowhere", root / "broken", ec);
        fs::create_directory_symlink(root, root / "real" / "up", ec);

        DirectoryWalk::Result r;
        const Trace now = walk(root, true, false, &r);
        const Trace ref = reference(root, true, false);
        const std::vector<fs::path> want = { root / "filelink.txt", root / "real" / "a.txt" };
        CHECK("F1 only the real folder is entered", now.folders == std::vector<fs::path>{ root / "real" });
        CHECK("F2 file link kept, nothing through the folder link or the loop", sorted(now.files) == sorted(want));
        CHECK("F3 broken link and loop are not errors", r.unreadableFolders == 0 && r.status == DirectoryWalk::Status::Done);
        CHECK("F4 same as the old walk", same(ref, now));
    }

    std::printf("\n=== G) deep nesting ===\n");
    {
        fs::path deep = base / "g";
        for (int i = 0; i < 64; ++i) deep /= "d";
        fs::create_directories(deep);
        touch(deep / "leaf.txt");
        DirectoryWalk::Result r;
        const Trace now = walk(base / "g", true, false, &r);
        CHECK("G1 64 levels: the leaf file is found", now.files == std::vector<fs::path>{ deep / "leaf.txt" } && r.unreadableFolders == 0);
        CHECK("G2 same as the old walk", same(reference(base / "g", true, false), now));
    }

    std::error_code ec;
    fs::remove_all(base, ec);
    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
