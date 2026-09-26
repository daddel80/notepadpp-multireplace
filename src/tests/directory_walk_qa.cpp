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

// QA harness for the Find/Replace in Files folder walk (DirectoryWalk) and the folder
// rules of collectScanFiles.
// Portable, on a model file system (DirectoryWalk::walk with a model lister):
//   M. random trees against an independent recursive walk of the model: same files,
//      order, folders judged, entry count and unreadable count, under random folder
//      verdicts and cancel points, with folders that cannot be opened and listings
//      that fail midway (which a real file system does not produce on demand)
//   H. random trees and filters: folder rules decided once per folder (excluded
//      folders are not listed) keep exactly the files matchPath kept, in order
// Windows only, on real folders (DirectoryWalk::collect) next to the walk it replaced,
// recursive_directory_iterator with a hidden check per folder (reference()):
//   A. random trees: same files, folder order and entry count, with and without the
//      hidden-folder rule and without subfolders
//   B. a folder that disappears between listing and opening (the full-drive abort
//      guy038 reported): counted, the walk goes on
//   C. a folder without list permission: counted instead of dropped silently
//   D. root errors are reported, never thrown
//   E. cancel stops at the requested entry
//   F. links: a folder link or a link loop is not entered, a file link is kept,
//      a broken link is ignored
//   G. deep nesting
// F needs symlink rights (developer mode) and compares with the MSVC std::filesystem,
// C an ACL that is enforced; both report SKIP where not available (Wine enforces neither).
// PathMatchSpecW is replaced by a plain '*'/'?' glob outside Windows.
//
// Build (from src/tests):
//   g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. -o directory_walk_qa directory_walk_qa.cpp
//   cl /std:c++20 /EHsc /I.. directory_walk_qa.cpp ..\DirectoryWalk.cpp shlwapi.lib advapi32.lib /Fe:directory_walk_qa.exe
// Usage: directory_walk_qa [seed] [trees]
#include "../DirectoryWalk.h"

#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <cwctype>
#include <fstream>
#include <functional>
#include <map>
#include <random>
#include <string>
#include <string_view>
#include <system_error>
#include <vector>

#ifdef _WIN32
#include <windows.h>
#include <shlwapi.h>
#include <aclapi.h>
#else
static bool globMatch(const wchar_t* s, const wchar_t* p)
{
    if (*p == L'\0') return *s == L'\0';
    if (*p == L'*') return globMatch(s, p + 1) || (*s != L'\0' && globMatch(s + 1, p));
    if (*s == L'\0') return false;
    if (*p != L'?' && std::towlower(*p) != std::towlower(*s)) return false;
    return globMatch(s + 1, p + 1);
}
#define PathMatchSpecW(name, pat) globMatch(name, pat)
#endif

namespace fs = std::filesystem;
using DirectoryWalk::Take;

static int failures = 0;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }

// ================================================================ model file system
struct Node {
    enum Kind { File, Folder, LinkToFile, LinkToFolder, BrokenLink };
    std::wstring name;
    Kind kind = File;
    bool hidden = false;
    bool unreadable = false;   // folder: cannot be opened
    int failAfter = -1;        // folder: the listing breaks off after this many entries
    std::vector<Node> children;
};

class ModelLister {
public:
    class Listing {
    public:
        bool hasEntry() const { return _items && _index < _items->size() && !_lost; }
        DirectoryWalk::Entry entry() const
        {
            const Node& n = (*_items)[_index];
            return { n.name.c_str(), n.kind == Node::Folder, n.kind >= Node::LinkToFile, n.hidden };
        }
        bool next()
        {
            ++_index;
            if (_failAfter >= 0 && _index >= static_cast<size_t>(_failAfter) && _index < _items->size()) {
                _lost = true;
                return false;
            }
            return true;
        }
    private:
        friend class ModelLister;
        const std::vector<Node>* _items = nullptr;
        size_t _index = 0;
        int _failAfter = -1;
        bool _lost = false;
    };

    explicit ModelLister(const Node& root, const fs::path& rootPath) { add(root, rootPath); }

    // A listing that breaks off before its first entry cannot be listed at all (as on Windows)
    bool open(const fs::path& folder, Listing& listing, std::error_code& error)
    {
        ++opened;
        const auto it = _nodes.find(folder.wstring());
        if (it == _nodes.end() || it->second->kind != Node::Folder || it->second->unreadable || it->second->failAfter == 0) {
            error = std::make_error_code(std::errc::permission_denied);
            return false;
        }
        listing._items = &it->second->children;
        listing._failAfter = it->second->failAfter;
        return true;
    }

    bool linksToFile(const fs::path& link) const
    {
        const auto it = _nodes.find(link.wstring());
        return it != _nodes.end() && it->second->kind == Node::LinkToFile;
    }

    size_t opened = 0;

private:
    void add(const Node& n, const fs::path& p)
    {
        _nodes[p.wstring()] = &n;
        for (const Node& c : n.children) add(c, p / c.name);
    }
    std::map<std::wstring, const Node*> _nodes;
};

static Node growModel(std::mt19937& rng, int depth, bool faults, const std::vector<std::wstring>& names)
{
    Node folder;
    folder.kind = Node::Folder;
    const int count = static_cast<int>(rng() % 7);
    for (int i = 0; i < count; ++i) {
        Node c;
        const unsigned r = rng() % 100;
        if (depth > 0 && r < 30) c = growModel(rng, depth - 1, faults, names);
        else if (r < 34) c.kind = Node::LinkToFile;
        else if (r < 38) c.kind = Node::LinkToFolder;
        else if (r < 41) c.kind = Node::BrokenLink;
        c.name = names[rng() % names.size()] + std::to_wstring(i);   // unique within the folder
        c.hidden = rng() % 6 == 0;
        folder.children.push_back(std::move(c));
    }
    if (faults) {
        if (rng() % 12 == 0) folder.unreadable = true;
        else if (!folder.children.empty() && rng() % 8 == 0) folder.failAfter = static_cast<int>(rng() % folder.children.size());
    }
    return folder;
}

// ---------------------------------------------------------------- M: reference walk
struct Policy {
    Take root;
    std::function<Take(const fs::path&, bool)> verdict;
    std::function<bool(const wchar_t*)> keep;
    size_t cancelAt = 0;
};

struct Trace {
    std::vector<fs::path> files, judged;
    size_t seen = 0, unreadable = 0, opened = 1;   // folders listed or tried, the root included
    bool canceled = false;
};

// Independent of DirectoryWalk: recursion in listing order
static bool referenceWalk(const Node& folder, const fs::path& path, Take take, const Policy& p, Trace& t)
{
    const size_t n = folder.children.size();
    const size_t delivered = (folder.failAfter >= 0 && static_cast<size_t>(folder.failAfter) < n) ? static_cast<size_t>(folder.failAfter) : n;
    for (size_t i = 0; i < delivered; ++i) {
        const Node& c = folder.children[i];
        if (++t.seen == p.cancelAt) { t.canceled = true; return false; }
        const fs::path child = path / c.name;
        if (c.kind == Node::Folder) {
            if (!take.subfolders) continue;
            t.judged.push_back(child);
            const Take sub = p.verdict(child, c.hidden);
            if (!sub.files && !sub.subfolders) continue;
            ++t.opened;
            if (c.unreadable || c.failAfter == 0) { ++t.unreadable; continue; }
            if (!referenceWalk(c, child, sub, p, t)) return false;
        }
        else if (take.files && p.keep(c.name.c_str()) && (c.kind == Node::File || c.kind == Node::LinkToFile)) {
            t.files.push_back(child);
        }
    }
    if (delivered < n) ++t.unreadable;
    return true;
}

static Trace modelWalk(const Node& root, const fs::path& rootPath, const Policy& p, DirectoryWalk::Result& r)
{
    Trace t;
    ModelLister lister(root, rootPath);
    DirectoryWalk::Options o;
    o.root = p.root;
    o.subfolder = [&](const fs::path& f, bool hidden) { t.judged.push_back(f); return p.verdict(f, hidden); };
    o.keepFile = p.keep;
    o.keepGoing = [&](size_t n) { t.seen = n; return n != p.cancelAt; };
    r = DirectoryWalk::walk(lister, rootPath, o);
    t.files = r.files;
    t.unreadable = r.unreadableFolders;
    t.opened = lister.opened;
    t.canceled = r.status == DirectoryWalk::Status::Canceled;
    return t;
}

// ---------------------------------------------------------------- H: filter rules
// VERBATIM: StringUtils::trim, StringUtils::splitFilterPatterns
static std::wstring trim(const std::wstring& str) {
    const auto first = str.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos) return L"";
    const auto last = str.find_last_not_of(L" \t\r\n");
    return str.substr(first, last - first + 1);
}

static std::vector<std::wstring> splitFilterPatterns(const std::wstring& filter) {
    std::vector<std::wstring> out;
    size_t pos = 0;
    while (pos <= filter.size()) {
        const size_t sep = filter.find(L';', pos);
        const size_t end = (sep == std::wstring::npos) ? filter.size() : sep;
        std::wstring pattern = trim(filter.substr(pos, end - pos));
        if (!pattern.empty()) out.push_back(std::move(pattern));
        if (sep == std::wstring::npos) break;
        pos = sep + 1;
    }
    return out;
}

// VERBATIM: HiddenSciGuard::parseFilter and its pattern lists
struct Filter {
    std::vector<std::wstring> include_patterns, exclude_patterns, exclude_folders, exclude_folders_recursive;

    void parseFilter(const std::wstring& filterString) {
        include_patterns.clear(); exclude_patterns.clear();
        exclude_folders.clear(); exclude_folders_recursive.clear();
        for (const std::wstring& tok : splitFilterPatterns(filterString)) {
            if (tok.rfind(L"!+", 0) == 0) {
                exclude_folders_recursive.push_back(tok.substr(2));
            }
            else if (tok.rfind(L"!", 0) == 0) {
                if (tok.size() > 1 && tok[1] == L'\\')
                    exclude_folders.push_back(tok.substr(2));
                else
                    exclude_patterns.push_back(tok.substr(1));
            }
            else {
                include_patterns.push_back(tok);
            }
        }
        if (include_patterns.empty() &&
            (!exclude_patterns.empty() || !exclude_folders.empty() || !exclude_folders_recursive.empty()))
            include_patterns.push_back(L"*.*");
    }

    // VERBATIM (before): HiddenSciGuard::matchPath, per file found
    bool matchPath(const fs::path& path) const
    {
        const std::wstring fname = path.filename().wstring();
        const fs::path parentPath = path.parent_path();
        if (!parentPath.empty()) {
            const std::wstring parentName = parentPath.filename().wstring();
            for (const auto& pat : exclude_folders)
                if (PathMatchSpecW(parentName.c_str(), pat.c_str()))
                    return false;
        }
        for (auto dir = parentPath; !dir.empty() && dir != dir.root_path(); dir = dir.parent_path()) {
            const std::wstring dirName = dir.filename().wstring();
            for (const auto& rawPat : exclude_folders_recursive) {
                std::wstring_view pat = rawPat;
                if (!pat.empty() && (pat.front() == L'\\' || pat.front() == L'/'))
                    pat.remove_prefix(1);
                if (PathMatchSpecW(dirName.c_str(), std::wstring{ pat }.c_str()))
                    return false;
            }
        }
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return false;
        if (include_patterns.empty())
            return true;
        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return true;
        return false;
    }

    // VERBATIM (after): HiddenSciGuard::excludesFolder, excludesRoot, excludesFilesIn, matchFileName
    bool excludesFolder(const std::wstring& folderName) const
    {
        for (const auto& rawPat : exclude_folders_recursive) {
            std::wstring_view pat = rawPat;
            if (!pat.empty() && (pat.front() == L'\\' || pat.front() == L'/'))
                pat.remove_prefix(1);

            if (PathMatchSpecW(folderName.c_str(), std::wstring{ pat }.c_str()))
                return true;
        }
        return false;
    }

    bool excludesRoot(const fs::path& root) const
    {
        for (auto dir = root; !dir.empty() && dir != dir.root_path(); dir = dir.parent_path())
            if (excludesFolder(dir.filename().wstring()))
                return true;
        return false;
    }

    bool excludesFilesIn(const std::wstring& folderName) const
    {
        for (const auto& pat : exclude_folders)
            if (PathMatchSpecW(folderName.c_str(), pat.c_str()))
                return true;
        return false;
    }

    bool matchFileName(const wchar_t* fileName) const
    {
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fileName, pat.c_str()))
                return false;

        if (include_patterns.empty())
            return true;

        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fileName, pat.c_str()))
                return true;

        return false;
    }
};

// Before: every file of the old walk (hidden folders pruned), then matchPath
static std::vector<fs::path> scanBefore(const Node& model, const std::wstring& dir, bool recurse, bool includeHidden, const Filter& guard, size_t& seen)
{
    const fs::path root = fs::absolute(fs::path(dir)).lexically_normal();
    Policy p;
    p.root = { true, recurse };
    p.verdict = [&](const fs::path&, bool hidden) { return (hidden && !includeHidden) ? Take{ false, false } : Take{}; };
    p.keep = [](const wchar_t*) { return true; };
    Trace t;
    referenceWalk(model, root, p.root, p, t);
    seen = t.seen;
    std::vector<fs::path> out;
    for (const auto& f : t.files)
        if (guard.matchPath(f)) out.push_back(f);
    return out;
}

// After: VERBATIM options of MultiReplace::collectScanFiles
static std::vector<fs::path> scanAfter(const Node& model, const fs::path& modelRoot, const std::wstring& dir, bool recurse, bool includeHidden, const Filter& guard, size_t& seen)
{
    DirectoryWalk::Options options;
    options.subfolder = [&](const fs::path& folder, bool hidden) {
        const std::wstring name = folder.filename().wstring();
        if ((hidden && !includeHidden) || guard.excludesFolder(name))
            return DirectoryWalk::Take{ false, false };
        return DirectoryWalk::Take{ !guard.excludesFilesIn(name), true };
        };
    options.keepFile = [&](const wchar_t* name) { return guard.matchFileName(name); };
    options.keepGoing = [&](size_t n) { seen = n; return true; };

    fs::path root = fs::absolute(fs::path(dir)).lexically_normal();
    if (!root.has_filename() && root.has_relative_path())
        root = root.parent_path();
    if (guard.excludesRoot(root))
        options.root = { false, false };
    else
        options.root = { !guard.excludesFilesIn(root.filename().wstring()), recurse };

    seen = 0;
    ModelLister lister(model, modelRoot);
    return DirectoryWalk::walk(lister, root, options).files;
}

// ================================================================ real folders (Windows)
#ifdef _WIN32
static bool g_links = false;   // symlinks can be created here
static void touch(const fs::path& p) { std::ofstream(p) << "x"; }
static bool isHidden(const fs::path& p) { const DWORD a = GetFileAttributesW(p.c_str()); return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_HIDDEN); }
static void hide(const fs::path& p) { SetFileAttributesW(p.c_str(), GetFileAttributesW(p.c_str()) | FILE_ATTRIBUTE_HIDDEN); }
static std::vector<fs::path> sorted(std::vector<fs::path> v) { std::sort(v.begin(), v.end()); return v; }

static bool makeLink(const fs::path& target, const fs::path& link, bool folder)
{
    const DWORD flags = (folder ? SYMBOLIC_LINK_FLAG_DIRECTORY : 0) | SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE;
    if (!CreateSymbolicLinkW(link.c_str(), target.c_str(), flags)) return false;
    const DWORD a = GetFileAttributesW(link.c_str());   // Wine reports success without a link
    return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_REPARSE_POINT);
}

struct RealTrace {
    std::vector<fs::path> files, folders;   // kept files; real folders reached (candidates to enter)
    std::vector<size_t> fileAt;             // entry number of every kept file
    size_t seen = 0;
};

// The walk before DirectoryWalk: every entry of recursive_directory_iterator,
// a folder's hidden attribute read per folder
static RealTrace reference(const fs::path& root, bool recurse, bool skipHidden)
{
    RealTrace t;
    if (recurse) {
        auto it = fs::recursive_directory_iterator(root, fs::directory_options::skip_permission_denied);
        for (auto end = fs::recursive_directory_iterator(); it != end; ++it) {
            ++t.seen;
            if (it->symlink_status().type() == fs::file_type::directory) t.folders.push_back(it->path());
            if (it->is_directory() && skipHidden && isHidden(it->path())) {
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

static RealTrace walk(const fs::path& root, bool recurse, bool skipHidden, DirectoryWalk::Result* out = nullptr, size_t cancelAt = 0)
{
    RealTrace t;
    DirectoryWalk::Options o;
    o.root = { true, recurse };
    o.subfolder = [&](const fs::path& p, bool hidden) {
        t.folders.push_back(p);
        return (skipHidden && hidden) ? Take{ false, false } : Take{};
        };
    o.keepFile = [&](const wchar_t*) { t.fileAt.push_back(t.seen); return true; };
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
    if (rng() % 4 == 0) { touch(dir / "hf.txt"); hide(dir / "hf.txt"); }   // hidden files are always kept
    for (int i = 0; i < dirs; ++i) {
        const bool hidden = rng() % 3 == 0;
        const fs::path sub = dir / ((hidden ? "h" : "d") + std::to_string(i));
        fs::create_directory(sub);
        if (hidden) hide(sub);
        grow(rng, sub, depth - 1);
    }
    if (!g_links) return;
    if (rng() % 4 == 0) makeLink(dir, dir / "loop", true);
    if (rng() % 4 == 0) makeLink(dir / "gone.txt", dir / "broken", false);
    if (rng() % 3 == 0) { touch(dir / "target.txt"); makeLink(dir / "target.txt", dir / "link.txt", false); }
    if (depth > 0 && rng() % 3 == 0) {
        fs::create_directory(dir / "real");
        touch(dir / "real" / "r.txt");
        makeLink(dir / "real", dir / "dirlink", true);
    }
}

static bool same(const RealTrace& a, const RealTrace& b)
{
    return a.files == b.files && a.folders == b.folders && a.fileAt == b.fileAt && a.seen == b.seen;
}

static void realFolders(unsigned seed, int trees)
{
    std::mt19937 rng(seed);
    const fs::path base = fs::temp_directory_path() / ("mr_directory_walk_qa_" + std::to_string(seed) + "_" + std::to_string(rng() % 100000));
    std::error_code ec;
    fs::remove_all(base, ec);
    fs::create_directories(base);
    {
        touch(base / "probe_target.txt");
        g_links = makeLink(base / "probe_target.txt", base / "probe", false);
        fs::remove(base / "probe", ec);
        fs::remove(base / "probe_target.txt", ec);
    }
    std::printf("\nreal folders: links %s, base %s\n", g_links ? "on" : "off", base.string().c_str());

    std::printf("\n=== A) random trees: new walk == old walk ===\n");
    {
        int bad[3] = { 0, 0, 0 };
        size_t entries = 0, counted = 0, hiddenSeen = 0;
        for (int n = 0; n < trees; ++n) {
            const fs::path root = base / ("a" + std::to_string(n));
            fs::create_directory(root);
            grow(rng, root, 1 + static_cast<int>(rng() % 4));
            const bool modes[3][2] = { { true, false }, { true, true }, { false, false } };
            for (int m = 0; m < 3; ++m) {
                DirectoryWalk::Result r;
                const RealTrace ref = reference(root, modes[m][0], modes[m][1]);
                const RealTrace now = walk(root, modes[m][0], modes[m][1], &r);
                if (m == 0) {
                    entries += ref.seen;
                    for (const auto& f : ref.folders) hiddenSeen += isHidden(f);
                }
                counted += r.unreadableFolders;
                if (!same(ref, now) || r.status != DirectoryWalk::Status::Done) {
                    if (++bad[m] <= 3)
                        std::printf("  MISMATCH tree %d mode %d: files %zu/%zu folders %zu/%zu seen %zu/%zu\n", n, m,
                            ref.files.size(), now.files.size(), ref.folders.size(), now.folders.size(), ref.seen, now.seen);
                }
            }
        }
        std::printf("  %zu entries walked, %zu hidden folders\n", entries, hiddenSeen);
        CHECK("A1 all subfolders: files, folder order, entry numbers, mismatches: " + std::to_string(bad[0]), bad[0] == 0);
        CHECK("A2 hidden folders pruned from the listing's attribute: mismatches: " + std::to_string(bad[1]), bad[1] == 0 && hiddenSeen > 0);
        CHECK("A3 without subfolders: mismatches: " + std::to_string(bad[2]), bad[2] == 0);
        CHECK("A4 healthy trees report no unreadable folder", counted == 0);
    }

    std::printf("\n=== B) folder disappears between listing and opening ===\n");
    {
        const fs::path root = base / "b";
        fs::create_directories(root / "a");
        fs::create_directories(root / "b" / "sub");
        fs::create_directories(root / "c");
        touch(root / "a" / "1.txt");
        touch(root / "b" / "2.txt");
        touch(root / "b" / "sub" / "3.txt");
        touch(root / "c" / "4.txt");
        touch(root / "5.txt");

        DirectoryWalk::Options o;
        o.subfolder = [](const fs::path& p, bool) { if (p.filename() == "b") fs::remove_all(p); return Take{}; };
        DirectoryWalk::Result r;
        bool threw = false;
        try { r = DirectoryWalk::collect(root, o); }
        catch (...) { threw = true; }
        const std::vector<fs::path> want = { root / "5.txt", root / "a" / "1.txt", root / "c" / "4.txt" };
        CHECK("B1 the walk does not throw and finishes", !threw && r.status == DirectoryWalk::Status::Done);
        CHECK("B2 the vanished folder is counted once", r.unreadableFolders == 1);
        CHECK("B3 every file outside it is found", sorted(r.files) == sorted(want));
    }

    std::printf("\n=== C) folder without list permission ===\n");
    {
        const fs::path root = base / "c";
        fs::create_directories(root / "open");
        fs::create_directories(root / "locked" / "deeper");
        touch(root / "open" / "1.txt");
        touch(root / "locked" / "2.txt");
        touch(root / "locked" / "deeper" / "3.txt");
        touch(root / "4.txt");

        ACL empty;   // an empty DACL grants nothing; the owner can still restore it
        const bool denied = InitializeAcl(&empty, sizeof(empty), ACL_REVISION)
            && SetNamedSecurityInfoW(const_cast<wchar_t*>((root / "locked").c_str()), SE_FILE_OBJECT,
                DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION, nullptr, nullptr, &empty, nullptr) == ERROR_SUCCESS;
        WIN32_FIND_DATAW d;
        const HANDLE probe = FindFirstFileExW((root / "locked" / "*").c_str(), FindExInfoBasic, &d, FindExSearchNameMatch, nullptr, 0);
        if (probe != INVALID_HANDLE_VALUE) FindClose(probe);
        if (!denied || probe != INVALID_HANDLE_VALUE) {
            std::printf("SKIP C (the ACL is not enforced here)\n");
        }
        else {
            DirectoryWalk::Result r;
            const RealTrace now = walk(root, true, false, &r);
            const std::vector<fs::path> want = { root / "4.txt", root / "open" / "1.txt" };
            CHECK("C1 the files outside it are found", sorted(now.files) == sorted(want) && r.status == DirectoryWalk::Status::Done);
            CHECK("C2 and the folder it could not read is counted", r.unreadableFolders == 1);
        }
        SetNamedSecurityInfoW(const_cast<wchar_t*>((root / "locked").c_str()), SE_FILE_OBJECT,
            DACL_SECURITY_INFORMATION | UNPROTECTED_DACL_SECURITY_INFORMATION, nullptr, nullptr, nullptr, nullptr);
    }

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
        const RealTrace ref = reference(root, true, false);
        const size_t k = ref.seen / 2 + 1;
        DirectoryWalk::Result r;
        const RealTrace now = walk(root, true, false, &r, k);
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
        makeLink(root / "real", root / "dirlink", true);
        makeLink(root / "real" / "a.txt", root / "filelink.txt", false);
        makeLink(root / "nowhere", root / "broken", false);
        makeLink(root, root / "real" / "up", true);

        DirectoryWalk::Result r;
        const RealTrace now = walk(root, true, false, &r);
        const RealTrace ref = reference(root, true, false);
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
        const RealTrace now = walk(base / "g", true, false, &r);
        CHECK("G1 64 levels: the leaf file is found", now.files == std::vector<fs::path>{ deep / "leaf.txt" } && r.unreadableFolders == 0);
        CHECK("G2 same as the old walk", same(reference(base / "g", true, false), now));
    }

    fs::remove_all(base, ec);
}
#endif

int main(int argc, char** argv)
{
    const unsigned seed = argc > 1 ? static_cast<unsigned>(std::strtoul(argv[1], nullptr, 10)) : 20260926u;
    const int trees = argc > 2 ? std::atoi(argv[2]) : 200;
    std::printf("seed %u, %d trees\n", seed, trees);
#ifdef _WIN32
    const fs::path modelBase = L"C:\\model";
#else
    const fs::path modelBase = "/model";
#endif
    const std::vector<std::wstring> plain = { L"a", L"b", L"c", L"sub" };
    const std::vector<std::wstring> named = { L"d", L"sub", L"obj", L"logs", L"h", L"f", L"g.log" };

    std::printf("\n=== M) model trees: the walk == an independent recursive walk ===\n");
    {
        std::mt19937 rng(seed);
        int bad[4] = { 0, 0, 0, 0 };
        size_t entries = 0, unreadable = 0, lost = 0, canceled = 0, notOpened = 0;
        for (int n = 0; n < trees * 5; ++n) {
            Node root = growModel(rng, 1 + static_cast<int>(rng() % 4), true, plain);
            root.unreadable = false;
            root.failAfter = -1;
            const fs::path rootPath = modelBase / (L"m" + std::to_wstring(n));
            for (int mode = 0; mode < 4; ++mode) {
                const unsigned salt = static_cast<unsigned>(rng());
                Policy p;
                p.root = { mode != 1, mode != 2 };
                p.verdict = [salt](const fs::path& f, bool hidden) {
                    const size_t h = std::hash<std::wstring>{}(f.wstring()) ^ salt;
                    if (hidden && (h & 1)) return Take{ false, false };
                    return Take{ (h % 5) != 0, (h % 7) != 0 };
                    };
                p.keep = [salt](const wchar_t* name) { return ((std::hash<std::wstring>{}(name) ^ salt) % 4) != 0; };
                Trace ref;
                referenceWalk(root, rootPath, p.root, p, ref);
                if (mode == 3 && ref.seen > 0) {   // cancel somewhere in the middle
                    p.cancelAt = 1 + static_cast<size_t>(rng() % ref.seen);
                    ref = Trace{};
                    referenceWalk(root, rootPath, p.root, p, ref);
                }
                DirectoryWalk::Result r;
                const Trace now = modelWalk(root, rootPath, p, r);
                const bool ok = now.canceled == ref.canceled && now.seen == ref.seen && now.judged == ref.judged
                    && now.opened == ref.opened
                    && (now.canceled ? std::equal(now.files.begin(), now.files.end(), ref.files.begin(), ref.files.end())
                                     : (now.files == ref.files && now.unreadable == ref.unreadable));
                if (!ok && ++bad[mode] <= 3)
                    std::printf("  MISMATCH tree %d mode %d: files %zu/%zu judged %zu/%zu seen %zu/%zu unreadable %zu/%zu\n", n, mode,
                        ref.files.size(), now.files.size(), ref.judged.size(), now.judged.size(), ref.seen, now.seen, ref.unreadable, now.unreadable);
                entries += ref.seen;
                unreadable += ref.unreadable;
                canceled += ref.canceled;
                notOpened += ref.judged.size() + 1 - ref.opened;
            }
            std::function<void(const Node&)> countLost = [&](const Node& f) {
                if (f.failAfter > 0) ++lost;
                for (const Node& c : f.children) if (c.kind == Node::Folder) countLost(c);
            };
            countLost(root);
        }
        std::printf("  %zu entries, %zu unreadable folders, %zu listings breaking off, %zu canceled walks\n", entries, unreadable, lost, canceled);
        CHECK("M1 all of the root: files, order, folders judged and opened, unreadable count, mismatches: " + std::to_string(bad[0]), bad[0] == 0);
        CHECK("M2 root without its files: mismatches: " + std::to_string(bad[1]), bad[1] == 0);
        CHECK("M3 root without subfolders: mismatches: " + std::to_string(bad[2]), bad[2] == 0);
        CHECK("M4 canceled walks keep exactly the files before the cancel, mismatches: " + std::to_string(bad[3]), bad[3] == 0 && canceled > 0);
        CHECK("M5 faults were exercised (unreadable folders, listings breaking off)", unreadable > 0 && lost > 0);
        CHECK("M6 folders judged 'nothing' were among them, and never opened", notOpened > 0 && bad[0] + bad[1] + bad[2] + bad[3] == 0);
    }

    std::printf("\n=== H) filters: folder rules per folder == matchPath per file ===\n");
    {
        static const wchar_t* pool[] = {
            L"*.*", L"*.txt", L"f*", L"*.bin", L"a?.txt", L"*",
            L"!*.bin", L"!f1*", L"!*.log", L"!g*",
            L"!\\d0", L"!\\h*", L"!\\sub*", L"!\\d1\\", L"!\\root*", L"!\\*", L"!\\up",
            L"!+d1", L"!+\\h0", L"!+\\d0\\", L"!+*2", L"!+sub*", L"!+\\root*", L"!+obj*", L"!+up", L"!+\\zz",
        };
        const size_t poolSize = sizeof(pool) / sizeof(pool[0]);
        std::mt19937 rng(seed + 7);
        int compared = 0, bad = 0;
        size_t seenBefore = 0, seenAfter = 0, keptTotal = 0;
        for (int n = 0; n < trees; ++n) {
            // the root sits below folders the filter can name, so !+x above the root is covered
            const Node model = growModel(rng, 1 + static_cast<int>(rng() % 3), false, named);
            const fs::path root = modelBase / (n % 2 ? L"up" : L"zz") / (L"root" + std::to_wstring(n));
            for (int k = 0; k < 6; ++k) {
                std::wstring filter;
                const int parts = 1 + static_cast<int>(rng() % 4);
                for (int p = 0; p < parts; ++p) filter += std::wstring(pool[rng() % poolSize]) + L";";
                Filter guard;
                guard.parseFilter(filter);
                const std::wstring dir = root.wstring() + (k % 3 == 0 ? std::wstring(1, fs::path::preferred_separator) : L"");
                for (int mode = 0; mode < 4; ++mode) {
                    const bool recurse = (mode & 1) == 0, includeHidden = (mode & 2) != 0;
                    size_t sb = 0, sa = 0;
                    const auto before = scanBefore(model, dir, recurse, includeHidden, guard, sb);
                    const auto after = scanAfter(model, root, dir, recurse, includeHidden, guard, sa);
                    ++compared;
                    seenBefore += sb; seenAfter += sa; keptTotal += after.size();
                    if (before != after && ++bad <= 5)
                        std::printf("  MISMATCH root %d filter '%ls' recurse %d hidden %d: %zu vs %zu files\n",
                            n, filter.c_str(), recurse, includeHidden, before.size(), after.size());
                }
            }
        }
        std::printf("  %d scans compared, %zu files kept, entries listed %zu before / %zu after\n", compared, keptTotal, seenBefore, seenAfter);
        CHECK("H1 same files in the same order for every filter and mode, mismatches: " + std::to_string(bad), bad == 0 && keptTotal > 0);
        CHECK("H2 excluded folders are not listed any more", seenAfter < seenBefore);
    }

#ifdef _WIN32
    realFolders(seed, trees);
#else
    std::printf("\nSKIP A-G (real folders: Windows only)\n");
#endif

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
