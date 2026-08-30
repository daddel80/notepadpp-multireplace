// QA harness for the file-filter syntax.
//
// Covers the fix for the separator bug: the Files path used to normalise
// ';' -> ' ' and then tokenise on whitespace, which made any pattern that
// contains a space ("my report.txt") impossible to express - while the
// open-documents path split on ';' correctly. Both now share
// StringUtils::splitFilterPatterns.
//
// Mirrors verbatim:
//   StringUtils::splitFilterPatterns        (StringUtils.cpp)
//   HiddenSciGuard::parseFilter + matchPath (HiddenSciGuard.h, file-level part)
//   MultiReplace::matchesDocFilter          (MultiReplacePanel.cpp)
//
// PathMatchSpecW is replaced by a plain '*'/'?' glob. Test data avoids the
// one case where the two are known to differ (MS-DOS '*.*' handling of
// extension-less names), so no assertion here depends on that quirk.
//
// Build: g++ -std=c++20 -Wall -Wextra -o filter_syntax_qa filter_syntax_qa.cpp
#include <cstdio>
#include <string>
#include <vector>
#include <cwctype>

// ------------------------------------------------------- glob stand-in
static bool globMatch(const std::wstring& s, const std::wstring& p,
                      size_t si = 0, size_t pi = 0)
{
    while (pi < p.size()) {
        if (p[pi] == L'*') {
            for (size_t k = si; k <= s.size(); ++k)
                if (globMatch(s, p, k, pi + 1)) return true;
            return false;
        }
        if (si >= s.size()) return false;
        if (p[pi] != L'?' && towlower(p[pi]) != towlower(s[si])) return false;
        ++si; ++pi;
    }
    return si == s.size();
}
#define PathMatchSpecW(name, pat) globMatch(name, pat)

// ------------------------------- VERBATIM: StringUtils::trim / split
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

// The pre-fix Files behaviour, kept so the tests can show the difference.
static std::vector<std::wstring> oldSplitOnWhitespace(const std::wstring& filter) {
    std::wstring f = filter;
    for (auto& c : f) if (c == L';') c = L' ';
    std::vector<std::wstring> out;
    size_t i = 0;
    while (i < f.size()) {
        while (i < f.size() && (f[i] == L' ' || f[i] == L'\t')) ++i;
        size_t j = i;
        while (j < f.size() && f[j] != L' ' && f[j] != L'\t') ++j;
        if (j > i) out.push_back(f.substr(i, j - i));
        i = j;
    }
    return out;
}

// --------------------------- VERBATIM: HiddenSciGuard::parseFilter
struct Guard {
    std::vector<std::wstring> include_patterns, exclude_patterns,
                              exclude_folders, exclude_folders_recursive;

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
            (!exclude_patterns.empty() || !exclude_folders.empty()
             || !exclude_folders_recursive.empty()))
        {
            include_patterns.push_back(L"*.*");
        }
    }

    // file-level part of matchPath (folder rules need a real path, tested separately)
    bool matchFile(const std::wstring& fname) const {
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fname, pat)) return false;
        if (include_patterns.empty()) return true;
        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fname, pat)) return true;
        return false;
    }
};

// ----------------------- VERBATIM: MultiReplace::matchesDocFilter
static bool matchesDocFilter(const std::wstring& fileName, const std::wstring& filter) {
    if (filter.empty() || filter == L"*.*" || filter == L"*") return true;

    bool hasPositivePattern = false, matchedPositive = false;
    for (const std::wstring& pattern : splitFilterPatterns(filter)) {
        if (pattern[0] == L'!') {
            std::wstring excl = pattern.substr(1);
            if (!excl.empty() && PathMatchSpecW(fileName, excl)) return false;
        }
        else {
            hasPositivePattern = true;
            if (PathMatchSpecW(fileName, pattern)) matchedPositive = true;
        }
    }
    return hasPositivePattern ? matchedPositive : true;
}

// ----------------------------------------------------------- checks
static int failures = 0;
static void CHECK(const std::string& d, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", d.c_str());
    if (!ok) ++failures;
}
static std::string join(const std::vector<std::wstring>& v) {
    std::string s = "[";
    for (size_t i = 0; i < v.size(); ++i) {
        if (i) s += ", ";
        s += "'"; for (wchar_t c : v[i]) s += static_cast<char>(c); s += "'";
    }
    return s + "]";
}

int main() {
    std::printf("=== the reported bug: a pattern containing a space ===\n\n");
    {
        const std::wstring f = L"my report.txt";
        const auto now = splitFilterPatterns(f);
        const auto before = oldSplitOnWhitespace(f);
        std::printf("  filter  : my report.txt\n  before  : %s\n  now     : %s\n\n",
                    join(before).c_str(), join(now).c_str());
        CHECK("B1 'my report.txt' is one pattern, not two",
              now.size() == 1 && now[0] == L"my report.txt");
        CHECK("B1 the old behaviour really did split it (regression guard)",
              before.size() == 2);

        Guard g; g.parseFilter(f);
        CHECK("B2 the file itself matches", g.matchFile(L"my report.txt"));
        CHECK("B2 a file named just 'report.txt' does NOT match",
              !g.matchFile(L"report.txt"));
    }
    {
        const std::wstring f = L"*.*; !my report.txt";
        Guard g; g.parseFilter(f);
        CHECK("B3 exclusion with a space excludes exactly that file",
              !g.matchFile(L"my report.txt"));
        CHECK("B3 and does not swallow every 'report.txt'",
              g.matchFile(L"report.txt"));
    }

    std::printf("\n=== ordinary filters keep working ===\n\n");
    {
        Guard g; g.parseFilter(L"*.cpp; *.h; *.txt");
        CHECK("O1 three include patterns", g.include_patterns.size() == 3);
        CHECK("O1 *.cpp matches", g.matchFile(L"a.cpp"));
        CHECK("O1 *.bak does not", !g.matchFile(L"a.bak"));
    }
    {
        CHECK("O2 spaces around separators are ignored",
              splitFilterPatterns(L"  *.cpp ;  *.h  ;*.txt ").size() == 3);
        CHECK("O3 empty patterns are dropped",
              splitFilterPatterns(L"*.*;;;!*.md").size() == 2);
        CHECK("O4 a lone separator yields nothing",
              splitFilterPatterns(L" ; ; ").empty());
        CHECK("O5 empty filter yields nothing",
              splitFilterPatterns(L"").empty());
    }
    {
        Guard g; g.parseFilter(L"*.*; !\\tests\\; !+\\logs\\; !*.bak");
        CHECK("O6 non-recursive folder exclude classified",
              g.exclude_folders.size() == 1 && g.exclude_folders[0] == L"tests\\");
        CHECK("O6 recursive folder exclude classified",
              g.exclude_folders_recursive.size() == 1
              && g.exclude_folders_recursive[0] == L"\\logs\\");
        CHECK("O6 file exclude classified",
              g.exclude_patterns.size() == 1 && g.exclude_patterns[0] == L"*.bak");
    }
    {
        Guard g; g.parseFilter(L"!*.bak");
        CHECK("O7 exclusion-only filter gets the implicit *.* base",
              g.include_patterns.size() == 1 && g.include_patterns[0] == L"*.*");
    }

    std::printf("\n=== the two filter paths must agree ===\n\n");
    {
        // file-level filters only; folder rules exist in Files mode alone
        const std::wstring filters[] = {
            L"*.cpp; *.h", L"!*.bak", L"*.*; !*.exe; !*.obj",
            L"my report.txt", L"*.*; !my report.txt", L"*.log",
        };
        const std::wstring names[] = {
            L"a.cpp", L"a.h", L"a.bak", L"a.exe", L"my report.txt",
            L"report.txt", L"a.log", L"notes.txt",
        };
        int compared = 0, disagreed = 0;
        for (const auto& f : filters) {
            Guard g; g.parseFilter(f);
            for (const auto& n : names) {
                ++compared;
                if (g.matchFile(n) != matchesDocFilter(n, f)) {
                    ++disagreed;
                    std::printf("     divergence: filter='%ls' file='%ls' files=%d docs=%d\n",
                                f.c_str(), n.c_str(), g.matchFile(n), matchesDocFilter(n, f));
                }
            }
        }
        std::printf("  %d combinations compared\n\n", compared);
        CHECK("A1 Files and Docs filter agree on every combination", disagreed == 0);
    }

    std::printf("\n%s (%d failure%s)\n",
                failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
