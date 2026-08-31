// QA harness for the "Find/Replace in Docs" transparency headers.
//
// Covers the counters added for guy038's follow-up (issue #133): the dock
// header of Find in Docs now reports "[N document(s) searched]" plus an
// ", M excluded by filter" clause, and the Replace-in-Docs status line
// appends the same information when the doc filter excluded documents.
//
// Mirrors verbatim:
//   StringUtils::splitFilterPatterns          (StringUtils.cpp)
//   MultiReplace::matchesDocFilter            (MultiReplacePanel.cpp)
//   LanguageManager::get substitution         (LanguageManager.cpp:
//                                              <br/> -> CRLF, $REPLACE_STRINGn
//                                              high->low, then $REPLACE_STRING)
//   the counting skeleton of findAllInOpenedDocs / replaceAllInOpenedDocs
//     (filter gate -> ++docsFilteredOut/continue, else search -> ++docsSearched,
//      uniqueFiles/hit aggregation only for hits > 0)
//
// Template strings are copied verbatim from language_mapping.cpp (EN) and
// languages.ini [german].
//
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined
//        -o docs_header_qa docs_header_qa.cpp
#include <cstdio>
#include <random>
#include <set>
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

// -------------- VERBATIM (logic): LanguageManager::get substitution
static std::wstring lmGet(const std::wstring& tpl, const std::vector<std::wstring>& repl) {
    std::wstring result = tpl;
    const std::wstring base = L"$REPLACE_STRING";

    for (size_t p = result.find(L"<br/>");
        p != std::wstring::npos;
        p = result.find(L"<br/>", p))
    {
        result.replace(p, 5, L"\r\n");
        p += 2;
    }
    for (size_t i = repl.size(); i > 0; --i) {
        const std::wstring ph = base + std::to_wstring(i);
        const std::wstring& vv = repl[i - 1];
        for (size_t p = result.find(ph);
            p != std::wstring::npos;
            p = result.find(ph, p))
        {
            result.replace(p, ph.size(), vv);
            p += vv.size();
        }
    }
    if (!repl.empty()) {
        for (size_t p = result.find(base);
            p != std::wstring::npos;
            p = result.find(base, p))
        {
            result.replace(p, base.size(), repl[0]);
            p += repl[0].size();
        }
    }
    return result;
}

// ------- template strings, verbatim from language_mapping.cpp (EN) and
// ------- languages.ini [german]
static const std::wstring kDockDocsSuffixEN   = L" [$REPLACE_STRING1 document(s) searched$REPLACE_STRING2]";
static const std::wstring kDockDocsFilteredEN = L", $REPLACE_STRING1 excluded by filter";
static const std::wstring kStatusDocsFilteredEN = L" $REPLACE_STRING1 document(s) searched, $REPLACE_STRING2 excluded by filter.";
static const std::wstring kReplaceSummaryEN   = L"Replace in documents: $REPLACE_STRING occurrences replaced.";
static const std::wstring kDockDocsSuffixDE   = L" [$REPLACE_STRING1 Dokument(e) durchsucht$REPLACE_STRING2]";
static const std::wstring kDockDocsFilteredDE = L", $REPLACE_STRING1 durch Filter ausgeschlossen";

// -------------------- MIRROR: the counting skeleton of the two doc loops
struct Doc { std::wstring name; int hits; };

struct ScanResult {
    int totalHits = 0;
    size_t uniqueFiles = 0;
    size_t docsSearched = 0;
    size_t docsFilteredOut = 0;
    std::wstring suffix;   // dock header suffix as composed in findAllInOpenedDocs
};

static ScanResult simulateFindInDocs(const std::vector<Doc>& mainView,
                                     const std::vector<Doc>& subView,
                                     bool filterDocs, const std::wstring& docFilter)
{
    ScanResult r;
    std::set<std::wstring> unique;
    auto oneView = [&](const std::vector<Doc>& docs) {
        for (const Doc& d : docs) {
            if (filterDocs) {
                if (!matchesDocFilter(d.name, docFilter)) { ++r.docsFilteredOut; continue; }
            }
            // processCurrentBuffer: aggregate only when the doc has hits
            if (d.hits > 0) { unique.insert(d.name); r.totalHits += d.hits; }
            ++r.docsSearched;
        }
    };
    oneView(mainView);
    oneView(subView);
    r.uniqueFiles = unique.size();

    std::wstring filterClause;
    if (r.docsFilteredOut > 0)
        filterClause = lmGet(kDockDocsFilteredEN, { std::to_wstring(r.docsFilteredOut) });
    r.suffix = lmGet(kDockDocsSuffixEN, { std::to_wstring(r.docsSearched), filterClause });
    return r;
}

static std::wstring simulateReplaceSummary(int grandTotal, size_t searched, size_t filteredOut)
{
    std::wstring docsSummary = lmGet(kReplaceSummaryEN, { std::to_wstring(grandTotal) });
    if (filteredOut > 0)
        docsSummary += lmGet(kStatusDocsFilteredEN,
            { std::to_wstring(searched), std::to_wstring(filteredOut) });
    return docsSummary;
}

// ----------------------------------------------------------- checks
static int failures = 0;
static void CHECK(const std::string& d, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", d.c_str());
    if (!ok) ++failures;
}
static std::string n(const std::wstring& w) {           // narrow for printf
    std::string s; for (wchar_t c : w) s += (c < 128) ? static_cast<char>(c) : '?';
    return s;
}

int main() {
    std::printf("=== D1 guy038's E:\\Test case: 7 open docs, all with hits ===\n\n");
    {
        std::vector<Doc> m = { {L"Mark_Style.txt",2}, {L"license_890.txt",1}, {L"Xomx.txt",10},
                               {L"langs.xml",72}, {L"notepad++.exe",162},
                               {L"npp.8.9.portable.x64.7z",98}, {L"Marginalize.dll",27} };
        auto r = simulateFindInDocs(m, {}, false, L"");
        CHECK("D1 totalHits 372", r.totalHits == 372);
        CHECK("D1 unique == searched == 7 (every doc hit)",
              r.uniqueFiles == 7 && r.docsSearched == 7 && r.docsFilteredOut == 0);
        CHECK("D1 suffix ' [7 document(s) searched]'",
              r.suffix == L" [7 document(s) searched]");
        std::printf("     header: (%d hits in %zu file(s))%s\n\n",
                    r.totalHits, r.uniqueFiles, n(r.suffix).c_str());
    }

    std::printf("=== D2 divergence: docs without hits ===\n\n");
    {
        std::vector<Doc> m = { {L"a.txt",5}, {L"b.txt",0}, {L"c.txt",0}, {L"d.txt",7},
                               {L"e.txt",0}, {L"f.txt",0}, {L"g.txt",38} };
        std::vector<Doc> s = { {L"h.txt",0}, {L"i.txt",0}, {L"j.txt",0} };
        auto r = simulateFindInDocs(m, s, false, L"");
        CHECK("D2 3 files with hits out of 10 searched",
              r.uniqueFiles == 3 && r.docsSearched == 10);
        CHECK("D2 suffix ' [10 document(s) searched]'",
              r.suffix == L" [10 document(s) searched]");
        std::printf("     header: (%d hits in %zu file(s))%s\n\n",
                    r.totalHits, r.uniqueFiles, n(r.suffix).c_str());
    }

    std::printf("=== D3 doc filter active ===\n\n");
    {
        std::vector<Doc> m = { {L"a.txt",4}, {L"b.txt",0}, {L"c.log",9},
                               {L"d.exe",50}, {L"e.txt",2} };
        auto r = simulateFindInDocs(m, {}, true, L"*.txt");
        CHECK("D3 searched 3 (.txt), filtered 2 (.log/.exe)",
              r.docsSearched == 3 && r.docsFilteredOut == 2);
        CHECK("D3 hits only from searched docs (4+2, .exe's 50 excluded)",
              r.totalHits == 6 && r.uniqueFiles == 2);
        CHECK("D3 suffix ' [3 document(s) searched, 2 excluded by filter]'",
              r.suffix == L" [3 document(s) searched, 2 excluded by filter]");
    }
    {
        // negation filter: everything except logs
        std::vector<Doc> m = { {L"a.txt",1}, {L"b.log",1}, {L"c.log",1} };
        auto r = simulateFindInDocs(m, {}, true, L"*.*; !*.log");
        CHECK("D3 exclusion filter: 1 searched, 2 filtered",
              r.docsSearched == 1 && r.docsFilteredOut == 2 && r.totalHits == 1);
    }

    std::printf("\n=== D4 filter excludes everything ===\n\n");
    {
        std::vector<Doc> m = { {L"a.txt",3}, {L"b.txt",1}, {L"c.txt",0},
                               {L"d.txt",0}, {L"e.txt",9} };
        auto r = simulateFindInDocs(m, {}, true, L"*.md");
        CHECK("D4 0 searched, 5 filtered, 0 hits",
              r.docsSearched == 0 && r.docsFilteredOut == 5 && r.totalHits == 0);
        CHECK("D4 suffix ' [0 document(s) searched, 5 excluded by filter]'",
              r.suffix == L" [0 document(s) searched, 5 excluded by filter]");
    }

    std::printf("\n=== D5 clone: same file open in both views ===\n\n");
    {
        std::vector<Doc> m = { {L"clone.txt",3} };
        std::vector<Doc> s = { {L"clone.txt",3} };
        auto r = simulateFindInDocs(m, s, false, L"");
        CHECK("D5 searched 2 (per view), unique 1 (documented behavior)",
              r.docsSearched == 2 && r.uniqueFiles == 1);
    }

    std::printf("\n=== D6 invariants over randomized scenarios ===\n\n");
    {
        std::mt19937 rng(133);   // seed = the issue number
        std::uniform_int_distribution<int> nDocs(0, 12), hit(0, 5), pick(0, 3);
        const std::wstring exts[4] = { L".txt", L".log", L".xml", L".exe" };
        const std::wstring filters[4] = { L"*.txt", L"*.txt; *.xml", L"*.*; !*.exe", L"*.md" };
        int bad = 0;
        for (int t = 0; t < 2000; ++t) {
            std::vector<Doc> m, s;
            const int nm = nDocs(rng), ns = nDocs(rng) / 2;
            for (int i = 0; i < nm; ++i)
                m.push_back({ L"m" + std::to_wstring(i) + exts[pick(rng)], hit(rng) });
            for (int i = 0; i < ns; ++i)
                s.push_back({ L"s" + std::to_wstring(i) + exts[pick(rng)], hit(rng) });
            const bool useFilter = (t % 2) == 1;
            const std::wstring f = filters[pick(rng)];
            auto r = simulateFindInDocs(m, s, useFilter, f);

            const size_t open = m.size() + s.size();
            if (r.uniqueFiles > r.docsSearched) ++bad;                       // hits only from searched docs
            if (r.docsSearched + r.docsFilteredOut != open) ++bad;           // every open doc accounted for
            if (!useFilter && r.docsFilteredOut != 0) ++bad;                 // no filter, no exclusions
            if (r.totalHits == 0 && r.uniqueFiles != 0) ++bad;               // no hits, no hit-files
        }
        CHECK("D6 2000 scenarios: unique<=searched, searched+filtered==open, "
              "no phantom exclusions", bad == 0);
    }

    std::printf("\n=== D7 German templates ===\n\n");
    {
        const std::wstring clause = lmGet(kDockDocsFilteredDE, { L"4" });
        const std::wstring suffix = lmGet(kDockDocsSuffixDE, { L"3", clause });
        CHECK("D7 ' [3 Dokument(e) durchsucht, 4 durch Filter ausgeschlossen]'",
              suffix == L" [3 Dokument(e) durchsucht, 4 durch Filter ausgeschlossen]");
        CHECK("D7 no-filter DE suffix ' [7 Dokument(e) durchsucht]'",
              lmGet(kDockDocsSuffixDE, { L"7", L"" }) == L" [7 Dokument(e) durchsucht]");
    }

    std::printf("\n=== D8 Replace-in-Docs status line ===\n\n");
    {
        CHECK("D8 with filter clause",
              simulateReplaceSummary(12, 3, 2) ==
              L"Replace in documents: 12 occurrences replaced."
              L" 3 document(s) searched, 2 excluded by filter.");
        CHECK("D8 without exclusions: summary unchanged (bare $REPLACE_STRING form)",
              simulateReplaceSummary(7, 5, 0) ==
              L"Replace in documents: 7 occurrences replaced.");
    }

    std::printf("\n%s (%d failure%s)\n",
                failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
