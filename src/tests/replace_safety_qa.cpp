// QA harness for the third guy038 follow-up round (issue #133):
//
//   1. Replace in Files skips files that are open in N++ with unsaved
//      changes (fork prevention) and reports them in the summary.
//   2. "Clear all marks" also deletes bookmarks when "Bookmark matched
//      lines" is checked - N++ parity (FindReplaceDlg::clearMarks).
//   3. View-constant correction in replaceAllInOpenedDocs: NPPM_ACTIVATEDOC
//      and NPPM_GETCURRENTDOCINDEX take MAIN_VIEW(0)/SUB_VIEW(1);
//      PRIMARY_VIEW(1)/SECOND_VIEW(2) belong to NPPM_GETNBOPENFILES only.
//
// Mirrors the panel's pathKey/dirty-set logic, summary composition and
// LM.get substitution verbatim; N++'s ACTIVATEDOC/GETCURRENTDOCINDEX
// dispatch is transcribed from NppBigSwitch.cpp (2026-08-31) as decision
// tables, so a constant regression fails a test instead of shipping.
//
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined
//        -o replace_safety_qa replace_safety_qa.cpp
#include <cstdio>
#include <cwctype>
#include <string>
#include <unordered_set>
#include <vector>

// ------------------------------------------------ view constants (verbatim)
static constexpr int MAIN_VIEW = 0;
static constexpr int SUB_VIEW = 1;
static constexpr int PRIMARY_VIEW = 1;   // NPPM_GETNBOPENFILES only
static constexpr int SECOND_VIEW = 2;    // NPPM_GETNBOPENFILES only

// N++ NppBigSwitch.cpp, case NPPM_ACTIVATEDOC (transcribed):
//   whichView = (wParam != MAIN_VIEW && wParam != SUB_VIEW) ? currentView() : wParam
// and switchEditViewTo refuses invisible views (returns currentView()).
static int nppActivateDocResolvesTo(int wParam, int currentView, bool subVisible)
{
    int whichView = (wParam != MAIN_VIEW && wParam != SUB_VIEW) ? currentView : wParam;
    if (whichView == SUB_VIEW && !subVisible) return currentView; // switchEditViewTo refusal
    return whichView;
}

// N++ NppBigSwitch.cpp, case NPPM_GETCURRENTDOCINDEX (transcribed):
//   lParam == SUB_VIEW -> sub index; EVERYTHING else -> main index.
static int nppDocIndexQueryReads(int lParam) { return (lParam == SUB_VIEW) ? SUB_VIEW : MAIN_VIEW; }

// ------------------------- MIRROR: pathKey + dirty-set skip (Replace in Files)
static std::wstring pathKey(std::wstring s) {
    for (auto& c : s) c = towlower(c);
    return s;
}

struct SkipResult { size_t openUnsavedSkipped = 0; std::vector<std::wstring> processed; };

static SkipResult simulateReplaceLoop(const std::vector<std::wstring>& scanFiles,
                                      const std::vector<std::wstring>& dirtyOpen)
{
    std::unordered_set<std::wstring> dirtyOpenPaths;
    for (const auto& p : dirtyOpen) dirtyOpenPaths.insert(pathKey(p));

    SkipResult r;
    for (const auto& fp : scanFiles) {
        if (!dirtyOpenPaths.empty() && dirtyOpenPaths.count(pathKey(fp)) != 0) {
            ++r.openUnsavedSkipped;
            continue;
        }
        r.processed.push_back(fp);
    }
    return r;
}

// ----------------- MIRROR: LM.get substitution + summary composition
static std::wstring lmGet(const std::wstring& tpl, const std::vector<std::wstring>& repl) {
    std::wstring result = tpl;
    const std::wstring base = L"$REPLACE_STRING";
    for (size_t i = repl.size(); i > 0; --i) {
        const std::wstring ph = base + std::to_wstring(i);
        for (size_t p = result.find(ph); p != std::wstring::npos; p = result.find(ph, p)) {
            result.replace(p, ph.size(), repl[i - 1]); p += repl[i - 1].size();
        }
    }
    if (!repl.empty())
        for (size_t p = result.find(base); p != std::wstring::npos; p = result.find(base, p)) {
            result.replace(p, base.size(), repl[0]); p += repl[0].size();
        }
    return result;
}

static const std::wstring kReadonlyEN   = L" $REPLACE_STRING read-only file(s) skipped.";
static const std::wstring kOpenUnsavedEN = L" $REPLACE_STRING file(s) skipped: open in Notepad++ with unsaved changes.";

// searched-files denominator, mirrored from the status tail
static size_t searchedFiles(size_t reached, size_t readOnly, size_t openUnsaved, size_t guardSkips) {
    const size_t notSearched = readOnly + openUnsaved + guardSkips;
    return (reached > notSearched) ? (reached - notSearched) : 0;
}

// --------- MIRROR: Clear-all-marks bookmark decision (N++ clearMarks parity)
static bool clearAlsoDeletesBookmarks(bool bookmarkCheckboxChecked) {
    return bookmarkCheckboxChecked;   // keyed on state at clear time, like N++
}

// ----------------------------------------------------------- checks
static int failures = 0;
static void CHECK(const std::string& d, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", d.c_str());
    if (!ok) ++failures;
}

int main() {
    std::printf("=== S1 dirty-open files are skipped, case-insensitively ===\n\n");
    {
        auto r = simulateReplaceLoop(
            { L"C:\\proj\\a.txt", L"C:\\proj\\B.TXT", L"C:\\proj\\c.txt" },
            { L"c:\\PROJ\\b.txt" });
        CHECK("S1 one file skipped despite case differences",
              r.openUnsavedSkipped == 1 && r.processed.size() == 2);
        CHECK("S1 the right files survived",
              r.processed[0] == L"C:\\proj\\a.txt" && r.processed[1] == L"C:\\proj\\c.txt");
    }
    {
        auto r = simulateReplaceLoop({ L"C:\\x\\a.txt" }, {});
        CHECK("S1 no dirty docs: nothing skipped, zero-cost path",
              r.openUnsavedSkipped == 0 && r.processed.size() == 1);
    }
    {
        // unsaved "new 1" style docs have no on-disk path and never match
        auto r = simulateReplaceLoop({ L"C:\\x\\a.txt" }, { L"new 1" });
        CHECK("S1 untitled docs never collide with scanned paths",
              r.openUnsavedSkipped == 0);
    }

    std::printf("\n=== S2 summary sentence and denominator ===\n\n");
    {
        const std::wstring s = lmGet(kOpenUnsavedEN, { L"2" });
        CHECK("S2 sentence renders",
              s == L" 2 file(s) skipped: open in Notepad++ with unsaved changes.");
        const std::wstring both = lmGet(kReadonlyEN, { L"1" }) + lmGet(kOpenUnsavedEN, { L"2" });
        CHECK("S2 appends after the read-only sentence, same pattern",
              both == L" 1 read-only file(s) skipped."
                     L" 2 file(s) skipped: open in Notepad++ with unsaved changes.");
        CHECK("S2 denominator subtracts ALL panel-side skips",
              searchedFiles(/*reached*/10, /*ro*/1, /*openUnsaved*/2, /*guard*/3) == 4);
        CHECK("S2 denominator clamps at zero",
              searchedFiles(2, 1, 2, 3) == 0);
    }

    std::printf("\n=== S3 Clear all marks vs bookmarks (N++ clearMarks parity) ===\n\n");
    {
        CHECK("S3 checkbox checked at clear time -> bookmarks deleted too",
              clearAlsoDeletesBookmarks(true));
        CHECK("S3 checkbox unchecked -> bookmarks stay",
              !clearAlsoDeletesBookmarks(false));
    }

    std::printf("\n=== S4 view-constant semantics (transcribed from N++ sources) ===\n\n");
    {
        // The dormant bug: PRIMARY_VIEW(1) sent to ACTIVATEDOC is read as SUB_VIEW.
        CHECK("S4 PRIMARY_VIEW(1) == SUB_VIEW(1): the aliasing that caused the bug",
              PRIMARY_VIEW == SUB_VIEW);
        CHECK("S4 old code, two views visible: 'main' loop landed in the SUB view",
              nppActivateDocResolvesTo(PRIMARY_VIEW, MAIN_VIEW, true) == SUB_VIEW);
        CHECK("S4 old code, single view: worked only via the invisible-view refusal",
              nppActivateDocResolvesTo(PRIMARY_VIEW, MAIN_VIEW, false) == MAIN_VIEW);
        CHECK("S4 old code, SECOND_VIEW(2): fell back to whatever view was current",
              nppActivateDocResolvesTo(SECOND_VIEW, SUB_VIEW, true) == SUB_VIEW
              && nppActivateDocResolvesTo(SECOND_VIEW, MAIN_VIEW, true) == MAIN_VIEW);
        CHECK("S4 fixed code: MAIN_VIEW/SUB_VIEW resolve to themselves",
              nppActivateDocResolvesTo(MAIN_VIEW, SUB_VIEW, true) == MAIN_VIEW
              && nppActivateDocResolvesTo(SUB_VIEW, MAIN_VIEW, true) == SUB_VIEW);
        CHECK("S4 GETCURRENTDOCINDEX: SECOND_VIEW(2) read the MAIN index (old bug)",
              nppDocIndexQueryReads(SECOND_VIEW) == MAIN_VIEW);
        CHECK("S4 GETCURRENTDOCINDEX: SUB_VIEW(1) reads the sub index (fixed)",
              nppDocIndexQueryReads(SUB_VIEW) == SUB_VIEW);
    }

    std::printf("\n%s (%d failure%s)\n",
                failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
