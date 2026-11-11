// Start positions per run scope and editor selection, against Notepad++'s FindReplaceDlg
// (cr = SCI_GETSELECTIONSTART/END): processAll begins at cr.cpMin, processFindNext at cr.cpMax.
// Find Previous starts at cr.cpMin, the mirror of Find Next, not at Notepad++'s cr.cpMax - 1 (what it
// then finds: find_prev_engine_qa). VERBATIM decision cores of computeAllStartPos (non-selection branch),
// the RunScope gating in handleReplaceAllButton (one start per run, passed to every list entry),
// handleFindNext/PrevButton and the selection handleReplaceButton works from.
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -o f5_start_pos_qa f5_start_pos_qa.cpp
#include <algorithm>
#include <cstdio>
#include <utility>
#include <vector>
using Pos = long long;
enum class RunScope { Editor, Docs, Files };

struct Sel { Pos anchor, caret; };                                         // anchor == caret: no selection
Pos selectionStart(const Sel& s) { return std::min(s.anchor, s.caret); }  // SCI_GETSELECTIONSTART
Pos selectionEnd(const Sel& s) { return std::max(s.anchor, s.caret); }    // SCI_GETSELECTIONEND

Pos computeAllStartPos_core(bool wrapEnabled, bool fromCursorEnabled, Pos selStart) {
    return wrapEnabled ? 0 : (fromCursorEnabled ? selStart : 0);
}
bool fromCursorFor(RunScope scope, bool optionOn) { return scope == RunScope::Editor && optionOn; }
bool selectionFor(RunScope scope, bool selectionRadio) { return scope != RunScope::Files && selectionRadio; }
Pos effectiveStart(RunScope scope, bool optionOn, bool wrap, const Sel& s) {
    return computeAllStartPos_core(wrap, fromCursorFor(scope, optionOn), selectionStart(s));
}
Pos findNextStart(const Sel& s) { return selectionEnd(s); }
Pos findPrevStart(const Sel& s) { return selectionStart(s); }
Pos findAllInDocsStart(bool selMode, Pos selStart) { return selMode ? selStart : 0; }            // handleFindAllInDocsButton

Pos nppAllStart(bool wrap, const Sel& s) { return wrap ? 0 : selectionStart(s); }
Pos nppFindNextStart(const Sel& s) { return selectionEnd(s); }
Pos nppFindPrevStart(const Sel& s) { return selectionEnd(s) - 1; }   // differs on purpose

using Range = std::pair<Pos, Pos>;
struct Multi { std::vector<Range> ranges; bool block; };                     // creation order: the last is the main one
Range currentSelection(const Multi& m) {                                      // SCI_GETSELECTIONSTART/END
    if (!m.block) return m.ranges.back();
    Range r = m.ranges.front();
    for (const Range& x : m.ranges) r = { std::min(r.first, x.first), std::max(r.second, x.second) };
    return r;
}
std::vector<Range> selectedRanges(const Multi& m) {                           // short block rows and added carets are empty
    std::vector<Range> v;
    for (const Range& r : m.ranges) if (r.second > r.first) v.push_back(r);
    return v;
}
Range firstRange(const Multi& m) { const auto v = selectedRanges(m); return *std::min_element(v.begin(), v.end()); }   // getSelectionInfo(false)
Range lastRange(const Multi& m) { const auto v = selectedRanges(m); return *std::max_element(v.begin(), v.end()); }    // getSelectionInfo(true)
Range replaceSelection(bool selectionScope, const Multi& m) { return selectionScope ? firstRange(m) : currentSelection(m); }
Pos findNextScopeStart(bool followUp, Range sel) { return followUp ? sel.second : sel.first; }   // Selection scope
Pos findPrevScopeStart(bool followUp, Range sel) { return followUp ? sel.first : sel.second; }

int fails = 0;
#define CHECK(d,c) do{bool ok=(c);printf("%-4s %s\n",ok?"PASS":"FAIL",d);if(!ok)++fails;}while(0)
int main() {
    const Sel NONE{ 18, 18 }, LTR{ 10, 14 }, RTL{ 14, 10 };   // left to right: caret at the end
    CHECK("editor, option ON,  wrap off, no selection -> caret",  effectiveStart(RunScope::Editor, true,  false, NONE) == 18);
    CHECK("editor, option ON,  wrap off, LTR -> selection start (was the caret)", effectiveStart(RunScope::Editor, true, false, LTR) == 10);
    CHECK("editor, option ON,  wrap off, RTL -> selection start", effectiveStart(RunScope::Editor, true,  false, RTL) == 10);
    CHECK("editor, option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Editor, true,  true,  LTR) == 0);
    CHECK("editor, option OFF, wrap off -> 0",                    effectiveStart(RunScope::Editor, false, false, LTR) == 0);
    CHECK("docs,   option ON,  wrap off -> 0",                    effectiveStart(RunScope::Docs,   true,  false, LTR) == 0);
    CHECK("docs,   option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Docs,   true,  true,  LTR) == 0);
    CHECK("docs,   option OFF, wrap off -> 0",                    effectiveStart(RunScope::Docs,   false, false, LTR) == 0);
    CHECK("files,  option ON,  wrap off -> 0 (F5)",               effectiveStart(RunScope::Files,  true,  false, NONE) == 0);
    CHECK("files,  option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Files,  true,  true,  NONE) == 0);
    CHECK("files,  option OFF, wrap off -> 0",                    effectiveStart(RunScope::Files,  false, false, NONE) == 0);
    CHECK("docs: Replace All starts where Find All in Docs does", effectiveStart(RunScope::Docs, true, false, LTR) == findAllInDocsStart(false, 10));
    CHECK("editor keeps the selection scope",                     selectionFor(RunScope::Editor, true));
    CHECK("docs keep each document's selection scope",            selectionFor(RunScope::Docs,   true));
    CHECK("files never use a selection scope",                    !selectionFor(RunScope::Files, true));
    CHECK("find next, RTL -> behind the selection (was its start)", findNextStart(RTL) == 14);
    CHECK("find prev, no selection -> the cursor (N++: one before)", findPrevStart(NONE) == 18 && nppFindPrevStart(NONE) == 17);
    CHECK("find prev, LTR -> before the selection (N++: in it)",  findPrevStart(LTR) == 10 && nppFindPrevStart(LTR) == 13);
    CHECK("find prev, RTL -> before the selection, as LTR",       findPrevStart(RTL) == 10);
    CHECK("find prev at document start stays 0",                  findPrevStart(Sel{ 0, 0 }) == 0);
    CHECK("selection scope, follow-up: prev before, next behind the last find",
          findPrevScopeStart(true, Range(10, 14)) == 10 && findNextScopeStart(true, Range(10, 14)) == 14);
    CHECK("selection scope, new selection: prev from its end, next from its start",
          findPrevScopeStart(false, Range(10, 14)) == 14 && findNextScopeStart(false, Range(10, 14)) == 10);

    const Multi SEVERAL{ { { 4, 7 }, { 20, 23 } }, false }, BLOCK{ { { 2, 6 }, { 17, 21 }, { 32, 36 } }, true };
    CHECK("replace, all text, several -> the main one (was the first)", replaceSelection(false, SEVERAL) == Range(20, 23));
    CHECK("replace, all text, block -> the whole block (was row 1)", replaceSelection(false, BLOCK) == Range(2, 36));
    CHECK("replace, selection scope, several -> first range (kept)", replaceSelection(true, SEVERAL) == Range(4, 7));
    CHECK("replace, selection scope, block -> first row (kept)",   replaceSelection(true, BLOCK) == Range(2, 6));

    const Multi SHORT_ROWS{ { { 2, 2 }, { 17, 21 }, { 32, 36 }, { 40, 40 } }, true }, PLUS_CARET{ { { 4, 7 }, { 20, 23 }, { 1, 1 } }, false };
    CHECK("selection scope, block with short first and last rows: next from row 2, prev from the end of row 3",
          findNextScopeStart(false, firstRange(SHORT_ROWS)) == 17 && findPrevScopeStart(false, lastRange(SHORT_ROWS)) == 36);
    CHECK("replace, selection scope, block with a short first row -> first selected row", replaceSelection(true, SHORT_ROWS) == Range(17, 21));
    CHECK("selection scope, several plus an added caret: the caret does not count",
          firstRange(PLUS_CARET) == Range(4, 7) && lastRange(PLUS_CARET) == Range(20, 23));

    int mismatches = 0;
    for (Pos a = 0; a <= 12; ++a)
        for (Pos c = 0; c <= 12; ++c) {
            const Sel s{ a, c };
            for (bool wrap : { false, true })
                mismatches += effectiveStart(RunScope::Editor, true, wrap, s) != nppAllStart(wrap, s);
            mismatches += findNextStart(s) != nppFindNextStart(s);
            mismatches += findPrevStart(s) != selectionStart(s) || findPrevStart(s) > findNextStart(s);
        }
    CHECK("all anchor/caret pairs: run starts and Find Next as in N++, Find Previous before the selection", mismatches == 0);
    // belt and suspenders: with SCI_GOTOPOS 0 after load, even an ungated read is safe
    CHECK("defined caret 0 after load: ungated worst case still 0", computeAllStartPos_core(false, true, 0) == 0);
    printf("\n%s (%d)\n", fails==0?"ALL CHECKS PASSED":"CHECKS FAILED", fails);
    return fails==0?0:1;
}
