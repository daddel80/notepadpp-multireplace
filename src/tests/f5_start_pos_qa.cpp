// Start-position matrix for Replace All per run scope (F5 and the Docs cursor start).
// VERBATIM decision core of computeAllStartPos (non-selection branch) plus the
// RunScope gating in handleReplaceAllButton/replaceAll: only the editor starts at the
// cursor. Runs over open documents or files cover whole documents, as Notepad++ does
// (replaceInOpenedFiles, findInOpenedFiles and the Files runs pass isEntireDoc = true).
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -o f5_start_pos_qa f5_start_pos_qa.cpp
#include <cstdio>
using Pos = long long;
enum class RunScope { Editor, Docs, Files };

Pos computeAllStartPos_core(bool wrapEnabled, bool fromCursorEnabled, Pos caretPos) {
    return wrapEnabled ? 0 : (fromCursorEnabled ? caretPos : 0);
}
bool fromCursorFor(RunScope scope, bool optionOn) { return scope == RunScope::Editor && optionOn; }      // gating under test
bool selectionFor(RunScope scope, bool selectionRadio) { return scope != RunScope::Files && selectionRadio; }
Pos effectiveStart(RunScope scope, bool optionOn, bool wrap, Pos caret) {
    return computeAllStartPos_core(wrap, fromCursorFor(scope, optionOn), caret);
}
Pos findAllInDocsStart(bool selMode, Pos selStart) { return selMode ? selStart : 0; }   // handleFindAllInDocsButton

int fails = 0;
#define CHECK(d,c) do{bool ok=(c);printf("%-4s %s\n",ok?"PASS":"FAIL",d);if(!ok)++fails;}while(0)
int main() {
    const Pos CARET = 18;   // editor caret, or the caret after SCI_ADDTEXT (file end)
    CHECK("editor, option ON,  wrap off -> caret",                effectiveStart(RunScope::Editor, true,  false, CARET) == CARET);
    CHECK("editor, option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Editor, true,  true,  CARET) == 0);
    CHECK("editor, option OFF, wrap off -> 0",                    effectiveStart(RunScope::Editor, false, false, CARET) == 0);
    CHECK("docs,   option ON,  wrap off -> 0 (was the caret)",    effectiveStart(RunScope::Docs,   true,  false, CARET) == 0);
    CHECK("docs,   option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Docs,   true,  true,  CARET) == 0);
    CHECK("docs,   option OFF, wrap off -> 0",                    effectiveStart(RunScope::Docs,   false, false, CARET) == 0);
    CHECK("files,  option ON,  wrap off -> 0 (F5)",               effectiveStart(RunScope::Files,  true,  false, CARET) == 0);
    CHECK("files,  option ON,  wrap on  -> 0",                    effectiveStart(RunScope::Files,  true,  true,  CARET) == 0);
    CHECK("files,  option OFF, wrap off -> 0",                    effectiveStart(RunScope::Files,  false, false, CARET) == 0);
    CHECK("docs: Replace All starts where Find All in Docs does", effectiveStart(RunScope::Docs, true, false, CARET) == findAllInDocsStart(false, CARET));
    CHECK("editor keeps the selection scope",                     selectionFor(RunScope::Editor, true));
    CHECK("docs keep each document's selection scope",            selectionFor(RunScope::Docs,   true));
    CHECK("files never use a selection scope",                    !selectionFor(RunScope::Files, true));
    // belt and suspenders: with SCI_GOTOPOS 0 after load, even an ungated read is safe
    CHECK("defined caret 0 after load: ungated worst case still 0", computeAllStartPos_core(false, true, 0) == 0);
    printf("\n%s (%d)\n", fails==0?"ALL CHECKS PASSED":"CHECKS FAILED", fails);
    return fails==0?0:1;
}
