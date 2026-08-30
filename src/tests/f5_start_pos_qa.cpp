// F5 QA: start-position matrix for Replace All (editor vs. file scope).
// VERBATIM decision core of computeAllStartPos (non-selection branch) plus the
// fileScope gating introduced in handleReplaceAllButton/replaceAll.
#include <cstdio>
using Pos = long long;

Pos computeAllStartPos_core(bool wrapEnabled, bool fromCursorEnabled, Pos caretPos) {
    return wrapEnabled ? 0 : (fromCursorEnabled ? caretPos : 0);
}
Pos effectiveStart(bool fileScope, bool optionOn, bool wrap, Pos caret) {
    const bool fromCursor = !fileScope && optionOn;   // gating under test
    return computeAllStartPos_core(wrap, fromCursor, caret);
}
int fails = 0;
#define CHECK(d,c) do{bool ok=(c);printf("%-4s %s\n",ok?"PASS":"FAIL",d);if(!ok)++fails;}while(0)
int main() {
    const Pos END = 18; // caret after SCI_ADDTEXT (file end) resp. editor caret
    CHECK("file scope, option ON,  wrap off -> 0 (F5 fixed; was END)", effectiveStart(true,  true,  false, END) == 0);
    CHECK("file scope, option ON,  wrap on  -> 0",                     effectiveStart(true,  true,  true,  END) == 0);
    CHECK("file scope, option OFF, wrap off -> 0",                     effectiveStart(true,  false, false, END) == 0);
    CHECK("editor,     option ON,  wrap off -> caret (unchanged)",     effectiveStart(false, true,  false, END) == END);
    CHECK("editor,     option ON,  wrap on  -> 0 (unchanged)",         effectiveStart(false, true,  true,  END) == 0);
    CHECK("editor,     option OFF, wrap off -> 0 (unchanged)",         effectiveStart(false, false, false, END) == 0);
    // belt and suspenders: with SCI_GOTOPOS 0 after load, even an ungated read is safe
    CHECK("defined caret 0 after load: ungated worst case still 0",    computeAllStartPos_core(false, true, 0) == 0);
    printf("\n%s (%d)\n", fails==0?"ALL CHECKS PASSED":"CHECKS FAILED", fails);
    return fails==0?0:1;
}
