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

// QA for ColumnTabs (Flow Tabs) on the real Scintilla Document: padding in and out, undo and redo across switches, edits in the
// padding, save and gate flows, the fallback when the history no longer matches, and the save point.
// VERBATIM: the editor messages as Editor::WndProc hands them to the Document; text width is a fixed 8 px per character.
// Build: object files as for onepass_engine_qa (Notepad++ sources: scintilla + boostregex, Boost.Regex), then
//   g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src -I.. column_tabs_qa.cpp ../ColumnTabs.cpp ../NumericToken.cpp *.o -lboost_regex -o column_tabs_qa
#include <cstddef>
#include <cstdlib>
#include <cstdint>
#include <cstdarg>
#include <cassert>
#include <cstring>
#include <cstdio>
#include <cmath>
#include <stdexcept>
#include <new>
#include <utility>
#include <string>
#include <string_view>
#include <vector>
#include <array>
#include <map>
#include <set>
#include <forward_list>
#include <optional>
#include <algorithm>
#include <iterator>
#include <memory>
#include <chrono>
#include <regex>
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
#include "UniConversion.h"
#include "Scintilla.h"
#include "ColumnTabs.h"
using namespace Scintilla; using namespace Scintilla::Internal;
static constexpr int kCpUtf8 = 65001;
static constexpr int IND = 30;
static bool g_changeHistory = false;    // the document records change history, as Notepad++ sets it by default
static const char* g_pass = "";

// Platform stubs so Document links headless
namespace Scintilla::Internal {
void Platform::Assert(const char* c, const char* file, int line) noexcept { std::fprintf(stderr, "Assertion [%s] failed at %s %d\n", c, file, line); std::abort(); }
void Platform::DebugPrintf(const char*, ...) noexcept {}
void Platform::DebugDisplay(const char*) noexcept {}
bool Platform::ShowAssertionPopUps(bool) noexcept { return false; }
ColourRGBA Platform::Chrome() { return ColourRGBA(0xe0, 0xe0, 0xe0); }
ColourRGBA Platform::ChromeHighlight() { return ColourRGBA(0xff, 0xff, 0xff); }
const char* Platform::DefaultFont() { return "monospace"; }
int Platform::DefaultFontSize() { return 10; }
unsigned int Platform::DoubleClickTime() { return 500; }
}

// A document with the editor state Flow Tabs reaches through messages
struct Editor : DocWatcher {
    Document* d;
    ColumnTabs::Aligner aligner;
    Sci::Position targetStart = 0, targetEnd = 0;
    std::intptr_t mask = SC_MODEVENTMASKALL;
    std::intptr_t zoom = 0;
    std::map<std::intptr_t, std::vector<int>> stops;
    std::vector<std::pair<std::intptr_t, bool>> markers;   // container notifications: token, undo
    int announced = 0;                                      // text changes the container was told about
    int savePointLeft = 0, savePointReached = 0;
    bool cloned = false;                                    // the document also shows in the other view
    std::intptr_t changeHistoryOption = 0;

    explicit Editor(const std::string& s) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(kCpUtf8);
        d->InsertString(0, s.c_str(), static_cast<Sci::Position>(s.size()));
        d->DeleteUndoHistory();
        if (g_changeHistory) {
            changeHistoryOption = SC_CHANGE_HISTORY_ENABLED | SC_CHANGE_HISTORY_MARKERS;
            d->ChangeHistorySet(true);
        }
        d->AddWatcher(this, nullptr);
        aligner.setIndicator(IND);
    }
    ~Editor() override { d->RemoveWatcher(this, nullptr); d->Release(); }
    Editor(const Editor&) = delete;
    Editor& operator=(const Editor&) = delete;

    ColumnTabs::Send send() {
        return [this](unsigned int msg, std::uintptr_t w, std::intptr_t l) -> std::intptr_t {
            const auto pos = static_cast<Sci::Position>(w);
            switch (msg) {
            case SCI_GETLENGTH: return d->Length();
            case SCI_GETLINECOUNT: return d->LinesTotal();
            case SCI_POSITIONFROMLINE: return pos > d->LinesTotal() ? -1 : d->LineStart(pos);
            case SCI_GETLINEENDPOSITION: return pos >= d->LinesTotal() ? d->Length() : d->LineEnd(pos);
            case SCI_LINEFROMPOSITION: return d->SciLineFromPosition(pos);
            case SCI_GETRANGEPOINTER: return reinterpret_cast<std::intptr_t>(d->RangePointer(pos, l));
            case SCI_GETCHARACTERPOINTER: return reinterpret_cast<std::intptr_t>(d->BufferPointer());
            case SCI_GETDOCPOINTER: return reinterpret_cast<std::intptr_t>(d);
            case SCI_SETTARGETRANGE: targetStart = pos; targetEnd = l; return 0;
            case SCI_REPLACETARGET: {   // Editor::ReplaceTarget, basic
                const std::string_view text(reinterpret_cast<const char*>(l), static_cast<std::size_t>(w));
                d->BeginUndoAction();
                if (targetEnd > targetStart) d->DeleteChars(targetStart, targetEnd - targetStart);
                targetEnd = targetStart;
                const Sci::Position n = d->InsertString(targetStart, text.data(), static_cast<Sci::Position>(text.size()));
                targetEnd = targetStart + n;
                d->EndUndoAction();
                return static_cast<std::intptr_t>(text.size());
            }
            case SCI_DELETERANGE: d->DeleteChars(pos, l); return 0;
            case SCI_GETUNDOCOLLECTION: return d->IsCollectingUndo();
            case SCI_SETUNDOCOLLECTION: d->SetUndoCollection(w != 0); return 0;
            case SCI_ADDUNDOACTION: d->AddUndoAction(pos, (l & UNDO_MAY_COALESCE) != 0); return 0;
            case SCI_EMPTYUNDOBUFFER: d->DeleteUndoHistory(); return 0;
            case SCI_SETUNDOSAVEPOINT: d->SetUndoSavePoint(static_cast<int>(w)); return 0;
            case SCI_GETUNDOSAVEPOINT: return d->UndoSavePoint();
            case SCI_GETUNDOCURRENT: return d->UndoCurrent();
            case SCI_GETMODIFY: return !d->IsSavePoint();
            case SCI_SETSAVEPOINT: d->SetSavePoint(); return 0;
            case SCI_GETREADONLY: return d->IsReadOnly();
            case SCI_SETREADONLY: d->SetReadOnly(w != 0); return 1;
            case SCI_GETCHANGEHISTORY: return changeHistoryOption;
            case SCI_SETCHANGEHISTORY: changeHistoryOption = static_cast<std::intptr_t>(w); d->ChangeHistorySet((w & 1) != 0); return 0;
            case SCI_CANUNDO: return d->CanUndo();
            case SCI_GETMODEVENTMASK: return mask;
            case SCI_SETMODEVENTMASK: mask = static_cast<std::intptr_t>(w); return 0;
            case SCI_GETZOOM: return zoom;
            case SCI_SETINDICATORCURRENT: d->DecorationSetCurrentIndicator(static_cast<int>(w)); return 0;
            case SCI_GETINDICATORCURRENT: return d->decorations->GetCurrentIndicator();
            case SCI_SETINDICATORVALUE: d->decorations->SetCurrentValue(static_cast<int>(w)); return 0;
            case SCI_GETINDICATORVALUE: return d->decorations->GetCurrentValue();
            case SCI_INDICATORFILLRANGE: d->DecorationFillRange(pos, d->decorations->GetCurrentValue(), l); return 0;
            case SCI_INDICATORCLEARRANGE: d->DecorationFillRange(pos, 0, l); return 0;
            case SCI_INDICATORVALUEAT: return d->decorations->ValueAt(static_cast<int>(w), l);
            case SCI_INDICATOREND: return d->decorations->End(static_cast<int>(w), l);
            case SCI_CLEARTABSTOPS: stops.erase(static_cast<std::intptr_t>(w)); return 0;
            case SCI_ADDTABSTOP: stops[static_cast<std::intptr_t>(w)].push_back(static_cast<int>(l)); return 0;
            case SCI_GETNEXTTABSTOP: {
                int next = 0;
                const auto it = stops.find(static_cast<std::intptr_t>(w));
                if (it != stops.end())
                    for (const int x : it->second) if (x > l && (next == 0 || x < next)) next = x;
                return next;
            }
            case SCI_TEXTWIDTH: {   // fixed 8 px per UTF-8 character
                int n = 0;
                for (const char* p = reinterpret_cast<const char*>(l); *p; ++p)
                    if ((static_cast<unsigned char>(*p) & 0xC0) != 0x80) ++n;
                return 8 * n;
            }
            default: std::fprintf(stderr, "unexpected message %u\n", msg); std::abort();
            }
        };
    }

    // Editor::NotifyModified: the container hears what the mask lets through
    void NotifyModified(Document*, DocModification mh, void*) override {
        const int type = static_cast<int>(mh.modificationType);
        if (!(type & mask)) return;
        if (type & SC_MOD_CONTAINER) {
            markers.push_back({ mh.token, (type & SC_PERFORMED_UNDO) != 0 });
            return;
        }
        if (type & (SC_MOD_INSERTTEXT | SC_MOD_DELETETEXT)) ++announced;
        aligner.onModified(send(), type, mh.position, mh.length, mh.text, mh.linesAdded);
        if (cloned) aligner.onModified(send(), type, mh.position, mh.length, mh.text, mh.linesAdded);   // the other view
    }
    void NotifyModifyAttempt(Document*, void*) override {}
    void NotifySavePoint(Document*, void*, bool atSavePoint) override { ++(atSavePoint ? savePointReached : savePointLeft); }
    void NotifyDeleted(Document*, void*) noexcept override {}
    void NotifyStyleNeeded(Document*, void*, Sci::Position) override {}
    void NotifyErrorOccurred(Document*, void*, Scintilla::Status) override {}
    void NotifyGroupCompleted(Document*, void*) noexcept override {}

    std::string text() const { return std::string(d->BufferPointer(), static_cast<size_t>(d->Length())); }
    bool marked(Sci::Position p) const { return d->decorations->ValueAt(IND, p) != 0; }
    int markedCount() const { int n = 0; for (Sci::Position p = 0; p < d->Length(); ++p) n += marked(p); return n; }
    void type(Sci::Position p, const std::string& s) { d->InsertString(p, s.c_str(), static_cast<Sci::Position>(s.size())); }
    void erase(Sci::Position p, Sci::Position n) { d->DeleteChars(p, n); }
    void replace(Sci::Position p, Sci::Position n, const std::string& s) { d->BeginUndoAction(); erase(p, n); type(p, s); d->EndUndoAction(); }

    // Ctrl+Z / Ctrl+Y as in Notepad++: the step, then a marker the step reached is replayed
    ColumnTabs::Replayed step(bool undo) {
        markers.clear();
        if (undo) d->Undo(); else d->Redo();
        ColumnTabs::Replayed r = ColumnTabs::Replayed::NotOurs;
        for (const auto& [token, wasUndo] : markers) r = aligner.replay(send(), token, wasUndo, nullptr);
        markers.clear();
        return r;
    }
    ColumnTabs::SwitchedOn on(const std::string& delim = ";", bool numbers = true, char quote = '\0') {
        ColumnTabs::Settings s; s.delimiter = delim; s.alignNumbers = numbers; s.quote = quote;
        return aligner.switchOn(send(), s, nullptr);
    }
    bool off() { return aligner.switchOff(send(), nullptr); }
};

static int fails = 0, checks = 0;
static void check(const char* name, bool ok) { ++checks; if (!ok) { ++fails; std::printf("FAIL %s%s\n", g_pass, name); } }
static std::string vis(std::string s) { for (auto& c : s) { if (c == '\t') c = '>'; else if (c == '\n') c = '|'; } return s; }
static void checkText(const char* name, const Editor& e, const std::string& want) {
    const bool ok = e.text() == want;
    check(name, ok);
    if (!ok) std::printf("     got  \"%s\"\n     want \"%s\"\n", vis(e.text()).c_str(), vis(want).c_str());
}

static const std::string DECIMAL = "Item;Price\nA;5\nB;1,25\nC;10";
static const std::string DECIMAL_PADDED = "Item\t;Price\nA\t;  5\nB\t;  1,25\nC\t; 10";

static void runAll() {
    using ColumnTabs::SwitchedOn; using ColumnTabs::Replayed;

    // Change history: undo across the switches, deletions next to the padding and in it. Scintilla's change history
    // breaks on text changes outside the undo history, Flow Tabs take it off the document before they change the text
    {
        Editor e(DECIMAL);
        e.on();
        if (g_changeHistory) check("history: view keeps its setting", e.changeHistoryOption == (SC_CHANGE_HISTORY_ENABLED | SC_CHANGE_HISTORY_MARKERS));
        e.off();
        check("history: undo off", e.step(true) == Replayed::Padded);
        checkText("history: padded again", e, DECIMAL_PADDED);
        check("history: undo on", e.step(true) == Replayed::Unpadded);
        checkText("history: plain again", e, DECIMAL);
        check("history: redo on", e.step(false) == Replayed::Padded);
        check("history: redo off", e.step(false) == Replayed::Unpadded);
        checkText("history: plain at the end", e, DECIMAL);
    }
    {
        Editor e(DECIMAL);
        e.on();
        e.erase(12, 1);                 // 'A', right before its padding
        e.off();
        check("history: deletion, undo off", e.step(true) == Replayed::Padded);
        e.step(true);
        checkText("history: deletion undone", e, DECIMAL_PADDED);
        check("history: deletion, undo on", e.step(true) == Replayed::Unpadded);
        checkText("history: deletion, plain", e, DECIMAL);
        check("history: deletion, redo all", e.step(false) == Replayed::Padded && (e.step(false), e.step(false)) == Replayed::Unpadded);
        checkText("history: deletion, redone", e, "Item;Price\n;5\nB;1,25\nC;10");
    }
    {
        Editor e(DECIMAL);
        e.on();
        e.erase(13, 1);                 // a padding TAB, by hand
        e.off();
        e.step(true);
        e.step(true);
        checkText("history: padding deletion undone", e, DECIMAL_PADDED);
        check("history: padding marked again", e.marked(13));
        e.step(true);
        checkText("history: padding deletion, plain", e, DECIMAL);
    }

    // Padding in and out: TAB before each delimiter, spaces before numbers, no ".00"
    {
        Editor e(DECIMAL);
        std::vector<ColumnTabs::Range> ins;
        ColumnTabs::Settings s; s.delimiter = ";";
        check("on: padded", e.aligner.switchOn(e.send(), s, &ins) == SwitchedOn::Padded);
        checkText("on: text", e, DECIMAL_PADDED);
        check("on: padding marked", e.markedCount() == 4 + 2 + 2 + 1);
        check("on: inserted ranges", ins.size() == 7 && ins[0].start == 4 && ins[0].end == 5);
        check("on: stops on every line", e.stops.size() == 4 && e.stops[0].size() == 1);
        check("on: stop after widest cell", e.stops[0][0] == 8 * 4 + 2 * 8 + 4);
        check("on: one marker", e.d->CanUndo());
        std::vector<ColumnTabs::Range> rem;
        check("off", e.aligner.switchOff(e.send(), &rem));
        checkText("off: text", e, DECIMAL);
        check("off: removed ranges before removal", rem.size() == 7 && rem[0].start == 4 && rem[1].start == 13 && rem[2].start == 15 && rem[2].end == 17);
        check("off: no stops", e.stops.empty());
        check("off: no marks", e.markedCount() == 0);
    }

    // Undo and redo across switches put the padding back exactly
    {
        Editor e(DECIMAL);
        e.on(); e.off();
        check("undo off -> padded", e.step(true) == Replayed::Padded);
        checkText("undo off: text", e, DECIMAL_PADDED);
        check("undo off: aligned again", e.aligner.padded(e.send()) && !e.stops.empty());
        check("undo on -> unpadded", e.step(true) == Replayed::Unpadded);
        checkText("undo on: text", e, DECIMAL);
        check("redo on -> padded", e.step(false) == Replayed::Padded);
        checkText("redo on: text", e, DECIMAL_PADDED);
        check("redo off -> unpadded", e.step(false) == Replayed::Unpadded);
        checkText("redo off: text", e, DECIMAL);
    }

    // K2: typed in aligned text, switched off, Ctrl+Z three times: every step lands where it belongs
    {
        Editor e(DECIMAL);
        e.d->SetSavePoint();
        e.on();
        const Sci::Position five = static_cast<Sci::Position>(DECIMAL_PADDED.find("5\n"));
        e.type(five + 1, "55");                                    // "A;555"
        const std::string typed = e.text();
        e.off();
        checkText("k2: off keeps typed digits", e, "Item;Price\nA;555\nB;1,25\nC;10");
        e.step(true);
        checkText("k2: undo off", e, typed);
        e.step(true);
        checkText("k2: undo typing", e, DECIMAL_PADDED);
        check("k2: back at the save point, right after on", e.d->IsSavePoint());
        e.step(true);
        checkText("k2: undo on", e, DECIMAL);
    }

    // K1: a saved document stays saved through the switches, and the next edit is announced
    {
        Editor e(DECIMAL);
        e.d->SetSavePoint();
        e.on();
        check("k1: saved after on", e.d->IsSavePoint());
        e.off();
        check("k1: saved after off", e.d->IsSavePoint());
        const int left = e.savePointLeft;
        e.type(0, "x");
        check("k1: edit leaves the save point, announced", !e.d->IsSavePoint() && e.savePointLeft == left + 1);
    }
    {
        Editor e(DECIMAL);
        e.d->SetSavePoint();
        e.on(); e.off();
        e.step(true);
        check("k1: undo off keeps it saved", e.d->IsSavePoint());
        e.step(true);
        check("k1: undo on keeps it saved", e.d->IsSavePoint());
        e.step(false); e.step(false);
        check("k1: redo keeps it saved", e.d->IsSavePoint());
        e.on(); e.type(0, "x");
        const int reached = e.savePointReached;
        e.step(true);
        check("k1: undo an edit back to the save point", e.d->IsSavePoint() && e.savePointReached == reached + 1);
        e.step(true);
        check("k1: undo on after it keeps it saved", e.d->IsSavePoint());
    }

    // An unsaved document stays unsaved through the switches and the fallback
    {
        Editor e(DECIMAL);
        e.d->SetSavePoint();
        e.type(0, "Z");
        e.on(); e.off();
        check("unsaved stays unsaved", !e.d->IsSavePoint());
        e.d->SetUndoCollection(false); e.type(1, "?"); e.d->SetUndoCollection(true);   // changed outside undo
        check("history no longer matches -> reset", e.step(true) == Replayed::Reset);
        check("reset: unsaved", !e.d->IsSavePoint() && !e.d->CanUndo());
        e.type(0, "q"); e.d->Undo();
        check("reset: unsaved after type and undo", !e.d->IsSavePoint());
    }

    // The padding went another way (Notepad++ reloads the file): switching off leaves no step for undo
    {
        Editor e(DECIMAL);
        e.on();
        e.erase(0, e.d->Length()); e.type(0, DECIMAL); e.d->DeleteUndoHistory();
        check("reloaded: off", e.off());
        checkText("reloaded: text", e, DECIMAL);
        check("reloaded: no step", !e.d->CanUndo() && e.stops.empty());
        check("reloaded: on again", e.on() == SwitchedOn::Padded);
        checkText("reloaded: padded", e, DECIMAL_PADDED);
    }

    // Progress: rising, in steps, both ways; a short run reports nothing
    {
        std::string big;
        for (int i = 0; i < 5000; ++i) big += "row" + std::to_string(i) + ";" + std::to_string(i % 97) + "\n";
        Editor e(big);
        std::vector<std::pair<bool, int>> seen;
        // Rising below 100 while running, then 100 once: the end of the run
        const auto rising = [&] {
            for (std::size_t i = 1; i < seen.size(); ++i) if (seen[i].second < seen[i - 1].second) return false;
            for (std::size_t i = 0; i + 1 < seen.size(); ++i) if (seen[i].second >= 100) return false;
            return !seen.empty() && seen.back().second == 100;
        };
        e.aligner.setProgress([&](bool padding, int percent) { seen.push_back({ padding, percent }); }, 0, 0);
        e.on();
        check("progress: on, rising", rising() && seen.front().first && seen.size() >= 10);
        seen.clear();
        e.step(true);
        check("progress: undo on, rising", rising() && !seen.front().first);
        seen.clear();
        e.step(false);
        check("progress: redo on, rising", rising() && seen.front().first);
        seen.clear();
        e.off();
        check("progress: off, rising", rising() && !seen.front().first);
        seen.clear();
        e.aligner.setProgress([&](bool padding, int percent) { seen.push_back({ padding, percent }); }, 60000, 0);
        e.on();
        check("progress: short run reports nothing", seen.empty());
        checkText("progress: padded as without", e, [&] { Editor f(big); f.on(); return f.text(); }());
    }

    // A cell typed wider than its column moves only its own line on, and as little as possible
    {
        Editor e(DECIMAL);
        e.on();
        const std::vector<int> shared{ 52 };
        e.type(13, "XXXXXX");                                      // "AXXXXXX" before the padding TAB
        check("fit: wider cell, own stops", e.stops[1] == std::vector<int>{ 76 });
        check("fit: other lines keep theirs", e.stops[0] == shared && e.stops[2] == shared && e.stops[3] == shared);
        e.erase(13, 6);
        check("fit: narrow again, shared stops", e.stops[1] == shared);
        e.step(true);
        check("fit: undo brings the wide cell and its stops", e.stops[1] == std::vector<int>{ 76 });
        e.step(false);
        check("fit: redo narrows again", e.stops[1] == shared);
    }
    {
        Editor e("aaaa;bbbbbbbbbb;c\nx;b;c");
        e.on();
        const std::vector<int> shared{ 52, 160 };
        check("fit: shared stops", e.stops[0] == shared && e.stops[1] == shared);
        e.type(21, "yyyyy");                                       // "xyyyyy", the next cell still fits
        check("fit: the next column takes the overflow up", e.stops[1] == (std::vector<int>{ 68, 160 }));
        e.type(e.d->Length(), "\nzz\t;q");
        check("fit: a new line with a TAB in front of the delimiter gets the shared stops", e.stops[2] == shared);
        e.off();
        check("fit: off, no stops", e.stops.empty());
        e.step(true);
        check("fit: undo off, the wide line keeps its own stops", e.stops[1] == (std::vector<int>{ 68, 160 }) && e.stops[0] == shared);
        e.stops.clear();                                           // the other view, showing the same document
        e.aligner.ensureStops(e.send());
        check("fit: a view without stops gets them", e.stops[0] == shared && e.stops[1] == (std::vector<int>{ 68, 160 }));
        e.step(false);
        e.stops[0] = shared;                                       // stale stops in the other view
        e.aligner.restoreStops(e.send());
        check("fit: no longer aligned, stops cleared", e.stops.empty());
    }
    {
        Editor e("aaaa;bbbbbbbbbb;c\nx;b;c");
        e.on();
        const std::vector<int> shared{ 52, 160 };
        e.type(22, "yyyyy");                                       // "x\tyyyyy;b": typed behind the padding TAB
        check("fit: text behind the padding TAB keeps the stops", e.stops[1] == shared);
        e.type(27, "zzzzzzzzzz");                                  // past the next column
        check("fit: text behind the padding TAB moves only the next column", e.stops[1] == (std::vector<int>{ 52, 208 }));
    }
    {
        Editor e("aaaa;bbbbbbbbbb;c\nx;b;c");
        e.on();
        e.erase(21, 1);                                            // "x;b\t;c": the first cell lost its padding
        check("fit: a cell without padding uses no stop", e.stops[1] == std::vector<int>{ 160 });
    }

    // Typed text never becomes padding, also inside a padding run
    {
        Editor e(DECIMAL);
        e.on();
        const Sci::Position spaces = static_cast<Sci::Position>(DECIMAL_PADDED.find("  5"));
        e.type(spaces + 1, "7");                                   // between the two spaces before 5
        check("typed inside padding is not padding", !e.marked(spaces + 1) && e.marked(spaces) && e.marked(spaces + 2));
        e.off();
        checkText("typed digit stays", e, "Item;Price\nA;75\nB;1,25\nC;10");
    }

    // K8: padding TAB deleted and brought back by Ctrl+Z is padding again
    {
        Editor e(DECIMAL);
        e.on();
        e.erase(4, 1);                                             // backspace on the TAB of "Item\t;"
        e.d->Undo();
        check("k8: tab back and marked", e.text() == DECIMAL_PADDED && e.marked(4));
        e.off();
        checkText("k8: off", e, DECIMAL);
        e.on();
        e.erase(0, e.d->Length());                                 // select all, delete
        e.d->Undo();
        check("k8: all back, all marked", e.text() == DECIMAL_PADDED && e.markedCount() == 9);
        e.off();
        checkText("k8: all back, off", e, DECIMAL);
    }

    // K7: lines rewritten in another order keep their padding (sort lines)
    {
        Editor e(DECIMAL);
        e.on();
        const std::string t = e.text();
        const size_t l2 = t.find('\n') + 1;
        const std::string body = t.substr(l2);                       // "A..|B..|C.."
        const std::string a = body.substr(0, body.find('\n') + 1);
        const std::string rest = body.substr(a.size());
        const std::string b = rest.substr(0, rest.find('\n') + 1);
        const std::string c = rest.substr(b.size());
        e.replace(static_cast<Sci::Position>(l2), static_cast<Sci::Position>(body.size()), c + "\n" + b + a.substr(0, a.size() - 1));
        e.off();
        checkText("k7: sorted lines without padding", e, "Item;Price\nC;10\nB;1,25\nA;5");
    }
    {
        // The document in both views: each change is reported twice, the padding follows it once
        Editor e(DECIMAL);
        e.cloned = true;
        e.on();
        const std::string t = e.text();
        const size_t l2 = t.find('\n') + 1;
        const std::string body = t.substr(l2);
        const std::string a = body.substr(0, body.find('\n') + 1);
        const std::string rest = body.substr(a.size());
        const std::string b = rest.substr(0, rest.find('\n') + 1);
        const std::string c = rest.substr(b.size());
        e.replace(static_cast<Sci::Position>(l2), static_cast<Sci::Position>(body.size()), c + "\n" + b + a.substr(0, a.size() - 1));
        e.erase(5, 1);                                             // the padding TAB after "Item"
        e.step(true);
        check("clone: undo marks the padding again", e.marked(4));
        e.off();
        checkText("clone: sorted lines without padding", e, "Item;Price\nC;10\nB;1,25\nA;5");
    }

    // Exceptions: a deleted padding TAB, a new line, a wider number; switching off and Ctrl+Z restore exactly
    {
        Editor e(DECIMAL);
        e.on();
        e.erase(4, 1);
        e.type(e.d->Length(), "\nD;1234");
        const Sci::Position ten = static_cast<Sci::Position>(e.text().find("10"));
        e.type(ten, "9");
        const std::string before = e.text();
        e.off();
        checkText("exceptions: off", e, "Item;Price\nA;5\nB;1,25\nC;910\nD;1234");
        check("exceptions: undo off -> padded", e.step(true) == Replayed::Padded);
        checkText("exceptions: exact", e, before);
        e.step(true); e.step(true); e.step(true);
        checkText("exceptions: back to the start", e, DECIMAL_PADDED);
        check("exceptions: undo on", e.step(true) == Replayed::Unpadded);
        checkText("exceptions: original", e, DECIMAL);
    }

    // Saving: out, save, the same padding back; the history goes on across it
    {
        Editor e(DECIMAL);
        e.on();
        e.type(0, "x");
        const std::string edited = e.text();
        e.off();
        e.d->SetSavePoint();
        std::vector<ColumnTabs::Range> ins;
        ColumnTabs::Settings s; s.delimiter = ";";
        check("save: realign exact", e.aligner.realign(e.send(), s, &ins) == SwitchedOn::Padded);
        checkText("save: padded again", e, edited);
        check("save: still saved", e.d->IsSavePoint());
        check("save: inserted ranges", ins.size() == 7);
        e.step(true);
        check("save: undo realign keeps it saved", e.d->IsSavePoint());
        e.step(true);
        checkText("save: undo realign and save", e, edited);
        check("save: undo the switch off keeps it saved", e.d->IsSavePoint());
        e.step(true);
        checkText("save: undo edit", e, DECIMAL_PADDED);
        check("save: undo edit leaves the save point", !e.d->IsSavePoint());
        e.step(true);
        checkText("save: undo on", e, DECIMAL);
        e.step(false); e.step(false); e.step(false); e.step(false);
        checkText("save: redo all", e, edited);
        check("save: saved again at the save point", e.d->IsSavePoint());
    }

    // Gate: out, change the text, align anew; Ctrl+Z walks back through all of it
    {
        Editor e(DECIMAL);
        e.on();
        e.off();
        e.replace(5, 5, "Cost");                                    // "Item;Cost"
        ColumnTabs::Settings s; s.delimiter = ";";
        check("gate: realign fresh", e.aligner.realign(e.send(), s, nullptr) == SwitchedOn::Padded);
        checkText("gate: padded", e, "Item\t;Cost\nA\t;  5\nB\t;  1,25\nC\t; 10");
        e.step(true);
        checkText("gate: undo realign", e, "Item;Cost\nA;5\nB;1,25\nC;10");
        e.step(true);
        checkText("gate: undo change", e, DECIMAL);
        e.step(true);
        checkText("gate: undo off", e, DECIMAL_PADDED);
        e.step(true);
        checkText("gate: undo on", e, DECIMAL);
    }

    // The strip is announced once, the padding not
    {
        Editor e(DECIMAL);
        e.announced = 0;
        e.on();
        check("announce: on silent", e.announced == 0);
        e.off();
        check("announce: off once", e.announced == 1);
    }

    // Tab-delimited: stops only, no marker; numbers get spaces
    {
        Editor e("a\tb\n1\t22\n333\t4");
        check("tsv: stops only", e.on("\t", false) == SwitchedOn::StopsOnly);
        check("tsv: text unchanged, no marker", e.text() == "a\tb\n1\t22\n333\t4" && !e.d->CanUndo());
        check("tsv: stops", e.stops.size() == 3 && e.stops[0][0] == 8 * 3 + 20);
        e.off();
        check("tsv: stops gone", e.stops.empty());
        check("tsv numbers: padded", e.on("\t", true) == SwitchedOn::Padded);
        checkText("tsv numbers: text", e, "a\tb\n   1\t 22\n 333\t  4");
        e.off();
        checkText("tsv numbers: off", e, "a\tb\n1\t22\n333\t4");
    }

    // Quotes, multi-character delimiters, read-only, nothing to align
    {
        Editor q("\"a;b\";c\nx;y");
        q.on(";", true, '"');
        checkText("quoted delimiter stays", q, "\"a;b\"\t;c\nx\t;y");
        Editor m("a::b::c\n1::22::333");
        m.on("::", false);
        checkText("two-character delimiter", m, "a\t::b\t::c\n1\t::22\t::333");
        Editor r(DECIMAL);
        r.d->SetReadOnly(true);
        check("read-only refused", r.on() == SwitchedOn::ReadOnly && r.text() == DECIMAL);
        Editor n("no delimiter here");
        check("nothing to align", n.on() == SwitchedOn::NothingToAlign && !n.d->CanUndo());
    }

    // A read-only document still loses its padding
    {
        Editor e(DECIMAL);
        e.on();
        e.d->SetReadOnly(true);
        e.off();
        check("read-only: padding out, flag kept", e.text() == DECIMAL && e.d->IsReadOnly());
    }

    // New lines get the stops their cells need
    {
        Editor e(DECIMAL);
        e.on();
        e.type(e.d->Length(), "\nE;7");
        check("new line without padding: no stops of its own", e.stops.count(4) == 0);
        e.type(e.d->Length() - 2, "\t");                         // "E\t;7"
        check("new line with a TAB in front of the delimiter: the shared stop", e.stops[4] == std::vector<int>{ 52 });
    }

    // Text without padding; markers
    {
        Editor e(DECIMAL);
        e.on();
        e.type(e.d->Length(), "\nE;7");
        check("without padding", e.aligner.withoutPadding(e.send(), 0, e.d->Length()) == "Item;Price\nA;5\nB;1,25\nC;10\nE;7");
        check("marker tokens", ColumnTabs::Aligner::isMarker(0x4D520001) && !ColumnTabs::Aligner::isMarker(0x4D52464C + 0x10000)
            && !ColumnTabs::Aligner::isMarker(0x4D520000) && !ColumnTabs::Aligner::isMarker(42));
        e.aligner.forget(reinterpret_cast<std::intptr_t>(e.d));
        e.step(true);
        const std::string padded = e.text();
        check("forgotten marker resets", e.step(true) == Replayed::Reset && e.text() == padded && !e.d->CanUndo());
    }

    // Redo of typing is typing: not padding
    {
        Editor e(DECIMAL);
        e.on();
        const Sci::Position spaces = static_cast<Sci::Position>(DECIMAL_PADDED.find("  5"));
        e.type(spaces + 1, "7");
        e.d->Undo();
        e.d->Redo();
        check("redo typed: not padding", !e.marked(spaces + 1));
    }

    // Another document in the view: stops gone; stops only means aligned no more, padding stays aligned
    {
        Editor t("a\tb\n1\t22");
        t.on("\t", false);
        t.stops.clear();
        t.aligner.leave(reinterpret_cast<std::intptr_t>(t.d));
        check("leave: stops only ends", !t.aligner.aligned(t.send()));
        Editor e(DECIMAL);
        e.on();
        e.stops.clear();
        e.aligner.leave(reinterpret_cast<std::intptr_t>(e.d));
        check("leave: padding stays aligned", e.aligner.padded(reinterpret_cast<std::intptr_t>(e.d)));
        e.aligner.restoreStops(e.send());
        check("restore stops", e.stops.size() == 4 && e.stops[1][0] == 8 * 4 + 20);
    }

    // Zoom: stops measured again for the new zoom when the padding comes back
    {
        Editor e(DECIMAL);
        e.on(); e.off();
        e.zoom = 2;
        e.step(true);
        check("zoom: stops present", !e.stops.empty());
    }

    // Timing on 100,000 rows
}

// FT_BENCH=<file>: one round trip on a big file, with change history as in Notepad++
static void bench() {
    using ColumnTabs::Replayed;
    if (const char* path = std::getenv("FT_BENCH")) {
        FILE* f = std::fopen(path, "rb");
        if (f) {
            std::string big;
            char buf[1 << 16];
            for (size_t n; (n = std::fread(buf, 1, sizeof buf, f)) > 0;) big.append(buf, n);
            std::fclose(f);
            Editor e(big);
            const auto t0 = std::chrono::steady_clock::now();
            e.on();
            const auto t1 = std::chrono::steady_clock::now();
            e.off();
            const auto t2 = std::chrono::steady_clock::now();
            e.step(true);
            const auto t3 = std::chrono::steady_clock::now();
            const auto ms = [](auto a, auto b) { return std::chrono::duration<double, std::milli>(b - a).count(); };
            std::printf("bench: on %.0f ms, off %.0f ms, undo off %.0f ms, length %ld\n", ms(t0, t1), ms(t1, t2), ms(t2, t3), static_cast<long>(e.d->Length()));
            check("bench: roundtrip", e.step(true) == Replayed::Unpadded && e.text() == big);
        }
    }
}

int main() {
    runAll();
    g_changeHistory = true;
    g_pass = "[change history] ";
    runAll();
    bench();
    std::printf("%d checks, %d failed\n", checks, fails);
    return fails ? 1 : 0;
}
