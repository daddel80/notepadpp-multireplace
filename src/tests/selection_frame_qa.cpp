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

// QA for SelectionFrame on the real Scintilla Document: the area the Selection scope remembers follows edits
// the way Scintilla moves indicators, announced edits keep replaced text framed, generations retire old frames,
// and frames belong to their document.
// VERBATIM: the indicator messages as Editor::WndProc hands them to the Document.
// Build: object files as for onepass_engine_qa (Notepad++ sources: scintilla + boostregex, Boost.Regex), then
//   g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src -I.. selection_frame_qa.cpp ../SelectionFrame.cpp *.o -lboost_regex -o selection_frame_qa
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
#include "../SelectionFrame.h"
using namespace Scintilla; using namespace Scintilla::Internal;
using SelectionFrame::Span;
static constexpr int kCpUtf8 = 65001;

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

struct Doc {
    Document* d;
    explicit Doc(const std::string& s) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(kCpUtf8);
        d->InsertString(0, s.c_str(), static_cast<Sci::Position>(s.size()));
        d->DeleteUndoHistory();
    }
    ~Doc() { d->Release(); }
    Doc(const Doc&) = delete;
    Doc& operator=(const Doc&) = delete;

    // The messages SelectionFrame sends, as Editor::WndProc passes them on
    SelectionFrame::Send send() {
        return [this](unsigned int msg, std::uintptr_t w, std::intptr_t l) -> std::intptr_t {
            switch (msg) {
            case SCI_GETLENGTH: return d->Length();
            case SCI_SETINDICATORCURRENT: d->DecorationSetCurrentIndicator(static_cast<int>(w)); return 0;
            case SCI_SETINDICATORVALUE: d->decorations->SetCurrentValue(static_cast<int>(w)); return 0;
            case SCI_GETINDICATORCURRENT: return d->decorations->GetCurrentIndicator();
            case SCI_GETINDICATORVALUE: return d->decorations->GetCurrentValue();
            case SCI_INDICATORFILLRANGE: d->DecorationFillRange(static_cast<Sci::Position>(w), d->decorations->GetCurrentValue(), l); return 0;
            case SCI_INDICATORCLEARRANGE: d->DecorationFillRange(static_cast<Sci::Position>(w), 0, l); return 0;
            case SCI_INDICATORVALUEAT: return d->decorations->ValueAt(static_cast<int>(w), l);
            case SCI_INDICATOREND: return d->decorations->End(static_cast<int>(w), l);
            case SCI_GETDOCPOINTER: return reinterpret_cast<std::intptr_t>(d);
            default: std::fprintf(stderr, "unexpected message %u\n", msg); std::abort();
            }
        };
    }
    Sci::Position line(Sci::Line n) const { return d->LineStart(n); }
    std::string text(Span s) const {
        std::string out(static_cast<size_t>(s.end - s.start), '\0');
        d->GetCharRange(out.data(), s.start, s.end - s.start);
        return out;
    }
    void insert(Sci::Position pos, const std::string& s) { d->InsertString(pos, s.c_str(), static_cast<Sci::Position>(s.size())); }
    void erase(Sci::Position pos, Sci::Position len) { d->DeleteChars(pos, len); }
    // Editor::ReplaceTarget without patterns: delete, then insert
    void replace(Sci::Position pos, Sci::Position len, const std::string& s) { d->BeginUndoAction(); erase(pos, len); insert(pos, s); d->EndUndoAction(); }
};

// Edits announced the way MR passes SCN_MODIFIED on: deletions before, insertions after; off for MR's own Replace
struct Announce : DocWatcher {
    Doc& doc;
    SelectionFrame::Frame& frame;
    bool on = true;
    Announce(Doc& d, SelectionFrame::Frame& f) : doc(d), frame(f) { doc.d->AddWatcher(this, nullptr); }
    ~Announce() override { doc.d->RemoveWatcher(this, nullptr); }
    Announce(const Announce&) = delete;
    Announce& operator=(const Announce&) = delete;
    void NotifyModified(Document*, DocModification mh, void*) override {
        const int type = static_cast<int>(mh.modificationType);
        if (!on) return;
        if (type & SC_MOD_BEFOREDELETE) frame.noteDelete(doc.send(), mh.position, mh.length);
        else if (type & SC_MOD_INSERTTEXT) frame.keepInsert(doc.send(), mh.position, mh.length);
    }
    void NotifyModifyAttempt(Document*, void*) override {}
    void NotifySavePoint(Document*, void*, bool) override {}
    void NotifyDeleted(Document*, void*) noexcept override {}
    void NotifyStyleNeeded(Document*, void*, Sci::Position) override {}
    void NotifyErrorOccurred(Document*, void*, Scintilla::Status) override {}
    void NotifyGroupCompleted(Document*, void*) noexcept override {}
};

static int fails = 0, checks = 0;
static void check(const char* name, bool ok) { ++checks; if (!ok) { ++fails; std::printf("FAIL %s\n", name); } }

static std::vector<std::string> texts(Doc& doc, const SelectionFrame::Frame& f) {
    std::vector<std::string> v;
    for (const Span s : f.read(doc.send())) v.push_back(doc.text(s));
    return v;
}
// Rows `lines` long from `first`, `width` characters from column `col`, like a rectangular selection
static std::vector<Span> block(const Doc& doc, int first, int lines, int col, int width) {
    std::vector<Span> v;
    for (int l = first; l < first + lines; ++l) v.push_back({ doc.line(l) + col, doc.line(l) + col + width });
    return v;
}

int main() {
    SelectionFrame::Frame f;
    f.setIndicator(9);

    {
        std::map<unsigned int, std::intptr_t> style;
        f.applyStyle([&](unsigned int msg, std::uintptr_t w, std::intptr_t l) -> std::intptr_t { if (w == 9) style[msg] = l; return 0; }, false);
        check("style: straight box, light fill, opaque thin outline under the text",
            style[SCI_INDICSETSTYLE] == INDIC_STRAIGHTBOX && style[SCI_INDICSETALPHA] > 0 && style[SCI_INDICSETALPHA] < 64 && style[SCI_INDICSETOUTLINEALPHA] == 255
            && style[SCI_INDICSETSTROKEWIDTH] == 100 && style[SCI_INDICSETUNDER] == 1);
    }

    const std::string four = "ab foo foo cd\nab foo foo cd\nab foo foo cd\nab foo foo cd";
    const std::vector<std::string> rows(4, "foo foo");

    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7));
      check("draw and read back the rows", texts(doc, f) == rows); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(2) + 6, "XY");
      check("typing in row 3 widens only that row", texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo", "fooXY foo", "foo foo" }); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(1) + 3, "Z");
      check("typing at the start of a row stays outside", texts(doc, f) == rows); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(0) + 10, "W");
      check("typing at the end of a row stays outside", texts(doc, f) == rows); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(2) + 6, "\n");
      check("a line break in a row adds no row, the row runs over it",
          texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo", "foo\n foo", "foo foo" } && doc.d->LinesTotal() == 5); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 1, 4);
      check("deleting across the start of a row keeps its rest", texts(doc, f) == std::vector<std::string>{ "foo foo", "o foo", "foo foo", "foo foo" }); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 7);
      check("deleting a whole row drops it", texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo", "foo foo" }); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 7); doc.d->Undo();
      check("plain Scintilla: undo brings the text back but not the dropped row",
          texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo", "foo foo" } && doc.d->Length() == static_cast<Sci::Position>(four.size())); }
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(2) + 6, "XY"); doc.d->Undo();
      check("undo of typing restores the row", texts(doc, f) == rows); }

    // MR's Replace: Scintilla deletes and inserts, the panel draws its adjusted copy again
    { Doc doc(four); std::vector<Span> copy = block(doc, 0, 4, 3, 7); f.draw(doc.send(), copy);
      doc.replace(doc.line(1) + 3, 3, "QUUXX");
      check("plain Scintilla: a replacement at the start of a row falls out", texts(doc, f) == std::vector<std::string>{ "foo foo", " foo", "foo foo", "foo foo" });
      copy[1].end += 2; for (size_t i = 2; i < copy.size(); ++i) { copy[i].start += 2; copy[i].end += 2; }
      f.draw(doc.send(), copy);
      check("drawn from the panel's copy the replacement belongs to its row", texts(doc, f) == std::vector<std::string>{ "foo foo", "QUUXX foo", "foo foo", "foo foo" }); }
    { Doc doc("ab foo cd\nab foo cd"); f.draw(doc.send(), { { 3, 6 }, { doc.line(1) + 3, doc.line(1) + 6 } });
      doc.replace(3, 3, "ZZ");
      check("plain Scintilla: a row that was only the hit is gone after a replacement", texts(doc, f) == std::vector<std::string>{ "foo" }); }

    // Undo, redo and typing as MR announces them; MR's own Replace is left out and drawn again as above
    { Doc doc(four); Announce a(doc, f); std::vector<Span> copy = block(doc, 0, 4, 3, 7); f.draw(doc.send(), copy);
      a.on = false; doc.replace(doc.line(1) + 3, 3, "QUUXX"); a.on = true;
      copy[1].end += 2; for (size_t i = 2; i < copy.size(); ++i) { copy[i].start += 2; copy[i].end += 2; }
      f.draw(doc.send(), copy);
      doc.d->Undo();
      check("undo of a Replace at the start of a row keeps the old text in the row", texts(doc, f) == rows);
      doc.d->Redo();
      check("redo keeps the replacement in the row", texts(doc, f) == std::vector<std::string>{ "foo foo", "QUUXX foo", "foo foo", "foo foo" }); }
    { Doc doc("ab foo cd\nab foo cd"); Announce a(doc, f); f.draw(doc.send(), { { 3, 6 }, { doc.line(1) + 3, doc.line(1) + 6 } });
      a.on = false; doc.replace(3, 3, "ZZ"); a.on = true;
      f.draw(doc.send(), { { 3, 5 }, { doc.line(1) + 3, doc.line(1) + 6 } });
      doc.d->Undo();
      check("undo brings back a row that was only the hit", texts(doc, f) == std::vector<std::string>{ "foo", "foo" });
      doc.d->Redo();
      check("redo keeps that row", texts(doc, f) == std::vector<std::string>{ "ZZ", "foo" }); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 7); doc.d->Undo();
      check("undo of deleting a whole row brings the row back", texts(doc, f) == rows); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.replace(doc.line(1) + 3, 3, "QUUXX");
      check("typing over the start of a row stays in the row", texts(doc, f) == std::vector<std::string>{ "foo foo", "QUUXX foo", "foo foo", "foo foo" }); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.insert(doc.line(1) + 3, "Z");
      check("announced typing at the start of a row still stays outside", texts(doc, f) == rows); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.replace(doc.line(1) + 1, 2, "XY");
      check("replacing the text just before a row stays outside", texts(doc, f) == rows); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.replace(doc.line(1) + 2, 2, "XY");
      check("replacing across the start of a row stays outside", texts(doc, f) == std::vector<std::string>{ "foo foo", "oo foo", "foo foo", "foo foo" }); }
    { Doc a(four), b(four); f.draw(a.send(), block(a, 0, 4, 3, 7));
      f.noteDelete(a.send(), a.line(1) + 3, 3); f.keepInsert(b.send(), b.line(1) + 3, 3);
      check("an insertion in another document is not framed", f.read(b.send()).empty()); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 3);
      a.on = false; doc.insert(doc.line(3), "XX"); a.on = true; doc.insert(doc.line(1) + 3, "foo");
      check("an edit not announced in between leaves the insertion outside", texts(doc, f)[1] == " foo"); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 7); f.erase(doc.send()); doc.d->Undo();
      check("after erase undo frames nothing again", f.read(doc.send()).empty()); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7)); doc.erase(doc.line(1) + 3, 7); f.retire(); doc.d->Undo();
      check("after retire undo frames nothing again", f.read(doc.send()).empty()); f.dropStale(doc.send()); }
    { Doc doc(four); Announce a(doc, f); f.draw(doc.send(), block(doc, 0, 4, 3, 7));
      doc.d->DecorationSetCurrentIndicator(12); doc.d->decorations->SetCurrentValue(5);
      doc.replace(doc.line(1) + 3, 3, "QUUXX");
      check("keeping a replacement framed leaves the current indicator and value as found",
          doc.d->decorations->GetCurrentIndicator() == 12 && doc.d->decorations->GetCurrentValue() == 5 && texts(doc, f)[1] == "QUUXX foo"); }

    // Generations: a retired frame no longer counts and each document drops it
    { Doc doc(four); f.draw(doc.send(), block(doc, 0, 2, 3, 7)); f.retire();
      check("a retired frame is not read", f.read(doc.send()).empty());
      f.dropStale(doc.send());
      check("dropStale removes the retired frame", doc.d->decorations->ValueAt(9, doc.line(0) + 4) == 0 && doc.d->decorations->ValueAt(9, doc.line(1) + 4) == 0);
      f.draw(doc.send(), block(doc, 2, 2, 3, 7));
      check("a frame drawn after retire reads back", texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo" });
      f.dropStale(doc.send());
      check("dropStale keeps the current frame", texts(doc, f) == std::vector<std::string>{ "foo foo", "foo foo" });
      f.erase(doc.send());
      check("erase removes the frame", f.read(doc.send()).empty()); }

    // Frames belong to their document
    { Doc a(four), b("other text"); f.draw(a.send(), block(a, 0, 4, 3, 7));
      check("another document holds no frame", f.read(b.send()).empty());
      f.dropStale(b.send());
      check("the first document keeps its frame", texts(a, f) == rows); }

    { Doc doc("plain"); f.dropStale(doc.send());
      check("a document without the indicator reads empty", f.read(doc.send()).empty()); }
    { Doc doc(""); check("an empty document reads empty", f.read(doc.send()).empty()); }

    // Another indicator current: draw and erase choose their own
    { Doc doc(four); doc.d->DecorationSetCurrentIndicator(12); f.draw(doc.send(), block(doc, 0, 1, 3, 7));
      doc.d->DecorationSetCurrentIndicator(12); f.erase(doc.send());
      check("draw and erase set their indicator themselves", f.read(doc.send()).empty()); }

    // Notepad++ fills with the document's current indicator and value: the frame leaves both as found
    { Doc doc(four); doc.d->DecorationSetCurrentIndicator(12); doc.d->decorations->SetCurrentValue(5);
      const auto asFound = [&] { return doc.d->decorations->GetCurrentIndicator() == 12 && doc.d->decorations->GetCurrentValue() == 5; };
      f.draw(doc.send(), block(doc, 0, 2, 3, 7)); const bool afterDraw = asFound() && texts(doc, f).size() == 2;
      f.retire(); f.dropStale(doc.send()); const bool afterDrop = asFound() && doc.d->decorations->ValueAt(9, doc.line(0) + 4) == 0;
      f.draw(doc.send(), block(doc, 2, 2, 3, 7)); f.erase(doc.send());
      check("current indicator and value stay as found", afterDraw && afterDrop && asFound() && f.read(doc.send()).empty()); }

    // A block over 20,000 rows
    { std::string big; for (int i = 0; i < 20000; ++i) big += "ab foo foo cd\n";
      Doc doc(big);
      const auto t0 = std::chrono::steady_clock::now();
      f.draw(doc.send(), block(doc, 0, 20000, 3, 7));
      const size_t n = f.read(doc.send()).size();
      const auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - t0).count();
      check("20,000 rows draw and read back", n == 20000);
      std::printf("20,000 rows: draw + read %lld ms\n", static_cast<long long>(ms)); }

    std::printf("%s: %d checks, %d failed\n", fails ? "FAIL" : "PASS", checks, fails);
    return fails ? 1 : 0;
}
