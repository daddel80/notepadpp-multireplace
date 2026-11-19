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

// QA for ReplaceMarks on the real Scintilla Document: replaced text keeps a mark with what it replaced, marks
// merge when their text is replaced again, text inserted otherwise loses the mark, marks belong to their document,
// deleted text leaves a mark on the next character, the navigation steps through all and revert puts them back;
// undo and redo bring the marks back with their text, also those of deleted text.
// VERBATIM: the indicator messages as Editor::WndProc hands them to the Document.
// Build: Scintilla object files as for selection_frame_qa, Document.cxx without SCI_OWNREGEX (no Boost), then
//   g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src -I.. replace_marks_qa.cpp ../ReplaceMarks.cpp ../EditorPalette.cpp *.o -o replace_marks_qa
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
#include "../ReplaceMarks.h"
using namespace Scintilla; using namespace Scintilla::Internal;
static constexpr int kCpUtf8 = 65001;
static constexpr int kId = 9;
static constexpr int kCutId = 10;

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
    std::pair<Sci::Position, Sci::Position> target{ 0, 0 };
    int undoSteps = 0;   // the marks' own
    explicit Doc(const std::string& s) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(kCpUtf8);
        d->InsertString(0, s.c_str(), static_cast<Sci::Position>(s.size()));
        d->DeleteUndoHistory();
    }
    ~Doc() { d->Release(); }
    Doc(const Doc&) = delete;
    Doc& operator=(const Doc&) = delete;

    // The messages ReplaceMarks sends, as Editor::WndProc passes them on
    ReplaceMarks::Send send() {
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
            case SCI_INDICATORSTART: return d->decorations->Start(static_cast<int>(w), l);
            case SCI_INDICATOREND: return d->decorations->End(static_cast<int>(w), l);
            case SCI_GETDOCPOINTER: return reinterpret_cast<std::intptr_t>(d);
            case SCI_GETRANGEPOINTER: return reinterpret_cast<std::intptr_t>(d->RangePointer(static_cast<Sci::Position>(w), l));
            case SCI_BEGINUNDOACTION: d->BeginUndoAction(); return 0;
            case SCI_GETUNDOCOLLECTION: return d->IsCollectingUndo() ? 1 : 0;
            case SCI_ADDUNDOACTION: d->AddUndoAction(static_cast<Sci::Position>(w), (l & UNDO_MAY_COALESCE) != 0); ++undoSteps; return 0;
            case SCI_ENDUNDOACTION: d->EndUndoAction(); return 0;
            case SCI_SETTARGETRANGE: target = { static_cast<Sci::Position>(w), l }; return 0;
            case SCI_REPLACETARGET: replace(target.first, target.second - target.first, std::string(reinterpret_cast<const char*>(l), w)); return static_cast<std::intptr_t>(w);
            default: std::fprintf(stderr, "unexpected message %u\n", msg); std::abort();
            }
        };
    }
    std::string all() const {
        std::string out(static_cast<size_t>(d->Length()), '\0');
        d->GetCharRange(out.data(), 0, d->Length());
        return out;
    }
    std::string text(Sci::Position pos, Sci::Position len) const {
        std::string out(static_cast<size_t>(len), '\0');
        d->GetCharRange(out.data(), pos, len);
        return out;
    }
    void insert(Sci::Position pos, const std::string& s) { d->InsertString(pos, s.c_str(), static_cast<Sci::Position>(s.size())); }
    void erase(Sci::Position pos, Sci::Position len) { d->DeleteChars(pos, len); }
    // Editor::ReplaceTarget without patterns: delete, then insert
    void replace(Sci::Position pos, Sci::Position len, const std::string& s) { d->BeginUndoAction(); erase(pos, len); insert(pos, s); d->EndUndoAction(); }
};

// Changes reported the way MR passes SCN_MODIFIED on (the token as int); off during MR's own replacement
struct Watch : DocWatcher {
    Doc& doc;
    ReplaceMarks::Marks& marks;
    bool on = true;
    Watch(Doc& d, ReplaceMarks::Marks& m) : doc(d), marks(m) { doc.d->AddWatcher(this, nullptr); }
    ~Watch() override { doc.d->RemoveWatcher(this, nullptr); }
    Watch(const Watch&) = delete;
    Watch& operator=(const Watch&) = delete;
    void NotifyModified(Document*, DocModification mh, void*) override {
        if (!on) return;
        const int type = static_cast<int>(mh.modificationType);
        const ReplaceMarks::Step step = (type & SC_PERFORMED_UNDO) ? ReplaceMarks::Step::Undo
            : (type & SC_PERFORMED_REDO) ? ReplaceMarks::Step::Redo : ReplaceMarks::Step::Edit;
        if (type & SC_MOD_CONTAINER) marks.replayed(doc.send(), static_cast<int>(mh.token), step == ReplaceMarks::Step::Undo);
        else if (type & SC_MOD_INSERTTEXT) marks.inserted(doc.send(), mh.position, mh.length, step);
        else if (type & SC_MOD_BEFOREDELETE) marks.deleting(doc.send(), mh.position, mh.length, step);
        else if (type & SC_MOD_DELETETEXT) marks.deleted(doc.send(), mh.position, mh.length, step);
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

// MR's performReplace with marks: one undo action
static bool mrReplace(ReplaceMarks::Marks& m, Doc& doc, Watch& w, Sci::Position pos, Sci::Position len, const std::string& s) {
    doc.d->BeginUndoAction();
    m.beforeReplace(doc.send(), pos, len);
    w.on = false;
    doc.replace(pos, len, s);
    w.on = true;
    const bool first = m.afterReplace(doc.send(), static_cast<std::intptr_t>(s.size()));
    doc.d->EndUndoAction();
    return first;
}

using Run = std::pair<std::string, std::string>;   // marked text, what it replaced

static std::vector<Run> runs(Doc& doc, const ReplaceMarks::Marks& m) {
    std::vector<Run> v;
    const auto send = doc.send();
    const std::intptr_t length = doc.d->Length();
    for (std::intptr_t pos = 0; pos < length;) {
        const std::intptr_t end = send(SCI_INDICATOREND, kId, pos);
        if (end <= pos) break;
        if (send(SCI_INDICATORVALUEAT, kId, pos)) {
            const ReplaceMarks::Span s = m.at(send, pos);
            v.push_back({ doc.text(pos, end - pos), s.start == pos && s.end == end ? std::string(m.originalOf(send, s)) : std::string("?") });
        }
        pos = end;
    }
    return v;
}

static std::intptr_t valueAt(Doc& doc, Sci::Position pos) { return doc.send()(SCI_INDICATORVALUEAT, kId, pos); }

using Cut = std::pair<std::intptr_t, std::string>;   // where text was deleted, what it was

static std::vector<Cut> cuts(Doc& doc, const ReplaceMarks::Marks& m) {
    std::vector<Cut> v;
    const auto send = doc.send();
    const std::intptr_t length = doc.d->Length();
    for (std::intptr_t pos = 0; pos < length;) {
        const std::intptr_t end = send(SCI_INDICATOREND, kCutId, pos);
        if (end <= pos) break;
        if (send(SCI_INDICATORVALUEAT, kCutId, pos))
            v.push_back({ pos, end == pos + 1 && m.deletionAt(send, pos).found() ? std::string(m.originalOf(send, { pos, pos })) : std::string("?") });
        pos = end;
    }
    return v;
}

using Place = std::array<std::intptr_t, 4>;   // start, end, index, count

static Place step(Doc& doc, const ReplaceMarks::Marks& m, std::intptr_t from, std::intptr_t to, bool forward) {
    const ReplaceMarks::Stop s = m.step(doc.send(), from, to, forward);
    return { s.span.start, s.span.end, static_cast<std::intptr_t>(s.index), static_cast<std::intptr_t>(s.count) };
}

static std::vector<std::pair<std::intptr_t, std::intptr_t>> within(Doc& doc, const ReplaceMarks::Marks& m, std::intptr_t from, std::intptr_t to) {
    std::vector<std::pair<std::intptr_t, std::intptr_t>> v;
    for (const ReplaceMarks::Span& s : m.within(doc.send(), from, to)) v.push_back({ s.start, s.end });
    return v;
}

static std::pair<std::intptr_t, std::intptr_t> revert(Doc& doc, ReplaceMarks::Marks& m, std::intptr_t from, std::intptr_t to) {
    const ReplaceMarks::Span s = m.revert(doc.send(), from, to);
    return { s.start, s.end };
}

static bool drawn(Doc& doc) {
    const auto send = doc.send();
    for (const int id : { kId, kCutId }) {
        const std::intptr_t end = send(SCI_INDICATOREND, id, 0);
        if (send(SCI_INDICATORVALUEAT, id, 0) != 0 || (end > 0 && end < doc.d->Length())) return true;
    }
    return false;
}

int main() {
    ReplaceMarks::Marks m;
    m.setIndicators(kId, kCutId);

    {
        // The editor theme decides, not Notepad++'s dark mode: the default themes, a black theme as in #139
        // and MossyLawn, where the light grey stands out too little on the background
        using Styles = std::map<unsigned int, std::intptr_t>;
        Styles light, dark, black, mossy, cutLight, cutDark, cutBlack, cutMossy;
        const auto styled = [&](Styles& marks, Styles& cuts, const EditorPalette::Theme& theme) {
            m.applyStyle([pm = &marks, pc = &cuts](unsigned int msg, std::uintptr_t w, std::intptr_t l) -> std::intptr_t {
                (w == kId ? *pm : *pc)[msg] = l; return 0; }, theme);
        };
        const EditorPalette::Theme mossyLawn = EditorPalette::makeTheme(0x3D6958, 0x76C4F2, 0x397A71);
        styled(light, cutLight, EditorPalette::makeTheme(0xFFFFFF, 0x000000, 0xFFE8E8));
        styled(dark, cutDark, EditorPalette::makeTheme(0x3F3F3F, 0xCCDCDC, 0x101010));
        styled(black, cutBlack, EditorPalette::makeTheme(0x000000, 0xFFFFFF, 0x333333));
        styled(mossy, cutMossy, mossyLawn);
        check("style: grey dotted underline, darker on light, lighter on dark",
            light[SCI_INDICSETSTYLE] == INDIC_DOTS && dark[SCI_INDICSETSTYLE] == INDIC_DOTS && light[SCI_INDICSETSTROKEWIDTH] == 100
            && light[SCI_INDICSETFORE] == 0x505050 && dark[SCI_INDICSETFORE] == 0xB0B0B0);
        check("style: a black theme gets the light grey", black[SCI_INDICSETFORE] == 0xB0B0B0);
        const auto mossyDots = static_cast<EditorPalette::Color>(mossy[SCI_INDICSETFORE]);
        check("style: a grey too close to the theme is lightened until it stands out",
            mossyDots != 0xB0B0B0 && EditorPalette::contrast(mossyDots, mossyLawn.background) >= EditorPalette::MARK_CONTRAST
            && EditorPalette::contrast(mossyDots, mossyLawn.caretLine) >= EditorPalette::MARK_CONTRAST);
        check("style: deletions a triangle in the same grey",
            cutLight[SCI_INDICSETSTYLE] == INDIC_POINT && cutDark[SCI_INDICSETSTYLE] == INDIC_POINT
            && cutLight[SCI_INDICSETFORE] == 0x505050 && cutDark[SCI_INDICSETFORE] == 0xB0B0B0
            && cutBlack[SCI_INDICSETFORE] == 0xB0B0B0 && cutMossy[SCI_INDICSETFORE] == mossy[SCI_INDICSETFORE]);
    }

    { Doc doc("one two three"); Watch w(doc, m);
      const bool first = mrReplace(m, doc, w, 4, 3, "2");
      const bool second = mrReplace(m, doc, w, 0, 3, "1");
      check("a replacement is marked with what it replaced", runs(doc, m) == std::vector<Run>{ { "1", "one" }, { "2", "two" } } && doc.all() == "1 2 three");
      check("afterReplace reports only the first mark of a document", first && !second);
      check("text around a mark stays unmarked", !m.at(doc.send(), 1).found() && !m.at(doc.send(), 3).found());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a a a"); Watch w(doc, m);
      for (Sci::Position pos : { 0, 3, 6 }) mrReplace(m, doc, w, pos, 1, "bb");
      check("Replace All from left to right: one mark per replacement",
          runs(doc, m) == std::vector<Run>{ { "bb", "a" }, { "bb", "a" }, { "bb", "a" } } && doc.all() == "bb bb bb");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 1, 1, "Y");
      check("adjacent replacements keep their own marks", runs(doc, m) == std::vector<Run>{ { "X", "a" }, { "Y", "b" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 2, "Q");
      check("replacing marked text with its neighbour merges: the mark shows the text before both",
          runs(doc, m) == std::vector<Run>{ { "Qb", "xa" } } && doc.all() == "Qb y");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 4, "Q");
      check("a replacement around a whole mark takes it in", runs(doc, m) == std::vector<Run>{ { "Q", "xa " } } && doc.all() == "Qy");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb");
      const std::intptr_t before = valueAt(doc, 1);
      mrReplace(m, doc, w, 2, 1, "Z");
      check("replacing inside a mark keeps it and its record",
          runs(doc, m) == std::vector<Run>{ { "bZb", "a" } } && valueAt(doc, 1) == before && valueAt(doc, 3) == before);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("aXb"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "A1"); mrReplace(m, doc, w, 3, 1, "B1");
      mrReplace(m, doc, w, 1, 3, "-");
      check("a replacement across two marks and the text between them",
          runs(doc, m) == std::vector<Run>{ { "A-1", "aXb" } } && doc.all() == "A-1");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m);
      const bool marked = mrReplace(m, doc, w, 4, 4, "");
      check("a deletion leaves a mark on the next character", runs(doc, m).empty() && marked && doc.all() == "one three"
          && cuts(doc, m) == std::vector<Cut>{ { 4, "two " } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); mrReplace(m, doc, w, 2, 1, "");
      check("deleting inside a mark keeps the rest marked", runs(doc, m) == std::vector<Run>{ { "bb", "a" } });
      mrReplace(m, doc, w, 1, 2, "");
      check("deleting a whole mark leaves a deletion with its original", runs(doc, m).empty() && doc.all() == "--"
          && cuts(doc, m) == std::vector<Cut>{ { 1, "a" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 0, "X");
      const ReplaceMarks::Span s = m.at(doc.send(), 1);
      check("an insertion is marked with an empty original", runs(doc, m) == std::vector<Run>{ { "X", "" } } && s.start == 1 && m.originalOf(doc.send(), s).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); mrReplace(m, doc, w, 2, 0, "Z");
      check("an insertion inside a mark joins it", runs(doc, m) == std::vector<Run>{ { "bZbb", "a" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 1, 0, "S"); mrReplace(m, doc, w, 4, 0, "E");
      check("insertions at the edges of a mark get their own", runs(doc, m) == std::vector<Run>{ { "S", "" }, { "bb", "a" }, { "E", "" } } && doc.all() == "-SbbE-");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); doc.insert(2, "Z");
      check("typed text inside a mark joins it", runs(doc, m) == std::vector<Run>{ { "bZbb", "a" } } && doc.all() == "-bZbb-");
      doc.insert(5, "E"); doc.insert(1, "S");
      check("typed text at the edges stays outside", runs(doc, m) == std::vector<Run>{ { "bZbb", "a" } } && doc.all() == "-SbZbbE-");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 1, 1, "Y"); doc.insert(1, "z");
      check("typed text between two marks joins neither", runs(doc, m) == std::vector<Run>{ { "X", "a" }, { "Y", "b" } } && doc.all() == "XzY");
      doc.insert(3, "e");
      check("typed text after a mark at the document end stays outside", runs(doc, m) == std::vector<Run>{ { "X", "a" }, { "Y", "b" } } && doc.all() == "XzYe");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); doc.erase(2, 1); doc.d->Undo();
      check("undo of a deletion inside a mark brings the text back into it", runs(doc, m) == std::vector<Run>{ { "bbb", "a" } } && doc.all() == "-bbb-");
      mrReplace(m, doc, w, 2, 1, "Z"); doc.d->Undo();
      check("undo of a replacement inside a mark leaves the mark as it was", runs(doc, m) == std::vector<Run>{ { "bbb", "a" } } && doc.all() == "-bbb-");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); mrReplace(m, doc, w, 0, 2, "Q");   // "Qbb y": merged, takes "xa"
      doc.erase(1, 1); doc.d->Undo();
      check("undo of a deletion inside a merged mark brings the text back into it", runs(doc, m) == std::vector<Run>{ { "Qbb", "xa" } } && doc.all() == "Qbb y");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 3, "2"); doc.d->Undo();
      check("undo of a replacement takes its mark", runs(doc, m).empty() && doc.all() == "one two three");
      doc.d->Redo();
      check("redo brings the replacement back with its mark", runs(doc, m) == std::vector<Run>{ { "2", "two" } } && doc.all() == "one 2 three");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y aXb"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 2, "Q");                      // "Qb y aXb"
      mrReplace(m, doc, w, 5, 1, "A1"); mrReplace(m, doc, w, 8, 1, "B1"); mrReplace(m, doc, w, 6, 3, "-");
      const bool merged = runs(doc, m) == std::vector<Run>{ { "Qb", "xa" }, { "A-1", "aXb" } };
      while (doc.d->CanUndo()) doc.d->Undo();
      check("after undoing every replacement no mark is left", merged && runs(doc, m).empty() && doc.all() == "xa y aXb");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a b"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 2, 1, "Y"); mrReplace(m, doc, w, 0, 1, "ZZ"); mrReplace(m, doc, w, 4, 0, "E");
      check("marks at the start and the end of the document", runs(doc, m) == std::vector<Run>{ { "ZZ", "a" }, { "Y", "b" }, { "E", "" } } && doc.all() == "ZZ YE");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y aXb"); Watch w(doc, m);
      doc.d->BeginUndoAction();
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 2, "Q");
      mrReplace(m, doc, w, 5, 1, "A1"); mrReplace(m, doc, w, 8, 1, "B1"); mrReplace(m, doc, w, 6, 3, "-");
      doc.d->EndUndoAction();
      doc.d->Undo();
      check("one undo of a whole run leaves no mark", runs(doc, m).empty() && doc.all() == "xa y aXb" && !doc.d->CanUndo());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 3, "2"); m.clear(doc.send());
      check("clear removes the marks of the document", runs(doc, m).empty() && !drawn(doc) && !m.at(doc.send(), 4).found());
      check("the next replacement is the first again", mrReplace(m, doc, w, 0, 3, "1") && runs(doc, m) == std::vector<Run>{ { "1", "one" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc a("one two"); Doc b("one two"); Watch wa(a, m); Watch wb(b, m);
      mrReplace(m, a, wa, 4, 3, "2"); mrReplace(m, b, wb, 0, 3, "1");
      check("marks belong to their document", runs(a, m) == std::vector<Run>{ { "2", "two" } } && runs(b, m) == std::vector<Run>{ { "1", "one" } }
          && valueAt(a, 4) == 1 && valueAt(b, 0) == 1);
      m.retire();
      check("retired marks show no original but stay drawn until dropped", m.empty() && !m.at(a.send(), 4).found() && drawn(a) && drawn(b));
      m.dropStale(a.send());
      check("dropStale clears one document", !drawn(a) && drawn(b));
      mrReplace(m, b, wb, 2, 3, "2");
      check("a replacement in a stale document drops the old marks first", runs(b, m) == std::vector<Run>{ { "2", "two" } } && valueAt(b, 0) == 0 && valueAt(b, 2) == 1);
      m.dropStale(a.send());
      check("a dropped document stays as it is", !drawn(a));
      m.forget(reinterpret_cast<std::intptr_t>(b.d));
      check("a forgotten document shows no original", m.empty() && !m.at(b.send(), 2).found()); }

    { Doc doc("one"); Watch w(doc, m);
      check("a range past the end marks nothing", !mrReplace(m, doc, w, 2, 0, "") && runs(doc, m).empty());
      m.beforeReplace(doc.send(), 1, 5);
      check("a range past the end is not read", !m.afterReplace(doc.send(), 1) && runs(doc, m).empty()); }

    { Doc doc("one two"); Watch w(doc, m);
      const auto send = doc.send();
      send(SCI_SETINDICATORCURRENT, 5, 0); send(SCI_SETINDICATORVALUE, 7, 0);
      mrReplace(m, doc, w, 4, 3, "2"); doc.insert(4, "x"); m.clear(send);
      check("the document's current indicator and value stay as found", send(SCI_GETINDICATORCURRENT, 0, 0) == 5 && send(SCI_GETINDICATORVALUE, 0, 0) == 7);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("gr\xC3\xB6\xC3\x9F" "e"); Watch w(doc, m);
      mrReplace(m, doc, w, 2, 4, "oess");
      check("multi-byte originals come back byte for byte", runs(doc, m) == std::vector<Run>{ { "oess", "\xC3\xB6\xC3\x9F" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc a("one two"); Doc b("one two"); Watch wa(a, m); Watch wb(b, m);
      m.beforeReplace(a.send(), 4, 3);
      wb.on = false; b.replace(4, 3, "2"); wb.on = true;
      const bool marked = m.afterReplace(b.send(), 1);
      check("afterReplace in another document marks nothing", !marked && runs(a, m).empty() && runs(b, m).empty());
      m.beforeReplace(a.send(), 0, 3);
      mrReplace(m, a, wa, 4, 3, "2");
      check("an unfinished replacement leaves no trace", runs(a, m) == std::vector<Run>{ { "2", "two" } });
      m.forget(reinterpret_cast<std::intptr_t>(a.d)); }

    // Deletions
    { Doc doc("ab  \r\ncd"); Watch w(doc, m);
      mrReplace(m, doc, w, 2, 2, "");
      check("a deletion at the line end sits on the line break", cuts(doc, m) == std::vector<Cut>{ { 2, "  " } } && doc.all() == "ab\r\ncd");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab  "); Watch w(doc, m);
      check("a deletion at the document end has no character to sit on", !mrReplace(m, doc, w, 2, 2, "") && cuts(doc, m).empty() && doc.all() == "ab");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 1, 1, "B");
      check("replacing the character after a deletion takes the deletion in", runs(doc, m) == std::vector<Run>{ { "B", ",b" } } && cuts(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 1, 0, "-");
      check("an insertion at a deletion takes it in", runs(doc, m) == std::vector<Run>{ { "-", "," } } && cuts(doc, m).empty() && doc.all() == "a-b");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 0, 1, "A");
      check("a replacement right before a deletion leaves it alone",
          runs(doc, m) == std::vector<Run>{ { "A", "a" } } && cuts(doc, m) == std::vector<Cut>{ { 1, "," } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("XYZ!"); Watch w(doc, m);
      mrReplace(m, doc, w, 2, 1, ""); mrReplace(m, doc, w, 1, 1, "");
      check("two deletions at one place become one", cuts(doc, m) == std::vector<Cut>{ { 1, "YZ" } } && doc.all() == "X!");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b;c"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 2, 1, "");   // "abc"
      doc.insert(1, "-");
      check("typed text at a deletion goes before it", cuts(doc, m) == std::vector<Cut>{ { 2, "," }, { 3, ";" } } && doc.all() == "a-bc");
      doc.insert(3, "x");
      check("typed text between two deletions takes neither",
          cuts(doc, m) == std::vector<Cut>{ { 2, "," }, { 4, ";" } } && doc.all() == "a-bxc" && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 4, ""); doc.d->Undo();
      check("undo of a deletion takes its mark", cuts(doc, m).empty() && doc.all() == "one two three");
      doc.d->Redo();
      check("redo deletes again with its mark", cuts(doc, m) == std::vector<Cut>{ { 4, "two " } } && doc.all() == "one three");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b,c"); Watch w(doc, m);
      doc.d->BeginUndoAction();
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 2, 1, "");
      doc.d->EndUndoAction();
      const bool both = cuts(doc, m) == std::vector<Cut>{ { 1, "," }, { 2, "," } };
      doc.d->Undo();
      check("one undo of a run takes all its deletions", both && cuts(doc, m).empty() && doc.all() == "a,b,c");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); doc.erase(1, 1);
      check("deleting the character after a deletion takes the mark", cuts(doc, m).empty() && doc.all() == "a");
      doc.d->Undo();
      check("undo of that brings the deletion back", cuts(doc, m) == std::vector<Cut>{ { 1, "," } } && doc.all() == "ab");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b;c"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); mrReplace(m, doc, w, 2, 1, ""); mrReplace(m, doc, w, 0, 1, "A");   // "Abc"
      const auto send = doc.send();
      check("deletions are found at their boundary", m.deletionAt(send, 1).found() && m.deletionAt(send, 2).found()
          && !m.deletionAt(send, 0).found() && !m.at(send, 1).found() && m.originalOf(send, { 1, 1 }) == "," && m.originalOf(send, { 2, 2 }) == ";");
      check("at a caret: the deletion there, the mark after it, the one before",
          m.atCaret(send, 1).start == 1 && m.atCaret(send, 1).end == 1 && m.atCaret(send, 0).end == 1 && !m.atCaret(send, 3).found());
      check("deletions take their place in the navigation",
          step(doc, m, 0, 0, true) == Place{ 0, 1, 1, 3 } && step(doc, m, 0, 1, true) == Place{ 1, 1, 2, 3 }
          && step(doc, m, 1, 1, true) == Place{ 2, 2, 3, 3 } && step(doc, m, 2, 2, true) == Place{ 0, 1, 1, 3 }
          && step(doc, m, 2, 2, false) == Place{ 1, 1, 2, 3 } && step(doc, m, 1, 1, false) == Place{ 0, 1, 1, 3 });
      check("a selection takes the deletions strictly inside it",
          within(doc, m, 0, 3).size() == 3 && within(doc, m, 0, 1).size() == 1 && within(doc, m, 1, 2).empty());
      check("revert of a deletion puts its text back",
          m.revert(send, ReplaceMarks::Span{ 1, 1 }).end == 2 && doc.all() == "A,bc" && cuts(doc, m) == std::vector<Cut>{ { 3, ";" } });
      check("revert of a selection with deletions", revert(doc, m, 0, 4) == std::pair<std::intptr_t, std::intptr_t>{ 0, 5 } && doc.all() == "a,b;c"
          && runs(doc, m).empty() && cuts(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 2, 1, "B"); mrReplace(m, doc, w, 1, 1, "");   // "aB": the deletion sits on the mark's first character
      const auto send = doc.send();
      const bool shared = runs(doc, m) == std::vector<Run>{ { "B", "b" } } && cuts(doc, m) == std::vector<Cut>{ { 1, "," } };
      m.revert(send, ReplaceMarks::Span{ 1, 2 });
      check("revert of a mark keeps the deletion before it", shared && doc.all() == "ab" && cuts(doc, m) == std::vector<Cut>{ { 1, "," } });
      m.revert(send, ReplaceMarks::Span{ 1, 1 });
      check("then the deletion goes back in front of it", doc.all() == "a,b" && cuts(doc, m).empty() && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 2, 1, "B"); mrReplace(m, doc, w, 1, 1, "");
      revert(doc, m, 0, 2);
      check("revert of a mark and the deletion before it", doc.all() == "a,b" && cuts(doc, m).empty() && runs(doc, m).empty());
      check("a reverted mark is no longer there to revert", !m.revert(doc.send(), ReplaceMarks::Span{ 1, 1 }).found());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    // Undo and redo
    { Doc doc("xa y aXb"); Watch w(doc, m);
      doc.d->BeginUndoAction();
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 2, "Q");
      mrReplace(m, doc, w, 5, 1, "A1"); mrReplace(m, doc, w, 8, 1, "B1"); mrReplace(m, doc, w, 6, 3, "-");
      doc.d->EndUndoAction();
      const std::vector<Run> done = runs(doc, m);
      doc.d->Undo(); doc.d->Redo();
      check("redo of a run brings every mark back, merged ones as they were",
          done == std::vector<Run>{ { "Qb", "xa" }, { "A-1", "aXb" } } && runs(doc, m) == done && doc.all() == "Qb y A-1" && cuts(doc, m).empty());
      doc.d->Undo(); doc.d->Redo(); doc.d->Undo();
      check("undo and redo again and again", runs(doc, m).empty() && doc.all() == "xa y aXb" && !doc.d->CanUndo());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("aXb"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "A1"); mrReplace(m, doc, w, 3, 1, "B1"); mrReplace(m, doc, w, 1, 3, "-");
      doc.d->Undo();
      check("undo of a merge brings back the marks it took in", runs(doc, m) == std::vector<Run>{ { "A1", "a" }, { "B1", "b" } } && doc.all() == "A1XB1");
      doc.d->Redo();
      check("redo merges them again", runs(doc, m) == std::vector<Run>{ { "A-1", "aXb" } } && doc.all() == "A-1");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 3, "2"); doc.erase(3, 3);
      check("deleted text takes its mark along", runs(doc, m).empty() && doc.all() == "onethree");
      doc.d->Undo();
      check("undo of the deletion brings the mark back", runs(doc, m) == std::vector<Run>{ { "2", "two" } } && doc.all() == "one 2 three");
      doc.d->Redo();
      check("redo deletes it again", runs(doc, m).empty() && doc.all() == "onethree");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-colour-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 6, "color"); doc.erase(1, 2); doc.d->Undo();
      check("undo of a deletion at a mark's edge gives the mark its text back", runs(doc, m) == std::vector<Run>{ { "color", "colour" } } && doc.all() == "-color-");
      revert(doc, m, 1, 6);
      check("so revert puts the original back once", doc.all() == "-colour-");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-abcd-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 4, "wxyz");
      doc.d->SetSavePoint();
      const int steps = doc.undoSteps;
      doc.erase(4, 1); doc.erase(3, 1);   // two Backspaces
      const bool both = doc.undoSteps == steps + 2;
      doc.d->Undo();
      check("Backspaces in a mark still undo as one step, back to the save point",
          both && doc.all() == "-wxyz-" && runs(doc, m) == std::vector<Run>{ { "wxyz", "abcd" } } && doc.d->IsSavePoint());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m); Watch clone(doc, m);
      mrReplace(m, doc, w, 4, 3, "2");
      const int steps = doc.undoSteps;
      doc.erase(4, 1);
      const bool once = doc.undoSteps == steps + 1;
      doc.d->Undo();
      check("a document in two views: one undo step for a deletion, its mark back",
          once && runs(doc, m) == std::vector<Run>{ { "2", "two" } } && doc.all() == "one 2 three");
      doc.d->Undo();
      const bool none = runs(doc, m).empty() && doc.all() == "one two three";
      doc.d->Redo();
      const bool again = runs(doc, m) == std::vector<Run>{ { "2", "two" } } && doc.all() == "one 2 three";
      doc.d->Redo();
      check("a document in two views: undo and redo of the replacement", none && again && runs(doc, m).empty() && doc.all() == "one  three");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a,b"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, ""); doc.insert(1, "-"); doc.d->Undo(); doc.d->Redo();
      check("redo of typed text at a deletion keeps the deletion", cuts(doc, m) == std::vector<Cut>{ { 2, "," } } && doc.all() == "a-b");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one three"); Watch w(doc, m);
      mrReplace(m, doc, w, 3, 0, " two");
      m.revert(doc.send(), ReplaceMarks::Span{ 3, 7 });
      const bool reverted = doc.all() == "one three" && runs(doc, m).empty();
      doc.d->Undo();
      check("undo of the revert of an insertion: its mark is back", reverted && doc.all() == "one two three" && runs(doc, m) == std::vector<Run>{ { " two", "" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two three"); Watch w(doc, m);
      mrReplace(m, doc, w, 3, 4, "");
      m.revert(doc.send(), ReplaceMarks::Span{ 3, 3 });
      const bool back = doc.all() == "one two three" && cuts(doc, m).empty();
      doc.d->Undo();
      check("undo of a deletion's revert: the triangle is back", back && doc.all() == "one three" && cuts(doc, m) == std::vector<Cut>{ { 3, " two" } });
      doc.d->Redo();
      check("redo reverts it again", doc.all() == "one two three" && cuts(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 3, "2"); m.clear(doc.send());
      doc.d->Undo(); doc.d->Redo();
      check("cleared marks stay away through undo and redo", runs(doc, m).empty() && !drawn(doc) && doc.all() == "one 2");
      mrReplace(m, doc, w, 0, 3, "1");
      doc.d->Undo(); doc.d->Undo();
      const bool old = runs(doc, m).empty() && doc.all() == "one two";
      doc.d->Redo(); doc.d->Redo();
      check("after clear: the old step finds nothing, the new one its mark", old && runs(doc, m) == std::vector<Run>{ { "1", "one" } } && doc.all() == "1 2");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 3, "1"); doc.d->Undo(); m.retire(); doc.d->Redo();
      check("retired marks stay away through redo", runs(doc, m).empty() && doc.all() == "1");
      m.dropStale(doc.send());
      check("a token of Flow Tabs is not the marks'", !m.replayed(doc.send(), (0x4D52 << 16) | 1, true) && m.replayed(doc.send(), 0x52000000, true)); }

    { Doc doc("a b c"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 2, 1, "Y");
      doc.d->Undo(); doc.insert(5, "!"); mrReplace(m, doc, w, 4, 1, "Z");
      doc.d->Undo();
      const bool undone = doc.all() == "X b c!" && runs(doc, m) == std::vector<Run>{ { "X", "a" } };
      doc.d->Redo();
      check("a step the history dropped makes room for the next", undone && doc.all() == "X b Z!" && runs(doc, m) == std::vector<Run>{ { "X", "a" }, { "Z", "c" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab  "); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "A");
      const int steps = doc.undoSteps;
      mrReplace(m, doc, w, 2, 2, ""); mrReplace(m, doc, w, 1, 0, "");
      doc.erase(0, 0); doc.erase(1, 1);
      check("no undo step where no marks were or come", doc.undoSteps == steps && doc.all() == "A");
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two"); Watch w(doc, m);
      doc.d->SetUndoCollection(false);
      mrReplace(m, doc, w, 4, 3, "2");
      doc.d->SetUndoCollection(true);
      check("without undo collection the marks add no undo step", !doc.d->CanUndo() && doc.undoSteps == 0 && runs(doc, m) == std::vector<Run>{ { "2", "two" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    // The list entry of a replacement
    { Doc doc("colour colour"); Watch w(doc, m);
      const auto send = doc.send();
      m.setSource({ 2, 3, "colour" }); mrReplace(m, doc, w, 0, 6, "color");
      m.setSource({ 2, 0, "ignored" }); mrReplace(m, doc, w, 6, 6, "color");
      const ReplaceMarks::Source a = m.sourceOf(send, { 0, 5 });
      check("a mark names the list and the entry that made it",
          a == ReplaceMarks::Source{ 2, 3, "colour" } && m.sourceOf(send, { 6, 11 }) == ReplaceMarks::Source{});
      m.setSource({ 2, 5, "color" }); mrReplace(m, doc, w, 0, 5, "hue");
      const ReplaceMarks::Source c = m.sourceOf(send, { 0, 3 });
      check("a whole mark replaced again keeps its original, the entry is the new one",
          c.entry == 5 && c.find == "color" && m.originalOf(send, { 0, 3 }) == "colour" && doc.all() == "hue color");
      doc.d->Undo();
      check("undo brings back the mark with its own entry", m.sourceOf(send, { 0, 5 }).entry == 3 && doc.all() == "color color");
      m.setSource({ 2, 7, {} }); mrReplace(m, doc, w, 6, 5, "tint");
      const ReplaceMarks::Source plain = m.sourceOf(send, { 6, 10 });
      check("an entry without a find text keeps its id", plain == ReplaceMarks::Source{ 2, 7, {} } && doc.all() == "color tint");
      m.setSource({ 4, 7, {} }); mrReplace(m, doc, w, 0, 5, "tone");
      check("the same entry id in another list is another entry",
          m.sourceOf(send, { 0, 4 }) == ReplaceMarks::Source{ 4, 7, {} } && m.sourceOf(send, { 5, 9 }) == ReplaceMarks::Source{ 2, 7, {} }
          && doc.all() == "tone tint");
      m.setSource({});
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    // The dots grow with the line height
    { std::intptr_t height = 15, stroke = 100;
      int set = 0;
      const ReplaceMarks::Send view = [&](unsigned int msg, std::uintptr_t, std::intptr_t l) -> std::intptr_t {
          if (msg == SCI_TEXTHEIGHT) return height;
          if (msg == SCI_INDICGETSTROKEWIDTH) return stroke;
          if (msg == SCI_INDICSETSTROKEWIDTH) { stroke = l; ++set; }
          return 0;
      };
      m.applyStroke(view);
      const bool thin = stroke == 100 && set == 0;
      height = 24; m.applyStroke(view);
      const bool two = stroke == 200;
      height = 40; m.applyStroke(view);
      check("dots: one pixel per 16 of line height, set only when they change", thin && two && stroke == 300 && set == 2); }

    // Navigation
    { Doc doc("a b c d"); Watch w(doc, m);
      check("no marks: nothing to find", !m.any(doc.send()) && step(doc, m, 0, 0, true) == Place{ -1, -1, 0, 0 });
      for (Sci::Position pos : { 0, 2, 4, 6 }) mrReplace(m, doc, w, pos, 1, "X");
      check("a document with marks has some", m.any(doc.send()));
      check("next from the caret at a mark: that mark", step(doc, m, 0, 0, true) == Place{ 0, 1, 1, 4 });
      check("next from a selected mark: the following one", step(doc, m, 0, 1, true) == Place{ 2, 3, 2, 4 });
      check("next after the last mark: around to the first", step(doc, m, 6, 7, true) == Place{ 0, 1, 1, 4 });
      check("previous from a selected mark: the one before", step(doc, m, 4, 5, false) == Place{ 2, 3, 2, 4 });
      check("previous before the first mark: around to the last", step(doc, m, 0, 1, false) == Place{ 6, 7, 4, 4 });
      check("previous from the caret right after a mark: that mark", step(doc, m, 3, 3, false) == Place{ 2, 3, 2, 4 });
      check("a selection over marks: next from its end, previous up to its start",
          step(doc, m, 1, 5, true) == Place{ 6, 7, 4, 4 } && step(doc, m, 1, 5, false) == Place{ 0, 1, 1, 4 });
      check("the mark holding a character", m.at(doc.send(), 2).start == 2 && m.at(doc.send(), 2).end == 3
          && !m.at(doc.send(), 1).found() && !m.at(doc.send(), 3).found() && !m.at(doc.send(), 7).found() && !m.at(doc.send(), -1).found());
      check("the marks overlapping a range",
          within(doc, m, 1, 5) == std::vector<std::pair<std::intptr_t, std::intptr_t>>{ { 2, 3 }, { 4, 5 } } && within(doc, m, 0, 7).size() == 4
          && within(doc, m, 2, 3).size() == 1 && within(doc, m, 1, 2).empty() && within(doc, m, 2, 2).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab c"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 1, 1, "Y"); mrReplace(m, doc, w, 3, 1, "long");
      check("adjacent marks: the caret between them", step(doc, m, 1, 1, true) == Place{ 1, 2, 2, 3 } && step(doc, m, 1, 1, false) == Place{ 0, 1, 1, 3 });
      check("the caret inside a mark: next and previous pass it", step(doc, m, 5, 5, true) == Place{ 0, 1, 1, 3 } && step(doc, m, 5, 5, false) == Place{ 1, 2, 2, 3 });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 3, "1");
      check("a single mark: next and previous come back to it", step(doc, m, 0, 1, true) == Place{ 0, 1, 1, 1 } && step(doc, m, 0, 1, false) == Place{ 0, 1, 1, 1 });
      m.retire();
      check("retired marks are not found", !m.any(doc.send()) && step(doc, m, 0, 0, true) == Place{ -1, -1, 0, 0 } && !m.at(doc.send(), 0).found()
          && !m.revert(doc.send(), 0, 1).found() && doc.all() == "1");
      m.dropStale(doc.send()); }

    // Revert
    { Doc doc("one two three"); Watch w(doc, m);
      mrReplace(m, doc, w, 4, 3, "2");
      check("revert puts back what a mark replaced", revert(doc, m, 4, 5) == std::pair<std::intptr_t, std::intptr_t>{ 4, 7 }
          && doc.all() == "one two three" && runs(doc, m).empty() && !m.any(doc.send()));
      doc.d->Undo();
      check("one undo takes the revert back, with the mark", doc.all() == "one 2 three" && runs(doc, m) == std::vector<Run>{ { "2", "two" } });
      doc.d->Redo();
      check("redo reverts again", doc.all() == "one two three" && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a b c"); Watch w(doc, m);
      for (Sci::Position pos : { 4, 2, 0 }) mrReplace(m, doc, w, pos, 1, "long");
      check("revert of a selection: every mark in it, the range of the restored text",
          revert(doc, m, 0, 14) == std::pair<std::intptr_t, std::intptr_t>{ 0, 5 } && doc.all() == "a b c" && runs(doc, m).empty());
      doc.d->Undo();
      check("the marks of a selection come back in one undo", doc.all() == "long long long"
          && runs(doc, m) == std::vector<Run>{ { "long", "a" }, { "long", "b" }, { "long", "c" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("x a b c y"); Watch w(doc, m);
      for (Sci::Position pos : { 6, 4, 2 }) mrReplace(m, doc, w, pos, 1, "long");
      check("revert of a wider selection: it keeps the text around the marks",
          revert(doc, m, 0, 18) == std::pair<std::intptr_t, std::intptr_t>{ 0, 9 } && doc.all() == "x a b c y" && !m.revert(doc.send(), 0, 0).found());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("a b c"); Watch w(doc, m);
      for (Sci::Position pos : { 4, 2, 0 }) mrReplace(m, doc, w, pos, 1, "XY");
      check("revert of part of a selection's marks: the selection keeps its text", revert(doc, m, 4, 6) == std::pair<std::intptr_t, std::intptr_t>{ 3, 5 }
          && doc.all() == "XY b XY" && runs(doc, m) == std::vector<Run>{ { "XY", "a" }, { "XY", "c" } });
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 1, 1, "Y");
      revert(doc, m, 0, 1);
      check("revert next to a mark leaves it, the restored text stays unmarked", doc.all() == "aY" && runs(doc, m) == std::vector<Run>{ { "Y", "b" } } && valueAt(doc, 0) == 0);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 0, 1, "X"); mrReplace(m, doc, w, 1, 1, "Y");
      revert(doc, m, 1, 2);
      check("revert at the document end after a mark", doc.all() == "Xb" && runs(doc, m) == std::vector<Run>{ { "X", "a" } } && valueAt(doc, 1) == 0);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("xa y"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bb"); mrReplace(m, doc, w, 0, 2, "Q");
      revert(doc, m, 0, 2);
      check("revert of a merged mark: the text before both replacements", doc.all() == "xa y" && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("-a-"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "bbb"); doc.insert(2, "Z");
      revert(doc, m, 1, 5);
      check("revert of a mark with typed text in it", doc.all() == "-a-" && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("ab"); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 0, "X");
      check("revert of an insertion removes it", revert(doc, m, 1, 2) == std::pair<std::intptr_t, std::intptr_t>{ 1, 1 } && doc.all() == "ab" && runs(doc, m).empty());
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { const std::string text("a\0" "b gr\xC3\xB6\xC3\x9F" "e", 11);
      Doc doc(text); Watch w(doc, m);
      mrReplace(m, doc, w, 1, 1, "0"); mrReplace(m, doc, w, 6, 4, "oess");
      const bool replaced = doc.all() == "a0b groesse";
      revert(doc, m, 0, doc.d->Length());
      check("revert brings back NUL and multi-byte text byte for byte", replaced && doc.all() == text);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { Doc doc("one two"); Watch w(doc, m);
      const auto send = doc.send();
      mrReplace(m, doc, w, 4, 3, "2");
      send(SCI_SETINDICATORCURRENT, 5, 0); send(SCI_SETINDICATORVALUE, 7, 0);
      const bool none = !m.revert(send, 0, 3).found() && !m.revert(send, 4, 4).found() && doc.all() == "one 2";
      revert(doc, m, 4, 5);
      check("nothing to revert outside the marks", none);
      check("revert keeps the document's current indicator and value", send(SCI_GETINDICATORCURRENT, 0, 0) == 5 && send(SCI_GETINDICATORVALUE, 0, 0) == 7);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { std::string text; for (int i = 0; i < 20000; ++i) text += "ab ";
      Doc doc(text); Watch w(doc, m);
      const auto t0 = std::chrono::steady_clock::now();
      for (Sci::Position pos = 0, i = 0; i < 20000; ++i, pos += 5) mrReplace(m, doc, w, pos, 2, "abcd");
      const double s = std::chrono::duration<double>(std::chrono::steady_clock::now() - t0).count();
      const std::vector<Run> v = runs(doc, m);
      check("20000 replacements: 20000 marks", v.size() == 20000 && v.front() == Run{ "abcd", "ab" } && v.back() == Run{ "abcd", "ab" });
      std::printf("  20000 replacements with marks: %.3f s\n", s);
      const auto t1 = std::chrono::steady_clock::now();
      const Place middle = step(doc, m, 50000, 50000, true);
      const bool any = m.any(doc.send());
      const auto restored = revert(doc, m, middle[0], middle[1]);
      const double n = std::chrono::duration<double>(std::chrono::steady_clock::now() - t1).count();
      check("20000 marks: the place of a mark in the middle, its revert",
          middle == Place{ 50000, 50004, 10001, 20000 } && any && restored == std::pair<std::intptr_t, std::intptr_t>{ 50000, 50002 }
          && step(doc, m, 0, 0, false)[3] == 19999);
      std::printf("  step, any and revert among 20000 marks: %.3f s\n", n);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { std::string text; for (int i = 0; i < 20000; ++i) text += "ab ";
      Doc doc(text); Watch w(doc, m);
      const auto t0 = std::chrono::steady_clock::now();
      for (Sci::Position pos = 1, i = 0; i < 20000; ++i, pos += 2) mrReplace(m, doc, w, pos, 1, "");
      const Place middle = step(doc, m, 20001, 20001, true);
      const double s = std::chrono::duration<double>(std::chrono::steady_clock::now() - t0).count();
      check("20000 deletions: 20000 marks, the place of one in the middle", cuts(doc, m).size() == 20000 && middle == Place{ 20003, 20003, 10002, 20000 });
      std::printf("  20000 deletions with marks and a step: %.3f s\n", s);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    { std::string text; for (int i = 0; i < 20000; ++i) text += "ab ";
      Doc doc(text); Watch w(doc, m);
      doc.d->BeginUndoAction();
      for (Sci::Position pos = 0, i = 0; i < 20000; ++i, pos += 5) mrReplace(m, doc, w, pos, 2, "abcd");
      doc.d->EndUndoAction();
      const auto t0 = std::chrono::steady_clock::now();
      doc.d->Undo();
      const bool undone = runs(doc, m).empty() && doc.all() == text;
      const auto t1 = std::chrono::steady_clock::now();
      doc.d->Redo();
      const double u = std::chrono::duration<double>(t1 - t0).count();
      const double r = std::chrono::duration<double>(std::chrono::steady_clock::now() - t1).count();
      const std::vector<Run> v = runs(doc, m);
      check("20000 replacements in one run: undo takes all marks, redo brings all back",
          undone && v.size() == 20000 && v.front() == Run{ "abcd", "ab" } && v.back() == Run{ "abcd", "ab" });
      std::printf("  undo and redo of 20000 marked replacements: %.3f s, %.3f s\n", u, r);
      m.forget(reinterpret_cast<std::intptr_t>(doc.d)); }

    // The calltip text
    const auto tipText = [](std::string_view label, std::string_view empty, std::string_view space, std::string_view original,
        bool utf8, std::string_view place = {}, std::string_view entry = {}, std::string_view find = {}) {
        return ReplaceMarks::tipText(label, empty, space, original, utf8, place, entry, find).text;
    };
    check("tip: one line", tipText("Before:", "(empty)", "_", "two", true) == "Before: two");
    check("tip: empty original", tipText("Before:", "(empty)", "_", "", true) == "Before: (empty)");
    check("tip: spaces at the edges made visible", tipText("B:", "-", "_", "  a b  ", true) == "B: __a b__"
        && tipText("B:", "-", "\xC2\xB7", "   ", true) == "B: \xC2\xB7\xC2\xB7\xC2\xB7" && tipText("B:", "-", "_", "a \nb", true) == "B: a \\n\nb");
    check("tip: a visible space counts as one column",
        tipText("B:", "-", "\xC2\xB7", std::string(99, 'x') + "  ", true) == "B: " + std::string(99, 'x') + "\xC2\xB7" + "...");
    check("tip: after a jump the place leads between the arrows", tipText("Before:", "(empty)", "_", "two", true, "3 of 12") == "\0013 of 12\002 Before: two");
    check("tip: CRLF visible, then a new line", tipText("B:", "-", "_", "a\r\nb", true) == "B: a\\r\\n\nb");
    check("tip: no empty last line", tipText("B:", "-", "_", "a\nb\n", true) == "B: a\\n\nb\\n");
    check("tip: a lone CR breaks the line too", tipText("B:", "-", "_", "a\rb", true) == "B: a\\r\nb");
    check("tip: control characters visible, arrows never",
        tipText("B:", "-", "_", std::string("a\tb\0c\x01" "d\x02", 8), true) == "B: a\\tb\\0c\\x01d\\x02");
    {
        std::string twelve, ten;
        for (int i = 1; i <= 12; ++i) twelve += std::to_string(i) + "\n";
        for (int i = 1; i <= 10; ++i) ten += std::to_string(i) + "\\n" + (i < 10 ? "\n" : "");
        check("tip: at most 10 lines, then ...", tipText("B:", "-", "_", twelve, true) == "B: " + ten + "\n...");
    }
    check("tip: long lines are cut", tipText("B:", "-", "_", std::string(150, 'x'), true) == "B: " + std::string(100, 'x') + "...");
    check("tip: UTF-8 cut between characters",
        tipText("B:", "-", "_", std::string(99, 'x') + "\xC3\xBC" + "y", true) == "B: " + std::string(99, 'x') + "\xC3\xBC" + "...");
    check("tip: other code pages count bytes",
        tipText("B:", "-", "_", std::string(99, 'x') + "\xFC\xFD", false) == "B: " + std::string(99, 'x') + "\xFC" + "...");
    check("tip: a cut line ends, the next one shows",
        tipText("B:", "-", "_", std::string(120, 'x') + "\ny", true) == "B: " + std::string(100, 'x') + "...\ny");
    check("tip: line breaks show even at the limit",
        tipText("B:", "-", "_", std::string(100, 'x') + "\r\nz", true) == "B: " + std::string(100, 'x') + "\\r\\n\nz");

    check("tip: the list entry with its find text", tipText("Before:", "(empty)", "_", "two", true, {}, "List entry #3:", "two")
        == "Before: two\nList entry #3: \"two\"");
    check("tip: an entry without a find text stands alone", tipText("Before:", "(empty)", "_", "two", true, {}, "List entry #3", "")
        == "Before: two\nList entry #3");
    check("tip: the entry after a jump and an empty original", tipText("B:", "(e)", "_", "", true, "1 of 2", "E:", "a\tb")
        == "\0011 of 2\002 B: (e)\nE: \"a\\tb\"");
    check("tip: a long find text is cut", tipText("B:", "-", "_", "x", true, {}, "E:", std::string(50, 'f'))
        == "B: x\nE: \"" + std::string(40, 'f') + "...\"");
    check("tip: the find text cut between UTF-8 characters", tipText("B:", "-", "_", "x", true, {}, "E:", std::string(39, 'f') + "\xC3\xBC\xC3\xBC")
        == "B: x\nE: \"" + std::string(39, 'f') + "\xC3\xBC...\"");
    {
        std::string twelve;
        for (int i = 1; i <= 12; ++i) twelve += "x\n";
        const std::string tip = tipText("B:", "-", "_", twelve, true, {}, "E:", "x");
        const std::string end = "\n...\nE: \"x\"";
        check("tip: the entry follows the cut original", tip.size() > end.size() && tip.compare(tip.size() - end.size(), end.size(), end) == 0);
    }

    // What the calltip highlights: the original alone, all its lines, never a placeholder
    const auto highlighted = [](const ReplaceMarks::Tip& tip) { return tip.text.substr(tip.start, tip.end - tip.start); };
    check("tip highlight: the original only, after a jump and before the entry",
        highlighted(ReplaceMarks::tipText("Before:", "(empty)", "_", "two", true, "3 of 12", "E:", "two")) == "two");
    check("tip highlight: an original over lines, cut and with visible spaces",
        highlighted(ReplaceMarks::tipText("B:", "-", "_", " a\nb" + std::string(120, 'x'), true, {}, "E:", "x"))
        == "_a\\n\nb" + std::string(99, 'x') + "...");
    {
        const ReplaceMarks::Tip empty = ReplaceMarks::tipText("Before:", "(empty)", "_", "", true, {}, "E:", "x");
        check("tip highlight: none on the placeholder of an empty original", empty.start == empty.end && empty.text.find("(empty)") != std::string::npos);
    }
    check("tip highlight: the original only, before an entry without a find text",
        highlighted(ReplaceMarks::tipText("Before:", "(empty)", "_", "two", true, {}, "List entry #3", "")) == "two");

    std::printf("%d/%d checks passed\n", checks - fails, checks);
    return fails ? 1 : 0;
}
