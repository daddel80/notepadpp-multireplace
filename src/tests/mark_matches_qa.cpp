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

// QA for MarkMatches on the real Scintilla Document: every match carries its color in one indicator, runs come back in
// document order with their colors and group by color in a given order, a theme change recolors exactly the colors it
// maps, clearing takes all marks out, and the document keeps its current indicator and value. A neighbour plugin that
// changes the current indicator and value whenever an indicator changes (Notepad++ passes every fill on to all plugins)
// costs no mark its place or color.
// VERBATIM: the indicator messages as Editor::WndProc hands them to the Document.
// Build: Scintilla object files as for replace_marks_qa, then
//   g++ -std=c++17 -O1 -g -Wall -Wextra -fsanitize=address,undefined -I$S/include -I$S/src -I.. mark_matches_qa.cpp ../MarkMatches.cpp *.o -o mark_matches_qa
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
#include <random>
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
#include "../MarkMatches.h"
using namespace Scintilla; using namespace Scintilla::Internal;
static constexpr int kCpUtf8 = 65001;
static constexpr int kId = 13;          // Mark Matches
static constexpr int kNeighbour = 9;    // another plugin's indicator

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

static int g_fail = 0, g_checks = 0;
static void check(bool ok, const std::string& what) {
    ++g_checks;
    if (!ok) { ++g_fail; std::printf("FAIL  %s\n", what.c_str()); }
}

struct Doc {
    Document* d;
    std::map<std::pair<unsigned int, std::uintptr_t>, std::intptr_t> view;   // style messages: (message, indicator) -> value
    explicit Doc(const std::string& s) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(kCpUtf8);
        d->InsertString(0, s.c_str(), static_cast<Sci::Position>(s.size()));
        d->DeleteUndoHistory();
    }
    ~Doc() { d->Release(); }
    Doc(const Doc&) = delete;
    Doc& operator=(const Doc&) = delete;

    MarkMatches::Send send() {
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
            case SCI_INDICSETSTYLE: case SCI_INDICSETFLAGS: case SCI_INDICSETALPHA: case SCI_INDICSETOUTLINEALPHA: case SCI_INDICSETUNDER:
                view[{ msg, w }] = l; return 0;
            default: std::fprintf(stderr, "unexpected message %u\n", msg); std::abort();
            }
        };
    }
    int valueAt(int id, Sci::Position pos) const { return d->decorations->ValueAt(id, pos); }
};

// Another plugin: on every indicator change it makes its own indicator current with its own value, marks a character
// and leaves both set
struct Neighbour : DocWatcher {
    Doc& doc;
    bool busy = false;
    int calls = 0;
    explicit Neighbour(Doc& d) : doc(d) { doc.d->AddWatcher(this, nullptr); }
    ~Neighbour() override { doc.d->RemoveWatcher(this, nullptr); }
    Neighbour(const Neighbour&) = delete;
    Neighbour& operator=(const Neighbour&) = delete;
    void NotifyModified(Document*, DocModification mh, void*) override {
        if (busy || !FlagSet(mh.modificationType, ModificationFlags::ChangeIndicator)) return;
        busy = true;
        ++calls;
        doc.d->DecorationSetCurrentIndicator(kNeighbour);
        doc.d->decorations->SetCurrentValue(7);
        doc.d->DecorationFillRange(0, 7, 1);
        busy = false;
    }
    void NotifyModifyAttempt(Document*, void*) override {}
    void NotifySavePoint(Document*, void*, bool) override {}
    void NotifyDeleted(Document*, void*) noexcept override {}
    void NotifyStyleNeeded(Document*, void*, Sci::Position) override {}
    void NotifyErrorOccurred(Document*, void*, Scintilla::Status) override {}
    void NotifyGroupCompleted(Document*, void*) noexcept override {}
};

static std::intptr_t valueOf(MarkMatches::Color c) { return SC_INDICVALUEBIT | static_cast<std::intptr_t>(c & 0xFFFFFF); }

static std::string describe(const std::vector<MarkMatches::Run>& runs)
{
    std::string out;
    char item[64];
    for (const auto& r : runs) {
        std::snprintf(item, sizeof item, "[%ld,%ld)%06X ", static_cast<long>(r.start), static_cast<long>(r.end), static_cast<unsigned>(r.color));
        out += item;
    }
    return out;
}

int main()
{
    constexpr MarkMatches::Color RED = 0x3030E0, GREEN = 0x30C030, BLUE = 0xE03030;
    const std::string text = "alpha beta gamma delta alpha beta epsilon";

    // Style
    {
        Doc doc(text);
        const MarkMatches::Send view = doc.send();
        MarkMatches::applyStyle(view, kId, 150, 0);
        check(doc.view[{ SCI_INDICSETSTYLE, kId }] == INDIC_ROUNDBOX, "style: rounded box");
        check(doc.view[{ SCI_INDICSETFLAGS, kId }] == SC_INDICFLAG_VALUEFORE, "style: the color comes from the value");
        check(doc.view[{ SCI_INDICSETALPHA, kId }] == 150 && doc.view[{ SCI_INDICSETOUTLINEALPHA, kId }] == 0, "style: alpha and outline");
        check(doc.view[{ SCI_INDICSETUNDER, kId }] == 1, "style: under the text");
        const std::size_t before = doc.view.size();
        MarkMatches::applyStyle(view, -1, 150, 0);
        check(doc.view.size() == before, "style: no indicator, no message");
    }

    // Painting, runs, the current indicator
    for (const bool neighbour : { false, true }) {
        const std::string with = neighbour ? " (with a neighbour plugin)" : "";
        Doc doc(text);
        const MarkMatches::Send send = doc.send();
        std::unique_ptr<Neighbour> other;
        if (neighbour) other = std::make_unique<Neighbour>(doc);
        doc.d->DecorationSetCurrentIndicator(kNeighbour);   // as the other plugin left it
        doc.d->decorations->SetCurrentValue(7);
        {
            MarkMatches::Painter paint(send, kId);
            paint.fill(0, 5, RED);       // alpha
            paint.fill(23, 5, RED);      // alpha
            paint.fill(6, 4, GREEN);     // beta
            paint.fill(29, 4, GREEN);    // beta
            paint.fill(11, 5, BLUE);     // gamma, touches nothing
            paint.fill(16, 1, BLUE);     // the space after gamma: the same color merges
            paint.fill(17, 5, GREEN);    // delta right behind it: a run of its own
            paint.fill(40, 0, RED);      // empty: nothing
        }
        check(doc.d->decorations->GetCurrentIndicator() == kNeighbour && doc.d->decorations->GetCurrentValue() == 7,
            "painter: the current indicator and value come back" + with);
        check(doc.valueAt(kId, 0) == valueOf(RED) && doc.valueAt(kId, 4) == valueOf(RED) && doc.valueAt(kId, 5) == 0,
            "painter: a match carries its color" + with);
        check(doc.valueAt(kId, 6) == valueOf(GREEN) && doc.valueAt(kId, 12) == valueOf(BLUE) && doc.valueAt(kId, 17) == valueOf(GREEN),
            "painter: each match its own color" + with);
        const std::vector<MarkMatches::Run> runs = MarkMatches::runs(send, kId);
        const std::string expected = "[0,5)3030E0 [6,10)30C030 [11,17)E03030 [17,22)30C030 [23,28)3030E0 [29,33)30C030 ";
        check(describe(runs) == expected, "runs in document order, neighbours of other colors apart: " + describe(runs) + with);
        if (neighbour) {
            check(other->calls > 0, "the neighbour was called");
            check(doc.valueAt(kNeighbour, 0) == 7 && doc.valueAt(kNeighbour, 6) == 0, "the neighbour's indicator holds only its own mark");
        }

        // Recolor: only the colors the map names
        const std::unordered_map<MarkMatches::Color, MarkMatches::Color> theme{ { RED, 0x1010A0 }, { GREEN, 0x108010 } };
        doc.d->DecorationSetCurrentIndicator(kNeighbour);
        doc.d->decorations->SetCurrentValue(7);
        const std::size_t changed = MarkMatches::recolor(send, kId, theme);
        check(changed == 5, "recolor: the runs of mapped colors, " + std::to_string(changed) + with);
        check(describe(MarkMatches::runs(send, kId)) == "[0,5)1010A0 [6,10)108010 [11,17)E03030 [17,22)108010 [23,28)1010A0 [29,33)108010 ",
            "recolor: new colors in place, unmapped keep theirs: " + describe(MarkMatches::runs(send, kId)) + with);
        check(doc.d->decorations->GetCurrentIndicator() == kNeighbour && doc.d->decorations->GetCurrentValue() == 7,
            "recolor: the current indicator and value come back" + with);
        check(MarkMatches::recolor(send, kId, theme) == 0, "recolor: done twice changes nothing" + with);
        check(MarkMatches::recolor(send, kId, { { BLUE, BLUE } }) == 0, "recolor: a color that stays is no change" + with);
        check(MarkMatches::recolor(send, kId, {}) == 0 && MarkMatches::recolor(send, -1, theme) == 0, "recolor: nothing to map, no indicator" + with);

        // Clear
        MarkMatches::clear(send, kId);
        check(MarkMatches::runs(send, kId).empty(), "clear: no marks left" + with);
        check(doc.d->decorations->GetCurrentIndicator() == kNeighbour && doc.d->decorations->GetCurrentValue() == 7,
            "clear: the current indicator and value come back" + with);
        if (neighbour) check(doc.valueAt(kNeighbour, 0) == 7, "clear: the neighbour keeps its mark");
    }

    // Grouped by color in a given order
    {
        using MarkMatches::Run;
        constexpr MarkMatches::Color GRAY = 0x808080;
        const std::vector<Run> runs = { { 0, 2, BLUE }, { 3, 5, RED }, { 6, 8, GRAY }, { 9, 11, BLUE }, { 12, 14, GREEN }, { 15, 17, RED } };
        check(describe(MarkMatches::byColor(runs, { RED, BLUE, GREEN })) == "[3,5)3030E0 [15,17)3030E0 [0,2)E03030 [9,11)E03030 [12,14)30C030 [6,8)808080 ",
            "by color: the given order, each color in document order, a color the order lacks after them");
        check(describe(MarkMatches::byColor(runs, {})) == "[0,2)E03030 [9,11)E03030 [3,5)3030E0 [15,17)3030E0 [6,8)808080 [12,14)30C030 ",
            "by color: no order, the colors as they first appear");
        check(describe(MarkMatches::byColor(runs, { GREEN, RED, GREEN, BLUE })) == "[12,14)30C030 [3,5)3030E0 [15,17)3030E0 [0,2)E03030 [9,11)E03030 [6,8)808080 ",
            "by color: a color twice in the order keeps its first place");
        check(MarkMatches::byColor({}, { RED }).empty(), "by color: no runs");
    }

    // Edge cases
    {
        Doc doc("");
        const MarkMatches::Send send = doc.send();
        check(MarkMatches::runs(send, kId).empty(), "empty document: no runs");
        MarkMatches::clear(send, kId);
        check(MarkMatches::runs(send, -1).empty(), "no indicator: no runs");
        Doc other("text");
        const MarkMatches::Send send2 = other.send();
        { MarkMatches::Painter paint(send2, -1); paint.fill(0, 4, RED); }
        check(other.valueAt(kId, 0) == 0, "no indicator: the painter marks nothing");
        check(MarkMatches::runs(send2, kId).empty(), "a document that never held the indicator: no runs");
    }

    // Random matches against a model, with the neighbour
    {
        std::mt19937 rng(20261120u);
        std::string big(20000, 'x');
        Doc doc(big);
        const MarkMatches::Send send = doc.send();
        Neighbour other(doc);
        std::vector<MarkMatches::Color> model(big.size(), 0);
        const MarkMatches::Color palette[] = { RED, GREEN, BLUE, 0x123456, 0xABCDEF };
        {
            MarkMatches::Painter paint(send, kId);
            for (int i = 0; i < 3000; ++i) {
                const std::size_t pos = rng() % big.size();
                const std::size_t len = std::min<std::size_t>(1 + rng() % 12, big.size() - pos);
                const MarkMatches::Color c = palette[rng() % 5];
                paint.fill(static_cast<std::intptr_t>(pos), static_cast<std::intptr_t>(len), c);
                for (std::size_t k = pos; k < pos + len; ++k) model[k] = c;
            }
        }
        bool same = true;
        for (std::size_t k = 0; k < big.size() && same; ++k)
            same = doc.valueAt(kId, static_cast<Sci::Position>(k)) == (model[k] ? valueOf(model[k]) : 0);
        check(same, "random: every character holds the color last painted on it");
        bool ordered = true, joined = true;
        const auto runs = MarkMatches::runs(send, kId);
        for (std::size_t i = 1; i < runs.size(); ++i) {
            ordered = ordered && runs[i - 1].end <= runs[i].start;
            joined = joined && !(runs[i - 1].end == runs[i].start && runs[i - 1].color == runs[i].color);
        }
        check(ordered && joined, "random: runs ordered, touching runs differ in color");
        std::unordered_map<MarkMatches::Color, MarkMatches::Color> swap{ { RED, BLUE }, { BLUE, RED } };
        MarkMatches::recolor(send, kId, swap);
        same = true;
        for (std::size_t k = 0; k < big.size() && same; ++k) {
            MarkMatches::Color want = model[k] == RED ? BLUE : model[k] == BLUE ? RED : model[k];
            same = doc.valueAt(kId, static_cast<Sci::Position>(k)) == (want ? valueOf(want) : 0);
        }
        check(same, "random: swapping two colors recolors exactly them");
    }

    // Speed: 100,000 matches
    {
        std::string big(600000, 'a');
        Doc doc(big);
        const MarkMatches::Send send = doc.send();
        const auto t0 = std::chrono::steady_clock::now();
        {
            MarkMatches::Painter paint(send, kId);
            for (std::intptr_t i = 0; i < 100000; ++i) paint.fill(i * 6, 3, (i % 7) ? RED : GREEN);
        }
        const auto t1 = std::chrono::steady_clock::now();
        const std::size_t runs = MarkMatches::runs(send, kId).size();
        const auto t2 = std::chrono::steady_clock::now();
        std::printf("  100,000 matches: paint %.3f s, runs %.3f s\n", std::chrono::duration<double>(t1 - t0).count(),
            std::chrono::duration<double>(t2 - t1).count());
        check(runs == 100000, "speed: all matches are runs");
    }

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
