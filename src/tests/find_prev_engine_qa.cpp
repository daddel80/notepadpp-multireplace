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

// Engine-level QA for Find Previous on the real Scintilla Document and the N++ Boost regex bridge.
// VERBATIM: the start of handleFindPrevButton (the start of the current selection, the cursor
// without one, an empty hit there being the current position; the end of a new selection in the
// Selection scope), performSingleSearch, performRegexBackwardFallback, performSearchSelection and for
// the column scope findDelimitersInLine, getColumnInfo and the backward walk of performSearchColumn.
// The reference is what a forward walk over the text, each range of a scope or each selected cell finds.
//   A  from every cursor position: the last hit that ends at or before the cursor, else (wrap) the last hit
//   B  pressing Find Previous again and again visits every hit from right to left, none twice
//   C  Selection scope (one range, ranges with a gap, touching ranges, short ranges cutting words): from a
//      new selection every hit right to left, none twice; from every cursor in the scope the last hit before it
//   D  CSV column scope (quotes, empty cells, CRLF, a two-character delimiter, UTF-8, literals spanning a
//      delimiter): pressing through from the end and from every cursor position the last hit before it
// Build: object files as for onepass_engine_qa (Notepad++ sources: scintilla + boostregex, Boost.Regex), then
//   g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src find_prev_engine_qa.cpp *.o -lboost_regex -o find_prev_engine_qa
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
using namespace Scintilla; using namespace Scintilla::Internal;
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

static constexpr int F_POSIX = 0x00400000, F_SKIPCRLF = 0x08000000, F_ALL = 0x40000000;
static constexpr int F_ALLOWATSTART = (int)0x80000000;
static FindOption fo(int v) { return static_cast<FindOption>(v); }
static int findFlags(bool regex) {   // buildSearchFlags(wholeWord=false, matchCase=true, regex, isReplaceAll=false)
    return (int)FindOption::MatchCase | (regex ? (int)FindOption::RegExp | F_POSIX | F_SKIPCRLF | F_ALL : 0);
}

struct Doc {
    Document* d;
    explicit Doc(const std::string& s) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(kCpUtf8); d->SetUndoCollection(false);
        d->SetCaseFolder(std::make_unique<CaseFolderUnicode>());
        d->InsertString(0, s.c_str(), s.size());
    }
    ~Doc() { d->Release(); }
    Sci::Position len() const { return d->Length(); }
    Sci::Position positionAfter(Sci::Position p) const { return d->MovePositionOutsideChar(p + 1, 1, true); }   // SCI_POSITIONAFTER
    bool isCursorPosition(Sci::Position p) const { return d->MovePositionOutsideChar(p, 1, true) == p; }      // not inside CRLF or a character
};

struct Hit {
    Sci::Position pos = -1, len = 0;
    bool operator==(const Hit& o) const { return pos == o.pos && len == o.len; }
};

static Hit regexBackward(Doc& doc, Sci::Position start, Sci::Position end, const std::string& pat, int flags, Sci::Position skipEmptyAt,
                         Sci::Position regionEnd);

// performSingleSearch: backward regex goes through the fallback, otherwise SCI_SETTARGETRANGE + SCI_SEARCHINTARGET,
// refusing what the panel refuses and an empty hit at skipEmptyAt
static Hit single(Doc& doc, Sci::Position from, Sci::Position to, const std::string& pat, int flags, Sci::Position skipEmptyAt = -1,
                  Sci::Position regionEnd = -1, bool isBackward = false) {
    if ((from > to || (isBackward && from == to)) && (flags & (int)FindOption::RegExp))
        return regexBackward(doc, from, to, pat, flags, skipEmptyAt, regionEnd);
    Sci::Position l = (Sci::Position)pat.size();
    const Sci::Position p = doc.d->FindText(from, to, pat.c_str(), fo(flags), &l);
    if (p < 0 || l < 0 || p + l > doc.len()) return {};
    if (from <= to && p + l > to) return {};
    if (l == 0 && p == skipEmptyAt) return {};
    return { p, l };
}
static Sci::Position advanceAfterMatch(Doc& doc, const Hit& r) {
    if (r.len > 0) return r.pos + r.len;
    const Sci::Position after = doc.positionAfter(r.pos);
    const Sci::Position next = (after > r.pos) ? after : (r.pos + 1);
    return (next > doc.len()) ? doc.len() : next;
}
// performRegexBackwardFallback
static Hit regexBackward(Doc& doc, Sci::Position start, Sci::Position end, const std::string& pat, int flags, Sci::Position skipEmptyAt,
                         Sci::Position regionEnd) {
    const Sci::Position scanEnd = start;
    const Sci::Position checkEnd = (regionEnd >= 0) ? regionEnd : doc.len();
    Sci::Position pos = end;
    Hit last;
    const int walkFlags = flags | F_ALLOWATSTART;
    while (pos <= scanEnd) {
        const Hit r = single(doc, pos, scanEnd, pat, walkFlags, skipEmptyAt);
        if (r.pos < 0) break;
        const Hit unclipped = single(doc, r.pos, checkEnd, pat, walkFlags, skipEmptyAt);
        if (!(unclipped == r)) break;
        if (unclipped.pos + unclipped.len > scanEnd) break;
        last = r;
        Sci::Position next = advanceAfterMatch(doc, r);
        if (next <= pos) next = pos + 1;
        pos = next;
    }
    return last;
}
// performSearchBackward, All Text
static Hit backward(Doc& doc, Sci::Position start, const std::string& pat, int flags, Sci::Position skipEmptyAt) {
    return single(doc, start, 0, pat, flags, skipEmptyAt, doc.len(), true);
}
// handleFindPrevButton, single mode, All Text, wrap around on
static Hit findPrevious(Doc& doc, Sci::Position selStart, const std::string& pat, int flags) {
    const Hit h = backward(doc, selStart, pat, flags, selStart);
    return (h.pos >= 0) ? h : backward(doc, doc.len(), pat, flags, -1);
}

using ScopeRange = std::pair<Sci::Position, Sci::Position>;
// performSearchSelection, backward: the ranges from the last one, each hit checked against its own range end
static Hit scopeBackward(Doc& doc, const std::vector<ScopeRange>& scope, Sci::Position start, const std::string& pat, int flags,
                         Sci::Position skipEmptyAt) {
    for (auto it = scope.rbegin(); it != scope.rend(); ++it) {
        if (start < it->first) continue;
        const Sci::Position from = std::min(start, it->second);
        if (from == it->first) continue;
        const Hit h = single(doc, from, it->first, pat, flags, skipEmptyAt, it->second, true);
        if (h.pos >= 0) return h;
        start = it->first;
    }
    return {};
}
// Reference for the Selection scope: what the forward walk finds in each range, a hit on a shared edge once
static std::vector<Hit> walkScope(Doc& doc, const std::vector<ScopeRange>& scope, const std::string& pat, int flags) {
    std::vector<Hit> v;
    for (const ScopeRange& r : scope)
        for (Sci::Position pos = r.first; pos <= r.second;) {
            const Hit h = single(doc, pos, r.second, pat, flags | F_ALLOWATSTART);
            if (h.pos < 0) break;
            if (v.empty() || !(v.back() == h)) v.push_back(h);
            if (h.len == 0 && h.pos >= r.second) break;
            const Sci::Position next = advanceAfterMatch(doc, h);
            pos = (next > pos) ? next : pos + 1;
        }
    return v;
}
// Selection scopes as selections leave them: non-empty ranges on cursor positions, sorted, one range, a gap, touching, short
static std::vector<std::vector<ScopeRange>> scopesFor(Doc& doc) {
    const Sci::Position n = doc.len();
    const std::vector<std::vector<ScopeRange>> raw = { { { 1, n - 1 } }, { { 0, n / 3 }, { n / 2, n } }, { { 0, n / 2 }, { n / 2, n } },
                                                  { { 1, 4 }, { 6, 9 }, { 11, 14 } } };
    std::vector<std::vector<ScopeRange>> out;
    for (const auto& scope : raw) {
        std::vector<ScopeRange> s;
        for (const ScopeRange& r : scope) {
            const Sci::Position a = doc.d->MovePositionOutsideChar(std::min(r.first, n), 1, true);
            const Sci::Position b = doc.d->MovePositionOutsideChar(std::min(r.second, n), 1, true);
            if (b > a) s.push_back({ a, b });
        }
        if (!s.empty()) out.push_back(s);
    }
    return out;
}

// Reference: every hit a forward walk sees, as Find All steps through the text
static std::vector<Hit> walk(Doc& doc, const std::string& pat, int flags) {
    std::vector<Hit> v;
    for (Sci::Position pos = 0; pos <= doc.len();) {
        const Hit r = single(doc, pos, doc.len(), pat, flags | F_ALLOWATSTART);
        if (r.pos < 0) break;
        v.push_back(r);
        if (r.len == 0 && r.pos >= doc.len()) break;
        const Sci::Position next = advanceAfterMatch(doc, r);
        pos = (next > pos) ? next : pos + 1;
    }
    return v;
}

// Column scope: the delimiter positions per line (findDelimitersInLine; SCI_LINELENGTH counts the line end)
struct LineDelims { std::vector<Sci::Position> positions; Sci::Position lineLength = 0; };
struct Columns { std::set<int> columns; std::string delimiter, quote; };
static std::vector<LineDelims> findDelimiters(Doc& doc, const Columns& cd) {
    std::vector<LineDelims> out;
    for (Sci::Line line = 0; line < doc.d->LinesTotal(); ++line) {
        LineDelims lineInfo;
        const Sci::Position lineStart = doc.d->LineStart(line);
        lineInfo.lineLength = doc.d->LineStart(line + 1) - lineStart;
        std::string buf(static_cast<size_t>(lineInfo.lineLength), '\0');
        doc.d->GetCharRange(buf.data(), lineStart, lineInfo.lineLength);
        const std::string_view lineContent(buf);
        const size_t delimLen = cd.delimiter.size();
        const std::string_view delimiter(cd.delimiter);
        const bool hasQuoteChar = !cd.quote.empty();
        const char currentQuoteChar = hasQuoteChar ? cd.quote[0] : '\0';
        const char delimChar = delimiter[0];
        size_t pos = 0;
        bool inQuotes = false;
        while (pos < lineContent.size()) {
            if (hasQuoteChar && lineContent[pos] == currentQuoteChar) { inQuotes = !inQuotes; ++pos; continue; }
            if (!inQuotes) {
                if (delimLen == 1) {
                    if (lineContent[pos] == delimChar) { lineInfo.positions.push_back(static_cast<Sci::Position>(pos)); ++pos; continue; }
                }
                else {
                    const size_t foundPos = lineContent.find(delimiter, pos);
                    if (foundPos == std::string_view::npos) break;
                    if (hasQuoteChar) {
                        const size_t nextQuote = lineContent.find(currentQuoteChar, pos);
                        if (nextQuote != std::string_view::npos && nextQuote < foundPos) { pos = nextQuote; continue; }
                    }
                    lineInfo.positions.push_back(static_cast<Sci::Position>(foundPos));
                    pos = foundPos + delimLen;
                    continue;
                }
            }
            ++pos;
        }
        out.push_back(lineInfo);
    }
    return out;
}
// The bounds of a cell as performSearchColumn takes them: the last column runs to the start of the next line
static ScopeRange cellBounds(Doc& doc, const Columns& cd, const std::vector<LineDelims>& delims, Sci::Line line, size_t column) {
    const LineDelims& lineInfo = delims[line];
    const size_t totalColumns = lineInfo.positions.size() + 1;
    const Sci::Position lineStartPos = doc.d->LineStart(line), lineEndPos = lineStartPos + lineInfo.lineLength;
    const Sci::Position startColumn = (column == 1) ? lineStartPos
        : lineStartPos + lineInfo.positions[column - 2] + static_cast<Sci::Position>(cd.delimiter.size());
    const Sci::Position endColumn = (column == totalColumns) ? lineEndPos : lineStartPos + lineInfo.positions[column - 1];
    return { startColumn, endColumn };
}
// getColumnInfo and performSearchColumn, backward
static Hit columnBackward(Doc& doc, const Columns& cd, const std::vector<LineDelims>& delims, Sci::Position start,
                          const std::string& pat, int flags, Sci::Position skipEmptyAt) {
    const Sci::Line startLine = doc.d->SciLineFromPosition(start);
    size_t startColumnIndex = 1;
    if (startLine < doc.d->LinesTotal() && startLine < static_cast<Sci::Line>(delims.size())) {
        const auto& linePositions = delims[startLine].positions;
        const Sci::Position lineStartPos = doc.d->LineStart(startLine);
        size_t i = 0;
        for (; i < linePositions.size(); ++i)
            if (start <= lineStartPos + linePositions[i]) { startColumnIndex = i + 1; break; }
        if (i == linePositions.size()) startColumnIndex = linePositions.size() + 1;
    }
    for (Sci::Line line = startLine; line >= 0; --line) {
        if (line >= static_cast<Sci::Line>(delims.size())) break;
        const size_t totalColumns = delims[line].positions.size() + 1;
        for (size_t column = (line == startLine) ? startColumnIndex : totalColumns; column >= 1; --column) {
            if (!cd.columns.count(static_cast<int>(column))) continue;
            const ScopeRange cell = cellBounds(doc, cd, delims, line, column);
            Sci::Position endColumn = cell.second;
            if (start < cell.first) continue;
            if (start >= cell.first && start <= endColumn) endColumn = start;
            const Hit h = single(doc, endColumn, cell.first, pat, flags, skipEmptyAt, cell.second, true);
            if (h.pos >= 0) return h;
        }
    }
    return {};
}

static std::string show(const std::string& s) { std::string o; for (char c : s) o += (c == '\n') ? std::string("|") : (c == '\r') ? std::string("~") : std::string(1, c); return o; }

int main() {
    const std::vector<std::string> texts = { "foo bar foo baz foo bar", "foofoofoo", "aa\nbb\n\ncc\n", "ab cd\r\nef\r\n\r\ng",
                                             "x  yy x", "f\xC3\xA4\xC3\xA4 f\xC3\xA4\xC3\xA4\n\xC3\xA4" };
    const std::vector<std::pair<std::string, bool>> pats = { { "foo", false }, { "a", false }, { "\xC3\xA4", false },
        { "^", true }, { "$", true }, { "\\b", true }, { "(?=b)", true }, { "x*", true }, { "\\w+", true }, { "fo+", true },
        { "b+|^", true }, { "^$", true }, { "\\r?\\n", true } };
    int failsA = 0, failsB = 0, failsC = 0, cases = 0, runs = 0, scopeRuns = 0, scopeClicks = 0;
    for (const auto& t : texts) for (const auto& [pat, regex] : pats) {
        const int flags = findFlags(regex);
        Doc doc(t);
        const std::vector<Hit> hits = walk(doc, pat, flags);
        for (Sci::Position c = 0; c <= doc.len(); ++c) {   // A: cursor without selection
            if (!doc.isCursorPosition(c)) continue;
            Hit want;
            for (const Hit& h : hits) if (h.pos + h.len <= c && !(h.len == 0 && h.pos == c)) want = h;
            if (want.pos < 0 && !hits.empty()) want = hits.back();
            const Hit got = findPrevious(doc, c, pat, flags);
            ++cases;
            if (!(got == want) && ++failsA <= 8)
                std::printf("A FAIL '%s' in '%s', cursor %lld: got %lld+%lld, want %lld+%lld\n", show(pat).c_str(), show(t).c_str(),
                    (long long)c, (long long)got.pos, (long long)got.len, (long long)want.pos, (long long)want.len);
        }
        std::vector<Hit> want;   // B: from the end, every press starts at the hit the last one marked
        for (const Hit& h : hits) if (!(h.len == 0 && h.pos == doc.len())) want.insert(want.begin(), h);
        std::vector<Hit> got;
        for (Sci::Position selStart = doc.len(); got.size() <= hits.size();) {
            const Hit h = backward(doc, selStart, pat, flags, selStart);
            if (h.pos < 0) break;
            got.push_back(h);
            selStart = h.pos;
        }
        ++runs;
        if (got != want && ++failsB <= 8)
            std::printf("B FAIL '%s' in '%s': %zu presses, %zu hits expected\n", show(pat).c_str(), show(t).c_str(), got.size(), want.size());

        for (const std::vector<ScopeRange>& scope : scopesFor(doc)) {   // C: Selection scope
            const std::vector<Hit> inScope = walkScope(doc, scope, pat, flags);
            const std::vector<Hit> wantC(inScope.rbegin(), inScope.rend());
            std::vector<Hit> gotC;   // a new selection starts at its end, where an empty hit counts; a follow-up at the hit
            for (Sci::Position selStart = scope.back().second, skip = -1; gotC.size() <= inScope.size();) {
                const Hit h = scopeBackward(doc, scope, selStart, pat, flags, skip);
                if (h.pos < 0) break;
                gotC.push_back(h);
                selStart = skip = h.pos;
            }
            ++scopeRuns;
            if (gotC != wantC && ++failsC <= 8)
                std::printf("C FAIL '%s' in '%s', %zu ranges: %zu presses, %zu hits expected\n", show(pat).c_str(), show(t).c_str(),
                    scope.size(), gotC.size(), wantC.size());
            for (Sci::Position c = scope.front().first; c <= scope.back().second; ++c) {   // a click inside the scope
                if (!doc.isCursorPosition(c)) continue;
                Hit wantClick;
                for (const Hit& h : inScope) if (h.pos + h.len <= c && !(h.len == 0 && h.pos == c)) wantClick = h;
                const Hit gotClick = scopeBackward(doc, scope, c, pat, flags, c);
                ++scopeClicks;
                if (!(gotClick == wantClick) && ++failsC <= 8)
                    std::printf("C FAIL '%s' in '%s', cursor %lld: got %lld+%lld, want %lld+%lld\n", show(pat).c_str(), show(t).c_str(),
                        (long long)c, (long long)gotClick.pos, (long long)gotClick.len, (long long)wantClick.pos, (long long)wantClick.len);
            }
        }
    }

    struct Csv { std::string text; std::string delimiter, quote; };   // D: column scope
    const std::vector<Csv> csvs = { { "a,foo,bar\nfoo,x,foofoo\n,,foo\nlast,foo", ",", "" },
        { "\"a,b\",foo,\"x\"\"y\"\nfoo,\"foo,bar\",z\n", ",", "\"" }, { "a;foo\r\nfoo;b;\r\n;foo bar;x\r\n", ";", "" },
        { "a::foo::b\nfoo::::c\nx::y::foo\n", "::", "" }, { "\xC3\xA4,f\xC3\xB6\xC3\xB6,x\nf\xC3\xB6\xC3\xB6,\xC3\xA4\n", ",", "" },
        { "xxab,cd\nab,c\nfoob,car\n", ",", "" } };
    const std::vector<std::set<int>> columnSets = { { 1 }, { 2 }, { 3 }, { 1, 3 }, { 2, 3 }, { 1, 2, 3 }, { 4 } };
    const std::vector<std::pair<std::string, bool>> csvPats = { { "foo", false }, { "o", false }, { "ab,c", false },
        { "fo+", true }, { "\\w+", true }, { ".+", true }, { "[^,;]+", true }, { "$", true }, { "^", true }, { "\\b", true },
        { "o*", true }, { "f\xC3\xB6+", true }, { "b,c", true } };
    int failsD = 0, columnRuns = 0, columnClicks = 0;
    for (const Csv& csv : csvs) for (const std::set<int>& cols : columnSets) for (const auto& [pat, regex] : csvPats) {
        const int flags = findFlags(regex);
        Doc doc(csv.text);
        const Columns cd{ cols, csv.delimiter, csv.quote };
        const std::vector<LineDelims> delims = findDelimiters(doc, cd);
        std::vector<ScopeRange> cells;
        for (Sci::Line line = 0; line < static_cast<Sci::Line>(delims.size()); ++line)
            for (size_t column = 1; column <= delims[line].positions.size() + 1; ++column)
                if (cols.count(static_cast<int>(column))) cells.push_back(cellBounds(doc, cd, delims, line, column));
        const std::vector<Hit> inCells = walkScope(doc, cells, pat, flags);
        std::vector<Hit> wantD;   // from the end of the text, where an empty hit is the cursor itself
        for (auto it = inCells.rbegin(); it != inCells.rend(); ++it) if (!(it->len == 0 && it->pos == doc.len())) wantD.push_back(*it);
        std::vector<Hit> gotD;
        for (Sci::Position selStart = doc.len(); gotD.size() <= inCells.size();) {
            const Hit h = columnBackward(doc, cd, delims, selStart, pat, flags, selStart);
            if (h.pos < 0) break;
            gotD.push_back(h);
            selStart = h.pos;
        }
        ++columnRuns;
        if (gotD != wantD && ++failsD <= 8)
            std::printf("D FAIL '%s' in '%s', columns %zu: %zu presses, %zu hits expected\n", show(pat).c_str(), show(csv.text).c_str(),
                cols.size(), gotD.size(), wantD.size());
        for (Sci::Position c = 0; c <= doc.len(); ++c) {
            if (!doc.isCursorPosition(c)) continue;
            Hit wantClick;
            for (const Hit& h : inCells) if (h.pos + h.len <= c && !(h.len == 0 && h.pos == c)) wantClick = h;
            const Hit gotClick = columnBackward(doc, cd, delims, c, pat, flags, c);
            ++columnClicks;
            if (!(gotClick == wantClick) && ++failsD <= 8)
                std::printf("D FAIL '%s' in '%s', cursor %lld: got %lld+%lld, want %lld+%lld\n", show(pat).c_str(), show(csv.text).c_str(),
                    (long long)c, (long long)gotClick.pos, (long long)gotClick.len, (long long)wantClick.pos, (long long)wantClick.len);
        }
    }
    std::printf("A  %d cursor positions: %d wrong\nB  %d texts x patterns pressed through: %d wrong\n", cases, failsA, runs, failsB);
    std::printf("C  %d scopes pressed through, %d cursor positions in them: %d wrong\n", scopeRuns, scopeClicks, failsC);
    std::printf("D  %d column runs pressed through, %d cursor positions in them: %d wrong\n", columnRuns, columnClicks, failsD);
    const bool ok = failsA == 0 && failsB == 0 && failsC == 0 && failsD == 0;
    std::printf("\n%s\n", ok ? "ALL CHECKS PASSED" : "CHECKS FAILED");
    return ok ? 0 : 1;
}
