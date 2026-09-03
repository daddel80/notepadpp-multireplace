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

// Engine-level QA for the one-pass Replace All walk (OnePassHits) on the real
// Scintilla Document and the N++ Boost regex bridge - the same code path
// SCI_SEARCHINTARGET takes inside Notepad++.
//   on     : OnePassHits with remember = true (next hits kept)
//   off    : OnePassHits with remember = false (every entry searched every step)
//   old    : the one-pass flow before OnePassHits (probe EMPTYMATCH_ALL, verify NOTAFTERMATCH)
//   classic: single-entry Replace All as N++ processRange does it (NOTAFTERMATCH, bridge watcher)
// Sections: A on == old on lists without empty matches, B one entry == classic
// (the old flow replaced nothing or never terminated for ^ $ \b x*), C edge lists
// on == off (lookarounds, size changes, CRLF, BOF/EOF), D fuzz on == off in the
// full text (and == classic for single entries), E column scope, F selection
// scope, G external edits and reset(), H custom word characters (SCI_SETWORDCHARS),
// I ANSI code page, J start position inside the text, K completeness of the
// context classification (every regex construct is measured against the engine:
// does a match at P depend on the bytes before P, and does the class know),
// L1-L3 entry classes in isolation (literals; plus regex without context; plus
// context-dependent). D-L also vary the replacement length per hit and disable entries.
// Usage: onepass_engine_qa [quick|full] [seed] [maxTextLen]
//        onepass_engine_qa trace '<text, | = LF, ~ = CR>' find repl r|l|w ...   (step trace, ML=1 for match list 1,3,4)
//        onepass_engine_qa bench [KB]                                          (on / off / old timings on prose)
// The logic check that builds without these sources is onepass_qa.cpp.
//
// Build (needs the Notepad++ sources: scintilla + boostregex, and Boost.Regex):
//   NPP=/path/to/notepad-plus-plus; S=$NPP/scintilla
//   for f in CellBuffer Document PerLine RunStyles Decoration CaseFolder CaseConvert CharClassify
//            CharacterCategoryMap CharacterType UniConversion DBCS ChangeHistory UndoHistory RESearch; do
//     g++ -std=c++17 -O2 -DNDEBUG -DSCI_OWNREGEX -I$S/include -I$S/src -c $S/src/$f.cxx; done
//   g++ -std=c++17 -O2 -DNDEBUG -DSCI_OWNREGEX -I$S/include -I$S/src -I$NPP/boostregex
//       -c $NPP/boostregex/BoostRegExSearch.cxx $NPP/boostregex/UTF8DocumentIterator.cxx
//   g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src -I.. onepass_engine_qa.cpp ../OnePassHits.cpp *.o -lboost_regex -o onepass_engine_qa
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
#include <chrono>
#include <functional>
#include <random>
#include <set>
#include "../OnePassHits.h"

static constexpr int F_POSIX = 0x00400000, F_SKIPCRLF = 0x08000000, F_NOTAFTER = 0x20000000, F_ALL = 0x40000000;
static constexpr int F_ALLOWATSTART = (int)0x80000000;
static FindOption fo(int v) { return static_cast<FindOption>(v); }

struct Entry { std::string find, repl; bool regex = false, wholeWord = false, matchCase = true, nullable = false; };   // nullable: can match empty (pool tag for the level sections)

static int baseFlags(const Entry& e) {
    int f = e.matchCase ? (int)FindOption::MatchCase : 0;
    if (e.wholeWord) f |= (int)FindOption::WholeWord;
    if (e.regex) f |= (int)FindOption::RegExp | F_POSIX | F_SKIPCRLF;
    return f;
}
static int replFlags(const Entry& e) { return baseFlags(e) | (e.regex ? F_NOTAFTER : 0); }          // buildSearchFlags(isReplaceAll=true)
static int probeFlags(const Entry& e) { return baseFlags(e) | (e.regex ? F_ALL : 0); }              // buildSearchFlags(isReplaceAll=false)
static int newFlags(const Entry& e) { return replFlags(e) | F_ALLOWATSTART; }                       // OnePassHits

struct DocOptions { int codePage = kCpUtf8; const char* wordChars = nullptr; };   // wordChars: as SCI_SETWORDCHARS would set them
struct Doc {
    Document* d;
    bool refused = false;   // the last find() was refused by the engine, not answered with "nothing here"
    Doc(const std::string& s, DocOptions o = {}) {
        d = new Document(DocumentOption::Default); d->AddRef(); d->SetDBCSCodePage(o.codePage); d->SetUndoCollection(false); d->SetCaseFolder(std::make_unique<CaseFolderUnicode>());
        if (o.wordChars) { d->SetDefaultCharClasses(false); d->SetCharClasses(reinterpret_cast<const unsigned char*>(o.wordChars), CharacterClass::word); }
        d->InsertString(0, s.c_str(), s.size());
    }
    ~Doc() { d->Release(); }
    Sci::Position len() const { return d->Length(); }
    Sci::Position find(Sci::Position from, Sci::Position to, const std::string& pat, int flags, Sci::Position& lenOut) {
        lenOut = (Sci::Position)pat.size();
        if (pat.empty() || from > to) return -1;
        Sci::Position p = d->FindText(from, to, pat.c_str(), fo(flags), &lenOut);
        if (p >= 0 && lenOut < 0) { refused = true; return -1; }   // performSingleSearch refuses a match ending before its start
        if (p == -2 || p == -3) { refused = true; return -1; }     // invalid pattern or an exception in the bridge
        return (p < 0 || p + lenOut > to) ? -1 : p;
    }
    std::string text() const { std::string s(d->Length(), '\0'); d->GetCharRange(s.data(), 0, d->Length()); return s; }
    Sci::Position replace(Sci::Position p, Sci::Position oldLen, std::string r, bool regex) {   // returns inserted length
        if (regex) {                                                                       // SCI_REPLACETARGETRE path: $1, \n, \U... substituted
            Sci::Position l = (Sci::Position)r.size();
            const char* sub = d->SubstituteByPosition(r.c_str(), &l);
            r.assign(sub ? sub : "", sub ? (size_t)l : 0);
        }
        d->DeleteChars(p, oldLen);
        d->InsertString(p, r.c_str(), r.size());
        return (Sci::Position)r.size();
    }
    Sci::Position next(Sci::Position p) const { return d->NextPosition(p, 1); }
    Sci::Position nextChar(Sci::Position p) const { return (d->CharAt(p) == '\r' && d->CharAt(p + 1) == '\n') ? p + 2 : d->NextPosition(p, 1); }   // bridge nextCharacter()
    void backwardSearch() { Sci::Position l = 1; d->FindText(d->Length(), 0, "a", fo((int)FindOption::RegExp | F_POSIX), &l); }   // a Find Previous before the run: the bridge keeps _lastDirection = -1
};

struct Hit { Sci::Position pos = -1, len = 0; };
struct Run { std::string text; std::vector<int> replaced, found; size_t searches = 0; };

static bool traceSteps = false;
static std::string showAll(std::string s) { std::string o; for (char c : s) { if (c == '\n') o += "|"; else if (c == '\r') o += "~"; else o += c; } return o; }

// ---------------------------------------------------------------- search scope as the panel applies it
// Full: whole text. Column: per line, only the selected delimiter-separated fields (performSearchColumn:
// a field runs to its delimiter, the last one to the start of the next line). Selection: stored ranges,
// moved after every replacement like adjustSelectionScope.
struct Scoping {
    OnePassHits::Scope scope = OnePassHits::Scope::Full;
    char delim = ','; std::set<int> columns;
    std::vector<std::pair<Sci::Position, Sci::Position>> sel;
    Hit search(Doc& doc, Sci::Position from, const std::string& pat, int flags) const {
        Hit h;
        if (scope == OnePassHits::Scope::Full) { h.pos = doc.find(from, doc.len(), pat, flags, h.len); return h; }
        if (scope == OnePassHits::Scope::Selection) {
            for (const auto& r : sel) {
                if (from > r.second) continue;
                const Sci::Position s = std::max(from, r.first);
                if (s == r.second) continue;
                h.pos = doc.find(s, r.second, pat, flags, h.len);
                if (h.pos >= 0) return h;
            }
            return Hit{};
        }
        const Sci::Position lines = doc.d->LinesTotal();
        for (Sci::Position l = doc.d->LineFromPosition(from); l < lines; ++l) {
            const Sci::Position ls = doc.d->LineStart(l), le = doc.d->LineStart(l + 1);
            Sci::Position fs = ls; int col = 1;
            for (Sci::Position p = ls; p <= le; ++p) {
                if (p < le && doc.d->CharAt(p) != delim) continue;
                const Sci::Position fe = (p < le) ? p : le;
                if (columns.count(col) && fe >= from) { h.pos = doc.find(std::max(fs, from), fe, pat, flags, h.len); if (h.pos >= 0) return h; }
                fs = p + 1; ++col;
            }
        }
        return Hit{};
    }
    void afterReplace(Sci::Position pos, Sci::Position oldLen, Sci::Position newLen) {
        const Sci::Position delta = newLen - oldLen; if (delta == 0) return;
        for (auto& r : sel) { if (pos + oldLen <= r.first) { r.first += delta; r.second += delta; } else if (pos < r.second) r.second += delta; }
    }
};

// ---------------------------------------------------------------- shared step body (replace or skip)
// matchList: replace only these hit numbers (per entry), empty = all
struct Walk {
    Doc& doc; const std::vector<Entry>& list; const std::set<int>& matchList; Run run; size_t steps = 0; bool capped = false;
    bool step() { return ++steps < 200000 || (capped = true, false); }
    Walk(Doc& d, const std::vector<Entry>& l, const std::set<int>& m) : doc(d), list(l), matchList(m) { run.replaced.assign(l.size(), 0); run.found.assign(l.size(), 0); }
    bool take(size_t w) { ++run.found[w]; return matchList.empty() || matchList.count(run.found[w]); }
};

// ---------------------------------------------------------------- the walk as onePassReplaceAll runs it, over OnePassHits
struct HitsPass : Walk {
    bool remember; Scoping sc; bool dynamicRepl = false; std::set<size_t> extEdits; Sci::Position startPos = 0;   // extEdits: steps with an outside insertion + reset()
    HitsPass(Doc& d, const std::vector<Entry>& l, const std::set<int>& m, bool rem, Scoping s = {}) : Walk(d, l, m), remember(rem), sc(std::move(s)) {}
    OnePassDoc view() {
        OnePassDoc v;
        v.search = [this](const std::string& pattern, int flags, Sci_Position from) { ++run.searches; doc.refused = false; const Hit h = sc.search(doc, from, pattern, flags); if (traceSteps) std::printf("      search '%s' from %lld -> [%lld,+%lld]%s\n", pattern.c_str(), (long long)from, (long long)h.pos, (long long)h.len, doc.refused ? " (refused)" : ""); return OnePassHit{ h.pos, (h.pos < 0 && doc.refused) ? -1 : h.len }; };
        v.searchRange = [this](const std::string& pattern, int flags, Sci_Position from, Sci_Position to) { ++run.searches; Hit h; h.pos = doc.find(from, to, pattern, flags, h.len); return OnePassHit{ h.pos, h.len }; };
        v.charAt = [this](Sci_Position pos) { return (int)(unsigned char)doc.d->CharAt(pos); };
        v.positionAfter = [this](Sci_Position pos) { return (Sci_Position)doc.next(pos); };
        v.lineCount = [this]() { return (Sci_Position)doc.d->LinesTotal(); };
        v.length = [this]() { return (Sci_Position)doc.len(); };
        v.utf8 = doc.d->CodePage() == kCpUtf8;
        return v;
    }
    std::string replFor(size_t w) const {   // formula-like entries: replacement varies per hit
        std::string r = list[w].repl; const int k = run.found[w];
        if (dynamicRepl) { if (k % 5 == 0) r.clear(); else if (k % 3 == 0) r += r; else if (k % 7 == 0) r += "\n"; }
        return r;
    }
    void go() {
        std::vector<OnePassEntry> entries(list.size());
        for (size_t i = 0; i < list.size(); ++i) { entries[i].findText = list[i].find; entries[i].flags = newFlags(list[i]); entries[i].active = !list[i].find.empty(); entries[i].regex = list[i].regex; entries[i].wholeWord = list[i].wholeWord; }
        OnePassDoc v = view();
        OnePassHits hits(v, entries, sc.scope, remember);
        Sci::Position pos = startPos;
        while (step()) {
            if (extEdits.count(steps) && pos < doc.len()) { doc.d->InsertString(doc.next(pos), "zz", 2); hits.reset(); }   // outside edit at a character boundary
            size_t w; const OnePassHit h = hits.next(pos, w); if (h.pos < 0) break;
            const bool doReplace = take(w);
            if (traceSteps) std::printf("   %s pos=%lld len=%lld win=%zu '%s' hit=[%lld,+%lld]%s\n", remember ? "on " : "off", (long long)pos, (long long)doc.len(), w, list[w].find.c_str(), (long long)h.pos, (long long)h.len, doReplace ? "" : " skip");
            if (doReplace) {
                hits.beforeReplace(h);
                // replaceOne verification: an independent search with the entry's own flags in its own scope,
                // starting where the class says the hit is visible from (its position, or the search start for \K)
                ++run.searches; const Hit vh = sc.search(doc, hits.verifyFrom(w, h.pos), list[w].find, newFlags(list[w]));
                if (!(vh.pos == h.pos && vh.len == h.len)) { if (traceSteps) std::printf("      verification mismatch entry %zu at %lld, skipped\n", w, (long long)h.pos); hits.afterSkip(w, h); pos = h.pos + h.len; continue; }
                const Sci::Position ins = doc.replace(h.pos, h.len, replFor(w), list[w].regex); ++run.replaced[w];
                sc.afterReplace(h.pos, h.len, ins);
                hits.afterReplace(w, h, ins - h.len); pos = h.pos + ins;
            }
            else { hits.afterSkip(w, h); pos = h.pos + h.len; }
        }
        run.text = doc.text();
    }
};
struct NewPass : HitsPass { NewPass(Doc& d, const std::vector<Entry>& l, const std::set<int>& m, Scoping s = {}) : HitsPass(d, l, m, true, std::move(s)) {} };
struct RefPass : HitsPass { RefPass(Doc& d, const std::vector<Entry>& l, const std::set<int>& m, Scoping s = {}) : HitsPass(d, l, m, false, std::move(s)) {} };

// ---------------------------------------------------------------- OLD: the one-pass flow before OnePassHits
struct OldPass : Walk {
    std::vector<bool> exhausted;
    OldPass(Doc& d, const std::vector<Entry>& l, const std::set<int>& m) : Walk(d, l, m), exhausted(l.size(), false) {}
    void go() {
        Sci::Position pos = 0;
        while (step()) {
            Hit best; size_t w = SIZE_MAX;
            for (size_t i = 0; i < list.size(); ++i) {
                if (exhausted[i]) continue;
                ++run.searches; Hit h; h.pos = doc.find(pos, doc.len(), list[i].find, probeFlags(list[i]), h.len);
                if (h.pos >= 0) { if (best.pos < 0 || h.pos < best.pos) { best = h; w = i; } }
                else if (!list[i].regex && !list[i].wholeWord) exhausted[i] = true;
            }
            if (best.pos < 0) break;
            if (take(w)) {
                ++run.searches; Sci::Position vl; Sci::Position v = doc.find(best.pos, doc.len(), list[w].find, replFlags(list[w]), vl);
                if (v == best.pos && vl == best.len) { const Sci::Position ins = doc.replace(best.pos, best.len, list[w].repl, list[w].regex); ++run.replaced[w]; pos = best.pos + ins; if (best.len == 0 && ins == 0) pos = doc.next(pos); }
                else pos = best.len > 0 ? best.pos + best.len : doc.next(best.pos);
            }
            else pos = best.len > 0 ? best.pos + best.len : doc.next(best.pos);
        }
        run.text = doc.text();
    }
};

// ---------------------------------------------------------------- CLASSIC single-entry replaceAll (N++ processRange, bridge watcher)
static Run classic(const std::string& text, const Entry& e, Sci::Position startPos = 0, DocOptions o = {}) {
    Doc doc(text, o); Run run; run.replaced.assign(1, 0); run.found.assign(1, 0);
    Sci::Position pos = startPos, endRange = doc.len();
    for (;;) {
        Sci::Position len; Sci::Position p = doc.find(pos, doc.len(), e.find, replFlags(e), len);
        if (p < 0) break;
        ++run.found[0];
        const Sci::Position ins = doc.replace(p, len, e.repl, e.regex); ++run.replaced[0];
        if (p + len == endRange) break;
        pos = p + ins; endRange += ins - len;
    }
    run.text = doc.text(); return run;
}

// ---------------------------------------------------------------- PANEL: MultiReplace's own replaceAll loop
// This is what the tab's one-pass option is switched against: with one-pass off every entry gets its own
// run of this loop from the start of the document. A single entry must give the same result either way.
static Run panelReplaceAll(const std::string& text, const Entry& e, Sci::Position startPos = 0, DocOptions o = {}) {
    Doc doc(text, o); Run run; run.replaced.assign(1, 0); run.found.assign(1, 0);
    auto ensureForwardProgress = [&](Sci::Position candidate, Sci::Position lastPos) {   // panel: ensureForwardProgress
        if (candidate > lastPos) return candidate;
        const Sci::Position next = std::max(doc.next(lastPos), lastPos + 1);
        return std::min(next, doc.len());
    };
    Sci::Position pos = startPos;
    for (int steps = 0; steps < 200000; ++steps) {
        Sci::Position len; const Sci::Position p = doc.find(pos, doc.len(), e.find, replFlags(e), len);
        if (p < 0) break;
        ++run.found[0];
        const Sci::Position ins = doc.replace(p, len, e.repl, e.regex); ++run.replaced[0];
        Sci::Position nextPos = p + ins;
        if (len == 0 || nextPos != p) nextPos = ensureForwardProgress(nextPos, p);
        pos = nextPos;
    }
    run.text = doc.text(); return run;
}

// ---------------------------------------------------------------- driver
static int failures = 0; static bool verboseOld = false;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }
static std::string show(std::string s) { for (char& c : s) if (c == '\n') c = '|'; else if (c == '\r') c = '~'; return s.size() > 90 ? s.substr(0, 90) + "..." : s; }
static bool sameRun(const Run& a, const Run& b) { return a.text == b.text && a.replaced == b.replaced && a.found == b.found; }

static void compareNewRef(const std::string& name, const std::string& text, const std::vector<Entry>& list, const std::set<int>& ml = {}, Scoping sc = {}, DocOptions o = {}) {
    Doc a(text, o), b(text, o);
    NewPass n(a, list, ml, sc); n.go();
    RefPass r(b, list, ml, sc); r.go();
    const bool same = sameRun(n.run, r.run);
    CHECK(name + "  [on == off]  searches " + std::to_string(n.run.searches) + " vs " + std::to_string(r.run.searches), same && !n.capped && !r.capped);
    if (!same) { std::printf("      on:  %s\n      off: %s\n", show(n.run.text).c_str(), show(r.run.text).c_str()); }
}
static void compareNewOld(const std::string& name, const std::string& text, const std::vector<Entry>& list, const std::set<int>& ml = {}) {
    Doc a(text), b(text);
    NewPass n(a, list, ml); n.go();
    OldPass t(b, list, ml); t.go();
    const bool same = n.run.text == t.run.text && n.run.replaced == t.run.replaced;
    CHECK(name + "  [on == old]", same);
    if (!same) { std::printf("      on:  %s\n      old: %s\n", show(n.run.text).c_str(), show(t.run.text).c_str()); }
}
static void compareNewClassic(const std::string& name, const std::string& text, const Entry& e) {
    const std::vector<Entry> one{ e }; const std::set<int> none;
    Doc a(text);
    NewPass n(a, one, none); n.go();
    Run c = classic(text, e);
    const bool same = n.run.text == c.text && n.run.replaced == c.replaced;
    CHECK(name + "  [on == classic replaceAll]", same && !n.capped);
    if (!same) { std::printf("      on:      %s\n      classic: %s\n", show(n.run.text).c_str(), show(c.text).c_str()); }
    Doc b(text); OldPass t(b, one, none); t.go();
    if (!same || verboseOld) std::printf("      (old one-pass: %d replaced%s, classic: %d)\n", t.run.replaced[0], t.capped ? ", NO TERMINATION (step cap hit)" : "", c.replaced[0]);
}

static std::string fuzzText(std::mt19937& rng, size_t n, bool csv = false, int exotic = 0) {
    static const char* atoms[] = { "a", "b", "ab", "ba", "x", " ", " ", "\n", "\r\n", "\r\n", "foo", "bar", "1", "_", "\xC3\xA4", "\xC3\x9F", "cat", "dog", "xa", "\r", "yy", "\f", "-", "\t", "x.y", ",", ",", ",\n" };
    static const std::vector<std::string> exoticAtoms = [] {   // exotic but valid UTF-8
        std::vector<std::string> e = { "a", "b", "ab", " ", "\n", "\r\n", "foo", "\f", "\xC2\x85", "\xE2\x80\xA8", "\xE2\x80\xA9", "\xC3\xA4", "\xC3\x84", "\xF0\x9F\x98\x80", "\xE2\x80\x8B",
            "\x7F", "\x0B", "\x1B", "\xEF\xBB\xBF", "x", "ab", "\r", "\t" };
        e.push_back(std::string("\0", 1)); return e;
    }();
    static const std::vector<std::string> brokenAtoms = { "a", "b", " ", "\n", "\xC3\xA4", "\xF0\x9F\x98\x80", "\xA4", "\xC3", "\xE2\x80", "\xFF", "\xF0\x9F", "\x80", "\xC0\x80", "x" };   // invalid sequences
    const size_t k = csv ? 28 : 25;
    std::string s;
    if (exotic == 1) { while (s.size() < n) s += exoticAtoms[rng() % exoticAtoms.size()]; return s; }
    if (exotic == 2) { while (s.size() < n) s += brokenAtoms[rng() % brokenAtoms.size()]; return s; }
    while (s.size() < n) s += atoms[rng() % k]; return s;
}
static const Entry kPool[] = {
    { "a", "" }, { "ab", "X" }, { "ba", "b" }, { "foo", "foobar" }, { "x", "xa" }, { "cat", "dog" }, { "dog", "c" }, { "-", "+" }, { "x.y", "z" },
    { "a", "A", false, true }, { "foo", "F", false, true }, { "\xC3\xA4", "ae" }, { "A", "z", false, false, false }, { "\xC3\xA4", "AE", false, true }, { "x-y", "Q", false, true }, { "foo", "F2", false, true, false },
    { "a+", "<$0>", true }, { "\\bfoo\\b", "F", true }, { "^", "> ", true, false, true, true }, { "$", ";", true, false, true, true }, { "\\b", "|", true, false, true, true },
    { "x*", "y", true, false, true, true }, { "b?", "Q", true, false, true, true }, { "(?=dog)", "!", true, false, true, true }, { "cat(?=\\s)", "C", true }, { "(?<=a)b", "B", true },
    { "(a|)", "-", true, false, true, true }, { "\\Bx", "X", true }, { "[ab]{2}", "2", true }, { "\\d+", "#", true }, { "(f)(o)o", "$2$1", true },
    { "a\\nb", "AB", true }, { "^a", "A", true }, { "b$", "B", true }, { "\\Aa", "S", true }, { "a\\z", "E", true },
    { "(?<!x)a", "N", true }, { "\\R", "\n", true }, { "\\Ga", "G", true }, { "(?<w>ab)", "[$1]", true },
    { "\\B", "#", true, false, true, true }, { "(?=\\r)", "!", true, false, true, true }, { "\\r?$", "!", true, false, true, true }, { "(?<=\\r)", "@", true, false, true, true }, { "\\r\\n", "\n", true }, { "\\n", "\r\n", true },
    { ".", "_", true }, { ".*", "*", true, false, true, true }, { ".*?", "?", true, false, true, true }, { "a*?", "L", true, false, true, true }, { "(?!a)", "!", true, false, true, true }, { "()", "E", true, false, true, true },
    { "\\Z", "Z", true, false, true, true }, { "\\z", "z", true, false, true, true }, { "(?m)^b", "M", true }, { "[[:<:]]", "<", true, false, true, true }, { "\\<", "<", true, false, true, true }, { "\\>", ">", true, false, true, true },
    { "a*+", "P", true, false, true, true }, { "(?>a*)", "T", true, false, true, true }, { "(x)\\1", "D", true }, { "\\s*", "S", true, false, true, true }, { "\\h*$", "H", true, false, true, true }, { "^\\s*$", "-", true, false, true, true },
    { "(?i)FOO", "I", true }, { "b|", "B", true, false, true, true }, { "(?=a|)", "?", true, false, true, true }, { "\\b\\w", "W", true }, { "(?<![a-z])x", "X", true },
    { "a", "aa" }, { "foo", "foofoo" }, { "\\n", "", true }, { "\\r", "", true }, { "b", "\r" }, { "a", "a\n" }, { "x", "\r\n" }, { "o", "$0$0", true }, { "b", "\f" }, { "a", "a," },
    { "y", "", false, true }, { "a\\Kb", "K", true }, { "\\r", "" }, { "\n", "" }, { "\f", "" }, { "\\f", "\n", true },
    { "(?s)a.b", "S", true }, { "FOO|BAR", "fb", true, false, false }, { "\\N+", "N", true }, { "\\h+", "H", true }, { "(?<=\\b)a", "A", true },
    { "\\bx*\\b", "V", true, false, true, true }, { "(?=x*)", "?", true, false, true, true }, { "a(?!\\n)", "A", true }, { "(?<=\\n)a", "L", true }, { "(?=,)", "!", true, false, true, true }, { ",", ";" },
    { "b", "\\n", true }, { "(a)(b)?", "[$1$2]", true }, { "x", "\\U$0\\E", true }, { "a", std::string(300, 'Q') }, { "\\bcat\\b", "", true }, { "^foo", "F", true }, { "\\bx-y\\b", "Q", true },
};
// entries for the exotic sections: raw bytes that break or complete multi-byte sequences, separators the
// engine knows beyond \r and \n, NUL, surrogate pairs, the constructs the probe cannot wrap, regex + whole word
static const std::vector<Entry> kExoticPool = [] {
    std::vector<Entry> p = {
        { "\xC3\xA4", "\xC3\x84" }, { "\xE2\x80\xA8", "\n" }, { "\xC2\x85", "" }, { "\xF0\x9F\x98\x80", ":)" }, { "a", "\xF0\x9F\x98\x80" }, { "\xE2\x80\x8B", "" }, { "\xEF\xBB\xBF", "" },
        { "\\x{2028}", "|", true }, { "\\x85", "N", true }, { "\\R", "|", true }, { "\\R", "", true }, { "^", ">", true, false, true, true }, { "$", ";", true, false, true, true }, { "\\Z", "Z", true, false, true, true },
        { ".", "_", true }, { "\\N", "", true }, { "\\x00", "0", true }, { "[^a]", "?", true }, { "\\W", "w", true }, { "\\bfoo", "F", true }, { "(?i)\xC3\xA4", "ae", true }, { "\\p{L}+", "L", true }, { "\\pL", "l", true },
        { "\\x{1F600}", "E", true }, { "(?!)", "never", true }, { "(?=)", "E", true, false, true, true }, { "a{0}", "0", true, false, true, true }, { "(|a)", "|", true, false, true, true }, { "(?(?=a)ab|b)", "C", true },
        { "(a(?1)?b)", "R", true }, { "(?<n>a)\\k<n>", "D", true }, { "(a|b(?R))", "Q", true }, { "(?x) a  b", "X", true }, { "(?x) a # c", "Y", true }, { "\\Qa.b", "Q", true }, { "foo", "F", true, true },
        { "a\\Kb", "K", true }, { "\\Ka", "K2", true }, { "(?<=\xC3\xA4)b", "B", true }, { "\\b\xC3\xA4", "AE", true }, { "\xC3\xA4\\b", "EA", true }, { "\\v", "V", true }, { "\\f", "F", true }, { "[\\x00-\\x1f]", "C", true },
        { "\\x7F", "D", true }, { "a", std::string(2000, 'z') }, { "\\bx*\\b", "V", true, false, true, true }, { "(?<!\\n)$", "!", true, false, true, true }, { "^(?=.)", "^", true, false, true, true }, { "\\G\\b", "G", true, false, true, true },
        { "\x7F", "" }, { "\x0B", "v" }, { "b", std::string("\0", 1) },
    };
    p.push_back({ std::string("\0", 1), "N" });
    p.push_back({ std::string("a\0", 2), "AN" });
    return p;
}();
// entries that cut multi-byte characters apart: only for the malformed-text section
static const std::vector<Entry> kBrokenPool = {
    { "\xC3", "\xA4" }, { "\xA4", "" }, { "a", "\xC3" }, { "b", "\xA4" }, { "\xC3\xA4", "\xC3" }, { "x", "\xF0\x9F" }, { "\xF0\x9F\x98\x80", "\xF0\x9F" }, { "\xFF", "y" },
    { "\\bb", "B", true }, { "\\b", "|", true }, { "^", ">", true, false, true, true }, { "$", ";", true, false, true, true }, { ".", "_", true }, { "\\w", "W", true }, { "a", "b" }, { "b", "" },
};
static bool hasCtxToken(const Entry& e) { return OnePassHits::contextSensitive(e.find, e.regex, e.wholeWord); }

// level 1: literals; 2: plus regex without context tokens; 3: plus context-dependent entries; 4: everything (nullable too)
static bool inLevel(const Entry& e, int level) {
    if (level >= 4) return true;
    if (e.nullable) return false;
    if (!e.regex && !e.wholeWord) return true;
    if (level == 1) return false;
    if (!hasCtxToken(e)) return true;
    return level >= 3;
}
static Entry fuzzEntry(std::mt19937& rng, int level = 4, int exotic = 0) {
    if (exotic == 2) return kBrokenPool[rng() % kBrokenPool.size()];
    if (exotic && rng() % 2) return kExoticPool[rng() % kExoticPool.size()];
    for (;;) { const Entry& e = kPool[rng() % (sizeof(kPool) / sizeof(kPool[0]))]; if (inLevel(e, level)) return e; }
}
static std::vector<Entry> fuzzList(std::mt19937& rng, int level = 4, size_t maxEntries = 8, int exotic = 0) {
    std::vector<Entry> list; const size_t n = 1 + rng() % maxEntries;
    for (size_t i = 0; i < n; ++i) { Entry e = fuzzEntry(rng, level, exotic); if (rng() % 6 == 0) e.find.clear(); list.push_back(e); }   // empty find = disabled entry
    return list;
}
static void printMismatch(int c, const std::string& tag, const std::string& text, const std::vector<Entry>& list, const std::set<int>& ml, const Run& on, const Run& off) {
    std::printf("  MISMATCH case %d%s%s text=%s\n   list:", c, tag.c_str(), ml.empty() ? "" : " [ml 1,3,4]", show(text).c_str());
    for (auto& e : list) std::printf(" [%s -> %s%s%s]", e.find.c_str(), show(e.repl).c_str(), e.regex ? " rx" : "", e.wholeWord ? " ww" : "");
    size_t k = 0; while (k < on.text.size() && k < off.text.size() && on.text[k] == off.text[k]) ++k;
    const size_t a = k > 30 ? k - 30 : 0;
    auto win = [&](const std::string& t) { return showAll(t.substr(a, 70)); };
    std::printf("\n   first diff at %zu\n   on:  ...%s\n   off: ...%s\n   counts on:", k, win(on.text).c_str(), win(off.text).c_str());
    for (int v : on.replaced) std::printf(" %d", v);
    std::printf("  off:");
    for (int v : off.replaced) std::printf(" %d", v);
    std::printf("\n");
}

// one fuzz section: `cases` random texts and lists, on vs off (and vs classic for single full-scope entries)
// exotic: 0 none, 1 exotic but valid UTF-8, 2 malformed UTF-8 (informational: see the M4 comment)
struct Variant { OnePassHits::Scope scope = OnePassHits::Scope::Full; bool extEdits = false; DocOptions doc; bool randomStart = false; int level = 4; int exotic = 0; size_t maxEntries = 8; };
static int fuzzSection(const char* name, std::mt19937& rng, int cases, size_t maxLen, Variant v, int* cappedOut = nullptr) {
    int bad = 0;
    for (int c = 0; c < cases; ++c) {
        const bool csv = v.scope == OnePassHits::Scope::Column;
        std::string text = fuzzText(rng, 20 + rng() % maxLen, csv, v.exotic);
        if (v.doc.codePage == 0) for (char& ch : text) if ((unsigned char)ch >= 0x80 && rng() % 2) ch = (char)(0xC0 + rng() % 0x3F);   // ANSI: single high bytes
        const std::vector<Entry> list = fuzzList(rng, v.level, v.maxEntries, v.exotic);
        std::set<int> ml; if (rng() % 5 == 0) ml = { 1, 3, 4 };
        Scoping sc; sc.scope = v.scope;
        if (csv) { const int k = 1 + rng() % 4; for (int i = 0; i < k; ++i) sc.columns.insert(1 + (int)(rng() % 5)); }
        if (v.scope == OnePassHits::Scope::Selection) {
            const Sci::Position L = (Sci::Position)text.size();
            for (int i = 0; i < 2; ++i) { Sci::Position a = rng() % (L + 1), b = rng() % (L + 1); if (a > b) std::swap(a, b); sc.sel.push_back({ a, b }); }
            std::sort(sc.sel.begin(), sc.sel.end()); if (sc.sel[0].second > sc.sel[1].first) sc.sel[1].first = sc.sel[0].second;   // panel keeps ranges disjoint
        }
        const bool dyn = rng() % 2 == 0;
        std::set<size_t> edits; if (v.extEdits) { edits.insert(2 + rng() % 5); edits.insert(9 + rng() % 9); }
        Doc a(text, v.doc), b(text, v.doc);
        const bool bwd = rng() % 2 == 0;
        if (bwd) { a.backwardSearch(); b.backwardSearch(); }   // half the runs start like a run after Find Previous
        Sci::Position start = 0; if (v.randomStart) start = a.d->MovePositionOutsideChar(rng() % (a.len() + 1), 1, false);
        NewPass on(a, list, ml, sc); on.dynamicRepl = dyn; on.extEdits = edits; on.startPos = start; on.go();
        RefPass off(b, list, ml, sc); off.dynamicRepl = dyn; off.extEdits = edits; off.startPos = start; off.go();
        bool okClassic = true;
        if (v.scope == OnePassHits::Scope::Full && !v.extEdits && !dyn && list.size() == 1 && ml.empty() && !list[0].find.empty()) {
            const Run cl = classic(text, list[0], start, v.doc); okClassic = cl.text == on.run.text && cl.replaced == on.run.replaced;
            if (!okClassic) std::printf("  (classic differs: %d vs on %d)\n", cl.replaced[0], on.run.replaced[0]);
        }
        if (cappedOut && (on.capped || off.capped)) ++*cappedOut;
        if (!sameRun(on.run, off.run) || on.capped || off.capped || !okClassic) {
            if (++bad <= (cappedOut ? 2 : 5)) {
                printMismatch(c, std::string(" [") + name + (dyn ? ", dynamic repl" : "") + (v.extEdits ? ", ext edits" : "") + (start ? ", start " + std::to_string((long long)start) : "") + (bwd ? ", after backward search" : "") + "]", text, list, ml, on.run, off.run);
                std::printf("   found on:"); for (int x : on.run.found) std::printf(" %d", x); std::printf("  off:"); for (int x : off.run.found) std::printf(" %d", x);
                std::printf("  capped %d/%d  columns:", (int)on.capped, (int)off.capped); for (int k : sc.columns) std::printf(" %d", k); std::printf("  sel:"); for (auto& r : sc.sel) std::printf(" [%lld,%lld)", (long long)r.first, (long long)r.second); std::printf("\n");
                if (getenv("TRACE")) { traceSteps = true; Doc a2(text, v.doc), b2(text, v.doc); NewPass n2(a2, list, ml, sc); n2.dynamicRepl = dyn; n2.extEdits = edits; n2.startPos = start; n2.go(); RefPass r2(b2, list, ml, sc); r2.dynamicRepl = dyn; r2.extEdits = edits; r2.startPos = start; r2.go(); traceSteps = false; std::printf("   bytes:"); for (unsigned char ch : text) std::printf(" %02x", ch); std::printf("\n"); }
            }
        }
    }
    return bad;
}

// K) the properties the bookkeeping rests on, measured against the engine instead of assumed.
// Every construct below (the pool's plus the exotic ones) is run over probe texts with the
// walk's own flags, at every byte position, case-sensitive and insensitive:
//   K1 a hit at P depends on the byte before P only for constructs the class classifies
//      (contextSensitive or neverRemembered), otherwise a remembered hit could survive a change before it
//   K2 gate: where the plain search finds a hit at P the anchored \G(?:...) search finds one at P too,
//      so the probe never blocks a real hit (the other direction is harmless since F3, keep() decides)
//   K3 class: two bytes of the same contextClass before P give the same hit at P, because a class
//      change is the only trigger of the probe (NonAscii always triggers it)
//   K4 none: no hit from p means no hit from any p' > p, because an entry without a hit is not searched again
//   K5 stability: the hit (q, l) found from p is the hit found from every p' in (p, q], because a remembered
//      hit is verified by a search from q while the walk moves through (p, q]
//   K6 the bridge treats an empty search range as backward after a backward search (_lastDirection);
//      the result of FindText(p, p) must not depend on it
static const std::vector<std::string> kConstructs = {
    // the pool's constructs
    "\\b", "\\B", "\\bfoo", "\\Ba", "\\<a", "a\\>", "[[:<:]]a", "^a", "\\Aa", "\\`a", "^", "(?m)^a",
    "(?<=a)b", "(?<!a)b", "(?<=\\s)a", "a\\Kb", "\\Ga", "\\G",
    "$", "a$", "\\Z", "\\z", "\\'", "(?=a)", "(?!a)", "a(?=b)", "a(?!b)",
    "a", "abc", "a+", "a*", "a?", "[abc]", "[^abc]", ".", ".*", "a{1,3}", "(a|b)", "(?:ab)", "(a)\\1",
    "\\d", "\\w", "\\s", "\\S", "\\W", "\\D", "\\h", "\\R", "[[:alpha:]]", "[[:space:]]", "[[:word:]]",
    "(?i)A", "(?s)a.b", "(?>a+)", "a++", "a*?", "(?<n>a)", "\\Qa.b\\E", "\\x61", "[\\d-]", "\\p{L}", "a(?#comment)b",
    "(?=a|)", "()", "(?:)", "b|", "\\N", "\\141",
    // conditionals, recursion, back references in every spelling, branch reset
    "(?(1)a|b)", "(a)?(?(1)b|c)", "(?(?=a)ab|b)", "(?(?<=a)b|c)", "(?(?!a)b|a)",
    "(a(?1)?b)", "(?<n>a)(?&n)", "(a|b(?R))", "(?:a|b(?0))", "(?P<n>a)(?P=n)", "(?P<n>a)(?P>n)",
    "(a)\\g{1}", "(a)\\g{-1}", "(a)\\g1", "(?<n>a)\\k<n>", "(?<n>a)\\k'n'", "(?<n>a)\\g{n}", "(a)(b)\\2\\1",
    "(?|a|b)", "(?|(a)|(b))\\1",
    // inline flags, free-spacing mode with its comments, quoting
    "(?x) a b", "(?x) a # comment", "(?x: a # c\n)", "(?ix)A", "(?-x)a", "(?i:A)b", "(?-i:a)", "(?s:.)", "(?m:^a)", "(?-m:^a)", "(?m)$",
    "\\Qa.b", "\\Q", "a\\Q", "\\E", "\\Qa\\E\\Qb",
    // escapes and classes
    "\\x{61}", "\\x{e4}", "\\x{2028}", "\\x85", "\\0", "\\cA", "\\e", "\\a", "\\t", "\\n", "\\r", "\\f", "\\v", "\\x00",
    "\\pL", "\\P{L}", "\\p{Lu}", "\\p{^L}", "[\\p{L}]", "[[:^alpha:]]", "[[=a=]]", "[[.a.]]", "[[.space.]]", "[a-\\x{ff}]", "[\\x00-\\x1f]",
    "\\N+", "\\H", "\\V", "\\X", "\\C",
    "[^a]", "[^\\n]", "[\\n]", "[\\r\\n]", "[\\f]", ".+", ".*?", ".?", "\\w+", "\\d+", "[[:alnum:]]", "[[:punct:]]", "[[:cntrl:]]", "[[:blank:]]",
    "[[:graph:]]", "[[:print:]]", "[[:lower:]]", "[[:upper:]]", "[[:xdigit:]]", "[[:unicode:]]",
    // nullable, never matching, degenerate and expensive
    "(?!)", "(?=)", "a{0}", "a{0,0}", "(|a)", "(?:|a)", "(?>)", "()*", "(a*)*", "(a*)+", "(a|ab)(c|bcd)(d*)", "(a+)+b", "(?:a?){3}a{3}",
    "a?+", "a{1,2}+", "(?>a+)b", "(?>a*)", "a|b", "|b", "a|", "(a|)",
    // anchors and lookarounds in combination
    "^$", "^\\s*$", "\\s+$", "$\\n", "\\r?$", "(?=\\r)", "\\R$", "^\\R", "(?<=^)a", "(?<=$)", "(?<=\\b)a", "(?<=\\r)", "(?<=\\n)a",
    "(?<=[ab])c", "(?<=a|bc)d", "(?<=(a))\\1", "\\Ka", "\\G^a", "\\G$", "(?:\\G|x)a", "a\\Z", "a\\z", "a\\'", "\\bx*\\b", "(?=x*)", "\\b\\w",
    // non-ASCII, separators the engine knows beyond \r and \n, invalid bytes
    "\xC3\xA4", "(?i)\xC3\xA4", "\xC3\xA4+", "[\xC3\xA4]", "\\x{e4}+", "\xA4", "\xC3", "\xC2\x85", "\xE2\x80\xA8", "\\x{2028}|\\x{2029}",
    "[\\x{2028}\\x{2029}\\x{85}]", "\\x{1F600}", "\xF0\x9F\x98\x80", "\xC3\xA4\\b", "\\b\xC3\xA4", "(?i)\\bA", "(?i)[a-z]", "(?i)ABC", "[A-Z]", "(?i)[[:upper:]]",
};
static const std::vector<std::string> kLiterals = { "a", "ab", "b c", " ", "\n", "\r\n", "\xC3\xA4", "\xA4", "\xC3", std::string("\0", 1), "A" };

static std::vector<std::string> probeTexts() {
    std::vector<std::string> t = {
        "abc def", "ab\ncd", "foo bar", "a1_2", "  x ", "ab\r\ncd", "\nab", "aab", "b c", "1 2 3", "9ab", " ab", "\tab", ".ab", "-ab", "A B", "\fab",
        "\xC3\xA4" "b", "_ab", "ab.cd", "a.b", "b\nc", "x\ty", "a", "ab\n", "", "a\r\n\r\nb\r\n", "a\rb\rc\r", "\n\n\n", "\r\n", "\f\f", "ab\f", "a\fb\fc",
        "\xC2\x85" "ab\xC2\x85", "x\xE2\x80\xA8y\xE2\x80\xA9z", "\xC3\xA4\xC3\xA4 b", "\xF0\x9F\x98\x80" "a", "\xA4" "b", "\xC3" "b", "a\xE2\x80" "b", "\xFF\xFE",
        "aaaa bbbb", "ababab", "x\vy", "\x7F" "a", "foo\n\nbar\n", "a b\r\nc d\r\n", "(a(b)c)", "a-b_c.d", "AbC", "aAaA", "abcd", "aabcbcd", "a\n\ra",
    };
    t.push_back(std::string("a\0b\0", 4));
    return t;
}
static const int kRx = (int)FindOption::RegExp | F_POSIX | F_SKIPCRLF | F_NOTAFTER | F_ALLOWATSTART;

struct KResult { int gaps = 0, gates = 0, classes = 0, nones = 0, stabs = 0, quirks = 0; int dependent = 0, independent = 0, untested = 0, unsupported = 0, f3 = 0; int invalidNones = 0, invalidStabs = 0, invalidGates = 0, invalidQuirks = 0; };
static bool validUtf8(const std::string& s) {
    for (size_t i = 0; i < s.size();) {
        const unsigned char c = (unsigned char)s[i]; size_t n = c < 0x80 ? 1 : (c >> 5) == 6 ? 2 : (c >> 4) == 14 ? 3 : (c >> 3) == 30 ? 4 : 0;
        if (n == 0 || i + n > s.size()) return false;
        for (size_t j = 1; j < n; ++j) if (((unsigned char)s[i + j] & 0xC0) != 0x80) return false;
        i += n;
    }
    return true;
}

static Sci::Position kFind(Doc& d, Sci::Position from, Sci::Position to, const std::string& pat, int flags, Sci::Position& len) {
    len = (Sci::Position)pat.size();
    return d.d->FindText(from, to, pat.c_str(), fo(flags), &len);   // raw: -1 none, -2 invalid, -3 exception
}

// K1 and K3: the hit at P = prefix length, over prefixes of every class
static bool kDependence(KResult& k, const std::string& pat, bool regex, bool wholeWord, int flags) {
    static const std::vector<std::string> prefixes = [] {
        std::vector<std::string> p = { "", "a", "1", "_", "-", " ", "\t", "\n", "\r", "\f", "\v", "\x80", "\xC3", "!", ".", ",", ")", "\\", "x", "A", "\x7F", "\x1F", "\xFF",
            "\xC2\x85", "\xE2\x80\xA8", "\xE2\x80\xA9", "\xC3\xA4", "\xC3\x84", "\xE2\x80\x8B", "\xEF\xBB\xBF", "\xF0\x9F\x98\x80", "\xE2\x80", "\r\n", "ab", "a " };
        p.push_back(std::string("\0", 1));
        return p;
    }();
    static const char* tails[] = { "abc def", "ab\ncd", "foo bar", "a1_2", "  x ", "ab\r\ncd", "\nab", "aab", "b c", "1 2 3", "9ab", " ab", "\tab", ".ab", "-ab", "A B", "\fab", "\xC3\xA4" "b", "_ab", "ab.cd", "a.b", "b\nc", "x\ty", "a", "ab\n", "", "\xC2\x85" "a", "\xE2\x80\xA8" "a", "ABC", "\xA4" "b" };
    bool dependent = false, exercised = false, unsupported = false, classGap = false;
    const bool covered = OnePassHits::contextSensitive(pat, regex, wholeWord) || OnePassHits::neverRemembered(pat, regex);
    for (const char* tail : tails) {
        std::map<OnePassHits::Ctx, std::pair<bool, long long>> byClass;   // class -> first result seen (hit length or -1)
        bool first = true; long long ref = -1;
        for (const std::string& pre : prefixes) {
            Doc d(pre + tail);
            const Sci::Position P = (Sci::Position)pre.size();
            if (d.d->MovePositionOutsideChar(P, 1, false) != P) continue;   // prefix and tail form one character: P is not a position the walk can be at
            Sci::Position len; const Sci::Position r = kFind(d, P, d.len(), pat, flags, len);
            if (r == -2 || r == -3) { unsupported = true; break; }
            const long long res = (r == P) ? (long long)len : -1;
            if (res >= 0) exercised = true;
            if (first) { ref = res; first = false; } else if (res != ref) dependent = true;
            // K3: ASCII prefixes only; Bof and NonAscii always trigger the probe
            if (pre.size() == 1 && (unsigned char)pre[0] < 0x80 && regex) {
                const OnePassHits::Ctx c = OnePassHits::contextClass((unsigned char)pre[0]);
                auto it = byClass.find(c);
                if (it == byClass.end()) byClass[c] = { true, res };
                else if (it->second.second != res) classGap = true;
            }
        }
        if (unsupported) break;
    }
    if (unsupported) { ++k.unsupported; return false; }
    if (dependent) { ++k.dependent; if (!covered) { ++k.gaps; std::printf("  K1 GAP: '%s' depends on the byte before the match but is classified as plain\n", show(pat).c_str()); } }
    else if (!exercised) ++k.untested;
    else if (!covered) ++k.independent;
    if (classGap && covered && !OnePassHits::neverRemembered(pat, regex)) { ++k.classes; std::printf("  K3 CLASS GAP: '%s' gives different hits after two bytes of the same class\n", show(pat).c_str()); }
    return dependent;
}

// K2, K4, K5, K6: every start position of every probe text. Texts with invalid UTF-8 are counted
// apart: there the bridge's iterator decodes differently depending on where it starts, so the
// engine itself answers inconsistently and nothing built on it can be consistent (see Stage 6).
static void kPositions(KResult& k, const std::string& pat, bool regex, int flags, const std::vector<std::string>& texts, bool dependent) {
    const bool uncached = OnePassHits::neverRemembered(pat, regex);
    // only context-sensitive entries get the anchored probe, and only a construct whose hit at P can be
    // created by a change before P needs it: for the others a hit at the replacement end existed before
    const bool probed = regex && !uncached && OnePassHits::contextSensitive(pat, regex, false) && dependent;
    const int gates0 = k.gates, stabs0 = k.stabs, nones0 = k.nones;
    const std::string anchored = "\\G(?:" + pat + ")";
    int shownGate = 0, shownNone = 0, shownStab = 0, shownQuirk = 0;
    for (const std::string& text : texts) {
        const bool valid = validUtf8(text);
        Doc d(text); const Sci::Position L = d.len();
        std::vector<Sci::Position> pos(L + 1), len(L + 1), apos(L + 1), alen(L + 1);
        bool unsupported = false;
        for (Sci::Position p = 0; p <= L; ++p) {
            pos[p] = kFind(d, p, L, pat, flags, len[p]);
            if (pos[p] == -2 || pos[p] == -3) { unsupported = true; break; }
            if (probed) apos[p] = kFind(d, p, L, anchored, flags, alen[p]);
        }
        if (unsupported) continue;
        for (Sci::Position p = 0; p <= L; ++p) {
            if (probed) {
                if (pos[p] == p && apos[p] != p) { ++(valid ? k.gates : k.invalidGates); if (valid && ++shownGate <= 2) std::printf("  K2 GATE: '%s' plain hit at %lld of '%s' but \\G(?:...) gives %lld\n", show(pat).c_str(), (long long)p, show(text).c_str(), (long long)apos[p]); }
                if (apos[p] == p && pos[p] != p) ++k.f3;
            }
            if (!uncached && pos[p] < 0) {
                for (Sci::Position q = p + 1; q <= L; ++q) if (pos[q] >= 0) { ++(valid ? k.nones : k.invalidNones); if (!valid && getenv("SHOWINVALID")) { std::printf("  K4 (malformed) '%s' nothing from %lld, [%lld,+%lld] from %lld; bytes:", show(pat).c_str(), (long long)p, (long long)pos[q], (long long)len[q], (long long)q); for (unsigned char ch : text) std::printf(" %02x", ch); std::printf("\n"); } if (valid && ++shownNone <= 2) std::printf("  K4 NONE: '%s' finds nothing from %lld but [%lld,+%lld] from %lld in '%s'\n", show(pat).c_str(), (long long)p, (long long)pos[q], (long long)len[q], (long long)q, show(text).c_str()); break; }
            }
            if (!uncached && pos[p] >= 0) {   // an uncached hit (\K, lookbehind, \G) is verified from the walk position it was found from
                for (Sci::Position q = p + 1; q <= pos[p]; ++q)
                    if (pos[q] != pos[p] || len[q] != len[p]) { ++(valid ? k.stabs : k.invalidStabs); if (valid && ++shownStab <= 2) std::printf("  K5 STABILITY: '%s' from %lld gives [%lld,+%lld], from %lld gives [%lld,+%lld] in '%s'\n", show(pat).c_str(), (long long)p, (long long)pos[p], (long long)len[p], (long long)q, (long long)pos[q], (long long)len[q], show(text).c_str()); break; }
            }
        }
        // K6: empty ranges after a backward search
        Doc f(text), b(text);
        { Sci::Position l; kFind(b, L, 0, "a", flags, l); }   // sets the bridge's _lastDirection to -1
        for (Sci::Position p = 0; p <= L; ++p) {
            Sci::Position lf, lb; const Sci::Position rf = kFind(f, p, p, pat, flags, lf), rb = kFind(b, p, p, pat, flags, lb);
            if (rf != rb || (rf >= 0 && lf != lb)) { ++(valid ? k.quirks : k.invalidQuirks); if (valid && ++shownQuirk <= 2) std::printf("  K6 QUIRK: '%s' at empty range %lld of '%s': forward %lld/%lld, after a backward search %lld/%lld\n", show(pat).c_str(), (long long)p, show(text).c_str(), (long long)rf, (long long)lf, (long long)rb, (long long)lb); }
        }
    }
    if (k.gates > gates0 || k.stabs > stabs0 || k.nones > nones0) std::printf("  '%s': K2 %d, K4 %d, K5 %d violations in valid text\n", show(pat).c_str(), k.gates - gates0, k.nones - nones0, k.stabs - stabs0);
}

// K7: the probe for a literal entry searches a window (4x the find text plus 4 bytes) instead of the
// whole document, so that a gate that fails costs nothing. The window must never hide a match that
// starts exactly at the probe position - case folding can make a match longer than its pattern.
static int probeWindow(const std::vector<std::string>& texts) {
    int bad = 0, checks = 0;
    for (const std::string& lit : kLiterals) {
        if (lit.empty()) continue;
        for (int mode = 0; mode < 4; ++mode) {
            const int flags = ((mode & 1) ? (int)FindOption::MatchCase : 0) | ((mode & 2) ? (int)FindOption::WholeWord : 0);
            for (const std::string& text : texts) {
                Doc d(text); const Sci::Position L = d.len();
                for (Sci::Position p = 0; p <= L; ++p) {
                    Sci::Position lw, lf;
                    const Sci::Position win = std::min(L, p + 4 * (Sci::Position)lit.size() + 4);
                    const bool inWindow = kFind(d, p, win, lit, flags, lw) == p;
                    const bool inFull = kFind(d, p, L, lit, flags, lf) == p;
                    ++checks;
                    if (inWindow != inFull) { if (++bad <= 5) std::printf("  K7 WINDOW: '%s' at %lld of '%s' (flags %d): window %d, full %d\n", show(lit).c_str(), (long long)p, show(text).c_str(), flags, (int)inWindow, (int)inFull); }
                }
            }
        }
    }
    std::printf("  K7 probe window: %d positions checked, %d disagreements with a search over the whole document\n", checks, bad);
    return bad;
}

// K8: the classification is a textual token test over a hand-written list of constructs. A construct
// nobody thought of is the residual. Random patterns from a grammar cover what the list does not:
// each one is measured the same way as K1 - does a hit at P change with the byte before P, and does
// the classification know? Patterns this Boost build rejects are skipped, not counted as safe.
static std::string randomPattern(std::mt19937& rng, int depth = 0) {
    static const char* atoms[] = { "a", "b", "x", "ab", ".", "\\w", "\\d", "\\s", "\\S", "\\W", "\\h", "\\R", "\\N", "[ab]", "[^ab]", "[a-z]", "[[:alpha:]]", "[[:space:]]",
        "\\b", "\\B", "^", "$", "\\A", "\\Z", "\\z", "\\<", "\\>", "[[:<:]]", "[[:>:]]", "\\`", "\\'", "\\G", "\\K", "\xC3\xA4", "\\x61", "\\n", "\\r", "\\f", "\\t" };
    static const char* quant[] = { "*", "+", "?", "{0,2}", "{1,3}", "*?", "+?", "??", "*+", "++" };
    const int k = rng() % (depth >= 2 ? 6 : 12);
    switch (k) {
    case 0: case 1: case 2: case 3: case 4: case 5: return atoms[rng() % (sizeof(atoms) / sizeof(atoms[0]))];
    case 6: return randomPattern(rng, depth + 1) + quant[rng() % 10];
    case 7: return "(" + randomPattern(rng, depth + 1) + ")";
    case 8: return "(?:" + randomPattern(rng, depth + 1) + "|" + randomPattern(rng, depth + 1) + ")";
    case 9: return randomPattern(rng, depth + 1) + randomPattern(rng, depth + 1);
    case 10: return (rng() % 2 ? "(?=" : "(?!") + randomPattern(rng, depth + 1) + ")";
    default: return (rng() % 2 ? "(?<=" : "(?<!") + std::string(atoms[rng() % 8]) + ")" + randomPattern(rng, depth + 1);
    }
}
static int randomClassification(std::mt19937& rng, int count) {
    KResult k;
    int measured = 0;
    for (int i = 0; i < count; ++i) {
        const std::string pat = randomPattern(rng);
        const int before = k.unsupported;
        kDependence(k, pat, true, false, kRx | (int)FindOption::MatchCase);
        if (k.unsupported == before) ++measured;
    }
    std::printf("  K8 random patterns: %d generated, %d measured (%d rejected by this Boost), %d depend on the byte before, %d gaps\n",
        count, measured, k.unsupported, k.dependent, k.gaps);
    return k.gaps;
}

static int engineProperties() {
    KResult k; const std::vector<std::string> texts = probeTexts();
    for (const std::string& pat : kConstructs) {
        const bool dependent = kDependence(k, pat, true, false, kRx | (int)FindOption::MatchCase);
        for (int ci = 0; ci < 2; ++ci) kPositions(k, pat, true, kRx | (ci ? 0 : (int)FindOption::MatchCase), texts, dependent);
    }
    for (const std::string& lit : kLiterals) {
        for (int ww = 0; ww < 2; ++ww) {
            const int flags = (int)FindOption::MatchCase | (ww ? (int)FindOption::WholeWord : 0);
            const bool dependent = kDependence(k, lit, false, ww != 0, flags);
            kPositions(k, lit, false, flags, texts, dependent);
        }
    }
    std::printf("  %zu regex constructs, %zu literals: %d depend on the byte before (K1 gaps %d), %d proven independent and remembered, %d untested, %d unsupported by this Boost\n",
        kConstructs.size(), kLiterals.size(), k.dependent, k.gaps, k.independent, k.untested, k.unsupported);
    std::printf("  valid UTF-8: K2 gate violations %d (harmless F3-direction disagreements %d), K3 class gaps %d, K4 none violations %d, K5 stability violations %d, K6 direction quirks %d\n",
        k.gates, k.f3, k.classes, k.nones, k.stabs, k.quirks);
    std::printf("  invalid UTF-8 (engine inconsistency, informational): K2 %d, K4 %d, K5 %d, K6 %d\n", k.invalidGates, k.invalidNones, k.invalidStabs, k.invalidQuirks);
    int extra = probeWindow(texts);
    std::mt19937 pr(0xC0FFEE);
    extra += randomClassification(pr, 4000);
    return k.gaps + k.gates + k.classes + k.nones + k.stabs + k.quirks + extra;
}

// Every pattern the harness knows, as a single entry, on every probe text: the one-pass walk must give the
// same text and the same count as Notepad++'s Replace All (processRange) and as the panel's own replaceAll
// loop. No pattern and no text is excluded - the \K defect (F7) survived four rounds behind one exclusion.
static int singleEntryEquivalence(const std::vector<Entry>& pool, const std::vector<std::string>& texts, const char* what) {
    int bad = 0, runs = 0, differsPanelClassic = 0;
    const std::set<int> none;
    for (const Entry& e : pool) {
        if (e.find.empty()) continue;   // an empty find is a disabled entry, it never reaches the walk
        for (const std::string& text : texts) {
            for (int ci = 0; ci < 2; ++ci) {
                DocOptions o; o.codePage = ci ? 0 : kCpUtf8;
                { Doc probe(text, o); Sci::Position l = (Sci::Position)e.find.size();
                  const Sci::Position r = probe.d->FindText(0, probe.len(), e.find.c_str(), fo(replFlags(e)), &l);
                  if (r == -2 || r == -3) continue; }        // this Boost build rejects the pattern: nothing to compare
                ++runs;
                Doc a(text, o); const std::vector<Entry> one{ e };
                NewPass n(a, one, none); n.go();
                const Run cl = classic(text, e, 0, o), pl = panelReplaceAll(text, e, 0, o);
                if (cl.text != pl.text || cl.replaced != pl.replaced) ++differsPanelClassic;
                // the reference is Notepad++'s own Replace All. Where the panel's replaceAll loop disagrees
                // with it the panel loop is reported separately: that is not the one-pass walk's business.
                if (cl.text != pl.text || cl.replaced != pl.replaced) {
                    if (++differsPanelClassic <= 3)
                        std::printf("  NOTE the panel's own replaceAll differs from Notepad++ Replace All (not a one-pass question): '%s' -> '%s'%s on '%s': panel %d replaced, Notepad++ %d, one-pass %d; text %s\n",
                            show(e.find).c_str(), show(e.repl).c_str(), e.regex ? " rx" : "", show(text).c_str(), pl.replaced[0], cl.replaced[0], n.run.replaced[0],
                            cl.text == pl.text ? "identical" : "DIFFERENT");
                }
                if (n.run.text != cl.text || n.run.replaced != cl.replaced || n.capped) {
                    if (++bad <= 10)
                        std::printf("  DIFF '%s' -> '%s'%s%s on '%s' (cp %d): one-pass %d replaced '%s', Notepad++ Replace All %d '%s'%s\n",
                            show(e.find).c_str(), show(e.repl).c_str(), e.regex ? " rx" : "", e.wholeWord ? " ww" : "", show(text).c_str(), o.codePage,
                            n.run.replaced[0], show(n.run.text).c_str(), cl.replaced[0], show(cl.text).c_str(), n.capped ? " NO TERMINATION" : "");
                }
            }
        }
    }
    std::printf("  %s: %d runs, %d differences from Notepad++ Replace All; the panel's own replaceAll loop differs from Notepad++ in %d of them (pre-existing, see the notes above)\n", what, runs, bad, differsPanelClassic);
    return bad;
}

static std::string prose(std::mt19937& rng, size_t bytes) {
    static const char* words[] = { "the", "quick", "brown", "fox", "jumps", "over", "lazy", "dog", "lorem", "ipsum", "dolor", "sit", "amet", "value", "index", "count", "buffer", "string", "result", "error", "line", "file", "path", "name", "data", "item", "list", "entry", "search", "replace", "text", "match", "flag", "mode", "column", "table", "row", "cell", "user", "input" };
    std::string s; size_t col = 0;
    while (s.size() < bytes) { const char* w = words[rng() % 40]; s += w; col += strlen(w); if (col > 70) { s += "\r\n"; col = 0; } else { s += ' '; ++col; } if (rng() % 300 == 0) s += "TAG_" + std::to_string(rng() % 100) + " "; }
    return s;
}
static void bench(size_t kb) {
    std::mt19937 rng(42); const std::string text = prose(rng, kb * 1024);
    std::printf("text: %zu bytes (CRLF prose)\n", text.size());
    struct Case { const char* name; std::vector<Entry> list; bool runOld; };
    const std::vector<Case> cases = {
        { "10 literals, all frequent", { { "the", "THE" }, { "fox", "wolf" }, { "dog", "cat" }, { "value", "val" }, { "index", "idx" }, { "count", "cnt" }, { "buffer", "buf" }, { "string", "str" }, { "result", "res" }, { "error", "err" } }, true },
        { "5 frequent literals + 3 rare regex (every ~3000 chars) + 2 regex without hits", { { "the", "THE" }, { "fox", "wolf" }, { "dog", "cat" }, { "value", "val" }, { "index", "idx" }, { "TAG_1\\d", "T", true }, { "TAG_2\\d", "U", true }, { "\\bTAG_3\\d\\b", "V", true }, { "\\bTODO\\b", "X", true }, { "foo\\d{5}", "Y", true } }, true },
        { "1 frequent literal + 1 regex without hits", { { "the", "THE" }, { "\\bTODO\\b", "X", true } }, true },
        { "^ -> '> ' + 2 frequent literals (empty match on every line)", { { "^", "> ", true }, { "the", "THE" }, { "dog", "cat" } }, false },
        { "\\s+$ -> '' + 3 literals (context-sensitive entry)", { { "\\s+$", "", true }, { "the", "THE" }, { "fox", "wolf" }, { "dog", "cat" } }, true },
    };
    const std::set<int> none;
    for (const Case& c : cases) {
        auto ms = [](auto&& f) { const auto t0 = std::chrono::steady_clock::now(); f(); return std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - t0).count(); };
        int rOn = 0, rOff = 0, rOld = 0; std::string tOn, tOff; size_t sOn = 0, sOff = 0;
        const double mOn = ms([&] { Doc d(text); NewPass p(d, c.list, none); p.go(); for (int v : p.run.replaced) rOn += v; tOn = p.run.text; sOn = p.run.searches; });
        const double mOff = ms([&] { Doc d(text); RefPass p(d, c.list, none); p.go(); for (int v : p.run.replaced) rOff += v; tOff = p.run.text; sOff = p.run.searches; });
        const double mOld = c.runOld ? ms([&] { Doc d(text); OldPass p(d, c.list, none); p.go(); for (int v : p.run.replaced) rOld += v; }) : -1;
        std::printf("\n%s\n  on  %8.0f ms  %d replaced  %zu searches%s\n  off %8.0f ms  %d replaced  %zu searches\n", c.name, mOn, rOn, sOn, tOn == tOff ? "" : "  TEXT DIFFERS FROM off", mOff, rOff, sOff);
        if (c.runOld) std::printf("  old %8.0f ms  %d replaced\n", mOld, rOld);
    }
}

int main(int argc, char** argv) {
    if (argc > 2 && std::string(argv[1]) == "trace") {
        std::string text = argv[2]; for (char& c : text) { if (c == '|') c = '\n'; else if (c == '~') c = '\r'; }
        std::vector<Entry> list; for (int i = 3; i + 2 < argc; i += 3) { Entry e; e.find = argv[i]; e.repl = argv[i + 1]; for (char& c : e.repl) { if (c == '|') c = '\n'; else if (c == '~') c = '\r'; } e.regex = argv[i + 2][0] == 'r'; e.wholeWord = argv[i + 2][0] == 'w'; list.push_back(e); }
        traceSteps = true; Doc a(text), b(text); std::set<int> ml; if (getenv("ML")) { ml.insert(1); ml.insert(3); ml.insert(4); }
        NewPass n(a, list, ml); n.go(); RefPass r(b, list, ml); r.go();
        std::printf("on:  %s\noff: %s\n%s\n", showAll(n.run.text).c_str(), showAll(r.run.text).c_str(), n.run.text == r.run.text ? "SAME" : "DIFF");
        return 0;
    }
    if (argc > 1 && std::string(argv[1]) == "bench") { bench(argc > 2 ? std::stoul(argv[2]) : 300); return 0; }
    if (argc > 3 && std::string(argv[1]) == "find") {   // find '<text, | = LF, ~ = CR>' pattern [from [to]] [l|w|i]: raw engine results, plain and anchored
        std::string text = argv[2]; for (char& c : text) { if (c == '|') c = '\n'; else if (c == '~') c = '\r'; }
        const std::string pat = argv[3]; const char mode = argc > 6 ? argv[6][0] : 'r';
        Doc d(text); const Sci::Position from = argc > 4 ? std::stoll(argv[4]) : 0, to = argc > 5 ? std::stoll(argv[5]) : d.len();
        const int flags = mode == 'l' ? (int)FindOption::MatchCase : mode == 'w' ? ((int)FindOption::MatchCase | (int)FindOption::WholeWord) : mode == 'i' ? kRx : (kRx | (int)FindOption::MatchCase);
        Sci::Position len; const Sci::Position r = kFind(d, from, to, pat, flags, len);
        std::printf("plain    '%s' in [%lld,%lld] of '%s' (%lld bytes): %lld +%lld\n", pat.c_str(), (long long)from, (long long)to, show(text).c_str(), (long long)d.len(), (long long)r, (long long)len);
        if (mode != 'l' && mode != 'w') { Sci::Position al; const Sci::Position ar = kFind(d, from, to, "\\G(?:" + pat + ")", flags, al); std::printf("anchored \\G(?:%s): %lld +%lld\n", pat.c_str(), (long long)ar, (long long)al); }
        return 0;
    }
    const bool quick = argc > 1 && std::string(argv[1]) == "quick";
    const unsigned seed = argc > 2 ? (unsigned)std::stoul(argv[2]) : 12345u;
    const size_t maxLen = argc > 3 ? (size_t)std::stoul(argv[3]) : 120;
    std::printf("=== A) on vs old (lists without empty-matching patterns must be identical) ===\n");
    const std::string t1 = "the quick brown fox jumps over the lazy dog\nfoo bar foo baz\nalpha beta gamma\nxaxbxc dogx xdog\n";
    compareNewOld("A1 literals mixed", t1, { { "the", "THE" }, { "fox", "wolf" }, { "foo", "" }, { "zz", "q" } });
    compareNewOld("A2 regex non-empty", t1, { { "\\bfoo\\b", "F" }, { "b[a-z]+", "<$0>", true }, { "o{2}", "0", true }, { "zq", "1", true } });
    compareNewOld("A3 overlapping/adjacent/ties", t1, { { "abc", "1" }, { "bcd", "2" }, { "xa", "-" }, { "ax", "+" }, { "x", "" } });
    compareNewOld("A4 boundary creation", t1, { { "x", " " }, { "\\bdog\\b", "CAT", true }, { "the", "THE", false, true } });
    compareNewOld("A5 whole word + case", t1, { { "the", "T", false, true }, { "FOO", "f", false, false, false }, { "\xC3\xA4", "ae" } });
    compareNewOld("A6 match list 1,3", t1, { { "o", "0" }, { "a", "4" } }, { 1, 3 });
    compareNewOld("A7 newline in replacement", t1, { { "fox", "fox\n" }, { "\\bjumps\\b", "J", true }, { "^alpha", "A", true } });
    compareNewOld("A8 lookahead/lookbehind", t1, { { "fo(?=o)", "F", true }, { "(?<=la)zy", "ZY", true }, { "(?<!x)dog", "D", true } });

    std::printf("\n=== B) one-pass with one entry must equal classic Replace All (the old flow did not for empty matches) ===\n");
    const std::string t2 = "alpha\nbeta\n\ngamma delta\nxx yy\n";
    compareNewClassic("B1 ^ -> '> '", t2, { "^", "> ", true });
    compareNewClassic("B2 $ -> ';'", t2, { "$", ";", true });
    compareNewClassic("B3 \\b -> '|'", t2, { "\\b", "|", true });
    compareNewClassic("B4 x* -> 'y'", t2, { "x*", "y", true });
    compareNewClassic("B5 (?=a) -> '!'", t2, { "(?=a)", "!", true });
    compareNewClassic("B6 a? -> ''", t2, { "a?", "", true });
    compareNewClassic("B7 ^$ empty line -> '#'", t2, { "^$", "#", true });
    compareNewClassic("B8 literal 'a' -> ''", t2, { "a", "" });
    compareNewClassic("B9 \\Ga -> G", t2, { "\\Ga", "G", true });
    compareNewClassic("B9a a\\Kb -> K", "abc abx ab", { "a\\Kb", "K", true });
    compareNewClassic("B9b \\Ka -> K", "aa ba a", { "\\Ka", "K", true });
    compareNewClassic("B9c (?<=a)b -> L", "abc abx ab", { "(?<=a)b", "L", true });
    const std::string t3 = "a\r\nb\r\n\r\nxx yy\r\nfoo bar\r\n";
    const Entry crazy[] = {
        { "$", ";", true }, { "^", ">", true }, { "x*", "y", true }, { "x*", "", true }, { "\\R", "|", true }, { "\\b", "|", true }, { "\\B", "#", true },
        { "(?=\\r)", "!", true }, { "\\r?$", "!", true }, { "^$", "#", true }, { "(?<=\\r)", "@", true }, { "\\r\\n", "\n", true }, { "\\n", "\r\n", true },
        { ".", "_", true }, { ".*", "*", true }, { ".*?", "?", true }, { "a*?", "L", true }, { "(?!a)", "!", true }, { "()", "E", true }, { "(?:)", "e", true },
        { "\\Z", "Z", true }, { "\\z", "z", true }, { "(?m)^b", "M", true }, { "[[:<:]]", "<", true }, { "[[:>:]]", ">", true }, { "\\<", "<", true }, { "\\>", ">", true },
        { "a\\Kb", "K", true }, { "\\Ka", "k", true }, { "\\K", "0", true }, { "(?<=a)b", "L", true }, { "\\G\\K", "g", true },
        { "a*+", "P", true }, { "(?>a*)", "T", true }, { "(x)\\1", "D", true }, { "\\s*", "S", true }, { "\\h*$", "H", true }, { "^\\s*$", "-", true },
        { "(?i)FOO", "I", true }, { "b|", "B", true }, { "|b", "B", true }, { "(?=a|)", "?", true }, { "\\b\\w", "W", true }, { "(?<![a-z])x", "X", true },
        { "a", "aa" }, { "foo", "foofoo" }, { "\\n", "", true }, { "\\r", "", true }, { "b", "\r" }, { "a", "a\n" }, { "x", "\r\n" }, { "o", "$0$0", true },
        { "a", "A", false, true }, { "FOO", "f", false, false, false }, { "y", "", false, true },
        { "(?s)a.b", "S", true }, { "\\N+", "N", true }, { "(?=x*)", "?", true }, { "x", "\\U$0\\E", true }, { "b", "\\n", true }, { "a", std::string(300, 'Q') },
    };
    int bi = 10;
    for (const char* txt : { t2.c_str(), t3.c_str(), "", "\r\n", "\r", "\n\r", "a\rb\nc\r\nd" })
        for (const Entry& e : crazy) compareNewClassic("B" + std::to_string(bi++) + " '" + e.find + "' -> '" + show(e.repl) + "' on '" + show(txt) + "'", txt, e);

    std::printf("\n=== C) on vs off, edge lists (lookarounds, size changes, CRLF, BOF/EOF) ===\n");
    compareNewRef("C1 empty-capable mix", t2, { { "^", "> ", true }, { "a", "" }, { "\\b", "|", true } });
    compareNewRef("C2 nullable before/after", t2, { { "x*", "y", true }, { "beta", "b" }, { "(?=g)", "!", true } });
    compareNewRef("C3 backrefs + named group", t1, { { "(f)(o)o", "$2$1", true }, { "(?<w>the)", "[$1]", true }, { "a", "" } });
    compareNewRef("C4 deletions at BOF/EOF", "aXbXc", { { "a", "" }, { "c", "" }, { "^X", "1", true }, { "X$", "2", true } });
    compareNewRef("C5 empty doc", "", { { "^", ">", true }, { "a", "b" } });
    compareNewRef("C6 no trailing newline, ^ and $", "one\ntwo", { { "^", "[", true }, { "$", "]", true } });
    compareNewRef("C7 CRLF with $ and .", "a\r\nb\r\n", { { "$", ";", true }, { "a.", "A", true }, { "\\R", "|", true } });
    compareNewRef("C8 non-ASCII around replacement", "gr\xC3\xBCn \xC3\xBC" "ber stra\xC3\x9F" "e", { { "gr\xC3\xBCn", "g", true }, { "\\bber", "B", true }, { "\xC3\x9F", "ss" } });
    compareNewRef("C9 replacement creates other entry's match", "ab ab ab", { { "a", "c" }, { "cb", "!" }, { "\\bcb\\b", "?", true } });
    compareNewRef("C10 uncached \\G and \\K", "aaa baa", { { "\\Ga", "G", true }, { "b\\Ka", "K", true }, { "a", "-" } });
    compareNewRef("C11 lookahead hit beyond a shrinking replacement", "aaaa bc bc aaaa bc", { { "aaaa", "a" }, { "b(?=c)", "X", true } });
    compareNewRef("C12 lookahead hit beyond a growing replacement", "a bc a bc", { { "a", "aaaaaaaaaa" }, { "b(?=c)", "X", true } });
    compareNewRef("C13 lookahead target consumed by an earlier entry", "abc abc", { { "ab", "x" }, { "b(?=c)", "Y", true }, { "(?<=x)c", "Z", true } });
    compareNewRef("C14 negative lookahead, following char deleted after the walk passed", "ab ab", { { "a(?!b)", "A", true }, { "b", "" } });
    compareNewRef("C15 10 KB insertion before remembered hits", "a bc bc bc", { { "a", std::string(10240, 'q') }, { "bc", "X" }, { "\\bc", "?", true } });
    compareNewRef("C16 everything before deleted, \\A becomes true", "xyz ab", { { "xyz ", "" }, { "\\Aab", "!", true }, { "^a", "A", true } });
    compareNewRef("C17 deletion to empty document", "ab", { { "ab", "" }, { "^", ">", true }, { "$", ";", true }, { "\\z", "z", true } });
    compareNewRef("C18 multi-byte deltas", "\xC3\xA4\xC3\xA4 b\xC3\xA4 \xC3\x9F" "b", { { "\xC3\xA4", "a" }, { "\\bb", "B", true }, { "(?<=a)b", "X", true }, { "\xC3\x9F", "sss" } });
    compareNewRef("C19 CRLF joined by a replacement", "ab\ncd\nab\ncd", { { "b", "\r" }, { "^c", "C", true }, { "$", ";", true } });
    compareNewRef("C20 CRLF split by a replacement", "a\r\nb\r\nc", { { "\\r", "\rX", true }, { "^b", "B", true }, { "\\bc", "C", true } });
    compareNewRef("C21 no-op replacement (delta 0) with context patterns", "a a a", { { "a", "a" }, { "\\ba", "A", true }, { "(?<=\\s)a", "S", true } });
    compareNewRef("C22 replacement recreates its own find text", "abbb abbb", { { "ab", "a" }, { "\\bb", "B", true } });
    compareNewRef("C23 $ and \\Z after a replacement ending in \\r", "ab\nc\nab\nc", { { "b", "\r" }, { "$", ";", true }, { "\\Z", "Z", true } });
    compareNewRef("C24 replacement text is the next entry's find text", "a b a b", { { "a", "b" }, { "b", "c" }, { "\\bc\\b", "!", true } });
    compareNewRef("C25 lookbehind over the replacement end", "xa xa xa", { { "x", "yy" }, { "(?<=yy)a", "A", true }, { "(?<=x)a", "B", true } });
    compareNewRef("C26 match list with size changes", "a1 a2 a3 a4 a5 a6", { { "a\\d", "<$0>", true }, { "\\s", "_", true }, { "\\b", "|", true } }, { 1, 3, 4 });
    Scoping col; col.scope = OnePassHits::Scope::Column; col.columns = { 1, 3 };
    compareNewRef("C27 column scope: same-line and later-line hits", "a,b,a\r\na,b,a\r\nab,ab,ab", { { "a", "xyz" }, { "\\ba", "A", true }, { "b", "" } }, {}, col);
    compareNewRef("C28 column scope: replacement removes a delimiter", "a,b,c\nx,y,z\n", { { ",b", "" }, { "c", "C" }, { "\\bz", "Z", true } }, {}, col);
    Scoping sel; sel.scope = OnePassHits::Scope::Selection; sel.sel = { { 2, 9 }, { 12, 20 } };
    compareNewRef("C29 selection scope: growing and shrinking inside ranges", "aa bb aa bb aa bb aa bb", { { "aa", "a" }, { "bb", "bbbb" }, { "\\ba", "A", true } }, {}, sel);
    // analytical findings: a form feed is a line separator for ^ and $ in Boost; whole-word follows the document's word characters
    compareNewRef("C30 form feed inserted before a remembered ^ entry", "a!b\nb", { { "a!", "\f" }, { "^b", "B", true } });
    compareNewRef("C31 form feed removed before a remembered $ entry", "a\fb a\fb", { { "\f", "" }, { "a$", "A", true }, { "b", "" } });
    DocOptions custom; custom.wordChars = "abcdefghijklmnopqrstuvwxyz0123456789_-";
    compareNewRef("C32 custom word chars: '-' replaced by ' ' frees a whole-word hit", "x-foo foo", { { "-", " " }, { "foo", "F", false, true } }, {}, {}, custom);
    compareNewRef("C33 custom word chars: ' ' replaced by '-' binds a whole-word hit", "x foo foo", { { " ", "-" }, { "foo", "F", false, true } }, {}, {}, custom);
    // Stage 6: the cases that can only be built here
    compareNewRef("C34 replacement ends with a lead byte, untouched continuation byte completes it", "x\xA4" "b x\xA4" "b", { { "x", "\xC3" }, { "\\bb", "B", true } });
    compareNewRef("C35 lone lead byte removed before a word-boundary entry", "\xC3" "b \xC3" "b", { { "\xC3", "" }, { "\\bb", "B", true }, { "b", "" } });
    compareNewRef("C36 U+2028 inserted before a remembered ^ entry", "a!b\nb", { { "a!", "\xE2\x80\xA8" }, { "^b", "B", true } });
    compareNewRef("C37 NEL and U+2028 removed before $ entries", "a\xE2\x80\xA8" "b a\xC2\x85" "b", { { "\xE2\x80\xA8", "" }, { "\xC2\x85", "" }, { "a$", "A", true }, { "b", "" } });
    compareNewRef("C38 never-matching, always-empty and degenerate patterns", "ab ab", { { "(?!)", "!", true }, { "(?=)", "E", true }, { "a{0}", "0", true }, { "(|a)", "|", true }, { "b", "" } });
    compareNewRef("C39 conditionals, recursion, named back references", "aabb ab abab aa", { { "(?(?=a)ab|b)", "X", true }, { "(a(?1)?b)", "R", true }, { "(?<n>a)\\k<n>", "D", true } });
    compareNewRef("C40 whole-pattern recursion and (?x) comments (never remembered)", "aabb ab a b", { { "^(a|b(?R))", "Q", true }, { "(?x) a # c", "Y", true }, { "\\Qa.b", "Q", true }, { "b", "B" } });
    compareNewRef("C41 \\K hits verified from the walk position", "abc abc ab", { { "a\\Kb", "K", true }, { "\\Kc", "C", true }, { "b", "x" } });
    compareNewRef("C42 NUL in text, find and replacement", std::string("a\0b a\0b", 7), { { std::string("\0", 1), "N" }, { "\\x00", "Z", true }, { "b", std::string("\0", 1) } });
    compareNewRef("C43 regex plus whole word (the engine ignores the word flag for regex)", "xfoo foo -foo", { { "foo", "F", true, true }, { "-", " " } });
    compareNewRef("C44 \\G with ^ and \\b, \\K at the pattern start", "aa\naa ab", { { "\\G^a", "G", true }, { "\\G\\b", "|", true }, { "\\Ka", "K", true } });
    compareNewRef("C45 emoji (surrogate pair) around replacements", "\xF0\x9F\x98\x80" "a \xF0\x9F\x98\x80" "b", { { "a", "\xF0\x9F\x98\x80" }, { ".", "_", true }, { "\\bb", "B", true } });
    compareNewRef("C46 64 KB replacement before remembered hits and EOF empty matches", "a bc bc", { { "a", std::string(65536, 'q') }, { "bc", "X" }, { "$", ";", true }, { "\\z", "z", true } });
    compareNewRef("C47 CR-only document", "a\rb\r\rc\r", { { "^", ">", true }, { "$", ";", true }, { "\\R", "|", true }, { ".", "_", true } });
    compareNewRef("C48 form feeds only", "\f\f\f", { { "^", ">", true }, { "$", ";", true }, { "\\Z", "Z", true }, { ".", "_", true } });
    compareNewRef("C49 every separator the engine knows in one text", "a\r\nb\fc\nd\re\xC2\x85" "f\xE2\x80\xA8" "g\xE2\x80\xA9" "h", { { "^", ">", true }, { "$", ";", true }, { "\\R", "|", true }, { "[a-h]", "x", true } });
    { std::string chain; for (int i = 0; i < 500; ++i) chain += "ab "; compareNewRef("C50 500 growing replacements shift remembered hits", chain, { { "a", "aaaa" }, { "\\bb", "B", true }, { "$", ";", true } }); }
    { std::string chain; for (int i = 0; i < 500; ++i) chain += "ab "; compareNewRef("C51 500 deletions shift remembered hits", chain, { { "ab", "" }, { "\\b ", "_", true }, { "^", ">", true } }); }
    { std::vector<Entry> many; for (int i = 0; i < 200; ++i) many.push_back({ std::string(1, (char)('a' + i % 26)) + std::to_string(i % 7), "<" + std::to_string(i) + ">" }); many.push_back({ "\\b", "|", true }); std::mt19937 r(7);
      compareNewRef("C52 201 entries on a 3000 byte text", fuzzText(r, 3000), many); }
    { std::vector<Entry> dup(10, Entry{ "a", "b" }); compareNewRef("C53 ten identical entries with a match list", "aaaa aaaa", dup, { 1, 3, 4 }); }
    { const std::string t = "a b a b a b"; const std::vector<Entry> l = { { "^", ">>>>>>>>>>>>>>>>>>>>", true }, { "\\b", "|", true }, { "a", "" } }; const std::set<int> none;
      Doc a(t), b(t); NewPass n(a, l, none); n.extEdits = { 2, 4, 6 }; n.go(); RefPass r(b, l, none); r.extEdits = { 2, 4, 6 }; r.go();
      CHECK("C55 growing replacements, then edits from outside and reset()  [on == off]", sameRun(n.run, r.run) && !n.capped && !r.capped); }
    { Doc a("ab ab"), b("ab ab"); a.backwardSearch(); b.backwardSearch(); const std::vector<Entry> l = { { "$", ";", true }, { "\\z", "z", true }, { "x*", "y", true } }; const std::set<int> none;
      NewPass n(a, l, none); n.startPos = 5; n.go(); RefPass r(b, l, none); r.startPos = 5; r.go();
      CHECK("C54 start at EOF after a Find Previous (empty search range, bridge direction quirk)  [on == off]", sameRun(n.run, r.run) && !n.capped && !r.capped); }

    std::printf("\n=== P) every pattern as a single entry on every probe text: one-pass == Replace All ===\n");
    {
        std::vector<Entry> all(std::begin(kPool), std::end(kPool));
        all.insert(all.end(), kExoticPool.begin(), kExoticPool.end());
        all.insert(all.end(), kBrokenPool.begin(), kBrokenPool.end());
        for (const Entry& e : crazy) all.push_back(e);
        std::vector<std::string> texts = probeTexts();
        for (const char* t : { t1.c_str(), t2.c_str(), t3.c_str(), "", "\r\n", "\r", "\n\r", "a\rb\nc\r\nd", "a\f\fb", "\xC3\xA4\xC3\xA4", "x\xA4" "b x\xA4" "b", "a,b,c\nx,y,z\n" }) texts.push_back(t);
        CHECK("P single entry == Replace All for every pattern and text", singleEntryEquivalence(all, texts, "P") == 0);
    }

    std::printf("\n=== K) the properties the bookkeeping rests on, measured against the engine ===\n");
    CHECK("K1-K8 classification complete (hand-written list and 4000 random patterns), gate complete, classes sufficient, none monotone, hits stable, no direction quirk, probe window sufficient", engineProperties() == 0);

    std::mt19937 rng(seed); const int cases = quick ? 150 : 600;
    static const char* kWordChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-.";
    struct Section { const char* letter; const char* title; Variant v; int caseDiv = 1; size_t maxLenOverride = 0; };
    const Section sections[] = {
        { "D", "full text (and vs classic for single entries)", {} },
        { "E", "column scope", { OnePassHits::Scope::Column } },
        { "F", "selection scope", { OnePassHits::Scope::Selection } },
        { "G", "external edits during the walk + reset()", { OnePassHits::Scope::Full, true } },
        { "H", "custom word characters (- and . are word characters)", { OnePassHits::Scope::Full, false, { kCpUtf8, kWordChars } } },
        { "I", "ANSI code page (single high bytes)", { OnePassHits::Scope::Full, false, { 0, nullptr } } },
        { "J", "start position inside the text", { OnePassHits::Scope::Full, false, {}, true } },
        { "L1", "literals only", { OnePassHits::Scope::Full, false, {}, false, 1 } },
        { "L2", "plus regex without context tokens", { OnePassHits::Scope::Full, false, {}, false, 2 } },
        { "L3", "plus context-dependent entries, nothing nullable", { OnePassHits::Scope::Full, false, {}, false, 3 } },
        { "M1", "exotic but valid UTF-8: NEL, U+2028/29, ZWSP, BOM, emoji, NUL, control bytes, exotic constructs", { OnePassHits::Scope::Full, false, {}, false, 4, 1 } },
        { "M2", "the same in an ANSI document", { OnePassHits::Scope::Full, false, { 0, nullptr }, false, 4, 1 } },
        { "M3", "the same with external edits", { OnePassHits::Scope::Full, true, {}, false, 4, 1 } },
        { "M5", "exotic entries in column scope", { OnePassHits::Scope::Column, false, {}, false, 4, 1 } },
        { "M6", "exotic entries in selection scope", { OnePassHits::Scope::Selection, false, {}, false, 4, 1 } },
        { "N", "scale: up to 48 entries on texts up to 3000 bytes", { OnePassHits::Scope::Full, false, {}, false, 4, 0, 48 }, 10, 3000 },
    };
    for (const Section& sec : sections) {
        std::printf("\n=== %s) fuzz: %s ===\n", sec.letter, sec.title);
        const int n = cases / sec.caseDiv; const size_t len = sec.maxLenOverride ? sec.maxLenOverride : maxLen;
        const int bad = fuzzSection(sec.letter, rng, n, len, sec.v);
        CHECK(std::string(sec.letter) + " fuzz seed " + std::to_string(seed) + ", " + std::to_string(n) + " cases, max text " + std::to_string(len) + ", mismatches: " + std::to_string(bad), bad == 0);
    }

    // M4: malformed UTF-8. The bridge decodes an invalid sequence from wherever the search starts, so the
    // engine's own answer depends on the start position (measured in K4 for such texts). Remembering and the
    // plain walk search from different positions and may therefore disagree about a hit; neither is "right".
    // What must hold is the safety property: every replacement passed the verification search, and the run
    // terminates. Differences are reported as a number, not as a failure.
    std::printf("\n=== M4) malformed UTF-8: engine answers depend on the search start, differences are informational ===\n");
    // M4a: one entry at a time, remembering / plain walk / Notepad++'s own Replace All (processRange) on the
    // same malformed text. Counts how often each pair differs, to place the bookkeeping's share of the blame.
    {
        std::mt19937 r(seed ^ 0x4d34u); int onOff = 0, onCl = 0, offCl = 0, degenerate = 0, total = 0;
        for (int c = 0; c < cases * 4; ++c) {
            const std::string text = fuzzText(r, 20 + r() % maxLen, false, 2);
            const std::vector<Entry> one{ kBrokenPool[r() % kBrokenPool.size()] };
            const std::set<int> none;
            Doc a(text), b(text); NewPass on(a, one, none); on.go(); RefPass off(b, one, none); off.go();
            const Run cl = classic(text, one[0]);
            ++total;
            if (!sameRun(on.run, off.run)) ++onOff;
            if (on.run.text != cl.text || on.run.replaced != cl.replaced) ++onCl;
            if (off.run.text != cl.text || off.run.replaced != cl.replaced) ++offCl;
            { Doc d(text); Sci::Position len = (Sci::Position)one[0].find.size();
              for (Sci::Position p = 0; p <= d.len(); ++p) { Sci::Position l = len; const Sci::Position q = d.d->FindText(p, d.len(), one[0].find.c_str(), fo(one[0].regex ? (kRx | (int)FindOption::MatchCase) : (int)FindOption::MatchCase), &l); if (q >= 0 && l < 0) { ++degenerate; break; } } }
        }
        std::printf("  %d single-entry runs on malformed text: remembering vs plain walk differ %d, remembering vs N++ Replace All %d, plain walk vs N++ Replace All %d; the engine returned a match ending before its start in %d\n",
            total, onOff, onCl, offCl, degenerate);
    }
    { int capped = 0; const int diff = fuzzSection("M4", rng, cases, maxLen, { OnePassHits::Scope::Full, false, {}, false, 4, 2 }, &capped);
      std::printf("  %d of %d cases differ between remembering and the plain walk; every replacement was verified\n", diff - capped, cases);
      CHECK("M4 malformed UTF-8: both walks terminate (differences informational: " + std::to_string(diff - capped) + " of " + std::to_string(cases) + ")", capped == 0); }

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
