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

// QA harness for the one-pass Replace All walk (OnePassHits). Builds the real
// class against a fake document with a small pattern set that reproduces the
// N++ Boost bridge rules the walk relies on:
//   - leftmost match, empty matches allowed at the search start
//   - ^ and $ around CRLF (neither matches between \r and \n), a form feed ends a line, \b \B \A \z
//   - the end of a search range counts as the end of the text ($ \z \b lookahead)
//   - a barred empty match is stepped over CRLF-aware (bridge nextCharacter)
// and checks, on random texts and lists:
//   A. fixed expectations for the empty-match cases the old one-pass got
//      wrong (^ / $ / \b / x* not replaced or looping), termination
//   B. one entry == N++ Replace All (processRange + NOTAFTERMATCH continuation)
//   C. remember = true == remember = false in the full text
//   D. the same in column scope (delimiter-separated fields, like performSearchColumn)
//   E. the same in selection scope (ranges moved like adjustSelectionScope)
//   F. the same with outside edits during the walk followed by reset()
//   G. the same with custom word characters (SCI_SETWORDCHARS: - and . are word characters)
//   H. the same with a start position inside the text
//   L1-L3. entry classes in isolation: literals; plus regex without context; plus context-dependent
// C-L also vary the replacement length per hit (formula-like) and disable entries.
// The same checks on the real Scintilla document and N++ Boost bridge are in
// onepass_engine_qa.cpp (needs the N++ sources).
//
// Build (from src/tests):
//   g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. -o onepass_qa onepass_qa.cpp ../OnePassHits.cpp
//   cl /std:c++20 /EHsc /I.. onepass_qa.cpp ..\OnePassHits.cpp /Fe:onepass_qa.exe
// Usage: onepass_qa [seed] [cases]
#include "../OnePassHits.h"

#include <algorithm>
#include <cstdio>
#include <cstring>
#include <random>
#include <set>
#include <string>
#include <vector>

// ------------------------------------------------ fake document + mini engine
static constexpr int kWholeWord = 0x2, kMatchCase = 0x4, kRegExp = 0x00200000;   // SCFIND_* (Scintilla.h)

struct Entry { std::string find, repl; bool regex = false, wholeWord = false, matchCase = true; };
static int flagsOf(const Entry& e) { return (e.regex ? kRegExp : 0) | (e.wholeWord ? kWholeWord : 0) | (e.matchCase ? kMatchCase : 0); }

static std::string g_extraWordChars;   // SCI_SETWORDCHARS additions for the whole-word check (regex \b keeps the engine's fixed classes)
static bool isWordRx(const std::string& t, long i) { return i >= 0 && i < (long)t.size() && (isalnum((unsigned char)t[i]) || t[i] == '_'); }
static bool isWord(const std::string& t, long i) { return isWordRx(t, i) || (i >= 0 && i < (long)t.size() && g_extraWordChars.find(t[i]) != std::string::npos); }

struct Doc {
    std::string t;
    long searches = 0;
    long len() const { return (long)t.size(); }
    char at(long i) const { return (i >= 0 && i < len()) ? t[i] : '\0'; }
    bool lineEndAt(long i) const { return t[i] == '\n' || (t[i] == '\r' && at(i + 1) != '\n'); }
    long lineCount() const { long n = 1; for (long i = 0; i < len(); ++i) if (lineEndAt(i)) ++n; return n; }
    long nextChar(long p) const { return (at(p) == '\r' && at(p + 1) == '\n') ? p + 2 : p + 1; }   // bridge nextCharacter()
    long lineOf(long p) const { long n = 0; for (long i = 0; i < p && i < len(); ++i) if (lineEndAt(i)) ++n; return n; }
    long lineStart(long line) const { long n = 0; for (long i = 0; i < len(); ++i) { if (n == line) return i; if (lineEndAt(i)) ++n; } return len(); }

    // pattern atoms: literal char, '.', 'c*', \b \B ^ $ \A \z, (?=c) (?!c) (?<=c) (?<!c); \G(?:...) anchors at the search start
    struct Atom { char kind; char c; };
    static char lit(const std::string& p, size_t& i) {   // one literal character at p[i] (\r and \n as escapes); advances i past it
        if (p[i] == '\\' && i + 1 < p.size()) { const char c = p[i + 1]; i += 2; return c == 'r' ? '\r' : c == 'n' ? '\n' : c; }
        return p[i++];
    }
    static std::vector<Atom> parse(const std::string& p) {
        std::vector<Atom> a;
        for (size_t i = 0; i < p.size();) {
            if (p.compare(i, 3, "(?=") == 0 || p.compare(i, 3, "(?!") == 0) { const char k = p[i + 2]; i += 3; a.push_back({ k, lit(p, i) }); ++i; }
            else if (p.compare(i, 4, "(?<=") == 0 || p.compare(i, 4, "(?<!") == 0) { const char k = p[i + 3] == '=' ? '<' : '>'; i += 4; a.push_back({ k, lit(p, i) }); ++i; }
            else if (p[i] == '\\' && i + 1 < p.size() && strchr("bBAz", p[i + 1])) { a.push_back({ p[i + 1], 0 }); i += 2; }
            else if (p[i] == '^' || p[i] == '$' || p[i] == '.') { a.push_back({ p[i], 0 }); ++i; }
            else { const char c = lit(p, i); if (i < p.size() && p[i] == '*') { a.push_back({ '*', c }); ++i; } else a.push_back({ 'L', c }); }
        }
        return a;
    }
    // `to` is the end of the search range: the engine sees nothing behind it, but everything before `from`
    bool atLineStart(long p) const { return p == 0 || at(p - 1) == '\n' || at(p - 1) == '\f' || (at(p - 1) == '\r' && at(p) != '\n'); }
    bool lineEnd(long p, long to) const { return p == to || at(p) == '\r' || at(p) == '\f' || (at(p) == '\n' && (p == 0 || at(p - 1) != '\r')); }
    bool matchAt(const std::vector<Atom>& a, size_t i, long p, bool ic, long to, long& end) const {
        if (i == a.size()) { end = p; return true; }
        const Atom& x = a[i];
        auto eq = [&](char d, char c) { return ic ? tolower((unsigned char)d) == tolower((unsigned char)c) : d == c; };
        auto wordAt = [&](long q) { return q < to && isWordRx(t, q); };
        switch (x.kind) {
        case 'L': return p < to && eq(t[p], x.c) && matchAt(a, i + 1, p + 1, ic, to, end);
        case '.': return p < to && t[p] != '\n' && t[p] != '\r' && t[p] != '\f' && matchAt(a, i + 1, p + 1, ic, to, end);
        case '*': { long n = 0; while (p + n < to && eq(t[p + n], x.c)) ++n; for (; n >= 0; --n) if (matchAt(a, i + 1, p + n, ic, to, end)) return true; return false; }
        case 'b': return (isWordRx(t, p - 1) != wordAt(p)) && matchAt(a, i + 1, p, ic, to, end);
        case 'B': return (isWordRx(t, p - 1) == wordAt(p)) && matchAt(a, i + 1, p, ic, to, end);
        case '^': return atLineStart(p) && matchAt(a, i + 1, p, ic, to, end);
        case '$': return lineEnd(p, to) && matchAt(a, i + 1, p, ic, to, end);
        case 'A': return p == 0 && matchAt(a, i + 1, p, ic, to, end);
        case 'z': return p == to && matchAt(a, i + 1, p, ic, to, end);
        case '=': return p < to && eq(t[p], x.c) && matchAt(a, i + 1, p, ic, to, end);
        case '!': return !(p < to && eq(t[p], x.c)) && matchAt(a, i + 1, p, ic, to, end);
        case '<': return p > 0 && eq(t[p - 1], x.c) && matchAt(a, i + 1, p, ic, to, end);
        case '>': return !(p > 0 && eq(t[p - 1], x.c)) && matchAt(a, i + 1, p, ic, to, end);
        }
        return false;
    }
    // leftmost match in [from, to]; empty matches allowed at from (EMPTYMATCH_ALLOWATSTART)
    OnePassHit find(long from, long to, std::string pat, int flags) {
        ++searches;
        OnePassHit h;
        if (from > to) return h;
        if (flags & kRegExp) {
            bool anchored = false;
            if (pat.rfind("\\G(?:", 0) == 0 && pat.back() == ')') { anchored = true; pat = pat.substr(5, pat.size() - 6); }
            const std::vector<Atom> a = parse(pat);
            for (long p = from; p <= (anchored ? from : to); ++p) { long end; if (matchAt(a, 0, p, !(flags & kMatchCase), to, end)) { h.pos = p; h.len = end - p; return h; } }
            return h;
        }
        for (long p = from; p + (long)pat.size() <= to; ++p) {
            bool ok = true;
            for (size_t k = 0; k < pat.size() && ok; ++k) ok = (flags & kMatchCase) ? t[p + k] == pat[k] : tolower((unsigned char)t[p + k]) == tolower((unsigned char)pat[k]);
            if (ok && (flags & kWholeWord)) ok = !isWord(t, p - 1) && !(p + (long)pat.size() < to && isWord(t, p + (long)pat.size()));
            if (ok) { h.pos = p; h.len = (long)pat.size(); return h; }
        }
        return h;
    }
    long replace(long p, long n, const std::string& r) { t.replace((size_t)p, (size_t)n, r); return (long)r.size(); }
};

// ------------------------------------------------ search scope as the panel applies it
// Column: per line only the selected delimiter-separated fields (a field runs to its delimiter, the last
// one to the start of the next line). Selection: stored ranges, moved after every replacement.
struct Scoping {
    OnePassHits::Scope scope = OnePassHits::Scope::Full;
    char delim = ','; std::set<int> columns;
    std::vector<std::pair<long, long>> sel;
    OnePassHit search(Doc& doc, long from, const std::string& pat, int flags) const {
        if (scope == OnePassHits::Scope::Full) return doc.find(from, doc.len(), pat, flags);
        if (scope == OnePassHits::Scope::Selection) {
            for (const auto& r : sel) {
                if (from > r.second) continue;
                const long s = std::max(from, r.first);
                if (s == r.second) continue;
                const OnePassHit h = doc.find(s, r.second, pat, flags);
                if (h.pos >= 0) return h;
            }
            return {};
        }
        const long lines = doc.lineCount();
        for (long l = doc.lineOf(from); l < lines; ++l) {
            const long ls = doc.lineStart(l), le = doc.lineStart(l + 1);
            long fs = ls; int col = 1;
            for (long p = ls; p <= le; ++p) {
                if (p < le && doc.t[p] != delim) continue;
                const long fe = (p < le) ? p : le;
                if (columns.count(col) && fe >= from) { const OnePassHit h = doc.find(std::max(fs, from), fe, pat, flags); if (h.pos >= 0) return h; }
                fs = p + 1; ++col;
            }
        }
        return {};
    }
    void afterReplace(long pos, long oldLen, long newLen) {
        const long delta = newLen - oldLen; if (delta == 0) return;
        for (auto& r : sel) { if (pos + oldLen <= r.first) { r.first += delta; r.second += delta; } else if (pos < r.second) r.second += delta; }
    }
};

struct Run { std::string text; std::vector<int> replaced, found; long searches = 0; bool capped = false; };

// ------------------------------------------------ the walk as onePassReplaceAll runs it
struct WalkOptions { std::set<int> matchList; bool dynamicRepl = false; std::set<int> extEdits; Scoping sc; long startPos = 0; std::string wordChars; };
static Run walk(const std::string& text, const std::vector<Entry>& list, bool remember, WalkOptions opt = {}) {
    Doc d{ text }; Run run; run.replaced.assign(list.size(), 0); run.found.assign(list.size(), 0);
    g_extraWordChars = opt.wordChars;
    std::vector<OnePassEntry> entries(list.size());
    for (size_t i = 0; i < list.size(); ++i) { entries[i].findText = list[i].find; entries[i].flags = flagsOf(list[i]); entries[i].active = !list[i].find.empty(); entries[i].regex = list[i].regex; entries[i].wholeWord = list[i].wholeWord; }
    OnePassDoc view;
    view.search = [&](const std::string& pattern, int flags, Sci_Position from) { return opt.sc.search(d, from, pattern, flags); };
    view.searchRange = [&](const std::string& pattern, int flags, Sci_Position from, Sci_Position to) { return d.find(from, to, pattern, flags); };
    view.charAt = [&](Sci_Position pos) { return (int)(unsigned char)d.at(pos); };
    view.positionAfter = [&](Sci_Position pos) { return pos < d.len() ? pos + 1 : pos; };
    view.lineCount = [&]() { return (Sci_Position)d.lineCount(); };
    view.length = [&]() { return (Sci_Position)d.len(); };
    view.utf8 = false;   // the fake engine treats every byte as one character, like an ANSI document
    OnePassHits hits(view, entries, opt.sc.scope, remember);
    Sci_Position pos = opt.startPos;
    for (int steps = 1;; ++steps) {
        if (steps > 100000) { run.capped = true; break; }
        if (opt.extEdits.count(steps) && pos < d.len()) { d.t.insert((size_t)pos + 1, "zz"); hits.reset(); }   // outside edit behind the walk
        size_t w; const OnePassHit h = hits.next(pos, w); if (h.pos < 0) break;
        ++run.found[w];
        if (opt.matchList.empty() || opt.matchList.count(run.found[w])) {
            hits.beforeReplace(h);
            const OnePassHit v = opt.sc.search(d, h.pos, list[w].find, flagsOf(list[w]));   // replaceOne verifies the hit with the entry's own search
            if (v.pos != h.pos || v.len != h.len) { hits.afterSkip(w, h); pos = h.pos + h.len; continue; }
            std::string r = list[w].repl; const int k = run.found[w];
            if (opt.dynamicRepl) { if (k % 5 == 0) r.clear(); else if (k % 3 == 0) r += r; else if (k % 7 == 0) r += "\n"; }
            const long ins = d.replace(h.pos, h.len, r); ++run.replaced[w];
            opt.sc.afterReplace(h.pos, h.len, ins);
            hits.afterReplace(w, h, ins - h.len); pos = h.pos + ins;
        }
        else { hits.afterSkip(w, h); pos = h.pos + h.len; }
    }
    run.text = d.t; run.searches = d.searches; return run;
}

// N++ processRange with EMPTYMATCH_NOTAFTERMATCH: the next search starts at the
// end of the replacement; an empty match right there is refused and the bridge
// retries from nextCharacter(); the loop ends when a match reached the range end.
static Run classic(const std::string& text, const Entry& e) {
    Doc doc{ text }; Run run; run.replaced.assign(1, 0); run.found.assign(1, 0);
    long pos = 0, endRange = doc.len(), lastEnd = -1;
    for (int steps = 0; steps < 100000; ++steps) {
        OnePassHit h = doc.find(pos, doc.len(), e.find, flagsOf(e));
        if (h.pos >= 0 && h.len == 0 && h.pos == lastEnd) { const long after = doc.nextChar(h.pos); h = (after <= doc.len()) ? doc.find(after, doc.len(), e.find, flagsOf(e)) : OnePassHit{}; }
        if (h.pos < 0) break;
        ++run.found[0];
        const long ins = doc.replace(h.pos, h.len, e.repl); ++run.replaced[0];
        if (h.pos + h.len == endRange) break;
        pos = lastEnd = h.pos + ins; endRange += ins - h.len;
    }
    run.text = doc.t; return run;
}

// ------------------------------------------------ driver
static int failures = 0;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }
static std::string show(std::string s) { for (char& c : s) { if (c == '\n') c = '|'; else if (c == '\r') c = '~'; } return s.size() > 80 ? s.substr(0, 80) + "..." : s; }
static bool sameRun(const Run& a, const Run& b) { return a.text == b.text && a.replaced == b.replaced && a.found == b.found && !a.capped && !b.capped; }

static void expect(const std::string& name, const std::string& text, const std::vector<Entry>& list, const std::string& want, std::vector<int> counts = {}, std::set<int> ml = {}) {
    WalkOptions opt; opt.matchList = ml;
    const Run r = walk(text, list, true, opt);
    const bool ok = r.text == want && !r.capped && (counts.empty() || r.replaced == counts);
    CHECK(name, ok);
    if (!ok) std::printf("      got  %s\n      want %s\n", show(r.text).c_str(), show(want).c_str());
}

static std::string fuzzText(std::mt19937& rng, size_t n, bool csv) {
    static const char* atoms[] = { "a", "b", "ab", "x", " ", "\n", "\r\n", "\r\n", "foo", "1", "_", "\xC3\xA4", "cat", "xa", "\r", "yy", "\f", "-", "x.y", ",", ",", ",\n" };
    const size_t k = csv ? 22 : 19;
    std::string s; while (s.size() < n) s += atoms[rng() % k]; return s;
}
struct PoolEntry { Entry e; bool nullable; };
static const PoolEntry kPool[] = {
    { { "a", "" }, false }, { { "ab", "X" }, false }, { { "foo", "foobar" }, false }, { { "x", "xa" }, false }, { { "cat", "c" }, false }, { { "a", "aa" }, false }, { { "b", "\r" }, false }, { { "a", "a\n" }, false }, { { "x", "\r\n" }, false },
    { { "a", "A", false, true }, false }, { { "FOO", "f", false, false, false }, false }, { { "\n", "" }, false }, { { "\r", "" }, false }, { { "yy", "", false, true }, false }, { { ",", ";" }, false }, { { "a", std::string(200, 'Q') }, false },
    { { "-", " " }, false }, { { " ", "-" }, false }, { { "x-y", "Q", false, true }, false }, { { "foo", "F", false, true }, false }, { { "b", "\f" }, false }, { { "\f", "" }, false }, { { "a", "a," }, false }, { { "x.y", "z" }, false },
    { { "^", "> ", true }, true }, { { "$", ";", true }, true }, { { "\\b", "|", true }, true }, { { "\\B", "#", true }, true }, { { "x*", "y", true }, true }, { { "x*", "", true }, true }, { { "a*", "-", true }, true },
    { { "(?=a)", "!", true }, true }, { { "(?!a)", "?", true }, true }, { { "(?<=a)b", "B", true }, false }, { { "(?<!x)a", "N", true }, false }, { { "\\Aa", "S", true }, false }, { { "a\\z", "E", true }, false },
    { { "^a", "A", true }, false }, { { "b$", "B", true }, false }, { { "\\bfoo", "F", true }, false }, { { "a.", "_", true }, false }, { { ".", "_", true }, false }, { { "^$", "#", true }, true }, { { "\\r$", "!", true }, true },
    { { "(?=\\r)", "!", true }, true }, { { "(?<=\\r)", "@", true }, true }, { { "x*\\b", "W", true }, true }, { { "\\bx*", "V", true }, true }, { { "a*$", "Z", true }, true }, { { "^x*", "Q", true }, true },
    { { "\\z", "z", true }, true }, { { "(?=,)", "!", true }, true }, { { "(?<=,)a", "C", true }, false }, { { "a(?!\\n)", "A", true }, false }, { { "b*$", "T", true }, true }, { { "^b", "B", true }, false }, { { "\\bx", "X", true }, false },
};
static bool hasCtxToken(const Entry& e) {
    if (e.wholeWord) return true;
    if (!e.regex) return false;
    for (const char* t : { "\\b", "\\B", "^", "$", "\\A", "\\Z" }) if (e.find.find(t) != std::string::npos) return true;
    return false;
}
// level 1: literals; 2: plus regex without context tokens; 3: plus context-dependent entries; 4: everything (nullable too)
static bool inLevel(const PoolEntry& p, int level) {
    if (level >= 4) return true;
    if (p.nullable) return false;
    if (!p.e.regex && !p.e.wholeWord) return true;
    if (level == 1) return false;
    return level >= 3 || !hasCtxToken(p.e);
}
static Entry fuzzEntry(std::mt19937& rng, int level = 4) {
    for (;;) { const PoolEntry& p = kPool[rng() % (sizeof(kPool) / sizeof(kPool[0]))]; if (inLevel(p, level)) return p.e; }
}
static std::vector<Entry> fuzzList(std::mt19937& rng, int level = 4) {
    std::vector<Entry> list; const size_t n = 1 + rng() % 8;
    for (size_t i = 0; i < n; ++i) { Entry e = fuzzEntry(rng, level); if (rng() % 6 == 0) e.find.clear(); list.push_back(e); }   // empty find = disabled entry
    return list;
}
struct Variant { OnePassHits::Scope scope = OnePassHits::Scope::Full; bool extEdits = false; const char* wordChars = ""; bool randomStart = false; int level = 4; };
static int fuzzSection(const char* name, std::mt19937& rng, int cases, Variant v) {
    int bad = 0;
    for (int c = 0; c < cases; ++c) {
        const OnePassHits::Scope scope = v.scope; const bool extEdits = v.extEdits;
        const bool csv = scope == OnePassHits::Scope::Column;
        const std::string text = fuzzText(rng, 10 + rng() % 200, csv);
        const std::vector<Entry> list = fuzzList(rng, v.level);
        WalkOptions opt; opt.sc.scope = scope; opt.wordChars = v.wordChars;
        if (v.randomStart) { opt.startPos = rng() % (text.size() + 1); while (opt.startPos > 0 && ((unsigned char)text[opt.startPos] & 0xC0) == 0x80) --opt.startPos; }
        if (rng() % 5 == 0) opt.matchList = { 1, 3, 4 };
        opt.dynamicRepl = rng() % 2 == 0;
        if (csv) { const int k = 1 + rng() % 4; for (int i = 0; i < k; ++i) opt.sc.columns.insert(1 + (int)(rng() % 5)); }
        if (scope == OnePassHits::Scope::Selection) {
            const long L = (long)text.size();
            for (int i = 0; i < 2; ++i) { long a = rng() % (L + 1), b = rng() % (L + 1); if (a > b) std::swap(a, b); opt.sc.sel.push_back({ a, b }); }
            std::sort(opt.sc.sel.begin(), opt.sc.sel.end()); if (opt.sc.sel[0].second > opt.sc.sel[1].first) opt.sc.sel[1].first = opt.sc.sel[0].second;
        }
        if (extEdits) { opt.extEdits.insert(2 + (int)(rng() % 5)); opt.extEdits.insert(9 + (int)(rng() % 9)); }
        const Run on = walk(text, list, true, opt), off = walk(text, list, false, opt);
        if (!sameRun(on, off)) {
            if (++bad <= 3) {
                std::printf("  MISMATCH [%s%s%s]%s on %s\n   list:", name, opt.dynamicRepl ? ", dynamic repl" : "", opt.startPos ? ", start > 0" : "", opt.matchList.empty() ? "" : " [ml 1,3,4]", show(text).c_str());
                for (auto& e : list) std::printf(" [%s -> %s%s%s]", e.find.c_str(), show(e.repl).c_str(), e.regex ? " rx" : "", e.wholeWord ? " ww" : "");
                std::printf("\n   on:  %s\n   off: %s\n   found on:", show(on.text).c_str(), show(off.text).c_str());
                for (int v : on.found) std::printf(" %d", v);
                std::printf("  off:");
                for (int v : off.found) std::printf(" %d", v);
                std::printf("\n");
            }
        }
    }
    return bad;
}

int main(int argc, char** argv) {
    const unsigned seed = argc > 1 ? (unsigned)std::stoul(argv[1]) : 7u;
    const int cases = argc > 2 ? std::stoi(argv[2]) : 3000;

    std::printf("=== A) fixed expectations (empty matches replaced like Replace All, no endless loop) ===\n");
    const std::string t = "alpha\nbeta\n\ngamma\n";
    expect("A1 ^ -> '> '", t, { { "^", "> ", true } }, "> alpha\n> beta\n> \n> gamma\n> ", { 5 });
    expect("A2 $ -> ';'", t, { { "$", ";", true } }, "alpha;\nbeta;\n;\ngamma;\n;", { 5 });
    expect("A3 x* -> 'y' inserts once per position, once at EOF", "ab", { { "x*", "y", true } }, "yayby", { 3 });
    expect("A4 two nullables terminate, lowest index wins a tie", "ab", { { "x*", "y", true }, { "(?=b)", "!", true } }, "yay!by", { 3, 1 });
    expect("A5 $ on CRLF once per line", "a\r\nb\r\n", { { "$", ";", true } }, "a;\r\nb;\r\n;", { 3 });
    expect("A6 ^ on CRLF, never inside the pair", "a\r\nb\r\n", { { "^", ">", true } }, ">a\r\n>b\r\n>", { 3 });
    expect("A7 empty replacement of empty match is a no-op walk", "a\r\nb", { { "x*", "", true } }, "a\r\nb", { 4 });
    expect("A8 \\b after deletion re-checked", "ab ab", { { "a", "" }, { "\\bb", "B", true } }, "B B", { 2, 2 });
    expect("A9 replacement creates a boundary for a later entry", "xab", { { "x", " " }, { "\\bab", "!", true } }, " !", { 1, 1 });
    expect("A10 match list: a skipped empty hit does not move the walk past other hits", "ab\nab", { { "^", ">", true }, { "a", "A" } }, "ab\n>Ab", { 1, 1 }, { 2 });
    expect("A11 disabled entry (empty find) is ignored", "ab", { { "", "!" }, { "b", "B" } }, "aB", { 0, 1 });
    expect("A12 lookahead hit beyond a shrinking replacement", "aaaa bc bc", { { "aaaa", "a" }, { "b(?=c)", "X", true } }, "a Xc Xc", { 1, 2 });
    expect("A13 lookahead hit beyond a growing replacement", "a bc a bc", { { "a", "aaaaaaaa" }, { "b(?=c)", "X", true } }, "aaaaaaaa Xc aaaaaaaa Xc", { 2, 2 });
    expect("A14 everything before deleted, \\A becomes true", "xyz ab", { { "xyz ", "" }, { "\\Aab", "!", true } }, "!", { 1, 1 });
    expect("A15 form feed inserted before a remembered ^ entry", "a!b\nb", { { "a!", "\f" }, { "^b", "B", true } }, "\fB\nB", { 1, 2 });
    { WalkOptions o; o.wordChars = "-"; const Run r = walk("x-foo foo", { { "-", " " }, { "foo", "F", false, true } }, true, o); CHECK("A16 custom word chars: '-' replaced by ' ' frees a whole-word hit", r.text == "x F F"); if (r.text != "x F F") std::printf("      got  %s\n", show(r.text).c_str()); }

    std::printf("\n=== B) one entry == N++ Replace All ===\n");
    std::mt19937 rng(seed); int badB = 0;
    for (int c = 0; c < cases; ++c) {
        const std::string text = fuzzText(rng, 10 + rng() % 120, false); const Entry e = fuzzEntry(rng);
        const Run n = walk(text, { e }, true), cl = classic(text, e);
        if (n.text != cl.text || n.replaced != cl.replaced || n.capped) {
            if (++badB <= 3) std::printf("  MISMATCH '%s' -> '%s'%s on %s\n   walk:    %s\n   classic: %s\n", e.find.c_str(), show(e.repl).c_str(), e.regex ? " rx" : "", show(text).c_str(), show(n.text).c_str(), show(cl.text).c_str());
        }
    }
    CHECK("B single entry vs N++ processRange, " + std::to_string(cases) + " cases, mismatches: " + std::to_string(badB), badB == 0);

    struct Section { const char* letter; const char* title; Variant v; };
    const Section sections[] = {
        { "C", "full text", {} },
        { "D", "column scope", { OnePassHits::Scope::Column } },
        { "E", "selection scope", { OnePassHits::Scope::Selection } },
        { "F", "outside edits + reset()", { OnePassHits::Scope::Full, true } },
        { "G", "custom word characters (- and . are word characters)", { OnePassHits::Scope::Full, false, "-." } },
        { "H", "start position inside the text", { OnePassHits::Scope::Full, false, "", true } },
        { "L1", "literals only", { OnePassHits::Scope::Full, false, "", false, 1 } },
        { "L2", "plus regex without context tokens", { OnePassHits::Scope::Full, false, "", false, 2 } },
        { "L3", "plus context-dependent entries, nothing nullable", { OnePassHits::Scope::Full, false, "", false, 3 } },
    };
    for (const Section& sec : sections) {
        std::printf("\n=== %s) remember on vs off, %s ===\n", sec.letter, sec.title);
        const int bad = fuzzSection(sec.letter, rng, cases, sec.v);
        CHECK(std::string(sec.letter) + " " + sec.title + ", " + std::to_string(cases) + " cases, mismatches: " + std::to_string(bad), bad == 0);
    }
    {
        long on = 0, off = 0; std::mt19937 r2(seed + 1);
        for (int c = 0; c < 200; ++c) { const std::string text = fuzzText(r2, 300, false); const std::vector<Entry> list = fuzzList(r2); on += walk(text, list, true).searches; off += walk(text, list, false).searches; }
        CHECK("remembering searches less (" + std::to_string(on) + " vs " + std::to_string(off) + ")", on < off);
    }

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
