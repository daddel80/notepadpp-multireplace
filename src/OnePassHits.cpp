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

#include "OnePassHits.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <initializer_list>
#include <utility>

static bool patternHas(const std::string& p, std::initializer_list<const char*> tokens)
{
    for (const char* t : tokens)
        if (p.find(t) != std::string::npos) return true;
    return false;
}

// an inline flag group that switches on free-spacing mode: (?x) (?ix) (?x-i: ...
static bool hasExtendedFlag(const std::string& p)
{
    for (size_t i = p.find("(?"); i != std::string::npos; i = p.find("(?", i + 2)) {
        size_t j = i + 2;
        bool x = false;
        while (j < p.size() && (std::isalpha(static_cast<unsigned char>(p[j])) || p[j] == '-')) x |= (p[j++] == 'x');
        if (x && j < p.size() && (p[j] == ')' || p[j] == ':')) return true;
    }
    return false;
}

// a \Q that runs to the end of the pattern quotes everything after it
static bool hasOpenQuote(const std::string& p)
{
    const size_t q = p.rfind("\\Q");
    return q != std::string::npos && p.find("\\E", q) == std::string::npos;
}

// Textual token tests. A false positive (the token inside a class or escaped) only costs
// a probe; the completeness of the lists is checked against the engine in onepass_engine_qa.
// (?R) and (?0) recurse into the whole pattern, which the probe's \G(?:...) wrapper would
// change; a # comment in free-spacing mode or an open \Q would swallow the wrapper's bracket.
bool OnePassHits::neverRemembered(const std::string& findText, bool regex)
{
    return regex && (patternHas(findText, { "(?<=", "(?<!", "\\K", "\\G", "(?R)", "(?0)" })
        || hasExtendedFlag(findText) || hasOpenQuote(findText));
}

bool OnePassHits::contextSensitive(const std::string& findText, bool regex, bool wholeWord)
{
    return wholeWord || (regex && patternHas(findText,
        { "\\b", "\\B", "\\<", "\\>", "^", "$", "\\A", "\\Z", "\\`", "[[:<:]]", "[[:>:]]" }));
}

OnePassHits::OnePassHits(OnePassDoc& doc, std::vector<OnePassEntry> entries, Scope scope, bool remember)
    : _doc(doc), _scope(scope), _remember(remember)
{
    _e.reserve(entries.size());
    for (OnePassEntry& in : entries) {
        Entry e;
        static_cast<OnePassEntry&>(e) = std::move(in);
        if (e.active) {
            e.uncached = neverRemembered(e.findText, e.regex);
            e.ctxStart = contextSensitive(e.findText, e.regex, e.wholeWord);
            if (e.regex && e.ctxStart) e.anchored = "\\G(?:" + e.findText + ")";
        }
        _e.push_back(std::move(e));
    }
}

// everything remembered is dropped, the empty-match bars are kept: they belong to the walk, not to a hit
void OnePassHits::forget()
{
    for (Entry& e : _e) { e.valid = false; e.none = false; }
}

OnePassHit OnePassHits::next(Sci_Position pos, size_t& index)
{
    OnePassHit best;
    index = SIZE_MAX;
    // A search decodes the text from where it starts. In well-formed text every search start is a
    // character boundary and two searches see the same characters, which is what lets a hit be
    // remembered and an entry without a hit be left alone. A walk position inside a character (only
    // reachable in malformed text) breaks that, so nothing remembered from before it survives.
    if (insideChar(pos)) forget();
    for (size_t i = 0; i < _e.size(); ++i) {
        Entry& e = _e[i];
        if (!e.active || e.none) continue;
        if (!e.valid || e.pos < pos || (e.sidestepped && pos != e.from)) {
            if (!keep(e, search(e, pos))) continue;
        }
        if (best.pos < 0 || e.pos < best.pos) { best = { e.pos, e.len }; index = i; }
    }
    return best;
}

// every remembered hit is the result of a plain search from some position (or its shift):
// the walk's verification search must never disagree with what was remembered
bool OnePassHits::keep(Entry& e, const OnePassHit& r)
{
    if (r.pos < 0) {
        e.valid = false;
        // "nothing from here" means nothing later only if the engine answered; a refused search says nothing.
        // A barred empty match that could not be sidestepped (end of text) may become legal later.
        e.none = r.len >= 0 && _remember && _scope == Scope::Full && !e.uncached && !e.sidestepped;
        return false;
    }
    e.pos = r.pos; e.len = r.len; e.none = false;
    // a column or selection range ends where the engine sees the end of the text, so an
    // empty hit at that end exists only for a search started inside the range: not kept
    e.valid = _remember && !e.uncached && !(r.len == 0 && _scope != Scope::Full);
    return true;
}

OnePassHit OnePassHits::search(Entry& e, Sci_Position from)
{
    OnePassHit r = _doc.search(e.findText, e.flags, from);
    e.from = e.searchedFrom = from;
    e.sidestepped = e.regex && r.pos == from && r.len == 0 && emptyBarred(e, from);
    if (e.sidestepped) {
        const Sci_Position after = nextChar(from);
        e.searchedFrom = after;
        r = (after > from) ? _doc.search(e.findText, e.flags, after) : OnePassHit{};
    }
    if (r.pos < 0 && r.len < 0) e.sidestepped = false;   // a refused search is not a sidestep
    return r;
}

bool OnePassHits::anchoredAt(const Entry& e, Sci_Position at, OnePassHit& out) const
{
    const Sci_Position docLen = _doc.length();
    const Sci_Position to = e.regex ? docLen : (std::min)(docLen, at + 4 * static_cast<Sci_Position>(e.findText.size()) + 4);
    out = _doc.searchRange(e.regex ? e.anchored : e.findText, e.flags, at, to);
    return out.pos == at;
}

OnePassHits::Ctx OnePassHits::contextClass(int c)
{
    if (c == '\n' || c == '\f') return Ctx::Lf;   // Boost: form feed ends a line for ^ and $ too
    if (c == '\r') return Ctx::Cr;
    if (c >= 0x80) return Ctx::NonAscii;            // letters, separators (U+0085, U+2028) and invalid bytes alike
    return (std::isalnum(c) || c == '_') ? Ctx::Word : Ctx::Other;
}

OnePassHits::Ctx OnePassHits::classAt(Sci_Position pos) const
{
    return pos <= 0 ? Ctx::Bof : contextClass(_doc.charAt(pos - 1));
}

// the bridge's nextCharacter(): a CRLF pair is stepped over as one character
Sci_Position OnePassHits::nextChar(Sci_Position pos) const
{
    if (_doc.charAt(pos) == '\r' && _doc.charAt(pos + 1) == '\n') return pos + 2;
    return _doc.positionAfter(pos);
}

void OnePassHits::beforeReplace(const OnePassHit& hit)
{
    if (!_remember) return;
    const Sci_Position end = hit.pos + hit.len;
    _classBefore = classAt(end);
    _byteBefore = end > 0 ? _doc.charAt(end - 1) : -1;
    _linesBefore = _doc.lineCount();
}

void OnePassHits::afterReplace(size_t w, const OnePassHit& hit, Sci_Position delta)
{
    const Sci_Position oldEnd = hit.pos + hit.len, end = oldEnd + delta;
    Entry& win = _e[w];
    win.valid = false; win.none = false; win.consumedEnd = end;
    if (hit.len == 0) win.lastEmptyOrig = hit.pos - _netDelta;
    _netDelta += delta;
    if (!_remember) return;

    // Nothing remembered survives when the line structure changed (inserted, removed or
    // split line ends) or when the search scope itself can move: a selection range grows
    // or shrinks, a column field gains a delimiter and the rest of the line changes column.
    // Nor when the replacement ends inside a character: the untouched bytes after it then
    // belong to a sequence starting inside the replacement and can decode to other characters
    const bool broken = _scope != Scope::Full || _doc.lineCount() != _linesBefore || insideChar(end);
    const Ctx classAfter = classAt(end);
    const int byteAfter = end > 0 ? _doc.charAt(end - 1) : -1;
    // regex context (\b ^ $ ...) follows the engine's fixed ASCII classes; the whole-word
    // check of a literal follows the document's word characters, which the host can change,
    // so for it any other byte counts as a change
    const bool ctxChanged = classAfter != _classBefore || classAfter == Ctx::NonAscii || _classBefore == Ctx::NonAscii;
    const bool byteChanged = byteAfter != _byteBefore || byteAfter >= 0x80 || _byteBefore >= 0x80;

    for (size_t j = 0; j < _e.size(); ++j) {
        Entry& e = _e[j];
        if (j == w || !e.active || e.uncached) continue;
        if (broken) { e.valid = false; e.none = false; continue; }
        if (e.valid) {
            // untouched text with an untouched character before it: only the position moves;
            // a sidestepped hit is tied to a bar at a fixed position and must be searched again
            if (e.pos > oldEnd && !e.sidestepped) e.pos += delta;
            else e.valid = false;
        }
        if (!e.ctxStart) continue;
        // a new match can only start where the replacement ended; the anchored probe is a
        // filter, the plain search decides (the engine's search heuristics may skip a position
        // the anchored matcher accepts, and the verification follows the plain search)
        if (!(e.regex ? ctxChanged : byteChanged)) continue;
        OnePassHit a;
        if (anchoredAt(e, end, a)) keep(e, search(e, end));
    }
}

void OnePassHits::afterSkip(size_t w, const OnePassHit& hit)
{
    _e[w].valid = false;
    _e[w].consumedEnd = hit.pos + hit.len;
    if (hit.len == 0) _e[w].lastEmptyOrig = hit.pos - _netDelta;
}

void OnePassHits::reset()
{
    for (Entry& e : _e) { e.valid = false; e.none = false; e.consumedEnd = -1; e.lastEmptyOrig = -1; e.searchedFrom = -1; }
}
