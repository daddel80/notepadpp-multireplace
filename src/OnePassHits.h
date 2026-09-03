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

#pragma once

#include <cstdint>
#include <functional>
#include <string>
#include <vector>

#include "Sci_Position.h"

// One-pass Replace All: the text is walked once, at every position the nearest
// hit of any enabled entry wins and replaced text is never touched again.
// OnePassHits answers "which entry hits next from here" and is told what
// happened to each hit. It applies the N++ Replace All rules for empty matches:
// never directly after the entry's own previous match, a CRLF pair is stepped
// over as one character, and an entry sees an empty match at most once per
// original text position, so nullable patterns terminate.
//
// Every entry keeps its next hit and is searched again only when the last
// replacement could have changed it. The constructor's remember parameter is
// a test lever, not a setting: with false every step searches every entry
// from the walk position, which the QA harnesses use as the reference result.
// The plugin always uses the default.

struct OnePassHit {
    Sci_Position pos = -1;
    Sci_Position len = 0;   // pos < 0 and len < 0: the engine refused the search (invalid pattern, exception,
};                          // a match ending before its start) - nothing found, but not proof that nothing is there

// The walk's view of the document, bound by the caller (Scintilla in the
// plugin, a fake document in the QA). Searches run with the entry's own flags,
// empty matches allowed at the start position.
struct OnePassDoc {
    std::function<OnePassHit(const std::string& pattern, int flags, Sci_Position from)> search;                   // next hit at/after from inside the search scope (text, columns, selections), len < 0 = refused
    std::function<OnePassHit(const std::string& pattern, int flags, Sci_Position from, Sci_Position to)> searchRange; // plain search in [from, to]
    std::function<int(Sci_Position pos)> charAt;                  // byte value 0..255, 0 past the end
    std::function<Sci_Position(Sci_Position pos)> positionAfter;  // next character position
    std::function<Sci_Position()> lineCount;
    std::function<Sci_Position()> length;
    bool utf8 = true;                                             // document code page is UTF-8 (an ANSI document has no multi-byte characters)
};

struct OnePassEntry {
    std::string findText;   // document encoding, extended-mode escapes resolved
    int flags = 0;          // Scintilla search flags including SCFIND_REGEXP_EMPTYMATCH_ALLOWATSTART
    bool active = false;    // enabled with a non-empty find text
    bool regex = false;
    bool wholeWord = false;
};

class OnePassHits {
public:
    enum class Scope { Full, Column, Selection };

    OnePassHits(OnePassDoc& doc, std::vector<OnePassEntry> entries, Scope scope, bool remember = true);

    OnePassHit next(Sci_Position pos, size_t& index);                       // nearest hit at/after pos, index = its entry (tie: lowest index)
    void beforeReplace(const OnePassHit& hit);                              // right before the hit is replaced
    void afterReplace(size_t w, const OnePassHit& hit, Sci_Position delta); // hit of entry w replaced, text length changed by delta
    void afterSkip(size_t w, const OnePassHit& hit);                        // hit of entry w counted but left in place
    void reset();                                                           // the text changed behind the walk's back
    const OnePassEntry& entry(size_t i) const { return _e[i]; }
    // where the caller's verification search has to start for this entry's hit. Normally the hit's own
    // position: a search from there finds the same hit again (measured as K5 in onepass_engine_qa).
    // An entry bound to the search start (\K, \G, lookbehind) is the exception: its hit is only
    // visible from the position it was searched from.
    Sci_Position verifyFrom(size_t i, Sci_Position hitPos) const { return _e[i].uncached ? _e[i].searchedFrom : hitPos; }

    // classification, public so a test can check it against the engine:
    // contextSensitive: a match may depend on the character before its start (\b ^ $ ..., whole word)
    // neverRemembered:  reaches back an unknown distance, is bound to the search start ((?<= \K \G),
    //                   or cannot be wrapped in \G(?:...) for the probe (whole-pattern recursion, (?x) comments)
    // contextClass:     how those constructs see the byte before a match start
    enum class Ctx { Bof, Cr, Lf, Word, Other, NonAscii };
    static bool contextSensitive(const std::string& findText, bool regex, bool wholeWord);
    static bool neverRemembered(const std::string& findText, bool regex);
    static Ctx contextClass(int byteBefore);

private:
    struct Entry : OnePassEntry {
        std::string anchored;            // \G(?:findText) for regex entries whose match depends on the character before it
        bool ctxStart = false, uncached = false;
        Sci_Position pos = -1, len = 0, from = -1, searchedFrom = -1, consumedEnd = -1, lastEmptyOrig = -1;
        bool valid = false, none = false, sidestepped = false;   // sidestepped: found past a barred empty match at from
    };

    OnePassDoc& _doc;
    std::vector<Entry> _e;
    Scope _scope;
    bool _remember;
    Ctx _classBefore = Ctx::Bof;
    int _byteBefore = -1;
    Sci_Position _linesBefore = 0;
    Sci_Position _netDelta = 0;   // maps a current position to its original one: pos - _netDelta

    bool emptyBarred(const Entry& e, Sci_Position at) const { return at == e.consumedEnd || at - _netDelta == e.lastEmptyOrig; }
    bool insideChar(Sci_Position at) const { return _doc.utf8 && (_doc.charAt(at) & 0xC0) == 0x80; }   // a trail byte: the position lies inside a character for the regex bridge
    void forget();
    bool keep(Entry& e, const OnePassHit& r);
    OnePassHit search(Entry& e, Sci_Position from);
    bool anchoredAt(const Entry& e, Sci_Position at, OnePassHit& out) const;
    Ctx classAt(Sci_Position pos) const;
    Sci_Position nextChar(Sci_Position pos) const;
};
