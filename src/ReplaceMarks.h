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
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <vector>

// Text MultiReplace replaced, kept in the document as indicators: Scintilla moves them with every edit and draws a
// dotted underline, or a small triangle where text was deleted. A mark's value names the record of what it replaced.
// Undo steps of their own in the document's history bring the marks back with undo and redo.
namespace ReplaceMarks {

    // A message to one Scintilla view
    using Send = std::function<std::intptr_t(unsigned int, std::uintptr_t, std::intptr_t)>;

    // The text a mark covers, [start, end); a deletion has start == end; start is -1 where no mark is
    struct Span {
        std::intptr_t start = -1;
        std::intptr_t end = -1;
        bool found() const noexcept { return start >= 0; }
    };

    // A mark the navigation reaches, with its place among the marks of the document (counted from 1)
    struct Stop {
        Span span;
        std::size_t index = 0;
        std::size_t count = 0;
    };

    // The list entry a replacement came from: its number (from 1) and find text (UTF-8); entry 0 for none
    struct Source {
        std::size_t entry = 0;
        std::string_view find;
    };

    // How a change reached the text: typed or done by a command, or Scintilla's undo or redo of an earlier one
    enum class Step { Edit, Undo, Redo };

    class Marks {
    public:
        // One indicator underlines replaced text, the other puts a triangle on the character after a deletion
        void setIndicators(int marks, int deletions) noexcept { id_ = marks; deletionId_ = deletions; }
        int indicator() const noexcept { return id_; }
        bool empty() const noexcept { return stores_.empty(); }

        // Per view: a dotted underline, the usual sign of text with a hint behind it, and a triangle in the same grey.
        // The dots grow with the line height, so zoom, font and DPI changes call applyStroke again.
        void applyStyle(const Send& view, bool dark) const;
        void applyStroke(const Send& view) const;

        // The list entry that makes the replacements that follow
        void setSource(std::size_t entry, std::string_view find);

        // Around each replacement of [pos, pos + length): the new text gets a mark with what the area held before
        // MultiReplace replaced anything in it, no text a deletion mark. The caller makes both calls and the replacement
        // one undo action: afterReplace adds the step that brings the marks back. It is true for a document's first mark.
        void beforeReplace(const Send& doc, std::intptr_t pos, std::intptr_t length);
        bool afterReplace(const Send& doc, std::intptr_t newLength);

        // Changes to the text as Scintilla reports them. Text inserted otherwise loses the mark it took from its
        // neighbours, inside one mark it joins it; deleted marks go into the undo history with their text.
        void inserted(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step);
        void deleting(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step);
        void deleted(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step);

        // An undo step (SC_MOD_CONTAINER) reached by undo or redo; false where the token is not one of the marks'
        bool replayed(const Send& doc, std::intptr_t token, bool undo);

        // The marks of a document: any at all, the one holding the character at pos, the deletion at the boundary pos,
        // the one at a caret (a deletion there, the mark after it, the one before), those in the selection [from, to)
        bool any(const Send& doc) const;
        Span at(const Send& doc, std::intptr_t pos) const;
        Span deletionAt(const Send& doc, std::intptr_t pos) const;
        Span atCaret(const Send& doc, std::intptr_t pos) const;
        std::vector<Span> within(const Send& doc, std::intptr_t from, std::intptr_t to) const;

        // What a mark replaced, and the list entry that did it
        std::string_view originalOf(const Send& doc, Span mark) const;
        Source sourceOf(const Send& doc, Span mark) const;

        // The first mark after the selection [from, to), or the last one before it, around the document end; a deletion
        // at an empty selection counts as reached
        Stop step(const Send& doc, std::intptr_t from, std::intptr_t to, bool forward) const;

        // Marks get back the text they replaced, in one undo step: those in the selection [from, to), which then also
        // covers the restored text, or one mark, whose restored text is returned
        Span revert(const Send& doc, std::intptr_t from, std::intptr_t to);
        Span revert(const Send& doc, Span mark);

        // The marks of one document, or of all: the others drop theirs in dropStale when they show again
        void clear(const Send& doc);
        void retire();
        void dropStale(const Send& doc);

        // The document closes
        void forget(std::intptr_t document);

    private:
        struct Record {
            std::uint32_t offset;
            std::uint32_t length;
            std::uint32_t source;          // sources[source - 1], 0 for none
        };

        struct Origin {
            std::size_t entry = 0;
            std::string find;
        };

        struct Store {
            std::string pool;              // the replaced texts, one after the other
            std::vector<Record> records;   // mark value v names records[v - 1]
            std::vector<Origin> sources;
        };

        // A run of one indicator in a snapshot, relative to its start; the top bit of the value marks a deletion
        struct Piece {
            std::uint32_t offset;
            std::uint32_t length;
            std::uint32_t value;
        };

        // A change of the text with the marks around it before and after, for undo and redo to put back
        struct Entry {
            std::intptr_t pos = 0;         // where the text changed
            std::intptr_t start = 0;       // where both snapshots begin
            std::uint32_t removed = 0;     // the text taken out and put in
            std::uint32_t inserted = 0;
            std::uint32_t before = 0;      // the snapshot lengths
            std::uint32_t after = 0;
            std::uint32_t first = 0;       // its pieces: those before, then those after
            std::uint32_t beforePieces = 0;
            std::uint32_t afterPieces = 0;
        };

        // A document's entries; the token of entries[i] is base + i, and it keeps counting after a clear
        struct Journal {
            std::uint32_t base = 0;
            std::uint32_t redoFrom = UINT32_MAX;   // entries undone from here on: the history drops them with the next step
            std::vector<Entry> entries;
            std::vector<Piece> pieces;
        };

        // The replacement under way: the area its mark covers (the replaced text and the marks it touches)
        struct Pending {
            std::intptr_t doc = 0;
            std::intptr_t pos = -1;        // the replaced text
            std::intptr_t start = -1;
            std::intptr_t end = -1;
            std::intptr_t length = 0;      // of the replaced text
            std::intptr_t reuse = 0;       // the area is one whole mark: its record stays
            bool deletionAfter = false;    // an insertion took in the deletion at its place, which then follows it
            std::string original;
            std::intptr_t before = 0;      // the area and the character after it, as they were
            std::vector<Piece> pieces;
        };

        // An undo under way: the entry whose marks come back once its text has; done, the document length after that
        struct Replay {
            std::intptr_t doc = 0;
            std::uint32_t index = 0;
            std::intptr_t done = -1;
        };

        // Marks in text about to be deleted, for the undo step that follows the deletion
        struct Capture {
            std::intptr_t doc = 0;
            std::intptr_t pos = -1;
            std::intptr_t length = 0;
            std::vector<Piece> pieces;
        };

        // A run of one indicator with a record, [start, end); start is -1 where none is
        struct Run {
            std::intptr_t start = -1;
            std::intptr_t end = -1;
            std::intptr_t value = 0;
        };

        static bool known(const Store& store, std::intptr_t value) noexcept;
        static std::string_view recordText(const Store& store, std::intptr_t value);
        static std::intptr_t newRecord(Store& store, const std::string& original, std::uint32_t source);
        static std::intptr_t shareRecord(Store& store, std::intptr_t value, std::uint32_t source);
        std::uint32_t sourceIn(Store& store) const;
        const Store* storeOf(std::intptr_t document) const;
        const Store* storeOf(const Send& doc) const;
        const Record* recordOf(const Send& doc, Span mark, const Store*& store) const;
        void dropStaleOf(const Send& doc, std::intptr_t document);
        Run nextRun(const Send& doc, const Store& store, int id, std::intptr_t pos) const;
        Span revertAll(const Send& doc, const std::vector<Span>& marks);

        std::uint32_t snapshot(const Send& doc, std::intptr_t from, std::intptr_t to, std::vector<Piece>& pieces) const;
        void restore(const Send& doc, std::intptr_t from, std::uint32_t length, const Piece* pieces, std::uint32_t count) const;
        bool record(const Send& doc, std::intptr_t document, Entry entry, const std::vector<Piece>& before, bool coalesce);
        bool settle(const Send& doc, std::intptr_t document, bool insertion, std::intptr_t pos, std::intptr_t length);
        void dropJournal(std::intptr_t document);

        // Calls visit(span, value) for each mark with a record, in document order from the one at pos on, while it
        // returns true; a deletion comes before a mark starting at its place
        template <typename Visit>
        void eachMark(const Send& doc, const Store& store, std::intptr_t pos, Visit visit) const;

        int id_ = -1;
        int deletionId_ = -1;
        std::unordered_map<std::intptr_t, Store> stores_;     // Scintilla document -> its records
        std::unordered_map<std::intptr_t, Journal> journals_; // Scintilla document -> its undo steps
        std::unordered_set<std::intptr_t> stale_;             // documents still showing marks of dropped records
        Origin source_;
        Pending pending_;
        Replay replay_;
        Capture capture_;
        bool busy_ = false;                                    // reverting: its own changes are no edits
    };

    // A calltip's text; [start, end) holds the original, drawn in the calltip's highlight color
    struct Tip { std::string text; std::size_t start = 0; std::size_t end = 0; };

    // The calltip: label and the original with control characters, line breaks and the spaces at its edges (as space)
    // made visible, cut to a few lines (utf8 counts characters); after a jump, place ("3 of 12") between arrows leads;
    // a line with the list entry (its label) and its find text in quotes follows
    Tip tipText(std::string_view label, std::string_view empty, std::string_view space, std::string_view original,
        bool utf8, std::string_view place = {}, std::string_view entry = {}, std::string_view find = {});
}
