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

#include "ReplaceMarks.h"

#include "Scintilla.h"

#include <algorithm>
#include <cstdio>
#include <limits>

namespace ReplaceMarks {

    namespace {

        // Grey (COLORREF, 0x00BBGGRR): dark on light, light on dark editor themes
        constexpr EditorPalette::Color LIGHT_COLOR = 0x505050;
        constexpr EditorPalette::Color DARK_COLOR = 0xB0B0B0;

        // One pixel of dot per 16 pixels of line height: thin at the usual sizes, growing with zoom, font and DPI
        constexpr std::intptr_t LINE_PER_DOT = 16;

        // Mark values stay below SC_INDICVALUEBIT, record offsets fit 32 bits
        constexpr std::size_t VALUE_LIMIT = 0xFFFFFF;
        constexpr std::size_t POOL_LIMIT = std::numeric_limits<std::uint32_t>::max();

        // Undo steps of the marks: 0x52 in the top byte of the token (Flow Tabs use 0x4D52....), the entry below
        constexpr std::intptr_t TOKEN_TAG = 0x52;
        constexpr std::uint32_t TOKEN_INDEXES = 0xFFFFFF;
        constexpr std::uint32_t DELETION_PIECE = 0x80000000u;

        // How much of an original the calltip shows, and of a find text
        constexpr int TIP_LINES = 10;
        constexpr std::size_t TIP_COLUMNS = 100;
        constexpr std::size_t FIND_COLUMNS = 40;

        std::string_view textAt(const Send& doc, std::intptr_t pos, std::intptr_t length)
        {
            if (length <= 0) return {};
            const auto* text = reinterpret_cast<const char*>(doc(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(pos), length));
            return text ? std::string_view(text, static_cast<std::size_t>(length)) : std::string_view();
        }

        // Notepad++ and other plugins fill with the document's current indicator and value: put both back
        class KeepCurrent {
        public:
            explicit KeepCurrent(const Send& doc)
                : doc_(doc), indicator_(doc(SCI_GETINDICATORCURRENT, 0, 0)), value_(doc(SCI_GETINDICATORVALUE, 0, 0)) {}
            ~KeepCurrent()
            {
                doc_(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(indicator_), 0);
                doc_(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(value_), 0);
            }
            KeepCurrent(const KeepCurrent&) = delete;
            KeepCurrent& operator=(const KeepCurrent&) = delete;

        private:
            const Send& doc_;
            const std::intptr_t indicator_;
            const std::intptr_t value_;
        };

        // [pos, pos + length) of one indicator gets value, 0 clears it; the caller keeps the current indicator
        void fill(const Send& doc, int id, std::intptr_t pos, std::intptr_t length, std::intptr_t value)
        {
            if (id < 0 || length <= 0) return;
            doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id), 0);
            if (value) {
                doc(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(value), 0);
                doc(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(pos), length);
            }
            else {
                doc(SCI_INDICATORCLEARRANGE, static_cast<std::uintptr_t>(pos), length);
            }
        }

        bool holds(const Send& doc, int id, std::intptr_t pos, std::intptr_t length)
        {
            if (id < 0 || length <= 0) return false;
            const auto indicator = static_cast<std::uintptr_t>(id);
            if (doc(SCI_INDICATORVALUEAT, indicator, pos)) return true;
            const std::intptr_t end = doc(SCI_INDICATOREND, indicator, pos);
            return end > pos && end < pos + length;
        }

        void clearAll(const Send& doc, int marks, int deletions)
        {
            const KeepCurrent keep(doc);
            const std::intptr_t length = doc(SCI_GETLENGTH, 0, 0);
            fill(doc, marks, 0, length, 0);
            fill(doc, deletions, 0, length, 0);
        }

        void appendVisible(std::string& tip, unsigned char c)
        {
            switch (c) {
            case '\n': tip += "\\n"; break;
            case '\r': tip += "\\r"; break;
            case '\t': tip += "\\t"; break;
            case '\0': tip += "\\0"; break;
            default:
                if (c < 0x20) {   // 1 and 2 would show as calltip arrows
                    char hex[8];
                    std::snprintf(hex, sizeof(hex), "\\x%02X", c);
                    tip += hex;
                }
                else {
                    tip += static_cast<char>(c);
                }
            }
        }

        bool fits32(std::intptr_t value) noexcept
        {
            return value >= 0 && static_cast<std::uintmax_t>(value) <= std::numeric_limits<std::uint32_t>::max();
        }
    }

    void Marks::applyStyle(const Send& view, const EditorPalette::Theme& theme) const
    {
        const auto color = static_cast<std::intptr_t>(EditorPalette::visible(theme.dark ? DARK_COLOR : LIGHT_COLOR, theme));
        if (id_ >= 0) {
            const auto id = static_cast<std::uintptr_t>(id_);
            view(SCI_INDICSETSTYLE, id, INDIC_DOTS);
            view(SCI_INDICSETFORE, id, color);
            applyStroke(view);
        }
        if (deletionId_ >= 0) {   // Scintilla sizes the triangle by the line's descent itself
            const auto id = static_cast<std::uintptr_t>(deletionId_);
            view(SCI_INDICSETSTYLE, id, INDIC_POINT);
            view(SCI_INDICSETFORE, id, color);
        }
    }

    void Marks::applyStroke(const Send& view) const
    {
        if (id_ < 0) return;
        const auto id = static_cast<std::uintptr_t>(id_);
        const std::intptr_t dots = std::max<std::intptr_t>(1, (view(SCI_TEXTHEIGHT, 0, 0) + LINE_PER_DOT / 2) / LINE_PER_DOT);
        if (view(SCI_INDICGETSTROKEWIDTH, id, 0) != dots * 100) view(SCI_INDICSETSTROKEWIDTH, id, dots * 100);
    }

    void Marks::setSource(Source source)
    {
        if (source.entry == 0) source = {};   // without an entry the rest names nothing
        if (source == source_.source()) return;
        source_ = { source.list, source.entry, std::string(source.find) };
    }

    void Marks::beforeReplace(const Send& doc, std::intptr_t pos, std::intptr_t length)
    {
        pending_ = {};
        const std::intptr_t docLength = doc(SCI_GETLENGTH, 0, 0);
        if (id_ < 0 || pos < 0 || length < 0 || pos + length > docLength) return;
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        dropStaleOf(doc, document);
        pending_.doc = document;
        pending_.pos = pos;
        pending_.start = pos;
        pending_.end = pos + length;
        pending_.length = length;

        const Store* store = storeOf(document);
        if (!store) {   // the document has no marks yet
            pending_.original.assign(textAt(doc, pos, length));
            pending_.before = std::min(pending_.end + 1, docLength) - pos;
            return;
        }

        // The area takes in every mark the replaced text touches: marks give their record, the rest its text
        const auto id = static_cast<std::uintptr_t>(id_);
        std::vector<Run> pieces;   // the marks, and the text between them with value 0
        const std::intptr_t stop = pos + length;
        std::intptr_t at = pos;
        if (pos > 0) {   // a mark reaching into the replaced text, or around an insertion point
            const std::intptr_t value = doc(SCI_INDICATORVALUEAT, id, pos - 1);
            const std::intptr_t end = value ? doc(SCI_INDICATOREND, id, pos - 1) : 0;
            if (end > pos) {
                pending_.start = doc(SCI_INDICATORSTART, id, pos - 1);
                pieces.push_back({ pending_.start, end, value });
                at = end;
            }
        }
        while (at < stop) {
            const std::intptr_t value = doc(SCI_INDICATORVALUEAT, id, at);
            std::intptr_t end = doc(SCI_INDICATOREND, id, at);
            if (end <= at) end = stop;   // the document never held a mark
            if (!value) end = std::min(end, stop);
            pieces.push_back({ at, end, value });
            at = end;
        }
        pending_.end = std::max(stop, at);

        // Deletions in the area, or at the place of an insertion, give their text where they were
        std::vector<Run> deletions;
        for (Run cut = nextRun(doc, *store, deletionId_, pending_.start);
             cut.start >= pending_.start && (cut.start < pending_.end || cut.start == pending_.start);
             cut = nextRun(doc, *store, deletionId_, cut.end))
            deletions.push_back(cut);

        std::size_t next = 0;
        for (const Run& piece : pieces) {
            if (piece.value) {   // the deletions on a mark's characters lie before its text
                for (; next < deletions.size() && deletions[next].start < piece.end; ++next)
                    pending_.original.append(recordText(*store, deletions[next].value));
                if (known(*store, piece.value)) pending_.original.append(recordText(*store, piece.value));
                else pending_.original.append(textAt(doc, piece.start, piece.end - piece.start));
                continue;
            }
            std::intptr_t from = piece.start;
            for (; next < deletions.size() && deletions[next].start < piece.end; ++next) {
                pending_.original.append(textAt(doc, from, deletions[next].start - from));
                pending_.original.append(recordText(*store, deletions[next].value));
                from = deletions[next].start;
            }
            pending_.original.append(textAt(doc, from, piece.end - from));
        }
        for (; next < deletions.size(); ++next) pending_.original.append(recordText(*store, deletions[next].value));

        if (pieces.size() + deletions.size() == 1) {   // one whole mark or deletion: its record stays
            const std::intptr_t only = pieces.empty() ? deletions.front().value : pieces.front().value;
            pending_.reuse = known(*store, only) ? only : 0;
        }
        pending_.deletionAfter = pieces.empty() && !deletions.empty();

        // The area with the character after it, for undo: the replacement changes the marks of nothing else
        const std::intptr_t to = std::min(pending_.end + 1, docLength);
        pending_.before = to - pending_.start;
        snapshot(doc, pending_.start, to, pending_.pieces);
    }

    bool Marks::afterReplace(const Send& doc, std::intptr_t newLength)
    {
        Pending pending = std::move(pending_);
        pending_ = {};
        if (id_ < 0 || pending.doc == 0 || newLength < 0 || doc(SCI_GETDOCPOINTER, 0, 0) != pending.doc) return false;
        if (pending.length == 0 && newLength == 0) return false;   // nothing replaced
        const std::intptr_t end = pending.end + newLength - pending.length;
        const std::intptr_t length = doc(SCI_GETLENGTH, 0, 0);
        const bool removed = end <= pending.start;   // no text left: the mark goes on the character after it

        bool first = false;
        if (!removed || (deletionId_ >= 0 && !pending.original.empty() && pending.start < length)) {
            const auto found = stores_.find(pending.doc);
            first = found == stores_.end();
            Store& store = first ? stores_[pending.doc] : found->second;
            const KeepCurrent keep(doc);
            if (removed) {   // a deletion right after the area now has the same place: it joins this one
                const auto deletion = static_cast<std::uintptr_t>(deletionId_);
                const std::intptr_t after = doc(SCI_INDICATORVALUEAT, deletion, pending.start);
                if (known(store, after) && doc(SCI_INDICATORSTART, deletion, pending.start) == pending.start) {
                    pending.original.append(recordText(store, after));
                    pending.reuse = 0;
                }
            }
            const std::uint32_t source = sourceIn(store);
            std::intptr_t value = 0;   // without a record the area stays unmarked
            if (!pending.reuse) value = newRecord(store, pending.original, source);
            else if (store.records[static_cast<std::size_t>(pending.reuse) - 1].source == source) value = pending.reuse;
            else value = shareRecord(store, pending.reuse, source);
            if (removed) {
                fill(doc, deletionId_, pending.start, 1, value);
            }
            else {
                fill(doc, id_, pending.start, end - pending.start, value);
                fill(doc, deletionId_, pending.start, end - pending.start, 0);   // the deletions it took in, or one it inherited
                if (pending.deletionAfter && end < length) fill(doc, deletionId_, end, 1, 0);
            }
        }

        // Undo and redo put back the marks of the area and the character after it
        Entry entry;
        entry.pos = pending.pos;
        entry.start = pending.start;
        if (fits32(pending.length) && fits32(newLength) && fits32(pending.before)) {
            entry.removed = static_cast<std::uint32_t>(pending.length);
            entry.inserted = static_cast<std::uint32_t>(newLength);
            entry.before = static_cast<std::uint32_t>(pending.before);
            entry.after = static_cast<std::uint32_t>(std::min(std::max(end, pending.start) + 1, length) - pending.start);
            record(doc, pending.doc, entry, pending.pieces, false);
        }
        return first;
    }

    void Marks::inserted(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step)
    {
        if (step != Step::Undo) replay_ = {};
        if (id_ < 0 || length <= 0 || stores_.empty() || busy_) return;
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        if (document == pending_.doc) return;   // MultiReplace's own replacement
        if (step == Step::Undo && settle(doc, document, true, pos, length)) return;

        const KeepCurrent keep(doc);
        const std::intptr_t end = pos + length;
        if (holds(doc, deletionId_, pos, length)) fill(doc, deletionId_, pos, length, 0);   // a neighbour's deletion

        const auto id = static_cast<std::uintptr_t>(id_);
        const std::intptr_t around = (pos > 0 && end < doc(SCI_GETLENGTH, 0, 0)) ? doc(SCI_INDICATORVALUEAT, id, pos - 1) : 0;
        if (around && around == doc(SCI_INDICATORVALUEAT, id, end)) return;   // inside one mark: Scintilla filled it in
        if (holds(doc, id_, pos, length)) fill(doc, id_, pos, length, 0);
    }

    void Marks::deleting(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step)
    {
        capture_ = {};
        if (step != Step::Edit || id_ < 0 || length <= 0 || busy_) return;
        replay_ = {};
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        if (document == pending_.doc || !storeOf(document)) return;
        std::vector<Piece> pieces;
        snapshot(doc, pos, pos + length, pieces);
        if (!pieces.empty()) capture_ = { document, pos, length, std::move(pieces) };
    }

    void Marks::deleted(const Send& doc, std::intptr_t pos, std::intptr_t length, Step step)
    {
        if (step != Step::Undo) replay_ = {};
        if (id_ < 0 || length <= 0 || busy_) return;
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        if (step == Step::Undo) {
            settle(doc, document, false, pos, length);
            return;
        }
        if (step != Step::Edit || capture_.doc != document || capture_.pos != pos || capture_.length != length) return;
        const Capture capture = std::move(capture_);
        capture_ = {};
        if (!fits32(length)) return;
        Entry entry;
        entry.pos = pos;
        entry.start = pos;
        entry.removed = static_cast<std::uint32_t>(length);
        entry.before = static_cast<std::uint32_t>(length);
        record(doc, document, entry, capture.pieces, true);   // with the deletion: typing still coalesces
    }

    bool Marks::replayed(const Send& doc, std::intptr_t token, bool undo)
    {
        if ((token >> 24) != TOKEN_TAG) return false;
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        const auto found = journals_.find(document);
        const auto index = static_cast<std::uint32_t>(token & TOKEN_INDEXES);
        replay_ = {};
        if (found == journals_.end() || index < found->second.base || index - found->second.base >= found->second.entries.size())
            return true;   // from before a clear
        Journal& journal = found->second;
        if (undo) {   // the step comes before the change's text: the marks follow once the text is back
            journal.redoFrom = index;
            replay_ = { document, index, -1 };
            return true;
        }
        journal.redoFrom = index + 1;
        const Entry& entry = journal.entries[index - journal.base];
        restore(doc, entry.start, entry.after, journal.pieces.data() + entry.first + entry.beforePieces, entry.afterPieces);
        return true;
    }

    // An undo under way: the entry's text is back with its last change, then its marks; the other view reports that again
    bool Marks::settle(const Send& doc, std::intptr_t document, bool insertion, std::intptr_t pos, std::intptr_t length)
    {
        if (replay_.doc == 0) return false;
        const auto found = journals_.find(replay_.doc);
        if (replay_.doc != document || found == journals_.end() || replay_.index - found->second.base >= found->second.entries.size()) {
            replay_ = {};
            return false;
        }
        const Journal& journal = found->second;
        const Entry& entry = journal.entries[replay_.index - journal.base];
        const bool last = pos == entry.pos
            && (insertion ? (entry.removed > 0 && static_cast<std::uintmax_t>(length) == entry.removed)
                          : (entry.removed == 0 && static_cast<std::uintmax_t>(length) == entry.inserted));
        const std::intptr_t now = doc(SCI_GETLENGTH, 0, 0);
        if (replay_.done >= 0) {
            if (last && now == replay_.done) return true;
            replay_ = {};
            return false;
        }
        if (last) {
            restore(doc, entry.start, entry.before, journal.pieces.data() + entry.first, entry.beforePieces);
            replay_.done = now;
            return true;
        }
        const bool firstOfTwo = !insertion && entry.removed > 0 && pos == entry.pos && static_cast<std::uintmax_t>(length) == entry.inserted;
        if (!firstOfTwo) replay_ = {};   // the inserted text goes first, then the removed text comes back
        return false;
    }

    Marks::Run Marks::nextRun(const Send& doc, const Store& store, int id, std::intptr_t pos) const
    {
        const std::intptr_t length = doc(SCI_GETLENGTH, 0, 0);
        if (id < 0 || pos >= length) return {};
        const auto indicator = static_cast<std::uintptr_t>(id);
        pos = pos > 0 ? doc(SCI_INDICATORSTART, indicator, pos) : 0;   // the start of the run at pos
        while (pos < length) {
            const std::intptr_t end = doc(SCI_INDICATOREND, indicator, pos);
            if (end <= pos) return {};   // the document never held the indicator
            const std::intptr_t value = doc(SCI_INDICATORVALUEAT, indicator, pos);
            if (known(store, value)) return { pos, end, value };
            pos = end;
        }
        return {};
    }

    template <typename Visit>
    void Marks::eachMark(const Send& doc, const Store& store, std::intptr_t pos, Visit visit) const
    {
        Run mark = nextRun(doc, store, id_, pos);
        Run deletion = nextRun(doc, store, deletionId_, pos);
        while (mark.start >= 0 || deletion.start >= 0) {
            if (deletion.start >= 0 && (mark.start < 0 || deletion.start <= mark.start)) {
                if (!visit(Span{ deletion.start, deletion.start }, deletion.value)) return;
                deletion = nextRun(doc, store, deletionId_, deletion.end);
            }
            else {
                if (!visit(Span{ mark.start, mark.end }, mark.value)) return;
                mark = nextRun(doc, store, id_, mark.end);
            }
        }
    }

    bool Marks::any(const Send& doc) const
    {
        const Store* store = storeOf(doc);
        bool found = false;
        if (store) eachMark(doc, *store, 0, [&](Span, std::intptr_t) { found = true; return false; });
        return found;
    }

    Span Marks::at(const Send& doc, std::intptr_t pos) const
    {
        const Store* store = storeOf(doc);
        if (!store || pos < 0 || pos >= doc(SCI_GETLENGTH, 0, 0)) return {};
        const auto id = static_cast<std::uintptr_t>(id_);
        if (!known(*store, doc(SCI_INDICATORVALUEAT, id, pos))) return {};
        return { doc(SCI_INDICATORSTART, id, pos), doc(SCI_INDICATOREND, id, pos) };
    }

    Span Marks::deletionAt(const Send& doc, std::intptr_t pos) const
    {
        const Store* store = deletionId_ < 0 ? nullptr : storeOf(doc);
        if (!store || pos < 0 || pos >= doc(SCI_GETLENGTH, 0, 0)) return {};
        const auto id = static_cast<std::uintptr_t>(deletionId_);
        if (!known(*store, doc(SCI_INDICATORVALUEAT, id, pos)) || doc(SCI_INDICATORSTART, id, pos) != pos) return {};
        return { pos, pos };
    }

    Span Marks::atCaret(const Send& doc, std::intptr_t pos) const
    {
        Span mark = deletionAt(doc, pos);
        if (!mark.found()) mark = at(doc, pos);
        if (!mark.found() && pos > 0) mark = at(doc, pos - 1);
        return mark;
    }

    std::vector<Span> Marks::within(const Send& doc, std::intptr_t from, std::intptr_t to) const
    {
        std::vector<Span> marks;
        const Store* store = from < to ? storeOf(doc) : nullptr;
        if (store) eachMark(doc, *store, from, [&](Span span, std::intptr_t) {
            if (span.start >= to) return false;
            if (span.start < span.end ? span.end > from : span.start > from) marks.push_back(span);   // deletions inside
            return true;
        });
        return marks;
    }

    const Marks::Record* Marks::recordOf(const Send& doc, Span mark, const Store*& store) const
    {
        const int id = mark.start < mark.end ? id_ : deletionId_;
        store = (id < 0 || mark.start < 0) ? nullptr : storeOf(doc);
        if (!store || mark.start >= doc(SCI_GETLENGTH, 0, 0)) return nullptr;
        const std::intptr_t value = doc(SCI_INDICATORVALUEAT, static_cast<std::uintptr_t>(id), mark.start);
        return known(*store, value) ? &store->records[static_cast<std::size_t>(value) - 1] : nullptr;
    }

    std::string_view Marks::originalOf(const Send& doc, Span mark) const
    {
        const Store* store = nullptr;
        const Record* record = recordOf(doc, mark, store);
        return record ? std::string_view(store->pool).substr(record->offset, record->length) : std::string_view();
    }

    Source Marks::sourceOf(const Send& doc, Span mark) const
    {
        const Store* store = nullptr;
        const Record* record = recordOf(doc, mark, store);
        if (!record || record->source == 0) return {};
        return store->sources[record->source - 1].source();
    }

    Stop Marks::step(const Send& doc, std::intptr_t from, std::intptr_t to, bool forward) const
    {
        Stop stop;
        const Store* store = storeOf(doc);
        if (!store) return stop;
        const bool selected = from < to;
        Span first, last;
        eachMark(doc, *store, 0, [&](Span span, std::intptr_t) {
            ++stop.count;
            if (stop.count == 1) first = span;
            last = span;
            const bool text = selected || span.start < span.end;   // a deletion at an empty selection is reached
            const bool after = span.start > to || (span.start == to && text);
            const bool before = span.end < from || (span.end == from && text);
            if ((forward && after && !stop.span.found()) || (!forward && before)) {
                stop.span = span;
                stop.index = stop.count;
            }
            return true;
        });
        if (!stop.span.found() && stop.count > 0) {   // around the document end
            stop.span = forward ? first : last;
            stop.index = forward ? 1 : stop.count;
        }
        return stop;
    }

    Span Marks::revert(const Send& doc, std::intptr_t from, std::intptr_t to)
    {
        const std::vector<Span> marks = within(doc, from, to);
        if (marks.empty()) return {};
        const Span restored = revertAll(doc, marks);
        return { std::min(from, restored.start), std::max(to, marks.back().end) + restored.end - marks.back().end };
    }

    Span Marks::revert(const Send& doc, Span mark)
    {
        const Span now = mark.start < mark.end ? at(doc, mark.start) : deletionAt(doc, mark.start);
        if (!now.found() || now.start != mark.start || now.end != mark.end) return {};
        return revertAll(doc, { mark });
    }

    Span Marks::revertAll(const Send& doc, const std::vector<Span>& marks)
    {
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        const Store& store = *storeOf(document);
        std::intptr_t shift = 0;   // how much longer the restored texts are than their marks
        const KeepCurrent keep(doc);
        busy_ = true;
        doc(SCI_BEGINUNDOACTION, 0, 0);
        for (auto mark = marks.rbegin(); mark != marks.rend(); ++mark) {   // from the last one: the earlier stay in place
            const bool deletion = mark->start == mark->end;
            const auto id = static_cast<std::uintptr_t>(deletion ? deletionId_ : id_);
            const std::intptr_t value = doc(SCI_INDICATORVALUEAT, id, mark->start);
            if (!known(store, value)) continue;
            std::intptr_t before = 0;   // a deletion right before a mark sits on its first character: it stays there
            if (!deletion && deletionAt(doc, mark->start).found())
                before = doc(SCI_INDICATORVALUEAT, static_cast<std::uintptr_t>(deletionId_), mark->start);
            const std::string_view original = recordText(store, value);
            const auto length = static_cast<std::intptr_t>(original.size());

            std::vector<Piece> was;   // the mark and the character after it, for undo
            const std::intptr_t wasEnd = std::min(mark->end + 1, doc(SCI_GETLENGTH, 0, 0));
            snapshot(doc, mark->start, wasEnd, was);

            doc(SCI_SETTARGETRANGE, static_cast<std::uintptr_t>(mark->start), mark->end);
            doc(SCI_REPLACETARGET, original.size(), reinterpret_cast<std::intptr_t>(original.data()));
            fill(doc, id_, mark->start, length, 0);   // the value Scintilla gave the text from its neighbours
            fill(doc, deletionId_, mark->start, deletion ? length + 1 : length, 0);   // with a deletion, the one undone
            const std::intptr_t now = doc(SCI_GETLENGTH, 0, 0);
            if (before && mark->start < now) fill(doc, deletionId_, mark->start, 1, before);
            shift += length - (mark->end - mark->start);

            if (fits32(mark->end - mark->start) && fits32(length) && fits32(wasEnd - mark->start)) {
                Entry entry;
                entry.pos = mark->start;
                entry.start = mark->start;
                entry.removed = static_cast<std::uint32_t>(mark->end - mark->start);
                entry.inserted = static_cast<std::uint32_t>(length);
                entry.before = static_cast<std::uint32_t>(wasEnd - mark->start);
                entry.after = static_cast<std::uint32_t>(std::min(mark->start + length + 1, now) - mark->start);
                record(doc, document, entry, was, false);
            }
        }
        doc(SCI_ENDUNDOACTION, 0, 0);
        busy_ = false;
        return { marks.front().start, marks.back().end + shift };
    }

    void Marks::clear(const Send& doc)
    {
        pending_ = {};
        if (id_ < 0) return;
        const std::intptr_t document = doc(SCI_GETDOCPOINTER, 0, 0);
        dropJournal(document);
        const bool marked = stores_.erase(document) > 0;
        if (stale_.erase(document) > 0 || marked) clearAll(doc, id_, deletionId_);
    }

    void Marks::retire()
    {
        pending_ = {};
        for (const auto& entry : stores_) stale_.insert(entry.first);
        stores_.clear();
        for (auto& journal : journals_) dropJournal(journal.first);
    }

    void Marks::dropStale(const Send& doc)
    {
        if (id_ < 0 || stale_.empty()) return;
        dropStaleOf(doc, doc(SCI_GETDOCPOINTER, 0, 0));
    }

    void Marks::forget(std::intptr_t document)
    {
        dropJournal(document);
        journals_.erase(document);
        stores_.erase(document);
        stale_.erase(document);
        if (pending_.doc == document) pending_ = {};
    }

    bool Marks::known(const Store& store, std::intptr_t value) noexcept
    {
        return value > 0 && static_cast<std::size_t>(value) <= store.records.size();
    }

    std::string_view Marks::recordText(const Store& store, std::intptr_t value)
    {
        const Record& record = store.records[static_cast<std::size_t>(value) - 1];
        return std::string_view(store.pool).substr(record.offset, record.length);
    }

    // 0 when the store has no room for another record
    std::intptr_t Marks::newRecord(Store& store, const std::string& original, std::uint32_t source)
    {
        if (store.records.size() >= VALUE_LIMIT || original.size() > POOL_LIMIT - store.pool.size()) return 0;
        store.records.push_back({ static_cast<std::uint32_t>(store.pool.size()), static_cast<std::uint32_t>(original.size()), source });
        store.pool.append(original);
        return static_cast<std::intptr_t>(store.records.size());
    }

    // Records never change, undo brings back their values: another source takes a record of its own, with the same text
    std::intptr_t Marks::shareRecord(Store& store, std::intptr_t value, std::uint32_t source)
    {
        if (store.records.size() >= VALUE_LIMIT) return 0;
        Record record = store.records[static_cast<std::size_t>(value) - 1];
        record.source = source;
        store.records.push_back(record);
        return static_cast<std::intptr_t>(store.records.size());
    }

    // The current source among those of the store, added on its first use; 0 for none
    std::uint32_t Marks::sourceIn(Store& store) const
    {
        if (source_.entry == 0) return 0;
        for (std::size_t i = store.sources.size(); i-- > 0;) {   // mostly the last one
            if (store.sources[i].source() == source_.source()) return static_cast<std::uint32_t>(i + 1);
        }
        store.sources.push_back(source_);
        return static_cast<std::uint32_t>(store.sources.size());
    }

    const Marks::Store* Marks::storeOf(std::intptr_t document) const
    {
        const auto found = stores_.find(document);
        return found == stores_.end() ? nullptr : &found->second;
    }

    const Marks::Store* Marks::storeOf(const Send& doc) const
    {
        return (id_ < 0 || stores_.empty()) ? nullptr : storeOf(doc(SCI_GETDOCPOINTER, 0, 0));
    }

    void Marks::dropStaleOf(const Send& doc, std::intptr_t document)
    {
        if (stale_.erase(document) > 0) clearAll(doc, id_, deletionId_);
    }

    // The runs of both indicators in [from, to), appended to pieces; their number
    std::uint32_t Marks::snapshot(const Send& doc, std::intptr_t from, std::intptr_t to, std::vector<Piece>& pieces) const
    {
        const std::size_t count = pieces.size();
        for (const int id : { id_, deletionId_ }) {
            if (id < 0) continue;
            const auto indicator = static_cast<std::uintptr_t>(id);
            for (std::intptr_t pos = from; pos < to;) {
                const std::intptr_t value = doc(SCI_INDICATORVALUEAT, indicator, pos);
                std::intptr_t end = doc(SCI_INDICATOREND, indicator, pos);
                if (end <= pos || end > to) end = to;   // the document never held it, or the run goes on
                if (value > 0 && static_cast<std::size_t>(value) <= VALUE_LIMIT) {
                    pieces.push_back({ static_cast<std::uint32_t>(pos - from), static_cast<std::uint32_t>(end - pos),
                        static_cast<std::uint32_t>(value) | (id == deletionId_ ? DELETION_PIECE : 0u) });
                }
                pos = end;
            }
        }
        return static_cast<std::uint32_t>(pieces.size() - count);
    }

    // [from, from + length) gets the marks of a snapshot again
    void Marks::restore(const Send& doc, std::intptr_t from, std::uint32_t length, const Piece* pieces, std::uint32_t count) const
    {
        if (length == 0 || from < 0 || from + static_cast<std::intptr_t>(length) > doc(SCI_GETLENGTH, 0, 0)) return;
        const KeepCurrent keep(doc);
        fill(doc, id_, from, length, 0);
        fill(doc, deletionId_, from, length, 0);
        for (std::uint32_t i = 0; i < count; ++i) {
            const Piece& piece = pieces[i];
            fill(doc, (piece.value & DELETION_PIECE) ? deletionId_ : id_, from + piece.offset, piece.length, piece.value & ~DELETION_PIECE);
        }
    }

    // A change goes into the document's undo history with the marks around it; nothing where none were or will be
    bool Marks::record(const Send& doc, std::intptr_t document, Entry entry, const std::vector<Piece>& before, bool coalesce)
    {
        if (!doc(SCI_GETUNDOCOLLECTION, 0, 0)) return false;
        Journal& journal = journals_[document];
        const std::size_t live = journal.entries.size();
        const std::size_t undone = journal.redoFrom >= journal.base ? journal.redoFrom - journal.base : live;
        if (undone < live) {   // the undo history dropped these with the change that reached it
            journal.pieces.resize(journal.entries[undone].first);
            journal.entries.resize(undone);
        }
        journal.redoFrom = UINT32_MAX;
        const std::size_t index = journal.base + journal.entries.size();
        if (index > TOKEN_INDEXES || journal.pieces.size() + before.size() > POOL_LIMIT / 2) return false;

        const std::size_t first = journal.pieces.size();
        journal.pieces.insert(journal.pieces.end(), before.begin(), before.end());
        entry.afterPieces = snapshot(doc, entry.start, entry.start + entry.after, journal.pieces);
        if (before.empty() && entry.afterPieces == 0) {   // no marks to put back either way
            journal.pieces.resize(first);
            return false;
        }
        entry.first = static_cast<std::uint32_t>(first);
        entry.beforePieces = static_cast<std::uint32_t>(before.size());
        journal.entries.push_back(entry);
        doc(SCI_ADDUNDOACTION, static_cast<std::uintptr_t>((TOKEN_TAG << 24) | static_cast<std::intptr_t>(index)),
            coalesce ? UNDO_MAY_COALESCE : UNDO_NONE);
        return true;
    }

    // Its tokens left in the undo history find no entry any more
    void Marks::dropJournal(std::intptr_t document)
    {
        const auto found = journals_.find(document);
        if (found != journals_.end()) {
            Journal& journal = found->second;
            const std::size_t next = journal.base + journal.entries.size();
            journal.base = static_cast<std::uint32_t>(std::min<std::size_t>(next, TOKEN_INDEXES + 1));
            journal.redoFrom = UINT32_MAX;
            journal.entries = {};
            journal.pieces = {};
        }
        if (replay_.doc == document) replay_ = {};
        if (capture_.doc == document) capture_ = {};
    }

    Tip tipText(std::string_view label, std::string_view empty, std::string_view space, std::string_view original,
        bool utf8, std::string_view place, std::string_view entry, std::string_view find)
    {
        std::string tip;
        if (!place.empty()) tip.append("\001").append(place).append("\002 ");   // the calltip draws 1 and 2 as arrows
        tip.append(label);
        tip += ' ';
        if (original.empty()) tip.append(empty);
        const std::size_t start = tip.size();

        // Spaces at the edges would not show
        const std::size_t lead = original.find_first_not_of(' ');
        const std::size_t trail = original.find_last_not_of(' ');
        int lines = 1;
        std::size_t columns = 0;
        bool cut = false;   // the rest of this line is left out
        for (std::size_t i = 0; i < original.size(); ++i) {
            const auto c = static_cast<unsigned char>(original[i]);
            const bool lineEnd = c == '\r' || c == '\n';
            if (!cut && !lineEnd && columns == TIP_COLUMNS && (!utf8 || (c & 0xC0) != 0x80)) {
                tip += "...";
                cut = true;
            }
            if (!cut) {
                if (c == ' ' && (lead == std::string_view::npos || i < lead || i > trail)) tip.append(space);
                else appendVisible(tip, c);
                if (!lineEnd && (!utf8 || (c & 0xC0) != 0x80)) ++columns;
            }
            const bool lineBreak = c == '\n' || (c == '\r' && (i + 1 == original.size() || original[i + 1] != '\n'));
            if (lineBreak && i + 1 < original.size()) {
                if (lines == TIP_LINES) {
                    tip.append("\n...");
                    break;
                }
                tip += '\n';
                ++lines;
                columns = 0;
                cut = false;
            }
        }
        const std::size_t end = tip.size();
        if (entry.empty()) return { std::move(tip), start, end };

        // The list entry on a line of its own, its find text (if any) on one line in quotes
        tip.append("\n").append(entry);
        if (find.empty()) return { std::move(tip), start, end };
        tip.append(" \"");
        columns = 0;
        for (std::size_t i = 0; i < find.size(); ++i) {
            const auto c = static_cast<unsigned char>(find[i]);
            const bool starts = !utf8 || (c & 0xC0) != 0x80;   // a character, not the rest of one
            if (starts && columns == FIND_COLUMNS) {
                tip += "...";
                break;
            }
            appendVisible(tip, c);
            if (starts) ++columns;
        }
        tip += '"';
        return { std::move(tip), start, end };
    }
}
