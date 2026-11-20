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

#include "SelectionFrame.h"

#include "Scintilla.h"
#include "SciIndicators.h"

namespace SelectionFrame {

    namespace {

        using SciIndicators::KeepCurrent;
        using SciIndicators::forEachRun;

        // Classic blue (COLORREF, 0x00BBGGRR): #4F81BD on light, #6A9BD1 on dark editor themes
        constexpr EditorPalette::Color LIGHT_COLOR = 0xBD814F;
        constexpr EditorPalette::Color DARK_COLOR = 0xD19B6A;

        // Light fill under the text: #4F81BD on white gives #E6EDF6
        constexpr int FILL_ALPHA = 36;

        // Indicator values stay below SC_INDICVALUEBIT
        constexpr int GENERATION_LIMIT = 0xFFFFFF;

        void clearAll(const Send& doc, int id)
        {
            doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id), 0);
            doc(SCI_INDICATORCLEARRANGE, 0, doc(SCI_GETLENGTH, 0, 0));
        }
    }

    void Frame::applyStyle(const Send& view, const EditorPalette::Theme& theme) const
    {
        if (id_ < 0) return;
        const auto id = static_cast<std::uintptr_t>(id_);
        const EditorPalette::Color color = EditorPalette::visible(theme.dark ? DARK_COLOR : LIGHT_COLOR, theme);
        view(SCI_INDICSETSTYLE, id, INDIC_STRAIGHTBOX);
        view(SCI_INDICSETFORE, id, static_cast<std::intptr_t>(color));
        view(SCI_INDICSETALPHA, id, EditorPalette::fillAlpha(color, FILL_ALPHA, theme));
        view(SCI_INDICSETOUTLINEALPHA, id, 255);
        view(SCI_INDICSETSTROKEWIDTH, id, 100);
        view(SCI_INDICSETUNDER, id, 1);
    }

    void Frame::draw(const Send& doc, const std::vector<Span>& spans)
    {
        deleted_ = {};
        if (id_ < 0) return;
        const KeepCurrent keep(doc);
        clearAll(doc, id_);
        doc(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(generation_), 0);
        for (const Span& span : spans) {
            if (span.end > span.start) doc(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(span.start), span.end - span.start);
        }
    }

    void Frame::erase(const Send& doc)
    {
        deleted_ = {};
        if (id_ < 0) return;
        const KeepCurrent keep(doc);
        clearAll(doc, id_);
    }

    std::vector<Span> Frame::read(const Send& doc) const
    {
        std::vector<Span> spans;
        if (id_ < 0) return spans;
        forEachRun(doc, id_, [&](std::intptr_t start, std::intptr_t end, std::intptr_t value) {
            if (value == generation_) spans.push_back({ start, end });
            });
        return spans;
    }

    void Frame::dropStale(const Send& doc) const
    {
        if (id_ < 0) return;
        std::vector<Span> stale;
        forEachRun(doc, id_, [&](std::intptr_t start, std::intptr_t end, std::intptr_t value) {
            if (value != 0 && value != generation_) stale.push_back({ start, end });
            });
        if (stale.empty()) return;
        const KeepCurrent keep(doc);
        doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id_), 0);
        for (const Span& span : stale) {
            doc(SCI_INDICATORCLEARRANGE, static_cast<std::uintptr_t>(span.start), span.end - span.start);
        }
    }

    void Frame::noteDelete(const Send& doc, std::intptr_t pos, std::intptr_t length)
    {
        deleted_ = {};
        if (id_ < 0 || length <= 0) return;
        const auto id = static_cast<std::uintptr_t>(id_);
        if (doc(SCI_INDICATORVALUEAT, id, pos) == generation_ && doc(SCI_INDICATOREND, id, pos) >= pos + length) {
            deleted_ = { doc(SCI_GETDOCPOINTER, 0, 0), pos, doc(SCI_GETLENGTH, 0, 0) - length };
        }
    }

    void Frame::keepInsert(const Send& doc, std::intptr_t pos, std::intptr_t length)
    {
        const Deleted deleted = deleted_;
        deleted_ = {};
        if (id_ < 0 || length <= 0 || deleted.pos != pos || deleted.doc != doc(SCI_GETDOCPOINTER, 0, 0)
            || deleted.docLength != doc(SCI_GETLENGTH, 0, 0) - length) return;
        const KeepCurrent keep(doc);
        doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id_), 0);
        doc(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(generation_), 0);
        doc(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(pos), length);
    }

    void Frame::retire() noexcept
    {
        generation_ = (generation_ >= GENERATION_LIMIT) ? 1 : generation_ + 1;
        deleted_ = {};
    }
}
