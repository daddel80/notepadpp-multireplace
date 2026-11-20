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

#include "MarkMatches.h"

#include <algorithm>

#include "Scintilla.h"

namespace MarkMatches {

    namespace {

        std::intptr_t valueOf(Color color) { return SC_INDICVALUEBIT | static_cast<std::intptr_t>(color & SC_INDICVALUEMASK); }
        Color colorOf(std::intptr_t value) { return static_cast<Color>(value & SC_INDICVALUEMASK); }

        void fillRange(const Send& doc, int id, std::intptr_t pos, std::intptr_t length, std::intptr_t value)
        {
            doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id), 0);
            doc(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(value), 0);
            doc(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(pos), length);
        }
    }

    void applyStyle(const Send& view, int id, int alpha, int outlineAlpha)
    {
        if (id < 0) return;
        const auto indicator = static_cast<std::uintptr_t>(id);
        view(SCI_INDICSETSTYLE, indicator, INDIC_ROUNDBOX);
        view(SCI_INDICSETFLAGS, indicator, SC_INDICFLAG_VALUEFORE);
        view(SCI_INDICSETALPHA, indicator, alpha);
        view(SCI_INDICSETOUTLINEALPHA, indicator, outlineAlpha);
        view(SCI_INDICSETUNDER, indicator, 1);
    }

    void Painter::fill(std::intptr_t pos, std::intptr_t length, Color color)
    {
        if (id_ < 0 || length <= 0) return;
        fillRange(doc_, id_, pos, length, valueOf(color));
    }

    std::vector<Run> runs(const Send& doc, int id)
    {
        std::vector<Run> out;
        if (id < 0) return out;
        SciIndicators::forEachRun(doc, id, [&](std::intptr_t start, std::intptr_t end, std::intptr_t value) {
            if (value) out.push_back({ start, end, colorOf(value) });
        });
        return out;
    }

    std::vector<Run> byColor(std::vector<Run> runs, const std::vector<Color>& order)
    {
        std::unordered_map<Color, std::size_t> place;
        for (const Color color : order) place.emplace(color, place.size());
        for (const Run& run : runs) place.emplace(run.color, place.size());
        std::stable_sort(runs.begin(), runs.end(), [&place](const Run& a, const Run& b) {
            return place.at(a.color) < place.at(b.color);
        });
        return runs;
    }

    std::size_t recolor(const Send& doc, int id, const std::unordered_map<Color, Color>& colors)
    {
        std::size_t changed = 0;
        if (id < 0 || colors.empty()) return changed;
        const SciIndicators::KeepCurrent keep(doc);
        for (const Run& run : runs(doc, id)) {
            const auto to = colors.find(run.color);
            if (to == colors.end() || to->second == run.color) continue;
            fillRange(doc, id, run.start, run.end - run.start, valueOf(to->second));
            ++changed;
        }
        return changed;
    }

    void clear(const Send& doc, int id)
    {
        if (id < 0) return;
        const SciIndicators::KeepCurrent keep(doc);
        doc(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(id), 0);
        doc(SCI_INDICATORCLEARRANGE, 0, doc(SCI_GETLENGTH, 0, 0));
    }
}
