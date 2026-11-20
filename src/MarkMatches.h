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

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "EditorPalette.h"
#include "SciIndicators.h"

// Mark Matches in one indicator: every match carries its fill color as the indicator value (SC_INDICFLAG_VALUEFORE),
// so any number of colors takes one indicator. Each fill sets the indicator and value anew: Notepad++ passes every
// fill on to all plugins, and their handlers may change both in between.
namespace MarkMatches {

    using Send = SciIndicators::Send;
    using Color = EditorPalette::Color;

    // In a view: a rounded fill under the text in the color its match carries, one alpha for all colors
    void applyStyle(const Send& view, int id, int alpha, int outlineAlpha);

    // Marks ranges in a color; the document keeps its current indicator and value
    class Painter {
    public:
        Painter(const Send& doc, int id) : doc_(doc), keep_(doc), id_(id) {}
        void fill(std::intptr_t pos, std::intptr_t length, Color color);

    private:
        const Send& doc_;
        SciIndicators::KeepCurrent keep_;
        int id_;
    };

    struct Run {
        std::intptr_t start;
        std::intptr_t end;
        Color color;
    };
    // The marked runs in document order; neighbouring matches of different colors are runs of their own
    std::vector<Run> runs(const Send& doc, int id);

    // The runs grouped by color, the colors in the given order and each group in document order; colors the order
    // lacks follow in the order they first appear
    std::vector<Run> byColor(std::vector<Run> runs, const std::vector<Color>& order);

    // Gives each run the color its color maps to, the others keep theirs. Returns the number of runs recolored
    std::size_t recolor(const Send& doc, int id, const std::unordered_map<Color, Color>& colors);

    void clear(const Send& doc, int id);
}
