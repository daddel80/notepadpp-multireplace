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

#include <vector>

// The list tab strip in the panel's tab row. The tab control always holds every tab at its natural
// width; the panel moves it and shows one window of it, with "..." on the side where tabs are hidden.
namespace TabStripLayout {

    // Horizontal extent of a tab in the tab control
    struct Edge {
        int left;
        int right;
    };

    // Tab row geometry in panel pixels
    struct Metrics {
        int left;        // where the strip starts
        int maxWidth;    // room for the strip while no "..." shows on the left
        int indicator;   // width a "..." takes, gap included
        int endMargin;   // room the tab control keeps after the last tab
        int plusGap;     // distance from the last tab to "+", the free "..." slot included
        int plusMax;     // rightmost position of "+"
    };

    struct Layout {
        bool overflow = false;     // not every tab fits
        int first = 0;             // first tab in view
        bool hiddenLeft = false;   // "..." left of the strip
        bool hiddenRight = false;  // "..." right of the strip
        int viewLeft = 0;          // panel x where the visible strip starts
        int viewWidth = 0;         // width of the visible strip
        int controlLeft = 0;       // panel x of the tab control
        int controlWidth = 0;      // every tab plus the end margin
        int clipLeft = 0;          // visible strip in tab control coordinates
        bool clipped = false;      // the view shows only part of the control
        int plusLeft = 0;          // panel x of "+"
    };

    // The row with `first` as the first tab in view, clamped to the scroll range
    Layout arrange(const std::vector<Edge>& tabs, const Metrics& m, int first);

    // First tab in view after scrolling by `steps` tabs
    int scrolled(const std::vector<Edge>& tabs, const Metrics& m, int first, int steps);

    // First tab in view that shows tab `index` completely, scrolling as little as possible
    int revealing(const std::vector<Edge>& tabs, const Metrics& m, int first, int index);
}
