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
#include <vector>

// Find/Replace in Open Documents: the tabs a run visits. While "All documents" is off, a tab
// counts only if its file name passes the Filter field and its tab color is one of the chosen.
namespace DocScope {

    // Tab colors as NPPM_GETTABCOLORID reports them: 0 for the first, -1 for none. How many there are,
    // Notepad++'s Apply Color menu tells (five so far); the mask has room for 32.
    constexpr int MAX_COLORS = 32;

    // Marked: every color, also one Notepad++ adds later
    constexpr unsigned ALL_COLORS = ~0u;

    // Entries of the Tab color dropdown: Any, Marked (any color), then one per color
    constexpr int CHOICE_ANY = 0;
    constexpr int CHOICE_MARKED = 1;
    constexpr int FIRST_COLOR_CHOICE = CHOICE_MARKED + 1;

    constexpr int clampColorCount(int colorCount)
    {
        return colorCount < 0 ? 0 : (colorCount > MAX_COLORS ? MAX_COLORS : colorCount);
    }

    constexpr int choiceCount(int colorCount)
    {
        return FIRST_COLOR_CHOICE + clampColorCount(colorCount);
    }

    // The tab color of a dropdown entry, -1 for Any, Marked and entries past the colors
    constexpr int colorFromChoice(int choice, int colorCount)
    {
        return (choice >= FIRST_COLOR_CHOICE && choice < choiceCount(colorCount)) ? choice - FIRST_COLOR_CHOICE : -1;
    }

    // One bit per color, every bit for Marked; 0 takes every tab, uncolored ones too
    constexpr unsigned maskFromChoice(int choice, int colorCount)
    {
        if (choice == CHOICE_MARKED) return ALL_COLORS;
        const int colorId = colorFromChoice(choice, colorCount);
        return colorId >= 0 ? 1u << colorId : 0;
    }

    // Several colors are Marked (with five colors Marked was 31); 0 and a color past the list are Any
    constexpr int choiceFromMask(unsigned mask, int colorCount)
    {
        if (mask == 0) return CHOICE_ANY;
        if ((mask & (mask - 1)) != 0) return CHOICE_MARKED;
        int colorId = 0;
        while ((mask >> colorId) != 1u) ++colorId;
        return colorId < clampColorCount(colorCount) ? FIRST_COLOR_CHOICE + colorId : CHOICE_ANY;
    }

    constexpr bool colorMatches(unsigned mask, int colorId)
    {
        return mask == 0 || (colorId >= 0 && colorId < MAX_COLORS && (mask & (1u << colorId)) != 0);
    }

    // An item of the Notepad++ menu that holds "Apply Color 1"
    struct ColorMenuItem {
        unsigned id = 0;
        bool other = false;    // separator or submenu
        bool sample = false;   // carries a color sample bitmap
    };

    // Positions of the color items in the order of the tab color ids: from "Apply Color 1" on, every item
    // with a sample or one of the ids first to last, past "Remove Color", up to the first other item.
    // A color Notepad++ adds comes with a sample, so it joins the list whatever its id.
    inline std::vector<size_t> colorMenuPositions(const std::vector<ColorMenuItem>& items,
        unsigned first, unsigned last, unsigned none)
    {
        std::vector<size_t> positions;
        size_t pos = 0;
        while (pos < items.size() && items[pos].id != first) ++pos;
        for (; pos < items.size() && positions.size() < static_cast<size_t>(MAX_COLORS); ++pos) {
            const ColorMenuItem& item = items[pos];
            if (item.other) break;
            if (item.id == none) continue;
            if (!item.sample && (item.id < first || item.id > last)) break;
            positions.push_back(pos);
        }
        return positions;
    }

    // NPPM_GETNPPVERSION with zero padding (8.6.8 is 8 and 680); NPPM_GETTABCOLORID came with 8.6.8
    constexpr bool tabColorsAvailable(int major, int paddedMinor)
    {
        return major > 8 || (major == 8 && paddedMinor >= 680);
    }

    struct Tab {
        int view;              // MAIN_VIEW or SUB_VIEW
        std::ptrdiff_t index;  // position in that view
    };

    struct Plan {
        std::vector<Tab> tabs;   // the tabs in scope, in the order a run visits them
        size_t total = 0;        // all tabs of the visited views
        size_t excluded() const { return total - tabs.size(); }
    };

    // views in visiting order; tabCount(view) and inScope(view, index) ask Notepad++
    template <class TabCount, class InScope>
    Plan plan(const std::vector<int>& views, TabCount tabCount, InScope inScope)
    {
        Plan scope;
        for (const int view : views) {
            const std::ptrdiff_t count = static_cast<std::ptrdiff_t>(tabCount(view));
            for (std::ptrdiff_t index = 0; index < count; ++index) {
                ++scope.total;
                if (inScope(view, index)) scope.tabs.push_back({ view, index });
            }
        }
        return scope;
    }
}
