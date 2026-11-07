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

    // Tab colors as NPPM_GETTABCOLORID reports them: 0 to 4, -1 for none
    constexpr int COLOR_COUNT = 5;
    constexpr unsigned ALL_COLORS = (1u << COLOR_COUNT) - 1;

    // Entries of the Tab color dropdown: Any, Marked (any color), Color 1 to Color 5
    constexpr int CHOICE_ANY = 0;
    constexpr int CHOICE_MARKED = 1;
    constexpr int CHOICE_COUNT = CHOICE_MARKED + 1 + COLOR_COUNT;

    // The tab color of a dropdown entry, -1 for Any and Marked
    constexpr int colorFromChoice(int choice)
    {
        return (choice > CHOICE_MARKED && choice < CHOICE_COUNT) ? choice - CHOICE_MARKED - 1 : -1;
    }

    // One bit per color; 0 takes every tab, uncolored ones too
    constexpr unsigned maskFromChoice(int choice)
    {
        if (choice == CHOICE_MARKED) return ALL_COLORS;
        const int colorId = colorFromChoice(choice);
        return colorId >= 0 ? 1u << colorId : 0;
    }

    // A mask the dropdown cannot show counts as Any
    constexpr int choiceFromMask(unsigned mask)
    {
        for (int choice = 0; choice < CHOICE_COUNT; ++choice)
            if (maskFromChoice(choice) == mask) return choice;
        return CHOICE_ANY;
    }

    constexpr bool colorMatches(unsigned mask, int colorId)
    {
        return mask == 0 || (colorId >= 0 && colorId < COLOR_COUNT && (mask & (1u << colorId)) != 0);
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
