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

#include "TabStripLayout.h"

#include <algorithm>

namespace TabStripLayout {

    namespace {

        // Shift that brings tab `first` to where the first tab sits unscrolled
        int offsetOf(const std::vector<Edge>& tabs, int first)
        {
            return tabs[first].left - tabs.front().left;
        }

        // Width the tabs from `first` to the end take in view
        int widthFrom(const std::vector<Edge>& tabs, int first)
        {
            return tabs.back().right - offsetOf(tabs, first);
        }

        int roomFor(const Metrics& m, int first)
        {
            return m.maxWidth - (first > 0 ? m.indicator : 0);
        }

        // The strip in view: as far as the room goes, the end margin while it fits
        int viewWidthFor(const std::vector<Edge>& tabs, const Metrics& m, int first)
        {
            return std::min(widthFrom(tabs, first) + m.endMargin, roomFor(m, first));
        }

        // End of the scroll range: from there on every following tab fits
        int lastFirst(const std::vector<Edge>& tabs, const Metrics& m)
        {
            const int count = static_cast<int>(tabs.size());
            for (int first = 0; first < count; ++first) {
                if (widthFrom(tabs, first) <= roomFor(m, first)) return first;
            }
            return count - 1;
        }
    }

    Layout arrange(const std::vector<Edge>& tabs, const Metrics& m, int first)
    {
        Layout lay;
        lay.viewLeft = m.left;
        lay.controlLeft = m.left;
        if (tabs.empty()) {
            lay.plusLeft = std::min(m.left + m.plusGap, m.plusMax);
            return lay;
        }

        lay.controlWidth = tabs.back().right + m.endMargin;
        lay.overflow = tabs.back().right > m.maxWidth;
        lay.first = std::clamp(first, 0, lastFirst(tabs, m));
        lay.hiddenLeft = lay.first > 0;
        // A cut end margin alone hides nothing
        lay.hiddenRight = widthFrom(tabs, lay.first) > roomFor(m, lay.first);
        lay.viewLeft = m.left + (lay.hiddenLeft ? m.indicator : 0);
        lay.viewWidth = viewWidthFor(tabs, m, lay.first);
        lay.clipLeft = offsetOf(tabs, lay.first);
        lay.clipped = lay.clipLeft > 0 || lay.viewWidth < lay.controlWidth;
        lay.controlLeft = lay.viewLeft - lay.clipLeft;
        lay.plusLeft = lay.hiddenRight ? m.plusMax
            : std::min(lay.controlLeft + tabs.back().right + m.plusGap, m.plusMax);
        return lay;
    }

    int scrolled(const std::vector<Edge>& tabs, const Metrics& m, int first, int steps)
    {
        if (tabs.empty()) return 0;
        const int last = lastFirst(tabs, m);
        return std::clamp(std::clamp(first, 0, last) + steps, 0, last);
    }

    int revealing(const std::vector<Edge>& tabs, const Metrics& m, int first, int index)
    {
        if (tabs.empty()) return 0;
        const int last = lastFirst(tabs, m);
        index = std::clamp(index, 0, static_cast<int>(tabs.size()) - 1);
        first = std::clamp(first, 0, last);
        if (index < first) return index;

        // Scroll right until the tab and the margin after it end in view, never past the tab itself
        const int stop = std::min(index, last);
        while (first < stop && tabs[index].right + m.endMargin - offsetOf(tabs, first) > viewWidthFor(tabs, m, first)) {
            ++first;
        }
        return first;
    }
}
