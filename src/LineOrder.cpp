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

#include "LineOrder.h"

#include <algorithm>
#include <cstddef>
#include <numeric>
#include <utility>

namespace LineOrder {

    void moved(Order& order, const std::vector<std::size_t>& moved)
    {
        const bool fits = order.size() == moved.size()
            && std::all_of(moved.begin(), moved.end(), [&](std::size_t from) { return from < order.size(); });
        if (!fits) {
            order = moved;
            return;
        }
        Order next(moved.size());
        for (std::size_t i = 0; i < moved.size(); ++i) {
            next[i] = order[moved[i]];
        }
        order = std::move(next);
    }

    void inserted(Order& order, std::size_t line, std::size_t count)
    {
        if (count == 0) return;
        const std::size_t next = order.empty() ? 0 : *std::max_element(order.begin(), order.end()) + 1;
        const auto at = static_cast<std::ptrdiff_t>(std::min(line, order.size()));
        order.insert(order.begin() + at, count, 0);
        std::iota(order.begin() + at, order.begin() + at + static_cast<std::ptrdiff_t>(count), next);
    }

    void deleted(Order& order, std::size_t line, std::size_t count)
    {
        if (line >= order.size()) return;
        const std::size_t end = (count < order.size() - line) ? line + count : order.size();
        order.erase(order.begin() + static_cast<std::ptrdiff_t>(line), order.begin() + static_cast<std::ptrdiff_t>(end));
    }

    void renewed(Order& order, std::size_t line)
    {
        if (line >= order.size()) return;
        order[line] = *std::max_element(order.begin(), order.end()) + 1;
    }

    std::vector<std::size_t> restored(const Order& order, bool emptyLastLine)
    {
        std::vector<std::size_t> from(order.size());
        std::iota(from.begin(), from.end(), std::size_t{ 0 });
        const std::size_t end = (emptyLastLine && !from.empty()) ? from.size() - 1 : from.size();
        std::stable_sort(from.begin(), from.begin() + static_cast<std::ptrdiff_t>(end),
            [&](std::size_t a, std::size_t b) { return order[a] < order[b]; });
        return from;
    }
}
