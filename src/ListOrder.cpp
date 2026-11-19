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

#include "ListOrder.h"

#include <algorithm>
#include <unordered_map>

namespace ListOrder {

    std::vector<std::size_t> arranged(const Ids& order, const Ids& current, bool newFirst)
    {
        std::unordered_map<std::size_t, std::size_t> rank;   // id -> its place in order; a repeated id keeps the first
        rank.reserve(order.size());
        for (std::size_t i = 0; i < order.size(); ++i) rank.emplace(order[i], i);

        std::vector<std::size_t> known, others, place(current.size());
        for (std::size_t i = 0; i < current.size(); ++i) {
            const auto found = rank.find(current[i]);
            if (found == rank.end()) {
                others.push_back(i);
            }
            else {
                place[i] = found->second;
                known.push_back(i);
            }
        }
        std::stable_sort(known.begin(), known.end(), [&](std::size_t a, std::size_t b) { return place[a] < place[b]; });
        std::stable_sort(others.begin(), others.end(), [&](std::size_t a, std::size_t b) {
            return newFirst ? current[a] > current[b] : current[a] < current[b];
        });

        const std::vector<std::size_t>& head = newFirst ? others : known;
        const std::vector<std::size_t>& tail = newFirst ? known : others;
        std::vector<std::size_t> from;
        from.reserve(current.size());
        from.insert(from.end(), head.begin(), head.end());
        from.insert(from.end(), tail.begin(), tail.end());
        return from;
    }
}
