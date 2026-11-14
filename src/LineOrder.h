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

// The original order of a sorted text: for each line its place before the first sort. It goes along with
// lines moved, inserted and deleted and gives the order the lines go back to.
namespace LineOrder {

    using Order = std::vector<std::size_t>;

    // Line i now holds the former line moved[i]; without an order that fits the text it starts from this text
    void moved(Order& order, const std::vector<std::size_t>& moved);

    // Inserted lines come after all others; deleted lines leave gaps; a line that lost its content counts as new
    void inserted(Order& order, std::size_t line, std::size_t count);
    void deleted(Order& order, std::size_t line, std::size_t count);
    void renewed(Order& order, std::size_t line);

    // For each line of the restored text the line it comes from; an empty last line stays last, as when sorting
    std::vector<std::size_t> restored(const Order& order, bool emptyLastLine);
}
