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

// The order of a list whose entries carry ids: a remembered order comes back with the entries added and removed since.
// Ids grow with every entry a list gets, so they also tell the order in which entries came in.
namespace ListOrder {

    using Ids = std::vector<std::size_t>;

    // For each place of the arranged list the place in current of the entry that goes there: the entries order names
    // follow its order, the others came in later and follow them oldest first, or with newFirst lead newest first;
    // ids current lacks drop out
    std::vector<std::size_t> arranged(const Ids& order, const Ids& current, bool newFirst);
}
