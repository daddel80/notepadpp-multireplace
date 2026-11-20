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

#include "EditorIndicators.h"

namespace EditorIndicators {

    int Pool::id(Kind kind)
    {
        int& id = ids_[static_cast<std::size_t>(kind)];
        if (id == UNASKED) id = ask();
        return id;
    }

    int Pool::peek(Kind kind) const noexcept
    {
        const int id = ids_[static_cast<std::size_t>(kind)];
        return id == UNASKED ? NONE_LEFT : id;
    }

    int Pool::ask()
    {
        if (!picking_) {
            const int got = allocate_ ? allocate_() : NOT_HANDED_OUT;
            if (got != NOT_HANDED_OUT) return got >= 0 ? got : NONE_LEFT;
            picking_ = true;
        }
        return nextPick_ <= LAST_PLUGIN ? nextPick_++ : NONE_LEFT;
    }

    Pool& shared()
    {
        static Pool pool;
        return pool;
    }
}
