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

#include <array>
#include <cstddef>
#include <functional>
#include <utility>

// The Scintilla indicators MR draws with in the editor views, one per kind. Notepad++ 8.5.6 and later hands them out
// (NPPM_ALLOCATEINDICATOR), so no other plugin that asks it for one gets the same; older versions leave MR to pick
// from the plugins' range itself.
namespace EditorIndicators {

    enum class Kind { ScopeFrame, ReplaceMarks, ReplaceDeletions, FlowTabs, MarkMatches, Duplicates, Count };

    // Notepad++'s indicators for plugins (resource.h: INDICATOR_PLUGINS, INDICATOR_PLUGINS_LIMIT)
    constexpr int FIRST_PLUGIN = 9;
    constexpr int LAST_PLUGIN = 20;

    // What asking Notepad++ answers besides an id
    constexpr int NONE_LEFT = -1;        // all of them are handed out
    constexpr int NOT_HANDED_OUT = -2;   // a version without NPPM_ALLOCATEINDICATOR

    class Pool {
    public:
        // One indicator from Notepad++: its id, NONE_LEFT or NOT_HANDED_OUT
        using Allocate = std::function<int()>;

        Pool() { ids_.fill(UNASKED); }
        void setAllocate(Allocate allocate) { allocate_ = std::move(allocate); }

        // The kind's indicator, asked for on its first use; -1 when none was left, then for good
        int id(Kind kind);
        // The kind's indicator without asking: -1 before its first use
        int peek(Kind kind) const noexcept;

    private:
        static constexpr int UNASKED = -3;
        int ask();

        Allocate allocate_;
        std::array<int, static_cast<std::size_t>(Kind::Count)> ids_{};
        bool picking_ = false;   // Notepad++ hands out none: MR picks
        int nextPick_ = FIRST_PLUGIN;
    };

    // The pool of the two editor views
    Pool& shared();
}
