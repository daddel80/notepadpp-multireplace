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
#include <cstdint>
#include <set>
#include <string>
#include <unordered_map>
#include <vector>

// The CSV view of each document while another one is active: the column highlight with the settings it was
// colored with, the sort state and the bookmarks Find Duplicates replaced. The panel settings stay in the panel.
namespace CsvDocState {

    using Buffer = std::uintptr_t;   // Notepad++ buffer id

    // The settings a highlight was colored with
    struct Coloring {
        std::set<int> columns;
        std::string delimiter;
        std::string quote;
        bool operator==(const Coloring&) const = default;
    };

    // The text as MultiReplace last saw it; a change made elsewhere leaves other values
    struct Snapshot {
        std::intptr_t length = -1;
        std::intptr_t lines = -1;
        std::intptr_t undoActions = -1;
        std::intptr_t undoCurrent = -1;
        bool operator==(const Snapshot&) const = default;
    };

    // The original line order and the direction the sort buttons show
    struct Sorting {
        std::vector<std::size_t> originalLineOrder;
        int direction = 0;       // the panel's SortDirection, 0 unsorted
        bool tracking = false;   // inserted and deleted lines keep the order up to date
        bool empty() const noexcept { return direction == 0 && !tracking && originalLineOrder.empty(); }
    };

    // What a document shows again when it becomes active
    struct View {
        bool highlighted = false;   // its colors still fit: set the style colors, color nothing
        bool stale = false;         // colors from before a change or from before CSV ended: clear them
        Coloring coloring;
        Sorting sorting;
    };

    class Store {
    public:
        // Leaving a document hands over its sort state; arriving returns what it shows if its text is the one seen
        void park(Buffer doc, Sorting&& sorting);
        View resume(Buffer doc, const Snapshot& now);

        void highlight(Buffer doc, const Coloring& coloring, const Snapshot& now);
        void unhighlight(Buffer doc);
        bool highlighted(Buffer doc) const;

        void seen(Buffer doc, const Snapshot& now);   // the active document as MultiReplace last saw it
        bool reloaded(Buffer doc, const Snapshot& now) const;   // Notepad++ reloaded it: the undo history is gone
        void csvEnded();                              // the highlights of all documents end, sort states stay
        void forget(Buffer doc);                      // the document closes; Notepad++ may give its id to the next one

        void replacedBookmarks(Buffer doc);
        bool takeReplacedBookmarks(Buffer doc);

    private:
        struct Record {
            bool highlighted = false;
            bool fits = true;         // the highlight still fits text and CSV mode
            Coloring coloring;
            Snapshot seen;
            Sorting sorting;
            bool bookmarks = false;   // Find Duplicates replaced the bookmarks
        };

        std::unordered_map<Buffer, Record> docs_;
    };
}
