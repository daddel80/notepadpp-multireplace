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

#include "CsvDocState.h"

#include <utility>

namespace CsvDocState {

    void Store::park(Buffer doc, Sorting&& sorting)
    {
        if (doc == 0) return;
        const auto it = docs_.find(doc);
        if (it == docs_.end()) {
            // Never seen: the text it belongs to is unknown, the next resume drops it
            if (!sorting.empty()) docs_[doc].sorting = std::move(sorting);
            return;
        }
        Record& r = it->second;
        r.sorting = std::move(sorting);
        if (!r.highlighted && !r.bookmarks && r.sorting.empty()) docs_.erase(it);
    }

    View Store::resume(Buffer doc, const Snapshot& now)
    {
        View view;
        const auto it = docs_.find(doc);
        if (it == docs_.end()) return view;

        Record& r = it->second;
        if (r.seen != now) {
            r.fits = false;
            r.sorting = {};
        }
        view.highlighted = r.highlighted && r.fits;
        view.stale = r.highlighted && !r.fits;
        if (view.highlighted) view.coloring = r.coloring;
        view.sorting = std::move(r.sorting);
        r.sorting = {};
        if (view.stale) {
            r.highlighted = false;
            r.fits = true;
        }
        r.seen = now;
        return view;
    }

    void Store::highlight(Buffer doc, const Coloring& coloring, const Snapshot& now)
    {
        Record& r = docs_[doc];
        r.highlighted = true;
        r.fits = true;
        r.coloring = coloring;
        r.seen = now;
    }

    void Store::unhighlight(Buffer doc)
    {
        const auto it = docs_.find(doc);
        if (it == docs_.end()) return;
        it->second.highlighted = false;
        it->second.fits = true;
    }

    bool Store::highlighted(Buffer doc) const
    {
        const auto it = docs_.find(doc);
        return it != docs_.end() && it->second.highlighted;
    }

    void Store::seen(Buffer doc, const Snapshot& now)
    {
        if (doc != 0) docs_[doc].seen = now;
    }

    bool Store::reloaded(Buffer doc, const Snapshot& now) const
    {
        const auto it = docs_.find(doc);
        return it != docs_.end() && it->second.seen.undoActions > 0 && now.undoActions == 0;
    }

    void Store::csvEnded()
    {
        for (auto& entry : docs_) {
            if (entry.second.highlighted) entry.second.fits = false;
        }
    }

    void Store::forget(Buffer doc)
    {
        docs_.erase(doc);
    }

    void Store::replacedBookmarks(Buffer doc)
    {
        docs_[doc].bookmarks = true;
    }

    bool Store::takeReplacedBookmarks(Buffer doc)
    {
        const auto it = docs_.find(doc);
        if (it == docs_.end() || !it->second.bookmarks) return false;
        it->second.bookmarks = false;
        return true;
    }
}
