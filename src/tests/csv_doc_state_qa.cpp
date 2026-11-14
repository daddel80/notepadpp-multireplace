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

// QA for CsvDocState: what a document keeps while another one is active.
// Build: g++ -std=c++20 -O1 -g -fsanitize=address,undefined -I.. csv_doc_state_qa.cpp ../CsvDocState.cpp -o csv_doc_state_qa
#include <cstdio>
#include "CsvDocState.h"

using namespace CsvDocState;

static int g_fail = 0;
static void check(bool ok, const char* what) {
    std::printf("%s  %s\n", ok ? "ok  " : "FAIL", what);
    if (!ok) ++g_fail;
}

static Snapshot snap(long length, long lines, long actions, long current) { return { length, lines, actions, current }; }
static Coloring coloring(const char* delimiter) { return { { 1, 3 }, delimiter, "\"" }; }
static Sorting sorted(int direction) { return { { 2, 0, 1 }, direction, true }; }

int main() {
    const Buffer A = 0x1F2'3456'789AULL, B = 0x1F2'0000'0010ULL, C = 0x77;
    const Snapshot s1 = snap(100, 5, 3, 3), s2 = snap(101, 5, 4, 4);

    {   // Unknown document: nothing to show
        Store st;
        const View v = st.resume(A, s1);
        check(!v.highlighted && !v.stale && v.sorting.empty(), "unknown document shows nothing");
    }
    {   // Highlight and sort state come back with an unchanged text
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.park(A, sorted(1));
        const View v = st.resume(A, s1);
        check(v.highlighted && !v.stale && v.coloring == coloring(";"), "highlight comes back with its settings");
        check(v.sorting.direction == 1 && v.sorting.tracking && v.sorting.originalLineOrder.size() == 3, "sort state comes back");
        check(st.highlighted(A), "still highlighted after resume");
    }
    {   // Text changed elsewhere: colors to clear, sort state dropped
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.park(A, sorted(2));
        const View v = st.resume(A, s2);
        check(!v.highlighted && v.stale && v.sorting.empty(), "changed text: stale colors, no sort state");
        check(!st.highlighted(A), "no longer highlighted after a stale resume");
        const View again = st.resume(A, s2);
        check(!again.highlighted && !again.stale, "a second resume has nothing left to clear");
    }
    {   // The active document as last seen counts, not as first highlighted
        Store st;
        st.highlight(A, coloring(","), s1);
        st.seen(A, s2);
        st.park(A, {});
        const View v = st.resume(A, s2);
        check(v.highlighted, "seen() moves the snapshot along");
    }
    {   // CSV ends: highlights stale everywhere, sort states stay
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.park(A, sorted(1));
        st.highlight(B, coloring(","), s2);
        st.csvEnded();
        const View va = st.resume(A, s1);
        check(va.stale && !va.highlighted && va.sorting.direction == 1, "CSV ended: colors stale, sort state kept");
        const View vb = st.resume(B, s2);
        check(vb.stale, "CSV ended for every document");
    }
    {   // A highlight switched off is not cleared again
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.unhighlight(A);
        st.park(A, {});
        const View v = st.resume(A, s1);
        check(!v.highlighted && !v.stale, "unhighlighted document has nothing to show or clear");
    }
    {   // Closing forgets; the next document with the same id starts clean
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.park(A, sorted(1));
        st.replacedBookmarks(A);
        st.forget(A);
        const View v = st.resume(A, s1);
        check(!v.highlighted && !v.stale && v.sorting.empty() && !st.takeReplacedBookmarks(A), "forget clears everything");
    }
    {   // Sort state of a document never seen: dropped on resume
        Store st;
        st.park(C, sorted(1));
        const View v = st.resume(C, s1);
        check(v.sorting.empty(), "unseen text: sort state dropped");
    }
    {   // Nothing parked, nothing kept
        Store st;
        st.seen(A, s1);
        st.park(A, {});
        const View v = st.resume(A, s1);
        check(!v.highlighted && !v.stale && v.sorting.empty(), "empty record goes");
    }
    {   // Bookmarks Find Duplicates replaced, per document
        Store st;
        st.replacedBookmarks(A);
        check(!st.takeReplacedBookmarks(B), "other document has none");
        check(st.takeReplacedBookmarks(A), "replaced bookmarks taken once");
        check(!st.takeReplacedBookmarks(A), "and not twice");
        st.replacedBookmarks(B);
        st.park(B, {});
        check(st.takeReplacedBookmarks(B), "park keeps the bookmark record");
    }
    {   // Reload: the undo history is gone, edits and undo never empty it
        Store st;
        st.seen(A, snap(100, 5, 3, 3));
        check(st.reloaded(A, snap(120, 6, 0, 0)), "emptied undo history is a reload");
        check(!st.reloaded(A, snap(100, 5, 3, 0)), "undo back to the start is no reload");
        check(!st.reloaded(A, snap(101, 5, 1, 1)), "an edit after undo is no reload");
        check(!st.reloaded(B, snap(100, 5, 0, 0)), "unknown document is never reloaded");
        st.seen(C, snap(10, 1, 0, 0));
        check(!st.reloaded(C, snap(12, 1, 0, 0)), "without history a reload cannot be told");
    }
    {   // Parking a document with its highlight keeps the record without sort state
        Store st;
        st.highlight(A, coloring(";"), s1);
        st.park(A, {});
        check(st.highlighted(A), "highlighted record survives an empty park");
        st.park(0, sorted(1));
        check(!st.resume(0, s1).highlighted, "buffer 0 is never parked");
    }

    std::printf(g_fail ? "%d FAILED\n" : "all passed\n", g_fail);
    return g_fail ? 1 : 0;
}
