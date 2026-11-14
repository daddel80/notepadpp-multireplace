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

// QA for LineOrder: the original order goes along with sorting and editing and gives the order to go back to.
// Build: g++ -std=c++20 -O1 -g -fsanitize=address,undefined -I.. line_order_qa.cpp ../LineOrder.cpp -o line_order_qa
#include <algorithm>
#include <cstdio>
#include <numeric>
#include <string>
#include <utility>
#include <vector>
#include "LineOrder.h"

static int g_fail = 0;
static void check(bool ok, const char* what) {
    std::printf("%s  %s\n", ok ? "ok  " : "FAIL", what);
    if (!ok) ++g_fail;
}

static std::vector<std::string> take(const std::vector<std::string>& lines, const std::vector<std::size_t>& from) {
    std::vector<std::string> out;
    for (std::size_t i : from) out.push_back(lines[i]);
    return out;
}

// A text as the panel handles it: one header line, lines without their line ends, the original order beside it
struct Text {
    std::vector<std::string> lines;
    LineOrder::Order order;
    bool sorted = false;

    // As sortRowsByColumn: the header stays first, an empty last line stays last
    void sort(bool descending) {
        std::vector<std::size_t> moved(lines.size());
        std::iota(moved.begin(), moved.end(), std::size_t{ 0 });
        const std::size_t end = (lines.size() > 2 && lines.back().empty()) ? lines.size() - 1 : lines.size();
        std::stable_sort(moved.begin() + 1, moved.begin() + static_cast<std::ptrdiff_t>(end), [&](std::size_t a, std::size_t b) {
            return descending ? lines[a] > lines[b] : lines[a] < lines[b];
        });
        LineOrder::moved(order, moved);
        lines = take(lines, moved);
        sorted = true;
    }
    // A whole line in front of line `at`: Enter at the end of the line before, or a pasted line at a line start
    void insert(std::size_t at, const std::string& text) {
        lines.insert(lines.begin() + static_cast<std::ptrdiff_t>(at), text);
        if (sorted) LineOrder::inserted(order, at, 1);
    }
    void erase(std::size_t at) {
        lines.erase(lines.begin() + static_cast<std::ptrdiff_t>(at));
        if (sorted) LineOrder::deleted(order, at, 1);
    }
    // As restoreOriginalLineOrder
    void unsort() {
        lines = take(lines, LineOrder::restored(order, lines.back().empty()));
        order.clear();
        sorted = false;
    }
    std::string str() const {
        std::string s;
        for (std::size_t i = 0; i < lines.size(); ++i) s += (i ? "|" : "") + lines[i];
        return s;
    }
};

static Text text(std::vector<std::string> lines) { Text t; t.lines = std::move(lines); return t; }
static Text withEnd() { return text({ "Item;Price", "A;5", "B;7", "C;9", "D;11", "" }); }   // file ends with a line break
static Text withoutEnd() { return text({ "Key;Value", "k1;v1", "k2;v2", "k3;v3" }); }

int main() {
    {   // Sorting and going back without edits
        Text t = withEnd();
        t.sort(true);
        check(t.str() == "Item;Price|D;11|C;9|B;7|A;5|", "sorted descending, empty last line stays last");
        t.unsort();
        check(t.str() == withEnd().str(), "back to the original");
    }
    {   // Sorting again in the other direction keeps the very first order
        Text t = withEnd();
        t.sort(false);
        t.sort(true);
        t.unsort();
        check(t.str() == withEnd().str(), "ascending, descending, back: the original");
    }
    {   // The reported case: a line typed in between while sorted
        Text t = withEnd();
        t.sort(true);
        t.insert(3, "X;1");                      // Enter behind C;9
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|X;1|", "new line at the end, before the empty last line");
    }
    {   // Enter at the end of the last sorted line
        Text t = withEnd();
        t.sort(true);
        t.insert(5, "X;1");
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|X;1|", "new last row goes to the end, empty line last");
    }
    {   // A file without line break at the end
        Text t = withoutEnd();
        t.sort(true);
        t.insert(2, "k9;v9");
        t.unsort();
        check(t.str() == "Key;Value|k1;v1|k2;v2|k3;v3|k9;v9", "no line break at the end: new line last");
    }
    {   // Several new lines keep the order they were inserted in
        Text t = withEnd();
        t.sort(true);
        t.insert(4, "X;1");
        t.insert(1, "Y;2");
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|X;1|Y;2|", "new lines in the order they came");
    }
    {   // Deleted lines leave no trace
        Text t = withEnd();
        t.sort(true);
        t.erase(2);                              // C;9
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|D;11|", "deleted line gone, the rest in original order");
    }
    {   // Deleting and inserting mixed, sorted again in between
        Text t = withEnd();
        t.sort(false);
        t.erase(1);                              // A;5
        t.insert(2, "X;1");
        t.sort(true);
        t.insert(1, "Y;2");
        t.unsort();
        check(t.str() == "Item;Price|B;7|C;9|D;11|X;1|Y;2|", "mixed edits across two sorts");
    }
    {   // A line above the header is a new line as well
        Text t = withEnd();
        t.sort(true);
        t.insert(0, "Z;0");
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|Z;0|", "line above the header goes to the end");
    }
    {   // Typing into the empty last line makes it a row of its own place
        Text t = withEnd();
        t.sort(true);
        t.insert(2, "X;1");
        t.lines.back() = "E;3";
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|E;3|X;1", "former empty last line keeps its place, new line after it");
    }
    {   // Enter on the empty last line: the file still ends with a line break
        Text t = withEnd();
        t.sort(true);
        t.lines.back() = "X;1";
        t.insert(6, "");
        t.unsort();
        check(t.str() == "Item;Price|A;5|B;7|C;9|D;11|X;1|", "row typed on the last line stays before the line break");
    }
    {   // A block of sorted rows replaced as a whole, as by sorting a selection with Notepad++: the rows count as new
        Text t = text({ "H", "b", "a", "d", "c", "" });
        t.sort(true);                                         // H d c b a ""
        LineOrder::deleted(t.order, 2, 2);                    // from the start of "c" to the end of the text of "a"
        LineOrder::renewed(t.order, 2);                       // "a" lost its content
        LineOrder::inserted(t.order, 3, 2);                   // "a b c" into the empty line: it takes the first line
        t.lines = { "H", "d", "a", "b", "c", "" };
        t.unsort();
        check(t.str() == "H|d|a|b|c|", "replaced block keeps its order after the other rows");
    }
    {   // Pasted into the empty last line: the line takes the first pasted row
        Text t = text({ "H", "x", "y", "" });
        t.sort(true);                                         // H y x ""
        LineOrder::inserted(t.order, 4, 1);
        t.lines = { "H", "y", "x", "p", "q" };
        t.unsort();
        check(t.str() == "H|x|y|p|q", "rows pasted into the last line stay in their order");
    }
    {   // A line renewed twice or past the end
        LineOrder::Order order{ 2, 0, 1 };
        LineOrder::renewed(order, 0);
        LineOrder::renewed(order, 0);
        check(order == LineOrder::Order({ 4, 0, 1 }), "renewed takes the next free number");
        LineOrder::renewed(order, 7);
        check(order.size() == 3, "renew past the end does nothing");
    }
    {   // An order that no longer fits the text starts over instead of reading past it
        LineOrder::Order order{ 2, 0, 1 };
        LineOrder::moved(order, { 4, 3, 2, 1, 0 });
        check(order == LineOrder::Order({ 4, 3, 2, 1, 0 }), "misfit: starts from this text");
        order = { 1, 0 };
        LineOrder::moved(order, { 1, 5 });
        check(order == LineOrder::Order({ 1, 5 }), "misfit index: starts over");
    }
    {   // Line numbers past the end
        LineOrder::Order order{ 1, 0 };
        LineOrder::inserted(order, 9, 2);
        check(order == LineOrder::Order({ 1, 0, 2, 3 }), "insert past the end appends");
        LineOrder::deleted(order, 9, 1);
        check(order.size() == 4, "delete past the end does nothing");
        LineOrder::deleted(order, 3, 5);
        check(order == LineOrder::Order({ 1, 0, 2 }), "delete running past the end stops there");
        LineOrder::inserted(order, 1, 0);
        check(order.size() == 3, "inserting no lines changes nothing");
    }
    {   // Small texts
        check(LineOrder::restored({}, true).empty(), "empty order, empty result");
        check(LineOrder::restored({ 7 }, true) == std::vector<std::size_t>({ 0 }), "single line stays");
        check(LineOrder::restored({ 9, 4, 6 }, false) == std::vector<std::size_t>({ 1, 2, 0 }), "gaps from deleted lines");
    }

    std::printf(g_fail ? "%d FAILED\n" : "all passed\n", g_fail);
    return g_fail ? 1 : 0;
}
