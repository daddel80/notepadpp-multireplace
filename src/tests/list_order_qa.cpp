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

// QA for ListOrder: the third click on a sorted column brings back the order from before sorting, with the entries
// added since where new entries go and those deleted left out; undo and redo of a sort bring back a recorded order.
// Build: g++ -std=c++17 -O1 -g -Wall -Wextra -Wshadow -fsanitize=address,undefined -fno-sanitize-recover=undefined -I.. list_order_qa.cpp ../ListOrder.cpp -o list_order_qa
#include <algorithm>
#include <cstdio>
#include <random>
#include <string>
#include <utility>
#include <vector>
#include "ListOrder.h"

static int g_fail = 0;
static void check(bool ok, const char* what) {
    std::printf("%s  %s\n", ok ? "ok  " : "FAIL", what);
    if (!ok) ++g_fail;
}

using Places = std::vector<std::size_t>;

struct Entry {
    std::size_t id;
    std::string find;
    std::string comment;
};

// A list as the panel keeps it: entries with ids, the sorted column and the order from before sorting
struct List {
    std::vector<Entry> entries;
    int column = -1;                 // the sorted column, -1 for none
    bool descending = false;
    ListOrder::Ids unsorted;         // remembered when sorting starts from the unsorted list
    bool newFirst = false;           // Add new entries at the top
    std::size_t nextId = 1;

    ListOrder::Ids ids() const {
        ListOrder::Ids out;
        for (const Entry& e : entries) out.push_back(e.id);
        return out;
    }
    void insert(std::size_t at, const std::string& find, const std::string& comment = {}) {
        entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(std::min(at, entries.size())), Entry{ nextId++, find, comment });
    }
    // As Add into List
    void add(const std::string& find) { insert(newFirst ? 0 : entries.size(), find); }
    void erase(std::size_t at) { entries.erase(entries.begin() + static_cast<std::ptrdiff_t>(at)); }
    void move(std::size_t from, std::size_t to) {
        Entry e = entries[from];
        erase(from);
        entries.insert(entries.begin() + static_cast<std::ptrdiff_t>(to), e);
    }
    void arrange(const Places& from) {
        std::vector<Entry> next;
        for (std::size_t i : from) next.push_back(entries[i]);
        entries = std::move(next);
    }
    // As ListSort: the order with its sort state, which undo and redo of a sort bring back together
    struct State {
        ListOrder::Ids order;
        int column;
        bool descending;
        ListOrder::Ids unsorted;
    };
    State state() const { return { ids(), column, descending, unsorted }; }
    void apply(const State& s) {
        arrange(ListOrder::arranged(s.order, ids(), false));
        column = s.column;
        descending = s.descending;
        unsorted = s.unsorted;
    }
    // As sortReplaceListData: a header click sorts ascending, then descending, then goes back to unsorted
    void click(int col) {
        if (column < 0) unsorted = ids();
        if (column != col) { column = col; descending = false; }
        else if (!descending) descending = true;
        else column = -1;
        if (column < 0) {
            arrange(ListOrder::arranged(unsorted, ids(), newFirst));
            return;
        }
        std::stable_sort(entries.begin(), entries.end(), [&](const Entry& a, const Entry& b) {
            const std::string& x = col == 0 ? a.find : a.comment;
            const std::string& y = col == 0 ? b.find : b.comment;
            return descending ? x > y : x < y;
        });
    }
    std::string str() const {
        std::string s;
        for (std::size_t i = 0; i < entries.size(); ++i) s += (i ? "|" : "") + entries[i].find;
        return s;
    }
};

static List list(std::initializer_list<const char*> finds, bool newFirst = false) {
    List l;
    l.newFirst = newFirst;
    for (const char* f : finds) l.insert(l.entries.size(), f);
    return l;
}

static bool isPermutation(Places p, std::size_t n) {
    std::sort(p.begin(), p.end());
    for (std::size_t i = 0; i < p.size(); ++i) if (p[i] != i) return false;
    return p.size() == n;
}

int main() {
    {   // Nothing remembered: all entries count as added, in the order they came in
        check(ListOrder::arranged({}, { 5, 3, 9 }, false) == Places({ 1, 0, 2 }), "no remembered order: oldest first");
        check(ListOrder::arranged({}, { 5, 3, 9 }, true) == Places({ 2, 0, 1 }), "no remembered order, new ones first: newest first");
        check(ListOrder::arranged({}, {}, false).empty() && ListOrder::arranged({ 1, 2 }, {}, true).empty(), "empty list, empty result");
        check(ListOrder::arranged({ 4, 2, 8 }, { 4, 2, 8 }, false) == Places({ 0, 1, 2 }), "the remembered order itself: unchanged");
    }
    {   // The header cycle
        List l = list({ "c", "a", "b" });
        l.click(0);
        const bool asc = l.str() == "a|b|c";
        l.click(0);
        const bool desc = l.str() == "c|b|a";
        l.click(0);
        check(asc && desc && l.str() == "c|a|b" && l.column < 0, "three clicks: ascending, descending, the order from before");
    }
    {   // Entries added while sorted go where new entries go
        List l = list({ "c", "a", "b" });
        l.click(0); l.add("d"); l.add("e"); l.click(0); l.click(0);
        check(l.str() == "c|a|b|d|e", "added while sorted: after the others, in the order they were added");
        List t = list({ "c", "a", "b" }, true);
        t.click(0); t.add("d"); t.add("e"); t.click(0); t.click(0);
        check(t.str() == "e|d|c|a|b", "with new entries at the top: before the others, newest first");
        List p = list({ "c", "a", "b" });
        p.click(0); p.insert(1, "y"); p.insert(0, "x"); p.click(0); p.click(0);
        check(p.str() == "c|a|b|y|x", "pasted while sorted: in the order they came in, not where they were pasted");
    }
    {   // Deleted entries drop out
        List l = list({ "c", "a", "b", "d" });
        l.click(0); l.erase(1); l.erase(0); l.click(0); l.click(0);
        check(l.str() == "c|d", "deleted while sorted: gone from the restored order");
        List all = list({ "b", "a" });
        all.click(0); all.erase(0); all.erase(0); all.add("z"); all.click(0); all.click(0);
        check(all.str() == "z", "all deleted, one added: just that one");
    }
    {   // The order from right before sorting counts, not the one from the first sort ever
        List l = list({ "c", "a", "b" });
        l.click(0); l.click(0); l.click(0);
        l.move(2, 0);
        l.click(0); l.click(0); l.click(0);
        check(l.str() == "b|c|a", "a move between two sorts is kept");
        List s = list({ "c", "a", "b" });
        s.click(0); s.move(2, 0);
        s.click(0); s.click(0);
        check(s.str() == "c|a|b", "a move while sorted does not change the order to go back to");
    }
    {   // Another column while sorted: still the order from before the first of them
        List l = list({ "c", "a", "b" });
        l.entries[0].comment = "2"; l.entries[1].comment = "3"; l.entries[2].comment = "1";
        l.click(0);
        l.click(1);
        const bool byComment = l.str() == "b|c|a";
        l.click(1); l.click(1);
        check(byComment && l.str() == "c|a|b", "sorting another column keeps the order from before sorting");
    }
    {   // Equal keys keep their order, so the round trip is exact
        List l = list({ "b", "a", "b", "a" });
        const ListOrder::Ids before = l.ids();
        l.click(0);
        const ListOrder::Ids asc = l.ids();
        check(asc == ListOrder::Ids({ 2, 4, 1, 3 }), "equal entries keep their order when sorted");
        l.click(0); l.click(0);
        check(l.ids() == before, "the round trip brings back each entry to its place");
    }
    {   // Undo and redo of a sort: a recorded order comes back exactly
        List l = list({ "d", "b", "a", "c" });
        const ListOrder::Ids before = l.ids();
        l.click(0);
        l.arrange(ListOrder::arranged(before, l.ids(), false));
        check(l.str() == "d|b|a|c", "a recorded order comes back exactly");
        check(ListOrder::arranged({ 1, 2, 3, 4 }, { 4, 2 }, false) == Places({ 1, 0 }), "recorded entries the list lacks drop out");
        check(ListOrder::arranged({ 2, 1 }, { 1, 7, 2 }, false) == Places({ 2, 0, 1 }), "an entry the record lacks stays at the end");
        check(ListOrder::arranged({ 2, 1 }, { 1, 7, 2 }, true) == Places({ 1, 2, 0 }), "or at the top");
    }
    {   // Undo back past a sort brings back the order to go back to that held then
        struct Step { bool sort; List::State before, after; std::size_t from, to; };
        std::vector<Step> undo;
        List l = list({ "c", "a", "b" });
        auto click = [&](int col) { const List::State before = l.state(); l.click(col); undo.push_back({ true, before, l.state(), 0, 0 }); };
        auto move = [&](std::size_t from, std::size_t to) { l.move(from, to); undo.push_back({ false, {}, {}, from, to }); };
        auto back = [&] {
            const Step s = undo.back();
            undo.pop_back();
            if (s.sort) l.apply(s.before);
            else l.move(s.to, s.from);
        };
        click(0); click(0); click(0);
        move(2, 0);
        click(0);
        back();
        const bool sortUndone = l.str() == "b|c|a" && l.column < 0;
        back();
        back();
        const bool descAgain = l.str() == "c|b|a" && l.column == 0 && l.descending;
        click(0);
        check(sortUndone && descAgain && l.str() == "c|a|b", "undo past a sort: the third click restores the order from before that sort");
    }
    {   // Ids that should not repeat still give every entry one place
        const Places p = ListOrder::arranged({ 1, 2 }, { 2, 2, 1 }, false);
        check(p == Places({ 2, 0, 1 }), "a repeated id: both copies in their current order");
        check(ListOrder::arranged({ 3, 1, 3 }, { 1, 3 }, false) == Places({ 1, 0 }), "a repeated id in the record keeps its first place");
        ListOrder::Ids many;
        for (int i = 0; i < 60; ++i) many.push_back(i % 3 ? 4 : 9);
        Places expected;
        for (std::size_t i = 0; i < many.size(); ++i) if (many[i] == 9) expected.push_back(i);
        for (std::size_t i = 0; i < many.size(); ++i) if (many[i] == 4) expected.push_back(i);
        check(ListOrder::arranged({ 9, 4 }, many, false) == expected, "many copies of an id keep their order");
    }
    {   // Random lists and edits: the remembered order, the others together where new entries go, oldest or newest first
        std::mt19937 rng(20261119);
        int checked = 0, bad = 0;
        for (int round = 0; round < 40000; ++round) {
            List l;
            l.newFirst = rng() % 2;
            const int n = static_cast<int>(rng() % 9);
            for (int i = 0; i < n; ++i) {
                l.insert(l.entries.size(), std::string(1, static_cast<char>('a' + rng() % 4)), std::string(1, static_cast<char>('a' + rng() % 3)));
            }
            for (int s = static_cast<int>(rng() % 10); s > 0; --s) {
                const int col = static_cast<int>(rng() % 2);
                switch (rng() % 6) {
                case 0: l.add("n"); break;
                case 1: if (!l.entries.empty()) l.erase(rng() % l.entries.size()); break;
                case 2: if (l.entries.size() > 1) l.move(rng() % l.entries.size(), rng() % l.entries.size()); break;
                case 3: l.insert(rng() % (l.entries.size() + 1), "p"); break;
                default: {
                    const bool back = l.column == col && l.descending;
                    const ListOrder::Ids remembered = l.unsorted;
                    const ListOrder::Ids current = l.ids();
                    if (back) {
                        ++checked;
                        const Places p = ListOrder::arranged(remembered, current, l.newFirst);
                        bool ok = isPermutation(p, current.size());
                        std::vector<std::size_t> rank(current.size(), SIZE_MAX);
                        for (std::size_t i = 0; i < current.size(); ++i) {
                            const auto it = std::find(remembered.begin(), remembered.end(), current[i]);
                            if (it != remembered.end()) rank[i] = static_cast<std::size_t>(it - remembered.begin());
                        }
                        // Walk the result: the known ones rise in rank, the others in id (falling with newFirst), no mixing
                        bool seenKnown = false, seenOther = false;
                        std::size_t lastRank = 0, lastId = 0;
                        for (std::size_t k = 0; ok && k < p.size(); ++k) {
                            const std::size_t i = p[k];
                            if (rank[i] != SIZE_MAX) {
                                if ((seenKnown && rank[i] < lastRank) || (seenOther && !l.newFirst)) ok = false;
                                lastRank = rank[i];
                                seenKnown = true;
                            }
                            else {
                                const bool wrongWay = l.newFirst ? current[i] > lastId : current[i] < lastId;
                                if ((seenOther && wrongWay) || (seenKnown && l.newFirst)) ok = false;
                                lastId = current[i];
                                seenOther = true;
                            }
                        }
                        if (!ok) ++bad;
                    }
                    l.click(col);
                }
                }
            }
        }
        std::printf("      %d restores checked\n", checked);
        check(checked > 1000 && bad == 0, "random: a permutation, remembered order kept, the others together where new entries go");
    }
    {   // Random header clicks without edits: every way back to unsorted gives the start order
        std::mt19937 rng(7);
        int trips = 0, bad = 0;
        for (int round = 0; round < 20000; ++round) {
            List l;
            const int n = static_cast<int>(rng() % 9);
            for (int i = 0; i < n; ++i) {
                l.insert(l.entries.size(), std::string(1, static_cast<char>('a' + rng() % 3)), std::string(1, static_cast<char>('a' + rng() % 3)));
            }
            const ListOrder::Ids start = l.ids();
            for (int s = static_cast<int>(rng() % 12); s > 0; --s) {
                l.click(static_cast<int>(rng() % 2));
                if (l.column < 0) {
                    ++trips;
                    if (l.ids() != start) ++bad;
                }
            }
        }
        std::printf("      %d ways back to unsorted\n", trips);
        check(trips > 1000 && bad == 0, "random: sorting and back without edits gives the start order");
    }

    std::printf(g_fail ? "%d FAILED\n" : "all passed\n", g_fail);
    return g_fail ? 1 : 0;
}
