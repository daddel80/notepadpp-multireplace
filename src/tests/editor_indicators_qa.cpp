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

// QA for EditorIndicators: a kind gets its indicator from Notepad++ on first use and asks only once, also when none
// was left; a Notepad++ without allocation leaves MR to pick from the plugins' range, in the order of first use.
// Build: g++ -std=c++17 -O1 -g -Wall -Wextra -Wshadow -fsanitize=address,undefined -fno-sanitize-recover=undefined -I.. editor_indicators_qa.cpp ../EditorIndicators.cpp -o editor_indicators_qa
#include <cstdio>
#include <vector>
#include "EditorIndicators.h"

using EditorIndicators::Kind;
using EditorIndicators::Pool;

static int g_fail = 0;
static void check(bool ok, const char* what) {
    std::printf("%s  %s\n", ok ? "ok  " : "FAIL", what);
    if (!ok) ++g_fail;
}

// Notepad++ as NPPM_ALLOCATEINDICATOR answers: the next id of its range, NONE_LEFT once it is used up
struct Npp {
    int next = EditorIndicators::FIRST_PLUGIN;
    int asked = 0;
    int answer() {
        ++asked;
        return next <= EditorIndicators::LAST_PLUGIN ? next++ : EditorIndicators::NONE_LEFT;
    }
};

int main()
{
    {
        Npp npp;
        npp.next = 11;   // Columns++ and Search++ were first
        Pool pool;
        pool.setAllocate([&] { return npp.answer(); });
        check(pool.peek(Kind::ReplaceMarks) == -1 && npp.asked == 0, "peek before the first use: -1, nothing asked");
        check(pool.id(Kind::ReplaceMarks) == 11, "first use asks Notepad++: the next free id");
        check(pool.id(Kind::ReplaceDeletions) == 12, "the next kind gets the next id");
        check(pool.id(Kind::ReplaceMarks) == 11 && npp.asked == 2, "a kind keeps its id and asks only once");
        check(pool.peek(Kind::ReplaceDeletions) == 12, "peek after the first use: the id");
        check(pool.id(Kind::MarkMatches) == 13 && pool.id(Kind::ScopeFrame) == 14, "ids follow the order of first use");
        check(pool.peek(Kind::FlowTabs) == -1 && pool.peek(Kind::Duplicates) == -1, "kinds never used have none");
    }
    {
        Npp npp;
        npp.next = 20;   // one left
        Pool pool;
        pool.setAllocate([&] { return npp.answer(); });
        check(pool.id(Kind::MarkMatches) == 20, "the last free id");
        check(pool.id(Kind::Duplicates) == -1, "none left: -1");
        check(pool.id(Kind::Duplicates) == -1 && npp.asked == 2, "none left stays -1 without asking again");
        check(pool.peek(Kind::Duplicates) == -1, "peek of a kind left without: -1");
        check(pool.id(Kind::MarkMatches) == 20, "the kind that got one keeps it");
    }
    {
        int asked = 0;
        Pool pool;
        pool.setAllocate([&] { ++asked; return EditorIndicators::NOT_HANDED_OUT; });
        check(pool.id(Kind::FlowTabs) == EditorIndicators::FIRST_PLUGIN, "Notepad++ before 8.5.6: MR picks the first plugin id");
        check(pool.id(Kind::ScopeFrame) == EditorIndicators::FIRST_PLUGIN + 1, "picks follow the order of first use");
        check(asked == 1, "a Notepad++ that hands out none is asked once");
        std::vector<int> all;
        for (int k = 0; k < static_cast<int>(Kind::Count); ++k) all.push_back(pool.id(static_cast<Kind>(k)));
        bool inRange = true;
        for (std::size_t i = 0; i < all.size(); ++i) {
            if (all[i] < EditorIndicators::FIRST_PLUGIN || all[i] > EditorIndicators::LAST_PLUGIN) inRange = false;
            for (std::size_t j = 0; j < i; ++j) if (all[i] == all[j]) inRange = false;
        }
        check(inRange, "all kinds picked: distinct ids of the plugins' range");
    }
    {
        Pool pool;   // no allocator set
        check(pool.id(Kind::MarkMatches) == EditorIndicators::FIRST_PLUGIN, "without an allocator MR picks");
    }
    {
        Npp npp;
        Pool pool;
        pool.setAllocate([&] { return npp.answer(); });
        int ids[static_cast<int>(Kind::Count)];
        for (int k = 0; k < static_cast<int>(Kind::Count); ++k) ids[k] = pool.id(static_cast<Kind>(k));
        bool distinct = true;
        for (int i = 0; i < static_cast<int>(Kind::Count); ++i)
            for (int j = 0; j < i; ++j) if (ids[i] == ids[j] || ids[i] < 0) distinct = false;
        check(distinct && npp.asked == static_cast<int>(Kind::Count), "every kind its own id, one request each");
    }
    check(&EditorIndicators::shared() == &EditorIndicators::shared(), "one shared pool");

    if (g_fail) { std::printf("%d check(s) FAILED\n", g_fail); return 1; }
    std::printf("all checks passed\n");
    return 0;
}
