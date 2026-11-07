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

// QA harness for TabStripLayout, the arrangement of the list tab strip. On fixed and random strips:
//   A. a strip that fits: no "...", nothing clipped, "+" right after the last tab
//   B. every scroll position: first tab at the start inset, the view stays out of the "..." slot,
//      "..." exactly where tabs are cut (a cut end margin alone shows none), "+" right after the
//      last tab once it is in view, the control clipped exactly when the view shows part of it
//   C. the end of the scroll range: every following tab in view, one step back hides the last tab
//   D. scrolling moves one tab per step and stops at both ends
//   E. revealing shows the tab completely with the least scrolling, never past the tab
//   F. a growing strip with the new tab revealed keeps its row full (no shrinking view)
// Build (from src/tests):
//   g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. -o tab_strip_layout_qa tab_strip_layout_qa.cpp ../TabStripLayout.cpp
//   cl /std:c++20 /EHsc /I.. tab_strip_layout_qa.cpp ..\TabStripLayout.cpp /Fe:tab_strip_layout_qa.exe
// Usage: tab_strip_layout_qa [seed] [cases]
#include "../TabStripLayout.h"

#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <random>
#include <string>
#include <vector>

using namespace TabStripLayout;

namespace {

    int g_checks = 0;
    int g_failures = 0;

    void check(bool ok, const char* what, const std::string& context)
    {
        ++g_checks;
        if (ok) return;
        if (++g_failures <= 25) std::printf("FAIL %s | %s\n", what, context.c_str());
    }

    // Tabs of the given widths, the first one inset by `lead` like the Windows tab control places it
    std::vector<Edge> strip(const std::vector<int>& widths, int lead)
    {
        std::vector<Edge> tabs;
        int x = lead;
        for (int w : widths) {
            tabs.push_back({ x, x + w });
            x += w;
        }
        return tabs;
    }

    // Row metrics as the panel derives them: "+" after the "..." slot that follows the strip
    Metrics row(int left, int maxWidth, int indicator, int gap)
    {
        return { left, maxWidth, indicator, gap, gap + indicator, left + maxWidth + gap + indicator };
    }

    std::string describe(const std::vector<Edge>& tabs, const Metrics& m, int first)
    {
        std::string s = "left=" + std::to_string(m.left) + " max=" + std::to_string(m.maxWidth)
            + " ind=" + std::to_string(m.indicator) + " margin=" + std::to_string(m.endMargin)
            + " first=" + std::to_string(first) + " tabs=";
        for (const Edge& e : tabs) s += std::to_string(e.left) + "-" + std::to_string(e.right) + " ";
        return s;
    }

    int roomFor(const Metrics& m, int first) { return m.maxWidth - (first > 0 ? m.indicator : 0); }

    // Brute-force end of the scroll range: from there on the last tab fits
    int lastFirstOf(const std::vector<Edge>& tabs, const Metrics& m)
    {
        const int count = static_cast<int>(tabs.size());
        for (int k = 0; k < count; ++k) {
            const int need = tabs.back().right - (tabs[k].left - tabs.front().left);
            if (need <= roomFor(m, k)) return k;
        }
        return count - 1;
    }

    // Tab `index` lies completely in the view
    bool showsTab(const Layout& lay, const std::vector<Edge>& tabs, int index)
    {
        return lay.controlLeft + tabs[index].left >= lay.viewLeft
            && lay.controlLeft + tabs[index].right <= lay.viewLeft + lay.viewWidth;
    }

    // ... and the end margin after it, what revealing aims for
    bool showsWithMargin(const Layout& lay, const std::vector<Edge>& tabs, const Metrics& m, int index)
    {
        return showsTab(lay, tabs, index) && lay.controlLeft + tabs[index].right + m.endMargin <= lay.viewLeft + lay.viewWidth;
    }

    void checkLayout(const std::vector<Edge>& tabs, const Metrics& m, int first)
    {
        const std::string ctx = describe(tabs, m, first);
        const Layout lay = arrange(tabs, m, first);
        if (tabs.empty()) {
            check(!lay.overflow && !lay.hiddenLeft && !lay.hiddenRight && lay.controlWidth == 0, "empty strip", ctx);
            return;
        }

        const int last = lastFirstOf(tabs, m);
        const int lead = tabs.front().left;
        const int lastRight = tabs.back().right;
        const int viewRight = lay.viewLeft + lay.viewWidth;

        check(lay.first == std::clamp(first, 0, last), "first clamped to the scroll range", ctx);
        check(lay.overflow == (lastRight > m.maxWidth), "overflow", ctx);
        check(lay.controlWidth == lastRight + m.endMargin, "control holds every tab", ctx);
        check(lay.hiddenLeft == (lay.first > 0), "left ... iff scrolled", ctx);
        check(lay.viewLeft == m.left + (lay.hiddenLeft ? m.indicator : 0), "view starts after the left ...", ctx);
        check(lay.controlLeft + tabs[lay.first].left == lay.viewLeft + lead, "first tab at the start inset", ctx);
        check(lay.clipLeft == lay.viewLeft - lay.controlLeft, "clip matches the control position", ctx);
        check(viewRight <= m.left + m.maxWidth, "view stays out of the ... slot", ctx);
        check(lay.hiddenRight == (lay.controlLeft + lastRight > viewRight), "right ... iff the last tab is cut", ctx);
        check(lay.clipped == (lay.clipLeft > 0 || lay.viewWidth < lay.controlWidth), "clipped iff part of the control shows", ctx);
        check(lay.plusLeft <= m.plusMax, "+ within its anchor", ctx);

        if (!lay.overflow) {
            check(lay.first == 0 && lay.clipLeft == 0 && lay.controlLeft == m.left
                && lay.viewWidth == std::min(lay.controlWidth, m.maxWidth), "fitting strip from the start", ctx);
        }
        if (lay.hiddenRight) {
            check(lay.viewWidth == roomFor(m, lay.first), "cut strip fills its room", ctx);
            check(lay.plusLeft == m.plusMax, "+ at its anchor while tabs hide right", ctx);
            check(lay.plusLeft - m.indicator >= viewRight, "right ... after the strip", ctx);
        }
        else {
            check(viewRight == std::min(lay.controlLeft + lastRight + m.endMargin, m.left + m.maxWidth), "view hugs the last tab", ctx);
            check(lay.plusLeft == std::min(lay.controlLeft + lastRight + m.plusGap, m.plusMax), "+ after the last tab", ctx);
            for (int i = lay.first; i < static_cast<int>(tabs.size()); ++i)
                check(showsTab(lay, tabs, i), "every tab from the first in view", ctx);
        }
        if (lay.overflow && lay.first == last && last > 0 && !lay.hiddenRight) {
            check(arrange(tabs, m, last - 1).hiddenRight, "one step back hides the last tab", ctx);
        }
    }

    void checkScrolling(const std::vector<Edge>& tabs, const Metrics& m)
    {
        const std::string ctx = describe(tabs, m, 0);
        if (tabs.empty()) {
            check(scrolled(tabs, m, 3, 1) == 0 && revealing(tabs, m, 3, 1) == 0, "empty strip scrolls nowhere", ctx);
            return;
        }
        const int last = lastFirstOf(tabs, m);
        for (int f = -2; f <= last + 2; ++f) {
            const int fc = std::clamp(f, 0, last);
            check(scrolled(tabs, m, f, 1) == std::min(fc + 1, last), "one step right", ctx);
            check(scrolled(tabs, m, f, -1) == std::max(fc - 1, 0), "one step left", ctx);
        }

        // Walking right from the start visits every position once, walking back returns
        int first = 0, steps = 0;
        while (scrolled(tabs, m, first, 1) != first) { first = scrolled(tabs, m, first, 1); ++steps; }
        check(first == last && steps == last, "walk right ends at the scroll end", ctx);
        while (scrolled(tabs, m, first, -1) != first) first = scrolled(tabs, m, first, -1);
        check(first == 0, "walk left ends at the start", ctx);
    }

    void checkRevealing(const std::vector<Edge>& tabs, const Metrics& m, int first, int index)
    {
        const std::string ctx = describe(tabs, m, first) + " index=" + std::to_string(index);
        const int last = lastFirstOf(tabs, m);
        const int fc = std::clamp(first, 0, last);
        const int r = revealing(tabs, m, first, index);
        check(r >= 0 && r <= last, "revealed position in range", ctx);

        const Layout lay = arrange(tabs, m, r);
        const bool fits = tabs[index].right - tabs[index].left + tabs.front().left <= roomFor(m, index);
        if (fits || index > last) check(showsTab(lay, tabs, index), "revealed tab in view", ctx);
        else check(r == index, "oversized tab starts the view", ctx);

        if (index < fc) {
            check(r == index, "tab left of the view becomes the first", ctx);
            return;
        }
        check(r >= fc && r <= std::min(index, last), "scrolls right, never past the tab", ctx);
        check(showsWithMargin(lay, tabs, m, index) || r == std::min(index, last), "margin in view unless at the bound", ctx);
        for (int k = fc; k < r; ++k)
            check(!showsWithMargin(arrange(tabs, m, k), tabs, m, index), "least scrolling", ctx);
    }

    void checkStrip(const std::vector<Edge>& tabs, const Metrics& m, std::mt19937& rng)
    {
        const int count = static_cast<int>(tabs.size());
        for (int f = -1; f <= count + 1; ++f) checkLayout(tabs, m, f);
        checkScrolling(tabs, m);
        if (count == 0) return;
        std::uniform_int_distribution<int> anyFirst(-1, count + 1);
        for (int index = 0; index < count; ++index) checkRevealing(tabs, m, anyFirst(rng), index);
    }

    // F: tabs added one by one, each new tab revealed as addNewTab does
    void checkGrowth(const std::vector<int>& widths, int lead, const Metrics& m)
    {
        std::vector<int> grown;
        int first = 0;
        for (int w : widths) {
            grown.push_back(w);
            const std::vector<Edge> tabs = strip(grown, lead);
            first = revealing(tabs, m, first, static_cast<int>(tabs.size()) - 1);
            const Layout lay = arrange(tabs, m, first);
            const std::string ctx = describe(tabs, m, first);
            check(!lay.hiddenRight || lay.first == static_cast<int>(tabs.size()) - 1, "new tab in view", ctx);
            check(lay.first == lastFirstOf(tabs, m), "full row after adding a tab", ctx);
        }
    }
}

int main(int argc, char** argv)
{
    const unsigned seed = argc > 1 ? static_cast<unsigned>(std::strtoul(argv[1], nullptr, 10)) : 1u;
    const int cases = argc > 2 ? std::atoi(argv[2]) : 20000;
    std::mt19937 rng(seed);

    // Fixed strips at 100 % and 200 %: 15 tabs of about 75 px in a 460 px panel row
    for (int s : { 1, 2 }) {
        const Metrics m = row(12 * s, 460 * s - 12 * s - 40 * s - 2 * (22 * s + 2 * s) - (20 * s + 2 * s), 20 * s, 2 * s);
        std::vector<int> widths;
        for (int i = 1; i <= 15; ++i) widths.push_back((i < 10 ? 70 : 78) * s);
        for (int n = 0; n <= 15; ++n) {
            const std::vector<int> part(widths.begin(), widths.begin() + n);
            checkStrip(strip(part, 2 * s), m, rng);
        }
        checkGrowth(widths, 2 * s, m);
    }

    // A fitting strip: no indicators, "+" glued
    {
        const Metrics m = row(12, 400, 20, 2);
        const std::vector<Edge> tabs = strip({ 60, 80, 70 }, 2);
        const Layout lay = arrange(tabs, m, 0);
        check(!lay.overflow && !lay.hiddenLeft && !lay.hiddenRight && lay.plusLeft == 12 + 212 + 22,
            "fitting strip with + after it", describe(tabs, m, 0));
    }

    // The last tab whole and only its end margin cut: nothing hidden on the right
    {
        const Metrics m = row(12, 300, 21, 3);
        const std::vector<Edge> tabs = strip({ 50, 120, 158 }, 0);
        const Layout lay = arrange(tabs, m, 1);
        check(lay.first == 1 && !lay.hiddenRight && lay.clipped && lay.viewWidth == 279,
            "cut end margin shows no right ...", describe(tabs, m, 1));
    }

    std::uniform_int_distribution<int> countDist(0, 40), widthDist(40, 180), leadDist(0, 4);
    std::uniform_int_distribution<int> leftDist(0, 40), maxDist(60, 1600), indicatorDist(10, 44), gapDist(1, 6);
    for (int c = 0; c < cases; ++c) {
        std::vector<int> widths(countDist(rng));
        for (int& w : widths) w = widthDist(rng);
        const int lead = leadDist(rng);
        const Metrics m = row(leftDist(rng), maxDist(rng), indicatorDist(rng), gapDist(rng));
        checkStrip(strip(widths, lead), m, rng);
        if (c % 16 == 0) checkGrowth(widths, lead, m);
    }

    std::printf("tab_strip_layout_qa: %d checks, %d failures (seed %u, %d cases)\n", g_checks, g_failures, seed, cases);
    return g_failures == 0 ? 0 : 1;
}
