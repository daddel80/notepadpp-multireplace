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

// QA for TextReplaceAnimation: steps run like the entries of a list, each on the text the step before left, from its
// first match on. At every tick the plain, marked and plain parts match a model of the steps that apply; each step
// reports its result once and holds it for its time; steps without a match are left out.
// Build: g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. text_replace_animation_qa.cpp -o text_replace_animation_qa
#include <algorithm>
#include <cstddef>
#include <cstdio>
#include <string>
#include <vector>

#include "../TextReplaceAnimation.h"

namespace {

    using Step = TextReplaceAnimation::Step;
    using Tick = TextReplaceAnimation::Tick;
    using Phase = TextReplaceAnimation::Phase;

    int g_checks = 0;
    int g_fail = 0;

    void check(bool ok, const char* what)
    {
        ++g_checks;
        if (!ok) {
            ++g_fail;
            std::printf("FAIL  %s\n", what);
        }
    }

    struct Seen {
        std::vector<std::wstring> statuses;   // in the order the steps reported their result
        std::vector<int> holds;               // ticks each result stood
    };

    // Runs a start to its end and compares every tick with a model of the steps that apply
    bool run(const std::wstring& text, const std::vector<Step>& steps, const std::wstring& result, Seen* seen = nullptr)
    {
        struct Applied { std::wstring source; std::wstring next; std::size_t at; Step step; };
        std::vector<Applied> model;
        std::wstring current = text;
        for (const Step& step : steps) {
            const std::size_t at = step.find.empty() ? std::wstring::npos : current.find(step.find);
            if (at == std::wstring::npos) continue;
            const std::wstring next = current.substr(0, at) + step.replace + current.substr(at + step.find.size());
            model.push_back({ current, next, at, step });
            current = next;
        }

        TextReplaceAnimation animation;
        animation.start(text, steps);
        bool ok = current == result && animation.result() == result;
        if (model.empty()) return ok && !animation.isActive() && animation.tick() == Tick::Over;

        for (std::size_t i = 0; i < model.size(); ++i) {
            const Applied& m = model[i];
            const std::wstring tail = m.source.substr(m.at + m.step.find.size());
            ok = ok && animation.phase() == Phase::Find;

            int ticks = 0;
            while (animation.phase() == Phase::Find && ++ticks < 1000) {
                ok = ok && animation.before() == m.source.substr(0, m.at) && animation.marked() == m.step.find
                    && animation.after() == tail;
                ok = ok && animation.tick() == Tick::Running;
            }

            std::size_t chars = 0;
            Tick last = Tick::Running;
            while (animation.phase() == Phase::Replace && chars < 1000) {
                ok = ok && animation.before() == m.next.substr(0, m.at + (std::min)(chars, m.step.replace.size()))
                    && animation.marked() == m.step.find.substr((std::min)(chars, m.step.find.size()))
                    && animation.after() == tail;
                last = animation.tick();
                ++chars;
                ok = ok && (last == Tick::Running) == (animation.phase() == Phase::Replace);
            }
            ok = ok && last == Tick::StepDone && chars == (std::max)(m.step.find.size(), m.step.replace.size());
            if (seen) seen->statuses.push_back(animation.step().status);

            int hold = 0;
            while (animation.phase() == Phase::Status && hold < 1000) {
                ok = ok && animation.before() == m.next && animation.marked().empty() && animation.after().empty();
                last = animation.tick();
                ++hold;
            }
            if (seen) seen->holds.push_back(hold);
            ok = ok && hold == (std::max)(1, m.step.holdTicks);
            const bool lastStep = i + 1 == model.size();
            ok = ok && last == (lastStep ? Tick::Over : Tick::Running) && animation.isActive() != lastStep;
        }
        return ok && animation.tick() == Tick::Over && animation.before().empty() && animation.marked().empty()
            && animation.after().empty() && animation.result() == result;
    }
}

int main()
{
    // Two steps like two list entries: the first replaces a word and leaves the rest, the second the whole result
    {
        const std::vector<Step> chain = { { L"alpha", L"epsilon", L"one", 34 },
                                          { L"epsilon beta", L"delta zeta and eta", L"two" } };
        Seen seen;
        check(run(L"alpha beta", chain, L"delta zeta and eta", &seen), "chain: every tick as the model");
        check(seen.statuses == std::vector<std::wstring>{ L"one", L"two" }, "chain: each status once, in order");
        check(seen.holds == std::vector<int>{ 34, TextReplaceAnimation::HOLD_TICKS }, "chain: each result holds its time");

        TextReplaceAnimation animation;
        animation.start(L"alpha beta", chain);
        check(animation.before().empty() && animation.marked() == L"alpha" && animation.after() == L" beta",
            "chain: only the first word is marked, the rest stands plain");
        check(run(L"delta zeta and eta", { { L"delta zeta and eta", L"alpha beta", L"back" } }, L"alpha beta"),
            "chain: back in one step");
    }

    // Where the match is and how long the replacement is
    check(run(L"a alpha b", { { L"alpha", L"epsilon", L"s" } }, L"a epsilon b"), "match in the middle: text on both sides stays");
    check(run(L"alpha beta", { { L"alpha", L"al", L"s" } }, L"al beta"), "replacement shorter than the match");
    check(run(L"al", { { L"al", L"epsilon beta", L"s" } }, L"epsilon beta"), "replacement longer than the match");
    check(run(L"alpha alpha", { { L"alpha", L"x", L"s" } }, L"x alpha"), "only the first match");
    check(run(L"alpha b\u00E9ta", { { L"b\u00E9ta", L"beta", L"s" } }, L"alpha beta"), "characters beyond ASCII");
    check(run(L"alpha", { { L"alpha", L"", L"s" } }, L""), "replaced by nothing");

    // Steps that do not apply
    {
        Seen seen;
        check(run(L"alpha beta", { { L"omega", L"x", L"a" }, { L"", L"y", L"b" }, { L"alpha", L"epsilon", L"c" } },
            L"epsilon beta", &seen), "steps without a match and with an empty find text are left out");
        check(seen.statuses == std::vector<std::wstring>{ L"c" }, "only the step that applies reports");
        check(run(L"alpha beta", { { L"alpha", L"x", L"a" }, { L"alpha beta", L"y", L"b" } }, L"x beta"),
            "a step sees the text the step before left");
        check(run(L"alpha", { { L"x", L"y", L"s" } }, L"alpha"), "nothing applies: no run, the text stays");
        check(run(L"alpha", {}, L"alpha"), "no steps: no run");
        check(run(L"alpha", { { L"alpha", L"x", L"s", 0 } }, L"x"), "no hold time: the result stands one tick");
    }

    // Stopping and starting again
    {
        TextReplaceAnimation animation;
        check(animation.result().empty() && !animation.isActive() && animation.tick() == Tick::Over, "before a start: idle");
        animation.start(L"alpha beta", { { L"alpha", L"epsilon", L"one" } });
        for (int i = 0; i < 20; ++i) animation.tick();
        animation.stop();
        check(!animation.isActive() && animation.tick() == Tick::Over, "stop: the run is over");
        check(animation.before().empty() && animation.marked().empty() && animation.after().empty(), "stop: nothing drawn from the run");
        check(animation.result() == L"epsilon beta", "stop: the result stays known");
        animation.start(L"epsilon beta", { { L"epsilon beta", L"alpha beta", L"back" } });
        check(animation.isActive() && animation.marked() == L"epsilon beta" && animation.result() == L"alpha beta",
            "a new start begins from its own text");
    }

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
