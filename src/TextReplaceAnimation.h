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

#include <algorithm>
#include <cstddef>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

// A text replaced like a list run, one entry after the other: each step marks the first match of its find text
// (Find), the replacement takes over from the left while the rest of the match stays marked (Replace), then the
// result stands while the step's status shows (Status). Driven by a fixed tick, free of window code.
class TextReplaceAnimation {
public:
    enum class Phase { Idle, Find, Replace, Status };
    // What a tick did: the run goes on, a step's result stands from now on, or the run is over
    enum class Tick { Running, StepDone, Over };

    static constexpr unsigned int TICK_MS = 35;
    static constexpr int HOLD_TICKS = 57;   // about two seconds

    struct Step {
        std::wstring find;
        std::wstring replace;
        std::wstring status;   // shown while the result stands
        int holdTicks = HOLD_TICKS;
    };

    // Each step works on the text the step before left; a step without a match there is left out
    void start(std::wstring text, std::vector<Step> steps) {
        _texts.assign(1, std::move(text));
        _steps.clear();
        _at.clear();
        for (Step& step : steps) {
            const std::wstring& current = _texts.back();
            const std::size_t at = step.find.empty() ? std::wstring::npos : current.find(step.find);
            if (at == std::wstring::npos) continue;
            std::wstring next = current.substr(0, at) + step.replace + current.substr(at + step.find.size());
            _texts.push_back(std::move(next));
            _at.push_back(at);
            _steps.push_back(std::move(step));
        }
        _index = 0;
        _chars = 0;
        _ticks = 0;
        _phase = _steps.empty() ? Phase::Idle : Phase::Find;
    }

    Tick tick() {
        switch (_phase) {
        case Phase::Find:
            if (++_ticks == FIND_TICKS) {
                _phase = Phase::Replace;
                _chars = 0;
            }
            return Tick::Running;
        case Phase::Replace:
            if (++_chars < (std::max)(step().find.size(), step().replace.size())) return Tick::Running;
            _phase = Phase::Status;
            _ticks = 0;
            return Tick::StepDone;
        case Phase::Status:
            if (++_ticks < step().holdTicks) return Tick::Running;
            if (_index + 1 == _steps.size()) {
                _phase = Phase::Idle;
                return Tick::Over;
            }
            ++_index;
            _phase = Phase::Find;
            _ticks = 0;
            return Tick::Running;
        default:
            return Tick::Over;
        }
    }

    void stop() { _phase = Phase::Idle; }

    Phase phase() const { return _phase; }
    bool isActive() const { return _phase != Phase::Idle; }
    // The step running or last run; only after a start with a step that applies
    const Step& step() const { return _steps[_index]; }
    // The text once all steps ran
    const std::wstring& result() const { return _texts.back(); }

    // The text as it stands, in three parts: plain before the mark, the marked rest of the match, plain after it
    std::wstring_view before() const {
        switch (_phase) {
        case Phase::Find: return std::wstring_view(_texts[_index]).substr(0, _at[_index]);
        case Phase::Replace: return std::wstring_view(_texts[_index + 1]).substr(0, _at[_index] + (std::min)(_chars, step().replace.size()));
        case Phase::Status: return _texts[_index + 1];
        default: return {};
        }
    }
    std::wstring_view marked() const {
        switch (_phase) {
        case Phase::Find: return step().find;
        case Phase::Replace: return std::wstring_view(step().find).substr((std::min)(_chars, step().find.size()));
        default: return {};
        }
    }
    std::wstring_view after() const {
        if (_phase != Phase::Find && _phase != Phase::Replace) return {};
        return std::wstring_view(_texts[_index]).substr(_at[_index] + step().find.size());
    }

private:
    static constexpr int FIND_TICKS = 16;   // about half a second

    Phase _phase = Phase::Idle;
    std::vector<Step> _steps;
    std::vector<std::wstring> _texts{ std::wstring() };   // the text before each step, then the result
    std::vector<std::size_t> _at;                          // where each step's match starts
    std::size_t _index = 0;
    std::size_t _chars = 0;
    int _ticks = 0;
};
