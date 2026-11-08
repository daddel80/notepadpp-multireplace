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
#include <string>
#include <string_view>
#include <utility>

// A text replaced like a hit: the whole source is marked first (Find), then the target takes over
// from the left while the rest of the source stays marked (Replace), then the result stands while
// a status shows (Status). Driven by a fixed tick, free of window code.
class TextReplaceAnimation {
public:
    enum class Phase { Idle, Find, Replace, Status };

    static constexpr unsigned int TICK_MS = 35;

    void start(std::wstring from, std::wstring to) {
        _from = std::move(from);
        _to = std::move(to);
        _phase = Phase::Find;
        _ticks = 0;
        _step = 0;
    }

    // Advances one tick; false once the animation is over
    bool tick() {
        switch (_phase) {
        case Phase::Find:
            if (++_ticks == FIND_TICKS) _phase = Phase::Replace;
            return true;
        case Phase::Replace:
            if (++_step >= (std::max)(_from.size(), _to.size())) {
                _phase = Phase::Status;
                _ticks = 0;
            }
            return true;
        case Phase::Status:
            if (++_ticks == STATUS_TICKS) _phase = Phase::Idle;
            return _phase != Phase::Idle;
        default:
            return false;
        }
    }

    void stop() { _phase = Phase::Idle; }

    Phase phase() const { return _phase; }
    bool isActive() const { return _phase != Phase::Idle; }
    const std::wstring& target() const { return _to; }

    // Already replaced, drawn plain
    std::wstring_view done() const {
        if (_phase == Phase::Find) return {};
        if (_phase == Phase::Replace) return std::wstring_view(_to).substr(0, (std::min)(_step, _to.size()));
        return _to;
    }

    // Not yet replaced, drawn marked
    std::wstring_view rest() const {
        if (_phase == Phase::Find) return _from;
        if (_phase == Phase::Replace) return std::wstring_view(_from).substr((std::min)(_step, _from.size()));
        return {};
    }

private:
    static constexpr int FIND_TICKS = 16;     // about half a second
    static constexpr int STATUS_TICKS = 57;   // about two seconds

    Phase _phase = Phase::Idle;
    std::wstring _from;
    std::wstring _to;
    size_t _step = 0;
    int _ticks = 0;
};
