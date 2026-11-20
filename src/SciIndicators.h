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

#include <cstdint>
#include <functional>

#include "Scintilla.h"

// Indicators of a Scintilla document, through send(message, wParam, lParam) to a view that shows it
namespace SciIndicators {

    using Send = std::function<std::intptr_t(unsigned int, std::uintptr_t, std::intptr_t)>;

    // Notepad++ and other plugins fill with the document's current indicator and value: both go back on leaving
    class KeepCurrent {
    public:
        explicit KeepCurrent(const Send& doc)
            : doc_(doc), indicator_(doc(SCI_GETINDICATORCURRENT, 0, 0)), value_(doc(SCI_GETINDICATORVALUE, 0, 0)) {}
        ~KeepCurrent()
        {
            doc_(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(indicator_), 0);
            doc_(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(value_), 0);
        }
        KeepCurrent(const KeepCurrent&) = delete;
        KeepCurrent& operator=(const KeepCurrent&) = delete;

    private:
        const Send& doc_;
        const std::intptr_t indicator_;
        const std::intptr_t value_;
    };

    // Calls visit(start, end, value) for every run of the indicator, the unmarked ones with value 0
    template <class Visit>
    void forEachRun(const Send& doc, int id, Visit visit)
    {
        const std::intptr_t length = doc(SCI_GETLENGTH, 0, 0);
        for (std::intptr_t pos = 0; pos < length;) {
            const std::intptr_t end = doc(SCI_INDICATOREND, static_cast<std::uintptr_t>(id), pos);
            if (end <= pos) break;  // the document never held this indicator
            visit(pos, end, doc(SCI_INDICATORVALUEAT, static_cast<std::uintptr_t>(id), pos));
            pos = end;
        }
    }
}
