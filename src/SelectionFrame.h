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
#include <vector>

// The area the Selection scope remembers, kept in the document as an indicator: Scintilla moves it with
// every edit and draws it as a thin frame. Only frames of the current generation count.
namespace SelectionFrame {

    // A message to one Scintilla view
    using Send = std::function<std::intptr_t(unsigned int, std::uintptr_t, std::intptr_t)>;

    struct Span {
        std::intptr_t start;
        std::intptr_t end;
    };

    class Frame {
    public:
        void setIndicator(int id) noexcept { id_ = id; }
        int indicator() const noexcept { return id_; }

        // Per view: a thin outline over a light fill, under the text
        void applyStyle(const Send& view, bool dark) const;

        // Per document, leaving its current indicator and value as found; draw replaces the frame
        void draw(const Send& doc, const std::vector<Span>& spans) const;
        void erase(const Send& doc) const;
        std::vector<Span> read(const Send& doc) const;
        void dropStale(const Send& doc) const;

        // The frames drawn so far stop counting; each document drops its own in dropStale
        void retire() noexcept;

    private:
        int id_ = -1;
        int generation_ = 1;
    };
}
