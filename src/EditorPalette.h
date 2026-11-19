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

// Colors MR draws into the text, fitted to the editor's theme. Notepad++'s dark mode only tells how
// the interface looks; the theme decides what lies behind the text.
namespace EditorPalette {

    using Color = std::uint32_t;                 // a Windows COLORREF, 0x00BBGGRR
    constexpr Color NO_COLOR = 0xFFFFFFFF;

    // Least contrast ratios (WCAG)
    constexpr double MARK_CONTRAST = 3.0;        // dots, frames and colored labels against their background
    constexpr double TEXT_CONTRAST = 4.5;        // text on a background MR paints itself
    constexpr double FILL_CONTRAST = 3.0;        // the theme's text over a fill MR lays under it, at most the theme's own

    // What a view paints behind the text; dark when white stands out more than black
    struct Theme {
        Color background = 0xFFFFFF;
        Color text = 0x000000;
        Color caretLine = NO_COLOR;              // the current line's background, if it is filled
        bool dark = false;
    };
    Theme makeTheme(Color background, Color text, Color caretLine = NO_COLOR);

    double contrast(Color a, Color b);           // 1 to 21
    bool isDark(Color background);
    Color blend(Color over, Color under, int alpha);   // alpha 0 to 255, as Scintilla lays a fill over the background

    // A mark at least MARK_CONTRAST off the background and, where possible, off the current line
    Color visible(Color color, const Theme& theme);
    // Text with at least the given contrast to its background
    Color readable(Color text, Color background, double minimum = TEXT_CONTRAST);
    // The highest alpha up to the given one at which a fill under the theme's text keeps FILL_CONTRAST
    int fillAlpha(Color fill, int alpha, const Theme& theme);
}
