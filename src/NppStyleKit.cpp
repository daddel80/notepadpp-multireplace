// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025 Thomas Knoefel
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

#include "NppStyleKit.h"
#include "Notepad_plus_msgs.h"
#include "Scintilla.h"

#include <shlwapi.h>

#include <algorithm>
#include <iterator>

namespace NppStyleKit {

    // ---------- ThemeUtils ----------
    namespace {
        // NppDarkMode::Colors, the layout NPPM_GETDARKMODECOLORS fills
        struct DarkModeColors
        {
            COLORREF background, softerBackground, hotBackground, pureBackground, errorBackground;
            COLORREF text, darkerText, disabledText, linkText;
            COLORREF edge, hotEdge, disabledEdge;
        };

        // Colors of the active dark mode tone; Notepad++'s defaults if the query fails
        DarkModeColors darkModeColors(HWND hNpp)
        {
            DarkModeColors colors{};
            if (!::SendMessage(hNpp, NPPM_GETDARKMODECOLORS, sizeof(colors), reinterpret_cast<LPARAM>(&colors)))
            {
                colors.disabledText = RGB(0x80, 0x80, 0x80);
                colors.linkText = RGB(0xFF, 0xFF, 0x00);
            }
            return colors;
        }
    }

    bool ThemeUtils::isDarkMode(HWND hNpp)
    {
        return (::SendMessage(hNpp, NPPM_ISDARKMODEENABLED, 0, 0) != 0);
    }

    COLORREF ThemeUtils::linkTextColor(HWND hNpp)
    {
        return isDarkMode(hNpp) ? darkModeColors(hNpp).linkText : ::GetSysColor(COLOR_HOTLIGHT);
    }

    COLORREF ThemeUtils::disabledTextColor(HWND hNpp)
    {
        return isDarkMode(hNpp) ? darkModeColors(hNpp).disabledText : ::GetSysColor(COLOR_GRAYTEXT);
    }

    EditorPalette::Theme ThemeUtils::editorTheme(HWND hSci)
    {
        if (!hSci) return EditorPalette::makeTheme(RGB(0xFF, 0xFF, 0xFF), RGB(0, 0, 0));
        const auto back = static_cast<EditorPalette::Color>(::SendMessage(hSci, SCI_STYLEGETBACK, STYLE_DEFAULT, 0));
        const auto text = static_cast<EditorPalette::Color>(::SendMessage(hSci, SCI_STYLEGETFORE, STYLE_DEFAULT, 0));

        // The current line counts when it is filled; as a frame it leaves the background as it is
        EditorPalette::Color caretLine = EditorPalette::NO_COLOR;
        if (::SendMessage(hSci, SCI_GETELEMENTISSET, SC_ELEMENT_CARET_LINE_BACK, 0) && ::SendMessage(hSci, SCI_GETCARETLINEFRAME, 0, 0) == 0) {
            const auto argb = static_cast<EditorPalette::Color>(::SendMessage(hSci, SCI_GETELEMENTCOLOUR, SC_ELEMENT_CARET_LINE_BACK, 0));
            caretLine = EditorPalette::blend(argb & 0xFFFFFF, back, static_cast<int>(argb >> 24));
        }
        return EditorPalette::makeTheme(back, text, caretLine);
    }

    // ---------- ColorTools ----------
    COLORREF ColorTools::tabColorSwatch(int colorId, bool darkMode)
    {
        // Default tab colors of Notepad++ (stylers.model.xml, the same in every bundled theme)
        static constexpr COLORREF kLight[] = { RGB(0xF3, 0xF0, 0xCB), RGB(0xDB, 0xF3, 0xCB), RGB(0xCB, 0xDB, 0xF3), RGB(0xF3, 0xDB, 0xCB), RGB(0xF3, 0xCB, 0xEE) };
        static constexpr COLORREF kDark[]  = { RGB(0x80, 0x78, 0x48), RGB(0x56, 0x80, 0x48), RGB(0x50, 0x70, 0x94), RGB(0x80, 0x48, 0x49), RGB(0x75, 0x48, 0x80) };
        if (colorId < 0 || colorId >= static_cast<int>(std::size(kLight)))
            return CLR_INVALID;

        // Lightness and saturation as Notepad++ sets them for its menu swatches
        WORD hue = 0, lightness = 0, saturation = 0;
        ::ColorRGBToHLS((darkMode ? kDark : kLight)[colorId], &hue, &lightness, &saturation);
        lightness = darkMode ? 146 : 140;
        saturation = static_cast<WORD>((std::min)(240, saturation + (darkMode ? 100 : 30)));
        return ::ColorHLSToRGB(hue, lightness, saturation);
    }

} // namespace NppStyleKit
