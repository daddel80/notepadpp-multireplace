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
#include "DocScope.h"
#include "Notepad_plus_msgs.h"
#include "Scintilla.h"
#include "menuCmdID.h"

#include <shlwapi.h>

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
    namespace {
        // The menu below menu (depth levels down) that holds the item id itself
        HMENU findMenuHolding(HMENU menu, UINT id, int depth)
        {
            const int count = ::GetMenuItemCount(menu);
            for (int pos = 0; pos < count; ++pos)
                if (::GetMenuItemID(menu, pos) == id) return menu;
            if (depth > 0) {
                for (int pos = 0; pos < count; ++pos) {
                    const HMENU sub = ::GetSubMenu(menu, pos);
                    const HMENU found = sub ? findMenuHolding(sub, id, depth - 1) : nullptr;
                    if (found) return found;
                }
            }
            return nullptr;
        }

        // The color inside the black frame of a Notepad++ menu color sample
        COLORREF sampleColor(HBITMAP sample)
        {
            BITMAP info{};
            if (!sample || ::GetObject(sample, sizeof(info), &info) == 0 || info.bmWidth < 3 || info.bmHeight < 3)
                return CLR_INVALID;
            COLORREF color = CLR_INVALID;
            if (const HDC hdc = ::CreateCompatibleDC(nullptr)) {
                if (const HGDIOBJ old = ::SelectObject(hdc, sample)) {
                    color = ::GetPixel(hdc, info.bmWidth / 2, info.bmHeight / 2);
                    ::SelectObject(hdc, old);
                }
                ::DeleteDC(hdc);
            }
            return color;
        }
    }

    std::vector<COLORREF> ColorTools::nppTabColors(HWND hNpp)
    {
        // View > Tab holds the Apply Color items; kept while Notepad++ keeps that menu
        static HMENU tabMenu = nullptr;
        if (!tabMenu || !::IsMenu(tabMenu) || findMenuHolding(tabMenu, IDM_VIEW_TAB_COLOUR_1, 0) != tabMenu) {
            const auto mainMenu = reinterpret_cast<HMENU>(::SendMessage(hNpp, NPPM_GETMENUHANDLE, NPPMAINMENU, 0));
            tabMenu = mainMenu ? findMenuHolding(mainMenu, IDM_VIEW_TAB_COLOUR_1, 3) : nullptr;
        }
        if (!tabMenu) return {};

        // Notepad++ hangs the same sample on both check marks of an Apply Color item
        std::vector<DocScope::ColorMenuItem> items;
        std::vector<HBITMAP> samples;
        const int count = ::GetMenuItemCount(tabMenu);
        for (int pos = 0; pos < count; ++pos) {
            MENUITEMINFOW mii{};
            mii.cbSize = sizeof(mii);
            mii.fMask = MIIM_ID | MIIM_FTYPE | MIIM_SUBMENU | MIIM_CHECKMARKS;
            DocScope::ColorMenuItem item;
            HBITMAP sample = nullptr;
            if (::GetMenuItemInfoW(tabMenu, static_cast<UINT>(pos), TRUE, &mii)) {
                sample = mii.hbmpUnchecked ? mii.hbmpUnchecked : mii.hbmpChecked;
                item = { mii.wID, (mii.fType & MFT_SEPARATOR) != 0 || mii.hSubMenu != nullptr, sample != nullptr };
            }
            else {
                item.other = true;
            }
            items.push_back(item);
            samples.push_back(sample);
        }

        std::vector<COLORREF> colors;
        for (const size_t pos : DocScope::colorMenuPositions(items, IDM_VIEW_TAB_COLOUR_1, IDM_VIEW_TAB_COLOUR_5, IDM_VIEW_TAB_COLOUR_NONE))
            colors.push_back(sampleColor(samples[pos]));
        return colors;
    }

} // namespace NppStyleKit
