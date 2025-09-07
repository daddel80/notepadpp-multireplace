// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
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

#define NOMINMAX
#define WM_UPDATE_FOCUS (WM_APP + 2)

#include "StaticDialog/StaticDialog.h"
#include "BatchUIGuard.h"
#include "MultiReplacePanel.h"
#include "Notepad_plus_msgs.h"
#include "PluginDefinition.h"
#include "Scintilla.h"
#include "DPIManager.h"
#include "luaEmbedded.h"
#include "HiddenSciGuard.h"
#include "ResultDock.h"
#include "Encoding.h"
#include "LanguageManager.h"
#include "ConfigManager.h"
#include "UndoRedoManager.h"
#include <algorithm>
#include <bitset>
#include <fstream>
#include <map>
#include <numeric>
#include <regex>
#include <set>
#include <sstream>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <filesystem>

#include <windows.h>
#include <Commctrl.h>
#include <iomanip>
#include <lua.hpp>
#include <sdkddkver.h>


static LanguageManager& LM = LanguageManager::instance();
static ConfigManager& CFG = ConfigManager::instance();
static UndoRedoManager& URM = UndoRedoManager::instance();


#pragma region Initialization

void MultiReplace::initializeWindowSize()
{
    loadUIConfigFromIni();

    HMONITOR hMonitor = MonitorFromRect(&windowRect, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo = { sizeof(monitorInfo) };

    if (GetMonitorInfo(hMonitor, &monitorInfo))
    {
        int monitorLeft = monitorInfo.rcWork.left;
        int monitorTop = monitorInfo.rcWork.top;
        int monitorRight = monitorInfo.rcWork.right;
        int monitorBottom = monitorInfo.rcWork.bottom;

        int windowWidth = windowRect.right - windowRect.left;
        int windowHeight = windowRect.bottom - windowRect.top;

        const int visibilityMargin = 10;

        bool isCompletelyOffScreen =
            (windowRect.right <= monitorLeft || windowRect.left >= monitorRight ||
                windowRect.bottom <= monitorTop || windowRect.top >= monitorBottom);

        if (isCompletelyOffScreen)
        {
            windowRect.left = monitorLeft + visibilityMargin;
            windowRect.top = monitorTop + visibilityMargin;
        }
        else
        {
            if (windowRect.left < monitorLeft + visibilityMargin)
                windowRect.left = monitorLeft + visibilityMargin;
            if (windowRect.top < monitorTop + visibilityMargin)
                windowRect.top = monitorTop + visibilityMargin;
            if (windowRect.left + windowWidth > monitorRight - visibilityMargin)
                windowRect.left = monitorRight - windowWidth - visibilityMargin;
            if (windowRect.top + windowHeight > monitorBottom - visibilityMargin)
                windowRect.top = monitorBottom - windowHeight - visibilityMargin;
        }

        windowRect.right = windowRect.left + windowWidth;
        windowRect.bottom = windowRect.top + windowHeight;
    }

    SetWindowPos(
        _hSelf,
        NULL,
        windowRect.left,
        windowRect.top,
        windowRect.right - windowRect.left,
        windowRect.bottom - windowRect.top,
        SWP_NOZORDER
    );
}

void MultiReplace::initializeFontStyles() {
    if (!dpiMgr) return;

    // Helper lambda to create a font
    auto createFont = [&](int height, int weight, const wchar_t* fontName) {
        return ::CreateFont(
            dpiMgr->scaleY(height),  // Scale font height
            0,                       // Default font width
            0,                       // Escapement
            0,                       // Orientation
            weight,                  // Font weight
            FALSE,                   // Italic
            FALSE,                   // Underline
            FALSE,                   // Strikeout
            DEFAULT_CHARSET,         // Character set
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            fontName                 // Font name
        );
        };

    // Define standard and normal fonts
    _hStandardFont = createFont(13, FW_NORMAL, L"MS Shell Dlg 2");
    _hNormalFont1 = createFont(14, FW_NORMAL, L"MS Shell Dlg 2");
    _hNormalFont2 = createFont(12, FW_NORMAL, L"Courier New");
    _hNormalFont3 = createFont(14, FW_NORMAL, L"Courier New");
    _hNormalFont4 = createFont(16, FW_NORMAL, L"Courier New");
    _hNormalFont5 = createFont(18, FW_NORMAL, L"Courier New");
    _hNormalFont6 = createFont(22, FW_NORMAL, L"Courier New");

    // Apply standard font to all controls in ctrlMap
    for (const auto& pair : ctrlMap) {
        SendMessage(GetDlgItem(_hSelf, pair.first), WM_SETFONT, (WPARAM)_hStandardFont, TRUE);
    }

    // Specific controls using normal fonts
    for (int controlId : { IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_STATUS_MESSAGE, IDC_PATH_DISPLAY, IDC_STATS_DISPLAY }) {
        SendMessage(GetDlgItem(_hSelf, controlId), WM_SETFONT, (WPARAM)_hNormalFont1, TRUE);
    }
    for (int controlId : { IDC_COLUMN_DROP_BUTTON, IDC_COLUMN_HIGHLIGHT_BUTTON }) {
        SendMessage(GetDlgItem(_hSelf, controlId), WM_SETFONT, (WPARAM)_hNormalFont2, TRUE);
    }

    SendMessage(GetDlgItem(_hSelf, IDC_SAVE_BUTTON), WM_SETFONT, (WPARAM)_hNormalFont3, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_COLUMN_COPY_BUTTON), WM_SETFONT, (WPARAM)_hNormalFont3, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), WM_SETFONT, (WPARAM)_hNormalFont4, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_BUTTON), WM_SETFONT, (WPARAM)_hNormalFont5, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_ALL_SMALL_BUTTON), WM_SETFONT, (WPARAM)_hNormalFont6, TRUE);

    // Define bold fonts
    _hBoldFont1 = createFont(22, FW_BOLD, L"Courier New");
    _hBoldFont2 = createFont(12, FW_BOLD, L"MS Shell Dlg 2");

    // Specific controls using bold fonts, adjusted to match the correct sizes
    SendMessage(GetDlgItem(_hSelf, IDC_SWAP_BUTTON), WM_SETFONT, (WPARAM)_hBoldFont1, TRUE);
    SendMessage(GetDlgItem(_hSelf, ID_EDIT_EXPAND_BUTTON), WM_SETFONT, (WPARAM)_hBoldFont2, TRUE);

    // For ListView: calculate widths of special characters and add padding
    checkMarkWidth_scaled = getCharacterWidth(IDC_REPLACE_LIST, L"\u2714") + 15;
    crossWidth_scaled = getCharacterWidth(IDC_REPLACE_LIST, L"\u2716") + 15;
    boxWidth_scaled = getCharacterWidth(IDC_REPLACE_LIST, L"\u2610") + 15;

    // Set delete button column width
    deleteButtonColumnWidth = crossWidth_scaled;
}

RECT MultiReplace::calculateMinWindowFrame(HWND hwnd) {
    // Use local variables to avoid modifying windowRect
    RECT tempWindowRect;
    GetWindowRect(hwnd, &tempWindowRect);

    // Measure the window's borders and title bar
    RECT clientRect;
    GetClientRect(hwnd, &clientRect);

    int borderWidth = ((tempWindowRect.right - tempWindowRect.left) - (clientRect.right - clientRect.left)) / 2;
    int titleBarHeight = (tempWindowRect.bottom - tempWindowRect.top) - (clientRect.bottom - clientRect.top) - borderWidth;

    // Base minimum content height (list on/off)
    int minContentHeight = useListEnabled ? MIN_HEIGHT_scaled : SHRUNK_HEIGHT_scaled;

    // Add extra room if “Replace in Files” panel is visible AND we're not in Two-Buttons-Mode
    bool twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;
    int panelExtra = ((isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode) ? sy(REPLACE_FILES_PANEL_HEIGHT) : 0;

    int minHeight = minContentHeight + panelExtra;
    int minWidth = MIN_WIDTH_scaled;

    // Adjust for window borders and title bar
    minHeight += borderWidth + titleBarHeight;
    minWidth += 2 * borderWidth;

    RECT minSize = { 0, 0, minWidth, minHeight };

    return minSize;
}

void MultiReplace::positionAndResizeControls(int windowWidth, int windowHeight)
{
    // Helper functions sx() and sy() are used to scale values for DPI awareness.
    // sx(value): scales the x-axis value based on DPI settings.
    // sy(value): scales the y-axis value based on DPI settings.

    // DPI Aware System Metrics for Checkboxes and Radio Buttons
    UINT dpi = dpiMgr->getDPIX(); // Get the DPI from DPIManager

    // Calculate appropriate height for checkboxes and radiobuttons
    int checkboxBaseHeight = dpiMgr->getCustomMetricOrFallback(SM_CYMENUCHECK, dpi, 14);
    int radioButtonBaseHeight = dpiMgr->getCustomMetricOrFallback(SM_CYMENUCHECK, dpi, 14);

    // Get the font height from the standard font
    int fontHeight = getFontHeight(_hSelf, _hStandardFont);
    fontHeight = fontHeight + sy(8); // Padding Font

    // Choose the larger value between the font height and the base height
    int checkboxHeight = std::max(checkboxBaseHeight, fontHeight);
    int radioButtonHeight = std::max(radioButtonBaseHeight, fontHeight);

    // Calculate dimensions without scaling
    BOOL twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;
    int filesOffsetY = ((isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode) ? sy(REPLACE_FILES_PANEL_HEIGHT) : 0;
    int buttonX = windowWidth - sx(33 + 128);
    int checkbox2X = buttonX + sx(134);
    int useListButtonX = buttonX + sx(133);
    int swapButtonX = windowWidth - sx(33 + 128 + 26);
    int comboWidth = windowWidth - sx(289);
    int listWidth = windowWidth - sx(207);
    int listHeight = std::max(windowHeight - sy(245) - filesOffsetY, sy(20)); // Minimum listHeight to prevent IDC_PATH_DISPLAY from overlapping with IDC_STATUS_MESSAGE
    int useListButtonY = windowHeight - sy(34);

    // Apply scaling only when assigning to ctrlMap
    ctrlMap[IDC_STATIC_FIND] = { sx(11), sy(18), sx(80), sy(19), WC_STATIC, LM.getLPCW(L"panel_find_what"), SS_RIGHT, NULL };
    ctrlMap[IDC_STATIC_REPLACE] = { sx(11), sy(47), sx(80), sy(19), WC_STATIC, LM.getLPCW(L"panel_replace_with"), SS_RIGHT };

    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { sx(16), sy(76), sx(158), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_match_whole_word_only"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { sx(16), sy(101), sx(158), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_match_case"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_CHECKBOX] = { sx(16), sy(126), sx(134), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_use_variables"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_HELP] = { sx(152), sy(126), sx(20), sy(20), WC_BUTTON, LM.getLPCW(L"panel_help"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_WRAP_AROUND_CHECKBOX] = { sx(16), sy(151), sx(158), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_wrap_around"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_AT_MATCHES_CHECKBOX] = { sx(16), sy(176), sx(112), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_replace_at_matches"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_HIT_EDIT] = { sx(130), sy(176), sx(41), sy(16), WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL,  LM.getLPCW(L"tooltip_replace_at_matches") };

    ctrlMap[IDC_SEARCH_MODE_GROUP] = { sx(180), sy(79), sx(173), sy(104), WC_BUTTON, LM.getLPCW(L"panel_search_mode"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_NORMAL_RADIO] = { sx(188), sy(101), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_normal"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXTENDED_RADIO] = { sx(188), sy(126), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_extended"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REGEX_RADIO] = { sx(188), sy(150), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_regular_expression"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_SCOPE_GROUP] = { sx(367), sy(79), sx(203), sy(125), WC_BUTTON, LM.getLPCW(L"panel_scope"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_ALL_TEXT_RADIO] = { sx(375), sy(101), sx(189), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_all_text"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_SELECTION_RADIO] = { sx(375), sy(126), sx(189), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_selection"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COLUMN_MODE_RADIO] = { sx(375), sy(150), sx(45), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_csv"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_COLUMN_NUM_STATIC] = { sx(369), sy(181), sx(30), sy(20), WC_STATIC, LM.getLPCW(L"panel_cols"), SS_RIGHT, NULL };
    ctrlMap[IDC_COLUMN_NUM_EDIT] = { sx(400), sy(181), sx(41), sy(16), WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_columns") };
    ctrlMap[IDC_DELIMITER_STATIC] = { sx(442), sy(181), sx(38), sy(20), WC_STATIC, LM.getLPCW(L"panel_delim"), SS_RIGHT, NULL };
    ctrlMap[IDC_DELIMITER_EDIT] = { sx(481), sy(181), sx(25), sy(16), WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_delimiter") };
    ctrlMap[IDC_QUOTECHAR_STATIC] = { sx(506), sy(181), sx(37), sy(20), WC_STATIC, LM.getLPCW(L"panel_quote"), SS_RIGHT, NULL };
    ctrlMap[IDC_QUOTECHAR_EDIT] = { sx(544), sy(181), sx(15), sy(16), WC_EDIT, NULL, ES_CENTER | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_quote") };

    ctrlMap[IDC_COLUMN_SORT_DESC_BUTTON] = { sx(418), sy(149), sx(19), sy(20), WC_BUTTON, symbolSortDesc, BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_sort_descending") };
    ctrlMap[IDC_COLUMN_SORT_ASC_BUTTON] = { sx(437), sy(149), sx(19), sy(20), WC_BUTTON, symbolSortAsc, BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_sort_ascending") };
    ctrlMap[IDC_COLUMN_DROP_BUTTON] = { sx(459), sy(149), sx(25), sy(20), WC_BUTTON, L"✖", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_drop_columns") };
    ctrlMap[IDC_COLUMN_COPY_BUTTON] = { sx(487), sy(149), sx(25), sy(20), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_copy_columns") }; // 
    ctrlMap[IDC_COLUMN_HIGHLIGHT_BUTTON] = { sx(515), sy(149), sx(45), sy(20), WC_BUTTON, L"🖍", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_column_highlight") };

    // Dynamic positions and sizes
    ctrlMap[IDC_FIND_EDIT] = { sx(96), sy(14), comboWidth, sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_EDIT] = { sx(96), sy(44), comboWidth, sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_SWAP_BUTTON] = { swapButtonX, sy(26), sx(22), sy(27), WC_BUTTON, L"⇅", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, sy(14), sx(128), sy(52), WC_BUTTON, LM.getLPCW(L"panel_add_into_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, sy(91), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_replace_all"), BS_SPLITBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_BUTTON] = { buttonX, sy(91), sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_replace"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_SMALL_BUTTON] = { buttonX + sx(100), sy(91), sx(28), sy(24), WC_BUTTON, L"↻", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_replace_all") };
    ctrlMap[IDC_2_BUTTONS_MODE] = { checkbox2X, sy(91), sx(20), sy(20), WC_BUTTON, L"", BS_AUTOCHECKBOX | WS_TABSTOP, LM.getLPCW(L"tooltip_2_buttons_mode") };
    ctrlMap[IDC_FIND_ALL_BUTTON] = { buttonX, sy(119), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_find_all"), BS_SPLITBUTTON | WS_TABSTOP, NULL };

    findNextButtonText = L"▼ " + LM.get(L"panel_find_next_small");
    ctrlMap[IDC_FIND_NEXT_BUTTON] = ControlInfo{ buttonX + sx(32), sy(119), sx(96), sy(24), WC_BUTTON, findNextButtonText.c_str(), BS_PUSHBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_FIND_PREV_BUTTON] = { buttonX, sy(119), sx(28), sy(24), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_BUTTON] = { buttonX, sy(147), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_mark_matches"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, sy(147), sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_mark_matches_small"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX + sx(100), sy(147), sx(28), sy(24), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_copy_marked_text") }; // 
    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, sy(175), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_clear_all_marks"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_STATUS_MESSAGE] = { sx(19), sy(205) + filesOffsetY, listWidth - sx(5), sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_ENDELLIPSIS | SS_NOPREFIX | SS_OWNERDRAW, NULL };
    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, sy(227) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_LOAD_LIST_BUTTON] = { buttonX, sy(227) + filesOffsetY, sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_NEW_LIST_BUTTON] = { buttonX + sx(100), sy(227) + filesOffsetY, sx(28), sy(24), WC_BUTTON, L"➕", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_new_list") };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, sy(255) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_save_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_SAVE_BUTTON] = { buttonX, sy(255) + filesOffsetY, sx(28), sy(24), WC_BUTTON, L"💾", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_save") };
    ctrlMap[IDC_SAVE_AS_BUTTON] = { buttonX + sx(32), sy(255) + filesOffsetY, sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_save_as"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, sy(283) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_export_to_bash"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_UP_BUTTON] = { buttonX + sx(4), sy(323) + filesOffsetY, sx(24), sy(24), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_DOWN_BUTTON] = { buttonX + sx(4), sy(323 + 24 + 4) + filesOffsetY, sx(24), sy(24), WC_BUTTON, L"▼", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_SHIFT_FRAME] = { buttonX, sy(323 - 11) + filesOffsetY, sx(128), sy(68), WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_SHIFT_TEXT] = { buttonX + sx(30), sy(323 + 16) + filesOffsetY, sx(96), sy(16), WC_STATIC, LM.getLPCW(L"panel_move_lines"), SS_LEFT, NULL };
    ctrlMap[IDC_REPLACE_LIST] = { sx(14), sy(227) + filesOffsetY, listWidth, listHeight, WC_LISTVIEW, NULL, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS, NULL };
    ctrlMap[IDC_PATH_DISPLAY] = { sx(14), sy(225) + listHeight + sy(5) + filesOffsetY, listWidth, sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_NOTIFY, NULL };
    ctrlMap[IDC_STATS_DISPLAY] = { sx(14) + listWidth, sy(225) + listHeight + sy(5) + filesOffsetY, 0, sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_NOTIFY, NULL };
    ctrlMap[IDC_USE_LIST_BUTTON] = { useListButtonX, useListButtonY , sx(22), sy(22), WC_BUTTON, L"-", BS_PUSHBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_CANCEL_REPLACE_BUTTON] = { buttonX, sy(260), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_cancel_replace"), BS_PUSHBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_FILE_OPS_GROUP] = { sx(14), sy(210), listWidth, sy(80), WC_BUTTON,LM.getLPCW(L"panel_replace_in_files"), BS_GROUPBOX, NULL };

    ctrlMap[IDC_FILTER_STATIC] = { sx(15),  sy(230), sx(75),  sy(19), WC_STATIC, LM.getLPCW(L"panel_filter"), SS_RIGHT, NULL };
    ctrlMap[IDC_FILTER_EDIT] = { sx(96),  sy(230), comboWidth - sx(170),  sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_FILTER_HELP] = { sx(96) + comboWidth - sx(170) + sx(5), sy(228), sx(20), sy(20), WC_STATIC, L"(?)", SS_CENTER | SS_OWNERDRAW | SS_NOTIFY, LM.getLPCW(L"tooltip_filter_help") };
    ctrlMap[IDC_DIR_STATIC] = { sx(15),  sy(257), sx(75),  sy(19), WC_STATIC, LM.getLPCW(L"panel_directory"), SS_RIGHT, NULL };
    ctrlMap[IDC_DIR_EDIT] = { sx(96),  sy(257), comboWidth - sx(170),  sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_BROWSE_DIR_BUTTON] = { comboWidth - sx(70), sy(257), sx(20),  sy(20), WC_BUTTON, L"...", BS_PUSHBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_SUBFOLDERS_CHECKBOX] = { comboWidth - sx(20), sy(230), sx(120), sy(13), WC_BUTTON, LM.getLPCW(L"panel_in_subfolders"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_HIDDENFILES_CHECKBOX] = { comboWidth - sx(20), sy(257), sx(120), sy(13), WC_BUTTON, LM.getLPCW(L"panel_in_hidden_folders"),BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
}

void MultiReplace::initializeCtrlMap() {

    hInstance = (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE);
    s_hDlg = _hSelf;

    // Get the client rectangle
    RECT rcClient;
    GetClientRect(_hSelf, &rcClient);
    // Extract width and height from the RECT
    int windowWidth = rcClient.right - rcClient.left;
    int windowHeight = rcClient.bottom - rcClient.top;

    // Define Position for all Elements
    positionAndResizeControls(windowWidth, windowHeight);

    // Now iterate over the controls and create each one.
    if (!createAndShowWindows()) {
        return;
    }

    // Show or hide the "Export to Bash" button according to the ExportToBash option
    ShowWindow(GetDlgItem(_hSelf, IDC_EXPORT_BASH_BUTTON), exportToBashEnabled ? SW_SHOW : SW_HIDE);

    // immediately hide or show the files sub-panel
    updateFilesPanel();

    // Initialize the tooltip for the "Use List" button with dynamic text
    updateUseListState(false);

    // Limit the input for IDC_QUOTECHAR_EDIT to one character
    SendMessage(GetDlgItem(_hSelf, IDC_QUOTECHAR_EDIT), EM_SETLIMITTEXT, (WPARAM)1, 0);

    // Enable IDC_SELECTION_RADIO based on text selection
    SelectionInfo selection = getSelectionInfo(false);
    bool isTextSelected = (selection.length > 0);
    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), isTextSelected);

    isWindowOpen = true;
}

bool MultiReplace::createAndShowWindows() {
    // IDs of all controls in the “Replace/Find in Files” panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    const bool twoButtonsMode = (IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED);
    const bool initialShow = (isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode;

    auto isRepInFilesId = [&](int id) {
        return std::find(repInFilesIds.begin(), repInFilesIds.end(), id) != repInFilesIds.end();
        };

    for (auto& pair : ctrlMap)
    {
        const bool isFilesCtrl = isRepInFilesId(pair.first);

        // Create all controls as children, but only set WS_VISIBLE if needed
        DWORD style = pair.second.style | WS_CHILD;
        if (!isFilesCtrl || initialShow) {
            style |= WS_VISIBLE;
        }

        HWND hwndControl = CreateWindowEx(
            0,
            pair.second.className,
            pair.second.windowName,
            style,
            pair.second.x, pair.second.y, pair.second.cx, pair.second.cy,
            _hSelf,
            (HMENU)(INT_PTR)pair.first,
            hInstance,
            NULL
        );
        if (!hwndControl) {
            return false; // Handle creation error
        }

        // Only create tooltips if enabled and tooltip text is available
        if ((tooltipsEnabled || pair.first == IDC_FILTER_HELP)
            && pair.second.tooltipText != nullptr
            && pair.second.tooltipText[0] != L'\0')
        {
            // Create a tooltip window for this control
            HWND hwndTooltip = CreateWindowEx(
                0,                                   // no extended styles
                TOOLTIPS_CLASS,                      // tooltip class
                nullptr,
                WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON | TTS_NOPREFIX,
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                _hSelf,                              // parent = our panel/dialog
                nullptr,
                hInstance,
                nullptr
            );

            if (hwndTooltip)
            {
                // Limit width only for the "?" help tooltip; 0 = unlimited
                DWORD maxWidth = (pair.first == IDC_FILTER_HELP) ? 200 : 0;
                SendMessage(hwndTooltip, TTM_SETMAXTIPWIDTH, 0, maxWidth);

                // Bind the tooltip to the specific child control (by HWND)
                TOOLINFO ti = { 0 };
                ti.cbSize = sizeof(ti);
                ti.hwnd = _hSelf;                         // parent window
                ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;    // subclass the control to show tips
                ti.uId = (UINT_PTR)hwndControl;          // identify by child HWND
                ti.lpszText = const_cast<LPWSTR>(pair.second.tooltipText);
                SendMessage(hwndTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
            }
        }
    }
    return true;
}

void MultiReplace::initializeMarkerStyle() {
    // Initialize for non-list marker
    long standardMarkerColor = MARKER_COLOR;
    int standardMarkerStyle = textStyles[0];
    colorToStyleMap[standardMarkerColor] = standardMarkerStyle;

    send(SCI_SETINDICATORCURRENT, standardMarkerStyle, 0);
    send(SCI_INDICSETSTYLE, standardMarkerStyle, INDIC_STRAIGHTBOX);
    send(SCI_INDICSETFORE, standardMarkerStyle, standardMarkerColor);
    send(SCI_INDICSETALPHA, standardMarkerStyle, 100);
}

void MultiReplace::initializeListView() {
    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    SetWindowSubclass(_replaceListView, &MultiReplace::ListViewSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));
    createListViewColumns();
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);


    // Set tooltips based on isHoverTextEnabled
    DWORD extendedStyle = LVS_EX_FULLROWSELECT;
    if (isHoverTextEnabled) {
        extendedStyle |= LVS_EX_INFOTIP; // Enable tooltips
    }
    ListView_SetExtendedListViewStyle(_replaceListView, extendedStyle);

    // Initialize columnSortOrder with Unsorted state for all sortable columns
    columnSortOrder[ColumnID::FIND_COUNT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::REPLACE_COUNT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::FIND_TEXT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::REPLACE_TEXT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::COMMENTS] = SortDirection::Unsorted;

}

void MultiReplace::initializeDragAndDrop() {
    dropTarget = new DropTarget(_replaceListView, this);  // Create an instance of DropTarget
    HRESULT hr = ::RegisterDragDrop(_replaceListView, dropTarget);  // Register the ListView as a drop target

    if (FAILED(hr)) {
        // Safely release the DropTarget instance to avoid memory leaks
        delete dropTarget;
        dropTarget = nullptr;  // Set to nullptr to prevent any unintended usage
    }
}

void MultiReplace::moveAndResizeControls() {
    // IDs of controls to be moved or resized
    const int controlIds[] = {
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_SWAP_BUTTON, IDC_STATIC_FRAME, IDC_COPY_TO_LIST_BUTTON, IDC_REPLACE_ALL_BUTTON,
        IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON, IDC_2_BUTTONS_MODE, IDC_FIND_ALL_BUTTON, IDC_FIND_NEXT_BUTTON,
        IDC_FIND_PREV_BUTTON, IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON,
        IDC_USE_LIST_BUTTON, IDC_CANCEL_REPLACE_BUTTON, IDC_LOAD_FROM_CSV_BUTTON, IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON, IDC_SAVE_TO_CSV_BUTTON,
        IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT, IDC_EXPORT_BASH_BUTTON, 
        IDC_PATH_DISPLAY, IDC_STATS_DISPLAY, IDC_FILTER_EDIT, IDC_FILTER_HELP, IDC_SUBFOLDERS_CHECKBOX, 
        IDC_HIDDENFILES_CHECKBOX, IDC_DIR_EDIT, IDC_BROWSE_DIR_BUTTON, IDC_FILE_OPS_GROUP, IDC_DIR_PROGRESS_BAR, IDC_STATUS_MESSAGE
    };

    std::unordered_map<int, HWND> hwndMap;  // Store HWNDs to avoid multiple calls to GetDlgItem

    // Move and resize controls
    for (int ctrlId : controlIds) {
        const ControlInfo& ctrlInfo = ctrlMap[ctrlId];
        HWND resizeHwnd = GetDlgItem(_hSelf, ctrlId);
        hwndMap[ctrlId] = resizeHwnd;  // Store HWND

        RECT rc;
        GetClientRect(resizeHwnd, &rc);

        DWORD startSelection = 0, endSelection = 0;
        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT 
            || ctrlId == IDC_DIR_EDIT || ctrlId == IDC_FILTER_EDIT) {
            SendMessage(resizeHwnd, CB_GETEDITSEL, (WPARAM)&startSelection, (LPARAM)&endSelection);
        }

        int height = ctrlInfo.cy;
        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT
            || ctrlId == IDC_DIR_EDIT || ctrlId == IDC_FILTER_EDIT) {
            COMBOBOXINFO cbi = { sizeof(COMBOBOXINFO) };
            if (GetComboBoxInfo(resizeHwnd, &cbi)) {
                height = cbi.rcItem.bottom - cbi.rcItem.top;
            }
        }

        MoveWindow(resizeHwnd, ctrlInfo.x, ctrlInfo.y, ctrlInfo.cx, height, TRUE);

        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT
            || ctrlId == IDC_DIR_EDIT || ctrlId == IDC_FILTER_EDIT) {
            SendMessage(resizeHwnd, CB_SETEDITSEL, 0, MAKELPARAM(startSelection, endSelection));
        }
    }

    showListFilePath();

    /*
    // IDs of controls to be redrawn
    const int redrawIds[] = {
        IDC_USE_LIST_BUTTON, IDC_COPY_TO_LIST_BUTTON, IDC_REPLACE_ALL_BUTTON, IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON,
        IDC_2_BUTTONS_MODE, IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON, IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON,
        IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT
    };

    // Redraw controls using stored HWNDs
    for (int ctrlId : redrawIds) {
        InvalidateRect(hwndMap[ctrlId], NULL, TRUE);
    }*/
}

void MultiReplace::updateTwoButtonsVisibility() {
    BOOL twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;

    auto setVisibility = [this](const std::vector<int>& elements, bool condition) {
        for (int id : elements) {
            ShowWindow(GetDlgItem(_hSelf, id), condition ? SW_SHOW : SW_HIDE);
        }
        };

    // Replace-Buttons
    setVisibility({ IDC_REPLACE_ALL_SMALL_BUTTON, IDC_REPLACE_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_REPLACE_ALL_BUTTON }, !twoButtonsMode);

    // Find-Buttons
    setVisibility({ IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_FIND_ALL_BUTTON }, !twoButtonsMode);

    // Mark-Buttons
    setVisibility({ IDC_MARK_MATCHES_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_MARK_BUTTON }, !twoButtonsMode);

    // Load-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_LOAD_FROM_CSV_BUTTON }, !twoButtonsMode);

    // Save-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_SAVE_TO_CSV_BUTTON }, !twoButtonsMode);

    updateFilesPanel();
}

void MultiReplace::updateListViewFrame()
{
    HWND lv = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    if (!lv) return;

    const ControlInfo& ci = ctrlMap[IDC_REPLACE_LIST];
    MoveWindow(lv, ci.x, ci.y, ci.cx, ci.cy, TRUE);
}

void MultiReplace::repaintPanelContents(HWND hGrp, const std::wstring& title)
{
    // All controls inside the "Replace/Find in Files" panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    // Keep the caption in sync
    SetDlgItemText(_hSelf, IDC_FILE_OPS_GROUP, title.c_str());

    // Never repaint the frame if the group is hidden. This avoids "ghost frames".
    if (!IsWindowVisible(hGrp)) {
        return; // Title updated for next show; no drawing while hidden.
    }

    // Map group rect to parent coordinates
    RECT rcGrp;
    GetWindowRect(hGrp, &rcGrp);
    MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcGrp), 2);

    // Clean the parent background in the group area
    RedrawWindow(_hSelf, &rcGrp, NULL,
        RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_NOCHILDREN);

    // Redraw the groupbox (including non-client frame)
    RedrawWindow(hGrp, NULL, NULL,
        RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE | RDW_NOCHILDREN | RDW_FRAME);

    // Redraw visible children (no erase to prevent flicker)
    for (int id : repInFilesIds) {
        if (id == IDC_FILE_OPS_GROUP) continue;
        HWND hChild = GetDlgItem(_hSelf, id);
        if (IsWindow(hChild) && IsWindowVisible(hChild)) {
            RedrawWindow(hChild, NULL, NULL,
                RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
        }
    }
}

void MultiReplace::updateFilesPanel()
{
    // All controls inside the "Replace/Find in Files" panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    // Persisted UI state to avoid unnecessary work
    static bool lastShow = false;
    static std::wstring lastTitleKey;

    const bool twoButtonsMode = (IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED);
    const bool show = (isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode;

    // Determine title
    std::wstring titleKey;
    std::wstring titleText;
    if (isReplaceInFiles && isFindAllInFiles) {
        titleKey = L"panel_find_replace_in_files";
        titleText = LM.get(titleKey);
    }
    else if (isFindAllInFiles) {
        titleKey = L"panel_find_in_files";
        titleText = LM.get(titleKey);
    }
    else {
        titleKey = L"panel_replace_in_files";
        titleText = LM.get(titleKey);
    }

    HWND hGrp = GetDlgItem(_hSelf, IDC_FILE_OPS_GROUP);
    HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);

    // Capture geometry BEFORE layout changes (used to clean freed area when hiding)
    RECT rcGrpBefore{};
    if (IsWindow(hGrp)) {
        GetWindowRect(hGrp, &rcGrpBefore);
        MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcGrpBefore), 2);
    }
    RECT rcStatusBefore{};
    if (IsWindow(hStatus)) {
        GetWindowRect(hStatus, &rcStatusBefore);
        MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcStatusBefore), 2);
    }

    // Case A) Visibility toggles: full show/hide + relayout + localized repaint
    if (show != lastShow)
    {

        // Toggle visibility of all panel controls
        for (int id : repInFilesIds) {
            ShowWindow(GetDlgItem(_hSelf, id), show ? SW_SHOW : SW_HIDE);
        }

        // Layout after visibility change
        RECT rcClient; GetClientRect(_hSelf, &rcClient);
        positionAndResizeControls(rcClient.right - rcClient.left, rcClient.bottom - rcClient.top);
        updateListViewFrame();
        moveAndResizeControls();
        adjustWindowSize();
        onSelectionChanged();

        // Thaw parent painting
        SendMessage(_hSelf, WM_SETREDRAW, TRUE, 0);

        // Compute "after" rect for status (for union repaint when hiding)
        RECT rcStatusAfter{};
        if (IsWindow(hStatus)) {
            GetWindowRect(hStatus, &rcStatusAfter);
            MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcStatusAfter), 2);
        }

        if (show) {
            EnableWindow(GetDlgItem(_hSelf, IDC_CANCEL_REPLACE_BUTTON), FALSE);
            // Showing: repaint panel (title + frame + children) safely
            repaintPanelContents(hGrp, titleText);
            // Clear selection in both edit fields after opening the panel
            SendMessage(GetDlgItem(_hSelf, IDC_FILTER_EDIT), CB_SETEDITSEL, 0, 0);
            SendMessage(GetDlgItem(_hSelf, IDC_DIR_EDIT), CB_SETEDITSEL, 0, 0);
        }
        else {

            // Force controls that moved up to repaint (no erase to avoid flicker)
            const int idsShiftedUp[] = {
                IDC_REPLACE_LIST, IDC_STATUS_MESSAGE, IDC_PATH_DISPLAY, IDC_STATS_DISPLAY,
                IDC_LOAD_FROM_CSV_BUTTON, IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON,
                IDC_SAVE_TO_CSV_BUTTON,  IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON,
                IDC_EXPORT_BASH_BUTTON, IDC_SHIFT_FRAME, IDC_SHIFT_TEXT,
                IDC_UP_BUTTON, IDC_DOWN_BUTTON
            };
            for (int id : idsShiftedUp) {
                HWND hCtrl = GetDlgItem(_hSelf, id);
                if (IsWindow(hCtrl) && IsWindowVisible(hCtrl)) {
                    RedrawWindow(hCtrl, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
                }
            }

            // IMPORTANT: Do NOT change the group title while hidden.
            // No SetDlgItemText() here — prevents any accidental frame repaint.
            lastTitleKey.clear();

        }
        RedrawWindow(_hSelf, NULL, NULL, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);

        lastShow = show;
        if (show) lastTitleKey = titleKey;
        return;
    }

    // Case B) Still hidden: nothing to do (keep state clean)
    if (!show) {
        lastShow = false;
        lastTitleKey.clear();
        return;
    }

    // Case C) Visible and only the title changed
    if (titleKey != lastTitleKey)
    {
        // Repaint only the panel contents/title without relayout
        repaintPanelContents(hGrp, titleText);
        lastTitleKey = titleKey;
    }
}

void MultiReplace::setUIElementVisibility() {
    // Determine the state of mode radio buttons
    bool regexChecked = SendMessage(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_GETCHECK, 0, 0) == BST_CHECKED;
    bool columnModeChecked = SendMessage(GetDlgItem(_hSelf, IDC_COLUMN_MODE_RADIO), BM_GETCHECK, 0, 0) == BST_CHECKED;

    // Update the Whole Word checkbox visibility based on the Regex mode
    EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), !regexChecked);

    const std::vector<int> columnRadioDependentElements = {
        IDC_COLUMN_SORT_DESC_BUTTON, IDC_COLUMN_SORT_ASC_BUTTON,
        IDC_COLUMN_DROP_BUTTON, IDC_COLUMN_COPY_BUTTON, IDC_COLUMN_HIGHLIGHT_BUTTON
    };

    // Update the UI elements based on Column mode
    for (int id : columnRadioDependentElements) {
        EnableWindow(GetDlgItem(_hSelf, id), columnModeChecked);
    }

    // Update the FIND_PREV_BUTTON state based on Regex and Selection mode
    // EnableWindow(GetDlgItem(_hSelf, IDC_FIND_PREV_BUTTON), !regexChecked);
}

void MultiReplace::drawGripper() {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(_hSelf, &ps);

    // Get the size of the client area
    RECT rect;
    GetClientRect(_hSelf, &rect);

    // Determine where to draw the gripper
    int gripperAreaSize = sx(11); // Total size of the gripper area
    POINT startPoint = { rect.right - gripperAreaSize, rect.bottom - gripperAreaSize };

    // Define the new size and reduced gap of the gripper dots
    int dotSize = sx(2); // Increased dot size
    int gap = std::max(sx(1), 1); // Reduced gap between dots

    // Brush Color for Gripper
    HBRUSH hBrush = CreateSolidBrush(RGB(200, 200, 200));

    // Matrix to identify the points to draw
    int positions[3][3] = {
        {0, 0, 1},
        {0, 1, 1},
        {1, 1, 1}
    };

    for (int row = 0; row < 3; ++row)
    {
        for (int col = 0; col < 3; ++col)
        {
            // Skip drawing for omitted positions
            if (positions[row][col] == 0) continue;

            int x = startPoint.x + col * (dotSize + gap);
            int y = startPoint.y + row * (dotSize + gap);
            RECT dotRect = { x, y, x + dotSize, y + dotSize };
            FillRect(hdc, &dotRect, hBrush);
        }
    }

    DeleteObject(hBrush);

    EndPaint(_hSelf, &ps);
}

void MultiReplace::SetWindowTransparency(HWND hwnd, BYTE alpha) {
    LONG style = GetWindowLong(hwnd, GWL_EXSTYLE);
    SetWindowLong(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED);
    SetLayeredWindowAttributes(hwnd, 0, alpha, LWA_ALPHA);
}

void MultiReplace::adjustWindowSize() {

    // Get the minimum allowed window size
    RECT minSize = calculateMinWindowFrame(_hSelf);
    int minHeight = minSize.bottom; // minSize.bottom is minHeight

    // Get the current window position and size
    RECT currentRect;
    GetWindowRect(_hSelf, &currentRect);
    int currentWidth = currentRect.right - currentRect.left;
    int currentX = currentRect.left;
    int currentY = currentRect.top;

    int newHeight = useListEnabled ? std::max(useListOnHeight, minHeight) : useListOffHeight;

    // Adjust the window size while keeping the current position and width
    SetWindowPos(_hSelf, NULL, currentX, currentY, currentWidth, newHeight, SWP_NOZORDER);
}

void MultiReplace::updateUseListState(bool isUpdate)
{   // isUpdate - specifies whether to update the existing tooltip (true) or create a new one (false)

    // Set the button text based on the current state
    SetDlgItemText(_hSelf, IDC_USE_LIST_BUTTON, useListEnabled ? L"˄" : L"˅");

    // Set the status message based on the list state
    showStatusMessage(useListEnabled ? LM.get(L"status_enable_list") : LM.get(L"status_disable_list"), MessageStatus::Info);

    // Return early if tooltips are disabled
    if (!tooltipsEnabled) {
        return;
    }

    // Create the tooltip window if it doesn't exist yet
    if (!isUpdate && !_hUseListButtonTooltip)
    {
        _hUseListButtonTooltip = CreateWindowEx(
            NULL, TOOLTIPS_CLASS, NULL,
            WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON,  // Use the same styles as in createAndShowWindows
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            _hSelf, NULL, hInstance, NULL);

        if (!_hUseListButtonTooltip)
        {
            // Handle error if tooltip creation fails
            return;
        }

        // Activate the tooltip
        SendMessage(_hUseListButtonTooltip, TTM_ACTIVATE, TRUE, 0);
    }

    // Prepare the TOOLINFO structure
    TOOLINFO ti = { 0 };
    ti.cbSize = sizeof(TOOLINFO);
    ti.hwnd = _hSelf;  // Parent window handle
    ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    ti.uId = (UINT_PTR)GetDlgItem(_hSelf, IDC_USE_LIST_BUTTON);  // Associate with the specific control

    // Get tooltip text based on the list state
    LPCWSTR tooltipText = useListEnabled ? LM.getLPCW(L"tooltip_disable_list") : LM.getLPCW(L"tooltip_enable_list");
    ti.lpszText = const_cast<LPWSTR>(tooltipText);  // Assign the tooltip text

    // If it's an update, delete the old tooltip first
    if (isUpdate)
    {
        SendMessage(_hUseListButtonTooltip, TTM_DELTOOL, 0, (LPARAM)&ti);
    }

    // Add or update the tooltip
    SendMessage(_hUseListButtonTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
}

#pragma endregion


#pragma region Undo stack

void MultiReplace::addItemsToReplaceList(const std::vector<ReplaceItemData>& items, size_t insertPosition = std::numeric_limits<size_t>::max()) {
    // Determine the insert position
    if (insertPosition > replaceListData.size()) {
        insertPosition = replaceListData.size();
    }

    size_t startIndex = insertPosition;
    size_t endIndex = startIndex + items.size() - 1;

    // Insert items at the specified position
    replaceListData.insert(replaceListData.begin() + insertPosition, items.begin(), items.end());

    // Update the ListView
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    // Create undo and redo actions
    UndoRedoAction action;

    // Undo action: Remove the added items
    action.undoAction = [this, startIndex, endIndex]() {
        // Remove items from the replace list
        replaceListData.erase(replaceListData.begin() + startIndex, replaceListData.begin() + endIndex + 1);

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Scroll to the position where items were removed
        scrollToIndices(startIndex, startIndex);
        };

    // Redo action: Re-insert the items
    std::vector<ReplaceItemData> itemsToRedo(items); // Copy of items to redo
    action.redoAction = [this, itemsToRedo, startIndex, endIndex]() {
        replaceListData.insert(replaceListData.begin() + startIndex, itemsToRedo.begin(), itemsToRedo.end());

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the re-inserted items
        for (size_t i = startIndex; i <= endIndex; ++i) {
            ListView_SetItemState(_replaceListView, static_cast<int>(i), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the inserted items
        scrollToIndices(startIndex, endIndex);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Add items");
}

void MultiReplace::removeItemsFromReplaceList(const std::vector<size_t>& indicesToRemove) {
    // Sort indices in descending order to handle shifting issues during removal
    std::vector<size_t> sortedIndices = indicesToRemove;
    std::sort(sortedIndices.rbegin(), sortedIndices.rend());

    // Store removed items along with their indices for undo purposes
    std::vector<std::pair<size_t, ReplaceItemData>> removedItemsWithIndices;
    for (size_t index : sortedIndices) {
        if (index < replaceListData.size()) {
            removedItemsWithIndices.emplace_back(index, replaceListData[index]);
            replaceListData.erase(replaceListData.begin() + index);
        }
    }

    // Update the ListView to reflect changes
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    // Create undo and redo actions
    UndoRedoAction action;

    // Undo action: Re-insert the removed items at their original positions
    action.undoAction = [this, removedItemsWithIndices]() {
        for (auto it = removedItemsWithIndices.rbegin(); it != removedItemsWithIndices.rend(); ++it) {
            const auto& [index, item] = *it;
            if (index <= replaceListData.size()) {
                replaceListData.insert(replaceListData.begin() + index, item);
            }
        }

        // Update the ListView to include restored items
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select restored rows
        for (const auto& [index, _] : removedItemsWithIndices) {
            if (index < replaceListData.size()) {
                ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
            }
        }

        // Calculate visible range and ensure restored rows are visible
        std::vector<size_t> indices;
        for (const auto& [index, _] : removedItemsWithIndices) {
            indices.push_back(index);
        }
        size_t firstIndex = *std::min_element(indices.begin(), indices.end());
        size_t lastIndex = *std::max_element(indices.begin(), indices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    // Redo action: Remove the same items again
    action.redoAction = [this, removedItemsWithIndices]() {
        // Sort indices in descending order to avoid shifting issues
        std::vector<size_t> sortedIndices;
        for (const auto& [index, _] : removedItemsWithIndices) {
            sortedIndices.push_back(index);
        }
        std::sort(sortedIndices.rbegin(), sortedIndices.rend());

        for (size_t index : sortedIndices) {
            if (index < replaceListData.size()) {
                replaceListData.erase(replaceListData.begin() + index);
            }
        }

        // Update the ListView to reflect removal
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Scroll to the position where deletion occurred
        if (!sortedIndices.empty()) {
            size_t minIndex = sortedIndices.back(); // Since sortedIndices is descending
            scrollToIndices(minIndex, minIndex);
        }
        };

    // Push the action to the undoStack
    URM.push(action.undoAction, action.redoAction, L"Remove items");
}

void MultiReplace::modifyItemInReplaceList(size_t index, const ReplaceItemData& newData) {
    // Store the original data
    ReplaceItemData originalData = replaceListData[index];

    // Modify the item
    replaceListData[index] = newData;

    // Update the ListView item
    updateListViewItem(index);

    // Create Undo/Redo actions
    UndoRedoAction action;

    // Undo action: Restore the original data
    action.undoAction = [this, index, originalData]() {
        replaceListData[index] = originalData;
        updateListViewItem(index);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select and ensure visibility of the modified item
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
        scrollToIndices(index, index);

        // Set focus to the ListView
        SetFocus(_replaceListView);
        };

    // Redo action: Apply the new data again
    action.redoAction = [this, index, newData]() {
        replaceListData[index] = newData;
        updateListViewItem(index);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select and ensure visibility of the modified item
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
        scrollToIndices(index, index);

        // Set focus to the ListView
        SetFocus(_replaceListView);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Modify item");
}

bool MultiReplace::moveItemsInReplaceList(std::vector<size_t>& indices, Direction direction) {
    if (indices.empty()) {
        return false; // No items to move
    }

    // Check the bounds
    if ((direction == Direction::Up && indices.front() == 0) ||
        (direction == Direction::Down && indices.back() == replaceListData.size() - 1)) {
        return false; // Out of bounds, do nothing
    }

    // Store pre-move indices for undo
    std::vector<size_t> preMoveIndices = indices;

    // Adjust indices for the move
    if (direction == Direction::Up) {
        for (size_t& idx : indices) {
            idx -= 1; // Adjust indices upwards
        }
    }
    else { // Direction::Down
        for (size_t& idx : indices) {
            idx += 1; // Adjust indices downwards
        }
    }

    // Perform the move
    for (size_t i = 0; i < preMoveIndices.size(); ++i) {
        std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
    }

    // Update the ListView
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    // Create Undo/Redo actions
    UndoRedoAction action;

    action.undoAction = [this, preMoveIndices, indices]() {
        // Swap back the moved items
        for (size_t i = 0; i < preMoveIndices.size(); ++i) {
            std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
        }

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the original positions
        for (size_t idx : preMoveIndices) {
            ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the moved items
        size_t firstIndex = *std::min_element(preMoveIndices.begin(), preMoveIndices.end());
        size_t lastIndex = *std::max_element(preMoveIndices.begin(), preMoveIndices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    action.redoAction = [this, preMoveIndices, indices]() {
        // Swap items to their new positions again
        for (size_t i = 0; i < preMoveIndices.size(); ++i) {
            std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
        }

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the moved positions
        for (size_t idx : indices) {
            ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the moved items
        size_t firstIndex = *std::min_element(indices.begin(), indices.end());
        size_t lastIndex = *std::max_element(indices.begin(), indices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Move items");

    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select the moved rows in their new positions
    for (size_t idx : indices) {
        ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
    }

    // Ensure visibility of the moved items
    size_t firstIndex = *std::min_element(indices.begin(), indices.end());
    size_t lastIndex = *std::max_element(indices.begin(), indices.end());
    scrollToIndices(firstIndex, lastIndex);

    return true;
}

void MultiReplace::sortItemsInReplaceList(const std::vector<size_t>& originalOrder, const std::vector<size_t>& newOrder, const std::map<int, SortDirection>& previousColumnSortOrder, int columnID, SortDirection direction) {
    UndoRedoAction action;

    // Undo action: Restore original order and sort state
    action.undoAction = [this, originalOrder, previousColumnSortOrder]() {
        // Restore original order
        std::unordered_map<size_t, ReplaceItemData> idToItemMap;
        for (const auto& item : replaceListData) {
            idToItemMap[item.id] = item;
        }

        replaceListData.clear();
        for (size_t id : originalOrder) {
            replaceListData.push_back(idToItemMap[id]);
        }

        // Restore previous sort order
        columnSortOrder = previousColumnSortOrder;

        // Update UI
        updateHeaderSortDirection();
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);
        };

    // Redo action: Restore sorted order and current sort state
    action.redoAction = [this, newOrder, columnID, direction]() {
        // Restore sorted order
        std::unordered_map<size_t, ReplaceItemData> idToItemMap;
        for (const auto& item : replaceListData) {
            idToItemMap[item.id] = item;
        }

        replaceListData.clear();
        for (size_t id : newOrder) {
            replaceListData.push_back(idToItemMap[id]);
        }

        // Update sort state
        columnSortOrder.clear();
        columnSortOrder[columnID] = direction;

        // Update UI
        updateHeaderSortDirection();
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);
        };

    // Push the action onto the undo stack
    URM.push(action.undoAction, action.redoAction, L"Sort items");
}

void MultiReplace::scrollToIndices(size_t firstIndex, size_t lastIndex) {
    if (firstIndex > lastIndex) {
        std::swap(firstIndex, lastIndex);
    }
    if (lastIndex >= replaceListData.size()) {
        lastIndex = replaceListData.empty() ? 0 : replaceListData.size() - 1;
    }

    // Get the client area of the ListView
    RECT rcClient;
    GetClientRect(_replaceListView, &rcClient);

    // Calculate the height of a single item in the ListView
    int itemHeight = ListView_GetItemSpacing(_replaceListView, TRUE) >> 16;

    // Calculate the number of fully visible items in the client area
    int visibleItemCount = rcClient.bottom / itemHeight;

    // Calculate the middle index of the range
    size_t middleIndex = firstIndex + (lastIndex - firstIndex) / 2;

    // Determine the desired top index so that the middle of the range is centered
    int desiredTopIndex = static_cast<int>(middleIndex) - (visibleItemCount / 2);

    // Ensure the desired top index is within valid bounds
    if (desiredTopIndex < 0) {
        desiredTopIndex = 0;
    }
    else if (desiredTopIndex + visibleItemCount > static_cast<int>(replaceListData.size())) {
        desiredTopIndex = static_cast<int>(replaceListData.size()) - visibleItemCount;
        if (desiredTopIndex < 0) {
            desiredTopIndex = 0;
        }
    }

    // Get the current top index of the ListView
    int currentTopIndex = ListView_GetTopIndex(_replaceListView);

    // Calculate the scroll offset in items
    int scrollItems = desiredTopIndex - currentTopIndex;

    // Scroll the ListView if necessary
    if (scrollItems != 0) {
        int scrollPixels = scrollItems * itemHeight;
        ListView_Scroll(_replaceListView, 0, scrollPixels);
    }
}

#pragma endregion


#pragma region ListView

HWND MultiReplace::CreateHeaderTooltip(HWND hwndParent)
{
    HWND hwndTT = CreateWindowEx(
        WS_EX_TOPMOST,
        TOOLTIPS_CLASS,
        NULL,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        hwndParent,
        NULL,
        GetModuleHandle(NULL),
        NULL
    );

    if (hwndTT)
    {
        SendMessage(hwndTT, TTM_ACTIVATE, TRUE, 0);
    }

    return hwndTT;
}

void MultiReplace::AddHeaderTooltip(HWND hwndTT, HWND hwndHeader, int columnIndex, LPCTSTR pszText) {
    RECT rect;
    Header_GetItemRect(hwndHeader, columnIndex, &rect);

    TOOLINFO ti = { 0 };
    ti.cbSize = sizeof(TOOLINFO);
    ti.uFlags = TTF_SUBCLASS;
    ti.hwnd = hwndHeader;
    ti.hinst = GetModuleHandle(NULL);
    ti.uId = columnIndex;
    ti.lpszText = const_cast<LPWSTR>(pszText);
    ti.rect = rect;

    SendMessage(hwndTT, TTM_DELTOOL, 0, (LPARAM)&ti);
    SendMessage(hwndTT, TTM_ADDTOOL, 0, (LPARAM)&ti);
}

void MultiReplace::createListViewColumns() {
    LVCOLUMN lvc;
    ZeroMemory(&lvc, sizeof(lvc));

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

    // Retrieve column widths
    findCountColumnWidth = getColumnWidth(ColumnID::FIND_COUNT);
    replaceCountColumnWidth = getColumnWidth(ColumnID::REPLACE_COUNT);
    findColumnWidth = getColumnWidth(ColumnID::FIND_TEXT);
    replaceColumnWidth = getColumnWidth(ColumnID::REPLACE_TEXT);
    commentsColumnWidth = getColumnWidth(ColumnID::COMMENTS);
    deleteButtonColumnWidth = getColumnWidth(ColumnID::DELETE_BUTTON);

    // Delete all existing columns
    int columnCount = Header_GetItemCount(ListView_GetHeader(_replaceListView));
    for (int i = columnCount - 1; i >= 0; --i) {
        ListView_DeleteColumn(_replaceListView, i);
    }
    columnIndices.clear();  // Clear previous indices after deleting

    const ControlInfo& listCtrlInfo = ctrlMap[IDC_REPLACE_LIST];


    // Define the ResizableColWidths struct to pass to calcDynamicColWidth
    ResizableColWidths widths = {
        _replaceListView,
        listCtrlInfo.cx,
        (isFindCountVisible) ? findCountColumnWidth : 0,
        (isReplaceCountVisible) ? replaceCountColumnWidth : 0,
        findColumnWidth,
        replaceColumnWidth,
        (isCommentsColumnVisible) ? commentsColumnWidth : 0,
        (isDeleteButtonVisible) ? deleteButtonColumnWidth : 0,
        GetSystemMetrics(SM_CXVSCROLL)
    };


    // Calculate dynamic width for Find, Replace, and Comments columns
    int perColumnWidth = calcDynamicColWidth(widths);

    int currentIndex = 0;  // Start tracking columns

    // Column 1: Find Count
    if (isFindCountVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getLPW(L"header_find_count");
        lvc.cx = findCountColumnWidth;
        lvc.fmt = LVCFMT_RIGHT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::FIND_COUNT] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::FIND_COUNT] = -1;  // Mark as not visible
    }

    // Column 2: Replace Count
    if (isReplaceCountVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getLPW(L"header_replace_count");
        lvc.cx = replaceCountColumnWidth;
        lvc.fmt = LVCFMT_RIGHT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::REPLACE_COUNT] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::REPLACE_COUNT] = -1;  // Mark as not visible
    }

    // Column 3: Selection Checkbox
    lvc.iSubItem = currentIndex;
    lvc.pszText = L"\u2610";
    lvc.cx = getColumnWidth(ColumnID::SELECTION);
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::SELECTION] = currentIndex;
    ++currentIndex;

    // Column 4: Find Text (dynamic width)
    lvc.iSubItem = currentIndex;
    lvc.pszText = LM.getLPW(L"header_find");
    lvc.cx = (findColumnLockedEnabled ? findColumnWidth : perColumnWidth);
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::FIND_TEXT] = currentIndex;
    ++currentIndex;

    // Column 5: Replace Text (dynamic width)
    lvc.iSubItem = currentIndex;
    lvc.pszText = LM.getLPW(L"header_replace");
    lvc.cx = (replaceColumnLockedEnabled ? replaceColumnWidth : perColumnWidth);
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::REPLACE_TEXT] = currentIndex;
    ++currentIndex;

    // Columns 6 to 10: Options (fixed width)
    const std::wstring options[] = {
        L"header_whole_word",
        L"header_match_case",
        L"header_use_variables",
        L"header_extended",
        L"header_regex"
    };
    for (int i = 0; i < 5; ++i) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getLPW(options[i]);
        lvc.cx = checkMarkWidth_scaled;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[static_cast<ColumnID>(static_cast<int>(ColumnID::WHOLE_WORD) + i)] = currentIndex;
        ++currentIndex;
    }

    // Column 11: Comments (dynamic width)
    if (isCommentsColumnVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getLPW(L"header_comments");
        lvc.cx = (commentsColumnLockedEnabled ? commentsColumnWidth : perColumnWidth);
        lvc.fmt = LVCFMT_LEFT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::COMMENTS] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::COMMENTS] = -1;  // Mark as not visible
    }

    // Column 12: Delete Button (fixed width)
    if (isDeleteButtonVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = L"";
        lvc.cx = crossWidth_scaled;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::DELETE_BUTTON] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::DELETE_BUTTON] = -1;
    }

    updateHeaderSortDirection();
    updateHeaderSelection();
    updateListViewTooltips();
}

void MultiReplace::insertReplaceListItem(const ReplaceItemData& itemData) {
    int useVariables = IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED ? 1 : 0;

    // Return early if findText is empty and "Use Variables" is not enabled
    if (itemData.findText.empty() && useVariables == 0) {
        showStatusMessage(LM.get(L"status_no_find_string"), MessageStatus::Error);
        return;
    }

    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    // Check if itemData is already in the vector
    bool isDuplicate = false;
    for (const auto& existingItem : replaceListData) {
        if (existingItem == itemData) {
            isDuplicate = true;
            break;
        }
    }

    std::vector<ReplaceItemData> itemsToAdd = { itemData };
    addItemsToReplaceList(itemsToAdd);

    // Show a status message indicating the value added to the list
    std::wstring message;
    if (isDuplicate) {
        message = LM.get(L"status_duplicate_entry") + itemData.findText;
    }
    else {
        message = LM.get(L"status_value_added");
    }
    showStatusMessage(message, MessageStatus::Success);

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Select newly added line
    size_t newItemIndex = replaceListData.size() - 1;
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED); // Delete previous selection
    ListView_SetItemState(_replaceListView, static_cast<int>(newItemIndex), LVIS_SELECTED, LVIS_SELECTED);
    ListView_EnsureVisible(_replaceListView, static_cast<int>(newItemIndex), FALSE);

    // Update Header if there might be any changes
    updateHeaderSelection();
}

int MultiReplace::getCharacterWidth(int elementID, const wchar_t* character) {
    // Get the HWND of the element by its ID
    HWND hwndElement = GetDlgItem(_hSelf, elementID);

    // Get the font used by the element
    HFONT hFont = (HFONT)SendMessage(hwndElement, WM_GETFONT, 0, 0);

    // Get the device context for measuring text
    HDC hdc = GetDC(hwndElement);
    SelectObject(hdc, hFont);  // Use the font of the given element

    // Measure the width of the character
    SIZE size;
    GetTextExtentPoint32W(hdc, character, 1, &size);

    // Release the device context
    ReleaseDC(hwndElement, hdc);

    // Return the width of the character
    return size.cx;
}

std::wstring getColumnIDText(ColumnID columnID) {
    switch (columnID) {
    case ColumnID::INVALID:          return L"INVALID";
    case ColumnID::FIND_COUNT:       return L"FIND_COUNT";
    case ColumnID::REPLACE_COUNT:    return L"REPLACE_COUNT";
    case ColumnID::SELECTION:        return L"SELECTION";
    case ColumnID::FIND_TEXT:        return L"FIND_TEXT";
    case ColumnID::REPLACE_TEXT:     return L"REPLACE_TEXT";
    case ColumnID::WHOLE_WORD:       return L"WHOLE_WORD";
    case ColumnID::MATCH_CASE:       return L"MATCH_CASE";
    case ColumnID::USE_VARIABLES:    return L"USE_VARIABLES";
    case ColumnID::EXTENDED:         return L"EXTENDED";
    case ColumnID::REGEX:            return L"REGEX";
    case ColumnID::COMMENTS:         return L"COMMENTS";
    case ColumnID::DELETE_BUTTON:    return L"DELETE_BUTTON";
    default:                         return L"UNKNOWN";
    }
}

int MultiReplace::getColumnWidth(ColumnID columnID) {


    int width = 0;

    switch (columnID) {
    case ColumnID::DELETE_BUTTON:
        width = crossWidth_scaled;
        break;
    case ColumnID::SELECTION:
        width = boxWidth_scaled;
        break;
    case ColumnID::WHOLE_WORD:
    case ColumnID::MATCH_CASE:
    case ColumnID::USE_VARIABLES:
    case ColumnID::EXTENDED:
    case ColumnID::REGEX:
        width = checkMarkWidth_scaled;
        break;
    default:
        // For dynamic columns (Find, Replace, Comments, Find Count, Replace Count)
        auto it = columnIndices.find(columnID);
        if (it != columnIndices.end() && it->second != -1) {
            // Column exists, get width from ListView
            int index = it->second;
            width = ListView_GetColumnWidth(_replaceListView, index);
        }
        else {
            // Column does not exist, use default column width
            switch (columnID) {
            case ColumnID::FIND_COUNT:
                width = findCountColumnWidth;
                break;
            case ColumnID::REPLACE_COUNT:
                width = replaceCountColumnWidth;
                break;
            case ColumnID::FIND_TEXT:
                width = findColumnWidth;
                break;
            case ColumnID::REPLACE_TEXT:
                width = replaceColumnWidth;
                break;
            case ColumnID::COMMENTS:
                width = commentsColumnWidth;
                break;
            default:
                width = MIN_GENERAL_WIDTH_scaled; // Fallback to minimum width
                break;
            }
        }

        // Ensure the width is at least MIN_GENERAL_WIDTH_scaled
        width = std::max(width, MIN_GENERAL_WIDTH_scaled);

        break;
    }

    return width;
}

int MultiReplace::calcDynamicColWidth(const ResizableColWidths& widths) {
    // Calculate the total width of fixed columns
    int totalWidthFixedColumns = boxWidth_scaled + (checkMarkWidth_scaled * 5)
        + widths.findCountWidth
        + widths.replaceCountWidth
        + widths.deleteWidth;

    // Calculate the remaining width by subtracting fixed widths and visible static column widths
    int totalRemainingWidth = widths.listViewWidth
        - widths.margin  // Adjust for the vertical scrollbar margin
        - totalWidthFixedColumns
        - (findColumnLockedEnabled ? widths.findWidth : 0)
        - (replaceColumnLockedEnabled ? widths.replaceWidth : 0)
        - (commentsColumnLockedEnabled ? widths.commentsWidth : 0);

    // Calculate the number of dynamic columns
    int dynamicColumnCount = (!findColumnLockedEnabled)
        + (!replaceColumnLockedEnabled)
        + (!commentsColumnLockedEnabled && isCommentsColumnVisible);

    // Ensure at least one dynamic column is present for safety
    dynamicColumnCount = std::max(dynamicColumnCount, 1);

    // Calculate the width for each dynamic column
    int perColumnWidth = std::max(totalRemainingWidth / dynamicColumnCount, MIN_GENERAL_WIDTH_scaled);

    return perColumnWidth; // Return the calculated width for each dynamic column
}

void MultiReplace::updateListViewAndColumns() {
    // Retrieve control information
    const ControlInfo& listCtrlInfo = ctrlMap[IDC_REPLACE_LIST];
    HWND listView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    // Retrieve column widths
    findCountColumnWidth = getColumnWidth(ColumnID::FIND_COUNT);
    replaceCountColumnWidth = getColumnWidth(ColumnID::REPLACE_COUNT);
    commentsColumnWidth = getColumnWidth(ColumnID::COMMENTS);
    deleteButtonColumnWidth = getColumnWidth(ColumnID::DELETE_BUTTON);
    findColumnWidth = getColumnWidth(ColumnID::FIND_TEXT);
    replaceColumnWidth = getColumnWidth(ColumnID::REPLACE_TEXT);


    // Prepare the ResizableColWidths struct for dynamic width calculations
    ResizableColWidths widths = {
        listView,
        listCtrlInfo.cx,
        (isFindCountVisible) ? findCountColumnWidth : 0,
        (isReplaceCountVisible) ? replaceCountColumnWidth : 0,
        findColumnWidth,
        replaceColumnWidth,
        (isCommentsColumnVisible) ? commentsColumnWidth : 0,
        (isDeleteButtonVisible) ? deleteButtonColumnWidth : 0,
        GetSystemMetrics(SM_CXVSCROLL)
    };

    // Calculate dynamic column width
    int perColumnWidth = calcDynamicColWidth(widths);

    // Disable redraw to prevent flickering during updates
    SendMessage(listView, WM_SETREDRAW, FALSE, 0);

    // Update widths of dynamic columns
    if (columnIndices[ColumnID::FIND_TEXT] != -1) {
        int width = findColumnLockedEnabled ? findColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::FIND_TEXT], width);
    }
    if (columnIndices[ColumnID::REPLACE_TEXT] != -1) {
        int width = replaceColumnLockedEnabled ? replaceColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::REPLACE_TEXT], width);
    }
    if (columnIndices[ColumnID::COMMENTS] != -1) {
        int width = commentsColumnLockedEnabled ? commentsColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::COMMENTS], width);
    }

    if (columnIndices[ColumnID::FIND_COUNT] != -1) {
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::FIND_COUNT], findCountColumnWidth);
    }
    if (columnIndices[ColumnID::REPLACE_COUNT] != -1) {
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::REPLACE_COUNT], replaceCountColumnWidth);
    }

    // Resize the ListView control to the updated width and height
    updateListViewFrame();

    // Re-enable redraw after updates
    SendMessage(listView, WM_SETREDRAW, TRUE, 0);
}

void MultiReplace::updateListViewItem(size_t index) {
    if (index >= replaceListData.size()) return;

    const ReplaceItemData& item = replaceListData[index];

    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::FIND_TEXT], const_cast<LPWSTR>(item.findText.c_str()));
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::REPLACE_TEXT], const_cast<LPWSTR>(item.replaceText.c_str()));
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::COMMENTS], const_cast<LPWSTR>(item.comments.c_str()));

    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::WHOLE_WORD], item.wholeWord ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::MATCH_CASE], item.matchCase ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::USE_VARIABLES], item.useVariables ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::EXTENDED], item.extended ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::REGEX], item.regex ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::SELECTION], item.isEnabled ? L"\u25A0" : L"\u2610");

    ListView_RedrawItems(_replaceListView, static_cast<int>(index), static_cast<int>(index));
}

void MultiReplace::updateListViewTooltips() {

    if (!tooltipsEnabled) {
        return;
    }

    HWND hwndHeader = ListView_GetHeader(_replaceListView);
    if (!hwndHeader)
        return;

    // Destroy the old tooltip window to ensure all previous tooltips are removed
    if (_hHeaderTooltip) {
        DestroyWindow(_hHeaderTooltip);
        _hHeaderTooltip = nullptr;
    }

    // Create a new tooltip window
    _hHeaderTooltip = CreateHeaderTooltip(hwndHeader);

    // Re-add tooltips for columns 6 to 10
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::WHOLE_WORD], LM.getLPW(L"tooltip_header_whole_word"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::MATCH_CASE], LM.getLPW(L"tooltip_header_match_case"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::USE_VARIABLES], LM.getLPW(L"tooltip_header_use_variables"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::EXTENDED], LM.getLPW(L"tooltip_header_extended"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::REGEX], LM.getLPW(L"tooltip_header_regex"));
}

void MultiReplace::handleCopyBack(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }

    // Copy the data from the selected item back to the source interfaces
    ReplaceItemData& itemData = replaceListData[pnmia->iItem];

    // Update the controls directly
    SetWindowTextW(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText.c_str());
    SetWindowTextW(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText.c_str());
    SendMessageW(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, itemData.wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, itemData.matchCase ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, itemData.useVariables ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_NORMAL_RADIO), BM_SETCHECK, (!itemData.regex && !itemData.extended) ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_EXTENDED_RADIO), BM_SETCHECK, itemData.extended ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_SETCHECK, itemData.regex ? BST_CHECKED : BST_UNCHECKED, 0);
}

void MultiReplace::shiftListItem(const Direction& direction) {
    std::vector<size_t> selectedIndices;
    int i = -1;

    // Collect selected indices
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(i);
    }

    if (selectedIndices.empty()) {
        showStatusMessage(LM.get(L"status_no_rows_selected_to_shift"), MessageStatus::Error);
        return;
    }

    // Pass the selected indices to moveItemsInReplaceList
    if (!moveItemsInReplaceList(selectedIndices, direction)) {
        return;
    }

    // Deselect all items
    for (int j = 0; j < ListView_GetItemCount(_replaceListView); ++j) {
        ListView_SetItemState(_replaceListView, j, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Re-select the shifted items
    for (size_t index : selectedIndices) {
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Show status message when rows are successfully shifted
    showStatusMessage(LM.get(L"status_rows_shifted", { std::to_wstring(selectedIndices.size()) }), MessageStatus::Success);

    // Show Statistics
    showListFilePath();
}

void MultiReplace::handleDeletion(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }
    // Store the index of the item to be deleted
    size_t itemIndex = static_cast<size_t>(pnmia->iItem);

    // Prepare a vector containing the index of the item to remove
    std::vector<size_t> indicesToRemove = { itemIndex };

    // Delegate the deletion and undo/redo logic to removeItemsFromReplaceList
    removeItemsFromReplaceList(indicesToRemove);

    // Update Header if there might be any changes
    updateHeaderSelection();

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(LM.get(L"status_one_line_deleted"), MessageStatus::Success);
}

void MultiReplace::deleteSelectedLines() {
    // Collect selected indices
    std::vector<size_t> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(static_cast<size_t>(i));
    }

    if (selectedIndices.empty()) {
        showStatusMessage(LM.get(L"status_no_rows_selected_to_delete"), MessageStatus::Error);
        return;
    }

    // Ensure indices are valid before calling removeItemsFromReplaceList
    for (size_t index : selectedIndices) {
        if (index >= replaceListData.size()) {
            showStatusMessage(LM.get(L"status_invalid_indices"), MessageStatus::Error);
            return;
        }
    }

    // Call the removeItemsFromReplaceList function
    removeItemsFromReplaceList(selectedIndices);

    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select the next available line
    size_t nextIndexToSelect = selectedIndices.back() < replaceListData.size()
        ? selectedIndices.back()
        : replaceListData.size() - 1;
    if (nextIndexToSelect < replaceListData.size()) {
        ListView_SetItemState(_replaceListView, static_cast<int>(nextIndexToSelect), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Update Header if there might be any changes
    updateHeaderSelection();

    // Show status message
    showStatusMessage(LM.get(L"status_lines_deleted", { std::to_wstring(selectedIndices.size()) }), MessageStatus::Success);
}

void MultiReplace::sortReplaceListData(int columnID) {
    // Check if the column is sortable
    if (columnID != ColumnID::FIND_COUNT &&
        columnID != ColumnID::REPLACE_COUNT &&
        columnID != ColumnID::FIND_TEXT &&
        columnID != ColumnID::REPLACE_TEXT &&
        columnID != ColumnID::COMMENTS) {
        return;
    }

    // Preserve the selection
    auto selectedIDs = getSelectedRows();

    // Ensure IDs are assigned
    for (auto& item : replaceListData) {
        if (item.id == 0) {
            item.id = generateUniqueID();
        }
    }

    // Capture the original order and previous sort state
    std::vector<size_t> originalOrder;
    for (const auto& item : replaceListData) {
        originalOrder.push_back(item.id);
    }
    auto previousColumnSortOrder = columnSortOrder;

    // Determine the new sort direction
    SortDirection direction = SortDirection::Ascending;
    auto it = columnSortOrder.find(columnID);
    if (it != columnSortOrder.end() && it->second == SortDirection::Ascending) {
        direction = SortDirection::Descending;
    }

    // Update the column sort order
    columnSortOrder.clear();
    columnSortOrder[columnID] = direction;

    // return -1 for empty or non-numeric strings
    auto safeToInt = [](const std::wstring& s) -> int {
        try { return s.empty() ? -1 : std::stoi(s); }
        catch (...) { return -1; }
    };

    // Perform the sorting
    std::sort(replaceListData.begin(), replaceListData.end(),
        [this, columnID, direction, safeToInt]
    (const ReplaceItemData& a, const ReplaceItemData& b) -> bool
        {
            switch (columnID)
            {
            case ColumnID::FIND_COUNT: {
                int numA = safeToInt(a.findCount);
                int numB = safeToInt(b.findCount);
                return direction == SortDirection::Ascending ? numA < numB : numA > numB;
            }
            case ColumnID::REPLACE_COUNT: {
                int numA = safeToInt(a.replaceCount);
                int numB = safeToInt(b.replaceCount);
                return direction == SortDirection::Ascending ? numA < numB : numA > numB;
            }
            case ColumnID::FIND_TEXT:
                return direction == SortDirection::Ascending ? a.findText < b.findText : a.findText > b.findText;
            case ColumnID::REPLACE_TEXT:
                return direction == SortDirection::Ascending ? a.replaceText < b.replaceText : a.replaceText > b.replaceText;
            case ColumnID::COMMENTS:
                return direction == SortDirection::Ascending ? a.comments < b.comments : a.comments > b.comments;
            default:
                return false;
            }
        });

    // Capture the new order after sorting
    std::vector<size_t> newOrder;
    for (const auto& item : replaceListData) {
        newOrder.push_back(item.id);
    }

    // Update the UI and restore the selection
    updateHeaderSortDirection();
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
    selectRows(selectedIDs);

    // Delegate undo/redo creation to the new function
    sortItemsInReplaceList(originalOrder, newOrder, previousColumnSortOrder, columnID, direction);
}

std::vector<size_t> MultiReplace::getSelectedRows() {
    std::vector<size_t> selectedIDs;
    int index = -1; // Use int to properly handle -1 case
    while ((index = ListView_GetNextItem(_replaceListView, index, LVNI_SELECTED)) != -1) {
        if (index >= 0 && static_cast<size_t>(index) < replaceListData.size()) {
            selectedIDs.push_back(replaceListData[index].id);
        }
    }
    return selectedIDs;
}

size_t MultiReplace::generateUniqueID() {
    static size_t currentID = 0;
    return ++currentID;
}

void MultiReplace::selectRows(const std::vector<size_t>& selectedIDs) {
    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Reselect rows based on IDs
    for (size_t i = 0; i < replaceListData.size(); ++i) {
        if (std::find(selectedIDs.begin(), selectedIDs.end(), replaceListData[i].id) != selectedIDs.end()) {
            ListView_SetItemState(_replaceListView, i, LVIS_SELECTED, LVIS_SELECTED);
        }
    }
}

void MultiReplace::handleCopyToListButton() {
    ReplaceItemData itemData;

    itemData.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
    itemData.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

    itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
    itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
    itemData.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
    itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
    itemData.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);

    insertReplaceListItem(itemData);

    // Add the entered text to the combo box history
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText);
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText);

    // Enable the ListView accordingly
    useListEnabled = true;
    updateUseListState(true);
    adjustWindowSize();
}

void MultiReplace::resetCountColumns() {
    // Reset the find and replace count columns in the list data
    for (auto& itemData : replaceListData) {
        itemData.findCount = L"";
        itemData.replaceCount = L"";
    }

    // Update the list view to immediately reflect the changes
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

void MultiReplace::updateCountColumns(const size_t itemIndex, const int findCount, int replaceCount)
{
    ReplaceItemData& itemData = replaceListData[itemIndex];

    if (findCount == -2) {
        itemData.findCount.clear();
    }
    else if (findCount != -1) {
        itemData.findCount = std::to_wstring(findCount);
    }

    if (replaceCount != -1) {
        itemData.replaceCount = std::to_wstring(replaceCount);
    }
}

void MultiReplace::clearList() {
    // Check for unsaved changes before clearing the list
    if (checkForUnsavedChanges() == IDCANCEL) {
        return;
    }

    // Clear the replace list data vector
    replaceListData.clear();

    // Update the ListView to reflect the cleared list
    ListView_SetItemCountEx(_replaceListView, 0, LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    // Reset listFilePath to an empty string
    listFilePath.clear();

    // Call showListFilePath to update the UI with the cleared file path
    showListFilePath();

    // Set the original list hash to a default value when list is cleared
    originalListHash = 0;
}

std::size_t MultiReplace::computeListHash(const std::vector<ReplaceItemData>& list) {
    std::size_t combinedHash = 0;
    ReplaceItemDataHasher hasher;

    for (const auto& item : list) {
        combinedHash ^= hasher(item) + golden_ratio_constant + (combinedHash << 6) + (combinedHash >> 2);
    }

    return combinedHash;
}

void MultiReplace::refreshUIListView()
{
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

void MultiReplace::handleColumnVisibilityToggle(UINT menuId) {
    // Toggle the corresponding visibility flag
    switch (menuId) {
    case IDM_TOGGLE_FIND_COUNT:
        isFindCountVisible = !isFindCountVisible;
        break;
    case IDM_TOGGLE_REPLACE_COUNT:
        isReplaceCountVisible = !isReplaceCountVisible;
        break;
    case IDM_TOGGLE_COMMENTS:
        isCommentsColumnVisible = !isCommentsColumnVisible;
        break;
    case IDM_TOGGLE_DELETE:
        isDeleteButtonVisible = !isDeleteButtonVisible;
        break;
    default:
        return; // Unhandled menu ID
    }

    // Recreate the ListView columns to reflect the changes
    HWND listView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    createListViewColumns();

    // Refresh the ListView (if necessary)
    InvalidateRect(listView, NULL, TRUE);
}

ColumnID MultiReplace::getColumnIDFromIndex(int columnIndex) const {
    auto it = std::find_if(
        columnIndices.begin(),
        columnIndices.end(),
        [columnIndex](const auto& pair) { return pair.second == columnIndex; });

    return (it != columnIndices.end()) ? it->first : ColumnID::INVALID;
}

int MultiReplace::getColumnIndexFromID(ColumnID columnID) const {
    auto it = std::find_if(
        columnIndices.begin(),
        columnIndices.end(),
        [columnID](const auto& pair) { return pair.first == columnID; });

    return (it != columnIndices.end()) ? it->second : -1; // Return -1 if not found
}

void MultiReplace::onPathDisplayDoubleClick()
{
    if (listFilePath.empty())
        return;

    std::wstring cmdParam = L"/select,\"" + listFilePath + L"\"";
    ShellExecuteW(nullptr, L"open", L"explorer.exe", cmdParam.c_str(), nullptr, SW_SHOWNORMAL);
}

#pragma endregion


#pragma region Contextmenu Display Columns

void MultiReplace::showColumnVisibilityMenu(HWND hWnd, POINT pt) {
    // Create a popup menu
    HMENU hMenu = CreatePopupMenu();

    // Add menu items with checkmarks based on current visibility
    AppendMenu(hMenu, MF_STRING | (isFindCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_FIND_COUNT, LM.getLPCW(L"header_find_count"));
    AppendMenu(hMenu, MF_STRING | (isReplaceCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_REPLACE_COUNT, LM.getLPCW(L"header_replace_count"));
    AppendMenu(hMenu, MF_STRING | (isCommentsColumnVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_COMMENTS, LM.getLPCW(L"header_comments"));
    AppendMenu(hMenu, MF_STRING | (isDeleteButtonVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_DELETE, LM.getLPCW(L"header_delete_button"));

    // Display the menu at the specified location
    TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, hWnd, NULL);

    // Destroy the menu after use
    DestroyMenu(hMenu);
}

#pragma endregion


#pragma region Contextmenu List

void MultiReplace::toggleBooleanAt(int itemIndex, ColumnID columnID) {
    if (itemIndex < 0 || itemIndex >= static_cast<int>(replaceListData.size())) {
        return; // Early return for invalid item index
    }

    // Store the original data
    ReplaceItemData originalData = replaceListData[itemIndex];

    // Create a new data object to represent the modified state
    ReplaceItemData newData = originalData;

    // Toggle the boolean field based on the ColumnID
    switch (columnID) {
    case ColumnID::SELECTION:
        newData.isEnabled = !newData.isEnabled;
        break;
    case ColumnID::WHOLE_WORD:
        newData.wholeWord = !newData.wholeWord;
        break;
    case ColumnID::MATCH_CASE:
        newData.matchCase = !newData.matchCase;
        break;
    case ColumnID::USE_VARIABLES:
        newData.useVariables = !newData.useVariables;
        break;
    case ColumnID::EXTENDED:
        newData.extended = !newData.extended;
        break;
    case ColumnID::REGEX:
        newData.regex = !newData.regex;
        break;
    default:
        return; // Not a toggleable boolean column
    }

    // Use modifyItemInReplaceList to handle the change and Undo/Redo
    modifyItemInReplaceList(static_cast<size_t>(itemIndex), newData);

    if (columnID == ColumnID::SELECTION)
        updateHeaderSelection();   // Re-evaluate all/none/some enabled for the header
}

void MultiReplace::editTextAt(int itemIndex, ColumnID columnID) {
    // Determine column index and validate
    int column = getColumnIndexFromID(columnID);
    if (column == -1)
        return;

    // Suppress hover text and disable tooltips
    isHoverTextSuppressed = true;
    DWORD extendedStyle = ListView_GetExtendedListViewStyle(_replaceListView);
    extendedStyle &= ~LVS_EX_INFOTIP;
    ListView_SetExtendedListViewStyle(_replaceListView, extendedStyle);

    // Calculate X position of the column based on scrolling and widths
    int totalWidthBeforeColumn = 0;
    for (int i = 0; i < column; ++i) {
        totalWidthBeforeColumn += ListView_GetColumnWidth(_replaceListView, i);
    }
    int columnWidth = ListView_GetColumnWidth(_replaceListView, column);

    SCROLLINFO siHorz = { sizeof(siHorz), SIF_POS };
    GetScrollInfo(_replaceListView, SB_HORZ, &siHorz);
    int correctedX = totalWidthBeforeColumn - siHorz.nPos;

    // Retrieve Y position and height of the row
    RECT itemRect;
    ListView_GetItemRect(_replaceListView, itemIndex, &itemRect, LVIR_BOUNDS);
    int correctedY = itemRect.top;
    int editHeight = itemRect.bottom - itemRect.top;

    // Calculate scaled button dimensions
    int buttonWidth = sx(20);
    int buttonHeight = editHeight + 2;

    // Adjust edit control width to reserve space for the button
    int editWidth = columnWidth - buttonWidth;

    // Create multi-line edit control with vertical and horizontal scrollbars
    hwndEdit = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN,
        correctedX,
        correctedY,
        editWidth,
        editHeight,
        _replaceListView,
        NULL,
        (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE),
        NULL
    );

    // Create an expand/collapse toggle button next to the edit control
    hwndExpandBtn = CreateWindowEx(
        0,
        L"BUTTON",
        L"↓", // Indicator for expand/collapse
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        correctedX + editWidth,
        correctedY - 1, // Adjust for button alignment
        buttonWidth,
        buttonHeight,
        _replaceListView,
        (HMENU)ID_EDIT_EXPAND_BUTTON,
        (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE),
        NULL
    );

    // Set the _hBoldFont2 to the expand/collapse button
    if (_hBoldFont2) {
        SendMessage(hwndExpandBtn, WM_SETFONT, (WPARAM)_hBoldFont2, TRUE);
    }

    // Set the initial text of the edit control
    const std::wstring* src = nullptr;
    if (itemIndex >= 0 && itemIndex < static_cast<int>(replaceListData.size())) {
        switch (columnID) {
        case ColumnID::FIND_TEXT:    src = &replaceListData[itemIndex].findText;    break;
        case ColumnID::REPLACE_TEXT: src = &replaceListData[itemIndex].replaceText; break;
        case ColumnID::COMMENTS:     src = &replaceListData[itemIndex].comments;    break;
        default: break;
        }
    }
    SetWindowTextW(hwndEdit, (src ? src->c_str() : L""));

    // Get the ListView font and set it for the Edit control
    HFONT hListViewFont = (HFONT)SendMessage(_replaceListView, WM_GETFONT, 0, 0);
    if (hListViewFont) {
        SendMessage(hwndEdit, WM_SETFONT, (WPARAM)hListViewFont, TRUE);
    }

    // Focus the edit control and select all its text
    SetFocus(hwndEdit);
    SendMessage(hwndEdit, EM_SETSEL, 0, -1);

    // Subclass the edit control for custom keyboard handling
    SetWindowSubclass(hwndEdit, EditControlSubclassProc, 1, (DWORD_PTR)this);

    // Store the editing state
    _editingItemIndex = itemIndex;
    _editingColumnIndex = column;
    _editingColumnID = columnID;
    _editIsExpanded = false; // Initial state is collapsed
}

void MultiReplace::closeEditField(bool commitChanges)
{
    if (!hwndEdit) return;

    if (commitChanges &&
        _editingColumnID != ColumnID::INVALID &&
        _editingItemIndex >= 0 &&
        _editingItemIndex < static_cast<int>(replaceListData.size()))
    {
        // Read edited text dynamically
        int len = GetWindowTextLengthW(hwndEdit);
        std::wstring newText;
        if (len > 0) {
            newText.resize(static_cast<size_t>(len) + 1, L'\0');
            int written = GetWindowTextW(hwndEdit, &newText[0], len + 1);
            if (written < 0) written = 0;
            if (!newText.empty() && newText.back() == L'\0') newText.pop_back();
            newText.resize(static_cast<size_t>(written));
        }
        else {
            newText.clear();
        }

        // Prepare old/new data
        ReplaceItemData originalData = replaceListData[_editingItemIndex];
        ReplaceItemData newData = originalData;
        bool hasChanged = false;

        switch (_editingColumnID) {
        case ColumnID::FIND_TEXT:
            if (originalData.findText != newText) {
                newData.findText = std::move(newText);
                hasChanged = true;
            }
            break;

        case ColumnID::REPLACE_TEXT:
            if (originalData.replaceText != newText) {
                newData.replaceText = std::move(newText);
                hasChanged = true;
            }
            break;

        case ColumnID::COMMENTS:
            if (originalData.comments != newText) {
                newData.comments = std::move(newText);
                hasChanged = true;
            }
            break;

        default:
            break;
        }

        if (hasChanged) {
            modifyItemInReplaceList(_editingItemIndex, newData);
        }
    }

    // Tear down edit controls (unchanged)
    DestroyWindow(hwndEdit);
    hwndEdit = nullptr;
    if (hwndExpandBtn && IsWindow(hwndExpandBtn)) {
        DestroyWindow(hwndExpandBtn);
        hwndExpandBtn = nullptr;
    }
    _editIsExpanded = false;
    _editingItemIndex = -1;
    _editingColumnIndex = -1;
    _editingColumnID = ColumnID::INVALID;

    if (isHoverTextSuppressed) {
        isHoverTextSuppressed = false;
        if (isHoverTextEnabled) {
            DWORD es = ListView_GetExtendedListViewStyle(_replaceListView);
            es |= LVS_EX_INFOTIP;
            ListView_SetExtendedListViewStyle(_replaceListView, es);
        }
    }
}

LRESULT CALLBACK MultiReplace::EditControlSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    // Access the instance of MultiReplace
    MultiReplace* pThis = reinterpret_cast<MultiReplace*>(dwRefData);

    switch (msg) {
    case WM_KEYDOWN:
        if (wParam == VK_ESCAPE || wParam == VK_TAB) {
            pThis->closeEditField(true);
            RemoveWindowSubclass(hwnd, EditControlSubclassProc, uIdSubclass);
            return 0;
        }
        break;

    case WM_KILLFOCUS:
    {
        HWND newFocus = GetFocus();
        if (newFocus == hwndExpandBtn) {
            return 0;
        }

        pThis->closeEditField(true);
        RemoveWindowSubclass(hwnd, EditControlSubclassProc, uIdSubclass);
        return 0;
    }

    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK MultiReplace::ListViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    // Retrieve the 'this' pointer from the safe reference data, not from GWLP_USERDATA.
    MultiReplace* pThis = reinterpret_cast<MultiReplace*>(dwRefData);

    // Defensive check: If the pointer is somehow invalid, fall back to the default procedure.
    if (!pThis || !IsWindow(pThis->_hSelf)) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    // Handle WM_NCDESTROY to remove the subclass before the window is finally destroyed.
    if (msg == WM_NCDESTROY) {
        RemoveWindowSubclass(hwnd, &MultiReplace::ListViewSubclassProc, uIdSubclass);
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }
    switch (msg) {
    case WM_VSCROLL:
    case WM_MOUSEWHEEL:
    case WM_HSCROLL: {
        // Close the edit control if it exists and is visible
        if (pThis->hwndEdit && IsWindow(pThis->hwndEdit)) {
            DestroyWindow(pThis->hwndEdit);
            pThis->hwndEdit = NULL;
        }
        // Allow the default list view procedure to handle standard scrolling
        break;
    }
    case WM_NOTIFY: {
        NMHDR* pnmhdr = reinterpret_cast<NMHDR*>(lParam);
        // Check notifications from header
        if (pnmhdr->hwndFrom == ListView_GetHeader(hwnd)) {
            int code = static_cast<int>(static_cast<short>(pnmhdr->code));

            if (code == HDN_ITEMCHANGED) {
                pThis->updateListViewTooltips(); //Updating Positions of Column Tooltips
            }

            // Handle right-click (NM_RCLICK) on the header
            if (code == NM_RCLICK) {
                // Get the current cursor position
                POINT ptScreen;
                GetCursorPos(&ptScreen);

                // Show the column visibility menu at the cursor position
                pThis->showColumnVisibilityMenu(pThis->_hSelf, ptScreen);

                return TRUE; // Indicate that the message has been handled
            }

            if (code == HDN_DIVIDERDBLCLICK) {
                NMHEADER* phdn = reinterpret_cast<NMHEADER*>(lParam);

                // Identify the clicked column based on phdn->iItem
                int clickedColumn = phdn->iItem;

                // Lock or unlock the column based on the clicked divider
                if (clickedColumn == pThis->columnIndices[ColumnID::FIND_TEXT]) {
                    pThis->findColumnLockedEnabled = !pThis->findColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }
                else if (clickedColumn == pThis->columnIndices[ColumnID::REPLACE_TEXT]) {
                    pThis->replaceColumnLockedEnabled = !pThis->replaceColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }
                else if (clickedColumn == pThis->columnIndices[ColumnID::COMMENTS]) {
                    pThis->commentsColumnLockedEnabled = !pThis->commentsColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }

                return TRUE; // Indicate that the message has been handled
            }

            // These values are derived from the HDN_ITEMCHANGEDW and HDN_ITEMCHANGEDA constants:
            // HDN_ITEMCHANGEDW = 0U - 300U - 21 = -321
            // HDN_ITEMCHANGEDA = 0U - 300U - 1  = -301
            // The constants are not used directly to avoid unsigned arithmetic overflow warnings.
            if (code == (int(0) - 300 - 21) || code == (int(0) - 300 - 1)) {
                // If there is an active edit control, destroy it when the header is changed
                if (pThis->hwndEdit && IsWindow(pThis->hwndEdit)) {
                    DestroyWindow(pThis->hwndEdit);
                    pThis->hwndEdit = NULL;
                }

                // Set a timer to defer the tooltip update, ensuring the column resize is complete
                SetTimer(hwnd, 1, 100, NULL);  // Timer ID 1, 100ms delay

            }
        }
        break;
    }

    case WM_MOUSEMOVE: {

        if (!pThis->isHoverTextEnabled || pThis->isHoverTextSuppressed) {
            return CallWindowProc(pThis->originalListViewProc, hwnd, msg, wParam, lParam);
        }

        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);

        LVHITTESTINFO hitTestInfo = {};
        hitTestInfo.pt = pt;

        int hitResult = ListView_HitTest(hwnd, &hitTestInfo);
        if (hitResult != -1) {
            int currentRow = hitTestInfo.iItem;
            int currentSubItem = hitTestInfo.iSubItem;

            // Update only if the row, sub-item, or mouse position has significantly changed
            if (currentRow != pThis->lastTooltipRow || currentSubItem != pThis->lastTooltipSubItem ||
                abs(pt.x - pThis->lastMouseX) > 5 || abs(pt.y - pThis->lastMouseY) > 5) {

                pThis->lastTooltipRow = currentRow;
                pThis->lastTooltipSubItem = currentSubItem;
                pThis->lastMouseX = pt.x;
                pThis->lastMouseY = pt.y;

                // Temporarily disable LVS_EX_INFOTIP to force tooltip update
                DWORD extendedStyle = ListView_GetExtendedListViewStyle(hwnd);
                ListView_SetExtendedListViewStyle(hwnd, extendedStyle & ~LVS_EX_INFOTIP);

                // Set a timer to re-enable LVS_EX_INFOTIP after 10ms
                SetTimer(hwnd, 1, 10, NULL);
            }
        }
        break;
    }

    case WM_SYSKEYDOWN:
    {
        if ((GetKeyState(VK_MENU) & 0x8000) && wParam == VK_UP) {
            int iItem = ListView_GetNextItem(hwnd, -1, LVNI_SELECTED);
            if (iItem >= 0) { NMITEMACTIVATE nmia{}; nmia.iItem = iItem; pThis->handleCopyBack(&nmia); }
            return 0;
        }
        // Optional fallback for Alt+E / Alt+D
        if (GetKeyState(VK_MENU) & 0x8000) {
            if (wParam == 'E') { pThis->setSelections(true, ListView_GetSelectedCount(hwnd) > 0);  return 0; }
            if (wParam == 'D') { pThis->setSelections(false, ListView_GetSelectedCount(hwnd) > 0);  return 0; }
        }
        break;
    }

    case WM_TIMER: {
        if (wParam == 1) { // Tooltip re-enable timer   
            KillTimer(hwnd, 1); // Kill the timer first to prevent it from firing again
            if (!pThis->isHoverTextEnabled || pThis->isHoverTextSuppressed) {
                return 0;
            }
            // Re-enable LVS_EX_INFOTIP
            DWORD extendedStyle = ListView_GetExtendedListViewStyle(hwnd);
            ListView_SetExtendedListViewStyle(hwnd, extendedStyle | LVS_EX_INFOTIP);
        }
        break;
    }

    case WM_COMMAND: {
        // Check if the source is our expand button
        WORD wId = LOWORD(wParam);
        WORD wEvent = HIWORD(wParam);

        if (wId == ID_EDIT_EXPAND_BUTTON && wEvent == BN_CLICKED) {
            // user clicked the expand button
            pThis->toggleEditExpand();
            return 0;
        }
        break;
    }

    case WM_UPDATE_FOCUS:
    {
        pThis->showListFilePath();
        return 0;
    }

    default:
        break;
    }

    // For all unhandled messages, call the default subclass procedure.
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void MultiReplace::toggleEditExpand()
{
    if (!hwndEdit || !hwndExpandBtn)
        return;

    // Get current position and size of the edit control in ListView coordinates
    RECT rc;
    GetWindowRect(hwndEdit, &rc);
    POINT ptLT = { rc.left, rc.top };
    POINT ptRB = { rc.right, rc.bottom };
    MapWindowPoints(NULL, _replaceListView, &ptLT, 1);
    MapWindowPoints(NULL, _replaceListView, &ptRB, 1);

    int curWidth = ptRB.x - ptLT.x;
    int curHeight = ptRB.y - ptLT.y;

    // Calculate new height (expand or collapse)
    int newHeight;
    if (_editIsExpanded) {
        newHeight = curHeight / editFieldSize; // Collapse
        SetWindowText(hwndExpandBtn, L"↓");
    }
    else {
        newHeight = curHeight * editFieldSize; // Expand
        SetWindowText(hwndExpandBtn, L"↑");
    }
    _editIsExpanded = !_editIsExpanded;

    // Set the _hBoldFont2 to the expand/collapse button
    SendMessage(hwndExpandBtn, WM_SETFONT, (WPARAM)_hBoldFont2, TRUE);

    // Update position and size of edit control and button
    MoveWindow(hwndEdit, ptLT.x, ptLT.y, curWidth, newHeight, TRUE);
    MoveWindow(hwndExpandBtn, ptLT.x + curWidth, ptLT.y - 1, sx(20), newHeight + 2, TRUE);

    // Bring controls to the top and ensure edit has focus
    SetWindowPos(hwndEdit, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    SetWindowPos(hwndExpandBtn, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    SetFocus(hwndEdit);
}

void MultiReplace::createContextMenu(HWND hwnd, POINT ptScreen, MenuState state) {
    HMENU hMenu = CreatePopupMenu();
    if (hMenu) {
        AppendMenu(hMenu, MF_STRING | (state.canUndo ? MF_ENABLED : MF_GRAYED), IDM_UNDO, LM.get(L"ctxmenu_undo").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canRedo ? MF_ENABLED : MF_GRAYED), IDM_REDO, LM.get(L"ctxmenu_redo").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_CUT_LINES_TO_CLIPBOARD, LM.get(L"ctxmenu_cut").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_COPY_LINES_TO_CLIPBOARD, LM.get(L"ctxmenu_copy").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canPaste ? MF_ENABLED : MF_GRAYED), IDM_PASTE_LINES_FROM_CLIPBOARD, LM.get(L"ctxmenu_paste").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_SELECT_ALL, LM.get(L"ctxmenu_select_all").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.canEdit ? MF_ENABLED : MF_GRAYED), IDM_EDIT_VALUE, LM.get(L"ctxmenu_edit").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_DELETE_LINES, LM.get(L"ctxmenu_delete").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_ADD_NEW_LINE, LM.get(L"ctxmenu_add_new_line").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.clickedOnItem ? MF_ENABLED : MF_GRAYED), IDM_COPY_DATA_TO_FIELDS, LM.get(L"ctxmenu_transfer_to_input_fields").c_str());
        AppendMenu(hMenu, MF_STRING | (state.listNotEmpty ? MF_ENABLED : MF_GRAYED), IDM_SEARCH_IN_LIST, LM.get(L"ctxmenu_search_in_list").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allEnabled ? MF_ENABLED : MF_GRAYED), IDM_ENABLE_LINES, LM.get(L"ctxmenu_enable").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allDisabled ? MF_ENABLED : MF_GRAYED), IDM_DISABLE_LINES, LM.get(L"ctxmenu_disable").c_str());
        TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, NULL);
        DestroyMenu(hMenu); // Clean up
    }
}

MenuState MultiReplace::checkMenuConditions(POINT ptScreen) {
    MenuState state;

    // Convert screen coordinates to client coordinates
    POINT ptClient = ptScreen;
    ScreenToClient(_replaceListView, &ptClient);

    // Perform hit testing to determine the exact location of the click within the ListView
    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = ptClient;
    int hitTestResult = ListView_HitTest(_replaceListView, &hitInfo);
    state.clickedOnItem = (hitTestResult != -1);

    // Determine the clicked column
    int clickedColumn = -1;
    int totalWidth = 0;
    HWND header = ListView_GetHeader(_replaceListView);
    int columnCount = Header_GetItemCount(header);

    // Iterate through the columns and calculate the total width
    for (int i = 0; i < columnCount; ++i) {
        totalWidth += ListView_GetColumnWidth(_replaceListView, i);
        if (ptClient.x < totalWidth) {
            clickedColumn = i;
            break;
        }
    }

    // Map the clicked column to ColumnID, using a direct lookup in columnIndices
    int columnID = getColumnIDFromIndex(clickedColumn);

    // Enable editing if the clicked column is editable and currently visible
    state.canEdit = state.clickedOnItem && (
        columnID == ColumnID::SELECTION ||
        columnID == ColumnID::WHOLE_WORD ||
        columnID == ColumnID::MATCH_CASE ||
        columnID == ColumnID::USE_VARIABLES ||
        columnID == ColumnID::EXTENDED ||
        columnID == ColumnID::REGEX ||
        columnID == ColumnID::FIND_TEXT ||
        columnID == ColumnID::REPLACE_TEXT ||
        columnID == ColumnID::COMMENTS
        );

    // Rest of the checks remain the same
    state.listNotEmpty = (ListView_GetItemCount(_replaceListView) > 0);
    state.canPaste = canPasteFromClipboard();
    state.hasSelection = (ListView_GetSelectedCount(_replaceListView) > 0);

    unsigned int enabledCount = 0, disabledCount = 0;

    int itemIndex = -1;
    while ((itemIndex = ListView_GetNextItem(_replaceListView, itemIndex, LVNI_SELECTED)) != -1) {
        auto& itemData = replaceListData[itemIndex];
        if (itemData.isEnabled) {
            ++enabledCount;
        }
        else {
            ++disabledCount;
        }
    }

    state.allEnabled = (enabledCount == ListView_GetSelectedCount(_replaceListView));
    state.allDisabled = (disabledCount == ListView_GetSelectedCount(_replaceListView));

    state.canUndo = URM.canUndo();
    state.canRedo = URM.canRedo();

    return state;
}

void MultiReplace::performItemAction(POINT pt, ItemAction action) {
    // Get the horizontal scroll offset
    SCROLLINFO si = {};
    si.cbSize = sizeof(si);
    si.fMask = SIF_POS;
    GetScrollInfo(_replaceListView, SB_HORZ, &si);
    int scrollX = si.nPos;

    // Adjust point based on horizontal scroll
    POINT ptAdjusted = pt;
    ptAdjusted.x += scrollX;

    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = ptAdjusted; // Use adjusted coordinates
    int hitTestResult = ListView_HitTest(_replaceListView, &hitInfo);

    // Calculate the clicked column based on the adjusted X coordinate
    int clickedColumn = -1;
    int totalWidth = 0;
    HWND header = ListView_GetHeader(_replaceListView);
    int columnCount = Header_GetItemCount(header);

    for (int i = 0; i < columnCount; ++i) {
        totalWidth += ListView_GetColumnWidth(_replaceListView, i);
        if (ptAdjusted.x < totalWidth) {
            clickedColumn = i;
            break; // Column found
        }
    }

    // Correct mapping from clickedColumn to ColumnID
    ColumnID columnID = getColumnIDFromIndex(clickedColumn);

    switch (action) {
    case ItemAction::Undo:
        URM.undo();
        showListFilePath();
        break;
    case ItemAction::Redo:
        URM.redo();
        showListFilePath();
        break;
    case ItemAction::Search:
        performSearchInList();
        break;
    case ItemAction::Cut:
        copySelectedItemsToClipboard();
        deleteSelectedLines();
        showListFilePath();
        break;
    case ItemAction::Copy:
        copySelectedItemsToClipboard();
        break;
    case ItemAction::Paste:
        pasteItemsIntoList();
        showListFilePath();
        break;
    case ItemAction::Edit:
        if (columnID == ColumnID::FIND_TEXT ||
            columnID == ColumnID::REPLACE_TEXT ||
            columnID == ColumnID::COMMENTS) {
            editTextAt(hitTestResult, columnID);
        }
        else if (columnID == ColumnID::SELECTION ||
            columnID == ColumnID::WHOLE_WORD ||
            columnID == ColumnID::MATCH_CASE ||
            columnID == ColumnID::USE_VARIABLES ||
            columnID == ColumnID::EXTENDED ||
            columnID == ColumnID::REGEX) {
            toggleBooleanAt(hitTestResult, columnID);
            showListFilePath();
        }
        break;
    case ItemAction::Delete: {
        int selectedCount = ListView_GetSelectedCount(_replaceListView);
        std::wstring confirmationMessage;

        if (selectedCount == 1) {
            confirmationMessage = LM.get(L"msgbox_confirm_delete_single");
        }
        else if (selectedCount > 1) {
            confirmationMessage = LM.get(L"msgbox_confirm_delete_multiple", { std::to_wstring(selectedCount) });
        }

        int msgBoxID = MessageBox(nppData._nppHandle, confirmationMessage.c_str(), LM.get(L"msgbox_title_confirm").c_str(), MB_ICONWARNING | MB_YESNO);
        if (msgBoxID == IDYES) {
            deleteSelectedLines();
            showListFilePath();
        }
        break;
    }
    case ItemAction::Add: {
        int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
        if (insertPosition != -1) {
            ++insertPosition;
        }
        else {
            insertPosition = ListView_GetItemCount(_replaceListView);
        }
        ReplaceItemData newItem; // Default-initialized
        std::vector<ReplaceItemData> itemsToAdd = { newItem };
        addItemsToReplaceList(itemsToAdd, static_cast<size_t>(insertPosition));

        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);
        ListView_SetItemState(_replaceListView, insertPosition, LVIS_SELECTED, LVIS_SELECTED);
        ListView_EnsureVisible(_replaceListView, insertPosition, FALSE);
        showListFilePath();
        break;
    }
    }
}

void MultiReplace::copySelectedItemsToClipboard() {
    const int itemCount = ListView_GetItemCount(_replaceListView);
    const int columnCount = Header_GetItemCount(ListView_GetHeader(_replaceListView));
    std::wstring csvData;
    csvData.reserve(itemCount * columnCount * 16);

    for (int i = 0; i < itemCount; ++i) {
        if (ListView_GetItemState(_replaceListView, i, LVIS_SELECTED) & LVIS_SELECTED) {
            const ReplaceItemData& item = replaceListData[i];
            csvData += std::to_wstring(item.isEnabled) + L",";
            csvData += escapeCsvValue(item.findText) + L",";
            csvData += escapeCsvValue(item.replaceText) + L",";
            csvData += std::to_wstring(item.wholeWord) + L",";
            csvData += std::to_wstring(item.matchCase) + L",";
            csvData += std::to_wstring(item.useVariables) + L",";
            csvData += std::to_wstring(item.extended) + L",";
            csvData += std::to_wstring(item.regex) + L",";
            csvData += escapeCsvValue(item.comments);
            csvData += L'\n';
        }
    }

    if (csvData.empty()) return;

    if (OpenClipboard(nullptr)) {
        EmptyClipboard();
        const SIZE_T size = (csvData.size() + 1) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
        if (hMem) {
            if (void* p = GlobalLock(hMem)) {
                memcpy(p, csvData.c_str(), size);
                GlobalUnlock(hMem);
                if (!SetClipboardData(CF_UNICODETEXT, hMem)) {
                    GlobalFree(hMem);
                }
            }
            else {
                GlobalFree(hMem);
            }
        }
        CloseClipboard();
    }
}

bool MultiReplace::canPasteFromClipboard() {
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !OpenClipboard(nullptr)) {
        return false;
    }
    struct ClipboardGuard { ~ClipboardGuard() { CloseClipboard(); } } guard;

    HGLOBAL hData = GetClipboardData(CF_UNICODETEXT);
    if (!hData) {
        return false;
    }
    wchar_t* raw = static_cast<wchar_t*>(GlobalLock(hData));
    if (!raw) {
        return false;
    }
    std::wstring content{ raw };
    GlobalUnlock(hData);

    std::wistringstream stream(content);
    std::wstring line;
    while (std::getline(stream, line)) {
        if (line.empty()) {
            continue;
        }
        auto columns = parseCsvLine(line);
        if (columns.size() == 8 || columns.size() == 9) {
            return true;
        }
    }
    return false;
}

void MultiReplace::pasteItemsIntoList() {
    if (!OpenClipboard(nullptr)) return;

    struct ClipboardGuard { ~ClipboardGuard() { CloseClipboard(); } } guard;

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (!hData) return;

    auto pData = static_cast<const wchar_t*>(GlobalLock(hData));
    if (!pData) return;

    std::wstring content(pData);
    GlobalUnlock(hData);

    std::wstringstream contentStream(content);
    std::wstring line;
    std::vector<ReplaceItemData> itemsToInsert; // Collect items to insert

    // Determine insert position based on focused item or default to end if none is focused
    int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
    if (insertPosition != -1) {
        // Increase by one to insert after the focused item
        ++insertPosition;
    }
    else {
        // If no item is focused, consider the paste position as the end of the list
        insertPosition = ListView_GetItemCount(_replaceListView);
    }

    while (std::getline(contentStream, line)) {
        if (line.empty()) continue; // Skip empty lines

        std::vector<std::wstring> columns = parseCsvLine(line);

        // Check for proper column count before adding to the list
        if ((columns.size() != 8 && columns.size() != 9)) continue;

        ReplaceItemData item;
        try {
            item.isEnabled    = std::stoi(columns[0]) != 0;
            item.findText     = columns[1];
            item.replaceText  = columns[2];
            item.wholeWord    = std::stoi(columns[3]) != 0;
            item.matchCase    = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended     = std::stoi(columns[6]) != 0;
            item.regex        = std::stoi(columns[7]) != 0;
            item.comments     = (columns.size() == 9 ? columns[8] : L"");
        }
        catch (const std::exception&) {
            continue; // Silently ignore lines with conversion errors
        }

        itemsToInsert.push_back(item);
    }

    // No items were collected, so nothing to insert
    if (itemsToInsert.empty()) return;

    // Use addItemsToReplaceList to insert items at the specified position
    addItemsToReplaceList(itemsToInsert, static_cast<size_t>(insertPosition));

    // Deselect all items first
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select newly inserted items in the list view
    for (size_t i = 0; i < itemsToInsert.size(); ++i) {
        size_t idx = static_cast<size_t>(insertPosition) + i;
        ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
    }

    // Ensure the first inserted item is visible
    ListView_EnsureVisible(_replaceListView, insertPosition, FALSE);
}

void MultiReplace::performSearchInList() {
    std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
    std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

    // Exit if both fields are empty
    if (findText.empty() && replaceText.empty()) {
        showStatusMessage(LM.get(L"status_no_find_replace_list_input"), MessageStatus::Error);
        return;
    }

    int startIdx = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED); // Get selected item or -1 if no selection
    int matchIdx = searchInListData(startIdx, findText, replaceText);

    if (matchIdx != -1) {
        // Deselect all items first
        int itemCount = ListView_GetItemCount(_replaceListView);
        for (int i = 0; i < itemCount; ++i) {
            ListView_SetItemState(_replaceListView, i, 0, LVIS_SELECTED | LVIS_FOCUSED);
        }

        // Highlight the matched item
        ListView_SetItemState(_replaceListView, matchIdx, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(_replaceListView, matchIdx, FALSE);
        showStatusMessage(LM.get(L"status_found_in_list"), MessageStatus::Success);
    }
    else {
        // Show failure status message if no match found
        showStatusMessage(LM.get(L"status_not_found_in_list"), MessageStatus::Error);
    }
}

int MultiReplace::searchInListData(int startIdx, const std::wstring& findText, const std::wstring& replaceText) {
    int listSize = static_cast<int>(replaceListData.size());
    bool searchFromStartAgain = true;

    // If startIdx is -1 (no selection), start from the beginning
    int i = (startIdx == -1) ? 0 : startIdx + 1;

    while (true) {
        if (i == listSize) {
            if (searchFromStartAgain) {
                // Restart search from the beginning until the original start index
                i = 0;
                searchFromStartAgain = false;
            }
            else {
                // No match found
                return -1;
            }
        }

        if (i == startIdx) break; // Stop if we have searched the entire list

        const auto& item = replaceListData[i];
        bool findMatch = findText.empty() || item.findText.find(findText) != std::wstring::npos;
        bool replaceMatch = replaceText.empty() || item.replaceText.find(replaceText) != std::wstring::npos;

        if (findMatch && replaceMatch) {
            // Match found
            return i;
        }

        ++i;
    }

    // No match found
    return -1;
}

void MultiReplace::handleEditOnDoubleClick(int itemIndex, ColumnID columnID) {
    // Perform the appropriate action based on the ColumnID
    if (columnID == ColumnID::FIND_TEXT || columnID == ColumnID::REPLACE_TEXT || columnID == ColumnID::COMMENTS) {
        editTextAt(itemIndex, columnID); // Pass ColumnID directly
    }
    else if (columnID == ColumnID::SELECTION ||
        columnID == ColumnID::WHOLE_WORD ||
        columnID == ColumnID::MATCH_CASE ||
        columnID == ColumnID::USE_VARIABLES ||
        columnID == ColumnID::EXTENDED ||
        columnID == ColumnID::REGEX) {
        toggleBooleanAt(itemIndex, columnID); // Pass ColumnID directly
    }
}

#pragma endregion


#pragma region Dialog

INT_PTR CALLBACK MultiReplace::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        // checkFullStackInfo();
        dpiMgr = new DPIManager(_hSelf);
        //loadLanguage();

        //  a) extract paths
        wchar_t pluginDir[MAX_PATH] = {};
        SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, MAX_PATH,
            reinterpret_cast<LPARAM>(pluginDir));

        wchar_t langXml[MAX_PATH] = {};
        SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH,
            reinterpret_cast<LPARAM>(langXml));
        wcscat_s(langXml, L"\\..\\..\\nativeLang.xml");

        LanguageManager::instance().load(pluginDir, langXml);
        initializeWindowSize();
        pointerToScintilla();
        initializeMarkerStyle();
        initializeCtrlMap();
        initializeFontStyles();
        applyThemePalette();
        loadSettings();
        updateTwoButtonsVisibility();
        initializeListView();
        initializeDragAndDrop();
        adjustWindowSize();

        // Activate Dark Mode
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME,
            static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(_hSelf));

        // Post a custom message to perform post-initialization tasks after the dialog is shown
        PostMessage(_hSelf, WM_POST_INIT, 0, 0);

        return TRUE;
    }

    case WM_POST_INIT:
    {
        checkForFileChangesAtStartup();
        return TRUE;
    }

    case WM_GETMINMAXINFO:
    {
        MINMAXINFO* pMMI = reinterpret_cast<MINMAXINFO*>(lParam);
        RECT adjustedSize = calculateMinWindowFrame(_hSelf);

        // Set minimum window width
        pMMI->ptMinTrackSize.x = adjustedSize.right;
        // Allow horizontal resizing up to a maximum value
        pMMI->ptMaxTrackSize.x = MAXLONG;

        if (useListEnabled) {
            // Set minimum window height
            pMMI->ptMinTrackSize.y = adjustedSize.bottom;
            // Allow vertical resizing up to a maximum value
            pMMI->ptMaxTrackSize.y = MAXLONG;
        }
        else {
            // Set fixed window height
            pMMI->ptMinTrackSize.y = adjustedSize.bottom;
            pMMI->ptMaxTrackSize.y = adjustedSize.bottom; // Fix the height when Use List is unchecked
        }

        return 0;
    }

    case WM_ACTIVATE:
    {
        if (_keepOnTopDuringBatch) {
            // Force foreground alpha during batch
            SetWindowTransparency(_hSelf, foregroundTransparency);
            return 0;
        }

        if (LOWORD(wParam) == WA_INACTIVE) {
            // The window loses focus
            SetWindowTransparency(_hSelf, backgroundTransparency); // Use the loaded value
        }
        else {
            // The window gains focus
            SetWindowTransparency(_hSelf, foregroundTransparency); // Use the loaded value
        }
        return 0;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_STATUS_MESSAGE)) {
            SetTextColor(hdcStatic, _statusMessageColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return (LRESULT)GetStockObject(NULL_BRUSH); // Return a brush handle
        }

        return FALSE;
    }

    case WM_DESTROY:
    {

        if (_replaceListView && originalListViewProc) {
            SetWindowLongPtr(_replaceListView, GWLP_WNDPROC, (LONG_PTR)originalListViewProc);
        }

        saveSettings(); // Save any settings before destroying

        // Unregister Drag-and-Drop
        if (dropTarget) {
            RevokeDragDrop(_replaceListView);
            delete dropTarget;
            dropTarget = nullptr;
        }

        if (hwndEdit) {
            DestroyWindow(hwndEdit);
        }

        DeleteObject(_hStandardFont);
        DeleteObject(_hBoldFont1);
        DeleteObject(_hNormalFont1);
        DeleteObject(_hNormalFont2);
        DeleteObject(_hNormalFont3);
        DeleteObject(_hNormalFont4);
        DeleteObject(_hNormalFont5);
        DeleteObject(_hNormalFont6);

        // Close the debug window if open
        if (hDebugWnd != NULL) {
            RECT rect;
            if (GetWindowRect(hDebugWnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }
            PostMessage(hDebugWnd, WM_CLOSE, 0, 0);
            hDebugWnd = NULL; // Reset the handle after closing
        }

        // Clean up DPIManager
        if (dpiMgr) {
            delete dpiMgr;
            dpiMgr = nullptr;
        }

        DestroyWindow(_hSelf); // Destroy the main window

        // Post a quit message to ensure the application terminates cleanly
        PostQuitMessage(0);

        return 0;
    }

    case WM_SIZE:
    {
        if (isWindowOpen) {
            // Force the edit control of the right mouse click to lose focus by setting focus to the main window
            if (isWindowOpen && hwndEdit && GetFocus() == hwndEdit) {
                HWND hwndListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
                SetFocus(hwndListView);
            }

            int newWidth = LOWORD(lParam);
            int newHeight = HIWORD(lParam);

            // Calculate Position for all Elements
            positionAndResizeControls(newWidth, newHeight);

            // Move and resize the List
            updateListViewAndColumns();

            // Move all Elements
            moveAndResizeControls();

            // Refresh UI and gripper by invalidating window
            InvalidateRect(_hSelf, NULL, TRUE);

            if (useListEnabled) {
                // Update useListOnHeight with the new height
                RECT currentRect;
                GetWindowRect(_hSelf, &currentRect);
                int currentHeight = currentRect.bottom - currentRect.top;

                // Get the minimum allowed window height
                RECT minSize = calculateMinWindowFrame(_hSelf);
                int minHeight = minSize.bottom;

                // Ensure useListOnHeight is at least the minimum height
                useListOnHeight = std::max(currentHeight, minHeight);
            }
        }
        return 0;
    }

    case WM_NOTIFY:
    {
        NMHDR* pnmh = reinterpret_cast<NMHDR*>(lParam);

        if (pnmh->code == BCN_DROPDOWN && pnmh->hwndFrom == GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON))
        {
            // split-button menu for Replace All
            RECT rc; ::GetWindowRect(pnmh->hwndFrom, &rc);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_ALL_OPTION, LM.getLPW(L"split_menu_replace_all"));
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_IN_ALL_DOCS_OPTION, LM.getLPW(L"split_menu_replace_all_in_docs"));
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_IN_FILES_OPTION, LM.getLPW(L"split_menu_replace_all_in_files"));
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, rc.left, rc.bottom, 0, _hSelf, NULL);
            DestroyMenu(hMenu);
            return TRUE;
        }

        if (pnmh->code == BCN_DROPDOWN && pnmh->hwndFrom == GetDlgItem(_hSelf, IDC_FIND_ALL_BUTTON))
        {
            // split-button menu for Find All
            RECT rc; ::GetWindowRect(pnmh->hwndFrom, &rc);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_OPTION, LM.getLPW(L"split_menu_find_all"));
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_IN_ALL_DOCS_OPTION, LM.getLPW(L"split_menu_find_all_in_docs"));
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_IN_FILES_OPTION, LM.getLPW(L"split_menu_find_all_in_files"));
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, rc.left, rc.bottom, 0, _hSelf, nullptr);
            DestroyMenu(hMenu);
            return TRUE;
        }

        if (pnmh->idFrom == IDC_REPLACE_LIST) {
            switch (pnmh->code) {
            case NM_CLICK:
            {
                NMITEMACTIVATE* pnmia = reinterpret_cast<NMITEMACTIVATE*>(lParam);
                int subItem = pnmia->iSubItem;
                int itemIndex = pnmia->iItem;

                // Ensure valid item index, as subItem 0 could refer to the first column or an invalid click area.
                // This prevents accidental actions when clicking outside valid list items.
                if (itemIndex >= 0 && itemIndex < static_cast<int>(replaceListData.size())) {
                    ColumnID columnID = getColumnIDFromIndex(subItem);

                    switch (columnID) {
                    case ColumnID::DELETE_BUTTON:
                        handleDeletion(pnmia);
                        break;

                    case ColumnID::SELECTION:
                    {
                        bool currentSelectionStatus = replaceListData[itemIndex].isEnabled;
                        setSelections(!currentSelectionStatus, true);
                    }
                    break;

                    default:
                        break;
                    }

                    showListFilePath();
                }
                return TRUE;
            }

            case NM_DBLCLK:
            {
                NMITEMACTIVATE* pnmia = reinterpret_cast<NMITEMACTIVATE*>(lParam);
                int itemIndex = pnmia->iItem;
                int clickedColumn = pnmia->iSubItem;

                if (itemIndex != -1 && clickedColumn != -1) {
                    if (doubleClickEditsEnabled) {
                        ColumnID columnID = getColumnIDFromIndex(clickedColumn);
                        handleEditOnDoubleClick(itemIndex, columnID);
                    }
                    else {
                        handleCopyBack(pnmia);
                    }
                }
                return TRUE;
            }

            case LVN_GETDISPINFO:
            {
                NMLVDISPINFO* plvdi = reinterpret_cast<NMLVDISPINFO*>(lParam);
                const int itemIndex = plvdi->item.iItem;
                const int subItem = plvdi->item.iSubItem;

                // Provide text only when requested
                if ((plvdi->item.mask & LVIF_TEXT) == 0)
                    return TRUE;

                if (itemIndex < 0 || itemIndex >= static_cast<int>(replaceListData.size()))
                {
                    if (plvdi->item.pszText && plvdi->item.cchTextMax > 0)
                        plvdi->item.pszText[0] = L'\0';
                    return TRUE;
                }

                const ReplaceItemData& d = replaceListData[itemIndex];

                // Fill text depending on which column is asking
                if (columnIndices[ColumnID::FIND_COUNT] != -1 && subItem == columnIndices[ColumnID::FIND_COUNT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.findCount.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::REPLACE_COUNT] != -1 && subItem == columnIndices[ColumnID::REPLACE_COUNT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.replaceCount.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::SELECTION] != -1 && subItem == columnIndices[ColumnID::SELECTION]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.isEnabled ? L"\u25A0" : L"\u2610", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::FIND_TEXT] != -1 && subItem == columnIndices[ColumnID::FIND_TEXT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.findText.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::REPLACE_TEXT] != -1 && subItem == columnIndices[ColumnID::REPLACE_TEXT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.replaceText.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::WHOLE_WORD] != -1 && subItem == columnIndices[ColumnID::WHOLE_WORD]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.wholeWord ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::MATCH_CASE] != -1 && subItem == columnIndices[ColumnID::MATCH_CASE]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.matchCase ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::USE_VARIABLES] != -1 && subItem == columnIndices[ColumnID::USE_VARIABLES]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.useVariables ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::EXTENDED] != -1 && subItem == columnIndices[ColumnID::EXTENDED]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.extended ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::REGEX] != -1 && subItem == columnIndices[ColumnID::REGEX]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.regex ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::COMMENTS] != -1 && subItem == columnIndices[ColumnID::COMMENTS]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.comments.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::DELETE_BUTTON] != -1 && subItem == columnIndices[ColumnID::DELETE_BUTTON]) {
                    (void)lstrcpynW(plvdi->item.pszText, L"\u2716", plvdi->item.cchTextMax);
                }
                else {
                    if (plvdi->item.pszText && plvdi->item.cchTextMax > 0)
                        plvdi->item.pszText[0] = L'\0';
                }
                return TRUE;
            }

            case LVN_COLUMNCLICK:
            {
                NMLISTVIEW* pnmv = reinterpret_cast<NMLISTVIEW*>(lParam);

                int clickedColumn = pnmv->iSubItem;

                // Map the clicked column index to ColumnID
                int columnID = getColumnIDFromIndex(clickedColumn);

                if (columnID == ColumnID::INVALID) {
                    // If no valid column is found, ignore the click
                    return TRUE;
                }

                if (columnID == ColumnID::SELECTION) {
                    setSelections(!allSelected);
                }
                else {
                    sortReplaceListData(columnID);
                }
                return TRUE;
            }

            case LVN_KEYDOWN:
            {
                LPNMLVKEYDOWN pnkd = reinterpret_cast<LPNMLVKEYDOWN>(pnmh);
                int iItem = -1;

                // For arrow keys and Page Up/Page Down, post a custom message to update the focus
                if (pnkd->wVKey == VK_UP || pnkd->wVKey == VK_DOWN ||
                    pnkd->wVKey == VK_PRIOR || pnkd->wVKey == VK_NEXT)
                {
                    PostMessage(_replaceListView, WM_UPDATE_FOCUS, 0, 0);
                    return TRUE;
                }

                // For non-arrow keys, proceed with existing focus update.
                PostMessage(_replaceListView, WM_SETFOCUS, 0, 0);

                // Handling keyboard shortcuts for menu actions
                if (GetKeyState(VK_CONTROL) & 0x8000) { // If Ctrl is pressed
                    switch (pnkd->wVKey) {
                    case 'Z': // Ctrl+Z for Undo
                        URM.undo();
                        break;
                    case 'Y': // Ctrl+Y for Redo
                        URM.redo();
                        break;
                    case 'F': // Ctrl+F for Search in List
                        performItemAction(_contextMenuClickPoint, ItemAction::Search);
                        break;
                    case 'X': // Ctrl+X for Cut
                        performItemAction(_contextMenuClickPoint, ItemAction::Cut);
                        break;
                    case 'C': // Ctrl+C for Copy
                        performItemAction(_contextMenuClickPoint, ItemAction::Copy);
                        break;
                    case 'V': // Ctrl+V for Paste
                        performItemAction(_contextMenuClickPoint, ItemAction::Paste);
                        break;
                    case 'A': // Ctrl+A for Select All
                        ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
                        showListFilePath();
                        break;
                    case 'I': // Ctrl+I for Adding new Line
                        performItemAction(_contextMenuClickPoint, ItemAction::Add);
                        break;
                    }
                } 
                else {
                    switch (pnkd->wVKey) {
                    case VK_DELETE: // Delete key for deleting selected lines
                        performItemAction(_contextMenuClickPoint, ItemAction::Delete);
                        break;
                    case VK_F12: // F12 key
                    {
                        showDPIAndFontInfo();  // Show the DPI and font information
                        break;
                    }
                    case VK_SPACE: // Spacebar key
                        iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                        if (iItem >= 0) {
                            // Get current selection status of the item
                            bool currentSelectionStatus = replaceListData[iItem].isEnabled;
                            // Set the selection status to its opposite
                            setSelections(!currentSelectionStatus, true);
                        }
                        break;
                    }
                }
                return TRUE;
            }
            }
        }
        return FALSE;
    }

    case WM_CONTEXTMENU:
    {
        // Check if the right-click is on the ListView (not the header)
        if ((HWND)wParam == _replaceListView) {
            POINT ptScreen;
            ptScreen.x = LOWORD(lParam);
            ptScreen.y = HIWORD(lParam);

            // Get the size of the virtual screen (the total width and height of all monitors)
            int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            // Check if the x-coordinates are outside the expected range and adjust
            if (ptScreen.x > virtualWidth) {
                ptScreen.x -= 65536; // Adjust for potential overflow
            }

            // Check if the y-coordinates are outside the expected range and adjust
            if (ptScreen.y > virtualHeight) {
                ptScreen.y -= 65536; // Adjust for potential overflow
            }

            _contextMenuClickPoint = ptScreen; // Store initial click point for later action determination
            ScreenToClient(_replaceListView, &_contextMenuClickPoint); // Convert to client coordinates for hit testing
            MenuState state = checkMenuConditions(ptScreen);
            createContextMenu(_hSelf, ptScreen, state); // Show context menu
            return TRUE;
        }
        return FALSE;
    }

    case WM_SHOWWINDOW:
    {
        if (wParam == TRUE) {
            std::wstring wstr = getSelectedText();

            // Set selected text in IDC_FIND_EDIT
            if (!wstr.empty()) {
                SetWindowTextW(GetDlgItem(_hSelf, IDC_FIND_EDIT), wstr.c_str());
            }
        }
        else {
            handleClearTextMarksButton();
            handleClearDelimiterState();
        }
        return 0;
    }

    case WM_PAINT:
    {
        drawGripper();
        return 0;
    }

    case WM_DRAWITEM:
    {
        DRAWITEMSTRUCT* pdis = (DRAWITEMSTRUCT*)lParam;

        if (pdis->CtlID == IDC_STATUS_MESSAGE) {
            wchar_t buffer[256];
            GetWindowTextW(pdis->hwndItem, buffer, 256);

            SetTextColor(pdis->hDC, _statusMessageColor);
            SetBkMode(pdis->hDC, TRANSPARENT);

            RECT textRect = pdis->rcItem;
            DrawTextW(pdis->hDC, buffer, -1, &textRect, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            return TRUE;
        }
        else if (pdis->CtlID == IDC_FILTER_HELP) {

            wchar_t buffer[16];
            GetWindowTextW(pdis->hwndItem, buffer, 16);

            SetTextColor(pdis->hDC, _filterHelpColor);
            SetBkMode(pdis->hDC, TRANSPARENT);

            DrawTextW(pdis->hDC, buffer, -1, &pdis->rcItem, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            return TRUE;
        }

        return FALSE;
    }

    case WM_CTLCOLORDLG:
    {
        // Store the handle to the brush that the dialog will use to paint its background
        _hDlgBrush = (HBRUSH)wParam;
        return (INT_PTR)_hDlgBrush;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {

        case IDC_PATH_DISPLAY:
        {
            if (HIWORD(wParam) == STN_DBLCLK)
            {
                onPathDisplayDoubleClick();
                return TRUE;
            }
            return FALSE;
        }

        case IDC_USE_VARIABLES_HELP:
        {
            auto n = SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, 0, 0);
            std::wstring path(n, 0);
            SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, n + 1, reinterpret_cast<LPARAM>(path.data()));
            path += L"\\MultiReplace";
            BOOL isDarkMode = SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != FALSE;
            std::wstring filename = isDarkMode ? L"\\help_use_variables_dark.html" : L"\\help_use_variables_light.html";
            path += filename;
            ShellExecute(NULL, L"open", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
            return TRUE;
        }

        case IDCANCEL:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            EndDialog(_hSelf, 0);
            _MultiReplace.display(false);
            return TRUE;
        }

        case IDC_2_BUTTONS_MODE:
        {
            // Check if the Find checkbox has been clicked
            if (HIWORD(wParam) == BN_CLICKED)
            {
                updateTwoButtonsVisibility();
                return TRUE;
            }
            return FALSE;
        }

        case IDC_REGEX_RADIO:
        {
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_NORMAL_RADIO:
        case IDC_EXTENDED_RADIO:
        {
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), TRUE);
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_ALL_TEXT_RADIO:
        {
            setUIElementVisibility();
            handleClearDelimiterState();
            return TRUE;
        }

        case IDC_SELECTION_RADIO:
        {
            setUIElementVisibility();
            handleClearDelimiterState();
            return TRUE;
        }

        case IDC_COLUMN_NUM_EDIT:
        case IDC_DELIMITER_EDIT:
        case IDC_QUOTECHAR_EDIT:
        case IDC_COLUMN_MODE_RADIO:
        {
            CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_COLUMN_MODE_RADIO);
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_COLUMN_SORT_ASC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortStateAndSort(SortDirection::Ascending);
                UpdateSortButtonSymbols();
            }
            return TRUE;
        }

        case IDC_COLUMN_SORT_DESC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortStateAndSort(SortDirection::Descending);
                UpdateSortButtonSymbols();
            }
            return TRUE;
        }

        case IDC_COLUMN_DROP_BUTTON:
        {
            if (confirmColumnDeletion()) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleDeleteColumns();
                }
            }
            return TRUE;
        }

        case IDC_COLUMN_COPY_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleCopyColumnsToClipboard();
            }
            return TRUE;
        }

        case IDC_COLUMN_HIGHLIGHT_BUTTON:
        {
            int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

            if (!highlightedTabs.isHighlighted(currentBufferID)) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleHighlightColumnsInDocument();
                }
            }
            else {
                handleClearColumnMarks();
                showStatusMessage(LanguageManager::instance().get(L"status_column_marks_cleared"), MessageStatus::Success);
            }
            return TRUE;
        }

        case IDC_USE_LIST_BUTTON:
        {
            useListEnabled = !useListEnabled;
            updateUseListState(true);
            adjustWindowSize();
            return TRUE;
        }

        case IDC_SWAP_BUTTON:
        {
            std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

            // Swap the content of the two text fields
            SetDlgItemTextW(_hSelf, IDC_FIND_EDIT, replaceText.c_str());
            SetDlgItemTextW(_hSelf, IDC_REPLACE_EDIT, findText.c_str());
            return TRUE;
        }

        case IDC_COPY_TO_LIST_BUTTON:
        {
            handleCopyToListButton();
            return TRUE;
        }

        case IDC_FIND_ALL_BUTTON:
        {
            CloseDebugWindow();
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);

            if (isFindAllInFiles)
                handleFindInFiles();
            else if (isFindAllInDocs)
                handleFindAllInDocsButton();
            else
                handleFindAllButton();

            return TRUE;
        }

        case ID_FIND_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getLPW(L"split_button_find_all"));
            isFindAllInDocs = false;
            isFindAllInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_FIND_ALL_IN_ALL_DOCS_OPTION: 
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getLPW(L"split_button_find_all_in_docs"));
            isFindAllInDocs = true;
            isFindAllInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_FIND_ALL_IN_FILES_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getLPW(L"split_button_find_all_in_files"));
            isFindAllInDocs = false;
            isFindAllInFiles = true;
            updateFilesPanel();
            return TRUE;
        }

        case IDC_FIND_NEXT_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindNextButton();
            return TRUE;
        }

        case IDC_FIND_PREV_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindPrevButton();
            return TRUE;
        }

        case IDC_REPLACE_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceButton();
            return TRUE;
        }

        case IDC_REPLACE_ALL_SMALL_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceAllButton();
            return TRUE;
        }

        case IDC_REPLACE_ALL_BUTTON:
        {
            CloseDebugWindow();

            if (isReplaceAllInDocs)
            {
                replaceAllInOpenedDocs();
            }
            else if (isReplaceInFiles) {
                std::wstring filter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
                std::wstring dir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);

                addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FILTER_EDIT), filter);
                addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), dir);
                handleReplaceInFiles();
            }
            else
            {
                resetCountColumns();
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                handleReplaceAllButton();
            }
            return TRUE;
        }

        case IDC_BROWSE_DIR_BUTTON:
        {
            handleBrowseDirectoryButton();
            return TRUE;
        }

        case IDC_MARK_MATCHES_BUTTON:
        case IDC_MARK_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleClearTextMarksButton();
            handleMarkMatchesButton();
            return TRUE;
        }

        case IDC_CLEAR_MARKS_BUTTON:
        {
            resetCountColumns();
            handleClearTextMarksButton();
            showStatusMessage(LM.get(L"status_all_marks_cleared"), MessageStatus::Success);
            return TRUE;
        }

        case IDC_COPY_MARKED_TEXT_BUTTON:
        {
            handleCopyMarkedTextToClipboardButton();
            return TRUE;
        }

        case IDC_SAVE_AS_BUTTON:
        case IDC_SAVE_TO_CSV_BUTTON:
        {
            std::wstring filePath = promptSaveListToCsv();

            if (!filePath.empty()) {
                saveListToCsv(filePath, replaceListData);

                // For "Save As", update the global listFilePath to the new file path
                if (wParam == IDC_SAVE_AS_BUTTON) {
                    listFilePath = filePath;
                }
            }
            ;

            return TRUE;
        }

        case IDC_SAVE_BUTTON:
        {
            if (!listFilePath.empty()) {
                saveListToCsv(listFilePath, replaceListData);
            }
            else {
                // If no file path is set, prompt the user to select a file path
                std::wstring filePath = promptSaveListToCsv();

                if (!filePath.empty()) {
                    saveListToCsv(filePath, replaceListData);
                }
            }
            return TRUE;
        }

        case IDC_LOAD_LIST_BUTTON:
        case IDC_LOAD_FROM_CSV_BUTTON:
        {
            // Check for unsaved changes before proceeding
            int userChoice = checkForUnsavedChanges();
            if (userChoice == IDCANCEL) {
                return TRUE;
            }

            std::wstring csvDescription = LM.get(L"filetype_csv");  // "CSV Files (*.csv)"
            std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {csvDescription, L"*.csv"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = LM.get(L"panel_load_list");
            std::wstring filePath = openFileDialog(false, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST, L"csv", L"");

            if (!filePath.empty()) {
                loadListFromCsv(filePath);
            }
            return TRUE;
        }

        case IDC_NEW_LIST_BUTTON:
        {
            clearList();
            showStatusMessage(LM.get(L"status_new_list_created"), MessageStatus::Success);
            return TRUE;
        }

        case IDC_UP_BUTTON:
        {
            shiftListItem(Direction::Up);
            return TRUE;
        }

        case IDC_DOWN_BUTTON:
        {
            shiftListItem(Direction::Down);
            return TRUE;
        }

        case IDC_EXPORT_BASH_BUTTON:
        {
            std::wstring bashDescription = LM.get(L"filetype_bash");  // "Bash Scripts (*.sh)"
            std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {bashDescription, L"*.sh"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = LM.get(L"panel_export_to_bash");

            // Set a default filename if none is provided
            static int scriptCounter = 1;
            std::wstring defaultFileName = L"Replace_Script_" + std::to_wstring(++scriptCounter) + L".sh";

            // Open the file dialog with the default filename for bash scripts
            std::wstring filePath = openFileDialog(true, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"sh", defaultFileName);

            if (!filePath.empty()) {
                exportToBashScript(filePath);
            }
            return TRUE;
        }

        case ID_REPLACE_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getLPW(L"split_button_replace_all"));
            isReplaceAllInDocs = false;
            isReplaceInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_REPLACE_IN_ALL_DOCS_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getLPW(L"split_button_replace_all_in_docs"));
            isReplaceAllInDocs = true;
            isReplaceInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_REPLACE_IN_FILES_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getLPW(L"split_button_replace_all_in_files"));
            isReplaceAllInDocs = false;
            isReplaceInFiles = true;
            updateFilesPanel();
            return TRUE;
        }

        case IDC_CANCEL_REPLACE_BUTTON:
        {
            _isCancelRequested = true;
            return TRUE;
        }

        case IDM_SEARCH_IN_LIST:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Search);
            return TRUE;
        }

        case IDM_UNDO: {
            performItemAction(_contextMenuClickPoint, ItemAction::Undo);
            return TRUE;
        }

        case IDM_REDO: {
            performItemAction(_contextMenuClickPoint, ItemAction::Redo);
            return TRUE;
        }

        case IDM_COPY_DATA_TO_FIELDS:
        {
            NMITEMACTIVATE nmia = {};
            nmia.iItem = ListView_HitTest(_replaceListView, &_contextMenuClickPoint);
            handleCopyBack(&nmia);
            return TRUE;
        }

        case IDM_CUT_LINES_TO_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Cut);
            return TRUE;
        }

        case IDM_COPY_LINES_TO_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Copy);
            return TRUE;
        }

        case IDM_PASTE_LINES_FROM_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Paste);
            return TRUE;
        }

        case IDM_EDIT_VALUE:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Edit);
            return TRUE;
        }

        case IDM_DELETE_LINES:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Delete);
            return TRUE;
        }

        case IDM_SELECT_ALL:
        {
            ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
            return TRUE;
        }

        case IDM_ENABLE_LINES:
        {
            setSelections(true, ListView_GetSelectedCount(_replaceListView) > 0);
            return TRUE;
        }

        case IDM_DISABLE_LINES:
        {
            setSelections(false, ListView_GetSelectedCount(_replaceListView) > 0);
            return TRUE;
        }

        case IDM_TOGGLE_FIND_COUNT:
        case IDM_TOGGLE_REPLACE_COUNT:
        case IDM_TOGGLE_COMMENTS:
        case IDM_TOGGLE_DELETE:
        {
            // Toggle the visibility based on the menu selection
            handleColumnVisibilityToggle(LOWORD(wParam));
            return TRUE;
        }

        case IDM_ADD_NEW_LINE:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Add);
            return TRUE;
        }

        default:
            return FALSE;
        }
    }

    default:
        return FALSE;
    }
}

#pragma endregion


#pragma region Replace

void MultiReplace::replaceAllInOpenedDocs()
{
    if (MessageBox(
        nppData._nppHandle,
        LM.get(L"msgbox_confirm_replace_all").c_str(),
        LM.get(L"msgbox_title_confirm").c_str(),
        MB_ICONWARNING | MB_OKCANCEL
    ) != IDOK)
    {
        return;
    }

    // Reset the UI counters before starting
    resetCountColumns();
    std::vector<int> listFindTotals(replaceListData.size(), 0);
    std::vector<int> listReplaceTotals(replaceListData.size(), 0);

    // How many docs are open in each view?
    LRESULT docCountMain = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, PRIMARY_VIEW);
    LRESULT docCountSecondary = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, SECOND_VIEW);

    bool visibleMain = IsWindowVisible(nppData._scintillaMainHandle);
    bool visibleSecond = IsWindowVisible(nppData._scintillaSecondHandle);

    // Remember which doc was active
    LRESULT currentDocIndex = ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTDOCINDEX, 0, MAIN_VIEW);

    // Local lambda to process one document
    auto processOneDoc = [&](int view, LRESULT idx) -> bool {
        ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, view, idx);
        handleDelimiterPositions(DelimiterOperation::LoadAll);
        if (!handleReplaceAllButton()) return false; // aborted via Stop/Error

        // Accumulate the per-list-entry counters from the UI
        for (size_t j = 0; j < replaceListData.size(); ++j) {
            int f = replaceListData[j].findCount.empty() ? 0 : std::stoi(replaceListData[j].findCount);
            int r = replaceListData[j].replaceCount.empty() ? 0 : std::stoi(replaceListData[j].replaceCount);
            listFindTotals[j] += f;
            listReplaceTotals[j] += r;
        }
        resetCountColumns();
        return true;
        };

    // Iterate main view
    if (visibleMain) {
        for (LRESULT i = 0; i < docCountMain; ++i) {
            if (!processOneDoc(PRIMARY_VIEW, i)) break;
        }
    }
    // Iterate secondary view
    if (visibleSecond) {
        for (LRESULT i = 0; i < docCountSecondary; ++i) {
            if (!processOneDoc(SECOND_VIEW, i)) break;
        }
    }

    // Restore the originally active document
    ::SendMessage(nppData._nppHandle,
        NPPM_ACTIVATEDOC,
        visibleMain ? PRIMARY_VIEW : SECOND_VIEW,
        currentDocIndex);

    // Write back only the enabled entries
    for (size_t j = 0; j < replaceListData.size(); ++j) {
        if (!replaceListData[j].isEnabled) continue;
        updateCountColumns(j, listFindTotals[j], listReplaceTotals[j]);
    }
    refreshUIListView();

    // Reset the scope radios to a consistent default
    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
    ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
    ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
}

bool MultiReplace::handleReplaceAllButton(bool showCompletionMessage, const std::filesystem::path* explicitPath) {

    if (!validateDelimiterData()) {
        return false;
    }

    // In "Replace in All Docs" + "Selection" scope, skip documents that have no selection.
    if (isReplaceAllInDocs &&
        IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED &&
        getSelectionInfo(false).length == 0)
    {
        return true;          // just jump into next Document
    }

    // First check if the document is read-only
    LRESULT isReadOnly = send(SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(LM.get(L"status_cannot_replace_read_only"), MessageStatus::Error);
        return false;
    }

    if (!initLuaState()) {
        // Fallback: The _luaInitialized flag remains false, 
        // so resolveLuaSyntax(...) calls will fail safely inside
        // and effectively do nothing. We just continue.
    }

    // Read Filename and Path for LUA
    updateFilePathCache(explicitPath);

    int totalReplaceCount = 0;
    bool replaceSuccess = true;

    if (useListEnabled)
    {
        // Check if the replaceListData is empty and warn the user if so
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_instructions"), MessageStatus::Error);
            return false;
        }

        // Check status for initial call if stopped by DEBUG, don't highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/false)) {
            return false;
        }

        // snapshot a stable start and current selection (only computed once)
        const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
        const bool allFromCursor = allFromCursorEnabled;

        SearchContext startCtx{};
        startCtx.docLength = send(SCI_GETLENGTH, 0, 0);
        startCtx.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
        startCtx.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
        startCtx.retrieveFoundText = false;
        startCtx.highlightMatch = false;

        const SelectionInfo fixedSel = getSelectionInfo(false);
        const Sci_Position  fixedStart = computeAllStartPos(startCtx, wrapAroundEnabled, allFromCursor);

        send(SCI_BEGINUNDOACTION, 0, 0);
        for (size_t i = 0; i < replaceListData.size(); ++i)
        {
            if (replaceListData[i].isEnabled)
            {
                if (!wrapAroundEnabled && allFromCursor)
                {
                    const Sci_Position docLenNow = send(SCI_GETLENGTH, 0, 0);
                    auto clamp = [&](Sci_Position p) {
                        return (p < 0) ? 0 : (p > docLenNow ? docLenNow : p);
                        };

                    if (startCtx.isSelectionMode) {
                        Sci_Position s = clamp(fixedSel.startPos);
                        Sci_Position e = clamp(fixedSel.endPos);
                        if (e < s) std::swap(s, e);
                        send(SCI_SETSEL, s, e);
                    }
                    else {
                        const Sci_Position s = clamp(fixedStart);
                        send(SCI_GOTOPOS, s, 0);
                    }
                }

                int findCount = 0;
                int replaceCount = 0;

                // Call replaceAll and break out if there is an error or a Debug Stop
                replaceSuccess = replaceAll(replaceListData[i], findCount, replaceCount, i);

                // Refresh ListView to show updated statistics
                refreshUIListView();

                // Accumulate total replacements
                totalReplaceCount += replaceCount;

                // cosmetic caret restore only for single-document run
                if (!replaceSuccess) {
                    break;
                }
            }
        }
        send(SCI_ENDUNDOACTION, 0, 0);
    }
    else
    {
        ReplaceItemData itemData;
        itemData.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        itemData.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        itemData.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
        itemData.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        send(SCI_BEGINUNDOACTION, 0, 0);
        int findCount = 0;
        replaceSuccess = replaceAll(itemData, findCount, totalReplaceCount);
        send(SCI_ENDUNDOACTION, 0, 0);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText);
    }
    // Display status message
    if (replaceSuccess && showCompletionMessage) {
        showStatusMessage(LM.get(L"status_occurrences_replaced", { std::to_wstring(totalReplaceCount) }), MessageStatus::Success);
    }

    // Only reset the scope radio buttons for a single-document "Replace All"
    if (!isReplaceAllInDocs) {
        SelectionInfo selection = getSelectionInfo(false);
        if (selection.length == 0 &&
            IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)
        {
            ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
            ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
            ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
        }

    }

    return replaceSuccess;

}

void MultiReplace::handleReplaceButton() {

    if (!validateDelimiterData()) {
        return;
    }

    // First check if the document is read-only
    LRESULT isReadOnly = send(SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(LM.getLPW(L"status_cannot_replace_read_only"), MessageStatus::Error);
        return;
    }

    if (!initLuaState()) {
        // Fallback: The _luaInitialized flag remains false, 
        // so resolveLuaSyntax(...) calls will fail safely inside
        // and effectively do nothing. We just continue.
    }

    // Read Filename and Path for LUA
    updateFilePathCache();

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SearchResult searchResult;
    searchResult.pos = -1;
    searchResult.length = 0;
    searchResult.foundText = "";

    SelectionInfo selection = getSelectionInfo(false);
    Sci_Position newPos = (selection.length > 0) ? selection.startPos : send(SCI_GETCURRENTPOS, 0, 0);

    size_t matchIndex = std::numeric_limits<size_t>::max();

    SearchContext context;
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_uncheck"), MessageStatus::Error);
            return;
        }

        // Check status of initial if stopped by DEBUG, highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/true)) {
            return;
        }

        bool wasReplaced = false;  // Detection for eplacements
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled) {
                context.findText = convertAndExtendW(replaceListData[i].findText, replaceListData[i].extended);
                context.searchFlags = (replaceListData[i].wholeWord * SCFIND_WHOLEWORD) |
                    (replaceListData[i].matchCase * SCFIND_MATCHCASE) |
                    (replaceListData[i].regex * SCFIND_REGEXP);

                // Set search flags before calling replaceOne
                send(SCI_SETSEARCHFLAGS, context.searchFlags);

                wasReplaced = replaceOne(replaceListData[i], selection, searchResult, newPos, i, context);
                if (wasReplaced) {
                    refreshUIListView(); // Refresh the ListView to show updated statistic
                    break;
                }
            }
        }

        if (!(wasReplaced && stayAfterReplaceEnabled)) {
            searchResult = performListSearchForward(replaceListData, newPos, matchIndex, context);

            if (searchResult.pos < 0 && wrapAroundEnabled) {
                searchResult = performListSearchForward(replaceListData, 0, matchIndex, context);
            }
        }

        // Build and show message based on results
        if (wasReplaced) {
            if (stayAfterReplaceEnabled) {
                refreshUIListView();
                showStatusMessage(LM.get(L"status_replace_one"), MessageStatus::Success);
            }
            else if (searchResult.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                showStatusMessage(LM.get(L"status_replace_one_next_found"), MessageStatus::Info);
            }
            else {
                showStatusMessage(LM.get(L"status_replace_one_none_left"), MessageStatus::Info);
            }
        }

        else {
            if (searchResult.pos < 0) {
                showStatusMessage(LM.get(L"status_no_occurrence_found"), MessageStatus::Error, true);
            }
            else {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_found_text_not_replaced"), MessageStatus::Info);
            }
        }
    }
    else {
        ReplaceItemData replaceItem;
        replaceItem.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        replaceItem.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        replaceItem.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        replaceItem.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        replaceItem.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
        replaceItem.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        replaceItem.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        context.findText = convertAndExtendW(replaceItem.findText, replaceItem.extended);
        context.searchFlags = (replaceItem.wholeWord * SCFIND_WHOLEWORD) |
            (replaceItem.matchCase * SCFIND_MATCHCASE) |
            (replaceItem.regex * SCFIND_REGEXP);

        // Set search flags before calling `performSearchForward` and 'replaceOne' which contains 'performSearchForward'
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        bool wasReplaced = replaceOne(replaceItem, selection, searchResult, newPos, SIZE_MAX, context);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), replaceItem.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceItem.replaceText);

        if (!(wasReplaced && stayAfterReplaceEnabled)) {
            if (searchResult.pos < 0 && wrapAroundEnabled) {
                searchResult = performSearchForward(context, 0);
            }
            else if (searchResult.pos >= 0) {
                searchResult = performSearchForward(context, newPos);
            }
        }

        if (wasReplaced) {
            if (stayAfterReplaceEnabled) {
                showStatusMessage(LM.get(L"status_replace_one"), MessageStatus::Success);
            }
            else if (searchResult.pos >= 0) {
                showStatusMessage(LM.get(L"status_replace_one_next_found"), MessageStatus::Success);
            }
            else {
                showStatusMessage(LM.get(L"status_replace_one_none_left"), MessageStatus::Info);
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(LM.get(L"status_no_occurrence_found"), MessageStatus::Error, true);
            }
            else {
                showStatusMessage(LM.get(L"status_found_text_not_replaced"), MessageStatus::Info);
            }
        }
    }

    // Disable selection radio and switch to "All Text" if it was Replaced and no selection left, or search will be trapped
    selection = getSelectionInfo(false);
    if (selection.length == 0 && IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
    }

}

bool MultiReplace::replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos, size_t itemIndex, const SearchContext& context)
{
    // Get the document's codepage once at the beginning.
    int documentCodepage = getCurrentDocCodePage();

    searchResult = performSearchForward(context, selection.startPos);

    // Only proceed if the found match is exactly the same as the initial selection.
    if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
        bool skipReplace = false;

        if (itemIndex != SIZE_MAX) {
            updateCountColumns(itemIndex, 1);
            selectListItem(itemIndex);
        }

        // This block contains the core replacement logic, structured to be consistent with replaceAll.
        {
            std::string finalReplaceText; // Will hold the final text in the document's native encoding.

            // --- Lua Variable Expansion ---
            if (itemData.useVariables) {
                std::string luaTemplateUtf8 = Encoding::wstringToUtf8(itemData.replaceText);
                if (!compileLuaReplaceCode(luaTemplateUtf8)) {
                    return false;
                }

                LuaVariables vars;
                // Fill vars struct with context for the current match.
                int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
                int previousLineStartPos = (currentLineIndex == 0) ? 0 : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));
                setLuaFileVars(vars);
                if (context.isColumnMode) {
                    ColumnInfo columnInfo = getColumnInfo(searchResult.pos);
                    vars.COL = static_cast<int>(columnInfo.startColumnIndex);
                }
                vars.CNT = 1;
                vars.LCNT = 1;
                vars.APOS = static_cast<int>(searchResult.pos) + 1;
                vars.LINE = currentLineIndex + 1;
                vars.LPOS = static_cast<int>(searchResult.pos) - previousLineStartPos + 1;

                // Convert the MATCH variable to UTF-8 for Lua if the document is ANSI.
                vars.MATCH = searchResult.foundText;
                if (documentCodepage != SC_CP_UTF8) {
                    vars.MATCH = Encoding::wstringToUtf8(Encoding::ansiToWString(vars.MATCH, documentCodepage));
                }

                if (!resolveLuaSyntax(luaTemplateUtf8, vars, skipReplace, itemData.regex)) {
                    return false;
                }

                // Convert the result from Lua (UTF-8) to the final document codepage.
                finalReplaceText = convertAndExtendW(Encoding::utf8ToWString(luaTemplateUtf8), itemData.extended, documentCodepage);
            }
            else {
                // Case without variables: convert once using the safe helper.
                finalReplaceText = convertAndExtendW(itemData.replaceText, itemData.extended, documentCodepage);
            }

            // --- Final Replacement Execution ---
            if (!skipReplace) {
                newPos = itemData.regex
                    ? performRegexReplace(finalReplaceText, searchResult.pos, searchResult.length)
                    : performReplace(finalReplaceText, searchResult.pos, searchResult.length);
                newPos = ensureForwardProgress(newPos, searchResult);
                send(SCI_SETSEL, newPos, newPos);
                if (itemIndex != SIZE_MAX) {
                    updateCountColumns(itemIndex, -2, 1);
                }
                return true; // A replacement was made.
            }
        }

        // Replacement was skipped by Lua logic.
        newPos = ensureForwardProgress(searchResult.pos + searchResult.length, searchResult);
        send(SCI_SETSEL, newPos, newPos);
    }
    return false; // No replacement was made.
}

bool MultiReplace::replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount, size_t itemIndex)
{
    if (itemData.findText.empty() && !itemData.useVariables) {
        findCount = replaceCount = 0;
        return true;
    }

    // Get the document's codepage once at the beginning.
    const int documentCodepage = getCurrentDocCodePage();

    // --- Search bytes MUST match document codepage---
    SearchContext context;
    context.findText = convertAndExtendW(itemData.findText, itemData.extended);
    context.searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED;
    context.isSelectionMode = IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED;
    context.retrieveFoundText = itemData.useVariables;
    context.highlightMatch = false;

    send(SCI_SETSEARCHFLAGS, context.searchFlags);

    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    Sci_Position startPos = computeAllStartPos(context, wrapAroundEnabled, allFromCursorEnabled);

    SearchResult searchResult = performSearchForward(context, startPos);

    // --- Replace at matches---
    bool useMatchList = IsDlgButtonChecked(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX) == BST_CHECKED;
    std::vector<int> matchList;
    if (useMatchList) {
        std::wstring sel = getTextFromDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT);
        if (sel.empty()) {
            showStatusMessage(LM.get(L"status_missing_match_selection"), MessageStatus::Error);
            return false;
        }
        matchList = parseNumberRanges(sel, LM.get(L"status_invalid_range_in_match_data"));
        if (matchList.empty()) return false;
    }

    // --- Prepare replacement text template (only if needed for Lua) ---
    std::string luaTemplateUtf8;
    if (itemData.useVariables) {
        luaTemplateUtf8 = Encoding::wstringToUtf8(itemData.replaceText);
        if (!compileLuaReplaceCode(luaTemplateUtf8)) {
            return false;
        }
    }

    std::string fixedReplace;
    if (!itemData.useVariables) {
        fixedReplace = convertAndExtendW(itemData.replaceText, itemData.extended, documentCodepage);
    }

    int prevLineIdx = -1;
    int lineFindCount = 0;

    // --- Main replacement loop ---
    while (searchResult.pos >= 0)
    {
        bool skipReplace = false;
        ++findCount;
        if (itemIndex != SIZE_MAX) updateCountColumns(itemIndex, findCount);

        const bool replaceThisHit =
            !useMatchList || (std::find(matchList.begin(), matchList.end(), findCount) != matchList.end());

        Sci_Position nextPos; // declared before both branches use it

        // This block contains the core replacement logic, structured to be consistent with replaceOne.
        {
            std::string finalReplaceText; // Will hold the final text in the document's native encoding.

            // --- Lua Variable Expansion ---
            if (itemData.useVariables) {
                std::string luaWorkingUtf8 = luaTemplateUtf8; // Use a fresh copy of the template.
                LuaVariables vars;
                // (Fill vars struct - this part is identical in both functions)
                int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
                int previousLineStartPos = (currentLineIndex == 0) ? 0 : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));
                setLuaFileVars(vars);
                if (context.isColumnMode) {
                    ColumnInfo columnInfo = getColumnInfo(searchResult.pos);
                    vars.COL = static_cast<int>(columnInfo.startColumnIndex);
                }
                if (currentLineIndex != prevLineIdx) { lineFindCount = 0; prevLineIdx = currentLineIndex; }
                ++lineFindCount;
                vars.CNT = findCount;
                vars.LCNT = lineFindCount;
                vars.APOS = static_cast<int>(searchResult.pos) + 1;
                vars.LINE = currentLineIndex + 1;
                vars.LPOS = static_cast<int>(searchResult.pos) - previousLineStartPos + 1;

                // Convert the MATCH variable to UTF-8 for Lua if the document is ANSI.
                vars.MATCH = searchResult.foundText;
                if (documentCodepage != SC_CP_UTF8) {
                    vars.MATCH = Encoding::wstringToUtf8(Encoding::ansiToWString(vars.MATCH, documentCodepage));
                }

                if (!resolveLuaSyntax(luaWorkingUtf8, vars, skipReplace, itemData.regex)) {
                    return false;
                }

                // Convert the result from Lua (UTF-8) to the final document codepage.
                if (replaceThisHit && !skipReplace) {
                    finalReplaceText = convertAndExtendW(Encoding::utf8ToWString(luaWorkingUtf8), itemData.extended, documentCodepage);
                }
            }

            if (replaceThisHit && !skipReplace) {
                if (!itemData.useVariables) {
                    nextPos = itemData.regex
                        ? performRegexReplace(fixedReplace, searchResult.pos, searchResult.length)
                        : performReplace(fixedReplace, searchResult.pos, searchResult.length);
                }
                else {
                    nextPos = itemData.regex
                        ? performRegexReplace(finalReplaceText, searchResult.pos, searchResult.length)
                        : performReplace(finalReplaceText, searchResult.pos, searchResult.length);
                }

                ++replaceCount;
                if (itemIndex != SIZE_MAX) updateCountColumns(itemIndex, -1, replaceCount);
                context.docLength = send(SCI_GETLENGTH); // Update doc length after modification.
            }
            else {
                nextPos = searchResult.pos + searchResult.length;
                send(SCI_SETSELECTIONSTART, nextPos, 0);
                send(SCI_SETSELECTIONEND, nextPos, 0);
            }
            nextPos = ensureForwardProgress(nextPos, searchResult);
            searchResult = performSearchForward(context, nextPos);
        }
    }
    return true;
}

Sci_Position MultiReplace::performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    send(SCI_SETTARGETRANGE, pos, pos + length);

    return pos + send(
        SCI_REPLACETARGET,
        replaceTextUtf8.size(),
        reinterpret_cast<sptr_t>(replaceTextUtf8.c_str())
    );
}

Sci_Position MultiReplace::performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    send(SCI_SETTARGETRANGE, pos, pos + length);

    sptr_t replacedLen = send(
        SCI_REPLACETARGETRE,
        static_cast<WPARAM>(-1),
        reinterpret_cast<sptr_t>(replaceTextUtf8.c_str())
    );

    Sci_Position newPos = pos + static_cast<Sci_Position>(replacedLen);
    return newPos;
}

bool MultiReplace::preProcessListForReplace(bool highlight) {
    if (!replaceListData.empty()) {
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled && replaceListData[i].useVariables) {
                if (replaceListData[i].findText.empty()) {
                    if (highlight) {
                        selectListItem(i);  // Highlight the list entry
                    }
                    std::string localReplaceTextUtf8 = Encoding::wstringToUtf8(replaceListData[i].replaceText);

                    // Compile the Lua code once and cache it
                    if (!compileLuaReplaceCode(localReplaceTextUtf8)) {
                        return false;
                    }

                    bool skipReplace = false;
                    LuaVariables vars;
                    setLuaFileVars(vars);   // Setting FNAME and FPATH
                    if (!resolveLuaSyntax(localReplaceTextUtf8, vars, skipReplace, replaceListData[i].regex)) {
                        return false;
                    }
                }
            }
        }
    }
    return true;
}

SelectionInfo MultiReplace::getSelectionInfo(bool isBackward) {
    // Get the number of selections
    LRESULT selectionCount = send(SCI_GETSELECTIONS, 0, 0);
    Sci_Position selectedStart = 0;
    Sci_Position correspondingEnd = 0;
    std::vector<SelectionRange> selections; // Store all selections for sorting

    if (selectionCount > 0) {
        selections.resize(selectionCount);

        // Retrieve all selections
        for (LRESULT i = 0; i < selectionCount; ++i) {
            selections[i].start = send(SCI_GETSELECTIONNSTART, i, 0);
            selections[i].end = send(SCI_GETSELECTIONNEND, i, 0);
        }

        // Sort selections based on direction
        if (isBackward) {
            std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) {
                return a.start > b.start;
                });
        }
        else {
            std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) {
                return a.start < b.start;
                });
        }

        // Set the start and end based on the sorted selection
        selectedStart = selections[0].start;
        correspondingEnd = selections[0].end;
    }
    else {
        // No selection case: use current cursor position
        selectedStart = send(SCI_GETCURRENTPOS, 0, 0);
        correspondingEnd = selectedStart;
    }

    // Calculate the selection length
    Sci_Position selectionLength = correspondingEnd - selectedStart;
    return SelectionInfo{ selectedStart, correspondingEnd, selectionLength };
}

void MultiReplace::captureLuaGlobals(lua_State* L) {
    lua_pushglobaltable(L);
    lua_pushnil(L);
    while (lua_next(L, -2) != 0) {
        const char* key = lua_tostring(L, -2);
        LuaVariable luaVar;
        luaVar.name = key;

        int type = lua_type(L, -1);
        if (type == LUA_TNUMBER) {
            luaVar.type = LuaVariableType::Number;
            luaVar.numberValue = lua_tonumber(L, -1);
        }
        else if (type == LUA_TSTRING) {
            luaVar.type = LuaVariableType::String;
            luaVar.stringValue = lua_tostring(L, -1);
        }
        else if (type == LUA_TBOOLEAN) {
            luaVar.type = LuaVariableType::Boolean;
            luaVar.booleanValue = lua_toboolean(L, -1);
        }
        else {
            // Skipping unknown types
            lua_pop(L, 1);
            continue;
        }

        globalLuaVariablesMap[key] = luaVar;
        lua_pop(L, 1);
    }

    lua_pop(L, 1);
}

std::string MultiReplace::escapeForRegex(const std::string& input) {
    std::string escaped;
    for (char c : input) {
        switch (c) {
        case '\\':
        case '^': case '$': case '.': case '|':
        case '?': case '*': case '+': case '(': case ')':
        case '[': case ']': case '{': case '}':
            escaped.push_back('\\'); // Prepend a backslash to escape regex special chars
            escaped.push_back(c);
            break;
        default:
            escaped.push_back(c);
        }
    }
    return escaped;
}

void MultiReplace::setLuaVariable(lua_State* L, const std::string& varName, std::string value) {
    // Check if the input string is a number
    bool isNumber = normalizeAndValidateNumber(value);
    if (isNumber) {
        double doubleVal = std::stod(value);
        int intVal = static_cast<int>(doubleVal);
        if (doubleVal == static_cast<double>(intVal)) {
            lua_pushinteger(L, intVal); // Push as integer if value is integral
        }
        else {
            lua_pushnumber(L, doubleVal); // Push as floating-point number otherwise
        }
    }
    else {
        lua_pushstring(L, value.c_str());
    }
    lua_setglobal(L, varName.c_str()); // Set the global variable in Lua
}

void MultiReplace::updateFilePathCache(const std::filesystem::path* explicitPath) {
    if (explicitPath) {
        // A specific path was provided (during "Replace in Files"). Use it.
        // Convert reliably to UTF-8 for Lua.
        cachedFilePath = Encoding::wstringToUtf8(explicitPath->wstring());
        cachedFileName = Encoding::wstringToUtf8(explicitPath->filename().wstring());
    }
    else {
        // No specific path was provided. Default to the active Notepad++ document.
        wchar_t filePathBuffer[MAX_PATH] = { 0 };
        wchar_t fileNameBuffer[MAX_PATH] = { 0 };
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(filePathBuffer));
        ::SendMessage(nppData._nppHandle, NPPM_GETFILENAME, MAX_PATH, reinterpret_cast<LPARAM>(fileNameBuffer));

        // Convert reliably to UTF-8 for Lua.
        cachedFilePath = Encoding::wstringToUtf8(std::wstring(filePathBuffer));
        cachedFileName = Encoding::wstringToUtf8(std::wstring(fileNameBuffer));
    }
}

void MultiReplace::setLuaFileVars(LuaVariables& vars) {
    // Use cached file path if valid
    if (!cachedFilePath.empty() &&
        (cachedFilePath.find('\\') != std::string::npos || cachedFilePath.find('/') != std::string::npos)) {
        vars.FPATH = cachedFilePath;
    }
    else {
        vars.FPATH.clear();
    }

    // Use cached file name
    vars.FNAME = cachedFileName;
}

bool MultiReplace::initLuaState()
{
    // Reset the Lua state if it already exists.
    if (_luaState) {
        lua_close(_luaState);
        _luaState = nullptr;
    }

    // Invalidate Lua code cache
    _lastCompiledLuaCode.clear();
    _luaCompiledReplaceRef = LUA_NOREF;

    // Create a new Lua state.
    _luaState = luaL_newstate();
    if (!_luaState) {
        MessageBox(nppData._nppHandle, L"Failed to create Lua state", L"Lua Error", MB_OK | MB_ICONERROR);
        return false;
    }

    luaL_openlibs(_luaState);

    if (luaSafeModeEnabled) {
        applyLuaSafeMode(_luaState);
    }

    // Load Lua source code directly (fallback mode)
    if (luaL_loadstring(_luaState, luaSourceCode) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        MessageBox(nppData._nppHandle, Encoding::utf8ToWString(errMsg).c_str(), L"Lua Script Load Error", MB_OK | MB_ICONERROR);
        lua_pop(_luaState, 1);
        lua_close(_luaState);
        _luaState = nullptr;
        return false;
    }

    // Execute the loaded Lua code.
    if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        MessageBox(nppData._nppHandle, Encoding::utf8ToWString(errMsg).c_str(), L"Lua Execution Error", MB_OK | MB_ICONERROR);
        lua_pop(_luaState, 1);
        lua_close(_luaState);
        _luaState = nullptr;
        return false;
    }

    // Any Lua call to safeLoadFileSandbox(...) invokes the C++ code
    lua_pushcfunction(_luaState, &MultiReplace::safeLoadFileSandbox);
    lua_setglobal(_luaState, "safeLoadFileSandbox");

    return true;
}

bool MultiReplace::compileLuaReplaceCode(const std::string& luaCode)
{
    // Compile only if changed or not compiled yet
    if (luaCode == _lastCompiledLuaCode && _luaCompiledReplaceRef != LUA_NOREF) {
        return true; // already compiled, reuse
    }

    // Remove old reference if exists
    if (_luaCompiledReplaceRef != LUA_NOREF) {
        luaL_unref(_luaState, LUA_REGISTRYINDEX, _luaCompiledReplaceRef);
        _luaCompiledReplaceRef = LUA_NOREF;
    }

    // Compile Lua code
    if (luaL_loadstring(_luaState, luaCode.c_str()) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        if (errMsg && isLuaErrorDialogEnabled) {
            std::wstring error_message = Encoding::utf8ToWString(errMsg);
            MessageBox(nppData._nppHandle,
                error_message.c_str(),
                LM.get(L"msgbox_title_use_variables_syntax_error").c_str(),
                MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        }
        lua_pop(_luaState, 1); // pop error message
        return false;
    }

    // Store compiled Lua chunk
    _luaCompiledReplaceRef = luaL_ref(_luaState, LUA_REGISTRYINDEX);
    _lastCompiledLuaCode = luaCode; // Cache the compiled Lua code
    return true;
}

bool MultiReplace::resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex) 
{
    // 1) Stack-checkpoint
    const int stackBase = lua_gettop(_luaState);
    auto restoreStack = [this, stackBase]() { lua_settop(_luaState, stackBase); };

    // 2) Ensure Lua state exists
    if (!_luaState) { return false; }

    // 3) Numeric globals
    lua_pushinteger(_luaState, vars.CNT);  lua_setglobal(_luaState, "CNT");
    lua_pushinteger(_luaState, vars.LCNT); lua_setglobal(_luaState, "LCNT");
    lua_pushinteger(_luaState, vars.LINE); lua_setglobal(_luaState, "LINE");
    lua_pushinteger(_luaState, vars.LPOS); lua_setglobal(_luaState, "LPOS");
    lua_pushinteger(_luaState, vars.APOS); lua_setglobal(_luaState, "APOS");
    lua_pushinteger(_luaState, vars.COL);  lua_setglobal(_luaState, "COL");

    // 4) String globals
    setLuaVariable(_luaState, "FPATH", vars.FPATH);
    setLuaVariable(_luaState, "FNAME", vars.FNAME);
    setLuaVariable(_luaState, "MATCH", vars.MATCH);

    // 5) REGEX flag
    lua_pushboolean(_luaState, regex);
    lua_setglobal(_luaState, "REGEX");

    // 6) CAP# globals (regex only)
    std::vector<std::string> capNames;
    if (regex) {
        for (int i = 1; i <= MAX_CAP_GROUPS; ++i) {
            sptr_t len = send(SCI_GETTAG, i, 0, true);
            if (len < 0) { break; }

            std::string capVal;
            if (len > 0) {
                std::vector<char> buf(len + 1, '\0');
                if (send(SCI_GETTAG, i,
                         reinterpret_cast<sptr_t>(buf.data()), false) >= 0)
                    capVal.assign(buf.data());
            }
            std::string capName = "CAP" + std::to_string(i);
            setLuaVariable(_luaState, capName, capVal);
            capNames.push_back(capName);
        }
    }

    // 7) Run pre-compiled chunk
    lua_rawgeti(_luaState, LUA_REGISTRYINDEX, _luaCompiledReplaceRef);
    if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
        const char* err = lua_tostring(_luaState, -1);
        if (err && isLuaErrorDialogEnabled) {
            MessageBox(nppData._nppHandle,
                       Encoding::utf8ToWString(err).c_str(),
                       LM.get(L"msgbox_title_use_variables_syntax_error").c_str(),
                       MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        }
        restoreStack();
        return false;
    }

    // 8) resultTable
    lua_getglobal(_luaState, "resultTable");
    if (!lua_istable(_luaState, -1)) {
        if (isLuaErrorDialogEnabled) {
            std::wstring msg =
                LM.get(L"msgbox_use_variables_execution_error",
                       { Encoding::utf8ToWString(inputString.c_str()) });
            MessageBox(nppData._nppHandle, msg.c_str(),
                       LM.get(L"msgbox_title_use_variables_execution_error").c_str(),
                       MB_OK);
        }
        restoreStack();
        return false;
    }

    // 9) result & skip
    lua_getfield(_luaState, -1, "result");                 // push result
    if (lua_isnil(_luaState, -1)) {
        inputString.clear();
    } else if (lua_isstring(_luaState, -1) || lua_isnumber(_luaState, -1)) {
        std::string res = lua_tostring(_luaState, -1);
        if (regex) { res = escapeForRegex(res); }
        inputString = res;
    }
    lua_pop(_luaState, 1);                                 // pop result

    lua_getfield(_luaState, -1, "skip");                   // push skip
    skip = lua_isboolean(_luaState, -1) && lua_toboolean(_luaState, -1);
    lua_pop(_luaState, 1);                                 // pop skip

    // (resultTable left on stack until the very end, then dropped by restoreStack)

    // 10) CAP variable dump & cleanup
    std::string capVariablesStr;
    for (const auto& capName : capNames) {
        lua_getglobal(_luaState, capName.c_str());         // push CAP value

        if (lua_isnumber(_luaState, -1)) {
            double n = lua_tonumber(_luaState, -1);
            std::ostringstream os; os << std::fixed << std::setprecision(8) << n;
            capVariablesStr += capName + "\tNumber\t" + os.str() + "\n\n";
        } else if (lua_isboolean(_luaState, -1)) {
            bool b = lua_toboolean(_luaState, -1);
            capVariablesStr += capName + "\tBoolean\t" + (b ? "true" : "false") + "\n\n";
        } else if (lua_isstring(_luaState, -1)) {
            capVariablesStr += capName + "\tString\t" + lua_tostring(_luaState, -1) + "\n\n";
        } else {
            capVariablesStr += capName + "\t<nil>\n\n";
        }
        lua_pop(_luaState, 1);                             // pop CAP value

        lua_pushnil(_luaState);                            // clear global
        lua_setglobal(_luaState, capName.c_str());
    }

    // 11) DEBUG flag & window
    lua_getglobal(_luaState, "DEBUG");                     // push DEBUG
    bool debugOn = lua_isboolean(_luaState, -1) && lua_toboolean(_luaState, -1);
    lua_pop(_luaState, 1);                                 // pop DEBUG

    if (debugOn) {
        globalLuaVariablesMap.clear();
        captureLuaGlobals(_luaState);

        std::string globalsStr = "Global Lua variables:\n\n";
        for (const auto& p : globalLuaVariablesMap) {
            const LuaVariable& v = p.second;
            if (v.type == LuaVariableType::String) {
                globalsStr += v.name + "\tString\t" + v.stringValue + "\n\n";
            } else if (v.type == LuaVariableType::Number) {
                std::ostringstream os; os << std::fixed << std::setprecision(8) << v.numberValue;
                globalsStr += v.name + "\tNumber\t" + os.str() + "\n\n";
            } else if (v.type == LuaVariableType::Boolean) {
                globalsStr += v.name + "\tBoolean\t" + (v.booleanValue ? "true" : "false") + "\n\n";
            }
        }

        refreshUIListView();
        int resp = ShowDebugWindow(capVariablesStr + globalsStr);
        if (resp == 3) { restoreStack(); return false; }   // “Stop”
        if (resp == -1) { restoreStack(); return false; }  // window closed
    }

    // 12) Success
    restoreStack();
    return true;
}

int MultiReplace::safeLoadFileSandbox(lua_State* L)
{
    const char* path = luaL_checkstring(L, 1);

    // Read file (your existing UTF-8/ANSI logic)
    std::wstring wpath = Encoding::utf8ToWString(path);
    std::ifstream in(std::filesystem::path(wpath), std::ios::binary);
    if (!in) {
        lua_pushboolean(L, false);
        lua_pushfstring(L, "Cannot open file: %s", path);
        return 2;
    }
    std::string raw((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    in.close();

    bool rawIsUtf8 = Encoding::isValidUtf8(raw);
    std::string utf8_buf;
    if (rawIsUtf8) {
        utf8_buf = std::move(raw);
    }
    else {
        int wlen = MultiByteToWideChar(CP_ACP, 0, raw.data(), (int)raw.size(), nullptr, 0);
        std::wstring wide(wlen, L'\0');
        MultiByteToWideChar(CP_ACP, 0, raw.data(), (int)raw.size(), &wide[0], wlen);

        int u8len = WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, nullptr, 0, nullptr, nullptr);
        utf8_buf.resize(u8len);
        WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, &utf8_buf[0], u8len, nullptr, nullptr);
    }

    // Strip UTF-8 BOM if present
    if (utf8_buf.size() >= 3 &&
        (unsigned char)utf8_buf[0] == 0xEF &&
        (unsigned char)utf8_buf[1] == 0xBB &&
        (unsigned char)utf8_buf[2] == 0xBF)
    {
        utf8_buf.erase(0, 3);
    }

    // Load as text chunk only (avoid binary precompiled chunks)
#if LUA_VERSION_NUM >= 503
    if (luaL_loadbufferx(L, utf8_buf.data(), utf8_buf.size(), path, "t") != LUA_OK)
#else
    if (luaL_loadbuffer(L, utf8_buf.data(), utf8_buf.size(), path) != LUA_OK)
#endif
    {
        lua_pushboolean(L, false);
        lua_pushstring(L, lua_tostring(L, -2)); // copy error
        lua_remove(L, -3); // remove original error under the two pushes
        return 2;
    }

    if (lua_pcall(L, 0, 1, 0) != LUA_OK) {
        lua_pushboolean(L, false);
        lua_pushstring(L, lua_tostring(L, -2));
        lua_remove(L, -3);
        return 2;
    }

    // Success: stack has [ table ]
    lua_pushboolean(L, true);
    lua_insert(L, -2); // [ true, table ]
    return 2;
}

void MultiReplace::applyLuaSafeMode(lua_State* L)
{
    auto removeGlobal = [&](const char* name) {
        lua_pushnil(L);
        lua_setglobal(L, name);
        };

    // Remove dangerous base functions
    removeGlobal("dofile");
    removeGlobal("load");
    removeGlobal("loadfile");
    removeGlobal("require");
    removeGlobal("collectgarbage");

    // Remove whole libraries
    removeGlobal("os");
    removeGlobal("io");
    removeGlobal("package");
    removeGlobal("debug");

    // Keep string/table/math/utf8/base intact.
}

Sci_Position MultiReplace::computeAllStartPos(const SearchContext& context, bool wrapEnabled, bool fromCursorEnabled)
{
    SelectionInfo selInfo = getSelectionInfo(false);
    if (context.isSelectionMode) {
        return selInfo.startPos;
    }
    const Sci_Position caretPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
    return wrapEnabled ? 0 : (fromCursorEnabled ? caretPos : 0);
}

#pragma endregion


#pragma region Replace in Files

bool MultiReplace::selectDirectoryDialog(HWND owner, std::wstring& outPath)
{
    // Initialize COM and store the result.
    HRESULT hrInit = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hrInit)) {
        // If critical COM error, we cannot proceed at all.
        return false;
    }

    IFileDialog* pfd = nullptr;
    // Use a different HRESULT variable for the instance creation.
    HRESULT hrCreate = CoCreateInstance(
        CLSID_FileOpenDialog, nullptr,
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd)
    );

    // Only proceed if the dialog instance was created successfully.
    if (SUCCEEDED(hrCreate)) {
        // Tell it to pick folders only.
        FILEOPENDIALOGOPTIONS opts;
        pfd->GetOptions(&opts);
        pfd->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

        // Show the dialog.
        if (SUCCEEDED(pfd->Show(owner)))
        {
            IShellItem* psi = nullptr;
            if (SUCCEEDED(pfd->GetResult(&psi)) && psi)
            {
                // Retrieve the selected folder’s file system path.
                PWSTR pszPath = nullptr;
                if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath)))
                {
                    outPath = pszPath;
                    CoTaskMemFree(pszPath);
                }
                psi->Release();
            }
        }
        pfd->Release();
    }

    // Uninitialize COM only if our initial call was the one that actually
    // initialized it (returned S_OK).
    if (hrInit == S_OK) {
        CoUninitialize();
    }

    return !outPath.empty();
}

bool MultiReplace::handleBrowseDirectoryButton()
{
    std::wstring dir;
    if (selectDirectoryDialog(_hSelf, dir))
    {
        // User picked one — set the combo-box text & keep history
        SetDlgItemTextW(_hSelf, IDC_DIR_EDIT, dir.c_str());
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), dir);
    }
    return true;  // always return TRUE so the dialog proc knows we handled it
}

void MultiReplace::handleReplaceInFiles() {
    HiddenSciGuard guard;
    if (!guard.create()) {
        showStatusMessage(LM.get(L"status_error_hidden_buffer"), MessageStatus::Error);
        return;
    }

    // Read all inputs once at the beginning.
    auto wDir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);
    auto wFilter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
    const bool recurse = (IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED);
    const bool hide = (IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED);

    if (wFilter.empty()) {
        wFilter = L"*.*";
        SetDlgItemTextW(_hSelf, IDC_FILTER_EDIT, wFilter.c_str());
    }
    guard.parseFilter(wFilter);

    if (wDir.empty() || !std::filesystem::exists(wDir)) {
        showStatusMessage(LM.get(L"status_error_invalid_directory"), MessageStatus::Error);
        return;
    }

    std::vector<std::filesystem::path> files;
    try {
        namespace fs = std::filesystem;
        if (recurse) {
            for (auto& e : fs::recursive_directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
        else {
            for (auto& e : fs::directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
    }
    catch (const std::exception& ex) {
        std::wstring wideReason = Encoding::utf8ToWString(ex.what());
        showStatusMessage(LM.get(L"status_error_scanning_directory", { wideReason }), MessageStatus::Error);
        return;
    }

    if (files.empty()) {
        MessageBox(_hSelf, LM.getLPW(L"msgbox_no_files"), LM.getLPW(L"msgbox_title_confirm"), MB_OK);
        return;
    }

    // --- Confirmation Dialog Setup ---
    // Manually shorten the directory path to prevent ugly wrapping.
    HDC dialogHdc = GetDC(_hSelf);
    HFONT dialogHFont = (HFONT)SendMessage(_hSelf, WM_GETFONT, 0, 0);
    SelectObject(dialogHdc, dialogHFont);
    std::wstring shortenedDirectory = getShortenedFilePath(wDir, 400, dialogHdc);
    ReleaseDC(_hSelf, dialogHdc);

    std::wstring message = LM.get(L"msgbox_confirm_replace_in_files", { std::to_wstring(files.size()), shortenedDirectory, wFilter });
    if (MessageBox(_hSelf, message.c_str(), LM.getLPW(L"msgbox_title_confirm"), MB_OKCANCEL | MB_SETFOREGROUND) != IDOK)
        return;

    // RAII-based UI State Management
    BatchUIGuard uiGuard(this, _hSelf);
    _isCancelRequested = false;

    if (useListEnabled) {
        resetCountColumns();
    }

    std::vector<int> listFindTotals;
    std::vector<int> listReplaceTotals;
    if (useListEnabled) {
        listFindTotals.assign(replaceListData.size(), 0);
        listReplaceTotals.assign(replaceListData.size(), 0);
    }

    int total = static_cast<int>(files.size()), idx = 0, changed = 0;
    showStatusMessage(L"Progress: [  0%]", MessageStatus::Info);

    // Per-file binding guard
    struct SciBindingGuard {
        MultiReplace* self;
        HWND oldSci; SciFnDirect oldFn; sptr_t oldData;
        HiddenSciGuard& g;
        SciBindingGuard(MultiReplace* s, HiddenSciGuard& guard) : self(s), g(guard) {
            oldSci = s->_hScintilla; oldFn = s->pSciMsg; oldData = s->pSciWndData;
            s->_hScintilla = g.hSci; s->pSciMsg = g.fn; s->pSciWndData = g.pData;
        }
        ~SciBindingGuard() {
            self->_hScintilla = oldSci; self->pSciMsg = oldFn; self->pSciWndData = oldData;
        }
    };

    bool aborted = false;

    for (const auto& fp : files) {
        MSG m;
        while (::PeekMessage(&m, nullptr, 0, 0, PM_REMOVE)) {
            ::TranslateMessage(&m);
            ::DispatchMessage(&m);
        }

        if (_isShuttingDown) { aborted = true; break; }
        if (_isCancelRequested) { aborted = true; break; }

        ++idx;

        int percent = static_cast<int>((static_cast<double>(idx) / (std::max)(1, total)) * 100.0);
        std::wstring prefix = L"Progress: [" + std::to_wstring(percent) + L"%] ";
        HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
        HDC hdc = GetDC(hStatus);
        HFONT hFont = (HFONT)SendMessage(hStatus, WM_GETFONT, 0, 0);
        SelectObject(hdc, hFont);
        SIZE sz{}; GetTextExtentPoint32W(hdc, prefix.c_str(), (int)prefix.length(), &sz);
        RECT rc{}; GetClientRect(hStatus, &rc);
        int avail = (rc.right - rc.left) - sz.cx;
        std::wstring shortPath = getShortenedFilePath(fp.wstring(), avail, hdc);
        ReleaseDC(hStatus, hdc);
        showStatusMessage(prefix + shortPath, MessageStatus::Info);

        std::string original;
        if (!guard.loadFile(fp, original)) { continue; }

        DWORD attrs = GetFileAttributesW(fp.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) { continue; }

        EncodingInfo enc = detectEncoding(original);
        std::string u8in;
        if (!convertBufferToUtf8(original, enc, u8in)) { continue; }

        // Bind hidden buffer for the file scope
        {
            SciBindingGuard bind(this, guard);

            send(SCI_CLEARALL, 0, 0);
            send(SCI_SETCODEPAGE, SC_CP_UTF8, 0);
            send(SCI_ADDTEXT, (WPARAM)u8in.length(), reinterpret_cast<sptr_t>(u8in.data()));

            handleDelimiterPositions(DelimiterOperation::LoadAll);

            if (!handleReplaceAllButton(false, &fp)) { _isCancelRequested = true; aborted = true; }

            // Keep list counters in sync (per-item find/replace tallies)
            if (useListEnabled) {
                for (size_t i = 0; i < replaceListData.size(); ++i) {
                    if (!replaceListData[i].isEnabled) continue;
                    const int f = replaceListData[i].findCount.empty() ? 0 : std::stoi(replaceListData[i].findCount);
                    const int r = replaceListData[i].replaceCount.empty() ? 0 : std::stoi(replaceListData[i].replaceCount);
                    listFindTotals[i] += f;
                    listReplaceTotals[i] += r;
                }
            }

            // Write back only if content changed
            std::string u8out = guard.getText();  // uses directfunction of hidden buffer  :contentReference[oaicite:6]{index=6}
            if (u8out != u8in) {
                std::string finalBuf;
                if (convertUtf8ToOriginal(u8out, enc, original, finalBuf))
                    if (guard.writeFile(fp, finalBuf)) ++changed;
            }
        }

        if (aborted) break; // ensures RAII restored before leaving loop
    }

    if (useListEnabled) {
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;
            updateCountColumns(i, listFindTotals[i], listReplaceTotals[i]);
        }
        refreshUIListView();
    }

    // status line
    if (!_isShuttingDown) {
        const bool wasCanceled = (_isCancelRequested || aborted);

        std::wstring msg = LM.get(L"status_replace_summary",
            { std::to_wstring(changed), std::to_wstring(files.size()) });
        if (wasCanceled) {
            msg += L" - " + LM.get(L"status_canceled");
        }

        MessageStatus ms = MessageStatus::Success;
        if (wasCanceled || changed == 0) {
            // neutral if canceled or no changes were written
            ms = MessageStatus::Info;
        }

        showStatusMessage(msg, ms);
    }
    _isCancelRequested = false;


}

bool MultiReplace::convertBufferToUtf8(const std::string& original_buf, const EncodingInfo& enc_info, std::string& utf8_output) {
    const char* data_ptr = original_buf.data() + enc_info.bom_length;
    int data_len = static_cast<int>(original_buf.size() - enc_info.bom_length);
    if (data_len < 0) return false;

    if (enc_info.sc_codepage == SC_CP_UTF8) {
        utf8_output.assign(data_ptr, data_len);
        return true;
    }

    std::wstring wbuf;

    switch (enc_info.sc_codepage) {
    case 1200: { // UTF-16 LE
        if (data_len % 2 != 0) return false;
        wbuf.assign(reinterpret_cast<const wchar_t*>(data_ptr), data_len / sizeof(wchar_t));
        break;
    }
    case 1201: { // UTF-16 BE
        if (data_len % 2 != 0) return false;
        std::string temp_le_buf(data_ptr, data_len);
        for (size_t i = 0; i < temp_le_buf.length(); i += 2) {
            std::swap(temp_le_buf[i], temp_le_buf[i + 1]);
        }
        wbuf.assign(reinterpret_cast<const wchar_t*>(temp_le_buf.data()), temp_le_buf.length() / sizeof(wchar_t));
        break;
    }
    default: { // ANSI
        int wide_len = MultiByteToWideChar(enc_info.sc_codepage, 0, data_ptr, data_len, nullptr, 0);
        if (wide_len <= 0) return false;
        wbuf.resize(wide_len);
        MultiByteToWideChar(enc_info.sc_codepage, 0, data_ptr, data_len, &wbuf[0], wide_len);
        break;
    }
    }

    if (wbuf.empty() && data_len > 0) return false; // Conversion to wstring failed

    int utf8_len = WideCharToMultiByte(CP_UTF8, 0, &wbuf[0], (int)wbuf.size(), nullptr, 0, nullptr, nullptr);
    if (utf8_len <= 0) return false;
    utf8_output.resize(utf8_len);
    WideCharToMultiByte(CP_UTF8, 0, &wbuf[0], (int)wbuf.size(), &utf8_output[0], utf8_len, nullptr, nullptr);

    return true;
}

bool MultiReplace::convertUtf8ToOriginal(const std::string& utf8_input, const EncodingInfo& original_enc_info, const std::string& original_buf_with_bom, std::string& final_output_with_bom) {
    std::string final_output_no_bom;

    if (original_enc_info.sc_codepage == SC_CP_UTF8) {
        final_output_no_bom = utf8_input;
    }
    else {
        // Step 1: Convert the modified UTF-8 string back to a standard wstring (UTF-16 LE on Windows)
        std::wstring wbuf_out;
        int wide_len_out = MultiByteToWideChar(CP_UTF8, 0, utf8_input.c_str(), (int)utf8_input.size(), nullptr, 0);
        if (wide_len_out <= 0) return false;
        wbuf_out.resize(wide_len_out);
        MultiByteToWideChar(CP_UTF8, 0, utf8_input.c_str(), (int)utf8_input.size(), &wbuf_out[0], wide_len_out);

        // Step 2: Convert the wstring to the original target encoding
        switch (original_enc_info.sc_codepage) {
        case 1200: { // Original was UTF-16 LE
            final_output_no_bom.assign(reinterpret_cast<const char*>(wbuf_out.data()), wbuf_out.size() * sizeof(wchar_t));
            break;
        }
        case 1201: { // Original was UTF-16 BE
            // To convert from our wstring (UTF-16 LE) to UTF-16 BE, we must byte-swap each character.
            final_output_no_bom.assign(reinterpret_cast<const char*>(wbuf_out.data()), wbuf_out.size() * sizeof(wchar_t));
            for (size_t i = 0; i < final_output_no_bom.length(); i += 2) {
                std::swap(final_output_no_bom[i], final_output_no_bom[i + 1]);
            }
            break;
        }
        default: { // Original was ANSI
            int final_len = WideCharToMultiByte(original_enc_info.sc_codepage, 0, wbuf_out.c_str(), (int)wbuf_out.size(), nullptr, 0, nullptr, nullptr);
            if (final_len <= 0) return false;
            final_output_no_bom.resize(final_len);
            WideCharToMultiByte(original_enc_info.sc_codepage, 0, wbuf_out.c_str(), (int)wbuf_out.size(), &final_output_no_bom[0], final_len, nullptr, nullptr);
            break;
        }
        }
    }

    // Step 3: Prepend the original BOM
    final_output_with_bom.clear();
    if (original_enc_info.bom_length > 0) {
        final_output_with_bom.append(original_buf_with_bom.data(), original_enc_info.bom_length);
    }
    final_output_with_bom.append(final_output_no_bom);
    return true;
}

EncodingInfo MultiReplace::detectEncoding(const std::string& buffer) {
    if (buffer.empty()) { return { CP_ACP, 0 }; }
    if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) { return { SC_CP_UTF8, 3 }; }
    if (buffer.size() >= 2 && static_cast<unsigned char>(buffer[0]) == 0xFF && static_cast<unsigned char>(buffer[1]) == 0xFE) { return { 1200, 2 }; }
    if (buffer.size() >= 2 && static_cast<unsigned char>(buffer[0]) == 0xFE && static_cast<unsigned char>(buffer[1]) == 0xFF) { return { 1201, 2 }; }
    if (Encoding::isValidUtf8(buffer)) { return { SC_CP_UTF8, 0 }; }
    int check_len = std::min((int)buffer.size(), 512);
    if (check_len > 1) {
        if (check_len % 2 != 0) check_len--;
        int nulls_at_odd_pos = 0;
        for (int i = 1; i < check_len; i += 2) { if (buffer[i] == '\0') { nulls_at_odd_pos++; } }
        double null_ratio = (check_len > 0) ? (static_cast<double>(nulls_at_odd_pos) / (check_len / 2.0)) : 0.0;
        if (null_ratio > 0.40) { return { 1200, 0 }; }
    }
    return { CP_ACP, 0 };
}

#pragma endregion


#pragma region DebugWindow from resolveLuaSyntax

int MultiReplace::ShowDebugWindow(const std::string& message) {
    const int buffer_size = 4096;
    wchar_t wMessage[buffer_size];
    debugWindowResponse = -1;

    // Convert the message from UTF-8 to UTF-16
    int result = MultiByteToWideChar(CP_UTF8, 0, message.c_str(), -1, wMessage, buffer_size);
    if (result == 0) {
        MessageBox(nppData._nppHandle, L"Error converting message", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    static bool isClassRegistered = false;

    if (!isClassRegistered) {
        WNDCLASS wc = { 0 };

        wc.lpfnWndProc = DebugWindowProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = L"DebugWindowClass";
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);

        if (!RegisterClass(&wc)) {
            MessageBox(nppData._nppHandle, L"Error registering class", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
            return -1;
        }

        isClassRegistered = true;
    }

    std::wstringstream formattedMessage;
    std::wstringstream inputMessageStream(wMessage);
    std::wstring line;

    while (std::getline(inputMessageStream, line)) {
        std::wistringstream iss(line);
        std::wstring variable, type, value;

        if (std::getline(iss, variable, L'\t') &&
            std::getline(iss, type, L'\t') &&
            std::getline(iss, value)) {

            // Trim whitespace
            type.erase(0, type.find_first_not_of(L" \t"));
            type.erase(type.find_last_not_of(L" \t") + 1);
            value.erase(0, value.find_first_not_of(L" \t"));
            value.erase(value.find_last_not_of(L" \t") + 1);

            if (type == L"Number") {
                try {
                    double num = std::stod(value);

                    // If the number is integral, output as an integer
                    if (num == std::floor(num)) {
                        value = std::to_wstring(static_cast<long long>(num));
                    }
                    else {
                        // Format the number with fixed precision (up to 6 decimals)
                        std::wstringstream numStream;
                        numStream << std::fixed << std::setprecision(8) << num;
                        std::wstring formatted = numStream.str();
                        // Remove trailing zeros
                        size_t pos = formatted.find_last_not_of(L'0');
                        if (pos != std::wstring::npos) {
                            formatted.erase(pos + 1);
                        }
                        // If the last character is a decimal point, remove it
                        if (!formatted.empty() && formatted.back() == L'.') {
                            formatted.pop_back();
                        }
                        value = formatted;
                    }
                }
                catch (...) {
                    // If conversion fails, retain the original value
                }
            }

            formattedMessage << variable << L"\t" << type << L"\t" << value << L"\n";
        }
    }

    std::wstring finalMessage = formattedMessage.str();

    // Use the saved position and size if set, otherwise use default position and size
    int width = debugWindowSizeSet ? debugWindowSize.cx : sx(260); // Set initial width
    int height = debugWindowSizeSet ? debugWindowSize.cy : sy(400); // Set initial height
    int x = debugWindowPositionSet ? debugWindowPosition.x : (GetSystemMetrics(SM_CXSCREEN) - width) / 2; // Center horizontally
    int y = debugWindowPositionSet ? debugWindowPosition.y : (GetSystemMetrics(SM_CYSCREEN) - height) / 2; // Center vertically

    HWND hwnd = CreateWindowEx(
        WS_EX_TOPMOST, // Always on top
        L"DebugWindowClass",
        L"Debug Information",
        WS_OVERLAPPEDWINDOW,
        x, y, width, height,
        nppData._nppHandle, NULL, hInstance, (LPVOID)finalMessage.c_str()
    );

    if (hwnd == NULL) {
        MessageBoxW(nppData._nppHandle, L"Error creating window", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    // Save the handle of the debug window
    hDebugWnd = hwnd;

    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg = { 0 };

    // Scintilla needs seperate key handling
    while (IsWindow(hwnd))
    {
        // Check for the global shutdown signal from Notepad++
        if (_isShuttingDown) {
            // If N++ is closing, we must break our loop immediately to allow a clean shutdown.
            DestroyWindow(hwnd); // This will post a WM_QUIT message and terminate the loop.
            debugWindowResponse = 3; // Simulate a "Stop" press.
            continue;
        }

        // Process messages in a non-blocking way
        if (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT) {
                break; // Exit loop on WM_QUIT
            }
            if (!IsDialogMessage(hwnd, &msg)) {
                // Check if the Debug window is not focused and forward key combinations to Notepad++
                if (GetForegroundWindow() != hwnd) {
                    if (msg.message == WM_KEYDOWN && (GetKeyState(VK_CONTROL) & 0x8000)) {
                        // Handle Ctrl-C, Ctrl-V, and Ctrl-X
                        if (msg.wParam == 'C') {
                            SendMessage(nppData._scintillaMainHandle, SCI_COPY, 0, 0);
                            continue;
                        }
                        else if (msg.wParam == 'V') {
                            SendMessage(nppData._scintillaMainHandle, SCI_PASTE, 0, 0);
                            continue;
                        }
                        else if (msg.wParam == 'X') {
                            SendMessage(nppData._scintillaMainHandle, SCI_CUT, 0, 0);
                            continue;
                        }
                        // Handle Ctrl-U for lowercase
                        else if (msg.wParam == 'U' && !(GetKeyState(VK_SHIFT) & 0x8000)) {
                            SendMessage(nppData._scintillaMainHandle, SCI_LOWERCASE, 0, 0);
                            continue;
                        }
                        // Handle Ctrl-Shift-U for uppercase
                        else if (msg.wParam == 'U' && (GetKeyState(VK_SHIFT) & 0x8000)) {
                            SendMessage(nppData._scintillaMainHandle, SCI_UPPERCASE, 0, 0);
                            continue;
                        }
                    }
                }
                // Process other messages normally
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
        else
        {
            // If there are no messages, wait for the next one without consuming 100% CPU
            WaitMessage();
        }
    }

    // Get the window position and size before destroying it
    RECT rect;
    if (GetWindowRect(hwnd, &rect)) {
        debugWindowPosition.x = rect.left;
        debugWindowPosition.y = rect.top;
        debugWindowPositionSet = true;

        debugWindowSize.cx = rect.right - rect.left;
        debugWindowSize.cy = rect.bottom - rect.top;
        debugWindowSizeSet = true;
    }

    DestroyWindow(hwnd);
    hDebugWnd = NULL; // Reset the handle after the window is destroyed

    if (debugWindowResponse == 2) // ID of "Next" is 2
    {
        MSG m;
        // Check for and remove any pending mouse messages from the queue.
        // This prevents a lingering WM_LBUTTONUP from re-triggering the "Next" button
        // on the next debug window that will be created.
        while (PeekMessage(&m, nullptr, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
        {
            // Do nothing with the message, effectively discarding it.
        }
    }

    return debugWindowResponse;
}

LRESULT CALLBACK MultiReplace::DebugWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hListView;
    static HWND hNextButton;
    static HWND hStopButton;
    static HWND hCopyButton;

    switch (msg) {
    case WM_CREATE: {
        hListView = CreateWindowW(WC_LISTVIEW, L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            10, 10, 360, 140,
            hwnd, (HMENU)1, NULL, NULL);

        // Initialize columns
        LVCOLUMN lvCol;
        lvCol.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        lvCol.pszText = L"Variable";
        lvCol.cx = 120;
        ListView_InsertColumn(hListView, 0, &lvCol);
        lvCol.pszText = L"Type";
        lvCol.cx = 80;
        ListView_InsertColumn(hListView, 1, &lvCol);
        lvCol.pszText = L"Value";
        lvCol.cx = 160;
        ListView_InsertColumn(hListView, 2, &lvCol);

        hNextButton = CreateWindowW(L"BUTTON", L"Next",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            50, 160, 80, 30,
            hwnd, (HMENU)2, NULL, NULL);

        hStopButton = CreateWindowW(L"BUTTON", L"Stop",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            150, 160, 80, 30,
            hwnd, (HMENU)3, NULL, NULL);

        hCopyButton = CreateWindowW(L"BUTTON", L"Copy",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            250, 160, 80, 30,
            hwnd, (HMENU)4, NULL, NULL);

        // Extract the debug message from lParam
        LPCWSTR debugMessage = reinterpret_cast<LPCWSTR>(reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams);

        // Populate ListView with the debug message
        std::wstringstream ss(debugMessage);
        std::wstring line;
        int itemIndex = 0;
        while (std::getline(ss, line)) {
            std::wstring variable, type, value;
            std::wistringstream iss(line);
            if (std::getline(iss, variable, L'\t') && std::getline(iss, type, L'\t') && std::getline(iss, value)) {
                LVITEM lvItem;
                lvItem.mask = LVIF_TEXT;
                lvItem.iItem = itemIndex;
                lvItem.iSubItem = 0;
                lvItem.pszText = const_cast<LPWSTR>(variable.c_str());
                ListView_InsertItem(hListView, &lvItem);
                ListView_SetItemText(hListView, itemIndex, 1, const_cast<LPWSTR>(type.c_str()));
                ListView_SetItemText(hListView, itemIndex, 2, const_cast<LPWSTR>(value.c_str()));
                ++itemIndex;
            }
        }

        break;
    }

    case WM_SIZE: {
        // Get the new window size
        RECT rect;
        GetClientRect(hwnd, &rect);

        // Calculate new positions for the buttons
        int buttonWidth = 80;
        int buttonHeight = 30;
        int listHeight = rect.bottom - buttonHeight - 40; // 40 for padding

        SetWindowPos(hListView, NULL, 10, 10, rect.right - 20, listHeight, SWP_NOZORDER);
        SetWindowPos(hNextButton, NULL, 10, rect.bottom - buttonHeight - 10, buttonWidth, buttonHeight, SWP_NOZORDER);
        SetWindowPos(hStopButton, NULL, buttonWidth + 20, rect.bottom - buttonHeight - 10, buttonWidth, buttonHeight, SWP_NOZORDER);
        SetWindowPos(hCopyButton, NULL, 2 * (buttonWidth + 20), rect.bottom - buttonHeight - 10, buttonWidth, buttonHeight, SWP_NOZORDER);

        // Adjust the column widths
        ListView_SetColumnWidth(hListView, 0, LVSCW_AUTOSIZE_USEHEADER); // Variable column
        ListView_SetColumnWidth(hListView, 1, LVSCW_AUTOSIZE_USEHEADER); // Type column
        ListView_SetColumnWidth(hListView, 2, LVSCW_AUTOSIZE_USEHEADER); // Value column
        break;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == 2 || LOWORD(wParam) == 3) {
            debugWindowResponse = LOWORD(wParam);

            // Save the window position and size before closing
            RECT rect;
            if (GetWindowRect(hwnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }

            DestroyWindow(hwnd);
        }
        else if (LOWORD(wParam) == 4) {  // Copy button was pressed
            CopyListViewToClipboard(hListView);
        }
        break;

    case WM_CLOSE:
        // Only set debugWindowResponse to -1 if not already set by a button
        if (debugWindowResponse == -1) {
            // Closed by X, Alt+F4, or CloseDebugWindow()
            debugWindowResponse = -1;
        }

        // Save the window position and size before closing
        RECT rect;
        if (GetWindowRect(hwnd, &rect)) {
            debugWindowPosition.x = rect.left;
            debugWindowPosition.y = rect.top;
            debugWindowPositionSet = true;
            debugWindowSize.cx = rect.right - rect.left;
            debugWindowSize.cy = rect.bottom - rect.top;
            debugWindowSizeSet = true;
        }

        DestroyWindow(hwnd);
        break;

    case WM_DESTROY:
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void MultiReplace::CopyListViewToClipboard(HWND hListView) {
    const int itemCount = ListView_GetItemCount(hListView);
    if (itemCount <= 0) {
        return;
    }

    const int columnCount = Header_GetItemCount(ListView_GetHeader(hListView));
    std::wstring clipboardText;
    clipboardText.reserve(static_cast<size_t>(itemCount) * columnCount * 64);

    wchar_t buffer[512];

    for (int i = 0; i < itemCount; ++i) {
        for (int j = 0; j < columnCount; ++j) {
            LVITEMW li{};
            li.iSubItem = j;
            li.cchTextMax = std::size(buffer);
            li.pszText = buffer;

            buffer[0] = L'\0';
            SendMessageW(hListView, LVM_GETITEMTEXTW, (WPARAM)i, (LPARAM)&li);
            buffer[std::size(buffer) - 1] = L'\0';

            int len = lstrlenW(buffer);
            if (len > 0) {
                clipboardText.append(buffer, len);
            }

            if (j < columnCount - 1) {
                clipboardText.push_back(L'\t');
            }
        }
        clipboardText.push_back(L'\n');
    }

    if (clipboardText.empty()) {
        return;
    }

    if (OpenClipboard(nullptr)) {
        EmptyClipboard();
        size_t bytes = (clipboardText.size() + 1) * sizeof(wchar_t);
        HGLOBAL hData = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (hData) {
            void* ptr = GlobalLock(hData);
            if (ptr) {
                memcpy(ptr, clipboardText.c_str(), bytes);
                GlobalUnlock(hData);
                SetClipboardData(CF_UNICODETEXT, hData);
            }
            else {
                GlobalFree(hData);
            }
        }
        CloseClipboard();
    }
}

void MultiReplace::CloseDebugWindow()
{
    // Triggers the WM_CLOSE message for the debug window, handled in DebugWindowProc
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // Save the window position and size before closing
    RECT rc;
    if (GetWindowRect(hDebugWnd, &rc)) {
        debugWindowPosition = { rc.left, rc.top };
        debugWindowSize = { rc.right - rc.left, rc.bottom - rc.top };
        debugWindowPositionSet = debugWindowSizeSet = true;
    }

    SendMessage(hDebugWnd, WM_CLOSE, 0, 0);  // synchronous → window really gone
}

#pragma endregion


#pragma region Find All

std::wstring MultiReplace::sanitizeSearchPattern(const std::wstring& raw) {
    // Escape newlines directly in the wstring (no encoding conversion needed)
    std::wstring escaped = raw;
    // Replace CR/LF as before, but for wstring
    size_t pos = 0;
    while ((pos = escaped.find(L"\r", pos)) != std::wstring::npos)
        escaped.replace(pos, 1, L"\\r"), pos += 2;
    pos = 0;
    while ((pos = escaped.find(L"\n", pos)) != std::wstring::npos)
        escaped.replace(pos, 1, L"\\n"), pos += 2;
    return escaped;
}

void MultiReplace::trimHitToFirstLine(
    const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend,
    ResultDock::Hit& h)
{
    // Determine EOL sequence based on document setting
    int eolMode = (int)sciSend(SCI_GETEOLMODE, 0, 0);
    std::string eolSeq =
        eolMode == SC_EOL_CRLF ? "\r\n" :
        (eolMode == SC_EOL_CR ? "\r" : "\n");

    // Get line index and start position
    int lineZero = (int)sciSend(SCI_LINEFROMPOSITION, h.pos, 0);
    Sci_Position lineStart = sciSend(SCI_POSITIONFROMLINE, lineZero, 0);

    // Read raw line including EOL
    int rawLen = (int)sciSend(SCI_LINELENGTH, lineZero, 0);
    std::string raw;
    raw.resize(rawLen);
    sciSend(SCI_GETLINE, lineZero, (LPARAM)raw.data());
    raw.resize(strnlen(raw.c_str(), rawLen));

    // Find EOL position
    size_t eolPos = raw.find(eolSeq);
    if (eolPos == std::string::npos) {
        size_t posCR = raw.find('\r');
        size_t posLF = raw.find('\n');
        if (posCR != std::string::npos && posLF != std::string::npos)
            eolPos = std::min(posCR, posLF);
        else
            eolPos = (posCR != std::string::npos ? posCR : posLF);
    }

    // Trim match length to not exceed first line
    if (eolPos != std::string::npos) {
        Sci_Position matchOffset = h.pos - lineStart;
        Sci_Position maxLen = (Sci_Position)eolPos - matchOffset;
        if (matchOffset < (Sci_Position)eolPos) {
            h.length = std::max<Sci_Position>(0, std::min(h.length, maxLen));
        }
        else {
            h.length = 0;
        }
    }
}

void MultiReplace::handleFindAllButton()
{
    // 1) sanity 
    if (!validateDelimiterData())
        return;

    // 2) result dock 
    ResultDock& dock = ResultDock::instance();
    dock.ensureCreatedAndVisible(nppData);

    // 3) helper lambdas 
    auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT {
        return ::SendMessage(_hScintilla, m, w, l);
        };

    // 4) current file path 
    wchar_t buf[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, (LPARAM)buf);
    std::wstring wFilePath = *buf ? buf : L"<untitled>";
    std::string  utf8FilePath = Encoding::wstringToUtf8(wFilePath);

    // 5) search context 
    SearchContext context;
    context.docLength = sciSend(SCI_GETLENGTH);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.retrieveFoundText = false;
    context.highlightMatch = false;

    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
    Sci_Position scanStart = computeAllStartPos(context, wrapAroundEnabled, allFromCursorEnabled);

    // 6) containers 
    ResultDock::FileMap fileMap;
    int totalHits = 0;

    // ===== LIST MODE =====
    if (useListEnabled)
    {
        if (replaceListData.empty())
        {
            showStatusMessage(LM.get(L"status_add_values_or_uncheck"), MessageStatus::Error);
            return;
        }

        resetCountColumns();

        for (size_t idx = 0; idx < replaceListData.size(); ++idx)
        {
            auto& item = replaceListData[idx];
            if (!item.isEnabled || item.findText.empty())
                continue;

            // Sanitize pattern for this criterion's header
            std::wstring sanitizedPattern = this->sanitizeSearchPattern(item.findText);

            // (a) Set up search flags & pattern 
            context.findText = convertAndExtendW(item.findText, item.extended);
            context.searchFlags =
                (item.wholeWord ? SCFIND_WHOLEWORD : 0)
                | (item.matchCase ? SCFIND_MATCHCASE : 0)
                | (item.regex ? SCFIND_REGEXP : 0);
            sciSend(SCI_SETSEARCHFLAGS, context.searchFlags);

            // (b) Collect hits
            std::vector<ResultDock::Hit> rawHits;
            LRESULT pos = scanStart;
            while (true)
            {
                SearchResult r = performSearchForward(context, pos);
                if (r.pos < 0) break;
                pos = advanceAfterMatch(r);

                ResultDock::Hit h{};
                h.fullPathUtf8 = utf8FilePath;
                h.pos = (Sci_Position)r.pos;
                h.length = (Sci_Position)r.length;
                this->trimHitToFirstLine(sciSend, h);
                if (h.length > 0)
                    rawHits.push_back(std::move(h));
            }

            // (c) Write per-criterion counters (even if zero)
            const int hitCnt = static_cast<int>(rawHits.size());
            item.findCount = std::to_wstring(hitCnt);
            updateCountColumns(idx, hitCnt);

            // (d) Aggregate only when there are hits
            if (hitCnt > 0) {
                auto& agg = fileMap[utf8FilePath];
                agg.wPath = wFilePath;
                agg.hitCount += hitCnt;
                agg.crits.push_back({ sanitizedPattern, std::move(rawHits) });
                totalHits += hitCnt;
            }
        }

        refreshUIListView();
    }
    // ===== SINGLE MODE =====
    else
    {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findW);

        // Prepare header pattern & flags
        std::wstring headerPattern = this->sanitizeSearchPattern(findW);
        context.findText = convertAndExtendW(
            findW,
            IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED
        );
        context.searchFlags =
            (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0)
            | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0)
            | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);
        sciSend(SCI_SETSEARCHFLAGS, context.searchFlags);

        // Collect hits
        std::vector<ResultDock::Hit> rawHits;
        LRESULT pos = scanStart;
        while (true)
        {
            SearchResult r = performSearchForward(context, pos);
            if (r.pos < 0) break;
            pos = advanceAfterMatch(r);

            ResultDock::Hit h{};
            h.fullPathUtf8 = utf8FilePath;
            h.pos = r.pos;
            h.length = r.length;
            this->trimHitToFirstLine(sciSend, h);
            if (h.length > 0)
                rawHits.push_back(std::move(h));
        }

        // Aggregate into one file entry only when there are hits
        if (!rawHits.empty())
        {
            auto& agg = fileMap[utf8FilePath];
            agg.wPath = wFilePath;
            agg.hitCount = static_cast<int>(rawHits.size());
            agg.crits.push_back({ headerPattern, std::move(rawHits) });
            totalHits += agg.hitCount;
        }
    }

    // ===== unified Search Result API calls — ALWAYS show a header =====
    const size_t fileCount = fileMap.size(); // counts files *with* hits
    const std::wstring header = useListEnabled
        ? LM.get(L"dock_list_header", { std::to_wstring(totalHits), std::to_wstring(fileCount) })
        : LM.get(L"dock_single_header", { this->sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)),
                                           std::to_wstring(totalHits),
                                           std::to_wstring(fileCount) });

    dock.startSearchBlock(header,
        useListEnabled ? groupResultsEnabled : false,
        dock.purgeEnabled());

    // In the current-doc case we have at most one file block to append
    if (fileCount > 0)
        dock.appendFileBlock(fileMap, sciSend);

    dock.closeSearchBlock(totalHits, static_cast<int>(fileCount));

    // Status: 0 hits → Error, else Success
    showStatusMessage(
        (totalHits == 0)
        ? LM.get(L"status_no_matches_found")
        : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) }),
        (totalHits == 0) ? MessageStatus::Error : MessageStatus::Success);
}

void MultiReplace::handleFindAllInDocsButton()
{
    // 1) sanity + dock setup
    if (!validateDelimiterData())
        return;

    ResultDock& dock = ResultDock::instance();
    dock.ensureCreatedAndVisible(nppData);

    // 2) counters
    int totalHits = 0;
    std::unordered_set<std::string> uniqueFiles; // files *with* hits

    if (useListEnabled) resetCountColumns();
    std::vector<int> listHitTotals(useListEnabled ? replaceListData.size() : 0, 0);

    // 3) placeholder header (will be updated in closeSearchBlock)
    std::wstring placeholder = useListEnabled
        ? LM.get(L"dock_list_header", { L"0", L"0" })
        : LM.get(L"dock_single_header", {
              sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)),
              L"0", L"0" });

    // open block (pending text; commit happens in closeSearchBlock)
    dock.startSearchBlock(placeholder,
        groupResultsEnabled,
        dock.purgeEnabled());

    // 4) scan each tab
    auto processCurrentBuffer = [&]()
        {
            pointerToScintilla();
            auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0)->LRESULT {
                return ::SendMessage(_hScintilla, m, w, l);
                };

            wchar_t wBuf[MAX_PATH] = {};
            ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, (LPARAM)wBuf);
            std::wstring wPath = *wBuf ? wBuf : L"<untitled>";
            std::string  u8Path = Encoding::wstringToUtf8(wPath);

            const bool selMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
            SelectionInfo sel = getSelectionInfo(false);
            if (selMode && sel.length == 0) return;

            Sci_Position scanStart = selMode ? sel.startPos : 0;
            const bool columnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);

            ResultDock::FileMap fileMap;
            int hitsInFile = 0;

            auto collect = [&](size_t critIdx, const std::wstring& patt, SearchContext& ctx) {
                std::vector<ResultDock::Hit> raw;
                LRESULT pos = scanStart;
                while (true)
                {
                    SearchResult r = performSearchForward(ctx, pos);
                    if (r.pos < 0) break;
                    pos = advanceAfterMatch(r);

                    ResultDock::Hit h{};
                    h.fullPathUtf8 = u8Path;
                    h.pos = r.pos;
                    h.length = r.length;
                    this->trimHitToFirstLine(sciSend, h);
                    if (h.length > 0) raw.push_back(std::move(h));
                }
                const int hitCnt = static_cast<int>(raw.size());
                if (useListEnabled && critIdx < listHitTotals.size())
                    listHitTotals[critIdx] += hitCnt;
                if (hitCnt == 0) return;

                auto& agg = fileMap[u8Path];
                agg.wPath = wPath;
                agg.hitCount += hitCnt;
                agg.crits.push_back({ patt, std::move(raw) });
                hitsInFile += hitCnt;
                };

            if (useListEnabled)
            {
                for (size_t idx = 0; idx < replaceListData.size(); ++idx)
                {
                    const auto& it = replaceListData[idx];
                    if (!it.isEnabled || it.findText.empty()) continue;

                    SearchContext ctx;
                    ctx.docLength = sciSend(SCI_GETLENGTH);
                    ctx.isColumnMode = columnMode;
                    ctx.isSelectionMode = selMode;
                    ctx.findText = convertAndExtendW(it.findText, it.extended);
                    ctx.searchFlags = (it.wholeWord ? SCFIND_WHOLEWORD : 0)
                        | (it.matchCase ? SCFIND_MATCHCASE : 0)
                        | (it.regex ? SCFIND_REGEXP : 0);
                    sciSend(SCI_SETSEARCHFLAGS, ctx.searchFlags);
                    collect(idx, sanitizeSearchPattern(it.findText), ctx);
                }
            }
            else
            {
                std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
                if (!findW.empty()) {
                    SearchContext ctx;
                    ctx.docLength = sciSend(SCI_GETLENGTH);
                    ctx.isColumnMode = columnMode;
                    ctx.isSelectionMode = selMode;
                    ctx.findText = convertAndExtendW(findW,
                        IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
                    ctx.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0)
                        | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0)
                        | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);
                    sciSend(SCI_SETSEARCHFLAGS, ctx.searchFlags);
                    collect(0, sanitizeSearchPattern(findW), ctx);
                }
            }

            if (hitsInFile > 0)
            {
                // Commit this file and count it as "with hits"
                dock.appendFileBlock(fileMap, sciSend);
                totalHits += hitsInFile;
                uniqueFiles.insert(u8Path);
            }
        };

    // iterate tabs (unchanged)
    LRESULT savedIdx = ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTDOCINDEX, 0, MAIN_VIEW);
    const bool mainVis = !!::IsWindowVisible(nppData._scintillaMainHandle);
    const bool subVis = !!::IsWindowVisible(nppData._scintillaSecondHandle);
    LRESULT nbMain = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, PRIMARY_VIEW);
    LRESULT nbSub = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, SECOND_VIEW);

    if (mainVis)
        for (LRESULT i = 0; i < nbMain; ++i)
        {
            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, MAIN_VIEW, i);
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            processCurrentBuffer();
        }
    if (subVis)
        for (LRESULT i = 0; i < nbSub; ++i)
        {
            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, SUB_VIEW, i);
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            processCurrentBuffer();
        }
    ::SendMessage(nppData._nppHandle,
        NPPM_ACTIVATEDOC,
        mainVis ? MAIN_VIEW : SUB_VIEW,
        savedIdx);

    // write per-criterion totals back to list and repaint
    if (useListEnabled)
    {
        for (size_t i = 0; i < listHitTotals.size(); ++i)
        {
            if (!replaceListData[i].isEnabled || replaceListData[i].findText.empty())
                continue;
            replaceListData[i].findCount = std::to_wstring(listHitTotals[i]);
            updateCountColumns(i, listHitTotals[i]);
        }
        refreshUIListView();
    }

    // 6) finalise — ALWAYS close block (header even with zero hits)
    dock.closeSearchBlock(totalHits, static_cast<int>(uniqueFiles.size()));

    // Status: 0 hits → Error, else Success
    showStatusMessage(
        (totalHits == 0)
        ? LM.get(L"status_no_matches_found")
        : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) }),
        (totalHits == 0) ? MessageStatus::Error : MessageStatus::Success);
}

void MultiReplace::handleFindInFiles() {
    // 1) sanity + hidden buffer
    if (!validateDelimiterData())
        return;

    HiddenSciGuard guard;
    if (!guard.create()) {
        showStatusMessage(LM.get(L"status_error_hidden_buffer"), MessageStatus::Error);
        return;
    }

    // Read all inputs once at the beginning.
    auto wDir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);
    auto wFilter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
    const bool recurse = (IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED);
    const bool hide = (IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED);

    // If the filter is empty, default to "*.*" and update the UI.
    if (wFilter.empty()) {
        wFilter = L"*.*";
        SetDlgItemTextW(_hSelf, IDC_FILTER_EDIT, wFilter.c_str());
    }
    guard.parseFilter(wFilter);

    if (wDir.empty() || !std::filesystem::exists(wDir)) {
        showStatusMessage(LM.get(L"status_error_invalid_directory"), MessageStatus::Error);
        return;
    }

    // Build the file list using the same traversal and matching logic as ReplaceInFiles.
    std::vector<std::filesystem::path> files;
    try {
        namespace fs = std::filesystem;
        if (recurse) {
            for (auto& e : fs::recursive_directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
        else {
            for (auto& e : fs::directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
    }
    catch (const std::exception& ex) {
        std::wstring wideReason = Encoding::utf8ToWString(ex.what());
        showStatusMessage(LM.get(L"status_error_scanning_directory", { wideReason }), MessageStatus::Error);
        return;
    }

    if (files.empty()) {
        MessageBox(_hSelf, LM.getLPW(L"msgbox_no_files"), LM.getLPW(L"msgbox_title_confirm"), MB_OK);
        return;
    }

    // 2) result dock
    ResultDock& dock = ResultDock::instance();
    dock.ensureCreatedAndVisible(nppData);  // ok if already visible  :contentReference[oaicite:0]{index=0}

    // counters
    int totalHits = 0;
    std::unordered_set<std::string> uniqueFiles; // files *with* hits only

    if (useListEnabled) resetCountColumns();
    std::vector<int> listHitTotals(useListEnabled ? replaceListData.size() : 0, 0);

    // placeholder header (updated in closeSearchBlock)
    std::wstring placeholder = useListEnabled
        ? LM.get(L"dock_list_header", { L"0", L"0" })
        : LM.get(L"dock_single_header", {
            sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)),
            L"0", L"0" });

    // Open the search block (pending buffer; committed in closeSearchBlock)
    dock.startSearchBlock(placeholder,
        useListEnabled ? groupResultsEnabled : false,
        dock.purgeEnabled());  // may clear view first  :contentReference[oaicite:1]{index=1}

    // 3) RAII-based UI lock identical to ReplaceInFiles
    BatchUIGuard uiGuard(this, _hSelf);
    _isCancelRequested = false;

    // Progress
    int idx = 0;
    const int total = static_cast<int>(files.size());
    showStatusMessage(L"Progress: [  0%]", MessageStatus::Info);

    // Small RAII guard to bind/unbind the hidden Scintilla safely per file
    struct SciBindingGuard {
        MultiReplace* self;
        HWND oldSci; SciFnDirect oldFn; sptr_t oldData;
        HiddenSciGuard& g;
        SciBindingGuard(MultiReplace* s, HiddenSciGuard& guard) : self(s), g(guard) {
            oldSci = s->_hScintilla; oldFn = s->pSciMsg; oldData = s->pSciWndData;
            s->_hScintilla = g.hSci; s->pSciMsg = g.fn; s->pSciWndData = g.pData;
        }
        ~SciBindingGuard() {
            self->_hScintilla = oldSci; self->pSciMsg = oldFn; self->pSciWndData = oldData;
        }
    };

    // 4) per-file processing
    bool aborted = false;

    for (const auto& fp : files)
    {
        // Pump messages to keep UI responsive and allow cancel clicks
        MSG m;
        while (::PeekMessage(&m, nullptr, 0, 0, PM_REMOVE)) {
            ::TranslateMessage(&m);
            ::DispatchMessage(&m);
        }

        if (_isShuttingDown) { aborted = true; break; }
        if (_isCancelRequested) { aborted = true; break; }

        ++idx;

        // progress line with shortened path
        const int percent = static_cast<int>((static_cast<double>(idx) / (std::max)(1, total)) * 100.0);
        const std::wstring prefix = L"Progress: [" + std::to_wstring(percent) + L"%] ";

        HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
        HDC hdc = GetDC(hStatus);
        HFONT hFont = (HFONT)SendMessage(hStatus, WM_GETFONT, 0, 0);
        SelectObject(hdc, hFont);
        SIZE sz{}; GetTextExtentPoint32W(hdc, prefix.c_str(), (int)prefix.length(), &sz);
        RECT rc{}; GetClientRect(hStatus, &rc);
        const int avail = (rc.right - rc.left) - sz.cx;
        std::wstring shortPath = getShortenedFilePath(fp.wstring(), avail, hdc);
        ReleaseDC(hStatus, hdc);
        showStatusMessage(prefix + shortPath, MessageStatus::Info);

        // read file; read-only/not-readable files are simply skipped
        std::string original;
        if (!guard.loadFile(fp, original)) { continue; }

        // detect and convert to UTF-8 for the hidden buffer
        EncodingInfo enc = detectEncoding(original);
        std::string u8;
        if (!convertBufferToUtf8(original, enc, u8)) { continue; }

        // Bind hidden buffer for the whole per-file scope
        SciBindingGuard bind(this, guard);

        // set text into hidden buffer
        send(SCI_CLEARALL, 0, 0);
        send(SCI_SETCODEPAGE, SC_CP_UTF8, 0);
        send(SCI_ADDTEXT, (WPARAM)u8.length(), reinterpret_cast<sptr_t>(u8.data()));

        // delimiters snapshot
        handleDelimiterPositions(DelimiterOperation::LoadAll);

        // Aggregation for *this* file only
        const std::wstring wPath = fp.wstring();
        const std::string  u8Path = Encoding::wstringToUtf8(wPath);

        bool columnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);

        ResultDock::FileMap fileMap;   // one file entry
        int hitsInFile = 0;

        auto collect = [&](size_t critIdx, const std::wstring& pattW, SearchContext& ctx) {
            std::vector<ResultDock::Hit> raw;
            LRESULT pos = 0;

            while (true) {
                SearchResult r = performSearchForward(ctx, pos);
                if (r.pos < 0) break;
                pos = advanceAfterMatch(r);

                ResultDock::Hit h{};
                h.fullPathUtf8 = u8Path;
                h.pos = r.pos;
                h.length = r.length;

                // Trim to first line for the dock (operates on currently bound hidden buffer)
                this->trimHitToFirstLine([this](UINT m, WPARAM w, LPARAM l)->LRESULT { return send(m, w, l); }, h);

                if (h.length > 0) raw.push_back(std::move(h));
            }

            const int n = (int)raw.size();
            if (n == 0) return;

            auto& agg = fileMap[u8Path];
            agg.wPath = wPath;
            agg.hitCount += n;
            agg.crits.push_back({ sanitizeSearchPattern(pattW), std::move(raw) });

            hitsInFile += n;
            totalHits += n;
            if (useListEnabled && critIdx < listHitTotals.size())
                listHitTotals[critIdx] += n;
            };

        // search setup
        if (useListEnabled)
        {
            for (size_t i = 0; i < replaceListData.size(); ++i)
            {
                const auto& it = replaceListData[i];
                if (!it.isEnabled || it.findText.empty()) continue;

                SearchContext ctx{};
                ctx.docLength = send(SCI_GETLENGTH);
                ctx.isColumnMode = columnMode;
                ctx.isSelectionMode = false;
                ctx.findText = convertAndExtendW(it.findText, it.extended);
                ctx.searchFlags = (it.wholeWord ? SCFIND_WHOLEWORD : 0)
                    | (it.matchCase ? SCFIND_MATCHCASE : 0)
                    | (it.regex ? SCFIND_REGEXP : 0);
                send(SCI_SETSEARCHFLAGS, ctx.searchFlags, 0);
                collect(i, it.findText, ctx);
            }
        }
        else
        {
            std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            if (!findW.empty())
            {
                SearchContext ctx{};
                ctx.docLength = send(SCI_GETLENGTH);
                ctx.isColumnMode = columnMode;
                ctx.isSelectionMode = false;
                ctx.findText = convertAndExtendW(findW, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
                ctx.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0)
                    | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0)
                    | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);
                send(SCI_SETSEARCHFLAGS, ctx.searchFlags, 0);
                collect(0, findW, ctx);
            }
        }

        // Commit per-file block
        if (hitsInFile > 0) {
            auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0)->LRESULT { return send(m, w, l); };
            dock.appendFileBlock(fileMap, sciSend);  // adds to pending text  :contentReference[oaicite:2]{index=2}
            uniqueFiles.insert(u8Path);              // count only files with hits
        }
    }

    // 5) finalize: ALWAYS close the search block, even with zero hits or cancel
    dock.closeSearchBlock(totalHits, static_cast<int>(uniqueFiles.size()));  // commits header + pending text  :contentReference[oaicite:3]{index=3}

    // update list counters + UI
    if (useListEnabled)
    {
        for (size_t i = 0; i < listHitTotals.size(); ++i)
        {
            if (!replaceListData[i].isEnabled || replaceListData[i].findText.empty())
                continue;
            replaceListData[i].findCount = std::to_wstring(listHitTotals[i]);
            updateCountColumns(i, listHitTotals[i]);
        }
        refreshUIListView();
    }

    // status line
// --- status line (+ optional " - Canceled" suffix from LM) ---
    const bool wasCanceled = (_isCancelRequested || aborted);
    const std::wstring canceledSuffix = wasCanceled ? (L" - " + LM.get(L"status_canceled")) : L"";

    std::wstring msg = (totalHits == 0)
        ? LM.get(L"status_no_matches_found")
        : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) });

    MessageStatus ms = MessageStatus::Success;
    if (wasCanceled) {
        ms = MessageStatus::Info;
    }
    else if (totalHits == 0) {
        ms = MessageStatus::Error;
    }

    showStatusMessage(msg + canceledSuffix, ms, true);
    _isCancelRequested = false;
}

#pragma endregion


#pragma region Find

void MultiReplace::handleFindNextButton() {

    if (!validateDelimiterData()) {
        return;
    }

    size_t matchIndex = std::numeric_limits<size_t>::max();
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
    SelectionInfo selection = getSelectionInfo(false);

    // Determine the starting search position:
    // If there is a selection and the selection radio is checked, start at the beginning of the selection.
    // Otherwise, use the current cursor position.
    Sci_Position searchPos = (selection.length > 0 && IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)
        ? selection.startPos
        : send(SCI_GETCURRENTPOS, 0, 0);

    // Initialize SearchContext
    SearchContext context;
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_find_directly"), MessageStatus::Error);
            return;
        }

        // Call performListSearchForward with the complete SearchContext
        SearchResult result = performListSearchForward(replaceListData, searchPos, matchIndex, context);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performListSearchForward(replaceListData, 0, matchIndex, context);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_wrapped"), MessageStatus::Info);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            updateCountColumns(matchIndex, 1);
            refreshUIListView();
            selectListItem(matchIndex);
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found"), MessageStatus::Error, true);
        }
    }
    else {
        // Read search text and check search options
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);

        // Determine if extended mode is active
        bool isExtended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        context.findText = convertAndExtendW(findText, isExtended);

        // Set search flags
        context.searchFlags =
            (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);

        // Set search flags before calling `performSearchForward`
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        // Perform forward search
        SearchResult result = performSearchForward(context, searchPos);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performSearchForward(context, 0);
            if (result.pos >= 0) {
                showStatusMessage(LM.get(L"status_wrapped"), MessageStatus::Info);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found_for", { findText }), MessageStatus::Error, true);
        }
    }
}

void MultiReplace::handleFindPrevButton() {

    if (!validateDelimiterData()) {
        return;
    }

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SelectionInfo selection = getSelectionInfo(true);
    Sci_Position searchPos;

    // Determine starting position for backward search:
    // If there is a selection and the selection radio is checked, start from the end of the selection.
    // Otherwise, use the current cursor position.
    if (selection.length > 0 && (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)) {
        searchPos = selection.endPos;
    }
    else {
        searchPos = send(SCI_GETCURRENTPOS, 0, 0);
    }
    // Move one position backward if possible
    searchPos = (searchPos > 0) ? send(SCI_POSITIONBEFORE, searchPos, 0) : searchPos;

    // Create a fully initialized `SearchContext`
    SearchContext context;
    context.findText = convertAndExtendW(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT),
        IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
    context.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0) |
        (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0) |
        (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        size_t matchIndex = std::numeric_limits<size_t>::max();
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_find_directly"), MessageStatus::Error);
            return;
        }
        SearchResult result = performListSearchBackward(replaceListData, searchPos, matchIndex, context);
        if (result.pos < 0 && wrapAroundEnabled) {
            searchPos = send(SCI_GETLENGTH, 0, 0);
            result = performListSearchBackward(replaceListData, searchPos, matchIndex, context);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_wrapped"), MessageStatus::Info);
                if (context.isSelectionMode) {
                    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
                }
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            updateCountColumns(matchIndex, 1);
            refreshUIListView();
            selectListItem(matchIndex);
            if (context.isSelectionMode) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found"), MessageStatus::Error, true);
        }
    }
    else {

        // Set search flags before calling 'performSearchBackward'
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        SearchResult result = performSearchBackward(context, searchPos);
        if (result.pos < 0 && wrapAroundEnabled) {
            searchPos = context.isSelectionMode ? selection.endPos : send(SCI_GETLENGTH, 0, 0);
            result = performSearchBackward(context, searchPos);
            if (result.pos >= 0) {
                showStatusMessage(LM.get(L"status_wrapped"), MessageStatus::Info);
                if (context.isSelectionMode) {
                    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
                }
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            if (context.isSelectionMode) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found_for", { getTextFromDialogItem(_hSelf, IDC_FIND_EDIT) }),
                MessageStatus::Error, true);
        }
    }
}

SearchResult MultiReplace::performSingleSearch(const SearchContext& context, SelectionRange range)
{
    // Early exit if the search string is empty
    if (context.findText.empty()) {
        return {};  // Return default-initialized SearchResult
    }

    // Set the target range and search flags via Scintilla
    send(SCI_SETTARGETRANGE, range.start, range.end);

    // Perform the search using SCI_SEARCHINTARGET
    Sci_Position pos = send(SCI_SEARCHINTARGET, context.findText.size(), reinterpret_cast<sptr_t>(context.findText.c_str()));
    Sci_Position matchEnd = send(SCI_GETTARGETEND);

    // Validate the search result using cached document length from context
    if (pos < 0 || matchEnd <= pos || matchEnd > context.docLength) {
        return {};  // Return empty result if no match is found
    }

    // Construct the search result
    SearchResult result;
    result.pos = pos;
    result.length = matchEnd - pos;

    // Retrieve found text only when requested
    if (context.retrieveFoundText) {
        const int codepage = static_cast<int>(send(SCI_GETCODEPAGE));
        const size_t bytesPerChar = (codepage == SC_CP_UTF8) ? 4u : 1u;

        const size_t charsInTarget = static_cast<size_t>(result.length);
        const size_t cap = charsInTarget * bytesPerChar + 1;

        std::string buf(cap, '\0');
        LRESULT textLength = send(SCI_GETTARGETTEXT, 0, reinterpret_cast<LPARAM>(buf.data()));

        if (textLength < 0) textLength = 0;
        if (static_cast<size_t>(textLength) >= cap)
            textLength = static_cast<LRESULT>(cap - 1);

        buf.resize(static_cast<size_t>(textLength));
        result.foundText = std::move(buf);
    }

    // Highlight the match if required by context
    if (context.highlightMatch) {
        displayResultCentered(pos, matchEnd, true);
    }

    return result;
}

SearchResult MultiReplace::performSearchForward(const SearchContext& context, LRESULT start)
{
    bool isBackward = false;
    SelectionRange targetRange;
    SearchResult result;

    // Use cached mode states from the context to decide the search method
    if (context.isColumnMode && columnDelimiterData.isValid()) {
        result = performSearchColumn(context, start, isBackward);
    }
    else if (context.isSelectionMode) {
        result = performSearchSelection(context, start, isBackward);
    }
    else {
        targetRange.start = start;
        targetRange.end = context.docLength;
        result = performSingleSearch(context, targetRange);
    }
    return result;
}

SearchResult MultiReplace::performSearchBackward(const SearchContext& context, LRESULT start)
{
    bool isBackward = true;
    SelectionRange targetRange;
    SearchResult result;

    if (context.isSelectionMode) {
        result = performSearchSelection(context, start, isBackward);
    }
    else if (context.isColumnMode && columnDelimiterData.isValid()) {
        result = performSearchColumn(context, start, isBackward);
    }
    else {
        targetRange.start = start;
        targetRange.end = 0; // Beginning of the document for backward search
        result = performSingleSearch(context, targetRange);
    }
    return result;
}

SearchResult MultiReplace::performSearchSelection(const SearchContext& context, LRESULT start, bool isBackward) {
    SearchResult result;
    SelectionRange targetRange;
    std::vector<SelectionRange> selections;

    LRESULT selectionCount = send(SCI_GETSELECTIONS, 0, 0);
    if (selectionCount == 0) {
        return SearchResult(); // No selections to search
    }

    // Gather all selection positions
    selections.resize(selectionCount);
    for (int i = 0; i < selectionCount; ++i) {
        selections[i].start = send(SCI_GETSELECTIONNSTART, i, 0);
        selections[i].end = send(SCI_GETSELECTIONNEND, i, 0);
    }

    // Sort selections based on the search direction
    if (isBackward) {
        std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) { return a.start > b.start; });
    }
    else {
        std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) { return a.start < b.start; });
    }

    // Iterate through each selection range based on the search direction
    for (const auto& selection : selections) {
        if ((isBackward && start < selection.start) || (!isBackward && start > selection.end)) {
            continue; // Skip selections outside the current search range
        }

        // Set the target range within the selection
        targetRange.start = isBackward ? std::min(start, selection.end) : std::max(start, selection.start);
        targetRange.end = isBackward ? selection.start : selection.end;

        // Skip if the range is invalid
        if (targetRange.start == targetRange.end) continue;

        // Perform the search within the target range
        result = performSingleSearch(context, targetRange);
        if (result.pos >= 0) return result;

        // Update the starting position for the next iteration
        start = isBackward ? selection.start - 1 : selection.end + 1;
    }

    return result; // No match found
}

SearchResult MultiReplace::performSearchColumn(const SearchContext& context, LRESULT start, bool isBackward)
{
    SearchResult result;
    SelectionRange targetRange;

    // Identify column start information based on the starting position
    ColumnInfo columnInfo = getColumnInfo(start);
    LRESULT startLine = columnInfo.startLine;
    SIZE_T startColumnIndex = columnInfo.startColumnIndex;
    LRESULT totalLines = columnInfo.totalLines;

    // Set line iteration based on search direction
    LRESULT line = startLine;
    while (isBackward ? (line >= 0) : (line < totalLines)) {
        // Avoid out-of-bounds access in lineDelimiterPositions
        if (line >= static_cast<LRESULT>(lineDelimiterPositions.size())) {
            break;
        }

        // Retrieve local info for the line
        const auto& lineInfo = lineDelimiterPositions[line];
        const auto& linePositions = lineInfo.positions;
        SIZE_T totalColumns = linePositions.size() + 1;

        // Calculate absolute line start and end
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, line, 0);
        LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        // Set column iteration range and step based on direction
        SIZE_T column = isBackward ? (line == startLine ? startColumnIndex : totalColumns) : startColumnIndex;
        SIZE_T endColumnIdx = isBackward ? 1 : totalColumns;
        int columnStep = isBackward ? -1 : 1;

        // Iterate over columns in the specified direction
        for (; (isBackward ? (column >= endColumnIdx) : (column <= endColumnIdx)); column += columnStep) {
            LRESULT startColumn = 0;
            LRESULT endColumn = 0;

            // Define absolute start of this column
            if (column == 1) {
                startColumn = lineStartPos;
            }
            else {
                startColumn = lineStartPos
                    + lineInfo.positions[column - 2].offsetInLine
                    + columnDelimiterData.delimiterLength;
            }

            // Define absolute end of this column
            if (column == totalColumns) {
                endColumn = lineEndPos;
            }
            else {
                endColumn = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
            }

            // Skip columns not specified in columnDelimiterData
            if (columnDelimiterData.columns.find(static_cast<int>(column)) == columnDelimiterData.columns.end()) {
                continue;
            }

            // Adjust the target range based on start position and search direction
            if (isBackward && start >= startColumn && start <= endColumn) {
                endColumn = start;
            }
            else if (!isBackward && start >= startColumn && start <= endColumn) {
                startColumn = start;
            }

            // Define target range for the search
            targetRange.start = isBackward ? endColumn : startColumn;
            targetRange.end = isBackward ? startColumn : endColumn;

            // Perform search within the target range
            result = std::move(performSingleSearch(context, targetRange));

            // Return immediately if a match is found
            if (result.pos >= 0) {
                return result;
            }
        }

        // Move to the next line based on search direction
        line += (isBackward ? -1 : 1);

        // Reset column index for subsequent lines
        startColumnIndex = isBackward ? totalColumns : 1;
    }

    return result; // No match found in column mode
}

SearchResult MultiReplace::performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max(); // Initialize with a value that represents "no index".

    for (size_t i = 0; i < list.size(); ++i) {
        if (!list[i].isEnabled) {
            continue; // Skip disabled entries
        }

        // Use the existing search context and only override `findText` and `searchFlags`
        SearchContext localContext = context;
        localContext.findText = convertAndExtendW(list[i].findText, list[i].extended);
        localContext.searchFlags = (list[i].wholeWord ? SCFIND_WHOLEWORD : 0) |
            (list[i].matchCase ? SCFIND_MATCHCASE : 0) |
            (list[i].regex ? SCFIND_REGEXP : 0);

        // Set search flags before calling 'performSearchBackward'
        send(SCI_SETSEARCHFLAGS, localContext.searchFlags);

        // Perform backward search using the updated search context
        SearchResult result = performSearchBackward(localContext, cursorPos);

        // If a match was found and it's closer to the cursor than the current closest match, update the closest match
        if (result.pos >= 0 && (closestMatch.pos < 0 || (result.pos + result.length) >(closestMatch.pos + closestMatch.length))) {
            closestMatch = result;
            closestMatchIndex = i; // Update the index of the closest match
        }
    }

    if (closestMatch.pos >= 0) { // Check if a match was found
        displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, false);
    }

    return closestMatch;
}

SearchResult MultiReplace::performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max(); // Initialize with a value representing "no index".

    for (size_t i = 0; i < list.size(); ++i) {
        if (!list[i].isEnabled) {
            continue; // Skip disabled entries
        }

        // Use the existing search context and only override `findText` and `searchFlags`
        SearchContext localContext = context;
        localContext.findText = convertAndExtendW(list[i].findText, list[i].extended);
        localContext.searchFlags = (list[i].wholeWord ? SCFIND_WHOLEWORD : 0) |
            (list[i].matchCase ? SCFIND_MATCHCASE : 0) |
            (list[i].regex ? SCFIND_REGEXP : 0);

        // Set search flags before calling `performSearchForward`
        send(SCI_SETSEARCHFLAGS, localContext.searchFlags);

        // Perform forward search using the updated search context
        SearchResult result = performSearchForward(localContext, cursorPos);

        // If a match is found that is closer to the cursor than the current closest match, update the closest match
        if (result.pos >= 0 && (closestMatch.pos < 0 || result.pos < closestMatch.pos)) {
            closestMatch = result;
            closestMatchIndex = i; // Update the index of the closest match
        }
    }

    if (closestMatch.pos >= 0) { // Check if a match was found
        displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, true);
    }

    return closestMatch;
}

void MultiReplace::displayResultCentered(size_t posStart, size_t posEnd, bool isDownwards)
{
    // Make sure target lines are unfolded
    send(SCI_ENSUREVISIBLE, send(SCI_LINEFROMPOSITION, posStart, 0), 0);
    send(SCI_ENSUREVISIBLE, send(SCI_LINEFROMPOSITION, posEnd, 0), 0);

    // Jump-scroll to center, if current position is out of view
    send(SCI_SETVISIBLEPOLICY, CARET_JUMPS | CARET_EVEN, 0);
    send(SCI_ENSUREVISIBLEENFORCEPOLICY, send(SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // When searching up, the beginning of the (possible multiline) result is important, when scrolling down the end
    send(SCI_GOTOPOS, isDownwards ? posEnd : posStart, 0);
    send(SCI_SETVISIBLEPOLICY, CARET_EVEN, 0);
    send(SCI_ENSUREVISIBLEENFORCEPOLICY, send(SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // Adjust so that we see the entire match; primarily horizontally
    send(SCI_SCROLLRANGE, posStart, posEnd);

    // Move cursor to end of result and select result
    send(SCI_GOTOPOS, posEnd, 0);
    send(SCI_SETANCHOR, posStart, 0);

    // Update Scintilla's knowledge about what column the caret is in, so that if user
    // does up/down arrow as first navigation after the search result is selected,
    // the caret doesn't jump to an unexpected column
    send(SCI_CHOOSECARETX, 0, 0);

}

void MultiReplace::selectListItem(size_t matchIndex) {
    if (!highlightMatchEnabled) {
        return;
    }

    HWND hListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    if (hListView && matchIndex != std::numeric_limits<size_t>::max()) {
        // Deselect all items
        ListView_SetItemState(hListView, -1, 0, LVIS_SELECTED);

        // Select the item at matchIndex
        ListView_SetItemState(hListView, static_cast<int>(matchIndex), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

        // Ensure the item is visible, but only scroll if absolutely necessary
        ListView_EnsureVisible(hListView, static_cast<int>(matchIndex), TRUE);
    }
}

#pragma endregion


#pragma region Mark

void MultiReplace::handleMarkMatchesButton() {

    if (!validateDelimiterData()) {
        return;
    }

    int totalMatchCount = 0;
    markedStringsCount = 0;

    // Read wrap state once
    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_mark_directly"), MessageStatus::Error);
            return;
        }

        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;

            const ReplaceItemData& item = replaceListData[i];

            // Build SearchContext for list-based marking
            SearchContext context;
            context.findText = convertAndExtendW(item.findText, item.extended);
            context.searchFlags = (item.wholeWord * SCFIND_WHOLEWORD)
                | (item.matchCase * SCFIND_MATCHCASE)
                | (item.regex * SCFIND_REGEXP);
            context.docLength = send(SCI_GETLENGTH, 0, 0);
            context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
            context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
            context.retrieveFoundText = false;
            context.highlightMatch = false;

            // --- Start position logic (mirrors Find All / Replace All)
            Sci_Position startPos = 0;
            if (context.isSelectionMode) {
                const SelectionInfo selInfo = getSelectionInfo(false);
                startPos = selInfo.startPos;
            }
            else if (wrapAroundEnabled) {
                startPos = 0;
            }
            else {
                const Sci_Position caretPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
                startPos = allFromCursorEnabled ? caretPos : 0;
            }

            const int matchCount = markString(context, startPos);
            if (matchCount > 0) {
                totalMatchCount += matchCount;
                updateCountColumns(i, matchCount);
                refreshUIListView();  // Refresh UI only when necessary
            }
        }
    }
    else {
        // Retrieve search parameters from UI
        const std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        const bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        const bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        const bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        const bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        // Build SearchContext for direct marking
        SearchContext context;
        context.findText = convertAndExtendW(findText, extended);
        context.searchFlags = (wholeWord * SCFIND_WHOLEWORD)
            | (matchCase * SCFIND_MATCHCASE)
            | (regex * SCFIND_REGEXP);
        context.docLength = send(SCI_GETLENGTH, 0, 0);
        context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
        context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
        context.retrieveFoundText = false;
        context.highlightMatch = false;

        // --- Start position logic (mirrors Find All / Replace All)
        Sci_Position startPos = 0;
        if (context.isSelectionMode) {
            const SelectionInfo selInfo = getSelectionInfo(false);
            startPos = selInfo.startPos;
        }
        else if (wrapAroundEnabled) {
            startPos = 0;
        }
        else {
            const Sci_Position caretPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
            startPos = allFromCursorEnabled ? caretPos : 0;
        }

        totalMatchCount = markString(context, startPos);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }

    // Display total number of marked occurrences
    showStatusMessage(LM.get(L"status_occurrences_marked", { std::to_wstring(totalMatchCount) }), MessageStatus::Info);
}

int MultiReplace::markString(const SearchContext& context, Sci_Position initialStart)
{
    if (context.findText.empty()) return 0;

    int markCount = 0;
    LRESULT pos = initialStart;
    send(SCI_SETSEARCHFLAGS, context.searchFlags);

    for (SearchResult r = performSearchForward(context, pos);
        r.pos >= 0;
        r = performSearchForward(context, pos))
    {
        if (r.length > 0) {
            highlightTextRange(r.pos, r.length, context.findText);
            ++markCount;
        }
        pos = advanceAfterMatch(r);

        if (pos >= context.docLength) break;
    }

    if (useListEnabled && markCount > 0) ++markedStringsCount;
    return markCount;
}

void MultiReplace::highlightTextRange(LRESULT pos, LRESULT len, const std::string& findText)
{
    int color = useListEnabled ? generateColorValue(findText) : MARKER_COLOR;
    auto it = colorToStyleMap.find(color);
    int indicatorStyle;

    if (it == colorToStyleMap.end()) {
        indicatorStyle = useListEnabled ? textStyles[(colorToStyleMap.size() % (textStyles.size() - 1)) + 1] : textStyles[0];
        colorToStyleMap[color] = indicatorStyle;
        send(SCI_INDICSETFORE, indicatorStyle, color);
    }
    else {
        indicatorStyle = it->second;
    }

    send(SCI_SETINDICATORCURRENT, indicatorStyle, 0);
    send(SCI_INDICSETSTYLE, indicatorStyle, INDIC_STRAIGHTBOX);
    send(SCI_INDICSETALPHA, indicatorStyle, 100);
    send(SCI_INDICATORFILLRANGE, pos, len);
}

int MultiReplace::generateColorValue(const std::string& str) {
    // DJB2 hash
    unsigned long hash = 5381;
    for (char c : str) {
        hash = ((hash << 5) + hash) + c;  // hash * 33 + c
    }

    // Create an RGB color using the hash
    int r = (hash >> 16) & 0xFF;
    int g = (hash >> 8) & 0xFF;
    int b = hash & 0xFF;

    // Convert RGB to long
    int color = (r << 16) | (g << 8) | b;

    return color;
}

void MultiReplace::handleClearTextMarksButton()
{
    for (int style : textStyles)
    {
        send(SCI_SETINDICATORCURRENT, style, 0);
        send(SCI_INDICATORCLEARRANGE, 0, send(SCI_GETLENGTH, 0, 0));
    }

    markedStringsCount = 0;
    colorToStyleMap.clear();
}

void MultiReplace::handleCopyMarkedTextToClipboardButton()
{

    if (!validateDelimiterData()) {
        return;
    }

    bool wasLastCharMarked = false;
    size_t markedTextCount = 0;

    std::string markedText;
    std::string styleText;
    std::string eol = getEOLStyle();

    for (int style : textStyles)
    {
        send(SCI_SETINDICATORCURRENT, style, 0);
        LRESULT pos = 0;
        LRESULT nextPos = send(SCI_INDICATOREND, style, pos);

        while (nextPos > pos) // check if nextPos has advanced
        {
            bool atEndOfIndic = send(SCI_INDICATORVALUEAT, style, pos) != 0;

            if (atEndOfIndic)
            {
                if (!wasLastCharMarked)
                {
                    ++markedTextCount;
                }

                wasLastCharMarked = true;

                for (LRESULT i = pos; i < nextPos; ++i)
                {
                    char ch = static_cast<char>(send(SCI_GETCHARAT, i, 0));
                    styleText += ch;
                }

                markedText += styleText + eol; // Append marked text and EOL;
                styleText.clear();
            }
            else
            {
                wasLastCharMarked = false;
            }

            pos = nextPos;
            nextPos = send(SCI_INDICATOREND, style, pos);
        }
    }

    // Remove the last EOL if necessary
    if (!markedText.empty() && markedText.length() >= eol.length()) {
        markedText.erase(markedText.length() - eol.length());
    }

    // Convert encoding to wide string
    std::wstring wstr = Encoding::bytesToWString(markedText, getCurrentDocCodePage());;

    copyTextToClipboard(wstr, static_cast<int>(markedTextCount));
}

void MultiReplace::copyTextToClipboard(const std::wstring& text, int textCount)
{
    if (text.empty()) {
        showStatusMessage(LM.get(L"status_no_text_to_copy"), MessageStatus::Error);
        return;
    }

    if (!OpenClipboard(nullptr))
        return;
    EmptyClipboard();

    const size_t byteSize = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, byteSize);
    if (!hMem) {
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_allocate_memory"), MessageStatus::Error);
        return;
    }

    void* p = GlobalLock(hMem);
    if (!p) {
        GlobalFree(hMem);
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_allocate_memory"), MessageStatus::Error);
        return;
    }

    memcpy(p, text.c_str(), byteSize);
    GlobalUnlock(hMem);

    if (SetClipboardData(CF_UNICODETEXT, hMem) == nullptr) {
        GlobalFree(hMem);
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_to_copy"), MessageStatus::Error);
        return;
    }

    CloseClipboard();
    showStatusMessage(LM.get(L"status_items_copied_to_clipboard", { std::to_wstring(textCount) }), MessageStatus::Success);
}

#pragma endregion


#pragma region CSV

bool MultiReplace::confirmColumnDeletion() {
    // Attempt to parse columns and delimiters
    if (!parseColumnAndDelimiterData()) {
        return false;  // Parsing failed, exit with false indicating no confirmation
    }

    // Now columnDelimiterData should be populated with the parsed column data
    size_t columnCount = columnDelimiterData.columns.size();
    std::wstring confirmMessage = LM.get(L"msgbox_confirm_delete_columns", { std::to_wstring(columnCount) });

    // Display a message box with Yes/No options and a question mark icon
    int msgboxID = MessageBox(
        nppData._nppHandle,
        confirmMessage.c_str(),
        LM.get(L"msgbox_title_confirm").c_str(),
        MB_ICONWARNING | MB_YESNO
    );

    return (msgboxID == IDYES);  // Return true if user confirmed, else false
}

void MultiReplace::handleDeleteColumns()
{
    if (!validateDelimiterData()) {
        return;
    }

    send(SCI_BEGINUNDOACTION, 0, 0);

    int deletedFieldsCount = 0;
    SIZE_T lineCount = lineDelimiterPositions.size();

    // Iterate over lines from last to first.
    for (SIZE_T i = lineCount; --i > 0; ) {
        const auto& lineInfo = lineDelimiterPositions[i];

        // Get absolute start and end positions for this line.
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, i, 0);
        LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        // Get the actual EOL length for this line
        LRESULT eolLength = getEOLLengthForLine(i);

        // Vector to collect deletion ranges for the current line.
        std::vector<std::pair<LRESULT, LRESULT>> deletionRanges;

        // Iterate over columns using the set's reverse_iterator.
        for (auto it = columnDelimiterData.columns.rbegin(); it != columnDelimiterData.columns.rend(); ++it) {
            SIZE_T column = *it;

            // Only process columns that exist in this line.
            if (column > lineInfo.positions.size() + 1)
                continue;

            LRESULT startPos = 0, endPos = 0;
            // Calculate the absolute start position:
            if (column == 1) {
                startPos = lineStartPos;
            }
            else if (column - 2 < lineInfo.positions.size()) {
                startPos = lineStartPos + lineInfo.positions[column - 2].offsetInLine;
            }
            else {
                continue;  // Skip if the column index is invalid for this line.
            }

            // Calculate the absolute end position:
            if (column - 1 < lineInfo.positions.size()) {
                if (column == 1) {
                    // When deleting the first column, also delete the following delimiter.
                    endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine
                        + columnDelimiterData.delimiterLength;
                }
                else {
                    endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
                }
            }
            else {
                // Last column: delete until the end of the line.
                // For non-last lines, subtract actual EOL length to preserve the line break.
                if (i < lineCount - 1)
                    endPos = lineEndPos - eolLength;
                else
                    endPos = lineEndPos;
            }

            deletionRanges.push_back({ startPos, endPos });
        }

        // If deletion ranges were computed for this line, merge contiguous or overlapping ranges.
        if (!deletionRanges.empty()) {
            // Sort ranges by start position.
            std::sort(deletionRanges.begin(), deletionRanges.end(),
                [](const std::pair<LRESULT, LRESULT>& a, const std::pair<LRESULT, LRESULT>& b) {
                    return a.first < b.first;
                });
            std::vector<std::pair<LRESULT, LRESULT>> mergedRanges;
            auto currentRange = deletionRanges[0];
            for (size_t j = 1; j < deletionRanges.size(); ++j) {
                const auto& nextRange = deletionRanges[j];
                // If the next range starts before or exactly when the current range ends, merge them.
                if (nextRange.first <= currentRange.second) {
                    currentRange.second = std::max(currentRange.second, nextRange.second);
                }
                else {
                    mergedRanges.push_back(currentRange);
                    currentRange = nextRange;
                }
            }
            mergedRanges.push_back(currentRange);

            // Execute the deletions for all merged ranges in reverse order.
            // (Deleting from rightmost to leftmost prevents shifting of absolute positions.)
            for (auto it = mergedRanges.rbegin(); it != mergedRanges.rend(); ++it) {
                LRESULT lengthToDelete = it->second - it->first;
                if (lengthToDelete > 0) {
                    send(SCI_DELETERANGE, it->first, lengthToDelete, false);
                    ++deletedFieldsCount;
                }
            }
        }
    }

    send(SCI_ENDUNDOACTION, 0, 0);

    // Display a status message with the number of deleted fields.
    showStatusMessage(LM.get(L"status_deleted_fields_count", { std::to_wstring(deletedFieldsCount) }), MessageStatus::Success);
}

void MultiReplace::handleCopyColumnsToClipboard()
{
    if (!validateDelimiterData()) {
        return;
    }

    std::string combinedText;
    int copiedFieldsCount = 0;
    size_t lineCount = lineDelimiterPositions.size();


    // Iterate through each line
    for (size_t i = 0; i < lineCount; ++i) {
        const auto& lineInfo = lineDelimiterPositions[i];

        // Calculate absolute start/end for this line
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, i, 0);
        LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        bool isFirstCopiedColumn = true;
        std::string lineText;

        // Process each selected column
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= lineInfo.positions.size() + 1) {
                LRESULT startPos = 0;
                LRESULT endPos = 0;

                // Determine the absolute start position for this column
                if (column == 1) {
                    startPos = lineStartPos;
                    isFirstCopiedColumn = false;
                }
                else if (column - 2 < lineInfo.positions.size()) {
                    startPos = lineStartPos + lineInfo.positions[column - 2].offsetInLine;
                    // Drop the first delimiter if copied as the first column
                    if (isFirstCopiedColumn) {
                        startPos += columnDelimiterData.delimiterLength;
                        isFirstCopiedColumn = false;
                    }
                }
                else {
                    break; // No more valid columns
                }

                // Determine the absolute end position for this column
                if (column - 1 < lineInfo.positions.size()) {
                    endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
                }
                else {
                    endPos = lineEndPos;
                }

                // Buffer to hold the extracted text
                std::vector<char> buffer(static_cast<size_t>(endPos - startPos) + 1, '\0');

                // Prepare TextRange structure for Scintilla
                Sci_TextRangeFull tr;
                tr.chrg.cpMin = startPos;
                tr.chrg.cpMax = endPos;
                tr.lpstrText = buffer.data();

                // Extract text for the column
                send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
                lineText += std::string(buffer.data());

                ++copiedFieldsCount;
            }
        }

        combinedText += lineText;

        // Append standard EOL only if the last column is not included
        if (i < lineCount - 1 && (lineText.empty() || (combinedText.back() != '\n' && combinedText.back() != '\r'))) {
            combinedText += getEOLStyle();
        }
    }

    // Convert to Wide String and copy to clipboard
    std::wstring wstr = Encoding::bytesToWString(combinedText, getCurrentDocCodePage());
    copyTextToClipboard(wstr, copiedFieldsCount);
}

#pragma endregion


#pragma region CSV Sort

std::vector<CombinedColumns> MultiReplace::extractColumnData(SIZE_T startLine, SIZE_T endLine)
{
    std::vector<CombinedColumns> combinedData;
    combinedData.reserve(endLine - startLine); // Preallocate space to avoid reallocations

    for (SIZE_T i = startLine; i < endLine; ++i) {
        // Retrieve the LineInfo object
        const auto& lineInfo = lineDelimiterPositions[i];

        // Calculate absolute start position for this line in the Scintilla document
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, i, 0);

        // Calculate absolute end position (excluding EOL)
        LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        // Use the reusable lineBuffer from the class
        // Read the entire line content in one call to minimize SendMessage overhead
        size_t currentLineLength = static_cast<size_t>(lineInfo.lineLength);
        lineBuffer.resize(currentLineLength + 1, '\0'); // Adjust size for the current line

        {
            Sci_TextRangeFull tr;
            tr.chrg.cpMin = lineStartPos;
            tr.chrg.cpMax = lineEndPos;
            tr.lpstrText = lineBuffer.data();
            send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
        }

        CombinedColumns rowData;
        rowData.columns.resize(columnDelimiterData.inputColumns.size());

        for (size_t columnIndex = 0; columnIndex < columnDelimiterData.inputColumns.size(); ++columnIndex) {
            SIZE_T columnNumber = columnDelimiterData.inputColumns[columnIndex];

            // Determine the absolute start and end positions of this column in the document
            LRESULT startPos;
            if (columnNumber == 1) {
                // First column starts at the beginning of the line
                startPos = lineStartPos;
            }
            else if (columnNumber - 2 < lineInfo.positions.size()) {
                // Column after the first
                startPos = lineStartPos + lineInfo.positions[columnNumber - 2].offsetInLine
                    + columnDelimiterData.delimiterLength;
            }
            else {
                // Skip invalid columns if the delimiter doesn't exist
                continue;
            }

            // Determine the absolute end position
            LRESULT endPos;
            if (columnNumber - 1 < lineInfo.positions.size()) {
                // End of the current column = next delimiter offset
                endPos = lineStartPos + lineInfo.positions[columnNumber - 1].offsetInLine;
            }
            else {
                // Last column goes to the line's end
                endPos = lineEndPos;
            }

            // Map absolute positions back to local offsets in the line buffer
            const size_t localStart = static_cast<size_t>(startPos - lineStartPos);
            const size_t localEnd = static_cast<size_t>(endPos - lineStartPos);

            // Extract column text using local offsets
            if (localStart < localEnd && localEnd <= currentLineLength) {
                const size_t colSize = localEnd - localStart;
                std::string columnText(lineBuffer.data() + localStart, colSize);

                // Delete \n and \r at the end of the row
                while (!columnText.empty() && (columnText.back() == '\n' || columnText.back() == '\r')) {
                    columnText.pop_back();
                }

                rowData.columns[columnIndex].text = columnText;
            }
            else {
                // Empty column if the range is invalid
                rowData.columns[columnIndex].text.clear();
            }

        }

        // Add parsed row to the result
        combinedData.push_back(std::move(rowData));
    }

    // Optional step: detectNumericColumns if you want to parse columns as numeric
    detectNumericColumns(combinedData);

    return combinedData;
}

void MultiReplace::sortRowsByColumn(SortDirection sortDirection)
{
    // Validate delimiters before sorting
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(LM.get(L"status_invalid_column_or_delimiter"), MessageStatus::Error);
        return;
    }

    send(SCI_BEGINUNDOACTION, 0, 0);

    size_t lineCount = lineDelimiterPositions.size();
    if (lineCount <= CSVheaderLinesCount) {
        send(SCI_ENDUNDOACTION);
        return;
    }

    // Create an index array for all lines
    std::vector<size_t> tempOrder(lineCount);

    // Initialize vector with numeric sorted values
    std::iota(tempOrder.begin(), tempOrder.end(), 0);

    // Extract the columns after any header lines
    std::vector<CombinedColumns> combinedData = extractColumnData(CSVheaderLinesCount, lineCount);

    // Single-pass sort using multi-column comparison
    std::sort(tempOrder.begin() + CSVheaderLinesCount, tempOrder.end(),
        [&](size_t a, size_t b) {
            size_t indexA = a - CSVheaderLinesCount;
            size_t indexB = b - CSVheaderLinesCount;
            const auto& rowA = combinedData[indexA];
            const auto& rowB = combinedData[indexB];

            // Compare each column in priority order
            for (size_t colIndex = 0; colIndex < columnDelimiterData.inputColumns.size(); ++colIndex) {
                // If needed, map to actual column (e.g., size_t realIndex = columnDelimiterData.inputColumns[colIndex] - 1;)
                int cmp = compareColumnValue(rowA.columns[colIndex], rowB.columns[colIndex]);
                if (cmp != 0) {
                    return (sortDirection == SortDirection::Ascending) ? (cmp < 0) : (cmp > 0);
                }
            }
            // If all columns match, keep original order
            return false;
        }
    );

    // Update originalLineOrder if tracking is used
    if (!originalLineOrder.empty()) {
        std::vector<size_t> newOrder(originalLineOrder.size());
        for (size_t i = 0; i < tempOrder.size(); ++i) {
            newOrder[i] = originalLineOrder[tempOrder[i]];
        }
        originalLineOrder = std::move(newOrder);
    }
    else {
        originalLineOrder = tempOrder;
    }

    // Reorder lines in the editor based on sorted indices
    reorderLinesInScintilla(tempOrder);

    send(SCI_ENDUNDOACTION, 0, 0);
}

void MultiReplace::reorderLinesInScintilla(const std::vector<size_t>& sortedIndex) {
    std::string lineBreak = getEOLStyle();

    isSortedColumn = false; // Stop logging changes
    // Extract the text of each line based on the sorted index and include a line break after each
    std::string combinedLines;
    for (size_t i = 0; i < sortedIndex.size(); ++i) {
        size_t idx = sortedIndex[i];
        LRESULT lineStart = send(SCI_POSITIONFROMLINE, idx, 0);
        LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, idx, 0);
        std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1); // Buffer size includes space for null terminator
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = lineStart;
        tr.chrg.cpMax = lineEnd;
        tr.lpstrText = buffer.data();
        send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
        combinedLines += std::string(buffer.data(), buffer.size() - 1); // Exclude null terminator from the string
        if (i < sortedIndex.size() - 1) {
            combinedLines += lineBreak; // Add line break after each line except the last
        }
    }

    // Clear all content from Scintilla
    send(SCI_CLEARALL);

    // Re-insert the combined lines
    send(SCI_APPENDTEXT, combinedLines.length(), reinterpret_cast<sptr_t>(combinedLines.c_str()));

    isSortedColumn = true; // Ready for logging changes
}

void MultiReplace::restoreOriginalLineOrder(const std::vector<size_t>& originalOrder) {

    // Determine the total number of lines in the document
    size_t totalLineCount = static_cast<size_t>(send(SCI_GETLINECOUNT));

    // Ensure the size of the originalOrder vector matches the number of lines in the document
    auto maxElementIt = std::max_element(originalOrder.begin(), originalOrder.end());
    if (maxElementIt == originalOrder.end() || *maxElementIt != totalLineCount - 1) {
        return;
    }

    // Create a vector for the new sorted content of the document
    std::vector<std::string> sortedLines(totalLineCount);
    std::string lineBreak = getEOLStyle();

    // Iterate through each line in the document and fill sortedLines according to originalOrder
    for (size_t i = 0; i < totalLineCount; ++i) {
        size_t newPosition = originalOrder[i];
        LRESULT lineStart = send(SCI_POSITIONFROMLINE, i, 0);
        LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, i, 0);
        std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = lineStart;
        tr.chrg.cpMax = lineEnd;
        tr.lpstrText = buffer.data();
        send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
        sortedLines[newPosition] = std::string(buffer.data(), buffer.size() - 1); // Exclude null terminator
    }

    // Clear the content of the editor
    send(SCI_CLEARALL);

    // Re-insert the lines in their original order
    for (size_t i = 0; i < sortedLines.size(); ++i) {
        //std::string message = "Inserting line at position: " + std::to_string(i) + "\nContent: " + sortedLines[i];
        //MessageBoxA(NULL, message.c_str(), "Debug Insert Line", MB_OK);
        send(SCI_APPENDTEXT, sortedLines[i].length(), reinterpret_cast<sptr_t>(sortedLines[i].c_str()));
        // Add a line break after each line except the last one
        if (i < sortedLines.size() - 1) {
            send(SCI_APPENDTEXT, lineBreak.length(),reinterpret_cast<sptr_t>(lineBreak.c_str()));
        }
    }
}

void MultiReplace::extractLineContent(size_t idx, std::string& content, const std::string& lineBreak) {
    LRESULT lineStart = send(SCI_POSITIONFROMLINE, idx);
    LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, idx);
    std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
    Sci_TextRangeFull tr{ lineStart, lineEnd, buffer.data() };
    send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
    content.assign(buffer.begin(), buffer.end() - 1); // Remove the null terminator
    content += lineBreak;
}

void MultiReplace::UpdateSortButtonSymbols() {
    HWND hwndAscButton = GetDlgItem(_hSelf, IDC_COLUMN_SORT_ASC_BUTTON);
    HWND hwndDescButton = GetDlgItem(_hSelf, IDC_COLUMN_SORT_DESC_BUTTON);

    switch (currentSortState) {
    case SortDirection::Unsorted:
        SetWindowText(hwndAscButton, symbolSortAsc);
        SetWindowText(hwndDescButton, symbolSortDesc);
        break;
    case SortDirection::Ascending:
        SetWindowText(hwndAscButton, symbolSortAscUnsorted);
        SetWindowText(hwndDescButton, symbolSortDesc);
        break;
    case SortDirection::Descending:
        SetWindowText(hwndAscButton, symbolSortAsc);
        SetWindowText(hwndDescButton, symbolSortDescUnsorted);
        break;
    }
}

void MultiReplace::handleSortStateAndSort(SortDirection direction) {

    if (!validateDelimiterData()) {
        return;
    }

    if ((direction == SortDirection::Ascending && currentSortState == SortDirection::Ascending) ||
        (direction == SortDirection::Descending && currentSortState == SortDirection::Descending)) {
        isSortedColumn = false; //Disable logging of changes
        restoreOriginalLineOrder(originalLineOrder);
        currentSortState = SortDirection::Unsorted;
        originalLineOrder.clear();
    }
    else {
        currentSortState = (direction == SortDirection::Ascending) ? SortDirection::Ascending : SortDirection::Descending;
        if (columnDelimiterData.isValid()) {
            sortRowsByColumn(direction);
        }
    }
}

void MultiReplace::updateUnsortedDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType) {
    if (!isSortedColumn || lineNumber > originalLineOrder.size()) {
        return; // Invalid line number, return early
    }

    switch (changeType) {
    case ChangeType::Insert: {
        // Find the current maximum value in originalLineOrder and add one for the new line placeholder
        size_t maxIndex = originalLineOrder.empty() ? 0 : (*std::max_element(originalLineOrder.begin(), originalLineOrder.end())) + 1;

        //Insert multiple placeholders instead of just one
        std::vector<size_t> newIndices;
        newIndices.reserve(blockCount);

        for (SIZE_T i = 0; i < blockCount; ++i) {
            newIndices.push_back(maxIndex + i); 
        }

        // Insert them at the specified position
        originalLineOrder.insert(
            originalLineOrder.begin() + lineNumber,
            newIndices.begin(),
            newIndices.end()
        );
        break;
    }

    case ChangeType::Delete: {
        // Ensure lineNumber is within the range before attempting to delete
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > originalLineOrder.size()) {
            endPos = originalLineOrder.size();
        }

        if (lineNumber < originalLineOrder.size()) { // Safety check
            // Collect the removed indices 
            std::vector<size_t> removed(
                originalLineOrder.begin() + lineNumber,
                originalLineOrder.begin() + endPos
            );

            // Directly remove that range
            originalLineOrder.erase(
                originalLineOrder.begin() + lineNumber,
                originalLineOrder.begin() + endPos
            );

            // Adjust subsequent indices to reflect the deletion
            if (!removed.empty()) {
                size_t maxRemoved = *std::max_element(removed.begin(), removed.end());
                for (size_t i = 0; i < originalLineOrder.size(); ++i) {
                    if (originalLineOrder[i] > maxRemoved) {
                        originalLineOrder[i] -= (maxRemoved - lineNumber + 1);
                    }
                }
            }
        }
        break;
    }
    case ChangeType::Modify:
    default:
        // Do nothing for Modify
        break;
    }
}

void MultiReplace::detectNumericColumns(std::vector<CombinedColumns>& data)
{
    if (data.empty()) return;
    size_t colCount = data[0].columns.size();

    for (size_t col = 0; col < colCount; ++col) {
        for (auto& row : data) {
            std::string tmp = row.columns[col].text;

            // Skip empty fields to prevent false string classification
            if (tmp.empty()) continue;

            // If column is still numeric, set numeric values
            if (normalizeAndValidateNumber(tmp)) {
                row.columns[col].isNumeric = true;
                row.columns[col].numericValue = std::stod(tmp);
                row.columns[col].text = tmp; // keep '.' version if you like
            }
        }
    }
}

int MultiReplace::compareColumnValue(const ColumnValue& left, const ColumnValue& right)
{
    // If one is numeric and the other is not, numeric values come first
    if (left.isNumeric != right.isNumeric) {
        return left.isNumeric ? -1 : 1;
    }

    // If both are numeric, compare numerical values
    if (left.isNumeric) {
        if (left.numericValue < right.numericValue) return -1;
        if (left.numericValue > right.numericValue) return 1;
        return 0; // Values are equal
    }

    // If both are strings, compare text values
    return left.text.compare(right.text);
}

#pragma endregion


#pragma region Scope

bool MultiReplace::parseColumnAndDelimiterData() {
    // Retrieve data from dialog items
    std::wstring columnDataString = getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT);
    std::wstring delimiterData = getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT);
    std::wstring quoteCharString = getTextFromDialogItem(_hSelf, IDC_QUOTECHAR_EDIT);

    // Remove invalid delimiter characters (\n, \r)
    std::vector<std::wstring> stringsToRemove = { L"\\n", L"\\r" };
    for (const auto& strToRemove : stringsToRemove) {
        std::wstring::size_type pos = 0;
        while ((pos = delimiterData.find(strToRemove, pos)) != std::wstring::npos) {
            delimiterData.erase(pos, strToRemove.length());
        }
    }

    // Convert delimiter and quote character to standard strings
    std::string extendedDelimiter = convertAndExtendW(delimiterData, true);
    std::string quoteCharConverted = convertAndExtendW(quoteCharString, false);

    // Check for changes BEFORE modifying existing values
    // Cannot be used in LUA as it is Scintilla encoded and not UTF8
    bool delimiterChanged = (columnDelimiterData.extendedDelimiter != extendedDelimiter);
    bool quoteCharChanged = (columnDelimiterData.quoteChar != quoteCharConverted);

    // Trim leading and trailing commas from column data
    columnDataString.erase(0, columnDataString.find_first_not_of(L','));
    columnDataString.erase(columnDataString.find_last_not_of(L',') + 1);

    // Ensure that columnDataString and delimiter are not empty
    if (columnDataString.empty() || delimiterData.empty()) {
        showStatusMessage(LM.get(L"status_missing_column_or_delimiter_data"), MessageStatus::Error);
        return false;
    }

    // Use parseNumberRanges() to process column data
    std::vector<int> parsedColumns = parseNumberRanges(columnDataString, LM.get(L"status_invalid_range_in_column_data"));
    if (parsedColumns.empty()) return false; // Abort if parsing failed

    // Convert parsedColumns to set for uniqueness
    std::set<int> uniqueColumns(parsedColumns.begin(), parsedColumns.end());

    // Validate delimiter
    if (extendedDelimiter.empty()) {
        showStatusMessage(LM.get(L"status_extended_delimiter_empty"), MessageStatus::Error);
        return false;
    }

    // Validate quote character
    if (!quoteCharString.empty() && (quoteCharString.length() != 1 ||
        !(quoteCharString[0] == L'"' || quoteCharString[0] == L'\''))) {
        showStatusMessage(LM.get(L"status_invalid_quote_character"), MessageStatus::Error);
        return false;
    }

    // Check if columns have changed BEFORE updating stored data
    bool columnChanged = (columnDelimiterData.columns != uniqueColumns);

    // Update columnDelimiterData values
    columnDelimiterData.delimiterChanged = delimiterChanged;
    columnDelimiterData.quoteCharChanged = quoteCharChanged;
    columnDelimiterData.columnChanged = columnChanged;

    columnDelimiterData.inputColumns = std::move(parsedColumns);
    columnDelimiterData.columns = std::move(uniqueColumns);
    columnDelimiterData.extendedDelimiter = std::move(extendedDelimiter);
    columnDelimiterData.delimiterLength = columnDelimiterData.extendedDelimiter.length();
    columnDelimiterData.quoteChar = std::move(quoteCharConverted);

    return true;
}

bool MultiReplace::validateDelimiterData() {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED) {
        return parseColumnAndDelimiterData();
    }

    return true;
}

void MultiReplace::findAllDelimitersInDocument() {

    lineDelimiterPositions.clear();

    textModified = false;
    logChanges.clear();
    isLoggingEnabled = true;

    LRESULT totalLines = send(SCI_GETLINECOUNT, 0, 0);

    lineDelimiterPositions.reserve(totalLines);

    // Find and store delimiter positions for each line
    for (LRESULT line = 0; line < totalLines; ++line) {
        findDelimitersInLine(line);
    }

    lineBuffer.shrink_to_fit();
    logChanges.clear();

}

void MultiReplace::findDelimitersInLine(LRESULT line) {
    // Initialize line information
    LineInfo lineInfo;
    lineInfo.lineIndex = static_cast<size_t>(line);

    // Get line length
    LRESULT lineLength = send(SCI_LINELENGTH, line, 0);
    lineInfo.lineLength = lineLength;

    // Ensure buffer is large enough
    if (lineBuffer.size() < static_cast<size_t>(lineLength + 1))
        lineBuffer.resize(lineLength + 1);

    // Read line into buffer and create a string view
    send(SCI_GETLINE, line, reinterpret_cast<sptr_t>(lineBuffer.data()));
    std::string_view lineContent(lineBuffer.data(), static_cast<size_t>(lineLength));

    // Cache delimiter and quote character
    size_t delimLen = columnDelimiterData.delimiterLength;
    std::string_view delimiter(columnDelimiterData.extendedDelimiter);
    bool hasQuoteChar = !columnDelimiterData.quoteChar.empty();
    char currentQuoteChar = hasQuoteChar ? columnDelimiterData.quoteChar[0] : '\0';

    size_t pos = 0;
    bool inQuotes = false;

    while (pos < lineContent.size()) {
        // Toggle quote status if a quote character is found
        if (hasQuoteChar && lineContent[pos] == currentQuoteChar) {
            inQuotes = !inQuotes;
            ++pos;
            continue;
        }

        // Process delimiter if outside quotes
        if (!inQuotes) {
            if (delimLen == 1) {
                // Single-character delimiter check
                if (lineContent[pos] == delimiter[0]) {
                    lineInfo.positions.push_back({ static_cast<LRESULT>(pos) });
                    ++pos;
                    continue;
                }
            }
            else {
                // Multi-character delimiter search
                size_t foundPos = lineContent.find(delimiter, pos);
                if (foundPos != std::string_view::npos) {
                    if (hasQuoteChar) {
                        size_t nextQuote = lineContent.find(currentQuoteChar, pos);
                        if (nextQuote != std::string_view::npos && nextQuote < foundPos) {
                            pos = nextQuote;
                            continue;
                        }
                    }
                    lineInfo.positions.push_back({ static_cast<LRESULT>(foundPos) });
                    pos = foundPos + delimLen;
                    continue;
                }
                else {
                    break; // No more delimiters found
                }
            }
        }
        ++pos;
    }

    // Store line info in vector
    if (line < static_cast<LRESULT>(lineDelimiterPositions.size()))
        lineDelimiterPositions[line] = std::move(lineInfo);
    else {
        lineDelimiterPositions.resize(line + 1);
        lineDelimiterPositions[line] = std::move(lineInfo);
    }
}

ColumnInfo MultiReplace::getColumnInfo(LRESULT startPosition) {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED ||
        columnDelimiterData.columns.empty() ||
        columnDelimiterData.extendedDelimiter.empty() ||
        lineDelimiterPositions.empty())
    {
        return { 0, 0, 0 };
    }

    // Determine how many lines are in Scintilla
    LRESULT totalLines = send(SCI_GETLINECOUNT, 0, 0);

    // Determine which line in Scintilla corresponds to startPosition
    LRESULT startLine = send(SCI_LINEFROMPOSITION, startPosition, 0);
    SIZE_T  startColumnIndex = 1;

    // Check if the line index is valid in lineDelimiterPositions
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());
    if (startLine < totalLines && startLine < listSize) {
        const auto& lineInfo = lineDelimiterPositions[startLine];
        const auto& linePositions = lineInfo.positions;

        // Calculate absolute start of this line
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, startLine, 0);

        SIZE_T i = 0;
        for (; i < linePositions.size(); ++i) {
            // Absolute position of the i-th delimiter
            LRESULT delimAbsPos = lineStartPos + linePositions[i].offsetInLine;

            // If the caret is before or exactly on this delimiter, we've found our column
            if (startPosition <= delimAbsPos) {
                startColumnIndex = i + 1; // 1-based column
                break;
            }
        }

        // If we did not break, the caret is beyond the last delimiter => last column
        if (i == linePositions.size()) {
            startColumnIndex = linePositions.size() + 1;
        }
    }

    return { totalLines, startLine, startColumnIndex };
}

LRESULT MultiReplace::adjustForegroundForDarkMode(LRESULT textColor, LRESULT backgroundColor) {
    // Extract RGB components from text color
    int redText = (textColor & 0xFF);
    int greenText = ((textColor >> 8) & 0xFF);
    int blueText = ((textColor >> 16) & 0xFF);

    // Extract RGB components from background color
    int redBg = (backgroundColor & 0xFF);
    int greenBg = ((backgroundColor >> 8) & 0xFF);
    int blueBg = ((backgroundColor >> 16) & 0xFF);

    // Blend with background color to improve contrast
    float blendFactor = 0.8f;  // Higher value = stronger tint from background
    redText = static_cast<int>(redText * (1.0f - blendFactor) + redBg * blendFactor);
    greenText = static_cast<int>(greenText * (1.0f - blendFactor) + greenBg * blendFactor);
    blueText = static_cast<int>(blueText * (1.0f - blendFactor) + blueBg * blendFactor);

    // Increase brightness for better readability
    float brightnessBoost = 1.9f;  // Higher value = lighter text
    redText = std::min(255, static_cast<int>(redText * brightnessBoost));
    greenText = std::min(255, static_cast<int>(greenText * brightnessBoost));
    blueText = std::min(255, static_cast<int>(blueText * brightnessBoost));

    // Reconstruct adjusted color
    return (blueText << 16) | (greenText << 8) | redText;
}

void MultiReplace::initializeColumnStyles() {

    int IDM_LANG_TEXT = 46016;  // Switch off Languages - > Normal Text
    ::SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_LANG_TEXT);

    // Get default text color from Notepad++
    LRESULT fgColor = send(SCI_STYLEGETFORE, STYLE_DEFAULT);

    // Check if Notepad++ is in dark mode
    bool isDarkMode = SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0);

    // Select color scheme based on mode
    const auto& columnColors = isDarkMode ? darkModeColumnColors : lightModeColumnColors;

    for (SIZE_T column = 0; column < hColumnStyles.size(); ++column) {
        long bgColor = columnColors[column % columnColors.size()];
        LRESULT adjustedFgColor = isDarkMode ? adjustForegroundForDarkMode(fgColor, bgColor) : fgColor;

        send(SCI_STYLESETBACK, hColumnStyles[column], bgColor);
        send(SCI_STYLESETFORE, hColumnStyles[column], adjustedFgColor);
    }
}

void MultiReplace::handleHighlightColumnsInDocument() {
    if (!validateDelimiterData()) {
        return;
    }

    int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);
    highlightedTabs.mark(currentBufferID);
    initializeColumnStyles();
    LRESULT totalLines = static_cast<LRESULT>(lineDelimiterPositions.size());

    for (LRESULT line = 0; line < totalLines; ++line) {
        highlightColumnsInLine(line);
    }

    if (!lineDelimiterPositions.empty()) {
        LRESULT startPosition = send(SCI_GETCURRENTPOS, 0, 0);
        showStatusMessage(LM.get(L"status_actual_position", { addLineAndColumnMessage(startPosition) }), MessageStatus::Success);
    }

    isColumnHighlighted = true;
    isCaretPositionEnabled = true;
}

void MultiReplace::highlightColumnsInLine(LRESULT line) {
    // Retrieve the pre-parsed line information
    const auto& lineInfo = lineDelimiterPositions[line];

    // Skip empty lines
    if (lineInfo.lineLength == 0) {
        return;
    }

    // Create a styles vector with a size equal to the line length
    std::vector<char> styles(static_cast<size_t>(lineInfo.lineLength), 0);

    // If no delimiters are found and column 1 is defined, fill the entire line with column 1's style
    if (lineInfo.positions.empty() &&
        std::find(columnDelimiterData.columns.begin(), columnDelimiterData.columns.end(), 1) != columnDelimiterData.columns.end())
    {
        // Mask the style value to fit into 8 bits
        char style = static_cast<char>(hColumnStyles[0 % hColumnStyles.size()] & 0xFF);
        std::fill(styles.begin(), styles.end(), style);
    }
    else {
        // For each defined column, calculate the start and end offsets within the line
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= lineInfo.positions.size() + 1) {
                LRESULT start = 0;
                LRESULT end = 0;

                if (column == 1) {
                    start = 0;
                }
                else {
                    start = lineInfo.positions[column - 2].offsetInLine + columnDelimiterData.delimiterLength;
                }

                if (column == lineInfo.positions.size() + 1) {
                    end = lineInfo.lineLength;
                }
                else {
                    end = lineInfo.positions[column - 1].offsetInLine;
                }

                // Apply the style if the range is valid
                if (start < end && end <= static_cast<LRESULT>(styles.size())) {
                    char style = static_cast<char>(hColumnStyles[(column - 1) % hColumnStyles.size()] & 0xFF);
                    std::fill(styles.begin() + start, styles.begin() + end, style);
                }
            }
        }
    }

    // Determine the absolute starting position of the line in the document
    LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, line, 0);

    // Apply the computed styles to the document via Scintilla's API
    send(SCI_STARTSTYLING, lineStartPos, 0);
    send(SCI_SETSTYLINGEX, styles.size(), reinterpret_cast<sptr_t>(styles.data()));
}

void MultiReplace::handleClearColumnMarks() {
    // Get the current buffer ID
    int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

    // If the tab was not highlighted, exit early
    if (!highlightedTabs.isHighlighted(currentBufferID)) {
        return;
    }

    // Get total document length
    LRESULT textLength = send(SCI_GETLENGTH);

    // Reset all styling to default
    send(SCI_STARTSTYLING, 0);
    send(SCI_SETSTYLING, textLength, STYLE_DEFAULT);

    isColumnHighlighted = false;
    isCaretPositionEnabled = false;

    // Force Scintilla to recalculate word wrapping if highlighting affected layout
    int originalWrapMode = static_cast<int>(send(SCI_GETWRAPMODE));
    if (originalWrapMode != SC_WRAP_NONE) {
        send(SCI_SETWRAPMODE, SC_WRAP_NONE, 0);
        send(SCI_SETWRAPMODE, originalWrapMode, 0);
    }

    // Remove tab from tracked highlighted tabs
    highlightedTabs.clear(currentBufferID);
}

std::wstring MultiReplace::addLineAndColumnMessage(LRESULT pos) {
    if (!columnDelimiterData.isValid()) {
        return L"";
    }
    ColumnInfo startInfo = getColumnInfo(pos);
    std::wstring lineAndColumnMessage = LM.get(L"status_line_and_column_position",
        { std::to_wstring(startInfo.startLine + 1),
          std::to_wstring(startInfo.startColumnIndex) });

    return lineAndColumnMessage;
}

void MultiReplace::processLogForDelimiters() {
    // Check if logChanges is accessible
    if (!textModified || logChanges.empty()) {
        return;
    }

    std::vector<LogEntry> modifyLogEntries;

    // Loop through the log entries in chronological order
    for (auto& logEntry : logChanges) {
        switch (logEntry.changeType) {
        case ChangeType::Insert:
        {
            Sci_Position insertPos = logEntry.lineNumber;
            Sci_Position blockCount = logEntry.blockSize;

            // Shift existing modifies if needed
            for (auto& m : modifyLogEntries) {
                if (m.lineNumber >= insertPos) {
                    m.lineNumber += blockCount;
                }
            }

            // Insert new lines
            updateDelimitersInDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Insert
            );

            // Immediately parse them so they're recognized
            updateDelimitersInDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Modify
            );

            updateUnsortedDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Insert
            );

            // Optionally highlight all new lines
            if (isColumnHighlighted) {
                LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);
                for (Sci_Position offset = 0; offset < blockCount; ++offset) {
                    Sci_Position lineToHighlight = insertPos + offset;
                    if (lineToHighlight >= 0
                        && static_cast<size_t>(lineToHighlight) < lineDelimiterPositions.size()
                        && lineToHighlight < docLineCount)
                    {
                        highlightColumnsInLine(lineToHighlight);
                    }
                }
            }

            // Convert Insert to Modify for the final pass
            logEntry.changeType = ChangeType::Modify;
            modifyLogEntries.push_back(logEntry);
            break;
        }

        case ChangeType::Delete: 
        {
            Sci_Position deletePos = logEntry.lineNumber;
            Sci_Position blockCount = logEntry.blockSize;

            // remove modifies in [deletePos..deletePos+blockCount)
            for (auto& m : modifyLogEntries) {
                if (m.lineNumber >= deletePos && m.lineNumber < (deletePos + blockCount)) {
                    m.lineNumber = -1; // Mark for removal
                }
                else if (m.lineNumber >= (deletePos + blockCount)) {
                    m.lineNumber -= blockCount;
                }
            }

            updateDelimitersInDocument(
                static_cast<SIZE_T>(deletePos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Delete
            );

            updateUnsortedDocument(
                static_cast<SIZE_T>(deletePos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Delete
            );

            // Re-parse lines after deletePos to keep everything in sync
            if (deletePos < (Sci_Position)lineDelimiterPositions.size()) {
                Sci_Position linesToReparse = (Sci_Position)lineDelimiterPositions.size() - deletePos;
                updateDelimitersInDocument(
                    static_cast<SIZE_T>(deletePos),
                    static_cast<SIZE_T>(linesToReparse),
                    ChangeType::Modify
                );
            }
            break;
        }

        case ChangeType::Modify: 
        {
            modifyLogEntries.push_back(logEntry);
            break;
        }

        default:
            break;
        }
    }

    // Apply the saved "Modify" entries
    for (const auto& m : modifyLogEntries)
    {
        if (m.lineNumber == -1) {
            continue; // skip removed lines
        }

        if (static_cast<size_t>(m.lineNumber) < lineDelimiterPositions.size()) {
            updateDelimitersInDocument(
                static_cast<size_t>(m.lineNumber),
                1,
                ChangeType::Modify
            );

            if (isColumnHighlighted) {
                LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);
                if (m.lineNumber >= 0
                    && m.lineNumber < docLineCount
                    && static_cast<size_t>(m.lineNumber) < lineDelimiterPositions.size())
                {
                    highlightColumnsInLine(m.lineNumber);
                }
            }
        }
    }

    // Workaround: Highlight last lines to fix N++ bug causing loss of styling 
    if (isColumnHighlighted) {
        size_t lastLine = lineDelimiterPositions.size();
        LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);

        if (lastLine >= 2) {
            size_t highlightLine1 = lastLine - 2;
            if (highlightLine1 < (size_t)docLineCount) {
                highlightColumnsInLine((LRESULT)highlightLine1);
            }
            size_t highlightLine2 = lastLine - 1;
            if (highlightLine2 < (size_t)docLineCount) {
                highlightColumnsInLine((LRESULT)highlightLine2);
            }
        }
    }

    // Clear Log queue
    logChanges.clear();
    textModified = false;
}

void MultiReplace::updateDelimitersInDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType) {
    // Return early if the line number is invalid
    if (lineNumber > lineDelimiterPositions.size()) {
        return;
    }

    switch (changeType) {
    case ChangeType::Insert:
    {
        // Insert multiple empty lines instead of just one
        std::vector<LineInfo> newLines;
        newLines.reserve(blockCount);

        for (SIZE_T i = 0; i < blockCount; ++i) {
            LineInfo newLine;
            newLine.lineIndex = lineNumber;
            newLine.lineLength = 0; // New line has no content initially
            newLine.positions.clear();
            newLines.push_back(newLine);
        }

        // Add them all at the specified position
        lineDelimiterPositions.insert(
            lineDelimiterPositions.begin() + lineNumber,
            newLines.begin(),
            newLines.end()
        );

        // Update the indices of all subsequent lines
        for (SIZE_T i = lineNumber; i < lineDelimiterPositions.size(); ++i) {
            lineDelimiterPositions[i].lineIndex = i;
        }
    }
    break;

    case ChangeType::Delete:
    {
        // Remove the specified number of lines if they exist
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > lineDelimiterPositions.size()) {
            endPos = lineDelimiterPositions.size();
        }

        if (lineNumber < lineDelimiterPositions.size()) {
            lineDelimiterPositions.erase(
                lineDelimiterPositions.begin() + lineNumber,
                lineDelimiterPositions.begin() + endPos
            );

            // Update the indices of all subsequent lines
            for (SIZE_T i = lineNumber; i < lineDelimiterPositions.size(); ++i) {
                lineDelimiterPositions[i].lineIndex = i;
            }
        }
    }
    break;

    case ChangeType::Modify:
    {
        // Reanalyze multiple lines if they exist
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > lineDelimiterPositions.size()) {
            endPos = lineDelimiterPositions.size();
        }

        for (SIZE_T i = lineNumber; i < endPos; ++i) {
            findDelimitersInLine(i);
        }
    }
    break;

    default:
        break;
    }
}

void MultiReplace::handleDelimiterPositions(DelimiterOperation operation) {
    // Check if IDC_COLUMN_MODE_RADIO checkbox is not checked, and if so, return early
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED) {
        return;
    }

    // If there's been a change in the active window within Notepad++, reset all delimiter settings
    if (documentSwitched) {
        handleClearDelimiterState();
        documentSwitched = false;
    }

    if (operation == DelimiterOperation::LoadAll) {
        // Parse column and delimiter data; exit if parsing fails or if delimiter is empty
        if (!parseColumnAndDelimiterData()) {
            return;
        }

        // Trigger scan only if necessary
        if (columnDelimiterData.isValid() &&
            (columnDelimiterData.delimiterChanged ||
                columnDelimiterData.quoteCharChanged ||
                lineDelimiterPositions.empty()))
        {
            findAllDelimitersInDocument();
        }
    }
    else if (operation == DelimiterOperation::Update) {
        // Ensure the columns and delimiter data is present before processing updates
        if (columnDelimiterData.isValid()) {
            processLogForDelimiters();
        }
    }
}

void MultiReplace::handleClearDelimiterState() {
    lineDelimiterPositions.clear();
    isLoggingEnabled = false;
    textModified = false;
    logChanges.clear();
    if (isColumnHighlighted) {
        handleClearColumnMarks();
    }
    isCaretPositionEnabled = false;
}

/* For testing purposes only
void MultiReplace::displayLogChangesInMessageBox() {

    // Helper function to convert std::string to std::wstring using Windows API
    auto utf8ToWString = [](const std::string& input) -> std::wstring {
        if (input.empty()) return std::wstring();
        int size = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, 0, 0);
        std::wstring result(size, 0);
        MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, &result[0], size);
        return result;
        };

    // Create a helper function to convert list to string
    auto listToString = [this](const std::vector<LineInfo>& list) {
        std::stringstream ss;
        LRESULT listSize = static_cast<LRESULT>(list.size()); // Convert to signed type
        for (LRESULT i = 0; i < listSize; ++i) {
            ss << "Line: " << i << " Start: " << list[i].startPosition << " Positions: ";
            for (const auto& data : list[i].positions) {
                ss << data.position << " ";
            }
            ss << " End: " << list[i].endPosition << "\n";
        }
        return ss.str();
        };

    std::string startContent = listToString(lineDelimiterPositions);
    std::wstring wideStartContent = utf8ToWString(startContent);
    MessageBox(NULL, wideStartContent.c_str(), L"Content at the beginning", MB_OK);

    std::ostringstream oss;

    for (const auto& entry : logChanges) {
        switch (entry.changeType) {
        case ChangeType::Insert:
            oss << "Insert: ";
            break;
        case ChangeType::Modify:
            oss << "Modify: ";
            break;
        case ChangeType::Delete:
            oss << "Delete: ";
            break;
        }
        oss << "Line Number: " << entry.lineNumber << "\n";
    }

    std::string logChangesStr = oss.str();
    std::wstring logChangesWStr = utf8ToWString(logChangesStr);
    MessageBox(NULL, logChangesWStr.c_str(), L"Log Changes", MB_OK);

    processLogForDelimiters();
    std::wstring wideMessageBoxContent = utf8ToWString(messageBoxContent);
    MessageBox(NULL, wideMessageBoxContent.c_str(), L"Final Result", MB_OK);
    messageBoxContent.clear();

    std::string endContent = listToString(lineDelimiterPositions);
    std::wstring wideEndContent = utf8ToWString(endContent);
    MessageBox(NULL, wideEndContent.c_str(), L"Content at the end", MB_OK);

    logChanges.clear();
}
*/

#pragma endregion


#pragma region Utilities

static inline bool decodeNumericEscape(const std::wstring& src,size_t pos, int  base, int digits,wchar_t& out)
{
    if (pos + digits > src.size())           // not enough characters
        return false;

    unsigned value = 0;
    for (int k = 0; k < digits; ++k)
    {
        wchar_t ch = src[pos + k];
        unsigned v;

        if (ch >= L'0' && ch <= L'9') v = ch - L'0';
        else if (ch >= L'A' && ch <= L'F') v = (ch - L'A') + 10;
        else if (ch >= L'a' && ch <= L'f') v = (ch - L'a') + 10;
        else return false;                  // invalid digit

        if (v >= static_cast<unsigned>(base))
            return false;                   // digit not allowed in this base

        value = value * base + v;
    }

    if (value > 0xFFFF)                     // outside BMP range
        return false;

    out = static_cast<wchar_t>(value);
    return true;
}

std::string MultiReplace::convertAndExtendW(const std::wstring& input, bool extended, UINT targetCodepage) const
{
    if (!extended)                          // fast path – no escapes
        return Encoding::wstringToString(input, targetCodepage);

    std::wstring out;
    out.reserve(input.size());

    for (size_t i = 0; i < input.size(); ++i)
    {
        wchar_t ch = input[i];

        if (ch != L'\\' || i + 1 >= input.size())
        {
            out.push_back(ch);
            continue;
        }

        wchar_t esc = input[++i];           // escape designator
        wchar_t decoded;

        switch (esc)
        {
        case L'r': out.push_back(L'\r');                    break;
        case L'n': out.push_back(L'\n');                    break;
        case L't': out.push_back(L'\t');                    break;
        case L'\\':out.push_back(L'\\');                    break;
        case L'0': out.push_back(L'\0');                    break;

        case L'o': // \oNNN  (octal)
            if (decodeNumericEscape(input, i + 1, 8, 3, decoded))
            {
                out.push_back(decoded); i += 3; break;
            }
            [[fallthrough]];                                // literal fallback

        case L'd': // \dNNN  (decimal)
            if (decodeNumericEscape(input, i + 1, 10, 3, decoded))
            {
                out.push_back(decoded); i += 3; break;
            }
            [[fallthrough]];

        case L'x': // \xHH   (hex-byte)
            if (decodeNumericEscape(input, i + 1, 16, 2, decoded))
            {
                out.push_back(decoded); i += 2; break;
            }
            [[fallthrough]];

        case L'u': // \uXXXX (Unicode BMP)
            if (decodeNumericEscape(input, i + 1, 16, 4, decoded))
            {
                out.push_back(decoded); i += 4; break;
            }
            [[fallthrough]];

        default:  // unknown or invalid sequence → keep literally
            out.append({ L'\\', esc });
            break;
        }
    }

    return Encoding::wstringToString(out, targetCodepage);
}

std::string MultiReplace::convertAndExtendW(const std::wstring& input, bool extended)
{
    UINT cp = getCurrentDocCodePage();
    return convertAndExtendW(input, extended, cp);
}

void MultiReplace::addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems)
{
    if (str.length() == 0)
    {
        // Skip adding empty strings to the combo box history
        return;
    }

    // Check if the string is already in the combo box
    int index = static_cast<int>(SendMessage(hComboBox, CB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(str.c_str())));

    // If the string is not found, insert it at the beginning
    if (index == CB_ERR)
    {
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str.c_str()));

        // Remove the last item if the list exceeds maxItems
        if (SendMessage(hComboBox, CB_GETCOUNT, 0, 0) > maxItems)
        {
            SendMessage(hComboBox, CB_DELETESTRING, maxItems, 0);
        }
    }
    else
    {
        // If the string is found, move it to the beginning
        SendMessage(hComboBox, CB_DELETESTRING, index, 0);
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str.c_str()));
    }

    // Select the newly added string
    SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
}

std::wstring MultiReplace::getTextFromDialogItem(HWND hwnd, int itemID)
{
    HWND hCtrl = GetDlgItem(hwnd, itemID);
    if (!hCtrl) return L"";

    int len = GetWindowTextLengthW(hCtrl);
    if (len <= 0) return L"";

    std::wstring wbuf(static_cast<size_t>(len) + 1, L'\0');
    int written = GetDlgItemTextW(hwnd, itemID, &wbuf[0], len + 1);
    if (written < 0) written = 0;
    if (!wbuf.empty() && wbuf.back() == L'\0') wbuf.pop_back();
    wbuf.resize(static_cast<size_t>(written));
    return wbuf;
}

void MultiReplace::setSelections(bool select, bool onlySelected) {
    if (replaceListData.empty()) {
        return;
    }

    // Store the original data for all affected items
    std::vector<std::pair<size_t, ReplaceItemData>> originalDataList;
    for (size_t i = 0; i < replaceListData.size(); ++i) {
        if (!onlySelected || ListView_GetItemState(_replaceListView, static_cast<int>(i), LVIS_SELECTED)) {
            originalDataList.emplace_back(i, replaceListData[i]);
            replaceListData[i].isEnabled = select;
        }
    }

    // Update the allSelected flag if all items were selected/deselected
    if (!onlySelected) {
        allSelected = select;
    }

    // Update the header after changing the selection status of the items
    updateHeaderSelection();

    // Update the ListView
    for (const auto& [index, _] : originalDataList) {
        updateListViewItem(index);
    }

    // Create Undo/Redo actions
    UndoRedoAction action;

    // Undo action: Restore the original data
    action.undoAction = [this, originalDataList]() {
        for (const auto& [index, data] : originalDataList) {
            replaceListData[index] = data;
            updateListViewItem(index);
        }
        updateHeaderSelection();
        };

    // Redo action: Apply the new selection status
    action.redoAction = [this, originalDataList]() {
        for (const auto& [index, data] : originalDataList) {
            replaceListData[index].isEnabled = !data.isEnabled;
            updateListViewItem(index);
        }
        updateHeaderSelection();
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Set enabled");

    // Show Select Statisics
    showListFilePath();
}

void MultiReplace::updateHeaderSelection() {
    bool anySelected = false;
    allSelected = !replaceListData.empty();

    // Check if any or all items in the replaceListData vector are selected
    for (const auto& itemData : replaceListData) {
        if (itemData.isEnabled) {
            anySelected = true;
        }
        else {
            allSelected = false;
        }
    }

    // Initialize the LVCOLUMN structure
    LVCOLUMN lvc = { 0 };
    lvc.mask = LVCF_TEXT;

    // Determine the symbol to show in the header
    if (allSelected) {
        lvc.pszText = L"\u25A0"; // Ballot box with check
    }
    else if (anySelected) {
        lvc.pszText = L"\u25A3"; // Black square containing small white square
    }
    else {
        lvc.pszText = L"\u2610"; // Ballot box without check
    }

    // Update the Selection column header dynamically, if it's enabled
    if (columnIndices[ColumnID::SELECTION] != -1) {
        ListView_SetColumn(_replaceListView, columnIndices[ColumnID::SELECTION], &lvc);
    }

}

void MultiReplace::updateHeaderSortDirection() {
    const wchar_t* ascendingSymbol = L" ▲";
    const wchar_t* descendingSymbol = L" ▼";
    const wchar_t* lockedSymbol = L" 🔒"; // Lock symbol for locked columns

    // Iterate through all columns in columnIndices
    for (const auto& [columnID, columnIndex] : columnIndices) {
        // Skip columns that are not visible (columnIndex == -1)
        if (columnIndex == -1) {
            continue;
        }

        // Only update headers for sortable columns
        if (columnID != ColumnID::FIND_COUNT &&
            columnID != ColumnID::REPLACE_COUNT &&
            columnID != ColumnID::FIND_TEXT &&
            columnID != ColumnID::REPLACE_TEXT &&
            columnID != ColumnID::COMMENTS) {
            continue;
        }

        // Get the base header text
        std::wstring headerText;
        switch (columnID) {
        case ColumnID::FIND_COUNT:
            headerText = LM.get(L"header_find_count");
            break;
        case ColumnID::REPLACE_COUNT:
            headerText = LM.get(L"header_replace_count");
            break;
        case ColumnID::FIND_TEXT:
            headerText = LM.get(L"header_find");
            // Append lock symbol if the FIND_TEXT column is locked
            if (findColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::REPLACE_TEXT:
            headerText = LM.get(L"header_replace");
            // Append lock symbol if the REPLACE_TEXT column is locked
            if (replaceColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::COMMENTS:
            headerText = LM.get(L"header_comments");
            if (commentsColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        default:
            continue; // Skip if it's not a sortable column
        }

        // Append sort symbol if the column is currently sorted
        auto sortIt = columnSortOrder.find(columnID);
        if (sortIt != columnSortOrder.end()) {
            SortDirection direction = sortIt->second;
            if (direction == SortDirection::Ascending) {
                headerText += ascendingSymbol;
            }
            else if (direction == SortDirection::Descending) {
                headerText += descendingSymbol;
            }
            // Do not append any symbol if direction is Unsorted
        }

        // Prepare the LVCOLUMN structure for updating the header
        LVCOLUMN lvc = {};
        lvc.mask = LVCF_TEXT;
        lvc.pszText = const_cast<LPWSTR>(headerText.c_str());

        // Update the column header using the correct index
        ListView_SetColumn(_replaceListView, columnIndex, &lvc);
    }
}

void MultiReplace::showStatusMessage(const std::wstring& messageText, MessageStatus status, bool isNotFound)
{
    const size_t MAX_DISPLAY_LENGTH = 150;

    std::wstring strMessage;
    strMessage.reserve(messageText.length());
    for (wchar_t ch : messageText) {
        if (iswprint(ch)) {
            strMessage += ch;
        }
    }

    if (strMessage.size() > MAX_DISPLAY_LENGTH) {
        strMessage = strMessage.substr(0, MAX_DISPLAY_LENGTH - 3) + L"...";
    }

    // Store the TYPE of the message for later theme switches
    _lastMessageStatus = status;

    // Set the active color based on the new status
    switch (status) {
    case MessageStatus::Success:
        _statusMessageColor = COLOR_SUCCESS;
        break;
    case MessageStatus::Error:
        _statusMessageColor = COLOR_ERROR;
        break;
    case MessageStatus::Info:
    default:
        _statusMessageColor = COLOR_INFO;
        break;
    }

    HWND hStatusMessage = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
    SetWindowText(hStatusMessage, strMessage.c_str());

    // --- KEY CHANGE: Invalidate the PARENT window's area BEHIND the control ---
    // This forces the N++ themed dialog to erase the background for us with the correct color.
    RECT rect;
    GetWindowRect(hStatusMessage, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hStatusMessage), (LPPOINT)&rect, 2);
    InvalidateRect(GetParent(hStatusMessage), &rect, TRUE); // TRUE means "erase background"
    UpdateWindow(GetParent(hStatusMessage));

    if (isNotFound && alertNotFoundEnabled)
    {
        FLASHWINFO fwInfo = { 0 };
        fwInfo.cbSize = sizeof(FLASHWINFO);
        fwInfo.hwnd = _hSelf;
        fwInfo.dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG;
        fwInfo.uCount = 0;
        fwInfo.dwTimeout = 0;
        FlashWindowEx(&fwInfo);
    }
}

void MultiReplace::applyThemePalette()
{
    // Check if Notepad++ is currently in dark mode
    BOOL isDarkMode = (SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0);

    // Assign colours from the predefined palettes in the header file
    if (isDarkMode) {
        COLOR_SUCCESS = DMODE_SUCCESS;
        COLOR_ERROR = DMODE_ERROR;
        COLOR_INFO = DMODE_INFO;
        _filterHelpColor = DMODE_FILTER_HELP;
    }
    else {
        COLOR_SUCCESS = LMODE_SUCCESS;
        COLOR_ERROR = LMODE_ERROR;
        COLOR_INFO = LMODE_INFO;
        _filterHelpColor = LMODE_FILTER_HELP;
    }

    // Update the active colour based on the last message status
    switch (_lastMessageStatus) {
    case MessageStatus::Success: _statusMessageColor = COLOR_SUCCESS; break;
    case MessageStatus::Error:   _statusMessageColor = COLOR_ERROR;   break;
    default:                     _statusMessageColor = COLOR_INFO;    break;
    }

    // Repaint status control and "(?)" tooltip so they pick up the new palette
    InvalidateRect(GetDlgItem(_hSelf, IDC_STATUS_MESSAGE), NULL, TRUE);
    InvalidateRect(GetDlgItem(_hSelf, IDC_FILTER_HELP), NULL, TRUE);
}

void MultiReplace::refreshColumnStylesIfNeeded()
{
    if (isColumnHighlighted) {          // flag is managed by highlight/clear handlers
        initializeColumnStyles();
    }
}

std::wstring MultiReplace::getShortenedFilePath(const std::wstring& path, int maxLength, HDC hDC) {
    bool hdcProvided = true;

    // If no HDC is provided, get the one for the main window (_hSelf)
    if (hDC == nullptr) {
        hDC = GetDC(_hSelf);  // Get the device context for _hSelf (less accurate)
        hdcProvided = false;  // Mark that HDC was not provided externally
    }

    double dotWidth = 0.0;
    SIZE charSize;
    std::vector<double> characterWidths;

    // Calculate the width of each character in the path and the width of the dots ("...")
    for (wchar_t ch : path) {
        GetTextExtentPoint32(hDC, &ch, 1, &charSize);
        characterWidths.push_back(static_cast<double>(charSize.cx));
        if (ch == L'.') {
            dotWidth = static_cast<double>(charSize.cx);  // Store width of '.' separately
        }
    }

    double totalDotsWidth = dotWidth * 3;  // Width for "..."

    // Split the directory and filename
    size_t lastSlashPos = path.find_last_of(L"\\/");
    std::wstring directoryPart = (lastSlashPos != std::wstring::npos) ? path.substr(0, lastSlashPos + 1) : L"";
    std::wstring fileName = (lastSlashPos != std::wstring::npos) ? path.substr(lastSlashPos + 1) : path;

    // Calculate widths for directory and file name
    double directoryWidth = 0.0, fileNameWidth = 0.0;

    for (size_t i = 0; i < directoryPart.size(); ++i) {
        directoryWidth += characterWidths[i];
    }

    for (size_t i = directoryPart.size(); i < path.size(); ++i) {
        fileNameWidth += characterWidths[i];
    }

    std::wstring displayPath;
    double currentWidth = 0.0;

    // Shorten the file name if necessary
    if (fileNameWidth + totalDotsWidth > maxLength) {
        for (size_t i = directoryPart.size(); i < path.size(); ++i) {
            if (currentWidth + characterWidths[i] + totalDotsWidth > maxLength) {
                break;
            }
            displayPath += path[i];
            currentWidth += characterWidths[i];
        }
        displayPath += L"...";
    }
    // Shorten the directory part if necessary
    else if (directoryWidth + fileNameWidth > maxLength) {
        for (size_t i = 0; i < directoryPart.size(); ++i) {
            if (currentWidth + characterWidths[i] + totalDotsWidth + fileNameWidth > maxLength) {
                break;
            }
            displayPath += directoryPart[i];
            currentWidth += characterWidths[i];
        }
        displayPath += L"...";
        displayPath += fileName;
    }
    else {
        displayPath = path; // No shortening needed
    }

    // If we obtained HDC ourselves, release it
    if (!hdcProvided) {
        ReleaseDC(_hSelf, hDC);
    }

    return displayPath;
}

void MultiReplace::showListFilePath()
{
    HWND hPathDisplay = GetDlgItem(_hSelf, IDC_PATH_DISPLAY);
    HWND hStatsDisplay = GetDlgItem(_hSelf, IDC_STATS_DISPLAY);
    HWND hListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    if (!hPathDisplay || !hListView)
        return;

    HDC hDC = GetDC(hPathDisplay);
    HFONT hFont = (HFONT)SendMessage(hPathDisplay, WM_GETFONT, 0, 0);
    SelectObject(hDC, hFont);

    // Get ListView dimensions
    RECT rcListView;
    GetWindowRect(hListView, &rcListView);
    MapWindowPoints(NULL, _hSelf, reinterpret_cast<LPPOINT>(&rcListView), 2);
    int listWidth = rcListView.right - rcListView.left;

    const int spacing = sx(10);
    const int verticalOffset = sy(2);

    // Calculate Y positions (keeping original heights)
    RECT rcPathDisplay;
    GetClientRect(hPathDisplay, &rcPathDisplay);
    int fieldHeight = rcPathDisplay.bottom - rcPathDisplay.top;
    int fieldY = rcListView.bottom + verticalOffset;

    int statsWidth = 0;

    if (listStatisticsEnabled && hStatsDisplay)
    {
        // Gather statistics
        int totalItems = static_cast<int>(replaceListData.size());
        int selectedItems = ListView_GetSelectedCount(_replaceListView);
        int activatedCount = static_cast<int>(std::count_if(replaceListData.begin(), replaceListData.end(),
            [](const ReplaceItemData& item) { return item.isEnabled; }));
        int focusedRow = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
        int displayRow = (selectedItems > 0 && focusedRow != -1) ? (focusedRow + 1) : 0;

        // Compose stats string
        std::wstring statsString = L"A:" + std::to_wstring(activatedCount)
            + L"  L:" + std::to_wstring(totalItems)
            + L"  |  R:" + std::to_wstring(displayRow)
            + L"  S:" + std::to_wstring(selectedItems);

        // Measure stats string width
        SIZE sz;
        GetTextExtentPoint32(hDC, statsString.c_str(), static_cast<int>(statsString.size()), &sz);
        statsWidth = sz.cx + sx(5); // padding

        // Position stats field
        int statsX = rcListView.left + listWidth - statsWidth;
        MoveWindow(hStatsDisplay, statsX, fieldY, statsWidth, fieldHeight, TRUE);
        SetWindowTextW(hStatsDisplay, statsString.c_str());
        ShowWindow(hStatsDisplay, SW_SHOW);
    }
    else if (hStatsDisplay)
    {
        // Pragmatic solution: Set width to zero instead of hiding
        statsWidth = 0;
        MoveWindow(hStatsDisplay, rcListView.right, fieldY, 0, fieldHeight, TRUE);
        SetWindowTextW(hStatsDisplay, L"");
        ShowWindow(hStatsDisplay, SW_HIDE);
    }

    // Adjust path field to use remaining space
    int pathWidth = listWidth - statsWidth - (listStatisticsEnabled ? spacing : 0);
    pathWidth = std::max(pathWidth, 0);
    MoveWindow(hPathDisplay, rcListView.left, fieldY, pathWidth, fieldHeight, TRUE);

    // Update path display text
    std::wstring shortenedPath = getShortenedFilePath(listFilePath, pathWidth, hDC);
    SetWindowTextW(hPathDisplay, shortenedPath.c_str());

    ReleaseDC(hPathDisplay, hDC);

    // Immediate redraw
    InvalidateRect(hPathDisplay, NULL, TRUE);
    UpdateWindow(hPathDisplay);
    if (hStatsDisplay) {
        InvalidateRect(hStatsDisplay, NULL, TRUE);
        UpdateWindow(hStatsDisplay);
    }
}

std::wstring MultiReplace::getSelectedText()
{
    Sci_Position lengthWithNul = static_cast<Sci_Position>(SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0, 0));
    if (lengthWithNul <= 0) return L"";

    std::string buffer(static_cast<size_t>(lengthWithNul), '\0');
    SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0,
        reinterpret_cast<LPARAM>(buffer.data()));

    if (!buffer.empty() && buffer.back() == '\0') buffer.pop_back();

    UINT sciCp = getCurrentDocCodePage();
    return Encoding::bytesToWString(buffer, sciCp);
}

LRESULT MultiReplace::getEOLLengthForLine(LRESULT line) {
    LRESULT lineLen = send(SCI_LINELENGTH, line, 0);
    if (lineLen == 0) {
        return 0; // Empty line, no EOL
    }

    // We'll check up to the last 2 chars only:
    LRESULT lineStart = send(SCI_POSITIONFROMLINE, line, 0);
    LRESULT checkCount = (lineLen >= 2) ? 2 : 1;

    // Read those chars via SCI_GETCHARAT
    char last[2] = { 0, 0 };
    for (LRESULT i = 0; i < checkCount; ++i) {
        last[i] = static_cast<char>(
            send(SCI_GETCHARAT, lineStart + lineLen - checkCount + i, 0)
            );
    }

    // If we have 2 chars and they are '\r\n', that's EOL=2
    if (checkCount == 2 && last[0] == '\r' && last[1] == '\n') {
        return 2;
    }
    // Otherwise check if the last char is '\r' or '\n'
    else if (last[checkCount - 1] == '\r' || last[checkCount - 1] == '\n') {
        return 1;
    }
    return 0;
}

std::string MultiReplace::getEOLStyle() {
    LRESULT eolMode = static_cast<LRESULT>(send(SCI_GETEOLMODE));
    switch (eolMode) {
    case SC_EOL_CRLF:
        return "\r\n";
    case SC_EOL_CR:
        return "\r";
    case SC_EOL_LF:
        return "\n";
    default:
        return "\n";  // Defaulting to LF
    }
}

sptr_t MultiReplace::send(unsigned int iMessage, uptr_t wParam, sptr_t lParam, bool useDirect) const {
    if (useDirect && pSciMsg) {
        return pSciMsg(pSciWndData, iMessage, wParam, lParam);
    }
    else {
        return ::SendMessage(_hScintilla, iMessage, wParam, lParam);
    }
}

bool MultiReplace::normalizeAndValidateNumber(std::string& str) {
    if (str.empty()) {
        return false;  // An empty string is not a valid number
    }

    if (str == "." || str == ",") {
        return false;
    }

    int dotCount = 0;
    std::string tempStr = str; // Temporary string to hold potentially modified string
    for (char& c : tempStr) {
        if (c == '.') {
            ++dotCount;
        }
        else if (c == ',') {
            ++dotCount;
            c = '.';  // Potentially replace comma with dot in tempStr
        }
        else if (!iswdigit(c)) {
            return false;  // Contains non-numeric characters
        }

        if (dotCount > 1) {
            return false;  // Contains more than one separator
        }
    }

    str = tempStr;
    return true;  // String is a valid number
}

std::vector<WCHAR> MultiReplace::createFilterString(const std::vector<std::pair<std::wstring, std::wstring>>& filters) {
    // Calculate the required size for the filter string
    size_t totalSize = 0;
    for (const auto& filter : filters) {
        totalSize += filter.first.size() + 1; // Description + null terminator
        totalSize += filter.second.size() + 1; // Pattern + null terminator
    }
    totalSize += 1; // Double null terminator at the end

    // Create the array
    std::vector<WCHAR> filterString;
    filterString.reserve(totalSize);

    // Fill the array
    for (const auto& filter : filters) {
        filterString.insert(filterString.end(), filter.first.begin(), filter.first.end());
        filterString.push_back(L'\0');
        filterString.insert(filterString.end(), filter.second.begin(), filter.second.end());
        filterString.push_back(L'\0');
    }

    // End with an additional null terminator
    filterString.push_back(L'\0');

    return filterString;
}

int MultiReplace::getFontHeight(HWND hwnd, HFONT hFont) {
    // Get the font size from the specified font and window
    TEXTMETRIC tm;
    HDC hdc = GetDC(hwnd);  // Get the device context for the specified window
    SelectObject(hdc, hFont);  // Select the specified font into the DC
    GetTextMetrics(hdc, &tm);  // Retrieve the text metrics for the font
    int fontHeight = tm.tmHeight;  // Extract the font height
    ReleaseDC(hwnd, hdc);  // Release the device context
    return fontHeight;  // Return the font height
}

std::vector<int> MultiReplace::parseNumberRanges(const std::wstring& input, const std::wstring& errorMessage)
{
    std::vector<int> result;
    if (input.empty()) return result;  // nothing to parse

    // use a hash set to filter out duplicates, but preserve insertion order in 'result'
    std::unordered_set<int> seen;

    std::wistringstream stream(input);
    std::wstring token;

    // helper to add a number only once, in the order encountered
    auto pushUnique = [&](int n) {
        if (seen.insert(n).second)      // if n was not already present
            result.push_back(n);
        };

    // process each comma-separated token, either a single number or a range "start-end"
    auto processToken = [&](const std::wstring& token) -> bool
        {
            if (token.empty()) return true;  // skip empty entries

            try {
                size_t dashPos = token.find(L'-');
                if (dashPos != std::wstring::npos) {
                    // RANGE: parse start and end values
                    int startRange = std::stoi(token.substr(0, dashPos));
                    int endRange = std::stoi(token.substr(dashPos + 1));
                    if (startRange < 1 || endRange < 1)
                        return false;

                    // push each number in the range, preserving order
                    if (endRange >= startRange) {
                        // ascending range
                        for (int i = startRange; i <= endRange; ++i)
                            pushUnique(i);
                    }
                    else {
                        // descending range
                        for (int i = startRange; i >= endRange; --i)
                            pushUnique(i);
                    }
                }
                else {
                    // SINGLE NUMBER
                    int number = std::stoi(token);
                    if (number < 1)
                        return false;  // invalid value

                    // add the single number
                    pushUnique(number);
                }
            }
            catch (const std::exception&) {
                return false;  // parsing failure
            }
            return true;
        };

    // split the input string on commas and feed each token to our processor
    while (std::getline(stream, token, L',')) {
        if (!processToken(token)) {
            showStatusMessage(errorMessage, MessageStatus::Error);
            return {};  // abort on invalid syntax
        }
    }

    // 'result' contains the right values in the correct order
    return result;
}

UINT MultiReplace::getCurrentDocCodePage()
{
    // Ask Scintilla for the current code page (always non‑negative)
    LRESULT cp = send(SCI_GETCODEPAGE, 0, 0);

    // If Scintilla ever returned 0 (undefined), fall back to ANSI
    return static_cast<UINT>(cp != 0 ? cp : CP_ACP);
}

Sci_Position MultiReplace::advanceAfterMatch(const SearchResult& r)  {
    if (r.length > 0) return r.pos + r.length;
    const Sci_Position after = static_cast<Sci_Position>(send(SCI_POSITIONAFTER, r.pos, 0));
    const Sci_Position next = (after > r.pos) ? after : (r.pos + 1);
    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    return (next > docLen) ? docLen : next;
}

Sci_Position MultiReplace::ensureForwardProgress(Sci_Position candidate, const SearchResult& last)  {
    if (candidate > last.pos) return candidate;
    const Sci_Position after = static_cast<Sci_Position>(send(SCI_POSITIONAFTER, last.pos, 0));
    const Sci_Position next = (after > last.pos) ? after : (last.pos + 1);
    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    return (next > docLen) ? docLen : next;
}

#pragma endregion


#pragma region StringHandling

#pragma endregion


#pragma region FileOperations

std::wstring MultiReplace::openFileDialog(bool saveFile, const std::vector<std::pair<std::wstring, std::wstring>>& filters, const WCHAR* title, DWORD flags, const std::wstring& fileExtension, const std::wstring& defaultFilePath) {
    flags |= OFN_NOCHANGEDIR;
    OPENFILENAME ofn = { 0 };
    WCHAR szFile[MAX_PATH] = { 0 };

    // Safely copy the default file path into the buffer and ensure null-termination
    if (!defaultFilePath.empty()) {
        wcsncpy_s(szFile, defaultFilePath.c_str(), MAX_PATH);
    }

    std::vector<WCHAR> filter = createFilterString(filters);

    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = _hSelf;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile) / sizeof(WCHAR);
    ofn.lpstrFilter = filter.data();
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = title;
    ofn.Flags = flags;

    if (saveFile ? GetSaveFileName(&ofn) : GetOpenFileName(&ofn)) {
        std::wstring filePath(szFile);

        // If no extension is provided, append the default one
        if (filePath.find_last_of(L".") == std::wstring::npos) {
            filePath += L"." + fileExtension;
        }

        return filePath;
    }
    else {
        return std::wstring();
    }
}

std::wstring MultiReplace::promptSaveListToCsv() {
    std::wstring csvDescription = LM.get(L"filetype_csv");  // "CSV Files (*.csv)"
    std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

    std::vector<std::pair<std::wstring, std::wstring>> filters = {
        {csvDescription, L"*.csv"},
        {allFilesDescription, L"*.*"}
    };

    std::wstring dialogTitle = LM.get(L"panel_save_list");
    std::wstring defaultFileName;

    if (!listFilePath.empty()) {
        // If a file path already exists, use its directory and filename
        defaultFileName = listFilePath;
    }
    else {
        // If no file path is set, provide a default file name with a sequential number
        static int fileCounter = 1;
        defaultFileName = L"Replace_List_" + std::to_wstring(++fileCounter) + L".csv";
    }

    // Call openFileDialog with the default file path and name
    std::wstring filePath = openFileDialog(
        true,
        filters,
        dialogTitle.c_str(),
        OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT,
        L"csv",
        defaultFileName
    );

    return filePath;
}

bool MultiReplace::saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    std::ofstream outFile(std::filesystem::path(filePath), std::ios::binary);

    if (!outFile.is_open()) {
        return false;
    }

    // [BOM] Write UTF-8 BOM for CSV
    outFile.write("\xEF\xBB\xBF", 3);

    // Convert and Write CSV header
    std::string utf8Header = Encoding::wstringToUtf8(L"Selected,Find,Replace,WholeWord,MatchCase,UseVariables,Extended,Regex,Comments\n");
    outFile << utf8Header;

    // Write list items to CSV file
    for (const ReplaceItemData& item : list) {
        std::wstring line = std::to_wstring(item.isEnabled) + L"," +
            escapeCsvValue(item.findText) + L"," +
            escapeCsvValue(item.replaceText) + L"," +
            std::to_wstring(item.wholeWord) + L"," +
            std::to_wstring(item.matchCase) + L"," +
            std::to_wstring(item.useVariables) + L"," +
            std::to_wstring(item.extended) + L"," +
            std::to_wstring(item.regex) + L"," +
            escapeCsvValue(item.comments) + L"\n";
        std::string utf8Line = Encoding::wstringToUtf8(line);
        outFile << utf8Line;
    }

    outFile.close();

    return !outFile.fail();;
}

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (!saveListToCsvSilent(filePath, list)) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    showStatusMessage(LM.get(L"status_saved_items_to_csv", { std::to_wstring(list.size()) }), MessageStatus::Success);

    // Update the file path and original hash after a successful save
    listFilePath = filePath;
    originalListHash = computeListHash(list);

    // Update the displayed file path below the list
    showListFilePath();
}

int MultiReplace::checkForUnsavedChanges() {
    std::size_t currentListHash = computeListHash(replaceListData);

    if (currentListHash != originalListHash) {

        std::wstring message;
        if (!listFilePath.empty()) {
            // Get the shortened file path and build the message
            std::wstring shortenedFilePath = getShortenedFilePath(listFilePath, 500);
            message = LM.get(L"msgbox_unsaved_changes_file", { shortenedFilePath });
        }
        else {
            // If no file is associated, use the alternative message
            message = LM.get(L"msgbox_unsaved_changes");
        }

        // Show the MessageBox with the appropriate message
        int result = MessageBox(
            nppData._nppHandle,
            message.c_str(),
            LM.get(L"msgbox_title_save_list").c_str(),
            MB_ICONWARNING | MB_YESNOCANCEL
        );

        if (result == IDYES) {
            if (!listFilePath.empty()) {
                saveListToCsv(listFilePath, replaceListData);
                return IDYES;  // Proceed if saved successfully
            }
            else {
                std::wstring filePath = promptSaveListToCsv();

                if (!filePath.empty()) {
                    saveListToCsv(filePath, replaceListData);
                    return IDYES;  // Proceed with clear after save
                }
                else {
                    return IDCANCEL;  // Cancel if no file was selected
                }
            }
        }
        else if (result == IDNO) {
            return IDNO;  // Allow proceeding without saving
        }
        else if (result == IDCANCEL) {
            return IDCANCEL;  // Cancel the action
        }
    }

    return IDYES;  // No unsaved changes, allow proceeding
}

void MultiReplace::loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list) {
    // Open the CSV file
    std::ifstream inFile(std::filesystem::path(filePath), std::ios::binary);
    if (!inFile.is_open()) {
        std::wstring shortenedFilePathW = getShortenedFilePath(filePath, 500);
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_unable_to_open_file", { shortenedFilePathW })));
    }

    // Read raw bytes
    std::string raw((std::istreambuf_iterator<char>(inFile)), std::istreambuf_iterator<char>());
    inFile.close();

    // [BOM REMOVAL] strip BOM bytes if present
    size_t offset = 0;
    UINT cp = CP_UTF8;                                    // assume UTF-8
    if (raw.size() >= 3
        && static_cast<unsigned char>(raw[0]) == 0xEF
        && static_cast<unsigned char>(raw[1]) == 0xBB
        && static_cast<unsigned char>(raw[2]) == 0xBF)
    {
        offset = 3;                                       // skip BOM
    }
    else if (!Encoding::isValidUtf8(raw)) {
        cp = CP_ACP;                                      // fallback to ANSI
    }

    // Convert UTF-8 or ANSI bytes to std::wstring
    int wlen = MultiByteToWideChar(cp, 0, raw.c_str() + offset, (int)raw.size() - (int)offset, nullptr, 0);
    std::wstring content(wlen, L'\0');
    MultiByteToWideChar(cp, 0, raw.c_str() + offset, (int)raw.size() - (int)offset, &content[0], wlen);

    std::wstringstream contentStream(content);

    // Read the header line
    std::wstring headerLine;
    if (!std::getline(contentStream, headerLine)) {
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_column_count")));
    }

    // Temporary list to hold parsed items
    std::vector<ReplaceItemData> tempList;
    std::wstring line;

    // Read and process each line in the file
    while (std::getline(contentStream, line)) {
        std::vector<std::wstring> columns = parseCsvLine(line);

        // Check if column count is valid
        if (columns.size() < 8 || columns.size() > 9) {
            throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_column_count")));
        }

        try {
            ReplaceItemData item;
            item.isEnabled = std::stoi(columns[0]) != 0;
            item.findText = columns[1];
            item.replaceText = columns[2];
            item.wholeWord = std::stoi(columns[3]) != 0;
            item.matchCase = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended = std::stoi(columns[6]) != 0;
            item.regex = std::stoi(columns[7]) != 0;
            item.comments = (columns.size() == 9) ? columns[8] : L"";
            tempList.push_back(item);
        }
        catch (const std::exception&) {
            throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_data_in_columns")));
        }
    }

    // Check if the file contains valid data rows
    if (tempList.empty()) {
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_no_valid_items_in_csv")));
    }

    // Transfer parsed data to the target list
    list = std::move(tempList);
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {
    try {
        loadListFromCsvSilent(filePath, replaceListData);

        // Update the file path and display it
        listFilePath = filePath;
        showListFilePath();

        // Compute the hash of the loaded list
        originalListHash = computeListHash(replaceListData);

        // Clear Undo and Redo stacks
        URM.clear();

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Display success or error message based on the loaded data
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_no_valid_items_in_csv"), MessageStatus::Error);
        }
        else {
            showStatusMessage(LM.get(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), MessageStatus::Success);
        }
    }
    catch (const CsvLoadException& ex) {
        showStatusMessage(Encoding::utf8ToWString(ex.what()), MessageStatus::Error);
    }
}

void MultiReplace::checkForFileChangesAtStartup() {
    if (listFilePath.empty()) {
        return;
    }

    try {
        std::vector<ReplaceItemData> tempListFromFile;
        loadListFromCsvSilent(listFilePath, tempListFromFile);

        std::size_t newFileHash = computeListHash(tempListFromFile);
        if (newFileHash != originalListHash) {
            std::wstring shortenedFilePath = getShortenedFilePath(listFilePath, 500);
            std::wstring message = LM.get(L"msgbox_file_modified_prompt", { shortenedFilePath });

            int response = MessageBox(nppData._nppHandle, message.c_str(), LM.get(L"msgbox_title_reload").c_str(), MB_YESNO | MB_ICONWARNING | MB_SETFOREGROUND);
            if (response == IDYES) {
                replaceListData = tempListFromFile;
                originalListHash = newFileHash;
            }
        }
    }
    catch (const CsvLoadException& ex) {
        showStatusMessage(Encoding::utf8ToWString(ex.what()), MessageStatus::Error);
    }

    if (replaceListData.empty()) {
        showStatusMessage(LM.get(L"status_no_valid_items_in_csv"), MessageStatus::Error);
    }
    else {
        showStatusMessage(LM.get(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), MessageStatus::Success);
        ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);
    }
}

std::wstring MultiReplace::escapeCsvValue(const std::wstring& value) {
    std::wstring escapedValue = L"\"";

    for (const wchar_t& ch : value) {
        switch (ch) {
        case L'"':
            escapedValue += L"\"\"";
            break;
        case L'\n':
            escapedValue += L"\\n";
            break;
        case L'\r':
            escapedValue += L"\\r";
            break;
        case L'\\':
            escapedValue += L"\\\\";
            break;
        default:
            escapedValue += ch;
            break;
        }
    }

    escapedValue += L"\""; 

    return escapedValue;
}

std::wstring MultiReplace::unescapeCsvValue(const std::wstring& value) {
    std::wstring unescapedValue;
    if (value.empty()) return unescapedValue;

    // Detect outer quotes: only if both first and last are quotes
    const bool hadOuterQuotes = (value.front() == L'"' && value.back() == L'"');
    size_t start = hadOuterQuotes ? 1 : 0;
    size_t end = hadOuterQuotes ? value.size() - 1 : value.size();

    for (size_t i = start; i < end; ++i) {
        if (i < end - 1 && value[i] == L'\\') {
            // Handle backslash escapes
            switch (value[i + 1]) {
            case L'n': unescapedValue += L'\n'; ++i; break;
            case L'r': unescapedValue += L'\r'; ++i; break;
            case L'\\': unescapedValue += L'\\'; ++i; break;
            default: unescapedValue += value[i]; break;
            }
        }
        // IMPORTANT: Only collapse CSV double quotes ("") to (")
        // if the original field had outer quotes (i.e., it was a quoted CSV field).
        else if (hadOuterQuotes && i < end - 1 && value[i] == L'"' && value[i + 1] == L'"') {
            unescapedValue += L'"';
            ++i;
        }
        else {
            unescapedValue += value[i];
        }
    }
    return unescapedValue;
}

std::vector<std::wstring> MultiReplace::parseCsvLine(const std::wstring& line) {
    std::vector<std::wstring> columns;
    std::wstring currentValue;
    bool insideQuotes = false;

    for (size_t i = 0; i < line.length(); ++i) {
        const wchar_t& ch = line[i];
        if (ch == L'"') {
            if (insideQuotes && i + 1 < line.length() && line[i + 1] == L'"') {
                // Escaped quote
                currentValue += L'"';
                ++i; // Skip the next quote
            }
            else {
                // Toggle the state
                insideQuotes = !insideQuotes;
            }
        }
        else if (ch == L',' && !insideQuotes) {
            // When not inside quotes, treat comma as column separator
            columns.push_back(unescapeCsvValue(currentValue));
            currentValue.clear();
        }
        else {
            currentValue += ch; // Append the character to the current value
        }
    }
    columns.push_back(unescapeCsvValue(currentValue)); // Add the last value
    return columns;
}

#pragma endregion


#pragma region Export

void MultiReplace::exportToBashScript(const std::wstring& fileName) {
    std::ofstream file(fileName);
    if (!file.is_open()) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    // Create date
    time_t currentTime = time(nullptr);
    struct tm* localTime = localtime(&currentTime);
    char dateBuffer[80];
    strftime(dateBuffer, sizeof(dateBuffer), "%Y-%m-%d", localTime);

    file << "#!/bin/bash\n";
    file << "# Auto-generated by MultiReplace Notepad++\n";
    file << "# Created on: " << dateBuffer << "\n\n";
    file << "inputFile=\"$1\"\n";
    file << "outputFile=\"$2\"\n\n";

    file << "processLine() {\n";
    file << "    local findString=\"$1\"\n";
    file << "    local replaceString=\"$2\"\n";
    file << "    local wholeWord=\"$3\"\n";
    file << "    local matchCase=\"$4\"\n";
    file << "    local normal=\"$5\"\n";
    file << "    local extended=\"$6\"\n";
    file << "    local regex=\"$7\"\n";
    file << R"(
    if [[ "$wholeWord" -eq 1 ]]; then
        findString='\b'${findString}'\b'
    fi
    if [[ "$matchCase" -eq 1 ]]; then
        template='s|'${findString}'|'${replaceString}'|g'
    else
        template='s|'${findString}'|'${replaceString}'|gi'
    fi
    case 1 in
        $normal)
            sed -i "${template}" "$outputFile"
            ;;
        $extended)
            sed -i -e ':a' -e 'N' -e '$!ba' -e 's/\n/__NEWLINE__/g' -e 's/\r/__CARRIAGERETURN__/g' "$outputFile"
            sed -i "${template}" "$outputFile"
            sed -i 's/__NEWLINE__/\n/g; s/__CARRIAGERETURN__/\r/g' "$outputFile"
            ;;
        $regex)
            sed -i -r "${template}" "$outputFile"
            ;;
    esac
)";
    file << "}\n\n";

    file << "cp $inputFile $outputFile\n\n";

    file << "# processLine arguments: \"findString\" \"replaceString\" wholeWord matchCase normal extended regex\n";

    bool hasExcludedItems = false;
    for (const auto& itemData : replaceListData) {
        if (!itemData.isEnabled) continue; // Skip if this item is not selected
        if (itemData.useVariables) {
            hasExcludedItems = true; // Mark as excluded
            continue;
        }

        std::string find;
        std::string replace;
        if (itemData.extended) {
            find = replaceNewline(translateEscapes(escapeSpecialChars(Encoding::wstringToUtf8(itemData.findText), true)), ReplaceMode::Extended);
            replace = replaceNewline(translateEscapes(escapeSpecialChars(Encoding::wstringToUtf8(itemData.replaceText), true)), ReplaceMode::Extended);
        }
        else if (itemData.regex) {
            find = replaceNewline(Encoding::wstringToUtf8(itemData.findText), ReplaceMode::Regex);
            replace = replaceNewline(Encoding::wstringToUtf8(itemData.replaceText), ReplaceMode::Regex);
        }
        else {
            find = replaceNewline(escapeSpecialChars(Encoding::wstringToUtf8(itemData.findText), false), ReplaceMode::Normal);
            replace = replaceNewline(escapeSpecialChars(Encoding::wstringToUtf8(itemData.replaceText), false), ReplaceMode::Normal);
        }

        std::string wholeWord = itemData.wholeWord ? "1" : "0";
        std::string matchCase = itemData.matchCase ? "1" : "0";
        std::string normal = (!itemData.regex && !itemData.extended) ? "1" : "0";
        std::string extended = itemData.extended ? "1" : "0";
        std::string regex = itemData.regex ? "1" : "0";

        file << "processLine \"" << find << "\" \"" << replace << "\" " << wholeWord << " " << matchCase << " " << normal << " " << extended << " " << regex << "\n";
    }

    file.close();

    if (file.fail()) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    showStatusMessage(LM.get(L"status_list_exported_to_bash"), MessageStatus::Success);

    // Show message box if excluded items were found
    if (hasExcludedItems) {
        MessageBox(_hSelf,
            LM.get(L"msgbox_use_variables_not_exported").c_str(),
            LM.get(L"msgbox_title_warning").c_str(),
            MB_OK | MB_ICONWARNING);
    }

}

std::string MultiReplace::escapeSpecialChars(const std::string& input, bool extended) {
    std::string output = input;

    // Define the escape characters that should not be masked
    std::string supportedEscapes = "nrt0xubd";

    // Mask all characters that could be interpreted by sed or the shell, including escape sequences.
    std::string specialChars = "$.*[]^&\\{}()?+|<>\"'`~;#";

    for (char c : specialChars) {
        std::string str(1, c);
        size_t pos = output.find(str);

        while (pos != std::string::npos) {
            // Check if the current character is an escape character
            if (str == "\\" && (pos == 0 || output[pos - 1] != '\\')) {
                // Skip masking if it is a supported or unsupported escape
                if (extended && (pos + 1 < output.size() && supportedEscapes.find(output[pos + 1]) != std::string::npos)) {
                    pos = output.find(str, pos + 1);
                    continue;
                }
            }

            output.insert(pos, "\\");
            pos = output.find(str, pos + 2);

        }
    }

    return output;
}

void MultiReplace::handleEscapeSequence(const std::regex& regex, const std::string& input, std::string& output, std::function<char(const std::string&)> converter) {
    std::sregex_iterator begin = std::sregex_iterator(input.begin(), input.end(), regex);
    std::sregex_iterator end;
    for (std::sregex_iterator i = begin; i != end; ++i) {
        std::smatch match = *i;
        std::string escape = match.str();
        try {
            char actualChar = converter(escape);
            size_t pos = output.find(escape);
            if (pos != std::string::npos) {
                output.replace(pos, escape.length(), 1, actualChar);
            }
        }
        catch (...) {
            // no errors caught during tests yet
        }

    }
}

std::string MultiReplace::translateEscapes(const std::string& input) {
    std::string output = input;

    std::regex octalRegex("\\\\o([0-7]{3})");
    std::regex decimalRegex("\\\\d([0-9]{3})");
    std::regex hexRegex("\\\\x([0-9a-fA-F]{2})");
    std::regex binaryRegex("\\\\b([01]{8})");
    std::regex unicodeRegex("\\\\u([0-9a-fA-F]{4})");
    std::regex newlineRegex("\\\\n");
    std::regex carriageReturnRegex("\\\\r");
    std::regex nullCharRegex("\\\\0");

    handleEscapeSequence(octalRegex, input, output, [](const std::string& octalEscape) {
        return static_cast<char>(std::stoi(octalEscape.substr(2), nullptr, 8));
        });

    handleEscapeSequence(decimalRegex, input, output, [](const std::string& decimalEscape) {
        return static_cast<char>(std::stoi(decimalEscape.substr(2)));
        });

    handleEscapeSequence(hexRegex, input, output, [](const std::string& hexEscape) {
        return static_cast<char>(std::stoi(hexEscape.substr(2), nullptr, 16));
        });

    handleEscapeSequence(binaryRegex, input, output, [](const std::string& binaryEscape) {
        return static_cast<char>(std::bitset<8>(binaryEscape.substr(2)).to_ulong());
        });

    handleEscapeSequence(unicodeRegex, input, output, [this](const std::string& unicodeEscape) -> char {
        int codepoint = std::stoi(unicodeEscape.substr(2), nullptr, 16);
        wchar_t unicodeChar = static_cast<wchar_t>(codepoint);
        std::wstring unicodeString = { unicodeChar };
        std::string result = Encoding::wstringToUtf8(unicodeString);
        return result.empty() ? 0 : result.front();
        });


    output = std::regex_replace(output, newlineRegex, "__NEWLINE__");
    output = std::regex_replace(output, carriageReturnRegex, "__CARRIAGERETURN__");
    output = std::regex_replace(output, nullCharRegex, "");  // \0 will not be supported

    return output;
}

std::string MultiReplace::replaceNewline(const std::string& input, ReplaceMode mode) {
    std::string result = input;

    if (mode == ReplaceMode::Normal) {
        result.erase(std::remove(result.begin(), result.end(), '\n'), result.end());
        result.erase(std::remove(result.begin(), result.end(), '\r'), result.end());
    }
    else if (mode == ReplaceMode::Extended) {
        result = std::regex_replace(result, std::regex("\n"), "__NEWLINE__");
        result = std::regex_replace(result, std::regex("\r"), "__CARRIAGERETURN__");
    }
    else if (mode == ReplaceMode::Regex) {
        result = std::regex_replace(result, std::regex("\n"), "\\n");
        result = std::regex_replace(result, std::regex("\r"), "\\r");
    }

    return result;
}

#pragma endregion


#pragma region INI

std::pair<std::wstring, std::wstring> MultiReplace::generateConfigFilePaths() {
    wchar_t configDir[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM)configDir);
    configDir[MAX_PATH - 1] = '\0'; // Ensure null-termination

    std::wstring iniFilePath = std::wstring(configDir) + L"\\MultiReplace.ini";
    std::wstring csvFilePath = std::wstring(configDir) + L"\\MultiReplaceList.ini";
    return { iniFilePath, csvFilePath };
}

void MultiReplace::saveSettingsToIni(const std::wstring& iniFilePath) {
    std::ofstream outFile(iniFilePath, std::ios::binary);

    if (!outFile.is_open()) {
        throw std::runtime_error("Could not open settings file for writing.");
    }

    // [BOM] Write UTF-8 BOM so that editors/parsers recognize UTF-8
    outFile.write("\xEF\xBB\xBF", 3);

    // Get the current window rectangle
    RECT currentRect;
    GetWindowRect(_hSelf, &currentRect);
    int width = currentRect.right - currentRect.left;
    int height = currentRect.bottom - currentRect.top;
    int posX = currentRect.left;
    int posY = currentRect.top;


    // Update useListOnHeight if Use List is checked
    if (useListEnabled) {
        useListOnHeight = height;
    }

    outFile << Encoding::wstringToUtf8(L"[Window]\n");
    outFile << Encoding::wstringToUtf8(L"PosX=" + std::to_wstring(posX) + L"\n");
    outFile << Encoding::wstringToUtf8(L"PosY=" + std::to_wstring(posY) + L"\n");
    outFile << Encoding::wstringToUtf8(L"Width=" + std::to_wstring(width) + L"\n");
    outFile << Encoding::wstringToUtf8(L"Height=" + std::to_wstring(useListOnHeight) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ScaleFactor=" + std::to_wstring(dpiMgr->getCustomScaleFactor()).substr(0, std::to_wstring(dpiMgr->getCustomScaleFactor()).find(L'.') + 2) + L"\n");

    // Save transparency settings
    outFile << Encoding::wstringToUtf8(L"ForegroundTransparency=" + std::to_wstring(foregroundTransparency) + L"\n");
    outFile << Encoding::wstringToUtf8(L"BackgroundTransparency=" + std::to_wstring(backgroundTransparency) + L"\n");

    // Store column widths for "Find Count", "Replace Count", and "Comments"
    findCountColumnWidth = (columnIndices[ColumnID::FIND_COUNT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_COUNT]) : findCountColumnWidth;
    replaceCountColumnWidth = (columnIndices[ColumnID::REPLACE_COUNT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_COUNT]) : replaceCountColumnWidth;
    findColumnWidth = (columnIndices[ColumnID::FIND_TEXT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_TEXT]) : findColumnWidth;
    replaceColumnWidth = (columnIndices[ColumnID::REPLACE_TEXT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_TEXT]) : replaceColumnWidth;
    commentsColumnWidth = (columnIndices[ColumnID::COMMENTS] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::COMMENTS]) : commentsColumnWidth;

    outFile << Encoding::wstringToUtf8(L"[ListColumns]\n");
    outFile << Encoding::wstringToUtf8(L"FindCountWidth=" + std::to_wstring(findCountColumnWidth) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceCountWidth=" + std::to_wstring(replaceCountColumnWidth) + L"\n");
    outFile << Encoding::wstringToUtf8(L"FindWidth=" + std::to_wstring(findColumnWidth) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceWidth=" + std::to_wstring(replaceColumnWidth) + L"\n");
    outFile << Encoding::wstringToUtf8(L"CommentsWidth=" + std::to_wstring(commentsColumnWidth) + L"\n");

    // Save column visibility states
    outFile << Encoding::wstringToUtf8(L"FindCountVisible=" + std::to_wstring(isFindCountVisible) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceCountVisible=" + std::to_wstring(isReplaceCountVisible) + L"\n");
    outFile << Encoding::wstringToUtf8(L"CommentsVisible=" + std::to_wstring(isCommentsColumnVisible) + L"\n");
    outFile << Encoding::wstringToUtf8(L"DeleteButtonVisible=" + std::to_wstring(isDeleteButtonVisible ? 1 : 0) + L"\n");

    // Save column lock states
    outFile << Encoding::wstringToUtf8(L"FindColumnLocked=" + std::to_wstring(findColumnLockedEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceColumnLocked=" + std::to_wstring(replaceColumnLockedEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"CommentsColumnLocked=" + std::to_wstring(commentsColumnLockedEnabled ? 1 : 0) + L"\n");

    // Convert and Store the current "Find what" and "Replace with" texts
    std::wstring currentFindTextData = escapeCsvValue(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT));
    std::wstring currentReplaceTextData = escapeCsvValue(getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT));

    outFile << Encoding::wstringToUtf8(L"[Current]\n");
    outFile << Encoding::wstringToUtf8(L"FindText=" + currentFindTextData + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceText=" + currentReplaceTextData + L"\n");

    // Prepare and Store the current options
    int wholeWord = IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int matchCase = IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int extended = IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED ? 1 : 0;
    int regex = IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? 1 : 0;
    int replaceAtMatches = IsDlgButtonChecked(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX) == BST_CHECKED ? 1 : 0;
    std::wstring editAtMatchesText = L"\"" + getTextFromDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT) + L"\"";
    int wrapAround = IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int useVariables = IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int ButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED ? 1 : 0;
    int useList = useListEnabled ? 1 : 0;

    outFile << Encoding::wstringToUtf8(L"[Options]\n");
    outFile << Encoding::wstringToUtf8(L"WholeWord=" + std::to_wstring(wholeWord) + L"\n");
    outFile << Encoding::wstringToUtf8(L"MatchCase=" + std::to_wstring(matchCase) + L"\n");
    outFile << Encoding::wstringToUtf8(L"Extended=" + std::to_wstring(extended) + L"\n");
    outFile << Encoding::wstringToUtf8(L"Regex=" + std::to_wstring(regex) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ReplaceAtMatches=" + std::to_wstring(replaceAtMatches) + L"\n");
    outFile << Encoding::wstringToUtf8(L"EditAtMatches=" + editAtMatchesText + L"\n");
    outFile << Encoding::wstringToUtf8(L"WrapAround=" + std::to_wstring(wrapAround) + L"\n");
    outFile << Encoding::wstringToUtf8(L"UseVariables=" + std::to_wstring(useVariables) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ButtonsMode=" + std::to_wstring(ButtonsMode) + L"\n");
    outFile << Encoding::wstringToUtf8(L"UseList=" + std::to_wstring(useList) + L"\n");
    outFile << Encoding::wstringToUtf8(L"DockWrap=" + std::to_wstring(ResultDock::wrapEnabled()) + L"\n");
    outFile << Encoding::wstringToUtf8(L"DockPurge=" + std::to_wstring(ResultDock::purgeEnabled()) + L"\n");
    outFile << Encoding::wstringToUtf8(L"HighlightMatch=" + std::to_wstring(highlightMatchEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ExportToBash=" + std::to_wstring(exportToBashEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"Tooltips=" + std::to_wstring(tooltipsEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"AlertNotFound=" + std::to_wstring(alertNotFoundEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"DoubleClickEdits=" + std::to_wstring(doubleClickEditsEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"HoverText=" + std::to_wstring(isHoverTextEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"EditFieldSize=" + std::to_wstring(editFieldSize) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ListStatistics=" + std::to_wstring(listStatisticsEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"StayAfterReplace=" + std::to_wstring(stayAfterReplaceEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"AllFromCursor=" + std::to_wstring(allFromCursorEnabled ? 1 : 0) + L"\n");
    outFile << Encoding::wstringToUtf8(L"GroupResults=" + std::to_wstring(groupResultsEnabled) + L"\n");

    // Lua runtime options
    outFile << Encoding::wstringToUtf8(L"[Lua]\n");
    outFile << Encoding::wstringToUtf8(L"SafeMode=" + std::to_wstring(luaSafeModeEnabled ? 1 : 0) + L"\n");

    // Convert and Store the scope options
    int selection = IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED ? 1 : 0;
    int columnMode = IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED ? 1 : 0;
    std::wstring columnNum = L"\"" + getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT) + L"\"";
    std::wstring delimiter = L"\"" + getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT) + L"\"";
    std::wstring quoteChar = L"\"" + getTextFromDialogItem(_hSelf, IDC_QUOTECHAR_EDIT) + L"\"";
    std::wstring headerLines = std::to_wstring(CSVheaderLinesCount);

    outFile << Encoding::wstringToUtf8(L"[Scope]\n");
    outFile << Encoding::wstringToUtf8(L"Selection=" + std::to_wstring(selection) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ColumnMode=" + std::to_wstring(columnMode) + L"\n");
    outFile << Encoding::wstringToUtf8(L"ColumnNum=" + columnNum + L"\n");
    outFile << Encoding::wstringToUtf8(L"Delimiter=" + delimiter + L"\n");
    outFile << Encoding::wstringToUtf8(L"QuoteChar=" + quoteChar + L"\n");
    outFile << Encoding::wstringToUtf8(L"HeaderLines=" + headerLines + L"\n");

    // Save “Replace in Files” settings
    std::wstring filterText = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
    std::wstring dirText = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);
    int inSub = IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int inHidden = IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED ? 1 : 0;
    outFile << Encoding::wstringToUtf8(L"[ReplaceInFiles]\n");
    outFile << Encoding::wstringToUtf8(L"Filter=\"" + filterText + L"\"\n");
    outFile << Encoding::wstringToUtf8(L"Directory=\"" + dirText + L"\"\n");
    outFile << Encoding::wstringToUtf8(L"InSubfolders=" + std::to_wstring(inSub) + L"\n");
    outFile << Encoding::wstringToUtf8(L"InHiddenFolders=" + std::to_wstring(inHidden) + L"\n");

    // Save the list file path and original hash
    outFile << Encoding::wstringToUtf8(L"[File]\n");
    outFile << Encoding::wstringToUtf8(L"ListFilePath=" + escapeCsvValue(listFilePath) + L"\n");
    outFile << Encoding::wstringToUtf8(L"OriginalListHash=" + std::to_wstring(originalListHash) + L"\n");

    // Save the "Find what" history
    LRESULT findWhatCount = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETCOUNT, 0, 0);
    int itemsToSave = std::min(static_cast<int>(findWhatCount), maxHistoryItems);
    outFile << Encoding::wstringToUtf8(L"[History]\n");
    outFile << Encoding::wstringToUtf8(L"FindTextHistoryCount=" + std::to_wstring(itemsToSave) + L"\n");

    // Save only the newest maxHistoryItems entries (starting from index 0)
    for (LRESULT i = 0; i < itemsToSave; ++i) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring findTextData = escapeCsvValue(std::wstring(buffer.data()));
        outFile << Encoding::wstringToUtf8(L"FindTextHistory" + std::to_wstring(i) + L"=" + findTextData + L"\n");
    }

    // Save the "Replace with" history
    LRESULT replaceWithCount = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETCOUNT, 0, 0);
    int replaceItemsToSave = std::min(static_cast<int>(replaceWithCount), maxHistoryItems);
    outFile << Encoding::wstringToUtf8(L"ReplaceTextHistoryCount=" + std::to_wstring(replaceItemsToSave) + L"\n");

    // Save only the newest maxHistoryItems entries (starting from index 0)
    for (LRESULT i = 0; i < replaceItemsToSave; ++i) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring replaceTextData = escapeCsvValue(std::wstring(buffer.data()));
        outFile << Encoding::wstringToUtf8(L"ReplaceTextHistory" + std::to_wstring(i) + L"=" + replaceTextData + L"\n");
    }

    // Save Filter history
    LRESULT filterCount = SendMessage(GetDlgItem(_hSelf, IDC_FILTER_EDIT), CB_GETCOUNT, 0, 0);
    int filterItemsToSave = std::min(static_cast<int>(filterCount), maxHistoryItems);
    outFile << Encoding::wstringToUtf8(L"FilterHistoryCount=" + std::to_wstring(filterItemsToSave) + L"\n");
    for (int i = 0; i < filterItemsToSave; ++i) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_FILTER_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buf(len + 1);
        SendMessage(GetDlgItem(_hSelf, IDC_FILTER_EDIT), CB_GETLBTEXT, i, (LPARAM)buf.data());
        outFile << Encoding::wstringToUtf8(L"FilterHistory" + std::to_wstring(i) + L"=" + escapeCsvValue(buf.data()) + L"\n");
    }

    //  Save Dir history
    LRESULT dirCount = SendMessage(GetDlgItem(_hSelf, IDC_DIR_EDIT), CB_GETCOUNT, 0, 0);
    int dirItemsToSave = std::min(static_cast<int>(dirCount), maxHistoryItems);
    outFile << Encoding::wstringToUtf8(L"DirHistoryCount=" + std::to_wstring(dirItemsToSave) + L"\n");
    for (int i = 0; i < dirItemsToSave; ++i) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_DIR_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buf(len + 1);
        SendMessage(GetDlgItem(_hSelf, IDC_DIR_EDIT), CB_GETLBTEXT, i, (LPARAM)buf.data());
        outFile << Encoding::wstringToUtf8(L"DirHistory" + std::to_wstring(i) + L"=" + escapeCsvValue(buf.data()) + L"\n");
    }

    outFile.close();
}

void MultiReplace::saveSettings() {
    static bool settingsSaved = false;
    if (settingsSaved) {
        return;  // Check as WM_DESTROY will be 28 times triggered
    }

    // Generate the paths to the configuration files
    auto [iniFilePath, csvFilePath] = generateConfigFilePaths();

    // Try to save the settings in the INI file
    try {
        saveSettingsToIni(iniFilePath);
        saveListToCsvSilent(csvFilePath, replaceListData);
    }
    catch (const std::exception& ex) {
        // If an error occurs while writing to the INI file, we show an error message
        std::wstring errorMessage = LM.get(L"msgbox_error_saving_settings", { std::wstring(ex.what(), ex.what() + strlen(ex.what())) });
        MessageBox(nppData._nppHandle, errorMessage.c_str(), LM.get(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
    }
    settingsSaved = true;
}

void MultiReplace::loadSettingsFromIni() {

    // Loading the history for the Find text field in reverse order
    int findHistoryCount = CFG.readInt(L"History", L"FindTextHistoryCount", 0);
    for (int i = findHistoryCount - 1; i >= 0; --i) {
        std::wstring findHistoryItem = CFG.readString(L"History", L"FindTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findHistoryItem);
    }

    // Loading the history for the Replace text field in reverse order
    int replaceHistoryCount = CFG.readInt(L"History", L"ReplaceTextHistoryCount", 0);
    for (int i = replaceHistoryCount - 1; i >= 0; --i) {
        std::wstring replaceHistoryItem = CFG.readString(L"History", L"ReplaceTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceHistoryItem);
    }

    // Load Filter history
    int filterHistoryCount = CFG.readInt(L"History", L"FilterHistoryCount", 0);
    for (int i = filterHistoryCount - 1; i >= 0; --i) {
        std::wstring item = CFG.readString(L"History", L"FilterHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FILTER_EDIT), item);
    }

    // Load Dir history
    int dirHistoryCount = CFG.readInt(L"History", L"DirHistoryCount", 0);
    for (int i = dirHistoryCount - 1; i >= 0; --i) {
        std::wstring item = CFG.readString(L"History", L"DirHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), item);
    }

    // Reading and setting current "Find what" and "Replace with" texts
    std::wstring findText = CFG.readString(L"Current", L"FindText", L"");
    std::wstring replaceText = CFG.readString(L"Current", L"ReplaceText", L"");
    setTextInDialogItem(_hSelf, IDC_FIND_EDIT, findText);
    setTextInDialogItem(_hSelf, IDC_REPLACE_EDIT, replaceText);

    // Setting options based on the INI file
    bool wholeWord = CFG.readBool(L"Options", L"WholeWord", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);

    bool matchCase = CFG.readBool(L"Options", L"MatchCase", false);
    SendMessage(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, matchCase ? BST_CHECKED : BST_UNCHECKED, 0);

    bool useVariables = CFG.readBool(L"Options", L"UseVariables", false);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, useVariables ? BST_CHECKED : BST_UNCHECKED, 0);

    // Selecting the appropriate search mode radio button based on the settings
    bool extended = CFG.readBool(L"Options", L"Extended", false);
    bool regex = CFG.readBool(L"Options", L"Regex", false);
    if (regex) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_REGEX_RADIO);
    }
    else if (extended) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_EXTENDED_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);
    }

    // Setting additional options
    bool wrapAround = CFG.readBool(L"Options", L"WrapAround", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WRAP_AROUND_CHECKBOX), BM_SETCHECK, wrapAround ? BST_CHECKED : BST_UNCHECKED, 0);

    bool replaceAtMatches = CFG.readBool(L"Options", L"ReplaceAtMatches", false);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX), BM_SETCHECK, replaceAtMatches ? BST_CHECKED : BST_UNCHECKED, 0);

    std::wstring editAtMatchesText = CFG.readString(L"Options", L"EditAtMatches", L"1");
    setTextInDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT, editAtMatchesText);

    bool replaceButtonsMode = CFG.readBool(L"Options", L"ButtonsMode", false);
    SendMessage(GetDlgItem(_hSelf, IDC_2_BUTTONS_MODE), BM_SETCHECK, replaceButtonsMode ? BST_CHECKED : BST_UNCHECKED, 0);

    useListEnabled = CFG.readBool(L"Options", L"UseList", true);
    updateUseListState(false);

    ResultDock::setWrapEnabled(CFG.readBool(L"Options", L"DockWrap", false));
    ResultDock::setPurgeEnabled(CFG.readBool(L"Options", L"DockPurge", false));

    highlightMatchEnabled = CFG.readBool(L"Options", L"HighlightMatch", true);
    exportToBashEnabled = CFG.readBool(L"Options", L"ExportToBash", false);
    alertNotFoundEnabled = CFG.readBool(L"Options", L"AlertNotFound", true);
    doubleClickEditsEnabled = CFG.readBool(L"Options", L"DoubleClickEdits", true);
    isHoverTextEnabled = CFG.readBool(L"Options", L"HoverText", true);

    // CFG.readInt for editFieldSize
    editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 5);
    editFieldSize = std::clamp(editFieldSize, MIN_EDIT_FIELD_SIZE, MAX_EDIT_FIELD_SIZE);

    listStatisticsEnabled = CFG.readBool(L"Options", L"ListStatistics", false);
    stayAfterReplaceEnabled = CFG.readBool(L"Options", L"StayAfterReplace", false);
    allFromCursorEnabled = CFG.readBool(L"Options", L"AllFromCursor", false);
    groupResultsEnabled = CFG.readBool(L"Options", L"GroupResults", false);

    // Lua runtime options
    luaSafeModeEnabled = CFG.readBool(L"Lua", L"SafeMode", false);

    // Loading and setting the scope
    int selection = CFG.readInt(L"Scope", L"Selection", 0);
    int columnMode = CFG.readInt(L"Scope", L"ColumnMode", 0);

    // Reading and setting specific scope settings
    std::wstring columnNum = CFG.readString(L"Scope", L"ColumnNum", L"1-50");
    setTextInDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT, columnNum);

    std::wstring delimiter = CFG.readString(L"Scope", L"Delimiter", L",");
    setTextInDialogItem(_hSelf, IDC_DELIMITER_EDIT, delimiter);

    std::wstring quoteChar = CFG.readString(L"Scope", L"QuoteChar", L"\"");
    setTextInDialogItem(_hSelf, IDC_QUOTECHAR_EDIT, quoteChar);

    CSVheaderLinesCount = CFG.readInt(L"Scope", L"HeaderLines", 1);

    // --- Load “Replace in Files” settings --------------------------
    std::wstring filter = CFG.readString(L"ReplaceInFiles", L"Filter", L"");
    setTextInDialogItem(_hSelf, IDC_FILTER_EDIT, filter);

    std::wstring dir = CFG.readString(L"ReplaceInFiles", L"Directory", L"");
    setTextInDialogItem(_hSelf, IDC_DIR_EDIT, dir);

    bool inSub = CFG.readBool(L"ReplaceInFiles", L"InSubfolders", false);
    SendMessage(GetDlgItem(_hSelf, IDC_SUBFOLDERS_CHECKBOX), BM_SETCHECK, inSub ? BST_CHECKED : BST_UNCHECKED, 0);

    bool inHidden = CFG.readBool(L"ReplaceInFiles", L"InHiddenFolders", false);
    SendMessage(GetDlgItem(_hSelf, IDC_HIDDENFILES_CHECKBOX), BM_SETCHECK, inHidden ? BST_CHECKED : BST_UNCHECKED, 0);


    // Load file path and original hash
    listFilePath = CFG.readString(L"File", L"ListFilePath", L"");
    originalListHash = CFG.readSizeT(L"File", L"OriginalListHash", 0);

    if (selection) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_SELECTION_RADIO);
        onSelectionChanged();
    }
    else if (columnMode) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_COLUMN_MODE_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_ALL_TEXT_RADIO);
    }

    setUIElementVisibility();
}

void MultiReplace::loadSettings() {
    auto [_, csvFilePath] = generateConfigFilePaths();

    // Read all settings from the cache
    loadSettingsFromIni();

    try {
        loadListFromCsvSilent(csvFilePath, replaceListData);
    }
    catch (const CsvLoadException& ex) {
        std::wstring errorMessage = L"An error occurred while loading the settings: ";
        errorMessage += std::wstring(ex.what(), ex.what() + strlen(ex.what()));
        // MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
    updateHeaderSelection();
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    showListFilePath();
}

void MultiReplace::loadUIConfigFromIni() {
    auto [iniFilePath, _] = generateConfigFilePaths(); // Generating config file paths
    CFG.load(iniFilePath); (iniFilePath);

    // Load DPI Scaling factor from INI file
    float customScaleFactor = CFG.readFloat(L"Window", L"ScaleFactor", 1.0f);
    dpiMgr->setCustomScaleFactor(customScaleFactor);

    // Scale Window and List Size after loading ScaleFactor
    MIN_WIDTH_scaled = sx(MIN_WIDTH);
    MIN_HEIGHT_scaled = sy(MIN_HEIGHT);
    SHRUNK_HEIGHT_scaled = sy(SHRUNK_HEIGHT);
    DEFAULT_COLUMN_WIDTH_FIND_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND);
    DEFAULT_COLUMN_WIDTH_REPLACE_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE);
    DEFAULT_COLUMN_WIDTH_COMMENTS_scaled = sx(DEFAULT_COLUMN_WIDTH_COMMENTS);
    DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND_COUNT);
    DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE_COUNT);
    MIN_GENERAL_WIDTH_scaled = sx(MIN_GENERAL_WIDTH);

    // Load window position (using CFG.readInt)
    windowRect.left = CFG.readInt(L"Window", L"PosX", POS_X);
    windowRect.top = CFG.readInt(L"Window", L"PosY", POS_Y);

    // Load the state of UseList
    useListEnabled = CFG.readBool(L"Options", L"UseList", true);
    updateUseListState(false);

    // add extra width to align right edge of List/Edit with Scope group box at initial startup
    int DEFAULT_WIDTH_EXTRA = 23;

    // Load window width
    int savedWidth = CFG.readInt(L"Window", L"Width", sx(MIN_WIDTH + DEFAULT_WIDTH_EXTRA));
    int width = std::max(savedWidth, MIN_WIDTH_scaled);

    // Load useListOnHeight
    useListOnHeight = CFG.readInt(L"Window", L"Height", MIN_HEIGHT_scaled);
    useListOnHeight = std::max(useListOnHeight, MIN_HEIGHT_scaled);

    int height = useListEnabled ? useListOnHeight : useListOffHeight;
    windowRect.right = windowRect.left + width;
    windowRect.bottom = windowRect.top + height;

    // Read column widths
    findColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindWidth", DEFAULT_COLUMN_WIDTH_FIND_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceWidth", DEFAULT_COLUMN_WIDTH_REPLACE_scaled), MIN_GENERAL_WIDTH_scaled);
    commentsColumnWidth = std::max(CFG.readInt(L"ListColumns", L"CommentsWidth", DEFAULT_COLUMN_WIDTH_COMMENTS_scaled), MIN_GENERAL_WIDTH_scaled);
    findCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindCountWidth", DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceCountWidth", DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);

    // Load column visibility states
    isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);

    // Load column lock states
    findColumnLockedEnabled = CFG.readBool(L"ListColumns", L"FindColumnLocked", true);
    replaceColumnLockedEnabled = CFG.readBool(L"ListColumns", L"ReplaceColumnLocked", false);
    commentsColumnLockedEnabled = CFG.readBool(L"ListColumns", L"CommentsColumnLocked", true);

    // Load transparency settings (readByteFromIniCache)
    foregroundTransparency = std::clamp(CFG.readByte(L"Window", L"ForegroundTransparency", DEFAULT_FOREGROUND_TRANSPARENCY), MIN_TRANSPARENCY, MAX_TRANSPARENCY);
    backgroundTransparency = std::clamp(CFG.readByte(L"Window", L"BackgroundTransparency", DEFAULT_BACKGROUND_TRANSPARENCY), MIN_TRANSPARENCY, MAX_TRANSPARENCY);

    // Load Tooltip setting
    tooltipsEnabled = CFG.readBool(L"Options", L"Tooltips", true);
}

void MultiReplace::setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text) {
    ::SetDlgItemTextW(hDlg, itemID, text.c_str());
}

#pragma endregion


#pragma region Event Handling -- triggered in beNotified() in MultiReplace.cpp

void MultiReplace::processTextChange(SCNotification* notifyCode) {
    if (!isLoggingEnabled) {
        return;
    }

    Sci_Position cursorPosition = notifyCode->position;
    Sci_Position addedLines = notifyCode->linesAdded;
    Sci_Position notifyLength = notifyCode->length;

    Sci_Position lineNumber = ::SendMessage(getScintillaHandle(), SCI_LINEFROMPOSITION, cursorPosition, 0);
    if (notifyCode->modificationType & SC_MOD_INSERTTEXT) {
        if (addedLines != 0) {
            // Set the first entry as Modify
            logChanges.push_back({ ChangeType::Modify, lineNumber });

            // Instead of pushing multiple Insert entries in a loop, just push one block
            LogEntry insertEntry;
            insertEntry.changeType = ChangeType::Insert;
            insertEntry.lineNumber = lineNumber;
            insertEntry.blockSize  = static_cast<Sci_Position>(std::abs(addedLines));
            logChanges.push_back(insertEntry);
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (logChanges.empty() || logChanges.back().changeType != ChangeType::Modify || logChanges.back().lineNumber != lineNumber) {
                logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
    else if (notifyCode->modificationType & SC_MOD_DELETETEXT) {
        if (addedLines != 0) {
            // Special handling for deletions at position 0
            if (cursorPosition == 0 && notifyLength == 0) {
                logChanges.push_back({ ChangeType::Delete, 0 });
                return;
            }
            // Then, log the deletes in one block
            LogEntry deleteEntry;
            // Set the first entry as Modify for the smallest lineNumber
            logChanges.push_back({ ChangeType::Modify, lineNumber });

            deleteEntry.changeType = ChangeType::Delete;
            deleteEntry.lineNumber = lineNumber;
            deleteEntry.blockSize  = static_cast<Sci_Position>(std::abs(addedLines));
            logChanges.push_back(deleteEntry);
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (logChanges.empty() || logChanges.back().changeType != ChangeType::Modify || logChanges.back().lineNumber != lineNumber) {
                logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
}

void MultiReplace::processLog() {

    instance->handleDelimiterPositions(DelimiterOperation::Update);
}

void MultiReplace::onDocumentSwitched() {
    if (!isWindowOpen) {
        return;
    }

    // Get the current buffer ID of the newly opened document
    int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

    // Check if the document has changed
    if (currentBufferID != scannedDelimiterBufferID) {
        documentSwitched = true;
        isCaretPositionEnabled = false;
        scannedDelimiterBufferID = currentBufferID;

        if (instance != nullptr) {
            // Clear highlighting only if the tab was previously highlighted
            instance->handleClearColumnMarks();
            instance->showStatusMessage(L"", MessageStatus::Info);
        }
    }

    // Reset sorting state when switching documents
    originalLineOrder.clear();
    currentSortState = SortDirection::Unsorted;
    isSortedColumn = false;
    instance->UpdateSortButtonSymbols();
}

void MultiReplace::pointerToScintilla() {

    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which != -1) {
        instance->_hScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
        s_hScintilla = instance->_hScintilla;
    }

    if (instance->_hScintilla) { // just to supress Warning
        instance->pSciMsg = (SciFnDirect)::SendMessage(instance->_hScintilla, SCI_GETDIRECTFUNCTION, 0, 0);
        instance->pSciWndData = (sptr_t)::SendMessage(instance->_hScintilla, SCI_GETDIRECTPOINTER, 0, 0);
    }
}

void MultiReplace::onSelectionChanged()
{
    static bool wasTextSelected = false;   // remember previous selection state

    HWND hDlg = getDialogHandle();         // dialog handle once for reuse

    // -----------------------------------------------------------------------
    // 1) “Replace in Files” mode:
    //    - Selection-Radio ist dort nutzlos → immer deaktivieren
    //    - Nur wenn er noch angehakt ist, auf All Text umschalten
    //    - Dann sofort zurückkehren, damit er nicht erneut aktiviert wird
    // -----------------------------------------------------------------------
    if (instance && (instance->isReplaceInFiles || instance->isFindAllInFiles))
    {
        HWND hSel = ::GetDlgItem(hDlg, IDC_SELECTION_RADIO);
        ::EnableWindow(hSel, FALSE);

        if (::SendMessage(hSel, BM_GETCHECK, 0, 0) == BST_CHECKED)
        {
            ::CheckRadioButton(
                hDlg,
                IDC_ALL_TEXT_RADIO,     // first in radio group
                IDC_COLUMN_MODE_RADIO,  // last  in radio group
                IDC_ALL_TEXT_RADIO      // button to check
            );
        }
        return;    // nothing else must re-enable Selection in this mode
    }

    // -----------------------------------------------------------------------
    // 2) Normal Replace-All / Replace-in-Opened-Docs modes
    // -----------------------------------------------------------------------
    Sci_Position start = ::SendMessage(getScintillaHandle(),
        SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(getScintillaHandle(),
        SCI_GETSELECTIONEND, 0, 0);
    bool isTextSelected = (start != end);

    HWND hSel = ::GetDlgItem(hDlg, IDC_SELECTION_RADIO);
    ::EnableWindow(hSel, isTextSelected);

    // If no text is selected but Selection is still checked → switch to All Text
    if (!isTextSelected &&
        ::SendMessage(hSel, BM_GETCHECK, 0, 0) == BST_CHECKED)
    {
        ::CheckRadioButton(
            hDlg,
            IDC_ALL_TEXT_RADIO,
            IDC_COLUMN_MODE_RADIO,
            IDC_ALL_TEXT_RADIO
        );
    }

    // Inform other UI parts when we just lost a selection
    if (wasTextSelected && !isTextSelected)
    {
        if (instance)
            instance->setUIElementVisibility();
    }
    wasTextSelected = isTextSelected;
}

void MultiReplace::onTextChanged() {
    textModified = true;
}

void MultiReplace::onCaretPositionChanged()
{
    if (!isCaretPositionEnabled) {
        return;
    }

    LRESULT startPosition = ::SendMessage(getScintillaHandle(), SCI_GETCURRENTPOS, 0, 0);
    if (instance != nullptr) {
        instance->showStatusMessage(LanguageManager::instance().get(L"status_actual_position", { instance->addLineAndColumnMessage(startPosition) }), MessageStatus::Success);
    }

}

void MultiReplace::onThemeChanged()
{
    // Update all theme-related colors (status messages and columns)
    if (instance) {
        {
            instance->applyThemePalette();          // status colours, tooltip repaint
            instance->refreshColumnStylesIfNeeded(); // guarded lexer reset
        }
    }
}

void MultiReplace::signalShutdown() {
    // This static method can be called from the global beNotified function
    if (instance) {
        instance->_isShuttingDown = true;
        instance->_isCancelRequested = true; // Also trigger the cancel flag
    }
}

#pragma endregion


#pragma region Debug DPI Information

BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData) {
    (void)hdcMonitor;  // Mark hdcMonitor as unused
    (void)lprcMonitor; // Mark lprcMonitor as unused

    MonitorEnumData* pData = reinterpret_cast<MonitorEnumData*>(dwData);

    MONITORINFOEX monitorInfoEx;
    monitorInfoEx.cbSize = sizeof(MONITORINFOEX);
    GetMonitorInfo(hMonitor, &monitorInfoEx);

    int screenWidth = monitorInfoEx.rcMonitor.right - monitorInfoEx.rcMonitor.left;
    int screenHeight = monitorInfoEx.rcMonitor.bottom - monitorInfoEx.rcMonitor.top;
    BOOL isPrimary = (monitorInfoEx.dwFlags & MONITORINFOF_PRIMARY) != 0;

    // Add monitor info to the buffer
    pData->monitorInfo += L"Monitor " + std::to_wstring(pData->monitorCount + 1) + L": "
        + (isPrimary ? L"Primary, " : L"Secondary, ")
        + std::to_wstring(screenWidth) + L"x" + std::to_wstring(screenHeight) + L"\n";

    // Increment the monitor counter
    pData->monitorCount++;

    // Check if this is the primary monitor
    if (isPrimary) {
        pData->primaryMonitorIndex = pData->monitorCount;
    }

    // Check if this is the current monitor where the window is
    if (MonitorFromWindow(GetForegroundWindow(), MONITOR_DEFAULTTONEAREST) == hMonitor) {
        pData->currentMonitor = pData->monitorCount;
    }

    return TRUE;  // Continue enumerating monitors
}

void MultiReplace::showDPIAndFontInfo()
{
    RECT sizeWindowRect;
    GetClientRect(_hSelf, &sizeWindowRect);  // Get window dimensions

    HDC hDC = GetDC(_hSelf);  // Get the device context (DC) for the current window
    if (hDC)
    {
        // Get current font of the window
        HFONT currentFont = (HFONT)SendMessage(_hSelf, WM_GETFONT, 0, 0);
        SelectObject(hDC, currentFont);  // Select current font into the DC

        TEXTMETRIC tmCurrent;
        GetTextMetrics(hDC, &tmCurrent);  // Retrieve text metrics for current font

        SelectObject(hDC, _hStandardFont);  // Select standard font into the DC
        TEXTMETRIC tmStandard;
        GetTextMetrics(hDC, &tmStandard);  // Retrieve text metrics for standard font

        SIZE size;
        GetTextExtentPoint32W(hDC, L"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", 52, &size);
        int baseUnitX = (size.cx / 26 + 1) / 2;
        int baseUnitY = tmCurrent.tmHeight;
        int duWidth = MulDiv(sizeWindowRect.right, 4, baseUnitX);
        int duHeight = MulDiv(sizeWindowRect.bottom, 8, baseUnitY);

        // Get DPI values from the DPIManager
        int dpix = dpiMgr->getDPIX();
        int dpiy = dpiMgr->getDPIY();
        float customScaleFactor = dpiMgr->getCustomScaleFactor();
        int scaledDpiX = dpiMgr->scaleX(96);
        int scaledDpiY = dpiMgr->scaleY(96);

        wchar_t scaleBuffer[10];
        swprintf(scaleBuffer, 10, L"%.1f", customScaleFactor);

        MonitorEnumData monitorData = {};
        monitorData.monitorCount = 0;
        monitorData.currentMonitor = 0;
        monitorData.primaryMonitorIndex = 0;

        EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, reinterpret_cast<LPARAM>(&monitorData));

        std::wstring message =
            L"On Monitor " + std::to_wstring(monitorData.currentMonitor) + L"\n"

            + monitorData.monitorInfo + L"\n"

            L"Window Size DUs: " + std::to_wstring(duWidth) + L"x" + std::to_wstring(duHeight) + L"\n"
            L"Scaled DPI: " + std::to_wstring(dpix) + L"x" + std::to_wstring(dpiy) + L" * " + scaleBuffer + L" = "
            + std::to_wstring(scaledDpiX) + L"x" + std::to_wstring(scaledDpiY) + L"\n\n"

            L"Windows Font: Height=" + std::to_wstring(tmCurrent.tmHeight) + L", Ascent=" + std::to_wstring(tmCurrent.tmAscent) +
            L", Descent=" + std::to_wstring(tmCurrent.tmDescent) + L", Weight=" + std::to_wstring(tmCurrent.tmWeight) + L"\n"
            L"Plugin Font: Height=" + std::to_wstring(tmStandard.tmHeight) + L", Ascent=" + std::to_wstring(tmStandard.tmAscent) +
            L", Descent=" + std::to_wstring(tmStandard.tmDescent) + L", Weight=" + std::to_wstring(tmStandard.tmWeight);

        MessageBox(_hSelf, message.c_str(), L"Window, Monitor, DPI, and Font Info", MB_ICONINFORMATION | MB_OK);

        ReleaseDC(_hSelf, hDC);
    }
    else
    {
        MessageBox(_hSelf, L"Failed to retrieve device context (HDC).", L"Error", MB_OK);
    }
}

#pragma endregion