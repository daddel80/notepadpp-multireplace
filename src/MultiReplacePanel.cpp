// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#define NOMINMAX


#include "MultiReplacePanel.h"
#include "Notepad_plus_msgs.h"
#include "PluginDefinition.h"
#include "Scintilla.h"
#include "DPIManager.h"

#include <algorithm>
#include <bitset>
#include <codecvt>
#include <Commctrl.h>
#include <fstream>
#include <functional>
#include <iostream>
#include <locale>
#include <map>
#include <regex>
#include <set>
#include <sstream>
#include <string>
#include <unordered_map>
#include <vector>
#include <windows.h>
#include <filesystem> 
#include <lua.hpp> 
#include <unordered_set>

#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif

HWND MultiReplace::s_hScintilla = nullptr;
HWND MultiReplace::s_hDlg = nullptr;
bool MultiReplace::isWindowOpen = false;
bool MultiReplace::isLoggingEnabled = true;
bool MultiReplace::textModified = true;
bool MultiReplace::documentSwitched = false;
bool MultiReplace::isCaretPositionEnabled = false;
bool MultiReplace::isLuaErrorDialogEnabled = true;
int MultiReplace::scannedDelimiterBufferID = -1;
std::map<int, ControlInfo> MultiReplace::ctrlMap;
std::vector<MultiReplace::LogEntry> MultiReplace::logChanges;
MultiReplace* MultiReplace::instance = nullptr;

std::vector<size_t> MultiReplace::originalLineOrder;
SortDirection MultiReplace::currentSortState = SortDirection::Unsorted;
bool MultiReplace::isSortedColumn = false;
POINT MultiReplace::debugWindowPosition = { CW_USEDEFAULT, CW_USEDEFAULT };
bool MultiReplace::debugWindowPositionSet = false;
int MultiReplace::debugWindowResponse = -1;
SIZE MultiReplace::debugWindowSize = { 400, 250 };
bool MultiReplace::debugWindowSizeSet = false;
HWND MultiReplace::hDebugWnd = NULL;


#pragma warning(disable: 6262)

#pragma region Initialization

void MultiReplace::initializeWindowSize() {

    loadUIConfigFromIni(); // Loads the UI configuration, including window size and position

    // Set the window position and size based on the loaded settings
    SetWindowPos(_hSelf, NULL, windowRect.left, windowRect.top,
        windowRect.right - windowRect.left,
        windowRect.bottom - windowRect.top, SWP_NOZORDER);
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
    for (int controlId : { IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_STATUS_MESSAGE, IDC_PATH_DISPLAY }) {
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

    int minHeight = useListEnabled ? MIN_HEIGHT_scaled : SHRUNK_HEIGHT_scaled;
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
    int buttonX = windowWidth - sx(33 + 128);
    int checkbox2X = buttonX + sx(134);
    int useListButtonX = buttonX + sx(133);
    int swapButtonX = windowWidth - sx(33 + 128 + 26);
    int comboWidth = windowWidth - sx(289);
    int listWidth = windowWidth - sx(202);
    int listHeight = std::max(windowHeight - sy(245), sy(20)); // Minimum listHeight to prevent IDC_PATH_DISPLAY from overlapping with IDC_STATUS_MESSAGE
    int useListButtonY = windowHeight - sy(34);
    
    // Apply scaling only when assigning to ctrlMap
    ctrlMap[IDC_STATIC_FIND] = { sx(11), sy(18), sx(80), sy(19), WC_STATIC, getLangStrLPCWSTR(L"panel_find_what"), SS_RIGHT, NULL };
    ctrlMap[IDC_STATIC_REPLACE] = { sx(11), sy(47), sx(80), sy(19), WC_STATIC, getLangStrLPCWSTR(L"panel_replace_with"), SS_RIGHT };

    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { sx(16), sy(76), sx(158), checkboxHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_match_whole_word_only"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { sx(16), sy(101), sx(158), checkboxHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_match_case"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_CHECKBOX] = { sx(16), sy(126), sx(134), checkboxHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_use_variables"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_HELP] = { sx(152), sy(126), sx(20), sy(20), WC_BUTTON, getLangStrLPCWSTR(L"panel_help"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_FIRST_CHECKBOX] = { sx(16), sy(151), sx(158), checkboxHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_replace_first_match_only"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_WRAP_AROUND_CHECKBOX] = { sx(16), sy(176), sx(158), checkboxHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_wrap_around"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };

    ctrlMap[IDC_SEARCH_MODE_GROUP] = { sx(180), sy(79), sx(173), sy(104), WC_BUTTON, getLangStrLPCWSTR(L"panel_search_mode"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_NORMAL_RADIO] = { sx(188), sy(101), sx(162), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_normal"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXTENDED_RADIO] = { sx(188), sy(126), sx(162), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_extended"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REGEX_RADIO] = { sx(188), sy(150), sx(162), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_regular_expression"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_SCOPE_GROUP] = { sx(367), sy(79), sx(203), sy(125), WC_BUTTON, getLangStrLPCWSTR(L"panel_scope"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_ALL_TEXT_RADIO] = { sx(375), sy(101), sx(189), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_all_text"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_SELECTION_RADIO] = { sx(375), sy(126), sx(189), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_selection"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COLUMN_MODE_RADIO] = { sx(375), sy(150), sx(45), radioButtonHeight, WC_BUTTON, getLangStrLPCWSTR(L"panel_csv"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_COLUMN_NUM_STATIC] = { sx(369), sy(181), sx(30), sy(20), WC_STATIC, getLangStrLPCWSTR(L"panel_cols"), SS_RIGHT, NULL };
    ctrlMap[IDC_COLUMN_NUM_EDIT] = { sx(400), sy(181), sx(41), sy(16), WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, getLangStrLPCWSTR(L"tooltip_columns") };
    ctrlMap[IDC_DELIMITER_STATIC] = { sx(442), sy(181), sx(38), sy(20), WC_STATIC, getLangStrLPCWSTR(L"panel_delim"), SS_RIGHT, NULL };
    ctrlMap[IDC_DELIMITER_EDIT] = { sx(481), sy(181), sx(25), sy(16), WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, getLangStrLPCWSTR(L"tooltip_delimiter") };
    ctrlMap[IDC_QUOTECHAR_STATIC] = { sx(506), sy(181), sx(37), sy(20), WC_STATIC, getLangStrLPCWSTR(L"panel_quote"), SS_RIGHT, NULL };
    ctrlMap[IDC_QUOTECHAR_EDIT] = { sx(544), sy(181), sx(15), sy(16), WC_EDIT, NULL, ES_CENTER | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, getLangStrLPCWSTR(L"tooltip_quote") };

    ctrlMap[IDC_COLUMN_SORT_DESC_BUTTON] = { sx(418), sy(149), sx(19), sy(20), WC_BUTTON, symbolSortDesc, BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_sort_descending") };
    ctrlMap[IDC_COLUMN_SORT_ASC_BUTTON] = { sx(437), sy(149), sx(19), sy(20), WC_BUTTON, symbolSortAsc, BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_sort_ascending") };
    ctrlMap[IDC_COLUMN_DROP_BUTTON] = { sx(459), sy(149), sx(25), sy(20), WC_BUTTON, L"✖", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_drop_columns") };
    ctrlMap[IDC_COLUMN_COPY_BUTTON] = { sx(487), sy(149), sx(25), sy(20), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_copy_columns") }; //🗍
    ctrlMap[IDC_COLUMN_HIGHLIGHT_BUTTON] = { sx(515), sy(149), sx(45), sy(20), WC_BUTTON, L"🖍", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_column_highlight") };

    ctrlMap[IDC_STATUS_MESSAGE] = { sx(19), sy(208), sx(530), sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT, NULL };
    
    // Dynamic positions and sizes
    ctrlMap[IDC_FIND_EDIT] = { sx(96), sy(14), comboWidth, sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_EDIT] = { sx(96), sy(44), comboWidth, sy(160), WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_SWAP_BUTTON] = { swapButtonX, sy(26), sx(22), sy(27), WC_BUTTON, L"⇅", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, sy(14), sx(128), sy(52), WC_BUTTON, getLangStrLPCWSTR(L"panel_add_into_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, sy(91), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_replace_all"), BS_SPLITBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_BUTTON] = { buttonX, sy(91), sx(96), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_replace"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_SMALL_BUTTON] = { buttonX + sx(100), sy(91), sx(28), sy(24), WC_BUTTON, L"↻", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_replace_all") };
    ctrlMap[IDC_2_BUTTONS_MODE] = { checkbox2X, sy(91), sx(20), sy(20), WC_BUTTON, L"", BS_AUTOCHECKBOX | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_2_buttons_mode") };
    ctrlMap[IDC_FIND_BUTTON] = { buttonX, sy(119), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_find_next"), BS_PUSHBUTTON | WS_TABSTOP, NULL };

    findNextButtonText = L"▼ " + getLangStr(L"panel_find_next_small");
    ctrlMap[IDC_FIND_NEXT_BUTTON] = ControlInfo{ buttonX + sx(32), sy(119), sx(96), sy(24), WC_BUTTON, findNextButtonText.c_str(), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    
    ctrlMap[IDC_FIND_PREV_BUTTON] = { buttonX, sy(119), sx(28), sy(24), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_BUTTON] = { buttonX, sy(147), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_mark_matches"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, sy(147), sx(96), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_mark_matches_small"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX + sx(100), sy(147), sx(28), sy(24), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_copy_marked_text") }; //🗍
    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, sy(175), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_clear_all_marks"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, sy(227), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_LOAD_LIST_BUTTON] = { buttonX, sy(227), sx(96), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_NEW_LIST_BUTTON] = { buttonX + sx(100), sy(227), sx(28), sy(24), WC_BUTTON, L"➕", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_new_list") };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, sy(255), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_save_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_SAVE_BUTTON] = { buttonX, sy(255), sx(28), sy(24), WC_BUTTON, L"💾", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_save") };
    ctrlMap[IDC_SAVE_AS_BUTTON] = { buttonX + sx(32), sy(255), sx(96), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_save_as"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, sy(283), sx(128), sy(24), WC_BUTTON, getLangStrLPCWSTR(L"panel_export_to_bash"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_UP_BUTTON] = { buttonX + sx(4), sy(323), sx(24), sy(24), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_DOWN_BUTTON] = { buttonX + sx(4), sy(323 + 24 + 4), sx(24), sy(24), WC_BUTTON, L"▼", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_SHIFT_FRAME] = { buttonX, sy(323 - 11), sx(128), sy(68), WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_SHIFT_TEXT] = { buttonX + sx(30), sy(323 + 16), sx(96), sy(16), WC_STATIC, getLangStrLPCWSTR(L"panel_move_lines"), SS_LEFT, NULL };
    ctrlMap[IDC_REPLACE_LIST] = { sx(14), sy(227), listWidth, listHeight, WC_LISTVIEW, NULL, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS, NULL };
    ctrlMap[IDC_PATH_DISPLAY] = { sx(14), sy(225) + listHeight + sy(5), listWidth, sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT, NULL };
    ctrlMap[IDC_USE_LIST_BUTTON] = { useListButtonX, useListButtonY, sx(22), sy(22), WC_BUTTON, L"-", BS_PUSHBUTTON | WS_TABSTOP, NULL };
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

    for (auto& pair : ctrlMap)
    {
        HWND hwndControl = CreateWindowEx(
            0,                          // No additional window styles.
            pair.second.className,      // Window class
            pair.second.windowName,     // Window text
            pair.second.style | WS_CHILD | WS_VISIBLE, // Window style
            pair.second.x,              // x position
            pair.second.y,              // y position
            pair.second.cx,             // width
            pair.second.cy,             // height
            _hSelf,                     // Parent window    
            (HMENU)(INT_PTR)pair.first, // Menu, or child-window identifier
            hInstance,                  // The window instance.
            NULL                        // Additional application data.
        );

        if (hwndControl == NULL)
        {
            DWORD dwError = GetLastError();
            std::wstring errorMsg = getLangStr(L"msgbox_failed_create_control", { std::to_wstring(pair.first), std::to_wstring(dwError) });
            MessageBox(nppData._nppHandle, errorMsg.c_str(), getLangStr(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
            return false;
        }

        // Only create tooltips if enabled and text is available
        if (tooltipsEnabled && pair.second.tooltipText != nullptr && pair.second.tooltipText[0] != '\0')
        {
            HWND hwndTooltip = CreateWindowEx(
                NULL, TOOLTIPS_CLASS, NULL,
                WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON,
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                _hSelf, NULL, hInstance, NULL);

            if (hwndTooltip)
            {
                // Associate the tooltip with the control
                TOOLINFO toolInfo = { 0 };
                toolInfo.cbSize = sizeof(toolInfo);
                toolInfo.hwnd = _hSelf;
                toolInfo.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
                toolInfo.uId = (UINT_PTR)hwndControl;
                toolInfo.lpszText = (LPWSTR)pair.second.tooltipText;
                SendMessage(hwndTooltip, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);
            }
        }
        // Show the window
        ShowWindow(hwndControl, SW_SHOW);
        UpdateWindow(hwndControl);
    }
    return true;
}

void MultiReplace::initializeMarkerStyle() {
    // Initialize for non-list marker
    long standardMarkerColor = MARKER_COLOR;
    int standardMarkerStyle = textStyles[0];
    colorToStyleMap[standardMarkerColor] = standardMarkerStyle;

    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, standardMarkerStyle, 0);
    ::SendMessage(_hScintilla, SCI_INDICSETSTYLE, standardMarkerStyle, INDIC_STRAIGHTBOX);
    ::SendMessage(_hScintilla, SCI_INDICSETFORE, standardMarkerStyle, standardMarkerColor);
    ::SendMessage(_hScintilla, SCI_INDICSETALPHA, standardMarkerStyle, 100);
}

void MultiReplace::initializeListView() {
    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    originalListViewProc = (WNDPROC)SetWindowLongPtr(_replaceListView, GWLP_WNDPROC, (LONG_PTR)ListViewSubclassProc);
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
        IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON, IDC_2_BUTTONS_MODE, IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON,
        IDC_FIND_PREV_BUTTON, IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON,
        IDC_USE_LIST_BUTTON, IDC_LOAD_FROM_CSV_BUTTON, IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON, IDC_SAVE_TO_CSV_BUTTON,
        IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT, IDC_EXPORT_BASH_BUTTON, IDC_PATH_DISPLAY
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
        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT) {
            SendMessage(resizeHwnd, CB_GETEDITSEL, (WPARAM)&startSelection, (LPARAM)&endSelection);
        }

        int height = ctrlInfo.cy;
        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT) {
            COMBOBOXINFO cbi = { sizeof(COMBOBOXINFO) };
            if (GetComboBoxInfo(resizeHwnd, &cbi)) {
                height = cbi.rcItem.bottom - cbi.rcItem.top;
            }
        }

        MoveWindow(resizeHwnd, ctrlInfo.x, ctrlInfo.y, ctrlInfo.cx, height, TRUE);

        if (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT) {
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
    setVisibility({ IDC_FIND_BUTTON }, !twoButtonsMode);

    // Mark-Buttons
    setVisibility({ IDC_MARK_MATCHES_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_MARK_BUTTON }, !twoButtonsMode);

    // Load-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_LOAD_FROM_CSV_BUTTON }, !twoButtonsMode);

    // Save-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_SAVE_TO_CSV_BUTTON }, !twoButtonsMode);

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
    EnableWindow(GetDlgItem(_hSelf, IDC_FIND_PREV_BUTTON), !regexChecked );
}

void MultiReplace::drawGripper() {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(_hSelf, &ps);

    // Get the size of the client area
    RECT rect;
    GetClientRect(_hSelf, &rect);

    // Determine where to draw the gripper
    int gripperAreaSize = 13; // Total size of the gripper area
    POINT startPoint = { rect.right - gripperAreaSize, rect.bottom - gripperAreaSize };

    // Define the new size and reduced gap of the gripper dots
    int dotSize = 3; // Increased dot size
    int gap = 1; // Reduced gap between dots

    // Brush Color for Gripper
    HBRUSH hBrush = CreateSolidBrush(RGB(200, 200, 200));

    // Matrix to identify the points to draw
    int positions[3][3] = {
        {0, 0, 1},
        {0, 1, 1},
        {1, 1, 1}
    };

    for (int row = 0; row < 3; row++)
    {
        for (int col = 0; col < 3; col++)
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
    showStatusMessage(useListEnabled ? getLangStr(L"status_enable_list") : getLangStr(L"status_disable_list"), COLOR_INFO);

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
    LPCWSTR tooltipText = useListEnabled ? getLangStrLPCWSTR(L"tooltip_disable_list") : getLangStrLPCWSTR(L"tooltip_enable_list");
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

void MultiReplace::undo() {
    if (!undoStack.empty()) {
        // Get the last action
        UndoRedoAction action = undoStack.back();
        undoStack.pop_back();

        // Execute the undo action
        action.undoAction();

        // Push the action onto the redoStack
        redoStack.push_back(action);
    }
    else {
        // showStatusMessage(L"Nothing to undo.", COLOR_INFO);
    }
}

void MultiReplace::redo() {
    if (!redoStack.empty()) {
        // Get the last action
        UndoRedoAction action = redoStack.back();
        redoStack.pop_back();

        // Execute the redo action
        action.redoAction();

        // Push the action back onto the undoStack
        undoStack.push_back(action);
    }
    else {
        // showStatusMessage(L"Nothing to redo.", COLOR_INFO);
    }
}

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

    // Clear the redoStack since a new action invalidates the redo history
    redoStack.clear();

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
    undoStack.push_back(action);
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

    // Clear redoStack since a new action invalidates redo history
    redoStack.clear();

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
    undoStack.push_back(action);
}

void MultiReplace::modifyItemInReplaceList(size_t index, const ReplaceItemData& newData) {
    // Store the original data
    ReplaceItemData originalData = replaceListData[index];

    // Modify the item
    replaceListData[index] = newData;

    // Update the ListView item
    updateListViewItem(index);

    // Clear the redoStack
    redoStack.clear();

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
    undoStack.push_back(action);
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

    // Clear the redoStack to invalidate future actions
    redoStack.clear();

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
    undoStack.push_back(action);

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

void MultiReplace::sortItemsInReplaceList(const std::vector<size_t>& originalOrder,
    const std::vector<size_t>& newOrder,
    const std::map<int, SortDirection>& previousColumnSortOrder,
    int columnID,
    SortDirection direction) {
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
    undoStack.push_back(action);
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
    for (int i = columnCount - 1; i >= 0; i--) {
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
        lvc.pszText = getLangStrLPWSTR(L"header_find_count");
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
        lvc.pszText = getLangStrLPWSTR(L"header_replace_count");
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
    lvc.pszText = getLangStrLPWSTR(L"header_find");
    lvc.cx = (findColumnLockedEnabled ? findColumnWidth : perColumnWidth);
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::FIND_TEXT] = currentIndex;
    ++currentIndex;

    // Column 5: Replace Text (dynamic width)
    lvc.iSubItem = currentIndex;
    lvc.pszText = getLangStrLPWSTR(L"header_replace");
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
        lvc.pszText = getLangStrLPWSTR(options[i]);
        lvc.cx = checkMarkWidth_scaled;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[static_cast<ColumnID>(static_cast<int>(ColumnID::WHOLE_WORD) + i)] = currentIndex;
        ++currentIndex;
    }

    // Column 11: Comments (dynamic width)
    if (isCommentsColumnVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = getLangStrLPWSTR(L"header_comments");
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
        showStatusMessage(getLangStr(L"status_no_find_string"), COLOR_ERROR);
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
        message = getLangStr(L"status_duplicate_entry") + itemData.findText;
    }
    else {
        message = getLangStr(L"status_value_added");
    }
    showStatusMessage(message, COLOR_SUCCESS);

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
    /*
    ResizableColWidths widths = {
        listView,
        listCtrlInfo.cx,  // Width of the ListView control
        findCountColumnWidth,
        replaceCountColumnWidth,
        findColumnWidth,
        replaceColumnWidth,
        commentsColumnWidth,
        deleteButtonColumnWidth,
        GetSystemMetrics(SM_CXVSCROLL)  // Width of the vertical scrollbar
    };
    */

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
    MoveWindow(listView, listCtrlInfo.x, listCtrlInfo.y, listCtrlInfo.cx, listCtrlInfo.cy, TRUE);

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
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::WHOLE_WORD], getLangStrLPWSTR(L"tooltip_header_whole_word"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::MATCH_CASE], getLangStrLPWSTR(L"tooltip_header_match_case"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::USE_VARIABLES], getLangStrLPWSTR(L"tooltip_header_use_variables"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::EXTENDED], getLangStrLPWSTR(L"tooltip_header_extended"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::REGEX], getLangStrLPWSTR(L"tooltip_header_regex"));
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
        showStatusMessage(getLangStr(L"status_no_rows_selected_to_shift"), COLOR_ERROR);
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
    showStatusMessage(getLangStr(L"status_rows_shifted", { std::to_wstring(selectedIndices.size()) }), COLOR_SUCCESS);
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

    showStatusMessage(getLangStr(L"status_one_line_deleted"), COLOR_SUCCESS);
}

void MultiReplace::deleteSelectedLines() {
    // Collect selected indices
    std::vector<size_t> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(static_cast<size_t>(i));
    }

    if (selectedIndices.empty()) {
        showStatusMessage(getLangStr(L"status_no_rows_selected_to_delete"), COLOR_ERROR);
        return;
    }

    // Ensure indices are valid before calling removeItemsFromReplaceList
    for (size_t index : selectedIndices) {
        if (index >= replaceListData.size()) {
            showStatusMessage(getLangStr(L"status_invalid_indices"), COLOR_ERROR);
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
    showStatusMessage(getLangStr(L"status_lines_deleted", { std::to_wstring(selectedIndices.size()) }), COLOR_SUCCESS);
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

    // Perform the sorting
    std::sort(replaceListData.begin(), replaceListData.end(),
        [this, columnID, direction](const ReplaceItemData& a, const ReplaceItemData& b) -> bool {
            switch (columnID) {
            case ColumnID::FIND_COUNT: {
                int numA = a.findCount.empty() ? -1 : std::stoi(a.findCount);
                int numB = b.findCount.empty() ? -1 : std::stoi(b.findCount);
                return direction == SortDirection::Ascending ? numA < numB : numA > numB;
            }
            case ColumnID::REPLACE_COUNT: {
                int numA = a.replaceCount.empty() ? -1 : std::stoi(a.replaceCount);
                int numB = b.replaceCount.empty() ? -1 : std::stoi(b.replaceCount);
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
    // Access the item data from the list
    ReplaceItemData& itemData = replaceListData[itemIndex];

    // Update findCount if provided
    if (findCount != -1) {
        itemData.findCount = std::to_wstring(findCount);
    }

    // Update replaceCount if provided
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

    // Show a status message to indicate the list was cleared
    showStatusMessage(getLangStr(L"status_list_cleared"), COLOR_SUCCESS);

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

#pragma endregion


#pragma region Contextmenu Display Columns

void MultiReplace::showColumnVisibilityMenu(HWND hWnd, POINT pt) {
    // Create a popup menu
    HMENU hMenu = CreatePopupMenu();

    // Add menu items with checkmarks based on current visibility
    AppendMenu(hMenu, MF_STRING | (isFindCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_FIND_COUNT, getLangStrLPCWSTR(L"header_find_count"));
    AppendMenu(hMenu, MF_STRING | (isReplaceCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_REPLACE_COUNT, getLangStrLPCWSTR(L"header_replace_count"));
    AppendMenu(hMenu, MF_STRING | (isCommentsColumnVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_COMMENTS, getLangStrLPCWSTR(L"header_comments"));
    AppendMenu(hMenu, MF_STRING | (isDeleteButtonVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_DELETE, getLangStrLPCWSTR(L"header_delete_button"));

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
    wchar_t itemText[MAX_TEXT_LENGTH];
    ListView_GetItemText(_replaceListView, itemIndex, column, itemText, MAX_TEXT_LENGTH);
    itemText[MAX_TEXT_LENGTH - 1] = L'\0';
    SetWindowText(hwndEdit, itemText);

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

void MultiReplace::closeEditField(bool commitChanges) {
    if (!hwndEdit) {
        return; // No active edit field present
    }

    if (commitChanges &&
        _editingColumnID != ColumnID::INVALID &&
        _editingItemIndex >= 0 &&
        _editingItemIndex < static_cast<int>(replaceListData.size())) {

        // Retrieve the new text from the edit field
        wchar_t newText[MAX_TEXT_LENGTH];
        GetWindowText(hwndEdit, newText, MAX_TEXT_LENGTH);

        // Store the original data
        ReplaceItemData originalData = replaceListData[_editingItemIndex];

        // Create new data based on the edited column
        ReplaceItemData newData = originalData;
        bool hasChanged = false;
        switch (_editingColumnID) {
        case ColumnID::FIND_TEXT:
            if (wcscmp(originalData.findText.c_str(), newText) != 0) {
                newData.findText = newText;
                hasChanged = true;
            }
            break;
        case ColumnID::REPLACE_TEXT:
            if (wcscmp(originalData.replaceText.c_str(), newText) != 0) {
                newData.replaceText = newText;
                hasChanged = true;
            }
            break;
        case ColumnID::COMMENTS:
            if (wcscmp(originalData.comments.c_str(), newText) != 0) {
                newData.comments = newText;
                hasChanged = true;
            }
            break;
        default:
            break; // Non-editable column
        }

        if (hasChanged) {
            // Apply changes and manage Undo/Redo
            modifyItemInReplaceList(_editingItemIndex, newData);
        }
    }

    // Destroy the edit field
    DestroyWindow(hwndEdit);
    hwndEdit = nullptr;

    // Destroy the expand button if it exists
    if (hwndExpandBtn && IsWindow(hwndExpandBtn)) {
        DestroyWindow(hwndExpandBtn);
        hwndExpandBtn = nullptr;
    }
    _editIsExpanded = false;

    // Reset editing state
    _editingItemIndex = -1;
    _editingColumnIndex = -1;
    _editingColumnID = ColumnID::INVALID;

    // Restore hover text only if we suppressed it
    if (isHoverTextSuppressed) {
        isHoverTextSuppressed = false;

        // If user has hover text enabled in general, re-add LVS_EX_INFOTIP
        if (isHoverTextEnabled) {
            DWORD extendedStyle = ListView_GetExtendedListViewStyle(_replaceListView);
            extendedStyle |= LVS_EX_INFOTIP;
            ListView_SetExtendedListViewStyle(_replaceListView, extendedStyle);
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

LRESULT CALLBACK MultiReplace::ListViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    MultiReplace* pThis = reinterpret_cast<MultiReplace*>(GetWindowLongPtr(GetParent(hwnd), GWLP_USERDATA));

    switch (msg) {
    case WM_VSCROLL:
    case WM_MOUSEWHEEL:
    case WM_HSCROLL:
        // Close the edit control if it exists and is visible
        if (pThis->hwndEdit && IsWindow(pThis->hwndEdit)) {
            DestroyWindow(pThis->hwndEdit);
            pThis->hwndEdit = NULL;
        }
        // Allow the default list view procedure to handle standard scrolling
        break;

    case WM_NOTIFY: {
        NMHDR* pnmhdr = reinterpret_cast<NMHDR*>(lParam);
        if (pnmhdr->hwndFrom == ListView_GetHeader(hwnd)) {
            int code = static_cast<int>(static_cast<short>(pnmhdr->code));

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

    default:
        break;
    }

    return CallWindowProc(pThis->originalListViewProc, hwnd, msg, wParam, lParam);
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
        newHeight = curHeight / 3; // Collapse
        SetWindowText(hwndExpandBtn, L"↓");
    }
    else {
        newHeight = curHeight * 3; // Expand
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
        AppendMenu(hMenu, MF_STRING | (state.canUndo ? MF_ENABLED : MF_GRAYED), IDM_UNDO, getLangStr(L"ctxmenu_undo").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canRedo ? MF_ENABLED : MF_GRAYED), IDM_REDO, getLangStr(L"ctxmenu_redo").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_CUT_LINES_TO_CLIPBOARD, getLangStr(L"ctxmenu_cut").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_COPY_LINES_TO_CLIPBOARD, getLangStr(L"ctxmenu_copy").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canPaste ? MF_ENABLED : MF_GRAYED), IDM_PASTE_LINES_FROM_CLIPBOARD, getLangStr(L"ctxmenu_paste").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_SELECT_ALL, getLangStr(L"ctxmenu_select_all").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.canEdit ? MF_ENABLED : MF_GRAYED), IDM_EDIT_VALUE, getLangStr(L"ctxmenu_edit").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_DELETE_LINES, getLangStr(L"ctxmenu_delete").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_ADD_NEW_LINE, getLangStr(L"ctxmenu_add_new_line").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.clickedOnItem ? MF_ENABLED : MF_GRAYED), IDM_COPY_DATA_TO_FIELDS, getLangStr(L"ctxmenu_transfer_to_input_fields").c_str());
        AppendMenu(hMenu, MF_STRING | (state.listNotEmpty ? MF_ENABLED : MF_GRAYED), IDM_SEARCH_IN_LIST, getLangStr(L"ctxmenu_search_in_list").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allEnabled ? MF_ENABLED : MF_GRAYED), IDM_ENABLE_LINES, getLangStr(L"ctxmenu_enable").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allDisabled ? MF_ENABLED : MF_GRAYED), IDM_DISABLE_LINES, getLangStr(L"ctxmenu_disable").c_str());
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
    for (int i = 0; i < columnCount; i++) {
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

    state.canUndo = !undoStack.empty();
    state.canRedo = !redoStack.empty();

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

    for (int i = 0; i < columnCount; i++) {
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
        undo();
        break;
    case ItemAction::Redo:
        redo();
        break;
    case ItemAction::Search:
        performSearchInList();
        break;
    case ItemAction::Cut:
        copySelectedItemsToClipboard();
        deleteSelectedLines();
        break;
    case ItemAction::Copy:
        copySelectedItemsToClipboard();
        break;
    case ItemAction::Paste:
        pasteItemsIntoList();
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
        }
        break;
    case ItemAction::Delete: {
        int selectedCount = ListView_GetSelectedCount(_replaceListView);
        std::wstring confirmationMessage;

        if (selectedCount == 1) {
            confirmationMessage = getLangStr(L"msgbox_confirm_delete_single");
        }
        else if (selectedCount > 1) {
            confirmationMessage = getLangStr(L"msgbox_confirm_delete_multiple", { std::to_wstring(selectedCount) });
        }

        int msgBoxID = MessageBox(nppData._nppHandle, confirmationMessage.c_str(), getLangStr(L"msgbox_title_confirm").c_str(), MB_ICONWARNING | MB_YESNO);
        if (msgBoxID == IDYES) {
            deleteSelectedLines();
        }
        break;
    }
    case ItemAction::Add: {
        int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
        if (insertPosition != -1) {
            insertPosition++;
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

        break;
    }
    }
}

void MultiReplace::copySelectedItemsToClipboard() {
    std::wstring csvData;
    int itemCount = ListView_GetItemCount(_replaceListView);
    int selectedCount = ListView_GetSelectedCount(_replaceListView);

    if (selectedCount > 0) {
        for (int i = 0; i < itemCount; ++i) {
            if (ListView_GetItemState(_replaceListView, i, LVIS_SELECTED) & LVIS_SELECTED) {
                // Assuming the index in the ListView corresponds to the index in replaceListData
                const ReplaceItemData& item = replaceListData[i];
                std::wstring line = std::to_wstring(item.isEnabled) + L"," +
                    escapeCsvValue(item.findText) + L"," +
                    escapeCsvValue(item.replaceText) + L"," +
                    std::to_wstring(item.wholeWord) + L"," +
                    std::to_wstring(item.matchCase) + L"," +
                    std::to_wstring(item.useVariables) + L"," +
                    std::to_wstring(item.extended) + L"," +
                    std::to_wstring(item.regex) + L"," +
                    escapeCsvValue(item.comments) + L"\n";  // Include Comments column
                csvData += line;
            }
        }
    }

    if (!csvData.empty()) {
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            HGLOBAL hClipboardData = GlobalAlloc(GMEM_DDESHARE, (csvData.size() + 1) * sizeof(wchar_t));
            if (hClipboardData) {
                wchar_t* pClipboardData = static_cast<wchar_t*>(GlobalLock(hClipboardData));
                if (pClipboardData != NULL) {
                    memcpy(pClipboardData, csvData.c_str(), (csvData.size() + 1) * sizeof(wchar_t));
                    GlobalUnlock(hClipboardData);
                    SetClipboardData(CF_UNICODETEXT, hClipboardData);
                }
                else {
                    GlobalFree(hClipboardData);
                }
                CloseClipboard();
            }
        }
    }
}

bool MultiReplace::canPasteFromClipboard() {
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) {
        return false; // Clipboard does not contain text in a format that we can paste
    }

    if (!OpenClipboard(nullptr)) {
        return false; // Cannot open the clipboard
    }

    bool canPaste = false; // Assume we cannot paste until proven otherwise
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData) {
        wchar_t* pClipboardData = static_cast<wchar_t*>(GlobalLock(hData));
        if (pClipboardData) {
            std::wstring content = pClipboardData;
            std::wistringstream contentStream(content);
            std::wstring line;

            while (std::getline(contentStream, line) && !canPaste) {
                if (line.empty()) continue; // Skip empty lines

                std::vector<std::wstring> columns = parseCsvLine(line);

                // For the format to be considered valid, ensure each line has the correct number of columns
                if ((columns.size() == 8 || columns.size() == 9) ) {
                    canPaste = true; // Found at least one valid line, no need to check further
                }
            }

            GlobalUnlock(hData);
        }
    }

    CloseClipboard();
    return canPaste;
}

void MultiReplace::pasteItemsIntoList() {
    if (!OpenClipboard(NULL)) {
        return; // Abort if the clipboard cannot be opened
    }

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (!hData) {
        CloseClipboard(); // Close the clipboard if no data is present
        return;
    }

    wchar_t* pClipboardData = static_cast<wchar_t*>(GlobalLock(hData));
    if (!pClipboardData) {
        CloseClipboard(); // Close the clipboard if the data cannot be locked
        return;
    }

    std::wstring content = pClipboardData;
    GlobalUnlock(hData);
    CloseClipboard();

    std::wstringstream contentStream(content);
    std::wstring line;
    std::vector<ReplaceItemData> itemsToInsert; // Collect items to insert

    // Determine insert position based on focused item or default to end if none is focused
    int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
    if (insertPosition != -1) {
        // Increase by one to insert after the focused item
        insertPosition++;
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
            item.isEnabled = std::stoi(columns[0]) != 0;
            item.findText = columns[1];
            item.replaceText = columns[2];
            item.wholeWord = std::stoi(columns[3]) != 0;
            item.matchCase = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended = std::stoi(columns[6]) != 0;
            item.regex = std::stoi(columns[7]) != 0;
            // Handle Comments column if present
            if (columns.size() == 9) {
                item.comments = columns[8];
            }
            else {
                item.comments = L"";
            }
        }
        catch (const std::exception&) {
            continue; // Silently ignore lines with conversion errors
        }

        itemsToInsert.push_back(item);
    }

    if (itemsToInsert.empty()) {
        // No items were collected, so nothing to insert
        return;
    }

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
        showStatusMessage(getLangStr(L"status_no_find_replace_list_input"), COLOR_ERROR);
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
        showStatusMessage(getLangStr(L"status_found_in_list"), COLOR_SUCCESS);
    }
    else {
        // Show failure status message if no match found
        showStatusMessage(getLangStr(L"status_not_found_in_list"), COLOR_ERROR);
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

        i++;
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
        loadLanguage();
        initializeWindowSize();
        pointerToScintilla();
        initializeMarkerStyle();
        initializeCtrlMap();
        initializeFontStyles();
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

#pragma warning(push)
#pragma warning(disable:26454)  // Suppress the overflow warning due to BCN_DROPDOWN and NM_RDBLCLK definition

        if (pnmh->code == BCN_DROPDOWN && pnmh->hwndFrom == GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON))
        {
            RECT rc;
            ::GetWindowRect(pnmh->hwndFrom, &rc);  // Get screen coordinates of the button

            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_ALL_OPTION, getLangStrLPWSTR(L"split_menu_replace_all"));
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_IN_ALL_DOCS_OPTION, getLangStrLPWSTR(L"split_menu_replace_all_in_docs"));

            // Display the menu directly below the button
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, rc.left, rc.bottom, 0, _hSelf, NULL);
            DestroyMenu(hMenu);  // Clean up after displaying menu

            return TRUE;
        }

#pragma warning(pop)  // Restore the original warning settings

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
                int itemIndex = plvdi->item.iItem;
                int subItem = plvdi->item.iSubItem;

                // Check if the item index is valid
                if (itemIndex >= 0 && itemIndex < static_cast<int>(replaceListData.size())) {
                    ReplaceItemData& itemData = replaceListData[itemIndex];

                    // Display the data based on the subitem
                    if (columnIndices[ColumnID::FIND_COUNT] != -1 && subItem == columnIndices[ColumnID::FIND_COUNT]) {
                        plvdi->item.pszText = const_cast<LPWSTR>(itemData.findCount.c_str());
                    }
                    else if (columnIndices[ColumnID::REPLACE_COUNT] != -1 && subItem == columnIndices[ColumnID::REPLACE_COUNT]) {
                        plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceCount.c_str());
                    }
                    else if (columnIndices[ColumnID::SELECTION] != -1 && subItem == columnIndices[ColumnID::SELECTION]) {
                        plvdi->item.pszText = itemData.isEnabled ? L"\u25A0" : L"\u2610";  // Square or checkbox
                    }
                    else if (columnIndices[ColumnID::FIND_TEXT] != -1 && subItem == columnIndices[ColumnID::FIND_TEXT]) {
                        plvdi->item.pszText = const_cast<LPWSTR>(itemData.findText.c_str());
                    }
                    else if (columnIndices[ColumnID::REPLACE_TEXT] != -1 && subItem == columnIndices[ColumnID::REPLACE_TEXT]) {
                        plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceText.c_str());
                    }
                    else if (columnIndices[ColumnID::WHOLE_WORD] != -1 && subItem == columnIndices[ColumnID::WHOLE_WORD]) {
                        plvdi->item.pszText = itemData.wholeWord ? L"\u2714" : L"";
                    }
                    else if (columnIndices[ColumnID::MATCH_CASE] != -1 && subItem == columnIndices[ColumnID::MATCH_CASE]) {
                        plvdi->item.pszText = itemData.matchCase ? L"\u2714" : L"";
                    }
                    else if (columnIndices[ColumnID::USE_VARIABLES] != -1 && subItem == columnIndices[ColumnID::USE_VARIABLES]) {
                        plvdi->item.pszText = itemData.useVariables ? L"\u2714" : L"";
                    }
                    else if (columnIndices[ColumnID::EXTENDED] != -1 && subItem == columnIndices[ColumnID::EXTENDED]) {
                        plvdi->item.pszText = itemData.extended ? L"\u2714" : L"";
                    }
                    else if (columnIndices[ColumnID::REGEX] != -1 && subItem == columnIndices[ColumnID::REGEX]) {
                        plvdi->item.pszText = itemData.regex ? L"\u2714" : L"";
                    }
                    else if (columnIndices[ColumnID::COMMENTS] != -1 && subItem == columnIndices[ColumnID::COMMENTS]) {
                        plvdi->item.pszText = const_cast<LPWSTR>(itemData.comments.c_str());
                    }
                    else if (columnIndices[ColumnID::DELETE_BUTTON] != -1 && subItem == columnIndices[ColumnID::DELETE_BUTTON]) {
                        plvdi->item.pszText = L"\u2716";  // Cross mark for delete
                    }
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

                PostMessage(_replaceListView, WM_SETFOCUS, 0, 0);
                // Handling keyboard shortcuts for menu actions
                if (GetKeyState(VK_CONTROL) & 0x8000) { // If Ctrl is pressed
                    switch (pnkd->wVKey) {
                    case 'Z': // Ctrl+Z for Undo
                        undo();
                        break;
                    case 'Y': // Ctrl+Y for Redo
                        redo();
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
                        break;
                    case 'I': // Ctrl+I for Adding new Line
                        performItemAction(_contextMenuClickPoint, ItemAction::Add);
                        break;
                    }
                }
                else if (GetKeyState(VK_MENU) & 0x8000) { // If Alt is pressed
                    switch (pnkd->wVKey) {
                    case 'E': // Alt+E for Enable Line
                        setSelections(true, ListView_GetSelectedCount(_replaceListView) > 0);
                        break;
                    case 'D': // Alt+D for Disable Line
                        setSelections(false, ListView_GetSelectedCount(_replaceListView) > 0);
                        break;
                    case VK_UP: // Alt+ UP for Push Back
                        iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                        if (iItem >= 0) {
                            NMITEMACTIVATE nmia;
                            ZeroMemory(&nmia, sizeof(nmia));
                            nmia.iItem = iItem;
                            handleCopyBack(&nmia);
                        }
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

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
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
            if (!isColumnHighlighted) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleHighlightColumnsInDocument();
                }
            }
            else {
                handleClearColumnMarks();
                showStatusMessage(getLangStr(L"status_column_marks_cleared"), COLOR_SUCCESS);
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

        case IDC_FIND_BUTTON:
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
            if (isReplaceAllInDocs)
            {
                CloseDebugWindow(); // Close the Lua debug window if it is open
                int msgboxID = MessageBox(
                    nppData._nppHandle,
                    getLangStr(L"msgbox_confirm_replace_all").c_str(),
                    getLangStr(L"msgbox_title_confirm").c_str(),
                    MB_ICONWARNING | MB_OKCANCEL
                );

                if (msgboxID == IDOK)
                {
                    // Reset Count Columns once before processing multiple documents
                    resetCountColumns();

                    // Get the number of opened documents in each view
                    LRESULT docCountMain = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, PRIMARY_VIEW);
                    LRESULT docCountSecondary = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, SECOND_VIEW);

                    // Check the visibility of each view
                    bool visibleMain = IsWindowVisible(nppData._scintillaMainHandle);
                    bool visibleSecond = IsWindowVisible(nppData._scintillaSecondHandle);

                    // Save focused Document
                    LRESULT currentDocIndex = ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTDOCINDEX, 0, MAIN_VIEW);

                    // Process documents in the main view if it's visible
                    if (visibleMain) {
                        for (LRESULT i = 0; i < docCountMain; ++i) {
                            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, MAIN_VIEW, i);
                            handleDelimiterPositions(DelimiterOperation::LoadAll);
                            handleReplaceAllButton();
                        }
                    }

                    // Process documents in the secondary view if it's visible
                    if (visibleSecond) {
                        for (LRESULT i = 0; i < docCountSecondary; ++i) {
                            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, SUB_VIEW, i);
                            handleDelimiterPositions(DelimiterOperation::LoadAll);
                            handleReplaceAllButton();
                        }
                    }

                    // Restore opened Document
                    ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, visibleMain ? MAIN_VIEW : SUB_VIEW, currentDocIndex);
                }
            }
            else
            {
                CloseDebugWindow(); // Close the Lua debug window if it is open
                resetCountColumns();
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                handleReplaceAllButton();
            }
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
            showStatusMessage(getLangStr(L"status_all_marks_cleared"), COLOR_SUCCESS);
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

            std::wstring csvDescription = getLangStr(L"filetype_csv");  // "CSV Files (*.csv)"
            std::wstring allFilesDescription = getLangStr(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {csvDescription, L"*.csv"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = getLangStr(L"panel_load_list");
            std::wstring filePath = openFileDialog(false, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST, L"csv", L"");

            if (!filePath.empty()) {
                loadListFromCsv(filePath);
            }
            return TRUE;
        }

        case IDC_NEW_LIST_BUTTON:
        {
            clearList();
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
            std::wstring bashDescription = getLangStr(L"filetype_bash");  // "Bash Scripts (*.sh)"
            std::wstring allFilesDescription = getLangStr(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {bashDescription, L"*.sh"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = getLangStr(L"panel_export_to_bash");

            // Set a default filename if none is provided
            static int scriptCounter = 1;
            std::wstring defaultFileName = L"Replace_Script_" + std::to_wstring(scriptCounter++) + L".sh";

            // Open the file dialog with the default filename for bash scripts
            std::wstring filePath = openFileDialog(true, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"sh", defaultFileName);

            if (!filePath.empty()) {
                exportToBashScript(filePath);
            }
            return TRUE;
        }

        case ID_REPLACE_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, getLangStrLPWSTR(L"split_button_replace_all"));
            isReplaceAllInDocs = false;
            return TRUE;
        }

        case ID_REPLACE_IN_ALL_DOCS_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, getLangStrLPWSTR(L"split_button_replace_all_in_docs"));
            isReplaceAllInDocs = true;
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

void MultiReplace::handleReplaceAllButton() {

    // First check if the document is read-only
    LRESULT isReadOnly = ::SendMessage(_hScintilla, SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(getLangStr(L"status_cannot_replace_read_only"), COLOR_ERROR);
        return;
    }

    globalLuaVariablesMap.clear(); // Clear all stored Lua Global Variables
    
    hashTablesMap.clear(); // Clear all stored Lua Hash Tables

    int totalReplaceCount = 0;

    if (useListEnabled)
    {
        // Check if the replaceListData is empty and warn the user if so
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_instructions"), COLOR_ERROR);
            return;
        }

        // Check status for initial call if stopped by DEBUG, don't highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/false)) {
            return;
        }

        ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
        for (size_t i = 0; i < replaceListData.size(); ++i)
        {
            if (replaceListData[i].isEnabled)
            {
                int findCount = 0;
                int replaceCount = 0;

                // Call replaceAll and break out if there is an error or a Debug Stop
                bool success = replaceAll(replaceListData[i], findCount, replaceCount, i);

                // Refresh ListView to show updated statistics
                refreshUIListView();

                // Accumulate total replacements
                totalReplaceCount += replaceCount;

                if (!success) {
                    break;  // Exit loop on error or Debug Stop
                }
            }
        }
        ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
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

        ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
        int findCount = 0;
        replaceAll(itemData, findCount, totalReplaceCount);
        ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText);
    }
    // Display status message
    showStatusMessage(getLangStr(L"status_occurrences_replaced", { std::to_wstring(totalReplaceCount) }), COLOR_SUCCESS);

    // Disable selection radio and switch to "All Text" if it was Replaced an none selection left or it will be trapped
    SelectionInfo selection = getSelectionInfo(false);
    if (selection.length == 0 && IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
    }
}

void MultiReplace::handleReplaceButton() {

    // First check if the document is read-only
    LRESULT isReadOnly = ::SendMessage(_hScintilla, SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(getLangStrLPWSTR(L"status_cannot_replace_read_only"), COLOR_ERROR);
        return;
    }
    
    globalLuaVariablesMap.clear(); // Clear all stored Lua Global Variables
   
    hashTablesMap.clear();   // Clear all stored Lua Hash Tables

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SearchResult searchResult;
    searchResult.pos = -1;
    searchResult.length = 0;
    searchResult.foundText = "";

    SelectionInfo selection = getSelectionInfo(false);

    // If there is a selection, set newPos to the start of the selection; otherwise, use the current cursor position
    Sci_Position newPos = (selection.length > 0) ? selection.startPos : ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);

    size_t matchIndex = std::numeric_limits<size_t>::max();

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_uncheck"), COLOR_ERROR);
            return;
        }

        // Check status of initial if stopped by DEBUG, highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/true)) {
            return;
        }

        int replacements = 0;  // Counter for replacements
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled && replaceOne(replaceListData[i], selection, searchResult, newPos, i)) {
                replacements++;
                refreshUIListView(); // Refresh the ListView to show updated statistic
            }
        }

        searchResult = performListSearchForward(replaceListData, newPos, matchIndex);

        if (searchResult.pos < 0 && wrapAroundEnabled) {
            searchResult = performListSearchForward(replaceListData, 0, matchIndex);
        }

        // Build and show message based on results
        if (replacements > 0) {
            if (searchResult.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView(); // Refresh the ListView to show updated statistic
                selectListItem(matchIndex); // Highlight the matched item in the list
                showStatusMessage(getLangStr(L"status_replace_next_found", { std::to_wstring(replacements) }), COLOR_INFO);
            }
            else {
                showStatusMessage(getLangStr(L"status_replace_none_left", { std::to_wstring(replacements) }), COLOR_INFO);
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(getLangStr(L"status_no_occurrence_found"), COLOR_ERROR, true);
            }
            else {
                updateCountColumns(matchIndex, 1);
                refreshUIListView(); // Refresh the ListView to show updated statistic
                selectListItem(matchIndex); // Highlight the matched item in the list
                showStatusMessage(getLangStr(L"status_found_text_not_replaced"), COLOR_INFO);
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

        std::string findTextUtf8 = convertAndExtend(replaceItem.findText, replaceItem.extended);
        int searchFlags = (replaceItem.wholeWord * SCFIND_WHOLEWORD) | (replaceItem.matchCase * SCFIND_MATCHCASE) | (replaceItem.regex * SCFIND_REGEXP);

        bool wasReplaced = replaceOne(replaceItem, selection, searchResult, newPos);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), replaceItem.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceItem.replaceText);

        if (searchResult.pos < 0 && wrapAroundEnabled) {
            searchResult = performSearchForward(findTextUtf8, searchFlags, true, 0);
        }
        else if (searchResult.pos >= 0) {
            searchResult = performSearchForward(findTextUtf8, searchFlags, true, newPos);
        }

        if (wasReplaced) {
            if (searchResult.pos >= 0) {
                showStatusMessage(getLangStr(L"status_replace_one_next_found"), COLOR_SUCCESS);
            }
            else {
                showStatusMessage(getLangStr(L"status_replace_one_none_left"), COLOR_INFO);
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(getLangStr(L"status_no_occurrence_found"), COLOR_ERROR, true);
            }
            else {
                showStatusMessage(getLangStr(L"status_found_text_not_replaced"), COLOR_INFO);
            }
        }
    }

    // Disable selection radio and switch to "All Text" if it was Replaced an none selection left or it will be trapped
    selection = getSelectionInfo(false);
    if (selection.length == 0 && IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
        ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
    }
}

bool MultiReplace::replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos, size_t itemIndex)
{
    std::string findTextUtf8 = convertAndExtend(itemData.findText, itemData.extended);
    int searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
    searchResult = performSearchForward(findTextUtf8, searchFlags, false, selection.startPos);

    if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
        bool skipReplace = false;
        std::string replaceTextUtf8 = convertAndExtend(itemData.replaceText, itemData.extended);
        std::string localReplaceTextUtf8 = wstringToString(itemData.replaceText);

        if (itemIndex != SIZE_MAX) {
            updateCountColumns(itemIndex, 1); // no refreshUIListView() necessary as implemented in Debug Window
            selectListItem(itemIndex);
        }

        if (itemData.useVariables) {
            LuaVariables vars;

            int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
            int previousLineStartPosition = (currentLineIndex == 0) ? 0 : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));

            // Setting FNAME and FPATH
            setLuaFileVars(vars);

            if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED) {
                ColumnInfo columnInfo = getColumnInfo(searchResult.pos);
                vars.COL = static_cast<int>(columnInfo.startColumnIndex);
            }

            vars.CNT = 1;
            vars.LCNT = 1;
            vars.APOS = static_cast<int>(searchResult.pos) + 1;
            vars.LINE = currentLineIndex + 1;
            vars.LPOS = static_cast<int>(searchResult.pos) - previousLineStartPosition + 1;
            vars.MATCH = searchResult.foundText;

            if (!resolveLuaSyntax(localReplaceTextUtf8, vars, skipReplace, itemData.regex)) {
                return false;  // Exit the function if error in syntax
            }
            replaceTextUtf8 = convertAndExtend(localReplaceTextUtf8, itemData.extended);
        }

        if (!skipReplace) {
            if (itemData.regex) {
                newPos = performRegexReplace(replaceTextUtf8, searchResult.pos, searchResult.length);
            }
            else {
                newPos = performReplace(replaceTextUtf8, searchResult.pos, searchResult.length);
            }

            if (itemIndex != SIZE_MAX) {
                updateCountColumns(itemIndex, -1, 1);
            }

            return true;  // A replacement was made
        }
        else {
            newPos = searchResult.pos + searchResult.length;
            // Clear selection
            send(SCI_SETSELECTIONSTART, newPos, 0);
            send(SCI_SETSELECTIONEND, newPos, 0);
        }
    }
    return false;  // No replacement was made
}

bool MultiReplace::replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount, size_t itemIndex)
{
    if (itemData.findText.empty()) {
        findCount = 0;
        replaceCount = 0;
        return true;
    }

    bool isReplaceFirstEnabled = (IsDlgButtonChecked(_hSelf, IDC_REPLACE_FIRST_CHECKBOX) == BST_CHECKED);
    int searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);

    std::string findTextUtf8 = convertAndExtend(itemData.findText, itemData.extended);
    std::string replaceTextUtf8 = convertAndExtend(itemData.replaceText, itemData.extended);

    int previousLineIndex = -1;
    int lineFindCount = 0;

    SearchResult searchResult = performSearchForward(findTextUtf8, searchFlags, false, 0);

    while (searchResult.pos >= 0)
    {
        bool skipReplace = false;
        findCount++;
        std::string localReplaceTextUtf8 = wstringToString(itemData.replaceText);

        if (itemIndex != SIZE_MAX) {  // check if used in List
            updateCountColumns(itemIndex, findCount);
        }

        if (itemData.useVariables) {
            LuaVariables vars;

            if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED) {
                ColumnInfo columnInfo = getColumnInfo(searchResult.pos);
                vars.COL = static_cast<int>(columnInfo.startColumnIndex);
            }

            int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
            int previousLineStartPosition = (currentLineIndex == 0) ? 0 : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));

            // Reset lineReplaceCount if the line has changed
            if (currentLineIndex != previousLineIndex) {
                lineFindCount = 0;
                previousLineIndex = currentLineIndex;
            }

            // Setting FNAME and FPATH
            setLuaFileVars(vars);

            if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED) {
                ColumnInfo columnInfo = getColumnInfo(searchResult.pos);
                vars.COL = static_cast<int>(columnInfo.startColumnIndex);
            }

            lineFindCount++;

            vars.CNT = findCount;
            vars.LCNT = lineFindCount;
            vars.APOS = static_cast<int>(searchResult.pos) + 1;
            vars.LINE = currentLineIndex + 1;
            vars.LPOS = static_cast<int>(searchResult.pos) - previousLineStartPosition + 1;
            vars.MATCH = searchResult.foundText;

            if (!resolveLuaSyntax(localReplaceTextUtf8, vars, skipReplace, itemData.regex)) {
                return false;  // Exit the loop if error in syntax or process is stopped by debug
            }
            replaceTextUtf8 = convertAndExtend(localReplaceTextUtf8, itemData.extended);
        }

        Sci_Position newPos;
        if (!skipReplace) {
            if (itemData.regex) {
                newPos = performRegexReplace(replaceTextUtf8, searchResult.pos, searchResult.length);
            }
            else {
                newPos = performReplace(replaceTextUtf8, searchResult.pos, searchResult.length);
            }
            replaceCount++;

            if (itemIndex != SIZE_MAX) { // check if used in List
                updateCountColumns(itemIndex, -1, replaceCount);
            }
        }
        else {
            newPos = searchResult.pos + searchResult.length;
            // Clear selection
            send(SCI_SETSELECTIONSTART, newPos, 0);
            send(SCI_SETSELECTIONEND, newPos, 0);
        }

        if (isReplaceFirstEnabled) {
            break;  // Exit the loop after the first successful replacement
        }

        searchResult = performSearchForward(findTextUtf8, searchFlags, false, newPos);
    }

    return true;
}

Sci_Position MultiReplace::performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    // Set the target range for the replacement
    send(SCI_SETTARGETRANGE, pos, pos + length);

    // Get the codepage of the document
    int cp = static_cast<int>(send(SCI_GETCODEPAGE, 0, 0));

    // Convert the string from UTF-8 to the codepage of the document
    std::string replaceTextCp = utf8ToCodepage(replaceTextUtf8, cp);

    // Perform the replacement
    send(SCI_REPLACETARGET, replaceTextCp.size(), reinterpret_cast<sptr_t>(replaceTextCp.c_str()));

    // Get the end position after the replacement
    Sci_Position newTargetEnd = static_cast<Sci_Position>(send(SCI_GETTARGETEND, 0, 0));

    // Set the cursor to the end of the replaced text
    //send(SCI_SETCURRENTPOS, newTargetEnd, 0);

    // Clear selection
    //send(SCI_SETSELECTIONSTART, newTargetEnd, 0);
    //send(SCI_SETSELECTIONEND, newTargetEnd, 0);

    return newTargetEnd;
}

Sci_Position MultiReplace::performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    // Set the target range for the replacement
    send(SCI_SETTARGETRANGE, pos, pos + length);

    // Get the codepage of the document
    int cp = static_cast<int>(send(SCI_GETCODEPAGE, 0, 0));

    // Convert the string from UTF-8 to the codepage of the document
    std::string replaceTextCp = utf8ToCodepage(replaceTextUtf8, cp);

    // Perform the regex replacement
    send(SCI_REPLACETARGETRE, static_cast<WPARAM>(-1), reinterpret_cast<sptr_t>(replaceTextCp.c_str()));

    // Get the end position after the replacement
    Sci_Position newTargetEnd = static_cast<Sci_Position>(send(SCI_GETTARGETEND, 0, 0));

    // Set the cursor to the end of the replaced text
    //send(SCI_SETCURRENTPOS, newTargetEnd, 0);

    // Clear selection
    //send(SCI_SETSELECTIONSTART, newTargetEnd, 0);
    //send(SCI_SETSELECTIONEND, newTargetEnd, 0);

    return newTargetEnd;
}

bool MultiReplace::preProcessListForReplace(bool highlight) {
    if (!replaceListData.empty()) {
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled && replaceListData[i].useVariables) {
                if (replaceListData[i].findText.empty()) {
                    if (highlight) {
                        selectListItem(i);  // Highlight the list entry
                    }
                    std::string localReplaceTextUtf8 = wstringToString(replaceListData[i].replaceText);
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
    LRESULT selectionCount = ::SendMessage(_hScintilla, SCI_GETSELECTIONS, 0, 0);
    Sci_Position selectedStart = 0;
    Sci_Position correspondingEnd = 0;
    std::vector<SelectionRange> selections; // Store all selections for sorting

    if (selectionCount > 0) {
        selections.resize(selectionCount);

        // Retrieve all selections
        for (LRESULT i = 0; i < selectionCount; ++i) {
            selections[i].start = ::SendMessage(_hScintilla, SCI_GETSELECTIONNSTART, i, 0);
            selections[i].end = ::SendMessage(_hScintilla, SCI_GETSELECTIONNEND, i, 0);
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
        selectedStart = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
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

void MultiReplace::captureHashTables(lua_State* L) {

    // Get hashTables from Lua
    lua_getglobal(L, "hashTables");
    if (!lua_istable(L, -1)) {
        lua_pop(L, 1);
        return; // Not a valid table, do nothing
    }

    // Iterate over hashTables[hpath]
    lua_pushnil(L);
    while (lua_next(L, -2) != 0) {
        // key = hpath (string), value = hpathTable (table)
        if (lua_isstring(L, -2) && lua_istable(L, -1)) {
            std::string hpath = lua_tostring(L, -2);
            LuaHashTable hTable;
            hTable.hpath = hpath;

            lua_pushnil(L);
            while (lua_next(L, -2) != 0) {
                // key = entry key, value = entry value (both expected to be strings)
                if (lua_isstring(L, -2) && lua_isstring(L, -1)) {
                    std::string k = lua_tostring(L, -2);
                    std::string v = lua_tostring(L, -1);
                    hTable.entries[k] = v;
                }
                lua_pop(L, 1); // pop value, keep key for next iteration
            }

            hashTablesMap[hpath] = hTable;
        }
        lua_pop(L, 1); // pop hpathTable
    }

    lua_pop(L, 1); // pop hashTables table
}

void MultiReplace::loadLuaGlobals(lua_State* L) {
    for (const auto& pair : globalLuaVariablesMap) {
        const LuaVariable& var = pair.second;

        switch (var.type) {
        case LuaVariableType::String:
            lua_pushstring(L, var.stringValue.c_str());
            break;
        case LuaVariableType::Number:
            lua_pushnumber(L, var.numberValue);
            break;
        case LuaVariableType::Boolean:
            lua_pushboolean(L, var.booleanValue);
            break;
        default:
            continue;  // Skip None or unsupported types
        }

        lua_setglobal(L, var.name.c_str());
    }
}

void MultiReplace::loadHashTables(lua_State* L) {
    // Create a new table for hashTables in Lua
    lua_newtable(L); // stack: {..., hashTablesTable}

    // Iterate over all stored hash tables in hashTablesMap
    for (const auto& pair : hashTablesMap) {
        const std::string& hpath = pair.first;
        const LuaHashTable& hTable = pair.second;

        // Create a sub-table for this particular hash table
        lua_newtable(L); // stack: {..., hashTablesTable, hpathTable}

        // Populate the hpathTable with key-value pairs
        for (const auto& kv : hTable.entries) {
            lua_pushstring(L, kv.first.c_str());  // push key
            lua_pushstring(L, kv.second.c_str()); // push value
            lua_settable(L, -3); // hpathTable[key] = value
        }

        // Assign hpathTable to hashTables[hpath]
        lua_pushstring(L, hpath.c_str());
        lua_insert(L, -2); // move hpathTable below hpath key
        lua_settable(L, -3); // hashTables[hpath] = hpathTable
    }

    // Set hashTables as a global variable
    lua_setglobal(L, "hashTables");
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

bool MultiReplace::resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex)
{
    lua_State* L = luaL_newstate();  // Create a new Lua environment
    luaL_openlibs(L);  // Load standard libraries

    loadLuaGlobals(L); // Load global Lua variables
    loadHashTables(L);

    // Set variables
    lua_pushnumber(L, vars.CNT);
    lua_setglobal(L, "CNT");
    lua_pushnumber(L, vars.LCNT);
    lua_setglobal(L, "LCNT");
    lua_pushnumber(L, vars.LINE);
    lua_setglobal(L, "LINE");
    lua_pushnumber(L, vars.LPOS);
    lua_setglobal(L, "LPOS");
    lua_pushnumber(L, vars.APOS);
    lua_setglobal(L, "APOS");
    lua_pushnumber(L, vars.COL);
    lua_setglobal(L, "COL");
    lua_pushboolean(L, regex);
    lua_setglobal(L, "REGEX");

    // Convert numbers to integers
    luaL_dostring(L, "CNT = math.tointeger(CNT)");
    luaL_dostring(L, "LCNT = math.tointeger(LCNT)");
    luaL_dostring(L, "LINE = math.tointeger(LINE)");
    luaL_dostring(L, "LPOS = math.tointeger(LPOS)");
    luaL_dostring(L, "APOS = math.tointeger(APOS)");
    luaL_dostring(L, "COL = math.tointeger(COL)");

    setLuaVariable(L, "FPATH", vars.FPATH);
    setLuaVariable(L, "FNAME", vars.FNAME);
    setLuaVariable(L, "MATCH", vars.MATCH);
    // Get CAPs from Scintilla using SCI_GETTAG
    std::vector<std::string> capNames;  // Vector to store the names of the set CAP variables for DEBUG Window

    if (regex) {
        sptr_t len = 0;
        int capIndex = 1;  // Start numbering from CAP1

        // Loop to retrieve up to MAX_CAP_GROUPS capture groups
        for (int i = 1; i <= MAX_CAP_GROUPS; ++i) {
            // First, get the length of the capture group
            len = send(SCI_GETTAG, i, 0, true);

            if (len < 0) {
                // No more captures are available
                break;
            }

            if (len == 0) {
                // Empty capture, Ensure the index is correctly incremented even for empty captures
                capIndex++;
                continue;
            }

            // Allocate a buffer to hold the capture value
            std::vector<char> buffer(len + 1);  // +1 for null terminator

            // Now, retrieve the capture group
            sptr_t result = send(SCI_GETTAG, i, reinterpret_cast<sptr_t>(buffer.data()), false);

            if (result < 0) {
                // Error occurred, clean up and return false
                lua_close(L);
                return false;
            }

            // Null-terminate the string
            buffer[len] = '\0';

            std::string capValue(buffer.data());  // Convert to std::string

            // Generate the CAP variable name dynamically
            std::string globalVarName = "CAP" + std::to_string(capIndex);

            // Set the capture as a global Lua variable
            setLuaVariable(L, globalVarName, capValue);

            // Store the CAP variable name for later use
            capNames.push_back(globalVarName);

            capIndex++;
        }

    }

    // Declare cond statement function
    luaL_dostring(L,
        "function cond(cond, trueVal, falseVal)\n"
        "  local res = {result = '', skip = false}  -- Initialize result table with defaults\n"
        "  if cond == nil then  -- Check if cond is nil\n"
        "    error('cond: condition cannot be nil')\n"
        "    return res\n"
        "  end\n"

        "  if cond and trueVal == nil then  -- Check if trueVal is nil\n"
        "    error('cond: trueVal cannot be nil')\n"
        "    return res\n"
        "  end\n"

        "  if not cond and falseVal == nil then  -- Check if falseVal is missing\n"
        "    res.skip = true  -- Set skip to true ONLY IF falseVal is not provided\n"
        "  end\n"

        "  if type(trueVal) == 'function' then\n"
        "    trueVal = trueVal()\n"
        "  end\n"

        "  if type(falseVal) == 'function' then\n"
        "    falseVal = falseVal()\n"
        "  end\n"

        "  if cond then\n"
        "    if type(trueVal) == 'table' then\n"
        "      res.result = trueVal.result\n"
        "      res.skip = trueVal.skip\n"
        "    else\n"
        "      res.result = trueVal\n"
        "      res.skip = false\n"
        "    end\n"
        "  else\n"
        "    if not res.skip then\n"
        "      if type(falseVal) == 'table' then\n"
        "        res.result = falseVal.result\n"
        "        res.skip = falseVal.skip\n"
        "      else\n"
        "        res.result = falseVal\n"
        "      end\n"
        "    end\n"
        "  end\n"
        "  resultTable = res\n"
        "  return res  -- Return the table containing result and skip\n"
        "end\n");

    // Declare the set function
    luaL_dostring(L,
        "function set(strOrCalc)\n"
        "  local res = {result = '', skip = false}  -- Initialize result table with defaults\n"
        "  if strOrCalc == nil then\n"
        "    error('set: cannot be nil')\n"
        "    return\n"
        "  end\n"
        "  if type(strOrCalc) == 'string' then\n"
        "    res.result = strOrCalc  -- Setting res.result\n"
        "  elseif type(strOrCalc) == 'number' then\n"
        "    res.result = tostring(strOrCalc)  -- Convert number to string and set to res.result\n"
        "  else\n"
        "    error('set: Expected string or number')\n"
        "    return\n"
        "  end\n"
        "  resultTable = res\n"
        "  return res  -- Return the table containing result and skip\n"
        "end\n");

    // Declare formatNumber function
    luaL_dostring(L,
        "function fmtN(num, maxDecimals, fixedDecimals)\n"
        "  if num == nil then\n"
        "    error('fmtN: num cannot be nil')\n"
        "    return\n"
        "  elseif type(num) ~= 'number' then\n"
        "    error('fmtN: Invalid type for num. Expected a number')\n"
        "    return\n"
        "  end\n"
        "  if maxDecimals == nil then\n"
        "    error('fmtN: maxDecimals cannot be nil')\n"
        "    return\n"
        "  elseif type(maxDecimals) ~= 'number' then\n"
        "    error('fmtN: Invalid type for maxDecimals. Expected a number')\n"
        "    return\n"
        "  end\n"
        "  if fixedDecimals == nil then\n"
        "    error('fmtN: fixedDecimals cannot be nil')\n"
        "    return\n"
        "  elseif type(fixedDecimals) ~= 'boolean' then\n"
        "    error('fmtN: Invalid type for fixedDecimals. Expected a boolean')\n"
        "    return\n"
        "  end\n"
        "  local multiplier = 10 ^ maxDecimals\n"
        "  local rounded = math.floor(num * multiplier + 0.5) / multiplier\n"
        "  local output = ''\n"
        "  if fixedDecimals then\n"
        "    output = string.format('%.' .. maxDecimals .. 'f', rounded)\n"
        "  else\n"
        "    local intPart, fracPart = math.modf(rounded)\n"
        "    if fracPart == 0 then\n"
        "      output = tostring(intPart)\n"
        "    else\n"
        "      output = tostring(rounded)\n"
        "    end\n"
        "  end\n"
        "  return output\n"
        "end\n");

    // Declare the vars function
    luaL_dostring(L,
        "function vars(args)\n"
        "  for name, value in pairs(args) do\n"
        "    -- Set the global variable only if it does not already exist\n"
        "    if _G[name] == nil then\n"
        "      if type(name) ~= 'string' then\n"
        "        error('vars: Variable name must be a string')\n"
        "      end\n"
        "      if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then\n"
        "        error('vars: Invalid variable name')\n"
        "      end\n"
        "      if value == nil then\n"
        "        error('vars: Value missing')\n"
        "      end\n"
        "      -- Check if the value is a string and REGEX is true, then preprocess backslashes\n"
        "      if type(value) == 'string' and REGEX then\n"
        "        value = value:gsub('\\\\', '\\\\\\\\')\n"
        "      end\n"
        "      _G[name] = value\n"
        "    end\n"
        "  end\n"

        "  -- Forward or initialize resultTable\n"
        "  local res = {result = '', skip = true}  -- Defaults for resultTable\n"

        "  -- Ensure resultTable is not overwritten if combined with cond or set\n"
        "  if resultTable == nil then\n"
        "    resultTable = res  -- Initialize resultTable if it does not exist\n"
        "  else\n"
        "    -- Ensure required fields are present without overwriting existing values\n"
        "    if resultTable.result == nil then\n"
        "      resultTable.result = res.result\n"
        "    end\n"
        "    if resultTable.skip == nil then\n"
        "      resultTable.skip = res.skip\n"
        "    end\n"
        "  end\n"

        "  return resultTable  -- Return the existing or new resultTable\n"
        "end\n");

    // 'init' to the existing 'vars' function for compatibility
    luaL_dostring(L,
        "init = vars\n"
    );

    luaL_dostring(L,
        // 1) Helper function: Load file in a sandboxed environment
        "function safeLoadFileSandbox(path)\n"
        "  -- Minimal Environment: Only safe base functions are allowed\n"
        "  local safeEnv = {\n"
        "    pairs = pairs,\n"
        "    ipairs = ipairs,\n"
        "    type = type,\n"
        "    tonumber = tonumber,\n"
        "    tostring = tostring,\n"
        "    table = table,\n"
        "    math = math,\n"
        "    string = string,\n"
        "  }\n"
        "  setmetatable(safeEnv, { __index = function(_, k)\n"
        "    -- Explicitly block access to critical globals like _G, os, io, etc.\n"
        "    if k == '_G' or k == 'os' or k == 'io' or k == 'dofile' or k == 'require' then\n"
        "      return nil  -- Access is denied\n"
        "    end\n"
        "    return _G[k]  -- Allow specific standard functions if needed\n"
        "  end})\n"

        "  -- Lua 5.2+ supports loadfile(path, 't', environment)\n"
        "  local chunk, err = loadfile(path, 't', safeEnv)\n"
        "  if not chunk then\n"
        "    return false, err\n"
        "  end\n"

        "  return pcall(chunk)  -- Execute the code safely\n"
        "end\n"

        // 2) lkp function that uses the sandboxed file loader
        "function lkp(key, hpath, inner)\n"
        "  -- Initialize the result table\n"
        "  local res = { result = '', skip = false }\n"

        "  -- Convert numeric keys to strings\n"
        "  if type(key) == 'number' then\n"
        "    key = tostring(key)\n"
        "  end\n"

        "  -- Validate 'key'\n"
        "  if key == nil then\n"
        "    error('lkp: key passed to file is nil in ' .. tostring(hpath))\n"
        "  end\n"

        "  -- Validate 'hpath'\n"
        "  if hpath == nil or hpath == '' then\n"
        "    error('lkp: file path is invalid or empty')\n"
        "  end\n"

        "  -- Default 'inner' to false\n"
        "  if inner == nil then\n"
        "    inner = false\n"
        "  end\n"

        "  -- Load the hash table if it has not been loaded yet\n"
        "  if hashTables[hpath] == nil then\n"
        "    -- Use the sandboxed file loader to load the file\n"
        "    local success, dataEntries = safeLoadFileSandbox(hpath)\n"
        "    if not success then\n"
        "      error('lkp: failed to safely load file at ' .. tostring(hpath) .. ': ' .. tostring(dataEntries))\n"
        "    end\n"

        "    -- Ensure the file returns a table\n"
        "    if type(dataEntries) ~= 'table' then\n"
        "      error('lkp: invalid format in file at ' .. tostring(hpath))\n"
        "    end\n"

        "    local tbl = {}\n"
        "    -- Iterate through dataEntries, expected to be an array of {keys, value}\n"
        "    for _, entry in ipairs(dataEntries) do\n"
        "      local keys = entry[1]\n"
        "      local value = entry[2]\n"

        "      if value == nil then\n"
        "        goto continue\n"
        "      end\n"

        "      -- Process multiple or single keys\n"
        "      if type(keys) == 'table' then\n"
        "        for _, k in ipairs(keys) do\n"
        "          if type(k) == 'number' then\n"
        "            k = tostring(k)\n"
        "          end\n"
        "          tbl[k] = value\n"
        "        end\n"
        "      elseif type(keys) == 'string' or type(keys) == 'number' then\n"
        "        if type(keys) == 'number' then\n"
        "          keys = tostring(keys)\n"
        "        end\n"
        "        tbl[keys] = value\n"
        "      else\n"
        "        -- Skip unknown key types\n"
        "        goto continue\n"
        "      end\n"

        "      ::continue::\n"
        "    end\n"
        "    hashTables[hpath] = tbl\n"
        "  end\n"

        "  -- Look up the key\n"
        "  local val = hashTables[hpath][key]\n"
        
        "  -- Handle lookup result\n"
        "  if val == nil then\n"
        "    if inner then\n"
        "      res.result = nil\n"
        "    else\n"
        "      res.result = key\n"
        "    end\n"
        "  else\n"
        "    res.result = val\n"
        "  end\n"
        
        "  resultTable = res\n"
        "  return res\n"
        "end\n"
    );

    luaL_dostring(L,
        "function lvars(filePath)\n"
        "  local res = {result = '', skip = true}  -- Default values for resultTable\n"

        "  -- Validate the file path\n"
        "  if filePath == nil or filePath == '' then\n"
        "    error('lvars: file path is invalid or empty')\n"
        "    return res\n"
        "  end\n"

        "  -- Load file in a sandboxed environment\n"
        "  local success, dataTable = safeLoadFileSandbox(filePath)\n"
        "  if not success then\n"
        "    error('lvars: failed to safely load file at ' .. tostring(filePath) .. ': ' .. tostring(dataTable))\n"
        "  end\n"

        "  -- Ensure the loaded data is a table\n"
        "  if type(dataTable) ~= 'table' then\n"
        "    error('lvars: invalid data format in file at ' .. tostring(filePath))\n"
        "  end\n"

        "  -- Set variables in the global environment (_G)\n"
        "  for name, value in pairs(dataTable) do\n"
        "    if type(name) ~= 'string' then\n"
        "      error('lvars: Variable name must be a string. Found invalid key \"' .. tostring(name) .. '\"')\n"
        "    end\n"
        "    if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then\n"
        "      error('lvars: Invalid variable name \"' .. tostring(name) .. '\"')\n"
        "    end\n"
        "    if value == nil then\n"
        "      error('lvars: Value missing for variable \"' .. tostring(name) .. '\"')\n"
        "    end\n"

        "    -- Escape backslashes if REGEX is true and value is a string\n"
        "    if REGEX and type(value) == 'string' then\n"
        "      value = value:gsub('\\\\', '\\\\\\\\')\n"
        "    end\n"

        "    _G[name] = value\n"
        "  end\n"

        "  -- Forward or initialize resultTable\n"
        "  if resultTable == nil then\n"
        "    resultTable = res  -- Initialize resultTable if it does not exist\n"
        "  else\n"
        "    -- Ensure required fields are present without overwriting existing values\n"
        "    if resultTable.result == nil then\n"
        "      resultTable.result = res.result\n"
        "    end\n"
        "    if resultTable.skip == nil then\n"
        "      resultTable.skip = res.skip\n"
        "    end\n"
        "  end\n"

        "  return resultTable  -- Return the existing or new resultTable\n"
        "end\n");

    // Show syntax error
    if (luaL_dostring(L, inputString.c_str()) != LUA_OK) {
        const char* cstr = lua_tostring(L, -1);
        lua_pop(L, 1);
        if (isLuaErrorDialogEnabled) {
            std::wstring error_message = utf8ToWString(cstr);
            MessageBox(nppData._nppHandle, error_message.c_str(), getLangStr(L"msgbox_title_use_variables_syntax_error").c_str(), MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        }
        lua_close(L);
        return false;
    }

    // Retrieve the result from the table
    lua_getglobal(L, "resultTable");
    if (lua_istable(L, -1)) {
        lua_getfield(L, -1, "result");
        if (lua_isnil(L, -1)) {
            inputString.clear();  // Clear inputString to represent nil
        }
        else if (lua_isstring(L, -1) || lua_isnumber(L, -1)) {
            std::string result = lua_tostring(L, -1);  // Get the result as a string
            if (regex) {
                result = escapeForRegex(result);  // Escape result if it will be used in regex
            }
            inputString = result;  // Update inputString with the potentially escaped result
        }
        lua_pop(L, 1);  // Pop the 'result' field from the stack

        // Retrieve the skip flag from the table
        lua_getfield(L, -1, "skip");
        if (lua_isboolean(L, -1)) {
            skip = lua_toboolean(L, -1);
        }
        else {
            skip = false;
        }
        lua_pop(L, 1);  // Pop the 'skip' field from the stack
    }
    else {
        // Show Runtime error
        if (isLuaErrorDialogEnabled) {
            std::wstring errorMsg = getLangStr(L"msgbox_use_variables_execution_error", { utf8ToWString(inputString.c_str()) });
            std::wstring errorTitle = getLangStr(L"msgbox_title_use_variables_execution_error");
            MessageBox(nppData._nppHandle, errorMsg.c_str(), errorTitle.c_str(), MB_OK);
        }
        lua_close(L);
        return false;
    }
    lua_pop(L, 1);  // Pop the 'result' table from the stack

    // Remove CAP variables to prevent them from being set in the next call
    std::string capVariablesStr;

    // Loop through the stored CAP names
    for (const auto& globalVarName : capNames) {
        lua_getglobal(L, globalVarName.c_str());

        // Include CAP variables in the debug output
        if (lua_isstring(L, -1)) {
            capVariablesStr += globalVarName + "\tString\t" + lua_tostring(L, -1) + "\n";
        }
        else if (lua_isnumber(L, -1)) {
            capVariablesStr += globalVarName + "\tNumber\t" + std::to_string(lua_tonumber(L, -1)) + "\n";
        }
        else if (lua_isboolean(L, -1)) {
            capVariablesStr += globalVarName + "\tBoolean\t" + (lua_toboolean(L, -1) ? "true" : "false") + "\n";
        }
        else {
            capVariablesStr += globalVarName + "\t<nil>\n";  // Handle missing or nil values
        }
        capVariablesStr += "\n";
        lua_pop(L, 1); // Pop the variable value

        // Remove the CAP variable from Lua
        lua_pushnil(L);
        lua_setglobal(L, globalVarName.c_str());
    }

    // Read Lua global Variables
    captureLuaGlobals(L);

    // Capture updated HASH tables
    captureHashTables(L);

    std::string luaVariablesStr;
    for (const auto& pair : globalLuaVariablesMap) {
        const LuaVariable& var = pair.second;
        std::string value;
        switch (var.type) {
        case LuaVariableType::String:
            value = var.stringValue;
            break;
        case LuaVariableType::Number:
            value = std::to_string(var.numberValue);
            break;
        case LuaVariableType::Boolean:
            value = var.booleanValue ? "true" : "false";
            break;
        default:
            value = "None or Unsupported Type";
            break;
        }
        luaVariablesStr += var.name + "\t" +
            ((var.type == LuaVariableType::String) ? "String" :
                (var.type == LuaVariableType::Number) ? "Number" :
                (var.type == LuaVariableType::Boolean) ? "Boolean" : "Unknown") +
            "\t" + value + "\n";
    }

    // Combine global variables and CAP variables for debug output
    std::string combinedVariablesStr = luaVariablesStr + capVariablesStr;

    // Show debug window if 'DEBUG' is true and clean the stack
    lua_getglobal(L, "DEBUG");
    if (lua_isboolean(L, -1) && lua_toboolean(L, -1)) {

        // Refresh the ListView for Statistics
        refreshUIListView();

        int response = ShowDebugWindow(combinedVariablesStr);

        if (response == 3) { // Stop button was pressed
            lua_pop(L, 1);
            lua_close(L);
            return false; // Stop the process if "Stop" is clicked
        }
    }
    lua_pop(L, 1);

    lua_close(L);

    return true;

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

void MultiReplace::setLuaFileVars(LuaVariables& vars) {
    wchar_t filePathBuffer[MAX_PATH] = { 0 };
    wchar_t fileNameBuffer[MAX_PATH] = { 0 };

    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(filePathBuffer));
    ::SendMessage(nppData._nppHandle, NPPM_GETFILENAME, MAX_PATH, reinterpret_cast<LPARAM>(fileNameBuffer));

    std::string filePath = wstringToString(std::wstring(filePathBuffer));
    std::string fileName = wstringToString(std::wstring(fileNameBuffer));

    if (!filePath.empty() && (filePath.find('\\') != std::string::npos || filePath.find('/') != std::string::npos)) {
        vars.FPATH = filePath; // Assign the full path if valid
    }
    else {
        vars.FPATH.clear(); // Clear FPATH if it's not a valid path
    }

    vars.FNAME = fileName; // Assign the extracted file name
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
        nppData._nppHandle, NULL, hInstance, (LPVOID)wMessage
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
    while (GetMessage(&msg, NULL, 0, 0)) {
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

        // Check for WM_QUIT after dispatching the message
        if (msg.message == WM_QUIT) {
            break;
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
        // Handle the window close button (X)
        debugWindowResponse = 3; // Set to the value that indicates the "Stop" button was pressed

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
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void MultiReplace::CopyListViewToClipboard(HWND hListView) {
    int itemCount = ListView_GetItemCount(hListView);
    int columnCount = Header_GetItemCount(ListView_GetHeader(hListView));

    std::wstring clipboardText;

    for (int i = 0; i < itemCount; ++i) {
        for (int j = 0; j < columnCount; ++j) {
            wchar_t buffer[256];
            ListView_GetItemText(hListView, i, j, buffer, 256);
            buffer[255] = L'\0';
            clipboardText += buffer;
            if (j < columnCount - 1) {
                clipboardText += L'\t';
            }
        }
        clipboardText += L'\n';
    }

    if (OpenClipboard(0)) {
        EmptyClipboard();
        HGLOBAL hClipboardData = GlobalAlloc(GMEM_DDESHARE, sizeof(WCHAR) * (clipboardText.length() + 1));
        if (hClipboardData) {
            WCHAR* pchData = reinterpret_cast<WCHAR*>(GlobalLock(hClipboardData));
            if (pchData) {
                wcscpy(pchData, clipboardText.c_str());
                GlobalUnlock(hClipboardData);
                if (!SetClipboardData(CF_UNICODETEXT, hClipboardData)) {
                    GlobalFree(hClipboardData);
                }
            }
        }
        CloseClipboard();
    }
}

void MultiReplace::CloseDebugWindow() {
    // Triggers the WM_CLOSE message for the debug window, handled in DebugWindowProc
    if (hDebugWnd != NULL) {
        // Save the window position and size before closing
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
    }
}

#pragma endregion


#pragma region Find

void MultiReplace::handleFindNextButton() {
    size_t matchIndex = std::numeric_limits<size_t>::max();

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SelectionInfo selection = getSelectionInfo(false);

    // Check if there is a selection and if the selection radio button is checked
    Sci_Position searchPos;
    if (selection.length > 0 && (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)) {
        // Jump to the start of the selection if the selection exists and the radio button is checked
        searchPos = selection.startPos;
    }
    else {
        // Otherwise, use the current cursor position
        searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
    }

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_find_directly"), COLOR_ERROR);
            return;
        }

        SearchResult result = performListSearchForward(replaceListData, searchPos, matchIndex);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performListSearchForward(replaceListData, 0, matchIndex);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView(); // Refresh the ListView to show updated statistic
                selectListItem(matchIndex); // Highlight the matched item in the list
                showStatusMessage(getLangStr(L"status_wrapped"), COLOR_INFO);
                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", COLOR_SUCCESS);
            updateCountColumns(matchIndex, 1);
            refreshUIListView(); // Refresh the ListView to show updated statistic
            selectListItem(matchIndex); // Highlight the matched item in the list

            // Disable selection radio and switch to "All Text" if it was previously checked or Search will be trapped in new selection
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found"), COLOR_ERROR, true);
        }
    }
    else {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);

        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

        std::string findTextUtf8 = convertAndExtend(findText, extended);
        SearchResult result = performSearchForward(findTextUtf8, searchFlags, true, searchPos);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performSearchForward(findTextUtf8, searchFlags, true, 0);
            if (result.pos >= 0) {
                showStatusMessage(getLangStr(L"status_wrapped"), COLOR_INFO);
                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", COLOR_SUCCESS);

            // Disable selection radio and switch to "All Text" if it was previously checked or Search will be trapped in new selection
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }

        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found_for", { findText }), COLOR_ERROR, true);
        }
    }
}

void MultiReplace::handleFindPrevButton() {
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SelectionInfo selection = getSelectionInfo(true);
    Sci_Position searchPos;

    // Determine starting position based on selection
    if (selection.length > 0 && (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)) {
        // Start from the end of the selection for backward search
        searchPos = selection.endPos;
    }
    else {
        // Default to current cursor position if no selection
        searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
    }

    // Move back one position if possible
    searchPos = (searchPos > 0) ? ::SendMessage(_hScintilla, SCI_POSITIONBEFORE, searchPos, 0) : searchPos;

    if (useListEnabled) {
        size_t matchIndex = std::numeric_limits<size_t>::max();

        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_find_directly"), COLOR_ERROR);
            return;
        }

        SearchResult result = performListSearchBackward(replaceListData, searchPos, matchIndex);

        if (result.pos < 0 && wrapAroundEnabled) {
            searchPos = ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0);

            result = performListSearchBackward(replaceListData, searchPos, matchIndex);

            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView(); // Refresh the ListView to show updated statistic
                selectListItem(matchIndex); // Highlight the matched item in the list
                showStatusMessage(getLangStr(L"status_wrapped"), COLOR_INFO);

                // Disable selection radio and switch to "All Text" if it was previously checked or Search will be trapped in new selection
                if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
                }

                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", COLOR_SUCCESS);
            updateCountColumns(matchIndex, 1);
            refreshUIListView(); // Refresh the ListView to show updated statistic
            selectListItem(matchIndex); // Highlight the matched item in the list

            // Disable selection radio and switch to "All Text" if it was previously checked or Search will be trapped in new selection
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found"), COLOR_ERROR, true);
        }
    }
    else {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

        std::string findTextUtf8 = convertAndExtend(findText, extended);

        SearchResult result = performSearchBackward(findTextUtf8, searchFlags, true, searchPos);

        if (result.pos < 0 && wrapAroundEnabled) {
            // Wrap-around logic
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                // Wrap around to the end of the last selection
                searchPos = selection.endPos;
            }
            else {
                // Wrap around to the end of the document
                searchPos = ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0);
            }

            result = performSearchBackward(findTextUtf8, searchFlags, true, searchPos);

            if (result.pos >= 0) {
                showStatusMessage(getLangStr(L"status_wrapped"), COLOR_INFO);

                // Disable selection radio and switch to "All Text" if it was previously checked or Search will be trapped in new selection
                if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                    ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
                }

                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", COLOR_SUCCESS);

            // Disable selection radio and switch to "All Text" if previously checked
            if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
                ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), FALSE);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
                ::SendMessage(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
            }
        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found_for", { findText }), COLOR_ERROR, true);
        }

    }
}

SearchResult MultiReplace::performSingleSearch(const std::string& findTextUtf8, int searchFlags, bool selectMatch, SelectionRange range) {

    // Check if the search string is empty
    if (findTextUtf8.empty()) {
        return SearchResult();  // Returns the default-initialized SearchResult
    }

    send(SCI_SETTARGETSTART, range.start, 0);
    send(SCI_SETTARGETEND, range.end, 0);
    send(SCI_SETSEARCHFLAGS, searchFlags, 0);

    LRESULT pos = send(SCI_SEARCHINTARGET, findTextUtf8.length(), reinterpret_cast<sptr_t>(findTextUtf8.c_str()));

    SearchResult result;
    result.pos = pos;

    if (pos >= 0) {
        // If a match is found, set additional result data
        result.length = send(SCI_GETTARGETEND, 0, 0) - pos;

        // Consider the worst case for UTF-8, where one character could be up to 4 bytes.
        char buffer[MAX_TEXT_LENGTH * 4 + 1] = { 0 };  // Assuming UTF-8 encoding in Scintilla
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = static_cast<int>(result.pos);
        tr.chrg.cpMax = static_cast<int>(result.pos + result.length);

        if (tr.chrg.cpMax - tr.chrg.cpMin > sizeof(buffer) - 1) {
            // Safety check to avoid overflow.
            tr.chrg.cpMax = tr.chrg.cpMin + sizeof(buffer) - 1;
        }

        tr.lpstrText = buffer;
        send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));

        result.foundText = std::string(buffer);

        // If selectMatch is true, highlight the found text
        if (selectMatch) {
            displayResultCentered(result.pos, result.pos + result.length, true);
        }
    }

    return result;
}

SearchResult MultiReplace::performSearchForward(const std::string& findTextUtf8, int searchFlags, bool selectMatch, LRESULT start)
{
    // Set search direction to forward
    bool isBackward = false;

    SelectionRange targetRange;
    SearchResult result;

    // Check if selection mode is active
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        // Perform search within selection
        result = performSearchSelection(findTextUtf8, searchFlags, selectMatch, start, isBackward);
    }
    // Check if IDC_COLUMN_MODE_RADIO is enabled, selectMatch is false, and column delimiter data is set
    else if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED && columnDelimiterData.isValid()) {
        // Perform search within columns
        result = performSearchColumn(findTextUtf8, searchFlags, selectMatch, start, isBackward);
    }
    else {
        // If neither IDC_SELECTION_RADIO nor IDC_COLUMN_MODE_RADIO, perform search within the whole document
        targetRange.start = start;
        targetRange.end = send(SCI_GETLENGTH, 0, 0); // End of the document
        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
    }

    return result;
}

SearchResult MultiReplace::performSearchBackward(const std::string& findTextUtf8, int searchFlags, bool selectMatch, LRESULT start) {

    // Set search direction to backward
    bool isBackward = true;

    SelectionRange targetRange;
    SearchResult result;

    // Check if selection mode is active
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        // Perform search within selection
        result = performSearchSelection(findTextUtf8, searchFlags, selectMatch, start, isBackward);
    }
    else if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED && columnDelimiterData.isValid()) {
        // Perform search within columns
        result = performSearchColumn(findTextUtf8, searchFlags, selectMatch, start, isBackward);
    }
    else {
        // Default backward search in the whole document
        targetRange.start = start;
        targetRange.end = 0; // Beginning of the document
        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
    }

    return result;
}

SearchResult MultiReplace::performSearchSelection(const std::string& findTextUtf8, int searchFlags, bool selectMatch, LRESULT start, bool isBackward) {
    SearchResult result;
    SelectionRange targetRange;
    std::vector<SelectionRange> selections;

    LRESULT selectionCount = ::SendMessage(_hScintilla, SCI_GETSELECTIONS, 0, 0);
    if (selectionCount == 0) {
        return SearchResult(); // No selections to search
    }

    // Gather all selection positions
    selections.resize(selectionCount);
    for (int i = 0; i < selectionCount; i++) {
        selections[i].start = ::SendMessage(_hScintilla, SCI_GETSELECTIONNSTART, i, 0);
        selections[i].end = ::SendMessage(_hScintilla, SCI_GETSELECTIONNEND, i, 0);
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
        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
        if (result.pos >= 0) return result;

        // Update the starting position for the next iteration
        start = isBackward ? selection.start - 1 : selection.end + 1;
    }

    return result; // No match found
}

SearchResult MultiReplace::performSearchColumn(const std::string& findTextUtf8, int searchFlags, bool selectMatch, LRESULT start, bool isBackward)
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
        if (line >= static_cast<LRESULT>(lineDelimiterPositions.size())) {
            break; // Avoid out-of-bounds access
        }

        const auto& linePositions = lineDelimiterPositions[line].positions;
        SIZE_T totalColumns = linePositions.size() + 1;

        // Set column iteration range and step based on direction
        SIZE_T column = isBackward ? (line == startLine ? startColumnIndex : totalColumns) : startColumnIndex;
        SIZE_T endColumnIndex = isBackward ? 1 : totalColumns;
        int columnStep = isBackward ? -1 : 1;

        // Iterate over columns in the specified direction
        for (; (isBackward ? (column >= endColumnIndex) : (column <= endColumnIndex)); column += columnStep) {
            LRESULT startColumn = 0;
            LRESULT endColumn = 0;

            // Define start and end positions for the current column
            if (column == 1) {
                startColumn = lineDelimiterPositions[line].startPosition;
            }
            else {
                startColumn = linePositions[column - 2].position + columnDelimiterData.delimiterLength;
            }

            if (column == totalColumns) {
                endColumn = lineDelimiterPositions[line].endPosition;
            }
            else {
                endColumn = linePositions[column - 1].position;
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
            result = performSingleSearch(
                findTextUtf8,
                searchFlags ,
                selectMatch,
                targetRange
            );

            // Return if a match is found
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

SearchResult MultiReplace::performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max(); // Initialize with a value that represents "no index".

    for (size_t i = 0; i < list.size(); ++i) {
        if (list[i].isEnabled) {
            int searchFlags = (list[i].wholeWord * SCFIND_WHOLEWORD) |
                (list[i].matchCase * SCFIND_MATCHCASE) |
                (list[i].regex * SCFIND_REGEXP);
            std::string findTextUtf8 = convertAndExtend(list[i].findText, list[i].extended);
            SearchResult result = performSearchBackward(findTextUtf8, searchFlags, false, cursorPos);

            // If a match was found and it's closer to the cursor than the current closest match, update the closest match
            if (result.pos >= 0 && (closestMatch.pos < 0 || (result.pos + result.length) >(closestMatch.pos + closestMatch.length))) {
                closestMatch = result;
                closestMatchIndex = i; // Update the index of the closest match
            }
        }
    }

    if (closestMatch.pos >= 0) { // Check if a match was found
        displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, false);
    }

    return closestMatch;
}

SearchResult MultiReplace::performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max(); // Initialize with a value representing "no index".

    for (size_t i = 0; i < list.size(); i++) {
        if (list[i].isEnabled) {
            int searchFlags = (list[i].wholeWord * SCFIND_WHOLEWORD) | (list[i].matchCase * SCFIND_MATCHCASE) | (list[i].regex * SCFIND_REGEXP);
            std::string findTextUtf8 = convertAndExtend(list[i].findText, list[i].extended);
            SearchResult result = performSearchForward(findTextUtf8, searchFlags, false, cursorPos);

            // If a match is found that is closer to the cursor than the current closest match, update the closest match
            if (result.pos >= 0 && (closestMatch.pos < 0 || result.pos < closestMatch.pos)) {
                closestMatch = result;
                closestMatchIndex = i; // Update the index of the closest match
            }
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
    ::SendMessage(_hScintilla, SCI_ENSUREVISIBLE, ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, posStart, 0), 0);
    ::SendMessage(_hScintilla, SCI_ENSUREVISIBLE, ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, posEnd, 0), 0);

    // Jump-scroll to center, if current position is out of view
    ::SendMessage(_hScintilla, SCI_SETVISIBLEPOLICY, CARET_JUMPS | CARET_EVEN, 0);
    ::SendMessage(_hScintilla, SCI_ENSUREVISIBLEENFORCEPOLICY, ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // When searching up, the beginning of the (possible multiline) result is important, when scrolling down the end
    ::SendMessage(_hScintilla, SCI_GOTOPOS, isDownwards ? posEnd : posStart, 0);
    ::SendMessage(_hScintilla, SCI_SETVISIBLEPOLICY, CARET_EVEN, 0);
    ::SendMessage(_hScintilla, SCI_ENSUREVISIBLEENFORCEPOLICY, ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // Adjust so that we see the entire match; primarily horizontally
    ::SendMessage(_hScintilla, SCI_SCROLLRANGE, posStart, posEnd);

    // Move cursor to end of result and select result
    ::SendMessage(_hScintilla, SCI_GOTOPOS, posEnd, 0);
    ::SendMessage(_hScintilla, SCI_SETANCHOR, posStart, 0);

    // Update Scintilla's knowledge about what column the caret is in, so that if user
    // does up/down arrow as first navigation after the search result is selected,
    // the caret doesn't jump to an unexpected column
    ::SendMessage(_hScintilla, SCI_CHOOSECARETX, 0, 0);

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
    int totalMatchCount = 0;
    markedStringsCount = 0;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_mark_directly"), COLOR_ERROR);
            return;
        }

        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled) {
                std::string findTextUtf8 = convertAndExtend(replaceListData[i].findText, replaceListData[i].extended);
                int searchFlags = (replaceListData[i].wholeWord * SCFIND_WHOLEWORD)
                    | (replaceListData[i].matchCase * SCFIND_MATCHCASE)
                    | (replaceListData[i].regex * SCFIND_REGEXP);
                int matchCount = markString(findTextUtf8, searchFlags);
                totalMatchCount += matchCount;

                if (matchCount > 0) {
                    updateCountColumns(i, matchCount);
                    refreshUIListView(); // Refresh the ListView to show updated statistic
                }
            }
        }
    }
    else {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        std::string findTextUtf8 = convertAndExtend(findText, extended);
        int searchFlags = (wholeWord * SCFIND_WHOLEWORD)
            | (matchCase * SCFIND_MATCHCASE)
            | (regex * SCFIND_REGEXP);
        totalMatchCount = markString(findTextUtf8, searchFlags);

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
    showStatusMessage(getLangStr(L"status_occurrences_marked", { std::to_wstring(totalMatchCount) }), COLOR_INFO);
}

int MultiReplace::markString(const std::string& findTextUtf8, int searchFlags) {
    if (findTextUtf8.empty()) {
        return 0;
    }

    int markCount = 0;  // Counter for marked matches
    SearchResult searchResult = performSearchForward(findTextUtf8, searchFlags, false, 0);
    while (searchResult.pos >= 0) {
        highlightTextRange(searchResult.pos, searchResult.length, findTextUtf8);
        markCount++;
        searchResult = performSearchForward(findTextUtf8, searchFlags, false, searchResult.pos + searchResult.length);
    }

    if (useListEnabled && markCount > 0) {
        markedStringsCount++;
    }

    return markCount;
}

void MultiReplace::highlightTextRange(LRESULT pos, LRESULT len, const std::string& findTextUtf8)
{
    long color = useListEnabled ? generateColorValue(findTextUtf8) : MARKER_COLOR;

    // Check if the color already has an associated style
    int indicatorStyle;
    if (colorToStyleMap.find(color) == colorToStyleMap.end()) {
        // If not, assign a new style and store it in the map
        indicatorStyle = useListEnabled ? textStyles[(colorToStyleMap.size() % (textStyles.size() - 1)) + 1] : textStyles[0];
        colorToStyleMap[color] = indicatorStyle;
    }
    else {
        // If yes, use the existing style
        indicatorStyle = colorToStyleMap[color];
    }

    // Set and apply highlighting style
    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, indicatorStyle, 0);
    ::SendMessage(_hScintilla, SCI_INDICSETSTYLE, indicatorStyle, INDIC_STRAIGHTBOX);

    if (colorToStyleMap.size() < textStyles.size()) {
        ::SendMessage(_hScintilla, SCI_INDICSETFORE, indicatorStyle, color);
    }

    ::SendMessage(_hScintilla, SCI_INDICSETALPHA, indicatorStyle, 100);
    ::SendMessage(_hScintilla, SCI_INDICATORFILLRANGE, pos, len);
}

long MultiReplace::generateColorValue(const std::string& str) {
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
    long color = (r << 16) | (g << 8) | b;

    return color;
}

void MultiReplace::handleClearTextMarksButton()
{
    for (int style : textStyles)
    {
        ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, style, 0);
        ::SendMessage(_hScintilla, SCI_INDICATORCLEARRANGE, 0, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
    }

    markedStringsCount = 0;
    colorToStyleMap.clear();
}

void MultiReplace::handleCopyMarkedTextToClipboardButton()
{
    bool wasLastCharMarked = false;
    size_t markedTextCount = 0;

    std::string markedText;
    std::string styleText;
    std::string eol = getEOLStyle();

    for (int style : textStyles)
    {
        ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, style, 0);
        LRESULT pos = 0;
        LRESULT nextPos = ::SendMessage(_hScintilla, SCI_INDICATOREND, style, pos);

        while (nextPos > pos) // check if nextPos has advanced
        {
            bool atEndOfIndic = ::SendMessage(_hScintilla, SCI_INDICATORVALUEAT, style, pos) != 0;

            if (atEndOfIndic)
            {
                if (!wasLastCharMarked)
                {
                    ++markedTextCount;
                }

                wasLastCharMarked = true;

                for (LRESULT i = pos; i < nextPos; ++i)
                {
                    char ch = static_cast<char>(::SendMessage(_hScintilla, SCI_GETCHARAT, i, 0));
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
            nextPos = ::SendMessage(_hScintilla, SCI_INDICATOREND, style, pos);
        }
    }

    // Remove the last EOL if necessary
    if (!markedText.empty() && markedText.length() >= eol.length()) {
        markedText.erase(markedText.length() - eol.length());
    }

    // Convert encoding to wide string
    std::wstring wstr = stringToWString(markedText);

    copyTextToClipboard(wstr, static_cast<int>(markedTextCount));
}

void MultiReplace::copyTextToClipboard(const std::wstring& text, int textCount)
{
    if (text.empty()) {
        showStatusMessage(getLangStr(L"status_no_text_to_copy"), COLOR_ERROR);
        return;
    }

    if (OpenClipboard(0))
    {
        EmptyClipboard();
        HGLOBAL hClipboardData = GlobalAlloc(GMEM_DDESHARE, sizeof(WCHAR) * (text.length() + 1));
        if (hClipboardData)
        {
            WCHAR* pchData = reinterpret_cast<WCHAR*>(GlobalLock(hClipboardData));
            if (pchData)
            {
                wcscpy(pchData, text.c_str());
                GlobalUnlock(hClipboardData);
                if (SetClipboardData(CF_UNICODETEXT, hClipboardData) != NULL)
                {
                    showStatusMessage(getLangStr(L"status_items_copied_to_clipboard", { std::to_wstring(textCount) }), COLOR_SUCCESS);
                }
                else
                {
                    showStatusMessage(getLangStr(L"status_failed_to_copy"), COLOR_ERROR);
                }
            }
            else
            {
                showStatusMessage(getLangStr(L"status_failed_allocate_memory"), COLOR_ERROR);
                GlobalFree(hClipboardData);
            }
        }
        CloseClipboard();
    }
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
    std::wstring confirmMessage = getLangStr(L"msgbox_confirm_delete_columns", { std::to_wstring(columnCount) });

    // Display a message box with Yes/No options and a question mark icon
    int msgboxID = MessageBox(
        nppData._nppHandle,
        confirmMessage.c_str(),
        getLangStr(L"msgbox_title_confirm").c_str(),
        MB_ICONWARNING | MB_YESNO
    );

    return (msgboxID == IDYES);  // Return true if user confirmed, else false
}

void MultiReplace::handleDeleteColumns()
{
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), COLOR_ERROR);
        return;
    }

    ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);

    int deletedFieldsCount = 0;
    SIZE_T lineCount = lineDelimiterPositions.size();

    // Loop from the last element down to the first
    for (SIZE_T i = lineCount; i-- > 0; ) {
        const auto& lineInfo = lineDelimiterPositions[i];

        // Process each column in reverse
        for (auto it = columnDelimiterData.columns.rbegin(); it != columnDelimiterData.columns.rend(); ++it) {
            SIZE_T column = *it;

            // Only process columns within the valid range
            if (column <= lineInfo.positions.size() + 1) {

                LRESULT startPos, endPos;

                if (column == 1) {
                    startPos = lineInfo.startPosition;
                }
                else if (column - 2 < lineInfo.positions.size()) {
                    startPos = lineInfo.positions[column - 2].position;
                }
                else {
                    continue;
                }

                if (column - 1 < lineInfo.positions.size()) {
                    // Delete leading Delimiter if first column will be droped
                    if (column == 1) {
                        endPos = lineInfo.positions[column - 1].position + columnDelimiterData.delimiterLength;
                    }
                    else {
                        endPos = lineInfo.positions[column - 1].position;
                    }
                }
                else {
                    endPos = lineInfo.endPosition;
                }

                send(SCI_DELETERANGE, startPos, endPos - startPos, false);

                deletedFieldsCount++;
            }
        }

    }
    ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);

    // Show status message
    showStatusMessage(getLangStr(L"status_deleted_fields_count", { std::to_wstring(deletedFieldsCount) }), COLOR_SUCCESS);
}

void MultiReplace::handleCopyColumnsToClipboard()
{
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), COLOR_ERROR);
        return;
    }

    std::string combinedText;
    int copiedFieldsCount = 0;
    size_t lineCount = lineDelimiterPositions.size();

    // Iterate through each line
    for (size_t i = 0; i < lineCount; ++i) {
        const auto& lineInfo = lineDelimiterPositions[i];

        bool isFirstCopiedColumn = true;
        std::string lineText;

        // Process each column
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= lineInfo.positions.size() + 1) {
                LRESULT startPos, endPos;

                if (column == 1) {
                    startPos = lineInfo.startPosition;
                    isFirstCopiedColumn = false;
                }
                else if (column - 2 < lineInfo.positions.size()) {
                    startPos = lineInfo.positions[column - 2].position;
                    // Drop first Delimiter if copied as first column
                    if (isFirstCopiedColumn) {
                        startPos += columnDelimiterData.delimiterLength;
                        isFirstCopiedColumn = false;
                    }
                }
                else {
                    break;
                }

                if (column - 1 < lineInfo.positions.size()) {
                    endPos = lineInfo.positions[column - 1].position;
                }
                else {
                    endPos = lineInfo.endPosition;
                }

                // Buffer to hold the text
                std::vector<char> buffer(static_cast<size_t>(endPos - startPos) + 1);

                // Prepare TextRange structure for Scintilla
                Sci_TextRangeFull tr;
                tr.chrg.cpMin = startPos;
                tr.chrg.cpMax = endPos;
                tr.lpstrText = buffer.data();

                // Extract text for the column
                SendMessage(_hScintilla, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));
                lineText += std::string(buffer.data());

                copiedFieldsCount++;
            }
        }

        combinedText += lineText;
        // Add a newline except after the last line
        if (i < lineCount - 1) {
            combinedText += "\n";
        }
    }

    // Convert to Wide String and Copy to Clipboard (Ensure this line only occurs once and wstr is not redefined)
    std::wstring wstr = stringToWString(combinedText);
    copyTextToClipboard(wstr, copiedFieldsCount);
}

#pragma endregion


#pragma region CSV Sort

std::vector<CombinedColumns> MultiReplace::extractColumnData(SIZE_T startLine, SIZE_T lineCount) {
    std::vector<CombinedColumns> combinedData;
    for (SIZE_T i = startLine; i < lineCount; ++i) {
        const auto& lineInfo = lineDelimiterPositions[i]; // Stelle sicher, dass lineDelimiterPositions definiert ist
        CombinedColumns rowData;
        rowData.columns.resize(columnDelimiterData.inputColumns.size()); // Stelle sicher, dass columnDelimiterData definiert ist

        size_t columnIndex = 0;
        for (SIZE_T columnNumber : columnDelimiterData.inputColumns) {
            LRESULT startPos, endPos;

            // Berechne Start- und Endpositionen für jede Spalte
            if (columnNumber == 1) {
                startPos = lineInfo.startPosition;
            }
            else if (columnNumber - 2 < lineInfo.positions.size()) {
                startPos = lineInfo.positions[columnNumber - 2].position + columnDelimiterData.delimiterLength;
            }
            else {
                continue;
            }

            if (columnNumber - 1 < lineInfo.positions.size()) {
                endPos = lineInfo.positions[columnNumber - 1].position;
            }
            else {
                endPos = lineInfo.endPosition;
            }

            // Puffer, um den Text zu halten
            std::vector<char> buffer(static_cast<size_t>(endPos - startPos) + 1);
            Sci_TextRangeFull tr;
            tr.chrg.cpMin = startPos;
            tr.chrg.cpMax = endPos;
            tr.lpstrText = buffer.data();

            // Extrahiere Text für die Spalte
            SendMessage(_hScintilla, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));
            rowData.columns[columnIndex++] = std::string(buffer.data());
        }

        combinedData.push_back(rowData);
    }

    return combinedData;
}

void MultiReplace::sortRowsByColumn(SortDirection sortDirection) {
    // Validate column delimiter data
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), COLOR_ERROR);
        return;
    }
    SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);

    size_t lineCount = lineDelimiterPositions.size();

    // Check if there's nothing to sort or if the document has fewer lines than header lines
    if (lineCount <= CSVheaderLinesCount) {
        // Either nothing to sort or document consists only of header lines
        SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
        return;
    }

    std::vector<CombinedColumns> combinedData;
    combinedData.reserve(lineCount); // Reserve space for all lines, including headers

    // Initialize tempOrder with indices for all lines, including header lines
    std::vector<size_t> tempOrder(lineCount);
    for (size_t i = 0; i < lineCount; ++i) {
        tempOrder[i] = i; // Manually filling tempOrder with 0, 1, ..., lineCount-1
    }

    // Extract content of specified columns, starting after header lines
    combinedData = extractColumnData(CSVheaderLinesCount, lineDelimiterPositions.size());

    // Sort the tempOrder based on combinedData, excluding header lines during comparison
    std::sort(tempOrder.begin() + CSVheaderLinesCount, tempOrder.end(), [&](const size_t a, const size_t b) {
        size_t adjustedA = a - CSVheaderLinesCount;
        size_t adjustedB = b - CSVheaderLinesCount;
        // Implement the sorting logic here, only for lines beyond the header lines
        return sortDirection == SortDirection::Ascending ? combinedData[adjustedA].columns[0] < combinedData[adjustedB].columns[0] : combinedData[adjustedA].columns[0] > combinedData[adjustedB].columns[0];
        });

    // Adjust originalLineOrder based on the opposite sorting results
    if (!originalLineOrder.empty()) {
        std::vector<size_t> newOrder(originalLineOrder.size());
        for (size_t i = 0; i < tempOrder.size(); ++i) {
            size_t positionInOriginal = tempOrder[i];
            size_t valueForNewOrder = originalLineOrder[positionInOriginal];
            newOrder[i] = valueForNewOrder;
        }
        originalLineOrder = std::move(newOrder);
    }
    else {
        originalLineOrder = tempOrder;
    }

    // Use tempOrder to reorder lines in Scintilla, adjusting for header lines
    reorderLinesInScintilla(tempOrder);

    SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
}

void MultiReplace::reorderLinesInScintilla(const std::vector<size_t>& sortedIndex) {
    std::string lineBreak = getEOLStyle();

    isSortedColumn = false; // Stop logging changes
    // Extract the text of each line based on the sorted index and include a line break after each
    std::string combinedLines;
    for (size_t i = 0; i < sortedIndex.size(); ++i) {
        size_t idx = sortedIndex[i];
        LRESULT lineStart = SendMessage(_hScintilla, SCI_POSITIONFROMLINE, idx, 0);
        LRESULT lineEnd = SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, idx, 0);
        std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1); // Buffer size includes space for null terminator
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = lineStart;
        tr.chrg.cpMax = lineEnd;
        tr.lpstrText = buffer.data();
        SendMessage(_hScintilla, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));
        combinedLines += std::string(buffer.data(), buffer.size() - 1); // Exclude null terminator from the string
        if (i < sortedIndex.size() - 1) {
            combinedLines += lineBreak; // Add line break after each line except the last
        }
    }

    // Clear all content from Scintilla
    SendMessage(_hScintilla, SCI_CLEARALL, 0, 0);

    // Re-insert the combined lines
    SendMessage(_hScintilla, SCI_APPENDTEXT, combinedLines.length(), reinterpret_cast<LPARAM>(combinedLines.c_str()));

    isSortedColumn = true; // Ready for logging changes
}

void MultiReplace::restoreOriginalLineOrder(const std::vector<size_t>& originalOrder) {

    // Determine the total number of lines in the document
    size_t totalLineCount = SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);

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
        LRESULT lineStart = SendMessage(_hScintilla, SCI_POSITIONFROMLINE, i, 0);
        LRESULT lineEnd = SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, i, 0);
        std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = lineStart;
        tr.chrg.cpMax = lineEnd;
        tr.lpstrText = buffer.data();
        SendMessage(_hScintilla, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));
        sortedLines[newPosition] = std::string(buffer.data(), buffer.size() - 1); // Exclude null terminator
    }

    // Clear the content of the editor
    SendMessage(_hScintilla, SCI_CLEARALL, 0, 0);

    // Re-insert the lines in their original order
    for (size_t i = 0; i < sortedLines.size(); ++i) {
        //std::string message = "Inserting line at position: " + std::to_string(i) + "\nContent: " + sortedLines[i];
        //MessageBoxA(NULL, message.c_str(), "Debug Insert Line", MB_OK);
        SendMessage(_hScintilla, SCI_APPENDTEXT, sortedLines[i].length(), reinterpret_cast<LPARAM>(sortedLines[i].c_str()));
        // Add a line break after each line except the last one
        if (i < sortedLines.size() - 1) {
            SendMessage(_hScintilla, SCI_APPENDTEXT, lineBreak.length(), reinterpret_cast<LPARAM>(lineBreak.c_str()));
        }
    }
}

void MultiReplace::extractLineContent(size_t idx, std::string& content, const std::string& lineBreak) {
    LRESULT lineStart = SendMessage(_hScintilla, SCI_POSITIONFROMLINE, idx, 0);
    LRESULT lineEnd = SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, idx, 0);
    std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
    Sci_TextRangeFull tr{ lineStart, lineEnd, buffer.data() };
    SendMessage(_hScintilla, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<LPARAM>(&tr));
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

void MultiReplace::updateUnsortedDocument(SIZE_T lineNumber, ChangeType changeType) {
    if (!isSortedColumn || lineNumber > originalLineOrder.size()) {
        return; // Invalid line number, return early
    }

    switch (changeType) {
    case ChangeType::Insert: {
        // Find the current maximum value in originalLineOrder and add one for the new line placeholder
        size_t maxIndex = originalLineOrder.empty() ? 0 : (*std::max_element(originalLineOrder.begin(), originalLineOrder.end())) + 1;
        // Insert maxIndex for the new line at the specified position in originalLineOrder
        originalLineOrder.insert(originalLineOrder.begin() + lineNumber, maxIndex);
        break;
    }

    case ChangeType::Delete: {
        // Ensure lineNumber is within the range before attempting to delete
        if (lineNumber < originalLineOrder.size()) { // Safety check
            // Element at lineNumber represents the actual index to be deleted
            size_t actualIndexToDelete = originalLineOrder[lineNumber];
            // Directly remove the index at lineNumber
            originalLineOrder.erase(originalLineOrder.begin() + lineNumber);
            // Adjust subsequent indices to reflect the deletion
            for (size_t i = 0; i < originalLineOrder.size(); ++i) {
                if (originalLineOrder[i] > actualIndexToDelete) {
                    --originalLineOrder[i];
                }
            }
        }
        break;
    }
    }
}

#pragma endregion


#pragma region Scope

bool MultiReplace::parseColumnAndDelimiterData() {

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

    std::string tempExtendedDelimiter = convertAndExtend(delimiterData, true);

    columnDelimiterData.delimiterChanged = !(columnDelimiterData.extendedDelimiter == convertAndExtend(delimiterData, true));
    columnDelimiterData.quoteCharChanged = !(columnDelimiterData.quoteChar == wstringToString(quoteCharString));

    // Initilaize values in case it will return with an error
    columnDelimiterData.extendedDelimiter = "";
    columnDelimiterData.quoteChar = "";
    columnDelimiterData.delimiterLength = 0;

    // Parse column data
    columnDataString.erase(0, columnDataString.find_first_not_of(L','));
    columnDataString.erase(columnDataString.find_last_not_of(L',') + 1);

    if (columnDataString.empty() || delimiterData.empty()) {
        showStatusMessage(getLangStr(L"status_missing_column_or_delimiter_data"), COLOR_ERROR);
        columnDelimiterData.columns.clear();
        columnDelimiterData.extendedDelimiter = "";
        return false;
    }

    std::vector<int> inputColumns;
    std::set<int> columns;
    std::wstring::size_type start = 0;
    std::wstring::size_type end = columnDataString.find(',', start);

    while (end != std::wstring::npos) {
        std::wstring block = columnDataString.substr(start, end - start);
        size_t dashPos = block.find('-');

        // Check if block has range of columns (e.g., 1-3)
        if (dashPos != std::wstring::npos) {
            try {
                // Parse range and add each column to set
                int startRange = std::stoi(block.substr(0, dashPos));
                int endRange = std::stoi(block.substr(dashPos + 1));

                if (startRange < 1 || endRange < startRange) {
                    showStatusMessage(getLangStr(L"status_invalid_range_in_column_data"), COLOR_ERROR);
                    columnDelimiterData.columns.clear();
                    return false;
                }

                for (int i = startRange; i <= endRange; ++i) {
                    if (columns.insert(i).second) {
                        inputColumns.push_back(i);
                    }
                }
            }
            catch (const std::exception&) {
                showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), COLOR_ERROR);
                columnDelimiterData.columns.clear();
                return false;
            }
        }
        else {
            // Parse single column and add to set
            try {
                int column = std::stoi(block);

                if (column < 1) {
                    showStatusMessage(getLangStr(L"status_invalid_column_number"), COLOR_ERROR);
                    columnDelimiterData.columns.clear();
                    return false;
                }

                if (columns.insert(column).second) {
                    inputColumns.push_back(column);
                }
            }
            catch (const std::exception&) {
                showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), COLOR_ERROR);
                columnDelimiterData.columns.clear();
                return false;
            }
        }

        start = end + 1;
        end = columnDataString.find(',', start);
    }

    // Handle last block of column data
    std::wstring lastBlock = columnDataString.substr(start);
    size_t dashPos = lastBlock.find('-');

    // Similar processing to above, but for last block
    if (dashPos != std::wstring::npos) {
        try {
            int startRange = std::stoi(lastBlock.substr(0, dashPos));
            int endRange = std::stoi(lastBlock.substr(dashPos + 1));

            if (startRange < 1 || endRange < startRange) {
                showStatusMessage(getLangStr(L"status_invalid_range_in_column_data"), COLOR_ERROR);
                columnDelimiterData.columns.clear();
                return false;
            }

            for (int i = startRange; i <= endRange; ++i) {
                if (columns.insert(i).second) {
                    inputColumns.push_back(i);
                }
            }
        }
        catch (const std::exception&) {
            showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), COLOR_ERROR);
            columnDelimiterData.columns.clear();
            return false;
        }
    }
    else {
        try {
            int column = std::stoi(lastBlock);

            if (column < 1) {
                showStatusMessage(getLangStr(L"status_invalid_column_number"), COLOR_ERROR);
                columnDelimiterData.columns.clear();
                return false;
            }
            auto insertResult = columns.insert(column);
            if (insertResult.second) { // Check if the insertion was successful
                inputColumns.push_back(column); // Add to the inputColumns vector
            }
        }
        catch (const std::exception&) {
            showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), COLOR_ERROR);
            columnDelimiterData.columns.clear();
            return false;
        }
    }


    // Check delimiter data
    if (tempExtendedDelimiter.empty()) {
        showStatusMessage(getLangStr(L"status_extended_delimiter_empty"), COLOR_ERROR);
        return false;
    }

    // Check Quote Character
    if (!quoteCharString.empty() && (quoteCharString.length() != 1 || !(quoteCharString[0] == L'"' || quoteCharString[0] == L'\''))) {
        showStatusMessage(getLangStr(L"status_invalid_quote_character"), COLOR_ERROR);
        return false;
    }

    // Set dataChanged flag
    columnDelimiterData.columnChanged = !(columnDelimiterData.columns == columns);

    // Set columnDelimiterData values
    columnDelimiterData.inputColumns = inputColumns;
    columnDelimiterData.columns = columns;
    columnDelimiterData.extendedDelimiter = tempExtendedDelimiter;
    columnDelimiterData.delimiterLength = tempExtendedDelimiter.length();
    columnDelimiterData.quoteChar = wstringToString(quoteCharString);

    return true;
}

void MultiReplace::findAllDelimitersInDocument() {

    // Clear list for new data
    lineDelimiterPositions.clear();

    // Reset TextModified Trigger
    textModified = false;
    logChanges.clear();

    // Enable detailed logging for capturing delimiter positions
    isLoggingEnabled = true;

    // Get total line count in document
    LRESULT totalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);

    // Resize the list to fit total lines
    lineDelimiterPositions.resize(totalLines);

    // Find and store delimiter positions for each line
    for (LRESULT line = 0; line < totalLines; ++line) {
        findDelimitersInLine(line);
    }

    // Shrink the reusable buffer used to read each line to release unused memory
    lineBuffer.shrink_to_fit();

    // Clear log queue
    logChanges.clear();

}

void MultiReplace::findDelimitersInLine(LRESULT line) {
    // Initialize LineInfo for this line
    LineInfo lineInfo;

    // Get start and end positions of the line
    lineInfo.startPosition = send(SCI_POSITIONFROMLINE, line, 0);
    lineInfo.endPosition = send(SCI_GETLINEENDPOSITION, line, 0);

    // Get line length and allocate buffer
    LRESULT lineLength = send(SCI_LINELENGTH, line, 0);
    
    // Resize the lineBuffer only if needed
    if (lineBuffer.size() < static_cast<size_t>(lineLength + 1)) {
        lineBuffer.resize(lineLength + 1);  // Increase the buffer size if necessary
    }

    // Get line content
    send(SCI_GETLINE, line, reinterpret_cast<sptr_t>(lineBuffer.data()));
    std::string lineContent(lineBuffer.data(), lineLength);

    // Define structure to store delimiter position
    DelimiterPosition delimiterPos = { 0 };

    bool inQuotes = false;
    std::string::size_type pos = 0;

    bool hasQuoteChar = !columnDelimiterData.quoteChar.empty();
    char currentQuoteChar = hasQuoteChar ? columnDelimiterData.quoteChar[0] : 0;

    while (pos < lineContent.size()) {
        // If there's a defined quote character and it matches, toggle inQuotes
        if (hasQuoteChar && lineContent[pos] == currentQuoteChar) {
            inQuotes = !inQuotes;
            ++pos;
            continue;
        }

        if (!inQuotes && lineContent.compare(pos, columnDelimiterData.delimiterLength, columnDelimiterData.extendedDelimiter) == 0) {
            delimiterPos.position = pos + lineInfo.startPosition;
            lineInfo.positions.push_back(delimiterPos);
            pos += columnDelimiterData.delimiterLength;  // Skip delimiter for next iteration
            continue;
        }
        ++pos;
    }

    // Convert size of lineDelimiterPositions to signed integer
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());

    // Update lineDelimiterPositions with the LineInfo for this line
    if (line < listSize) {
        lineDelimiterPositions[line] = lineInfo;
    }
    else {
        // If the line index is greater than the current size of the list,
        // append new elements to the list
        lineDelimiterPositions.resize(line + 1);
        lineDelimiterPositions[line] = lineInfo;
    }
}

ColumnInfo MultiReplace::getColumnInfo(LRESULT startPosition) {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED ||
        columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty() ||
        lineDelimiterPositions.empty()) {
        return { 0, 0, 0 };
    }

    LRESULT totalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);
    LRESULT startLine = ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, startPosition, 0);
    SIZE_T startColumnIndex = 1;

    // Check if the line exists in lineDelimiterPositions
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());
    if (startLine < totalLines && startLine < listSize) {
        const auto& linePositions = lineDelimiterPositions[startLine].positions;

        SIZE_T i = 0;
        for (; i < linePositions.size(); ++i) {
            if (startPosition <= linePositions[i].position) {
                startColumnIndex = i + 1;
                break;
            }
        }

        // Check if startPosition is in the last column only if the loop ran to completion
        if (i == linePositions.size()) {
            startColumnIndex = linePositions.size() + 1;  // We're in the last column
        }

    }
    return { totalLines, startLine, startColumnIndex };
}

void MultiReplace::initializeColumnStyles() {

    int IDM_LANG_TEXT = 46016;  // Switch off Languages - > Normal Text
    ::SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_LANG_TEXT);

    for (SIZE_T column = 0; column < hColumnStyles.size(); column++) {
        int style = hColumnStyles[column];
        long color = columnColors[column];
        long fgColor = 0x000000;  // set Foreground always on black

        // Set the style background color
        ::SendMessage(_hScintilla, SCI_STYLESETBACK, style, color);

        // Set the style foreground color to black
        ::SendMessage(_hScintilla, SCI_STYLESETFORE, style, fgColor);

    }

}

void MultiReplace::handleHighlightColumnsInDocument() {
    // Return early if columnDelimiterData is empty
    if (columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty()) {
        return;
    }

    // Initialize column styles
    initializeColumnStyles();

    // Get the total number of lines
    LRESULT totalLines = static_cast<LRESULT>(lineDelimiterPositions.size());

    // Iterate over each line's delimiter positions
    for (LRESULT line = 0; line < totalLines; ++line) {
        highlightColumnsInLine(line);
    }

    // Show Row and Column Position
    if (!lineDelimiterPositions.empty()) {
        LRESULT startPosition = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
        showStatusMessage(getLangStr(L"status_actual_position", { addLineAndColumnMessage(startPosition) }), COLOR_SUCCESS);
    }

    isColumnHighlighted = true;
    isCaretPositionEnabled = true;  // Enable Position detection

}

void MultiReplace::highlightColumnsInLine(LRESULT line) {
    const auto& lineInfo = lineDelimiterPositions[line];

    // Check for empty line
    if (lineInfo.endPosition - lineInfo.startPosition == 0) {
        return; // It's an empty line, so exit early
    }

    // Prepare a vector for styles
    std::vector<char> styles((lineInfo.endPosition) - lineInfo.startPosition, 0);

    // If no delimiter present, highlight whole line as first column
    if (lineInfo.positions.empty() &&
        std::find(columnDelimiterData.columns.begin(), columnDelimiterData.columns.end(), 1) != columnDelimiterData.columns.end())
    {
        char style = static_cast<char>(hColumnStyles[0 % hColumnStyles.size()]);
        std::fill(styles.begin(), styles.end(), style);
    }
    else {
        // Highlight specific columns from columnDelimiterData
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= lineInfo.positions.size() + 1) {
                LRESULT start = 0;
                LRESULT end = 0;

                // Set start and end positions based on column index
                if (column == 1) {
                    start = 0;
                }
                else {
                    start = lineInfo.positions[column - 2].position + columnDelimiterData.delimiterLength - lineInfo.startPosition;
                }

                if (column == lineInfo.positions.size() + 1) {
                    end = (lineInfo.endPosition) - lineInfo.startPosition;
                }
                else {
                    end = lineInfo.positions[column - 1].position - lineInfo.startPosition;
                }

                // Apply style to the specific range within the styles vector
                char style = static_cast<char>(hColumnStyles[(column - 1) % hColumnStyles.size()]);
                std::fill(styles.begin() + start, styles.begin() + end, style);
            }

        }
    }

    send(SCI_STARTSTYLING, lineInfo.startPosition, 0);
    send(SCI_SETSTYLINGEX, styles.size(), reinterpret_cast<sptr_t>(&styles[0]));
}

void MultiReplace::handleClearColumnMarks() {
    LRESULT textLength = ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0);

    send(SCI_STARTSTYLING, 0, 0);
    send(SCI_SETSTYLING, textLength, STYLE_DEFAULT);

    isColumnHighlighted = false;

    // Disable Position detection
    isCaretPositionEnabled = false;

    // Force Scintilla to recalculate word wrapping as highlighting is affecting layout
    int originalWrapMode = static_cast<int>(::SendMessage(_hScintilla, SCI_GETWRAPMODE, 0, 0));
    if (originalWrapMode != SC_WRAP_NONE) {
        ::SendMessage(_hScintilla, SCI_SETWRAPMODE, SC_WRAP_NONE, 0);
        ::SendMessage(_hScintilla, SCI_SETWRAPMODE, originalWrapMode, 0);
    }
}

std::wstring MultiReplace::addLineAndColumnMessage(LRESULT pos) {
    if (!columnDelimiterData.isValid()) {
        return L"";
    }
    ColumnInfo startInfo = getColumnInfo(pos);
    std::wstring lineAndColumnMessage = getLangStr(L"status_line_and_column_position",
        { std::to_wstring(startInfo.startLine + 1),
          std::to_wstring(startInfo.startColumnIndex) });

    return lineAndColumnMessage;
}

void MultiReplace::processLogForDelimiters()
{
    // Check if logChanges is accessible
    if (!textModified || logChanges.empty()) {
        return;
    }

    std::vector<LogEntry> modifyLogEntries;

    // Loop through the log entries in chronological order
    for (auto& logEntry : logChanges) {
        switch (logEntry.changeType) {
        case ChangeType::Insert:
            for (auto& modifyLogEntry : modifyLogEntries) {
                // Check if the last entry in modifyLogEntries is one less than logEntry.lineNumber
                if (&modifyLogEntry == &modifyLogEntries.back() && modifyLogEntry.lineNumber == logEntry.lineNumber - 1) {
                    // Do nothing for the last entry if it is one less than logEntry.lineNumber as it has been produced by the Insert itself and should stay
                    continue;
                }
                if (modifyLogEntry.lineNumber >= logEntry.lineNumber - 1) {
                    ++modifyLogEntry.lineNumber;
                }
            }
            updateDelimitersInDocument(static_cast<int>(logEntry.lineNumber), ChangeType::Insert);
            updateUnsortedDocument(static_cast<int>(logEntry.lineNumber), ChangeType::Insert);
            // this->messageBoxContent += "Line " + std::to_string(static_cast<int>(logEntry.lineNumber)) + " inserted.\n";
            // Add Insert entry as a Modify entry in modifyLogEntries
            logEntry.changeType = ChangeType::Modify;  // Convert Insert to Modify
            modifyLogEntries.push_back(logEntry);
            break;
        case ChangeType::Delete:
            for (auto& modifyLogEntry : modifyLogEntries) {
                if (modifyLogEntry.lineNumber > logEntry.lineNumber) {
                    --modifyLogEntry.lineNumber;
                }
                else if (modifyLogEntry.lineNumber == logEntry.lineNumber) {
                    modifyLogEntry.lineNumber = -1;  // Mark for deletion
                }
            }
            updateDelimitersInDocument(static_cast<int>(logEntry.lineNumber), ChangeType::Delete);
            updateUnsortedDocument(static_cast<int>(logEntry.lineNumber), ChangeType::Delete);
            // this->messageBoxContent += "Line " + std::to_string(static_cast<int>(logEntry.lineNumber)) + " deleted.\n";
            break;
        case ChangeType::Modify:
            modifyLogEntries.push_back(logEntry);
            break;
        default:
            break;
        }
    }


    // Apply the saved "Modify" entries to the original delimiter list
    for (const auto& modifyLogEntry : modifyLogEntries) {
        if (modifyLogEntry.lineNumber != -1) {
            updateDelimitersInDocument(static_cast<int>(modifyLogEntry.lineNumber), ChangeType::Modify);
            if (isColumnHighlighted) {
                //clearMarksInLine(modifyLogEntry.lineNumber);
                highlightColumnsInLine(modifyLogEntry.lineNumber);
            }
            //this->messageBoxContent += "Line " + std::to_string(static_cast<int>(modifyLogEntry.lineNumber)) + " modified.\n";
        }
    }

    // Workaround: Highlight last line to fix N++ bug causing loss of styling on last character whwn modification in any other line
    if (isColumnHighlighted) {
        LRESULT lastLine = send(SCI_GETLINECOUNT, 0, 0) - 1;
        if (lastLine >= 0) {
            highlightColumnsInLine(lastLine);
        }
    }

    // Clear Log queue
    logChanges.clear();
    textModified = false;
}

void MultiReplace::updateDelimitersInDocument(SIZE_T lineNumber, ChangeType changeType) {

    if (lineNumber > lineDelimiterPositions.size()) {
        return; // invalid line number
    }

    LineInfo lineInfo;
    switch (changeType) {
    case ChangeType::Insert:
        // Insert an empty line at the specified index
        if (lineNumber > 0) { // not the first line
            lineInfo.startPosition = lineDelimiterPositions[lineNumber - 1].endPosition + eolLength;
            lineInfo.endPosition = lineInfo.startPosition;
        }
        else {
            lineInfo.startPosition = 0;
            lineInfo.endPosition = 0;
        }
        lineDelimiterPositions.insert(lineDelimiterPositions.begin() + lineNumber, lineInfo);
        break;

    case ChangeType::Delete:
        // Delete the specified line
        if (lineNumber < lineDelimiterPositions.size()) {
            // Calculate the length of the deleted line (including EOL)
            LRESULT deletedLineLength = lineDelimiterPositions[lineNumber].endPosition
                - lineDelimiterPositions[lineNumber].startPosition
                + eolLength;

            lineDelimiterPositions.erase(lineDelimiterPositions.begin() + lineNumber);

            // Update positions for subsequent lines
            for (SIZE_T i = lineNumber; i < lineDelimiterPositions.size(); ++i) {
                lineDelimiterPositions[i].startPosition -= deletedLineLength;
                lineDelimiterPositions[i].endPosition -= deletedLineLength;
                for (auto& delim : lineDelimiterPositions[i].positions) {
                    delim.position -= deletedLineLength;
                }
            }
        }
        break;

    case ChangeType::Modify:
        // Modify the content of the specified line
        if (lineNumber < lineDelimiterPositions.size()) {
            // Re-analyze the line to find delimiters
            findDelimitersInLine(lineNumber);

            // Only adjust following lines if not at the last line
            if (lineNumber < lineDelimiterPositions.size() - 1) {
                // Calculate the difference to the next line start position (considering EOL)
                LRESULT positionDifference = lineDelimiterPositions[lineNumber + 1].startPosition - lineDelimiterPositions[lineNumber].endPosition - eolLength;

                // Update positions for subsequent lines
                for (SIZE_T i = lineNumber + 1; i < lineDelimiterPositions.size(); ++i) {
                    if (positionDifference > 0) { // The distance is too large, need to reduce
                        lineDelimiterPositions[i].startPosition -= positionDifference;
                        lineDelimiterPositions[i].endPosition -= positionDifference;
                        for (auto& delim : lineDelimiterPositions[i].positions) {
                            delim.position -= positionDifference;
                        }
                    }
                    else if (positionDifference < 0) { // The distance is too small, need to increase
                        lineDelimiterPositions[i].startPosition += abs(positionDifference);
                        lineDelimiterPositions[i].endPosition += abs(positionDifference);
                        for (auto& delim : lineDelimiterPositions[i].positions) {
                            delim.position += abs(positionDifference);
                        }
                    }
                }
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
    //int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);
    LRESULT updatedEolLength = getEOLLength();

    // If EOL length has changed or if there's been a change in the active window within Notepad++, reset all delimiter settings
    if (updatedEolLength != eolLength || documentSwitched) {
        handleClearDelimiterState();
        documentSwitched = false;
    }

    eolLength = updatedEolLength;
    //scannedDelimiterBufferID = currentBufferID;

    if (operation == DelimiterOperation::LoadAll) {
        // Parse column and delimiter data; exit if parsing fails or if delimiter is empty
        if (!parseColumnAndDelimiterData()) {
            return;
        }

        // If any conditions that warrant a refresh of delimiters are met, proceed
        if (columnDelimiterData.isValid()) {
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
    auto stringToWString = [](const std::string& input) -> std::wstring {
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
    std::wstring wideStartContent = stringToWString(startContent);
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
    std::wstring logChangesWStr = stringToWString(logChangesStr);
    MessageBox(NULL, logChangesWStr.c_str(), L"Log Changes", MB_OK);

    processLogForDelimiters();
    std::wstring wideMessageBoxContent = stringToWString(messageBoxContent);
    MessageBox(NULL, wideMessageBoxContent.c_str(), L"Final Result", MB_OK);
    messageBoxContent.clear();

    std::string endContent = listToString(lineDelimiterPositions);
    std::wstring wideEndContent = stringToWString(endContent);
    MessageBox(NULL, wideEndContent.c_str(), L"Content at the end", MB_OK);

    logChanges.clear();
}
*/

#pragma endregion


#pragma region Utilities

int MultiReplace::convertExtendedToString(const std::string& query, std::string& result)
{
    auto readBase = [](const char* str, int* value, int base, int size) -> bool
        {
            int i = 0, temp = 0;
            *value = 0;
            char max = '0' + static_cast<char>(base) - 1;
            char current;
            while (i < size)
            {
                current = str[i];
                if (current >= 'A')
                {
                    current &= 0xdf;
                    current -= ('A' - '0' - 10);
                }
                else if (current > '9')
                    return false;

                if (current >= '0' && current <= max)
                {
                    temp *= base;
                    temp += (current - '0');
                }
                else
                {
                    return false;
                }
                ++i;
            }
            *value = temp;
            return true;
        };

    int i = 0, j = 0;
    int charLeft = static_cast<int>(query.length());
    char current;
    result.clear();
    result.resize(query.length()); // Preallocate memory for optimal performance

    while (i < static_cast<int>(query.length()))
    {
        current = query[i];
        --charLeft;
        if (current == '\\' && charLeft)
        {
            ++i;
            --charLeft;
            current = query[i];
            switch (current)
            {
            case 'r':
                result[j] = '\r';
                break;
            case 'n':
                result[j] = '\n';
                break;
            case '0':
                result[j] = '\0';
                break;
            case 't':
                result[j] = '\t';
                break;
            case '\\':
                result[j] = '\\';
                break;
            case 'b':
            case 'd':
            case 'o':
            case 'x':
            case 'u':
            {
                int size = 0, base = 0;
                if (current == 'b')
                {
                    size = 8, base = 2;
                }
                else if (current == 'o')
                {
                    size = 3, base = 8;
                }
                else if (current == 'd')
                {
                    size = 3, base = 10;
                }
                else if (current == 'x')
                {
                    size = 2, base = 16;
                }
                else if (current == 'u')
                {
                    size = 4, base = 16;
                }

                if (charLeft >= size)
                {
                    int res = 0;
                    if (readBase(query.c_str() + (i + 1), &res, base, size))
                    {
                        result[j] = static_cast<char>(res);
                        i += size;
                        break;
                    }
                }
                // not enough chars to make parameter, use default method as fallback
            }
            [[fallthrough]];
            default:
                // unknown sequence, treat as regular text
                result[j] = '\\';
                ++j;
                result[j] = current;
                break;
            }
        }
        else
        {
            result[j] = query[i];
        }
        ++i;
        ++j;
    }

    // Nullterminate the result-String
    result.resize(j);

    // Return length of result-Strings
    return j;
}

std::string MultiReplace::convertAndExtend(const std::wstring& input, bool extended)
{
    std::string output = wstringToString(input);

    if (extended)
    {
        std::string outputExtended;
        convertExtendedToString(output, outputExtended);
        output = outputExtended;
    }

    return output;
}

std::string MultiReplace::convertAndExtend(const std::string& input, bool extended)
{
    std::string output = input;

    if (extended)
    {
        std::string outputExtended;
        convertExtendedToString(output, outputExtended);
        output = outputExtended;
    }

    return output;
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

std::wstring MultiReplace::getTextFromDialogItem(HWND hwnd, int itemID) {
    int textLen = GetWindowTextLength(GetDlgItem(hwnd, itemID));
    // Limit the length of the text to the maximum value
    textLen = std::min(textLen, MAX_TEXT_LENGTH);
    std::vector<TCHAR> buffer(textLen + 1);  // +1 for null terminator
    GetDlgItemText(hwnd, itemID, &buffer[0], textLen + 1);

    // Construct the std::wstring from the buffer excluding the null character
    return std::wstring(buffer.begin(), buffer.begin() + textLen);
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

    // Clear the redoStack
    redoStack.clear();

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
    undoStack.push_back(action);
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
            headerText = getLangStr(L"header_find_count");
            break;
        case ColumnID::REPLACE_COUNT:
            headerText = getLangStr(L"header_replace_count");
            break;
        case ColumnID::FIND_TEXT:
            headerText = getLangStr(L"header_find");
            // Append lock symbol if the FIND_TEXT column is locked
            if (findColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::REPLACE_TEXT:
            headerText = getLangStr(L"header_replace");
            // Append lock symbol if the REPLACE_TEXT column is locked
            if (replaceColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::COMMENTS:
            headerText = getLangStr(L"header_comments");
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

void MultiReplace::showStatusMessage(const std::wstring& messageText, COLORREF color, bool isNotFound)
{
    const size_t MAX_DISPLAY_LENGTH = 150;  // Maximum length of the message to be displayed

    // Filter out non-printable characters while keeping all printable Unicode characters
    std::wstring strMessage;
    for (wchar_t ch : messageText) {
        if (iswprint(ch)) {
            strMessage += ch;
        }
    }

    if (strMessage.size() > MAX_DISPLAY_LENGTH) {
        strMessage = strMessage.substr(0, MAX_DISPLAY_LENGTH - 3) + L"...";
    }

    // Get the handle for the status message control
    HWND hStatusMessage = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);

    // Set the new message
    _statusMessageColor = color;
    SetWindowText(hStatusMessage, strMessage.c_str());

    // Invalidate the area of the parent where the control resides
    RECT rect;
    GetWindowRect(hStatusMessage, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hStatusMessage), (LPPOINT)&rect, 2);
    InvalidateRect(GetParent(hStatusMessage), &rect, TRUE);
    UpdateWindow(GetParent(hStatusMessage));

    if (isNotFound && alertNotFoundEnabled)
    {
        // Play the default beep sound
        MessageBeep(MB_ICONEXCLAMATION);

        // Flash the window caption and taskbar button
        FLASHWINFO fwInfo = { 0 };
        fwInfo.cbSize = sizeof(FLASHWINFO);
        fwInfo.hwnd = _hSelf;          // Handle to your window
        fwInfo.dwFlags = FLASHW_ALL;   // Flash both the caption and the taskbar button
        fwInfo.uCount = 3;             // Number of times to flash the window
        fwInfo.dwTimeout = 100;          // Default cursor blink rate

        FlashWindowEx(&fwInfo);
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

void MultiReplace::showListFilePath() { 
    std::wstring path = listFilePath;

    // Obtain handle and device context for the path display control
    HWND hPathDisplay = GetDlgItem(_hSelf, IDC_PATH_DISPLAY);
    HDC hDC = GetDC(hPathDisplay);
    HFONT hFont = (HFONT)SendMessage(hPathDisplay, WM_GETFONT, 0, 0);
    SelectObject(hDC, hFont);

    // Get display width for IDC_PATH_DISPLAY
    RECT rcPathDisplay;
    GetClientRect(hPathDisplay, &rcPathDisplay);
    int pathDisplayWidth = rcPathDisplay.right - rcPathDisplay.left;

    // Call the new function to get the shortened file path
    std::wstring shortenedPath = getShortenedFilePath(path, pathDisplayWidth, hDC);

    // Display the shortened path
    SetWindowTextW(hPathDisplay, shortenedPath.c_str());

    // Update the parent window area where the control resides
    RECT rect;
    GetWindowRect(hPathDisplay, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hPathDisplay), (LPPOINT)&rect, 2);
    InvalidateRect(GetParent(hPathDisplay), &rect, TRUE);
    UpdateWindow(GetParent(hPathDisplay));
}

std::wstring MultiReplace::getSelectedText() {
    SIZE_T length = SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0, 0);

    if (length > MAX_TEXT_LENGTH) {
        return L"";
    }

    char* buffer = new char[length + 1];  // Add 1 for null terminator
    SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0, (LPARAM)buffer);
    buffer[length] = '\0';

    std::string str(buffer);
    std::wstring wstr = stringToWString(str);

    delete[] buffer;

    return wstr;
}

LRESULT MultiReplace::getEOLLength() {
    LRESULT eolMode = ::SendMessage(getScintillaHandle(), SCI_GETEOLMODE, 0, 0);
    switch (eolMode) {
    case SC_EOL_CRLF:
        return 2;
    case SC_EOL_CR:
    case SC_EOL_LF:
        return 1;
    default:
        return 2; // Default to CRLF
    }
}

std::string MultiReplace::getEOLStyle() {
    LRESULT eolMode = SendMessage(_hScintilla, SCI_GETEOLMODE, 0, 0);
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

sptr_t MultiReplace::send(unsigned int iMessage, uptr_t wParam, sptr_t lParam, bool useDirect) {
    if (useDirect && pSciMsg) {
        return pSciMsg(pSciWndData, iMessage, wParam, lParam);
    }
    else {
        return ::SendMessage(_hScintilla, iMessage, wParam, lParam);
    }
}

/*
sptr_t MultiReplace::send(unsigned int iMessage, uptr_t wParam, sptr_t lParam, bool useDirect) {
    sptr_t result;

    if (useDirect && pSciMsg) {
        result = pSciMsg(pSciWndData, iMessage, wParam, lParam);
    }
    else {
        result = ::SendMessage(_hScintilla, iMessage, wParam, lParam);
    }

    // Check Scintilla's error status
    int status = static_cast<int>(::SendMessage(_hScintilla, SCI_GETSTATUS, 0, 0));

    if (status != SC_STATUS_OK) {
        wchar_t buffer[512];
        switch (status) {
        case SC_STATUS_FAILURE:
            wcscpy(buffer, L"Error: Generic failure");
            break;
        case SC_STATUS_BADALLOC:
            wcscpy(buffer, L"Error: Memory is exhausted");
            break;
        case SC_STATUS_WARN_REGEX:
            wcscpy(buffer, L"Warning: Regular expression is invalid");
            break;
        default:
            swprintf(buffer, L"Error/Warning with status code: %d", status);
            break;
        }

        // Append the function call details
        wchar_t callDetails[512];
        #if defined(_WIN64)
            swprintf(callDetails, L"\niMessage: %u\nwParam: %llu\nlParam: %lld", iMessage, static_cast<unsigned long long>(wParam), static_cast<long long>(lParam));
        #else
            swprintf(callDetails, L"\niMessage: %u\nwParam: %lu\nlParam: %ld", iMessage, static_cast<unsigned long>(wParam), static_cast<long>(lParam));
        #endif


        wcscat(buffer, callDetails);

        MessageBox(NULL, buffer, L"Scintilla Error/Warning", MB_OK | (status >= SC_STATUS_WARN_START ? MB_ICONWARNING : MB_ICONERROR));

        // Clear the error status
        ::SendMessage(_hScintilla, SCI_SETSTATUS, SC_STATUS_OK, 0);
    }

    return result;
}
*/

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
            dotCount++;
        }
        else if (c == ',') {
            dotCount++;
            c = '.';  // Potentially replace comma with dot in tempStr
        }
        else if (!isdigit(c)) {
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

#pragma endregion


#pragma region StringHandling

std::wstring MultiReplace::stringToWString(const std::string& rString) const {
    int codePage = static_cast<int>(::SendMessage(_hScintilla, SCI_GETCODEPAGE, 0, 0));

    int requiredSize = MultiByteToWideChar(codePage, 0, rString.c_str(), -1, NULL, 0);
    if (requiredSize == 0)
        return std::wstring();

    std::vector<wchar_t> wideStringResult(requiredSize);
    MultiByteToWideChar(codePage, 0, rString.c_str(), -1, &wideStringResult[0], requiredSize);

    return std::wstring(&wideStringResult[0]);
}

std::string MultiReplace::wstringToString(const std::wstring& input) const {
    if (input.empty()) return std::string();

    int codePage = static_cast<int>(::SendMessage(_hScintilla, SCI_GETCODEPAGE, 0, 0));
    if (codePage == 0) codePage = CP_ACP;

    int size_needed = WideCharToMultiByte(codePage, 0, &input[0], (int)input.size(), NULL, 0, NULL, NULL);
    if (size_needed == 0) return std::string();

    std::string strResult(size_needed, 0);
    WideCharToMultiByte(codePage, 0, &input[0], (int)input.size(), &strResult[0], size_needed, NULL, NULL);

    return strResult;
}

std::wstring MultiReplace::utf8ToWString(const char* cstr) const {
    if (cstr == nullptr) {
        return std::wstring();
    }

    int requiredSize = MultiByteToWideChar(CP_UTF8, 0, cstr, -1, NULL, 0);
    if (requiredSize == 0) {
        return std::wstring();
    }

    std::vector<wchar_t> wideStringResult(requiredSize);
    MultiByteToWideChar(CP_UTF8, 0, cstr, -1, &wideStringResult[0], requiredSize);

    return std::wstring(&wideStringResult[0]);
}

std::string MultiReplace::utf8ToCodepage(const std::string& utf8Str, int codepage) const {
    // Convert the UTF-8 string to a wide string
    int lenWc = MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(), -1, nullptr, 0);
    if (lenWc == 0) {
        // Handle error
        return std::string();
    }
    std::vector<wchar_t> wideStr(lenWc);
    MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(), -1, &wideStr[0], lenWc);

    // Convert the wide string to the specific codepage
    int lenMbcs = WideCharToMultiByte(codepage, 0, &wideStr[0], -1, nullptr, 0, nullptr, nullptr);
    if (lenMbcs == 0) {
        // Handle error
        return std::string();
    }
    std::vector<char> cpStr(lenMbcs);
    WideCharToMultiByte(codepage, 0, &wideStr[0], -1, &cpStr[0], lenMbcs, nullptr, nullptr);

    return std::string(cpStr.data(), lenMbcs - 1);  // -1 to exclude the null character
}

std::wstring MultiReplace::trim(const std::wstring& str) {
    // Find the first character that is not whitespace, tab, newline, or carriage return
    const auto strBegin = str.find_first_not_of(L" \t\n\r");

    if (strBegin == std::wstring::npos) {
        // If the entire string consists of whitespace, return an empty string
        return L"";
    }

    // Find the last character that is not whitespace, tab, newline, or carriage return
    const auto strEnd = str.find_last_not_of(L" \t\n\r");

    // Calculate the range of non-whitespace characters
    const auto strRange = strEnd - strBegin + 1;

    // Return the substring without leading and trailing whitespace
    return str.substr(strBegin, strRange);
}

#pragma endregion


#pragma region FileOperations

std::wstring MultiReplace::openFileDialog(bool saveFile, const std::vector<std::pair<std::wstring, std::wstring>>& filters, const WCHAR* title, DWORD flags, const std::wstring& fileExtension, const std::wstring& defaultFilePath) {
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
    std::wstring csvDescription = getLangStr(L"filetype_csv");  // "CSV Files (*.csv)"
    std::wstring allFilesDescription = getLangStr(L"filetype_all_files");  // "All Files (*.*)"

    std::vector<std::pair<std::wstring, std::wstring>> filters = {
        {csvDescription, L"*.csv"},
        {allFilesDescription, L"*.*"}
    };

    std::wstring dialogTitle = getLangStr(L"panel_save_list");
    std::wstring defaultFileName;

    if (!listFilePath.empty()) {
        // If a file path already exists, use its directory and filename
        defaultFileName = listFilePath;
    }
    else {
        // If no file path is set, provide a default file name with a sequential number
        static int fileCounter = 1;
        defaultFileName = L"Replace_List_" + std::to_wstring(fileCounter++) + L".csv";
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
    std::ofstream outFile(filePath);

    if (!outFile.is_open()) {
        return false;
    }

    // Convert and Write CSV header
    std::string utf8Header = wstringToString(L"Selected,Find,Replace,WholeWord,MatchCase,UseVariables,Regex,Extended,Comments\n");
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
        std::string utf8Line = wstringToString(line);
        outFile << utf8Line;
    }

    outFile.close();

    return !outFile.fail();;
}

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (!saveListToCsvSilent(filePath, list)) {
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), COLOR_ERROR);
        return;
    }

    showStatusMessage(getLangStr(L"status_saved_items_to_csv", { std::to_wstring(list.size()) }), COLOR_SUCCESS);

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
            message = getLangStr(L"msgbox_unsaved_changes_file", { shortenedFilePath });
        }
        else {
            // If no file is associated, use the alternative message
            message = getLangStr(L"msgbox_unsaved_changes");
        }

        // Show the MessageBox with the appropriate message
        int result = MessageBox(
            nppData._nppHandle,
            message.c_str(),
            getLangStr(L"msgbox_title_save_list").c_str(),
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
    // Open file in binary mode to read UTF-8 data
    std::ifstream inFile(filePath);
    if (!inFile.is_open()) {
        std::wstring shortenedFilePathW = getShortenedFilePath(filePath, 500);
        std::string errorMessage = wstringToString(getLangStr(L"status_unable_to_open_file", { shortenedFilePathW }));
        throw CsvLoadException(errorMessage);
    }

    std::vector<ReplaceItemData> tempList;  // Temporary list to hold items
    std::string utf8Content((std::istreambuf_iterator<char>(inFile)), std::istreambuf_iterator<char>());
    std::wstring content = stringToWString(utf8Content);
    std::wstringstream contentStream(content);

    std::wstring line;
    std::getline(contentStream, line); // Skip the CSV header

    while (std::getline(contentStream, line)) {
        std::vector<std::wstring> columns = parseCsvLine(line);

        if (columns.size() < 8 || columns.size() > 9) {
            throw CsvLoadException(wstringToString(getLangStr(L"status_invalid_column_count")));
        }

        ReplaceItemData item;
        try {
            item.isEnabled = std::stoi(columns[0]) != 0;
            item.findText = columns[1];
            item.replaceText = columns[2];
            item.wholeWord = std::stoi(columns[3]) != 0;
            item.matchCase = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended = std::stoi(columns[6]) != 0;
            item.regex = std::stoi(columns[7]) != 0;

            // Handle Comments column for compatibility with old format
            if (columns.size() == 9) {
                item.comments = columns[8];
            }
            else {
                item.comments = L"";  // Initialize as empty string if not present
            }

            tempList.push_back(item);
        }
        catch (const std::exception&) {
            throw CsvLoadException(wstringToString(getLangStr(L"status_invalid_data_in_columns")));
        }
    }

    inFile.close();
    list = tempList;  // Transfer data from temporary list to the final list
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {

    try {
        loadListFromCsvSilent(filePath, replaceListData);

        // Store the file path only if loading was successful
        listFilePath = filePath;

        // Display the path below the list
        showListFilePath();

        // Calculate the original list hash after loading
        originalListHash = computeListHash(replaceListData);

        // Clear the Undo and Redo stacks after successful load
        undoStack.clear();
        redoStack.clear();

        // Update the list view control
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, NULL, TRUE);

        // Show success message
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_no_valid_items_in_csv"), COLOR_ERROR);
        }
        else {
            showStatusMessage(getLangStr(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), COLOR_SUCCESS);
        }
    }
    catch (const CsvLoadException& ex) {
        // Resolve the error key to a localized string when displaying the message
        showStatusMessage(stringToWString(ex.what()), COLOR_ERROR);
        return;
    }
}

void MultiReplace::checkForFileChangesAtStartup() {
    if (listFilePath.empty()) {
        return;
    }

    std::wstring shortenedFilePath = getShortenedFilePath(listFilePath, 500);

    try {
        std::vector<ReplaceItemData> tempListFromFile;
        loadListFromCsvSilent(listFilePath, tempListFromFile);  // Load the list into a temporary list

        std::size_t newFileHash = computeListHash(tempListFromFile);  // Calculate the new file hash

        // Check if the file has been modified externally
        if (newFileHash != originalListHash) {
            std::wstring message = getLangStr(L"msgbox_file_modified_prompt", { shortenedFilePath });

            int response = MessageBox(
                nppData._nppHandle,
                message.c_str(),
                getLangStr(L"msgbox_title_reload").c_str(),
                MB_YESNO | MB_ICONWARNING | MB_SETFOREGROUND
            );

            if (response == IDYES) {
                replaceListData = tempListFromFile;
                originalListHash = newFileHash;
            }
        }
    }
    catch (const CsvLoadException& ex) {
        // Resolve the error key to a localized string when displaying the message
        showStatusMessage(stringToWString(ex.what()), COLOR_ERROR);
        return;
    }

    if (replaceListData.empty()) {
        showStatusMessage(getLangStr(L"status_no_valid_items_in_csv"), COLOR_ERROR);
    }
    else {
        showStatusMessage(getLangStr(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), COLOR_SUCCESS);

        // Update the list view control, if necessary
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

    escapedValue += L"\""; // Schließt den String mit einem Anführungszeichen ab

    return escapedValue;
}

std::wstring MultiReplace::unescapeCsvValue(const std::wstring& value) {
    std::wstring unescapedValue;
    if (value.empty()) {
        return unescapedValue;
    }

    size_t start = (value.front() == L'"' && value.back() == L'"') ? 1 : 0;
    size_t end = (start == 1) ? value.size() - 1 : value.size();

    for (size_t i = start; i < end; ++i) {
        if (i < end - 1 && value[i] == L'\\') {
            switch (value[i + 1]) {
            case L'n': unescapedValue += L'\n'; ++i; break;
            case L'r': unescapedValue += L'\r'; ++i; break;
            case L'\\': unescapedValue += L'\\'; ++i; break;
            default: unescapedValue += value[i]; break;
            }
        }
        else if (i < end - 1 && value[i] == L'"' && value[i + 1] == L'"') {
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
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), COLOR_ERROR);
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
            find = replaceNewline(translateEscapes(escapeSpecialChars(wstringToString(itemData.findText), true)), ReplaceMode::Extended);
            replace = replaceNewline(translateEscapes(escapeSpecialChars(wstringToString(itemData.replaceText), true)), ReplaceMode::Extended);
        }
        else if (itemData.regex) {
            find = replaceNewline(wstringToString(itemData.findText), ReplaceMode::Regex);
            replace = replaceNewline(wstringToString(itemData.replaceText), ReplaceMode::Regex);
        }
        else {
            find = replaceNewline(escapeSpecialChars(wstringToString(itemData.findText), false), ReplaceMode::Normal);
            replace = replaceNewline(escapeSpecialChars(wstringToString(itemData.replaceText), false), ReplaceMode::Normal);
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
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), COLOR_ERROR);
        return;
    }

    showStatusMessage(getLangStr(L"status_list_exported_to_bash"), COLOR_SUCCESS);

    // Show message box if excluded items were found
    if (hasExcludedItems) {
        MessageBox(_hSelf,
            getLangStr(L"msgbox_use_variables_not_exported").c_str(),
            getLangStr(L"msgbox_title_warning").c_str(),
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
        std::string result = wstringToString(unicodeString);
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
    std::ofstream outFile(iniFilePath);

    if (!outFile.is_open()) {
        throw std::runtime_error("Could not open settings file for writing.");
    }

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

    outFile << wstringToString(L"[Window]\n");
    outFile << wstringToString(L"PosX=" + std::to_wstring(posX) + L"\n");
    outFile << wstringToString(L"PosY=" + std::to_wstring(posY) + L"\n");
    outFile << wstringToString(L"Width=" + std::to_wstring(width) + L"\n");
    outFile << wstringToString(L"Height=" + std::to_wstring(useListOnHeight) + L"\n");
    outFile << wstringToString(L"ScaleFactor=" + std::to_wstring(dpiMgr->getCustomScaleFactor()).substr(0, std::to_wstring(dpiMgr->getCustomScaleFactor()).find(L'.') + 2) + L"\n");

    // Save transparency settings
    outFile << wstringToString(L"ForegroundTransparency=" + std::to_wstring(foregroundTransparency) + L"\n");
    outFile << wstringToString(L"BackgroundTransparency=" + std::to_wstring(backgroundTransparency) + L"\n");

    // Store column widths for "Find Count", "Replace Count", and "Comments"
    findCountColumnWidth = (columnIndices[ColumnID::FIND_COUNT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_COUNT]) : findCountColumnWidth;
    replaceCountColumnWidth = (columnIndices[ColumnID::REPLACE_COUNT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_COUNT]) : replaceCountColumnWidth;
    findColumnWidth = (columnIndices[ColumnID::FIND_TEXT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_TEXT]) : findColumnWidth;
    replaceColumnWidth = (columnIndices[ColumnID::REPLACE_TEXT] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_TEXT]) : replaceColumnWidth;
    commentsColumnWidth = (columnIndices[ColumnID::COMMENTS] != -1) ? ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::COMMENTS]) : commentsColumnWidth;

    outFile << wstringToString(L"[ListColumns]\n");
    outFile << wstringToString(L"FindCountWidth=" + std::to_wstring(findCountColumnWidth) + L"\n");
    outFile << wstringToString(L"ReplaceCountWidth=" + std::to_wstring(replaceCountColumnWidth) + L"\n");
    outFile << wstringToString(L"FindWidth=" + std::to_wstring(findColumnWidth) + L"\n");
    outFile << wstringToString(L"ReplaceWidth=" + std::to_wstring(replaceColumnWidth) + L"\n");
    outFile << wstringToString(L"CommentsWidth=" + std::to_wstring(commentsColumnWidth) + L"\n");

    // Save column visibility states
    outFile << wstringToString(L"FindCountVisible=" + std::to_wstring(isFindCountVisible) + L"\n");
    outFile << wstringToString(L"ReplaceCountVisible=" + std::to_wstring(isReplaceCountVisible) + L"\n");
    outFile << wstringToString(L"CommentsVisible=" + std::to_wstring(isCommentsColumnVisible) + L"\n");
    outFile << wstringToString(L"DeleteButtonVisible=" + std::to_wstring(isDeleteButtonVisible ? 1 : 0) + L"\n");

    // Save column lock states
    outFile << wstringToString(L"FindColumnLocked=" + std::to_wstring(findColumnLockedEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"ReplaceColumnLocked=" + std::to_wstring(replaceColumnLockedEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"CommentsColumnLocked=" + std::to_wstring(commentsColumnLockedEnabled ? 1 : 0) + L"\n");

    // Convert and Store the current "Find what" and "Replace with" texts
    std::wstring currentFindTextData = escapeCsvValue(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT));
    std::wstring currentReplaceTextData = escapeCsvValue(getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT));

    outFile << wstringToString(L"[Current]\n");
    outFile << wstringToString(L"FindText=" + currentFindTextData + L"\n");
    outFile << wstringToString(L"ReplaceText=" + currentReplaceTextData + L"\n");

    // Prepare and Store the current options
    int wholeWord = IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int matchCase = IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int extended = IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED ? 1 : 0;
    int regex = IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? 1 : 0;
    int replaceFirst = IsDlgButtonChecked(_hSelf, IDC_REPLACE_FIRST_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int wrapAround = IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int useVariables = IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int ButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED ? 1 : 0;
    int useList = useListEnabled ? 1 : 0;

    outFile << wstringToString(L"[Options]\n");
    outFile << wstringToString(L"WholeWord=" + std::to_wstring(wholeWord) + L"\n");
    outFile << wstringToString(L"MatchCase=" + std::to_wstring(matchCase) + L"\n");
    outFile << wstringToString(L"Extended=" + std::to_wstring(extended) + L"\n");
    outFile << wstringToString(L"Regex=" + std::to_wstring(regex) + L"\n");
    outFile << wstringToString(L"ReplaceFirst=" + std::to_wstring(replaceFirst) + L"\n");
    outFile << wstringToString(L"WrapAround=" + std::to_wstring(wrapAround) + L"\n");
    outFile << wstringToString(L"UseVariables=" + std::to_wstring(useVariables) + L"\n");
    outFile << wstringToString(L"ButtonsMode=" + std::to_wstring(ButtonsMode) + L"\n");
    outFile << wstringToString(L"UseList=" + std::to_wstring(useList) + L"\n");
    outFile << wstringToString(L"HighlightMatch=" + std::to_wstring(highlightMatchEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"Tooltips=" + std::to_wstring(tooltipsEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"AlertNotFound=" + std::to_wstring(alertNotFoundEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"DoubleClickEdits=" + std::to_wstring(doubleClickEditsEnabled ? 1 : 0) + L"\n");
    outFile << wstringToString(L"HoverText=" + std::to_wstring(isHoverTextEnabled ? 1 : 0) + L"\n");

    // Convert and Store the scope options
    int selection = IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED ? 1 : 0;
    int columnMode = IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED ? 1 : 0;
    std::wstring columnNum = L"\"" + getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT) + L"\"";
    std::wstring delimiter = L"\"" + getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT) + L"\"";
    std::wstring quoteChar = L"\"" + getTextFromDialogItem(_hSelf, IDC_QUOTECHAR_EDIT) + L"\"";
    std::wstring headerLines = std::to_wstring(CSVheaderLinesCount);

    outFile << wstringToString(L"[Scope]\n");
    outFile << wstringToString(L"Selection=" + std::to_wstring(selection) + L"\n");
    outFile << wstringToString(L"ColumnMode=" + std::to_wstring(columnMode) + L"\n");
    outFile << wstringToString(L"ColumnNum=" + columnNum + L"\n");
    outFile << wstringToString(L"Delimiter=" + delimiter + L"\n");
    outFile << wstringToString(L"QuoteChar=" + quoteChar + L"\n");
    outFile << wstringToString(L"HeaderLines=" + headerLines + L"\n");

    // Save the list file path and original hash
    outFile << wstringToString(L"[File]\n");
    outFile << wstringToString(L"ListFilePath=" + listFilePath + L"\n");
    outFile << wstringToString(L"OriginalListHash=" + std::to_wstring(originalListHash) + L"\n");

    // Save the "Find what" history
    LRESULT findWhatCount = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETCOUNT, 0, 0);
    int itemsToSave = std::min(static_cast<int>(findWhatCount), maxHistoryItems);
    outFile << wstringToString(L"[History]\n");
    outFile << wstringToString(L"FindTextHistoryCount=" + std::to_wstring(itemsToSave) + L"\n");

    // Save only the newest maxHistoryItems entries (starting from index 0)
    for (LRESULT i = 0; i < itemsToSave; i++) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring findTextData = escapeCsvValue(std::wstring(buffer.data()));
        outFile << wstringToString(L"FindTextHistory" + std::to_wstring(i) + L"=" + findTextData + L"\n");
    }

    // Save the "Replace with" history
    LRESULT replaceWithCount = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETCOUNT, 0, 0);
    int replaceItemsToSave = std::min(static_cast<int>(replaceWithCount), maxHistoryItems);
    outFile << wstringToString(L"ReplaceTextHistoryCount=" + std::to_wstring(replaceItemsToSave) + L"\n");

    // Save only the newest maxHistoryItems entries (starting from index 0)
    for (LRESULT i = 0; i < replaceItemsToSave; i++) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring replaceTextData = escapeCsvValue(std::wstring(buffer.data()));
        outFile << wstringToString(L"ReplaceTextHistory" + std::to_wstring(i) + L"=" + replaceTextData + L"\n");
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
        std::wstring errorMessage = getLangStr(L"msgbox_error_saving_settings", { std::wstring(ex.what(), ex.what() + strlen(ex.what())) });
        MessageBox(nppData._nppHandle, errorMessage.c_str(), getLangStr(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
    }
    settingsSaved = true;
}

void MultiReplace::loadSettingsFromIni(const std::wstring& iniFilePath) {
    // Loading the history for the Find text field in reverse order
    int findHistoryCount = readIntFromIniFile(iniFilePath, L"History", L"FindTextHistoryCount", 0);
    for (int i = findHistoryCount - 1; i >= 0; i--) {
        std::wstring findHistoryItem = readStringFromIniFile(iniFilePath, L"History", L"FindTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findHistoryItem);
    }

    // Loading the history for the Replace text field in reverse order
    int replaceHistoryCount = readIntFromIniFile(iniFilePath, L"History", L"ReplaceTextHistoryCount", 0);
    for (int i = replaceHistoryCount - 1; i >= 0; i--) {
        std::wstring replaceHistoryItem = readStringFromIniFile(iniFilePath, L"History", L"ReplaceTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceHistoryItem);
    }

    // Reading and setting current "Find what" and "Replace with" texts
    std::wstring findText = readStringFromIniFile(iniFilePath, L"Current", L"FindText", L"");
    std::wstring replaceText = readStringFromIniFile(iniFilePath, L"Current", L"ReplaceText", L"");
    setTextInDialogItem(_hSelf, IDC_FIND_EDIT, findText);
    setTextInDialogItem(_hSelf, IDC_REPLACE_EDIT, replaceText);
    
    // Setting options based on the INI file
    bool wholeWord = readBoolFromIniFile(iniFilePath, L"Options", L"WholeWord", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);

    bool matchCase = readBoolFromIniFile(iniFilePath, L"Options", L"MatchCase", false);
    SendMessage(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, matchCase ? BST_CHECKED : BST_UNCHECKED, 0);

    bool useVariables = readBoolFromIniFile(iniFilePath, L"Options", L"UseVariables", false);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, useVariables ? BST_CHECKED : BST_UNCHECKED, 0);

    // Selecting the appropriate search mode radio button based on the settings
    bool extended = readBoolFromIniFile(iniFilePath, L"Options", L"Extended", false);
    bool regex = readBoolFromIniFile(iniFilePath, L"Options", L"Regex", false);
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
    bool wrapAround = readBoolFromIniFile(iniFilePath, L"Options", L"WrapAround", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WRAP_AROUND_CHECKBOX), BM_SETCHECK, wrapAround ? BST_CHECKED : BST_UNCHECKED, 0);

    bool replaceFirst = readBoolFromIniFile(iniFilePath, L"Options", L"ReplaceFirst", false);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_FIRST_CHECKBOX), BM_SETCHECK, replaceFirst ? BST_CHECKED : BST_UNCHECKED, 0);

    bool replaceButtonsMode = readBoolFromIniFile(iniFilePath, L"Options", L"ButtonsMode", false);
    SendMessage(GetDlgItem(_hSelf, IDC_2_BUTTONS_MODE), BM_SETCHECK, replaceButtonsMode ? BST_CHECKED : BST_UNCHECKED, 0);
    
    useListEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"UseList", true);
    updateUseListState(false);

    highlightMatchEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"HighlightMatch", true);
    alertNotFoundEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"AlertNotFound", true);
    doubleClickEditsEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"DoubleClickEdits", true);
    isHoverTextEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"HoverText", true);

    // Loading and setting the scope with enabled state check
    int selection = readIntFromIniFile(iniFilePath, L"Scope", L"Selection", 0);
    int columnMode = readIntFromIniFile(iniFilePath, L"Scope", L"ColumnMode", 0);

    // Reading and setting specific scope settings
    std::wstring columnNum = readStringFromIniFile(iniFilePath, L"Scope", L"ColumnNum", L"1-50");
    setTextInDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT, columnNum);

    std::wstring delimiter = readStringFromIniFile(iniFilePath, L"Scope", L"Delimiter", L",");
    setTextInDialogItem(_hSelf, IDC_DELIMITER_EDIT, delimiter);

    std::wstring quoteChar = readStringFromIniFile(iniFilePath, L"Scope", L"QuoteChar", L"\"");
    setTextInDialogItem(_hSelf, IDC_QUOTECHAR_EDIT, quoteChar);

    CSVheaderLinesCount = readIntFromIniFile(iniFilePath, L"Scope", L"HeaderLines", 1);

    // Load file path and original hash from the INI file
    listFilePath = readStringFromIniFile(iniFilePath, L"File", L"ListFilePath", L"");
    originalListHash = readSizeTFromIniFile(iniFilePath, L"File", L"OriginalListHash", 0);
    showListFilePath();

    // Adjusting UI elements based on the selected scope


    if (selection) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_SELECTION_RADIO);
        onSelectionChanged(); // check selection for IDC_SELECTION_RADIO
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
    // Generate the paths to the configuration files
    auto [iniFilePath, csvFilePath] = generateConfigFilePaths();

    try {
        loadSettingsFromIni(iniFilePath);        
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

}

void MultiReplace::loadUIConfigFromIni() {
    auto [iniFilePath, _] = generateConfigFilePaths(); // Generating config file paths

    // Load DPI Scaling factor from INI file
    float customScaleFactor = readFloatFromIniFile(iniFilePath, L"Window", L"ScaleFactor", 1.0f);
    dpiMgr->setCustomScaleFactor(customScaleFactor);

    // Scale Window and List Size after loading ScaleFactor
    MIN_WIDTH_scaled = sx(MIN_WIDTH);               // MIN_WIDTH from resource.rc
    MIN_HEIGHT_scaled = sy(MIN_HEIGHT);              // MIN_HEIGHT from resource.rc
    SHRUNK_HEIGHT_scaled = sy(SHRUNK_HEIGHT);          // SHRUNK_HEIGHT from resource.rc
    DEFAULT_COLUMN_WIDTH_FIND_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND);   // Scaled Default Size of Find Column
    DEFAULT_COLUMN_WIDTH_REPLACE_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE); // Scaled Default Size of Replace Column
    DEFAULT_COLUMN_WIDTH_COMMENTS_scaled = sx(DEFAULT_COLUMN_WIDTH_COMMENTS); // Scaled Default Size of Comments Column
    DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND_COUNT); // Scaled Default Size of Find Count Column
    DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE_COUNT); // Scaled Default Size of Replace Count Column
    MIN_GENERAL_WIDTH_scaled = sx(MIN_GENERAL_WIDTH);

    // Load window position
    windowRect.left = readIntFromIniFile(iniFilePath, L"Window", L"PosX", POS_X);
    windowRect.top = readIntFromIniFile(iniFilePath, L"Window", L"PosY", POS_Y);

    // Load the state of the Use List checkbox from the ini file
    useListEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"UseList", true); // Default to true if not found
    updateUseListState(false);

    // Load window width
    int savedWidth = readIntFromIniFile(iniFilePath, L"Window", L"Width", MIN_WIDTH_scaled);
    int width = std::max(savedWidth, MIN_WIDTH_scaled);

    // Load useListOnHeight from INI file
    useListOnHeight = readIntFromIniFile(iniFilePath, L"Window", L"Height", MIN_HEIGHT_scaled);
    useListOnHeight = std::max(useListOnHeight, MIN_HEIGHT_scaled); // Ensure minimum height

    // Set windowRect based on Use List state
    int height = useListEnabled ? useListOnHeight : useListOffHeight;

    windowRect.right = windowRect.left + width;
    windowRect.bottom = windowRect.top + height;

    // Read column widths
    findColumnWidth = std::max(readIntFromIniFile(iniFilePath, L"ListColumns", L"FindWidth", DEFAULT_COLUMN_WIDTH_FIND_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceColumnWidth = std::max(readIntFromIniFile(iniFilePath, L"ListColumns", L"ReplaceWidth", DEFAULT_COLUMN_WIDTH_REPLACE_scaled), MIN_GENERAL_WIDTH_scaled);
    commentsColumnWidth = std::max(readIntFromIniFile(iniFilePath, L"ListColumns", L"CommentsWidth", DEFAULT_COLUMN_WIDTH_COMMENTS_scaled), MIN_GENERAL_WIDTH_scaled);
    findCountColumnWidth = std::max(readIntFromIniFile(iniFilePath, L"ListColumns", L"FindCountWidth", DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceCountColumnWidth = std::max(readIntFromIniFile(iniFilePath, L"ListColumns", L"ReplaceCountWidth", DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);

    // Load column visibility states
    isFindCountVisible = readBoolFromIniFile(iniFilePath, L"ListColumns", L"FindCountVisible", false);
    isReplaceCountVisible = readBoolFromIniFile(iniFilePath, L"ListColumns", L"ReplaceCountVisible", false);
    isCommentsColumnVisible = readBoolFromIniFile(iniFilePath, L"ListColumns", L"CommentsVisible", false);
    isDeleteButtonVisible = readBoolFromIniFile(iniFilePath, L"ListColumns", L"DeleteButtonVisible", true);

    // Load column lock states
    findColumnLockedEnabled = readBoolFromIniFile(iniFilePath, L"ListColumns", L"FindColumnLocked", true);
    replaceColumnLockedEnabled = readBoolFromIniFile(iniFilePath, L"ListColumns", L"ReplaceColumnLocked", false);
    commentsColumnLockedEnabled = readBoolFromIniFile(iniFilePath, L"ListColumns", L"CommentsColumnLocked", true);

    // Load transparency settings with defaults
    foregroundTransparency = std::clamp(readByteFromIniFile(iniFilePath, L"Window", L"ForegroundTransparency", DEFAULT_FOREGROUND_TRANSPARENCY), MIN_TRANSPARENCY, MAX_TRANSPARENCY);
    backgroundTransparency = std::clamp(readByteFromIniFile(iniFilePath, L"Window", L"BackgroundTransparency", DEFAULT_BACKGROUND_TRANSPARENCY), MIN_TRANSPARENCY, MAX_TRANSPARENCY);
    // Load Tooltip setting
    tooltipsEnabled = readBoolFromIniFile(iniFilePath, L"Options", L"Tooltips", true);
}

std::wstring MultiReplace::readStringFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, const std::wstring& defaultValue) {
    // Convert std::wstring path to std::string path using WideCharToMultiByte
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, iniFilePath.c_str(), (int)iniFilePath.size(), NULL, 0, NULL, NULL);
    std::string filePath(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, iniFilePath.c_str(), (int)iniFilePath.size(), &filePath[0], size_needed, NULL, NULL);

    // Open file in binary mode to read UTF-8 data
    std::ifstream iniFile(filePath, std::ios::binary);
    if (!iniFile.is_open()) {
        return defaultValue; // Return default value if file can't be opened
    }

    // Read the file content into a std::string
    std::string utf8Content((std::istreambuf_iterator<char>(iniFile)), std::istreambuf_iterator<char>());

    // Convert UTF-8 std::string to std::wstring
    int len = MultiByteToWideChar(CP_UTF8, 0, utf8Content.c_str(), -1, NULL, 0);
    std::wstring wContent(len, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8Content.c_str(), -1, &wContent[0], len);

    // Process the content line by line
    std::wstringstream contentStream(wContent);
    std::wstring line;
    bool correctSection = false;

    while (std::getline(contentStream, line)) {
        if (line[0] == L'[') {
            size_t closingBracketPos = line.find(L']');
            if (closingBracketPos != std::wstring::npos) {
                correctSection = (line.substr(1, closingBracketPos - 1) == section);
            }
        }
        else if (correctSection) {
            size_t equalPos = line.find(L'=');
            if (equalPos != std::wstring::npos) {
                std::wstring foundKey = trim(line.substr(0, equalPos));
                std::wstring value = line.substr(equalPos + 1);

                if (foundKey == key) {
                    size_t lastQuotePos = value.rfind(L'\"');
                    size_t semicolonPos = value.find(L';', lastQuotePos);
                    if (semicolonPos != std::wstring::npos) {
                        value = value.substr(0, semicolonPos);
                    }
                    value = trim(value);

                    return unescapeCsvValue(value);
                }
            }
        }
    }

    return defaultValue; // Return default value if key is not found
}

bool MultiReplace::readBoolFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, bool defaultValue) {
    std::wstring defaultValueStr = defaultValue ? L"1" : L"0";
    std::wstring value = readStringFromIniFile(iniFilePath, section, key, defaultValueStr);
    return value == L"1";
}

int MultiReplace::readIntFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, int defaultValue) {
    return ::GetPrivateProfileIntW(section.c_str(), key.c_str(), defaultValue, iniFilePath.c_str());
}

std::size_t MultiReplace::readSizeTFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, std::size_t defaultValue) {
    WCHAR buffer[256] = { 0 };
    GetPrivateProfileStringW(section.c_str(), key.c_str(), std::to_wstring(defaultValue).c_str(), buffer, sizeof(buffer) / sizeof(WCHAR), iniFilePath.c_str());

    try {
        return static_cast<std::size_t>(std::stoull(buffer));  // Convert and cast directly to size_t
    }
    catch (...) {
        return defaultValue;  // If conversion fails, return the default value
    }
}

BYTE MultiReplace::readByteFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, BYTE defaultValue) {
    int intValue = ::GetPrivateProfileIntW(section.c_str(), key.c_str(), defaultValue, iniFilePath.c_str());
    return static_cast<BYTE>(intValue);
}

float MultiReplace::readFloatFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, float defaultValue) {
    WCHAR buffer[256] = { 0 };
    std::wstring defaultStr = std::to_wstring(defaultValue);

    GetPrivateProfileStringW(section.c_str(), key.c_str(), defaultStr.c_str(), buffer, sizeof(buffer) / sizeof(WCHAR), iniFilePath.c_str());

    try {
        return std::stof(buffer);
    }
    catch (...) {
        return defaultValue;
    }
}

void MultiReplace::setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text) {
    ::SetDlgItemTextW(hDlg, itemID, text.c_str());
}

#pragma endregion


#pragma region Language

std::wstring MultiReplace::getLanguageFromNativeLangXML() {
    wchar_t configDir[MAX_PATH] = { 0 };
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, reinterpret_cast<LPARAM>(configDir));
    configDir[MAX_PATH - 1] = L'\0';

    std::wstring nativeLangFilePath = std::wstring(configDir) + L"\\..\\..\\nativeLang.xml";
    std::wifstream file(nativeLangFilePath);
    std::wstring line;
    std::wstring language = L"english"; // default to English

    try {
        if (file.is_open()) {
            std::wregex languagePattern(L"<Native-Langue .*? filename=\"(.*?)\\.xml\"");
            std::wsmatch matches;

            while (std::getline(file, line)) {
                if (std::regex_search(line, matches, languagePattern) && matches.size() > 1) {
                    language = matches[1];
                    break; // Language found
                }
            }
        }
    }
    catch (const std::exception&) {
        // Error handling
    }

    file.close();
    return language; // Return the language code as a wide string
}

void MultiReplace::loadLanguage() {
    try {
        wchar_t pluginHomePath[MAX_PATH];
        ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, MAX_PATH, reinterpret_cast<LPARAM>(pluginHomePath));
        pluginHomePath[MAX_PATH - 1] = '\0';

        std::wstring languageIniFilePath = std::wstring(pluginHomePath) + L"\\MultiReplace\\languages.ini";
        std::wstring languageCode = getLanguageFromNativeLangXML();

        loadLanguageFromIni(languageIniFilePath, languageCode);
    }
    catch (const std::exception&) {
        // Error handling
    }
}

void MultiReplace::loadLanguageFromIni(const std::wstring& iniFilePath, const std::wstring& languageCode) {
    std::wstring section = languageCode;

    for (auto& entry : languageMap) {
        std::wstring translatedString = readStringFromIniFile(iniFilePath, section, entry.first, entry.second);
        entry.second = translatedString;
    }

}

std::wstring MultiReplace::getLangStr(const std::wstring& id, const std::vector<std::wstring>& replacements) {
    auto it = languageMap.find(id);
    if (it != languageMap.end()) {
        std::wstring result = it->second;
        std::wstring basePlaceholder = L"$REPLACE_STRING";

        // Replace all occurrences of "<br/>" with "\r\n" for line breaks in MessageBox
        size_t breakPos = result.find(L"<br/>");
        while (breakPos != std::wstring::npos) {
            result.replace(breakPos, 5, L"\r\n");  // 5 is the length of "<br/>"
            breakPos = result.find(L"<br/>");
        }

        // Replace placeholders with numbers from highest to lowest
        for (size_t i = replacements.size(); i > 0; i--) {
            std::wstring placeholder = basePlaceholder + std::to_wstring(i);
            size_t pos = result.find(placeholder);
            while (pos != std::wstring::npos) {
                result.replace(pos, placeholder.size(), replacements[i - 1]);
                pos = result.find(placeholder, pos + replacements[i - 1].size());
            }
        }

        // Finally, handle the case for $REPLACE_STRING
        size_t pos = result.find(basePlaceholder);
        while (pos != std::wstring::npos) {
            result.replace(pos, basePlaceholder.size(), replacements.empty() ? L"" : replacements[0]);
            pos = result.find(basePlaceholder, pos + replacements[0].size());
        }

        return result;
    }
    else {
        return L"Text not found"; // Return a default message if ID not found
    }
}

LPCWSTR MultiReplace::getLangStrLPCWSTR(const std::wstring& id) {
    static std::map<std::wstring, std::wstring> cache; // Static cache to hold strings and extend their lifetimes
    auto& cachedString = cache[id]; // Reference to the possibly existing entry in the cache

    if (cachedString.empty()) { // If not already cached, retrieve and store it
        cachedString = getLangStr(id);
    }

    return cachedString.c_str(); // Return a pointer to the cached string
}

LPWSTR MultiReplace::getLangStrLPWSTR(const std::wstring& id) {
    static std::wstring localStr;        // Static variable to extend the lifetime of the returned string
    localStr = getLangStr(id);           // Copy the result of getLangStr to the static variable
    return &localStr[0];                 // Return a modifiable pointer to the static string's buffer
}

#pragma endregion


#pragma region Event Handling -- triggered in beNotified() in MultiReplace.cpp

void MultiReplace::processTextChange(SCNotification* notifyCode) {
    if (!isWindowOpen || !isLoggingEnabled) {
        return;
    }

    Sci_Position cursorPosition = notifyCode->position;
    Sci_Position addedLines = notifyCode->linesAdded;
    Sci_Position notifyLength = notifyCode->length;

    Sci_Position lineNumber = ::SendMessage(getScintillaHandle(), SCI_LINEFROMPOSITION, cursorPosition, 0);
    if (notifyCode->modificationType & SC_MOD_INSERTTEXT) {
        if (addedLines != 0) {
            // Set the first entry as Modify
            MultiReplace::logChanges.push_back({ ChangeType::Modify, lineNumber });
            for (Sci_Position i = 1; i <= abs(addedLines); i++) {
                MultiReplace::logChanges.push_back({ ChangeType::Insert, lineNumber + i });
            }
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (MultiReplace::logChanges.empty() || MultiReplace::logChanges.back().changeType != ChangeType::Modify || MultiReplace::logChanges.back().lineNumber != lineNumber) {
                MultiReplace::logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
    else if (notifyCode->modificationType & SC_MOD_DELETETEXT) {
        if (addedLines != 0) {
            // Special handling for deletions at position 0
            if (cursorPosition == 0 && notifyLength == 0) {
                MultiReplace::logChanges.push_back({ ChangeType::Delete, 0 });
                return;
            }
            // Then, log the deletes in descending order
            for (Sci_Position i = abs(addedLines); i > 0; i--) {
                MultiReplace::logChanges.push_back({ ChangeType::Delete, lineNumber + i });
            }
            // Set the first entry as Modify for the smallest lineNumber
            MultiReplace::logChanges.push_back({ ChangeType::Modify, lineNumber });
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (MultiReplace::logChanges.empty() || MultiReplace::logChanges.back().changeType != ChangeType::Modify || MultiReplace::logChanges.back().lineNumber != lineNumber) {
                MultiReplace::logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
}

void MultiReplace::processLog() {
    if (!isWindowOpen) {
        return;
    }

    if (instance != nullptr) {
        instance->handleDelimiterPositions(DelimiterOperation::Update);
    }
}

void MultiReplace::onDocumentSwitched() {
    if (!isWindowOpen) {
        return;
    }

    // for scanned delimiter
    int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);
    if (currentBufferID != scannedDelimiterBufferID) {
        documentSwitched = true;
        isCaretPositionEnabled = false;
        scannedDelimiterBufferID = currentBufferID;
        if (instance != nullptr) {
            instance->isColumnHighlighted = false;
            instance->showStatusMessage(L"", RGB(0, 0, 0));
        }
    }

    // for sorted Columns
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

void MultiReplace::onSelectionChanged() {

    if (!isWindowOpen) {
        return;
    }

    static bool wasTextSelected = false;  // This stores the previous state

    // Get the start and end of the selection
    Sci_Position start = ::SendMessage(getScintillaHandle(), SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(getScintillaHandle(), SCI_GETSELECTIONEND, 0, 0);

    // Enable or disable IDC_SELECTION_RADIO depending on whether text is selected
    bool isTextSelected = (start != end);
    ::EnableWindow(::GetDlgItem(getDialogHandle(), IDC_SELECTION_RADIO), isTextSelected);

    // If no text is selected and IDC_SELECTION_RADIO is checked, check IDC_ALL_TEXT_RADIO instead
    if (!isTextSelected && (::SendMessage(::GetDlgItem(getDialogHandle(), IDC_SELECTION_RADIO), BM_GETCHECK, 0, 0) == BST_CHECKED)) {
        ::SendMessage(::GetDlgItem(getDialogHandle(), IDC_ALL_TEXT_RADIO), BM_SETCHECK, BST_CHECKED, 0);
        ::SendMessage(::GetDlgItem(getDialogHandle(), IDC_SELECTION_RADIO), BM_SETCHECK, BST_UNCHECKED, 0);
    }

    // Check if there was a switch from selected to not selected
    if (wasTextSelected && !isTextSelected) {
        if (instance != nullptr) {
            instance->setUIElementVisibility();
        }
    }
    wasTextSelected = isTextSelected;  // Update the previous state
}

void MultiReplace::onTextChanged() {
    textModified = true;
}

void MultiReplace::onCaretPositionChanged()
{
    if (!isWindowOpen || !isCaretPositionEnabled) {
        return;
    }

    LRESULT startPosition = ::SendMessage(getScintillaHandle(), SCI_GETCURRENTPOS, 0, 0);
    if (instance != nullptr) {
        instance->showStatusMessage(instance->getLangStr(L"status_actual_position", { instance->addLineAndColumnMessage(startPosition) }), COLOR_SUCCESS);
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