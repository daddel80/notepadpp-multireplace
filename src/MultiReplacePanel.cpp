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

#pragma warning(disable: 6262)

#pragma region Initialization

void MultiReplace::initializeWindowSize() {
    loadUIConfigFromIni(); // Loads the UI configuration, including window size and position

    // Set the window position and size based on the loaded settings
    SetWindowPos(_hSelf, NULL, windowRect.left, windowRect.top,
        windowRect.right - windowRect.left,
        windowRect.bottom - windowRect.top, SWP_NOZORDER);
}

RECT MultiReplace::calculateMinWindowFrame(HWND hwnd) {
    // Measure the window's borders and title bar
    RECT clientRect;
    GetClientRect(hwnd, &clientRect);
    GetWindowRect(hwnd, &windowRect);

    int borderWidth = ((windowRect.right - windowRect.left) - (clientRect.right - clientRect.left)) / 2;
    int titleBarHeight = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top) - borderWidth;

    RECT adjustedSize = { 0, 0, MIN_WIDTH + 2 * borderWidth, MIN_HEIGHT + borderWidth + titleBarHeight };

    return adjustedSize;
}

void MultiReplace::positionAndResizeControls(int windowWidth, int windowHeight)
{
    int buttonX = windowWidth - 45 - 160;
    int checkbox2X = buttonX + 173;
    int swapButtonX = windowWidth - 45 - 160 - 33;
    int comboWidth = windowWidth - 365;
    int frameX = windowWidth  - 320;
    int listWidth = windowWidth - 260;
    int listHeight = windowHeight - 295;
    int checkboxX = buttonX - 105;

    // Static positions and sizes
    ctrlMap[IDC_STATIC_FIND] = { 14, 19, 100, 24, WC_STATIC, getLangStrLPCWSTR(L"panel_find_what"), SS_RIGHT, NULL };
    ctrlMap[IDC_STATIC_REPLACE] = { 14, 54, 100, 24, WC_STATIC, getLangStrLPCWSTR(L"panel_replace_with"), SS_RIGHT };

    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { 20, 95, 198, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_match_whole_word_only"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { 20, 126, 198, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_match_case"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_CHECKBOX] = { 20, 157, 167, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_use_variables"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_USE_VARIABLES_HELP] = { 190, 158, 25, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_help"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_FIRST_CHECKBOX] = { 20, 189, 198, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_replace_first_match_only"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_WRAP_AROUND_CHECKBOX] = { 20, 220, 198, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_wrap_around"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };

    ctrlMap[IDC_SEARCH_MODE_GROUP] = { 225, 99, 200, 130, WC_BUTTON,  getLangStrLPCWSTR(L"panel_search_mode"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_NORMAL_RADIO] = { 235, 126, 180, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_normal"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXTENDED_RADIO] = { 235, 157, 180, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_extended"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REGEX_RADIO] = { 235, 188, 180, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_regular_expression"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_SCOPE_GROUP] = { 440, 99, 247, 163, WC_BUTTON, getLangStrLPCWSTR(L"panel_scope"), BS_GROUPBOX, NULL };
    ctrlMap[IDC_ALL_TEXT_RADIO] = { 450, 126, 230, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_all_text"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_SELECTION_RADIO] = { 450, 157, 230, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_selection"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COLUMN_MODE_RADIO] = { 450, 188, 50, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_csv"), BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COLUMN_NUM_STATIC] = { 450, 227, 30, 25, WC_STATIC, getLangStrLPCWSTR(L"panel_cols"), SS_RIGHT, NULL};
    ctrlMap[IDC_COLUMN_NUM_EDIT] = { 482, 227, 50, 20, WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL , getLangStrLPCWSTR(L"tooltip_columns") };
    ctrlMap[IDC_DELIMITER_STATIC] = { 538, 227, 40, 25, WC_STATIC, getLangStrLPCWSTR(L"panel_delim"), SS_RIGHT, NULL };
    ctrlMap[IDC_DELIMITER_EDIT] = { 580, 227, 30, 20, WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL ,  getLangStrLPCWSTR(L"tooltip_delimiter") };
    ctrlMap[IDC_QUOTECHAR_STATIC] = { 616, 227, 40, 25, WC_STATIC, getLangStrLPCWSTR(L"panel_quote"), SS_RIGHT, NULL };
    ctrlMap[IDC_QUOTECHAR_EDIT] = { 658, 227, 15, 20, WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL , getLangStrLPCWSTR(L"tooltip_quote") };

    ctrlMap[IDC_COLUMN_SORT_DESC_BUTTON] = { 508, 186, 17, 25, WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_sort_descending") };
    ctrlMap[IDC_COLUMN_SORT_ASC_BUTTON] = { 526, 186, 17, 25, WC_BUTTON, L"▼", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_sort_ascending") };
    ctrlMap[IDC_COLUMN_DROP_BUTTON] = { 554, 186, 25, 25, WC_BUTTON, L"✖", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_drop_columns") };
    ctrlMap[IDC_COLUMN_COPY_BUTTON] = { 590, 186, 25, 25, WC_BUTTON, L"🗍", BS_PUSHBUTTON | WS_TABSTOP,  getLangStrLPCWSTR(L"tooltip_copy_columns") };
    ctrlMap[IDC_COLUMN_HIGHLIGHT_BUTTON] = { 626, 186, 50, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_show"), BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_column_highlight") };

    ctrlMap[IDC_STATUS_MESSAGE] = { 14, 260, 630, 24, WC_STATIC, L"", WS_VISIBLE | SS_LEFT, NULL };

    // Dynamic positions and sizes
    ctrlMap[IDC_FIND_EDIT] = { 120, 19, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_EDIT] = { 120, 54, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_SWAP_BUTTON] = { swapButtonX, 33, 28, 34, WC_BUTTON, L"⇅", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, 19, 160, 60, WC_BUTTON, getLangStrLPCWSTR(L"panel_add_into_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_STATIC_FRAME] = { frameX, 99, 285, 163, WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, 118, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_replace_all"), BS_SPLITBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_BUTTON] = { buttonX, 118, 120, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_replace"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_SMALL_BUTTON] = { buttonX + 125, 118, 35, 30, WC_BUTTON, L"٭", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_replace_all") };
    ctrlMap[IDC_2_BUTTONS_MODE] = { checkbox2X, 118, 25, 25, WC_BUTTON, L"", BS_AUTOCHECKBOX | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_2_buttons_mode") };
    ctrlMap[IDC_FIND_BUTTON] = { buttonX, 153, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_find_next"), BS_PUSHBUTTON | WS_TABSTOP, NULL };

    findNextButtonText = L"▼ " + getLangStr(L"panel_find_next_small");
    ctrlMap[IDC_FIND_NEXT_BUTTON] = ControlInfo{ buttonX + 40, 153, 120, 30, WC_BUTTON, findNextButtonText.c_str(), BS_PUSHBUTTON | WS_TABSTOP, NULL };

    ctrlMap[IDC_FIND_PREV_BUTTON] = { buttonX, 153, 35, 30, WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_BUTTON] = { buttonX, 188, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_mark_matches"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, 188, 120, 30, WC_BUTTON,getLangStrLPCWSTR(L"panel_mark_matches_small"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX + 125, 188, 35, 30, WC_BUTTON, L"🗍", BS_PUSHBUTTON | WS_TABSTOP, getLangStrLPCWSTR(L"tooltip_copy_marked_text") };
    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, 223, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_clear_all_marks"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, 284, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, 319, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_save_list"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, 354, 160, 30, WC_BUTTON, getLangStrLPCWSTR(L"panel_export_to_bash"), BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_UP_BUTTON] = { buttonX + 5, 404, 30, 30, WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_DOWN_BUTTON] = { buttonX + 5, 404 + 30 + 5, 30, 30, WC_BUTTON, L"▼", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_SHIFT_FRAME] = { buttonX, 404 - 14, 165, 85, WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_SHIFT_TEXT] = { buttonX + 38, 404 + 20, 120, 20, WC_STATIC, getLangStrLPCWSTR(L"panel_shift_lines"), SS_LEFT, NULL };
    ctrlMap[IDC_REPLACE_LIST] = { 20, 284, listWidth, listHeight, WC_LISTVIEW, NULL, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS, NULL };
    ctrlMap[IDC_USE_LIST_CHECKBOX] = { checkboxX, 175, 95, 25, WC_BUTTON, getLangStrLPCWSTR(L"panel_use_list"), BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[ID_STATISTICS_COLUMNS] = { 2, 285, 17, 24, WC_BUTTON, L"▶", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, getLangStrLPCWSTR(L"tooltip_display_statistics_columns") };
}

void MultiReplace::initializeCtrlMap()
{

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

    // Create the font
    _hFont = CreateFont(FONT_SIZE, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, FONT_NAME);

    // Set the font for each control in ctrlMap
    for (auto& pair : ctrlMap)
    {
        SendMessage(GetDlgItem(_hSelf, pair.first), WM_SETFONT, (WPARAM)_hFont, TRUE);
    }

    // Limit the input for IDC_QUOTECHAR_EDIT to one character
    SendMessage(GetDlgItem(_hSelf, IDC_QUOTECHAR_EDIT), EM_SETLIMITTEXT, (WPARAM)1, 0);

    // Set the larger, bolder font for the swap, copy and refresh button
    HFONT hLargerBolderFont = CreateFont(28, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, TEXT("Courier New"));
    SendMessage(GetDlgItem(_hSelf, IDC_SWAP_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont, TRUE);

    HFONT hLargerBolderFont1 = CreateFont(29, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, TEXT("Courier New"));
    SendMessage(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont1, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_COLUMN_COPY_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont1, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_ALL_SMALL_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont1, TRUE);

    // CheckBox to Normal
    CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);

    // CheckBox to All Text
    CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_ALL_TEXT_RADIO);
    setElementsState(columnRadioDependentElements, false);
    
    // Enable IDC_SELECTION_RADIO based on text selection
    Sci_Position start = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONEND, 0, 0);
    bool isTextSelected = (start != end);
    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), isTextSelected);

    // Enable the checkbox ID IDC_USE_LIST_CHECKBOX
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);

    // Prepopulate default values for CSV
    SetDlgItemText(_hSelf, IDC_COLUMN_NUM_EDIT, L"1-20");
    SetDlgItemText(_hSelf, IDC_DELIMITER_EDIT, L",");
    SetDlgItemText(_hSelf, IDC_QUOTECHAR_EDIT, L"\"");

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
            MessageBox(NULL, errorMsg.c_str(), getLangStr(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR);
            return false;
        }

        // Create the tooltip for this control if it has tooltip text
        if (pair.second.tooltipText != NULL && pair.second.tooltipText[0] != '\0')
        {
            HWND hwndTooltip = CreateWindowEx(NULL, TOOLTIPS_CLASS, NULL,
                WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON,
                CW_USEDEFAULT, CW_USEDEFAULT,
                CW_USEDEFAULT, CW_USEDEFAULT,
                _hSelf, NULL, hInstance, NULL);

            if (!hwndTooltip)
            {
                // Handle error
                continue;
            }

            // Associate the tooltip with the control
            TOOLINFO toolInfo = { 0 };
            toolInfo.cbSize = sizeof(toolInfo);
            toolInfo.hwnd = _hSelf;
            toolInfo.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
            toolInfo.uId = (UINT_PTR)hwndControl;
            toolInfo.lpszText = (LPWSTR)pair.second.tooltipText;
            SendMessage(hwndTooltip, TTM_ADDTOOL, 0, (LPARAM)&toolInfo);
        }

        // Show the window
        ShowWindow(hwndControl, SW_SHOW);
        UpdateWindow(hwndControl);
    }
    return true;
}

void MultiReplace::initializePluginStyle()
{
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
    createListViewColumns(_replaceListView);
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    ListView_SetExtendedListViewStyle(_replaceListView, LVS_EX_FULLROWSELECT | LVS_EX_SUBITEMIMAGES);
}

void MultiReplace::moveAndResizeControls() {
    // IDs of controls to be moved or resized
    const int controlIds[] = {
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_SWAP_BUTTON, IDC_STATIC_FRAME, IDC_COPY_TO_LIST_BUTTON, IDC_REPLACE_ALL_BUTTON,
        IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON, IDC_2_BUTTONS_MODE, IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON,
        IDC_FIND_PREV_BUTTON, IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON,
        IDC_USE_LIST_CHECKBOX, IDC_LOAD_FROM_CSV_BUTTON, IDC_SAVE_TO_CSV_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON,
        IDC_SHIFT_TEXT, IDC_EXPORT_BASH_BUTTON
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

    // IDs of controls to be redrawn
    const int redrawIds[] = {
        IDC_USE_LIST_CHECKBOX, IDC_COPY_TO_LIST_BUTTON, IDC_REPLACE_ALL_BUTTON, IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON,
        IDC_2_BUTTONS_MODE, IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON, IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON,
        IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT
    };

    // Redraw controls using stored HWNDs
    for (int ctrlId : redrawIds) {
        InvalidateRect(hwndMap[ctrlId], NULL, TRUE);
    }
}

void MultiReplace::updateButtonVisibilityBasedOnMode() {
    // Update visibility of the buttons based on IDC_2_BUTTONS_MODE and IDC_2_MARK_BUTTONS_MODE state
    BOOL twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE);

    // for replace buttons
    ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_SMALL_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON), twoButtonsMode ? SW_HIDE : SW_SHOW);

    // for find buttons
    ShowWindow(GetDlgItem(_hSelf, IDC_FIND_NEXT_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_FIND_PREV_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_FIND_BUTTON), twoButtonsMode ? SW_HIDE : SW_SHOW);

    // for mark buttons
    ShowWindow(GetDlgItem(_hSelf, IDC_MARK_MATCHES_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), twoButtonsMode ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_MARK_BUTTON), twoButtonsMode ? SW_HIDE : SW_SHOW);
}

void MultiReplace::updateStatisticsColumnButtonIcon() {
    LPCWSTR icon = isStatisticsColumnsExpanded ? L"◀" : L"▶";
    // LPCWSTR tooltip = isStatisticsColumnsExpanded ? L"Collapse statistics columns" : L"Expand statistics columns";

    SetDlgItemText(_hSelf, ID_STATISTICS_COLUMNS, icon);

    //updateButtonTooltip(ID_STATISTICS_COLUMNS, tooltip);
}

#pragma endregion


#pragma region ListView

HWND MultiReplace::CreateHeaderTooltip(HWND hwndParent) {
    HWND hwndTT = CreateWindowEx(WS_EX_TOPMOST,
        TOOLTIPS_CLASS,
        NULL,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        hwndParent,
        NULL,
        GetModuleHandle(NULL),
        NULL);

    SetWindowPos(hwndTT,
        HWND_TOPMOST,
        0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

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
    ti.lpszText = (LPWSTR)pszText;
    ti.rect = rect;

    SendMessage(hwndTT, TTM_ADDTOOL, 0, (LPARAM)&ti);
}

void MultiReplace::createListViewColumns(HWND listView) {
    LVCOLUMN lvc;
    ZeroMemory(&lvc, sizeof(lvc));

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvc.fmt = LVCFMT_LEFT;

    // Get the client rectangle
    RECT rcClient;
    GetClientRect(_hSelf, &rcClient);
    // Extract width from the RECT
    int windowWidth = rcClient.right - rcClient.left;

    // Calculate the remaining width for the first two columns
    int adjustedWidth = windowWidth - 279;

    // Calculate the total width of columns 5 to 10 (Options and Delete Button)
    int columns5to10Width = 30 * 7;

    // Calculate the remaining width after subtracting the widths of the specified columns
    int remainingWidth = adjustedWidth - findCountColumnWidth - replaceCountColumnWidth - columns5to10Width;

    // Ensure remainingWidth is not less than the minimum width
    remainingWidth = std::max(remainingWidth, 60);

    lvc.iSubItem = 0;
    lvc.pszText = L"";
    lvc.cx = 0;
    ListView_InsertColumn(listView, 0, &lvc);

    lvc.iSubItem = 1;
    lvc.pszText = getLangStrLPWSTR(L"header_find_count");;
    lvc.cx = findCountColumnWidth;
    lvc.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(listView, 1, &lvc);

    lvc.iSubItem = 2;
    lvc.pszText = getLangStrLPWSTR(L"header_replace_count");
    lvc.cx = replaceCountColumnWidth;
    lvc.fmt = LVCFMT_RIGHT;
    ListView_InsertColumn(listView, 2, &lvc);

    // Column for Selection
    lvc.iSubItem = 3;
    lvc.pszText = L"\u2610";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 3, &lvc);

    // Column for "Find" Text
    lvc.iSubItem = 4;
    lvc.pszText = getLangStrLPWSTR(L"header_find");
    lvc.cx = remainingWidth / 2;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(listView, 4, &lvc);

    // Column for "Replace" Text
    lvc.iSubItem = 5;
    lvc.pszText = getLangStrLPWSTR(L"header_replace");
    lvc.cx = remainingWidth / 2;
    ListView_InsertColumn(listView, 5, &lvc);

    // Columns for Options
    const std::wstring options[] = { L"header_whole_word", L"header_match_case", L"header_use_variables", L"header_extended", L"header_regex" };
    for (int i = 0; i < 5; i++) {
        lvc.iSubItem = 6 + i;
        lvc.pszText = getLangStrLPWSTR(options[i]);
        lvc.cx = 30;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(listView, 6 + i, &lvc);
    }

    // Column for Delete Button
    lvc.iSubItem = 11;
    lvc.pszText = L"";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 11, &lvc);

    //Adding Tooltips
    HWND hwndHeader = ListView_GetHeader(listView);
    HWND hwndTT = CreateHeaderTooltip(hwndHeader);

    AddHeaderTooltip(hwndTT, hwndHeader, 6, getLangStrLPWSTR(L"tooltip_header_whole_word"));
    AddHeaderTooltip(hwndTT, hwndHeader, 7, getLangStrLPWSTR(L"tooltip_header_match_case"));
    AddHeaderTooltip(hwndTT, hwndHeader, 8, getLangStrLPWSTR(L"tooltip_header_use_variables"));
    AddHeaderTooltip(hwndTT, hwndHeader, 9, getLangStrLPWSTR(L"tooltip_header_extended"));
    AddHeaderTooltip(hwndTT, hwndHeader, 10, getLangStrLPWSTR(L"tooltip_header_regex"));

}

void MultiReplace::insertReplaceListItem(const ReplaceItemData& itemData)
{
    // Return early if findText is empty
    if (itemData.findText.empty()) {
        showStatusMessage(getLangStr(L"status_no_find_string"), RGB(255, 0, 0));
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

    // Add the data to the vector
    ReplaceItemData newItemData = itemData;
    replaceListData.push_back(newItemData);

    // Show a status message indicating the value added to the list
    std::wstring message;
    if (isDuplicate) {
        message = getLangStr(L"status_duplicate_entry") + newItemData.findText;
    }
    else {
        message = getLangStr(L"status_value_added");
    }
    showStatusMessage(message, RGB(0, 128, 0));

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Update Header if there might be any changes
    updateHeaderSelection();
}

int MultiReplace::calcDynamicColWidth(const CountColWidths& widths) {
    int columns5to10Width = 210; // Simplified calculation (30 * 7).

    // Directly calculate the width available for each dynamic column.
    int totalRemainingWidth = widths.listViewWidth - widths.margin - columns5to10Width - widths.findCountWidth - widths.replaceCountWidth;
    int perColumnWidth = std::max(totalRemainingWidth, MIN_COLUMN_WIDTH * 2) / 2; // Ensure total min width is 120, then divide by 2.
    return perColumnWidth; // Return width for a single column.
}

void MultiReplace::updateListViewAndColumns(HWND listView, LPARAM lParam) {
	int newWidth = LOWORD(lParam); // calculate ListWidth as lParam return WindowWidth
    int newHeight = HIWORD(lParam);

    CountColWidths widths = {
        listView,
        newWidth - 279, // Direct use of newWidth for listViewWidth
        false, // This is not used for current calculation.
        ListView_GetColumnWidth(listView, 1), // Current Find Count Width
        ListView_GetColumnWidth(listView, 2), // Current Replace Count Width
        0 // No margin as already precalculated in newWidth
    };

    // Calculate width available for each dynamic column.
    int perColumnWidth = calcDynamicColWidth(widths);

    HWND listHwnd = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    SendMessage(widths.listView, WM_SETREDRAW, FALSE, 0);

    // Set column widths directly using calculated width.
    ListView_SetColumnWidth(listView, 4, perColumnWidth); // Find Text
    ListView_SetColumnWidth(listView, 5, perColumnWidth); // Replace Text

    MoveWindow(listHwnd, 20, 284, newWidth - 260, newHeight - 295, TRUE);

    SendMessage(widths.listView, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(widths.listView, NULL, TRUE);
    UpdateWindow(widths.listView);

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

void MultiReplace::shiftListItem(HWND listView, const Direction& direction) {

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);

    std::vector<size_t> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(listView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(i);
    }

    if (selectedIndices.empty()) {
        showStatusMessage(getLangStr(L"status_no_rows_selected_to_shift"), RGB(255, 0, 0));
        return;
    }

    // Check the bounds
    if ((direction == Direction::Up && selectedIndices.front() == 0) || (direction == Direction::Down && selectedIndices.back() == replaceListData.size() - 1)) {
        return; // Don't perform the move if it's out of bounds
    }

    // Perform the shift operation
    if (direction == Direction::Up) {
        for (size_t& index : selectedIndices) {
            size_t swapIndex = index - 1;
            std::swap(replaceListData[index], replaceListData[swapIndex]);
            index = swapIndex;
        }
    }
    else { // direction is Down
        for (auto it = selectedIndices.rbegin(); it != selectedIndices.rend(); ++it) {
            size_t swapIndex = *it + 1;
            std::swap(replaceListData[*it], replaceListData[swapIndex]);
            *it = swapIndex;
        }
    }

    ListView_SetItemCountEx(listView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Deselect all items
    for (int j = 0; j < ListView_GetItemCount(listView); ++j) {
        ListView_SetItemState(listView, j, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Re-select the shifted items
    for (size_t index : selectedIndices) {
        ListView_SetItemState(listView, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Show status message when rows are successfully shifted
    showStatusMessage(getLangStr(L"status_rows_shifted", { std::to_wstring(selectedIndices.size()) }), RGB(0, 128, 0));

}

void MultiReplace::handleDeletion(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }
    // Remove the item from the ListView
    ListView_DeleteItem(_replaceListView, pnmia->iItem);

    // Remove the item from the replaceListData vector
    replaceListData.erase(replaceListData.begin() + pnmia->iItem);

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Update Header if there might be any changes
    updateHeaderSelection();

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(getLangStr(L"status_one_line_deleted"), RGB(0, 128, 0));
}

void MultiReplace::deleteSelectedLines(HWND listView) {
    std::vector<int> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(listView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(i);
    }

    if (selectedIndices.empty()) {
        return;
    }

    size_t numDeletedLines = selectedIndices.size();
    //int firstSelectedIndex = selectedIndices.front();
    int lastSelectedIndex = selectedIndices.back();

    // Remove the selected lines from replaceListData
    for (auto it = selectedIndices.rbegin(); it != selectedIndices.rend(); ++it) {
        replaceListData.erase(replaceListData.begin() + *it);
    }

    ListView_SetItemCountEx(listView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Deselect all items
    for (int j = 0; j < ListView_GetItemCount(listView); ++j) {
        ListView_SetItemState(listView, j, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Select the next available line
    int nextIndexToSelect = lastSelectedIndex < ListView_GetItemCount(listView) ? lastSelectedIndex : ListView_GetItemCount(listView) - 1;
    if (nextIndexToSelect >= 0) {
        ListView_SetItemState(listView, nextIndexToSelect, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Update Header if there might be any changes
    updateHeaderSelection();

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(getLangStr(L"status_lines_deleted", { std::to_wstring(numDeletedLines) }), RGB(0, 128, 0));
}

void MultiReplace::sortReplaceListData(int column, SortDirection direction) {
    auto selectedRows = getSelectedRows(); // Preserve selection

    // Sort based on column and direction
    std::sort(replaceListData.begin(), replaceListData.end(),
        [this, column, direction](const ReplaceItemData& a, const ReplaceItemData& b) -> bool {
            switch (column) {
            case 1: { // Sort by findCount, converting "" to -1
                int numA = a.findCount.empty() ? -1 : std::stoi(a.findCount);
                int numB = b.findCount.empty() ? -1 : std::stoi(b.findCount);
                return direction == SortDirection::Ascending ? numA < numB : numA > numB;
            }
            case 2: { // Sort by replaceCount, converting "" to -1
                int numA = a.replaceCount.empty() ? -1 : std::stoi(a.replaceCount);
                int numB = b.replaceCount.empty() ? -1 : std::stoi(b.replaceCount);
                return direction == SortDirection::Ascending ? numA < numB : numA > numB;
            }
            case 4: { // Sort by findText
                if (direction == SortDirection::Ascending) {
                    return a.findText < b.findText;
                }
                else {
                    return a.findText > b.findText;
                }
            }
            case 5: { // Sort by replaceText
                if (direction == SortDirection::Ascending) {
                    return a.replaceText < b.replaceText;
                }
                else {
                    return a.replaceText > b.replaceText;
                }
            }
            default:
                return false; // In case of an unknown column
            }
        });

    // Update UI and restore selection
    updateHeaderSortDirection();
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
    selectRows(selectedRows);
}

std::vector<size_t> MultiReplace::getSelectedRows() {
    // Initialize row IDs
    size_t counter = 1;
    for (auto& row : replaceListData) {
        row.id = counter++;
    }

    // Collect IDs of selected rows
    std::vector<size_t> selectedIDs;
    int index = -1; // Use int to properly handle -1 case and comparison with ListView_GetNextItem return value
    while ((index = ListView_GetNextItem(_replaceListView, index, LVNI_SELECTED)) != -1) {
        if (index >= 0 && static_cast<size_t>(index) < replaceListData.size()) {
            selectedIDs.push_back(replaceListData[index].id);
        }
    }

    return selectedIDs;
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
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
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

void MultiReplace::updateCountColumns(size_t itemIndex, int findCount, int replaceCount)
{
    // Check if the itemIndex is valid
    if (itemIndex >= replaceListData.size()) {
        return;
    }

    // Access the item data from the list
    ReplaceItemData& itemData = replaceListData[itemIndex];

    // Update findCount if provided
    if (findCount != -1) {
        int currentFindCount = 0;
        if (!itemData.findCount.empty()) {
            currentFindCount = std::stoi(itemData.findCount);
        }
        itemData.findCount = std::to_wstring(currentFindCount + findCount);
    }

    // Update replaceCount if provided
    if (replaceCount != -1) {
        int currentReplaceCount = 0;
        if (!itemData.replaceCount.empty()) {
            currentReplaceCount = std::stoi(itemData.replaceCount);
        }
        itemData.replaceCount = std::to_wstring(currentReplaceCount + replaceCount);
    }

    // Update the list view to immediately reflect the changes
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

void MultiReplace::resizeCountColumns() {
    HWND listView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    RECT listRect;
    GetClientRect(listView, &listRect);
    int listViewWidth = listRect.right - listRect.left;

    LONG style = GetWindowLong(listView, GWL_STYLE);
    bool hasVerticalScrollbar = (style & WS_VSCROLL) != 0;
    int margin = hasVerticalScrollbar ? 0 : 16;

    CountColWidths widths = {
        listView,
        listViewWidth,
        hasVerticalScrollbar,
        ListView_GetColumnWidth(_replaceListView, 1), // Current Find Count Width
        ListView_GetColumnWidth(_replaceListView, 2), // Current Replace Count Width
        margin
    };

    // Determine the direction of the adjustment
    bool expandColumns = widths.findCountWidth < COUNT_COLUMN_WIDTH && widths.replaceCountWidth < COUNT_COLUMN_WIDTH;

    // Perform the adjustment in steps
    while ((expandColumns && (widths.findCountWidth < COUNT_COLUMN_WIDTH || widths.replaceCountWidth < COUNT_COLUMN_WIDTH)) ||
        (!expandColumns && (widths.findCountWidth > 0 || widths.replaceCountWidth > 0))) {

        if (expandColumns) {
            widths.findCountWidth = std::min(widths.findCountWidth + STEP_SIZE, COUNT_COLUMN_WIDTH);
            widths.replaceCountWidth = std::min(widths.replaceCountWidth + STEP_SIZE, COUNT_COLUMN_WIDTH);
        }
        else {
            widths.findCountWidth = std::max(widths.findCountWidth - STEP_SIZE, 0);
            widths.replaceCountWidth = std::max(widths.replaceCountWidth - STEP_SIZE, 0);
        }

        int perColumnWidth = calcDynamicColWidth(widths);

        SendMessage(widths.listView, WM_SETREDRAW, FALSE, 0);
        ListView_SetColumnWidth(widths.listView, 1, widths.findCountWidth);
        ListView_SetColumnWidth(widths.listView, 2, widths.replaceCountWidth);
        ListView_SetColumnWidth(widths.listView, 4, perColumnWidth);
        ListView_SetColumnWidth(widths.listView, 5, perColumnWidth);
        
        SendMessage(widths.listView, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(widths.listView, NULL, TRUE);
        UpdateWindow(widths.listView);
    }
}

#pragma endregion


#pragma region Contextmenu

void MultiReplace::toggleBooleanAt(int itemIndex, int column) {
    if (itemIndex < 0 || itemIndex >= static_cast<int>(replaceListData.size()) || !(column == 3 || (column >= 6 && column <= 10))) {
        return; // Early return for invalid column or item index
    }

    ReplaceItemData& item = replaceListData[itemIndex];
    switch (column) {
    case 3: item.isEnabled = !item.isEnabled; break;
    case 6: item.wholeWord = !item.wholeWord; break;
    case 7: item.matchCase = !item.matchCase; break;
    case 8: item.useVariables = !item.useVariables; break;
    case 9: item.extended = !item.extended; break;
    case 10: item.regex = !item.regex; break;
    }

    ListView_RedrawItems(_replaceListView, itemIndex, itemIndex);
}

void MultiReplace::editTextAt(int itemIndex, int column) {
    // Calculate the total width of previous columns to get the X coordinate for the start of the selected column
    int totalWidthBeforeColumn = 0;
    for (int i = 0; i < column; ++i) {
        totalWidthBeforeColumn += ListView_GetColumnWidth(_replaceListView, i);
    }

    // Determine the width of the current column
    int columnWidth = ListView_GetColumnWidth(_replaceListView, column);

    // Adjust Y position considering the scroll position and adding an offset for better alignment
    int rowHeight = 20; // Assumed row height, adjust as needed
    SCROLLINFO si = { sizeof(si), SIF_POS };
    GetScrollInfo(_replaceListView, SB_VERT, &si);
    int visibleRowIndex = itemIndex - si.nPos; // Calculate the visible row index
    int correctedY = (visibleRowIndex + 1) * rowHeight + 3; // Adjust Y position downwards by 3 pixels

    // Create the Edit window with corrected coordinates and size
    hwndEdit = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        totalWidthBeforeColumn, correctedY, columnWidth, rowHeight,
        _replaceListView, NULL, (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE), NULL);

    // Set the initial text for the Edit window
    wchar_t itemText[MAX_TEXT_LENGTH];
    ListView_GetItemText(_replaceListView, itemIndex, column, itemText, sizeof(itemText) / sizeof(wchar_t));
    itemText[MAX_TEXT_LENGTH - 1] = L'\0';
    SetWindowText(hwndEdit, itemText);

    // Adjust font size for the Edit window
    HFONT hFont = CreateFont(-11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY,
        DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    SendMessage(hwndEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Set focus and select all text
    SetFocus(hwndEdit);
    SendMessage(hwndEdit, EM_SETSEL, 0, -1);

    // Subclass the edit control to handle Enter and clicks outside
    SetWindowSubclass(hwndEdit, EditControlSubclassProc, 1, (DWORD_PTR)this);

    _editingItemIndex = itemIndex;
    _editingColumn = column;
}

LRESULT CALLBACK MultiReplace::EditControlSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    switch (msg) {
    case WM_KILLFOCUS: {
        MultiReplace* pThis = reinterpret_cast<MultiReplace*>(dwRefData);

        // Retrieve the new text from the edit control
        wchar_t newText[MAX_TEXT_LENGTH];
        GetWindowText(hwnd, newText, MAX_TEXT_LENGTH);

        // Check if the column and item index are valid before updating
        if (pThis->_editingColumn >= 0 && pThis->_editingItemIndex >= 0 && pThis->_editingItemIndex < static_cast<int>(pThis->replaceListData.size())) {
            // Update the replaceListData vector with the new text
            ReplaceItemData& item = pThis->replaceListData[pThis->_editingItemIndex];
            if (pThis->_editingColumn == 4) { // Assuming column 4 is for findText
                item.findText = newText;
            }
            else if (pThis->_editingColumn == 5) { // Assuming column 5 is for replaceText
                item.replaceText = newText;
            }

            // Reflect this change in the ListView
            ListView_SetItemText(pThis->_replaceListView, pThis->_editingItemIndex, pThis->_editingColumn, newText);
        }

        // Clean up and remove subclass
        RemoveWindowSubclass(hwnd, EditControlSubclassProc, uIdSubclass);
        DestroyWindow(hwnd);

        return 0;
    }
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void MultiReplace::createContextMenu(HWND hwnd, POINT ptScreen, MenuState state) {
    HMENU hMenu = CreatePopupMenu();
    if (hMenu) {

        AppendMenu(hMenu, MF_STRING | (state.clickedOnItem ? MF_ENABLED : MF_GRAYED), IDM_COPY_DATA_TO_FIELDS, L"&Transfer to Input Fields\tAlt+Up");
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_COPY_LINES_TO_CLIPBOARD, L"&Copy\tCtrl+C");
        AppendMenu(hMenu, MF_STRING | (state.canPaste ? MF_ENABLED : MF_GRAYED), IDM_PASTE_LINES_FROM_CLIPBOARD, L"&Paste\tCtrl+V");
        AppendMenu(hMenu, MF_STRING | (state.canEdit ? MF_ENABLED : MF_GRAYED), IDM_EDIT_VALUE, L"&Edit\t");
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_DELETE_LINES, L"&Delete\tDel");
        AppendMenu(hMenu, MF_STRING, IDM_SELECT_ALL, L"&Select All\tCtrl+A");
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allEnabled ? MF_ENABLED : MF_GRAYED), IDM_ENABLE_LINES, L"&Enable\tAlt+E");
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allDisabled ? MF_ENABLED : MF_GRAYED), IDM_DISABLE_LINES, L"&Disable\tAlt+D");
        TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, NULL);
        DestroyMenu(hMenu); // Clean up
    }
}

MenuState MultiReplace::checkMenuConditions(HWND listView, POINT ptScreen) {
    MenuState state;

    POINT ptClient = ptScreen;
    ScreenToClient(listView, &ptClient);

    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = ptClient;
    int hitTestResult = ListView_HitTest(listView, &hitInfo);

    // Check if column was clicked and set canEdit
    if (hitTestResult != -1) {
        int clickedColumn = -1;
        int totalWidth = 0;
        HWND header = ListView_GetHeader(listView);
        int columnCount = Header_GetItemCount(header);

        for (int i = 0; i < columnCount; i++) {
            totalWidth += ListView_GetColumnWidth(listView, i);
            if (ptClient.x < totalWidth) {
                clickedColumn = i;
                break;
            }
        }

        if ((clickedColumn == 3) || (clickedColumn >= 4 && clickedColumn <= 10)) {
            state.canEdit = true;
        }

        state.clickedOnItem = true; // Click was on an item
    }

    // Check if the clipboard content is valid for pasting
    state.canPaste = canPasteFromClipboard();

    // Check if any items are selected to enable the delete option
    int selectedCount = ListView_GetSelectedCount(listView);
    state.hasSelection = (selectedCount > 0);

    // Check if all selected items are enabled or disabled
    int enabledCount = 0;
    int disabledCount = 0;

    int itemIndex = -1;
    while ((itemIndex = ListView_GetNextItem(listView, itemIndex, LVNI_SELECTED)) != -1) {
        auto& itemData = replaceListData[itemIndex];
        if (itemData.isEnabled) {
            ++enabledCount;
        }
        else {
            ++disabledCount;
        }
    }

    state.allEnabled = (selectedCount == enabledCount);
    state.allDisabled = (selectedCount == disabledCount);

    return state;
}

void MultiReplace::performItemAction(POINT pt, ItemAction action) {
    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = pt; // Use event-provided coordinates
    int hitTestResult = ListView_HitTest(_replaceListView, &hitInfo);

    // Calculate the clicked column based on the X coordinate
    int clickedColumn = -1;
    int totalWidth = 0;
    HWND header = ListView_GetHeader(_replaceListView);
    int columnCount = Header_GetItemCount(header);

    for (int i = 0; i < columnCount; i++) {
        totalWidth += ListView_GetColumnWidth(_replaceListView, i);
        if (pt.x < totalWidth) {
            clickedColumn = i;
            break; // Column found
        }
    }

    switch (action) {
    case ItemAction::Edit: {
        // Exit if no item found at click position
        if (hitTestResult == -1) return;

        // Perform actions based on the clicked column
        if ((clickedColumn == 3) || (clickedColumn >= 6 && clickedColumn <= 10)) {
            toggleBooleanAt(hitTestResult, clickedColumn);
        }
        else if (clickedColumn == 4 || clickedColumn == 5) {
            editTextAt(hitTestResult, clickedColumn);
        }
        break;
    }
    case ItemAction::Copy: {
        // Copy selected items to clipboard
        copySelectedItemsToClipboard(_replaceListView);
        break;
    }
    case ItemAction::Paste: {
        // Determine insert position based on hit test; use hitTestResult directly if item was clicked
        int insertPosition = hitTestResult != -1 ? hitTestResult : -1;
        pasteItemsIntoList(insertPosition);
        break;
    }
    }
}

void MultiReplace::copySelectedItemsToClipboard(HWND listView) {
    std::wstring csvData;
    int itemCount = ListView_GetItemCount(listView);
    int selectedCount = ListView_GetSelectedCount(listView);

    if (selectedCount > 0) {
        for (int i = 0; i < itemCount; ++i) {
            if (ListView_GetItemState(listView, i, LVIS_SELECTED) & LVIS_SELECTED) {
                // Angenommen, der Index in der ListView entspricht dem Index in replaceListData
                const ReplaceItemData& item = replaceListData[i];
                std::wstring line = std::to_wstring(item.isEnabled) + L"," +
                    escapeCsvValue(item.findText) + L"," +
                    escapeCsvValue(item.replaceText) + L"," +
                    std::to_wstring(item.wholeWord) + L"," +
                    std::to_wstring(item.matchCase) + L"," +
                    std::to_wstring(item.useVariables) + L"," +
                    std::to_wstring(item.extended) + L"," +
                    std::to_wstring(item.regex) + L"\n";
                csvData += line;
            }
        }
    }

    if (!csvData.empty()) {
        // Convert std::wstring to std::string (UTF-8) as needed
        std::string utf8CsvData = wstringToString(csvData);

        // Copying to clipboard
        if (OpenClipboard(NULL)) {
            EmptyClipboard();
            HGLOBAL hClipboardData = GlobalAlloc(GMEM_DDESHARE, (utf8CsvData.size() + 1) * sizeof(wchar_t));
            if (hClipboardData) {
                wchar_t* pClipboardData = (wchar_t*)GlobalLock(hClipboardData);
                if (pClipboardData != NULL) {
                    memcpy(pClipboardData, utf8CsvData.c_str(), (utf8CsvData.size() + 1) * sizeof(char)); // Should be utf8CsvData and not csvData, and sizeof(char) for UTF-8
                    GlobalUnlock(hClipboardData);
                    SetClipboardData(CF_UNICODETEXT, hClipboardData);
                }
                else {
                    GlobalFree(hClipboardData);
                }
                CloseClipboard();
            }

            CloseClipboard();
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

                // For the format to be considered valid, ensure each line has the correct number of columns
                // and the second column (findText) must not be empty for a line to be valid
                if (columns.size() == 8 && !columns[1].empty()) {
                    canPaste = true; // Found at least one valid line, no need to check further
                }
            }

            GlobalUnlock(hData);
        }
    }

    CloseClipboard();
    return canPaste;
}

void MultiReplace::pasteItemsIntoList(int insertPosition) {
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
    std::vector<int> insertedItemsIndices; // To track indices of inserted items

    while (std::getline(contentStream, line)) {
        if (line.empty()) continue; // Skip empty lines

        std::vector<std::wstring> columns;
        std::wstring currentValue;
        bool insideQuotes = false;

        for (const wchar_t& ch : line) {
            if (ch == L'"') {
                insideQuotes = !insideQuotes;
                continue;
            }
            if (ch == L',' && !insideQuotes) {
                columns.push_back(unescapeCsvValue(currentValue));
                currentValue.clear();
            }
            else {
                currentValue += ch;
            }
        }
        columns.push_back(unescapeCsvValue(currentValue)); // Add the last value

        // Check for proper column count and non-empty findText before adding to the list
        if (columns.size() != 8 || columns[1].empty()) continue;

        ReplaceItemData item;
        item.isEnabled = std::stoi(columns[0]) != 0;
        item.findText = columns[1];
        item.replaceText = columns[2];
        item.wholeWord = std::stoi(columns[3]) != 0;
        item.matchCase = std::stoi(columns[4]) != 0;
        item.useVariables = std::stoi(columns[5]) != 0;
        item.extended = std::stoi(columns[6]) != 0;
        item.regex = std::stoi(columns[7]) != 0;

        // Inserting item into the list
        if (insertPosition >= 0 && insertPosition < replaceListData.size()) {
            replaceListData.insert(replaceListData.begin() + insertPosition, item);
            insertedItemsIndices.push_back(insertPosition); // Track inserted item index
            ++insertPosition; // Adjust position for the next insert
        }
        else {
            replaceListData.push_back(item);
            insertedItemsIndices.push_back(static_cast<int>(replaceListData.size() - 1));// Track inserted item index at the end
        }
    }

    // Selecting newly inserted items in the list view
    for (int idx : insertedItemsIndices) {
        ListView_SetItemState(_replaceListView, idx, LVIS_SELECTED, LVIS_SELECTED);
    }

    // Refresh the ListView to reflect the changes and new selections
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

#pragma endregion


#pragma region Dialog

INT_PTR CALLBACK MultiReplace::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{

    switch (message)
    {
    case WM_INITDIALOG:
    {
        loadLanguage();
        initializeWindowSize();
        pointerToScintilla();
        initializePluginStyle();
        initializeCtrlMap();
        initializeListView();
        loadSettings();
        updateButtonVisibilityBasedOnMode();
		updateStatisticsColumnButtonIcon();
        // Activate Dark Mode
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(_hSelf));
         return TRUE;
    }
    break;

    case WM_GETMINMAXINFO: {
        MINMAXINFO* pMMI = reinterpret_cast<MINMAXINFO*>(lParam);
        RECT adjustedSize = calculateMinWindowFrame(_hSelf);
        pMMI->ptMinTrackSize.x = adjustedSize.right;
        pMMI->ptMinTrackSize.y = adjustedSize.bottom;
        return 0;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_STATUS_MESSAGE) ) {
            SetTextColor(hdcStatic, _statusMessageColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return (LRESULT)GetStockObject(NULL_BRUSH);
        }

        break;
    }
    break;

    case WM_DESTROY:
    {
        saveSettings();
        if (hwndEdit) {
            DestroyWindow(hwndEdit);
        }
        DeleteObject(_hFont);
        DestroyWindow(_hSelf);
    }
    break;

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

            // Update the global windowRect dimensions
            GetWindowRect(_hSelf, &windowRect);

            // Move and resize the List
            updateListViewAndColumns(GetDlgItem(_hSelf, IDC_REPLACE_LIST), lParam);

            // Calculate Position for all Elements
            positionAndResizeControls(newWidth, newHeight);

            // Move all Elements
            moveAndResizeControls();
        }
        return 0;
    }
    break;

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
                if (pnmia->iSubItem == 11) { // Delete button column
                    handleDeletion(pnmia);
                }
                if (pnmia->iSubItem == 3) { // Select button column
                    // get current selection status of the item
                    bool currentSelectionStatus = replaceListData[pnmia->iItem].isEnabled;
                    // set the selection status to its opposite
                    setSelections(!currentSelectionStatus, true);
                }
            }
            break;

            case NM_DBLCLK:
            {
                NMITEMACTIVATE* pnmia = reinterpret_cast<NMITEMACTIVATE*>(lParam);
                handleCopyBack(pnmia);
            }
            break;

            case LVN_GETDISPINFO:
            {
                NMLVDISPINFO* plvdi = reinterpret_cast<NMLVDISPINFO*>(lParam);

                // Get the data from the vector
                ReplaceItemData& itemData = replaceListData[plvdi->item.iItem];

                // Display the data based on the subitem
                switch (plvdi->item.iSubItem)
                {
                case 1:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.findCount.c_str());
                    break;
                case 2:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceCount.c_str());
                    break;
                case 3:
                    if (itemData.isEnabled) {
                        plvdi->item.pszText = L"\u25A0";
                    }
                    else {
                        plvdi->item.pszText = L"\u2610";
                    }
                    break;
                case 4:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.findText.c_str());
                    break;
                case 5:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceText.c_str());
                    break;
                case 6:
                    if (itemData.wholeWord) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 7:
                    if (itemData.matchCase) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 8:
                    if (itemData.useVariables) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 9:
                    if (itemData.extended) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 10:
                    if (itemData.regex) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 11:
                    plvdi->item.mask |= LVIF_TEXT;
                    plvdi->item.pszText = L"\u2716";
                    break;
                }
            }
            break;

            case LVN_COLUMNCLICK:
            {
                NMLISTVIEW* pnmv = reinterpret_cast<NMLISTVIEW*>(lParam);

                if (pnmv->iSubItem == 3) {
                    setSelections(!allSelected);
                }
                else {
                    // Toggle or initialize the sort order for the clicked column
                    SortDirection newDirection = SortDirection::Ascending; // Default direction
                    if (columnSortOrder.find(pnmv->iSubItem) != columnSortOrder.end() && columnSortOrder[pnmv->iSubItem] == SortDirection::Ascending) {
                        newDirection = SortDirection::Descending;
                    }
                    columnSortOrder[pnmv->iSubItem] = newDirection;
                    lastColumn = pnmv->iSubItem;
                    sortReplaceListData(lastColumn, newDirection); // Now correctly passing SortDirection
                }
                break;
            }

            case LVN_KEYDOWN:
            {
                LPNMLVKEYDOWN pnkd = reinterpret_cast<LPNMLVKEYDOWN>(pnmh);
                HDC hDC = NULL;
                int iItem = -1;

                PostMessage(_replaceListView, WM_SETFOCUS, 0, 0);
                // Handling keyboard shortcuts for menu actions
                if (GetKeyState(VK_CONTROL) & 0x8000) { // If Ctrl is pressed
                    switch (pnkd->wVKey) {
                    case 'C': // Ctrl+C for Copy
                        performItemAction(_contextMenuClickPoint, ItemAction::Copy);
                        break;
                    case 'V': // Ctrl+V for Paste
                        performItemAction(_contextMenuClickPoint, ItemAction::Paste);
                        break;
                    case 'A': // Ctrl+A for Select All
                        ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
                        break;
                        // Add more Ctrl+ shortcuts here if needed
                    }
                }
                else if (GetKeyState(VK_MENU) & 0x8000) { // If Alt is pressed
                    switch (pnkd->wVKey) {
                    case 'A': // Alt+A for Enable Line
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
                        deleteSelectedLines(_replaceListView);
                        break;
                    case VK_F12: // F12 key
                        GetClientRect(_hSelf, &windowRect);

                        hDC = GetDC(_hSelf);
                        if (hDC)
                        {
                            // Get the current font of the window
                            HFONT currentFont = (HFONT)SendMessage(_hSelf, WM_GETFONT, 0, 0);
                            HFONT hOldFont = (HFONT)SelectObject(hDC, currentFont);

                            // Get the text metrics for the current font
                            TEXTMETRIC tm;
                            GetTextMetrics(hDC, &tm);

                            // Calculate the base units
                            SIZE size;
                            GetTextExtentPoint32W(hDC, L"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", 52, &size);
                            int baseUnitX = (size.cx / 26 + 1) / 2;
                            int baseUnitY = tm.tmHeight;

                            // Calculate the window size in dialog units
                            int duWidth = MulDiv(windowRect.right, 4, baseUnitX);
                            int duHeight = MulDiv(windowRect.bottom, 8, baseUnitY);

                            wchar_t sizeText[100];
                            wsprintfW(sizeText, L"Window Size: %ld x %ld DUs", duWidth, duHeight);

                            MessageBoxW(_hSelf, sizeText, L"Window Size", MB_OK);

                            // Cleanup
                            SelectObject(hDC, hOldFont);
                            ReleaseDC(_hSelf, hDC);
                        }
                        break;
                    case VK_SPACE: // Spacebar key
                        iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                        if (iItem >= 0) {
                            // get current selection status of the item
                            bool currentSelectionStatus = replaceListData[iItem].isEnabled;
                            // set the selection status to its opposite
                            setSelections(!currentSelectionStatus, true);
                        }
                        break;
                    }
                }
                
            }
            break;
            }
        }
     }
    break;

    case WM_CONTEXTMENU:
    {
        if ((HWND)wParam == _replaceListView) {
            POINT ptScreen;
            ptScreen.x = LOWORD(lParam);
            ptScreen.y = HIWORD(lParam);
            _contextMenuClickPoint = ptScreen; // Store initial click point for later action determination
            ScreenToClient(_replaceListView, &_contextMenuClickPoint); // Convert to client coordinates for hit testing
            MenuState state = checkMenuConditions(_replaceListView, ptScreen);
            createContextMenu(_hSelf, ptScreen, state); // Show context menu
            return TRUE;
        }
    }
    break;

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
    }
    break;

    case WM_COMMAND:
    {

        switch (LOWORD(wParam))
        {

        case IDC_USE_VARIABLES_HELP:
        {
            auto n = SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, 0, 0);
            std::wstring path(n, 0);
            SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, n + 1, reinterpret_cast<LPARAM>(path.data()));
            path += L"\\MultiReplace\\help_use_variables.html";
            ShellExecute(NULL, L"open", path.c_str(), NULL, NULL, SW_SHOWNORMAL);
        }
        break;

        case IDCANCEL:
        {

            EndDialog(_hSelf, 0);
            _MultiReplace.display(false);

        }
        break;

        case IDC_2_BUTTONS_MODE:
        {
            // Check if the Find checkbox has been clicked
            if (HIWORD(wParam) == BN_CLICKED)
            {
                updateButtonVisibilityBasedOnMode();
            }
        }
        break;

        case IDC_REGEX_RADIO:
        {
            // Check if the Regular expression radio button is checked
            bool regexChecked = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);

            // Enable or disable the Whole word checkbox accordingly
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), !regexChecked);

            // If the Regular expression radio button is checked, uncheck the Whole word checkbox
            if (regexChecked)
            {
                CheckDlgButton(_hSelf, IDC_WHOLE_WORD_CHECKBOX, BST_UNCHECKED);
            }
        }
        break;

        case IDC_NORMAL_RADIO:
        case IDC_EXTENDED_RADIO:
        {
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), TRUE);
        }
        break;

        case IDC_ALL_TEXT_RADIO:
        {
            setElementsState(columnRadioDependentElements, false);
            setElementsState(selectionRadioDisabledButtons, true);
            handleClearDelimiterState();
        }
        break;

        case IDC_SELECTION_RADIO:
        {
            setElementsState(columnRadioDependentElements, false);
            setElementsState(selectionRadioDisabledButtons, false);
            handleClearDelimiterState();
        }
        break;

        case IDC_COLUMN_MODE_RADIO:
        {
            setElementsState(columnRadioDependentElements, true);
            setElementsState(selectionRadioDisabledButtons, true);
        }
        break;

        case IDC_COLUMN_SORT_ASC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortColumns(SortDirection::Ascending);
            }
        }
        break;

        case IDC_COLUMN_SORT_DESC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortColumns(SortDirection::Descending);
            }
        }
        break;

        case IDC_COLUMN_DROP_BUTTON:
        {
            if (confirmColumnDeletion()) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleDeleteColumns();
                }
            }
        }
        break;

        case IDC_COLUMN_COPY_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleCopyColumnsToClipboard();
            }
        }
        break;

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
                showStatusMessage(getLangStr(L"status_column_marks_cleared"), RGB(0, 128, 0));
            }
        }
        break;

        case IDC_USE_LIST_CHECKBOX:
        {
            // Check if the Use List Checkbox is enabled
            bool listCheckboxChecked = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);

            // Enable or disable the ListView accordingly
            EnableWindow(_replaceListView, listCheckboxChecked);

        }
        break;

        case IDC_SWAP_BUTTON:
        {
            std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

            // Swap the content of the two text fields
            SetDlgItemTextW(_hSelf, IDC_FIND_EDIT, replaceText.c_str());
            SetDlgItemTextW(_hSelf, IDC_REPLACE_EDIT, findText.c_str());
        }
        break;

        case IDC_COPY_TO_LIST_BUTTON:
        {
            handleCopyToListButton();
        }
        break;

        case IDC_FIND_BUTTON:
        case IDC_FIND_NEXT_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindNextButton();
        }
        break;

        case IDC_FIND_PREV_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindPrevButton(); 
        }
        break;

        case IDC_REPLACE_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceButton();
        }
        break;

        case IDC_REPLACE_ALL_SMALL_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceAllButton();
        }
        break;

        case IDC_REPLACE_ALL_BUTTON:
        {
            if (isReplaceAllInDocs)
            {
                int msgboxID = MessageBox(
                    NULL,
                    getLangStr(L"msgbox_confirm_replace_all").c_str(),
                    getLangStr(L"msgbox_title_confirm").c_str(),
                    MB_OKCANCEL
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
                resetCountColumns();
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                handleReplaceAllButton();
            }
        }
        break;

        case IDC_MARK_MATCHES_BUTTON:
        case IDC_MARK_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleClearTextMarksButton();
            handleMarkMatchesButton();
        }
        break;

        case IDC_CLEAR_MARKS_BUTTON:
        {
            resetCountColumns();
            handleClearTextMarksButton();
            showStatusMessage(getLangStr(L"status_all_marks_cleared"), RGB(0, 128, 0));
        }
        break;

        case IDC_COPY_MARKED_TEXT_BUTTON:
        {
            handleCopyMarkedTextToClipboardButton();
        }
        break;

        case IDC_SAVE_TO_CSV_BUTTON:
        {
            std::wstring filePath = openFileDialog(true, L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0", L"Save List As", OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"csv");
            if (!filePath.empty()) {
                saveListToCsv(filePath, replaceListData);
            }
        }
        break;

        case IDC_LOAD_FROM_CSV_BUTTON:
        {
            std::wstring filePath = openFileDialog(false, L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0", L"Open List", OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST, L"csv");
            if (!filePath.empty()) {
                loadListFromCsv(filePath);
            }
        }
        break;

        case IDC_UP_BUTTON:
        {
            shiftListItem(_replaceListView, Direction::Up);
        }
        break;

        case IDC_DOWN_BUTTON:
        {
            shiftListItem(_replaceListView, Direction::Down);
        }
        break;

        case IDC_EXPORT_BASH_BUTTON:
        {
            std::wstring filePath = openFileDialog(true, L"Bash Files (*.sh)\0*.sh\0All Files (*.*)\0*.*\0", L"Export as Bash", OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"sh");
            if (!filePath.empty()) {
                exportToBashScript(filePath);
            }
        }
        break;

        case ID_REPLACE_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, getLangStrLPWSTR(L"split_button_replace_all"));
            isReplaceAllInDocs = false;
        }
        break;

        case ID_REPLACE_IN_ALL_DOCS_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, getLangStrLPWSTR(L"split_button_replace_all_in_docs"));
            isReplaceAllInDocs = true;
        }
        break;

        case ID_STATISTICS_COLUMNS:
        {
            isStatisticsColumnsExpanded = !isStatisticsColumnsExpanded;
            resizeCountColumns();
            updateStatisticsColumnButtonIcon();
        }
		break;

        case IDM_EDIT_VALUE: 
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Edit);
        }
        break;

        case IDM_COPY_DATA_TO_FIELDS:
        {
            NMITEMACTIVATE nmia = {};
            nmia.iItem = ListView_HitTest(_replaceListView, &_contextMenuClickPoint);
            handleCopyBack(&nmia);            
        }
        break;

		case IDM_COPY_LINES_TO_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Copy);
        }
		break;

        case IDM_PASTE_LINES_FROM_CLIPBOARD: {
            performItemAction(_contextMenuClickPoint, ItemAction::Paste);
        }
        break;

        case IDM_DELETE_LINES:
        {
            int selectedCount = ListView_GetSelectedCount(_replaceListView);
            // Renaming 'message' to 'confirmationMessage' to avoid hiding the function parameter
            std::wstring confirmationMessage = L"Are you sure you want to delete ";
            if (selectedCount == 1) {
                confirmationMessage += L"this line?";
            }
            else {
                confirmationMessage += std::to_wstring(selectedCount) + L" lines?";
            }

            int msgBoxID = MessageBox(NULL, confirmationMessage.c_str(), L"Confirmation", MB_ICONWARNING | MB_YESNO);
            if (msgBoxID == IDYES)
            {
                // If confirmed, proceed to delete
                deleteSelectedLines(_replaceListView);
            }
        }
        break;

        case IDM_SELECT_ALL: {
            ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
        }
        break;

        case IDM_ENABLE_LINES: {
            setSelections(true, ListView_GetSelectedCount(_replaceListView) > 0);
        }
        break;

        case IDM_DISABLE_LINES: {
            setSelections(false, ListView_GetSelectedCount(_replaceListView) > 0);
        }
        break;


        default:
            return FALSE;
        }

    }
    break;

    }
    return FALSE;
}

#pragma endregion


#pragma region Replace

void MultiReplace::handleReplaceAllButton() {

    // First check if the document is read-only
    LRESULT isReadOnly = ::SendMessage(_hScintilla, SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(getLangStr(L"status_cannot_replace_read_only"), RGB(255, 0, 0));
        return;
    }

    // Clear all stored Lua Global Variables
    globalLuaVariablesMap.clear();

    int totalReplaceCount = 0;
    // Check if the "In List" option is enabled
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);

    if (useListEnabled)
    {
        // Check if the replaceListData is empty and warn the user if so
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_instructions"), RGB(255, 0, 0));
            return;
        }
        ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
        for (size_t i = 0; i < replaceListData.size(); ++i)
        {
            if (replaceListData[i].isEnabled)
            {
                int findCount = 0;
                int replaceCount = 0;
                replaceAll(replaceListData[i], findCount, replaceCount);

                // Update counts in list item
                if (findCount > 0) {
                    updateCountColumns(i, findCount, replaceCount);
                }

                // Accumulate total replacements
                totalReplaceCount += replaceCount;
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
    showStatusMessage(getLangStr(L"status_occurrences_replaced", { std::to_wstring(totalReplaceCount) }), RGB(0, 128, 0));
}

void MultiReplace::handleReplaceButton() {

    // First check if the document is read-only
    LRESULT isReadOnly = ::SendMessage(_hScintilla, SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(getLangStrLPWSTR(L"status_cannot_replace_read_only"), RGB(255, 0, 0));
        return;
    }

    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SearchResult searchResult;
    searchResult.pos = -1;
    searchResult.length = 0;
    searchResult.foundText = "";

    Sci_Position newPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
    size_t matchIndex = std::numeric_limits<size_t>::max();

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_uncheck"), RGB(255, 0, 0));
            return;
        }

        SelectionInfo selection = getSelectionInfo();

        int replacements = 0;  // Counter for replacements
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled && replaceOne(replaceListData[i], selection, searchResult, newPos)) {
                replacements++;
                updateCountColumns(i, -1, 1);
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
                showStatusMessage(getLangStr(L"status_replace_next_found", { std::to_wstring(replacements) }), RGB(0, 128, 0));
            }
            else {
                showStatusMessage(getLangStr(L"status_replace_none_left", { std::to_wstring(replacements) }), RGB(255, 0, 0));
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(getLangStr(L"status_no_occurrence_found"), RGB(255, 0, 0));
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

        SelectionInfo selection = getSelectionInfo();
        bool wasReplaced = replaceOne(replaceItem, selection, searchResult, newPos);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), replaceItem.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceItem.replaceText);

        if (searchResult.pos < 0 && wrapAroundEnabled) {
            searchResult = performSearchForward(findTextUtf8, searchFlags, true, 0);
        }

        if (wasReplaced) {
            if (searchResult.pos >= 0) {
                showStatusMessage(getLangStr(L"status_replace_one_next_found"), RGB(0, 128, 0));
            }
            else {
                showStatusMessage(getLangStr(L"status_replace_one_none_left"), RGB(255, 0, 0));
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(getLangStr(L"status_no_occurrence_found"), RGB(255, 0, 0));
            }
        }
    }

}

bool MultiReplace::replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos)
{
    std::string findTextUtf8 = convertAndExtend(itemData.findText, itemData.extended);
    int searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
    searchResult = performSearchForward(findTextUtf8, searchFlags, true, selection.startPos);

    if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
        bool skipReplace = false;
        std::string replaceTextUtf8 = convertAndExtend(itemData.replaceText, itemData.extended);
        std::string localReplaceTextUtf8 = wstringToString(itemData.replaceText);
        if (itemData.useVariables) {
            LuaVariables vars;

            int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
            int previousLineStartPosition = (currentLineIndex == 0) ? 0 : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));

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

void MultiReplace::replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount)
{
    if (itemData.findText.empty()) {
        findCount = 0;
        replaceCount = 0;
        return;
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

            lineFindCount++;

            vars.CNT = findCount;
            vars.LCNT = lineFindCount;
            vars.APOS = static_cast<int>(searchResult.pos) + 1;
            vars.LINE = currentLineIndex + 1;
            vars.LPOS = static_cast<int>(searchResult.pos) - previousLineStartPosition + 1;
            vars.MATCH = searchResult.foundText;

            if (!resolveLuaSyntax(localReplaceTextUtf8, vars, skipReplace, itemData.regex)) {
                break;  // Exit the loop if error in syntax
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
    send(SCI_SETCURRENTPOS, newTargetEnd, 0);

    // Clear selection
    send(SCI_SETSELECTIONSTART, newTargetEnd, 0);
    send(SCI_SETSELECTIONEND, newTargetEnd, 0);

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
    send(SCI_SETCURRENTPOS, newTargetEnd, 0);

    // Clear selection
    send(SCI_SETSELECTIONSTART, newTargetEnd, 0);
    send(SCI_SETSELECTIONEND, newTargetEnd, 0);

    return newTargetEnd;
}

SelectionInfo MultiReplace::getSelectionInfo() {
    // Get selected text
    Sci_Position selectionStart = ::SendMessage(_hScintilla, SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position selectionEnd = ::SendMessage(_hScintilla, SCI_GETSELECTIONEND, 0, 0);
    std::vector<char> buffer(selectionEnd - selectionStart + 1);
    ::SendMessage(_hScintilla, SCI_GETSELTEXT, 0, reinterpret_cast<LPARAM>(&buffer[0]));
    std::string selectedText(&buffer[0]);

    // Calculate the length of the selected text
    Sci_Position selectionLength = selectionEnd - selectionStart;

    return SelectionInfo{ selectedText, selectionStart, selectionLength };
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

bool MultiReplace::resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex)
{
    lua_State* L = luaL_newstate();  // Create a new Lua environment
    luaL_openlibs(L);  // Load standard libraries

    loadLuaGlobals(L); // Load global Lua variables

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

    // Convert numbers to integers
    luaL_dostring(L, "CNT = math.tointeger(CNT)");
    luaL_dostring(L, "LCNT = math.tointeger(LCNT)");
    luaL_dostring(L, "LINE = math.tointeger(LINE)");
    luaL_dostring(L, "LPOS = math.tointeger(LPOS)");
    luaL_dostring(L, "APOS = math.tointeger(APOS)");
    luaL_dostring(L, "COL = math.tointeger(COL)");

    setLuaVariable(L, "MATCH", vars.MATCH);
    // Get CAPs from Scintilla using SCI_GETTAG
    std::vector<std::string> caps;  // Initialize an empty vector to store the captures

    if (regex) {
        sptr_t len = 0;
        for (int i = 1; ; ++i) {
            char buffer[1024] = { 0 };  // Buffer to hold the capture value
            len = send(SCI_GETTAG, i, reinterpret_cast<sptr_t>(buffer), true);

            if (len <= 0) {
                // If len is zero or negative, break the loop
                break;
            }

            if (len < sizeof(buffer)) {
                // If the first character is 0x00, break the loop
                if (buffer[0] == 0x00) {
                    break;
                }
                buffer[len] = '\0';  // Null-terminate the string
                std::string cap(buffer);  // Convert to std::string
                caps.push_back(cap);  // Add the capture to the vector
            }
            else {
                // Buffer overflow detected: This should be rare, but it's good to check
                lua_close(L);
                return false;
            }
        }
    }

    // Process the captures and set them as global variables
    for (size_t i = 0; i < caps.size(); ++i) {
        std::string cap = caps[i];
        std::string globalVarName = "CAP" + std::to_string(i + 1);
        setLuaVariable(L, globalVarName, cap);
    }

    // Declare cond statement function
    luaL_dostring(L,
        "function cond(cond, trueVal, falseVal)\n"
        "  local res = {result = '', skip = false}  -- Initialize result table with defaults\n"
        "  if cond == nil then  -- Check if cond is nil\n"
        "    error('cond cannot be nil')\n"
        "    return res\n"
        "  end\n"

        "  if trueVal == nil then  -- Check if trueVal is nil\n"
        "    error('trueVal cannot be nil')\n"
        "    return res\n"
        "  end\n"

        "  if falseVal == nil then  -- Check if falseVal is missing\n"
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
        "    error('cannot be nil')\n"
        "    return\n"
        "  end\n"
        "  if type(strOrCalc) == 'string' then\n"
        "    res.result = strOrCalc  -- Setting res.result\n"
        "  elseif type(strOrCalc) == 'number' then\n"
        "    res.result = tostring(strOrCalc)  -- Convert number to string and set to res.result\n"
        "  else\n"
        "    error('Expected string or number')\n"
        "    return\n"
        "  end\n"
        "  resultTable = res\n"
        "  return res  -- Return the table containing result and skip\n"
        "end\n");

    // Declare formatNumber function
    luaL_dostring(L,
        "function fmtN(num, maxDecimals, fixedDecimals)\n"
        "  if num == nil then\n"
        "    error('num cannot be nil')\n"
        "    return\n"
        "  elseif type(num) ~= 'number' then\n"
        "    error('Invalid type for num. Expected a number')\n"
        "    return\n"
        "  end\n"
        "  if maxDecimals == nil then\n"
        "    error('maxDecimals cannot be nil')\n"
        "    return\n"
        "  elseif type(maxDecimals) ~= 'number' then\n"
        "    error('Invalid type for maxDecimals. Expected a number')\n"
        "    return\n"
        "  end\n"
        "  if fixedDecimals == nil then\n"
        "    error('fixedDecimals cannot be nil')\n"
        "    return\n"
        "  elseif type(fixedDecimals) ~= 'boolean' then\n"
        "    error('Invalid type for fixedDecimals. Expected a boolean')\n"
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
        "end");

    // Declare the init function
    luaL_dostring(L,
        "function init(args)\n"
        "  for name, value in pairs(args) do\n"
        "    if _G[name] == nil then\n"
        "      if type(name) ~= 'string' then\n"
        "        error('Variable name must be a string')\n"
        "      end\n"
        "      if not string.match(name, '^[A-Za-z_][A-Za-z0-9_]*$') then\n"
        "        error('Invalid variable name')\n"
        "      end\n"
        "      if value == nil then\n"
        "        error('Value missing in Init')\n"
        "      end\n"
        "      _G[name] = value\n"
        "    end\n"
        "  end\n"
        "end\n");
    
    // Show syntax error
    if (luaL_dostring(L, inputString.c_str()) != LUA_OK) {
        const char* cstr = lua_tostring(L, -1);
        lua_pop(L, 1);
        if (isLuaErrorDialogEnabled) {
            std::wstring error_message = utf8ToWString(cstr);
            MessageBoxW(NULL, error_message.c_str(), getLangStr(L"msgbox_title_use_variables_syntax_error").c_str(), MB_OK);
        }
        lua_close(L);
        return false;
    }

    // Retrieve the result from the table
    lua_getglobal(L, "resultTable");
    if (lua_istable(L, -1)) {
        lua_getfield(L, -1, "result");
        if (lua_isstring(L, -1) || lua_isnumber(L, -1)) {
            inputString = lua_tostring(L, -1);  // Update inputString with the result
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
            MessageBoxW(NULL, errorMsg.c_str(), errorTitle.c_str(), MB_OK);
        }
        lua_close(L);
        return false;
    }
    lua_pop(L, 1);  // Pop the 'result' table from the stack

    // Read Lua global Variables
    captureLuaGlobals(L);
    std::string luaVariablesStr;
    for (const auto& pair : globalLuaVariablesMap) {
        const LuaVariable& var = pair.second;
        luaVariablesStr += var.name + ": ";

        switch (var.type) {
        case LuaVariableType::String:
            luaVariablesStr += "String, " + var.stringValue;
            break;
        case LuaVariableType::Number:
            luaVariablesStr += "Number, " + std::to_string(var.numberValue);
            break;
        case LuaVariableType::Boolean:
            luaVariablesStr += "Boolean, " + std::string(var.booleanValue ? "true" : "false");
            break;
        default:
            luaVariablesStr += "None or Unsupported Type";
            break;
        }
        luaVariablesStr += "\n";
    }

    //MessageBoxA(NULL, luaVariablesStr.c_str(), "Lua Variables", MB_OK);


    lua_close(L);

    return true;

}

void MultiReplace::setLuaVariable(lua_State* L, const std::string& varName, std::string value) {
    bool isNumber = normalizeAndValidateNumber(value);
    if (isNumber) {
        double doubleVal = std::stod(value);
        int intVal = static_cast<int>(doubleVal);
        if (doubleVal == static_cast<double>(intVal)) {
            lua_pushinteger(L, intVal);
        }
        else {
            lua_pushnumber(L, doubleVal);
        }
    }
    else {
        lua_pushstring(L, value.c_str());
    }
    lua_setglobal(L, varName.c_str());
}

#pragma endregion


#pragma region Find

void MultiReplace::handleFindNextButton() {
    size_t matchIndex = std::numeric_limits<size_t>::max();

    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    LRESULT searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_find_directly"), RGB(255, 0, 0));
            return;
        }

        SearchResult result = performListSearchForward(replaceListData, searchPos, matchIndex);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performListSearchForward(replaceListData, 0, matchIndex);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                showStatusMessage(getLangStr(L"status_wrapped"), RGB(0, 128, 0));
                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", RGB(0, 128, 0));
            updateCountColumns(matchIndex, 1);
        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found"), RGB(255, 0, 0));
        }
    }
    else {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
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
                showStatusMessage(getLangStr(L"status_wrapped"), RGB(0, 128, 0));
                return;
            }
        }

        if (result.pos >= 0) {
            showStatusMessage(L"", RGB(0, 128, 0));
        }
        else {
            showStatusMessage(getLangStr(L"status_no_matches_found_for", { findText }), RGB(255, 0, 0));
        }
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
}

void MultiReplace::handleFindPrevButton() {

    size_t matchIndex = std::numeric_limits<size_t>::max();

    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    LRESULT searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
    searchPos = (searchPos > 0) ? ::SendMessage(_hScintilla, SCI_POSITIONBEFORE, searchPos, 0) : searchPos;

    if (useListEnabled)
    {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_find_directly"), RGB(255, 0, 0));
            return;
        }

        SearchResult result = performListSearchBackward(replaceListData, searchPos, matchIndex);

        if (result.pos >= 0) {
            updateCountColumns(matchIndex, 1);
            showStatusMessage(L"" + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
        }
        else if (wrapAroundEnabled)
        {
            result = performListSearchBackward(replaceListData, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0), matchIndex);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                showStatusMessage(getLangStr(L"status_wrapped_position", { addLineAndColumnMessage(result.pos) }), RGB(0, 128, 0));
            }
            else {
                showStatusMessage(getLangStr(L"status_no_matches_after_wrap"), RGB(255, 0, 0));
            }
        }
        else
        {
            showStatusMessage(getLangStr(L"status_no_matches_found"), RGB(255, 0, 0));
        }

    }
    else
    {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

        std::string findTextUtf8 = convertAndExtend(findText, extended);

        SearchResult result = performSearchBackward(findTextUtf8, searchFlags, searchPos);

        if (result.pos >= 0) {
            showStatusMessage(L"" + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
        }
        else if (wrapAroundEnabled)
        {
            result = performSearchBackward(findTextUtf8, searchFlags, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
            if (result.pos >= 0) {
                showStatusMessage(getLangStr(L"status_wrapped_find", { findText, addLineAndColumnMessage(result.pos) }), RGB(0, 128, 0));
            }
            else {
                showStatusMessage(getLangStr(L"status_no_matches_after_wrap_for", { findText }), RGB(255, 0, 0));
            }
        }
        else
        {
            showStatusMessage(getLangStr(L"status_no_matches_found_for", { findText }), RGB(255, 0, 0));
        }

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
}

SearchResult MultiReplace::performSingleSearch(const std::string& findTextUtf8, int searchFlags, bool selectMatch, SelectionRange range) {

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
        Sci_TextRange tr;
        tr.chrg.cpMin = static_cast<int>(result.pos);
        tr.chrg.cpMax = static_cast<int>(result.pos + result.length);

        if (tr.chrg.cpMax - tr.chrg.cpMin > sizeof(buffer) - 1) {
            // Safety check to avoid overflow.
            tr.chrg.cpMax = tr.chrg.cpMin + sizeof(buffer) - 1;
        }

        tr.lpstrText = buffer;
        send(SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));

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
    SearchResult result;
    SelectionRange targetRange;

    // Check if IDC_SELECTION_RADIO is enabled and selectMatch is false
    if (!selectMatch && IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        LRESULT selectionCount = ::SendMessage(_hScintilla, SCI_GETSELECTIONS, 0, 0);
        std::vector<SelectionRange> selections(selectionCount);

        for (int i = 0; i < selectionCount; i++) {
            selections[i].start = ::SendMessage(_hScintilla, SCI_GETSELECTIONNSTART, i, 0);
            selections[i].end = ::SendMessage(_hScintilla, SCI_GETSELECTIONNEND, i, 0);
        }

        // Sort selections based on their start position
        std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) {
            return a.start < b.start;
            });

        // Perform search within each selection
        for (const auto& selection : selections) {
            if (start >= selection.start && start < selection.end) {
                // If the start position is within the current selection
                targetRange = { start, selection.end };
                result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
            }
            else if (start < selection.start) {
                // If the start position is lower than the current selection
                targetRange = selection;
                result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
            }

            // Check if a match was found
            if (result.pos >= 0) {
                return result;
            }
        }
    }
    // Check if IDC_COLUMN_MODE_RADIO is enabled, selectMatch is false, and column delimiter data is set
    else if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED && columnDelimiterData.isValid()) {

        // Identify Column to Start
        ColumnInfo columnInfo = getColumnInfo(start);
        LRESULT totalLines = columnInfo.totalLines;
        LRESULT startLine = columnInfo.startLine;
        SIZE_T startColumnIndex = columnInfo.startColumnIndex;

        // Iterate over each line
        for (LRESULT line = startLine; line < totalLines; ++line) {
            if (line < static_cast<LRESULT>(lineDelimiterPositions.size())) {
                const auto& linePositions = lineDelimiterPositions[line].positions;
                SIZE_T totalColumns = linePositions.size() + 1;

                // Handle search for specific columns from columnDelimiterData
                for (SIZE_T column = startColumnIndex; column <= totalColumns; ++column) {

                    LRESULT startColumn = 0;
                    LRESULT endColumn = 0;

                    // Set start and end positions based on column index
                    if (column == 1) {
                        startColumn = lineDelimiterPositions[line].startPosition;
                    }
                    else {
                        startColumn = linePositions[column - 2].position + columnDelimiterData.delimiterLength;
                    }

                    if (column == linePositions.size() + 1) {
                        endColumn = lineDelimiterPositions[line].endPosition;
                    }
                    else {
                        endColumn = linePositions[column - 1].position;
                    }

                    // Check if the current column is included in the specified columns
                    if (columnDelimiterData.columns.find(static_cast<int>(column)) == columnDelimiterData.columns.end()) {
                        // If it's not included, skip this iteration
                        continue;
                    }

                    // If start position is within the column range, adjust startColumn
                    if (start >= startColumn && start <= endColumn) {
                        startColumn = start;
                    }

                    // Perform search within the column range
                    if (start <= startColumn) {
                        targetRange = { startColumn, endColumn };
                        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);

                        // Check if a match was found
                        if (result.pos >= 0) {
                            return result;
                        }
                    }
                }
                // Reset startColumnIndex for the next lines
                startColumnIndex = 1;
            }
        }
    }
    else {
        // If neither IDC_SELECTION_RADIO nor IDC_COLUMN_MODE_RADIO, perform search within the whole document
        targetRange.start = start;
        targetRange.end = send(SCI_GETLENGTH, 0, 0);
        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
    }

    return result;
}

SearchResult MultiReplace::performSearchBackward(const std::string& findTextUtf8, int searchFlags, LRESULT start)
{
    SearchResult result;
    SelectionRange targetRange;

    // Check if IDC_COLUMN_MODE_RADIO is enabled, and column delimiter data is set
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED && columnDelimiterData.isValid()) {

        // Identify Column to Start
        ColumnInfo columnInfo = getColumnInfo(start);
        LRESULT startLine = columnInfo.startLine;
        SIZE_T startColumnIndex = columnInfo.startColumnIndex;

        // Iterate over each line in reverse
        for (LRESULT line = startLine; line >= 0; --line) {
            if (line < static_cast<LRESULT>(lineDelimiterPositions.size())) {
                const auto& linePositions = lineDelimiterPositions[line].positions;
                SIZE_T totalColumns = linePositions.size() + 1;

                // Handle search for specific columns from columnDelimiterData
                for (SIZE_T column = (line == startLine ? startColumnIndex : totalColumns); column >= 1; --column) {

                    // Set start and end positions based on column index
                    LRESULT startColumn = 0;
                    LRESULT endColumn = 0;

                    if (column == 1) {
                        startColumn = lineDelimiterPositions[line].startPosition;
                    }
                    else {
                        startColumn = linePositions[column - 2].position + columnDelimiterData.delimiterLength;
                    }

                    if (column == linePositions.size() + 1) {
                        endColumn = lineDelimiterPositions[line].endPosition;
                    }
                    else {
                        endColumn = linePositions[column - 1].position;
                    }

                    // Check if the current column is included in the specified columns
                    if (columnDelimiterData.columns.find(static_cast<int>(column)) == columnDelimiterData.columns.end()) {
                        // If it's not included, skip this iteration
                        continue;
                    }

                    // Perform search within the column range
                    if (start >= startColumn && start <= endColumn) {
                        endColumn = start;
                    }

                    // Perform search within the column range
                    if (start >= endColumn) {
                        targetRange = { endColumn, startColumn };
                        result = performSingleSearch(findTextUtf8, searchFlags, true, targetRange);

                        // Check if a match was found
                        if (result.pos >= 0) {
                            return result;
                        }
                    }

                }
            }
        }
    }
    else {
        // Setting up the range to search backward from 'start' to the beginning
        SelectionRange searchRange;
        searchRange.start = start;
        searchRange.end = 0;
        result = performSingleSearch(findTextUtf8, searchFlags, true, searchRange);
    }

    return result;
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
            SearchResult result = performSearchBackward(findTextUtf8, searchFlags, cursorPos);

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

    closestMatchIndex = std::numeric_limits<size_t>::max(); // Initialisiert mit einem Wert, der "keinen Index" darstellt.

    for (size_t i = 0; i < list.size(); i++) {
        if (list[i].isEnabled) {
            int searchFlags = (list[i].wholeWord * SCFIND_WHOLEWORD) | (list[i].matchCase * SCFIND_MATCHCASE) | (list[i].regex * SCFIND_REGEXP);
            std::string findTextUtf8 = convertAndExtend(list[i].findText, list[i].extended);
            SearchResult result = performSearchForward(findTextUtf8, searchFlags, false, cursorPos);

            // Wenn ein Treffer gefunden wurde, der näher am Cursor liegt als der aktuelle nächste Treffer, aktualisiere den nächstgelegenen Treffer
            if (result.pos >= 0 && (closestMatch.pos < 0 || result.pos < closestMatch.pos)) {
                closestMatch = result;
                closestMatchIndex = i; // Aktualisiere den Index des nächstgelegenen Treffers
            }
        }
    }

    if (closestMatch.pos >= 0) { // Überprüfe, ob ein Treffer gefunden wurde
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

#pragma endregion


#pragma region Mark

void MultiReplace::handleMarkMatchesButton() {
    int totalMatchCount = 0;
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    markedStringsCount = 0;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(getLangStr(L"status_add_values_or_mark_directly"), RGB(255, 0, 0));
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
    showStatusMessage(getLangStr(L"status_occurrences_marked", { std::to_wstring(totalMatchCount) }), RGB(0, 0, 128));
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

    if (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED && markCount > 0) {
        markedStringsCount++;
    }

    return markCount;
}

void MultiReplace::highlightTextRange(LRESULT pos, LRESULT len, const std::string& findTextUtf8)
{
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
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
        showStatusMessage(getLangStr(L"status_no_text_to_copy"), RGB(255, 0, 0));
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
                    showStatusMessage(getLangStr(L"status_items_copied_to_clipboard", { std::to_wstring(textCount) }), RGB(0, 128, 0));
                }
                else
                {
                    showStatusMessage(getLangStr(L"status_failed_to_copy"), RGB(255, 0, 0));
                }
            }
            else
            {
                showStatusMessage(getLangStr(L"status_failed_allocate_memory"), RGB(255, 0, 0));
                GlobalFree(hClipboardData);
            }
        }
        CloseClipboard();
    }
}

#pragma endregion


#pragma region CSV

void MultiReplace::handleSortColumns(SortDirection sortDirection)
{
    // Validate column delimiter data
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), RGB(255, 0, 0));
        return;
    }
    SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);

    std::vector<CombinedColumns> combinedData;
    std::vector<size_t> index;
    size_t lineCount = lineDelimiterPositions.size();
    combinedData.reserve(lineCount); // Optimizing memory allocation
    index.reserve(lineCount);

    // Iterate through each line to extract content of specified columns
    for (SIZE_T i = CSVheaderLinesCount; i < lineCount; ++i) {
        const auto& lineInfo = lineDelimiterPositions[i];
        CombinedColumns rowData;
        rowData.columns.resize(columnDelimiterData.inputColumns.size());

        size_t columnIndex = 0;
        for (SIZE_T columnNumber : columnDelimiterData.inputColumns) {
            LRESULT startPos, endPos;

            // Calculate start and end positions for each column
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

            // Buffer to hold the text
            std::vector<char> buffer(static_cast<size_t>(endPos - startPos) + 1);

            // Prepare TextRange structure for Scintilla
            Sci_TextRange tr;
            tr.chrg.cpMin = startPos;
            tr.chrg.cpMax = endPos;
            tr.lpstrText = buffer.data();

            // Extract text for the column
            send(SCI_GETTEXTRANGE, 0, reinterpret_cast<sptr_t>(&tr), true);
            rowData.columns[columnIndex++] = std::string(buffer.data());
        }

        combinedData.push_back(rowData);
        index.push_back(i);
    }

    // Sorting logic
    std::sort(index.begin(), index.end(), [&](const size_t& a, const size_t& b) {
        // Adjusted indexing for actual data after headers
        size_t adjustedA = a - CSVheaderLinesCount;
        size_t adjustedB = b - CSVheaderLinesCount;

        // Compare columns for sorting
        for (size_t i = 0; i < columnDelimiterData.inputColumns.size(); ++i) {
            if (combinedData[adjustedA].columns[i] != combinedData[adjustedB].columns[i]) {
                return sortDirection == SortDirection::Ascending ?
                    combinedData[adjustedA].columns[i] < combinedData[adjustedB].columns[i] :
                    combinedData[adjustedA].columns[i] > combinedData[adjustedB].columns[i];
            }
        }
        return false; // In case of a tie, maintain original order
        });

    // Reordering lines in Scintilla based on the sorted index
    reorderLinesInScintilla(index);

    SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
}

void MultiReplace::reorderLinesInScintilla(const std::vector<size_t>& sortedIndex) {
    std::string combinedHeaderLines;
    std::string lineBreak = getEOLStyle();

    // Adjust the number of header lines if it exceeds the total number of lines
    SIZE_T actualHeaderLinesCount = CSVheaderLinesCount;
    SIZE_T totalLineCount = SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);
    if (CSVheaderLinesCount > totalLineCount) actualHeaderLinesCount = totalLineCount;

    // Extract and save the header lines
    for (SIZE_T i = 0; i < actualHeaderLinesCount; ++i) {
        LRESULT start = SendMessage(_hScintilla, SCI_POSITIONFROMLINE, i, 0);
        LRESULT end = SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, i, 0);
        std::vector<char> buffer(static_cast<size_t>(end - start) + 1);
        Sci_TextRange tr{ start, end, buffer.data() };
        SendMessage(_hScintilla, SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
        combinedHeaderLines += std::string(buffer.data());
        // Add line break to each header, except the last if no sorted lines follow
        if (i < actualHeaderLinesCount - 1 || (i == actualHeaderLinesCount - 1 && !sortedIndex.empty())) {
            combinedHeaderLines += lineBreak;
        }
    }

    // Extract the text of each line based on the sorted index
    std::vector<std::string> lines;
    lines.reserve(sortedIndex.size());
    for (size_t i = 0; i < sortedIndex.size(); ++i) {
        size_t idx = sortedIndex[i];
        LRESULT lineStart = SendMessage(_hScintilla, SCI_POSITIONFROMLINE, idx, 0);
        LRESULT lineEnd = SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, idx, 0);
        std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
        Sci_TextRange tr{ lineStart, lineEnd, buffer.data() };
        SendMessage(_hScintilla, SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
        lines.push_back(std::string(buffer.data()));
        // Add line break except after the last sorted line
        if (i < sortedIndex.size() - 1) {
            lines.back() += lineBreak;
        }
    }

    // Clear all content from Scintilla
    SendMessage(_hScintilla, SCI_CLEARALL, 0, 0);

    // Re-insert the header lines first
    SendMessage(_hScintilla, SCI_APPENDTEXT, combinedHeaderLines.length(), reinterpret_cast<LPARAM>(combinedHeaderLines.c_str()));

    // Re-insert the sorted lines
    for (const auto& line : lines) {
        SendMessage(_hScintilla, SCI_APPENDTEXT, line.length(), reinterpret_cast<LPARAM>(line.c_str()));
    }
}

bool MultiReplace::confirmColumnDeletion() {
    // Attempt to parse columns and delimiters
    if (!parseColumnAndDelimiterData()) {
        return false;  // Parsing failed, exit with false indicating no confirmation
    }

    // Now columnDelimiterData should be populated with the parsed column data
    size_t columnCount = columnDelimiterData.columns.size();
    std::wstring confirmMessage = getLangStr(L"msgbox_confirm_delete_columns", { std::to_wstring(columnCount) });
    int msgboxID = MessageBox(NULL, confirmMessage.c_str(), getLangStr(L"msgbox_title_confirm").c_str(), MB_ICONQUESTION | MB_YESNO);

    return (msgboxID == IDYES);  // Return true if user confirmed, else false
}

void MultiReplace::handleDeleteColumns()
{
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), RGB(255, 0, 0));
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
    showStatusMessage(getLangStr(L"status_deleted_fields_count", { std::to_wstring(deletedFieldsCount) }), RGB(0, 255, 0));
}

void MultiReplace::handleCopyColumnsToClipboard()
{
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(getLangStr(L"status_invalid_column_or_delimiter"), RGB(255, 0, 0));
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
                Sci_TextRange tr;
                tr.chrg.cpMin = startPos;
                tr.chrg.cpMax = endPos;
                tr.lpstrText = buffer.data();

                // Extract text for the column
                SendMessage(_hScintilla, SCI_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
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
    columnDelimiterData.quoteCharChanged = !(columnDelimiterData.quoteChar == wstringToString(quoteCharString) );

    // Initilaize values in case it will return with an error
    columnDelimiterData.extendedDelimiter = "";
    columnDelimiterData.quoteChar = "";
    columnDelimiterData.delimiterLength = 0;

    // Parse column data
    columnDataString.erase(0, columnDataString.find_first_not_of(L','));
    columnDataString.erase(columnDataString.find_last_not_of(L',') + 1);

    if (columnDataString.empty() || delimiterData.empty()) {
        showStatusMessage(getLangStr(L"status_missing_column_or_delimiter_data"), RGB(255, 0, 0));
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
                    showStatusMessage(getLangStr(L"status_invalid_range_in_column_data"), RGB(255, 0, 0));
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
                showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                return false;
            }
        }
        else {
            // Parse single column and add to set
            try {
                int column = std::stoi(block);

                if (column < 1) {
                    showStatusMessage(getLangStr(L"status_invalid_column_number"), RGB(255, 0, 0));
                    columnDelimiterData.columns.clear();
                    return false;
                }

                if (columns.insert(column).second) {
                    inputColumns.push_back(column); 
                }
            }
            catch (const std::exception&) {
                showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), RGB(255, 0, 0));
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
                showStatusMessage(getLangStr(L"status_invalid_range_in_column_data"), RGB(255, 0, 0));
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
            showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), RGB(255, 0, 0));
            columnDelimiterData.columns.clear();
            return false;
        }
    }
    else {
        try {
            int column = std::stoi(lastBlock);

            if (column < 1) {
                showStatusMessage(getLangStr(L"status_invalid_column_number"), RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                return false;
            }
            auto insertResult = columns.insert(column);
            if (insertResult.second) { // Check if the insertion was successful
                inputColumns.push_back(column); // Add to the inputColumns vector
            }
        }
        catch (const std::exception&) {
            showStatusMessage(getLangStr(L"status_syntax_error_in_column_data"), RGB(255, 0, 0));
            columnDelimiterData.columns.clear();
            return false;
        }
    }


    // Check delimiter data
    if (tempExtendedDelimiter.empty()) {
        showStatusMessage(getLangStr(L"status_extended_delimiter_empty"), RGB(255, 0, 0));
        return false;
    }

    // Check Quote Character
    if (!quoteCharString.empty() && (quoteCharString.length() != 1 || !(quoteCharString[0] == L'"' || quoteCharString[0] == L'\''))) {
        showStatusMessage(getLangStr(L"status_invalid_quote_character"), RGB(255, 0, 0));
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

    // Reset TextModiefeid Trigger
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

        // Find delimiters in line
        findDelimitersInLine(line);

    }

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
    char* buf = new char[lineLength + 1];

    // Get line content
    send(SCI_GETLINE, line, reinterpret_cast<sptr_t>(buf));
    std::string lineContent(buf, lineLength);
    delete[] buf;

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

    // Iterate over each line's delimiter positions
    LRESULT line = 0;
    while (line < static_cast<LRESULT>(lineDelimiterPositions.size()) ) {
        highlightColumnsInLine(line);
        ++line;
    }

    // Show Row and Column Position
    if (!lineDelimiterPositions.empty() ) {
        LRESULT startPosition = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
        showStatusMessage(getLangStr(L"status_actual_position", { addLineAndColumnMessage(startPosition) }), RGB(0, 128, 0));
    }

    SetDlgItemText(_hSelf, IDC_COLUMN_HIGHLIGHT_BUTTON, getLangStrLPCWSTR(L"panel_hide"));

    isColumnHighlighted = true;

    // Enable Position detection
    isCaretPositionEnabled = true;
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
                    end = (lineInfo.endPosition )- lineInfo.startPosition;
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

    SetDlgItemText(_hSelf, IDC_COLUMN_HIGHLIGHT_BUTTON, getLangStrLPCWSTR(L"panel_show"));

    isColumnHighlighted = false;

    // Disable Position detection
    isCaretPositionEnabled = false;
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
                if (modifyLogEntry.lineNumber >= logEntry.lineNumber -1 ) {
                    ++modifyLogEntry.lineNumber;
                }
            }
            updateDelimitersInDocument(static_cast<int>(logEntry.lineNumber), ChangeType::Insert);
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
    handleClearColumnMarks();
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

    int i = 0;
    do {
        if (!onlySelected || ListView_GetItemState(_replaceListView, i, LVIS_SELECTED)) {
            replaceListData[i].isEnabled = select;
        }
    } while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_ALL)) != -1);

    // Update the allSelected flag if all items were selected/deselected
    if (!onlySelected) {
        allSelected = select;
    }

    // Update the header after changing the selection status of the items
    updateHeaderSelection();

    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
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

    if (allSelected) {
        lvc.pszText = L"\u25A0"; // Ballot box with check
    }
    else if (anySelected) {
        lvc.pszText = L"\u25A3"; // Black square containing small white square
    }
    else {
        lvc.pszText = L"\u2610"; // Ballot box without check
    }

    ListView_SetColumn(_replaceListView, 3, &lvc);
}

void MultiReplace::updateHeaderSortDirection() {
    const wchar_t* ascendingSymbol = L"▲ ";
    const wchar_t* descendingSymbol = L"▼ ";

    // Iterate through each column that has a sort order defined in columnSortOrder
    for (const auto& [columnIndex, direction] : columnSortOrder) {
        std::wstring symbol = direction == SortDirection::Ascending ? ascendingSymbol : descendingSymbol;

        std::wstring headerText = L"" + symbol;
        symbol = L" " + symbol;

        // Append the base column title, this should be adjusted according to your actual column titles
        switch (columnIndex) {
        case 1: headerText = getLangStr(L"header_find_count") + symbol; break;
        case 2: headerText = getLangStr(L"header_replace_count") + symbol; break;
        case 4: headerText = getLangStr(L"header_find") + symbol; break;
        case 5: headerText = getLangStr(L"header_replace") + symbol; break;
        default: continue; // Skip if it's not a sortable column
        }

        // Prepare the LVCOLUMN structure for updating the header
        LVCOLUMN lvc = { 0 };
        lvc.mask = LVCF_TEXT;
        lvc.pszText = const_cast<LPWSTR>(headerText.c_str());

        // Correctly update the column header using the actual columnIndex without decrementing
        ListView_SetColumn(_replaceListView, columnIndex, &lvc); // Use columnIndex directly
    }
}

void MultiReplace::showStatusMessage(const std::wstring& messageText, COLORREF color)
{
    const size_t MAX_DISPLAY_LENGTH = 120; // Maximum length of the message to be displayed
    // Cut the message and add "..." if it's too long
    std::wstring strMessage = messageText;
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

void MultiReplace::setElementsState(const std::vector<int>& elements, bool enable) {
    for (int id : elements) {
        EnableWindow(GetDlgItem(_hSelf, id), enable ? TRUE : FALSE);
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

std::wstring MultiReplace::openFileDialog(bool saveFile, const WCHAR* filter, const WCHAR* title, DWORD flags, const std::wstring& fileExtension) {
    OPENFILENAME ofn = { 0 };
    WCHAR szFile[MAX_PATH] = { 0 };

    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = _hSelf;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile) / sizeof(WCHAR);
    ofn.lpstrFilter = filter;
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = title;
    ofn.Flags = flags;

    if (saveFile ? GetSaveFileName(&ofn) : GetOpenFileName(&ofn)) {
        std::wstring filePath(szFile);

        // Ensure that the filename ends with the correct extension if no extension is provided
        if (filePath.find_last_of(L".") == std::wstring::npos) {
            filePath += L"." + fileExtension;
        }

        return filePath;
    }
    else {
        return std::wstring();
    }
}

bool MultiReplace::saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    std::ofstream outFile(filePath);

    if (!outFile.is_open()) {
        return false;
    }

    // Convert and Write CSV header
    std::string utf8Header = wstringToString(L"Selected,Find,Replace,WholeWord,MatchCase,UseVariables,Regex,Extended\n");
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
            std::to_wstring(item.regex) + L"\n";
        std::string utf8Line = wstringToString(line);
        outFile << utf8Line;
    }

    outFile.close();
    
    return !outFile.fail();;
}

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (!saveListToCsvSilent(filePath, list)) {
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), RGB(255, 0, 0));
        return;
    }

    showStatusMessage(getLangStr(L"status_saved_items_to_csv", { std::to_wstring(list.size()) }), RGB(0, 128, 0));

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
}

void MultiReplace::loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list) {
    // Open file in binary mode to read UTF-8 data
    std::ifstream inFile(filePath);
    if (!inFile.is_open()) {
        throw CsvLoadException("status_unable_to_open_file");
    }

    std::vector<ReplaceItemData> tempList;  // Temporary list to hold items
    std::string utf8Content((std::istreambuf_iterator<char>(inFile)), std::istreambuf_iterator<char>());
    std::wstring content = stringToWString(utf8Content);
    std::wstringstream contentStream(content);

    std::wstring line;
    std::getline(contentStream, line); // Skip the CSV header

    while (std::getline(contentStream, line)) {
        std::wstringstream lineStream(line);
        std::vector<std::wstring> columns;
        bool insideQuotes = false;
        std::wstring currentValue;

        for (const wchar_t& ch : lineStream.str()) {
            if (ch == L'"') {
                insideQuotes = !insideQuotes;
            }
            if (ch == L',' && !insideQuotes) {
                columns.push_back(unescapeCsvValue(currentValue));
                currentValue.clear();
            }
            else {
                currentValue += ch;
            }
        }
        columns.push_back(unescapeCsvValue(currentValue));

        if (columns.size() != 8) {
            throw CsvLoadException("status_invalid_column_count");
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

            tempList.push_back(item);
        }
        catch (const std::exception&) {
            throw CsvLoadException("status_invalid_data_in_columns");
        }
    }

    inFile.close();
    list = tempList;  // Transfer data from temporary list to the final list
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {
    try {
        loadListFromCsvSilent(filePath, replaceListData);
    }
    catch (const CsvLoadException& ex) {
        // Resolve the error key to a localized string when displaying the message
        showStatusMessage(getLangStr(stringToWString(ex.what())), RGB(255, 0, 0));
        return;
    }

    if (replaceListData.empty()) {
        showStatusMessage(getLangStr(L"status_no_valid_items_in_csv"), RGB(255, 0, 0));
    }
    else {
        showStatusMessage(getLangStr(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), RGB(0, 128, 0));
    }

    // Update the list view control, if necessary
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
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

#pragma endregion


#pragma region Export

void MultiReplace::exportToBashScript(const std::wstring& fileName) {
    std::ofstream file(fileName);
    if (!file.is_open()) {
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), RGB(255, 0, 0));
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
    for (const auto& itemData : replaceListData) {
        if (!itemData.isEnabled) continue; // Skip if this item is not selected

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
        showStatusMessage(getLangStr(L"status_unable_to_save_file"), RGB(255, 0, 0));
        return;
    }

    showStatusMessage(getLangStr(L"status_list_exported_to_bash"), RGB(0, 128, 0));

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
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

    // Store window size and position from the global windowRect
    GetWindowRect(_hSelf, &windowRect);
    int width = windowRect.right - windowRect.left;
    int height = windowRect.bottom - windowRect.top;
    int posX = windowRect.left;
    int posY = windowRect.top;

    outFile << wstringToString(L"[Window]\n");
    outFile << wstringToString(L"Width=" + std::to_wstring(width) + L"\n");
    outFile << wstringToString(L"Height=" + std::to_wstring(height) + L"\n");
    outFile << wstringToString(L"PosX=" + std::to_wstring(posX) + L"\n");
    outFile << wstringToString(L"PosY=" + std::to_wstring(posY) + L"\n");

    // Store column widths for "Find Count" and "Replace Count"
    findCountColumnWidth = ListView_GetColumnWidth(_replaceListView, 1);
    replaceCountColumnWidth = ListView_GetColumnWidth(_replaceListView, 2);

    outFile << wstringToString(L"[ListColumns]\n");
    outFile << wstringToString(L"FindCountWidth=" + std::to_wstring(findCountColumnWidth) + L"\n");
    outFile << wstringToString(L"ReplaceCountWidth=" + std::to_wstring(replaceCountColumnWidth) + L"\n");

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
    int useList = IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED ? 1 : 0;

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

    // Convert and Store "Find what" history
    LRESULT findWhatCount = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETCOUNT, 0, 0);
    outFile << wstringToString(L"[History]\n");
    outFile << wstringToString(L"FindTextHistoryCount=" + std::to_wstring(findWhatCount) + L"\n");
    for (LRESULT i = 0; i < findWhatCount; i++) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring findTextData = escapeCsvValue(std::wstring(buffer.data()));
        outFile << wstringToString(L"FindTextHistory" + std::to_wstring(i) + L"=" + findTextData + L"\n");
    }

    // Store "Replace with" history
    LRESULT replaceWithCount = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETCOUNT, 0, 0);
    outFile << wstringToString(L"ReplaceTextHistoryCount=" + std::to_wstring(replaceWithCount) + L"\n");
    for (LRESULT i = 0; i < replaceWithCount; i++) {
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
        MessageBox(NULL, errorMessage.c_str(), getLangStr(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR);
    }
    settingsSaved = true;
}

void MultiReplace::loadSettingsFromIni(const std::wstring& iniFilePath) {

    // Load combo box histories
    int findHistoryCount = readIntFromIniFile(iniFilePath, L"History", L"FindTextHistoryCount", 0);
    for (int i = 0; i < findHistoryCount; i++) {
        std::wstring findHistoryItem = readStringFromIniFile(iniFilePath, L"History", L"FindTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findHistoryItem);
    }

    int replaceHistoryCount = readIntFromIniFile(iniFilePath, L"History", L"ReplaceTextHistoryCount", 0);
    for (int i = 0; i < replaceHistoryCount; i++) {
        std::wstring replaceHistoryItem = readStringFromIniFile(iniFilePath, L"History", L"ReplaceTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceHistoryItem);
    }

    // Read the current "Find what" and "Replace with" texts
    std::wstring findText = readStringFromIniFile(iniFilePath, L"Current", L"FindText", L"");
    std::wstring replaceText = readStringFromIniFile(iniFilePath, L"Current", L"ReplaceText", L"");

    setTextInDialogItem(_hSelf, IDC_FIND_EDIT, findText);
    setTextInDialogItem(_hSelf, IDC_REPLACE_EDIT, replaceText);

    bool wholeWord = readBoolFromIniFile(iniFilePath, L"Options", L"WholeWord", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);

    bool matchCase = readBoolFromIniFile(iniFilePath, L"Options", L"MatchCase", false);
    SendMessage(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, matchCase ? BST_CHECKED : BST_UNCHECKED, 0);

    bool useVariables = readBoolFromIniFile(iniFilePath, L"Options", L"UseVariables", false);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, useVariables ? BST_CHECKED : BST_UNCHECKED, 0);

    bool extended = readBoolFromIniFile(iniFilePath, L"Options", L"Extended", false);
    bool regex = readBoolFromIniFile(iniFilePath, L"Options", L"Regex", false);

    // Select the appropriate radio button based on the settings
    if (regex) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_REGEX_RADIO);
        EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), SW_HIDE);
    }
    else if (extended) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_EXTENDED_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);
    }

    bool wrapAround = readBoolFromIniFile(iniFilePath, L"Options", L"WrapAround", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WRAP_AROUND_CHECKBOX), BM_SETCHECK, wrapAround ? BST_CHECKED : BST_UNCHECKED, 0);

	bool replaceFirst = readBoolFromIniFile(iniFilePath, L"Options", L"ReplaceFirst", false);
	SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_FIRST_CHECKBOX), BM_SETCHECK, replaceFirst ? BST_CHECKED : BST_UNCHECKED, 0);

    bool replaceButtonsMode = readBoolFromIniFile(iniFilePath, L"Options", L"ButtonsMode", false);
    SendMessage(GetDlgItem(_hSelf, IDC_2_BUTTONS_MODE), BM_SETCHECK, replaceButtonsMode ? BST_CHECKED : BST_UNCHECKED, 0);

    bool useList = readBoolFromIniFile(iniFilePath, L"Options", L"UseList", false);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, useList ? BST_CHECKED : BST_UNCHECKED, 0);
    EnableWindow(_replaceListView, useList);

    // Load Scope
    int selection = readIntFromIniFile(iniFilePath, L"Scope", L"Selection", 0);
    int columnMode = readIntFromIniFile(iniFilePath, L"Scope", L"ColumnMode", 0);
    BOOL isEnabled = ::IsWindowEnabled(GetDlgItem(_hSelf, IDC_SELECTION_RADIO));

    if (selection && isEnabled) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_SELECTION_RADIO);
    }
    else if (columnMode) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_COLUMN_MODE_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_ALL_TEXT_RADIO);
    }

    BOOL columnModeSelected = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_NUM_EDIT), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_DELIMITER_EDIT), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_QUOTECHAR_EDIT), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_SORT_DESC_BUTTON), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_SORT_ASC_BUTTON), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_DROP_BUTTON), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_COPY_BUTTON), columnModeSelected);
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_HIGHLIGHT_BUTTON), columnModeSelected);

    std::wstring columnNum = readStringFromIniFile(iniFilePath, L"Scope", L"ColumnNum", L"");
    setTextInDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT, columnNum);

    std::wstring delimiter = readStringFromIniFile(iniFilePath, L"Scope", L"Delimiter", L"");
    setTextInDialogItem(_hSelf, IDC_DELIMITER_EDIT, delimiter);

    std::wstring quoteChar = readStringFromIniFile(iniFilePath, L"Scope", L"QuoteChar", L"");
    setTextInDialogItem(_hSelf, IDC_QUOTECHAR_EDIT, quoteChar);

    CSVheaderLinesCount = readIntFromIniFile(iniFilePath, L"Scope", L"HeaderLines", 1);
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

    // Load window position and size
    windowRect.left = readIntFromIniFile(iniFilePath, L"Window", L"PosX", POS_X);
    windowRect.top = readIntFromIniFile(iniFilePath, L"Window", L"PosY", POS_Y);
    windowRect.right = windowRect.left + std::max(readIntFromIniFile(iniFilePath, L"Window", L"Width", MIN_WIDTH), MIN_WIDTH);
    windowRect.bottom = windowRect.top + std::max(readIntFromIniFile(iniFilePath, L"Window", L"Height", MIN_HEIGHT), MIN_HEIGHT);

    // Read column widths
    findCountColumnWidth = readIntFromIniFile(iniFilePath, L"ListColumns", L"FindCountWidth", 0);
    replaceCountColumnWidth = readIntFromIniFile(iniFilePath, L"ListColumns", L"ReplaceCountWidth", 0);

    isStatisticsColumnsExpanded = (findCountColumnWidth >= COUNT_COLUMN_WIDTH && replaceCountColumnWidth >= COUNT_COLUMN_WIDTH);
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

        // Start at the end of replacements and count down
        if (!replacements.empty()) { // Check if replacements vector is not empty
            for (size_t i = replacements.size(); i-- > 0; ) {
                std::wstring placeholder = basePlaceholder + (i == 0 ? L"" : std::to_wstring(i + 1));
                size_t pos = result.find(placeholder);
                while (pos != std::wstring::npos) {
                    // Determine the replacement value based on the index
                    std::wstring replacementValue = replacements[i];
                    result.replace(pos, placeholder.size(), replacementValue);
                    pos = result.find(placeholder, pos + replacementValue.size());
                }
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

    Sci_Position lineNumber = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_LINEFROMPOSITION, cursorPosition, 0);
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

    int currentBufferID = (int)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);
    if (currentBufferID != scannedDelimiterBufferID) {
        documentSwitched = true;
        isCaretPositionEnabled = false;
        scannedDelimiterBufferID = currentBufferID;
        SetDlgItemText(s_hDlg, IDC_COLUMN_HIGHLIGHT_BUTTON, _MultiReplace.getLangStrLPCWSTR(L"panel_show"));
        if (instance != nullptr) {
            instance->isColumnHighlighted = false;
            instance->showStatusMessage(L"", RGB(0, 0, 0));
        }
    }
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
    const std::vector<int> selectionRadioDisabledButtons = {
    IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON, IDC_REPLACE_BUTTON
    };

    // Get the start and end of the selection
    Sci_Position start = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONEND, 0, 0);

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
            instance->setElementsState(selectionRadioDisabledButtons, true);
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

    LRESULT startPosition = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETCURRENTPOS, 0, 0);
    if (instance != nullptr) {
        instance->showStatusMessage(instance->getLangStr(L"status_actual_position", { instance->addLineAndColumnMessage(startPosition) }), RGB(0, 128, 0));
    }

}

#pragma endregion