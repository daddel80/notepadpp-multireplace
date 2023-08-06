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
#include "PluginDefinition.h"
#include <codecvt>
#include <locale>
#include <regex>
#include <windows.h>
#include <sstream>
#include <Commctrl.h>
#include <vector>
#include <fstream>
#include <iostream>
#include <map>
#include "Notepad_plus_msgs.h"
#include <bitset>
#include <string>
#include <functional>
#include <algorithm>
#include <unordered_map>
#include <set>
#include <chrono>
#include <sstream>


#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif

std::map<int, ControlInfo> MultiReplace::ctrlMap;
bool MultiReplace::isWindowOpen = false;
bool MultiReplace::isLoggingEnabled = true;
bool MultiReplace::textModified = true;
HWND MultiReplace::s_hScintilla = nullptr;
HWND MultiReplace::s_hDlg = nullptr;
std::vector<MultiReplace::LogEntry> MultiReplace::logChanges;


#pragma region Initialization

void MultiReplace::positionAndResizeControls(int windowWidth, int windowHeight)
{
    int buttonX = windowWidth - 40 - 160;
    int checkbox2X = buttonX + 173;
    int swapButtonX = windowWidth - 40 - 160 - 37;
    int comboWidth = windowWidth - 360;
    int frameX = windowWidth  - 310;
    int listWidth = windowWidth - 255;
    int listHeight = windowHeight - 274;
    int checkboxX = buttonX - 100;

    // Static positions and sizes
    ctrlMap[IDC_STATIC_FIND] = { 14, 19, 100, 24, WC_STATIC, L"Find what : ", SS_RIGHT, NULL };
    ctrlMap[IDC_STATIC_REPLACE] = { 14, 54, 100, 24, WC_STATIC, L"Replace with : ", SS_RIGHT };

    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { 20, 122, 155, 28, WC_BUTTON, L"Match whole word only", BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { 20, 146, 100, 28, WC_BUTTON, L"Match case", BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
    ctrlMap[IDC_WRAP_AROUND_CHECKBOX] = { 20, 170, 155, 28, WC_BUTTON, L"Wrap around", BS_AUTOCHECKBOX | WS_TABSTOP, NULL };

    ctrlMap[IDC_SEARCH_MODE_GROUP] = { 192, 105, 190, 100, WC_BUTTON, L"Search Mode", BS_GROUPBOX, NULL };
    ctrlMap[IDC_NORMAL_RADIO] = { 202, 125, 100, 20, WC_BUTTON, L"Normal", BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXTENDED_RADIO] = { 202, 150, 175, 20, WC_BUTTON, L"Extended (\\n, \\r, \\t, \\0, \\x...)", BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REGEX_RADIO] = { 202, 175, 175, 20, WC_BUTTON, L"Regular expression", BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    
    ctrlMap[IDC_SCOPE_GROUP] = { 400, 105, 230, 100, WC_BUTTON, L"Scope", BS_GROUPBOX, NULL };
    ctrlMap[IDC_ALL_TEXT_RADIO] = { 410, 125, 100, 20, WC_BUTTON, L"All Text", BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, NULL };
    ctrlMap[IDC_SELECTION_RADIO] = { 410, 150, 100, 20, WC_BUTTON, L"Selection", BS_AUTORADIOBUTTON | WS_TABSTOP,  L"Selection works with 'Replace All' and 'Mark Matches'"  };
    ctrlMap[IDC_COLUMN_MODE_RADIO] = { 410, 175, 20, 20, WC_BUTTON, L"", BS_AUTORADIOBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COLUMN_NUM_STATIC] = { 426, 175, 30, 20, WC_STATIC, L"Col.:", SS_RIGHT, NULL };
    ctrlMap[IDC_COLUMN_NUM_EDIT] = { 456, 175, 40, 20, WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL , L"Columns: '1,3,5-12' (individuals, ranges)" };
    ctrlMap[IDC_DELIMITER_STATIC] = { 501, 175, 40, 20, WC_STATIC, L"Delim.:", SS_RIGHT, NULL };
    ctrlMap[IDC_DELIMITER_EDIT] = { 541, 175, 40, 20, WC_EDIT, NULL, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL ,  L"Delimiter: Single/combined chars, \\t for Tab" };
    ctrlMap[IDC_COLUMN_HIGHLIGHT_BUTTON] = { 590, 173, 30, 25, WC_BUTTON, L"H", BS_PUSHBUTTON | WS_TABSTOP, L"Column highlight: On/Off" };

    ctrlMap[IDC_STATIC_HINT] = { 14, 100, 500, 60, WC_STATIC, L"Please enlarge the window to view the controls.", SS_CENTER, NULL };
    ctrlMap[IDC_STATUS_MESSAGE] = { 14, 220, 450, 24, WC_STATIC, L"", WS_VISIBLE | SS_LEFT, NULL };
    ctrlMap[IDC_ACTION_DESCRIPTION] = { 14, 220, 100, 24, WC_STATIC, L"", WS_VISIBLE | SS_LEFT, NULL };
    ctrlMap[IDC_PROGRESS_BAR] = { 120, 222, 330, 12, PROGRESS_CLASS, L"", WS_VISIBLE | WS_CHILD, NULL };
    ctrlMap[IDC_CANCEL_LONGRUN_BUTTON] = { 460, 215, 50, 25, WC_BUTTON, L"Cancel", BS_PUSHBUTTON | WS_TABSTOP, NULL };

    // Dynamic positions and sizes
    ctrlMap[IDC_FIND_EDIT] = { 120, 19, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_EDIT] = { 120, 54, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, NULL };
    ctrlMap[IDC_SWAP_BUTTON] = { swapButtonX, 33, 28, 34, WC_BUTTON, L"\u21C5", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, 19, 160, 60, WC_BUTTON, L"Add into List", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, 93, 160, 30, WC_BUTTON, L"Replace All", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_BUTTON] = { buttonX, 93, 120, 30, WC_BUTTON, L"Replace", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_REPLACE_ALL_SMALL_BUTTON] = { buttonX + 125, 93, 35, 30, WC_BUTTON, L"\u066D", BS_PUSHBUTTON | WS_TABSTOP, L"Replace All" };
    ctrlMap[IDC_2_BUTTONS_MODE] = { checkbox2X, 93, 20, 30, WC_BUTTON, L"", BS_AUTOCHECKBOX | WS_TABSTOP, L"2 buttons mode" };
    ctrlMap[IDC_FIND_BUTTON] = { buttonX, 128, 160, 30, WC_BUTTON, L"Find Next", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_FIND_NEXT_BUTTON] = { buttonX + 35, 128, 120, 30, WC_BUTTON, L"\u25BC Find Next", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_FIND_PREV_BUTTON] = { buttonX, 128, 35, 30, WC_BUTTON, L"\u25B2", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_BUTTON] = { buttonX, 163, 160, 30, WC_BUTTON, L"Mark Matches", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, 163, 120, 30, WC_BUTTON, L"Mark Matches", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX + 125, 163, 35, 30, WC_BUTTON, L"\U0001F5CD", BS_PUSHBUTTON | WS_TABSTOP, L"Copy to Clipboard" };
    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, 198, 160, 30, WC_BUTTON, L"Clear all marks", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, 244, 160, 30, WC_BUTTON, L"Load from CSV", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, 279, 160, 30, WC_BUTTON, L"Save to CSV", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, 314, 160, 30, WC_BUTTON, L"Export to Bash", BS_PUSHBUTTON | WS_TABSTOP, NULL };
    ctrlMap[IDC_UP_BUTTON] = { buttonX + 5, 364, 30, 30, WC_BUTTON, L"\u25B2", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_DOWN_BUTTON] = { buttonX + 5, 364 + 30 + 5, 30, 30, WC_BUTTON, L"\u25BC", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, NULL };
    ctrlMap[IDC_SHIFT_FRAME] = { buttonX, 364 - 14, 160, 85, WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_SHIFT_TEXT] = { buttonX + 38, 364 + 20, 60, 20, WC_STATIC, L"Shift Lines", SS_LEFT, NULL };
    ctrlMap[IDC_STATIC_FRAME] = { frameX, 80, 280, 155, WC_BUTTON, L"", BS_GROUPBOX, NULL };
    ctrlMap[IDC_REPLACE_LIST] = { 14, 244, listWidth, listHeight, WC_LISTVIEW, NULL, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS, NULL };
    ctrlMap[IDC_USE_LIST_CHECKBOX] = { checkboxX, 150, 80, 20, WC_BUTTON, L"Use List", BS_AUTOCHECKBOX | WS_TABSTOP, NULL };
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
    hFont = CreateFont(FONT_SIZE, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, FONT_NAME);

    // Set the font for each control in ctrlMap
    for (auto& pair : ctrlMap)
    {
        SendMessage(GetDlgItem(_hSelf, pair.first), WM_SETFONT, (WPARAM)hFont, TRUE);
    }

    // Set the larger, bolder font for the swap, copy and refresh button
    HFONT hLargerBolderFont = CreateFont(28, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, TEXT("Courier New"));
    SendMessage(GetDlgItem(_hSelf, IDC_SWAP_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont, TRUE);

    HFONT hLargerBolderFont1 = CreateFont(29, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, TEXT("Courier New"));
    SendMessage(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont1, TRUE);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_ALL_SMALL_BUTTON), WM_SETFONT, (WPARAM)hLargerBolderFont1, TRUE);

    // CheckBox to Normal
    CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);

    // CheckBox to All Text
    CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_ALL_TEXT_RADIO);
    
    // Enable IDC_SELECTION_RADIO based on text selection
    Sci_Position start = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(MultiReplace::getScintillaHandle(), SCI_GETSELECTIONEND, 0, 0);
    bool isTextSelected = (start != end);
    ::EnableWindow(::GetDlgItem(_hSelf, IDC_SELECTION_RADIO), isTextSelected);

    // Hide Hint Text
    ShowWindow(GetDlgItem(_hSelf, IDC_STATIC_HINT), SW_HIDE);

    // Enable the checkbox ID IDC_USE_LIST_CHECKBOX
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);

    isWindowOpen = true;
}

bool MultiReplace::createAndShowWindows() {

    // Buffer size for the error message
    constexpr int BUFFER_SIZE = 256;

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
            wchar_t msg[BUFFER_SIZE];
            DWORD dwError = GetLastError();
            wsprintf(msg, L"Failed to create control with ID: %d, GetLastError returned: %lu", pair.first, dwError);
            MessageBox(NULL, msg, L"Error", MB_OK | MB_ICONERROR);
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

void MultiReplace::initializeScintilla() {
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which != -1) {
        _hScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
        s_hScintilla = _hScintilla;
    }
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
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_SWAP_BUTTON, IDC_STATIC_FRAME, IDC_COPY_TO_LIST_BUTTON,
        IDC_REPLACE_ALL_BUTTON, IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON, IDC_2_BUTTONS_MODE,
        IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON,
        IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON,
        IDC_USE_LIST_CHECKBOX, IDC_LOAD_FROM_CSV_BUTTON, IDC_SAVE_TO_CSV_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT, 
        IDC_EXPORT_BASH_BUTTON
    };

    // IDs of controls to be redrawn
    const int redrawIds[] = {
        IDC_USE_LIST_CHECKBOX, IDC_REPLACE_ALL_BUTTON, IDC_REPLACE_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON, IDC_2_BUTTONS_MODE,
        IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON,
        IDC_MARK_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON,
        IDC_COPY_MARKED_TEXT_BUTTON,
        IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT,
    };

    // Move and resize controls
    for (int ctrlId : controlIds) {
        const ControlInfo& ctrlInfo = ctrlMap[ctrlId];
        MoveWindow(GetDlgItem(_hSelf, ctrlId), ctrlInfo.x, ctrlInfo.y, ctrlInfo.cx, ctrlInfo.cy, TRUE);
    }

    // Redraw controls
    for (int ctrlId : redrawIds) {
        InvalidateRect(GetDlgItem(_hSelf, ctrlId), NULL, TRUE);
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

void MultiReplace::updateUIVisibility() {
    // Get the current window size
    RECT rect;
    GetClientRect(_hSelf, &rect);
    int currentWidth = rect.right - rect.left;
    int currentHeight = rect.bottom - rect.top;

    // Set the minimum width and height in resourece.h 
    int minWidth = MIN_WIDTH;
    int minHeight = MIN_HEIGHT;


    // Determine if the window is smaller than the minimum size
    bool isSmallerThanMinSize = (currentWidth < minWidth) || (currentHeight < minHeight);

    // Define the UI element IDs to be shown or hidden based on the window size
    const int elementIds[] = {
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_SEARCH_MODE_GROUP, IDC_NORMAL_RADIO, IDC_EXTENDED_RADIO, IDC_REGEX_RADIO,
        IDC_SCOPE_GROUP, IDC_ALL_TEXT_RADIO, IDC_SELECTION_RADIO, IDC_COLUMN_MODE_RADIO ,IDC_DELIMITER_EDIT, IDC_COLUMN_NUM_EDIT, IDC_DELIMITER_STATIC, IDC_COLUMN_NUM_STATIC,
        IDC_SWAP_BUTTON, IDC_REPLACE_LIST, IDC_COPY_TO_LIST_BUTTON, IDC_USE_LIST_CHECKBOX,IDC_STATIC_FRAME,
        IDC_STATIC_FIND, IDC_STATIC_REPLACE, IDC_COLUMN_HIGHLIGHT_BUTTON,
        IDC_MATCH_CASE_CHECKBOX, IDC_WHOLE_WORD_CHECKBOX, IDC_WRAP_AROUND_CHECKBOX,
        IDC_LOAD_FROM_CSV_BUTTON, IDC_SAVE_TO_CSV_BUTTON,
        IDC_CLEAR_MARKS_BUTTON, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_FRAME,
        IDC_SHIFT_TEXT, IDC_STATUS_MESSAGE, IDC_EXPORT_BASH_BUTTON, 
        IDC_2_BUTTONS_MODE
    };

    // Show or hide elements based on the window size
    for (int id : elementIds) {
        ShowWindow(GetDlgItem(_hSelf, id), isSmallerThanMinSize ? SW_HIDE : SW_SHOW);
    }

    // Define the UI element IDs to be shown or hidden contrary to the window size
    const int oppositeElementIds[] = {
        IDC_STATIC_HINT
    };

    // Show or hide elements contrary to the window size
    for (int id : oppositeElementIds) {
        ShowWindow(GetDlgItem(_hSelf, id), isSmallerThanMinSize ? SW_SHOW : SW_HIDE);
    }

    // Check the checkbox state and decide which buttons to show
    if (!isSmallerThanMinSize) {
        updateButtonVisibilityBasedOnMode();
    }
    else {
        ShowWindow(GetDlgItem(_hSelf, IDC_FIND_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_FIND_NEXT_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_FIND_PREV_BUTTON), SW_HIDE);

        ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_SMALL_BUTTON), SW_HIDE);

        ShowWindow(GetDlgItem(_hSelf, IDC_MARK_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_MARK_MATCHES_BUTTON), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), SW_HIDE);
    }

    // Show or hide the "Cancel Longrun Button" based on the active state
    ShowWindow(GetDlgItem(_hSelf, IDC_CANCEL_LONGRUN_BUTTON), progressDisplayActive ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_PROGRESS_BAR), progressDisplayActive ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_STATUS_MESSAGE), progressDisplayActive ? SW_HIDE : SW_SHOW);
    ShowWindow(GetDlgItem(_hSelf, IDC_ACTION_DESCRIPTION), progressDisplayActive ? SW_SHOW : SW_HIDE);
}

#pragma endregion


#pragma region ListView

void MultiReplace::createListViewColumns(HWND listView) {
    LVCOLUMN lvc;

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvc.fmt = LVCFMT_LEFT;

    // Get the client rectangle
    RECT rcClient;
    GetClientRect(_hSelf, &rcClient);
    // Extract width from the RECT
    int windowWidth = rcClient.right - rcClient.left;

    // Calculate the remaining width for the first two columns
    int remainingWidth = windowWidth - 280;

    // Calculate the total width of columns 3 to 7
    int columns3to7Width = 30 * 7; // Assuming fixed width of 30 for columns 3 to 7

    remainingWidth -= columns3to7Width;

    lvc.iSubItem = 0;
    lvc.pszText = L"";
    lvc.cx = 0;
    ListView_InsertColumn(listView, 0, &lvc);

    lvc.iSubItem = 1;
    lvc.pszText = L"\u2610";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 1, &lvc);

    // Column for "Find" Text
    lvc.iSubItem = 2;
    lvc.pszText = L"Find";
    lvc.cx = remainingWidth / 2;
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(listView, 2, &lvc);

    // Column for "Replace" Text
    lvc.iSubItem = 3;
    lvc.pszText = L"Replace";
    lvc.cx = remainingWidth / 2;
    ListView_InsertColumn(listView, 3, &lvc);

    // Column for Option: Whole Word
    lvc.iSubItem = 4;
    lvc.pszText = L"W";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 4, &lvc);

    // Column for Option: Match Case
    lvc.iSubItem = 5;
    lvc.pszText = L"C";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 5, &lvc);

    // Column for Option: Normal
    lvc.iSubItem = 6;
    lvc.pszText = L"N";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 6, &lvc);

    // Column for Option: Extended
    lvc.iSubItem = 7;
    lvc.pszText = L"E";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 7, &lvc);

    // Column for Option: Regex
    lvc.iSubItem = 8;
    lvc.pszText = L"R";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 8, &lvc);

    // Column for Delete Button
    lvc.iSubItem = 9;
    lvc.pszText = L"";
    lvc.cx = 30;
    ListView_InsertColumn(listView, 9, &lvc);
}

void MultiReplace::insertReplaceListItem(const ReplaceItemData& itemData)
{
    // Return early if findText is empty
    if (itemData.findText.empty()) {
        showStatusMessage(L"No 'Find String' entered. Please provide a value to add to the list.", RGB(255, 0, 0));
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
        message = L"Duplicate entry: " + newItemData.findText;
    }
    else {
        message = L"Value added to the list.";
    }
    showStatusMessage(message.c_str(), RGB(0, 128, 0));

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Update Header if there might be any changes
    updateHeader();
}

void MultiReplace::updateListViewAndColumns(HWND listView, LPARAM lParam)
{
    // Get the new width and height of the window from lParam
    int newWidth = LOWORD(lParam);
    int newHeight = HIWORD(lParam);

    // Calculate the total width of columns 3 to 7
    int columns3to7Width = 0;
    for (int i = 4; i < 10; i++)
    {
        columns3to7Width += ListView_GetColumnWidth(listView, i);
    }
    columns3to7Width += 30; // for the first column

    // Calculate the remaining width for the first two columns
    int remainingWidth = newWidth - 280 - columns3to7Width;

    static int prevWidth = newWidth; // Store the previous width

    // If the window is horizontally maximized, update the IDC_REPLACE_LIST size first
    if (newWidth > prevWidth) {
        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_LIST), 14, 244, newWidth - 255, newHeight - 274, TRUE);
    }

    ListView_SetColumnWidth(listView, 2, remainingWidth / 2);
    ListView_SetColumnWidth(listView, 3, remainingWidth / 2);

    // If the window is horizontally minimized or vetically changed the size
    MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_LIST), 14, 244, newWidth - 255, newHeight - 274, TRUE);

    // If the window size hasn't changed, no need to do anything

    prevWidth = newWidth;
}

void MultiReplace::handleSelection(NMITEMACTIVATE* pnmia) {
    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }

    replaceListData[pnmia->iItem].isSelected = !replaceListData[pnmia->iItem].isSelected;

    // Update the header after changing the selection status of an item
    updateHeader();

    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
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
        showStatusMessage(L"No rows selected to shift.", RGB(255, 0, 0));
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
    showStatusMessage(std::to_wstring(selectedIndices.size()) + L" rows successfully shifted.", RGB(0, 128, 0));

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
    updateHeader();

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(L"1 line deleted.", RGB(0, 128, 0));
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
    updateHeader();

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(std::to_wstring(numDeletedLines) + L" lines deleted.", RGB(0, 128, 0));
}

void MultiReplace::sortReplaceListData(int column) {
    // Get the currently selected rows
    auto selectedRows = getSelectedRows();

    if (column == 2) {
        // Sort by `findText`
        std::sort(replaceListData.begin(), replaceListData.end(),
            [this](const ReplaceItemData& a, const ReplaceItemData& b) {
                if (this->ascending)
                    return a.findText < b.findText;
                else
                    return a.findText > b.findText;
            });
        std::wstring statusMessage = L"Find column sorted in " + std::wstring(ascending ? L"ascending" : L"descending") + L" order.";
        showStatusMessage(statusMessage.c_str(), RGB(0, 0, 255));
    }
    else if (column == 3) {
        // Sort by `replaceText`
        std::sort(replaceListData.begin(), replaceListData.end(),
            [this](const ReplaceItemData& a, const ReplaceItemData& b) {
                if (this->ascending)
                    return a.replaceText < b.replaceText;
                else
                    return a.replaceText > b.replaceText;
            });
        std::wstring statusMessage = L"Replace column sorted in " + std::wstring(ascending ? L"ascending" : L"descending") + L" order.";
        showStatusMessage(statusMessage.c_str(), RGB(0, 0, 255));
    }

    // Update the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);

    // Select the previously selected rows
    selectRows(selectedRows);
}

std::vector<ReplaceItemData> MultiReplace::getSelectedRows() {
    std::vector<ReplaceItemData> selectedRows;
    int i = -1;
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedRows.push_back(replaceListData[i]);
    }
    return selectedRows;
}

void MultiReplace::selectRows(const std::vector<ReplaceItemData>& rowsToSelect) {
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);  // deselect all items

    for (size_t i = 0; i < replaceListData.size(); i++) {
        for (const auto& row : rowsToSelect) {
            if (replaceListData[i] == row) {
                ListView_SetItemState(_replaceListView, i, LVIS_SELECTED, LVIS_SELECTED);
                break;
            }
        }
    }
}

void MultiReplace::handleCopyToListButton() {
    ReplaceItemData itemData;

    itemData.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
    itemData.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

    itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
    itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
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

#pragma endregion


#pragma region Dialog

INT_PTR CALLBACK MultiReplace::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{

    switch (message)
    {
    case WM_INITDIALOG:
    {
        initializeScintilla();
        initializePluginStyle();
        initializeCtrlMap();
        updateUIVisibility();
        initializeListView();
        loadSettings();
        return TRUE;
    }
    break;

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_STATUS_MESSAGE) ||
            hwndStatic == GetDlgItem(_hSelf, IDC_ACTION_DESCRIPTION)) {
            SetTextColor(hdcStatic, _statusMessageColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return (LRESULT)GetStockObject(NULL_BRUSH);
        }

        break;
    }
    break;

    case WM_DESTROY:
    {
        isLongRunCancelled = true;
        saveSettings();
        DestroyWindow(_hSelf);
        DeleteObject(hFont);
    }
    break;

    case WM_SIZE:
    {
        int newWidth = LOWORD(lParam);
        int newHeight = HIWORD(lParam);

        // Show Hint Message if not releted to the Window Size
        updateUIVisibility();

        // Move and resize the List
        updateListViewAndColumns(GetDlgItem(_hSelf, IDC_REPLACE_LIST), lParam);

        // Calculate Position for all Elements
        positionAndResizeControls(newWidth, newHeight);

        // Move all Elements
        moveAndResizeControls();

        return 0;
    }
    break;

    case WM_NOTIFY:
    {
        NMHDR* pnmh = (NMHDR*)lParam;
        if (pnmh->idFrom == IDC_REPLACE_LIST) {
            switch (pnmh->code) {
            case NM_CLICK:
            {
                NMITEMACTIVATE* pnmia = (NMITEMACTIVATE*)lParam;
                if (pnmia->iSubItem == 9) { // Delete button column
                    handleDeletion(pnmia);
                }
                if (pnmia->iSubItem == 1) { // Select button column
                    // get current selection status of the item
                    bool currentSelectionStatus = replaceListData[pnmia->iItem].isSelected;
                    // set the selection status to its opposite
                    setSelections(!currentSelectionStatus, true);
                }
            }
            break;

            case NM_DBLCLK:
            {
                NMITEMACTIVATE* pnmia = (NMITEMACTIVATE*)lParam;
                handleCopyBack(pnmia);
            }
            break;

            case LVN_GETDISPINFO:
            {
                NMLVDISPINFO* plvdi = (NMLVDISPINFO*)lParam;

                // Get the data from the vector
                ReplaceItemData& itemData = replaceListData[plvdi->item.iItem];

                // Display the data based on the subitem
                switch (plvdi->item.iSubItem)
                {

                case 1:
                    if (itemData.isSelected) {
                        plvdi->item.pszText = L"\u25A0";
                    }
                    else {
                        plvdi->item.pszText = L"\u2610";
                    }
                    break;
                case 2:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.findText.c_str());
                    break;
                case 3:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceText.c_str());
                    break;
                case 4:
                    if (itemData.wholeWord) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 5:
                    if (itemData.matchCase) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 6:
                    if (!itemData.regex && !itemData.extended) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 7:
                    if (itemData.extended) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 8:
                    if (itemData.regex) {
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 9:
                    plvdi->item.mask |= LVIF_TEXT;
                    plvdi->item.pszText = L"\u2716";
                    break;
                }
            }
            break;

            case LVN_COLUMNCLICK:
            {
                NMLISTVIEW* pnmv = (NMLISTVIEW*)lParam;

                // Check if the column 1 header was clicked
                if (pnmv->iSubItem == 1) {
                    setSelections(!allSelected);
                }

                // Check if the column "Find" or "Replace" header was clicked
                if (pnmv->iSubItem == 2 || pnmv->iSubItem == 3) {
                    if (lastColumn == pnmv->iSubItem) {
                        ascending = !ascending;
                    }
                    else {
                        lastColumn = pnmv->iSubItem;
                        ascending = true;
                    }
                    sortReplaceListData(lastColumn);
                }
                break;
            }

            case LVN_KEYDOWN:
            {
                LPNMLVKEYDOWN pnkd = reinterpret_cast<LPNMLVKEYDOWN>(pnmh);

                PostMessage(_replaceListView, WM_SETFOCUS, 0, 0);
                // Alt+A key for Select All
                if (pnkd->wVKey == 'A' && GetKeyState(VK_MENU) & 0x8000) {
                    setSelections(true, ListView_GetSelectedCount(_replaceListView) > 0);
                }
                // Alt+D key for Deselect All
                else if (pnkd->wVKey == 'D' && GetKeyState(VK_MENU) & 0x8000) {
                    setSelections(false, ListView_GetSelectedCount(_replaceListView) > 0);
                }
                else if (pnkd->wVKey == VK_DELETE) { // Delete key
                    deleteSelectedLines(_replaceListView);
                }
                else if ((GetKeyState(VK_MENU) & 0x8000) && (pnkd->wVKey == VK_UP)) { // Alt/AltGr + Up key
                    int iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                    if (iItem >= 0) {
                        NMITEMACTIVATE nmia;
                        nmia.iItem = iItem;
                        handleCopyBack(&nmia);
                    }
                }
                else if (pnkd->wVKey == VK_F12) { // F12 key
                    RECT windowRect;
                    GetClientRect(_hSelf, &windowRect);

                    HDC hDC = GetDC(_hSelf);
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
                }
                else if (pnkd->wVKey == VK_SPACE) { // Spacebar key
                    int iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                    if (iItem >= 0) {
                        // get current selection status of the item
                        bool currentSelectionStatus = replaceListData[iItem].isSelected;
                        // set the selection status to its opposite
                        setSelections(!currentSelectionStatus, true);
                    }
                }
                else if (pnkd->wVKey == VK_F10) {
                    findAllDelimitersInDocument();
                    isLoggingEnabled = true;
                    MessageBox(NULL, L"Delimiter List initialized", L"Notification", MB_OK);
                }
                else if (pnkd->wVKey == VK_F11) {
                    displayLogChangesInMessageBox();
                }
            }
            break;
            }
        }
        else
            DockingDlgInterface::run_dlgProc( message, wParam, lParam );
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
            isLongRunCancelled = false;
        }
        else {
            isLongRunCancelled = true;
        }
    }
    break;


    case WM_COMMAND:
    {

        switch (LOWORD(wParam))
        {

        case IDCANCEL :
        {
            if (_MultiReplace.isFloating())
            {
                isLongRunCancelled = true;
                EndDialog(_hSelf, 0);
                _MultiReplace.display(false);
            }
            else
            {
                ::SetFocus( getCurScintilla() );
            }
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
            handleClearColumnMarks();
            handleDelimiterPositions();
        }
        break;

        case IDC_SELECTION_RADIO:
        {
            setElementsState(columnRadioDependentElements, false);
            setElementsState(selectionRadioDisabledButtons, false);
            handleClearColumnMarks();
            handleDelimiterPositions();
        }
        break;

        case IDC_COLUMN_MODE_RADIO:
        {
            setElementsState(columnRadioDependentElements, true);
            setElementsState(selectionRadioDisabledButtons, true);
        }
        break;

        case IDC_COLUMN_HIGHLIGHT_BUTTON:
        {
            if (!isColumnHighlighted) {
                handleDelimiterPositions();
                if (!columnDelimiterData.columns.empty() && !columnDelimiterData.extendedDelimiter.empty()) {
                    handleHighlightColumnsInDocument();
                    LRESULT startPosition = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
                    showStatusMessage(L"Actual Position " + addLineAndColumnMessage(startPosition), RGB(0, 128, 0));
                }
            }
            else {
                handleClearColumnMarks();
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
            handleDelimiterPositions();
            handleFindNextButton();
        }
        break;

        case IDC_FIND_PREV_BUTTON:
        {
            handleDelimiterPositions();
            handleFindPrevButton(); 
        }
        break;

        case IDC_REPLACE_BUTTON:
        {
            handleReplaceButton();
        }
        break;

        case IDC_REPLACE_ALL_SMALL_BUTTON:
        case IDC_REPLACE_ALL_BUTTON:
        {
            handleDelimiterPositions();
            handleReplaceAllButton();
        }
        break;

        case IDC_MARK_MATCHES_BUTTON:
        case IDC_MARK_BUTTON:
        {
            handleDelimiterPositions();
            handleClearTextMarksButton();
            handleMarkMatchesButton();
        }
        break;

        case IDC_CLEAR_MARKS_BUTTON:
        {
            handleClearTextMarksButton();
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
            break;
        }
        case IDC_DOWN_BUTTON:
        {
            shiftListItem(_replaceListView, Direction::Down);
            break;
        }

        case IDC_EXPORT_BASH_BUTTON:
        {
            std::wstring filePath = openFileDialog(true, L"Bash Files (*.sh)\0*.sh\0All Files (*.*)\0*.*\0", L"Export as Bash", OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"sh");
            if (!filePath.empty()) {
                exportToBashScript(filePath);
            }
        }
        break;

        case IDC_CANCEL_LONGRUN_BUTTON:
        {
            isLongRunCancelled = true;
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
        showStatusMessage(L"Cannot replace. Document is read-only.", RGB(255, 0, 0));
        return;
    }

    int replaceCount = 0;
    // Check if the "In List" option is enabled
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);

    if (useListEnabled)
    {
        // Check if the replaceListData is empty and warn the user if so
        if (replaceListData.empty()) {
            showStatusMessage(L"Add values into the list. Or uncheck 'Use in List' to replace directly.", RGB(255, 0, 0));
            return;
        }
        ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
        for (size_t i = 0; i < replaceListData.size(); i++)
        {
            ReplaceItemData& itemData = replaceListData[i];
            if (itemData.isSelected) {
                replaceCount += replaceString(
                    itemData.findText, itemData.replaceText,
                    itemData.wholeWord, itemData.matchCase,
                    itemData.regex, itemData.extended);
            }
        }
        ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
    }
    else
    {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
        replaceCount = replaceString(findText, replaceText, wholeWord, matchCase, regex, extended);
        ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);

        // Add the entered text to the combo box history
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);
    }
    // Display status message
    showStatusMessage(std::to_wstring(replaceCount) + L" occurrences were replaced.", RGB(0, 128, 0));
}

void MultiReplace::handleReplaceButton() {
    // First check if the document is read-only
    LRESULT isReadOnly = ::SendMessage(_hScintilla, SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(L"Cannot replace. Document is read-only.", RGB(255, 0, 0));
        return;
    }

    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SearchResult searchResult;
    searchResult.pos = -1;
    searchResult.length = 0;
    searchResult.foundText = "";

    // Get the cursor position
    LRESULT cursorPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(L"Add values into the list. Or uncheck 'Use in List' to replace directly.", RGB(255, 0, 0));
            return;
        }

        // Get selected text to check against the list
        SelectionInfo selection = getSelectionInfo();

        // Search through replaceListData to find a match with the selected text
        for (const ReplaceItemData& itemData : replaceListData) {
            // Only perform the search if the item is selected
            if (itemData.isSelected) {
                std::wstring findText = itemData.findText;
                std::wstring replaceText = itemData.replaceText;
                bool wholeWord = itemData.wholeWord;
                bool matchCase = itemData.matchCase;
                bool regex = itemData.regex;
                bool extended = itemData.extended;

                std::string findTextUtf8 = convertAndExtend(findText, extended);

                // Define searchFlags
                int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

                // Search from the selection position
                searchResult = performSearchForward(findTextUtf8, searchFlags, selection.startPos, true);

                if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
                    // If it does match, replace the selected string
                    std::string replaceTextUtf8 = convertAndExtend(replaceText, extended);
                    performReplace(replaceTextUtf8, selection.startPos, selection.length);
                    showStatusMessage((L"Replaced '" + stringToWString(selection.text) + L"' with '" + replaceText + L"'.").c_str(), RGB(0, 128, 0));
                    break;
                }
            }
        }

        searchResult = performListSearchForward(replaceListData, cursorPos);

        // Check search results and handle wrap-around
        if (searchResult.pos < 0 && wrapAroundEnabled) {
            // If no match was found, and wrap-around is enabled, start the search again from the start
            searchResult = performListSearchForward(replaceListData, 0);
            if (searchResult.pos >= 0) {
                showStatusMessage(L"Wrapped, found match.", RGB(0, 128, 0));
            }
            else {
                showStatusMessage(L"No further matches found.", RGB(255, 0, 0));
            }
        }
        else if (searchResult.pos >= 0) {
            showStatusMessage((L"Found match for '" + stringToWString(searchResult.foundText) + L"'.").c_str(), RGB(0, 128, 0));
        }
        else {
            showStatusMessage(L"No matches found.", RGB(255, 0, 0));
        }
    }
    else {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        SelectionInfo selection = getSelectionInfo();
        std::string findTextUtf8 = convertAndExtend(findText, extended);

        // Define searchFlags before if block
        int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);
        searchResult = performSearchForward(findTextUtf8, searchFlags, selection.startPos, true);

        if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
            // If it does match, replace the selected string
            std::string replaceTextUtf8 = convertAndExtend(replaceText, extended);
            performReplace(replaceTextUtf8, selection.startPos, selection.length);
            showStatusMessage((L"Replaced '" + findText + L"' with '" + replaceText + L"'.").c_str(), RGB(0, 128, 0));

            // Continue search after replace
            searchResult = performSearchForward(findTextUtf8, searchFlags, searchResult.pos + searchResult.length, true);
        }

        // Check search results and handle wrap-around
        if (searchResult.pos < 0 && wrapAroundEnabled) {
            // If no match was found, and wrap-around is enabled, start the search again from the start
            searchResult = performSearchForward(findTextUtf8, searchFlags, 0, true);
            if (searchResult.pos >= 0) {
                showStatusMessage((L"Wrapped, found match for '" + findText + L"'.").c_str(), RGB(0, 128, 0));
            }
            else {
                showStatusMessage((L"No further matches found for '" + findText + L"'.").c_str(), RGB(255, 0, 0));
            }
        }
        else if (searchResult.pos >= 0) {
            showStatusMessage((L"Found match for '" + findText + L"'.").c_str(), RGB(0, 128, 0));
        }
        else {
            showStatusMessage((L"No matches found for '" + findText + L"'.").c_str(), RGB(255, 0, 0));
        }

    }
}

int MultiReplace::replaceString(const std::wstring& findText, const std::wstring& replaceText, bool wholeWord, bool matchCase, bool regex, bool extended)
{
    if (findText.empty()) {
        return 0;
    }

    int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

    std::string findTextUtf8 = convertAndExtend(findText, extended);
    std::string replaceTextUtf8 = convertAndExtend(replaceText, extended);

    int replaceCount = 0;  // Counter for replacements
    SearchResult searchResult = performSearchForward(findTextUtf8, searchFlags, false, 0);
    while (searchResult.pos >= 0)
    {
        
        Sci_Position newPos = performReplace(replaceTextUtf8, searchResult.pos, searchResult.length);
        replaceCount++;

        // Update Delimiter and Marking if enabled
        handleDelimiterPositions();

        searchResult = performSearchForward(findTextUtf8, searchFlags, false, newPos);
    }

    return replaceCount;
}

Sci_Position MultiReplace::performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    // Set the target range for the replacement
    ::SendMessage(_hScintilla, SCI_SETTARGETRANGE, pos, pos + length);

    // Get the codepage of the document
    int cp = static_cast<int>(::SendMessage(_hScintilla, SCI_GETCODEPAGE, 0, 0));

    // Convert the string from UTF-8 to the codepage of the document
    std::string replaceTextCp = utf8ToCodepage(replaceTextUtf8, cp);

    // Perform the replacement
    ::SendMessage(_hScintilla, SCI_REPLACETARGET, replaceTextCp.size(), reinterpret_cast<LPARAM>(replaceTextCp.c_str()));

    return pos + static_cast<Sci_Position>(replaceTextCp.length());
}

std::string MultiReplace::utf8ToCodepage(const std::string& utf8Str, int codepage) {
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

#pragma endregion


#pragma region Find

void MultiReplace::handleFindNextButton() {
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    LRESULT searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);

    if (useListEnabled)
    {
        if (replaceListData.empty()) {
            showStatusMessage(L"Add values into the list. Or uncheck 'Use in List' to find directly.", RGB(255, 0, 0));
            return;
        }

        SearchResult result = performListSearchForward(replaceListData, searchPos);

        if (result.pos >= 0) {
            showStatusMessage(L"" + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
        }
        else if (wrapAroundEnabled)
        {
            result = performListSearchForward(replaceListData, 0);
            if (result.pos >= 0) {
                showStatusMessage(L"Wrapped " + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
            }
        }
        else
        {
            showStatusMessage(L"No matches found.", RGB(255, 0, 0));
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
        SearchResult result = performSearchForward(findTextUtf8, searchFlags, true, searchPos);

        if (result.pos >= 0) {
            showStatusMessage(L"" + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
        }
        else if (wrapAroundEnabled)
        {
            result = performSearchForward(findTextUtf8, searchFlags, true, 0);
            if (result.pos >= 0) {
                showStatusMessage((L"Wrapped " + addLineAndColumnMessage(result.pos)).c_str(), RGB(0, 128, 0));
            }
        }
        else
        {
             showStatusMessage((L"No matches found for '" + findText + L"'.").c_str(), RGB(255, 0, 0));
        }

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
}

void MultiReplace::handleFindPrevButton() {

    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    LRESULT searchPos = ::SendMessage(_hScintilla, SCI_GETCURRENTPOS, 0, 0);
    searchPos = (searchPos > 0) ? ::SendMessage(_hScintilla, SCI_POSITIONBEFORE, searchPos, 0) : searchPos;

    if (useListEnabled)
    {
        if (replaceListData.empty()) {
            showStatusMessage(L"Add values into the list. Or uncheck 'Use in List' to find directly.", RGB(255, 0, 0));
            return;
        }

        SearchResult result = performListSearchBackward(replaceListData, searchPos);

        if (result.pos >= 0) {
            showStatusMessage(L"" + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
        }
        else if (wrapAroundEnabled)
        {
            result = performListSearchBackward(replaceListData, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
            if (result.pos >= 0) {
                showStatusMessage(L"Wrapped " + addLineAndColumnMessage(result.pos), RGB(0, 128, 0));
            }
            else {
                showStatusMessage(L"No matches found after wrap.", RGB(255, 0, 0));
            }
        }
        else
        {
            showStatusMessage(L"No matches found.", RGB(255, 0, 0));
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
                showStatusMessage((L"Wrapped '" + findText + L"'." + addLineAndColumnMessage(result.pos)).c_str(), RGB(0, 128, 0));
            }
            else {
                showStatusMessage((L"No matches found for '" + findText + L"' after wrap.").c_str(), RGB(255, 0, 0));
            }
        }
        else
        {
            showStatusMessage((L"No matches found for '" + findText + L"'.").c_str(), RGB(255, 0, 0));
        }

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
}

SearchResult MultiReplace::performSingleSearch(const std::string& findTextUtf8, int searchFlags, bool selectMatch, SelectionRange range) {

    ::SendMessage(_hScintilla, SCI_SETTARGETSTART, range.start, 0);
    ::SendMessage(_hScintilla, SCI_SETTARGETEND, range.end, 0);
    ::SendMessage(_hScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

    LRESULT pos = ::SendMessage(_hScintilla, SCI_SEARCHINTARGET, findTextUtf8.length(), reinterpret_cast<LPARAM>(findTextUtf8.c_str()));

    SearchResult result;
    result.pos = pos;
    result.length = 0;
    result.foundText = "";

    if (pos >= 0) {
        // If a match is found, set additional result data
        result.length = ::SendMessage(_hScintilla, SCI_GETTARGETEND, 0, 0) - pos;
        result.foundText = findTextUtf8;

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
    result.pos = -1;
    SelectionRange targetRange = { 0, 0 };

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
    else if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED &&
        !(columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty())) {

        // Identify Column to Start
        StartColumnInfo columnInfo = getStartColumnInfo(start);
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
                    if (column <= linePositions.size() + 1) {
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
                }
                // Reset startColumnIndex for the next lines
                startColumnIndex = 1;
            }
        }
    }
    else {
        // If neither IDC_SELECTION_RADIO nor IDC_COLUMN_MODE_RADIO, perform search within the whole document
        targetRange.start = start;
        targetRange.end = ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0);
        result = performSingleSearch(findTextUtf8, searchFlags, selectMatch, targetRange);
    }

    return result;
}

SearchResult MultiReplace::performSearchBackward(const std::string& findTextUtf8, int searchFlags, LRESULT start)
{
    SearchResult result;
    result.pos = -1;
    result.length = 0;
    result.foundText = "";

    SelectionRange targetRange;

    // Check if IDC_COLUMN_MODE_RADIO is enabled, and column delimiter data is set
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED &&
        !(columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty())) {

        // Identify Column to Start
        StartColumnInfo columnInfo = getStartColumnInfo(start);
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
                        startColumn = start;
                    }

                    targetRange = { startColumn, endColumn };
                    result = performSingleSearch(findTextUtf8, searchFlags, false, targetRange);

                    // Check if a match was found
                    if (result.pos >= 0) {
                        return result;
                    }
                }
            }
        }
    }
    else {
        ::SendMessage(_hScintilla, SCI_SETTARGETSTART, start, 0);
        ::SendMessage(_hScintilla, SCI_SETTARGETEND, 0, 0);
        ::SendMessage(_hScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);
        LRESULT pos = ::SendMessage(_hScintilla, SCI_SEARCHINTARGET, findTextUtf8.length(), reinterpret_cast<LPARAM>(findTextUtf8.c_str()));
        result.pos = pos;

        if (pos >= 0) {
            result.length = ::SendMessage(_hScintilla, SCI_GETTARGETEND, 0, 0) - pos;
            result.foundText = findTextUtf8;
            displayResultCentered(result.pos, result.pos + result.length, false);
        }
        else {
            result.length = 0;
        }
    }

    return result;
}

SearchResult MultiReplace::performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos)
{
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    for (const ReplaceItemData& itemData : list)
    {
        if (itemData.isSelected) {
            int searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
            std::string findTextUtf8 = convertAndExtend(itemData.findText, itemData.extended);
            SearchResult result = performSearchForward(findTextUtf8, searchFlags, false, cursorPos);

            // If a match was found and it's closer to the cursor than the current closest match, update the closest match
            if (result.pos >= 0 && (closestMatch.pos < 0 || result.pos < closestMatch.pos)) {
                closestMatch = result;
            }
        }
    }
    if (closestMatch.pos >= 0) { // Check if a match was found
        displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, true);
    }
    return closestMatch;
}

SearchResult MultiReplace::performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos)
{
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    for (const ReplaceItemData& itemData : list)
    {
        if (itemData.isSelected) {
            int searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
            std::string findTextUtf8 = convertAndExtend(itemData.findText, itemData.extended);
            SearchResult result = performSearchBackward(findTextUtf8, searchFlags, cursorPos);

            // If a match was found and it's closer to the cursor than the current closest match, update the closest match
            if (result.pos >= 0 && (closestMatch.pos < 0 || (result.pos + result.length) >(closestMatch.pos + closestMatch.length))) {
                closestMatch = result;
            }
        }
    }
    if (closestMatch.pos >= 0) { // Check if a match was found
        displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, false);
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
    int matchCount = 0;
    bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);
    markedStringsCount = 0;

    if (useListEnabled)
    {
        if (replaceListData.empty()) {
            showStatusMessage(L"Add values into the list. Or uncheck 'Use in List' to mark directly.", RGB(255, 0, 0));
            return;
        }

        for (ReplaceItemData& itemData : replaceListData)
        {
            if (itemData.isSelected) {
                matchCount += markString(
                    itemData.findText, itemData.wholeWord,
                    itemData.matchCase, itemData.regex, itemData.extended);
            }
        }
    }
    else
    {
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        bool regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        bool extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        matchCount = markString(findText, wholeWord, matchCase, regex, extended);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
    showStatusMessage(std::to_wstring(matchCount) + L" occurrences were marked.", RGB(0, 0, 128));
}

int MultiReplace::markString(const std::wstring& findText, bool wholeWord, bool matchCase, bool regex, bool extended)
{
    if (findText.empty()) {
        return 0;
    }

    int searchFlags = (wholeWord * SCFIND_WHOLEWORD) | (matchCase * SCFIND_MATCHCASE) | (regex * SCFIND_REGEXP);

    std::string findTextUtf8 = convertAndExtend(findText, extended);

    int markCount = 0;  // Counter for marked matches
    SearchResult searchResult = performSearchForward(findTextUtf8, searchFlags, false, 0);
    while (searchResult.pos >= 0)
    {
        highlightTextRange(searchResult.pos, searchResult.length, findTextUtf8);
        markCount++;
        searchResult = performSearchForward(findTextUtf8, searchFlags, false, searchResult.pos + searchResult.length); // Use nextPos for the next search
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
    showStatusMessage(L"All marks cleared.", RGB(0, 128, 0));
}

void MultiReplace::handleCopyMarkedTextToClipboardButton()
{
    bool wasLastCharMarked = false;
    size_t markedTextCount = 0;

    std::string markedText;
    std::string styleText;

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

                markedText += styleText;
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

    // Convert encoding to wide string
    std::wstring wstr = stringToWString(markedText);

    if (markedTextCount > 0)
    {
        // Open Clipboard
        OpenClipboard(0);
        EmptyClipboard();

        // Copy data to clipboard
        HGLOBAL hClipboardData = GlobalAlloc(GMEM_DDESHARE, sizeof(WCHAR) * (wstr.length() + 1));
        if (hClipboardData)
        {
            WCHAR* pchData;
            pchData = (WCHAR*)GlobalLock(hClipboardData);
            if (pchData)
            {
                wcscpy(pchData, wstr.c_str());
                GlobalUnlock(hClipboardData);

                // Checking SetClipboardData return
                if (SetClipboardData(CF_UNICODETEXT, hClipboardData) != NULL)
                {
                    int markedTextCountInt = static_cast<int>(markedTextCount);
                    showStatusMessage(std::to_wstring(markedTextCountInt) + L" marked blocks copied into Clipboard.", RGB(0, 128, 0));
                }
                else
                {
                    showStatusMessage(L"Failed to copy marked text to Clipboard.", RGB(255, 0, 0));
                }
            }
            else
            {
                showStatusMessage(L"Failed to allocate memory for Clipboard.", RGB(255, 0, 0));
                GlobalFree(hClipboardData);
            }
        }

        CloseClipboard();
    }
    else
    {
        showStatusMessage(L"No marked text to copy.", RGB(255, 0, 0));
    }
}

#pragma endregion


#pragma region Scope

bool MultiReplace::parseColumnAndDelimiterData() {
    // Clear previous data
    // delimiterPositionData = DelimiterPositionData{ 0, 0, 0 };

    std::wstring columnDataString = getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT);
    std::wstring delimiterData = getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT);

    columnDataString.erase(0, columnDataString.find_first_not_of(L','));
    columnDataString.erase(columnDataString.find_last_not_of(L',') + 1);

    if (columnDataString.empty() || delimiterData.empty()) {
        showStatusMessage(L"Column data or delimiter data is missing", RGB(255, 0, 0));
        columnDelimiterData.columns.clear();
        columnDelimiterData.extendedDelimiter = "";
        return false;
    }

    std::set<int> columns;
    std::wstring::size_type start = 0;
    std::wstring::size_type end = columnDataString.find(',', start);

    // Parse column data
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
                    showStatusMessage(L"Invalid range in column data", RGB(255, 0, 0));
                    columnDelimiterData.columns.clear();
                    columnDelimiterData.extendedDelimiter = "";
                    return false;
                }

                for (int i = startRange; i <= endRange; ++i) {
                    columns.insert(i);
                }
            }
            catch (const std::exception&) {
                showStatusMessage(L"Syntax error in column data", RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                columnDelimiterData.extendedDelimiter = "";
                return false;
            }
        }
        else {
            // Parse single column and add to set
            try {
                int column = std::stoi(block);

                if (column < 1) {
                    showStatusMessage(L"Invalid column number", RGB(255, 0, 0));
                    columnDelimiterData.columns.clear();
                    columnDelimiterData.extendedDelimiter = "";
                    return false;
                }

                columns.insert(column);
            }
            catch (const std::exception&) {
                showStatusMessage(L"Syntax error in column data", RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                columnDelimiterData.extendedDelimiter = "";
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
                showStatusMessage(L"Invalid range in column data", RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                columnDelimiterData.extendedDelimiter = "";
                return false;
            }

            for (int i = startRange; i <= endRange; ++i) {
                columns.insert(i);
            }
        }
        catch (const std::exception&) {
            showStatusMessage(L"Syntax error in column data", RGB(255, 0, 0));
            columnDelimiterData.columns.clear();
            columnDelimiterData.extendedDelimiter = "";
            return false;
        }
    }
    else {
        try {
            int column = std::stoi(lastBlock);

            if (column < 1) {
                showStatusMessage(L"Invalid column number", RGB(255, 0, 0));
                columnDelimiterData.columns.clear();
                columnDelimiterData.extendedDelimiter = "";
                return false;
            }

            columns.insert(column);
        }
        catch (const std::exception&) {
            showStatusMessage(L"Syntax error in column data", RGB(255, 0, 0));
            columnDelimiterData.columns.clear();
            columnDelimiterData.extendedDelimiter = "";
            return false;
        }
    }

    // Set dataChanged flag
    columnDelimiterData.delimiterChanged = !(columnDelimiterData.extendedDelimiter == convertAndExtend(delimiterData, true));

    columnDelimiterData.columns = columns;
    columnDelimiterData.extendedDelimiter = convertAndExtend(delimiterData, true);
    columnDelimiterData.delimiterLength = columnDelimiterData.extendedDelimiter.length();
    return true;
}

void MultiReplace::findAllDelimitersInDocument() {

    // Clear list for new data
    lineDelimiterPositions.clear();

    // Reset TextModiefeid Trigger
    textModified = false;
    logChanges.clear();

    // Get total line count in document
    LRESULT totalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);

    // Resize the list to fit total lines
    lineDelimiterPositions.resize(totalLines);

    // Find and store delimiter positions for each line
    for (LRESULT line = 0; line < totalLines; ++line) {

        processLogForDelimiters();

        // If isLongRunCancelled is true, clear lineDelimiterPositions and break the loop
        if (isLongRunCancelled) {
            lineDelimiterPositions.clear();
            break;
        }

        // Adjust totalLines according to changes in document
        LRESULT newTotalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);
        if (newTotalLines != totalLines) {
            lineDelimiterPositions.resize(newTotalLines);
            totalLines = newTotalLines;
        }

        // Find delimiters in line
        findDelimitersInLine(line);

        // Update progress in status
        displayProgressInStatus(line, totalLines, L"Loading CSV");

        // Adjust totalLines according to changes in document
        totalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);
    }

    // Clear log queue
    logChanges.clear();

    // Hide the cancel button and reset the cancellation flag
    resetProgressBar();
}

void MultiReplace::findDelimitersInLine(LRESULT line) {
    // Initialize LineInfo for this line
    LineInfo lineInfo;

    // Get start and end positions of the line
    lineInfo.startPosition = ::SendMessage(_hScintilla, SCI_POSITIONFROMLINE, line, 0);
    lineInfo.endPosition = ::SendMessage(_hScintilla, SCI_GETLINEENDPOSITION, line, 0);

    // Get line length and allocate buffer
    LRESULT lineLength = ::SendMessage(_hScintilla, SCI_LINELENGTH, line, 0);
    char* buf = new char[lineLength + 1];

    // Get line content
    ::SendMessage(_hScintilla, SCI_GETLINE, line, reinterpret_cast<LPARAM>(buf));
    std::string lineContent = buf;
    delete[] buf;

    // Define structure to store delimiter position
    DelimiterPosition delimiterPos;

    bool inQuotes = false;
    std::string::size_type pos = 0;

    while (pos < lineContent.size()) {
        if (lineContent[pos] == '\"') {
            // Toggle the flag if quote is found
            inQuotes = !inQuotes;
        }
        else if (!inQuotes && lineContent.compare(pos, columnDelimiterData.extendedDelimiter.size(), columnDelimiterData.extendedDelimiter) == 0) {
            // Only store position of delimiter if it's not inside quotes
            delimiterPos.position = pos + lineInfo.startPosition;
            lineInfo.positions.push_back(delimiterPos);
            pos += columnDelimiterData.extendedDelimiter.size() - 1; // Skip delimiter for next iteration
        }
        pos++;
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

StartColumnInfo MultiReplace::getStartColumnInfo(LRESULT startPosition) {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED ||
        columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty()) {
        return { 0, 0, 0 };
    }

    LRESULT totalLines = ::SendMessage(_hScintilla, SCI_GETLINECOUNT, 0, 0);
    LRESULT startLine = ::SendMessage(_hScintilla, SCI_LINEFROMPOSITION, startPosition, 0);
    SIZE_T startColumnIndex = 1;

    // Check if the line exists in lineDelimiterPositions
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());
    if (startLine < totalLines && startLine < listSize) {
        const auto& linePositions = lineDelimiterPositions[startLine].positions;

        for (SIZE_T i = 0; i < linePositions.size(); ++i) {
            if (startPosition < linePositions[i].position) {
                startColumnIndex = i + 1;
                break;
            }
            else if (i == linePositions.size() - 1 && startPosition >= linePositions[i].position) {
                startColumnIndex = i + 2;  // We're in the last column
                break;
            }
        }
    }
    return { totalLines, startLine, startColumnIndex };
}

void MultiReplace::initializeColumnStyles() {
    for (SIZE_T column = 1; column <= columnStyles.size(); column++) {
        // Calculate indicator style
        int indicatorStyle = columnStyles[(column - 1) % columnStyles.size()];

        // Assign color for the style
        long color = columnColors[(column - 1) % columnColors.size()];
        if (colorToStyleMap.find(color) == colorToStyleMap.end()) {
            colorToStyleMap[color] = indicatorStyle;
        }

        // Set the indicator style
        ::SendMessage(_hScintilla, SCI_INDICSETSTYLE, indicatorStyle, INDIC_STRAIGHTBOX);
        ::SendMessage(_hScintilla, SCI_INDICSETFORE, indicatorStyle, color);
        ::SendMessage(_hScintilla, SCI_INDICSETALPHA, indicatorStyle, 100);
    }
}

void MultiReplace::handleHighlightColumnsInDocument() {
    // Return early if columnDelimiterData is empty
    if (columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty()) {
        return;
    }

    // Convert size of lineDelimiterPositions to signed integer
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());

    // Initialize column styles
    initializeColumnStyles();

    // Iterate over each line's delimiter positions
    LRESULT line = 0;
    while (line < static_cast<LRESULT>(lineDelimiterPositions.size()) && !isLongRunCancelled) {
        processLogForDelimiters();
        highlightColumnsInLine(line);
        displayProgressInStatus(line, listSize, L"Higlight Columns");
        ++line;
    }

    // Hide the cancel button and reset the cancellation flag
    resetProgressBar();

    isColumnHighlighted = true;
}

void MultiReplace::highlightColumnsInLine(LRESULT line) {
    const auto& lineInfo = lineDelimiterPositions[line];

    // If no delimiter present, highlight whole line as first column
    if (lineInfo.positions.empty() &&
        std::find(columnDelimiterData.columns.begin(), columnDelimiterData.columns.end(), 1) != columnDelimiterData.columns.end())
    {
        // Highlight text range
        highlightColumnRange(lineInfo.startPosition, lineInfo.endPosition, 1);
    }
    else {
        // Highlight specific columns from columnDelimiterData
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= lineInfo.positions.size() + 1) {
                LRESULT start = 0;
                LRESULT end = 0;

                // Set start and end positions based on column index
                if (column == 1) {
                    start = lineInfo.startPosition;
                }
                else {
                    start = lineInfo.positions[column - 2].position + columnDelimiterData.delimiterLength;
                }

                if (column == lineInfo.positions.size() + 1) {
                    end = lineInfo.endPosition;
                }
                else {
                    end = lineInfo.positions[column - 1].position;
                }

                // Highlight the text range
                highlightColumnRange(start, end, column);
            }
        }
    }
}

void MultiReplace::clearMarksInLine(LRESULT line) {
    LRESULT startPos = ::SendMessage(_hScintilla, SCI_POSITIONFROMLINE, line, 0);
    LRESULT lineLength = ::SendMessage(_hScintilla, SCI_LINELENGTH, line, 0);

    for (int style : columnStyles) {
        ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, style, 0);
        ::SendMessage(_hScintilla, SCI_INDICATORCLEARRANGE, startPos, lineLength);
    }
}

void MultiReplace::highlightColumnRange(LRESULT start, LRESULT end, SIZE_T column) {
    // Skip if end is only one position higher than start
    if (end - start <= 1) {
        return;
    }

    // Get the indicator style for the column
    int indicatorStyle = colorToStyleMap[columnColors[(column - 1) % columnColors.size()]];

    // Set the current indicator and apply highlighting style
    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, indicatorStyle, 0);
    ::SendMessage(_hScintilla, SCI_INDICATORFILLRANGE, start, end - start);
}

void MultiReplace::handleClearColumnMarks() {
    for (int style : columnStyles) {
        ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, style, 0);
        ::SendMessage(_hScintilla, SCI_INDICATORCLEARRANGE, 0, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
    }

    showStatusMessage(L"Column marks cleared.", RGB(0, 128, 0));
    isColumnHighlighted = false;
}

std::wstring MultiReplace::addLineAndColumnMessage(LRESULT pos) {
    if (columnDelimiterData.columns.empty() && columnDelimiterData.extendedDelimiter.empty()) {
        return L"";
    }
    std::wstring lineAndColumnMessage;
    StartColumnInfo startInfo = getStartColumnInfo(pos);
    lineAndColumnMessage = L" (Line: " + std::to_wstring(startInfo.startLine + 1) +
                           L", Column: " + std::to_wstring(startInfo.startColumnIndex) + L")";
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
            this->messageBoxContent += "Line " + std::to_string(static_cast<int>(logEntry.lineNumber)) + " inserted.\n";
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
            this->messageBoxContent += "Line " + std::to_string(static_cast<int>(logEntry.lineNumber)) + " deleted.\n";
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
                clearMarksInLine(modifyLogEntry.lineNumber);
                highlightColumnsInLine(modifyLogEntry.lineNumber);
            }
            this->messageBoxContent += "Line " + std::to_string(static_cast<int>(modifyLogEntry.lineNumber)) + " modified.\n";
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

            // Update the highlight if necessary
            if (isColumnHighlighted) {
                highlightColumnsInLine(lineNumber);
            }

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

void MultiReplace::handleDelimiterPositions() {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED) {
        lineDelimiterPositions.clear();
        isLoggingEnabled = false;
        textModified = false;
        logChanges.clear();  // not really necessary
        return;
    }

    int currentBufferID = static_cast<int>(::SendMessage(_hScintilla, NPPM_GETCURRENTBUFFERID, 0, 0));

    // Create a message box to display the current buffer ID
    std::wstring bufferIDMessage = L"Current Buffer ID: " + std::to_wstring(currentBufferID);
    //MessageBox(NULL, bufferIDMessage.c_str(), L"Buffer ID Information", MB_OK);

    if (!parseColumnAndDelimiterData()) {
        return;
    }

    // Prevent executing the operations if the columns set or extendedDelimiter string is empty
    if (columnDelimiterData.columns.empty() || columnDelimiterData.extendedDelimiter.empty()) {
        return;
    }

    // Check and update EOL length if necessary
    LRESULT updatedEolLength = updateEOLLength();
    bool eolLengthChanged = (updatedEolLength != eolLength);

    if (eolLengthChanged) eolLength = updatedEolLength;

    if (eolLengthChanged || currentBufferID != scannedDelimiterBufferID || lineDelimiterPositions.empty() || columnDelimiterData.delimiterChanged) {
        isLoggingEnabled = true;  // Enable detailed logging
        findAllDelimitersInDocument();
        scannedDelimiterBufferID = currentBufferID;
    } else {
        processLogForDelimiters();
        textModified = false; // Reset the textModified flag after processing changes
    }
}

void MultiReplace::displayLogChangesInMessageBox() {

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
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
    std::string startContent = listToString(lineDelimiterPositions);
    std::wstring wideStartContent = converter.from_bytes(startContent);
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

    std::wstring logChangesWStr(logChangesStr.begin(), logChangesStr.end());
    MessageBox(NULL, logChangesWStr.c_str(), L"Log Changes", MB_OK);

    processLogForDelimiters();
    std::wstring wideMessageBoxContent = converter.from_bytes(messageBoxContent);
    MessageBox(NULL, wideMessageBoxContent.c_str(), L"Final Result", MB_OK);
    messageBoxContent.clear();

    std::string endContent = listToString(lineDelimiterPositions);
    std::wstring wideEndContent = converter.from_bytes(endContent);
    MessageBox(NULL, wideEndContent.c_str(), L"Content at the end", MB_OK);

    logChanges.clear();
}

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
            replaceListData[i].isSelected = select;
        }
    } while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_ALL)) != -1);

    // Update the allSelected flag if all items were selected/deselected
    if (!onlySelected) {
        allSelected = select;
    }

    // Update the header after changing the selection status of the items
    updateHeader();

    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

void MultiReplace::updateHeader() {
    bool anySelected = false;
    allSelected = !replaceListData.empty();

    // Check if any or all items in the replaceListData vector are selected
    for (const auto& itemData : replaceListData) {
        if (itemData.isSelected) {
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
        lvc.pszText = L"\u2610"; // Ballot box
    }

    ListView_SetColumn(_replaceListView, 1, &lvc);
}

void MultiReplace::showStatusMessage(const std::wstring& messageText, COLORREF color)
{
    const size_t MAX_DISPLAY_LENGTH = 60; // Maximum length of the message to be displayed

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

std::wstring MultiReplace::addProgressBarMessage(LRESULT currentLine, LRESULT totalLines, const std::wstring& message, LRESULT threshold) {
    static std::wstring lastProgressBarMessage;
    const int progressBarWidth = 10;  // Width of the progress bar in characters

    if (totalLines > threshold && currentLine % (totalLines / 100) == 0) {
        int percentageComplete = static_cast<int>((static_cast<double>(currentLine) / totalLines) * 100);
        if (percentageComplete % 5 == 0) { // Only update every 5%
            // Only construct a new message when the percentage changes
            std::wstring progressBar = L"[" + std::wstring((percentageComplete / 10), '#') +
                std::wstring((progressBarWidth - (percentageComplete / 10)), ' ') + L"]";
            lastProgressBarMessage = message + progressBar + std::to_wstring(percentageComplete) + L"%";
        }
    }
    // Return the last message constructed, regardless of whether it changed this iteration
    return lastProgressBarMessage;
}

void MultiReplace::displayProgressInStatus(LRESULT current, LRESULT total, const std::wstring& message) {
    static std::wstring lastProgressBarMessage;
    static bool showProgressBar = false;

    if (current == 0) {
        // Reset the progress bar and decision when a new process starts
        progressDisplayActive = total >= 100;

        // Set the text color
        _statusMessageColor = RGB(0, 0, 128);

        // Show or hide the progress display components
        ShowWindow(GetDlgItem(_hSelf, IDC_CANCEL_LONGRUN_BUTTON), progressDisplayActive ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_PROGRESS_BAR), progressDisplayActive ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_STATUS_MESSAGE), progressDisplayActive ? SW_HIDE : SW_SHOW);
        ShowWindow(GetDlgItem(_hSelf, IDC_ACTION_DESCRIPTION), progressDisplayActive ? SW_SHOW : SW_HIDE);

        // Update the action description message
        ::SetWindowText(GetDlgItem(_hSelf, IDC_ACTION_DESCRIPTION), message.c_str());

        // Disable all elements except the Cancel Long Run button
        setElementsState(allElementsExceptCancelLongRun, false, true);
    }

    if (progressDisplayActive) {
        int percentageComplete = static_cast<int>((static_cast<double>(current) / total) * 100);

        // Update the progress bar
        ::SendMessage(GetDlgItem(_hSelf, IDC_PROGRESS_BAR), PBM_SETPOS, (WPARAM)percentageComplete, 0);
    }

    // Process all pending messages
    MSG msg;
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

void MultiReplace::resetProgressBar() {
    // Reset isLongRunCancelled to false, hide Button
    isLongRunCancelled = false;
    progressDisplayActive = false;
    ShowWindow(GetDlgItem(_hSelf, IDC_CANCEL_LONGRUN_BUTTON), SW_HIDE);

    // Hide the progress bar
    ShowWindow(GetDlgItem(_hSelf, IDC_PROGRESS_BAR), SW_HIDE);

    // Hide the action description
    ShowWindow(GetDlgItem(_hSelf, IDC_ACTION_DESCRIPTION), SW_HIDE);

    // Show the status message control
    ShowWindow(GetDlgItem(_hSelf, IDC_STATUS_MESSAGE), SW_SHOW);

    // Enable all elements except the Cancel Long Run button
    setElementsState(allElementsExceptCancelLongRun, true, true);
}

LRESULT MultiReplace::updateEOLLength() {
    LRESULT eolMode = ::SendMessage(getScintillaHandle(), SCI_GETEOLMODE, 0, 0);
    LRESULT currentEolLength;

    switch (eolMode) {
    case SC_EOL_CRLF: // CRLF
        currentEolLength = 2;
        break;
    case SC_EOL_CR: // CR
    case SC_EOL_LF: // LF
        currentEolLength = 1;
        break;
    default:
        currentEolLength = 2; // Default to CRLF if unknown
        break;
    }

    return currentEolLength;
}

void MultiReplace::setElementsState(const std::vector<int>& elements, bool enable, bool restoreOriginalState) {
    // Store the state and disable the elements if restoreOriginalState is true and enable is false
    if (restoreOriginalState && !enable) {
        for (int id : elements) {
            bool isEnabled = IsWindowEnabled(GetDlgItem(_hSelf, id));
            stateSnapshot[id] = isEnabled;
            // Disable the current element
            EnableWindow(GetDlgItem(_hSelf, id), FALSE);
        }
    }
    // Enable or disable the elements if restoreOriginalState is false
    else if (!restoreOriginalState) {
        for (int id : elements) {
            if (enable) {
                // Enable the current element
                EnableWindow(GetDlgItem(_hSelf, id), TRUE);
            }
            else {
                // Disable the current element
                EnableWindow(GetDlgItem(_hSelf, id), FALSE);
            }
        }
    }

    // Restore the state if restoreOriginalState is true and enable is true
    if (restoreOriginalState && enable) {
        for (const auto& item : stateSnapshot) {
            EnableWindow(GetDlgItem(_hSelf, item.first), item.second ? TRUE : FALSE);
        }
        // Clear the snapshot after restoring it
        stateSnapshot.clear();
    }
}

#pragma endregion


#pragma region StringHandling

std::wstring MultiReplace::stringToWString(const std::string& rString) {
    int codePage = static_cast<int>(::SendMessage(_hScintilla, SCI_GETCODEPAGE, 0, 0));

    int requiredSize = MultiByteToWideChar(codePage, 0, rString.c_str(), -1, NULL, 0);
    if (requiredSize == 0)
        return std::wstring();

    std::vector<wchar_t> wideStringResult(requiredSize);
    MultiByteToWideChar(codePage, 0, rString.c_str(), -1, &wideStringResult[0], requiredSize);

    return std::wstring(&wideStringResult[0]);
}

std::string MultiReplace::wstringToString(const std::wstring& input) {
    if (input.empty()) return std::string();

    int codePage = static_cast<int>(::SendMessage(_hScintilla, SCI_GETCODEPAGE, 0, 0));
    if (codePage == 0) codePage = CP_ACP;

    int size_needed = WideCharToMultiByte(codePage, 0, &input[0], (int)input.size(), NULL, 0, NULL, NULL);
    if (size_needed == 0) return std::string();

    std::string strResult(size_needed, 0);
    WideCharToMultiByte(codePage, 0, &input[0], (int)input.size(), &strResult[0], size_needed, NULL, NULL);

    return strResult;
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
    std::wofstream outFile(filePath);
    outFile.imbue(std::locale(outFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>()));

    if (!outFile.is_open()) {
        return false;
    }

    // Write CSV header
    outFile << L"Selected,Find,Replace,WholeWord,MatchCase,Regex,Extended" << std::endl;

    // If list is empty, only the header will be written to the file
    if (!list.empty()) {
        // Write list items to CSV file
        for (const ReplaceItemData& item : list) {
            outFile << item.isSelected << L"," << escapeCsvValue(item.findText) << L"," << escapeCsvValue(item.replaceText) << L"," << item.wholeWord << L"," << item.matchCase << L"," << item.extended << L"," << item.regex << std::endl;
        }
    }

    outFile.close();
    return true;
}

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (!saveListToCsvSilent(filePath, list)) {
        showStatusMessage(L"Error: Unable to open file for writing.", RGB(255, 0, 0));
        return;
    }

    showStatusMessage(std::to_wstring(list.size()) + L" items saved to CSV.", RGB(0, 128, 0));

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
}

bool MultiReplace::loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list) {
    std::wifstream inFile(filePath);
    inFile.imbue(std::locale(inFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>()));

    if (!inFile.is_open()) {
        return false;
    }

    list.clear(); // Clear the existing list

    std::wstring line;
    std::getline(inFile, line); // Skip the CSV header

    // Read list items from CSV file
    while (std::getline(inFile, line)) {
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

        // Check if the row has the correct number of columns
        if (columns.size() != 7) {
            continue;
        }

        ReplaceItemData item;

        // Assign columns to item properties
        item.isSelected = std::stoi(columns[0]) != 0;
        item.findText = columns[1];
        item.replaceText = columns[2];
        item.wholeWord = std::stoi(columns[3]) != 0;
        item.matchCase = std::stoi(columns[4]) != 0;
        item.extended = std::stoi(columns[5]) != 0;
        item.regex = std::stoi(columns[6]) != 0;

        // Add the item to the list
        list.push_back(item);
    }

    inFile.close();
    return true;
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {
    if (!loadListFromCsvSilent(filePath, replaceListData)) {
        showStatusMessage(L"Error: Unable to open file for reading.", RGB(255, 0, 0));
        return;
    }

    updateHeader();
    // Update the list view control
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    if (replaceListData.empty())
    {
        showStatusMessage(L"No valid items found in the CSV file.", RGB(255, 0, 0)); // Red color
    }
    else
    {
        showStatusMessage(std::to_wstring(replaceListData.size()) + L" items loaded from CSV.", RGB(0, 128, 0)); // Green color
    }

    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

std::wstring MultiReplace::escapeCsvValue(const std::wstring& value) {
    std::wstring escapedValue;
    bool needsQuotes = false;

    for (const wchar_t& ch : value) {
        // Check if value contains any character that requires escaping
        if (ch == L',' || ch == L'"') {
            needsQuotes = true;
        }
        escapedValue += ch;
        // Escape double quotes by adding another double quote
        if (ch == L'"') {
            escapedValue += ch;
        }
    }

    // Enclose the value in double quotes if necessary
    if (needsQuotes) {
        escapedValue = L'"' + escapedValue + L'"';
    }

    return escapedValue;
}

std::wstring MultiReplace::unescapeCsvValue(const std::wstring& value) {
    std::wstring unescapedValue;
    bool insideQuotes = false;

    for (size_t i = 0; i < value.length(); ++i) {
        wchar_t ch = value[i];

        if (ch == L'"') {
            insideQuotes = !insideQuotes;
            if (insideQuotes) {
                // Ignore the leading quote
                continue;
            }
            else {
                // Check for escaped quotes (two consecutive quotes)
                if (i + 1 < value.length() && value[i + 1] == L'"') {
                    unescapedValue += ch;
                    ++i; // Skip the next quote
                }
                // Otherwise, ignore the trailing quote
            }
        }
        else {
            unescapedValue += ch;
        }
    }

    return unescapedValue;
}

#pragma endregion


#pragma region Export

void MultiReplace::exportToBashScript(const std::wstring& fileName) {
    std::ofstream file(wstringToString(fileName));
    if (!file.is_open()) {
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
        if (!itemData.isSelected) continue; // Skip if this item is not selected

        std::string find;
        std::string replace;
        if (itemData.extended) {
            find = translateEscapes(escapeSpecialChars(wstringToString(itemData.findText), true));
            replace = translateEscapes(escapeSpecialChars(wstringToString(itemData.replaceText), true));
        }
        else if (itemData.regex) {
            find = wstringToString(itemData.findText);
            replace = wstringToString(itemData.replaceText);
        }
        else {
            find = escapeSpecialChars(wstringToString(itemData.findText), false);
            replace = escapeSpecialChars(wstringToString(itemData.replaceText), false);
        }

        std::string wholeWord = itemData.wholeWord ? "1" : "0";
        std::string matchCase = itemData.matchCase ? "1" : "0";
        std::string normal = (!itemData.regex && !itemData.extended) ? "1" : "0";
        std::string extended = itemData.extended ? "1" : "0";
        std::string regex = itemData.regex ? "1" : "0";

        file << "processLine \"" << find << "\" \"" << replace << "\" " << wholeWord << " " << matchCase << " " << normal << " " << extended << " " << regex << "\n";
    }

    file.close();

    showStatusMessage(L"List exported to BASH script.", RGB(0, 128, 0));

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

    handleEscapeSequence(unicodeRegex, input, output, [](const std::string& unicodeEscape) {
        int codepoint = std::stoi(unicodeEscape.substr(2), nullptr, 16);
        std::wstring_convert<std::codecvt_utf8<char32_t>, char32_t> convert;
        return convert.to_bytes(static_cast<char32_t>(codepoint)).front();
        });

    output = std::regex_replace(output, newlineRegex, "__NEWLINE__");
    output = std::regex_replace(output, carriageReturnRegex, "__CARRIAGERETURN__");
    output = std::regex_replace(output, nullCharRegex, "");  // \0 will not be supported

    return output;
}

#pragma endregion


#pragma region INI

void MultiReplace::saveSettingsToIni(const std::wstring& iniFilePath) {
    std::wofstream outFile(iniFilePath);
    outFile.imbue(std::locale(outFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>()));

    if (!outFile.is_open()) {
        throw std::runtime_error("Could not open settings file for writing.");
    }

    // Store the current "Find what" and "Replace with" texts
    std::wstring currentFindTextData = L"\"" + getTextFromDialogItem(_hSelf, IDC_FIND_EDIT) + L"\"";
    std::wstring currentReplaceTextData = L"\"" + getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT) + L"\"";
    outFile << L"[Current]\n";
    outFile << L"FindText=" << currentFindTextData << L"\n";
    outFile << L"ReplaceText=" << currentReplaceTextData << L"\n";

    // Prepare and Store the current options
    int wholeWord = IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int matchCase = IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int extended = IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED ? 1 : 0;
    int regex = IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? 1 : 0;
    int wrapAround = IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED ? 1 : 0;
    int ButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED ? 1 : 0;
    int useList = IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED ? 1 : 0;

    // Store Options
    outFile << L"[Options]\n";
    outFile << L"WholeWord=" << wholeWord << L"\n";
    outFile << L"MatchCase=" << matchCase << L"\n";
    outFile << L"Extended=" << extended << L"\n";
    outFile << L"Regex=" << regex << L"\n";
    outFile << L"WrapAround=" << wrapAround << L"\n";
    outFile << L"ButtonsMode=" << ButtonsMode << L"\n";
    outFile << L"UseList=" << useList << L"\n";

    // Prepare and Store the scope options
    int selection = IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED ? 1 : 0;
    int columnMode = IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED ? 1 : 0;
    std::wstring columnNum = L"\"" + getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT) + L"\"";
    std::wstring delimiter = L"\"" + getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT) + L"\"";

    // Store Scope
    outFile << L"[Scope]\n";
    outFile << L"Selection=" << selection << L"\n";
    outFile << L"ColumnMode=" << columnMode << L"\n";
    outFile << L"ColumnNum=" << columnNum << L"\n";
    outFile << L"Delimiter=" << delimiter << L"\n";

    // Store "Find what" history
    LRESULT findWhatCount = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETCOUNT, 0, 0);
    outFile << L"[History]\n";
    outFile << L"FindTextHistoryCount=" << findWhatCount << L"\n";
    for (LRESULT i = 0; i < findWhatCount; i++) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring findTextData = L"\"" + std::wstring(buffer.data()) + L"\"";
        outFile << L"FindTextHistory" << i << L"=" << findTextData << L"\n";
    }

    // Store "Replace with" history
    LRESULT replaceWithCount = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETCOUNT, 0, 0);
    outFile << L"ReplaceTextHistoryCount=" << replaceWithCount << L"\n";
    for (LRESULT i = 0; i < replaceWithCount; i++) {
        LRESULT len = SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1)); // +1 for the null terminator
        SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring replaceTextData = L"\"" + std::wstring(buffer.data()) + L"\"";
        outFile << L"ReplaceTextHistory" << i << L"=" << replaceTextData << L"\n";
    }

    outFile.close();
}

void MultiReplace::saveSettings() {
    static bool settingsSaved = false;
    if (settingsSaved) {
        return;  // Check as WM_DESTROY will be 28 times triggered
    }

    // Get the path to the plugin's configuration file
    wchar_t configDir[MAX_PATH] = {}; // Initialize all elements to 0
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM)configDir);
    configDir[MAX_PATH - 1] = '\0'; // Ensure the configDir is null-terminated

    // Form the path to the INI file
    std::wstring iniFilePath = std::wstring(configDir) + L"\\MultiReplace.ini";
    std::wstring csvFilePath = std::wstring(configDir) + L"\\MultiReplaceList.ini";

    // Try to save the settings in the INI file
    try {
        saveSettingsToIni(iniFilePath);
        saveListToCsvSilent(csvFilePath, replaceListData);
    }
    catch (const std::exception& ex) {
        // If an error occurs while writing to the INI file, we show an error message
        std::wstring errorMessage = L"An error occurred while saving the settings: ";
        errorMessage += std::wstring(ex.what(), ex.what() + strlen(ex.what()));
        MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
    settingsSaved = true;
}

void MultiReplace::loadSettingsFromIni(const std::wstring& iniFilePath) {

    std::wifstream inFile(iniFilePath);
    inFile.imbue(std::locale(inFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>()));

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

    bool extended = readBoolFromIniFile(iniFilePath, L"Options", L"Extended", false);
    bool regex = readBoolFromIniFile(iniFilePath, L"Options", L"Regex", false);

    // Select the appropriate radio button based on the settings
    if (regex) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_REGEX_RADIO);
    }
    else if (extended) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_EXTENDED_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);
    }

    bool wrapAround = readBoolFromIniFile(iniFilePath, L"Options", L"WrapAround", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WRAP_AROUND_CHECKBOX), BM_SETCHECK, wrapAround ? BST_CHECKED : BST_UNCHECKED, 0);

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
    EnableWindow(GetDlgItem(_hSelf, IDC_COLUMN_HIGHLIGHT_BUTTON), columnModeSelected);

    std::wstring columnNum = readStringFromIniFile(iniFilePath, L"Scope", L"ColumnNum", L"");
    setTextInDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT, columnNum);

    std::wstring delimiter = readStringFromIniFile(iniFilePath, L"Scope", L"Delimiter", L"");
    setTextInDialogItem(_hSelf, IDC_DELIMITER_EDIT, delimiter);

}

void MultiReplace::loadSettings() {
    // Initialize configDir with all elements set to 0
    wchar_t configDir[MAX_PATH] = {};

    // Get the path to the plugin's configuration directory
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, (LPARAM)configDir);

    // Ensure configDir is null-terminated
    configDir[MAX_PATH - 1] = '\0';

    std::wstring iniFilePath = std::wstring(configDir) + L"\\MultiReplace.ini";
    std::wstring csvFilePath = std::wstring(configDir) + L"\\MultiReplaceList.ini";

    // Verify that the settings file exists before attempting to read it
    DWORD ftypIni = GetFileAttributesW(iniFilePath.c_str());
    if (ftypIni == INVALID_FILE_ATTRIBUTES) {
        // MessageBox(NULL, L"The settings file does not exist.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Verify that the list file exists before attempting to read it
    DWORD ftypCsv = GetFileAttributesW(csvFilePath.c_str());
    if (ftypCsv == INVALID_FILE_ATTRIBUTES) {
        // MessageBox(NULL, L"The list file does not exist.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }

    try {
        loadSettingsFromIni(iniFilePath);
        loadListFromCsvSilent(csvFilePath, replaceListData);
    }
    catch (const std::exception& ex) {
        std::wstring errorMessage = L"An error occurred while loading the settings: ";
        errorMessage += std::wstring(ex.what(), ex.what() + strlen(ex.what()));
        // MessageBox(NULL, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
    updateHeader();

    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, NULL, TRUE);
}

std::wstring MultiReplace::readStringFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, const std::wstring& defaultValue) {
    std::wifstream iniFile(iniFilePath);
    // Set the locale of the file stream for UTF-8 to UTF-16 conversion
    iniFile.imbue(std::locale(iniFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>));
    std::wstring line;
    bool correctSection = false;
    std::wstring requiredKey = key + L"=";
    size_t keyLength = requiredKey.length();

    if (iniFile.is_open()) {
        while (std::getline(iniFile, line)) {
            std::wstringstream linestream(line);
            std::wstring segment;

            if (line[0] == L'[') {
                // This line contains section name
                std::getline(linestream, segment, L']');
                correctSection = (segment == L"[" + section);
            }
            else if (correctSection && line.compare(0, keyLength, requiredKey) == 0) {
                // This line contains the key
                std::wstring value = line.substr(keyLength);
                // Remove prefix number and quotes from the value
                size_t quotePos = value.find(L"\"");
                if (quotePos != std::wstring::npos) {
                    value = value.substr(quotePos + 1, value.length() - quotePos - 2);
                }
                return value;
            }
        }
    }
    return defaultValue;
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