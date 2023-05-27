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

#include "MultiReplace.h"
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
#include <Notepad_plus_msgs.h>
#include <bitset>
#include <string>
#include <functional>



#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif

std::map<int, ControlInfo> MultiReplace::ctrlMap;

#pragma region Initialization

void MultiReplace::positionAndResizeControls(int windowWidth, int windowHeight)
{
    int buttonX = windowWidth - 40 - 160;
    int comboWidth = windowWidth - 360;
    int frameX = windowWidth - 30 - 310;
    int listWidth = windowWidth - 255;
    int listHeight = windowHeight - 300;
    int checkboxX = buttonX - 20 - 100;

    // Static positions and sizes
    ctrlMap[IDC_STATIC_FIND] = { 14, 14, 100, 24, WC_STATIC, L"Find what : ", SS_RIGHT };
    ctrlMap[IDC_STATIC_REPLACE] = { 14, 58, 100, 24, WC_STATIC, L"Replace with : ", SS_RIGHT };
    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { 20, 126, 200, 28, WC_BUTTON, L"Match whole word only", BS_AUTOCHECKBOX | WS_TABSTOP };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { 20, 156, 100, 28, WC_BUTTON, L"Match case", BS_AUTOCHECKBOX | WS_TABSTOP };
    ctrlMap[IDC_SEARCH_MODE_GROUP] = { 240, 106, 218, 110, WC_BUTTON, L"Search Mode", BS_GROUPBOX };
    ctrlMap[IDC_NORMAL_RADIO] = { 250, 126, 100, 20, WC_BUTTON, L"Normal", BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP };
    ctrlMap[IDC_REGEX_RADIO] = { 250, 156, 200, 20, WC_BUTTON, L"Regular expression", BS_AUTORADIOBUTTON | WS_TABSTOP };
    ctrlMap[IDC_EXTENDED_RADIO] = { 250, 186, 200, 20, WC_BUTTON, L"Extended (\\n, \\r, \\t, \\0, \\x...)", BS_AUTORADIOBUTTON | WS_TABSTOP };
    ctrlMap[IDC_STATIC_HINT] = { 14, 100, 500, 60, WC_STATIC, L"Please enlarge the window to view the controls.", SS_CENTER };
    ctrlMap[IDC_STATUS_MESSAGE] = { 14, 250, 450, 24, WC_STATIC, L"", WS_VISIBLE | SS_LEFT };

    // Dynamic positions and sizes
    ctrlMap[IDC_FIND_EDIT] = { 120, 14, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP };
    ctrlMap[IDC_REPLACE_EDIT] = { 120, 58, comboWidth, 200, WC_COMBOBOX, NULL, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP };
    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, 14, 160, 60, WC_BUTTON, L"Add to Replace List", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, 102, 160, 30, WC_BUTTON, L"Replace All", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, 142, 160, 30, WC_BUTTON, L"Mark Matches", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, 182, 160, 30, WC_BUTTON, L"Clear all marks", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX, 222, 160, 30, WC_BUTTON, L"Copy Marked Text", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, 286, 160, 30, WC_BUTTON, L"Load from CSV", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, 326, 160, 30, WC_BUTTON, L"Save to CSV", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_UP_BUTTON] = { buttonX+5, 390, 30, 30, WC_BUTTON, L"\u25B2", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER };
    ctrlMap[IDC_DOWN_BUTTON] = { buttonX+5, 390 + 30 + 5, 30, 30, WC_BUTTON, L"\u25BC", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER };    
    ctrlMap[IDC_SHIFT_FRAME] = { buttonX, 390-14, 160, 85, WC_BUTTON, L"", BS_GROUPBOX };
    ctrlMap[IDC_SHIFT_TEXT] = { buttonX + 40, 390 + 20, 60, 20, WC_STATIC, L"Shift Lines", SS_LEFT };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, 485, 160, 30, WC_BUTTON, L"Export to Bash", BS_PUSHBUTTON | WS_TABSTOP };
    ctrlMap[IDC_STATIC_FRAME] = { frameX, 88, 310, 170, WC_BUTTON, L"", BS_GROUPBOX };
    ctrlMap[IDC_REPLACE_LIST] = { 14, 270, listWidth, listHeight, WC_LISTVIEW, NULL, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS };
    ctrlMap[IDC_USE_LIST_CHECKBOX] = { checkboxX, 164, 80, 20, WC_BUTTON, L"Use List", BS_AUTOCHECKBOX | WS_TABSTOP };
}

bool MultiReplace::createAndShowWindows(HINSTANCE hInstance) {

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

        // Show the window
        ShowWindow(hwndControl, SW_SHOW);
        UpdateWindow(hwndControl);
    }
    return true;
}

void MultiReplace::initializeCtrlMap()
{
    HINSTANCE hInstance = (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE);

    // Get the client rectangle
    RECT rcClient;
    GetClientRect(_hSelf, &rcClient);
    // Extract width and height from the RECT
    int windowWidth = rcClient.right - rcClient.left;
    int windowHeight = rcClient.bottom - rcClient.top;

    // Define Position for all Elements
    positionAndResizeControls(windowWidth, windowHeight);

    // Now iterate over the controls and create each one.
    if (!createAndShowWindows(hInstance)) {
        return;
    }

    // Create the font
    hFont = CreateFont(FONT_SIZE, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, FONT_NAME);

    // Set the font for each control in ctrlMap
    for (auto& pair : ctrlMap)
    {
        SendMessage(GetDlgItem(_hSelf, pair.first), WM_SETFONT, (WPARAM)hFont, TRUE);
    }

    // CheckBox to Normal
    CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_EXTENDED_RADIO, IDC_NORMAL_RADIO);

    // Hide Hint Text
    ShowWindow(GetDlgItem(_hSelf, IDC_STATIC_HINT), SW_HIDE);

    // Enable the checkbox ID IDC_USE_LIST_CHECKBOX
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);

    // Deactivate Export Bash as it is experimental
    //HWND hExportButton = GetDlgItem(_hSelf, IDC_EXPORT_BASH_BUTTON);
    //EnableWindow(hExportButton, FALSE);

}

void MultiReplace::initializeScintilla() {
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which != -1) {
        _hScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
    }
}

void MultiReplace::createImageList() {
    const int ImageListSize = 16;
    _himl = ImageList_Create(ImageListSize, ImageListSize, ILC_COLOR32 | ILC_MASK, 1, 1);

    _hDeleteIcon = LoadIcon(_hInst, MAKEINTRESOURCE(DELETE_ICON));
    _hEnabledIcon = LoadIcon(_hInst, MAKEINTRESOURCE(ENABLED_ICON));
    _hCopyBackIcon = LoadIcon(_hInst, MAKEINTRESOURCE(COPYBACK_ICON));

    copyBackIconIndex = ImageList_AddIcon(_himl, _hCopyBackIcon);
    enabledIconIndex = ImageList_AddIcon(_himl, _hEnabledIcon);
    deleteIconIndex = ImageList_AddIcon(_himl, _hDeleteIcon);

    ListView_SetImageList(_replaceListView, _himl, LVSIL_SMALL);
}

void MultiReplace::initializeListView() {
    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    createImageList();
    createListViewColumns(_replaceListView);
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    ListView_SetExtendedListViewStyle(_replaceListView, LVS_EX_FULLROWSELECT | LVS_EX_SUBITEMIMAGES);
}

void MultiReplace::moveAndResizeControls() {
    // IDs of controls to be moved or resized
    const int controlIds[] = {
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_STATIC_FRAME, IDC_COPY_TO_LIST_BUTTON,
        IDC_REPLACE_ALL_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON,
        IDC_COPY_MARKED_TEXT_BUTTON, IDC_USE_LIST_CHECKBOX, IDC_LOAD_FROM_CSV_BUTTON,
        IDC_SAVE_TO_CSV_BUTTON, IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT, 
        IDC_EXPORT_BASH_BUTTON
    };

    // IDs of controls to be redrawn
    const int redrawIds[] = {
        IDC_WHOLE_WORD_CHECKBOX, IDC_MATCH_CASE_CHECKBOX, IDC_NORMAL_RADIO,
        IDC_REGEX_RADIO, IDC_EXTENDED_RADIO, IDC_USE_LIST_CHECKBOX,
        IDC_REPLACE_ALL_BUTTON, IDC_MARK_MATCHES_BUTTON, IDC_CLEAR_MARKS_BUTTON,
        IDC_COPY_MARKED_TEXT_BUTTON,  IDC_SHIFT_FRAME, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_TEXT,
        IDC_EXPORT_BASH_BUTTON
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
        IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_REPLACE_LIST, IDC_COPY_TO_LIST_BUTTON, IDC_USE_LIST_CHECKBOX,
        IDC_STATIC_FRAME, IDC_SEARCH_MODE_GROUP, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_EXTENDED_RADIO,
        IDC_STATIC_FIND, IDC_STATIC_REPLACE, IDC_MATCH_CASE_CHECKBOX, IDC_WHOLE_WORD_CHECKBOX,
        IDC_LOAD_FROM_CSV_BUTTON, IDC_SAVE_TO_CSV_BUTTON, IDC_REPLACE_ALL_BUTTON, IDC_MARK_MATCHES_BUTTON,
        IDC_CLEAR_MARKS_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON, IDC_UP_BUTTON, IDC_DOWN_BUTTON, IDC_SHIFT_FRAME,
        IDC_SHIFT_TEXT, IDC_STATUS_MESSAGE, IDC_EXPORT_BASH_BUTTON
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
    int remainingWidth = windowWidth - 280 - 190;
    remainingWidth = remainingWidth;
    // Column for "Find" Text
    lvc.iSubItem = 0;
    lvc.pszText = L"Find";
    lvc.cx = 195;
    //lvc.cx = remainingWidth / 2;
    ListView_InsertColumn(listView, 0, &lvc);

    // Column for "Replace" Text
    lvc.iSubItem = 1;
    lvc.pszText = L"Replace";
    lvc.cx = 195;
    //lvc.cx = remainingWidth / 2;
    ListView_InsertColumn(listView, 1, &lvc);

    // Column for Option: Whole Word
    lvc.iSubItem = 2;
    lvc.pszText = L"W";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 2, &lvc);

    // Column for Option: Match Case
    lvc.iSubItem = 3;
    lvc.pszText = L"C";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 3, &lvc);

    // Column for Option: Normal
    lvc.iSubItem = 4;
    lvc.pszText = L"N";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 4, &lvc);

    // Column for Option: Regex
    lvc.iSubItem = 5;
    lvc.pszText = L"R";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 5, &lvc);

    // Column for Option: Extended
    lvc.iSubItem = 6;
    lvc.pszText = L"E";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 6, &lvc);

    // Column for Copy Back Button
    lvc.iSubItem = 7;
    lvc.pszText = L"";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 7, &lvc);

    // Column for Delete Button
    lvc.iSubItem = 8;
    lvc.pszText = L"";
    lvc.cx = 30;
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 8, &lvc);
}

void MultiReplace::insertReplaceListItem(const ReplaceItemData& itemData)
{
    // Return early if findText is empty
    if (itemData.findText.empty()) {
        showStatusMessage(0, L"No 'Find String' entered. Please provide a value to add to the list.", RGB(255, 0, 0));
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
    showStatusMessage(0, message.c_str(), RGB(0, 128, 0));

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
}

void MultiReplace::updateListViewAndColumns(HWND listView, LPARAM lParam)
{
    // Get the new width and height of the window from lParam
    int newWidth = LOWORD(lParam);
    int newHeight = HIWORD(lParam);

    // Calculate the total width of columns 3 to 8
    int columns3to7Width = 0;
    for (int i = 2; i < 9; i++)
    {
        columns3to7Width += ListView_GetColumnWidth(listView, i);
    }

    // Calculate the remaining width for the first two columns
    int remainingWidth = newWidth - 280 - columns3to7Width;

    static int prevWidth = newWidth; // Store the previous width

    // If the window is horizontally maximized, update the IDC_REPLACE_LIST size first
    if (newWidth > prevWidth) {
        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_LIST), 14, 270, newWidth - 255, newHeight - 300, TRUE);
    }

    ListView_SetColumnWidth(listView, 0, remainingWidth / 2);
    ListView_SetColumnWidth(listView, 1, remainingWidth / 2);

    // If the window is horizontally minimized or vetically changed the size
    MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_LIST), 14, 270, newWidth - 255, newHeight - 300, TRUE);

    // If the window size hasn't changed, no need to do anything

    prevWidth = newWidth;
}

void MultiReplace::handleDeletion(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || pnmia->iItem >= replaceListData.size()) {
        return;
    }
    // Remove the item from the ListView
    ListView_DeleteItem(_replaceListView, pnmia->iItem);

    // Remove the item from the replaceListData vector
    replaceListData.erase(replaceListData.begin() + pnmia->iItem);

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(0, L"1 line deleted.", RGB(0, 128, 0));
}

void MultiReplace::handleCopyBack(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || pnmia->iItem >= replaceListData.size()) {
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
    SendMessageW(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_SETCHECK, itemData.regex ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_EXTENDED_RADIO), BM_SETCHECK, itemData.extended ? BST_CHECKED : BST_UNCHECKED, 0);
}

void MultiReplace::shiftListItem(HWND listView, const Direction& direction) {
    
    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
    
    std::vector<int> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(listView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(i);
    }

    if (selectedIndices.empty()) {
        showStatusMessage(0, L"No rows selected to shift.", RGB(255, 0, 0));
        return;
    }

    // Check the bounds
    if ((direction == Direction::Up && selectedIndices.front() == 0) || (direction == Direction::Down && selectedIndices.back() == replaceListData.size() - 1)) {
        return; // Don't perform the move if it's out of bounds
    }

    // Perform the shift operation
    if (direction == Direction::Up) {
        for (int& index : selectedIndices) {
            int swapIndex = index - 1;
            std::swap(replaceListData[index], replaceListData[swapIndex]);
            index = swapIndex;
        }
    }
    else { // direction is Down
        for (auto it = selectedIndices.rbegin(); it != selectedIndices.rend(); ++it) {
            int swapIndex = *it + 1;
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
    for (int index : selectedIndices) {
        ListView_SetItemState(listView, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Show status message when rows are successfully shifted
    std::wstring statusMessage = std::to_wstring(selectedIndices.size()) + L" rows successfully shifted.";
    showStatusMessage(static_cast<int>(selectedIndices.size()), statusMessage.c_str(), RGB(0, 128, 0));
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

    InvalidateRect(_replaceListView, NULL, TRUE);

    showStatusMessage(static_cast<int>(numDeletedLines), L"%d lines deleted.", RGB(0, 128, 0));
}

#pragma endregion


#pragma region SearchReplace

INT_PTR CALLBACK MultiReplace::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{

    switch (message)
    {
    case WM_INITDIALOG:
    {
        initializeScintilla();
        initializeCtrlMap();
        updateUIVisibility();
        initializeListView();

        return TRUE;
    }
    break;

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_STATUS_MESSAGE)) {
            SetTextColor(hdcStatic, _statusMessageColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return (LRESULT)GetStockObject(NULL_BRUSH);
        }

        break;
    }
    break;

    case WM_DESTROY:
    {
        DestroyIcon(_hDeleteIcon);
        DestroyIcon(_hEnabledIcon);
        DestroyIcon(_hCopyBackIcon);
        ImageList_Destroy(_himl);
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
                if (pnmia->iSubItem == 8) { // Delete button column
                    handleDeletion(pnmia);
                }
                else if (pnmia->iSubItem == 7) { // Copy Back button column
                    handleCopyBack(pnmia);
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
                case 0:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.findText.c_str());
                    break;
                case 1:
                    plvdi->item.pszText = const_cast<LPWSTR>(itemData.replaceText.c_str());
                    break;
                case 2:
                    if (itemData.wholeWord) {
                        //plvdi->item.mask |= LVIF_IMAGE;
                        //plvdi->item.iImage = enabledIconIndex;
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 3:
                    if (itemData.matchCase) {
                        //plvdi->item.mask |= LVIF_IMAGE;
                        //plvdi->item.iImage = enabledIconIndex;
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 4:
                    if (!itemData.regex && !itemData.extended) {
                        //plvdi->item.mask |= LVIF_IMAGE;
                        //plvdi->item.iImage = enabledIconIndex;
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 5:
                    if (itemData.regex) {
                        //plvdi->item.mask |= LVIF_IMAGE;
                        //plvdi->item.iImage = enabledIconIndex;
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 6:
                    if (itemData.extended) {
                        //plvdi->item.mask |= LVIF_IMAGE;
                        //plvdi->item.iImage = enabledIconIndex;
                        plvdi->item.mask |= LVIF_TEXT;
                        plvdi->item.pszText = L"\u2714";
                    }
                    break;
                case 7:
                    //plvdi->item.mask |= LVIF_IMAGE;
                    //plvdi->item.iImage = copyBackIconIndex;
                    plvdi->item.mask |= LVIF_TEXT;
                    plvdi->item.pszText = L"\U0001F879";
                    break;
                case 8:
                    //plvdi->item.mask |= LVIF_IMAGE;
                    //plvdi->item.iImage = deleteIconIndex;
                    plvdi->item.mask |= LVIF_TEXT;
                    plvdi->item.pszText = L"\u2716";
                    break;

                }
            }
            break;

            case LVN_KEYDOWN:
            {
                LPNMLVKEYDOWN pnkd = reinterpret_cast<LPNMLVKEYDOWN>(pnmh);

                if (pnkd->wVKey == VK_DELETE) {
                    deleteSelectedLines(_replaceListView);
                }
            }
            break;

            }
        }
    }
    break;

    case WM_TIMER:
    {   //Refresh of DropDown due to a Bug in Notepad++ Plugin implementation of Dark Mode
        if (wParam == 1)
        {
            KillTimer(_hSelf, 1);

            BOOL isDarkModeEnabled = (BOOL)::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0);

            if (!isDarkModeEnabled)
            {
                int comboBoxIDs[] = { IDC_FIND_EDIT, IDC_REPLACE_EDIT };

                for (int id : comboBoxIDs)
                {
                    HWND hComboBox = GetDlgItem(_hSelf, id);
                    int itemCount = (int)SendMessage(hComboBox, CB_GETCOUNT, 0, 0);

                    for (int i = itemCount - 1; i >= 0; i--)
                    {
                        SendMessage(hComboBox, CB_SETCURSEL, (WPARAM)i, 0);
                    }
                }
            }
        }
    }
    break;

    case WM_COMMAND:
    {
        switch (HIWORD(wParam))
        {
        case CBN_DROPDOWN:
        {   //Refresh of DropDown due to a Bug in Notepad++ Plugin implementation of Dark Mode
            if (LOWORD(wParam) == IDC_FIND_EDIT || LOWORD(wParam) == IDC_REPLACE_EDIT)
            {
                SetTimer(_hSelf, 1, 1, NULL);
            }
        }
        break;
        }

        switch (LOWORD(wParam))
        {
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
            // Enable the Whole word checkbox
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), TRUE);
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

        case IDC_COPY_TO_LIST_BUTTON:
        {            
            onCopyToListButtonClick();

        }
        break;

        case IDC_REPLACE_ALL_BUTTON:
        {
            int replaceCount = 0;
            // Check if the "In List" option is enabled
            bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);

            if (useListEnabled)
            {
                // Check if the replaceListData is empty and warn the user if so
                if (replaceListData.empty()) {
                    showStatusMessage(0, L"Add values into the list. Or uncheck 'Use in List' to replace directly.", RGB(255, 0, 0));
                    break;
                }
                ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
                for (size_t i = 0; i < replaceListData.size(); i++)
                {
                    ReplaceItemData& itemData = replaceListData[i];
                    replaceCount += replaceString(
                        itemData.findText.c_str(), itemData.replaceText.c_str(),
                        itemData.wholeWord, itemData.matchCase,
                        itemData.regex, itemData.extended
                    );
                }
                ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);
            }
            else
            {
                TCHAR findText[256];
                TCHAR replaceText[256];
                GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
                GetDlgItemText(_hSelf, IDC_REPLACE_EDIT, replaceText, 256);
                bool regex = false;
                bool extended = false;

                // Get the state of the Whole word checkbox
                bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

                // Get the state of the Match case checkbox
                bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

                if (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED) {
                    regex = true;
                }
                else if (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED) {
                    extended = true;
                }

                ::SendMessage(_hScintilla, SCI_BEGINUNDOACTION, 0, 0);
                replaceCount = replaceString(findText, replaceText, wholeWord, matchCase, regex, extended);
                ::SendMessage(_hScintilla, SCI_ENDUNDOACTION, 0, 0);

                // Add the entered text to the combo box history
                addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
                addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);
            }
            // Display status message
            showStatusMessage(replaceCount, L"%d occurrences were replaced.", RGB(0, 128, 0));
        }
        break;

        case IDC_MARK_MATCHES_BUTTON:
        {
            int matchCount = 0;
            // Check if the "In List" option is enabled
            bool useListEnabled = (IsDlgButtonChecked(_hSelf, IDC_USE_LIST_CHECKBOX) == BST_CHECKED);

            if (useListEnabled)
            {
                // Check if the replaceListData is empty and warn the user if so
                if (replaceListData.empty()) {
                    showStatusMessage(0, L"Add values into the list. Or uncheck 'Use in List' to mark directly.", RGB(255, 0, 0));
                    break;
                }
                // Iterate through the list of items
                for (size_t i = 0; i < replaceListData.size(); i++)
                {
                    ReplaceItemData& itemData = replaceListData[i];
                    bool regex = itemData.regex;
                    bool extended = itemData.extended;

                    matchCount += markString(
                        itemData.findText.c_str(), itemData.wholeWord,
                        itemData.matchCase, regex, extended);
                }
            }
            else
            {
                TCHAR findText[256];
                GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
                bool regex = false;
                bool extended = false;

                // Get the state of the Whole word checkbox
                bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

                // Get the state of the Match case checkbox
                bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

                if (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED) {
                    regex = true;
                }
                else if (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED) {
                    extended = true;
                }

                // Perform the Mark Matching Strings operation if Find field has a value
                matchCount = markString(findText, wholeWord, matchCase, regex, extended);

                // Add the entered text to the combo box history
                addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
            }
            // Display status message
            showStatusMessage(matchCount, L"%d occurrences were marked.", RGB(0, 0, 128));  // Blue color
        }
        break;

        case IDC_CLEAR_MARKS_BUTTON:
        {
            clearAllMarks();
        }
        break;

        case IDC_COPY_MARKED_TEXT_BUTTON:
        {
            copyMarkedTextToClipboard();
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

        default:
            return FALSE;
        }

    }
    break;

    }
    return FALSE;
}

int MultiReplace::convertExtendedToString(const TCHAR* query, TCHAR* result, int length)
{
    auto readBase = [](const TCHAR* str, int* value, int base, int size) -> bool
    {
        int i = 0, temp = 0;
        *value = 0;
        TCHAR max = '0' + static_cast<TCHAR>(base) - 1;
        TCHAR current;
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
    int resultLength = 0;

    for (int i = 0; i < length; ++i)
    {
        if (query[i] == '\\' && (i + 1) < length)
        {
            ++i;
            TCHAR current = query[i];
            switch (current)
            {
            case 'n':
                result[resultLength++] = '\n';
                break;
            case 't':
                result[resultLength++] = '\t';
                break;
            case 'r':
                result[resultLength++] = '\r';
                break;
            case '0':
                result[resultLength++] = '\0';
                break;
            case '\\':
                result[resultLength++] = '\\';
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

                if (length - i >= size)
                {
                    int res = 0;
                    if (readBase(query + (i + 1), &res, base, size))
                    {
                        result[resultLength++] = static_cast<TCHAR>(res);
                        i += size;
                        break;
                    }
                }
                // not enough chars to make parameter, use default method as fallback
                /* fallthrough */
            }

            default:
                // unknown sequence, treat as regular text
                result[resultLength++] = '\\';
                result[resultLength++] = current;
                break;
            }
        }
        else
        {
            result[resultLength++] = query[i];
        }
    }

    result[resultLength] = 0;

    // Convert TCHAR string to UTF-8 string
    std::wstring_convert<std::codecvt_utf8_utf16<TCHAR>> converter;
    std::string utf8Result = converter.to_bytes(result);

    // Return the length of the UTF-8 string
    return static_cast<int>(utf8Result.length());

}

int MultiReplace::replaceString(const TCHAR* findText, const TCHAR* replaceText, bool wholeWord, bool matchCase, bool regex, bool extended)
{
    // Return early if the Find field is empty
    if (findText[0] == '\0') {
        return 0;
    }

    int searchFlags = 0;
    if (wholeWord)
        searchFlags |= SCFIND_WHOLEWORD;
    if (matchCase)
        searchFlags |= SCFIND_MATCHCASE;
    if (regex)
        searchFlags |= SCFIND_REGEXP;

    std::wstring_convert<std::codecvt_utf8_utf16<TCHAR>> converter;
    std::string findTextUtf8 = converter.to_bytes(findText);
    std::string replaceTextUtf8 = converter.to_bytes(replaceText);

    int findTextLength = static_cast<int>(findTextUtf8.length());
    int replaceTextLength = static_cast<int>(replaceTextUtf8.length());
    TCHAR findTextExtended[256];
    TCHAR replaceTextExtended[256];
    if (extended)
    {
        int findTextExtendedLength = convertExtendedToString(findText, findTextExtended, findTextLength);
        int replaceTextExtendedLength = convertExtendedToString(replaceText, replaceTextExtended, replaceTextLength);

        findTextLength = findTextExtendedLength;
        replaceTextLength = replaceTextExtendedLength;

        findTextUtf8 = converter.to_bytes(findTextExtended);
        replaceTextUtf8 = converter.to_bytes(replaceTextExtended);
    }

    Sci_Position pos = 0;
    Sci_Position matchLen = 0;

    int replaceCount = 0;  // Counter for replacements

    while (pos >= 0)
    {
        ::SendMessage(_hScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(_hScintilla, SCI_SETTARGETEND, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(_hScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);
        pos = ::SendMessage(_hScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));

        if (pos >= 0)
        {
            matchLen = ::SendMessage(_hScintilla, SCI_GETTARGETEND, 0, 0) - pos;
            ::SendMessage(_hScintilla, SCI_SETSEL, pos, pos + matchLen);
            ::SendMessage(_hScintilla, SCI_REPLACESEL, 0, (LPARAM)replaceTextUtf8.c_str());
            pos += replaceTextLength;
            replaceCount++;  // Increment the counter for each replacement
        }
    }

    return replaceCount;  // Return the count of replacements
}

int MultiReplace::markString(const TCHAR* findText, bool wholeWord, bool matchCase, bool regex, bool extended)
{
    // Return early if the Find field is empty
    if (findText[0] == '\0') {
        return 0;
    }

    int searchFlags = 0;
    if (wholeWord)
        searchFlags |= SCFIND_WHOLEWORD;
    if (matchCase)
        searchFlags |= SCFIND_MATCHCASE;
    if (regex)
        searchFlags |= SCFIND_REGEXP;

    std::wstring_convert<std::codecvt_utf8_utf16<TCHAR>> converter;
    std::string findTextUtf8 = converter.to_bytes(findText);

    int findTextLength = static_cast<int>(findTextUtf8.length());

    if (extended)
    {
        TCHAR findTextExtended[256];
        int findTextExtendedLength = convertExtendedToString(findText, findTextExtended, findTextLength);
        findTextLength = findTextExtendedLength;
        findTextUtf8 = converter.to_bytes(findTextExtended);
    }

    LRESULT pos = 0;
    LRESULT matchLen = 0;
    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(_hScintilla, SCI_INDICSETSTYLE, 0, INDIC_STRAIGHTBOX);
    ::SendMessage(_hScintilla, SCI_INDICSETFORE, 0, 0x007F00);
    ::SendMessage(_hScintilla, SCI_INDICSETALPHA, 0, 100);

    int markCount = 0;  // Counter for marked matches

    while (pos >= 0)
    {
        ::SendMessage(_hScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(_hScintilla, SCI_SETTARGETEND, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(_hScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

        pos = ::SendMessage(_hScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));
        if (pos >= 0)
        {
            matchLen = ::SendMessage(_hScintilla, SCI_GETTARGETEND, 0, 0) - pos;
            ::SendMessage(_hScintilla, SCI_SETINDICATORVALUE, 1, 0);
            ::SendMessage(_hScintilla, SCI_INDICATORFILLRANGE, pos, matchLen);
            pos += matchLen;

            markCount++;  // Increment the counter for each marked match
        }
    }

    return markCount;  // Return the count of marked matches
}

void MultiReplace::clearAllMarks()
{
    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(_hScintilla, SCI_INDICATORCLEARRANGE, 0, ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
    showStatusMessage(0, L"All marks cleared.", RGB(0, 128, 0));
}

void MultiReplace::copyMarkedTextToClipboard()
{

    LRESULT length = ::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0);
    std::string markedText;
    int markedTextCount = 0;

    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    for (int i = 0; i < length; ++i)
    {
        if (::SendMessage(_hScintilla, SCI_INDICATORVALUEAT, 0, i))
        {
            char ch = static_cast<char>(::SendMessage(_hScintilla, SCI_GETCHARAT, i, 0));
            markedText += ch;
            markedTextCount++;
        }
    }

    if (!markedText.empty())
    {
        const char* output = markedText.c_str();
        size_t outputLength = markedText.length();
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, outputLength + 1);
        if (hMem)
        {
            LPVOID lockedMem = GlobalLock(hMem);
            if (lockedMem)
            {
                memcpy(lockedMem, output, outputLength + 1);
                GlobalUnlock(hMem);
                OpenClipboard(0);
                EmptyClipboard();
                SetClipboardData(CF_TEXT, hMem);
                CloseClipboard();
                showStatusMessage(markedTextCount, L"%d marked text strings copied into Clipboard.", RGB(0, 128, 0));
            }
        }
    }
    else
    {
        showStatusMessage(0, L"No marked text to copy.", RGB(255, 0, 0));
    }
}

void MultiReplace::onCopyToListButtonClick() {
    ReplaceItemData itemData;

    TCHAR findText[256];
    TCHAR replaceText[256];
    GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
    GetDlgItemText(_hSelf, IDC_REPLACE_EDIT, replaceText, 256);
    itemData.findText = findText;
    itemData.replaceText = replaceText;

    itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
    itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
    itemData.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
    itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

    insertReplaceListItem(itemData);

    // Add the entered text to the combo box history
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
}

void MultiReplace::addStringToComboBoxHistory(HWND hComboBox, const TCHAR* str, int maxItems)
{
    if (_tcslen(str) == 0)
    {
        // Skip adding empty strings to the combo box history
        return;
    }

    // Check if the string is already in the combo box
    int index = static_cast<int>(SendMessage(hComboBox, CB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(str)));

    // If the string is not found, insert it at the beginning
    if (index == CB_ERR)
    {
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str));

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
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str));
    }

    // Select the newly added string
    SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
}

void MultiReplace::showStatusMessage(int count, const wchar_t* messageFormat, COLORREF color)
{
    wchar_t message[350];
    wsprintf(message, messageFormat, count);

    // Get the handle for the status message control
    HWND hStatusMessage = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);

    // Set the new message
    _statusMessageColor = color;
    SetWindowText(hStatusMessage, message);

    // Invalidate the area of the parent where the control resides
    RECT rect;
    GetWindowRect(hStatusMessage, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hStatusMessage), reinterpret_cast<POINT*>(&rect), 2);
    InvalidateRect(GetParent(hStatusMessage), &rect, TRUE);
    UpdateWindow(GetParent(hStatusMessage));
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

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (list.empty())
    {
        showStatusMessage(0, L"The list is empty. Nothing to save.", RGB(255, 0, 0));
        return;
    }

    std::wofstream outFile(filePath);
    outFile.imbue(std::locale(outFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>));

    if (!outFile.is_open()) {
        showStatusMessage(0, L"Error: Unable to open file for writing.", RGB(255, 0, 0)); // Red color
        return;
    }

    // Write CSV header
    outFile << L"Find,Replace,WholeWord,regex,MatchCase,Extended" << std::endl;

    // Write list items to CSV file
    for (const ReplaceItemData& item : list) {
        outFile << escapeCsvValue(item.findText) << L"," << escapeCsvValue(item.replaceText) << L"," << item.wholeWord << L"," << item.regex << L"," << item.matchCase << L"," << item.extended << std::endl;
    }

    outFile.close();

    showStatusMessage(static_cast<int>(list.size()), L"%d items saved to CSV.", RGB(0, 128, 0));

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {
    std::wifstream inFile(filePath);
    inFile.imbue(std::locale(inFile.getloc(), new std::codecvt_utf8_utf16<wchar_t>()));

    if (!inFile.is_open()) {
        showStatusMessage(0, L"Error: Unable to open file for reading.", RGB(255, 0, 0)); // Red color
        return;
    }

    replaceListData.clear(); // Clear the existing list

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
        if (columns.size() != 6) {
            continue;
        }

        ReplaceItemData item;

        // Assign columns to item properties
        item.findText = columns[0];
        item.replaceText = columns[1];
        item.wholeWord = std::stoi(columns[2]) != 0;
        item.regex = std::stoi(columns[3]) != 0;
        item.matchCase = std::stoi(columns[4]) != 0;
        item.extended = std::stoi(columns[5]) != 0;

        // Use insertReplaceListItem to insert the item to the list
        insertReplaceListItem(item);

    }

    inFile.close();

    // Update the list view control
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    if (replaceListData.empty())
    {
        showStatusMessage(0, L"No valid items found in the CSV file.", RGB(255, 0, 0)); // Red color
    }
    else
    {
        // Display status message
        showStatusMessage(static_cast<int>(replaceListData.size()), L"%d items loaded from CSV.", RGB(0, 128, 0)); // Green color
    }

    // Enable the ListView accordingly
    SendMessage(GetDlgItem(_hSelf, IDC_USE_LIST_CHECKBOX), BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(_replaceListView, TRUE);
}

std::wstring MultiReplace::escapeCsvValue(const std::wstring& value) {
    std::wstring escapedValue;
    bool needsQuotes = false;

    for (const wchar_t& ch : value) {
        // Check if value contains any character that requires escaping
        if (ch == L',' || ch == L'"' || ch == L'\n' || ch == L'\r') {
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

    file << "#!/bin/bash\n";
    file << "# Auto-generated by MultiReplace Notepad++ plugin (by Thomas Knoefel)\n\n";
    file << "textFile=\"test.txt\"\n\n";

    file << "processLine() {\n";
    file << "    local findString=\"$1\"\n";
    file << "    local replaceString=\"$2\"\n";
    file << "    local wholeWord=\"$3\"\n";
    file << "    local matchCase=\"$4\"\n";
    file << "    local normal=\"$5\"\n";
    file << "    local regex=\"$6\"\n";
    file << "    local extended=\"$7\"\n";
    file << R"(
if [[ "$wholeWord" -eq 1 ]]; then
    findString='\b'${findString}'\b'
fi
if [[ "$matchCase" -eq 1 ]]; then
    template='s/'${findString}'/'${replaceString}'/g'
else
    template='s/'${findString}'/'${replaceString}'/gi'
fi
case 1 in
    $normal)
        sed -i -e "${template}" "$textFile"
        ;;
    $regex)
        sed -i -r -e "${template}" "$textFile"
        ;;
    $extended)
        sed -i -e ':a' -e 'N' -e '$!ba' -e"${template}" "$textFile"
        ;;
esac
)";
    file << "}\n\n";

    for (const auto& itemData : replaceListData) {
        std::string find;
        std::string replace;
        if (itemData.extended) {
            find = translateUnsupportedEscapes(escapeSpecialChars(wstringToString(itemData.findText), true));
            replace = translateUnsupportedEscapes(escapeSpecialChars(wstringToString(itemData.replaceText), true));
        }
        else {
            find = escapeSpecialChars(wstringToString(itemData.findText), false);
            replace = escapeSpecialChars(wstringToString(itemData.replaceText), false);
        }

        std::string wholeWord = itemData.wholeWord ? "1" : "0";
        std::string matchCase = itemData.matchCase ? "1" : "0";
        std::string regex = itemData.regex ? "1" : "0";
        std::string normal = (!itemData.regex && !itemData.extended) ? "1" : "0";
        std::string extended = itemData.extended ? "1" : "0";

        file << "# processLine arguments: \"findString\" \"replaceString\" wholeWord matchCase normal regex extended\n";
        file << "processLine \"" << find << "\" \"" << replace << "\" " << wholeWord << " " << matchCase << " " << normal << " " << regex << " " << extended << "\n";
    }

    file.close();
}

std::string MultiReplace::wstringToString(const std::wstring& wstr)
{
    std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> converter;
    return converter.to_bytes(wstr);
}

std::string MultiReplace::escapeSpecialChars(const std::string& input, bool extended) {
    std::string output = input;

    // Define the escape characters that should not be masked
    std::string supportedEscapes = "nrt0xub";

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

std::string MultiReplace::translateUnsupportedEscapes(const std::string& input) {
    std::string output = input;

    std::regex octalRegex("\\\\o([0-7]{3})");
    std::regex decimalRegex("\\\\d([0-9]{3})");
    std::regex hexRegex("\\\\x([0-9a-fA-F]{2})");
    std::regex binaryRegex("\\\\b([01]{8})");
    std::regex unicodeRegex("\\\\u([0-9a-fA-F]{4})");

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

    return output;
}

#pragma endregion
