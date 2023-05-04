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

#include "MultiReplacePanel.h"
#include "PluginDefinition.h"
#include <codecvt>
#include <locale>
#include <regex>
#include <windows.h>
#include <sstream>
#include <Commctrl.h>
#include <vector>



extern NppData nppData;

#define RESIZE_TIMER_ID 1


#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif


int MultiReplacePanel::convertExtendedToString(const TCHAR* query, TCHAR* result, int length)
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


void MultiReplacePanel::findAndReplace(const TCHAR* findText, const TCHAR* replaceText, bool wholeWord, bool matchCase, bool regexSearch, bool extended)
{
    int searchFlags = 0;
    if (wholeWord)
        searchFlags |= SCFIND_WHOLEWORD;
    if (matchCase)
        searchFlags |= SCFIND_MATCHCASE;
    if (regexSearch)
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

    while (pos >= 0)
    {
        ::SendMessage(_curScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(_curScintilla, SCI_SETTARGETEND, ::SendMessage(_curScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(_curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);
        pos = ::SendMessage(_curScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));

        if (pos >= 0)
        {
            matchLen = ::SendMessage(_curScintilla, SCI_GETTARGETEND, 0, 0) - pos;
            ::SendMessage(_curScintilla, SCI_SETSEL, pos, pos + matchLen);
            ::SendMessage(_curScintilla, SCI_REPLACESEL, 0, (LPARAM)replaceTextUtf8.c_str());
            //pos = ::SendMessage(curScintilla, SCI_POSITIONAFTER, pos + matchLen, 0);
            pos += replaceTextLength;
        }
    }
}


void MultiReplacePanel::markMatchingStrings(const TCHAR* findText, bool wholeWord, bool matchCase, bool regexSearch, bool extended)
{

    int searchFlags = 0;
    if (wholeWord)
        searchFlags |= SCFIND_WHOLEWORD;
    if (matchCase)
        searchFlags |= SCFIND_MATCHCASE;
    if (regexSearch)
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
    ::SendMessage(_curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(_curScintilla, SCI_INDICSETSTYLE, 0, INDIC_STRAIGHTBOX);
    ::SendMessage(_curScintilla, SCI_INDICSETFORE, 0, 0x007F00);
    ::SendMessage(_curScintilla, SCI_INDICSETALPHA, 0, 100);

    while (pos >= 0)
    {
        ::SendMessage(_curScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(_curScintilla, SCI_SETTARGETEND, ::SendMessage(_curScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(_curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

        pos = ::SendMessage(_curScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));
        if (pos >= 0)
        {
            matchLen = ::SendMessage(_curScintilla, SCI_GETTARGETEND, 0, 0) - pos;
            ::SendMessage(_curScintilla, SCI_SETINDICATORVALUE, 1, 0);
            ::SendMessage(_curScintilla, SCI_INDICATORFILLRANGE, pos, matchLen);
            pos += findTextLength;
        }
    }
}


void MultiReplacePanel::clearAllMarks()
{
    ::SendMessage(_curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(_curScintilla, SCI_INDICATORCLEARRANGE, 0, ::SendMessage(_curScintilla, SCI_GETLENGTH, 0, 0));
}

void MultiReplacePanel::copyMarkedTextToClipboard()
{

    LRESULT length = ::SendMessage(_curScintilla, SCI_GETLENGTH, 0, 0);
    std::string markedText;

    ::SendMessage(_curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    for (int i = 0; i < length; ++i)
    {
        if (::SendMessage(_curScintilla, SCI_INDICATORVALUEAT, 0, i))
        {
            char ch = static_cast<char>(::SendMessage(_curScintilla, SCI_GETCHARAT, i, 0));
            markedText += ch;
        }
    }

    if (!markedText.empty())
    {
        const char* output = markedText.c_str();
        size_t outputLength = markedText.length();
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, outputLength + 1);
        if (hMem)
        {
            memcpy(GlobalLock(hMem), output, outputLength + 1);
            GlobalUnlock(hMem);
            OpenClipboard(0);
            EmptyClipboard();
            SetClipboardData(CF_TEXT, hMem);
            CloseClipboard();
        }
    }
}


// handle the Copy button click
void MultiReplacePanel::onCopyToListButtonClick() {
    ReplaceItemData itemData;

    TCHAR findText[256];
    TCHAR replaceText[256];
    GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
    GetDlgItemText(_hSelf, IDC_REPLACE_EDIT, replaceText, 256);
    itemData.findText = findText;
    itemData.replaceText = replaceText;

    itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
    itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
    itemData.regexSearch = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
    itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

    insertReplaceListItem(itemData);
}


// handle the Replace all in List button click
void MultiReplacePanel::onReplaceAllInListButtonClick() {
    // Add code to loop through the ListView items and perform Find and Replace operations using the stored options for each item

}

void MultiReplacePanel::insertReplaceListItem(const ReplaceItemData& itemData)
{
    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    // Add the data to the vector
    ReplaceItemData newItemData = itemData;


    deleteIconIndex = ImageList_AddIcon(_himl, _hDeleteIcon);
    enabledIconIndex = ImageList_AddIcon(_himl, _hEnabledIcon);
    replaceListData.push_back(newItemData);

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    InvalidateRect(_replaceListView, NULL, TRUE);
}


void MultiReplacePanel::createListViewColumns(HWND listView) {
    LVCOLUMN lvc;

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvc.fmt = LVCFMT_LEFT;

    // Column for "Find" Text
    lvc.iSubItem = 0;
    lvc.pszText = L"Find";
    lvc.cx = 347; // Spaltenbreite
    ListView_InsertColumn(listView, 0, &lvc);

    // Column for "Replace" Text
    lvc.iSubItem = 1;
    lvc.pszText = L"Replace";
    lvc.cx = 346; // Spaltenbreite
    ListView_InsertColumn(listView, 1, &lvc);

    // Column for Option: Whole Word
    lvc.iSubItem = 2;
    lvc.pszText = L"W";
    lvc.cx = 30; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 2, &lvc);

    // Column for Option: Match Case
    lvc.iSubItem = 3;
    lvc.pszText = L"C";
    lvc.cx = 30; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 3, &lvc);

    // Column for Option: Normal
    lvc.iSubItem = 4;
    lvc.pszText = L"N";
    lvc.cx = 30; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 4, &lvc);

    // Column for Option: Regex
    lvc.iSubItem = 5;
    lvc.pszText = L"R";
    lvc.cx = 30; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 5, &lvc);

    // Column for Option: Extended
    lvc.iSubItem = 6;
    lvc.pszText = L"E";
    lvc.cx = 30; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 6, &lvc);

    // Column for Copy Back Button
    lvc.iSubItem = 7;
    lvc.pszText = L"";
    lvc.cx = 20; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 7, &lvc); 

    // Column for Delete Button
    lvc.iSubItem = 8;
    lvc.pszText = L"";
    lvc.cx = 20; // Spaltenbreite
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(listView, 8, &lvc);
}


INT_PTR CALLBACK MultiReplacePanel::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    static HFONT hFont = NULL;

    switch (message)
    {
    case WM_INITDIALOG:
    {

        // Create the font
        hFont = CreateFont(20, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 0, 0, 0, 0, TEXT("MS Shell Dlg"));

        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_EXTENDED_RADIO, IDC_NORMAL_RADIO);

        // Set the font for the controls
        SendMessage(GetDlgItem(_hSelf, IDC_FIND_EDIT), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), WM_SETFONT, (WPARAM)hFont, TRUE);

        // Initialize curScintilla
        int which = -1;
        ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
        if (which != -1) {
            _curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
        }

        // Check if the ListView is created correctly
        _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

        // Creating ImageList
        _himl = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);

        // Load Delete Button
        _hDeleteIcon = LoadIcon(_hInst, MAKEINTRESOURCE(DELETE_ICON));

        // Load Enabled Icon
        _hEnabledIcon = LoadIcon(_hInst, MAKEINTRESOURCE(ENABLED_ICON));

        // Load Copy Back Icon
        _hCopyBackIcon = LoadIcon(_hInst, MAKEINTRESOURCE(COPYBACK_ICON));

        // Add the Copy Back Icon to the ImageList
        copyBackIconIndex = ImageList_AddIcon(_himl, _hCopyBackIcon);

        // Add the Set Icon to the ImageList
        enabledIconIndex = ImageList_AddIcon(_himl, _hEnabledIcon);

        // Assign the IconList object to the ListView control
        ListView_SetImageList(_replaceListView, _himl, LVSIL_SMALL);

        // Create columns first
        createListViewColumns(_replaceListView);

        // Update the item count in the ListView
        ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

        // Enable full row selection
        ListView_SetExtendedListViewStyle(_replaceListView, LVS_EX_FULLROWSELECT | LVS_EX_SUBITEMIMAGES);

        return TRUE;
    }
    break;

    case WM_DESTROY:
    {
        DestroyIcon(_hDeleteIcon);
        DestroyIcon(_hEnabledIcon);
        DestroyIcon(_hCopyBackIcon);
        ImageList_Destroy(_himl);
        DestroyWindow(_hSelf);
    }
    break;

    case WM_GETMINMAXINFO:
    {
        MessageBox(_hSelf, TEXT("WM_GETMINMAXINFO erhalten"), TEXT("Debug"), MB_OK);

        MINMAXINFO* minMaxInfo = (MINMAXINFO*)lParam;
        minMaxInfo->ptMinTrackSize.x = 400; // Minimum width
        minMaxInfo->ptMinTrackSize.y = 300; // Minimum height
        minMaxInfo->ptMaxTrackSize.x = 800; // Maximum width
        minMaxInfo->ptMaxTrackSize.y = 600; // Maximum height
        return 0;
    }
    break;

    case WM_SIZE:
    {
        int newWidth = LOWORD(lParam);
        int newHeight = HIWORD(lParam);

        // Move and resize Find and Replace text boxes
        MoveWindow(GetDlgItem(_hSelf, IDC_FIND_EDIT), 120, 14, newWidth - 360, 200, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), 120, 58, newWidth - 360, 200, TRUE);

        // Move and resize the List
        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_LIST), 14, 250, newWidth - 255, newHeight - 270, TRUE);

        // Move buttons
        int buttonGap = 40;
        int buttonX = newWidth - buttonGap - 160;

        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON), buttonX, 14, 160, 30, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_MARK_MATCHES_BUTTON), buttonX, 80, 160, 30, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_CLEAR_MARKS_BUTTON), buttonX, 120, 160, 30, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_COPY_MARKED_TEXT_BUTTON), buttonX, 160, 160, 30, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_COPY_TO_LIST_BUTTON), buttonX, 215, 160, 60, TRUE);
        MoveWindow(GetDlgItem(_hSelf, IDC_REPLACE_ALL_IN_LIST_BUTTON), buttonX, 300, 160, 30, TRUE);

        // Set a timer to delay column resizing
        SetTimer(_hSelf, RESIZE_TIMER_ID, 100, NULL); // 100 ms delay

        return 0;
    }
    break;

    case WM_TIMER:
    {
        if (wParam == RESIZE_TIMER_ID) {
            // Resize columns after timer expires
            KillTimer(_hSelf, RESIZE_TIMER_ID);
            updateColumnWidths(_replaceListView);
        }
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
                    // Remove the item from the ListView
                    ListView_DeleteItem(_replaceListView, pnmia->iItem);

                    // Remove the item from the replaceListData vector
                    replaceListData.erase(replaceListData.begin() + pnmia->iItem);

                    // Update the item count in the ListView
                    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

                    InvalidateRect(_replaceListView, NULL, TRUE);
                }
                else if (pnmia->iSubItem == 7) { // Copy Back button column
                    // Copy the data from the selected item back to the source interfaces
                    ReplaceItemData& itemData = replaceListData[pnmia->iItem];
                    // Update the controls directly
                    SetWindowTextW(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText.c_str());
                    SetWindowTextW(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText.c_str());
                    SendMessageW(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, itemData.wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);
                    SendMessageW(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, itemData.matchCase ? BST_CHECKED : BST_UNCHECKED, 0);
                    SendMessageW(GetDlgItem(_hSelf, IDC_NORMAL_RADIO), BM_SETCHECK, (!itemData.regexSearch && !itemData.extended) ? BST_CHECKED : BST_UNCHECKED, 0);
                    SendMessageW(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_SETCHECK, itemData.regexSearch ? BST_CHECKED : BST_UNCHECKED, 0);
                    SendMessageW(GetDlgItem(_hSelf, IDC_EXTENDED_RADIO), BM_SETCHECK, itemData.extended ? BST_CHECKED : BST_UNCHECKED, 0);
                }
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
                        plvdi->item.mask |= LVIF_IMAGE;
                        plvdi->item.iImage = enabledIconIndex;
                    }
                    break;
                case 3:
                    if (itemData.matchCase) {
                        plvdi->item.mask |= LVIF_IMAGE;
                        plvdi->item.iImage = enabledIconIndex;
                    }
                    break;
                case 4:
                    if (!itemData.regexSearch && !itemData.extended) {
                        plvdi->item.mask |= LVIF_IMAGE;
                        plvdi->item.iImage = enabledIconIndex;
                    }
                    break;
                case 5:
                    if (itemData.regexSearch) {
                        plvdi->item.mask |= LVIF_IMAGE;
                        plvdi->item.iImage = enabledIconIndex;
                    }
                    break;
                case 6:
                    if (itemData.extended) {
                        plvdi->item.mask |= LVIF_IMAGE;
                        plvdi->item.iImage = enabledIconIndex;
                    }
                    break;
                case 7:
                    plvdi->item.mask |= LVIF_IMAGE;
                    plvdi->item.iImage = copyBackIconIndex;
                    break;
                case 8:
                    plvdi->item.mask |= LVIF_IMAGE;
                    plvdi->item.iImage = deleteIconIndex;
                    break;

                }
            }
            break;
            }
        }
    }
    break;



    case WM_COMMAND:
    {
        switch (wParam)
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

        // Add these case blocks for IDC_NORMAL_RADIO and IDC_EXTENDED_RADIO
        case IDC_NORMAL_RADIO:
        case IDC_EXTENDED_RADIO:
        {
            // Enable the Whole word checkbox
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), TRUE);
        }
        break;

        case IDC_REPLACE_ALL_BUTTON:
        {
            TCHAR findText[256];
            TCHAR replaceText[256];
            GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
            GetDlgItemText(_hSelf, IDC_REPLACE_EDIT, replaceText, 256);
            bool regexSearch = false;
            bool extended = false;

            // Get the state of the Whole word checkbox
            bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

            // Get the state of the Match case checkbox
            bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

            if (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED) {
                regexSearch = true;
            }
            else if (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED) {
                extended = true;
            }

            // Perform the Find and Replace operation if Find field has a value
            if (findText[0] != '\0') {
                ::SendMessage(_curScintilla, SCI_BEGINUNDOACTION, 0, 0);
                findAndReplace(findText, replaceText, wholeWord, matchCase, regexSearch, extended);
                ::SendMessage(_curScintilla, SCI_ENDUNDOACTION, 0, 0);
            }

            // Add the entered text to the combo box history
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);
        }
        break;

        case IDC_MARK_MATCHES_BUTTON:
        {
            TCHAR findText[256];
            GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
            bool regexSearch = false;
            bool extended = false;

            // Get the state of the Whole word checkbox
            bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

            // Get the state of the Match case checkbox
            bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

            if (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED) {
                regexSearch = true;
            }
            else if (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED) {
                extended = true;
            }

            // Perform the Mark Matching Strings operation if Find field has a value
            if (findText[0] != '\0') {
                markMatchingStrings(findText, wholeWord, matchCase, regexSearch, extended);
            }

            // Add the entered text to the combo box history
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
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

        case IDC_COPY_TO_LIST_BUTTON:
        {
            onCopyToListButtonClick();

        }
        break;

        case IDC_REPLACE_ALL_IN_LIST_BUTTON:
        {
            onReplaceAllInListButtonClick();
        }
        break;

        default:
            return FALSE;
        }

    }
    break;

    default:
        return DockingDlgInterface::run_dlgProc(message, wParam, lParam);
    }
    return FALSE;
}


void MultiReplacePanel::updateColumnWidths(HWND listView)
{
    // Get the current ListView client area width
    RECT listViewRect;
    GetClientRect(listView, &listViewRect);
    int listViewWidth = listViewRect.right - listViewRect.left;

    // Calculate the total width of columns 3 to 8
    int columns3to7Width = 0;
    for (int i = 2; i < 9; i++)
    {
        columns3to7Width += ListView_GetColumnWidth(listView, i);
    }

    // Calculate the remaining width for the first two columns
    int remainingWidth = listViewWidth - columns3to7Width;

    // Set the width of the first two columns
    ListView_SetColumnWidth(listView, 0, remainingWidth / 2);
    ListView_SetColumnWidth(listView, 1, remainingWidth / 2);
}



void MultiReplacePanel::addStringToComboBoxHistory(HWND hComboBox, const TCHAR* str, int maxItems)
{
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
