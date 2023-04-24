// This file is part of Notepad++ project
// Copyright (C)2022 Don HO <don.h@free.fr>

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


extern NppData nppData;

#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif


void findAndReplace(const TCHAR* findText, const TCHAR* replaceText, bool wholeWord, bool matchCase, bool regexSearch)
{
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, (WPARAM)0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

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

    LRESULT pos = 0;
    while (pos >= 0)
    {
        ::SendMessage(curScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(curScintilla, SCI_SETTARGETEND, ::SendMessage(curScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

        pos = ::SendMessage(curScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));
        if (pos >= 0)
        {
            if (regexSearch) {
                ::SendMessage(curScintilla, SCI_SETSEL, pos, pos + ::SendMessage(curScintilla, SCI_GETTARGETEND, 0, 0));
                ::SendMessage(curScintilla, SCI_REPLACETARGETRE, (WPARAM)-1, (LPARAM)replaceTextUtf8.c_str());
            }
            else {
                ::SendMessage(curScintilla, SCI_SETSEL, pos, pos + findTextLength);
                ::SendMessage(curScintilla, SCI_REPLACESEL, 0, (LPARAM)replaceTextUtf8.c_str());
            }
            pos += replaceTextLength;
        }
    }
}


void markMatchingStrings(const TCHAR* findText, bool wholeWord, bool matchCase, bool regexSearch)
{
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

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

    LRESULT pos = 0;
    ::SendMessage(curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(curScintilla, SCI_INDICSETSTYLE, 0, INDIC_STRAIGHTBOX);
    ::SendMessage(curScintilla, SCI_INDICSETFORE, 0, 0x007F00);
    ::SendMessage(curScintilla, SCI_INDICSETALPHA, 0, 100);

    while (pos >= 0)
    {
        ::SendMessage(curScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(curScintilla, SCI_SETTARGETEND, ::SendMessage(curScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

        pos = ::SendMessage(curScintilla, SCI_SEARCHINTARGET, findTextLength, reinterpret_cast<LPARAM>(findTextUtf8.c_str()));
        if (pos >= 0)
        {
            ::SendMessage(curScintilla, SCI_SETINDICATORVALUE, 1, 0);
            ::SendMessage(curScintilla, SCI_INDICATORFILLRANGE, pos, findTextLength);
            pos += findTextLength;
        }
    }
}

void clearAllMarks()
{
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
    ::SendMessage(curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    ::SendMessage(curScintilla, SCI_INDICATORCLEARRANGE, 0, ::SendMessage(curScintilla, SCI_GETLENGTH, 0, 0));
}

void copyMarkedTextToClipboard()
{
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

    LRESULT length = ::SendMessage(curScintilla, SCI_GETLENGTH, 0, 0);
    std::string markedText;

    ::SendMessage(curScintilla, SCI_SETINDICATORCURRENT, 0, 0);
    for (int i = 0; i < length; ++i)
    {
        if (::SendMessage(curScintilla, SCI_INDICATORVALUEAT, 0, i))
        {
            char ch = static_cast<char>(::SendMessage(curScintilla, SCI_GETCHARAT, i, 0));
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


INT_PTR CALLBACK MultiReplacePanel::run_dlgProc(UINT message, WPARAM wParam, LPARAM /*lParam*/)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        // ... [Set focus and return] ...
        return TRUE;
    }
    break;

    case WM_COMMAND:
    {
        switch (wParam)
        {
        case IDC_REPLACE_ALL_BUTTON:
        {
            TCHAR findText[256];
            TCHAR replaceText[256];
            GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);
            GetDlgItemText(_hSelf, IDC_REPLACE_EDIT, replaceText, 256);

            // Get the state of the Whole word checkbox
            bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

            // Get the state of the Match case checkbox
            bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

            // Get the state of the Regex checkbox
            bool regexSearch = (IsDlgButtonChecked(_hSelf, IDC_REGEX_CHECKBOX) == BST_CHECKED);

            // Perform the Find and Replace operation
            ::findAndReplace(findText, replaceText, wholeWord, matchCase, regexSearch);

            // Add the entered text to the combo box history
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);
        }
        break;

        case IDC_MARK_MATCHES_BUTTON:
        {
            TCHAR findText[256];
            GetDlgItemText(_hSelf, IDC_FIND_EDIT, findText, 256);

            // Get the state of the Whole word checkbox
            bool wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);

            // Get the state of the Match case checkbox
            bool matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

            // Get the state of the Regex checkbox
            bool regexSearch = (IsDlgButtonChecked(_hSelf, IDC_REGEX_CHECKBOX) == BST_CHECKED);

            // Perform the Mark Matching Strings operation
            ::markMatchingStrings(findText, wholeWord, matchCase, regexSearch);

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

        }
    }
    break;


    }

    return FALSE;
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
