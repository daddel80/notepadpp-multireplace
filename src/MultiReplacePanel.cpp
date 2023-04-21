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

extern NppData nppData;

#ifdef UNICODE
#define generic_strtoul wcstoul
#define generic_sprintf swprintf
#else
#define generic_strtoul strtoul
#define generic_sprintf sprintf
#endif

#define BCKGRD_COLOR (RGB(255,102,102))
#define TXT_COLOR    (RGB(255,255,255))

// The findAndReplace function
void findAndReplace(const TCHAR* findText, const TCHAR* replaceText, bool wholeWord)
{
    // Get the current Scintilla view
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

    // Set the search flags
    int searchFlags = 0;
    if (wholeWord)
        searchFlags |= SCFIND_WHOLEWORD;

    // Start searching from the beginning of the document
    LRESULT pos = 0;
    while (pos >= 0)
    {
        // Set the search target
        ::SendMessage(curScintilla, SCI_SETTARGETSTART, pos, 0);
        ::SendMessage(curScintilla, SCI_SETTARGETEND, ::SendMessage(curScintilla, SCI_GETLENGTH, 0, 0), 0);
        ::SendMessage(curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

        // Find the text
        pos = ::SendMessage(curScintilla, SCI_SEARCHINTARGET, lstrlen(findText), (LPARAM)findText);
        if (pos >= 0)
        {
            // Select the text
            ::SendMessage(curScintilla, SCI_SETSEL, pos, pos + lstrlen(findText));

            // Replace the selected text
            ::SendMessage(curScintilla, SCI_REPLACESEL, 0, (LPARAM)replaceText);

            // Update the search position
            pos += lstrlen(replaceText);
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

            // Perform the Find and Replace operation
            ::findAndReplace(findText, replaceText, wholeWord);

            // Add the entered text to the combo box history
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
            addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceText);
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
