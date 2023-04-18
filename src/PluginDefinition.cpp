//this file is part of notepad++
//Copyright (C)2022 Don HO <don.h@free.fr>
//
//This program is free software; you can redistribute it and/or
//modify it under the terms of the GNU General Public License
//as published by the Free Software Foundation; either
//version 2 of the License, or (at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.
//
//You should have received a copy of the GNU General Public License
//along with this program; if not, write to the Free Software
//Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

#include "PluginDefinition.h"
#include "menuCmdID.h"
#include <resource.h>
#include <windowsx.h>

HWND hDlg;
HINSTANCE g_hInst;


INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);

void showReplaceDialog();
void findAndReplace(const TCHAR* findText, const TCHAR* replaceText);


//
// The plugin data that Notepad++ needs
//
FuncItem funcItem[nbFunc];

//
// The data of Notepad++ that you can use in your plugin commands
//
NppData nppData;

//
// Initialize your plugin data here
// It will be called while plugin loading   
void pluginInit(HINSTANCE hModule)
{
    g_hInst = hModule;
}

//
// Here you can do the clean up, save the parameters (if any) for the next session
//
void pluginCleanUp()
{
}

//
// Initialization of your plugin commands
// You should fill your plugins commands here
void commandMenuInit()
{

    //--------------------------------------------//
    //-- STEP 3. CUSTOMIZE YOUR PLUGIN COMMANDS --//
    //--------------------------------------------//
    // with function :
    // setCommand(int index,                      // zero based number to indicate the order of command
    //            TCHAR *commandName,             // the command name that you want to see in plugin menu
    //            PFUNCPLUGINCMD functionPointer, // the symbol of function (function pointer) associated with this command. The body should be defined below. See Step 4.
    //            ShortcutKey *shortcut,          // optional. Define a shortcut to trigger this command
    //            bool check0nInit                // optional. Make this menu item be checked visually
    //            );
    setCommand(0, TEXT("Hello Notepad++"), hello, NULL, false);
    //setCommand(1, TEXT("Hello (with dialog)"), helloDlg, NULL, false);
    setCommand(1, TEXT("Find and Replace"), showReplaceDialog, NULL, false);
    //setCommand(3, TEXT("Hello Notepad++"), hello, NULL, false);
}

//
// Here you can do the clean up (especially for the shortcut)
//
void commandMenuCleanUp()
{
	// Don't forget to deallocate your shortcut here
}


//
// This function help you to initialize your plugin commands
//
bool setCommand(size_t index, TCHAR *cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey *sk, bool check0nInit) 
{
    if (index >= nbFunc)
        return false;

    if (!pFunc)
        return false;

    lstrcpy(funcItem[index]._itemName, cmdName);
    funcItem[index]._pFunc = pFunc;
    funcItem[index]._init2Check = check0nInit;
    funcItem[index]._pShKey = sk;

    return true;
}

//----------------------------------------------//
//-- STEP 4. DEFINE YOUR ASSOCIATED FUNCTIONS --//
//----------------------------------------------//
void hello()
{
    // Open a new document
    ::SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_FILE_NEW);

    // Get the current scintilla
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0)?nppData._scintillaMainHandle:nppData._scintillaSecondHandle;

    // Say hello now :
    // Scintilla control has no Unicode mode, so we use (char *) here
    ::SendMessage(curScintilla, SCI_SETTEXT, 0, (LPARAM)"Hello, Notepad++!");
}

void helloDlg()
{
    ::MessageBox(NULL, TEXT("Hello, Notepad++!"), TEXT("Notepad++ Plugin Template"), MB_OK);
}

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    (void)lParam;
    switch (uMsg) {
    case WM_INITDIALOG:
        // ... [Set focus and return] ...
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_OK_BUTTON:
        {
            TCHAR findText[256];
            TCHAR replaceText[256];
            GetDlgItemText(hwndDlg, IDC_FIND_EDIT, findText, 256);
            GetDlgItemText(hwndDlg, IDC_REPLACE_EDIT, replaceText, 256);

            // Perform the Find and Replace operation
            findAndReplace(findText, replaceText);
        }
        EndDialog(hwndDlg, 0);
        break;

        case IDC_CANCEL_BUTTON:
            EndDialog(hwndDlg, 0);
            break;
        }
        break;

    case WM_CLOSE:
        EndDialog(hwndDlg, 0);
        break;

    default:
        return FALSE;
    }
    return TRUE;
}

void findAndReplace(const TCHAR* findText, const TCHAR* replaceText) {
    // Get the current Scintilla
    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&which);
    if (which == -1)
        return;
    HWND curScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;

    // Convert TCHAR strings to char strings (Scintilla uses char strings)
    int findTextLen = WideCharToMultiByte(CP_UTF8, 0, findText, -1, NULL, 0, NULL, NULL);
    int replaceTextLen = WideCharToMultiByte(CP_UTF8, 0, replaceText, -1, NULL, 0, NULL, NULL);
    char* findTextA = new char[findTextLen];
    char* replaceTextA = new char[replaceTextLen];
    WideCharToMultiByte(CP_UTF8, 0, findText, -1, findTextA, findTextLen, NULL, NULL);
    WideCharToMultiByte(CP_UTF8, 0, replaceText, -1, replaceTextA, replaceTextLen, NULL, NULL);

    // Set the search flags (you can customize these if needed)
    int searchFlags = SCFIND_MATCHCASE | SCFIND_WHOLEWORD;

    // Set the target range to the whole document
    SendMessage(curScintilla, SCI_SETTARGETSTART, 0, 0);
    SendMessage(curScintilla, SCI_SETTARGETEND, SendMessage(curScintilla, SCI_GETLENGTH, 0, 0), 0);

    // Set the search flags
    SendMessage(curScintilla, SCI_SETSEARCHFLAGS, searchFlags, 0);

    // Replace all occurrences
    while (SendMessage(curScintilla, SCI_SEARCHINTARGET, findTextLen - 1, (LPARAM)findTextA) != -1)
    {
        // Replace the found text with the replacement text
        SendMessage(curScintilla, SCI_REPLACETARGET, replaceTextLen - 1, (LPARAM)replaceTextA);

        // Update the target range to start from the end of the last replacement
        LRESULT targetStart = SendMessage(curScintilla, SCI_GETTARGETEND, 0, 0);
        LRESULT textLength = SendMessage(curScintilla, SCI_GETLENGTH, 0, 0);
        SendMessage(curScintilla, SCI_SETTARGETSTART, targetStart, 0);
        SendMessage(curScintilla, SCI_SETTARGETEND, textLength, 0);
    }

    // Cleanup
    delete[] findTextA;
    delete[] replaceTextA;
}

void showReplaceDialog() {
    hDlg = CreateDialog(g_hInst, MAKEINTRESOURCE(IDD_REPLACE_DIALOG), nppData._nppHandle, DialogProc);
    DWORD errCode = GetLastError();
    TCHAR errMsg[256];
    wsprintf(errMsg, TEXT("Error code: %d"), errCode);
    MessageBox(NULL, errMsg, TEXT("Debug"), MB_OK);

    if (hDlg) {
        ShowWindow(hDlg, SW_SHOW);
    }
}