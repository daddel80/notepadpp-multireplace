//this file is part of notepad++
//Copyright (C)2022 Thomas Knoefel
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
#include "MultiReplacePanel.h"
#include "AboutDialog.h"


MultiReplacePanel _MultiReplacePanel;

INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);


//
// The plugin data that Notepad++ needs
//
FuncItem funcItem[nbFunc];

//
// The data of Notepad++ that you can use in your plugin commands
//
NppData nppData;
HINSTANCE hInst;

//
// Initialize your plugin data here
// It will be called while plugin loading   
void pluginInit(HINSTANCE hModule)
{
    _MultiReplacePanel.init((HINSTANCE)hModule, NULL);
    hInst = (HINSTANCE)hModule;
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
    setCommand(0, TEXT("Multi Replace && Mark"), multiReplacePanel, NULL, false);
    setCommand(1, TEXT("About"), about, NULL, false);
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
bool setCommand(size_t index, TCHAR* cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey* sk, bool check0nInit)
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


void multiReplacePanel()
{
    _MultiReplacePanel.setParent(nppData._nppHandle);
    tTbData	data = { 0 };

    if (!_MultiReplacePanel.isCreated())
    {
        _MultiReplacePanel.create(&data);

        // define the default docking behaviour
        data.uMask = DWS_DF_FLOATING;

        data.pszModuleName = _MultiReplacePanel.getPluginFileName();

        // the dlgDlg should be the index of funcItem where the current function pointer is
        // in this case is DOCKABLE_DEMO_INDEX
        data.dlgID = 3;
        ::SendMessage(nppData._nppHandle, NPPM_DMMREGASDCKDLG, 0, (LPARAM)&data);
    }
    _MultiReplacePanel.display();
}

void about()
{
    ShowAboutDialog(nppData._nppHandle);
}
