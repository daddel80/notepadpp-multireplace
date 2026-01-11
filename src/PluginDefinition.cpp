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
#include "MultiReplaceConfigDialog.h"
#include "LanguageManager.h"
#include <string>


MultiReplace _MultiReplace;
MultiReplaceConfigDialog _MultiReplaceConfig;
AboutDialog _AboutDialog;

//INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);


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
    _MultiReplace.init((HINSTANCE)hModule, NULL);
    _MultiReplaceConfig.init((HINSTANCE)hModule, NULL);
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
    LanguageManager& LM = LanguageManager::instance();

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
    setCommand(0, const_cast<TCHAR*>(LM.getLPCW(L"menu_multiple_replacement")), multiReplace, NULL, false);
    setCommand(1, TEXT("SEPARATOR"), NULL, NULL, false);
    setCommand(2, const_cast<TCHAR*>(LM.getLPCW(L"menu_settings")), multiReplaceConfig, NULL, false);
    setCommand(3, const_cast<TCHAR*>(LM.getLPCW(L"menu_documentation")), openHelpLink, NULL, false);
    setCommand(4, const_cast<TCHAR*>(LM.getLPCW(L"menu_about")), about, NULL, false);
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

    const bool isSeparator = (cmdName && lstrcmpi(cmdName, TEXT("SEPARATOR")) == 0);
    if (!pFunc && !isSeparator)
        return false;

    if (cmdName)
        lstrcpy(funcItem[index]._itemName, cmdName);
    else
        funcItem[index]._itemName[0] = 0;

    funcItem[index]._pFunc = pFunc;
    funcItem[index]._init2Check = check0nInit;
    funcItem[index]._pShKey = sk;
    return true;
}


//----------------------------------------------//
//-- STEP 4. DEFINE YOUR ASSOCIATED FUNCTIONS --//
//----------------------------------------------//


void multiReplace()
{
    _MultiReplace.setParent(nppData._nppHandle);
    if (!_MultiReplace.isCreated())
    {
        _MultiReplace.create(IDD_REPLACE_DIALOG);
    }
    _MultiReplace.display();
}

void openHelpLink()
{
    ShellExecute(NULL, TEXT("open"), TEXT("https://github.com/daddel80/notepadpp-multireplace#readme"), NULL, NULL, SW_SHOWNORMAL);
}

void about()
{
    _AboutDialog.init(hInst, nppData._nppHandle);
    _AboutDialog.doDialog();
}

void multiReplaceConfig()
{
    // Ensure correct parent window (Notepad++ main window)
    _MultiReplaceConfig.init(hInst, nppData._nppHandle);

    if (!_MultiReplaceConfig.isCreated())
    {
        _MultiReplaceConfig.create(IDD_MULTIREPLACE_CONFIG);
    }

    _MultiReplaceConfig.display(true);
}

//
// Refresh plugin menu text when UI language changes (NPPN_NATIVELANGCHANGED)
// This updates the menu items without requiring Notepad++ restart
//
void refreshPluginMenu()
{
    LanguageManager& LM = LanguageManager::instance();

    // Mapping: funcItem index -> language key
    // Index 1 is SEPARATOR, we skip it
    static const struct {
        int index;
        const wchar_t* langKey;
    } menuMappings[] = {
        { 0, L"menu_multiple_replacement" },
        // Index 1 is SEPARATOR - skip
        { 2, L"menu_settings" },
        { 3, L"menu_documentation" },
        { 4, L"menu_about" },
    };

    // Get main menu handle from Notepad++
    HMENU hMainMenu = reinterpret_cast<HMENU>(
        ::SendMessage(nppData._nppHandle, NPPM_GETMENUHANDLE, 0, 0)
        );

    if (!hMainMenu) return;

    // Update each menu item using its command ID
    for (const auto& mapping : menuMappings) {
        if (mapping.index < 0 || mapping.index >= nbFunc) continue;

        int cmdId = funcItem[mapping.index]._cmdID;
        if (cmdId == 0) continue;  // Command ID not yet assigned

        std::wstring newText = LM.get(mapping.langKey);

        // Update the internal funcItem name (for consistency with Notepad++ internals)
        lstrcpyn(funcItem[mapping.index]._itemName, newText.c_str(), nbChar);

        // Update the actual menu item text
        MENUITEMINFOW mii = {};
        mii.cbSize = sizeof(MENUITEMINFOW);
        mii.fMask = MIIM_STRING;
        mii.dwTypeData = const_cast<wchar_t*>(newText.c_str());

        // SetMenuItemInfoW with FALSE = search by command ID (not by position)
        // This works because Notepad++ assigns unique command IDs to plugin menu items
        ::SetMenuItemInfoW(hMainMenu, cmdId, FALSE, &mii);
    }

    // Force the menu bar to redraw to show the changes immediately
    ::DrawMenuBar(nppData._nppHandle);
}