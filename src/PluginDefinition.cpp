// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.

#include "PluginDefinition.h"
#include "MultiReplacePanel.h"
#include "AboutDialog.h"
#include "MultiReplaceConfigDialog.h"
#include "LanguageManager.h"
#include "ResultDock.h"
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
    setCommand(MenuPos::Open, const_cast<TCHAR*>(LM.getLPCW(L"menu_open")), multiReplace, NULL, false);
    setCommand(MenuPos::Separator1, const_cast<TCHAR*>(TEXT("SEPARATOR")), NULL, NULL, false);
    setCommand(MenuPos::Settings, const_cast<TCHAR*>(LM.getLPCW(L"menu_settings")), multiReplaceConfig, NULL, false);
    setCommand(MenuPos::Documentation, const_cast<TCHAR*>(LM.getLPCW(L"menu_documentation")), openHelpLink, NULL, false);
    setCommand(MenuPos::Separator2, const_cast<TCHAR*>(TEXT("SEPARATOR")), NULL, NULL, false);

    // ResultDock navigation commands (assignable via Settings > Shortcut Mapper > Plugin Commands)
    setCommand(MenuPos::FocusResults, const_cast<TCHAR*>(LM.getLPCW(L"cmd_focus_results")), focusResultDock, NULL, false);
    setCommand(MenuPos::NextResult, const_cast<TCHAR*>(LM.getLPCW(L"cmd_next_result")), gotoNextFound, NULL, false);
    setCommand(MenuPos::PrevResult, const_cast<TCHAR*>(LM.getLPCW(L"cmd_prev_result")), gotoPrevFound, NULL, false);
    setCommand(MenuPos::Separator3, const_cast<TCHAR*>(TEXT("SEPARATOR")), NULL, NULL, false);

    // Mark Replacements and the commands for its marks
    setCommand(MenuPos::MarkReplacements, const_cast<TCHAR*>(LM.getLPCW(L"cmd_mark_replacements")), toggleMarkReplacements, NULL,
        MultiReplace::isMarkReplacementsEnabled());
    setCommand(MenuPos::NextReplacement, const_cast<TCHAR*>(LM.getLPCW(L"cmd_next_replacement")), gotoNextReplacement, NULL, false);
    setCommand(MenuPos::PrevReplacement, const_cast<TCHAR*>(LM.getLPCW(L"cmd_prev_replacement")), gotoPrevReplacement, NULL, false);
    setCommand(MenuPos::RevertReplacement, const_cast<TCHAR*>(LM.getLPCW(L"cmd_revert_replacement")), revertReplacement, NULL, false);
    setCommand(MenuPos::Separator4, const_cast<TCHAR*>(TEXT("SEPARATOR")), NULL, NULL, false);

    // Tandem Mode: docks MR to a N++ edge and follows N++ on move/resize.
    setCommand(MenuPos::TandemMode, const_cast<TCHAR*>(LM.getLPCW(L"menu_tandem_mode")), toggleTandemMode, NULL,
        MultiReplace::isTandemPersistedEnabled());

    // Reopen-on-Startup: reopens the MR panel on N++ launch if it
    // was open at the last shutdown. Opt-in via this menu entry.
    setCommand(MenuPos::ReopenOnStartup, const_cast<TCHAR*>(LM.getLPCW(L"menu_reopen_on_startup")),
        toggleReopenOnStartup, NULL,
        MultiReplace::isReopenOnStartupEnabled());
    setCommand(MenuPos::Separator5, const_cast<TCHAR*>(TEXT("SEPARATOR")), NULL, NULL, false);

    setCommand(MenuPos::About, const_cast<TCHAR*>(LM.getLPCW(L"menu_about")), about, NULL, false);
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
        _MultiReplace.display();
        return;
    }
    // Toggle: the same command (and its shortcut) opens MR when hidden
    // and closes it when visible.
    _MultiReplace.display(!_MultiReplace.isVisible());
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

void focusResultDock()
{
    ResultDock::instance().focusDock();
}

void gotoNextFound()
{
    ResultDock::instance().gotoNextHit();
}

void gotoPrevFound()
{
    ResultDock::instance().gotoPrevHit();
}

// Syncs the plugin menu checkmark with the panel's tandem state.
void syncTandemMenuCheckmark()
{
    const int tandemCmdId = funcItem[MenuPos::TandemMode]._cmdID;
    ::SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK,
        static_cast<WPARAM>(tandemCmdId),
        _MultiReplace.isTandemEnabled() ? TRUE : FALSE);
}

void toggleTandemMode()
{
    _MultiReplace.toggleTandemMode();
    syncTandemMenuCheckmark();
}

// Toggles the reopen-on-startup preference and syncs the menu
// checkmark so it reflects the flag immediately.
void toggleReopenOnStartup()
{
    MultiReplace::toggleReopenOnStartup();

    const int cmdId = funcItem[MenuPos::ReopenOnStartup]._cmdID;
    ::SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK,
        static_cast<WPARAM>(cmdId),
        MultiReplace::isReopenOnStartupEnabled() ? TRUE : FALSE);
}

// Toggles Mark Replacements by its persisted value, so it works the same
// with and without the panel, and syncs the checkmark.
void toggleMarkReplacements()
{
    _MultiReplace.setMarkReplacements(!MultiReplace::isMarkReplacementsEnabled());

    const int cmdId = funcItem[MenuPos::MarkReplacements]._cmdID;
    ::SendMessage(nppData._nppHandle, NPPM_SETMENUITEMCHECK,
        static_cast<WPARAM>(cmdId),
        MultiReplace::isMarkReplacementsEnabled() ? TRUE : FALSE);
}

void gotoNextReplacement()
{
    _MultiReplace.gotoReplacement(true);
}

void gotoPrevReplacement()
{
    _MultiReplace.gotoReplacement(false);
}

void revertReplacement()
{
    _MultiReplace.revertReplacement();
}

// The menu positions of the commands the editor's context menu repeats, in MarkCommand order
static constexpr int MARK_COMMAND_POS[] = { MenuPos::NextReplacement, MenuPos::PrevReplacement, MenuPos::RevertReplacement };

void syncMarkReplacementsMenu(HMENU menu)
{
    const UINT next = static_cast<UINT>(funcItem[MenuPos::NextReplacement]._cmdID);
    if (!menu || ::GetMenuState(menu, next, MF_BYCOMMAND) == static_cast<UINT>(-1)) return;

    const UINT state = MF_BYCOMMAND | (_MultiReplace.hasReplaceMarks() ? MF_ENABLED : MF_GRAYED);
    for (const int index : MARK_COMMAND_POS)
        ::EnableMenuItem(menu, static_cast<UINT>(funcItem[index]._cmdID), state);
}

std::wstring markCommandText(MarkCommand command)
{
    const FuncItem& item = funcItem[MARK_COMMAND_POS[static_cast<int>(command)]];
    const HMENU menu = reinterpret_cast<HMENU>(::SendMessage(nppData._nppHandle, NPPM_GETMENUHANDLE, NPPPLUGINMENU, 0));
    wchar_t text[nbChar * 2] = {};   // with the shortcut after a tab
    if (menu && ::GetMenuStringW(menu, static_cast<UINT>(item._cmdID), text, static_cast<int>(std::size(text)), MF_BYCOMMAND) > 0)
        return text;
    return item._itemName;
}

//
// Refresh plugin menu text when UI language changes (NPPN_NATIVELANGCHANGED)
// This updates the menu items without requiring Notepad++ restart
//
void refreshPluginMenu()
{
    LanguageManager& LM = LanguageManager::instance();

    // Mapping: menu position -> language key; separators have no text
    static const struct {
        int index;
        const wchar_t* langKey;
    } menuMappings[] = {
        { MenuPos::Open, L"menu_open" },
        { MenuPos::Settings, L"menu_settings" },
        { MenuPos::Documentation, L"menu_documentation" },
        { MenuPos::FocusResults, L"cmd_focus_results" },
        { MenuPos::NextResult, L"cmd_next_result" },
        { MenuPos::PrevResult, L"cmd_prev_result" },
        { MenuPos::MarkReplacements, L"cmd_mark_replacements" },
        { MenuPos::NextReplacement, L"cmd_next_replacement" },
        { MenuPos::PrevReplacement, L"cmd_prev_replacement" },
        { MenuPos::RevertReplacement, L"cmd_revert_replacement" },
        { MenuPos::TandemMode, L"menu_tandem_mode" },
        { MenuPos::ReopenOnStartup, L"menu_reopen_on_startup" },
        { MenuPos::About, L"menu_about" },
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
        (void)lstrcpyn(funcItem[mapping.index]._itemName, newText.c_str(), nbChar);

        // Keep the shortcut Notepad++ shows after a tab
        wchar_t current[nbChar * 2] = {};
        if (::GetMenuStringW(hMainMenu, cmdId, current, static_cast<int>(std::size(current)), MF_BYCOMMAND) > 0) {
            if (const wchar_t* shortcut = wcschr(current, L'\t')) newText += shortcut;
        }

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