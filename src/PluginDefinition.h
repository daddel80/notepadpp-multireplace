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

#ifndef PLUGINDEFINITION_H
#define PLUGINDEFINITION_H

//
// All difinitions of plugin interface
//
#include "PluginInterface.h"
#include <string>

//-------------------------------------//
//-- STEP 1. DEFINE YOUR PLUGIN NAME --//
//-------------------------------------//
// Here define your plugin name
//
const TCHAR NPP_PLUGIN_NAME[] = TEXT("MultiReplace");

//-----------------------------------------------//
//-- STEP 2. DEFINE YOUR PLUGIN COMMAND NUMBER --//
//-----------------------------------------------//
// The plugin menu from top to bottom. Notepad++ keeps a user's shortcut by position,
// so moving an entry hands its shortcut to whatever takes that place.
namespace MenuPos {
    enum : int {
        Open, Separator1, Settings, Documentation, Separator2,
        FocusResults, NextResult, PrevResult, Separator3,
        MarkReplacements, NextReplacement, PrevReplacement, RevertReplacement, Separator4,
        TandemMode, ReopenOnStartup, Separator5,
        About,
        Count
    };
}
const int nbFunc = MenuPos::Count;


//
// Initialization of your plugin data
// It will be called while plugin loading
//
void pluginInit(HINSTANCE hModule);

//
// Cleaning of your plugin
// It will be called while plugin unloading
//
void pluginCleanUp();

//
//Initialization of your plugin commands
//
void commandMenuInit();

//
//Clean up your plugin commands allocation (if any)
//
void commandMenuCleanUp();

//
// Function which sets your command 
//
bool setCommand(size_t index, TCHAR* cmdName, PFUNCPLUGINCMD pFunc, ShortcutKey* sk = NULL, bool check0nInit = false);


//
// Your plugin command functions
//
void multiReplace();
void openHelpLink();
void about();
void multiReplaceConfig();
void focusResultDock();
void gotoNextFound();
void gotoPrevFound();
void toggleTandemMode();
void syncTandemMenuCheckmark();
void toggleReopenOnStartup();
void toggleMarkReplacements();
void gotoNextReplacement();
void gotoPrevReplacement();
void revertReplacement();

// A menu with the Mark Replacements commands opens: grayed while the current document has no marks
void syncMarkReplacementsMenu(HMENU menu);

// The commands the editor's context menu repeats on a replacement mark, as the plugin menu shows them
enum class MarkCommand { Next, Previous, Revert };
std::wstring markCommandText(MarkCommand command);

//
// Plugin menu language refresh (called on NPPN_NATIVELANGCHANGED)
//
void refreshPluginMenu();

#endif //PLUGINDEFINITION_H