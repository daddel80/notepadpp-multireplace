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

extern FuncItem funcItem[nbFunc];
extern NppData nppData;

HINSTANCE g_inst;


BOOL APIENTRY DllMain(HINSTANCE hModule, DWORD  reasonForCall, LPVOID /*lpReserved*/)
{
	try {

		switch (reasonForCall)
		{
			case DLL_PROCESS_ATTACH:
				pluginInit(hModule);
                g_inst = HINSTANCE(hModule);
				break;

			case DLL_PROCESS_DETACH:
				pluginCleanUp();
				break;

			case DLL_THREAD_ATTACH:
				break;

			case DLL_THREAD_DETACH:
				break;
		}
	}
	catch (...) { return FALSE; }

    return TRUE;
}


extern "C" __declspec(dllexport) void setInfo(NppData notpadPlusData)
{
	nppData = notpadPlusData;
	commandMenuInit();
}

extern "C" __declspec(dllexport) const TCHAR * getName()
{
	return NPP_PLUGIN_NAME;
}

extern "C" __declspec(dllexport) FuncItem * getFuncsArray(int *nbF)
{
	*nbF = nbFunc;
	return funcItem;
}


extern "C" __declspec(dllexport) void beNotified(SCNotification * notifyCode)
{
    switch (notifyCode->nmhdr.code)
    {
    case NPPN_TBMODIFICATION:
    {
        toolbarIconsWithDarkMode tbIcons;
        tbIcons.hToolbarBmp = ::LoadBitmap(g_inst, MAKEINTRESOURCE(IDR_MR_BMP));
        tbIcons.hToolbarIcon = ::LoadIcon(g_inst, MAKEINTRESOURCE(IDI_MR_ICON));
        tbIcons.hToolbarIconDarkMode = ::LoadIcon(g_inst, MAKEINTRESOURCE(IDI_MR_DM_ICON));
        ::SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_FORDARKMODE, funcItem[0]._cmdID, (LPARAM)&tbIcons);
    }
    break;
    case NPPN_SHUTDOWN:
    {
        commandMenuCleanUp();
    }
    break;
    case SCN_UPDATEUI:
    {
        if (notifyCode->updated & SC_UPDATE_SELECTION)
        {
            MultiReplace::onSelectionChanged();
        }
        MultiReplace::onCaretPositionChanged();
    }
    break;
    case SCN_MODIFIED:
    {
        if (notifyCode->modificationType & (SC_MOD_INSERTTEXT | SC_MOD_DELETETEXT)) {
            MultiReplace::onTextChanged();
            MultiReplace::processTextChange(notifyCode);
            MultiReplace::processLog();
        }
    }
    break;

    case NPPN_BUFFERACTIVATED:
    {
        MultiReplace::onDocumentSwitched();
        MultiReplace::pointerToScintilla();
    }
    break;

    case NPPN_DARKMODECHANGED:
    {
        MultiReplace::onThemeChanged();
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfHandleChange), reinterpret_cast<LPARAM>(_MultiReplace.getHSelf()));
        ::SetWindowPos(_MultiReplace.getHSelf(), nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED); // to redraw titlebar and window
        break;
    }

    default:
        return;
    }
}


// Here you can process the Npp Messages 
// I will make the messages accessible little by little, according to the need of plugin development.
// Please let me know if you need to access to some messages :
// http://sourceforge.net/forum/forum.php?forum_id=482781
//
extern "C" __declspec(dllexport) LRESULT messageProc(UINT /*Message*/, WPARAM /*wParam*/, LPARAM /*lParam*/)
{/*
	if (Message == WM_MOVE)
	{
		::MessageBox(NULL, "move", "", MB_OK);
	}
*/
	return TRUE;
}

#ifdef UNICODE
extern "C" __declspec(dllexport) BOOL isUnicode()
{
    return TRUE;
}
#endif //UNICODE
