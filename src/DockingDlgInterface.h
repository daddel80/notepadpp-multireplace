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

#include "ResultDock.h"
#include "Scintilla.h"
#include "StaticDialog/Docking.h"
#include "PluginDefinition.h"
#include "StaticDialog/DockingDlgInterface.h"
#include "StaticDialog/resource.h"      // IDD_MULTIREPLACE_RESULT_DOCK, IDI_MR_ICON

extern HINSTANCE g_inst;   // defined in PluginDefinition.cpp
extern NppData   nppData;  // provided by Notepad++ core

/* -------------------------------------------------------------------------
   Singleton accessor
------------------------------------------------------------------------- */
ResultDock& ResultDock::instance()
{
    static ResultDock s{ g_inst };   // constructed exactly once
    return s;
}

/* -------------------------------------------------------------------------
   Local helper – UTF-16 → UTF-8, TU-local to avoid duplicate symbols
------------------------------------------------------------------------- */
static std::string wstringToUtf8(const std::wstring& w)
{
    if (w.empty()) return {};
    int len = ::WideCharToMultiByte(CP_UTF8, 0,
        w.data(), static_cast<int>(w.size()),
        nullptr, 0, nullptr, nullptr);
    std::string out(len, '\0');
    ::WideCharToMultiByte(CP_UTF8, 0,
        w.data(), static_cast<int>(w.size()),
        out.data(), len, nullptr, nullptr);
    return out;
}

/* -------------------------------------------------------------------------
   Public API – replace the buffer
------------------------------------------------------------------------- */
void ResultDock::setText(const std::wstring& wText)
{
    if (!_hSci)
        _createDock(nppData);

    std::string u8 = wstringToUtf8(wText);
    ::SendMessage(_hSci, SCI_SETTEXT, 0, reinterpret_cast<LPARAM>(u8.c_str()));
}

/* -------------------------------------------------------------------------
   Public API – show (and create if needed)
------------------------------------------------------------------------- */
void ResultDock::ensureShown(const NppData& npp)
{
    if (!_hSci)
        _createDock(npp);

    // _hDock verwenden, um das korrekte Container-Fenster anzuzeigen
    if (_hDock)
    {
        ::SendMessage(npp._nppHandle, NPPM_DMMSHOW, 0,
            reinterpret_cast<LPARAM>(_hDock));
    }
}

/* -------------------------------------------------------------------------
   Internal – create Scintilla, register as dock
------------------------------------------------------------------------- */
void ResultDock::_createDock(const NppData& npp)
{
    // 1) Scintilla control
    _hSci = ::CreateWindowExW(0, L"Scintilla", L"", WS_CHILD | WS_VISIBLE,
        0, 0, 100, 100,
        npp._nppHandle, nullptr, _hInst, nullptr);
    ::SendMessage(_hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
    _initFolding();

    // 2) Docking descriptor (fields available in your Docking.h)
    tTbData dock{};
    dock.hClient = _hSci;
    dock.pszName = L"MultiReplace – Search results";
    dock.dlgID = IDD_MULTIREPLACE_RESULT_DOCK;
    dock.uMask = DWS_DF_CONT_BOTTOM | DWS_ICONTAB;
    dock.rcFloat = { 200,200,600,400 };
    dock.hIconTab = ::LoadIcon(_hInst, MAKEINTRESOURCE(IDI_MR_ICON));
    dock.pszAddInfo = L"";
    dock.pszModuleName = NPP_PLUGIN_NAME;

    _hDock = (HWND)::SendMessage(npp._nppHandle, NPPM_DMMREGASDCKDLG,
        0, reinterpret_cast<LPARAM>(&dock));

    if (!_hDock) {
        MessageBoxW(nullptr, L"Dock registration failed",
            L"MultiReplace", MB_ICONERROR);
        return;                       // abort: no container
    }
    _initFolding();
}

/* -------------------------------------------------------------------------
   Internal – folding margin setup
------------------------------------------------------------------------- */
void ResultDock::_initFolding() const
{
    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { ::SendMessage(_hSci, m, w, l); };

    constexpr int MARGIN_FOLD = 2;
    S(SCI_SETMARGINTYPEN, MARGIN_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN, MARGIN_FOLD, SC_MASK_FOLDERS);
    S(SCI_SETMARGINWIDTHN, MARGIN_FOLD, 16);

    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER);

    S(SCI_SETPROPERTY, (sptr_t)"fold", (sptr_t)"1");
    S(SCI_SETPROPERTY, (sptr_t)"fold.compact", (sptr_t)"1");
}
