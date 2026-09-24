// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
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

// QA harness for the batch freeze (DialogFreeze) against the real user32. A dialog
// with the control kinds of the panel (labels, group box, owner-drawn text and
// clickable statics, push/split buttons, check and radio boxes, combo, edit, list
// view, tab strip, a hidden control, Cancel) is frozen and released:
//   A. every operable control is disabled, labels, group box and status text are not
//   B. controls disabled before the freeze (mode states) stay disabled afterwards
//   C. Cancel is enabled while frozen, disabled again afterwards; the focus waits on
//      it and returns to the control that had it
//   D. a frozen control takes no mouse input: WindowFromPoint passes it by, so drag
//      and drop and context menus cannot reach it; owner-draw sees ODS_DISABLED
//   E. a nested freeze changes nothing; focus moved to another window is not taken back
//   F. no Cancel, a Cancel already enabled, a dialog destroyed while frozen
//   G. a list view kept from WM_ENABLE (the panel's is, for its dark mode look) is
//      frozen all the same: the window manager, not the control, blocks the input
//
// Build (Developer Command Prompt):
//   cl /nologo /EHsc /std:c++20 /W4 /DUNICODE /D_UNICODE dialog_freeze_qa.cpp ..\DialogFreeze.cpp user32.lib comctl32.lib
// Build (Linux, runs under Wine):
//   x86_64-w64-mingw32-g++ -std=c++20 -Wall -Wextra -DUNICODE -D_UNICODE -static
//       -o dialog_freeze_qa.exe dialog_freeze_qa.cpp ../DialogFreeze.cpp -lcomctl32
//   xvfb-run -a wine64 dialog_freeze_qa.exe

#include "../DialogFreeze.h"

#include <commctrl.h>
#include <cstdio>
#include <map>
#include <string>
#include <vector>

#ifdef _MSC_VER   // the panel lives in N++ with common controls 6
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

static int failures = 0;
static void CHECK(const std::string& what, bool ok) { std::printf("%s %s\n", ok ? "PASS" : "FAIL", what.c_str()); if (!ok) ++failures; }

enum Id {
    ID_LABEL = 100, ID_STATUS, ID_ENGINE, ID_GROUP, ID_FIND_ALL, ID_SAVE, ID_OPEN, ID_BOOKMARK,
    ID_SELECTION, ID_WHOLE_WORD, ID_FIND_EDIT, ID_DIR_EDIT, ID_LIST, ID_TABS, ID_NEW_LIST, ID_LEGACY, ID_CANCEL
};

static std::map<UINT, UINT> g_drawState;   // CtlID -> itemState of the last WM_DRAWITEM

static INT_PTR CALLBACK dlgProc(HWND, UINT msg, WPARAM, LPARAM lParam)
{
    if (msg == WM_INITDIALOG) return TRUE;
    if (msg == WM_DRAWITEM) {
        const auto* dis = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam);
        g_drawState[dis->CtlID] = dis->itemState;
        return TRUE;
    }
    return FALSE;
}

static LRESULT CALLBACK ignoreEnable(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR)
{
    if (msg == WM_ENABLE) return 0;
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static void pump()
{
    MSG m;
    while (PeekMessageW(&m, nullptr, 0, 0, PM_REMOVE)) { TranslateMessage(&m); DispatchMessageW(&m); }
}

static HWND makeDialog(int x, int y)
{
    struct alignas(4) { DLGTEMPLATE t; WORD menu, cls, title; } tmpl{
        { WS_POPUP | WS_CAPTION | WS_VISIBLE, WS_EX_TOPMOST, 0, static_cast<short>(x), static_cast<short>(y), 330, 220 }, 0, 0, 0 };
    return CreateDialogIndirectParamW(GetModuleHandleW(nullptr), &tmpl.t, nullptr, dlgProc, 0);
}

static HWND add(HWND dlg, const wchar_t* cls, const wchar_t* text, DWORD style, int id, int x, int y, int w, int h, bool visible = true)
{
    return CreateWindowExW(0, cls, text, WS_CHILD | (visible ? WS_VISIBLE : 0) | style, x, y, w, h,
        dlg, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), GetModuleHandleW(nullptr), nullptr);
}

static void setFocus(HWND dlg, HWND control) { SendMessageW(dlg, WM_NEXTDLGCTL, reinterpret_cast<WPARAM>(control), TRUE); pump(); }
static bool enabled(HWND dlg, int id) { return IsWindowEnabled(GetDlgItem(dlg, id)) != FALSE; }

static POINT centerOf(HWND control)
{
    RECT r{};
    GetWindowRect(control, &r);
    return { (r.left + r.right) / 2, (r.top + r.bottom) / 2 };
}

static UINT redrawnState(HWND dlg, int id)
{
    g_drawState.erase(static_cast<UINT>(id));
    RedrawWindow(GetDlgItem(dlg, id), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
    pump();
    const auto it = g_drawState.find(static_cast<UINT>(id));
    return it == g_drawState.end() ? 0xFFFFFFFF : it->second;
}

// The panel's control kinds; Selection and Match whole word start disabled as in
// Files mode with a regex search, Cancel starts disabled as between two runs.
static HWND buildPanel(int x, int y)
{
    HWND dlg = makeDialog(x, y);
    add(dlg, WC_STATICW, L"Find what:", SS_RIGHT, ID_LABEL, 5, 5, 60, 16);
    add(dlg, WC_STATICW, L"", SS_LEFT | SS_OWNERDRAW, ID_STATUS, 5, 25, 200, 16);
    add(dlg, WC_STATICW, L"(L)", SS_CENTER | SS_OWNERDRAW | SS_NOTIFY, ID_ENGINE, 210, 25, 20, 16);
    add(dlg, WC_BUTTONW, L"Scope", BS_GROUPBOX, ID_GROUP, 5, 45, 150, 70);
    add(dlg, WC_BUTTONW, L"Find All", BS_SPLITBUTTON | WS_TABSTOP, ID_FIND_ALL, 240, 5, 110, 24);
    add(dlg, WC_BUTTONW, L"Save", BS_SPLITBUTTON | WS_TABSTOP, ID_SAVE, 240, 35, 110, 24);
    add(dlg, WC_BUTTONW, L"Open", BS_PUSHBUTTON | WS_TABSTOP, ID_OPEN, 240, 65, 110, 24);
    add(dlg, WC_BUTTONW, L"", BS_AUTOCHECKBOX | WS_TABSTOP, ID_BOOKMARK, 355, 5, 16, 16);
    add(dlg, WC_BUTTONW, L"Selection", BS_AUTORADIOBUTTON | WS_TABSTOP, ID_SELECTION, 10, 60, 100, 16);
    add(dlg, WC_BUTTONW, L"Match whole word", BS_AUTOCHECKBOX | WS_TABSTOP, ID_WHOLE_WORD, 10, 85, 130, 16);
    add(dlg, WC_COMBOBOXW, L"", CBS_DROPDOWN | WS_TABSTOP, ID_FIND_EDIT, 70, 5, 160, 120);
    add(dlg, WC_EDITW, L"C:\\", ES_LEFT | WS_BORDER | WS_TABSTOP, ID_DIR_EDIT, 5, 120, 200, 20);
    add(dlg, WC_LISTVIEWW, L"", LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP, ID_LIST, 5, 150, 300, 150);
    add(dlg, WC_TABCONTROLW, L"", TCS_BOTTOM | WS_TABSTOP, ID_TABS, 5, 305, 300, 24);
    add(dlg, WC_STATICW, L"+", SS_CENTER | SS_OWNERDRAW | SS_NOTIFY, ID_NEW_LIST, 310, 305, 20, 20);
    add(dlg, WC_BUTTONW, L"Legacy", BS_PUSHBUTTON, ID_LEGACY, 360, 65, 60, 24, false);
    add(dlg, WC_BUTTONW, L"Cancel", BS_PUSHBUTTON | WS_TABSTOP, ID_CANCEL, 240, 95, 110, 24);

    TCITEMW tab{};
    tab.mask = TCIF_TEXT;
    tab.pszText = const_cast<LPWSTR>(L"List 1");
    SendMessageW(GetDlgItem(dlg, ID_TABS), TCM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&tab));
    EnableWindow(GetDlgItem(dlg, ID_SELECTION), FALSE);
    EnableWindow(GetDlgItem(dlg, ID_WHOLE_WORD), FALSE);
    EnableWindow(GetDlgItem(dlg, ID_CANCEL), FALSE);
    UpdateWindow(dlg);
    pump();
    return dlg;
}

static const int kOperable[] = { ID_ENGINE, ID_FIND_ALL, ID_SAVE, ID_OPEN, ID_BOOKMARK, ID_FIND_EDIT, ID_DIR_EDIT, ID_LIST, ID_TABS, ID_NEW_LIST, ID_LEGACY };
static const int kPassive[] = { ID_LABEL, ID_STATUS, ID_GROUP };
static const int kModeDisabled[] = { ID_SELECTION, ID_WHOLE_WORD };

static std::string names(HWND dlg, const int* ids, size_t n, bool wantEnabled)
{
    std::string wrong;
    for (size_t i = 0; i < n; ++i)
        if (enabled(dlg, ids[i]) != wantEnabled) wrong += " " + std::to_string(ids[i]);
    return wrong.empty() ? "none" : wrong;
}

int main()
{
    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_STANDARD_CLASSES | ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icc);

    HWND dlg = buildPanel(10, 10);
    HWND cancel = GetDlgItem(dlg, ID_CANCEL);
    HWND findAll = GetDlgItem(dlg, ID_FIND_ALL);
    HWND list = GetDlgItem(dlg, ID_LIST);
    CHECK("setup: dialog and controls created", dlg && cancel && findAll && list);

    SetForegroundWindow(dlg);
    setFocus(dlg, findAll);
    CHECK("setup: focus on Find All", GetFocus() == findAll);
    const bool hitBefore = (WindowFromPoint(centerOf(list)) == list);
    CHECK("setup: the list takes the mouse before the freeze", hitBefore);
    CHECK("setup: owner-drawn statics draw enabled", (redrawnState(dlg, ID_ENGINE) & ODS_DISABLED) == 0);

    {
        DialogFreeze freeze(dlg, cancel);
        pump();

        std::string wrong = names(dlg, kOperable, std::size(kOperable), false);
        CHECK("A1 operable controls disabled, still enabled: " + wrong, wrong == "none");
        wrong = names(dlg, kPassive, std::size(kPassive), true);
        CHECK("A2 label, status text and group box untouched, disabled: " + wrong, wrong == "none");
        COMBOBOXINFO cbi{};
        cbi.cbSize = sizeof(cbi);
        GetComboBoxInfo(GetDlgItem(dlg, ID_FIND_EDIT), &cbi);
        CHECK("A3 the combo's edit field is disabled with it", cbi.hwndItem && !IsWindowEnabled(cbi.hwndItem));
        wrong = names(dlg, kModeDisabled, std::size(kModeDisabled), false);
        CHECK("B1 mode-disabled controls stay disabled, enabled: " + wrong, wrong == "none");
        CHECK("C1 Cancel enabled while frozen", IsWindowEnabled(cancel) != FALSE);
        CHECK("C2 the focus waits on Cancel", GetFocus() == cancel);
        CHECK("D1 WindowFromPoint passes the frozen list by", hitBefore && WindowFromPoint(centerOf(list)) != list);
        CHECK("D2 owner-drawn static sees ODS_DISABLED", (redrawnState(dlg, ID_ENGINE) & ODS_DISABLED) != 0);
        CHECK("D3 the + glyph sees ODS_DISABLED", (redrawnState(dlg, ID_NEW_LIST) & ODS_DISABLED) != 0);

        {
            DialogFreeze nested(dlg, cancel);
            pump();
            CHECK("E1 nested freeze: Cancel still enabled and focused", IsWindowEnabled(cancel) && GetFocus() == cancel);
        }
        pump();
        wrong = names(dlg, kOperable, std::size(kOperable), false);
        CHECK("E2 after the nested release everything is still frozen, enabled: " + wrong, wrong == "none");
        CHECK("E3 after the nested release Cancel is still enabled and focused", IsWindowEnabled(cancel) && GetFocus() == cancel);
    }
    pump();

    std::string wrong = names(dlg, kOperable, std::size(kOperable), true);
    CHECK("B2 operable controls enabled again, still disabled: " + wrong, wrong == "none");
    wrong = names(dlg, kModeDisabled, std::size(kModeDisabled), false);
    CHECK("B3 mode-disabled controls not enabled by the release, enabled: " + wrong, wrong == "none");
    CHECK("B4 the hidden control is enabled and still hidden", enabled(dlg, ID_LEGACY) && !IsWindowVisible(GetDlgItem(dlg, ID_LEGACY)));
    CHECK("C3 Cancel disabled again", IsWindowEnabled(cancel) == FALSE);
    CHECK("C4 the focus is back on Find All", GetFocus() == findAll);
    CHECK("D4 the list takes the mouse again", WindowFromPoint(centerOf(list)) == list);
    CHECK("D5 owner-drawn static draws enabled again", (redrawnState(dlg, ID_ENGINE) & ODS_DISABLED) == 0);

    // A run started from the keyboard: the focus sits in the combo's edit field
    COMBOBOXINFO findInfo{};
    findInfo.cbSize = sizeof(findInfo);
    GetComboBoxInfo(GetDlgItem(dlg, ID_FIND_EDIT), &findInfo);
    setFocus(dlg, GetDlgItem(dlg, ID_FIND_EDIT));
    const HWND findField = GetFocus();
    CHECK("C5 setup: focus in the Find field", findField && findField == findInfo.hwndItem);
    {
        DialogFreeze freeze(dlg, cancel);
        pump();
        CHECK("C6 the focus left the frozen Find field for Cancel", GetFocus() == cancel);
    }
    pump();
    CHECK("C7 the focus is back in the Find field", GetFocus() == findField);

    SetWindowSubclass(list, ignoreEnable, 1, 0);
    {
        DialogFreeze freeze(dlg, cancel);
        pump();
        CHECK("G1 a list kept from WM_ENABLE is disabled all the same", IsWindowEnabled(list) == FALSE);
        CHECK("G2 and takes no mouse input", WindowFromPoint(centerOf(list)) != list);
    }
    pump();
    CHECK("G3 enabled again after the release", IsWindowEnabled(list) != FALSE && WindowFromPoint(centerOf(list)) == list);
    RemoveWindowSubclass(list, ignoreEnable, 1);

    // Focus moved to another window during the run stays there
    HWND other = buildPanel(360, 10);
    HWND otherEdit = GetDlgItem(other, ID_DIR_EDIT);
    SetForegroundWindow(dlg);
    setFocus(dlg, findAll);
    {
        DialogFreeze freeze(dlg, cancel);
        pump();
        SetForegroundWindow(other);
        SetActiveWindow(other);
        SetFocus(otherEdit);
        pump();
        CHECK("E4 setup: focus moved to another window during the freeze", GetFocus() == otherEdit);
    }
    pump();
    CHECK("E5 the release does not take the focus back", GetFocus() == otherEdit);
    CHECK("E6 Cancel disabled, Find All enabled", !IsWindowEnabled(cancel) && IsWindowEnabled(findAll));

    // A freeze while the focus is elsewhere does not pull it into the dialog
    {
        DialogFreeze freeze(dlg, cancel);
        pump();
        CHECK("E7 focus outside the dialog is left alone", GetFocus() == otherEdit);
    }
    pump();
    CHECK("E8 still there after the release", GetFocus() == otherEdit);
    DestroyWindow(other);
    pump();

    // Without a kept control everything operable is frozen and restored
    SetForegroundWindow(dlg);
    setFocus(dlg, findAll);
    {
        DialogFreeze freeze(dlg, nullptr);
        pump();
        wrong = names(dlg, kOperable, std::size(kOperable), false);
        CHECK("F1 no kept control: operable controls disabled, enabled: " + wrong, wrong == "none");
        CHECK("F2 no kept control: the disabled Find All holds no focus", GetFocus() != findAll);
    }
    pump();
    wrong = names(dlg, kOperable, std::size(kOperable), true);
    CHECK("F3 no kept control: restored, disabled: " + wrong, wrong == "none");
    CHECK("F4 no kept control: Cancel untouched (disabled)", IsWindowEnabled(cancel) == FALSE);

    // A kept control that was enabled stays enabled
    EnableWindow(cancel, TRUE);
    {
        DialogFreeze freeze(dlg, cancel);
        pump();
    }
    pump();
    CHECK("F5 a Cancel enabled before stays enabled", IsWindowEnabled(cancel) != FALSE);
    EnableWindow(cancel, FALSE);

    // The dialog goes away while frozen (N++ shutting down)
    {
        HWND gone = buildPanel(20, 20);
        DialogFreeze freeze(gone, GetDlgItem(gone, ID_CANCEL));
        DestroyWindow(gone);
        pump();
    }
    CHECK("F6 release after the dialog was destroyed does not crash", true);

    DestroyWindow(dlg);
    pump();
    std::printf("\n%s: %d failure(s)\n", failures ? "FAILED" : "ALL PASSED", failures);
    return failures ? 1 : 0;
}
