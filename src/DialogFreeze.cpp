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

#include "DialogFreeze.h"

#include <commctrl.h>

namespace {

    // Labels, group boxes and plain owner-drawn text take no input
    bool isOperable(HWND control)
    {
        wchar_t cls[32] = {};
        GetClassNameW(control, cls, 32);
        const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
        if (lstrcmpiW(cls, WC_STATICW) == 0)
            return (style & SS_NOTIFY) != 0;
        if (lstrcmpiW(cls, WC_BUTTONW) == 0)
            return (style & BS_TYPEMASK) != BS_GROUPBOX;
        return true;
    }

    void focusControl(HWND dialog, HWND control)
    {
        SendMessageW(dialog, WM_NEXTDLGCTL, reinterpret_cast<WPARAM>(control), TRUE);
    }

} // namespace

DialogFreeze::DialogFreeze(HWND dialog, HWND keep)
    : _dialog(dialog), _keep(keep)
{
    for (HWND child = GetWindow(dialog, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT)) {
        if (child != keep && IsWindowEnabled(child) && isOperable(child))
            _frozen.push_back(child);
    }

    const HWND focus = GetFocus();
    if (focus && IsChild(dialog, focus))
        _focus = focus;

    if (keep) {
        _keepWasEnabled = IsWindowEnabled(keep) != FALSE;
        EnableWindow(keep, TRUE);
        if (_focus && IsWindowVisible(keep))
            focusControl(dialog, keep);     // first: disabling the focused control leaves no focus
    }

    for (HWND control : _frozen)
        EnableWindow(control, FALSE);
}

DialogFreeze::~DialogFreeze()
{
    for (HWND control : _frozen)
        EnableWindow(control, TRUE);

    // Focus returns unless it was moved away from the kept control meanwhile
    if (_focus && _keep && GetFocus() == _keep
        && IsWindow(_focus) && IsWindowEnabled(_focus) && IsWindowVisible(_focus))
        focusControl(_dialog, _focus);

    if (_keep && !_keepWasEnabled)
        EnableWindow(_keep, FALSE);
}
