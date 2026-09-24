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

#include "MultiReplacePanel.h"
#include <windows.h>

// Window state for batch operations; the controls are frozen by BatchUIGuard
void MultiReplace::setBatchWindowState(HWND hDlg, bool inProgress) {
    // Keep above owner during batch (not global topmost)
    _keepOnTopDuringBatch = inProgress;
    SetWindowLongPtr(hDlg, GWLP_HWNDPARENT, (LONG_PTR)nppData._nppHandle);

    if (inProgress) {
        SetWindowPos(hDlg, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        SetWindowTransparency(hDlg, foregroundTransparency);
    }
    else {
        const bool isActive = (GetActiveWindow() == hDlg);
        SetWindowTransparency(hDlg, isActive ? foregroundTransparency : backgroundTransparency);
    }
}