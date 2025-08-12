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
#include <vector>
#include <windows.h>

// Centralized UI state switch for batch operations
void MultiReplace::setBatchUIState(HWND hDlg, bool inProgress) {
    static const int controlsToDisable[] = {
        IDC_REPLACE_ALL_BUTTON, IDC_REPLACE_BUTTON, IDC_FIND_ALL_BUTTON, IDC_MARK_BUTTON,
        IDC_CLEAR_MARKS_BUTTON, IDC_COPY_TO_LIST_BUTTON, IDC_REPLACE_ALL_SMALL_BUTTON,
        IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON, IDC_MARK_MATCHES_BUTTON,
        IDC_COPY_MARKED_TEXT_BUTTON, IDC_LOAD_FROM_CSV_BUTTON, IDC_LOAD_LIST_BUTTON,
        IDC_NEW_LIST_BUTTON, IDC_SAVE_TO_CSV_BUTTON, IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON,
        IDC_EXPORT_BASH_BUTTON, IDC_BROWSE_DIR_BUTTON, IDC_UP_BUTTON, IDC_DOWN_BUTTON,
        IDC_USE_LIST_BUTTON, IDC_SWAP_BUTTON, IDC_COLUMN_SORT_DESC_BUTTON,
        IDC_COLUMN_SORT_ASC_BUTTON, IDC_COLUMN_DROP_BUTTON, IDC_COLUMN_COPY_BUTTON,
        IDC_COLUMN_HIGHLIGHT_BUTTON, IDC_FIND_EDIT, IDC_REPLACE_EDIT, IDC_FILTER_EDIT,
        IDC_DIR_EDIT, IDC_REPLACE_HIT_EDIT, IDC_COLUMN_NUM_EDIT, IDC_DELIMITER_EDIT,
        IDC_QUOTECHAR_EDIT, IDC_WHOLE_WORD_CHECKBOX, IDC_MATCH_CASE_CHECKBOX,
        IDC_USE_VARIABLES_CHECKBOX, IDC_WRAP_AROUND_CHECKBOX,
        IDC_REPLACE_AT_MATCHES_CHECKBOX, IDC_2_BUTTONS_MODE, IDC_SUBFOLDERS_CHECKBOX,
        IDC_HIDDENFILES_CHECKBOX, IDC_NORMAL_RADIO, IDC_EXTENDED_RADIO, IDC_REGEX_RADIO,
        IDC_ALL_TEXT_RADIO, IDC_SELECTION_RADIO, IDC_COLUMN_MODE_RADIO
    };

    // Bulk enable/disable
    for (int id : controlsToDisable)
        EnableWindow(GetDlgItem(hDlg, id), !inProgress);
    EnableWindow(GetDlgItem(hDlg, IDC_CANCEL_REPLACE_BUTTON), inProgress);

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
