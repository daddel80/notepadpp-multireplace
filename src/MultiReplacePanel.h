// This file is part of Notepad++ project
// Copyright (C)2022 Don HO <don.h@free.fr>

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#ifndef MULTI_REPLACE_PANEL_H
#define MULTI_REPLACE_PANEL_H

#include "DockingFeature\DockingDlgInterface.h"
#include "resource.h"
#include <string>

typedef std::basic_string<TCHAR> generic_string;

class MultiReplacePanel : public DockingDlgInterface
{
public:
    MultiReplacePanel() : DockingDlgInterface(IDD_REPLACE_DIALOG) {};

    virtual void display(bool toShow = true) const {
        DockingDlgInterface::display(toShow);
    };

    void setParent(HWND parent2set) {
        _hParent = parent2set;
    };

protected:
    virtual INT_PTR CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
    static void addStringToComboBoxHistory(HWND hComboBox, const TCHAR* str, int maxItems = 10);
};

#endif // MULTI_REPLACE_PANEL_H
