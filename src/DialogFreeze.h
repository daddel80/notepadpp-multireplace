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

#pragma once

#include <windows.h>
#include <vector>

// Disables every operable control of a dialog for its lifetime, except one that stays usable
// and takes the keyboard focus. On release only the controls it disabled are enabled again.
class DialogFreeze {
public:
    DialogFreeze(HWND dialog, HWND keep);
    ~DialogFreeze();

    DialogFreeze(const DialogFreeze&) = delete;
    DialogFreeze& operator=(const DialogFreeze&) = delete;

private:
    HWND _dialog;
    HWND _keep;
    bool _keepWasEnabled = false;
    HWND _focus = nullptr;          // focused control before the freeze
    std::vector<HWND> _frozen;
};
