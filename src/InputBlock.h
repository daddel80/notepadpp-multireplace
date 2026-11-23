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

// A window takes no input while the block lives, like the owner of a modal dialog.
// On release the window is enabled again only if this block disabled it.
class InputBlock {
public:
    explicit InputBlock(HWND window)
        : _window((window && IsWindowEnabled(window)) ? window : nullptr) {
        if (_window) EnableWindow(_window, FALSE);
    }
    ~InputBlock() {
        if (_window && IsWindow(_window)) EnableWindow(_window, TRUE);
    }

    InputBlock(const InputBlock&) = delete;
    InputBlock& operator=(const InputBlock&) = delete;

private:
    HWND _window;
};
