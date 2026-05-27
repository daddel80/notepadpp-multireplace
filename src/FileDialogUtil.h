// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025 Thomas Knoefel
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

#include <string>
#include <vector>
#include <windows.h>

// Thin wrapper around the modern IFileDialog COM API (Vista+) for the
// Save/Open dialogs in MultiReplace. The motivation over GetSaveFileName /
// GetOpenFileName is OnTypeChange: when the user switches the file-type
// dropdown the visible filename's extension follows the filter, matching
// the behavior of Notepad++ itself (see CustomFileDialog.cpp in the npp
// source tree, same mechanism).
//
// Scope is deliberately minimal: single-file Save / Open with a filter
// list, optional default filename and default extension, and the 0-based
// index of the filter the user picked on OK. Multi-select, customized
// controls, and folder dialogs are out of scope.

namespace FileDialogUtil {

    struct Filter {
        std::wstring name;  // e.g. L"MultiReplace List (*.mrl)"
        std::wstring ext;   // e.g. L"*.mrl"  or  L"*.mrl;*.csv"
    };

    struct Params {
        HWND                owner = nullptr;
        std::wstring        title;
        std::wstring        defaultPath;          // initial file name (with or without folder)
        std::vector<Filter> filters;
        int                 initialFilterIndex = 0;  // 0-based; clamped if out of range
        std::wstring        defaultExtension;     // without leading dot; appended when the user omits one
        bool                pathMustExist = true;
        bool                fileMustExist = false;   // typically true for Open, false for Save
    };

    struct Result {
        bool         ok = false;
        std::wstring path;
        int          filterIndex = 0;  // 0-based; matches the filter the user had selected on OK
    };

    Result showSave(const Params& p);
    Result showOpen(const Params& p);

} // namespace FileDialogUtil