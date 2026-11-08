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

// Plugin version, shared by the version resource and the About dialog
#define VER_MAJOR   6
#define VER_MINOR   3
#define VER_RELEASE 0
#define VER_BUILD   39

#define VER_STR_HELPER(x) #x
#define VER_STR(x)        VER_STR_HELPER(x)
#define VER_STRING        VER_STR(VER_MAJOR) "." VER_STR(VER_MINOR) "." VER_STR(VER_RELEASE) "." VER_STR(VER_BUILD)
#define VER_WSTRING       L"" VER_STRING
