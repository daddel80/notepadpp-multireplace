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

#pragma once
#include <string>
#include <Windows.h>

/// Collection of static helper functions for common text‑encoding tasks.

class Encoding final
{
public:
    // raw bytes(document code page) → UTF‑8 std::string
    static std::string bytesToUtf8(const std::string_view src, UINT codePage);
    static std::wstring Encoding::bytesToWString(const std::string& raw, UINT cp);

    // UTF‑8 ⇄ std::wstring
    static std::wstring utf8ToWString(const std::string& utf8);
    static std::string  wstringToUtf8(const std::wstring& ws);

    // ANSI code page ⇄ std::wstring
    static std::wstring ansiToWString(const std::string& input, UINT codePage);
    static std::string  wstringToString(const std::wstring& ws, UINT codePage);

    // utilities
    static std::wstring trim(const std::wstring& str);
    static bool         isValidUtf8(const std::string& data);

private:
    Encoding() = delete;  // static‑only class
    ~Encoding() = delete;
};
