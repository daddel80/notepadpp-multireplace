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
#include <string_view>
#include <cstddef>

// ============================================================
// Language Key-Value Pairs (English defaults)
// ============================================================
struct LangKV {
    std::wstring_view k;
    std::wstring_view v;
};

extern const LangKV kEnglishPairs[];
extern const size_t kEnglishPairsCount;

// ============================================================
// UI Control Mappings for Language Refresh
// ============================================================

// Control ID -> Language Key (for SetWindowText)
struct UITextMapping {
    int controlId;
    const wchar_t* langKey;
};

// Control ID -> Language Key (for Tooltips - future use)
struct UITooltipMapping {
    int controlId;
    const wchar_t* langKey;
};

// ColumnID -> Language Key (for ListView headers)
// Note: columnId corresponds to ColumnID enum values
struct UIHeaderMapping {
    int columnId;
    const wchar_t* langKey;
};

// ============================================================
// Main Panel Mappings
// ============================================================
extern const UITextMapping kControlTextMappings[];
extern const size_t kControlTextMappingsCount;

extern const UITooltipMapping kTooltipMappings[];
extern const size_t kTooltipMappingsCount;

extern const UIHeaderMapping kHeaderTextMappings[];
extern const size_t kHeaderTextMappingsCount;

extern const UIHeaderMapping kHeaderTooltipMappings[];
extern const size_t kHeaderTooltipMappingsCount;
