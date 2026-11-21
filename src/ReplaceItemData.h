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
//
// ReplaceItemData.h
// The list row data model: one Find/Replace entry plus its option
// flags and metadata. Kept in its own header (no Win32 / panel deps)
// so format/codec code can consume it without pulling in the panel.

#pragma once

#include <cstddef>
#include <string>

struct ReplaceItemData
{
    size_t id = 0;                  // Session-only: given when the entry comes into a list and kept there; 0 outside a list
    int findCount = -1;
    int replaceCount = -1;
    bool isEnabled = true;
    std::wstring findText;
    std::wstring replaceText;
    bool wholeWord = false;
    bool matchCase = false;
    bool formulaSupport = false;
    bool extended = false;
    bool regex = false;
    std::wstring comments = L"";
    bool isDirty = false;           // Session-only: marks row as modified since last save/load
    std::wstring lastModified;      // Persistent timestamp: set on content change, saved in CSV

    // Same rule: enabled state, comment and bookkeeping do not make an entry different
    bool isDuplicateOf(const ReplaceItemData& rhs) const {
        return
            findText == rhs.findText &&
            replaceText == rhs.replaceText &&
            wholeWord == rhs.wholeWord &&
            matchCase == rhs.matchCase &&
            formulaSupport == rhs.formulaSupport &&
            extended == rhs.extended &&
            regex == rhs.regex;
    }
};
