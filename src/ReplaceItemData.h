// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ReplaceItemData.h
// The list row data model: one Find/Replace entry plus its option
// flags and metadata. Kept in its own header (no Win32 / panel deps)
// so format/codec code can consume it without pulling in the panel.

#pragma once

#include <cstddef>
#include <functional>
#include <string>

struct ReplaceItemData
{
    size_t id = 0;
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

    bool operator==(const ReplaceItemData& rhs) const {
        return
            isEnabled == rhs.isEnabled &&
            findText == rhs.findText &&
            replaceText == rhs.replaceText &&
            wholeWord == rhs.wholeWord &&
            matchCase == rhs.matchCase &&
            extended == rhs.extended &&
            regex == rhs.regex;
    }

    bool operator!=(const ReplaceItemData& rhs) const {
        return !(*this == rhs);
    }
};

// Hash function for ReplaceItemData
struct ReplaceItemDataHasher {
    std::size_t operator()(const ReplaceItemData& item) const {
        std::size_t hash = std::hash<bool>{}(item.isEnabled);
        hash ^= std::hash<std::wstring>{}(item.findText) << 1;
        hash ^= std::hash<std::wstring>{}(item.replaceText) << 1;
        hash ^= std::hash<bool>{}(item.wholeWord) << 1;
        hash ^= std::hash<bool>{}(item.matchCase) << 1;
        hash ^= std::hash<bool>{}(item.formulaSupport) << 1;
        hash ^= std::hash<bool>{}(item.extended) << 1;
        hash ^= std::hash<bool>{}(item.regex) << 1;
        hash ^= std::hash<std::wstring>{}(item.comments) << 1;
        return hash;
    }
};
