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
// --------------------------------------------------------------
//  IniFileCache
//  • Parses a Windows‑style INI file (UTF‑8, UTF‑8‑BOM, or ANSI)
//  • Stores key/value pairs hierarchically in memory
//  • Offers typed getters (string, int, bool …)
// --------------------------------------------------------------

#include <string>
#include <set>
#include <map>
#include <windows.h>    // BYTE
#include "StringUtils.h"

class IniFileCache
{
public:
    // Load/replace the current cache; returns true on success
    bool load(const std::wstring& iniFile);

    // Typed getters --------------------------------------------------------
    std::wstring readString(const std::wstring& section,
        const std::wstring& key,
        const std::wstring& def = L"") const;

    bool   readBool(const std::wstring& section,
        const std::wstring& key,
        bool   def = false) const;

    int    readInt(const std::wstring& section,
        const std::wstring& key,
        int    def = 0) const;

    float  readFloat(const std::wstring& section,
        const std::wstring& key,
        float  def = 0.f) const;

    BYTE   readByte(const std::wstring& section,
        const std::wstring& key,
        BYTE   def = 0) const;

    size_t readSizeT(const std::wstring& section,
        const std::wstring& key,
        size_t def = 0) const;

    // Direct access (rarely needed) ---------------------------------------
    // Ordered map so keys within a section are written in a stable, alphabetical
    // order on save (instead of the hash-dependent order of unordered_map).
    using Section = std::map<std::wstring, std::wstring>;
    const std::map<std::wstring, Section>& raw() const { return _data; }

    friend class ConfigManager;

private:
    // Core parser. If stringKeys is provided, keys whose values were
    // quoted in the INI file are recorded as "Section|Key" so that
    // ConfigManager::save() can re-escape them correctly.
    bool parse(const std::wstring& iniFile, std::set<std::wstring>* stringKeys = nullptr);

    using IniData = std::map<std::wstring, Section>;
    IniData _data;
};