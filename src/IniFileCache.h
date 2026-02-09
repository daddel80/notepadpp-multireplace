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
//  • Tracks which keys were stored as quoted strings in the INI
//  • Offers typed getters (string, int, bool …)
// --------------------------------------------------------------

#include <string>
#include <set>
#include <unordered_map>
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

    // Keys that were quoted (i.e. string values) in the parsed INI file.
    // Format: "Section|Key". Used by ConfigManager to preserve escaping.
    const std::set<std::wstring>& quotedKeys() const { return _quotedKeys; }

    // Direct access (rarely needed) ---------------------------------------
    using Section = std::unordered_map<std::wstring, std::wstring>;
    const std::unordered_map<std::wstring, Section>& raw() const { return _data; }

    friend class ConfigManager;

private:
    bool parse(const std::wstring& iniFile);   // core parser

    using IniData = std::unordered_map<std::wstring, Section>;
    IniData _data;
    std::set<std::wstring> _quotedKeys;         // "Section|Key" for quoted values
};