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
//  ConfigManager  (singleton)
//  • Handles user-configurable settings stored in MultiReplace.ini
//  • Uses IniFileCache for parsing / writing
// --------------------------------------------------------------

#include "IniFileCache.h"

#include <string>
#include <windows.h>    // BYTE

class ConfigManager
{
public:
    static ConfigManager& instance();

    // ---------------------------------------------------------------------
    // Load / Save
    // ---------------------------------------------------------------------
    bool load(const std::wstring& iniFile);
    void forceReload(const std::wstring& iniFile);
    bool save(const std::wstring& iniFile) const;
    bool isLoaded() const { return _isLoaded; }

    // ---------------------------------------------------------------------
    // Typed getters
    // ---------------------------------------------------------------------
    std::wstring readString(const std::wstring& sec,
        const std::wstring& key,
        const std::wstring& def = L"") const
    {
        return _cache.readString(sec, key, def);
    }

    bool   readBool(const std::wstring& s, const std::wstring& k, bool   d = false) const
    {
        return _cache.readBool(s, k, d);
    }
    int    readInt(const std::wstring& s, const std::wstring& k, int    d = 0)     const
    {
        return _cache.readInt(s, k, d);
    }
    float  readFloat(const std::wstring& s, const std::wstring& k, float  d = 0.f)   const
    {
        return _cache.readFloat(s, k, d);
    }
    BYTE   readByte(const std::wstring& s, const std::wstring& k, BYTE   d = 0)     const
    {
        return _cache.readByte(s, k, d);
    }
    size_t readSizeT(const std::wstring& s, const std::wstring& k, size_t d = 0)     const
    {
        return _cache.readSizeT(s, k, d);
    }

    // ---------------------------------------------------------------------
    // Typed setters (symmetric to getters)
    // ---------------------------------------------------------------------
    void writeString(const std::wstring& sec, const std::wstring& key, const std::wstring& val);
    void writeBool(const std::wstring& sec, const std::wstring& key, bool val);
    void writeInt(const std::wstring& sec, const std::wstring& key, int val);
    void writeFloat(const std::wstring& sec, const std::wstring& key, float val);
    void writeByte(const std::wstring& sec, const std::wstring& key, BYTE val);
    void writeSizeT(const std::wstring& sec, const std::wstring& key, size_t val);

    // Raw access if absolutely necessary
    const IniFileCache& ini() const { return _cache; }

private:
    ConfigManager() = default;
    ~ConfigManager() = default;
    ConfigManager(const ConfigManager&) = delete;
    ConfigManager& operator=(const ConfigManager&) = delete;

    IniFileCache _cache;
    std::wstring _iniPath;
    bool _isLoaded = false;
};