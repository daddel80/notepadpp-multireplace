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

// -----------------------------------------------------------------------------
//  Handles user settings stored in “MultiReplace.ini”.
//  • wrap IniFileCache for typed read access
//  • simple write helpers modify the in‑memory cache
//  • save() serialises the full cache back to disk (UTF‑8 +BOM)
// -----------------------------------------------------------------------------

#include "ConfigManager.h"
#include "Encoding.h"  
#include "StringUtils.h" 
#include <fstream>
#include <sstream>
#include <windows.h>

//
//  Singleton access
//
ConfigManager& ConfigManager::instance()
{
    static ConfigManager mgr;
    return mgr;
}

//
//  Load settings from file
//
bool ConfigManager::load(const std::wstring& iniFile)
{
    // Skip if already loaded with same path
    if (_isLoaded && !_iniPath.empty() && _iniPath == iniFile) {
        return true;
    }

    _iniPath = iniFile;
    bool result = _cache.load(iniFile);
    _isLoaded = result;
    return result;
}

void ConfigManager::forceReload(const std::wstring& iniFile)
{
    _isLoaded = false;
    _iniPath.clear();
    load(iniFile);
}

//
//  Save current cache to disk (UTF‑8 with BOM)
//  unescapeCsvValue is applied in IniFileCache::parse() when loading.
//
bool ConfigManager::save(const std::wstring& file) const
{
    std::wstring path = file.empty() ? _iniPath : file;
    if (path.empty()) return false;

    // Use wstring path directly (Windows MSVC supports this)
    std::ofstream out(path, std::ios::binary);
    if (!out.is_open()) return false;

    // UTF-8 BOM
    out.write("\xEF\xBB\xBF", 3);

    const auto& data = _cache.raw();
    for (const auto& secPair : data) {
        const std::wstring& section = secPair.first;
        out << Encoding::wstringToUtf8(L"[" + section + L"]\n");

        for (const auto& kv : secPair.second) {
            // Escape value for proper roundtrip (unescaped on load by IniFileCache::parse)
            std::wstring line = kv.first + L"=" + StringUtils::escapeCsvValue(kv.second) + L"\n";
            out << Encoding::wstringToUtf8(line);
        }
        out << '\n';
    }
    out.close();
    return !out.fail();
}

//
//  Write helpers (update cache only – caller calls save() later)
//
void ConfigManager::writeString(const std::wstring& sec, const std::wstring& key,
    const std::wstring& val)
{
    _cache._data[sec][key] = val;   // direct because ConfigManager is friend
}

void ConfigManager::writeInt(const std::wstring& sec, const std::wstring& key,
    int val)
{
    writeString(sec, key, std::to_wstring(val));
}

void ConfigManager::writeSizeT(const std::wstring& sec, const std::wstring& key,
    size_t val)
{
    writeString(sec, key, std::to_wstring(val));
}

void ConfigManager::writeBool(const std::wstring& sec, const std::wstring& key,
    bool val)
{
    writeString(sec, key, val ? L"1" : L"0");
}
