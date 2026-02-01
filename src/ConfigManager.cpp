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
//  Handles user settings stored in "MultiReplace.ini".
//  • wrap IniFileCache for typed read access
//  • simple write helpers modify the in‑memory cache
//  • save() serialises the full cache back to disk (UTF‑8 +BOM)
//  • Differentiates between numeric values (no escaping) and strings (escaped)
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
    _stringKeys.clear();  // Clear on reload
    bool result = _cache.load(iniFile);
    _isLoaded = result;
    return result;
}

void ConfigManager::forceReload(const std::wstring& iniFile)
{
    _isLoaded = false;
    _iniPath.clear();
    _stringKeys.clear();
    load(iniFile);
}

//
//  Save current cache to disk (UTF‑8 with BOM)
//  - String values (user input) are escaped with escapeCsvValue()
//  - Numeric values (int, bool, float, size_t) are written as-is
//
bool ConfigManager::save(const std::wstring& file) const
{
    std::wstring path = file.empty() ? _iniPath : file;
    if (path.empty()) return false;

    std::ofstream out(path, std::ios::binary);
    if (!out.is_open()) return false;

    // UTF-8 BOM
    out.write("\xEF\xBB\xBF", 3);

    const auto& data = _cache.raw();
    for (const auto& secPair : data) {
        const std::wstring& section = secPair.first;
        out << Encoding::wstringToUtf8(L"[" + section + L"]\n");

        for (const auto& kv : secPair.second) {
            std::wstring line;
            std::wstring fullKey = section + L"|" + kv.first;

            // Check if this key was written as a string (needs escaping)
            if (_stringKeys.find(fullKey) != _stringKeys.end()) {
                // String value: escape for proper roundtrip
                line = kv.first + L"=" + StringUtils::escapeCsvValue(kv.second) + L"\n";
            }
            else {
                // Numeric value: write as-is
                line = kv.first + L"=" + kv.second + L"\n";
            }
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

// String values need escaping when saved
void ConfigManager::writeString(const std::wstring& sec, const std::wstring& key,
    const std::wstring& val)
{
    _cache._data[sec][key] = val;
    _stringKeys.insert(sec + L"|" + key);  // Mark as string for escaping
}

// Numeric types don't need escaping
void ConfigManager::writeInt(const std::wstring& sec, const std::wstring& key,
    int val)
{
    _cache._data[sec][key] = std::to_wstring(val);
    // Don't add to _stringKeys - numeric values are not escaped
}

void ConfigManager::writeSizeT(const std::wstring& sec, const std::wstring& key,
    size_t val)
{
    _cache._data[sec][key] = std::to_wstring(val);
    // Don't add to _stringKeys - numeric values are not escaped
}

void ConfigManager::writeBool(const std::wstring& sec, const std::wstring& key,
    bool val)
{
    _cache._data[sec][key] = val ? L"1" : L"0";
    // Don't add to _stringKeys - boolean values are not escaped
}

void ConfigManager::writeFloat(const std::wstring& sec, const std::wstring& key,
    float val)
{
    _cache._data[sec][key] = std::to_wstring(val);
    // Don't add to _stringKeys - numeric values are not escaped
}

void ConfigManager::writeByte(const std::wstring& sec, const std::wstring& key,
    BYTE val)
{
    _cache._data[sec][key] = std::to_wstring(static_cast<int>(val));
    // Don't add to _stringKeys - numeric values are not escaped
}