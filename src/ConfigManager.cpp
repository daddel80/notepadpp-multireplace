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
#include <vector>
#include <algorithm>
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

    // Clear and rebuild _stringKeys from the INI: the parser records
    // every quoted value as "Section|Key" directly into _stringKeys.
    // Later writeString() calls add new keys additively.
    _stringKeys.clear();
    _cache._data.clear();
    bool result = _cache.parse(iniFile, &_stringKeys);

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
//  - String values (user input) are escaped with escapeQuoted()
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

    // Write a single section (header + keys + trailing blank line). The
    // string-key escaping is identical to the legacy path; kept in one place.
    auto writeSection = [&](const std::wstring& section,
        const IniFileCache::Section& keys)
        {
            out << Encoding::wstringToUtf8(L"[" + section + L"]\n");
            for (const auto& kv : keys) {
                const std::wstring fullKey = section + L"|" + kv.first;
                std::wstring line;
                if (_stringKeys.find(fullKey) != _stringKeys.end()) {
                    line = kv.first + L"=" + StringUtils::escapeQuoted(kv.second) + L"\n";
                }
                else {
                    line = kv.first + L"=" + kv.second + L"\n";
                }
                out << Encoding::wstringToUtf8(line);
            }
            out << '\n';
        };

    // Preferred section order: editable settings first (grouped by function),
    // machine state last. Any section not listed here is written afterwards
    // in alphabetical order, so nothing is ever dropped.
    static const wchar_t* const kSectionOrder[] = {
        L"General",
        L"SearchAndReplace", L"Interface", L"ReplaceInFiles", L"Engines", L"Lua",
        L"Scope", L"Csv", L"List", L"ListView", L"ResultDock",
        L"Export", L"Appearance", L"Tandem",
        L"Window", L"Tabs", L"History",
    };

    std::vector<std::wstring> written;
    for (const wchar_t* sec : kSectionOrder) {
        auto it = data.find(sec);
        if (it != data.end()) {
            writeSection(it->first, it->second);
            written.push_back(it->first);
        }
    }
    // Safety net: emit any remaining sections (unlisted / future) so the
    // ordering change can never silently lose data.
    for (const auto& secPair : data) {
        if (std::find(written.begin(), written.end(), secPair.first) == written.end()) {
            writeSection(secPair.first, secPair.second);
        }
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

//
//  Erase helpers — remove entries from the cache so save() does not
//  serialize them on the next shutdown. Used for one-time migration
//  cleanup of legacy INI keys whose values have already been consumed
//  by the migration code path.
//
void ConfigManager::eraseKey(const std::wstring& sec, const std::wstring& key)
{
    auto it = _cache._data.find(sec);
    if (it == _cache._data.end()) return;
    it->second.erase(key);
    if (it->second.empty()) {
        _cache._data.erase(it);
    }
    _stringKeys.erase(sec + L"|" + key);
}

void ConfigManager::eraseSection(const std::wstring& sec)
{
    _cache._data.erase(sec);
    // Drop any _stringKeys entries that belonged to this section so
    // the tracking set stays consistent with _data.
    const std::wstring prefix = sec + L"|";
    for (auto it = _stringKeys.begin(); it != _stringKeys.end(); ) {
        if (it->compare(0, prefix.size(), prefix) == 0) {
            it = _stringKeys.erase(it);
        }
        else {
            ++it;
        }
    }
}

bool ConfigManager::hasKey(const std::wstring& sec, const std::wstring& key) const
{
    auto it = _cache._data.find(sec);
    if (it == _cache._data.end()) return false;
    return it->second.find(key) != it->second.end();
}

void ConfigManager::moveKey(const std::wstring& srcSec, const std::wstring& key,
    const std::wstring& dstSec)
{
    if (srcSec == dstSec) return;
    auto sIt = _cache._data.find(srcSec);
    if (sIt == _cache._data.end()) return;
    auto kIt = sIt->second.find(key);
    if (kIt == sIt->second.end()) return;

    _cache._data[dstSec][key] = kIt->second;

    const std::wstring srcFull = srcSec + L"|" + key;
    if (_stringKeys.find(srcFull) != _stringKeys.end()) {
        _stringKeys.insert(dstSec + L"|" + key);
        _stringKeys.erase(srcFull);
    }

    sIt->second.erase(kIt);
    if (sIt->second.empty()) {
        _cache._data.erase(sIt);
    }
}