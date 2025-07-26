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
#include "Encoding.h"        // Encoding::wstringToUtf8
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
    _iniPath = iniFile;
    return _cache.load(iniFile);
}

//
//  Save current cache to disk (UTF‑8 with BOM)
//
bool ConfigManager::save(const std::wstring& file) const
{
    std::wstring path = file.empty() ? _iniPath : file;
    if (path.empty()) return false;

    // to UTF‑8 narrow path
    int sz8 = WideCharToMultiByte(CP_UTF8, 0, path.c_str(), (int)path.size(),
        nullptr, 0, nullptr, nullptr);
    std::string narrow(sz8, 0);
    WideCharToMultiByte(CP_UTF8, 0, path.c_str(), (int)path.size(),
        narrow.data(), sz8, nullptr, nullptr);

    std::ofstream out(narrow, std::ios::binary);
    if (!out.is_open()) return false;

    // UTF‑8 BOM
    out.write("\xEF\xBB\xBF", 3);

    const auto& data = _cache.raw();
    for (const auto& secPair : data) {
        const std::wstring& section = secPair.first;
        out << Encoding::wstringToUtf8(L"[" + section + L"]\n");

        for (const auto& kv : secPair.second) {
            std::wstring line = kv.first + L"=" + kv.second + L"\n";
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

// Further writeBool / writeFloat / … could be added similarly
