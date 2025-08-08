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
//  LanguageManager  (singleton)
//  • Detects active UI language via nativeLang.xml
//  • Loads the matching section from languages.ini
//  • Supplies translated strings with placeholder replacement
// --------------------------------------------------------------

#include "IniFileCache.h"
#include <string>
#include <vector>
#include <unordered_map>
#include <windows.h>

class LanguageManager
{
public:
    // --- Singleton --------------------------------------------------------
    static LanguageManager& instance();

    // --- Loading ----------------------------------------------------------
    bool load(const std::wstring& pluginDir, const std::wstring& nativeLangXml);
    bool loadFromIni(const std::wstring& iniFile, const std::wstring& languageCode);

    // --- Strings ----------------------------------------------------------
    std::wstring get(const std::wstring& id,
        const std::vector<std::wstring>& repl = {}) const;
    LPCWSTR      getLPCW(const std::wstring& id,
        const std::vector<std::wstring>& repl = {}) const;
    LPWSTR       getLPW(const std::wstring& id,
        const std::vector<std::wstring>& repl = {}) const;

    const IniFileCache& ini() const { return _cache; }

private:
    LanguageManager() = default;
    ~LanguageManager() = default;
    LanguageManager(const LanguageManager&) = delete;
    LanguageManager& operator=(const LanguageManager&) = delete;

    static std::wstring detectLanguage(const std::wstring& nativeLangXmlPath);
    void invalidateCaches();

    IniFileCache _cache;                                      // languages.ini
    std::unordered_map<std::wstring, std::wstring> _table;    // id -> text
};