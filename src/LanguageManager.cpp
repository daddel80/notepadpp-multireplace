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

#include "LanguageManager.h"
#include "Encoding.h"
#include "language_mapping.h"

#include <fstream>
#include <regex>
#include <sstream>
#include <windows.h>
#include <map>

// -----------------------------------------------------------------
// Singleton
// -----------------------------------------------------------------
LanguageManager& LanguageManager::instance()
{
    static LanguageManager mgr;
    return mgr;
}

// -----------------------------------------------------------------
// Public loading helpers
// -----------------------------------------------------------------
bool LanguageManager::load(const std::wstring& pluginDir,
    const std::wstring& nativeLangXmlPath)
{
    std::wstring langCode = detectLanguage(nativeLangXmlPath);

    std::wstring ini = pluginDir;
    if (!ini.empty() && ini.back() != L'\\' && ini.back() != L'/')
        ini += L'\\';
    ini += L"MultiReplace\\languages.ini";

    return loadFromIni(ini, langCode);
}

bool LanguageManager::loadFromIni(const std::wstring& iniFile,
    const std::wstring& languageCode)
{
    // 1) fallback = English
    _table = languageMap;

    if (!_cache.load(iniFile))
        return false;

    const auto& data = _cache.raw();
    // 2) override with requested language
    if (auto it = data.find(languageCode); it != data.end())
        for (const auto& kv : it->second)
            _table[kv.first] = kv.second;

    return true;
}

// -----------------------------------------------------------------
// String getters
// -----------------------------------------------------------------
std::wstring LanguageManager::get(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    auto it = _table.find(id);
    if (it == _table.end()) return L"Text not found";

    std::wstring result = it->second;
    const std::wstring base = L"$REPLACE_STRING";

    // <br/>  →  CRLF
    for (size_t p = result.find(L"<br/>");
        p != std::wstring::npos;
        p = result.find(L"<br/>", p))
        result.replace(p, 5, L"\r\n");

    // numbered placeholders (highest first)
    for (size_t i = repl.size(); i > 0; --i) {
        std::wstring ph = base + std::to_wstring(i);
        for (size_t p = result.find(ph);
            p != std::wstring::npos;
            p = result.find(ph, p))
            result.replace(p, ph.size(), repl[i - 1]);
    }

    // plain $REPLACE_STRING
    for (size_t p = result.find(base);
        p != std::wstring::npos;
        p = result.find(base, p))
        result.replace(p, base.size(),
            repl.empty() ? L"" : repl[0]);

    return result;
}

LPCWSTR LanguageManager::getLPCW(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    static std::map<std::wstring, std::wstring> cache;
    auto& ref = cache[id];
    if (ref.empty())
        ref = get(id, repl);
    return ref.c_str();
}

LPWSTR LanguageManager::getLPW(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    static std::wstring buf;
    buf = get(id, repl);
    return &buf[0];
}

// -----------------------------------------------------------------
// Detect active language from Notepad++ nativeLang.xml
// -----------------------------------------------------------------
std::wstring LanguageManager::detectLanguage(const std::wstring& xmlPath)
{
    std::wifstream file(xmlPath);
    if (!file.is_open()) return L"english";

    std::wregex  rx(L"<Native-Langue .*? filename=\"(.*?)\\.xml\"");
    std::wsmatch m;
    std::wstring line, lang = L"english";

    try {
        while (std::getline(file, line))
            if (std::regex_search(line, m, rx) && m.size() > 1) {
                lang = m[1];
                break;
            }
    }
    catch (...) {
        // keep fallback
    }
    return lang;
}
