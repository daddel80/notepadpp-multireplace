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
#include <unordered_map>
#include <windows.h>

namespace {
    // --- Local helpers ----------------------------------------------------
    std::wstring makeKey(const std::wstring& id,
        const std::vector<std::wstring>& repl)
    {
        std::wstring key = id;
        key.push_back(L'\x1F');
        for (const auto& r : repl) { key += r; key.push_back(L'\x1F'); }
        return key;
    }

    std::unordered_map<std::wstring, std::wstring>& lpcwCache()
    {
        static std::unordered_map<std::wstring, std::wstring> cache;
        return cache;
    }
}

// --- Singleton ------------------------------------------------------------
LanguageManager& LanguageManager::instance()
{
    static LanguageManager mgr;
    return mgr;
}

// --- Loading --------------------------------------------------------------
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
    // 1) Defaults (English) -> _table
    _table.clear();
    _table.reserve(kEnglishPairsCount);
    for (size_t i = 0; i < kEnglishPairsCount; ++i) {
        _table.emplace(std::wstring(kEnglishPairs[i].k),
            std::wstring(kEnglishPairs[i].v));
    }
    if (!_cache.load(iniFile))
        return false;
    const auto& data = _cache.raw();
    if (auto it = data.find(languageCode); it != data.end())
        for (const auto& kv : it->second)
            _table[kv.first] = kv.second;

    invalidateCaches();
    return true;
}

void LanguageManager::invalidateCaches()
{
    lpcwCache().clear();
}

// --- Strings --------------------------------------------------------------
// <br/> -> CRLF, then $REPLACE_STRINGn (high->low), then $REPLACE_STRING.
std::wstring LanguageManager::get(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    auto it = _table.find(id);
    if (it == _table.end())
        return id;

    std::wstring result = it->second;
    const std::wstring base = L"$REPLACE_STRING";

    // <br/> -> CRLF
    for (size_t p = result.find(L"<br/>");
        p != std::wstring::npos;
        p = result.find(L"<br/>", p))
    {
        result.replace(p, 5, L"\r\n");
        p += 2;
    }

    // $REPLACE_STRINGn
    for (size_t i = repl.size(); i > 0; --i)
    {
        const std::wstring ph = base + std::to_wstring(i);
        const std::wstring& vv = repl[i - 1];

        for (size_t p = result.find(ph);
            p != std::wstring::npos;
            p = result.find(ph, p))
        {
            result.replace(p, ph.size(), vv);
            p += vv.size();
        }
    }

    // $REPLACE_STRING
    {
        const std::wstring& vv = repl.empty() ? std::wstring() : repl[0];

        for (size_t p = result.find(base);
            p != std::wstring::npos;
            p = result.find(base, p))
        {
            result.replace(p, base.size(), vv);
            p += vv.size();
        }
    }

    return result;
}

LPCWSTR LanguageManager::getLPCW(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    auto& cache = lpcwCache();
    const std::wstring key = makeKey(id, repl);

    auto it = cache.find(key);
    if (it == cache.end())
        it = cache.emplace(key, get(id, repl)).first;

    return it->second.c_str();
}

LPWSTR LanguageManager::getLPW(const std::wstring& id,
    const std::vector<std::wstring>& repl) const
{
    thread_local std::wstring buf;
    buf = get(id, repl);
    return buf.empty() ? nullptr : buf.data();
}

// --- nativeLang.xml detection --------------------------------------------
std::wstring LanguageManager::detectLanguage(const std::wstring& xmlPath)
{
    std::wifstream file(xmlPath);
    if (!file.is_open())
        return L"english";

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
    catch (...) { /* keep fallback */ }

    return lang;
}
