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

#include "IniFileCache.h"
#include "Encoding.h"               // Encoding::trim / isValidUtf8
#include <fstream>
#include <sstream>
#include <regex>
#include <limits>
#include <algorithm>
#include <codecvt>
#include <windows.h>

namespace {

    // unescapeCsvValue – unchanged from original code ------------------
    std::wstring unescapeCsvValue(const std::wstring& value)
    {
        std::wstring out;
        if (value.empty()) return out;

        size_t start = (value.front() == L'"' && value.back() == L'"') ? 1 : 0;
        size_t end = (start == 1) ? value.size() - 1 : value.size();

        for (size_t i = start; i < end; ++i) {
            if (i < end - 1 && value[i] == L'\\') {
                switch (value[i + 1]) {
                case L'n': out += L'\n'; ++i; break;
                case L'r': out += L'\r'; ++i; break;
                case L'\\': out += L'\\'; ++i; break;
                default:   out += value[i];   break;
                }
            }
            else if (i < end - 1 && value[i] == L'"' && value[i + 1] == L'"') {
                out += L'"'; ++i;
            }
            else {
                out += value[i];
            }
        }
        return out;
    }

} // anonymous namespace



// ------------------------------------------------------------------
//  Public loader
// ------------------------------------------------------------------
bool IniFileCache::load(const std::wstring& iniFile)
{
    _data.clear();
    return parse(iniFile);
}



// ------------------------------------------------------------------
//  Parser (direct copy of old logic)
// ------------------------------------------------------------------
bool IniFileCache::parse(const std::wstring& iniFilePath)
{
    // convert wide → UTF‑8 for std::ifstream
    int sz8 = WideCharToMultiByte(CP_UTF8, 0,
        iniFilePath.c_str(), (int)iniFilePath.size(),
        nullptr, 0, nullptr, nullptr);
    std::string filePath(sz8, 0);
    WideCharToMultiByte(CP_UTF8, 0,
        iniFilePath.c_str(), (int)iniFilePath.size(),
        &filePath[0], sz8, nullptr, nullptr);

    std::ifstream ini(filePath, std::ios::binary);
    if (!ini.is_open()) return false;

    // read whole file
    std::string raw((std::istreambuf_iterator<char>(ini)),
        std::istreambuf_iterator<char>());
    ini.close();

    // UTF‑8 / ANSI detection (unchanged)
    size_t offset = 0;
    UINT cp = CP_UTF8;
    if (raw.size() >= 3 &&
        (unsigned char)raw[0] == 0xEF &&
        (unsigned char)raw[1] == 0xBB &&
        (unsigned char)raw[2] == 0xBF)
    {
        offset = 3;
    }
    else if (!Encoding::isValidUtf8(raw)) {
        cp = CP_ACP;
    }

    // convert to wide
    int wlen = MultiByteToWideChar(cp, 0,
        raw.c_str() + offset,
        (int)raw.size() - (int)offset,
        nullptr, 0);
    std::wstring content(wlen, 0);
    MultiByteToWideChar(cp, 0,
        raw.c_str() + offset,
        (int)raw.size() - (int)offset,
        &content[0], wlen);

    std::wstringstream ss(content);
    std::wstring line, section;

    while (std::getline(ss, line)) {
        line = Encoding::trim(line);
        if (line.empty()) continue;

        wchar_t first = line[0];
        if (first == L';' || first == L'#') continue;

        if (first == L'[') {
            size_t close = line.find(L']');
            if (close != std::wstring::npos) {
                section = Encoding::trim(line.substr(1, close - 1));
            }
            continue;
        }

        size_t eq = line.find(L'=');
        if (eq == std::wstring::npos) continue;

        std::wstring key = Encoding::trim(line.substr(0, eq));
        std::wstring value = Encoding::trim(line.substr(eq + 1));
        _data[section][key] = unescapeCsvValue(value);
    }
    return true;
}



// ------------------------------------------------------------------
//  Typed getters  (identical logic, macro‑safe max() usage)
// ------------------------------------------------------------------
std::wstring IniFileCache::readString(const std::wstring& s, const std::wstring& k,
    const std::wstring& d) const
{
    auto sIt = _data.find(s);
    if (sIt == _data.end()) return d;
    auto kIt = sIt->second.find(k);
    return (kIt == sIt->second.end()) ? d : kIt->second;
}

bool IniFileCache::readBool(const std::wstring& s, const std::wstring& k, bool d) const
{
    std::wstring def = d ? L"1" : L"0";
    std::wstring v = readString(s, k, def);
    if (v == L"1" || _wcsicmp(v.c_str(), L"true") == 0) return true;
    if (v == L"0" || _wcsicmp(v.c_str(), L"false") == 0) return false;
    return d;
}

int IniFileCache::readInt(const std::wstring& s, const std::wstring& k, int d) const
{
    try { return std::stoi(readString(s, k, std::to_wstring(d))); }
    catch (...) { return d; }
}

float IniFileCache::readFloat(const std::wstring& s, const std::wstring& k, float d) const
{
    try { return std::stof(readString(s, k, std::to_wstring(d))); }
    catch (...) { return d; }
}

BYTE IniFileCache::readByte(const std::wstring& s, const std::wstring& k, BYTE d) const
{
    int v = readInt(s, k, d);
    return static_cast<BYTE>(std::clamp(v, 0, 255));
}

size_t IniFileCache::readSizeT(const std::wstring& s, const std::wstring& k, size_t d) const
{
    try {
        unsigned long long v = std::stoull(readString(s, k, std::to_wstring(d)));
        if (v > (std::numeric_limits<size_t>::max)()) return d;
        return static_cast<size_t>(v);
    }
    catch (...) { return d; }
}
