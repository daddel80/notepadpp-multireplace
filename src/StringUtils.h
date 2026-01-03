// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025 Thomas Knoefel
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

#include <string>
#include <vector>
#include <functional>
#include <regex>

namespace StringUtils {

    // Keep the mode used by MultiReplace.
    enum class ReplaceMode { Normal, Extended, Regex };

    // ---- String helpers extracted from MultiReplacePanel ----
    std::wstring sanitizeSearchPattern(const std::wstring& raw);

    std::wstring                 escapeCsvValue(const std::wstring& value);
    std::wstring                 unescapeCsvValue(const std::wstring& value);
    std::vector<std::wstring>    parseCsvLine(const std::wstring& line);

    bool        normalizeAndValidateNumber(std::string& str);

    std::string escapeForRegex(const std::string& input);

    std::string escapeSpecialChars(const std::string& input, bool extended);
    void        handleEscapeSequence(const std::regex& rx,
        const std::string& in,
        std::string& out,
        std::function<char(const std::string&)> conv);
    std::string translateEscapes(const std::string& input);

    // Replace '\n' / '\r' according to the active mode (Normal / Extended / Regex)
    std::string replaceNewline(const std::string& input, ReplaceMode mode);

    std::wstring trim(const std::wstring& str);

    // Escape control characters for debug display (makes \n, \r, \t visible)
    std::string escapeControlChars(const std::string& input);

} // namespace StringUtils