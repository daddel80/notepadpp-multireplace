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

namespace StringUtils {

    // Keep the mode used by MultiReplace.
    enum class ReplaceMode { Normal, Extended, Regex };

    // ---- String helpers extracted from MultiReplacePanel ----
    std::wstring sanitizeSearchPattern(const std::wstring& raw);

    // Wrap value in double quotes; escape ", \r, \n, \\ so the result
    // round-trips through unescapeQuoted. Used by the INI cache, the
    // settings block of MR list files, and the body of CSV cells.
    std::wstring escapeQuoted(const std::wstring& value);

    // Inverse of escapeQuoted. Tolerates unquoted input (returns as-is).
    std::wstring unescapeQuoted(const std::wstring& value);

    // Lighter variant: wraps in quotes and doubles up embedded quotes,
    // but does NOT escape \r, \n, \\. Used for editor exports where
    // newlines must remain real newlines.
    std::wstring quoteField(const std::wstring& value);

    bool        normalizeAndValidateNumber(std::string& str);

    std::string escapeForRegex(const std::string& input);

    // Escape special regex/sed characters. In extended mode, preserves valid escape sequences.
    std::string escapeSpecialChars(const std::string& input, bool extended);

    // Translate escape sequences (\n, \r, \t, \xHH, \oOOO, \dDDD, \bBBBBBBBB, \uHHHH) to their values.
    // \n and \r become __NEWLINE__ and __CARRIAGERETURN__ placeholders for bash export.
    std::string translateEscapes(const std::string& input);

    // Replace '\n' / '\r' according to the active mode (Normal / Extended / Regex)
    std::string replaceNewline(const std::string& input, ReplaceMode mode);

    std::wstring trim(const std::wstring& str);

    // Escape control characters for debug display (makes \n, \r, \t visible)
    std::string escapeControlChars(const std::string& input);

    // ---- Unicode-aware string operations ----
    // Convert UTF-8 string to lowercase using Windows locale-aware API
    // Correctly handles: German ß/ẞ, French É→é, Turkish İ→i, etc.
    std::string toLowerUtf8(const std::string& utf8Str);

    // ---- Locale-aware number formatting ----
    // Format number with locale-specific thousand separator (e.g., 1,234 or 1.234)
    std::wstring formatNumber(size_t number);

    // ---- Template helpers (moved from MultiReplacePanel) ----
    // Replace all occurrences of a placeholder variable in a template string
    std::wstring replaceTemplateVar(const std::wstring& tmpl, const std::wstring& var, const std::wstring& value);

    // Translate simple escape sequences (\t, \n, \r, \\) in a template string
    std::wstring processTemplateEscapes(const std::wstring& tmpl);

} // namespace StringUtils