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

#include "StringUtils.h"

#include <windows.h>    // MultiByteToWideChar, WideCharToMultiByte, CharLowerBuffW, GetLocaleInfoW
#include <algorithm>
#include <cctype>       // std::isdigit, std::isxdigit
#include <cstdlib>      // std::strtol, std::atoi

namespace StringUtils {

    // ----------------------------------------------------------------------------
    // Find-All header text sanitizing
    // ----------------------------------------------------------------------------
    std::wstring sanitizeSearchPattern(const std::wstring& raw) {
        std::wstring escaped = raw;
        size_t pos = 0;
        while ((pos = escaped.find(L"\r", pos)) != std::wstring::npos) {
            escaped.replace(pos, 1, L"\\r");
            pos += 2;
        }
        pos = 0;
        while ((pos = escaped.find(L"\n", pos)) != std::wstring::npos) {
            escaped.replace(pos, 1, L"\\n");
            pos += 2;
        }
        return escaped;
    }

    // ----------------------------------------------------------------------------
    // CSV helpers
    // ----------------------------------------------------------------------------
    std::wstring escapeCsvValue(const std::wstring& value) {
        std::wstring out = L"\"";
        for (wchar_t ch : value) {
            switch (ch) {
            case L'"':  out += L"\"\""; break;
            case L'\n': out += L"\\n";  break;
            case L'\r': out += L"\\r";  break;
            case L'\\': out += L"\\\\"; break;
            default:    out += ch;      break;
            }
        }
        out += L"\"";
        return out;
    }

    std::wstring unescapeCsvValue(const std::wstring& value) {
        std::wstring out;
        if (value.empty()) return out;

        const bool quoted = (value.front() == L'"' && value.back() == L'"');
        size_t start = quoted ? 1 : 0;
        size_t end = quoted ? value.size() - 1 : value.size();

        for (size_t i = start; i < end; ++i) {
            if (i < end - 1 && value[i] == L'\\') {
                switch (value[i + 1]) {
                case L'n':  out += L'\n'; ++i; break;
                case L'r':  out += L'\r'; ++i; break;
                case L'\\': out += L'\\'; ++i; break;
                default:    out += value[i];   break;
                }
            }
            else if (quoted && i < end - 1 && value[i] == L'"' && value[i + 1] == L'"') {
                out += L'"';
                ++i;
            }
            else {
                out += value[i];
            }
        }
        return out;
    }

    std::vector<std::wstring> parseCsvLine(const std::wstring& line) {
        // Remove trailing line ending characters (handles CRLF from Windows)
        std::wstring cleanLine = line;
        while (!cleanLine.empty() && (cleanLine.back() == L'\r' || cleanLine.back() == L'\n')) {
            cleanLine.pop_back();
        }

        std::vector<std::wstring> columns;
        std::wstring current;
        bool insideQuotes = false;

        for (size_t i = 0; i < cleanLine.size(); ++i) {
            wchar_t ch = cleanLine[i];
            if (ch == L'"') {
                if (insideQuotes && i + 1 < cleanLine.size() && cleanLine[i + 1] == L'"') {
                    current += L'"';
                    ++i;
                }
                else {
                    insideQuotes = !insideQuotes;
                }
            }
            else if (ch == L',' && !insideQuotes) {
                columns.push_back(unescapeCsvValue(current));
                current.clear();
            }
            else {
                current += ch;
            }
        }
        columns.push_back(unescapeCsvValue(current));
        return columns;
    }

    // ----------------------------------------------------------------------------
    // Number normalization (used where numeric input is accepted)
    // ----------------------------------------------------------------------------
    bool normalizeAndValidateNumber(std::string& str) {
        if (str.empty()) return false;
        if (str == "." || str == ",") return false;

        int dotCount = 0;
        std::string tmp = str;

        for (char& c : tmp) {
            if (c == '.') {
                ++dotCount;
            }
            else if (c == ',') {
                ++dotCount;
                c = '.';
            }
            else if (!std::isdigit(static_cast<unsigned char>(c))) {
                return false;
            }
            if (dotCount > 1) return false;
        }

        str = tmp;
        return true;
    }

    // ----------------------------------------------------------------------------
    // Regex escaping (plain text → safe regex)
    // ----------------------------------------------------------------------------
    std::string escapeForRegex(const std::string& input) {
        std::string out;
        out.reserve(input.size() * 2);
        for (char c : input) {
            switch (c) {
            case '\\':
            case '^': case '$': case '.': case '|':
            case '?': case '*': case '+': case '(': case ')':
            case '[': case ']': case '{': case '}':
                out.push_back('\\');
                out.push_back(c);
                break;
            default:
                out.push_back(c);
            }
        }
        return out;
    }

    // ----------------------------------------------------------------------------
    // Escape translation helpers
    // ----------------------------------------------------------------------------

    // Process all escape sequences in a single pass (efficient, handles duplicates correctly)
    std::string translateEscapes(const std::string& input) {
        if (input.empty()) return input;

        std::string output;
        output.reserve(input.size());

        size_t i = 0;
        while (i < input.size()) {
            // Check for backslash escape sequence
            if (input[i] == '\\' && i + 1 < input.size()) {
                char next = input[i + 1];

                // \n -> placeholder
                if (next == 'n') {
                    output += "__NEWLINE__";
                    i += 2;
                    continue;
                }
                // \r -> placeholder
                if (next == 'r') {
                    output += "__CARRIAGERETURN__";
                    i += 2;
                    continue;
                }
                // \t -> tab
                if (next == 't') {
                    output += '\t';
                    i += 2;
                    continue;
                }
                // \0 -> skip (not supported)
                if (next == '0' && (i + 2 >= input.size() || !std::isdigit(static_cast<unsigned char>(input[i + 2])))) {
                    i += 2;
                    continue;
                }
                // \xHH -> hex byte
                if (next == 'x' && i + 3 < input.size()) {
                    char h1 = input[i + 2], h2 = input[i + 3];
                    if (std::isxdigit(static_cast<unsigned char>(h1)) &&
                        std::isxdigit(static_cast<unsigned char>(h2))) {
                        char hex[3] = { h1, h2, '\0' };
                        output += static_cast<char>(std::strtol(hex, nullptr, 16));
                        i += 4;
                        continue;
                    }
                }
                // \oOOO -> octal byte (3 digits)
                if (next == 'o' && i + 4 < input.size()) {
                    bool valid = true;
                    for (int j = 0; j < 3 && valid; ++j) {
                        char c = input[i + 2 + j];
                        if (c < '0' || c > '7') valid = false;
                    }
                    if (valid) {
                        char oct[4] = { input[i + 2], input[i + 3], input[i + 4], '\0' };
                        output += static_cast<char>(std::strtol(oct, nullptr, 8));
                        i += 5;
                        continue;
                    }
                }
                // \dDDD -> decimal byte (3 digits)
                if (next == 'd' && i + 4 < input.size()) {
                    bool valid = true;
                    for (int j = 0; j < 3 && valid; ++j) {
                        if (!std::isdigit(static_cast<unsigned char>(input[i + 2 + j]))) valid = false;
                    }
                    if (valid) {
                        char dec[4] = { input[i + 2], input[i + 3], input[i + 4], '\0' };
                        int val = std::atoi(dec);
                        if (val <= 255) {
                            output += static_cast<char>(val);
                            i += 5;
                            continue;
                        }
                    }
                }
                // \bBBBBBBBB -> binary byte (8 digits)
                if (next == 'b' && i + 9 < input.size()) {
                    bool valid = true;
                    for (int j = 0; j < 8 && valid; ++j) {
                        char c = input[i + 2 + j];
                        if (c != '0' && c != '1') valid = false;
                    }
                    if (valid) {
                        unsigned char val = 0;
                        for (int j = 0; j < 8; ++j) {
                            val = (val << 1) | (input[i + 2 + j] - '0');
                        }
                        output += static_cast<char>(val);
                        i += 10;
                        continue;
                    }
                }
                // \uHHHH -> Unicode codepoint (4 hex digits) -> UTF-8
                if (next == 'u' && i + 5 < input.size()) {
                    bool valid = true;
                    for (int j = 0; j < 4 && valid; ++j) {
                        if (!std::isxdigit(static_cast<unsigned char>(input[i + 2 + j]))) valid = false;
                    }
                    if (valid) {
                        char hex[5] = { input[i + 2], input[i + 3], input[i + 4], input[i + 5], '\0' };
                        int codepoint = static_cast<int>(std::strtol(hex, nullptr, 16));
                        // Convert to UTF-8
                        if (codepoint < 0x80) {
                            output += static_cast<char>(codepoint);
                        }
                        else if (codepoint < 0x800) {
                            output += static_cast<char>(0xC0 | (codepoint >> 6));
                            output += static_cast<char>(0x80 | (codepoint & 0x3F));
                        }
                        else {
                            output += static_cast<char>(0xE0 | (codepoint >> 12));
                            output += static_cast<char>(0x80 | ((codepoint >> 6) & 0x3F));
                            output += static_cast<char>(0x80 | (codepoint & 0x3F));
                        }
                        i += 6;
                        continue;
                    }
                }
            }

            // No escape sequence matched - copy character as-is
            output += input[i];
            ++i;
        }

        return output;
    }

    // Escape special regex/sed characters in a string (single-pass, efficient)
    std::string escapeSpecialChars(const std::string& input, bool extended) {
        if (input.empty()) return input;

        // Characters that need escaping for sed/regex
        static const std::string specials = "$.*[]^&\\{}()?+|<>\"'`~;#";
        // Escape sequences to preserve in extended mode
        static const std::string extendedEscapes = "nrt0xubd";

        std::string output;
        output.reserve(input.size() * 2);  // Worst case: every char escaped

        for (size_t i = 0; i < input.size(); ++i) {
            char c = input[i];

            if (specials.find(c) != std::string::npos) {
                // Special handling for backslash in extended mode
                if (c == '\\' && extended && i + 1 < input.size()) {
                    char next = input[i + 1];
                    if (extendedEscapes.find(next) != std::string::npos) {
                        // Preserve escape sequence (don't escape the backslash)
                        output += c;
                        continue;
                    }
                }
                // Escape the special character
                output += '\\';
            }
            output += c;
        }

        return output;
    }

    // ----------------------------------------------------------------------------
    // Replace newlines according to ReplaceMode
    // ----------------------------------------------------------------------------
    std::string replaceNewline(const std::string& input, ReplaceMode mode) {
        std::string result = input;

        if (mode == ReplaceMode::Normal) {
            result.erase(std::remove(result.begin(), result.end(), '\n'), result.end());
            result.erase(std::remove(result.begin(), result.end(), '\r'), result.end());
        }
        else if (mode == ReplaceMode::Extended) {
            // Replace \n and \r with placeholders (no regex needed)
            for (size_t pos = 0; (pos = result.find('\n', pos)) != std::string::npos; pos += 11)
                result.replace(pos, 1, "__NEWLINE__");
            for (size_t pos = 0; (pos = result.find('\r', pos)) != std::string::npos; pos += 18)
                result.replace(pos, 1, "__CARRIAGERETURN__");
        }
        else if (mode == ReplaceMode::Regex) {
            // Replace \n and \r with escape sequences (no regex needed)
            for (size_t pos = 0; (pos = result.find('\n', pos)) != std::string::npos; pos += 2)
                result.replace(pos, 1, "\\n");
            for (size_t pos = 0; (pos = result.find('\r', pos)) != std::string::npos; pos += 2)
                result.replace(pos, 1, "\\r");
        }
        return result;
    }

    // ----------------------------------------------------------------------------
    // trim leading/trailing whitespace & line breaks
    // ----------------------------------------------------------------------------
    std::wstring trim(const std::wstring& str)
    {
        const auto first = str.find_first_not_of(L" \t\r\n");
        if (first == std::wstring::npos)
            return L"";

        const auto last = str.find_last_not_of(L" \t\r\n");
        return str.substr(first, last - first + 1);
    }

    // ----------------------------------------------------------------------------
    // Escape control characters for debug display (makes \n, \r, \t visible)
    // ----------------------------------------------------------------------------
    std::string escapeControlChars(const std::string& input) {
        std::string result;
        result.reserve(input.size() * 2);

        for (unsigned char ch : input) {
            switch (ch) {
            case '\n': result += "\\n"; break;
            case '\r': result += "\\r"; break;
            case '\t': result += "\\t"; break;
            case '\0': result += "\\0"; break;
            default:
                if (ch < 32) {
                    char buf[8];
                    snprintf(buf, sizeof(buf), "\\x%02X", ch);
                    result += buf;
                }
                else {
                    result += static_cast<char>(ch);
                }
            }
        }
        return result;
    }

    // ----------------------------------------------------------------------------
    // Wrap field in quotes and escape inner quotes (" -> "")
    // ----------------------------------------------------------------------------
    std::wstring quoteField(const std::wstring& value) {
        std::wstring out;
        out.reserve(value.size() + 4);
        out += L'"';

        for (wchar_t ch : value) {
            if (ch == L'"') {
                out += L"\"\"";
            }
            else {
                out += ch;
            }
        }

        out += L'"';
        return out;
    }

    // ----------------------------------------------------------------------------
    // Unicode-aware lowercase conversion using Windows API
    // Correctly handles all Unicode characters including:
    // - German: Ä→ä, Ö→ö, Ü→ü, ẞ→ß
    // - French: É→é, È→è, Ê→ê
    // - Turkish: İ→i, I→ı (with correct locale)
    // - Greek, Cyrillic, etc.
    // ----------------------------------------------------------------------------
    std::string toLowerUtf8(const std::string& utf8Str)
    {
        if (utf8Str.empty()) return utf8Str;

        // Convert UTF-8 to Wide String using Windows API
        int wideLen = MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(),
            static_cast<int>(utf8Str.length()), nullptr, 0);
        if (wideLen <= 0) return utf8Str;  // Fallback on error

        std::wstring wideStr(wideLen, 0);
        MultiByteToWideChar(CP_UTF8, 0, utf8Str.c_str(),
            static_cast<int>(utf8Str.length()), &wideStr[0], wideLen);

        // Convert to lowercase using Windows locale-aware function
        CharLowerBuffW(&wideStr[0], wideLen);

        // Convert back to UTF-8
        int utf8Len = WideCharToMultiByte(CP_UTF8, 0, wideStr.c_str(), wideLen,
            nullptr, 0, nullptr, nullptr);
        if (utf8Len <= 0) return utf8Str;  // Fallback on error

        std::string result(utf8Len, 0);
        WideCharToMultiByte(CP_UTF8, 0, wideStr.c_str(), wideLen,
            &result[0], utf8Len, nullptr, nullptr);

        return result;
    }

    // ----------------------------------------------------------------------------
    // Locale-aware number formatting with thousand separators
    // Uses Windows user locale settings:
    // - US/UK: 1,234,567
    // - DE/AT/CH: 1.234.567
    // - FR: 1 234 567 (with narrow no-break space)
    // ----------------------------------------------------------------------------
    std::wstring formatNumber(size_t number)
    {
        std::wstring numStr = std::to_wstring(number);

        // Get locale-specific thousand separator from Windows
        wchar_t thousandSep[8] = L",";
        if (GetLocaleInfoW(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, thousandSep, 8) == 0) {
            // Fallback to comma if API fails
            wcscpy_s(thousandSep, L",");
        }

        // Insert thousand separators (from right to left)
        int insertPosition = static_cast<int>(numStr.length()) - 3;
        while (insertPosition > 0) {
            numStr.insert(insertPosition, thousandSep);
            insertPosition -= 3;
        }

        return numStr;
    }


} // namespace StringUtils