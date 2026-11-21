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

namespace StringUtils {

    // ----------------------------------------------------------------------------
    // Quoted-string escape helpers (generic; used by INI cache,
    // CSV cell bodies, and MR settings-block string values).
    // ----------------------------------------------------------------------------
    std::wstring escapeQuoted(const std::wstring& value) {
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

    std::wstring unescapeQuoted(const std::wstring& value) {
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

    std::wstring quoteField(const std::wstring& value) {
        std::wstring out;
        out.reserve(value.size() + 4);
        out += L'"';
        for (wchar_t ch : value) {
            if (ch == L'"') out += L"\"\"";
            else            out += ch;
        }
        out += L'"';
        return out;
    }

    std::wstring rfcQuote(const std::wstring& value, wchar_t delimiter) {
        bool needsQuote = false;
        for (wchar_t ch : value) {
            if (ch == delimiter || ch == L'"' || ch == L'\n' || ch == L'\r') {
                needsQuote = true;
                break;
            }
        }
        if (!needsQuote) return value;

        std::wstring out;
        out.reserve(value.size() + 4);
        out += L'"';
        for (wchar_t ch : value) {
            if (ch == L'"') out += L"\"\"";
            else            out += ch;
        }
        out += L'"';
        return out;
    }

    std::wstring rfcUnquote(const std::wstring& value) {
        if (value.size() < 2 || value.front() != L'"' || value.back() != L'"') {
            return value;  // unquoted cell: verbatim
        }
        std::wstring out;
        out.reserve(value.size());
        for (size_t i = 1; i + 1 < value.size(); ++i) {
            if (value[i] == L'"' && i + 2 < value.size() && value[i + 1] == L'"') {
                out += L'"';  // "" -> "
                ++i;
            }
            else {
                out += value[i];
            }
        }
        return out;
    }

    std::wstring newlinesToCrlf(const std::wstring& value) {
        std::wstring out;
        out.reserve(value.size());
        for (size_t i = 0; i < value.size(); ++i) {
            const wchar_t c = value[i];
            if (c == L'\r') {
                if (i + 1 < value.size() && value[i + 1] == L'\n') {
                    out += L"\r\n";  // CRLF stays
                    ++i;
                }
                else {
                    out += L'\r';    // lone CR stays silent
                }
            }
            else if (c == L'\n') {
                out += L"\r\n";      // bare LF -> CRLF
            }
            else {
                out += c;
            }
        }
        return out;
    }

    std::wstring newlinesForPlainCsv(const std::wstring& value) {
        std::wstring out;
        out.reserve(value.size());
        for (size_t i = 0; i < value.size(); ++i) {
            const wchar_t c = value[i];
            if (c == L'\r') {
                if (i + 1 < value.size() && value[i + 1] == L'\n') {
                    out += L'\n';    // CRLF -> bare LF (real in-cell break)
                    ++i;
                }
                else {
                    out += L'\r';    // lone CR stays
                }
            }
            else if (c == L'\n') {
                out += L"\\n";       // foreign lone LF -> literal backslash-n
            }
            else {
                out += c;
            }
        }
        return out;
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

    std::vector<std::wstring> splitFilterPatterns(const std::wstring& filter)
    {
        std::vector<std::wstring> out;
        size_t pos = 0;
        while (pos <= filter.size()) {
            const size_t sep = filter.find(L';', pos);
            const size_t end = (sep == std::wstring::npos) ? filter.size() : sep;
            std::wstring pattern = trim(filter.substr(pos, end - pos));
            if (!pattern.empty())
                out.push_back(std::move(pattern));
            if (sep == std::wstring::npos) break;
            pos = sep + 1;
        }
        return out;
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
    // Unicode-aware lowercase via the Windows API - handles non-ASCII
    // (umlauts, accents, locale rules) correctly, unlike tolower().
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
    // Locale-aware thousand separators via the Windows user locale
    // (US: 1,234,567 / DE: 1.234.567).
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

    std::wstring replaceTemplateVar(const std::wstring& tmpl,
        const std::wstring& var,
        const std::wstring& value) {
        std::wstring result = tmpl;
        size_t pos = 0;
        while ((pos = result.find(var, pos)) != std::wstring::npos) {
            result.replace(pos, var.length(), value);
            pos += value.length();
        }
        return result;
    }

    std::wstring processTemplateEscapes(const std::wstring& tmpl) {
        std::wstring result;
        result.reserve(tmpl.size());

        for (size_t i = 0; i < tmpl.size(); ++i) {
            if (tmpl[i] == L'\\' && i + 1 < tmpl.size()) {
                switch (tmpl[i + 1]) {
                case L't': result += L'\t'; ++i; break;
                case L'n': result += L'\n'; ++i; break;
                case L'r': result += L'\r'; ++i; break;
                case L'\\': result += L'\\'; ++i; break;
                default: result += tmpl[i]; break;
                }
            }
            else {
                result += tmpl[i];
            }
        }
        return result;
    }


} // namespace StringUtils