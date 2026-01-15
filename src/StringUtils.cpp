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
#include "Encoding.h"   // only used in translateEscapes()

#include <bitset>
#include <algorithm>
#include <cwctype>      // iswdigit

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
        std::vector<std::wstring> columns;
        std::wstring current;
        bool insideQuotes = false;

        for (size_t i = 0; i < line.size(); ++i) {
            wchar_t ch = line[i];
            if (ch == L'"') {
                if (insideQuotes && i + 1 < line.size() && line[i + 1] == L'"') {
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
            else if (!iswdigit(static_cast<wint_t>(static_cast<unsigned char>(c)))) {
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
    void handleEscapeSequence(const std::regex& rx,
        const std::string& in,
        std::string& out,
        std::function<char(const std::string&)> conv)
    {
        std::sregex_iterator it(in.begin(), in.end(), rx);
        std::sregex_iterator end;
        for (; it != end; ++it) {
            std::smatch m = *it;
            std::string token = m.str();
            try {
                char ch = conv(token);
                size_t pos = out.find(token);
                if (pos != std::string::npos) {
                    out.replace(pos, token.size(), 1, ch);
                }
            }
            catch (...) {
                // keep literal on conversion failure
            }
        }
    }

    std::string escapeSpecialChars(const std::string& input, bool extended) {
        std::string output = input;

        const std::string supported = "nrt0xubd";
        const std::string specials = "$.*[]^&\\{}()?+|<>\"'`~;#";

        for (char c : specials) {
            std::string needle(1, c);
            size_t pos = output.find(needle);

            while (pos != std::string::npos) {
                if (needle == "\\" && (pos == 0 || output[pos - 1] != '\\')) {
                    if (extended && (pos + 1 < output.size() &&
                        supported.find(output[pos + 1]) != std::string::npos)) {
                        pos = output.find(needle, pos + 1);
                        continue;
                    }
                }
                output.insert(pos, "\\");
                pos = output.find(needle, pos + 2);
            }
        }
        return output;
    }

    std::string translateEscapes(const std::string& input) {
        std::string output = input;

        std::regex reOct("\\\\o([0-7]{3})");
        std::regex reDec("\\\\d([0-9]{3})");
        std::regex reHex("\\\\x([0-9a-fA-F]{2})");
        std::regex reBin("\\\\b([01]{8})");
        std::regex reUni("\\\\u([0-9a-fA-F]{4})");
        std::regex reNL("\\\\n");
        std::regex reCR("\\\\r");
        std::regex reNul("\\\\0");

        handleEscapeSequence(reOct, input, output,
            [](const std::string& s) { return static_cast<char>(std::stoi(s.substr(2), nullptr, 8)); });
        handleEscapeSequence(reDec, input, output,
            [](const std::string& s) { return static_cast<char>(std::stoi(s.substr(2))); });
        handleEscapeSequence(reHex, input, output,
            [](const std::string& s) { return static_cast<char>(std::stoi(s.substr(2), nullptr, 16)); });
        handleEscapeSequence(reBin, input, output,
            [](const std::string& s) { return static_cast<char>(std::bitset<8>(s.substr(2)).to_ulong()); });

        // \uXXXX → UTF-8, first byte (as used in panel)
        handleEscapeSequence(reUni, input, output,
            [](const std::string& s)->char {
                int cp = std::stoi(s.substr(2), nullptr, 16);
                std::wstring w(1, static_cast<wchar_t>(cp));
                std::string utf8 = Encoding::wstringToUtf8(w);
                return utf8.empty() ? 0 : utf8.front();
            });

        // simple escapes → placeholders (Extended-mode downstream)
        output = std::regex_replace(output, reNL, "__NEWLINE__");
        output = std::regex_replace(output, reCR, "__CARRIAGERETURN__");
        output = std::regex_replace(output, reNul, ""); // \0 not supported
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
            result = std::regex_replace(result, std::regex("\n"), "__NEWLINE__");
            result = std::regex_replace(result, std::regex("\r"), "__CARRIAGERETURN__");
        }
        else if (mode == ReplaceMode::Regex) {
            result = std::regex_replace(result, std::regex("\n"), "\\n");
            result = std::regex_replace(result, std::regex("\r"), "\\r");
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


} // namespace StringUtils
