// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
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

#include <cstddef>
#include <string>

// Extended search mode: \r \n \t \\ \0 and the numbers \oNNN \bNNNNNNNN \dNNN \xHH \uHHHH;
// a sequence that is unknown or does not read as its number stays literal
namespace ExtendedEscapes {

    namespace detail {

        // Value of the digits at pos in the given base; false if one is missing or no digit of that base
        inline bool readNumber(const std::wstring& src, std::size_t pos, unsigned base, std::size_t digits, wchar_t& out)
        {
            if (pos + digits > src.size()) return false;

            unsigned value = 0;
            for (std::size_t k = 0; k < digits; ++k) {
                const wchar_t ch = src[pos + k];
                unsigned v = 0;
                if (ch >= L'0' && ch <= L'9')      v = static_cast<unsigned>(ch - L'0');
                else if (ch >= L'A' && ch <= L'F') v = static_cast<unsigned>(ch - L'A') + 10;
                else if (ch >= L'a' && ch <= L'f') v = static_cast<unsigned>(ch - L'a') + 10;
                else return false;
                if (v >= base) return false;
                value = value * base + v;
            }
            if (value > 0xFFFF) return false;   // outside the BMP
            out = static_cast<wchar_t>(value);
            return true;
        }

        // Appends the number whose letter is at src[pos] and returns the digits it took;
        // one that does not read stays literal and takes none
        inline std::size_t appendNumber(std::wstring& out, const std::wstring& src, std::size_t pos, unsigned base, std::size_t digits)
        {
            wchar_t decoded = 0;
            if (readNumber(src, pos + 1, base, digits, decoded)) {
                out.push_back(decoded);
                return digits;
            }
            out.push_back(L'\\');
            out.push_back(src[pos]);
            return 0;
        }
    }

    inline std::wstring expand(const std::wstring& input)
    {
        std::wstring out;
        out.reserve(input.size());

        for (std::size_t i = 0; i < input.size(); ++i) {
            if (input[i] != L'\\' || i + 1 >= input.size()) {
                out.push_back(input[i]);
                continue;
            }
            switch (input[++i]) {
            case L'r':  out.push_back(L'\r');                            break;
            case L'n':  out.push_back(L'\n');                            break;
            case L't':  out.push_back(L'\t');                            break;
            case L'\\': out.push_back(L'\\');                            break;
            case L'0':  out.push_back(L'\0');                            break;
            case L'o':  i += detail::appendNumber(out, input, i, 8, 3);  break;
            case L'b':  i += detail::appendNumber(out, input, i, 2, 8);  break;
            case L'd':  i += detail::appendNumber(out, input, i, 10, 3); break;
            case L'x':  i += detail::appendNumber(out, input, i, 16, 2); break;
            case L'u':  i += detail::appendNumber(out, input, i, 16, 4); break;
            default:    out.push_back(L'\\'); out.push_back(input[i]);   break;
            }
        }
        return out;
    }
}
