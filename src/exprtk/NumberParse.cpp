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
//
// NumberParse.cpp
// ExprTkEngine::parseCaptureToDouble and the match history forward
// here; this is the single source of truth for reading numbers.

#include "NumberParse.h"

#include <charconv>
#include <limits>
#include <system_error>

namespace MultiReplaceEngine {

    double parseNumber(const std::string& s)
    {
        constexpr double nan = std::numeric_limits<double>::quiet_NaN();
        const auto isDigit = [](char c) { return c >= '0' && c <= '9'; };

        std::size_t i = 0;
        while (i < s.size() && (s[i] == ' ' || s[i] == '\t')) ++i;

        // Sign and mantissa go into buf with '.' as the decimal point.
        // Exactly one decimal separator is allowed, and only one followed
        // by a digit counts: in "1.5, rest" the comma belongs to the rest.
        // More - mixed ("1.234,56") or repeated ("1.000.000") - is a
        // thousands-grouped or ambiguous form, so it yields NaN rather
        // than a misleading partial value.
        std::string buf;
        if (i < s.size() && (s[i] == '+' || s[i] == '-')) {
            if (s[i] == '-') buf.push_back('-');
            ++i;
        }
        bool digits = false;
        int separators = 0;
        for (; i < s.size(); ++i) {
            const char c = s[i];
            if (isDigit(c)) {
                buf.push_back(c);
                digits = true;
            }
            else if ((c == '.' || c == ',') && i + 1 < s.size() && isDigit(s[i + 1])) {
                buf.push_back('.');
                ++separators;
            }
            else {
                break;
            }
        }
        if (!digits || separators > 1) {
            return nan;
        }

        // Optional exponent; anything after the number is ignored ("1.5abc" -> 1.5)
        if (i < s.size() && (s[i] == 'e' || s[i] == 'E')) {
            std::size_t j = i + 1;
            if (j < s.size() && (s[j] == '+' || s[j] == '-')) ++j;
            if (j < s.size() && isDigit(s[j])) {
                while (j < s.size() && isDigit(s[j])) ++j;
                buf.append(s, i, j - i);
            }
        }

        double value = 0.0;
        const auto res = std::from_chars(buf.data(), buf.data() + buf.size(), value);
        if (res.ec != std::errc{} || res.ptr != buf.data() + buf.size()) {
            return nan;
        }
        return value;
    }

} // namespace MultiReplaceEngine