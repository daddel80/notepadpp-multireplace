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

#include "NumericToken.h"
#include <stdexcept>

namespace mr { namespace num {

    // --- helpers -------------------------------------------------------------

    static inline bool is_ascii_digit(unsigned char c) noexcept {
        return c >= '0' && c <= '9';
    }

    // Find start: DIGIT or SIGN+DIGIT; fallback (opt.): SIGN? + ('.'|',') + DIGIT
    static std::size_t find_token_start(std::string_view s, bool allowLeadingSeparator)
    {
        const std::size_t n = s.size();

        // primary
        for (std::size_t i = 0; i < n; ++i) {
            const unsigned char c = (unsigned char)s[i];
            if (is_ascii_digit(c)) return i;
            if ((c == '+' || c == '-') && (i + 1) < n && is_ascii_digit((unsigned char)s[i + 1]))
                return i;
        }

        // fallback
        if (allowLeadingSeparator) {
            for (std::size_t i = 0; i + 1 < n; ++i) {
                const unsigned char c  = (unsigned char)s[i];
                const unsigned char c2 = (unsigned char)s[i + 1];

                if ((c == '+' || c == '-') && (i + 2) < n) {
                    const unsigned char c3 = (unsigned char)s[i + 2];
                    if ((c2 == '.' || c2 == ',') && is_ascii_digit(c3))
                        return i;
                }
                if ((c == '.' || c == ',') && is_ascii_digit(c2))
                    return i;
            }
        }
        return std::string::npos;
    }

    // Normalize for std::stod: ','->'.', ".5"->"0.5", "-.5"->"-0.5", "12."->"12"
    static bool normalize_token(std::string& token)
    {
        for (char& ch : token) if (ch == ',') ch = '.';

        if (!token.empty() && token[0] == '.') {
            token.insert(token.begin(), '0');
        } else if (token.size() >= 2 && (token[0] == '+' || token[0] == '-') && token[1] == '.') {
            token.insert(token.begin() + 1, '0');
        }

        if (!token.empty() && token.back() == '.') {
            token.pop_back();
            if (token.empty() || token == "+" || token == "-")
                return false;
        }
        return true;
    }

    // --- public API ----------------------------------------------------------

    NumericToken parse_first_numeric_token(std::string_view field,
                                           const ParseOptions& opt)
    {
        NumericToken out{};
        const std::size_t n = field.size();
        if (n == 0) return out;

        const std::size_t start = find_token_start(field, opt.allowLeadingSeparator);
        if (start == std::string::npos) return out;

        std::size_t p = start;
        if (field[p] == '+' || field[p] == '-') { out.hasSign = true; ++p; }

        const std::size_t intBeg = p;
        while (p < n && is_ascii_digit((unsigned char)field[p])) ++p;
        out.intDigits = (int)(p - intBeg);

        if (p < n && (field[p] == '.' || field[p] == ',')) {
            out.hasDecimal = true;
            ++p;
            while (p < n && is_ascii_digit((unsigned char)field[p])) ++p;
        }
        else if (out.intDigits == 0 && opt.allowLeadingSeparator) {
            if (p < n && (field[p] == '.' || field[p] == ',')) {
                out.hasDecimal = true;
                ++p;
                if (p >= n || !is_ascii_digit((unsigned char)field[p])) return out;
                while (p < n && is_ascii_digit((unsigned char)field[p])) ++p;
            }
        }

        if (out.intDigits == 0 && !out.hasDecimal) return out;

        out.start = start;
        out.end   = p;

        out.normalized.assign(field.substr(out.start, out.end - out.start));
        if (!normalize_token(out.normalized)) return out;

        try {
            out.value = std::stod(out.normalized);
        } catch (...) {
            return out;
        }

        out.ok = true;
        return out;
    }

    bool try_parse_first_numeric_value(std::string_view field,
                                       double& outValue,
                                       std::string* outNormalized,
                                       const ParseOptions& opt)
    {
        const NumericToken t = parse_first_numeric_token(field, opt);
        if (!t.ok) return false;
        outValue = t.value;
        if (outNormalized) *outNormalized = t.normalized;
        return true;
    }

}} // namespace mr::num
