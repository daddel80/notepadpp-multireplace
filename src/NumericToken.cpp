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

    bool mr::num::classify_numeric_field(std::string_view f,
        NumericToken& outTok,
        const ParseOptions& opt)
    {
        // Helper lambdas
        auto is_space = [](unsigned char c) noexcept { return c == ' ' || c == '\t'; };
        auto is_digit = [](unsigned char c) noexcept { return c >= '0' && c <= '9'; };
        auto is_letter = [](unsigned char c) noexcept { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'); };
        auto is_forbidden_affix = [&](unsigned char c) noexcept {
            return is_digit(c) || c == '+' || c == '-' || c == '.' || c == ',';
            };

        const std::size_t n = f.size();
        if (n == 0) return false;

        // The caller promised 'f' is already trimmed (no leading/trailing ws).
        // Guard anyway (reject if found).
        if (is_space((unsigned char)f.front()) || is_space((unsigned char)f.back()))
            return false;

        // First, find the numeric token using the tolerant tokenizer.
        NumericToken tok = parse_first_numeric_token(f, opt);
        if (!tok.ok)
            return false;

        // The token must be the *only* numeric content; we allow optional prefix/suffix
        // around it, nothing else. Split: [0..tok.start) [tok] [tok.end..n)
        std::size_t L0 = 0;
        std::size_t L1 = tok.start;  // left side (pre-token)
        std::size_t R0 = tok.end;    // right side (post-token)
        std::size_t R1 = n;

        // LEFT: contiguous non-space run directly before token (optional affix)
        bool left_ok = true;
        bool left_has_affix = false;
        bool left_affix_letters = false;

        std::size_t left_run_beg = 0, left_run_end = 0; // [beg, end)
        if (L1 > L0) {
            // Skip spaces just before the token
            std::size_t p = L1;
            while (p > L0 && is_space((unsigned char)f[p - 1])) --p;
            if (p > L0) {
                // Non-space run immediately before spaces (or token if adjacent)
                std::size_t q = p;
                while (q > L0 && !is_space((unsigned char)f[q - 1])) --q;
                left_run_beg = q;
                left_run_end = p;
                left_has_affix = true;

                // Everything before the run must be spaces only
                for (std::size_t i = L0; i < q; ++i) { if (!is_space((unsigned char)f[i])) { left_ok = false; break; } }

                // Validate run: length <= 4, no digits and no "+-.,"
                const std::size_t run_len = p - q;
                if (left_ok) {
                    if (run_len == 0 || run_len > 4) left_ok = false;
                    else {
                        bool any_forbidden = false, any_letter = false, any_symbol = false;
                        for (std::size_t i = q; i < p; ++i) {
                            unsigned char c = (unsigned char)f[i];
                            if (is_forbidden_affix(c)) { any_forbidden = true; break; }
                            if (is_letter(c)) any_letter = true; else any_symbol = true;
                        }
                        if (any_forbidden || (any_letter && any_symbol)) left_ok = false;
                        left_affix_letters = any_letter && !any_symbol;
                        // If letters, there must be at least one space between affix and token
                        if (left_ok && left_affix_letters) {
                            if (L1 == p || !is_space((unsigned char)f[L1 - 1])) left_ok = false;
                        }
                    }
                }
            }
            else {
                // Only spaces on left side
                for (std::size_t i = L0; i < L1; ++i) { if (!is_space((unsigned char)f[i])) { left_ok = false; break; } }
            }
        }

        if (!left_ok) return false;

        // RIGHT: contiguous non-space run immediately after token (optional affix)
        bool right_ok = true;
        bool right_has_affix = false;
        bool right_affix_letters = false;

        std::size_t right_run_beg = 0, right_run_end = 0; // [beg, end)
        if (R1 > R0) {
            // Skip spaces just after the token
            std::size_t p = R0;
            while (p < R1 && is_space((unsigned char)f[p])) ++p;
            if (p < R1) {
                // Non-space run immediately after spaces (or adjacent)
                std::size_t q = p;
                while (q < R1 && !is_space((unsigned char)f[q])) ++q;
                right_run_beg = p;
                right_run_end = q;
                right_has_affix = true;

                // Everything after the run must be spaces only
                for (std::size_t i = q; i < R1; ++i) { if (!is_space((unsigned char)f[i])) { right_ok = false; break; } }

                // Validate run: length <= 4, no digits and no "+-.,"
                const std::size_t run_len = q - p;
                if (right_ok) {
                    if (run_len == 0 || run_len > 4) right_ok = false;
                    else {
                        bool any_forbidden = false, any_letter = false, any_symbol = false;
                        for (std::size_t i = p; i < q; ++i) {
                            unsigned char c = (unsigned char)f[i];
                            if (is_forbidden_affix(c)) { any_forbidden = true; break; }
                            if (is_letter(c)) any_letter = true; else any_symbol = true;
                        }
                        if (any_forbidden || (any_letter && any_symbol)) right_ok = false;
                        right_affix_letters = any_letter && !any_symbol;
                        // If letters, there must be at least one space between token and affix
                        if (right_ok && right_affix_letters) {
                            if (p == R0) right_ok = false; // adjacent => not allowed for letter affix
                        }
                    }
                }
            }
            else {
                // Only spaces on right side
                for (std::size_t i = R0; i < R1; ++i) { if (!is_space((unsigned char)f[i])) { right_ok = false; break; } }
            }
        }

        if (!right_ok) return false;

        // Passed: field is a numeric field (affixes optional)
        outTok = tok; // start/end refer to 'f' (trimmed view)
        return true;
    }


}} // namespace mr::num
