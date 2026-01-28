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
#include <charconv>

namespace mr {
    namespace num {

        // =========================================================================
        // Internal Constants
        // =========================================================================

        // Maximum length for currency/unit prefix or suffix (e.g. "$", "USD", "€")
        static constexpr std::size_t MAX_AFFIX_LENGTH = 4;

        // =========================================================================
        // Helper Functions
        // =========================================================================

        static inline bool is_ascii_digit(unsigned char c) noexcept {
            return c >= '0' && c <= '9';
        }

        static inline bool is_space(unsigned char c) noexcept {
            return c == ' ' || c == '\t';
        }

        static inline bool is_letter(unsigned char c) noexcept {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        static inline bool is_sign(unsigned char c) noexcept {
            return c == '+' || c == '-';
        }

        static inline bool is_decimal_sep(unsigned char c) noexcept {
            return c == '.' || c == ',';
        }

        // Characters not allowed in prefix/suffix (would indicate another number)
        static inline bool is_forbidden_affix_char(unsigned char c) noexcept {
            return is_ascii_digit(c) || is_sign(c) || is_decimal_sep(c);
        }

        // =========================================================================
        // Token Finding
        // =========================================================================

        // Find start position of numeric token.
        // Recognizes: DIGIT, SIGN+DIGIT, SIGN+SEP+DIGIT, SEP+DIGIT
        static std::size_t find_token_start(std::string_view s) noexcept
        {
            const std::size_t n = s.size();
            if (n == 0) return std::string::npos;

            // Scan for first potential token start
            for (std::size_t i = 0; i < n; ++i) {
                const unsigned char c = static_cast<unsigned char>(s[i]);

                // Direct digit
                if (is_ascii_digit(c))
                    return i;

                // Sign followed by digit or decimal
                if (is_sign(c) && (i + 1) < n) {
                    const unsigned char c2 = static_cast<unsigned char>(s[i + 1]);
                    // Sign + digit
                    if (is_ascii_digit(c2))
                        return i;
                    // Sign + decimal + digit (e.g. "-.5")
                    if (is_decimal_sep(c2) && (i + 2) < n &&
                        is_ascii_digit(static_cast<unsigned char>(s[i + 2])))
                        return i;
                }

                // Decimal separator followed by digit (e.g. ".5")
                if (is_decimal_sep(c) && (i + 1) < n &&
                    is_ascii_digit(static_cast<unsigned char>(s[i + 1])))
                    return i;
            }

            return std::string::npos;
        }

        // =========================================================================
        // Token Normalization
        // =========================================================================

        // Normalize token for parsing:
        // - Replace ',' with '.'
        // - ".5" -> "0.5", "-.5" -> "-0.5"
        // - "12." -> "12"
        // Returns false if token becomes invalid after normalization.
        static bool normalize_token(std::string& token) noexcept
        {
            // Replace comma with dot
            for (char& ch : token) {
                if (ch == ',') ch = '.';
            }

            // Add leading zero for ".5" -> "0.5"
            if (!token.empty() && token[0] == '.') {
                token.insert(token.begin(), '0');
            }
            // Add leading zero for "+.5" -> "+0.5" or "-.5" -> "-0.5"
            else if (token.size() >= 2 && is_sign(token[0]) && token[1] == '.') {
                token.insert(token.begin() + 1, '0');
            }

            // Remove trailing dot: "12." -> "12"
            if (!token.empty() && token.back() == '.') {
                token.pop_back();
                // Check if anything meaningful remains
                if (token.empty() || token == "+" || token == "-")
                    return false;
            }

            return true;
        }

        // =========================================================================
        // Token Parsing
        // =========================================================================

        // Parse the first numeric token from a string.
        // Returns NumericToken with ok=true on success.
        static NumericToken parse_first_numeric_token(std::string_view field) noexcept
        {
            NumericToken out{};
            const std::size_t n = field.size();
            if (n == 0) return out;

            const std::size_t start = find_token_start(field);
            if (start == std::string::npos) return out;

            std::size_t p = start;

            // Optional sign
            if (is_sign(static_cast<unsigned char>(field[p]))) {
                out.hasSign = true;
                ++p;
            }

            // Check for leading decimal separator (e.g. ".5" or after sign "-.5")
            if (p < n && is_decimal_sep(static_cast<unsigned char>(field[p]))) {
                out.hasDecimal = true;
                ++p;
                // Must have at least one digit after separator
                if (p >= n || !is_ascii_digit(static_cast<unsigned char>(field[p])))
                    return out;
                while (p < n && is_ascii_digit(static_cast<unsigned char>(field[p]))) ++p;
            }
            else {
                // Integer part digits
                const std::size_t intBeg = p;
                while (p < n && is_ascii_digit(static_cast<unsigned char>(field[p]))) ++p;
                out.intDigits = static_cast<int>(p - intBeg);

                // Optional decimal part
                if (p < n && is_decimal_sep(static_cast<unsigned char>(field[p]))) {
                    out.hasDecimal = true;
                    ++p;
                    while (p < n && is_ascii_digit(static_cast<unsigned char>(field[p]))) ++p;
                }

                // Must have at least integer digits
                if (out.intDigits == 0)
                    return out;
            }

            out.start = start;
            out.end = p;

            // Extract and normalize the token
            out.normalized.assign(field.substr(out.start, out.end - out.start));
            if (!normalize_token(out.normalized))
                return out;

            // Parse the normalized value using std::from_chars (no exceptions)
            // Note: from_chars doesn't accept leading '+', so skip it
            const char* parseStart = out.normalized.data();
            if (!out.normalized.empty() && out.normalized[0] == '+')
                ++parseStart;

            auto [ptr, ec] = std::from_chars(
                parseStart,
                out.normalized.data() + out.normalized.size(),
                out.value);

            if (ec != std::errc())
                return out;

            out.ok = true;
            return out;
        }

        // =========================================================================
        // Affix Validation
        // =========================================================================

        // Validate a prefix or suffix.
        // Returns true if valid: letters only OR symbols only, max length, no forbidden chars.
        static bool validate_affix(std::string_view affix) noexcept
        {
            if (affix.empty() || affix.size() > MAX_AFFIX_LENGTH)
                return false;

            bool anyLetter = false;
            bool anySymbol = false;

            for (unsigned char c : affix) {
                if (is_forbidden_affix_char(c))
                    return false;
                if (is_letter(c))
                    anyLetter = true;
                else
                    anySymbol = true;
            }

            // Can't mix letters and symbols (e.g. "$U" is invalid)
            if (anyLetter && anySymbol)
                return false;

            return true;
        }

        // =========================================================================
        // Public API
        // =========================================================================

        bool classify_numeric_field(std::string_view f, NumericToken& outTok)
        {
            const std::size_t n = f.size();
            if (n == 0) return false;

            // Input must be pre-trimmed
            if (is_space(static_cast<unsigned char>(f.front())) ||
                is_space(static_cast<unsigned char>(f.back())))
                return false;

            // Parse the numeric token
            NumericToken tok = parse_first_numeric_token(f);
            if (!tok.ok)
                return false;

            // Validate left side (prefix)
            if (tok.start > 0) {
                std::size_t leftEnd = tok.start;

                // Skip trailing spaces before token
                while (leftEnd > 0 && is_space(static_cast<unsigned char>(f[leftEnd - 1])))
                    --leftEnd;

                if (leftEnd > 0) {
                    // Find start of non-space run (the affix)
                    std::size_t leftStart = leftEnd;
                    while (leftStart > 0 && !is_space(static_cast<unsigned char>(f[leftStart - 1])))
                        --leftStart;

                    // Everything before the affix must be spaces
                    for (std::size_t i = 0; i < leftStart; ++i) {
                        if (!is_space(static_cast<unsigned char>(f[i])))
                            return false;
                    }

                    // Validate the affix
                    std::string_view affix(f.data() + leftStart, leftEnd - leftStart);
                    if (!validate_affix(affix))
                        return false;
                }
            }

            // Validate right side (suffix)
            if (tok.end < n) {
                std::size_t rightStart = tok.end;

                // Skip leading spaces after token
                while (rightStart < n && is_space(static_cast<unsigned char>(f[rightStart])))
                    ++rightStart;

                if (rightStart < n) {
                    // Find end of non-space run (the affix)
                    std::size_t rightEnd = rightStart;
                    while (rightEnd < n && !is_space(static_cast<unsigned char>(f[rightEnd])))
                        ++rightEnd;

                    // Everything after the affix must be spaces
                    for (std::size_t i = rightEnd; i < n; ++i) {
                        if (!is_space(static_cast<unsigned char>(f[i])))
                            return false;
                    }

                    // Validate the affix
                    std::string_view affix(f.data() + rightStart, rightEnd - rightStart);
                    if (!validate_affix(affix))
                        return false;
                }
            }

            outTok = tok;
            return true;
        }

    }
} // namespace mr::num