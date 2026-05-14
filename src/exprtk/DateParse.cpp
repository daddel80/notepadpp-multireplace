// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.

#include "DateParse.h"

#include <cctype>
#include <cstring>

namespace MultiReplace {

    namespace {

        // ASCII fast-path digit check. std::isdigit is locale-aware
        // which we don't want here - input is always ASCII digits.
        inline bool isAsciiDigit(char c) noexcept {
            return c >= '0' && c <= '9';
        }

        inline bool isAsciiSpace(char c) noexcept {
            // Matches the C locale's isspace set without the locale call.
            return c == ' ' || c == '\t' || c == '\n'
                || c == '\v' || c == '\f' || c == '\r';
        }

        // Read up to `maxDigits` ASCII digits starting at input[i].
        // Stops early when a non-digit is hit or maxDigits is reached.
        // Returns false if no digit was read; otherwise writes the
        // numeric value into `out` and advances i past the digits.
        bool readInt(std::string_view input, std::size_t& i,
                     int maxDigits, int& out) noexcept
        {
            int value = 0;
            int digits = 0;
            while (i < input.size() && digits < maxDigits
                   && isAsciiDigit(input[i]))
            {
                value = value * 10 + (input[i] - '0');
                ++i;
                ++digits;
            }
            if (digits == 0) return false;
            out = value;
            return true;
        }

        // Skip zero-or-more ASCII whitespace at input[i].
        void skipSpaces(std::string_view input, std::size_t& i) noexcept {
            while (i < input.size() && isAsciiSpace(input[i])) {
                ++i;
            }
        }

        // Try to consume a single literal `expected` from input at i.
        // Whitespace is *not* special-cased here - that's done at the
        // format-loop level so a single space in the format can absorb
        // any run of whitespace in the input.
        inline bool consumeLiteral(std::string_view input, std::size_t& i,
                                   char expected) noexcept
        {
            if (i >= input.size() || input[i] != expected) return false;
            ++i;
            return true;
        }

        // Case-insensitive ASCII compare for AM/PM tokens. Returns true
        // if input[i..] starts with the 2-char tag and advances i past
        // it; false otherwise. Only ASCII letters are compared, so the
        // caller's tag must be 2 ASCII chars.
        bool consumeCaseInsensitive2(std::string_view input, std::size_t& i,
                                     char a, char b) noexcept
        {
            if (i + 1 >= input.size()) return false;
            const char c0 = input[i];
            const char c1 = input[i + 1];
            const auto toUpper = [](char c) {
                return (c >= 'a' && c <= 'z') ? static_cast<char>(c - 32) : c;
            };
            if (toUpper(c0) != a || toUpper(c1) != b) return false;
            i += 2;
            return true;
        }

        // Result of a single specifier dispatch. Lets handleSpecifier
        // signal "I rewrote the format" (for %F / %T) without using
        // exceptions or out-params for flow control.
        struct SpecResult {
            bool ok;          // false = parse failure, abort
            bool didRewrite;  // %F/%T expanded into a sub-format
        };

        // Forward declaration for the recursive-format case.
        bool parseImpl(std::string_view input, std::size_t& i,
                       std::string_view format, std::size_t& f,
                       std::tm& tm);

        // Handle one %-specifier at format[f] (f points to the char
        // after '%'). Updates tm, advances input cursor i. Format
        // cursor f is left pointing AT the specifier letter; the
        // outer loop bumps past it.
        SpecResult handleSpecifier(std::string_view input, std::size_t& i,
                                   std::string_view format, std::size_t& f,
                                   std::tm& tm) noexcept
        {
            // f currently points at the specifier letter itself.
            const char spec = format[f];
            int v = 0;

            switch (spec) {
            case '%':
                if (!consumeLiteral(input, i, '%')) return { false, false };
                return { true, false };

            case 'Y':
                if (!readInt(input, i, 4, v)) return { false, false };
                if (v < 1 || v > 9999) return { false, false };
                tm.tm_year = v - 1900;
                return { true, false };

            case 'y':
                if (!readInt(input, i, 2, v)) return { false, false };
                // POSIX rule: 00..68 -> 2000..2068, 69..99 -> 1969..1999.
                tm.tm_year = (v < 69) ? (v + 2000 - 1900) : (v + 1900 - 1900);
                return { true, false };

            case 'm':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 1 || v > 12) return { false, false };
                tm.tm_mon = v - 1;
                return { true, false };

            case 'd':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 1 || v > 31) return { false, false };
                tm.tm_mday = v;
                return { true, false };

            case 'H':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 0 || v > 23) return { false, false };
                tm.tm_hour = v;
                return { true, false };

            case 'I':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 1 || v > 12) return { false, false };
                // Store the 12-hour value. %p will normalise it to 24h
                // when it runs; if %p never appears, 12 means 12 and
                // 1..11 mean 1..11 - which is what callers asking for
                // an %I-only parse would expect.
                tm.tm_hour = (v == 12) ? 0 : v;
                return { true, false };

            case 'M':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 0 || v > 59) return { false, false };
                tm.tm_min = v;
                return { true, false };

            case 'S':
                if (!readInt(input, i, 2, v)) return { false, false };
                if (v < 0 || v > 60) return { false, false };  // leap sec
                tm.tm_sec = v;
                return { true, false };

            case 'p':
                // Expects exactly "AM" or "PM" (any case). If %I already
                // ran, tm_hour holds 0..11 - add 12 for PM. If %H ran
                // instead, %p has no effect (24h is already correct).
                if (consumeCaseInsensitive2(input, i, 'A', 'M')) {
                    // AM: nothing to do. 12 AM was already normalised
                    // to 0 by the %I handler.
                    return { true, false };
                }
                if (consumeCaseInsensitive2(input, i, 'P', 'M')) {
                    if (tm.tm_hour < 12) tm.tm_hour += 12;
                    return { true, false };
                }
                return { false, false };

            case 'F': {
                // %F expands to "%Y-%m-%d". Recurse on the sub-format
                // so all range-checks and digit-count limits stay in
                // one code path.
                constexpr std::string_view sub("%Y-%m-%d");
                std::size_t subF = 0;
                if (!parseImpl(input, i, sub, subF, tm)) {
                    return { false, false };
                }
                return { true, true };
            }

            case 'T': {
                constexpr std::string_view sub("%H:%M:%S");
                std::size_t subF = 0;
                if (!parseImpl(input, i, sub, subF, tm)) {
                    return { false, false };
                }
                return { true, true };
            }

            default:
                // Unknown specifier -> fail rather than silently skip.
                // Keeps user mistakes visible.
                return { false, false };
            }
        }

        // Main parse loop. Walks format and input in lockstep; tracks
        // both indices by reference so the recursive %F / %T cases
        // share state.
        bool parseImpl(std::string_view input, std::size_t& i,
                       std::string_view format, std::size_t& f,
                       std::tm& tm)
        {
            while (f < format.size()) {
                const char fc = format[f];

                if (fc == '%') {
                    if (f + 1 >= format.size()) {
                        // Trailing '%' with no specifier letter - malformed.
                        return false;
                    }
                    ++f;  // step onto specifier letter
                    const SpecResult r = handleSpecifier(input, i, format, f, tm);
                    if (!r.ok) return false;
                    ++f;  // step past specifier letter (or past the
                          // %F/%T letter; the sub-format inside ran
                          // independently, no extra advance needed).
                    continue;
                }

                if (isAsciiSpace(fc)) {
                    // Whitespace in format -> zero-or-more whitespace
                    // in input. POSIX strptime behaviour: a single
                    // space absorbs runs of any whitespace.
                    skipSpaces(input, i);
                    ++f;
                    continue;
                }

                // Literal character. Must match exactly.
                if (i >= input.size() || input[i] != fc) {
                    return false;
                }
                ++i;
                ++f;
            }
            return true;
        }

    }  // namespace

    bool parseDateTime(std::string_view input,
                       std::string_view format,
                       std::tm& tm)
    {
        std::size_t i = 0;
        std::size_t f = 0;
        if (!parseImpl(input, i, format, f, tm)) {
            return false;
        }
        // POSIX strptime allows trailing input. We do the same: the
        // caller may have appended a timezone, comments, etc. The
        // contract is "format was satisfied", not "input was emptied".
        return true;
    }

}  // namespace MultiReplace
