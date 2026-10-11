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
#include <cstdint>
#include <cstring>
#include <string_view>

#include "Scintilla.h"

// Find/Replace in Files: before a file is filled into Scintilla, can a search match in it at all?
// Only a necessary condition: false means Scintilla cannot match, true decides nothing, every hit
// still comes from Scintilla. A regex is never decided. With match case the exact bytes Scintilla
// searches for must occur. Without match case only ASCII text with an ASCII pattern is decided,
// since Scintilla folds beyond ASCII ("ß" matches "ss", the Kelvin sign "k", the long s "s").
class LiteralPrefilter {
public:
    explicit LiteralPrefilter(std::string_view text) noexcept : _text(text) {}

    // pattern: the bytes Scintilla gets (Extended escapes resolved, in the document's codepage)
    // searchFlags: as set with SCI_SETSEARCHFLAGS
    bool mayMatch(std::string_view pattern, int searchFlags) noexcept {
        if (pattern.empty() || (searchFlags & SCFIND_REGEXP))
            return true;
        if (searchFlags & SCFIND_MATCHCASE)
            return contains(pattern);
        if (!isAscii(pattern) || !textIsAscii())
            return true;
        return containsIgnoringAsciiCase(pattern);
    }

    static bool isAscii(std::string_view s) noexcept {
        const char* p = s.data();
        size_t n = s.size();
        for (; n >= 8; p += 8, n -= 8) {
            uint64_t block;
            std::memcpy(&block, p, 8);
            if (block & 0x8080808080808080ull) return false;
        }
        for (; n > 0; ++p, --n)
            if (static_cast<unsigned char>(*p) >= 0x80) return false;
        return true;
    }

private:
    bool textIsAscii() noexcept {
        if (_textAscii < 0) _textAscii = isAscii(_text) ? 1 : 0;
        return _textAscii == 1;
    }

    // Next occurrence of c in [from, last], or nullptr
    static const char* find(const char* from, const char* last, char c) noexcept {
        if (from > last) return nullptr;
        return static_cast<const char*>(std::memchr(from, c, static_cast<size_t>(last - from) + 1));
    }

    static char lowerAscii(char c) noexcept {
        return (c >= 'A' && c <= 'Z') ? static_cast<char>(c + ('a' - 'A')) : c;
    }

    bool contains(std::string_view pattern) const noexcept {
        if (pattern.size() > _text.size()) return false;
        const char* const last = _text.data() + (_text.size() - pattern.size());   // last possible start
        for (const char* p = find(_text.data(), last, pattern[0]); p; p = find(p + 1, last, pattern[0]))
            if (std::memcmp(p + 1, pattern.data() + 1, pattern.size() - 1) == 0)
                return true;
        return false;
    }

    // Both cases of the first byte by memchr, the rest compared case-insensitively
    bool containsIgnoringAsciiCase(std::string_view pattern) const noexcept {
        if (pattern.size() > _text.size()) return false;
        const char* const last = _text.data() + (_text.size() - pattern.size());
        const char lower = lowerAscii(pattern[0]);
        const char upper = (lower >= 'a' && lower <= 'z') ? static_cast<char>(lower - ('a' - 'A')) : lower;
        const char* atLower = find(_text.data(), last, lower);
        const char* atUpper = (upper != lower) ? find(_text.data(), last, upper) : nullptr;
        while (atLower || atUpper) {
            const bool takeLower = atLower && (!atUpper || atLower < atUpper);
            const char* p = takeLower ? atLower : atUpper;
            size_t i = 1;
            while (i < pattern.size() && lowerAscii(p[i]) == lowerAscii(pattern[i])) ++i;
            if (i == pattern.size()) return true;
            if (takeLower) atLower = find(atLower + 1, last, lower);
            else atUpper = find(atUpper + 1, last, upper);
        }
        return false;
    }

    std::string_view _text;
    int _textAscii = -1;   // not checked yet
};
