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
#include <string>
#include <vector>

#include "ReplaceItemData.h"

// Content hash of a list, to tell it from the list as last saved or loaded. of() is the hash since 6.3; legacy() is
// the one of the versions before, kept in the INI for them, which misses changes that cancel out within an entry
namespace ListHash {

    namespace detail {

        // FNV-1a as std::hash of the MSVC library runs it for a size_t of eight and of four bytes
        template <class Word> struct Fnv;
        template <> struct Fnv<std::uint64_t> {
            static constexpr std::uint64_t offset = 14695981039346656037ULL;
            static constexpr std::uint64_t prime = 1099511628211ULL;
        };
        template <> struct Fnv<std::uint32_t> {
            static constexpr std::uint32_t offset = 2166136261U;
            static constexpr std::uint32_t prime = 16777619U;
        };

        // Calls onUnit for every UTF-16 code unit of text, whatever the width of wchar_t
        template <class OnUnit>
        void forEachUnit(const std::wstring& text, OnUnit&& onUnit)
        {
            if constexpr (sizeof(wchar_t) == 2) {
                for (const wchar_t ch : text) onUnit(static_cast<std::uint32_t>(ch));
            }
            else {
                for (const wchar_t ch : text) {
                    const std::uint32_t c = static_cast<std::uint32_t>(ch);
                    if (c > 0xFFFF) {
                        onUnit(0xD800 + ((c - 0x10000) >> 10));
                        onUnit(0xDC00 + ((c - 0x10000) & 0x3FF));
                    }
                    else {
                        onUnit(c);
                    }
                }
            }
        }

        // The hash of 4.x to 6.2 for a size_t of Word: std::hash of every field, the same shift for all, joined by XOR
        template <class Word>
        Word legacyAs(const std::vector<ReplaceItemData>& list)
        {
            const Word offset = Fnv<Word>::offset;
            const Word prime = Fnv<Word>::prime;
            const auto flag = [offset, prime](bool value) -> Word {
                return (offset ^ (value ? 1u : 0u)) * prime;
            };
            const auto text = [offset, prime](const std::wstring& value) -> Word {
                Word hash = offset;
                forEachUnit(value, [&hash, prime](std::uint32_t unit) {
                    hash = (hash ^ (unit & 0xFF)) * prime;
                    hash = (hash ^ (unit >> 8)) * prime;
                });
                return hash;
            };

            Word combined = 0;
            for (const ReplaceItemData& item : list) {
                Word hash = flag(item.isEnabled);
                hash ^= text(item.findText) << 1;
                hash ^= text(item.replaceText) << 1;
                hash ^= flag(item.wholeWord) << 1;
                hash ^= flag(item.matchCase) << 1;
                hash ^= flag(item.formulaSupport) << 1;
                hash ^= flag(item.extended) << 1;
                hash ^= flag(item.regex) << 1;
                hash ^= text(item.comments) << 1;
                combined ^= hash + 0x9e3779b9u + (combined << 6) + (combined >> 2);
            }
            return combined;
        }

        inline bool hexValue(const std::wstring& text, std::size_t from, std::uint64_t& value)
        {
            value = 0;
            for (std::size_t i = from; i < from + 16; ++i) {
                const wchar_t ch = text[i];
                std::uint64_t digit = 0;
                if (ch >= L'0' && ch <= L'9')      digit = static_cast<std::uint64_t>(ch - L'0');
                else if (ch >= L'A' && ch <= L'F') digit = static_cast<std::uint64_t>(ch - L'A') + 10;
                else if (ch >= L'a' && ch <= L'f') digit = static_cast<std::uint64_t>(ch - L'a') + 10;
                else return false;
                value = (value << 4) | digit;
            }
            return true;
        }
    }

    // Every entry in its order: options, Find, Replace, Comments. 0 is the empty list, the value a new tab starts with
    inline std::uint64_t of(const std::vector<ReplaceItemData>& list)
    {
        if (list.empty()) return 0;

        std::uint64_t hash = detail::Fnv<std::uint64_t>::offset;
        const auto step = [&hash](std::uint64_t value) { hash = (hash ^ value) * detail::Fnv<std::uint64_t>::prime; };
        const auto text = [&step](const std::wstring& value) {
            detail::forEachUnit(value, step);
            step(0x10000);   // ends the text: no code unit has this value
        };

        step(list.size());
        for (const ReplaceItemData& item : list) {
            step((item.isEnabled ? 1u : 0u) | (item.wholeWord ? 2u : 0u) | (item.matchCase ? 4u : 0u)
                | (item.formulaSupport ? 8u : 0u) | (item.extended ? 16u : 0u) | (item.regex ? 32u : 0u));
            text(item.findText);
            text(item.replaceText);
            text(item.comments);
        }
        return hash;
    }

    // The value a version before 6.3 computes on this platform
    inline std::uint64_t legacy(const std::vector<ReplaceItemData>& list)
    {
        if constexpr (sizeof(std::size_t) == 8) return detail::legacyAs<std::uint64_t>(list);
        else return detail::legacyAs<std::uint32_t>(list);
    }

    // A list as last saved or loaded
    struct Baseline {
        std::uint64_t hash = 0;      // of()
        std::uint64_t legacy = 0;    // legacy(), the value of the INI key the versions before 6.3 read and write

        bool matches(const std::vector<ReplaceItemData>& list) const { return of(list) == hash; }
    };

    inline Baseline baselineOf(const std::vector<ReplaceItemData>& list)
    {
        return { of(list), legacy(list) };
    }

    // The baseline behind a legacy hash from the INI: the file if the hash stands for it, else the list if it stands
    // for that; else it is unknown and no list matches it
    inline Baseline rebase(std::uint64_t stored, const std::vector<ReplaceItemData>& list, const std::vector<ReplaceItemData>* file)
    {
        if (file && legacy(*file) == stored) return { of(*file), stored };
        if (legacy(list) == stored) return { of(list), stored };
        return { stored, stored };
    }

    // The INI value of a baseline: "hash:legacy", 16 hex digits each
    inline std::wstring toText(const Baseline& baseline)
    {
        static const wchar_t digits[] = L"0123456789ABCDEF";
        std::wstring text(33, L':');
        for (int i = 0; i < 16; ++i) {
            text[static_cast<std::size_t>(15 - i)] = digits[(baseline.hash >> (4 * i)) & 0xF];
            text[static_cast<std::size_t>(32 - i)] = digits[(baseline.legacy >> (4 * i)) & 0xF];
        }
        return text;
    }

    // Reads toText(). False if text is no such value, or stands beside another legacy hash than stored: then a
    // version before 6.3 has set the baseline since and text is out of date
    inline bool fromText(const std::wstring& text, std::uint64_t stored, Baseline& baseline)
    {
        std::uint64_t hash = 0, legacyHash = 0;
        if (text.size() != 33 || text[16] != L':') return false;
        if (!detail::hexValue(text, 0, hash) || !detail::hexValue(text, 17, legacyHash)) return false;
        if (legacyHash != stored) return false;
        baseline = { hash, legacyHash };
        return true;
    }

    // The baseline of a tab from its INI values: text if it is up to date, else worked out from the legacy hash.
    // loadFile fills in the entries of the tab's list file; false if the tab has none or it cannot be read
    template <class LoadFile>
    Baseline fromIni(const std::wstring& text, std::uint64_t stored, const std::vector<ReplaceItemData>& list, LoadFile&& loadFile)
    {
        Baseline baseline;
        if (fromText(text, stored, baseline)) return baseline;

        std::vector<ReplaceItemData> file;
        const bool hasFile = loadFile(file);
        return rebase(stored, list, hasFile ? &file : nullptr);
    }

    // Baseline of an untitled tab: the versions before 6.3 set it to the list at exit and kept the dirty state only in
    // a flag, so a tab flagged dirty whose list matches gets the baseline of a new tab back, the empty list
    inline Baseline untitledBaseline(const Baseline& loaded, const std::vector<ReplaceItemData>& list, bool flaggedDirty)
    {
        return flaggedDirty && loaded.matches(list) ? baselineOf({}) : loaded;
    }
}
