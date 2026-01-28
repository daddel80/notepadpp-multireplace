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

#pragma once

// -----------------------------------------------------------------------------
// NumericToken.h
// -----------------------------------------------------------------------------
// Purpose:
//   String-based parser for numeric fields in CSV data.
//   Used for CSV column sorting and decimal-point alignment.
//   Locale-free: ASCII digits only, '.' and ',' as decimal separators.
//
// Recognized numeric fields:
//   [prefix] [sign] DIGITS [sep DIGITS] [suffix]
//   [prefix] [sign] sep DIGITS [suffix]           (leading decimal: .5, -.5)
//   
//   Examples:  "123", "-45.67", ".5", "-.5", "$100", "100EUR", "€50.00"
//
// Rules:
//   - Prefix/suffix max 4 chars, must be symbols ($€£) OR letters (USD, EUR)
//   - Affixes can be adjacent or space-separated: "$100", "100EUR", "100 EUR"
//   - Mixed letters+symbols in affix not allowed: "$USD" fails
//
// Internal normalization (for sorting):
//   ',' -> '.', ".5" -> "0.5", "-.5" -> "-0.5", "12." -> "12"
// -----------------------------------------------------------------------------

#include <string>
#include <string_view>
#include <cstddef>

namespace mr {
    namespace num {

        struct NumericToken {
            bool        ok = false;
            std::size_t start = 0;      // inclusive index in input
            std::size_t end = 0;      // exclusive index in input
            bool        hasSign = false;
            bool        hasDecimal = false;
            int         intDigits = 0;      // count of digits before decimal
            std::string normalized;          // normalized ASCII form (e.g. "-123.45")
            double      value = 0.0;    // parsed numeric value
        };

        // Classifies a trimmed field as numeric.
        // Returns true if field contains exactly one numeric token with optional
        // prefix/suffix (e.g. currency symbols). outTok receives the parsed token.
        // Input must be pre-trimmed (no leading/trailing whitespace).
        bool classify_numeric_field(std::string_view trimmedField, NumericToken& outTok);

    }
} // namespace mr::num
