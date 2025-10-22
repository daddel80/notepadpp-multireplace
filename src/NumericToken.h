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
//   String-based parser for the *first* numeric token inside a field.
//   Shared between ColumnTabs (alignment) and MultiReplace (sorting).
//   Locale-free: ASCII digits only, '.' and ',' as decimal separators.
//
// Recognized token patterns (first match wins):
//   [sign] DIGITS [ ('.'|',') DIGITS? ]     e.g. "-12", "+300.34", "12.", "66,1"
//   [sign]? ('.'|',') DIGITS                e.g. ".5", "-.75"
// Everything before/after the token is ignored (prefix/suffix like currency).
// Normalization:
//   ',' -> '.', ".5" -> "0.5", "-.5" -> "-0.5", "12." -> "12"
// -----------------------------------------------------------------------------

#include <string>
#include <string_view>
#include <cstddef>

namespace mr { namespace num {

    struct ParseOptions {
        bool allowLeadingSeparator = true; // accept ".5" / ",5"
        int  maxCurrencyAffix = 4;         // reserved (not enforced yet)
    };

    struct NumericToken {
        bool        ok          = false;
        std::size_t start       = 0;     // inclusive
        std::size_t end         = 0;     // exclusive
        bool        hasSign     = false;
        bool        hasDecimal  = false;
        int         intDigits   = 0;
        std::string normalized;          // normalized ASCII form
        double      value       = 0.0;   // parsed from normalized
    };

    // Parse first numeric token; returns { ok=true } on success.
    NumericToken parse_first_numeric_token(std::string_view field,
                                           const ParseOptions& opt = {});

    // Convenience: parse only the value (and optionally the normalized token).
    bool try_parse_first_numeric_value(std::string_view field,
                                       double& outValue,
                                       std::string* outNormalized = nullptr,
                                       const ParseOptions& opt = {});

}} // namespace mr::num
