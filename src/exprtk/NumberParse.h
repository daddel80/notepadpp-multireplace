// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// NumberParse.h
// Locale-independent text-to-double parser shared between the engine
// (where it interprets regex capture text for num()) and the match
// history module (where it interprets cached capture text for the
// same num() reads at non-zero lookback). Lives in its own header so
// neither caller has to depend on the other.
//
// The parser is permissive in the same way the existing engine has
// always been: it accepts both '.' and ',' as the decimal separator,
// tolerates trailing junk after a leading number, and never throws -
// non-numeric input becomes NaN so the formula author can either
// propagate it or guard with an isfinite() check.

#pragma once

#include <string>

namespace MultiReplaceEngine {

    // Parse `s` as a double. Returns:
    //   - NaN for empty input.
    //   - NaN if the leading characters do not form a number.
    //   - NaN if the input has more than one separator ('.' or ','),
    //     whether mixed ("1.234,56") or repeated ("1.000.000"): the
    //     intended value is ambiguous (thousands grouping vs. decimal).
    //   - The parsed value otherwise (trailing junk after a valid
    //     numeric prefix is silently dropped: "1.5abc" -> 1.5).
    //
    // Decimal separator: a single separator type is accepted. '.' is
    // the decimal point; if only ',' is present it is treated as the
    // decimal point ("1,5" -> 1.5) for European-locale data, since the
    // regex pipeline upstream cannot know which convention is in use.
    // Thousands separators are not supported - strip them in the regex.
    double parseNumber(const std::string& s);

} // namespace MultiReplaceEngine