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
    //   - The parsed value otherwise (trailing junk after a valid
    //     numeric prefix is silently dropped: "1.5abc" -> 1.5).
    //
    // Decimal separator: if the input contains '.', commas are treated
    // as part of trailing junk; if it does not contain '.', any ','
    // becomes the decimal point. This matches the long-standing
    // behaviour of the engine on the kind of European-locale data
    // ("1,5" meaning 1.5) where the regex pipeline upstream cannot
    // know which convention is in use.
    double parseNumber(const std::string& s);

} // namespace MultiReplaceEngine