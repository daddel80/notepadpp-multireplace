// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// NumberParse.cpp
// Implementation lifted verbatim from the engine's original
// parseCaptureToDouble so existing behaviour is preserved bit-for-bit.
// ExprTkEngine::parseCaptureToDouble now forwards here; this is the
// single source of truth.

#include "NumberParse.h"

#include <algorithm>
#include <charconv>
#include <limits>
#include <system_error>

namespace MultiReplaceEngine {

    double parseNumber(const std::string& s)
    {
        if (s.empty()) {
            return std::numeric_limits<double>::quiet_NaN();
        }

        // Exactly one decimal separator is allowed. Anything with more
        // than one separator - mixed ('.' and ',') or repeated ('1.000.000',
        // '1,000,000') - is a thousands-grouped or otherwise ambiguous form
        // that cannot be read as a single number, so it yields NaN rather
        // than a misleading partial value. Thousands separators are not
        // supported - strip them in the regex first.
        const long long sepCount =
            std::count(s.begin(), s.end(), '.') +
            std::count(s.begin(), s.end(), ',');
        if (sepCount > 1) {
            return std::numeric_limits<double>::quiet_NaN();
        }

        // Rewrite a lone ',' to '.' for the comma-decimal path; a working
        // buffer is built solely for that case. Plain '.' numerics and
        // integers parse directly from the input - no allocation.
        std::string buf;
        const char* first = s.data();
        const char* last = s.data() + s.size();

        if (s.find(',') != std::string::npos) {
            buf.assign(s);
            std::replace(buf.begin(), buf.end(), ',', '.');
            first = buf.data();
            last = buf.data() + buf.size();
        }

        double value = 0.0;
        auto res = std::from_chars(first, last, value);
        if (res.ec != std::errc{}) {
            return std::numeric_limits<double>::quiet_NaN();
        }

        // Trailing junk after a valid numeric prefix is consumed silently.
        // "1.5abc" -> 1.5. This matches how most programming languages
        // treat a number embedded in a larger string.
        return value;
    }

} // namespace MultiReplaceEngine