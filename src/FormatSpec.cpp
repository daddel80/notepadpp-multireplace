// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.

#include "FormatSpec.h"

#include <array>
#include <charconv>
#include <cmath>
#include <cstdio>
#include <ctime>
#include <cwchar>
#include <string>

namespace FormatSpec {

    namespace {

        // ASCII-only trim for the small spec strings we parse.
        std::wstring trim(const std::wstring& s) {
            size_t a = 0, b = s.size();
            while (a < b && (s[a] == L' ' || s[a] == L'\t')) ++a;
            while (b > a && (s[b - 1] == L' ' || s[b - 1] == L'\t')) --b;
            return s.substr(a, b - a);
        }

        bool isDigit(wchar_t c) { return c >= L'0' && c <= L'9'; }

        // Read an unsigned decimal integer at index i. Advances i past
        // the digits and returns the value. Returns -1 if no digit at i.
        int readInt(const std::wstring& s, size_t& i) {
            if (i >= s.size() || !isDigit(s[i])) return -1;
            int v = 0;
            while (i < s.size() && isDigit(s[i])) {
                v = v * 10 + (s[i] - L'0');
                ++i;
            }
            return v;
        }

        Spec makeError(const std::string& msg) {
            Spec s;
            s.valid = false;
            s.errorMessage = msg;
            return s;
        }

        // ---- Duration parser ------------------------------------------
        // Form: t<unit>:<mode>. Both parts mandatory in v1.
        Spec parseDuration(const std::wstring& s) {
            // s starts with 't' here.
            if (s.size() < 4) {
                return makeError("Duration spec needs unit and mode (e.g. ts:hms)");
            }

            Spec out;
            out.kind = Kind::Duration;

            switch (s[1]) {
            case L's': out.durationUnit = DurationUnit::Seconds; break;
            case L'm': out.durationUnit = DurationUnit::Minutes; break;
            case L'h': out.durationUnit = DurationUnit::Hours;   break;
            case L'd': out.durationUnit = DurationUnit::Days;    break;
            default:
                return makeError("Unknown duration unit; expected ts/tm/th/td");
            }

            if (s[2] != L':') {
                return makeError("Duration spec needs ':<mode>' after the unit");
            }

            const std::wstring mode = s.substr(3);
            if (mode == L"ms")   out.durationMode = DurationMode::Ms;
            else if (mode == L"hm")   out.durationMode = DurationMode::Hm;
            else if (mode == L"hms")  out.durationMode = DurationMode::Hms;
            else if (mode == L"dh")   out.durationMode = DurationMode::Dh;
            else if (mode == L"dhm")  out.durationMode = DurationMode::Dhm;
            else if (mode == L"dhms") out.durationMode = DurationMode::Dhms;
            else {
                return makeError("Unknown duration mode; expected ms/hm/hms/dh/dhm/dhms");
            }

            out.valid = true;
            return out;
        }

        // ---- Numeric parser -------------------------------------------
        // Form: [+] [0] [width] [.prec | .min-max] [type]
        Spec parseNumeric(const std::wstring& s) {
            Spec out;
            out.kind = Kind::Numeric;

            size_t i = 0;

            if (i < s.size() && s[i] == L'+') { out.forceSign = true; ++i; }
            if (i < s.size() && s[i] == L'0') { out.zeroPad = true; ++i; }

            // Width is the next integer (if any). A leading zero already
            // consumed above does NOT contribute to width.
            if (i < s.size() && isDigit(s[i])) {
                out.width = readInt(s, i);
            }

            // Precision block: '.' followed by either N or min-max.
            if (i < s.size() && s[i] == L'.') {
                ++i;
                int p = readInt(s, i);
                if (p < 0) {
                    return makeError("Precision '.' must be followed by a digit");
                }
                out.precisionMin = p;
                if (i < s.size() && s[i] == L'-') {
                    ++i;
                    int pmax = readInt(s, i);
                    if (pmax < 0) {
                        return makeError("Range precision '.min-' must be followed by max digits");
                    }
                    if (pmax < p) {
                        return makeError("Range precision: max must be >= min");
                    }
                    out.precisionMax = pmax;
                }
            }

            // Type letter, optional. If missing, NumericType::Default
            // (= shortest round-trip via std::to_chars).
            if (i < s.size()) {
                if (i + 1 != s.size()) {
                    return makeError("Unexpected trailing characters in numeric spec");
                }
                switch (s[i]) {
                case L'f': out.numericType = NumericType::Fixed;       break;
                case L'e': out.numericType = NumericType::Scientific;  break;
                case L'g': out.numericType = NumericType::General;     break;
                case L'x': out.numericType = NumericType::Hex;         break;
                case L'b': out.numericType = NumericType::Binary;      break;
                case L'o': out.numericType = NumericType::Octal;       break;
                default:
                    return makeError("Unknown numeric type letter; expected f/e/g/x/b/o");
                }
            }
            else {
                // No type letter, but only legal if at least one modifier
                // was given. Empty spec is the caller's responsibility.
                if (!out.forceSign && !out.zeroPad
                    && out.width < 0 && out.precisionMin < 0) {
                    return makeError("Empty numeric spec");
                }
            }

            // Integer-output types (x/b/o) reject precision and force-sign
            // because they don't apply meaningfully to integers.
            const bool isIntegerType =
                (out.numericType == NumericType::Hex
                    || out.numericType == NumericType::Binary
                    || out.numericType == NumericType::Octal);
            if (isIntegerType) {
                if (out.precisionMin >= 0) {
                    return makeError("Precision not allowed with integer types (x/b/o)");
                }
                if (out.forceSign) {
                    return makeError("'+' flag not allowed with integer types (x/b/o)");
                }
            }

            out.valid = true;
            return out;
        }

        // ---- Numeric renderer -----------------------------------------

        // Build a printf-style format string from the parsed Spec, then
        // run snprintf. We use printf rather than std::format because
        // std::format does not support the 'b' (binary) and trims the
        // .min-max precision form differently. printf gives us a single
        // backend that covers f/e/g/x/o uniformly; binary is handled
        // manually below.
        //
        // Locale: snprintf("%f", ...) honors the C locale, not the user
        // locale, so we get dot-decimal regardless of system settings.
        std::wstring formatPrintf(const Spec& spec, double value) {
            std::string fmt = "%";
            if (spec.forceSign) fmt += "+";
            if (spec.zeroPad)   fmt += "0";
            if (spec.width >= 0) fmt += std::to_string(spec.width);

            // Precision: .N or .max (we apply min separately below for
            // trailing-zero trimming).
            const int pmax = (spec.precisionMax >= 0) ? spec.precisionMax : spec.precisionMin;
            if (pmax >= 0) {
                fmt += "." + std::to_string(pmax);
            }

            switch (spec.numericType) {
            case NumericType::Fixed:      fmt += "f"; break;
            case NumericType::Scientific: fmt += "e"; break;
            case NumericType::General:    fmt += "g"; break;
            case NumericType::Hex:        fmt = "%";
                if (spec.zeroPad) fmt += "0";
                if (spec.width >= 0) fmt += std::to_string(spec.width);
                fmt += "llx";
                break;
            case NumericType::Octal:      fmt = "%";
                if (spec.zeroPad) fmt += "0";
                if (spec.width >= 0) fmt += std::to_string(spec.width);
                fmt += "llo";
                break;
            case NumericType::Binary:
                // Handled separately below.
                break;
            case NumericType::Default:
                // No type letter: shortest round-trip via to_chars,
                // then pad/sign manually if requested.
            {
                std::array<char, 64> buf{};
                auto res = std::to_chars(buf.data(), buf.data() + buf.size(), value);
                std::string body(buf.data(), res.ptr);
                if (spec.forceSign && !body.empty() && body[0] != '-' && body[0] != '+') {
                    body.insert(body.begin(), '+');
                }
                if (spec.width > 0 && static_cast<int>(body.size()) < spec.width) {
                    const char pad = spec.zeroPad ? '0' : ' ';
                    const int need = spec.width - static_cast<int>(body.size());
                    if (spec.zeroPad && !body.empty() && (body[0] == '-' || body[0] == '+')) {
                        body.insert(1, need, pad);
                    }
                    else {
                        body.insert(0, need, pad);
                    }
                }
                return std::wstring(body.begin(), body.end());
            }
            }

            // Binary path. Integer cast (negative values get two's
            // complement of the 64-bit representation).
            if (spec.numericType == NumericType::Binary) {
                long long iv = static_cast<long long>(value);
                unsigned long long u = static_cast<unsigned long long>(iv);
                std::string body;
                if (u == 0) body = "0";
                else {
                    while (u != 0) { body.insert(body.begin(), '0' + (u & 1)); u >>= 1; }
                }
                if (spec.width > 0 && static_cast<int>(body.size()) < spec.width) {
                    const char pad = spec.zeroPad ? '0' : ' ';
                    body.insert(0, spec.width - body.size(), pad);
                }
                return std::wstring(body.begin(), body.end());
            }

            // printf path for f/e/g/x/o.
            std::array<char, 128> buf{};
            int n;
            if (spec.numericType == NumericType::Hex
                || spec.numericType == NumericType::Octal) {
                long long iv = static_cast<long long>(value);
                n = std::snprintf(buf.data(), buf.size(), fmt.c_str(),
                    static_cast<unsigned long long>(iv));
            }
            else {
                n = std::snprintf(buf.data(), buf.size(), fmt.c_str(), value);
            }
            if (n < 0) return L"";
            std::string body(buf.data(), buf.data() + (std::min<size_t>(n, buf.size() - 1)));

            // Apply min-side trailing-zero trim for the .min-max form.
            // printf gave us pmax decimals; we strip trailing zeros down
            // to pmin (but never further), and a trailing dot too.
            if (spec.precisionMax >= 0 && spec.precisionMin >= 0
                && (spec.numericType == NumericType::Fixed
                    || spec.numericType == NumericType::Scientific
                    || spec.numericType == NumericType::General)) {
                // Find the decimal point. Look for '.' before any 'e'/'E'.
                size_t e = body.find_first_of("eE");
                size_t end = (e == std::string::npos) ? body.size() : e;
                size_t dot = body.rfind('.', end == 0 ? 0 : end - 1);
                if (dot != std::string::npos && dot < end) {
                    size_t maxStrip = end - (dot + 1) - static_cast<size_t>(spec.precisionMin);
                    size_t stripped = 0;
                    while (stripped < maxStrip && body[end - 1 - stripped] == '0') {
                        ++stripped;
                    }
                    // If we'd be left with a bare '.', strip that too.
                    if (stripped > 0 && body[end - 1 - stripped] == '.'
                        && static_cast<int>(end - dot - 1 - stripped) == 0) {
                        ++stripped;
                    }
                    body.erase(end - stripped, stripped);
                }
            }

            return std::wstring(body.begin(), body.end());
        }

        // ---- Duration renderer ----------------------------------------

        // Split a value in spec.durationUnit into days/hours/minutes/
        // seconds, then format by spec.durationMode.
        std::wstring formatDuration(const Spec& spec, double value) {
            // Convert to seconds (the common base).
            double seconds = value;
            switch (spec.durationUnit) {
            case DurationUnit::Seconds: break;
            case DurationUnit::Minutes: seconds *= 60.0;        break;
            case DurationUnit::Hours:   seconds *= 3600.0;      break;
            case DurationUnit::Days:    seconds *= 86400.0;     break;
            }

            const bool negative = seconds < 0.0;
            if (negative) seconds = -seconds;

            long long totalSec = static_cast<long long>(seconds);
            long long days = totalSec / 86400;
            long long rem = totalSec % 86400;
            long long hours = rem / 3600;
            rem %= 3600;
            long long mins = rem / 60;
            long long secs = rem % 60;

            char buf[64];
            switch (spec.durationMode) {
            case DurationMode::Ms:
                // M:SS - minutes can grow large; no day/hour folding.
                std::snprintf(buf, sizeof(buf), "%lld:%02lld",
                    totalSec / 60, totalSec % 60);
                break;
            case DurationMode::Hms:
                // H:MM:SS - hours can grow large.
                std::snprintf(buf, sizeof(buf), "%lld:%02lld:%02lld",
                    totalSec / 3600, (totalSec % 3600) / 60, totalSec % 60);
                break;
            case DurationMode::Hm:
                // H:MM - hours can grow large; seconds dropped (rounded
                // down by the integer cast above, no extra rounding here).
                std::snprintf(buf, sizeof(buf), "%lld:%02lld",
                    totalSec / 3600, (totalSec % 3600) / 60);
                break;
            case DurationMode::Dh:
                std::snprintf(buf, sizeof(buf), "%lld %02lld",
                    days, hours);
                break;
            case DurationMode::Dhm:
                std::snprintf(buf, sizeof(buf), "%lld %02lld:%02lld",
                    days, hours, mins);
                break;
            case DurationMode::Dhms:
                std::snprintf(buf, sizeof(buf), "%lld %02lld:%02lld:%02lld",
                    days, hours, mins, secs);
                break;
            }

            std::string body = buf;
            if (negative) body.insert(body.begin(), '-');
            return std::wstring(body.begin(), body.end());
        }

        // ---- Date parser ----------------------------------------------
        // Form: D[<strftime fmt>] or D[!<strftime fmt>]
        Spec parseDate(const std::wstring& s) {
            // s starts with 'D' and was already trimmed.
            if (s.size() < 3 || s[1] != L'[' || s.back() != L']') {
                return makeError("Date spec must be D[<format>], e.g. D[%Y-%m-%d]");
            }

            Spec out;
            out.kind = Kind::Date;

            // Strip the D[ ... ] brackets.
            std::wstring body = s.substr(2, s.size() - 3);

            // Lua-style '!' prefix forces UTC.
            if (!body.empty() && body[0] == L'!') {
                out.dateUtc = true;
                body.erase(body.begin());
            }

            if (body.empty()) {
                return makeError("Date spec is empty (need a strftime pattern inside D[...])");
            }

            // Convert wstring -> utf8 narrow for strftime. The strftime
            // patterns we expect are ASCII (%Y, %m, %d, separators), so
            // a direct byte cast is enough; non-ASCII chars would only
            // appear as literal output text, which strftime passes
            // through verbatim.
            out.dateFormat.reserve(body.size());
            for (wchar_t wc : body) {
                out.dateFormat.push_back(static_cast<char>(wc));
            }

            out.valid = true;
            return out;
        }

        // ---- Date renderer --------------------------------------------
        // Treats `value` as a Unix timestamp (seconds since 1970-01-01
        // UTC). Negative timestamps and out-of-range conversions are
        // refused; strftime is asked for either gm-time or local-time
        // depending on spec.dateUtc.
        std::wstring formatDate(const Spec& spec, double value) {
            if (!(value >= 0.0)) {
                // Negative or NaN. (NaN is already filtered upstream,
                // but we guard defensively here too.)
                return L"";
            }
            // Drop subseconds; strftime only consumes integer seconds.
            // Cap to a sane range so we don't hit time_t overflow on
            // 32-bit ABIs.
            constexpr double kMaxSeconds = 253402300800.0;  // 9999-12-31
            if (value > kMaxSeconds) {
                return L"";
            }
            std::time_t t = static_cast<std::time_t>(value);

            std::tm tm{};
#if defined(_WIN32)
            // MSVC: gmtime_s / localtime_s return errno_t (0 on success).
            const int rc = spec.dateUtc ? gmtime_s(&tm, &t)
                : localtime_s(&tm, &t);
            if (rc != 0) {
                return L"";
            }
#else
            // POSIX: gmtime_r / localtime_r return a pointer (NULL on err).
            std::tm* res = spec.dateUtc ? gmtime_r(&t, &tm)
                : localtime_r(&t, &tm);
            if (res == nullptr) {
                return L"";
            }
#endif

            // strftime may produce more bytes than the pattern length
            // (full month names, %c output etc.). 256 covers any
            // realistic single-line date pattern.
            char buf[256];
            const std::size_t n = std::strftime(buf, sizeof(buf),
                spec.dateFormat.c_str(),
                &tm);
            if (n == 0) {
                // Either the pattern was empty (caught at parse time)
                // or the output exceeded the buffer; surface as empty
                // rather than truncating silently.
                return L"";
            }

            std::wstring out;
            out.reserve(n);
            for (std::size_t i = 0; i < n; ++i) {
                out.push_back(static_cast<wchar_t>(static_cast<unsigned char>(buf[i])));
            }
            return out;
        }

    }  // unnamed namespace

    Spec parse(const std::wstring& specText) {
        std::wstring s = trim(specText);

        if (s.empty()) {
            return makeError("Empty spec");
        }

        // Date: D[...] or D[!...]
        if (s[0] == L'D' && s.size() >= 2 && s[1] == L'[') {
            return parseDate(s);
        }

        // Duration: starts with 't' followed by a unit letter and ':'.
        if (s[0] == L't') {
            return parseDuration(s);
        }

        return parseNumeric(s);
    }

    std::wstring apply(const Spec& spec, double value) {
        if (!spec.valid) return L"";

        switch (spec.kind) {
        case Kind::Date:     return formatDate(spec, value);
        case Kind::Duration: return formatDuration(spec, value);
        case Kind::Numeric:  return formatPrintf(spec, value);
        }
        return L"";
    }

    // ---- Formula / spec splitter --------------------------------------

    namespace {
        // ASCII trim for narrow string (used by the splitter).
        std::string trimNarrow(const std::string& s) {
            size_t a = 0, b = s.size();
            while (a < b && (s[a] == ' ' || s[a] == '\t')) ++a;
            while (b > a && (s[b - 1] == ' ' || s[b - 1] == '\t')) --b;
            return s.substr(a, b - a);
        }
    }

    Split splitFormulaSpec(const std::string& blockBody) {
        // Forward scan tracking quote state and bracket depth. We
        // remember the position of the last '~' that occurs while we
        // are NOT inside a quoted string and at bracket depth 0.
        //
        // ExprTk string literals use either single or double quotes.
        // Quoted strings can contain backslash escapes; we handle the
        // typical \' \" \\ cases so an escaped quote does not flip
        // the quote state.
        //
        // Brackets: square brackets [...] are tracked so that future
        // date specs like D[%Y-%m-%d ~ %H:%M] do not get split inside.
        // Round and curly brackets do not need tracking because no '~'
        // can legitimately appear inside an ExprTk call or list.
        long long lastTilde = -1;

        wchar_t quoteCh = 0;        // 0 = not in quote; otherwise the opening quote
        int bracketDepth = 0;

        for (size_t i = 0; i < blockBody.size(); ++i) {
            char c = blockBody[i];

            if (quoteCh != 0) {
                if (c == '\\' && i + 1 < blockBody.size()) {
                    ++i;  // skip the escaped char
                    continue;
                }
                if (c == (char)quoteCh) {
                    quoteCh = 0;
                }
                continue;
            }

            if (c == '\'' || c == '"') {
                quoteCh = (wchar_t)c;
                continue;
            }
            if (c == '[') { ++bracketDepth; continue; }
            if (c == ']') { if (bracketDepth > 0) --bracketDepth; continue; }

            if (c == '~' && bracketDepth == 0) {
                lastTilde = static_cast<long long>(i);
            }
        }

        Split result;
        if (lastTilde < 0) {
            // No separator - whole body is the formula.
            result.formula = trimNarrow(blockBody);
            result.hasSpec = false;
            return result;
        }

        result.formula = trimNarrow(blockBody.substr(0, static_cast<size_t>(lastTilde)));
        result.spec = trimNarrow(blockBody.substr(static_cast<size_t>(lastTilde) + 1));
        result.hasSpec = true;
        return result;
    }

}  // namespace FormatSpec