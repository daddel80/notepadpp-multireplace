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

        // ---- Text parser ----------------------------------------------
        // Form: [fill][align][width][.maxlen]
        // Caller has already stripped the leading 't:' marker, so 's'
        // here is the body. An empty body is rejected because at least
        // one component (width, align, ...) needs to be present for the
        // spec to do anything.
        //
        // Fill / align disambiguation: a leading fill character is only
        // recognised if the second character is an align letter
        // ('<', '>', '^'). Without that lookahead, a spec like 't:*15'
        // would be ambiguous - is '*' a fill or just a stray character?
        // Requiring align after fill makes the intent explicit.
        //
        // Fill is captured as the leading UTF-8 codepoint (1-4 bytes),
        // not a single byte, so non-ASCII fill characters (umlauts,
        // accented letters, etc.) work even though they take two bytes
        // in UTF-8.
        Spec parseString(const std::wstring& s) {
            if (s.empty()) {
                return makeError("Text spec is empty (need width, align, or fill+align after 't:')");
            }

            Spec out;
            out.kind = Kind::Text;

            size_t i = 0;

            // Decide whether the leading codepoint is a fill character.
            // It is, only when the next codepoint is an align letter.
            auto isAlignChar = [](wchar_t c) {
                return c == L'<' || c == L'>' || c == L'^';
                };

            // Lookahead: fill takes the leading codepoint only if an
            // align letter follows it. This handles three cases cleanly:
            //   '*<'  -> fill '*', align left   (ordinary fill)
            //   '<<'  -> fill '<', align left   (align-char used as fill)
            //   '<5'  -> no fill, align left, width 5
            // With just one character available, no fill is possible.
            if (s.size() >= 2 && isAlignChar(s[1])) {
                // Fill is the first codepoint. Encode it as UTF-8 so it
                // can be appended cheaply during rendering.
                const wchar_t fillCp = s[0];
                out.textFill.clear();
                if (fillCp < 0x80) {
                    out.textFill.push_back(static_cast<char>(fillCp));
                }
                else if (fillCp < 0x800) {
                    out.textFill.push_back(static_cast<char>(0xC0 | (fillCp >> 6)));
                    out.textFill.push_back(static_cast<char>(0x80 | (fillCp & 0x3F)));
                }
                else {
                    out.textFill.push_back(static_cast<char>(0xE0 | (fillCp >> 12)));
                    out.textFill.push_back(static_cast<char>(0x80 | ((fillCp >> 6) & 0x3F)));
                    out.textFill.push_back(static_cast<char>(0x80 | (fillCp & 0x3F)));
                }
                i = 1;  // s[1] is an align letter (guaranteed by the lookahead)
            }

            // Align letter, optional. Defaults to Left when omitted. If a
            // fill was consumed above, i points at the align letter the
            // lookahead already verified, so this branch always matches there.
            if (i < s.size() && isAlignChar(s[i])) {
                switch (s[i]) {
                case L'<': out.textAlign = TextAlign::Left;   break;
                case L'>': out.textAlign = TextAlign::Right;  break;
                case L'^': out.textAlign = TextAlign::Center; break;
                }
                ++i;
            }

            // Width, optional.
            if (i < s.size() && isDigit(s[i])) {
                out.width = readInt(s, i);
            }

            // Max-length, optional. '.' followed by a non-negative integer.
            if (i < s.size() && s[i] == L'.') {
                ++i;
                int m = readInt(s, i);
                if (m < 0) {
                    return makeError("Text spec '.' must be followed by a digit");
                }
                out.textMaxLength = m;
            }

            // Anything left over is a parse error.
            if (i != s.size()) {
                return makeError("Unexpected trailing characters in text spec");
            }

            out.valid = true;
            return out;
        }

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
                case L'd': out.numericType = NumericType::Integer;     break;
                default:
                    return makeError("Unknown numeric type letter; expected f/e/g/x/b/o/d");
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

            // Precision applies to floating-point output only. Integer
            // outputs (x/b/o/d) reject it because there's no fractional
            // part to format.
            const bool acceptsPrecision =
                (out.numericType == NumericType::Default
                    || out.numericType == NumericType::Fixed
                    || out.numericType == NumericType::Scientific
                    || out.numericType == NumericType::General);
            if (!acceptsPrecision && out.precisionMin >= 0) {
                return makeError("Precision not allowed with integer types (x/b/o/d)");
            }

            // Force-sign is meaningful for signed outputs. Hex/binary/octal
            // are formatted as unsigned bit patterns (two's complement for
            // negatives), so '+' has no sensible meaning there; signed
            // base-10 ('d') keeps it.
            const bool acceptsForceSign =
                (out.numericType != NumericType::Hex
                    && out.numericType != NumericType::Binary
                    && out.numericType != NumericType::Octal);
            if (!acceptsForceSign && out.forceSign) {
                return makeError("'+' flag not allowed with integer types (x/b/o)");
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
            case NumericType::Integer:    fmt += "lld"; break;
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

            // printf path for f/e/g/x/o/d.
            std::array<char, 128> buf{};
            int n;
            if (spec.numericType == NumericType::Hex
                || spec.numericType == NumericType::Octal) {
                long long iv = static_cast<long long>(value);
                n = std::snprintf(buf.data(), buf.size(), fmt.c_str(),
                    static_cast<unsigned long long>(iv));
            }
            else if (spec.numericType == NumericType::Integer) {
                n = std::snprintf(buf.data(), buf.size(), fmt.c_str(),
                    static_cast<long long>(value));
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

            char buf[64] = {};
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

        // ---- Text renderer --------------------------------------------
        //
        // Codepoint-aware so that multi-byte UTF-8 sequences count as one
        // character. Width pads with spec.textFill on the appropriate
        // side(s); max-length truncates without ever splitting a
        // codepoint mid-sequence.
        //
        // UTF-8 lead-byte test: continuation bytes have the bit pattern
        // 10xxxxxx, so anything that does NOT match (byte & 0xC0) == 0x80
        // is a codepoint start. That's all the decoder we need for
        // counting and slicing.
        std::wstring formatString(const Spec& spec, const std::string& text) {
            auto isLeadByte = [](unsigned char b) {
                return (b & 0xC0) != 0x80;
                };

            // Count codepoints in the input.
            std::size_t codepointCount = 0;
            for (unsigned char b : text) {
                if (isLeadByte(b)) ++codepointCount;
            }

            // Truncate to textMaxLength codepoints, slicing at a lead-byte
            // boundary. byteCutoff sits at the start of the (maxLen+1)-th
            // codepoint, i.e. the first byte we drop.
            std::size_t byteCutoff = text.size();
            if (spec.textMaxLength >= 0
                && codepointCount > static_cast<std::size_t>(spec.textMaxLength)) {
                std::size_t cpSeen = 0;
                for (std::size_t i = 0; i < text.size(); ++i) {
                    if (isLeadByte(static_cast<unsigned char>(text[i]))) {
                        if (cpSeen == static_cast<std::size_t>(spec.textMaxLength)) {
                            byteCutoff = i;
                            break;
                        }
                        ++cpSeen;
                    }
                }
                codepointCount = static_cast<std::size_t>(spec.textMaxLength);
            }
            const std::string body(text.data(), byteCutoff);

            // Pad to spec.width codepoints. If already at or above width,
            // emit as-is.
            std::string out;
            const int width = spec.width;
            if (width > 0 && codepointCount < static_cast<std::size_t>(width)) {
                const std::size_t padCount =
                    static_cast<std::size_t>(width) - codepointCount;

                // Compute left/right pad counts based on alignment.
                std::size_t padLeft = 0;
                std::size_t padRight = 0;
                switch (spec.textAlign) {
                case TextAlign::Left:   padRight = padCount; break;
                case TextAlign::Right:  padLeft = padCount; break;
                case TextAlign::Center:
                    padLeft = padCount / 2;
                    padRight = padCount - padLeft;
                    break;
                }

                out.reserve(text.size() + padCount * spec.textFill.size());
                for (std::size_t k = 0; k < padLeft; ++k)  out.append(spec.textFill);
                out.append(body);
                for (std::size_t k = 0; k < padRight; ++k) out.append(spec.textFill);
            }
            else {
                out = body;
            }

            // Convert UTF-8 bytes to wstring. Each byte maps to one
            // wchar; the engine immediately collapses these back to char
            // on the output path. We could decode to real wide
            // codepoints, but that's redundant work for an output that
            // ends up as a byte stream anyway.
            return std::wstring(out.begin(), out.end());
        }

        // Form: d:<strftime fmt> or d:!<strftime fmt>
        // Caller has already verified that s starts with "d:".
        Spec parseDate(const std::wstring& s) {
            Spec out;
            out.kind = Kind::Date;

            // Strip the "d:" prefix.
            std::wstring body = s.substr(2);

            // Lua-style '!' prefix forces UTC.
            if (!body.empty() && body[0] == L'!') {
                out.dateUtc = true;
                body.erase(body.begin());
            }

            if (body.empty()) {
                return makeError("Date spec is empty (need a strftime pattern after 'd:')");
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

    // ---- Top-level dispatch -------------------------------------------
    //
    // Spec families are distinguished by the leading character (plus a
    // single lookahead where needed):
    //
    //   'd:'                  -> Date
    //   'n:'                  -> Numeric (explicit marker; equivalent to bare)
    //   't:'                  -> Text
    //   't' + unit letter     -> Duration (ts/tm/th/td)
    //   anything else         -> Numeric (default family, no marker)
    //
    // Numeric is the unmarked default so that the common case stays
    // short (`5.2f`, `04x`). The optional `n:` form is offered for
    // callers who prefer a marker on every spec for visual consistency
    // with `d:`, `t:` etc. Both forms produce identical output.
    //
    // The 't' prefix is shared between Text and Duration: the second
    // character decides. ':' -> Text body, unit letter -> Duration. This
    // lookahead lets txt() consumers map to a `t:` marker without
    // colliding with the existing `ts:` / `tm:` etc. specs.
    //
    // Other families wear a leading marker so their grammars can grow
    // independently without retroactively ambiguating the numeric
    // parser.
    Spec parse(const std::wstring& specText) {
        std::wstring s = trim(specText);

        if (s.empty()) {
            return makeError("Empty spec");
        }

        // Date: d:<fmt>  (also d:!<fmt> for UTC)
        if (s.size() >= 2 && s[0] == L'd' && s[1] == L':') {
            return parseDate(s);
        }

        // Numeric (explicit marker): strip 'n:' and continue as if it
        // had been bare. The trailing part still has to satisfy the
        // normal numeric grammar - 'n:' on its own is rejected as an
        // empty numeric spec.
        if (s.size() >= 2 && s[0] == L'n' && s[1] == L':') {
            return parseNumeric(s.substr(2));
        }

        // 't' family: text (t:...) or duration (t<unit>:...). The second
        // character disambiguates - ':' takes us to the text parser,
        // anything else stays on the duration path.
        if (s[0] == L't') {
            if (s.size() >= 2 && s[1] == L':') {
                return parseString(s.substr(2));
            }
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
        case Kind::Text:     return L"";  // type mismatch; caller routes via string overload
        }
        return L"";
    }

    std::wstring apply(const Spec& spec, const std::string& text) {
        if (!spec.valid) return L"";

        // Only Text specs make sense here. For numeric / date / duration
        // a string input is a type mismatch; we pass the input through
        // unchanged so the engine layer can surface it verbatim (the
        // caller is expected to have flagged the mismatch upstream).
        if (spec.kind == Kind::Text) {
            return formatString(spec, text);
        }
        return std::wstring(text.begin(), text.end());
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
        // Forward scan tracking quote state only. We remember the
        // position of the last '~' that occurs while we are NOT inside
        // a quoted string.
        //
        // ExprTk string literals use either single or double quotes.
        // Quoted strings can contain backslash escapes; we handle the
        // typical \' \" \\ cases so an escaped quote does not flip
        // the quote state.
        //
        // Picking the LAST '~' keeps a formula that legitimately contains
        // '~' (e.g. ExprTk's multi-sequence operator `~{a; b}`) unambiguous
        // - the trailing '~' is always the spec separator.
        long long lastTilde = -1;

        wchar_t quoteCh = 0;        // 0 = not in quote; otherwise the opening quote

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

            if (c == '~') {
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