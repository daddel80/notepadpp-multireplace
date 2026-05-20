// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// FormatSpec.h
// Output formatting for the (?= formula ~ spec ) syntax. Parses a
// spec string into a structured FormatSpec, then renders a numeric
// value through it. Locale-independent: all output uses the C
// locale (dot decimal) regardless of system settings, so the CSV
// list format stays stable across machines.
//
// Spec grammar (v1):
//
//   numeric:   [+] [0] [width] [.prec | .min-max] [type]
//              types: f e g x b o d
//              examples: 05.2f   +8.3e   .4g   .2-5g   08x   8b   05d
//              Optional explicit form: prepend 'n:' for visual symmetry
//              with the other markers; 'n:5.2f' is identical to '5.2f'.
//
//   text:      t:[fill][align][width][.maxlen]
//              align: '<' (left, default), '>' (right), '^' (center)
//              fill:  any single character, must be followed by an
//                     align letter to disambiguate from width
//              width: minimum codepoint length; pads with fill char
//              maxlen: maximum codepoint length; truncates if exceeded
//              examples: t:<15   t:>20.10   t:*<20   t:.<15   t:^10
//
//   duration:  t<unit>:<mode>
//              units: ts (sec) tm (min) th (hour) td (day)
//              modes: ms hm hms dh dhm dhms
//              examples: ts:hms   td:dhms   tm:hm
//
//              The 't' marker is shared with text: parser disambiguates
//              by the second character ('s'/'m'/'h'/'d' -> duration,
//              ':' -> text).
//
//   date:      d:<strftime fmt> or d:!<strftime fmt>
//              Value is a Unix timestamp (seconds since 1970-01-01 UTC).
//              Without '!': local time. With '!': UTC (Lua convention).
//              examples: d:%Y-%m-%d   d:!%Y-%m-%dT%H:%M:%SZ
//
//              The '~' character is the spec separator and therefore
//              cannot appear inside the strftime pattern. Use any
//              other literal separator (- / : space).
//
// Length semantics for text specs operate on UTF-8 codepoints, not
// bytes, so multi-byte characters count as one. Padding fill character
// is taken from the spec as a UTF-8 codepoint as well.
//
// Locale-dependent modifiers (thousand sep, currency) are not part of v1.

#pragma once

#include <string>

namespace FormatSpec {

    // High-level category. Drives which subfields are valid.
    enum class Kind {
        Numeric,
        Text,
        Duration,
        Date
    };

    enum class NumericType {
        Default,    // type omitted entirely - shortest round-trip (std::to_chars)
        Fixed,      // f
        Scientific, // e
        General,    // g
        Hex,        // x
        Binary,     // b
        Octal,      // o
        Integer     // d  - signed base-10, truncated toward zero
    };

    enum class TextAlign {
        Left,    // <  (default)
        Right,   // >
        Center   // ^
    };

    enum class DurationUnit {
        Seconds,  // ts
        Minutes,  // tm
        Hours,    // th
        Days      // td
    };

    enum class DurationMode {
        Ms,    // M:SS
        Hm,    // H:MM
        Hms,   // H:MM:SS
        Dh,    // D HH
        Dhm,   // D HH:MM
        Dhms   // D HH:MM:SS
    };

    // Parsed representation of a format spec. Filled by parse(); the
    // unused subfield is left default. Width/precision values of -1
    // mean "not specified".
    struct Spec {
        Kind kind = Kind::Numeric;

        // Numeric subfields
        NumericType numericType = NumericType::Default;
        bool forceSign = false;     // +
        bool zeroPad = false;       // 0
        int  width = -1;            // -1 = not set
        int  precisionMin = -1;     // -1 = not set; for .N also stored here
        int  precisionMax = -1;     // -1 = not set; only set for .min-max form

        // Duration subfields
        DurationUnit durationUnit = DurationUnit::Seconds;
        DurationMode durationMode = DurationMode::Hms;

        // Date subfields. dateFormat holds the strftime pattern that
        // follows 'd:', without the leading '!' (which is captured by
        // dateUtc). UTF-8 storage so it can be passed straight to
        // strftime() in the renderer.
        bool dateUtc = false;
        std::string dateFormat;

        // Text subfields. width above is reused as the minimum codepoint
        // length. textMaxLength holds the optional max (truncation) so it
        // doesn't share semantics with the numeric precisionMin/Max
        // (which mean decimal places). textFill is a UTF-8 byte sequence
        // for one codepoint - typically one byte (space) but we keep it
        // as a string to handle non-ASCII fill characters cleanly.
        TextAlign textAlign = TextAlign::Left;
        std::string textFill = " ";     // single codepoint as UTF-8
        int textMaxLength = -1;         // -1 = not set

        // True if parse() succeeded. Callers should check before apply().
        bool valid = false;

        // Human-readable error from parse() when valid==false. The
        // ExprTk engine surfaces this through the same dialog path as
        // ExprTk compile errors.
        std::string errorMessage;
    };

    // Parse a spec string into a Spec. Leading/trailing whitespace
    // is tolerated. An empty spec is rejected (caller should not
    // call parse() in that case; that's the no-spec path with
    // default ExprTk output). Strict on unknown type letters and
    // malformed grammar.
    Spec parse(const std::wstring& specText);

    // Render a numeric value through spec. Valid for kinds Numeric,
    // Duration, Date. For Text-kind specs this returns an empty string;
    // callers must route those through the string overload instead.
    //
    // For Duration kind, value is interpreted in spec.durationUnit
    // (seconds / minutes / hours / days). Negative values produce
    // a leading '-' on the whole output.
    std::wstring apply(const Spec& spec, double value);

    // Render a string through spec. Valid only for Text kind. Any other
    // kind returns the input unchanged - the engine is expected to
    // surface a type-mismatch diagnostic at parse / compile time, not
    // silently format a numeric spec against a string.
    //
    // The input is treated as UTF-8. Width and max-length are counted in
    // codepoints; truncation never splits a multi-byte sequence.
    std::wstring apply(const Spec& spec, const std::string& text);

    // ---- Formula / spec splitter --------------------------------------
    //
    // Splits the body of an (?= ... ) block into the formula part and
    // the optional spec part at the LAST unquoted '~' that sits outside
    // ExprTk string literals. Quote-aware: single and double quotes
    // inside the formula (ExprTk string literals like 'a~b' or "a~b")
    // are skipped. Picking the LAST '~' keeps a formula that legitimately
    // contains '~' (e.g. ExprTk's multi-sequence operator `~{a; b}`)
    // unambiguous - any earlier '~' belongs to the formula, the last one
    // is the spec separator.
    //
    // Returned pair:
    //   .first  = formula text (left of the separator), trimmed
    //   .second = spec text (right of the separator), trimmed; empty
    //             when no separator was found
    //
    // No '~' present -> the whole input goes into .first, .second
    // stays empty. That's the back-compat path: ExprTk evaluates the
    // formula and formats with the default shortest round-trip.
    struct Split {
        std::string formula;
        std::string spec;
        bool hasSpec = false;
    };
    Split splitFormulaSpec(const std::string& blockBody);

}  // namespace FormatSpec