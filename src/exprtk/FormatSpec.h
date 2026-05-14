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
//              types: f e g x b o
//              examples: 05.2f   +8.3e   .4g   .2-5g   08x   8b
//
//   duration:  t<unit>:<mode>
//              units: ts (sec) tm (min) th (hour) td (day)
//              modes: ms hm hms dh dhm dhms
//              examples: ts:hms   td:dhms   tm:hm
//
//   date:      D[<strftime fmt>] or D[!<strftime fmt>]
//              Value is a Unix timestamp (seconds since 1970-01-01 UTC).
//              Without '!': local time. With '!': UTC (Lua convention).
//              examples: D[%Y-%m-%d]   D[!%Y-%m-%dT%H:%M:%SZ]
//
// Locale-dependent modifiers (thousand sep, currency) are not part of v1.

#pragma once

#include <string>

namespace FormatSpec {

    // High-level category. Drives which subfields are valid.
    enum class Kind {
        Numeric,
        Duration,
        Date
    };

    enum class NumericType {
        Default,  // type omitted entirely - shortest round-trip (std::to_chars)
        Fixed,    // f
        Scientific, // e
        General,  // g
        Hex,      // x
        Binary,   // b
        Octal     // o
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

        // Date subfields. dateFormat holds the strftime pattern between
        // the brackets, without the leading '!' (which is captured by
        // dateUtc). UTF-8 storage so it can be passed straight to
        // strftime() in the renderer.
        bool dateUtc = false;
        std::string dateFormat;

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

    // Render value through spec. Returns the formatted string for
    // valid specs, or an empty string if spec.valid is false (which
    // the caller would have caught at parse time anyway).
    //
    // For Duration kind, value is interpreted in spec.durationUnit
    // (seconds / minutes / hours / days). Negative values produce
    // a leading '-' on the whole output.
    std::wstring apply(const Spec& spec, double value);

    // ---- Formula / spec splitter --------------------------------------
    //
    // Splits the body of an (?= ... ) block into the formula part and
    // the optional spec part at the LAST unquoted '~' that sits at
    // bracket depth 0. Quote-aware: single and double quotes inside
    // the formula (ExprTk string literals like 'a~b' or "a~b") are
    // skipped. Bracket-aware: square brackets used by future spec
    // forms (e.g. D[%Y-%m-%d]) keep their content together; the
    // splitter ignores '~' inside [...].
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