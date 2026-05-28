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
// Spec grammar (v2):
//
//   A spec is a universal frame, an optional/required type marker, then a
//   type-specific tail:
//
//     [ [fill] align ] [width]   [marker:]   <type tail>
//
//   fill/align/width form the FRAME and apply to EVERY kind, padding the
//   finished value. align: '<' left, '>' right, '^' center. fill is any
//   single char, recognised only when an align char follows it. width is
//   the minimum codepoint length. The frame's default align is per-kind:
//   numbers right, text/date/duration left. When a marker is present it
//   ALWAYS sits between the frame and the body (frame leads), with one
//   optional space allowed (">12 d:%Y" == ">12d:%Y", "<8 n:.2f"). A marker
//   before the frame ("n:<8.2f", "t:>8.3") is rejected.
//
//   numeric:   [ [fill]align ][width] [n:] [+] [0] [width] [.prec|.min-max] [type]
//              types: f e g x b o d ; n: is an OPTIONAL cosmetic marker
//              placed after the frame. width may be in the frame OR the
//              body, not both. examples: 5.2f  >8.2f  <8 n:.2f  +.4g  08x
//              '0' is the sign-aware zero-pad flag; mixing it with an
//              explicit align (e.g. >05d, or >8 n:05) is an error - pick one.
//
//   text:      [ [fill]align ][width] [t:] [.maxlen]
//              No marker needed: a bare spec on a string value is text.
//              't:' stays valid as an OPTIONAL assertion, placed after the
//              frame (">8 t:.3"). .N on text is the max codepoint length
//              (truncation); on a number it is decimal places - the same
//              .N, polymorphic by value type.
//              examples: >15   <20.10   *<20   ^10   >8 t:.3 (explicit)
//
//   duration:  [ [fill]align ] t<unit>:<mode>
//              units: ts (sec) tm (min) th (hour) td (day)
//              modes: ms hm hms dh dhm dhms
//              examples: ts:hms   >12 ts:hms   td:dhms   ^10tm:hm
//
//   date:      [ [fill]align ] d:<strftime fmt>  or  d:utc:<strftime fmt>
//              Value is a Unix timestamp (seconds since 1970-01-01 UTC).
//              Without 'utc:': local time. With 'utc:': UTC.
//              examples: d:%Y-%m-%d   >12 d:utc:%Y-%m-%dT%H:%M:%SZ
//
//              For every marker the frame is written BEFORE it (see the
//              top-level note above); one optional space may separate them
//              (">12 d:%Y" equals ">12d:%Y"). The '~' character is the spec
//              separator and so cannot appear inside a strftime pattern.
//
// Length semantics for text specs operate on UTF-8 codepoints, not
// bytes, so multi-byte characters count as one. Padding fill character
// is taken from the spec as a UTF-8 codepoint as well.
//
// Locale-dependent modifiers (thousand sep, currency) are not part of v2.

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
        Default, // no explicit align; Stage B picks per kind (num=right, text=left)
        Left,    // <
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
        // follows 'd:', without the leading 'utc:' keyword (which is
        // captured by dateUtc). Stored as a wide string and rendered via
        // wcsftime so locale-dependent names (%B/%A) and any literal
        // non-ASCII text round-trip as wchar_t without a codepage guess.
        bool dateUtc = false;
        std::wstring dateFormat;

        // Universal frame subfields (apply to every kind in the final pad
        // stage). width above is the minimum codepoint length. textAlign
        // defaults to Default so the renderer picks per kind (number=right,
        // text=left). textFill is a UTF-8 byte sequence for one codepoint -
        // typically one byte (space) but kept as a string to handle
        // non-ASCII fill characters cleanly. textMaxLength is text-only
        // (truncation); it does not share semantics with numeric precision.
        TextAlign textAlign = TextAlign::Default;
        std::string textFill = " ";     // single codepoint as UTF-8
        int textMaxLength = -1;         // -1 = not set; text only

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

    // True if spec is a PURE frame: parsed as Numeric (the default kind) but
    // carrying only fill/align/width/.precision - no numeric-only traits
    // (sign, zero-pad, a type letter, or a range precision). Such a spec is
    // type-neutral and may be applied to a string value (where .precision
    // means truncation) as well as to a number. Used by the engine to allow
    // a marker-free spec on a string-returning formula.
    bool isPureFrame(const Spec& spec);

    // Render a string through spec. Valid for Text kind and for a pure frame
    // (see isPureFrame); for a pure frame the precision is the max codepoint
    // length (truncation). Any other kind returns the input unchanged - a
    // genuine type mismatch the engine surfaces at compile time.
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