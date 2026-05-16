// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ExprTkPatternParser.h
// Splits a replace template into a sequence of literal and expression
// segments. Expressions are introduced by (?= and terminated by the
// matching ). Literal text passes through unchanged except for two
// escape sequences:
//
//     \\       ->  literal backslash
//     \(?=     ->  literal three-character sequence "(?="
//
// Inside an expression no escaping is performed; the contents are
// passed verbatim to the formula engine. The matching ) is found by
// depth-counting, so balanced parentheses inside an expression (e.g.
// "min(num(1), num(2))") are handled naturally.
//
// The parser is stateless and produces one structured ParseResult per
// call. It never throws; failures are reported via success=false plus
// errorMessage and errorPos.

#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace MultiReplaceEngine {

    class ExprTkPatternParser {
    public:
        enum class SegmentType {
            Literal,        // Passes through to the output verbatim
            Expression      // To be evaluated by the formula engine
        };

        struct Segment {
            SegmentType type = SegmentType::Literal;

            // For Literal segments: escape-decoded text (\\ -> \ etc.).
            // For Expression segments: raw expression text (between the
            // marker and the matching close paren), no escaping applied.
            std::string text;

            // Byte offset in the source template where this segment
            // begins. For Expression segments this points at the '(' of
            // the (?= marker, not at the first character of the
            // expression - this makes error messages line up with what
            // the user typed.
            std::size_t sourcePos = 0;
        };

        struct ParseResult {
            bool success = true;
            std::vector<Segment> segments;

            // Populated only when success == false.
            std::string errorMessage;
            std::size_t errorPos = 0;
        };

        // Parse `templ` into segments. Pure function, no shared state.
        static ParseResult parse(const std::string& templ);

        // Returns true if the template contains at least one (?=...)
        // marker that would produce an Expression segment. Useful as a
        // fast pre-check before invoking the heavier parse() pipeline -
        // a template with no markers can be passed through as-is by the
        // caller, no engine evaluation needed.
        static bool hasExpressions(const std::string& templ);

        // Non-instantiable utility class.
        ExprTkPatternParser() = delete;
    };

} // namespace MultiReplaceEngine