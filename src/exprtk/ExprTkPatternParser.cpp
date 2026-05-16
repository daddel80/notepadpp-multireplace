// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ExprTkPatternParser.cpp
// Single-pass scanner over the template string. UTF-8 is handled
// transparently because the marker characters ( ? = ) and \ all live
// in the ASCII range, so multi-byte sequences in literal text never
// need special handling - their bytes simply flow through.

#include "ExprTkPatternParser.h"

namespace MultiReplaceEngine {

    namespace {

        // Escape sequences recognised in literal context. Anything not listed
        // here passes through unchanged - so "\n" stays as "\n" (the caller's
        // extended-string layer can decide what \n means).
        constexpr const char ESCAPED_BACKSLASH[] = "\\\\";   // "\\"
        constexpr const char ESCAPED_MARKER[] = "\\(?=";  // "\(?="

        // Marker that opens an expression block. Not a generic regex feature
        // here - we recognise the exact 3-character sequence and nothing else,
        // so "( ?=" or "(?=" inside an already-open expression do not start a
        // new block.
        constexpr const char EXPRESSION_OPEN[] = "(?=";

        // Helper: case-sensitive prefix match at position pos.
        inline bool startsWith(const std::string& s, std::size_t pos, const char* needle, std::size_t needleLen)
        {
            if (pos + needleLen > s.size()) {
                return false;
            }
            for (std::size_t k = 0; k < needleLen; ++k) {
                if (s[pos + k] != needle[k]) {
                    return false;
                }
            }
            return true;
        }

    } // anonymous namespace


    bool ExprTkPatternParser::hasExpressions(const std::string& templ)
    {
        // Cheap scan that respects the same escape rules as parse(), so
        // \(?= is correctly NOT counted as an expression marker.
        const std::size_t n = templ.size();
        std::size_t i = 0;
        while (i < n) {
            // Skip an escaped backslash so the following two chars are
            // taken at face value.
            if (startsWith(templ, i, ESCAPED_BACKSLASH, 2)) {
                i += 2;
                continue;
            }
            // An escaped marker is a literal "(?=" - not an expression.
            if (startsWith(templ, i, ESCAPED_MARKER, 4)) {
                i += 4;
                continue;
            }
            if (startsWith(templ, i, EXPRESSION_OPEN, 3)) {
                return true;
            }
            ++i;
        }
        return false;
    }


    ExprTkPatternParser::ParseResult ExprTkPatternParser::parse(const std::string& templ)
    {
        ParseResult out;

        // Buffer for the literal text accumulated since the last segment
        // boundary. Flushed into a Literal segment whenever we hit an
        // expression marker or the end of the template.
        std::string literalBuf;
        std::size_t literalStart = 0;       // sourcePos for the buffered literal

        const std::size_t n = templ.size();
        std::size_t i = 0;

        while (i < n) {
            // ----- escaped backslash (\\ -> \) -----
            if (startsWith(templ, i, ESCAPED_BACKSLASH, 2)) {
                if (literalBuf.empty()) {
                    literalStart = i;
                }
                literalBuf.push_back('\\');
                i += 2;
                continue;
            }

            // ----- escaped marker (\(?= -> (?=) -----
            if (startsWith(templ, i, ESCAPED_MARKER, 4)) {
                if (literalBuf.empty()) {
                    literalStart = i;
                }
                literalBuf.append(EXPRESSION_OPEN, 3);   // append literal "(?="
                i += 4;
                continue;
            }

            // ----- expression open marker -----
            if (startsWith(templ, i, EXPRESSION_OPEN, 3)) {
                // Flush any accumulated literal first.
                if (!literalBuf.empty()) {
                    Segment seg;
                    seg.type = SegmentType::Literal;
                    seg.text = std::move(literalBuf);
                    seg.sourcePos = literalStart;
                    out.segments.push_back(std::move(seg));
                    literalBuf.clear();
                }

                const std::size_t markerPos = i;          // points at '('
                i += 3;                                   // skip "(?="
                const std::size_t exprStart = i;

                // Depth-counting scan to the matching ')'. Inside the
                // expression we deliberately do NOT honour escape
                // sequences - the formula engine sees the raw text.
                int depth = 1;
                while (i < n && depth > 0) {
                    const char c = templ[i];
                    if (c == '(') {
                        ++depth;
                    }
                    else if (c == ')') {
                        --depth;
                        if (depth == 0) {
                            break;
                        }
                    }
                    ++i;
                }

                if (depth != 0) {
                    out.success = false;
                    out.errorMessage = "Unmatched (?= - missing closing ')'";
                    out.errorPos = markerPos;
                    out.segments.clear();
                    return out;
                }

                const std::size_t exprEnd = i;            // points at ')'

                // Reject empty expression: (?=) is almost certainly a
                // user mistake, and ExprTk would fail to compile it
                // anyway. Catch it here with a clearer message.
                if (exprEnd == exprStart) {
                    out.success = false;
                    out.errorMessage = "Empty expression in (?= ... )";
                    out.errorPos = markerPos;
                    out.segments.clear();
                    return out;
                }

                Segment seg;
                seg.type = SegmentType::Expression;
                seg.text = templ.substr(exprStart, exprEnd - exprStart);
                seg.sourcePos = markerPos;
                out.segments.push_back(std::move(seg));

                ++i;                                      // skip the ')'
                continue;
            }

            // ----- ordinary literal byte -----
            if (literalBuf.empty()) {
                literalStart = i;
            }
            literalBuf.push_back(templ[i]);
            ++i;
        }

        // Flush trailing literal (if any).
        if (!literalBuf.empty()) {
            Segment seg;
            seg.type = SegmentType::Literal;
            seg.text = std::move(literalBuf);
            seg.sourcePos = literalStart;
            out.segments.push_back(std::move(seg));
        }

        return out;
    }

} // namespace MultiReplaceEngine