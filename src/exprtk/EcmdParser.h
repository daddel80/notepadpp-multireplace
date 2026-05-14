// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// EcmdParser.h
// Parser for .ecmd files - libraries of user-defined ExprTk functions.
// Each function definition is a textual block of the form
//
//     function NAME(PARAMS) [: RETURN] BODY end
//
// where PARAMS is a comma-separated list of "name" or "name: TYPE", and
// TYPE / RETURN are single uppercase letters T (scalar, the default) or
// S (string). Examples:
//
//     function double_it(n)              # T, T   - reine Mathe
//         return n * 2;
//     end
//
//     function num2rom(n) : S            # T, S   - skalar in, string out
//         var r := ''; var v := n;
//         while (v >= 1000) { r := r + 'M'; v := v - 1000; };
//         /* ... */
//         return r;
//     end
//
//     function rom2num(s: S)             # S, T
//         var total := 0;
//         /* ... */
//         return total;
//     end
//
//     function padleft(s: S, n, fill: S) : S   # S, T, S, S
//         /* ... */
//         return r;
//     end
//
// Body normalisation:
//   The user writes 'return EXPR;' the way every other language allows.
//   The parser rewrites this to ExprTk's native 'return [EXPR];' form
//   before handing the body off downstream. Users who already write
//   the bracketed form are left untouched.
//
// The parser performs strict syntactic validation and is independent of
// ExprTk - it produces structured FunctionDefs that a downstream wrapper
// hands to exprtk::parser. Errors abort the load with a message and a
// byte offset into the source.

#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace MultiReplaceEngine {

    class EcmdParser {
    public:
        // T = scalar (double), S = string. Stored as the same ASCII chars
        // ExprTk itself uses in its parameter-sequence strings, so the
        // downstream wrapper can pass them through unchanged.
        enum class ValueType : char {
            Scalar = 'T',
            String = 'S'
        };

        struct ParamDef {
            std::string name;
            ValueType   type = ValueType::Scalar;
            std::size_t sourcePos = 0;
        };

        struct FunctionDef {
            std::string             name;
            std::vector<ParamDef>   params;
            ValueType               returnType = ValueType::Scalar;
            std::string             body;        // normalised text (return-wrap applied), ready for exprtk::parser
            std::size_t             sourcePos = 0;  // offset of the 'function' keyword
        };

        struct ParseResult {
            bool                        success = true;
            std::vector<FunctionDef>    functions;

            // Populated only when success == false. errorPos is a byte
            // offset into the original input - the caller can translate
            // it into a line/column for the user dialog.
            std::string                 errorMessage;
            std::size_t                 errorPos = 0;
        };

        // Parse the full contents of a .ecmd file. Stateless and pure;
        // safe to call from any thread.
        static ParseResult parse(const std::string& fileContent);

        EcmdParser() = delete;
    };

} // namespace MultiReplaceEngine
