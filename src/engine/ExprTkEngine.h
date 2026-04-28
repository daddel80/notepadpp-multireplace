// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ExprTkEngine.h
// Concrete IFormulaEngine implementation backed by the ExprTk header-
// only mathematical expression library. The engine is intended for
// numeric work: arithmetic on capture groups, counter math, simple
// conditional logic with if/else.
//
// Surface visible to user scripts:
//
//   Variables (numeric, read-only):
//     CNT  LCNT  LINE  LPOS  APOS  COL
//     (mirrors the FormulaVars counter set)
//
//   Functions:
//     reg(N)  -> capture group N as double
//               * reg(0) = the full match
//               * reg(1) = first capture group
//               * reg(N) = N-th capture group
//               String-to-double uses std::from_chars (locale-
//               independent: only "." is a decimal separator). A
//               capture that does not parse as a number yields 0.0.
//
// Strings (MATCH / FPATH / FNAME) are intentionally NOT exposed - the
// engine is numeric-only. Users who need string manipulation should
// switch the tab to the Lua engine.
//
// Lifecycle:
//   1. Construct via EngineFactory::create(EngineType::ExprTk, host).
//      The host pointer is accepted for symmetry with LuaEngine but
//      currently unused; ExprTk needs no UI callbacks.
//   2. initialize() registers variables and functions in the symbol
//      table.
//   3. compile(template) splits the template into segments and pre-
//      compiles every (?=expr) block.
//   4. execute(template, vars, ...) updates the per-match variables
//      and evaluates each compiled expression in document order.
//   5. shutdown() releases the compiled expressions and symbol table.
//
// Threading:
//   Not thread-safe. Each tab owns its own engine instance.

#pragma once

#include "IFormulaEngine.h"
#include "ILuaEngineHost.h"
#include "ExprTkPatternParser.h"

// ExprTk is a single-header library that pulls in a fair amount of
// template machinery, so the include lives here in the engine TU
// only. Other code paths see ExprTkEngine through IFormulaEngine.
#include "../exprtk/exprtk.hpp"

#include <string>
#include <vector>

namespace MultiReplaceEngine {

    class ExprTkEngine final : public IFormulaEngine {
    public:
        // The host pointer is accepted for symmetry with LuaEngine; the
        // ExprTk path currently makes no callbacks back into the host. It
        // is stored for future use (e.g. an "evaluation log" hook) but
        // never dereferenced by the current implementation.
        explicit ExprTkEngine(ILuaEngineHost* host);
        ~ExprTkEngine() override;

        // ----- IFormulaEngine -----------------------------------------------

        bool initialize() override;
        void shutdown()   override;

        bool compile(const std::string& scriptUtf8) override;

        FormulaResult execute(
            const std::string& scriptUtf8,
            const FormulaVars& vars,
            bool isRegexMatch,
            int  documentCodepage
        ) override;

        EngineType   type()        const override { return EngineType::ExprTk; }
        std::wstring shortName()   const override { return L"ExprTk"; }
        std::wstring shortLetter() const override { return L"E"; }
        std::wstring helpUrl()     const override;

    private:
        // Type aliases for ExprTk's templated machinery. Restricting the
        // engine to double precision keeps the symbol table simple and is
        // sufficient for the use cases we target.
        using symbol_table_t = exprtk::symbol_table<double>;
        using expression_t = exprtk::expression<double>;
        using parser_t = exprtk::parser<double>;

        // ----- internal helpers --------------------------------------------

        // Format and surface an error to the host. Mirrors LuaEngine's
        // pattern: gated by ILuaEngineHost::isLuaErrorDialogEnabled() (the
        // setting is engine-agnostic in spirit; only the symbol name still
        // carries the legacy "Lua" prefix). The body is prefixed with
        // "ExprTk: " so the user can immediately tell which engine raised
        // the error in mixed-engine workspaces.
        void reportError(ILuaEngineHost::ErrorCategory category,
            const std::string& details);

        // Parse a UTF-8 capture string into a double. Returns 0.0 for
        // empty / non-numeric input. Locale-independent: only '.' is
        // recognised as decimal separator.
        static double parseCaptureToDouble(const std::string& s);

        // Format a double back into a string for insertion into the
        // replace output. Uses the shortest round-trip representation
        // available via std::to_chars, falling back to a fixed-precision
        // format for extreme magnitudes.
        static std::string formatDouble(double value);

        // ExprTk-callable wrapper: implements reg(N). Reads from the
        // _captures vector populated at the start of execute().
        // Out-of-range indices return 0.0 (consistent with the empty-
        // capture rule in parseCaptureToDouble).
        class RegFunction : public exprtk::ifunction<double> {
        public:
            explicit RegFunction(ExprTkEngine* owner)
                : exprtk::ifunction<double>(1)   // arity = 1
                , _owner(owner) {
            }

            double operator()(const double& index) override;

        private:
            ExprTkEngine* _owner;
        };

        // ----- state -------------------------------------------------------

        ILuaEngineHost* _host;            // accepted, currently unused

        // Pre-parsed segments of the most recently compiled template.
        // Kept alongside the expressions so execute() can iterate both in
        // lockstep without re-parsing.
        ExprTkPatternParser::ParseResult _parsedTemplate;

        // Compile cache: when compile() is called twice with the same
        // template, we skip re-parsing and re-compiling.
        std::string _lastCompiledScript;
        bool        _haveCompiled = false;

        // ExprTk plumbing
        symbol_table_t              _symbolTable;
        parser_t                    _parser;
        std::vector<expression_t>   _compiledExpressions;

        // Variables registered with the symbol table. Held as members
        // (not locals) because ExprTk binds them by reference - they must
        // outlive every expression that uses them.
        double _varCNT = 0.0;
        double _varLCNT = 0.0;
        double _varLINE = 0.0;
        double _varLPOS = 0.0;
        double _varAPOS = 0.0;
        double _varCOL = 0.0;

        // Captures are pre-parsed once per execute() into doubles. The
        // RegFunction reads from this vector. Index 0 holds the full
        // match (FormulaVars::MATCH); index 1..N the capture groups.
        std::vector<double> _captures;

        // The reg() callable, registered with the symbol table.
        RegFunction _regFunction;
    };

} // namespace MultiReplaceEngine