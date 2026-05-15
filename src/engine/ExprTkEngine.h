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
//     CNT  LCNT  LINE  LPOS  APOS  COL  HIT
//     (CNT/LCNT/LINE/LPOS/APOS/COL mirror the FormulaVars counter set;
//      HIT is the full match parsed as a number, equivalent to num(0).)
//
//   Variables (string, read-only - usable inside return [...] only):
//     FPATH  FNAME
//
//   Functions:
//     num(N)  -> capture group N as double
//                * num(0) = the full match (same as HIT)
//                * num(1) = first capture group
//                * num(N) = N-th capture group
//                String-to-double uses std::from_chars (locale-
//                independent: only "." is a decimal separator).
//                A capture that does not parse as a number yields NaN;
//                NaN propagates through arithmetic and is surfaced as
//                an explicit error rather than silently coerced to 0.
//     txt(N)  -> capture group N as raw string (for return [...])
//     skip()  -> mark the current match to be left untouched
//     seq([start, [inc]])
//             -> sequence value start + (CNT-1)*inc.
//                Both arguments default to 1 so seq() yields 1,2,3,...
//                Useful for numbering matches without writing the
//                start+(CNT-1)*inc formula by hand.
//     numcol(N) / numcol('name')
//             -> CSV column on the current row, parsed as double.
//                Index is 1-based; "name" looks up the document's
//                first line as a header. NaN on missing column or
//                non-numeric content. Requires CSV mode to be active.
//     txtcol(N) / txtcol('name')
//             -> CSV column on the current row, as raw string. Same
//                addressing rules as numcol; only useful inside a
//                return [...] list.
//     ecmd(p) -> load a library of user-defined functions from path p
//                (typically used in an empty-Find init slot)
//
// The match text is intentionally not exposed as a string variable.
// Use txt(0) for the raw match string, HIT or num(0) for the numeric
// form, or a Boost backref \0 in the literal part of the template.
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
#include "../exprtk/EcmdParser.h"
#include "../exprtk/FormatSpec.h"

// ExprTk is a single-header library that pulls in a fair amount of
// template machinery, so the include lives here in the engine TU
// only. Other code paths see ExprTkEngine through IFormulaEngine.
#include "../exprtk/exprtk.hpp"

#include <memory>
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

        // ecmd-loaded user libraries live for one Replace-All run only.
        // beginRun() discards any library state from the previous run so
        // a re-run reloads the .ecmd files fresh from disk and removing
        // the ecmd() init slot makes its functions disappear.
        // _errorSkipCount / _skipAllErrors are still reset by the base.
        void beginRun() override;

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
        // pattern: gated by ILuaEngineHost::isFormulaErrorDialogEnabled().
        // The body is prefixed with "ExprTk: " so the user can immediately
        // tell which engine raised the error in mixed-engine workspaces.
        void reportError(ILuaEngineHost::ErrorCategory category,
            const std::string& details);

        // Invalid-result handling for both the numeric and return-list
        // paths. Triggered on NaN or Inf: bumps the skip counter and may
        // show the recoverable dialog. The original match stays intact
        // when the user skips, so no data is lost.
        void handleInvalid(const std::string& exprText);

        // Parse a UTF-8 capture string into a double. Returns NaN for
        // empty / non-numeric input. Accepts both '.' and ',' as decimal
        // separator.
        static double parseCaptureToDouble(const std::string& s);

        // Format a double back into a string for insertion into the
        // replace output. Uses the shortest round-trip representation
        // available via std::to_chars, falling back to a fixed-precision
        // format for extreme magnitudes.
        static std::string formatDouble(double value);

        // Append the elements of an ExprTk return-value list onto the
        // output string. Used when the user invokes 'return [...]' inside
        // an (?= ...) expression - that lets a single expression emit
        // mixed string and numeric output, which is the only way to get
        // a string-typed variable (FNAME, FPATH) into the result.
        // Errors during element traversal are tolerated rather than
        // aborted; an unknown element type is silently skipped so a
        // future ExprTk extension cannot break a running replace.
        // A NaN or Inf scalar element sets _outputHadInvalid so the
        // caller can skip the match (consistent with the numeric-path
        // invalid-result handling).
        // When escapeForRegex is true, any \ or $ produced by the
        // expression is escaped so the output is safe to feed to a
        // regex replacement engine; literal text segments outside of
        // (?=...) blocks are not affected.
        void appendExprtkResults(const expression_t& expr, std::string& out,
            bool escapeForRegex);

        // ExprTk-callable wrapper: implements num(N). Reads from the
        // _captures vector populated at the start of execute().
        // Out-of-range indices return 0.0 (consistent with the empty-
        // capture rule in parseCaptureToDouble).
        class NumFunction : public exprtk::ifunction<double> {
        public:
            explicit NumFunction(ExprTkEngine* owner)
                : exprtk::ifunction<double>(1)   // arity = 1
                , _owner(owner) {
            }

            double operator()(const double& index) override;

        private:
            ExprTkEngine* _owner;
        };

        // ExprTk-callable wrapper: implements skip(). Arity zero, returns
        // 0.0 so it can sit inside a numeric expression (e.g. an if/else
        // branch), and signals via the engine's _wantSkip flag that the
        // current match should not be replaced. Mirrors LuaEngine's skip()
        // semantics so users get the same conditional-replace pattern in
        // both engines.
        class SkipFunction : public exprtk::ifunction<double> {
        public:
            explicit SkipFunction(ExprTkEngine* owner)
                : exprtk::ifunction<double>(0)   // arity = 0
                , _owner(owner) {
            }

            double operator()() override;

        private:
            ExprTkEngine* _owner;
        };

        // ExprTk-callable wrapper: implements txt(N), the string-side
        // counterpart to num(N). Returns the captured text at index N
        // as a string (txt(0) = full match, txt(1..N) = capture groups).
        // Built on igeneric_function rather than plain ifunction because
        // ifunction can only return doubles - and the whole point of
        // txt() is to surface the *string* form of a capture.
        //
        // Only meaningful inside an ExprTk return statement, e.g.
        //   (?=return ['<', txt(1), '>'])
        // outside a return ExprTk will refuse to compile because the
        // surrounding numeric expression cannot consume a string.
        //
        // The "T" arity-spec declares one Scalar argument (the index).
        // The e_rtrn_string flag tells ExprTk this function returns a
        // string, so it can be used in string contexts within an
        // expression (Section 15 of the ExprTk docs).
        class TxtFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using scalar_t = typename generic_t::scalar_view;

            explicit TxtFunction(ExprTkEngine* owner)
                : igenfunct_t("T", igenfunct_t::e_rtrn_string)
                , _owner(owner) {
            }

            // Override picks up the string-return overload because we
            // declared e_rtrn_string in the constructor. The first
            // parameter receives the result; the return double is
            // ignored by ExprTk in this mode but must be present to
            // match the base-class signature.
            double operator()(std::string& result,
                parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // ExprTk-callable wrapper: implements seq(start, increment), a
        // sequence generator that returns start, start+inc, start+2*inc,
        // ... over consecutive matches.
        //
        // Both arguments are optional and default to 1, so:
        //   seq()         -> 1, 2, 3, ...
        //   seq(10)       -> 10, 11, 12, ...
        //   seq(0, 10)    -> 0, 10, 20, ...
        //   seq(100, -10) -> 100, 90, 80, ...
        //
        // Built on ivararg_function so the variadic argument count can be
        // checked at call time rather than fixed at registration. The
        // result is start + (CNT - 1) * inc; CNT is read directly from
        // the engine's _varCNT member which is updated at the start of
        // each execute().
        class SeqFunction : public exprtk::ivararg_function<double> {
        public:
            explicit SeqFunction(ExprTkEngine* owner)
                : _owner(owner) {
                exprtk::enable_zero_parameters(*this);  // allow seq() with no args
            }

            double operator()(const std::vector<double>& args) override;

        private:
            ExprTkEngine* _owner;
        };

        // numcol(N) / numcol('name') - column value as double.
        // arity-spec "T|S": one argument, scalar (index) or string (header).
        // Dispatches to the host's CSV column accessors; returns NaN on
        // any failure (no CSV mode, missing column, non-numeric content).
        //
        // Multi-prototype functions ("T|S") route through the psi
        // variant of operator(); psi tells us which alternative the
        // parser matched (0=T, 1=S) but resolveCsvCell branches on the
        // actual parameter type anyway, so psi is unused here.
        class NumColFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;

            explicit NumColFunction(ExprTkEngine* owner)
                : igenfunct_t("T|S")
                , _owner(owner) {
            }

            double operator()(const std::size_t& /*psi*/,
                parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // txtcol(N) / txtcol('name') - column value as string. Same
        // dispatch as NumColFunction but returns the raw cell text;
        // only meaningful inside a return [...] list, mirroring txt().
        class TxtColFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;

            explicit TxtColFunction(ExprTkEngine* owner)
                : igenfunct_t("T|S", igenfunct_t::e_rtrn_string)
                , _owner(owner) {
            }

            double operator()(const std::size_t& /*psi*/,
                std::string& result,
                parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // ExprTk-callable: parsedate(str, fmt) -> Unix timestamp.
        //
        // Parses str against the strftime-style fmt and returns the
        // resulting time as seconds-since-epoch. With a leading '!'
        // in fmt the result is treated as UTC; otherwise as local
        // time (Lua convention, matches our D[...] output spec).
        //
        // Returns NaN on parse failure or out-of-range fields, so a
        // bad input flows through the same recoverable-error dialog
        // as any other invalid-number result.
        //
        // Two string args, scalar return: "SS" arity-spec, plain
        // igeneric_function (no e_rtrn_string).
        class ParsedateFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using string_t = typename generic_t::string_view;

            explicit ParsedateFunction(ExprTkEngine* owner)
                : igenfunct_t("SS")
                , _owner(owner) {
            }

            double operator()(parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // Base-conversion built-ins: num <-> hex/bin/oct.
        //
        // num2X(n)   takes a scalar, returns a bare lowercase string
        //            ("ff", "1010", "77") - no "0x" / "0b" / "0o" prefix.
        //            Negative inputs come out as "-f". Float inputs are
        //            truncated toward zero before conversion.
        // X2num(s)   takes a string, returns a scalar. Accepts the input
        //            case-insensitively and with or without the matching
        //            prefix; surrounding whitespace is trimmed.
        //            Invalid characters for the base yield NaN.
        class Num2BaseFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using scalar_t = typename generic_t::scalar_view;

            Num2BaseFunction(ExprTkEngine* owner, int base)
                : igenfunct_t("T", igenfunct_t::e_rtrn_string)
                , _owner(owner)
                , _base(base) {
            }

            double operator()(std::string& result,
                parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
            int _base;
        };

        class Base2NumFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using string_t = typename generic_t::string_view;

            Base2NumFunction(ExprTkEngine* owner, int base)
                : igenfunct_t("S")
                , _owner(owner)
                , _base(base) {
            }

            double operator()(parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
            int _base;
        };

        // Roman numerals: num2rom(n) / rom2num(s).
        //
        // num2rom: emits uppercase canonical form using subtractive
        //   pairs (IV, IX, XL, XC, CD, CM). Range 1..3999; outside
        //   that range, or for NaN/Inf inputs, returns empty string.
        //   Float inputs truncate toward zero.
        // rom2num: lenient parser. Accepts mixed case and tolerates
        //   non-canonical forms like "IIII" for 4. Invalid characters
        //   or surrounding garbage yield NaN.
        class Num2RomFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using scalar_t = typename generic_t::scalar_view;

            explicit Num2RomFunction(ExprTkEngine* owner)
                : igenfunct_t("T", igenfunct_t::e_rtrn_string)
                , _owner(owner) {
            }

            double operator()(std::string& result,
                parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        class Rom2NumFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using string_t = typename generic_t::string_view;

            explicit Rom2NumFunction(ExprTkEngine* owner)
                : igenfunct_t("S")
                , _owner(owner) {
            }

            double operator()(parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // ----- ecmd library plumbing ---------------------------------------
        //
        // A user-defined function loaded from a .ecmd file. One instance
        // per function. Holds a pre-compiled ExprTk expression (the user-
        // authored body) plus the argument slots bound to it by reference.
        //
        // Operation:
        //   - At Replace-All start, EcmdLibrary::load() builds one instance
        //     per function in the file.
        //   - ExprTk routes a call to this instance through operator(),
        //     where we marshal incoming params into the bound slots, run
        //     the inner expression, and hand back the result.
        //   - Re-entrancy: each operator() saves the current slot values
        //     before overwriting them and restores them on return, so a
        //     function can call itself (recursion) or call another ecmd
        //     function that calls it back (mutual recursion).
        //
        // The two operator() overloads cover scalar and string return
        // respectively; which one ExprTk picks depends on the rtrn_type
        // we declared in the igeneric_function constructor.
        class EcmdFunctionInstance : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;

            EcmdFunctionInstance(const EcmdParser::FunctionDef& def);

            // Register both the per-instance inner symbol table (for the
            // bound argument slots) and the shared library symbol table
            // (so cross-calls to other ecmd functions resolve). Must be
            // called before compileBody().
            void prepareSymbolTables(symbol_table_t& libTable);

            // Compile the user-authored body string. Must run after
            // prepareSymbolTables(), and after every function in the
            // library has been registered in libTable so cross-call name
            // lookups succeed at compile time.
            bool compileBody(parser_t& parser, std::string& errorOut);

            // Scalar-return path (declared when returnType == 'T').
            double operator()(parameter_list_t parameters) override;

            // String-return path (declared when returnType == 'S').
            double operator()(std::string& result,
                parameter_list_t parameters) override;

            const std::string& name() const { return _name; }

        private:
            // Build the ExprTk parameter-sequence string ("T", "S", "TS",
            // ...) from the parameter type list.
            static std::string makeParamSequence(
                const std::vector<EcmdParser::ParamDef>& params);

            static igenfunct_t::return_type makeReturnType(
                EcmdParser::ValueType rt);

            // Snapshot of all argument slots before a call, kept so the
            // restore step can put them back exactly as they were. Holds
            // values, not references, so it survives even when the slot
            // is overwritten by a recursive call.
            struct ArgSnapshot {
                std::vector<double>      scalars;
                std::vector<std::string> strings;
            };

            // Save current slot values, then write the incoming params
            // into them. Shared between the scalar and string overloads.
            void marshalArgs(parameter_list_t parameters, ArgSnapshot& out);

            // Restore slot values from a prior snapshot, consuming it in
            // the same order the marshal pushed entries.
            void restoreArgs(std::size_t paramCount, ArgSnapshot& snapshot);

            std::string                                 _name;
            std::string                                 _body;

            symbol_table_t                              _innerSymbols;
            expression_t                                _innerExpr;

            // Argument slots, bound into _innerSymbols by reference.
            // unique_ptrs because vector<T>::push_back may relocate the
            // backing buffer and invalidate ExprTk's reference bindings.
            std::vector<std::unique_ptr<double>>        _scalarSlots;
            std::vector<std::unique_ptr<std::string>>   _stringSlots;

            // Routing: for each parameter, kind (scalar/string) and index
            // into the matching slot vector. Used by operator() to copy
            // the i-th incoming param into the right slot.
            struct ArgRoute {
                bool        isString;
                std::size_t slotIndex;
            };
            std::vector<ArgRoute>                       _argRoutes;
        };

        // Owns all ecmd functions loaded during the current run. Lifetime
        // is one Replace-All: re-created in beginRun(). The library's own
        // symbol_table is what holds the function registrations; the
        // outer engine expression registers this table alongside its main
        // symbol_table so user-written (?=...) blocks can call any loaded
        // function.
        class EcmdLibrary {
        public:
            EcmdLibrary() = default;

            // Returns the library's symbol_table. The engine registers
            // this against every compiled (?=...) expression so calls to
            // ecmd-loaded functions resolve at compile time.
            symbol_table_t& symbolTable() { return _libTable; }

            // Load every function from the parsed file. Two-pass:
            //  1) construct and register all instances at the library's
            //     symbol_table (empty bodies still uncompiled).
            //  2) compile each body. Since all names are now visible in
            //     _libTable, cross-calls and recursion resolve.
            // Returns false on any parser, registration, or compile
            // failure, with errorOut filled by an actionable message.
            bool load(const std::string& fileContent,
                const std::string& sourceLabel,
                std::string& errorOut);

            bool empty() const { return _instances.empty(); }

        private:
            symbol_table_t                                       _libTable;
            std::vector<std::unique_ptr<EcmdFunctionInstance>>   _instances;
            // Parser is owned by the library so its diagnostic state
            // doesn't leak across loads from different files.
            parser_t                                             _parser;
        };

        // ExprTk-callable: ecmd("path") -> 0. The user invokes this from
        // an empty-Find init slot to load a library of functions for the
        // current Replace-All run.
        //
        // One string arg, scalar return: "S" arity-spec, plain
        // igeneric_function (no e_rtrn_string). The return value (always
        // 0.0) is the same convention skip() uses - the meaningful effect
        // is the side-effect of loading the file.
        class EcmdLoaderFunction : public exprtk::igeneric_function<double> {
        public:
            using igenfunct_t = exprtk::igeneric_function<double>;
            using generic_t = typename igenfunct_t::generic_type;
            using parameter_list_t = typename igenfunct_t::parameter_list_t;
            using string_t = typename generic_t::string_view;

            explicit EcmdLoaderFunction(ExprTkEngine* owner)
                : igenfunct_t("S")
                , _owner(owner) {
            }

            double operator()(parameter_list_t parameters) override;

        private:
            ExprTkEngine* _owner;
        };

        // Read `utf8Path` from disk and feed its contents into the
        // current ecmd library. Errors flow through reportError() with
        // the path mentioned in the diagnostic.
        void loadEcmdFile(const std::string& utf8Path);

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

        // Per-segment output spec, in lockstep with _compiledExpressions.
        // hasSpec=false means no "~ spec" suffix; engine uses formatDouble().
        // isString=true means the root node is string-producing (e.g.
        // (?=num2rom(num(1))) or (?='abc')); the engine reads the string
        // directly via expression_helper::get_string() instead of value().
        struct SegmentSpec {
            bool hasSpec = false;
            bool isString = false;
            FormatSpec::Spec spec;
        };
        std::vector<SegmentSpec>    _segmentSpecs;

        // Variables registered with the symbol table. Held as members
        // (not locals) because ExprTk binds them by reference - they must
        // outlive every expression that uses them.
        double _varCNT = 0.0;
        double _varLCNT = 0.0;
        double _varLINE = 0.0;
        double _varLPOS = 0.0;
        double _varAPOS = 0.0;
        double _varCOL = 0.0;
        double _varHIT = 0.0;

        // Captures are pre-parsed once per execute() into doubles. The
        // NumFunction reads from this vector. Index 0 holds the full
        // match (FormulaVars::MATCH); index 1..N the capture groups.
        std::vector<double> _captures;

        // String form of the capture groups, kept alongside the numeric
        // _captures so TxtFunction can hand them to ExprTk return-list
        // entries. Index 0 is *not* used here (txt(0) reads _strMATCH
        // directly); slots 0..N-1 hold capture[1..N].
        std::vector<std::string> _captureStrings;

        // Holds the current match's string-side metadata for ExprTk's
        // string-typed symbol table entries. ExprTk binds string vars by
        // reference (Section 13 of the ExprTk docs), so these have to
        // live as long as the registered expressions.
        std::string _strMATCH;
        std::string _strFPATH;
        std::string _strFNAME;

        // Per-match flags, reset at the start of each execute().
        // _wantStop, _skipAllErrors and _errorSkipCount live in
        // IFormulaEngine and are reset by the base beginRun() / per-match
        // by execute() (for _wantStop).
        bool _wantSkip = false;          // set by skip() in the formula
        bool _outputHadInvalid = false;  // set when NaN or Inf reaches the output

        // The num() callable, registered with the symbol table.
        NumFunction _numFunction;

        // The skip() callable, also registered with the symbol table.
        SkipFunction _skipFunction;

        // The txt() callable for string-typed capture access (used
        // inside ExprTk return lists).
        TxtFunction _txtFunction;

        // The seq() callable for sequence generation across matches.
        SeqFunction _seqFunction;

        // CSV column access via the host (numcol/txtcol).
        NumColFunction _numColFunction;
        TxtColFunction _txtColFunction;

        // The parsedate(str, fmt) callable for string-to-timestamp
        // parsing - the inverse of D[fmt] output.
        ParsedateFunction _parsedateFunction;

        // Base-conversion built-ins. Two parameterised templates serve
        // hex/bin/oct in both directions; each instance carries its base
        // (16/2/8) so the same operator() logic handles all six names.
        Num2BaseFunction _num2hexFunction;
        Num2BaseFunction _num2binFunction;
        Num2BaseFunction _num2octFunction;
        Base2NumFunction _hex2numFunction;
        Base2NumFunction _bin2numFunction;
        Base2NumFunction _oct2numFunction;

        // Roman numeral conversions. Separate classes because Roman
        // is not a positional system - the parameterised Base*Function
        // pattern doesn't fit.
        Num2RomFunction _num2romFunction;
        Rom2NumFunction _rom2numFunction;

        // The ecmd("path") loader callable, registered with the symbol
        // table at initialize().
        EcmdLoaderFunction _ecmdLoaderFunction;

        // The set of user functions loaded during this Replace-All run.
        // Discarded and re-created in beginRun() so .ecmd files are
        // re-read fresh each run. Null between shutdown() and the next
        // initialize().
        std::unique_ptr<EcmdLibrary> _ecmdLibrary;
    };

} // namespace MultiReplaceEngine