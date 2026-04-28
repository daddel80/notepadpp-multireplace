// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ExprTkEngine.cpp
// Implementation of the ExprTk-backed formula engine. See header for
// the surface contract; here we focus on the pipeline: parse template
// -> compile each expression -> per match update variables -> evaluate
// expressions -> assemble output.

#include "ExprTkEngine.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cmath>
#include <cstring>
#include <system_error>

namespace MultiReplaceEngine {

    // ---------------------------------------------------------------------
    // Construction / destruction
    // ---------------------------------------------------------------------

    ExprTkEngine::ExprTkEngine(ILuaEngineHost* host)
        : _host(host)
        , _regFunction(this)
    {
    }

    ExprTkEngine::~ExprTkEngine()
    {
        shutdown();
    }

    // ---------------------------------------------------------------------
    // Lifecycle
    // ---------------------------------------------------------------------

    bool ExprTkEngine::initialize()
    {
        // Register the numeric variables. ExprTk binds them by reference,
        // so the underlying members must live as long as any compiled
        // expression that uses them - which is fine, they are members of
        // the engine itself.
        _symbolTable.add_variable("CNT", _varCNT);
        _symbolTable.add_variable("LCNT", _varLCNT);
        _symbolTable.add_variable("LINE", _varLINE);
        _symbolTable.add_variable("LPOS", _varLPOS);
        _symbolTable.add_variable("APOS", _varAPOS);
        _symbolTable.add_variable("COL", _varCOL);

        // Register the reg(N) function. The wrapper holds a back-pointer
        // to the engine, so it can read from _captures during eval.
        _symbolTable.add_function("reg", _regFunction);

        // ExprTk's standard math constants (pi, epsilon, infinity).
        _symbolTable.add_constants();

        return true;
    }

    void ExprTkEngine::shutdown()
    {
        _compiledExpressions.clear();
        _parsedTemplate = ExprTkPatternParser::ParseResult();
        _lastCompiledScript.clear();
        _haveCompiled = false;

        // We deliberately do NOT clear the symbol table here - if the
        // engine gets compile()'d again later, the same variable bindings
        // are still valid. Symbol-table teardown happens implicitly when
        // the engine itself is destroyed.
    }

    // ---------------------------------------------------------------------
    // Error reporting
    // ---------------------------------------------------------------------

    void ExprTkEngine::reportError(ILuaEngineHost::ErrorCategory category,
        const std::string& details)
    {
        if (!_host || !_host->isLuaErrorDialogEnabled()) {
            return;
        }
        // Prefix with the engine name so the user can tell at a glance
        // which engine raised the error. "ExprTk:" is intentionally not
        // localised - "ExprTk" is a proper noun and the colon delimits
        // the localisable body that the host's translator will format.
        _host->showErrorMessage(category, "ExprTk: " + details);
    }

    // ---------------------------------------------------------------------
    // Compile
    // ---------------------------------------------------------------------

    bool ExprTkEngine::compile(const std::string& scriptUtf8)
    {
        // Cache hit: the same template was compiled last time, nothing to
        // do. Mirrors the behaviour of LuaEngine::ensureCompiled.
        if (_haveCompiled && scriptUtf8 == _lastCompiledScript) {
            return true;
        }

        // Drop any previous state before we attempt a new compile, so a
        // failed compile leaves the engine in a clean "no compile yet"
        // state rather than half-populated leftovers.
        _compiledExpressions.clear();
        _parsedTemplate = ExprTkPatternParser::ParseResult();
        _lastCompiledScript.clear();
        _haveCompiled = false;

        // Step 1: split the template into literal/expression segments.
        auto parseRes = ExprTkPatternParser::parse(scriptUtf8);
        if (!parseRes.success) {
            // Parser-level error (unmatched (?=, empty expression).
            // Surface to the user with positional context so they can
            // locate the problem in their replace template.
            std::string msg = parseRes.errorMessage;
            msg += " (at position ";
            msg += std::to_string(parseRes.errorPos);
            msg += ")";
            reportError(ILuaEngineHost::ErrorCategory::CompileError, msg);
            return false;
        }

        // Step 2: pre-compile every Expression segment. We allocate one
        // expression per segment; literals are not compiled.
        _compiledExpressions.resize(parseRes.segments.size());
        for (std::size_t i = 0; i < parseRes.segments.size(); ++i) {
            const auto& seg = parseRes.segments[i];
            if (seg.type != ExprTkPatternParser::SegmentType::Expression) {
                continue;
            }

            expression_t expr;
            expr.register_symbol_table(_symbolTable);

            if (!_parser.compile(seg.text, expr)) {
                // ExprTk failed to compile this expression. Surface the
                // ExprTk parser error verbatim - it is technical but
                // precise (e.g. "Unknown variable or function: foo").
                std::string msg = "Invalid expression \"";
                msg += seg.text;
                msg += "\": ";
                msg += _parser.error();
                reportError(ILuaEngineHost::ErrorCategory::CompileError, msg);

                _compiledExpressions.clear();
                return false;
            }

            _compiledExpressions[i] = std::move(expr);
        }

        // Step 3: cache for re-use.
        _parsedTemplate = std::move(parseRes);
        _lastCompiledScript = scriptUtf8;
        _haveCompiled = true;
        return true;
    }

    // ---------------------------------------------------------------------
    // Execute
    // ---------------------------------------------------------------------

    FormulaResult ExprTkEngine::execute(
        const std::string& scriptUtf8,
        const FormulaVars& vars,
        bool /*isRegexMatch*/,
        int  /*documentCodepage*/)
    {
        FormulaResult result;

        // Lazy compile: if compile() was never called, or the script has
        // changed since the last compile, run it now. compile() has
        // already shown any error dialog; we just propagate the failure
        // through the FormulaResult so the pipeline can stop the run.
        if (!_haveCompiled || scriptUtf8 != _lastCompiledScript) {
            if (!compile(scriptUtf8)) {
                result.success = false;
                result.errorMessage = "ExprTk: compile failed";
                return result;
            }
        }

        // Update per-match numeric variables. ExprTk reads them by
        // reference at eval time, so writing here is enough.
        _varCNT = static_cast<double>(vars.CNT);
        _varLCNT = static_cast<double>(vars.LCNT);
        _varLINE = static_cast<double>(vars.LINE);
        _varLPOS = static_cast<double>(vars.LPOS);
        _varAPOS = static_cast<double>(vars.APOS);
        _varCOL = static_cast<double>(vars.COL);

        // Pre-parse captures into doubles. Index 0 = full match (MATCH),
        // index 1..N = capture groups. Doing this once per match is
        // cheaper than re-parsing inside reg() for every reg(N) call.
        _captures.clear();
        _captures.reserve(vars.captures.size() + 1);
        _captures.push_back(parseCaptureToDouble(vars.MATCH));
        for (const auto& cap : vars.captures) {
            _captures.push_back(parseCaptureToDouble(cap));
        }

        // Walk segments in order, accumulating the output. Literals copy
        // straight through; expressions get evaluated and formatted.
        std::string out;
        out.reserve(scriptUtf8.size());

        const auto& segs = _parsedTemplate.segments;
        for (std::size_t i = 0; i < segs.size(); ++i) {
            const auto& seg = segs[i];

            if (seg.type == ExprTkPatternParser::SegmentType::Literal) {
                out.append(seg.text);
                continue;
            }

            // Expression segment - evaluate and format.
            const double value = _compiledExpressions[i].value();
            out.append(formatDouble(value));
        }

        result.output = std::move(out);
        result.success = true;
        return result;
    }

    // ---------------------------------------------------------------------
    // Help URL
    // ---------------------------------------------------------------------

    std::wstring ExprTkEngine::helpUrl() const
    {
        // Points at the user-facing documentation section. Resolved at
        // call time so we don't bake an outdated URL into a binary.
        return L"https://github.com/daddel80/notepadpp-multireplace#exprtk-engine";
    }

    // ---------------------------------------------------------------------
    // String <-> double helpers
    // ---------------------------------------------------------------------

    double ExprTkEngine::parseCaptureToDouble(const std::string& s)
    {
        if (s.empty()) {
            return 0.0;
        }

        // std::from_chars is locale-independent: it always treats '.' as
        // the decimal separator and never reads thousand-separators. That
        // is exactly what we want for a programming-language style number.
        double value = 0.0;
        const char* first = s.data();
        const char* last = s.data() + s.size();

        auto res = std::from_chars(first, last, value);
        if (res.ec != std::errc{}) {
            // Non-numeric input. Could be intentional ("price=$5.00")
            // or a user mistake; in either case, returning 0.0 keeps
            // expressions evaluable rather than aborting the whole match.
            return 0.0;
        }

        // We do not require all of `s` to have been consumed; trailing
        // junk like "1.5abc" yields 1.5. This is consistent with how
        // most programming languages read numeric prefixes.
        return value;
    }

    std::string ExprTkEngine::formatDouble(double value)
    {
        // Special floats that std::to_chars handles but we want to spell
        // out explicitly so the output is recognisable in a replace
        // context (rather than implementation-defined "nan" capitalisation).
        if (std::isnan(value)) {
            return "nan";
        }
        if (std::isinf(value)) {
            return value < 0.0 ? "-inf" : "inf";
        }

        // Use std::to_chars in its "shortest round-trip" mode (no format
        // argument) so 1.0 -> "1", 1.5 -> "1.5", 1e20 -> "1e+20", etc.
        // 64 bytes is more than enough for any double.
        std::array<char, 64> buf{};
        auto res = std::to_chars(buf.data(), buf.data() + buf.size(), value);
        if (res.ec != std::errc{}) {
            // Fallback (should not happen with the buffer size above).
            return "nan";
        }
        return std::string(buf.data(), res.ptr);
    }

    // ---------------------------------------------------------------------
    // reg() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::RegFunction::operator()(const double& index)
    {
        // Index is passed as double; floor to integer. Negative or non-
        // integer indices are treated as out-of-range (return 0.0) rather
        // than thrown, so a malformed expression like reg(-1) doesn't
        // abort the whole match.
        if (!std::isfinite(index)) {
            return 0.0;
        }

        // Truncate toward zero. reg(1.5) becomes reg(1); reg(-0.5) -> 0.
        const long long idx = static_cast<long long>(index);
        if (idx < 0) {
            return 0.0;
        }

        const auto& caps = _owner->_captures;
        if (static_cast<std::size_t>(idx) >= caps.size()) {
            return 0.0;
        }

        return caps[static_cast<std::size_t>(idx)];
    }

} // namespace MultiReplaceEngine