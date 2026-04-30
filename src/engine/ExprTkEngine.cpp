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
#include <iomanip>
#include <sstream>
#include <system_error>

// SU::escapeControlChars is shared with LuaEngine for debug-window
// display: capture strings may contain TAB/NL/etc, which would break
// the tab-separated table format the host's debug dialog renders.
#include "../StringUtils.h"
namespace SU = StringUtils;

namespace MultiReplaceEngine {

    // ---------------------------------------------------------------------
    // Construction / destruction
    // ---------------------------------------------------------------------

    ExprTkEngine::ExprTkEngine(ILuaEngineHost* host)
        : _host(host)
        , _regFunction(this)
        , _skipFunction(this)
        , _rstrFunction(this)
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

        // Register string-typed variables for match/file metadata. Same
        // reference-binding contract as numeric vars (Section 13 of the
        // ExprTk docs). The strings are refreshed at the start of every
        // execute() call so each match sees its own values.
        _symbolTable.add_stringvar("MATCH", _strMATCH);
        _symbolTable.add_stringvar("FPATH", _strFPATH);
        _symbolTable.add_stringvar("FNAME", _strFNAME);

        // Register the reg(N) function. The wrapper holds a back-pointer
        // to the engine, so it can read from _captures during eval.
        _symbolTable.add_function("reg", _regFunction);

        // Register skip() so users can express conditional replacement
        // in a numeric expression (e.g. inside if/else). The function
        // returns 0.0 and signals via _wantSkip; execute() picks that
        // up after eval and propagates it through FormulaResult::skip.
        _symbolTable.add_function("skip", _skipFunction);

        // Register rstr(N) for string-typed capture access. Only useful
        // inside an ExprTk return list, where its string output is
        // concatenated into the replace text.
        _symbolTable.add_function("rstr", _rstrFunction);

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
        // Engine identifier is passed separately so the host can compose
        // the user-visible "ExprTk: ..." prefix from a translation
        // template instead of a hardcoded literal here.
        _host->showErrorMessage(category, "ExprTk", details);
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
        bool /*isRegexMatch*/,    // see note in result-construction below
        int  /*documentCodepage*/)
    {
        FormulaResult result;

        // Note on isRegexMatch:
        // The host pipeline routes engine output through Scintilla's plain
        // replace path (SCI_REPLACETARGET) regardless of regex mode, so no
        // pre-escaping of format-string special characters is needed here.
        // The parameter is preserved on the IFormulaEngine boundary for
        // future use (e.g. diagnostics) but currently has no effect.

        // Lazy compile: if compile() was never called, or the script has
        // changed since the last compile, run it now. compile() has
        // already shown any error dialog; we just propagate the failure
        // through the FormulaResult so the pipeline can stop the run.
        if (!_haveCompiled || scriptUtf8 != _lastCompiledScript) {
            if (!compile(scriptUtf8)) {
                result.success = false;
                // Internal diagnostic only - the user-visible dialog was
                // already raised by compile() via reportError().
                result.errorMessage = "compile failed";
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

        // Update per-match string variables. ExprTk holds these by
        // reference too, so writing into the members is enough; any
        // compiled expression that mentions MATCH/FPATH/FNAME will
        // pick up the new value on the next eval.
        _strMATCH = vars.MATCH;
        _strFPATH = vars.FPATH;
        _strFNAME = vars.FNAME;

        // Reset skip flag so a previous match's skip() call cannot
        // leak into this one. skip() will set it back to true if the
        // user calls it during eval below.
        _wantSkip = false;

        // Pre-parse captures into doubles. Index 0 = full match (MATCH),
        // index 1..N = capture groups. Doing this once per match is
        // cheaper than re-parsing inside reg() for every reg(N) call.
        _captures.clear();
        _captures.reserve(vars.captures.size() + 1);
        _captures.push_back(parseCaptureToDouble(vars.MATCH));
        for (const auto& cap : vars.captures) {
            _captures.push_back(parseCaptureToDouble(cap));
        }

        // Mirror the captures as raw strings for rstr(N). RstrFunction
        // reads from this vector indexed 0..N-1 for capture groups
        // 1..N. The full match (rstr(0)) is served from _strMATCH
        // directly and not duplicated here.
        _captureStrings = vars.captures;

        // ----- Debug-window display ---------------------------------------
        // When debug mode is on, surface a per-match snapshot of the
        // numeric variables and capture strings before evaluating the
        // expressions. The format mirrors the LuaEngine output so the
        // host's debug dialog renders both engines identically:
        //   <name>\t<Type>\t<value>\n\n
        //
        // ExprTk has no per-script DEBUG override (no user-defined
        // variables), so the global host toggle is the only switch.
        if (_host && _host->isDebugModeEnabled()) {
            std::ostringstream dbg;
            // Match LuaEngine's format precision so both engines render
            // numbers identically in the host's debug dialog.
            dbg << std::fixed << std::setprecision(8);

            // Numeric variables. Always emitted in a fixed order so the
            // dialog reads consistently across matches.
            auto emitNumber = [&dbg](const char* name, double value) {
                dbg << name << "\tNumber\t" << value << "\n\n";
                };
            emitNumber("CNT", _varCNT);
            emitNumber("LCNT", _varLCNT);
            emitNumber("LINE", _varLINE);
            emitNumber("LPOS", _varLPOS);
            emitNumber("APOS", _varAPOS);
            emitNumber("COL", _varCOL);

            // String variables. Same Tab-separated format as numbers so
            // both render identically in the debug dialog. FPATH/FNAME
            // can be empty in some contexts (no file path known) - we
            // still emit them so the dialog is predictable.
            auto emitString = [&dbg](const char* name, const std::string& value) {
                dbg << name << "\tString\t"
                    << SU::escapeControlChars(value) << "\n\n";
                };
            emitString("MATCH", vars.MATCH);
            emitString("FPATH", vars.FPATH);
            emitString("FNAME", vars.FNAME);

            // Capture strings. reg(0) is the full match; reg(1..N) are the
            // capture groups. We emit the original capture string (not the
            // double the engine sees) so the user can see exactly what the
            // pattern matched, including non-numeric content.
            dbg << "reg(0)\tString\t"
                << SU::escapeControlChars(vars.MATCH) << "\n\n";
            for (std::size_t i = 0; i < vars.captures.size(); ++i) {
                dbg << "reg(" << (i + 1) << ")\tString\t"
                    << SU::escapeControlChars(vars.captures[i]) << "\n\n";
            }

            // Refresh the panel's list view so any state the debug dialog
            // shows reflects the current run state, then block on the
            // dialog. Response codes mirror Lua's contract:
            //   3  -> user pressed "Stop"
            //  -1  -> dialog closed (e.g. window-close button)
            //  any other -> continue
            _host->refreshUiListView();
            const int resp = _host->showDebugWindow(dbg.str());
            if (resp == 3 || resp == -1) {
                result.success = false;
                result.errorMessage = "Aborted via debug window";
                return result;
            }
        }

        // Walk segments in order, accumulating the output. Literals copy
        // straight through; expressions get evaluated and either formatted
        // as a number (legacy path) or unpacked from an ExprTk return
        // statement (new path - lets users emit mixed string/number
        // output, which is the only way to surface a string variable like
        // FNAME / FPATH / MATCH).
        std::string out;
        out.reserve(scriptUtf8.size());

        const auto& segs = _parsedTemplate.segments;
        for (std::size_t i = 0; i < segs.size(); ++i) {
            const auto& seg = segs[i];

            if (seg.type == ExprTkPatternParser::SegmentType::Literal) {
                out.append(seg.text);
                continue;
            }

            // Eval. ExprTk's return statement turns into a side-channel
            // result list - we must check return_invoked() to decide
            // which output path applies, because value() is 0.0 in that
            // case and would otherwise look like a normal numeric eval.
            auto& expr = _compiledExpressions[i];
            const double value = expr.value();

            if (expr.return_invoked()) {
                appendExprtkResults(expr, out);
                continue;
            }

            // Legacy numeric path. If the user accidentally placed a
            // string-typed variable (FNAME, FPATH, MATCH) into a numeric
            // expression - or expected ExprTk to coerce a string to a
            // number - the eval produces NaN. Surface a useful error
            // instead of writing "nan" into the document.
            if (std::isnan(value)) {
                std::string msg =
                    "Expression \"" + seg.text + "\" evaluated to NaN. "
                    "ExprTk cannot use a string variable (MATCH, FPATH, "
                    "FNAME) where a number is expected. To emit a string, "
                    "wrap it in a return statement: "
                    "(?=return ['prefix-', FNAME]).";
                reportError(ILuaEngineHost::ErrorCategory::ExecutionError,
                    msg);
                result.success = false;
                result.errorMessage = "Eval produced NaN";
                return result;
            }
            out.append(formatDouble(value));
        }

        result.output = std::move(out);
        result.success = true;
        // If the user invoked skip() anywhere in the eval (even nested
        // in a conditional), tell the pipeline to leave the current
        // match untouched. The output we built above is then thrown
        // away by the caller; this matches LuaEngine's contract.
        result.skip = _wantSkip;
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
    // ExprTk return-statement unpacking
    // ---------------------------------------------------------------------

    void ExprTkEngine::appendExprtkResults(const expression_t& expr,
        std::string& out)
    {
        // ExprTk's return statement (Section 20 of the docs) turns the
        // expression into one that produces an ordered, typed result list
        // instead of a single scalar. results() returns a results_context
        // describing that list; each element is a type_store carrying its
        // type tag plus a typed view onto the actual data.
        const auto& results = expr.results();

        using results_ctx_t = typename std::remove_reference<decltype(results)>::type;
        using type_store_t = typename results_ctx_t::type_store_t;
        using string_view_t = typename type_store_t::string_view;
        using scalar_view_t = typename type_store_t::scalar_view;

        for (std::size_t i = 0; i < results.count(); ++i) {
            const type_store_t& ts = results[i];

            switch (ts.type) {
            case type_store_t::e_scalar: {
                // Numeric element - format like a regular numeric
                // expression so 1.0 -> "1", 1.5 -> "1.5", etc.
                const scalar_view_t sv(ts);
                out.append(formatDouble(sv()));
                break;
            }

            case type_store_t::e_string: {
                // String element - append verbatim. ExprTk strings
                // are 8-bit char sequences (UTF-8 in our pipeline).
                const string_view_t sv(ts);
                out.append(&sv[0], sv.size());
                break;
            }

            case type_store_t::e_vector:
                // Vectors are unusual in a replace template. We could
                // unpack them element-wise, but the use case is thin
                // and the cost of guessing wrong is a corrupted
                // replace. Skip for now; if someone needs it, we add
                // it once a concrete user pattern shows up.
                break;

            default:
                // Forwards-compat: silently ignore unknown element
                // types so a future ExprTk extension cannot break a
                // running replace.
                break;
            }
        }
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

    // ---------------------------------------------------------------------
    // skip() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::SkipFunction::operator()()
    {
        // The actual work is a side-effect on the engine: flip the flag
        // so execute() can transmit it through FormulaResult::skip after
        // the eval loop completes. We return 0.0 because ExprTk requires
        // a numeric value; users typically place skip() inside an if/else
        // branch where the returned value is discarded anyway.
        if (_owner) {
            _owner->_wantSkip = true;
        }
        return 0.0;
    }

    // ---------------------------------------------------------------------
    // rstr() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::RstrFunction::operator()(
        std::string& result,
        parameter_list_t parameters)
    {
        // Default to empty: an out-of-range or malformed call yields ""
        // rather than aborting eval, mirroring reg(N)'s defensive style.
        result.clear();

        if (!_owner || parameters.size() != 1) {
            return 0.0;
        }

        // The "T" arity-spec promised one Scalar argument; wrap it in a
        // scalar_view to read the value. The view is a thin reference
        // wrapper so this is cheap.
        const scalar_t s(parameters[0]);
        const double index = s();

        if (!std::isfinite(index)) {
            return 0.0;
        }

        // Truncate toward zero so rstr(1.7) means rstr(1). Negative
        // indices fall through to the empty-string default.
        const long long idx = static_cast<long long>(index);
        if (idx < 0) {
            return 0.0;
        }

        const std::size_t uidx = static_cast<std::size_t>(idx);

        // rstr(0) is the full match. _strMATCH is refreshed at the top
        // of execute() before any expression is evaluated, so it always
        // holds the current match's text here.
        if (uidx == 0) {
            result = _owner->_strMATCH;
            return 0.0;
        }

        // rstr(N) for N >= 1 reads capture group N, stored 0-based in
        // _captureStrings. Beyond the last capture we return the empty
        // string - same defensive behaviour as reg() for numeric
        // out-of-range.
        if (uidx - 1 < _owner->_captureStrings.size()) {
            result = _owner->_captureStrings[uidx - 1];
        }

        // Numeric return value is discarded by ExprTk in e_rtrn_string
        // mode - the meaningful output is the string we just filled.
        return 0.0;
    }

} // namespace MultiReplaceEngine