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
#include <ctime>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <limits>
#include <sstream>
#include <system_error>

#include "../StringUtils.h"
#include "../exprtk/DateParse.h"
#include "../Encoding.h"
namespace SU = StringUtils;

namespace MultiReplaceEngine {

    // ---------------------------------------------------------------------
    // Construction / destruction
    // ---------------------------------------------------------------------

    ExprTkEngine::ExprTkEngine(ILuaEngineHost* host)
        : _host(host)
        , _numFunction(this)
        , _skipFunction(this)
        , _txtFunction(this)
        , _seqFunction(this)
        , _numColFunction(this)
        , _txtColFunction(this)
        , _parsedateFunction(this)
        , _ecmdLoaderFunction(this)
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
        // Each variable is registered under both upper- and lowercase
        // so users may write whichever style they prefer.
        _symbolTable.add_variable("CNT", _varCNT);
        _symbolTable.add_variable("cnt", _varCNT);
        _symbolTable.add_variable("LCNT", _varLCNT);
        _symbolTable.add_variable("lcnt", _varLCNT);
        _symbolTable.add_variable("LINE", _varLINE);
        _symbolTable.add_variable("line", _varLINE);
        _symbolTable.add_variable("LPOS", _varLPOS);
        _symbolTable.add_variable("lpos", _varLPOS);
        _symbolTable.add_variable("APOS", _varAPOS);
        _symbolTable.add_variable("apos", _varAPOS);
        _symbolTable.add_variable("COL", _varCOL);
        _symbolTable.add_variable("col", _varCOL);
        _symbolTable.add_variable("HIT", _varHIT);
        _symbolTable.add_variable("hit", _varHIT);

        _symbolTable.add_stringvar("FPATH", _strFPATH);
        _symbolTable.add_stringvar("fpath", _strFPATH);
        _symbolTable.add_stringvar("FNAME", _strFNAME);
        _symbolTable.add_stringvar("fname", _strFNAME);

        // Register the num(N) function. The wrapper holds a back-pointer
        // to the engine, so it can read from _captures during eval.
        _symbolTable.add_function("num", _numFunction);

        // Register skip() so users can express conditional replacement
        // in a numeric expression (e.g. inside if/else). The function
        // returns 0.0 and signals via _wantSkip; execute() picks that
        // up after eval and propagates it through FormulaResult::skip.
        _symbolTable.add_function("skip", _skipFunction);

        // Register txt(N) for string-typed capture access. Only useful
        // inside an ExprTk return list, where its string output is
        // concatenated into the replace text.
        _symbolTable.add_function("txt", _txtFunction);

        // Register seq([start, [inc]]) as a sequence generator. Reads
        // _varCNT directly so the value tracks the current match index;
        // both arguments are optional and default to 1, so seq() yields
        // 1,2,3,... seq(start) yields start, start+1, ... and seq(start,
        // inc) is fully parametrised. Built on ivararg_function so the
        // optional-argument count is checked at call time.
        _symbolTable.add_function("seq", _seqFunction);

        // CSV column accessors. Both delegate to the host, which knows
        // about the active CSV settings, the line of the current match,
        // and caches the row's column array for repeated lookups.
        _symbolTable.add_function("numcol", _numColFunction);
        _symbolTable.add_function("txtcol", _txtColFunction);

        // Register parsedate(str, fmt) - the inverse of D[fmt] output.
        // Returns Unix timestamp on success, NaN on parse failure.
        _symbolTable.add_function("parsedate", _parsedateFunction);

        // Register ecmd("path") - the library loader. Functions loaded
        // through it land in _ecmdLibrary's own symbol_table, which we
        // register against every compiled expression in compile().
        _symbolTable.add_function("ecmd", _ecmdLoaderFunction);
        _ecmdLibrary = std::make_unique<EcmdLibrary>();

        // ExprTk's standard math constants (pi, epsilon, infinity).
        _symbolTable.add_constants();

        return true;
    }

    void ExprTkEngine::shutdown()
    {
        _compiledExpressions.clear();
        _segmentSpecs.clear();
        _parsedTemplate = ExprTkPatternParser::ParseResult();
        _lastCompiledScript.clear();
        _haveCompiled = false;
        _ecmdLibrary.reset();

        // We deliberately do NOT clear the symbol table here - if the
        // engine gets compile()'d again later, the same variable bindings
        // are still valid. Symbol-table teardown happens implicitly when
        // the engine itself is destroyed.
    }

    // ---------------------------------------------------------------------
    // Per-run lifecycle
    // ---------------------------------------------------------------------

    void ExprTkEngine::beginRun()
    {
        IFormulaEngine::beginRun();

        // Drop the previous run's ecmd library and start fresh. This is
        // what makes ecmd("path") re-read the file on every Replace-All
        // (user edits to the .ecmd take effect on the next run) and what
        // makes removing the ecmd() init slot also remove its functions.
        //
        // Side effect: the compile cache is invalidated because the
        // library symbol_table now refers to a different object, so any
        // previously compiled expression that called an ecmd function
        // would point at freed registrations.
        _ecmdLibrary = std::make_unique<EcmdLibrary>();
        _compiledExpressions.clear();
        _segmentSpecs.clear();
        _parsedTemplate = ExprTkPatternParser::ParseResult();
        _lastCompiledScript.clear();
        _haveCompiled = false;
    }

    // ---------------------------------------------------------------------
    // Error reporting
    // ---------------------------------------------------------------------

    void ExprTkEngine::reportError(ILuaEngineHost::ErrorCategory category,
        const std::string& details)
    {
        if (!_host || !_host->isFormulaErrorDialogEnabled()) {
            return;
        }
        _host->showErrorMessage(category, "ExprTk", details);
    }

    void ExprTkEngine::handleInvalid(const std::string& exprText)
    {
        // Delegate the entire skip-state and dialog plumbing to the base
        // class. The engine-specific bit is the translation key for the
        // detail-text shown in the dialog body. Triggered on both NaN
        // and Inf, since both indicate an unusable numeric result and
        // letting either reach the output would silently corrupt the
        // text.
        handleRecoverableSkip(_host, L"ExprTk",
            L"msgbox_recoverable_error_details_exprtk", exprText);
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
        _segmentSpecs.clear();
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
        _segmentSpecs.assign(parseRes.segments.size(), SegmentSpec{});
        for (std::size_t i = 0; i < parseRes.segments.size(); ++i) {
            const auto& seg = parseRes.segments[i];
            if (seg.type != ExprTkPatternParser::SegmentType::Expression) {
                continue;
            }

            // Split the segment body at the last unquoted '~' outside
            // brackets. Anything left of it is the formula; anything
            // right is an optional output-format spec.
            auto split = FormatSpec::splitFormulaSpec(seg.text);

            if (split.hasSpec) {
                FormatSpec::Spec parsed = FormatSpec::parse(
                    std::wstring(split.spec.begin(), split.spec.end()));
                if (!parsed.valid) {
                    std::string msg = "Invalid format spec \"";
                    msg += split.spec;
                    msg += "\": ";
                    msg += parsed.errorMessage;
                    reportError(ILuaEngineHost::ErrorCategory::CompileError, msg);
                    _compiledExpressions.clear();
                    _segmentSpecs.clear();
                    return false;
                }
                _segmentSpecs[i].hasSpec = true;
                _segmentSpecs[i].spec = std::move(parsed);
            }

            expression_t expr;
            expr.register_symbol_table(_symbolTable);
            if (_ecmdLibrary) {
                // Calls into ecmd-loaded functions resolve against the
                // library's own symbol table. Registering it here lets
                // user expressions like (?=num2rom(num(1))) compile
                // against whatever the current Replace-All run has loaded.
                expr.register_symbol_table(_ecmdLibrary->symbolTable());
            }

            if (!_parser.compile(split.formula, expr)) {
                // ExprTk failed to compile this expression. Surface the
                // ExprTk parser error verbatim - it is technical but
                // precise (e.g. "Unknown variable or function: foo").
                std::string msg = "Invalid expression \"";
                msg += split.formula;
                msg += "\": ";
                msg += _parser.error();
                reportError(ILuaEngineHost::ErrorCategory::CompileError, msg);

                _compiledExpressions.clear();
                _segmentSpecs.clear();
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
        bool isRegexMatch,
        int  /*documentCodepage*/)
    {
        FormulaResult result;

        // When the caller is in regex mode the rendered output will be
        // routed through a regex replacement engine. Expression results
        // (formula output) need to be escaped so any literal \ or $ from
        // the formula stays literal and is not misread as a backreference.
        // Literal segments outside (?=...) blocks are left unescaped on
        // purpose - that is what lets the user write \1 / $1 verbatim in
        // the replace template and have them expand normally.
        const bool escapeOutput = isRegexMatch;

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

        _strMATCH = vars.MATCH;
        _strFPATH = vars.FPATH;
        _strFNAME = vars.FNAME;

        _wantSkip = false;
        _wantStop = false;
        _outputHadInvalid = false;

        _captures.clear();
        _captures.reserve(vars.captures.size() + 1);
        _captures.push_back(parseCaptureToDouble(vars.MATCH));
        for (const auto& cap : vars.captures) {
            _captures.push_back(parseCaptureToDouble(cap));
        }
        _varHIT = _captures[0];

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
            emitNumber("HIT", _varHIT);

            auto emitString = [&dbg](const char* name, const std::string& value) {
                dbg << name << "\tString\t"
                    << SU::escapeControlChars(value) << "\n\n";
                };
            emitString("FPATH", vars.FPATH);
            emitString("FNAME", vars.FNAME);

            dbg << "num(0)\tString\t"
                << SU::escapeControlChars(vars.MATCH) << "\n\n";
            for (std::size_t i = 0; i < vars.captures.size(); ++i) {
                dbg << "num(" << (i + 1) << ")\tString\t"
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
        // as a number or unpacked from an ExprTk return statement, which
        // lets users emit mixed string/number output (the only way to
        // surface a string variable like FNAME / FPATH).
        std::string out;
        out.reserve(scriptUtf8.size());

        // Helper: build the FormulaResult for an invalid-result match
        // (NaN or Inf). Used by both the numeric and return-list paths.
        // Sets success=false when the user picked "Stop" in the dialog;
        // otherwise marks the match as skipped so the rest of the run
        // can continue. Either way, the original match text stays
        // intact - no data is lost.
        const auto onInvalid = [&](FormulaResult& r, const std::string& exprText) {
            handleInvalid(exprText);
            if (_wantStop) {
                r.success = false;
                r.errorMessage = "Aborted via formula-error dialog";
                return;
            }
            r.skip = true;
            r.success = true;
            r.outputIsRegexSafe = isRegexMatch;
            };

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
                // return-lists emit strings; numeric spec makes no sense.
                if (_segmentSpecs[i].hasSpec) {
                    std::string msg = "Format spec cannot be combined with return statement: \"";
                    msg += seg.text;
                    msg += "\"";
                    reportError(ILuaEngineHost::ErrorCategory::ExecutionError, msg);
                    result.success = false;
                    result.errorMessage = std::move(msg);
                    return result;
                }
                appendExprtkResults(expr, out, escapeOutput);
                if (_outputHadInvalid) {
                    onInvalid(result, seg.text);
                    return result;
                }
                continue;
            }

            // NaN or Inf both indicate an unusable numeric result. Route
            // either through the recoverable-error dialog instead of
            // letting "nan" / "inf" leak into the output text.
            if (std::isnan(value) || std::isinf(value)) {
                onInvalid(result, seg.text);
                return result;
            }

            // Format through spec, or fall back to shortest round-trip.
            // FormatSpec only emits ASCII chars (digits, sign, dot,
            // letters, colon, space), so each wchar fits in a char.
            if (_segmentSpecs[i].hasSpec) {
                std::wstring formatted = FormatSpec::apply(_segmentSpecs[i].spec, value);
                out.reserve(out.size() + formatted.size());
                for (wchar_t wc : formatted) {
                    out.push_back(static_cast<char>(wc));
                }
            }
            else {
                out.append(formatDouble(value));
            }
        }

        result.output = std::move(out);
        result.success = true;
        result.skip = _wantSkip;
        result.outputIsRegexSafe = isRegexMatch;
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
        std::string& out,
        bool escapeForRegex)
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
                const scalar_view_t sv(ts);
                const double v = sv();
                // Same invalid-result rule as the numeric path: NaN or
                // Inf inside a return list aborts the whole list. The
                // caller picks this up via _outputHadInvalid and routes
                // through the recoverable-error dialog.
                if (std::isnan(v) || std::isinf(v)) {
                    _outputHadInvalid = true;
                    return;
                }
                out.append(formatDouble(v));
                break;
            }

            case type_store_t::e_string: {
                // String element - append verbatim. ExprTk strings
                // are 8-bit char sequences (UTF-8 in our pipeline).
                // When the host will route this through a regex
                // replace, escape \ and $ so any literal backslash or
                // dollar from the formula stays literal and is not
                // misread as a backreference (\1, $1, $&, ...).
                const string_view_t sv(ts);
                if (escapeForRegex && _host) {
                    std::string raw(&sv[0], sv.size());
                    out.append(_host->escapeForRegex(raw));
                }
                else {
                    out.append(&sv[0], sv.size());
                }
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
            return std::numeric_limits<double>::quiet_NaN();
        }

        // Accept both '.' and ',' as decimal separator. If the string
        // contains a '.', commas are left alone so from_chars stops at
        // the first comma (preserving trailing-junk semantics for inputs
        // like "1.5,extra"). If there is no '.', any ',' is normalised
        // to '.' so comma-decimal locales parse correctly.
        std::string buf;
        const char* first = s.data();
        const char* last = s.data() + s.size();

        if (s.find('.') == std::string::npos &&
            s.find(',') != std::string::npos) {
            buf.assign(s);
            std::replace(buf.begin(), buf.end(), ',', '.');
            first = buf.data();
            last = buf.data() + buf.size();
        }

        double value = 0.0;
        auto res = std::from_chars(first, last, value);
        if (res.ec != std::errc{}) {
            // Non-numeric input - return NaN so the formula author can
            // either propagate it (NaN-as-error) or guard explicitly
            // with a NaN check (x != x). Silent 0.0 fallback would
            // corrupt data when matches contain unexpected text.
            return std::numeric_limits<double>::quiet_NaN();
        }

        // Trailing junk like "1.5abc" yields 1.5. Consistent with how
        // most programming languages read numeric prefixes.
        return value;
    }

    std::string ExprTkEngine::formatDouble(double value)
    {
        // Callers are expected to filter NaN and Inf before reaching
        // here (see the isnan/isinf checks in execute() and
        // appendExprtkResults()). If one slips through anyway, return
        // an empty string instead of letting the literal "inf" / "nan"
        // bleed into the user's text - that would defeat the whole
        // recoverable-error story.
        if (std::isnan(value) || std::isinf(value)) {
            return std::string{};
        }

        std::array<char, 64> buf{};
        auto res = std::to_chars(buf.data(), buf.data() + buf.size(), value);
        if (res.ec != std::errc{}) {
            return std::string{};
        }
        return std::string(buf.data(), res.ptr);
    }

    // ---------------------------------------------------------------------
    // num() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::NumFunction::operator()(const double& index)
    {
        // Index is passed as double; floor to integer. Negative or non-
        // integer indices are treated as out-of-range (return 0.0) rather
        // than thrown, so a malformed expression like num(-1) doesn't
        // abort the whole match.
        if (!std::isfinite(index)) {
            return 0.0;
        }

        // Truncate toward zero. num(1.5) becomes num(1); num(-0.5) -> 0.
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
    // txt() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::TxtFunction::operator()(
        std::string& result,
        parameter_list_t parameters)
    {
        // Default to empty: an out-of-range or malformed call yields ""
        // rather than aborting eval, mirroring num(N)'s defensive style.
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

        // Truncate toward zero so txt(1.7) means txt(1). Negative
        // indices fall through to the empty-string default.
        const long long idx = static_cast<long long>(index);
        if (idx < 0) {
            return 0.0;
        }

        const std::size_t uidx = static_cast<std::size_t>(idx);

        // txt(0) is the full match. _strMATCH is refreshed at the top
        // of execute() before any expression is evaluated, so it always
        // holds the current match's text here.
        if (uidx == 0) {
            result = _owner->_strMATCH;
            return 0.0;
        }

        // txt(N) for N >= 1 reads capture group N, stored 0-based in
        // _captureStrings. Beyond the last capture we return the empty
        // string - same defensive behaviour as num() for numeric
        // out-of-range.
        if (uidx - 1 < _owner->_captureStrings.size()) {
            result = _owner->_captureStrings[uidx - 1];
        }

        // Numeric return value is discarded by ExprTk in e_rtrn_string
        // mode - the meaningful output is the string we just filled.
        return 0.0;
    }

    // ---------------------------------------------------------------------
    // parsedate(str, fmt) -> Unix timestamp (or NaN on failure)
    // ---------------------------------------------------------------------

    double ExprTkEngine::ParsedateFunction::operator()(
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        if (parameters.size() != 2) {
            return nanResult;
        }

        // Both args must be strings. The .type tag on a generic_t is
        // an enum from exprtk::type_store; we compare its raw value.
        // generic_t::e_string is the matching enumerator (same pattern
        // as appendExprtkResults() uses on results[i].type).
        const auto t0 = parameters[0].type;
        const auto t1 = parameters[1].type;
        if (t0 != generic_t::e_string || t1 != generic_t::e_string) {
            return nanResult;
        }

        const string_t sv0(parameters[0]);
        const string_t sv1(parameters[1]);
        std::string input(sv0.begin(), sv0.end());
        std::string fmt(sv1.begin(), sv1.end());

        // Lua-style '!' prefix in the format means "treat result as UTC".
        bool utc = false;
        if (!fmt.empty() && fmt[0] == '!') {
            utc = true;
            fmt.erase(fmt.begin());
        }

        // Run the parser. tm starts zeroed so unset fields contribute
        // 0/Jan/1/midnight - sensible defaults.
        std::tm tm{};
        tm.tm_isdst = -1;  // let mktime figure out DST for local time
        const bool ok = MultiReplace::parseDateTime(input, fmt, tm);
        if (!ok) {
            return nanResult;
        }

        // Convert tm to time_t. Local time uses mktime (which honors the
        // current TZ); UTC needs a manual roll since C++17 has no portable
        // timegm.
        std::time_t t;
        if (utc) {
            // Manual day count from 1970-01-01 plus h/m/s.
            static const int days_per_month[] =
            { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            const long long year = tm.tm_year + 1900;
            const long long month = tm.tm_mon;       // 0..11
            const long long day = tm.tm_mday - 1;  // 0-indexed

            if (year < 1970) return nanResult;

            long long days = 0;
            for (long long y = 1970; y < year; ++y) {
                const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
                days += leap ? 366 : 365;
            }
            for (long long m = 0; m < month; ++m) {
                int dm = days_per_month[m];
                if (m == 1) {
                    const bool leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
                    if (leap) dm = 29;
                }
                days += dm;
            }
            days += day;

            const long long total = days * 86400
                + tm.tm_hour * 3600
                + tm.tm_min * 60
                + tm.tm_sec;
            t = static_cast<std::time_t>(total);
        }
        else {
            t = std::mktime(&tm);
            if (t == static_cast<std::time_t>(-1)) {
                return nanResult;
            }
        }

        return static_cast<double>(t);
    }

    // ---------------------------------------------------------------------
    // EcmdFunctionInstance
    // ---------------------------------------------------------------------

    std::string ExprTkEngine::EcmdFunctionInstance::makeParamSequence(
        const std::vector<EcmdParser::ParamDef>& params)
    {
        std::string seq;
        seq.reserve(params.size());
        for (const auto& p : params) {
            seq.push_back(static_cast<char>(p.type));
        }
        return seq;
    }

    ExprTkEngine::EcmdFunctionInstance::igenfunct_t::return_type
        ExprTkEngine::EcmdFunctionInstance::makeReturnType(EcmdParser::ValueType rt)
    {
        return (rt == EcmdParser::ValueType::String)
            ? igenfunct_t::e_rtrn_string
            : igenfunct_t::e_rtrn_scalar;
    }

    ExprTkEngine::EcmdFunctionInstance::EcmdFunctionInstance(
        const EcmdParser::FunctionDef& def)
        : igenfunct_t(makeParamSequence(def.params), makeReturnType(def.returnType))
        , _name(def.name)
        , _body(def.body)
    {
        // Build the routing table and the per-kind slot vectors. Each
        // parameter gets one slot (unique_ptr for stable addresses) bound
        // into _innerSymbols under the parameter's name, so the user's
        // body can read it directly.
        _argRoutes.reserve(def.params.size());

        for (const auto& p : def.params) {
            if (p.type == EcmdParser::ValueType::String) {
                _stringSlots.push_back(std::make_unique<std::string>());
                _innerSymbols.add_stringvar(p.name, *_stringSlots.back());
                _argRoutes.push_back({ true, _stringSlots.size() - 1 });
            }
            else {
                _scalarSlots.push_back(std::make_unique<double>(0.0));
                _innerSymbols.add_variable(p.name, *_scalarSlots.back());
                _argRoutes.push_back({ false, _scalarSlots.size() - 1 });
            }
        }
        _innerSymbols.add_constants();

        // ExprTk's igeneric_function rejects zero-argument calls by
        // default - explicitly enable them so user functions declared
        // as 'function name() ... end' can be called as 'name()'.
        if (def.params.empty()) {
            this->allow_zero_parameters() = true;
        }
    }

    void ExprTkEngine::EcmdFunctionInstance::prepareSymbolTables(
        symbol_table_t& libTable)
    {
        _innerExpr.register_symbol_table(_innerSymbols);
        // Sharing the library table here is what enables cross-calls and
        // recursion: when this function's body parses a name like
        // 'num2rom', the parser looks it up in libTable and binds the
        // call to the EcmdFunctionInstance registered there.
        _innerExpr.register_symbol_table(libTable);
    }

    bool ExprTkEngine::EcmdFunctionInstance::compileBody(
        parser_t& parser, std::string& errorOut)
    {
        if (parser.compile(_body, _innerExpr)) {
            return true;
        }
        errorOut = "In function '" + _name + "': " + parser.error();
        return false;
    }

    // ---- Argument marshalling (shared between scalar and string paths) --

    void ExprTkEngine::EcmdFunctionInstance::marshalArgs(
        parameter_list_t parameters, ArgSnapshot& out)
    {
        out.scalars.reserve(_scalarSlots.size());
        out.strings.reserve(_stringSlots.size());

        // ExprTk guarantees that for a fixed-signature generic function
        // parameters.size() matches the declared parameter-sequence
        // length, so iterating up to that count is safe.
        for (std::size_t i = 0; i < parameters.size(); ++i) {
            const auto& route = _argRoutes[i];
            if (route.isString) {
                out.strings.push_back(*_stringSlots[route.slotIndex]);
                const generic_t::string_view sv(parameters[i]);
                _stringSlots[route.slotIndex]->assign(sv.begin(), sv.size());
            }
            else {
                out.scalars.push_back(*_scalarSlots[route.slotIndex]);
                const generic_t::scalar_view sc(parameters[i]);
                *_scalarSlots[route.slotIndex] = sc();
            }
        }
    }

    void ExprTkEngine::EcmdFunctionInstance::restoreArgs(
        std::size_t paramCount, ArgSnapshot& snapshot)
    {
        std::size_t scalarPos = 0;
        std::size_t stringPos = 0;
        for (std::size_t i = 0; i < paramCount; ++i) {
            const auto& route = _argRoutes[i];
            if (route.isString) {
                *_stringSlots[route.slotIndex] = std::move(snapshot.strings[stringPos++]);
            }
            else {
                *_scalarSlots[route.slotIndex] = snapshot.scalars[scalarPos++];
            }
        }
    }

    double ExprTkEngine::EcmdFunctionInstance::operator()(
        parameter_list_t parameters)
    {
        ArgSnapshot snapshot;
        marshalArgs(parameters, snapshot);

        // Scalar-return user functions may end either in a final
        // expression (value() carries the result) or in a 'return [x]'
        // statement (results() carries it and value() is 0); capture
        // both and pick whichever applies.
        const double valueResult = _innerExpr.value();
        double result = std::numeric_limits<double>::quiet_NaN();
        if (_innerExpr.return_invoked()) {
            if (_innerExpr.results().count() > 0) {
                _innerExpr.results().get_scalar(0, result);
            }
        }
        else {
            result = valueResult;
        }

        restoreArgs(parameters.size(), snapshot);
        return result;
    }

    double ExprTkEngine::EcmdFunctionInstance::operator()(
        std::string& result, parameter_list_t parameters)
    {
        result.clear();

        ArgSnapshot snapshot;
        marshalArgs(parameters, snapshot);

        _innerExpr.value();
        if (_innerExpr.return_invoked() && _innerExpr.results().count() > 0) {
            _innerExpr.results().get_string(0, result);
        }

        restoreArgs(parameters.size(), snapshot);
        return 0.0;
    }

    // ---------------------------------------------------------------------
    // EcmdLibrary
    // ---------------------------------------------------------------------

    bool ExprTkEngine::EcmdLibrary::load(const std::string& fileContent,
        const std::string& sourceLabel,
        std::string& errorOut)
    {
        auto parsed = EcmdParser::parse(fileContent);
        if (!parsed.success) {
            errorOut = sourceLabel + " (offset " + std::to_string(parsed.errorPos)
                + "): " + parsed.errorMessage;
            return false;
        }

        // Pass 1: build every instance and register it in _libTable.
        // Bodies stay uncompiled at this point so a later body that calls
        // an earlier-defined function (or itself, or another function
        // defined further down the file) finds the name resolved.
        std::vector<std::unique_ptr<EcmdFunctionInstance>> pending;
        pending.reserve(parsed.functions.size());

        for (auto& def : parsed.functions) {
            auto inst = std::make_unique<EcmdFunctionInstance>(def);
            if (!_libTable.add_function(inst->name(), *inst)) {
                errorOut = sourceLabel + ": function '" + inst->name()
                    + "' clashes with an existing name";
                // Roll back: remove anything we already added so the
                // library stays in a consistent state. add_function
                // returned true only for entries we own; iterate them.
                for (auto& priorInst : pending) {
                    _libTable.remove_function(priorInst->name());
                }
                return false;
            }
            pending.push_back(std::move(inst));
        }

        // Pass 2: compile bodies. Cross-calls and recursion resolve now
        // because all names from this load (plus any previously loaded)
        // are visible in _libTable.
        for (auto& inst : pending) {
            inst->prepareSymbolTables(_libTable);
            std::string err;
            if (!inst->compileBody(_parser, err)) {
                errorOut = sourceLabel + ": " + err;
                for (auto& priorInst : pending) {
                    _libTable.remove_function(priorInst->name());
                }
                return false;
            }
        }

        // Commit on success: transfer ownership of compiled instances.
        for (auto& inst : pending) {
            _instances.push_back(std::move(inst));
        }
        return true;
    }

    // ---------------------------------------------------------------------
    // EcmdLoaderFunction
    // ---------------------------------------------------------------------

    double ExprTkEngine::EcmdLoaderFunction::operator()(
        parameter_list_t parameters)
    {
        if (!_owner || parameters.size() != 1) {
            return 0.0;
        }
        if (parameters[0].type != generic_t::e_string) {
            return 0.0;
        }
        const string_t sv(parameters[0]);
        _owner->loadEcmdFile(std::string(sv.begin(), sv.size()));
        return 0.0;
    }

    // ---------------------------------------------------------------------
    // loadEcmdFile
    // ---------------------------------------------------------------------

    void ExprTkEngine::loadEcmdFile(const std::string& utf8Path)
    {
        if (!_ecmdLibrary) {
            // Defensive: should always exist between initialize() and
            // shutdown(); ignore loads otherwise.
            return;
        }

        // Windows-correct path handling: convert UTF-8 to wide so
        // non-ASCII directory names work the same as Lua's
        // safeLoadFileSandbox does.
        std::wstring wpath = Encoding::utf8ToWString(utf8Path);
        std::ifstream in(std::filesystem::path(wpath), std::ios::binary);
        if (!in) {
            reportError(ILuaEngineHost::ErrorCategory::ExecutionError,
                "ecmd: cannot open file '" + utf8Path + "'");
            return;
        }

        std::ostringstream buf;
        buf << in.rdbuf();
        std::string content = buf.str();

        // Strip a leading UTF-8 BOM if present, so the parser sees clean
        // ASCII at offset 0.
        if (content.size() >= 3
            && static_cast<unsigned char>(content[0]) == 0xEF
            && static_cast<unsigned char>(content[1]) == 0xBB
            && static_cast<unsigned char>(content[2]) == 0xBF)
        {
            content.erase(0, 3);
        }

        std::string err;
        if (!_ecmdLibrary->load(content, utf8Path, err)) {
            reportError(ILuaEngineHost::ErrorCategory::CompileError, err);
        }
    }

    // ---------------------------------------------------------------------
    // seq() function implementation
    // ---------------------------------------------------------------------

    double ExprTkEngine::SeqFunction::operator()(const std::vector<double>& args)
    {
        // Both arguments are optional. seq() yields 1, 2, 3, ... seq(s)
        // starts from s with step 1, seq(s, i) is fully parametrised.
        // We read CNT from the engine directly; it has already been
        // updated for the current match by execute() before any
        // expression is evaluated, so seq() returns start on the first
        // match (CNT=1), start+inc on the second (CNT=2), etc.
        if (!_owner) return 0.0;

        const double start = !args.empty() ? args[0] : 1.0;
        const double inc = args.size() > 1 ? args[1] : 1.0;
        return start + (_owner->_varCNT - 1.0) * inc;
    }

    // ---------------------------------------------------------------------
    // numcol() / txtcol() implementations
    // ---------------------------------------------------------------------

    namespace {

        // Resolves an ExprTk T|S parameter into the host's raw cell text.
        // Returns false on any failure (non-finite scalar, index < 1,
        // unknown header name, no CSV mode, ...). On false, cell is
        // guaranteed empty.
        bool resolveCsvCell(ILuaEngineHost* host,
            exprtk::igeneric_function<double>::parameter_list_t parameters,
            std::string& cell)
        {
            using gen_t = exprtk::igeneric_function<double>::generic_type;
            using scalar_t = gen_t::scalar_view;
            using string_t = gen_t::string_view;

            cell.clear();
            if (!host || parameters.size() != 1) return false;

            const auto& p = parameters[0];
            if (p.type == gen_t::e_scalar) {
                const scalar_t s(p);
                const double v = s();
                if (!std::isfinite(v)) return false;
                const long long idx = static_cast<long long>(v);
                if (idx < 1) return false;
                return host->readCurrentRowColumnByIndex(static_cast<int>(idx), cell);
            }
            if (p.type == gen_t::e_string) {
                const string_t sv(p);
                return host->readCurrentRowColumnByName(exprtk::to_str(sv), cell);
            }
            return false;
        }

    } // anonymous namespace

    double ExprTkEngine::NumColFunction::operator()(
        const std::size_t& /*psi*/,
        parameter_list_t parameters)
    {
        std::string cell;
        if (!_owner || !resolveCsvCell(_owner->_host, parameters, cell)) {
            return std::numeric_limits<double>::quiet_NaN();
        }
        return parseCaptureToDouble(cell);
    }

    double ExprTkEngine::TxtColFunction::operator()(
        const std::size_t& /*psi*/,
        std::string& result,
        parameter_list_t parameters)
    {
        // resolveCsvCell clears 'result' before doing anything else,
        // so a failed call still leaves the empty-string default.
        if (!_owner) { result.clear(); return 0.0; }
        resolveCsvCell(_owner->_host, parameters, result);
        return 0.0;
    }

} // namespace MultiReplaceEngine