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
#include "../exprtk/MatchHistoryAnalysis.h"
#include "../exprtk/NumberParse.h"
#include "../Encoding.h"
namespace SU = StringUtils;

namespace MultiReplaceEngine {

    namespace {

        // Shared helpers for the generic-function arity dispatch in the
        // history readers and the upgraded num/txt overloads. The
        // parameter slot type varies across overloads, so the readScalar
        // template lets the compiler pick whatever exact reference type
        // ExprTk hands us; constructing a scalar_view from it is the
        // documented way to read a Scalar parameter.
        template <typename P>
        inline double readScalar(const P& p)
        {
            using gen_t = typename exprtk::igeneric_function<double>::generic_type;
            using scalar_t = typename gen_t::scalar_view;
            return scalar_t(p)();
        }

        // Turn a possibly-fractional, possibly-negative double into a
        // non-negative integral index. Returns false (and leaves out
        // untouched) for NaN, infinity, or a negative value. Trims toward
        // zero on the positive side so num(1.7, ...) means num(1, ...).
        inline bool toIndex(double v, long long& out)
        {
            if (!std::isfinite(v)) return false;
            const long long ll = static_cast<long long>(v);
            if (ll < 0) return false;
            out = ll;
            return true;
        }

        // Same shape but signed: used for the p lookback argument where
        // negative values fall back to v at runtime rather than aborting.
        inline bool toSigned(double v, long long& out)
        {
            if (!std::isfinite(v)) return false;
            out = static_cast<long long>(v);
            return true;
        }

    } // anonymous namespace

    // ---------------------------------------------------------------------
    // Construction / destruction
    // ---------------------------------------------------------------------

    ExprTkEngine::ExprTkEngine(ILuaEngineHost* host)
        : _host(host)
        , _numFunction(this)
        , _skipFunction(this)
        , _txtFunction(this)
        , _seqFunction(this)
        , _numOutFunction(this)
        , _txtOutFunction(this)
        , _numPrevFunction(this)
        , _txtPrevFunction(this)
        , _numColFunction(this)
        , _txtColFunction(this)
        , _parsedateFunction(this)
        , _num2hexFunction(this, 16)
        , _num2binFunction(this, 2)
        , _num2octFunction(this, 8)
        , _hex2numFunction(this, 16)
        , _bin2numFunction(this, 2)
        , _oct2numFunction(this, 8)
        , _num2romFunction(this)
        , _rom2numFunction(this)
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

        // Match history readers. numout(N[, P[, V]]) and txtout(N[, P[, V]])
        // read block N's emission P matches ago; numprev / txtprev are the
        // same-block shorthand. Buffer is sized at compile time from the
        // analyzer's findings - zero-cost when the template uses none of
        // these.
        _symbolTable.add_function("numout", _numOutFunction);
        _symbolTable.add_function("txtout", _txtOutFunction);
        _symbolTable.add_function("numprev", _numPrevFunction);
        _symbolTable.add_function("txtprev", _txtPrevFunction);

        // CSV column accessors. Both delegate to the host, which knows
        // about the active CSV settings, the line of the current match,
        // and caches the row's column array for repeated lookups.
        _symbolTable.add_function("numcol", _numColFunction);
        _symbolTable.add_function("txtcol", _txtColFunction);

        // Register parsedate(str, fmt) - the inverse of D[fmt] output.
        // Returns Unix timestamp on success, NaN on parse failure.
        _symbolTable.add_function("parsedate", _parsedateFunction);

        // Base conversions: numeric <-> hex/bin/oct as built-ins. num2X
        // returns a bare lowercase string; X2num accepts the input case-
        // insensitively, with or without the matching prefix, NaN on
        // invalid characters.
        _symbolTable.add_function("num2hex", _num2hexFunction);
        _symbolTable.add_function("num2bin", _num2binFunction);
        _symbolTable.add_function("num2oct", _num2octFunction);
        _symbolTable.add_function("hex2num", _hex2numFunction);
        _symbolTable.add_function("bin2num", _bin2numFunction);
        _symbolTable.add_function("oct2num", _oct2numFunction);

        // Roman numerals: standard 1..3999 range. num2rom emits canonical
        // subtractive form; rom2num is lenient about non-canonical input.
        _symbolTable.add_function("num2rom", _num2romFunction);
        _symbolTable.add_function("rom2num", _rom2numFunction);

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

        // Discard match history from any previous run. Each Replace-All
        // starts with an empty ring; the first match in the run sees
        // _history.size() == 0, so numprev() / numout() bootstrap with
        // their v fallback (0 for numprev arity-0, NaN otherwise).
        _history.clear();
        _currentBlockIndex = 0;
        // _currentBlockOutputs and _captureSlotsForHistory keep their
        // capacity; their content is overwritten on the next match.
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
        // state rather than half-populated leftovers. This includes the
        // history-related members: if compile() succeeds the Step-2.5
        // block reassigns them with the new dimensions, and if compile()
        // fails partway through we'd otherwise leak a stale ring buffer
        // from the previous successful compile into the next attempt.
        // (Functionally tolerated because `_haveCompiled = false` keeps
        // execute() from running, but cleaner state simplifies reasoning
        // and aligns with the other members reset here.)
        _compiledExpressions.clear();
        _segmentSpecs.clear();
        _parsedTemplate = ExprTkPatternParser::ParseResult();
        _lastCompiledScript.clear();
        _haveCompiled = false;
        _history = MatchHistory{};
        _historyCaptureCap = 0;
        _currentBlockOutputs.clear();
        _captureSlotsForHistory.clear();

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

            // Detect string-producing root nodes (e.g. (?=num2rom(num(1))),
            // (?='abc'), (?=fpath + '.bak')). These are read via
            // expression_helper::get_string() at eval time, bypassing the
            // 'return [...]' detour. Cached at compile to avoid per-eval
            // type checks. The existing 'return [...]' path is unchanged
            // and remains available.
            _segmentSpecs[i].isString =
                exprtk::expression_helper<double>::is_string(expr);

            // Format specs are numeric formatters (e.g. %05.2f). Combining
            // them with a string-producing formula has no defined meaning.
            // Catch it at compile time, same diagnostic style as the
            // return-statement counterpart in the eval path.
            if (_segmentSpecs[i].isString && _segmentSpecs[i].hasSpec) {
                std::string msg = "Format spec cannot be combined with a string-returning formula: \"";
                msg += seg.text;
                msg += "\"";
                reportError(ILuaEngineHost::ErrorCategory::CompileError, msg);

                _compiledExpressions.clear();
                _segmentSpecs.clear();
                return false;
            }

            _compiledExpressions[i] = std::move(expr);
        }

        // Step 2.5: Match-history analysis.
        // Walk the parsed template, find every num/txt arity-2/3 and
        // every numout/txtout/numprev/txtprev call, and read out:
        //   - the maximum lookback depth (sizes the ring buffer)
        //   - the maximum capture index (sizes each entry's capture vec)
        //   - any literal-argument errors (negative p, hard-cap overflow,
        //     negative n in num/txt arity-2/3)
        //
        // Errors abort compile via reportError() with the analyzer's
        // diagnostic message verbatim; success only resizes the ring,
        // it does not allocate per-entry storage (that happens lazily
        // on the first push from execute()).
        {
            std::string historyErr;
            HistoryAnalysis ha = analyzeHistory(parseRes, historyErr);
            if (!historyErr.empty()) {
                reportError(ILuaEngineHost::ErrorCategory::CompileError,
                    historyErr);
                _compiledExpressions.clear();
                _segmentSpecs.clear();
                return false;
            }

            // Size the ring buffer. depth=0 path keeps pushSwap as a
            // no-op so templates without history pay no runtime cost.
            //
            // Block indexing convention: block N is the N-th Expression
            // segment in the template (0-based, in source order); Literal
            // segments don't count. This matches user-visible semantics:
            // `numout(0, 1)` reads the FIRST (?=...) block's previous
            // output, regardless of how much literal text precedes or
            // follows it. ha.blockCount comes from analyzeHistory which
            // counts only Expression segments.
            const std::size_t blockCount = ha.blockCount;
            const std::size_t captureSlots = ha.maxCaptureIndex + 1;
            _history = MatchHistory(ha.maxLookback, captureSlots, blockCount);
            _historyCaptureCap = captureSlots;

            // Size the per-match block-output vector to match the
            // expression count - one slot per (?=...) in the template.
            // The execute() loop maps from segment-index to expression-
            // index via a running counter. (Empty staging vectors at
            // this point - they were cleared at the top of compile().)
            _currentBlockOutputs.assign(blockCount, BlockOutput{});
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

        // Reset per-match history state. _currentBlockIndex tracks which
        // expression-block is evaluating so numprev / txtprev can resolve
        // their implicit n. _currentBlockOutputs is already sized to the
        // expression count from compile(); each slot is overwritten as
        // the matching block evaluates. Skip-cases leave the slot with
        // stale data from the previous match - that's fine because
        // skipped matches don't push to history.
        _currentBlockIndex = 0;

        // Maps segment-index to expression-index: only Expression segments
        // increment this. The user-visible "block index" addresses this
        // counter, not the raw segment position - so a template like
        //   foo (?=A) bar (?=B)
        // gives A=block 0 and B=block 1, with the "foo " and " bar "
        // literals contributing nothing.
        std::size_t expressionIdx = 0;

        const auto& segs = _parsedTemplate.segments;
        for (std::size_t i = 0; i < segs.size(); ++i) {
            const auto& seg = segs[i];

            if (seg.type == ExprTkPatternParser::SegmentType::Literal) {
                out.append(seg.text);
                continue;
            }

            // Set the running block index BEFORE eval so numprev/txtprev
            // inside this expression know their implicit n. The
            // within-match guard in lookupBlockOutputAt uses this value
            // to block reads of the running block or any later one.
            // expressionIdx counts only Expression segments - matching
            // the user-visible block-index convention.
            _currentBlockIndex = expressionIdx;

            // Eval. ExprTk's return statement turns into a side-channel
            // result list - we must check return_invoked() to decide
            // which output path applies, because value() is 0.0 in that
            // case and would otherwise look like a normal numeric eval.
            auto& expr = _compiledExpressions[i];

            // Direct-string path: when the root node is string-producing
            // (e.g. (?=num2rom(num(1))), (?='abc'), (?=fpath + '.bak')),
            // read the string straight from the root node via the
            // expression_helper::get_string() extension. This avoids the
            // user having to wrap every string result in 'return [...]',
            // while leaving the explicit return-list path untouched.
            //
            // The isString flag was cached at compile time; here we only
            // pay the cost of get_string() which internally triggers
            // value() exactly once. NaN/Inf cannot occur in this branch
            // because a string-producing root node does not carry a
            // numeric result through value() (it returns T but the
            // semantic payload is the string).
            if (_segmentSpecs[i].isString) {
                std::string strOut;
                if (!exprtk::expression_helper<double>::get_string(expr, strOut)) {
                    // is_string() was true at compile but get_string()
                    // failed at eval. Should not happen for any node type
                    // we know of, but report as a soft skip so a
                    // pathological case never crashes the run.
                    onInvalid(result, seg.text);
                    return result;
                }
                // Record for history. txtout/txtprev readers will see
                // this slot in subsequent matches; numout/numprev get a
                // type-mismatch fallback because the slot is String.
                if (expressionIdx < _currentBlockOutputs.size()) {
                    _currentBlockOutputs[expressionIdx].setString(strOut);
                }
                if (escapeOutput && _host) {
                    out.append(_host->escapeForRegex(strOut));
                }
                else {
                    out.append(strOut);
                }
                ++expressionIdx;
                continue;
            }

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
                // Capture the assembled return-list string into the
                // block-output slot for history. appendExprtkResults
                // writes to `out` directly, so we snapshot the portion
                // it just wrote by tracking the size delta. This is a
                // narrow path (only fires for the legacy 'return [...]'
                // style) but the steady-state cost is one substring.
                const std::size_t outSizeBefore = out.size();
                appendExprtkResults(expr, out, escapeOutput);
                if (_outputHadInvalid) {
                    onInvalid(result, seg.text);
                    return result;
                }
                if (expressionIdx < _currentBlockOutputs.size()) {
                    // Use the unescaped pre-append snapshot if available;
                    // for simplicity record the appended bytes, which is
                    // what the user-visible output saw. History readers
                    // see the same string the document received.
                    _currentBlockOutputs[expressionIdx].setString(
                        std::string_view(out.data() + outSizeBefore,
                            out.size() - outSizeBefore));
                }
                ++expressionIdx;
                continue;
            }

            // NaN or Inf both indicate an unusable numeric result. Route
            // either through the recoverable-error dialog instead of
            // letting "nan" / "inf" leak into the output text.
            if (std::isnan(value) || std::isinf(value)) {
                onInvalid(result, seg.text);
                return result;
            }

            // Record the numeric output for history before formatting.
            // The formatted string isn't what numout/numprev want - they
            // want the raw double so subsequent numeric reads keep full
            // precision.
            if (expressionIdx < _currentBlockOutputs.size()) {
                _currentBlockOutputs[expressionIdx].setNumber(value);
            }
            ++expressionIdx;

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

        // Match-history push: only successful, non-skipped matches make
        // it into the ring. Skipped matches are invisible to subsequent
        // numprev/numout lookups, which is exactly what users with skip-
        // based filtering want.
        //
        // The push is a no-op when the ring depth is 0 (no template uses
        // history) - so templates that don't touch history pay nothing.
        //
        // The capture-slot snapshot is built lazily here from the live
        // _strMATCH and _captureStrings. The full match goes into slot 0,
        // capture groups into slots 1..N. pushSwap then swaps the entire
        // vector into the ring so its string buffers rotate with the
        // entry that was just evicted.
        if (!_wantSkip && _history.depth() > 0) {
            // Build the snapshot. The vector is held as a member so its
            // capacity is preserved across matches. Slot 0 = full match;
            // 1..N = capture groups.
            const std::size_t cap = std::max<std::size_t>(
                _historyCaptureCap,
                _captureStrings.size() + 1);

            if (_captureSlotsForHistory.size() < cap) {
                _captureSlotsForHistory.resize(cap);
            }
            _captureSlotsForHistory[0].assign(_strMATCH);
            for (std::size_t k = 0; k < _captureStrings.size(); ++k) {
                if (k + 1 < _captureSlotsForHistory.size()) {
                    _captureSlotsForHistory[k + 1].assign(_captureStrings[k]);
                }
            }
            // Slots beyond _captureStrings.size()+1 must be cleared so a
            // ring entry from a previous, wider match doesn't leak its
            // strings via the swap. (Steady-state cost: assigning an
            // empty string reuses the existing capacity.)
            for (std::size_t k = _captureStrings.size() + 1;
                k < _captureSlotsForHistory.size(); ++k) {
                _captureSlotsForHistory[k].assign(std::string_view{});
            }

            // If the runtime regex produced more captures than the
            // compile-time literal scan predicted (only possible when
            // the template used a non-literal capture index), grow
            // every ring entry to match before the push.
            if (cap > _historyCaptureCap) {
                _history.resizeCaptureSlots(cap);
                _historyCaptureCap = cap;
            }

            _history.pushSwap(_captureSlotsForHistory, _currentBlockOutputs);
        }

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

    double ExprTkEngine::NumFunction::operator()(const std::size_t& /*psi*/,
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        const std::size_t arity = parameters.size();
        if (arity < 1 || arity > 3 || !_owner) {
            return 0.0;
        }

        // First arg = capture index. Negative or non-finite -> legacy
        // 0.0 for arity-1 (backwards compatibility with the original
        // num() semantics), v-fallback for arity-2/3.
        const double indexD = readScalar(parameters[0]);
        long long idx = 0;
        const bool indexOk = toIndex(indexD, idx);

        // -----------------------------------------------------------------
        // Arity 1: legacy num(N) - read capture from the current match.
        // Out-of-range indices return 0.0, matching the pre-history
        // behaviour every existing template depends on.
        // -----------------------------------------------------------------
        if (arity == 1) {
            if (!indexOk) return 0.0;
            const auto& caps = _owner->_captures;
            const std::size_t u = static_cast<std::size_t>(idx);
            if (u >= caps.size()) return 0.0;
            return caps[u];
        }

        // -----------------------------------------------------------------
        // Arity 2/3: num(N, P) / num(N, P, V) - read capture N from the
        // match P matches ago. Fallback is NaN (arity-2) or V (arity-3),
        // applied when:
        //   - the index isn't a valid non-negative integer,
        //   - the lookback fails (past history, dynamic n out of range),
        //   - or, for arity-3 only, the looked-up value is not finite.
        // -----------------------------------------------------------------
        double fallback = nanResult;
        if (arity == 3) {
            fallback = readScalar(parameters[2]);
        }

        if (!indexOk) return fallback;

        const double pD = readScalar(parameters[1]);
        long long pLookback = 0;
        if (!toSigned(pD, pLookback) || pLookback < 0) {
            return fallback;
        }

        double value = 0.0;
        if (!_owner->lookupCaptureAt(idx, pLookback, value)) {
            return fallback;
        }

        if (arity == 3 && !std::isfinite(value)) {
            return fallback;
        }
        return value;
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
        const std::size_t& /*psi*/,
        std::string& result,
        parameter_list_t parameters)
    {
        // Default to empty: an out-of-range or malformed call yields ""
        // rather than aborting eval, mirroring num(N)'s defensive style.
        result.clear();

        const std::size_t arity = parameters.size();
        if (!_owner || arity < 1 || arity > 3) {
            return 0.0;
        }

        // First arg = capture index.
        const double indexD = readScalar(parameters[0]);
        long long idx = 0;
        const bool indexOk = toIndex(indexD, idx);

        // -----------------------------------------------------------------
        // Arity 1: legacy txt(N) - capture from the current match. The
        // value is left empty on any failure (mirrors the original
        // pre-history defensive behaviour).
        // -----------------------------------------------------------------
        if (arity == 1) {
            if (!indexOk) return 0.0;
            const std::size_t u = static_cast<std::size_t>(idx);
            if (u == 0) {
                result = _owner->_strMATCH;
                return 0.0;
            }
            if (u - 1 < _owner->_captureStrings.size()) {
                result = _owner->_captureStrings[u - 1];
            }
            return 0.0;
        }

        // -----------------------------------------------------------------
        // Arity 2/3: txt(N, P) / txt(N, P, V) - read capture N text from
        // the match P matches ago. Fallback is "" (arity-2) or V (arity-3).
        // -----------------------------------------------------------------
        std::string fallback;
        if (arity == 3) {
            // V is a string in arity-3 (declared "T|TT|TTS"). The third
            // slot is a string_view, not a scalar.
            const string_view_t sv(parameters[2]);
            fallback.assign(sv.begin(), sv.size());
        }

        if (!indexOk) { result = fallback; return 0.0; }

        const double pD = readScalar(parameters[1]);
        long long pLookback = 0;
        if (!toSigned(pD, pLookback) || pLookback < 0) {
            result = fallback;
            return 0.0;
        }

        if (!_owner->lookupCaptureAt(idx, pLookback, result)) {
            result = fallback;
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
    // Base conversions: num <-> hex/bin/oct
    // ---------------------------------------------------------------------

    namespace {

        // Returns true and writes the value if 'c' is a valid digit for
        // the given base (2/8/16). Hex accepts upper- and lowercase.
        bool digitForBase(char c, int base, int& outValue)
        {
            if (c >= '0' && c <= '9') {
                outValue = c - '0';
            }
            else if (base == 16 && c >= 'a' && c <= 'f') {
                outValue = 10 + (c - 'a');
            }
            else if (base == 16 && c >= 'A' && c <= 'F') {
                outValue = 10 + (c - 'A');
            }
            else {
                return false;
            }
            return outValue < base;
        }

        // Strips ASCII whitespace from both ends. Used so hex2num(" ff ")
        // still parses cleanly.
        std::string trimAscii(const std::string& s)
        {
            const auto isSpace = [](unsigned char c) {
                return c == ' ' || c == '\t' || c == '\r' || c == '\n';
                };
            std::size_t first = 0;
            while (first < s.size() && isSpace(static_cast<unsigned char>(s[first]))) ++first;
            std::size_t last = s.size();
            while (last > first && isSpace(static_cast<unsigned char>(s[last - 1]))) --last;
            return s.substr(first, last - first);
        }

        // Removes a leading base-prefix if present (0x/0X for hex,
        // 0b/0B for binary, 0o/0O for octal). Returns the offset to
        // continue parsing from.
        std::size_t skipBasePrefix(const std::string& s, int base)
        {
            if (s.size() < 2 || s[0] != '0') return 0;
            const char p = s[1];
            if (base == 16 && (p == 'x' || p == 'X')) return 2;
            if (base == 2 && (p == 'b' || p == 'B')) return 2;
            if (base == 8 && (p == 'o' || p == 'O')) return 2;
            return 0;
        }

        // Renders a non-negative integer in the given base as lowercase
        // ASCII. Always emits at least "0" so num2hex(0) -> "0".
        std::string formatUnsignedBase(unsigned long long v, int base)
        {
            if (v == 0) return std::string("0");

            static const char kDigits[] = "0123456789abcdef";
            char buf[80];
            std::size_t pos = sizeof(buf);
            while (v > 0) {
                buf[--pos] = kDigits[v % static_cast<unsigned>(base)];
                v /= static_cast<unsigned>(base);
            }
            return std::string(buf + pos, sizeof(buf) - pos);
        }

    } // anonymous namespace

    double ExprTkEngine::Num2BaseFunction::operator()(
        std::string& result,
        parameter_list_t parameters)
    {
        result.clear();

        if (parameters.size() != 1) {
            return 0.0;
        }
        const scalar_t s(parameters[0]);
        const double v = s();

        // NaN or Inf -> empty string. Same recoverable-error contract
        // as formatDouble(): never let "nan"/"inf" leak into output.
        if (!std::isfinite(v)) {
            return 0.0;
        }

        // Truncate toward zero so num2hex(15.7) == num2hex(15) and
        // num2hex(-15.7) == num2hex(-15). Matches (int) cast semantics.
        const long long signedVal = static_cast<long long>(v);
        const bool negative = signedVal < 0;
        const unsigned long long magnitude = negative
            ? static_cast<unsigned long long>(-(signedVal + 1)) + 1ULL  // safe abs for INT64_MIN
            : static_cast<unsigned long long>(signedVal);

        std::string body = formatUnsignedBase(magnitude, _base);
        if (negative) {
            result.reserve(body.size() + 1);
            result.push_back('-');
            result.append(body);
        }
        else {
            result = std::move(body);
        }
        return 0.0;
    }

    double ExprTkEngine::Base2NumFunction::operator()(
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        if (parameters.size() != 1) {
            return nanResult;
        }
        const string_t sv(parameters[0]);
        const std::string raw = trimAscii(exprtk::to_str(sv));
        if (raw.empty()) {
            return nanResult;
        }

        std::size_t i = 0;
        bool negative = false;
        if (raw[i] == '+' || raw[i] == '-') {
            negative = (raw[i] == '-');
            ++i;
        }
        i += skipBasePrefix(raw.substr(i), _base);
        if (i >= raw.size()) {
            return nanResult;   // sign or prefix with no digits
        }

        // Accumulate as unsigned to avoid intermediate-signed overflow.
        // Cap at 2^53 (double's safe-integer ceiling) so the round-trip
        // through double doesn't silently lose precision; over the cap
        // we return NaN rather than a corrupted value.
        unsigned long long acc = 0;
        const unsigned long long kSafeMax = (1ULL << 53);

        for (; i < raw.size(); ++i) {
            int digit = 0;
            if (!digitForBase(raw[i], _base, digit)) {
                return nanResult;
            }
            if (acc > kSafeMax / static_cast<unsigned>(_base)) {
                return nanResult;
            }
            acc = acc * static_cast<unsigned>(_base) + static_cast<unsigned>(digit);
            if (acc > kSafeMax) {
                return nanResult;
            }
        }

        const double mag = static_cast<double>(acc);
        return negative ? -mag : mag;
    }

    // ---------------------------------------------------------------------
    // Roman numerals: num2rom / rom2num
    // ---------------------------------------------------------------------

    namespace {

        // Greedy emission from largest to smallest. Subtractive pairs
        // (CM, CD, XC, XL, IX, IV) sit in the table so a single loop
        // produces the canonical form without special-case branches.
        std::string romanFromInt(int n)
        {
            static const struct { int value; const char* glyph; } kTable[] = {
                {1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"},
                { 100, "C"}, { 90, "XC"}, { 50, "L"}, { 40, "XL"},
                {  10, "X"}, {  9, "IX"}, {  5, "V"}, {  4, "IV"},
                {   1, "I"}
            };
            std::string out;
            for (const auto& row : kTable) {
                while (n >= row.value) {
                    out += row.glyph;
                    n -= row.value;
                }
            }
            return out;
        }

        // Returns the integer value for a single Roman glyph, or -1 if
        // the character is not a Roman numeral. Case-insensitive.
        int romanGlyphValue(char c)
        {
            switch (c) {
            case 'I': case 'i': return 1;
            case 'V': case 'v': return 5;
            case 'X': case 'x': return 10;
            case 'L': case 'l': return 50;
            case 'C': case 'c': return 100;
            case 'D': case 'd': return 500;
            case 'M': case 'm': return 1000;
            default: return -1;
            }
        }

    } // anonymous namespace

    double ExprTkEngine::Num2RomFunction::operator()(
        std::string& result,
        parameter_list_t parameters)
    {
        result.clear();

        if (parameters.size() != 1) return 0.0;
        const scalar_t s(parameters[0]);
        const double v = s();
        if (!std::isfinite(v)) return 0.0;

        // Range check first: classical Roman covers 1..3999. Out of
        // range -> empty string (same recoverable-error pattern as
        // num2hex(NaN)).
        const long long iv = static_cast<long long>(v);
        if (iv < 1 || iv > 3999) return 0.0;

        result = romanFromInt(static_cast<int>(iv));
        return 0.0;
    }

    double ExprTkEngine::Rom2NumFunction::operator()(
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        if (parameters.size() != 1) return nanResult;
        const string_t sv(parameters[0]);
        const std::string raw = trimAscii(exprtk::to_str(sv));
        if (raw.empty()) return nanResult;

        // Left-to-right scan: when a glyph's value is smaller than the
        // next one, it counts as negative (IV = -1 + 5 = 4). Lenient
        // toward non-canonical forms - "IIII" sums to 4, same as "IV".
        long long total = 0;
        const std::size_t n = raw.size();
        for (std::size_t i = 0; i < n; ++i) {
            const int curr = romanGlyphValue(raw[i]);
            if (curr < 0) return nanResult;
            const int next = (i + 1 < n) ? romanGlyphValue(raw[i + 1]) : -1;
            if (next > curr) {
                total += (next - curr);
                ++i;  // consume the next glyph too
            }
            else {
                total += curr;
            }
        }
        return static_cast<double>(total);
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
    // Match history readers - numout / txtout / numprev / txtprev
    // ---------------------------------------------------------------------
    //
    // The shape is identical across the four: every overload takes its
    // own subset of (n, p, v). The lookup helpers below are typed
    // (numeric vs. string) so callers don't need to know how the ring
    // buffer stores block outputs internally.
    //
    // Failure routing:
    //   - past history / out-of-range block / type mismatch
    //     -> lookup returns false, caller substitutes fallback
    //   - within current match, reading the running block or any later
    //     one always fails (the slot doesn't carry a meaningful value
    //     until the block has finished evaluating)
    //   - for arity-3 numeric readers, a non-finite looked-up value
    //     also routes to fallback (keeps NaN out of running totals)
    // ---------------------------------------------------------------------

    bool ExprTkEngine::lookupCaptureAt(long long n,
        long long pLookback,
        double& out) const
    {
        if (n < 0) return false;
        if (pLookback < 0) return false;

        if (pLookback == 0) {
            // Current match: read from the live capture vector.
            const std::size_t u = static_cast<std::size_t>(n);
            if (u >= _captures.size()) return false;
            out = _captures[u];
            return true;
        }

        const MatchHistoryEntry* entry = _history.lookback(
            static_cast<std::size_t>(pLookback));
        if (!entry) return false;

        const std::size_t u = static_cast<std::size_t>(n);
        if (u >= entry->captures.size()) return false;

        out = entry->captures[u].asNumber();
        return true;
    }

    bool ExprTkEngine::lookupCaptureAt(long long n,
        long long pLookback,
        std::string& out) const
    {
        if (n < 0) return false;
        if (pLookback < 0) return false;

        if (pLookback == 0) {
            const std::size_t u = static_cast<std::size_t>(n);
            if (u == 0) { out = _strMATCH; return true; }
            if (u - 1 >= _captureStrings.size()) return false;
            out = _captureStrings[u - 1];
            return true;
        }

        const MatchHistoryEntry* entry = _history.lookback(
            static_cast<std::size_t>(pLookback));
        if (!entry) return false;

        const std::size_t u = static_cast<std::size_t>(n);
        if (u >= entry->captures.size()) return false;

        out = entry->captures[u].asString();
        return true;
    }

    bool ExprTkEngine::lookupBlockOutputAt(long long n,
        long long pLookback,
        double& out) const
    {
        if (n < 0) return false;
        if (pLookback < 0) return false;

        if (pLookback == 0) {
            // Within-match: a block can only read EARLIER finished blocks.
            // The currently-running block and any later one return false
            // so the caller routes to fallback.
            const std::size_t u = static_cast<std::size_t>(n);
            if (u >= _currentBlockIndex) return false;
            if (u >= _currentBlockOutputs.size()) return false;
            const BlockOutput& bo = _currentBlockOutputs[u];
            if (bo.type != BlockOutput::Number) return false;
            out = bo.numValue;
            return true;
        }

        const MatchHistoryEntry* entry = _history.lookback(
            static_cast<std::size_t>(pLookback));
        if (!entry) return false;

        const std::size_t u = static_cast<std::size_t>(n);
        if (u >= entry->blockOutputs.size()) return false;

        const BlockOutput& bo = entry->blockOutputs[u];
        if (bo.type != BlockOutput::Number) return false;
        out = bo.numValue;
        return true;
    }

    bool ExprTkEngine::lookupBlockOutputAt(long long n,
        long long pLookback,
        std::string& out) const
    {
        if (n < 0) return false;
        if (pLookback < 0) return false;

        if (pLookback == 0) {
            const std::size_t u = static_cast<std::size_t>(n);
            if (u >= _currentBlockIndex) return false;
            if (u >= _currentBlockOutputs.size()) return false;
            const BlockOutput& bo = _currentBlockOutputs[u];
            if (bo.type != BlockOutput::String) return false;
            out = bo.strValue;
            return true;
        }

        const MatchHistoryEntry* entry = _history.lookback(
            static_cast<std::size_t>(pLookback));
        if (!entry) return false;

        const std::size_t u = static_cast<std::size_t>(n);
        if (u >= entry->blockOutputs.size()) return false;

        const BlockOutput& bo = entry->blockOutputs[u];
        if (bo.type != BlockOutput::String) return false;
        out = bo.strValue;
        return true;
    }

    // ---------------------------------------------------------------------
    // numout(N), numout(N, P), numout(N, P, V)
    // ---------------------------------------------------------------------

    double ExprTkEngine::NumOutFunction::operator()(const std::size_t& /*psi*/,
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        const std::size_t arity = parameters.size();
        if (!_owner || arity < 1 || arity > 3) return 0.0;

        const double indexD = readScalar(parameters[0]);
        long long n = 0;
        const bool indexOk = toIndex(indexD, n);

        // Arity-1 default: P = 1 (previous match).
        long long pLookback = 1;
        if (arity >= 2) {
            const double pD = readScalar(parameters[1]);
            if (!toSigned(pD, pLookback) || pLookback < 0) {
                return arity == 3 ? readScalar(parameters[2]) : nanResult;
            }
        }

        const double fallback = arity == 3
            ? readScalar(parameters[2])
            : nanResult;

        if (!indexOk) return fallback;

        double value = 0.0;
        if (!_owner->lookupBlockOutputAt(n, pLookback, value)) {
            return fallback;
        }
        if (arity == 3 && !std::isfinite(value)) {
            return fallback;
        }
        return value;
    }

    // ---------------------------------------------------------------------
    // txtout(N), txtout(N, P), txtout(N, P, V)
    // ---------------------------------------------------------------------

    double ExprTkEngine::TxtOutFunction::operator()(
        const std::size_t& /*psi*/,
        std::string& result,
        parameter_list_t parameters)
    {
        result.clear();

        const std::size_t arity = parameters.size();
        if (!_owner || arity < 1 || arity > 3) return 0.0;

        const double indexD = readScalar(parameters[0]);
        long long n = 0;
        const bool indexOk = toIndex(indexD, n);

        // Arity-1 default: P = 1 (previous match).
        long long pLookback = 1;
        if (arity >= 2) {
            const double pD = readScalar(parameters[1]);
            if (!toSigned(pD, pLookback) || pLookback < 0) {
                if (arity == 3) {
                    const string_view_t sv(parameters[2]);
                    result.assign(sv.begin(), sv.size());
                }
                return 0.0;
            }
        }

        std::string fallback;
        if (arity == 3) {
            const string_view_t sv(parameters[2]);
            fallback.assign(sv.begin(), sv.size());
        }

        if (!indexOk) { result = fallback; return 0.0; }

        if (!_owner->lookupBlockOutputAt(n, pLookback, result)) {
            result = fallback;
        }
        return 0.0;
    }

    // ---------------------------------------------------------------------
    // numprev(), numprev(P), numprev(P, V)
    //
    // Same as numout(currentBlockIndex, P[, V]) - the block index is
    // implicit. Arity-0 carries the ergonomic v=0 default that makes
    // (?=num(1) + numprev()) bootstrap a running total without manually
    // covering the "no previous" case.
    // ---------------------------------------------------------------------

    double ExprTkEngine::NumPrevFunction::operator()(const std::size_t& /*psi*/,
        parameter_list_t parameters)
    {
        const double nanResult = std::numeric_limits<double>::quiet_NaN();

        const std::size_t arity = parameters.size();
        if (!_owner || arity > 2) return 0.0;

        // Arity 0: numprev() -> previous match, fallback 0.
        if (arity == 0) {
            const long long n = static_cast<long long>(_owner->_currentBlockIndex);
            double value = 0.0;
            if (!_owner->lookupBlockOutputAt(n, /*pLookback=*/1, value)) {
                return 0.0;  // ergonomic v=0 default
            }
            if (!std::isfinite(value)) return 0.0;
            return value;
        }

        // Arity 1: numprev(P) -> match P ago, fallback NaN.
        // Arity 2: numprev(P, V) -> match P ago, fallback V.
        const double pD = readScalar(parameters[0]);
        long long pLookback = 0;
        const bool pOk = toSigned(pD, pLookback) && pLookback >= 0;

        const double fallback = arity == 2
            ? readScalar(parameters[1])
            : nanResult;

        if (!pOk) return fallback;

        const long long n = static_cast<long long>(_owner->_currentBlockIndex);
        double value = 0.0;
        if (!_owner->lookupBlockOutputAt(n, pLookback, value)) {
            return fallback;
        }
        if (arity == 2 && !std::isfinite(value)) {
            return fallback;
        }
        return value;
    }

    // ---------------------------------------------------------------------
    // txtprev(), txtprev(P), txtprev(P, V)
    // ---------------------------------------------------------------------

    double ExprTkEngine::TxtPrevFunction::operator()(
        const std::size_t& /*psi*/,
        std::string& result,
        parameter_list_t parameters)
    {
        result.clear();

        const std::size_t arity = parameters.size();
        if (!_owner || arity > 2) return 0.0;

        if (arity == 0) {
            const long long n = static_cast<long long>(_owner->_currentBlockIndex);
            _owner->lookupBlockOutputAt(n, /*pLookback=*/1, result);
            // Whether the lookup succeeded or not, result is either the
            // string or empty - both are acceptable arity-0 outputs.
            return 0.0;
        }

        const double pD = readScalar(parameters[0]);
        long long pLookback = 0;
        const bool pOk = toSigned(pD, pLookback) && pLookback >= 0;

        std::string fallback;
        if (arity == 2) {
            const string_view_t sv(parameters[1]);
            fallback.assign(sv.begin(), sv.size());
        }

        if (!pOk) { result = fallback; return 0.0; }

        const long long n = static_cast<long long>(_owner->_currentBlockIndex);
        if (!_owner->lookupBlockOutputAt(n, pLookback, result)) {
            result = fallback;
        }
        return 0.0;
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