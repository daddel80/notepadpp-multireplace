// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// IFormulaEngine.h
// Abstract interface for the per-tab formula engine. The replace
// pipeline only ever talks to this interface; concrete engines
// (LuaEngine, ExprTkEngine, ...) live behind it.
//
// Lifecycle:
//   1. Construct via EngineFactory::create(EngineType).
//   2. Call initialize() once before the first compile/execute.
//   3. compile(script) when the user's replace text changes.
//   4. execute(script, vars, ...) per match.
//   5. shutdown() before destruction (RAII handles this in
//      destructors anyway, but explicit is fine for ordering).

#pragma once

#include "EngineTypes.h"
#include "ILuaEngineHost.h"

#include <memory>
#include <string>

namespace MultiReplaceEngine {

    class IFormulaEngine {
    public:
        virtual ~IFormulaEngine() = default;

        // Non-copyable, non-movable: engines own native resources
        // (lua_State, exprtk::expression) that aren't trivially copied.
        IFormulaEngine() = default;
        IFormulaEngine(const IFormulaEngine&) = delete;
        IFormulaEngine& operator=(const IFormulaEngine&) = delete;
        IFormulaEngine(IFormulaEngine&&) = delete;
        IFormulaEngine& operator=(IFormulaEngine&&) = delete;

        // ----- Lifecycle ---------------------------------------------------

        // Bring the engine into a ready state. Returns false on a hard
        // failure (e.g. luaL_newstate returned null). After a false
        // return, the engine must be discarded.
        virtual bool initialize() = 0;

        // Release any held resources. Safe to call multiple times.
        // Always called by the destructor.
        virtual void shutdown() = 0;

        // ----- Per-run lifecycle ------------------------------------------
        //
        // beginRun() runs before the first match, the end-of-run hooks
        // after the last. The defaults here implement a generic skip-
        // counter pattern; subclasses just need to populate the
        // protected counters via noteRecoverableSkip() etc. Engines that
        // don't need any of this leave the counter alone and the hooks
        // return empty automatically.
        virtual void beginRun() {
            _skipAllErrors = false;
            _errorSkipCount = 0;
        }

        virtual std::wstring endRunSummary() {
            if (_errorSkipCount == 0) {
                return std::wstring{};
            }
            return localiseCount(L"status_recoverable_errors_skipped_summary",
                _errorSkipCount);
        }

        virtual std::wstring endRunSkipAllNoticeText() {
            // Only surface the final notice when the user picked "Skip
            // all errors". In the per-match skip path the user
            // confirmed each one already.
            if (!_skipAllErrors || _errorSkipCount == 0) {
                return std::wstring{};
            }
            return localiseCount(L"msgbox_recoverable_errors_skipped_notice",
                _errorSkipCount);
        }

        // ----- Per-script compile -----------------------------------------

        // Prepare a script for repeated execution. Engines are expected to
        // cache internally so calling compile() with the same script twice
        // is cheap. Returns false on a syntax error; the caller can fetch
        // details via the next execute() (which will also fail), or via
        // dedicated diagnostics if added later.
        virtual bool compile(const std::string& scriptUtf8) = 0;

        // ----- Per-match execution ----------------------------------------

        // Evaluate the previously compiled (or to-be-compiled) script with
        // the given match variables. The script string is also passed so
        // engines can lazily compile on the first execute call - this
        // matches the existing Lua bridge's "ensureLuaCodeCompiled +
        // resolveLuaSyntax" pattern.
        //
        // Parameters:
        //   scriptUtf8        The user's replace text (UTF-8)
        //   vars              CAP1..n, MATCH, FPATH, FNAME, counters, ...
        //   isRegexMatch      Whether the surrounding rule uses regex
        //                     (controls escaping of the result)
        //   documentCodepage  Scintilla codepage of the active document
        //                     (-1 means "ask Scintilla yourself")
        //
        // Returns a FormulaResult; engines never throw across this boundary.
        virtual FormulaResult execute(
            const std::string& scriptUtf8,
            const FormulaVars& vars,
            bool isRegexMatch,
            int documentCodepage
        ) = 0;

        // ----- Metadata ---------------------------------------------------

        // Identity of this concrete engine.
        virtual EngineType type() const = 0;

        // Display strings used by the panel UI:
        //   shortName  -> "Lua" / "ExprTk" (used in popup menu, status)
        //   shortLetter-> "L"   / "E"      (used as the inline label)
        // Both are intentionally English-only and not localised, matching
        // the convention that engine names are proper nouns.
        virtual std::wstring shortName() const = 0;
        virtual std::wstring shortLetter() const = 0;

        // URL for the help button to open when this engine is active.
        // May be a local file:// path or a remote https:// URL; the
        // panel doesn't care.
        virtual std::wstring helpUrl() const = 0;
    protected:
        // ----- Shared recoverable-error skip state ------------------------
        //
        // Engines that emit per-match warnings the user can choose to skip
        // (ExprTk on NaN, later Lua on pcall errors) share this state and
        // the helpers below. Three pieces of state cover the lifecycle:
        //
        //   _wantStop       per-match flag, set when the user picks "Stop"
        //                   in the recoverable-error dialog. The engine's
        //                   execute() must check it and bail out with
        //                   FormulaResult::success = false.
        //
        //   _skipAllErrors  run-scoped flag, set when the user picks
        //                   "Skip all errors". Once set, the engine
        //                   suppresses further dialogs for the rest of
        //                   the run; the count keeps incrementing.
        //
        //   _errorSkipCount run-scoped counter, incremented for every
        //                   skipped match (whether or not a dialog was
        //                   shown). Surfaced via endRunSummary() and
        //                   endRunSkipAllNoticeText().
        //
        // beginRun() resets the run-scoped fields. _wantStop is per-match
        // and gets reset by the engine at the top of each execute().
        bool        _wantStop = false;
        bool        _skipAllErrors = false;
        std::size_t _errorSkipCount = 0;

        // Bumps the skip counter and routes the user through the
        // recoverable-error dialog. Engines call this on every recoverable
        // error and then check _wantStop / FormulaResult::skip to decide
        // how to return. exprText is shown in the dialog to identify the
        // failing expression. The detailKey points to a translation
        // template that takes the exprText as $REPLACE_STRING and produces
        // the engine-specific hint shown in the dialog body.
        //
        // Returns the user's choice (also reflected in _wantStop /
        // _skipAllErrors so callers can ignore the return value if they
        // only care about side effects).
        ILuaEngineHost::RecoverableErrorChoice handleRecoverableSkip(
            ILuaEngineHost* host,
            const std::wstring& engineName,
            const std::wstring& detailKey,
            const std::string& exprText);

        // Helper for the default endRun* hooks: pulls a translation key
        // and substitutes the count as $REPLACE_STRING.
        static std::wstring localiseCount(const std::wstring& key,
            std::size_t count);
    };

    // Convenience alias used throughout the rest of the codebase to keep
    // declarations short.
    using FormulaEnginePtr = std::unique_ptr<IFormulaEngine>;

} // namespace MultiReplaceEngine