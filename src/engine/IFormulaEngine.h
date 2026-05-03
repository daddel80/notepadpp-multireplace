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
        // Per-run hooks. beginRun() is called before the first match of a
        // Replace-All pass, the others after the last. Default no-ops.
        virtual void beginRun() {}
        virtual std::wstring endRunSummary() { return {}; }
        virtual std::wstring endRunSkipAllNoticeText() { return {}; }

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
    };

    // Convenience alias used throughout the rest of the codebase to keep
    // declarations short.
    using FormulaEnginePtr = std::unique_ptr<IFormulaEngine>;

} // namespace MultiReplaceEngine