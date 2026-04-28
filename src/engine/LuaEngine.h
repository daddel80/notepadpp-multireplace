// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// LuaEngine.h
// Concrete IFormulaEngine implementation backed by a Lua 5.x state.
//
// This is the encapsulation of MR's original Lua bridge. Behaviour is
// preserved one-to-one - same variables, same set/skip semantics, same
// per-match optimization caches (FPATH/FNAME/regex flag/cap count) - the
// surface just lives behind IFormulaEngine now so the replace pipeline
// can stay engine-agnostic.

#pragma once

#include "IFormulaEngine.h"
#include "ILuaEngineHost.h"

#include <lua.hpp>

#include <map>
#include <string>
#include <vector>

namespace MultiReplaceEngine {

    // Lua-specific execution state kept inside the engine. None of these
    // members appear in the public interface; they exist only as
    // implementation detail of the per-match optimisations the original
    // bridge already performed.
    class LuaEngine final : public IFormulaEngine {
    public:
        // The host pointer must outlive the engine; LuaEngine does not
        // take ownership.
        explicit LuaEngine(ILuaEngineHost* host);
        ~LuaEngine() override;

        // ----- IFormulaEngine ---------------------------------------------

        bool initialize() override;
        void shutdown() override;

        bool compile(const std::string& scriptUtf8) override;

        FormulaResult execute(
            const std::string& scriptUtf8,
            const FormulaVars& vars,
            bool isRegexMatch,
            int documentCodepage
        ) override;

        EngineType type() const override { return EngineType::Lua; }
        std::wstring shortName() const override { return L"Lua"; }
        std::wstring shortLetter() const override { return L"L"; }
        std::wstring helpUrl() const override;

        // ----- Lua-specific helpers (internal use & sandbox callbacks) ----
        //
        // These are public only because Lua's C-API requires C-style free
        // functions for hooks. External callers should not depend on them.
        static int  safeLoadFileSandbox(lua_State* L);
        static void applyLuaSafeMode(lua_State* L);

    private:
        // ----- Internal helpers -------------------------------------------

        // Push a string variable to the Lua global table. Centralised so
        // UTF-8 handling and registry conventions stay in one place.
        void setLuaVariable(lua_State* L, const std::string& varName, const std::string& value);

        // Mirror Lua globals into _globalLuaVariablesMap for the debug
        // window dump.
        void captureLuaGlobals(lua_State* L);

        // Lazy compile cache: re-uses the previously compiled chunk when
        // the script hasn't changed. Mirrors the behaviour of the former
        // ensureLuaCodeCompiled.
        bool ensureCompiled(const std::string& scriptUtf8);

        // ----- State ------------------------------------------------------

        ILuaEngineHost* _host;                     // Non-owning, must outlive engine
        lua_State* _luaState = nullptr;

        // Compile cache
        int             _compiledReplaceRef = LUA_NOREF;
        std::string     _lastCompiledScript;

        // Per-match optimisation caches: avoid re-pushing globals that
        // didn't change since the previous match.
        std::string     _lastFPATH;
        std::string     _lastFNAME;
        int             _lastRegexFlag = -1;       // -1 = unset
        int             _lastCapCount = 0;
        std::vector<std::string> _currentCapNames;

        // Snapshot of Lua globals captured for the debug window. Cleared
        // and rebuilt on every debug dump.
        struct LuaVariableSnapshot {
            enum class Type { String, Number, Boolean, None };
            std::string name;
            Type        type = Type::None;
            std::string stringValue;
            double      numberValue = 0.0;
            bool        booleanValue = false;
        };
        std::map<std::string, LuaVariableSnapshot> _globalLuaVariablesMap;
    };

} // namespace MultiReplaceEngine