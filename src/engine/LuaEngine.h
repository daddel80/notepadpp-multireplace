// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.
//
// LuaEngine.h
// Concrete IFormulaEngine implementation backed by a Lua 5.4 state.
//
// Every evaluation runs under a time and a memory limit, so a runaway
// script stops with a message instead of freezing Notepad++.

#pragma once

#include "IFormulaEngine.h"
#include "IFormulaEngineHost.h"

#include <lua.hpp>

#include <chrono>
#include <string>
#include <unordered_map>
#include <vector>

namespace MultiReplaceEngine {

    // One Lua state per document; the helper script and the variables of
    // each match live in it.
    class LuaEngine final : public IFormulaEngine {
    public:
        // The host pointer must outlive the engine; LuaEngine does not
        // take ownership.
        explicit LuaEngine(IFormulaEngineHost* host);
        ~LuaEngine() override;

        // ----- IFormulaEngine ---------------------------------------------

        bool initialize() override;
        void shutdown() override;

        // Each document gets a new Lua state: files loaded by lcmd(),
        // lvars() and lkp() are read again, and removing the lcmd() call
        // also removes its functions from the global namespace.
        void beginDocument() override;

        bool compile(const std::string& scriptUtf8, std::size_t ruleIndex = static_cast<std::size_t>(-1)) override;

        FormulaResult execute(
            const std::string& scriptUtf8,
            const FormulaVars& vars,
            ReplaceMode mode,
            int documentCodepage
        ) override;

        EngineType type() const override { return EngineType::Lua; }
        std::wstring shortName() const override { return L"Lua"; }
        std::wstring shortLetter() const override { return L"L"; }

        // ----- Lua-specific helpers (internal use & sandbox callbacks) ----
        //
        // These are public only because Lua's C-API requires C-style free
        // functions for hooks. External callers should not depend on them.

        // safeLoadFileSandbox(path) -> true, value | false, message. Runs a
        // Lua text file (UTF-8, or the ANSI codepage) for lkp/lvars/lcmd.
        static int  safeLoadFileSandbox(lua_State* L);
        static void applyLuaSafeMode(lua_State* L);

    private:
        // ----- Internal helpers -------------------------------------------

        // Push a string variable to the Lua global table. Centralised so
        // UTF-8 handling and registry conventions stay in one place.
        void setLuaVariable(lua_State* L, const std::string& varName, const std::string& value);

        // Globals the engine sets and reads skip metamethods: a metatable
        // the script puts on _G must not raise outside lua_pcall.
        static void setGlobalRaw(lua_State* L, const char* name);  // pops the value
        static int  getGlobalRaw(lua_State* L, const char* name);  // pushes the value

        // The debug window's text: the CAPs of this match first, then the
        // other plain-valued globals by name, each value once.
        std::string debugDump();

        // Registry reference of the compiled script, LUA_NOREF after a
        // compile error. Each script compiles once per document.
        int compiledRef(const std::string& scriptUtf8);

        // Allocator with a ceiling while the formula runs, and the count
        // hook that ends an evaluation running past its deadline.
        static void* limitedAlloc(void* ud, void* ptr, size_t osize, size_t nsize);
        static void  runawayHook(lua_State* L, lua_Debug* ar);

        // ----- State ------------------------------------------------------

        IFormulaEngineHost* _host;                     // Non-owning, must outlive engine
        lua_State* _luaState = nullptr;

        // Compiled scripts of this document, by script text
        std::unordered_map<std::string, int> _compiled;

        // Runaway limits of the running evaluation: deadline, memory
        // ceiling, bytes in use, and whether the deadline stopped it
        std::chrono::steady_clock::time_point _deadline;
        size_t          _memoryCeiling = static_cast<size_t>(-1);
        size_t          _memoryUsed = 0;
        bool            _timedOut = false;

        // CAP globals of the current match, and how many the previous one
        // set, so the CAPs of a match with more groups do not linger
        int             _lastCapCount = 0;
        std::vector<std::string> _currentCapNames;
    };

} // namespace MultiReplaceEngine