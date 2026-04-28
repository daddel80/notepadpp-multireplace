// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// LuaEngine.cpp
// Implementation of the Lua-backed IFormulaEngine. The behaviour and
// performance characteristics match the original resolveLuaSyntax-based
// bridge in MultiReplacePanel - only the surface changed.

#include "LuaEngine.h"

#include <iomanip>
#include <sstream>

// MR-internal helpers used by this engine. Including the actual headers
// (rather than forward-declaring) keeps the linkage robust against
// future signature changes.

// The bundled Lua helper script (set/skip/cond/lkp/...) is provided as
// a static string in luaEmbedded.h. Including the header here gives
// this TU its own visible copy of that static data.
#include "../luaEmbedded.h"

// SU::escapeControlChars is used to format Lua values for the debug
// window display. SU is an alias for StringUtils declared elsewhere
// in MR; we declare a minimal compatible alias here so the engine
// translation unit doesn't depend on the wider MR codebase layout.
#include "../StringUtils.h"
namespace SU = StringUtils;

namespace MultiReplaceEngine {

    // ---------------------------------------------------------------------
    // Lifecycle
    // ---------------------------------------------------------------------

    LuaEngine::LuaEngine(ILuaEngineHost* host)
        : _host(host)
    {
    }

    LuaEngine::~LuaEngine()
    {
        shutdown();
    }

    bool LuaEngine::initialize()
    {
        // Reset any prior state (idempotent re-init).
        shutdown();

        _luaState = luaL_newstate();
        if (!_luaState) {
            return false;
        }

        luaL_openlibs(_luaState);

        if (_host && _host->isLuaSafeModeEnabled()) {
            applyLuaSafeMode(_luaState);
        }

        // Load and execute the bundled helper script (set/skip/cond/lkp).
        if (luaL_loadstring(_luaState, luaSourceCode) != LUA_OK) {
            const char* errMsg = lua_tostring(_luaState, -1);
            if (_host && _host->isLuaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::CompileError,
                    errMsg ? errMsg : "Failed to load Lua helper script");
            }
            lua_close(_luaState);
            _luaState = nullptr;
            return false;
        }

        if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
            const char* errMsg = lua_tostring(_luaState, -1);
            if (_host && _host->isLuaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::ExecutionError,
                    errMsg ? errMsg : "Failed to execute Lua helper script");
            }
            lua_close(_luaState);
            _luaState = nullptr;
            return false;
        }

        // Wire the safe file-loader hook so user scripts can require()
        // helper files even when the rest of io/os is sandboxed away.
        lua_pushcfunction(_luaState, &LuaEngine::safeLoadFileSandbox);
        lua_setglobal(_luaState, "safeLoadFileSandbox");

        // Reset all per-match optimisation caches; a fresh state has no
        // globals so any "value last pushed" tracking is stale.
        _lastFPATH.clear();
        _lastFNAME.clear();
        _lastRegexFlag = -1;
        _lastCapCount = 0;
        _compiledReplaceRef = LUA_NOREF;
        _lastCompiledScript.clear();

        return true;
    }

    void LuaEngine::shutdown()
    {
        if (_luaState) {
            if (_compiledReplaceRef != LUA_NOREF) {
                luaL_unref(_luaState, LUA_REGISTRYINDEX, _compiledReplaceRef);
                _compiledReplaceRef = LUA_NOREF;
            }
            lua_close(_luaState);
            _luaState = nullptr;
        }
        _lastCompiledScript.clear();
        _lastFPATH.clear();
        _lastFNAME.clear();
        _lastRegexFlag = -1;
        _lastCapCount = 0;
        _currentCapNames.clear();
        _globalLuaVariablesMap.clear();
    }

    // ---------------------------------------------------------------------
    // Compile cache
    // ---------------------------------------------------------------------

    bool LuaEngine::ensureCompiled(const std::string& scriptUtf8)
    {
        if (!_luaState) { return false; }

        if (scriptUtf8 == _lastCompiledScript &&
            _compiledReplaceRef != LUA_NOREF) {
            return true;
        }

        if (_compiledReplaceRef != LUA_NOREF) {
            luaL_unref(_luaState, LUA_REGISTRYINDEX, _compiledReplaceRef);
            _compiledReplaceRef = LUA_NOREF;
        }

        if (luaL_loadstring(_luaState, scriptUtf8.c_str()) != LUA_OK) {
            const char* errMsg = lua_tostring(_luaState, -1);
            if (_host && _host->isLuaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::CompileError,
                    errMsg ? errMsg : "Lua compile error");
            }
            lua_pop(_luaState, 1);
            return false;
        }

        _compiledReplaceRef = luaL_ref(_luaState, LUA_REGISTRYINDEX);
        _lastCompiledScript = scriptUtf8;
        return true;
    }

    bool LuaEngine::compile(const std::string& scriptUtf8)
    {
        return ensureCompiled(scriptUtf8);
    }

    // ---------------------------------------------------------------------
    // Execute
    // ---------------------------------------------------------------------

    FormulaResult LuaEngine::execute(
        const std::string& scriptUtf8,
        const FormulaVars& vars,
        bool isRegexMatch,
        int /*documentCodepage*/)
    {
        FormulaResult result;
        result.output = scriptUtf8;  // default: pass-through if anything fails

        if (!_luaState) {
            result.success = false;
            result.errorMessage = "Lua state not initialized";
            return result;
        }

        if (!ensureCompiled(scriptUtf8)) {
            result.success = false;
            result.errorMessage = "Compile failed";
            return result;
        }

        // Stack-checkpoint so any early return cleanly drops what we pushed.
        const int stackBase = lua_gettop(_luaState);
        auto restoreStack = [this, stackBase]() {
            lua_settop(_luaState, stackBase);
            };

        // ----- Numeric globals --------------------------------------------
        lua_pushinteger(_luaState, vars.CNT);  lua_setglobal(_luaState, "CNT");
        lua_pushinteger(_luaState, vars.LCNT); lua_setglobal(_luaState, "LCNT");
        lua_pushinteger(_luaState, vars.LINE); lua_setglobal(_luaState, "LINE");
        lua_pushinteger(_luaState, vars.LPOS); lua_setglobal(_luaState, "LPOS");
        lua_pushinteger(_luaState, vars.APOS); lua_setglobal(_luaState, "APOS");
        lua_pushinteger(_luaState, vars.COL);  lua_setglobal(_luaState, "COL");

        // ----- String globals ---------------------------------------------
        // FPATH/FNAME change at most once per replaceAll run; skip the
        // push+setglobal pair when the value is unchanged since the
        // previous match.
        if (vars.FPATH != _lastFPATH) {
            setLuaVariable(_luaState, "FPATH", vars.FPATH);
            _lastFPATH = vars.FPATH;
        }
        if (vars.FNAME != _lastFNAME) {
            setLuaVariable(_luaState, "FNAME", vars.FNAME);
            _lastFNAME = vars.FNAME;
        }
        setLuaVariable(_luaState, "MATCH", vars.MATCH);

        // ----- REGEX flag -------------------------------------------------
        const int regexFlag = isRegexMatch ? 1 : 0;
        if (regexFlag != _lastRegexFlag) {
            lua_pushboolean(_luaState, isRegexMatch);
            lua_setglobal(_luaState, "REGEX");
            _lastRegexFlag = regexFlag;
        }

        // ----- CAP# globals (regex only) ----------------------------------
        // Captures arrive pre-extracted from the host; the engine just
        // pushes them as Lua globals. Encoding conversion already happened
        // in the pipeline (the host knows the document codepage; the engine
        // shouldn't need to).
        _currentCapNames.clear();
        if (isRegexMatch) {
            for (size_t i = 0; i < vars.captures.size(); ++i) {
                std::string capName = "CAP" + std::to_string(i + 1);
                setLuaVariable(_luaState, capName, vars.captures[i]);
                _currentCapNames.push_back(capName);
            }
        }

        // ----- Run pre-compiled chunk -------------------------------------
        lua_rawgeti(_luaState, LUA_REGISTRYINDEX, _compiledReplaceRef);
        if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
            const char* err = lua_tostring(_luaState, -1);
            if (_host && _host->isLuaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::CompileError,
                    err ? err : "Lua execution error");
            }
            result.success = false;
            result.errorMessage = err ? err : "Lua execution error";
            restoreStack();
            return result;
        }

        // ----- resultTable ------------------------------------------------
        lua_getglobal(_luaState, "resultTable");
        if (!lua_istable(_luaState, -1)) {
            if (_host && _host->isLuaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::ExecutionError,
                    scriptUtf8);
            }
            result.success = false;
            result.errorMessage = "Lua produced no resultTable";
            restoreStack();
            return result;
        }

        // ----- result & skip ----------------------------------------------
        lua_getfield(_luaState, -1, "result");
        if (lua_isnil(_luaState, -1)) {
            result.output.clear();
        }
        else if (lua_isstring(_luaState, -1) || lua_isnumber(_luaState, -1)) {
            std::string res = lua_tostring(_luaState, -1);
            if (isRegexMatch && _host) {
                res = _host->escapeForRegex(res);
            }
            result.output = res;
        }
        lua_pop(_luaState, 1); // pop result

        lua_getfield(_luaState, -1, "skip");
        result.skip = lua_isboolean(_luaState, -1)
            && lua_toboolean(_luaState, -1);
        lua_pop(_luaState, 1); // pop skip

        // ----- Debug-window decision --------------------------------------
        lua_getglobal(_luaState, "DEBUG");
        const bool luaDebugExists = !lua_isnil(_luaState, -1);
        const bool luaDebug = luaDebugExists && lua_toboolean(_luaState, -1);
        lua_pop(_luaState, 1);
        const bool hostDebug = _host && _host->isDebugModeEnabled();
        const bool debugOn = luaDebugExists ? luaDebug : hostDebug;
        const bool needCapDump = debugOn;

        // ----- CAP dump (only when debug is on) ---------------------------
        std::string capVariablesStr;
        if (needCapDump) {
            for (const auto& capName : _currentCapNames) {
                lua_getglobal(_luaState, capName.c_str());

                if (lua_isnumber(_luaState, -1)) {
                    double n = lua_tonumber(_luaState, -1);
                    std::ostringstream os;
                    os << std::fixed << std::setprecision(8) << n;
                    capVariablesStr += capName + "\tNumber\t" + os.str() + "\n\n";
                }
                else if (lua_isboolean(_luaState, -1)) {
                    bool b = lua_toboolean(_luaState, -1);
                    capVariablesStr += capName + "\tBoolean\t"
                        + (b ? "true" : "false") + "\n\n";
                }
                else if (lua_isstring(_luaState, -1)) {
                    capVariablesStr += capName + "\tString\t"
                        + SU::escapeControlChars(lua_tostring(_luaState, -1))
                        + "\n\n";
                }

                lua_pop(_luaState, 1);
            }
        }

        // ----- CAP cleanup ------------------------------------------------
        // Only nil the CAPs that were set last match but not this match.
        // CAPs set this match get overwritten next time, so they need no
        // cleanup; only stale CAPs from a longer previous match would leak.
        const int currentCapCount = static_cast<int>(_currentCapNames.size());
        for (int i = currentCapCount; i < _lastCapCount; ++i) {
            std::string capName = "CAP" + std::to_string(i + 1);
            lua_pushnil(_luaState);
            lua_setglobal(_luaState, capName.c_str());
        }
        _lastCapCount = currentCapCount;

        // ----- Debug-window display ---------------------------------------
        if (needCapDump && _host) {
            _globalLuaVariablesMap.clear();
            captureLuaGlobals(_luaState);

            std::string globalsStr = "Global Lua variables:\n\n";
            for (const auto& p : _globalLuaVariablesMap) {
                const LuaVariableSnapshot& v = p.second;
                switch (v.type) {
                case LuaVariableSnapshot::Type::String:
                    globalsStr += v.name + "\tString\t"
                        + SU::escapeControlChars(v.stringValue) + "\n\n";
                    break;
                case LuaVariableSnapshot::Type::Number: {
                    std::ostringstream os;
                    os << std::fixed << std::setprecision(8) << v.numberValue;
                    globalsStr += v.name + "\tNumber\t" + os.str() + "\n\n";
                    break;
                }
                case LuaVariableSnapshot::Type::Boolean:
                    globalsStr += v.name + "\tBoolean\t"
                        + (v.booleanValue ? "true" : "false") + "\n\n";
                    break;
                case LuaVariableSnapshot::Type::None:
                    break;
                }
            }

            _host->refreshUiListView();

            const int resp = _host->showDebugWindow(capVariablesStr + globalsStr);

            if (resp == 3 || resp == -1) {
                // "Stop" pressed (3) or window closed (-1)
                result.success = false;
                result.errorMessage = "Aborted via debug window";
                restoreStack();
                return result;
            }
        }

        restoreStack();
        return result;
    }

    // ---------------------------------------------------------------------
    // Metadata
    // ---------------------------------------------------------------------

    std::wstring LuaEngine::helpUrl() const
    {
        // Help target for the panel's "?" button when Lua is the active
        // engine. Resolved to a URL or local file path by the host's
        // launcher; the engine doesn't open it itself.
        return L"https://github.com/daddel80/notepadpp-multireplace#use-variables";
    }

    // ---------------------------------------------------------------------
    // Internal helpers
    // ---------------------------------------------------------------------

    void LuaEngine::setLuaVariable(lua_State* L, const std::string& varName, const std::string& value)
    {
        lua_pushstring(L, value.c_str());
        lua_setglobal(L, varName.c_str());
    }

    void LuaEngine::captureLuaGlobals(lua_State* L)
    {
        lua_pushglobaltable(L);
        lua_pushnil(L);
        while (lua_next(L, -2) != 0) {
            // Skip non-string keys: lua_tostring on a numeric key would
            // convert it in-place and break lua_next traversal.
            if (lua_type(L, -2) != LUA_TSTRING) {
                lua_pop(L, 1);
                continue;
            }

            const char* key = lua_tostring(L, -2);
            LuaVariableSnapshot snapshot;
            snapshot.name = key;

            const int valType = lua_type(L, -1);
            if (valType == LUA_TNUMBER) {
                snapshot.type = LuaVariableSnapshot::Type::Number;
                snapshot.numberValue = lua_tonumber(L, -1);
            }
            else if (valType == LUA_TSTRING) {
                snapshot.type = LuaVariableSnapshot::Type::String;
                snapshot.stringValue = lua_tostring(L, -1);
            }
            else if (valType == LUA_TBOOLEAN) {
                snapshot.type = LuaVariableSnapshot::Type::Boolean;
                snapshot.booleanValue = lua_toboolean(L, -1);
            }
            else {
                // Skip unsupported types (table, function, userdata, ...)
                lua_pop(L, 1);
                continue;
            }

            _globalLuaVariablesMap[key] = std::move(snapshot);
            lua_pop(L, 1);
        }
        lua_pop(L, 1); // pop the global table
    }

    // ---------------------------------------------------------------------
    // Static Lua hooks (require C-style entry points)
    // ---------------------------------------------------------------------

    int LuaEngine::safeLoadFileSandbox(lua_State* L)
    {
        // Forwards to the implementation that lives in MR's Lua-helper TU
        // (the file loader needs MR's Encoding utilities and is therefore
        // not duplicated here). Defined as extern "C++" to keep linkage
        // identical to the original static implementation.
        extern int luaSafeLoadFileSandbox_impl(lua_State * L);
        return luaSafeLoadFileSandbox_impl(L);
    }

    void LuaEngine::applyLuaSafeMode(lua_State* L)
    {
        auto removeGlobal = [&](const char* name) {
            lua_pushnil(L);
            lua_setglobal(L, name);
            };

        // Remove dangerous base functions
        removeGlobal("dofile");
        removeGlobal("load");
        removeGlobal("loadfile");
        removeGlobal("require");
        removeGlobal("collectgarbage");

        // Remove whole libraries
        removeGlobal("os");
        removeGlobal("io");
        removeGlobal("package");
        removeGlobal("debug");

        // Keep string/table/math/utf8/base intact.
    }

} // namespace MultiReplaceEngine