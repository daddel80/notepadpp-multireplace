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
// LuaEngine.cpp
// Implementation of the Lua-backed IFormulaEngine.

#include "LuaEngine.h"

#include <cstdio>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iterator>
#include <sstream>

// The bundled Lua helper script (set/skip/cond/lkp/...) is provided as
// a static string in luaEmbedded.h. Including the header here gives
// this TU its own visible copy of that static data.
#include "../luaEmbedded.h"

// SU::escapeControlChars formats Lua values for the debug window;
// Encoding reads the files behind lkp/lvars/lcmd.
#include "../StringUtils.h"
#include "../Encoding.h"
namespace SU = StringUtils;

namespace MultiReplaceEngine {

    namespace {

        // One evaluation (one match or init row) may run this long.
        constexpr int kMaxRunSeconds = 10;
        constexpr const char* kStopFormat = "Formula stopped after %d seconds";

        // Lua instructions between two deadline checks.
        constexpr int kHookInterval = 10000;

        // Memory one evaluation may add; outside the formula nothing is refused.
        constexpr size_t kMaxMemory = static_cast<size_t>(1024) * 1024 * 1024;
        constexpr size_t kNoCeiling = static_cast<size_t>(-1);

    } // anonymous namespace

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

        _memoryUsed = 0;
        _memoryCeiling = kNoCeiling;
        _luaState = lua_newstate(&LuaEngine::limitedAlloc, this);
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
            if (_host && _host->isFormulaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::CompileError,
                    "Lua",
                    errMsg ? errMsg : "Failed to load Lua helper script");
            }
            lua_close(_luaState);
            _luaState = nullptr;
            return false;
        }

        if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
            const char* errMsg = lua_tostring(_luaState, -1);
            if (_host && _host->isFormulaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::ExecutionError,
                    "Lua",
                    errMsg ? errMsg : "Failed to execute Lua helper script");
            }
            lua_close(_luaState);
            _luaState = nullptr;
            return false;
        }
        lua_settop(_luaState, 0);

        // Wire the safe file-loader hook so user scripts can require()
        // helper files even when the rest of io/os is sandboxed away.
        lua_pushcfunction(_luaState, &LuaEngine::safeLoadFileSandbox);
        setGlobalRaw(_luaState, "safeLoadFileSandbox");

        // The count hook finds the engine through the state's extra space;
        // coroutines get a copy of it and inherit the hook.
        *static_cast<LuaEngine**>(lua_getextraspace(_luaState)) = this;
        lua_sethook(_luaState, &LuaEngine::runawayHook, LUA_MASKCOUNT, kHookInterval);

        // Reset all per-match optimisation caches; a fresh state has no
        // globals so any "value last pushed" tracking is stale.
        _lastFPATH.clear();
        _lastFNAME.clear();
        _lastRegexFlag = -1;
        _lastCapCount = 0;

        return true;
    }

    void LuaEngine::shutdown()
    {
        // Closing the state also frees every compiled chunk it references.
        if (_luaState) {
            lua_close(_luaState);
            _luaState = nullptr;
        }
        _compiled.clear();
        _lastFPATH.clear();
        _lastFNAME.clear();
        _lastRegexFlag = -1;
        _lastCapCount = 0;
        _currentCapNames.clear();
        _globalLuaVariablesMap.clear();
    }

    // ---------------------------------------------------------------------
    // Per-run lifecycle
    // ---------------------------------------------------------------------

    void LuaEngine::beginRun()
    {
        IFormulaEngine::beginRun();

        // Tear the state down and rebuild from scratch. Cheap (sub-ms)
        // and guarantees a clean slate: any lcmd-loaded helper, any
        // user-set global, any side-effect from a previous run is gone.
        // Re-runs of the same list now also re-read .lcmd files from
        // disk, so the user's edits take effect on the next click.
        if (_luaState) {
            shutdown();
            initialize();
        }
    }

    // ---------------------------------------------------------------------
    // Compile cache
    // ---------------------------------------------------------------------

    int LuaEngine::compiledRef(const std::string& scriptUtf8)
    {
        if (!_luaState) { return LUA_NOREF; }

        const auto it = _compiled.find(scriptUtf8);
        if (it != _compiled.end()) {
            return it->second;
        }

        // The script text doubles as chunk name, so messages quote it;
        // text chunks only, no precompiled binary chunks.
        if (luaL_loadbufferx(_luaState, scriptUtf8.data(), scriptUtf8.size(),
            scriptUtf8.c_str(), "t") != LUA_OK) {
            const char* errMsg = lua_tostring(_luaState, -1);
            if (_host && _host->isFormulaErrorDialogEnabled()) {
                _host->showErrorMessage(
                    ILuaEngineHost::ErrorCategory::CompileError,
                    "Lua",
                    errMsg ? errMsg : "Lua compile error");
            }
            lua_pop(_luaState, 1);
            return LUA_NOREF;
        }

        const int ref = luaL_ref(_luaState, LUA_REGISTRYINDEX);
        _compiled.emplace(scriptUtf8, ref);
        return ref;
    }

    bool LuaEngine::compile(const std::string& scriptUtf8, std::size_t /*ruleIndex*/)
    {
        return compiledRef(scriptUtf8) != LUA_NOREF;
    }

    // ---------------------------------------------------------------------
    // Execute
    // ---------------------------------------------------------------------

    FormulaResult LuaEngine::execute(
        const std::string& scriptUtf8,
        const FormulaVars& vars,
        ReplaceMode mode,
        int /*documentCodepage*/)
    {
        const bool isRegexMatch = (mode == ReplaceMode::Regex);
        FormulaResult result;
        result.outputIsRegexSafe = false;  // inserted as is, also in regex mode

        if (!_luaState) {
            result.success = false;
            result.errorMessage = "Lua state not initialized";
            return result;
        }

        const int ref = compiledRef(scriptUtf8);
        if (ref == LUA_NOREF) {
            result.success = false;
            result.errorMessage = "Compile failed";
            return result;
        }

        // Stack-checkpoint so any early return cleanly drops what we pushed.
        const int stackBase = lua_gettop(_luaState);
        auto restoreStack = [this, stackBase]() {
            lua_settop(_luaState, stackBase);
            };
        auto fail = [&](ILuaEngineHost::ErrorCategory category, const std::string& message) {
            if (_host && _host->isFormulaErrorDialogEnabled()) {
                _host->showErrorMessage(category, "Lua", message);
            }
            result.success = false;
            result.errorMessage = message;
            restoreStack();
            return result;
            };

        // ----- Numeric globals --------------------------------------------
        // Each variable is exposed under both upper- and lowercase
        // so users may write whichever style they prefer.
        lua_pushinteger(_luaState, vars.CNT);  setGlobalRaw(_luaState, "CNT");
        lua_pushinteger(_luaState, vars.CNT);  setGlobalRaw(_luaState, "cnt");
        lua_pushinteger(_luaState, vars.LCNT); setGlobalRaw(_luaState, "LCNT");
        lua_pushinteger(_luaState, vars.LCNT); setGlobalRaw(_luaState, "lcnt");
        lua_pushinteger(_luaState, vars.LINE); setGlobalRaw(_luaState, "LINE");
        lua_pushinteger(_luaState, vars.LINE); setGlobalRaw(_luaState, "line");
        lua_pushinteger(_luaState, vars.LPOS); setGlobalRaw(_luaState, "LPOS");
        lua_pushinteger(_luaState, vars.LPOS); setGlobalRaw(_luaState, "lpos");
        lua_pushinteger(_luaState, vars.APOS); setGlobalRaw(_luaState, "APOS");
        lua_pushinteger(_luaState, vars.APOS); setGlobalRaw(_luaState, "apos");
        lua_pushinteger(_luaState, vars.COL);  setGlobalRaw(_luaState, "COL");
        lua_pushinteger(_luaState, vars.COL);  setGlobalRaw(_luaState, "col");

        // ----- String globals ---------------------------------------------
        // FPATH/FNAME change at most once per replaceAll run; skip the
        // push+setglobal pair when the value is unchanged since the
        // previous match.
        if (vars.FPATH != _lastFPATH) {
            setLuaVariable(_luaState, "FPATH", vars.FPATH);
            setLuaVariable(_luaState, "fpath", vars.FPATH);
            _lastFPATH = vars.FPATH;
        }
        if (vars.FNAME != _lastFNAME) {
            setLuaVariable(_luaState, "FNAME", vars.FNAME);
            setLuaVariable(_luaState, "fname", vars.FNAME);
            _lastFNAME = vars.FNAME;
        }
        setLuaVariable(_luaState, "MATCH", vars.MATCH);
        setLuaVariable(_luaState, "match", vars.MATCH);

        // ----- REGEX flag -------------------------------------------------
        const int regexFlag = isRegexMatch ? 1 : 0;
        if (regexFlag != _lastRegexFlag) {
            lua_pushboolean(_luaState, isRegexMatch);
            setGlobalRaw(_luaState, "REGEX");
            lua_pushboolean(_luaState, isRegexMatch);
            setGlobalRaw(_luaState, "regex");
            _lastRegexFlag = regexFlag;
        }

        // ----- CAP# globals (regex only) ----------------------------------
        // Captures arrive pre-extracted from the host; the engine just
        // pushes them as Lua globals. Encoding conversion already happened
        // in the pipeline (the host knows the document codepage; the engine
        // shouldn't need to).
        // Each capture is exposed under both CAP# and cap# names.
        _currentCapNames.clear();
        if (isRegexMatch) {
            for (size_t i = 0; i < vars.captures.size(); ++i) {
                std::string capName = "CAP" + std::to_string(i + 1);
                std::string capNameLower = "cap" + std::to_string(i + 1);
                setLuaVariable(_luaState, capName, vars.captures[i]);
                setLuaVariable(_luaState, capNameLower, vars.captures[i]);
                _currentCapNames.push_back(capName);
            }
        }

        // CAPs of a previous match with more groups must not be seen by
        // this one; both upper- and lowercase aliases are cleared.
        const int currentCapCount = static_cast<int>(_currentCapNames.size());
        for (int i = currentCapCount; i < _lastCapCount; ++i) {
            std::string capName = "CAP" + std::to_string(i + 1);
            std::string capNameLower = "cap" + std::to_string(i + 1);
            lua_pushnil(_luaState);
            setGlobalRaw(_luaState, capName.c_str());
            lua_pushnil(_luaState);
            setGlobalRaw(_luaState, capNameLower.c_str());
        }
        _lastCapCount = currentCapCount;

        // ----- Run pre-compiled chunk -------------------------------------
        // The result belongs to this match only; without set()/cond() the
        // match stays unchanged.
        lua_pushnil(_luaState);
        setGlobalRaw(_luaState, "resultTable");

        // Limits of this evaluation; a stop the script catches ends the run anyway.
        _timedOut = false;
        _deadline = std::chrono::steady_clock::now() + std::chrono::seconds(kMaxRunSeconds);
        _memoryCeiling = (_memoryUsed > kNoCeiling - kMaxMemory) ? kNoCeiling : _memoryUsed + kMaxMemory;
        lua_sethook(_luaState, &LuaEngine::runawayHook, LUA_MASKCOUNT, kHookInterval);
        lua_rawgeti(_luaState, LUA_REGISTRYINDEX, ref);
        const int status = lua_pcall(_luaState, 0, LUA_MULTRET, 0);
        _memoryCeiling = kNoCeiling;
        if (_timedOut) {
            char stopMessage[64];
            std::snprintf(stopMessage, sizeof(stopMessage), kStopFormat, kMaxRunSeconds);
            return fail(ILuaEngineHost::ErrorCategory::ExecutionError, stopMessage);
        }
        if (status != LUA_OK) {
            const char* err = lua_tostring(_luaState, -1);
            return fail(ILuaEngineHost::ErrorCategory::ExecutionError,
                err ? err : "Lua execution error");
        }

        // ----- resultTable ------------------------------------------------
        getGlobalRaw(_luaState, "resultTable");
        if (lua_isnil(_luaState, -1)) {
            result.skip = true;
        }
        else if (!lua_istable(_luaState, -1)) {
            return fail(ILuaEngineHost::ErrorCategory::ExecutionError,
                "resultTable must be a table");
        }
        else {
            // ----- result & skip ------------------------------------------
            lua_pushliteral(_luaState, "result");
            const int resultType = lua_rawget(_luaState, -2);
            if (resultType == LUA_TSTRING || resultType == LUA_TNUMBER) {
                size_t len = 0;
                const char* text = lua_tolstring(_luaState, -1, &len);
                result.output.assign(text, len);
            }
            else if (resultType != LUA_TNIL) {
                return fail(ILuaEngineHost::ErrorCategory::ExecutionError,
                    std::string("The result must be a string or a number, not a ")
                    + lua_typename(_luaState, resultType));
            }
            lua_pop(_luaState, 1); // pop result

            lua_pushliteral(_luaState, "skip");
            lua_rawget(_luaState, -2);
            result.skip = lua_isboolean(_luaState, -1)
                && lua_toboolean(_luaState, -1);
            lua_pop(_luaState, 1); // pop skip
        }

        // Extended mode resolves \n, \t, \\ in what follows; a backslash
        // from the formula stays one.
        if (mode == ReplaceMode::Extended && !result.skip) {
            std::string escaped;
            escaped.reserve(result.output.size());
            for (const char c : result.output) {
                if (c == '\\') escaped.push_back('\\');
                escaped.push_back(c);
            }
            result.output = std::move(escaped);
        }

        // ----- Debug-window decision --------------------------------------
        // A script's DEBUG overrides the menu toggle; "Run to End" in the
        // debug window silences both for the rest of the run.
        getGlobalRaw(_luaState, "DEBUG");
        const bool luaDebugExists = !lua_isnil(_luaState, -1);
        const bool luaDebug = luaDebugExists && lua_toboolean(_luaState, -1);
        lua_pop(_luaState, 1);
        const bool hostDebug = _host && _host->isDebugModeEnabled();
        const bool silenced = _host && _host->isDebugRunSilenced();
        const bool debugOn = !silenced && (luaDebugExists ? luaDebug : hostDebug);

        // ----- Debug-window display ---------------------------------------
        if (debugOn && _host) {
            std::string capVariablesStr;
            for (const auto& capName : _currentCapNames) {
                getGlobalRaw(_luaState, capName.c_str());

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
    // Internal helpers
    // ---------------------------------------------------------------------

    void LuaEngine::setLuaVariable(lua_State* L, const std::string& varName, const std::string& value)
    {
        lua_pushlstring(L, value.data(), value.size());
        setGlobalRaw(L, varName.c_str());
    }

    void LuaEngine::setGlobalRaw(lua_State* L, const char* name)
    {
        lua_pushglobaltable(L);
        lua_pushstring(L, name);
        lua_rotate(L, -3, -1);  // _G, name, value
        lua_rawset(L, -3);
        lua_pop(L, 1);
    }

    int LuaEngine::getGlobalRaw(lua_State* L, const char* name)
    {
        lua_pushglobaltable(L);
        lua_pushstring(L, name);
        const int type = lua_rawget(L, -2);
        lua_remove(L, -2);  // the value stays
        return type;
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
    // Runaway limits
    // ---------------------------------------------------------------------

    void* LuaEngine::limitedAlloc(void* ud, void* ptr, size_t osize, size_t nsize)
    {
        LuaEngine& self = *static_cast<LuaEngine*>(ud);
        const size_t oldSize = ptr ? osize : 0;  // osize is a type tag for new blocks
        if (nsize == 0) {
            std::free(ptr);
            self._memoryUsed -= oldSize;
            return nullptr;
        }
        // Refusing makes Lua raise "not enough memory" in the formula
        if (nsize > oldSize && self._memoryUsed - oldSize + nsize > self._memoryCeiling) {
            return nullptr;
        }
        void* block = std::realloc(ptr, nsize);
        if (block) {
            self._memoryUsed = self._memoryUsed - oldSize + nsize;
        }
        return block;
    }

    void LuaEngine::runawayHook(lua_State* L, lua_Debug* /*ar*/)
    {
        LuaEngine* self = *static_cast<LuaEngine**>(lua_getextraspace(L));
        if (self && std::chrono::steady_clock::now() > self->_deadline) {
            // From now on every instruction raises again, so a pcall in the
            // script cannot keep it running
            self->_timedOut = true;
            lua_sethook(L, &LuaEngine::runawayHook, LUA_MASKCOUNT, 1);
            lua_sethook(self->_luaState, &LuaEngine::runawayHook, LUA_MASKCOUNT, 1);
            luaL_error(L, kStopFormat, kMaxRunSeconds);
        }
    }

    // ---------------------------------------------------------------------
    // Static Lua hooks (require C-style entry points)
    // ---------------------------------------------------------------------

    int LuaEngine::safeLoadFileSandbox(lua_State* L)
    {
        const char* path = luaL_checkstring(L, 1);

        std::ifstream in(std::filesystem::path(Encoding::utf8ToWString(path)), std::ios::binary);
        if (!in) {
            lua_pushboolean(L, false);
            lua_pushfstring(L, "Cannot open file: %s", path);
            return 2;
        }
        std::string text((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        in.close();

        // A file that is not UTF-8 is read in the ANSI codepage
        if (!Encoding::isValidUtf8(text.data(), text.size())) {
            text = Encoding::bytesToUtf8(text.data(), text.size(), CP_ACP);
        }
        if (text.size() >= 3
            && static_cast<unsigned char>(text[0]) == 0xEF
            && static_cast<unsigned char>(text[1]) == 0xBB
            && static_cast<unsigned char>(text[2]) == 0xBF) {
            text.erase(0, 3);
        }

        // Text chunks only, no precompiled binary chunks
        if (luaL_loadbufferx(L, text.data(), text.size(), path, "t") != LUA_OK
            || lua_pcall(L, 0, 1, 0) != LUA_OK) {
            lua_pushboolean(L, false);
            lua_insert(L, -2);  // false, message
            return 2;
        }
        lua_pushboolean(L, true);
        lua_insert(L, -2);  // true, value
        return 2;
    }

    void LuaEngine::applyLuaSafeMode(lua_State* L)
    {
        auto removeGlobal = [&](const char* name) {
            lua_pushnil(L);
            setGlobalRaw(L, name);
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