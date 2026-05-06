// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ILuaEngineHost.h
// Callback interface the LuaEngine uses to talk back to its host
// (typically the MultiReplace panel). Kept in its own header so
// callers - notably MultiReplacePanel.h - can implement the interface
// without pulling <lua.hpp> through their own translation unit.

#pragma once

#include <string>

namespace MultiReplaceEngine {

    // Callback hooks the engine needs back into the host. Lua scripts can
    // trigger UI (the debug window, error message boxes), and the result of
    // set() may need to be regex-escaped before going back into a regex
    // replacement; both pieces of logic live in MultiReplace, not in the
    // engine. Passing them as a small interface keeps the engine free of
    // MR types and trivially mockable in tests.
    //
    // Implementations are allowed to be no-ops (e.g. in tests).
    class ILuaEngineHost {
    public:
        virtual ~ILuaEngineHost() = default;

        // Escape characters that have meaning in a regex replacement so
        // user-produced text from Lua doesn't accidentally introduce
        // backreferences or escape sequences. Pure string transformation,
        // no I/O or UI.
        virtual std::string escapeForRegex(const std::string& input) = 0;

        // Show the Lua debug window with the captured CAP/global dump.
        // Return values mirror the existing ShowDebugWindow contract:
        //    3 -> user pressed "Stop"
        //   -1 -> window closed
        //  any other -> continue
        virtual int showDebugWindow(const std::string& message) = 0;

        // Refresh the panel's list view after a debug-window interruption.
        // Called only when the debug window is about to be shown so the
        // list reflects any state the user-script altered.
        virtual void refreshUiListView() = 0;

        // Display a translated error message box. Title and body come from
        // the host's translation table; the engine supplies its own
        // identifier (e.g. "Lua", "ExprTk") plus the raw error text and a
        // category key. The host composes the final body from a translation
        // template such as "$REPLACE_STRING1: $REPLACE_STRING2", so the
        // user-visible "Lua: ..." or "ExprTk: ..." prefix never has to be
        // hardcoded on the engine side.
        enum class ErrorCategory { CompileError, ExecutionError };
        virtual void showErrorMessage(ErrorCategory category,
            const std::string& engineName,
            const std::string& details) = 0;

        // Three-button dialog for recoverable runtime errors (used by
        // ExprTk on NaN). The host wraps the engine-supplied details
        // with rule index and match position from its own state.
        enum class RecoverableErrorChoice { SkipOne = 0, SkipAll = 1, Stop = 2 };
        virtual RecoverableErrorChoice showRecoverableErrorDialog(
            const std::string& engineName,
            const std::string& details) = 0;

        // User-toggleable mode flags. Read every match (rather than copied
        // at init) so changes apply immediately to the next evaluation.
        virtual bool isFormulaErrorDialogEnabled() const = 0;
        virtual bool isLuaSafeModeEnabled()    const = 0;
        virtual bool isDebugModeEnabled()      const = 0;
    };

} // namespace MultiReplaceEngine