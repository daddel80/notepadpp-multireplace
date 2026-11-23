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
// IFormulaEngineHost.h
// Callback interface the formula engines use to talk back to their host
// (the MultiReplace panel). Kept in its own header so callers - notably
// MultiReplacePanel.h - can implement the interface without pulling
// <lua.hpp> or ExprTk through their own translation unit.

#pragma once

#include <string>

namespace MultiReplaceEngine {

    // Callback hooks the engines need back into the host. Formulas can
    // trigger UI (the debug window, error message boxes), and ExprTk output
    // is regex-escaped before going back into a regex replacement; both
    // pieces of logic live in MultiReplace, not in the
    // engine. Passing them as a small interface keeps the engine free of
    // MR types and trivially mockable in tests.
    //
    // Implementations are allowed to be no-ops (e.g. in tests).
    class IFormulaEngineHost {
    public:
        virtual ~IFormulaEngineHost() = default;

        // Escape characters that have meaning in a regex replacement so
        // formula output doesn't accidentally introduce backreferences or
        // escape sequences. Pure string transformation, no I/O or UI.
        virtual std::string escapeForRegex(const std::string& input) = 0;

        // Show the debug window with the engine's variable dump.
        // Return values mirror the existing ShowDebugWindow contract:
        //    3 -> "Stop", or the document changed while the window waited
        //   -1 -> window closed
        //  any other -> continue
        virtual int showDebugWindow(const std::string& message) = 0;

        // Refresh the panel's list view after a debug-window interruption.
        // Called only when the debug window is about to be shown so the
        // list reflects any state the user-script altered.
        virtual void refreshUiListView() = 0;

        // Report a formula error that stops the run: a message box, or a
        // status line when error dialogs are turned off. The engine supplies
        // its identifier (e.g. "Lua", "ExprTk"), the raw error text and a
        // category; the host composes the translated text, so the "Lua: ..."
        // or "ExprTk: ..." prefix never has to be hardcoded in an engine.
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
        // True after "Run to End" in the debug window until the next run.
        virtual bool isDebugRunSilenced()      const = 0;

        // CSV column access for the line containing the current match.
        // Index is 1-based; addressing by name uses the document's first
        // line as a header row, parsed lazily on first use. Returns false
        // if CSV mode is not active, no current match is set, the column
        // does not exist on this row, or (for the by-name variant) the
        // header has no matching field. On false, out is left untouched.
        // Cells are returned verbatim - no quote stripping - matching the
        // existing extractColumnData convention.
        virtual bool readCurrentRowColumnByIndex(int colIndex1Based,
            std::string& out) const = 0;
        virtual bool readCurrentRowColumnByName(const std::string& headerName,
            std::string& out) const = 0;
    };

} // namespace MultiReplaceEngine