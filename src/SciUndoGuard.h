// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025 Thomas Knoefel
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

#pragma once

#include <windows.h>
#include "Scintilla.h"

// ============================================================================
// SciUndoGuard - RAII guard for Scintilla undo actions with nesting support
// ============================================================================
//
// PURPOSE:
// Manages Scintilla's SCI_BEGINUNDOACTION / SCI_ENDUNDOACTION with automatic
// nesting detection. Multiple text operations wrapped in a single guard
// become ONE undo step for the user.
//
// NOTE: This is SEPARATE from UndoRedoManager, which handles plugin-level
// undo (list operations, UI state). SciUndoGuard handles DOCUMENT-level
// undo (text modifications via Scintilla).
//
// NESTING SUPPORT:
// A thread-local counter tracks nesting depth. Only the outermost guard
// sends BEGIN/END to Scintilla. Inner guards are no-ops.
//
// USAGE:
//   // Simple case:
//   {
//       SciUndoGuard undo(hScintilla);
//       // ... multiple SCI_INSERTTEXT, SCI_DELETERANGE calls ...
//   }
//   // All changes = ONE undo step
//
//   // Nested case (inner guards are automatic no-ops):
//   {
//       SciUndoGuard outer(hSci);
//       SomeFunction();  // May create its own SciUndoGuard - that's OK
//   }
//
// ============================================================================

namespace SciUndo {

    namespace detail {
        /// Thread-local nesting depth counter.
        /// 0 → 1 transition triggers BEGIN, 1 → 0 triggers END.
        inline thread_local int g_nestingDepth = 0;
    }

    /// RAII guard for Scintilla undo actions with automatic nesting support.
    class SciUndoGuard {
        HWND _hSci;
        bool _ownsAction;  // true only if this instance sent SCI_BEGINUNDOACTION

    public:
        /// Construct guard. Only the outermost guard (when nesting depth is 0)
        /// sends SCI_BEGINUNDOACTION to Scintilla.
        explicit SciUndoGuard(HWND hSci) noexcept
            : _hSci(hSci), _ownsAction(false)
        {
            if (!_hSci) return;

            // Only the first guard (depth 0 → 1) actually starts the undo action
            if (detail::g_nestingDepth == 0) {
                // Verify undo collection is enabled before sending
                if (::SendMessage(_hSci, SCI_GETUNDOCOLLECTION, 0, 0)) {
                    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);
                    _ownsAction = true;
                }
            }
            ++detail::g_nestingDepth;
        }

        /// Destructor. Only the guard that owns the action sends SCI_ENDUNDOACTION.
        ~SciUndoGuard() noexcept {
            if (detail::g_nestingDepth > 0) {
                --detail::g_nestingDepth;
            }
            if (_ownsAction && _hSci) {
                ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);
            }
        }

        // Non-copyable, non-movable
        SciUndoGuard(const SciUndoGuard&) = delete;
        SciUndoGuard& operator=(const SciUndoGuard&) = delete;
        SciUndoGuard(SciUndoGuard&&) = delete;
        SciUndoGuard& operator=(SciUndoGuard&&) = delete;

        /// Returns true if this guard instance owns the undo action.
        [[nodiscard]] bool ownsAction() const noexcept { return _ownsAction; }

        /// Returns current nesting depth (for debugging/testing).
        [[nodiscard]] static int nestingDepth() noexcept {
            return detail::g_nestingDepth;
        }
    };

} // namespace SciUndo

// Convenience alias - can be used without namespace qualification
using SciUndoGuard = SciUndo::SciUndoGuard;
