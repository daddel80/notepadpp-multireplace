// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
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
#include <vector>
#include <string>
#include <functional>
#include "Scintilla.h"

namespace ColumnTabs
{
    // ========================================================================
    // Data Model
    // ========================================================================

    /// Information about a single line in the column model.
    struct CT_ColumnLineInfo {
        int lineLength = 0;                 ///< Length in bytes (without CR/LF)
        std::vector<int> delimiterOffsets;  ///< Byte offsets of each delimiter

        /// Number of fields (columns) = delimiters + 1
        size_t FieldCount() const noexcept { return delimiterOffsets.size() + 1; }
    };

    /// Read-only view of a document range for columnar operations.
    struct CT_ColumnModelView {
        std::vector<CT_ColumnLineInfo> Lines;  ///< Line data (optional if getLineInfo is set)
        size_t docStartLine = 0;               ///< 0-based absolute document line

        bool delimiterIsTab = false;           ///< true if delimiter is '\t'
        int  delimiterLength = 1;              ///< Bytes per delimiter (>= 1)
        bool collapseTabRuns = true;           ///< Preserved for callers

        /// Optional callback for lazy line access; index = (line - docStartLine).
        std::function<CT_ColumnLineInfo(size_t)> getLineInfo;
    };

    /// Options for text-modifying alignment (padding insertion).
    struct CT_AlignOptions {
        int  firstLine = 0;        ///< Inclusive, model-relative
        int  lastLine = -1;        ///< -1 => last model line
        int  gapCells = 2;         ///< Visual gap in *spaces* between columns
        bool oneFlowTabOnly = true; ///< Collapse runs to one '\t' and add tabstops
    };

    // ========================================================================
    // Indicator Management
    // ========================================================================
    //
    // The indicator marks characters inserted by padding operations (TABs,
    // spaces, decimal points). This allows CT_RemoveAlignedPadding() to
    // precisely remove only artificial padding without affecting user content.
    //
    // Style: INDIC_HIDDEN with alpha=0 (invisible but programmatically detectable)
    // ========================================================================

    void CT_SetIndicatorId(int id) noexcept;
    int  CT_GetIndicatorId() noexcept;

    // ========================================================================
    // Document Padding State
    // ========================================================================
    //
    // Per-document flag indicating whether padding has been inserted.
    // This is an O(1) optimization to avoid scanning for indicators.
    //
    // The flag uses Scintilla's document pointer (SCI_GETDOCPOINTER) as key,
    // correctly tracking state even with multiple views on the same document.
    //
    // IMPORTANT: This flag is the source of truth. It's set by
    // CT_InsertAlignedPadding / CT_ApplyNumericPadding and cleared by
    // CT_RemoveAlignedPadding.
    // ========================================================================

    /// Set padding flag for a specific document pointer.
    void CT_SetDocHasPads(sptr_t docPtr, bool has) noexcept;

    /// Get padding flag for a specific document pointer.
    bool CT_GetDocHasPads(sptr_t docPtr) noexcept;

    /// Set padding flag for the CURRENT document (convenience wrapper).
    void CT_SetCurDocHasPads(HWND hSci, bool has) noexcept;

    /// Get padding flag for the CURRENT document (convenience wrapper).
    /// This is the recommended function to check if padding exists.
    bool CT_GetCurDocHasPads(HWND hSci) noexcept;

    /// Alias for CT_GetCurDocHasPads (for API compatibility).
    /// NOTE: This checks the cached flag, NOT actual document content.
    bool CT_HasAlignedPadding(HWND hSci) noexcept;

    // ========================================================================
    // Destructive API (Modifies Document Text)
    // ========================================================================
    //
    // These functions insert or remove padding characters in the document.
    // All inserted characters are marked with the indicator for later removal.
    //
    // UNDO HANDLING:
    // These functions use SciUndoGuard internally. To combine multiple
    // operations into ONE undo step, wrap them in your own SciUndoGuard:
    //
    //   {
    //       SciUndoGuard undo(hSci);           // Outer guard
    //       CT_ApplyNumericPadding(...);       // Inner guard = no-op
    //       CT_InsertAlignedPadding(...);      // Inner guard = no-op
    //   }
    //   // Result: ONE undo step for everything
    // ========================================================================

    /// Insert TAB padding for column alignment.
    ///
    /// Inserts TAB characters before delimiters to align columns visually.
    /// All inserted TABs are marked with the indicator.
    ///
    /// @param hSci             Scintilla window handle
    /// @param model            Column model describing document structure
    /// @param opt              Alignment options
    /// @param outNothingToAlign  [out] true if no alignment was needed
    /// @return true on success
    bool CT_InsertAlignedPadding(HWND hSci,
        const CT_ColumnModelView& model,
        const CT_AlignOptions& opt,
        bool* outNothingToAlign = nullptr);

    /// Remove ALL padding previously inserted by this module.
    ///
    /// Finds all indicator-marked ranges and removes them, restoring
    /// the document to its original state.
    ///
    /// @param hSci  Scintilla window handle
    /// @return true if padding was removed, false if none existed
    bool CT_RemoveAlignedPadding(HWND hSci);

    /// Insert numeric alignment padding (leading spaces, trailing decimals).
    ///
    /// Example: "42" → "  42.00", "3.14" → "  3.14", "100.5" → "100.50"
    ///
    /// @param hSci      Scintilla window handle
    /// @param model     Column model
    /// @param firstLine First line (doc-absolute)
    /// @param lastLine  Last line (doc-absolute)
    /// @return true on success
    bool CT_ApplyNumericPadding(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine);

    // ========================================================================
    // Visual API (Non-Destructive Tab Stops)
    // ========================================================================
    //
    // These functions set Scintilla's per-line tab stop positions without
    // modifying document text. Creates visual column alignment using
    // existing TAB characters.
    // ========================================================================

    /// Apply calculated tab stops for visual column alignment.
    bool CT_ApplyFlowTabStops(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,   // -1 = all model lines
        int paddingPx);

    /// Convenience: Apply tab stops for entire model.
    inline bool CT_ApplyFlowTabStopsAll(HWND hSci,
        const CT_ColumnModelView& model,
        int paddingPx)
    {
        return CT_ApplyFlowTabStops(hSci, model,
            static_cast<int>(model.docStartLine), -1, paddingPx);
    }

    /// Remove visual tab stops, optionally restoring previous manual stops.
    bool CT_DisableFlowTabStops(HWND hSci, bool restoreManual);

    /// Convenience: Clear Flow tab stops and restore manual ones.
    inline bool CT_ClearFlowTabStops(HWND hSci)
    {
        return CT_DisableFlowTabStops(hSci, /*restoreManual=*/true);
    }

    /// Remove ALL tab stops (manual and Flow).
    bool CT_ClearAllTabStops(HWND hSci);

    /// Reset cached visual state (call on buffer/document switch).
    void CT_ResetFlowVisualState() noexcept;

    /// Check if any line has Flow tab stops.
    bool CT_HasFlowTabStops() noexcept;

    // ========================================================================
    // Utilities
    // ========================================================================

    /// Calculate visual width accounting for tab expansion.
    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth);

    // ========================================================================
    // Cleanup Helpers
    // ========================================================================

    /// Remove visual tab stops only (no text modification).
    bool CT_CleanupVisuals(HWND hSci) noexcept;

    /// Full cleanup: remove padding text AND visual tab stops.
    bool CT_CleanupAllForDoc(HWND hSci) noexcept;

} // namespace ColumnTabs