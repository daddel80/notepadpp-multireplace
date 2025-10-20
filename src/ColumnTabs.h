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

    // --- Data model ----------------------------------------------------------

    // One physical line in the column model.
    struct CT_ColumnLineInfo
    {
        int lineLength = 0;                 // length in bytes (without CR/LF)
        std::vector<int> delimiterOffsets;  // byte offsets of each delimiter token

        size_t FieldCount() const noexcept { return delimiterOffsets.size() + 1; }
    };

    // Read-only view of the parsed document range.
    struct CT_ColumnModelView
    {
        std::vector<CT_ColumnLineInfo> Lines;    // optional if getLineInfo is set
        size_t docStartLine = 0;                 // 0-based absolute document line

        bool delimiterIsTab = false;             // true if delimiter is '\t'
        int  delimiterLength = 1;                // bytes per delimiter (>= 1)
        bool collapseTabRuns = true;             // preserved for callers

        // If set, used instead of Lines; index = (line - docStartLine).
        std::function<CT_ColumnLineInfo(size_t)> getLineInfo;
    };

    // Options for destructive alignment (text-changing padding).
    struct CT_AlignOptions
    {
        int  firstLine = 0;                  // inclusive, model-relative
        int  lastLine = -1;                 // -1 => last model line
        int  gapCells = 2;                  // visual gap *in spaces* between columns
        bool spacesOnlyIfTabDelimiter = true;
        bool oneFlowTabOnly = true; // collapse runs to one '\t' and add tabstops
    };

    // --- Indicator (tracks inserted padding) --------------------------------

    // Set/Get the indicator id used to mark inserted padding.
    void CT_SetIndicatorId(int id) noexcept;
    int  CT_GetIndicatorId() noexcept;

    // --- Destructive API (edits text) ---------------------------------------

    // Insert aligned padding (tabs/spaces) according to the model/options.
    bool CT_InsertAlignedPadding(HWND hSci,
        const CT_ColumnModelView& model,
        const CT_AlignOptions& opt);

    // Remove previously inserted aligned padding (by indicator).
    bool CT_RemoveAlignedPadding(HWND hSci);

    // Query whether the buffer currently contains aligned padding owned by us.
    bool CT_HasAlignedPadding(HWND hSci) noexcept;

    bool CT_ApplyNumericPadding(HWND hSci, const CT_ColumnModelView& model, int firstLine, int lastLine);

    // --- Visual API (does not edit text; manages Scintilla tab stops) -------

    // Apply Flow tab stops for [firstLine..lastLine] with a fixed pixel gap.
    // Units: paddingPx in *pixels* (e.g., 12).
    bool CT_ApplyFlowTabStops(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,   // -1 => all model lines
        int paddingPx /*pixels*/);

    // Convenience: apply to all lines in model.
    inline bool CT_ApplyFlowTabStopsAll(HWND hSci,
        const CT_ColumnModelView& model,
        int paddingPx /*pixels*/) {
        return CT_ApplyFlowTabStops(hSci, model, 0, -1, paddingPx);
    }

    // Convenience overload: gap expressed in *spaces*; converts to pixels internally.
    bool CT_ApplyFlowTabStopsSpaces(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,
        int gapSpaces /*spaces*/);

    bool CT_DisableFlowTabStops(HWND hSci, bool restoreManual);

    // Remove only Flow (ETS-owned) tab stops; restores any manual per-line stops.
    inline bool CT_ClearFlowTabStops(HWND hSci) {
        return CT_DisableFlowTabStops(hSci, /*restoreManual=*/true);
    }

    // Remove *all* tab stops in the buffer (manual and Flow).
    bool CT_ClearAllTabStops(HWND hSci);

    // Forget any cached visual state (must be called on buffer/document switch).
    void CT_ResetFlowVisualState() noexcept;

    bool CT_HasFlowTabStops() noexcept;

    // --- Utilities (kept for callers) ---------------------------------------

    // Measure a cell’s *visual* width assuming a given tab width (utility).
    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth);

    // Per-document padding state (O(1) gate to avoid scans on switches)
    void CT_SetDocHasPads(sptr_t docPtr, bool has) noexcept;
    bool CT_GetDocHasPads(sptr_t docPtr) noexcept;

    // Convenience: resolve doc pointer from HWND
    void CT_SetCurDocHasPads(HWND hSci, bool has) noexcept;
    bool CT_GetCurDocHasPads(HWND hSci) noexcept;

    bool CT_CleanupVisuals(HWND hSci) noexcept;
    bool CT_CleanupAllForDoc(HWND hSci) noexcept;

} // namespace ColumnTabs
