#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <vector>
#include <functional>
#include <string>
#include "Scintilla.h"

#ifdef max
#undef max
#endif
#ifdef min
#undef min
#endif

namespace ColumnTabs
{
#ifndef CT_TABSTOPS_PIXELS
#define CT_TABSTOPS_PIXELS 1
#endif

    // One line in the CSV model.
    struct CT_ColumnLineInfo {
        int lineLength = 0;                 // bytes, without CR/LF
        std::vector<int> delimiterOffsets;  // byte offsets of each delimiter
        size_t FieldCount() const noexcept { return delimiterOffsets.size() + 1; }
    };

    // Read-only view of the parsed CSV.
    struct CT_ColumnModelView {
        std::vector<CT_ColumnLineInfo> Lines;   // optional (if getLineInfo is set)
        size_t docStartLine = 0;                // 0-based

        bool delimiterIsTab = false;            // true if '\t'
        int  delimiterLength = 1;               // bytes per delimiter token (>=1)
        bool collapseTabRuns = true;            // kept for callers

        // If set, used instead of Lines (index = line - docStartLine).
        std::function<CT_ColumnLineInfo(size_t)> getLineInfo;
    };

    // Options for padding/alignment.
    struct CT_AlignOptions {
        int  firstLine = 0;                 // inclusive
        int  lastLine = -1;                // -1 => to last model line
        int  gapCells = 2;                 // visual gap after a column (in "space" units)
        bool spacesOnlyIfTabDelimiter = true;
        bool oneElasticTabOnly = true;      // replace run by one '\t' and set tabstops
    };

    // Indicator id used to mark inserted tabs.
    void   CT_SetIndicatorId(int id) noexcept;
    int    CT_GetIndicatorId() noexcept;

    // Utility (kept for callers).
    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth);

    // Core API.
    bool   CT_InsertAlignedPadding(HWND hSci, const CT_ColumnModelView& model, const CT_AlignOptions& opt);
    bool   CT_RemoveAlignedPadding(HWND hSci);

    bool   ApplyElasticTabStops(HWND hSci, const CT_ColumnModelView& model,
        int firstLine, int lastLine, int paddingPx /*pixels*/);

    bool   ClearTabStops(HWND hSci);
    bool   CT_HasAlignedPadding(HWND hSci) noexcept;

} // namespace ColumnTabs
