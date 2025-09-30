// ColumnTabs.h
// MultiReplace – Elastic Column Alignment (tabstops + aligned padding)
// - Non-destructive (editor tab stops): ApplyElasticTabStops / ClearTabStops
// - Destructive (tabs+spaces padding, removable): CT_InsertAlignedPadding / CT_RemoveAlignedPadding

#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <string>
#include <vector>
#include <cstdint>
#include "Scintilla.h"

namespace ColumnTabs {

    // --- CSV model used by this lib (kept minimal; keep in sync with your matrix) ---
    struct CT_ColumnLineInfo {
        int lineLength = 0;                    // absolute line length (bytes)
        std::vector<int> delimiterOffsets;     // delimiter byte offsets (line-relative)
        std::vector<int> delimiterRunEnds;     // optional (TAB runs), exclusive end

        size_t FieldCount() const { return delimiterOffsets.size() + 1; }

        // line-relative byte offsets
        int FieldStartInLine(size_t col) const {
            return (col == 0) ? 0 : (delimiterOffsets[col - 1] + 1);
        }
        int FieldEndInLine(size_t col) const {
            return (col + 1 < FieldCount()) ? delimiterOffsets[col] : lineLength;
        }
        int DelimiterPosInLine(size_t col) const { return delimiterOffsets[col]; }
    };

    struct CT_ColumnModelView {
        int  docStartLine = 0;
        bool delimiterIsTab = false;
        int  delimiterLength = 1;
        bool collapseTabRuns = true;

        std::vector<CT_ColumnLineInfo> Lines;

        // optional callback if you don’t prefill Lines
        CT_ColumnLineInfo(*getLineInfo)(size_t idx) = nullptr;
    };

    // ---------------- Indicator (destructive mode) ----------------
    void CT_SetIndicatorId(int id) noexcept;   // default 8
    int  CT_GetIndicatorId() noexcept;

    // ---------------- Non-destructive (editor tab stops) ----------
    bool ApplyElasticTabStops(HWND hSci, const CT_ColumnModelView& model,
        int firstLine, int lastLine, int paddingPx);
    bool ClearTabStops(HWND hSci);

    // ---------------- Destructive (aligned padding, removable) ----
    struct CT_AlignOptions {
        int  firstLine = 0;            // inclusive
        int  lastLine = -1;           // -1 => to last
        int  gapCells = 2;            // min spacing after column
        bool spacesOnlyIfTabDelimiter = true;
    };

    bool CT_InsertAlignedPadding(HWND hSci, const CT_ColumnModelView& model, const CT_AlignOptions& opt);
    bool CT_RemoveAlignedPadding(HWND hSci);

    // helpers (QS/unit)
    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth);
    inline size_t CT_VisualCellWidth(const std::string& s, int tabWidth) {
        return CT_VisualCellWidth(s.data(), s.size(), tabWidth);
    }

    // sanity tests
    bool CT_QS_SelfTest();
    bool CT_QS_AlignedOnOff(HWND hSci, const CT_ColumnModelView& model);

} // namespace ColumnTabs
