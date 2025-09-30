// ColumnTabs.cpp
#include "ColumnTabs.h"
#include <algorithm>

#ifndef NOMINMAX
#define NOMINMAX
#endif

// --- direct scintilla call ---
static inline sptr_t S(HWND hSci, UINT m, uptr_t w = 0, sptr_t l = 0) {
    auto fn = reinterpret_cast<SciFnDirect>(::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
    auto ptr = (sptr_t)       ::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
}

// redraw guard (UI batching)
struct RedrawGuard {
    HWND h;
    explicit RedrawGuard(HWND hwnd) : h(hwnd) { ::SendMessage(h, WM_SETREDRAW, FALSE, 0); }
    ~RedrawGuard() { ::SendMessage(h, WM_SETREDRAW, TRUE, 0); ::InvalidateRect(h, nullptr, TRUE); }
};

// -----------------------------------------------------------------------------
//               implementation inside namespace ColumnTabs
// -----------------------------------------------------------------------------
namespace ColumnTabs {

    static int g_indic = 8;

    void CT_SetIndicatorId(int id) noexcept { g_indic = id; }
    int  CT_GetIndicatorId() noexcept { return g_indic; }

    static inline int getTabWidth(HWND hSci) {
        const int w = (int)S(hSci, SCI_GETTABWIDTH);
        return w > 0 ? w : 8;
    }

    // Count visual cells (monospace). '\t' jumps to next tab stop; CR/LF ignored.
    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth) {
        size_t col = 0;
        for (size_t i = 0; i < n; ++i) {
            const unsigned char c = (unsigned char)s[i];
            if (c == '\t') {
                if (tabWidth <= 1) { ++col; continue; }
                const size_t mod = col % (size_t)tabWidth;
                col += (mod == 0) ? (size_t)tabWidth : ((size_t)tabWidth - mod);
            }
            else if (c != '\r' && c != '\n') {
                ++col;
            }
        }
        return col;
    }

    // Build shortest mix of tabs+spaces to reach targetCol from currentCol (cells).
    static inline void makePadding(size_t curCol, size_t tgtCol, int tabWidth, bool allowTabs, std::string& out) {
        out.clear();
        if (tgtCol <= curCol) return;
        if (!allowTabs || tabWidth <= 1) { out.assign(tgtCol - curCol, ' '); return; }
        size_t col = curCol;
        while (true) {
            const size_t next = ((col / (size_t)tabWidth) + 1) * (size_t)tabWidth;
            if (next > tgtCol) break;
            out.push_back('\t'); col = next;
        }
        if (tgtCol > col) out.append(tgtCol - col, ' ');
    }

    // ---------------- non-destructive: editor tab stops --------------------------

    bool ClearTabStops(HWND hSci) {
        S(hSci, SCI_CLEARTABSTOPS);
        return true;
    }

    // Pixel-accurate editor-global tab stops (minimal variant; safe default).
    bool ApplyElasticTabStops(HWND hSci, const CT_ColumnModelView& model,
        int firstLine, int lastLine, int paddingPx)
    {
        if (model.Lines.empty() && !model.getLineInfo) return false;

        const int total = (int)(model.getLineInfo ? (lastLine >= 0 ? (lastLine - firstLine + 1) : 0)
            : model.Lines.size());
        if (total == 0) return false;

        // Determine range
        const int line0 = (std::max)(0, firstLine);
        const int line1 = (lastLine < 0) ? (line0 + (int)model.Lines.size() - 1) : (std::max)(firstLine, lastLine);
        (void)line0; (void)line1;

        // Reset and keep base width
        S(hSci, SCI_CLEARTABSTOPS);
        const int baseTabW = (int)S(hSci, SCI_GETTABWIDTH);
        const int padPx = (paddingPx > 0 ? paddingPx : 8);
        S(hSci, SCI_SETTABWIDTH, (baseTabW > 0 ? baseTabW : 4));
        (void)padPx; // placeholder for future per-column pixel stops
        return true;
    }

    // ---------------- destructive: aligned padding (tabs+spaces) -----------------

   // Compute per-column maximum field widths (cells), WITHOUT gap.
// Handles multi-char delimiters via model.delimiterLength.
    static bool computeTargets(HWND hSci, const ColumnTabs::CT_ColumnModelView& model,
        int line0, int line1, int /*gapCells*/, int tabWidth,
        std::vector<size_t>& maxWidth /*out*/)
    {
        size_t maxCols = 0;
        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = model.getLineInfo
                ? model.getLineInfo((size_t)(ln - model.docStartLine))
                : model.Lines[(size_t)(ln - model.docStartLine)];
            maxCols = (std::max)(maxCols, L.FieldCount());
        }
        maxWidth.assign(maxCols, 0);

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = model.getLineInfo
                ? model.getLineInfo((size_t)(ln - model.docStartLine))
                : model.Lines[(size_t)(ln - model.docStartLine)];

            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            const size_t nDelims = L.delimiterOffsets.size();
            const size_t nFields = nDelims + 1;

            for (size_t col = 0; col < nFields; ++col) {
                // [s..e) field bytes
                Sci_Position s = base;
                if (col > 0) {
                    const int prevDel = L.delimiterOffsets[col - 1];
                    s = base + (Sci_Position)prevDel + (Sci_Position)model.delimiterLength;
                }
                Sci_Position e = base + (Sci_Position)L.lineLength;
                if (col < nDelims) {
                    const int thisDel = L.delimiterOffsets[col];
                    e = base + (Sci_Position)thisDel;
                }
                if (e <= s) continue;

                std::string buf; buf.resize((size_t)(e - s));
                Sci_TextRangeFull tr{ s, e, buf.data() };
                S(hSci, SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);

                const size_t w = ColumnTabs::CT_VisualCellWidth(buf.data(), buf.size(), tabWidth);
                maxWidth[col] = (std::max)(maxWidth[col], w);
            }
        }
        return true;
    }



    bool ColumnTabs::CT_InsertAlignedPadding(HWND hSci,
        const CT_ColumnModelView& model,
        const CT_AlignOptions& opt)
    {
        // basic guards
        const bool hasPrefilled = !model.Lines.empty();
        if (!hasPrefilled && !model.getLineInfo) return false;

        // range
        const int line0 = (std::max)(0, opt.firstLine);
        const int line1 = hasPrefilled
            ? ((opt.lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(opt.firstLine, opt.lastLine))
            : opt.lastLine; // when using callback you must pass a bounded range

        if (hasPrefilled && (line1 < line0 || model.Lines.empty())) return false;

        const int tabW = getTabWidth(hSci);

        // 1) max field widths per column (cells), WITHOUT gap
        std::vector<size_t> targets; // per-column max widths (cells)
        if (!computeTargets(hSci, model, line0, line1, /*gapCells*/0, tabW, targets))
            return false;

        // 2) absolute start of each column (cells), global per document
        //    start[0] = 0; start[c] = start[c-1] + targets[c-1] + gap
        std::vector<size_t> colStart;
        colStart.resize(targets.size());
        if (!targets.empty()) {
            colStart[0] = 0;
            for (size_t c = 1; c < targets.size(); ++c)
                colStart[c] = colStart[c - 1] + targets[c - 1] + (size_t)opt.gapCells;
        }

        const bool allowTabs = !(opt.spacesOnlyIfTabDelimiter && model.delimiterIsTab);

        RedrawGuard rg(hSci);
        S(hSci, SCI_BEGINUNDOACTION);
        S(hSci, SCI_SETINDICATORCURRENT, g_indic);
        S(hSci, SCI_INDICSETSTYLE, g_indic, INDIC_HIDDEN);
        S(hSci, SCI_INDICSETALPHA, g_indic, 0);

        // 3) per-line insert pass (top -> bottom)
        for (int ln = line0; ln <= line1; ++ln) {
            // get line info
            const CT_ColumnLineInfo L = model.getLineInfo
                ? model.getLineInfo((size_t)(ln - model.docStartLine))
                : model.Lines[(size_t)(ln - model.docStartLine)];

            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            Sci_Position delta = 0;

            const size_t nDelims = L.delimiterOffsets.size();
            const size_t nFields = nDelims + 1;

            for (size_t col = 0; col < nFields; ++col) {
                if (col >= nDelims) break; // last field has no trailing delimiter

                // delimiter byte-offset in line (start of token)
                const int thisDelOff = L.delimiterOffsets[col];

                // absolute insert position (before delimiter), account inserts in this line
                const Sci_Position d = base + (Sci_Position)thisDelOff + delta;

                // measure ABSOLUTE current delimiter column (cells from line start)
                size_t currentAbs = 0;
                if (d > base) {
                    std::string head; head.resize((size_t)(d - base));
                    Sci_TextRangeFull trHead{ base, d, head.data() };
                    S(hSci, SCI_GETTEXTRANGEFULL, 0, (sptr_t)&trHead);
                    currentAbs = CT_VisualCellWidth(head.data(), head.size(), tabW);
                }

                // ABSOLUTE target delimiter column: start[col] + targets[col] + gap
                size_t targetAbs = currentAbs;
                if (col < targets.size()) {
                    const size_t startCol = (col < colStart.size()) ? colStart[col] : 0;
                    targetAbs = startCol + targets[col] + (size_t)opt.gapCells;
                }

                if (currentAbs >= targetAbs) continue;

                std::string pad;
                makePadding(/*current*/currentAbs, /*target*/targetAbs, tabW, allowTabs, pad);
                if (pad.empty()) continue;

                S(hSci, SCI_INSERTTEXT, d, (sptr_t)pad.c_str());
                S(hSci, SCI_INDICATORFILLRANGE, d, (sptr_t)pad.size());
                delta += (Sci_Position)pad.size();
            }
        }

        S(hSci, SCI_ENDUNDOACTION);
        return true;
    }


    bool CT_RemoveAlignedPadding(HWND hSci)
    {
        RedrawGuard rg(hSci);
        S(hSci, SCI_BEGINUNDOACTION);
        S(hSci, SCI_SETINDICATORCURRENT, g_indic);

        const Sci_Position len = S(hSci, SCI_GETLENGTH);
        std::vector<std::pair<Sci_Position, Sci_Position>> spans;
        spans.reserve(1024);

        // Collect all indicator ranges left->right using VALUEAT
        Sci_Position pos = 0;
        while (pos < len) {
            const int v = (int)S(hSci, SCI_INDICATORVALUEAT, g_indic, pos);
            if (v != 0) {
                const Sci_Position start = pos;
                // advance until indicator ends
                while (pos < len && (int)S(hSci, SCI_INDICATORVALUEAT, g_indic, pos) != 0)
                    ++pos;
                const Sci_Position end = pos;
                if (end > start) spans.emplace_back(start, end);
            }
            else {
                ++pos;
            }
        }

        // Delete back->front to keep offsets stable
        for (auto it = spans.rbegin(); it != spans.rend(); ++it) {
            const Sci_Position start = it->first;
            const Sci_Position length = it->second - it->first;
            S(hSci, SCI_INDICATORCLEARRANGE, start, length);
            S(hSci, SCI_DELETERANGE, start, length);
        }

        S(hSci, SCI_ENDUNDOACTION);
        return true;
    }


    // ---------------- QS ----------------

    bool CT_QS_SelfTest()
    {
        if (!(CT_VisualCellWidth("a", 1, 4) == 1)) return false;
        if (!(CT_VisualCellWidth("\t", 1, 4) == 4)) return false;
        if (!(CT_VisualCellWidth("ab", 2, 4) == 2)) return false;
        if (!(CT_VisualCellWidth("ab\t", 3, 4) == 4)) return false;
        if (!(CT_VisualCellWidth("a\tb", 3, 4) == 5)) return false;
        return true;
    }

    bool CT_QS_AlignedOnOff(HWND hSci, const CT_ColumnModelView& model)
    {
        CT_RemoveAlignedPadding(hSci);

        CT_AlignOptions opt{};
        opt.firstLine = 0;
        opt.lastLine = (int)model.Lines.size() - 1;
        opt.gapCells = 2;
        opt.spacesOnlyIfTabDelimiter = true;

        const bool on = CT_InsertAlignedPadding(hSci, model, opt);
        S(hSci, SCI_SETINDICATORCURRENT, g_indic);
        const bool has = (S(hSci, SCI_INDICATORSTART, g_indic, S(hSci, SCI_GETLENGTH)) >= 0);

        const bool off = CT_RemoveAlignedPadding(hSci);
        const bool clr = (S(hSci, SCI_INDICATORSTART, g_indic, S(hSci, SCI_GETLENGTH)) >= 0);

        return on && has && off && !clr;
    }

} // namespace ColumnTabs
