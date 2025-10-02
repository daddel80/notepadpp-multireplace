#include "ColumnTabs.h"
#include <algorithm>

#ifndef NOMINMAX
#define NOMINMAX
#endif

// --- Scintilla direct --------------------------------------------------------
static inline sptr_t S(HWND hSci, UINT m, uptr_t w = 0, sptr_t l = 0)
{
    auto fn = reinterpret_cast<SciFnDirect>(::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
    auto ptr = (sptr_t)       ::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
}

struct RedrawGuard {
    HWND h{};
    explicit RedrawGuard(HWND hwnd) : h(hwnd) { ::SendMessage(h, WM_SETREDRAW, FALSE, 0); }
    ~RedrawGuard() { ::SendMessage(h, WM_SETREDRAW, TRUE, 0); ::InvalidateRect(h, nullptr, TRUE); }
};

namespace ColumnTabs
{

    // --- Globals -----------------------------------------------------------------
    static int g_CT_IndicatorId = 30;

    void CT_SetIndicatorId(int id) noexcept { g_CT_IndicatorId = id; }
    int  CT_GetIndicatorId() noexcept { return g_CT_IndicatorId; }

    // --- Helpers -----------------------------------------------------------------
    static inline int pxOfSpace(HWND hSci)
    {
        const int px = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        return (px > 0) ? px : 8;
    }

    size_t CT_VisualCellWidth(const char* s, size_t n, int tabWidth)
    {
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

    static inline CT_ColumnLineInfo fetchLine(const CT_ColumnModelView& model, int line)
    {
        const size_t idx = (size_t)(line - (int)model.docStartLine);
        return model.getLineInfo ? model.getLineInfo(idx) : model.Lines[idx];
    }

    // Compute pixel tabstops (N fields -> N-1 stops) with symmetric padding around delimiters,
 // using your cumulative rule PLUS an exact EOL clamp that is computed with the SAME layout
 // model as the preferred stops. This guarantees that if any line ends at column m (no m+1),
 // all stops from m onward are >= that line's EOL, so the last delimiters of other lines
 // align flush with that EOL.
 //
 // Preferred rule (your formula):
 //   STOP_pref(c) = sum_{k=0..c}   (maxCellWidthPx[k]       + gapBeforePx)
 //                + sum_{k=0..c-1} (maxDelimiterWidthPx[k] + gapAfterPx)
 //
 // Our addition (exact EOL clamp):
 //   For each line with last column index L (i.e., FieldCount = L+1):
 //     EOL_X(line) = sum_{k=0..L-1} (cellW_line[k] + gapBeforePx + delimW_line[k] + gapAfterPx)
 //                 + cellW_line[L]         // no gap/delimiter after the last cell
 //   For every c >= L, eolClamp[c] = max(eolClamp[c], EOL_X(line)).
 //
 // Final stop:
 //   STOP(c) = max(STOP_pref(c), eolClamp(c))   // no artificial +1; monotonic non-decreasing.
    static bool computeStopsFromWidthsPx(HWND hSci,
        const ColumnTabs::CT_ColumnModelView& model,
        int line0, int line1,
        std::vector<int>& stopPx /*out*/,
        int gapPx)
    {
        using namespace ColumnTabs;

        if (line1 < line0) std::swap(line0, line1);

        // Symmetric gaps (you can make them asymmetric if needed)
        const int gapBeforePx = (gapPx >= 0) ? gapPx : 0; // left padding before a visible delimiter
        const int gapAfterPx = (gapPx >= 0) ? gapPx : 0; // right padding after the delimiter

        // --- determine maximum number of fields in range
        size_t maxCols = 0;
        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            maxCols = (std::max)(maxCols, L.FieldCount()); // FieldCount = nDelims + 1
        }
        if (maxCols < 2) { stopPx.clear(); return true; }

        const size_t stopsCount = maxCols - 1;

        // We do NOT rely on current tabstops for measuring; we measure text widths explicitly.
        // However, to avoid per-line stop contamination elsewhere, clear stops in range.
        for (int ln = line0; ln <= line1; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

        // ---------- helpers to fetch substrings (Scintilla 5 API) and measure pixel width ----------
        auto getSubText = [&](Sci_Position pos0, Sci_Position pos1, std::string& out) {
            if (pos1 <= pos0) { out.clear(); return; }
            const ptrdiff_t need = (ptrdiff_t)(pos1 - pos0);
            out.assign((size_t)need + 1u, '\0');

            Sci_TextRangeFull tr{};             // requires Scintilla.h with full API
            tr.chrg.cpMin = (Sci_PositionCR)pos0;
            tr.chrg.cpMax = (Sci_PositionCR)pos1;
            tr.lpstrText = out.data();
            S(hSci, SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);

            out.resize(std::strlen(out.c_str()));
            };

        auto measurePx = [&](const std::string& s) -> int {
            if (s.empty()) return 0;
            return (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)s.c_str());
            };

        // ---------- PASS 1: collect maxima for preferred stops AND per-line EOL clamp ----------
        std::vector<int> maxCellWidthPx(maxCols, 0);         // per column k
        std::vector<int> maxDelimiterWidthPx(stopsCount, 0); // per delimiter k (between k and k+1)
        std::vector<int> eolClamp(stopsCount, 0);            // cumulative EOL requirement per stop

        // Temporary buffers per line to compute its EOL in *our layout model*
        std::vector<int> cellW_line;    cellW_line.reserve(maxCols);
        std::vector<int> delimW_line;   delimW_line.reserve(stopsCount);

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);

            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            const size_t nDelims = L.delimiterOffsets.size();
            const size_t nFields = nDelims + 1;

            cellW_line.assign(nFields, 0);
            delimW_line.assign(nDelims, 0);

            // --- measure each cell's pixel width on this line
            for (size_t k = 0; k < nFields; ++k) {
                Sci_Position s = base;
                if (k > 0)
                    s = base + (Sci_Position)L.delimiterOffsets[k - 1]
                    + (Sci_Position)model.delimiterLength;

                Sci_Position e = (k < nDelims)
                    ? (base + (Sci_Position)L.delimiterOffsets[k])
                    : (base + (Sci_Position)L.lineLength);

                std::string cell;
                getSubText(s, e, cell);
                const int w = measurePx(cell);
                cellW_line[k] = w;
                if (w > maxCellWidthPx[k]) maxCellWidthPx[k] = w;
            }

            // --- measure each delimiter width on this line (if not a tab)
            if (!model.delimiterIsTab && model.delimiterLength > 0) {
                for (size_t d = 0; d < nDelims && d < stopsCount; ++d) {
                    Sci_Position d0 = base + (Sci_Position)L.delimiterOffsets[d];
                    Sci_Position d1 = d0 + (Sci_Position)model.delimiterLength;
                    std::string del;
                    getSubText(d0, d1, del);
                    const int dw = measurePx(del);
                    delimW_line[d] = dw;
                    if (dw > maxDelimiterWidthPx[d]) maxDelimiterWidthPx[d] = dw;
                }
            }
            else {
                // delimiter is a TAB -> visual width is governed by stops; treat as 0
                std::fill(delimW_line.begin(), delimW_line.end(), 0);
            }

            // --- compute this line's EOL in our layout model (exact)
            const int lastIdx = (int)nFields - 1;
            int eolX = 0;
            // all complete boundaries before the last column:
            for (int k = 0; k < lastIdx; ++k) {
                eolX += cellW_line[(size_t)k] + gapBeforePx;
                eolX += delimW_line[(size_t)k] + gapAfterPx;
            }
            // finally add only the last cell's width (no gap/delimiter after it)
            eolX += cellW_line[(size_t)lastIdx];

            // this EOL clamps all subsequent stops c >= lastIdx
            for (size_t c = (size_t)lastIdx; c < stopsCount; ++c)
                if (eolX > eolClamp[c]) eolClamp[c] = eolX;
        }

        // ---------- PASS 2: build preferred stops from maxima (your rule) ----------
        std::vector<int> stopPref(stopsCount, 0);
        {
            int acc = 0;
            for (size_t c = 0; c < stopsCount; ++c) {
                // add full previous boundaries
                acc += maxCellWidthPx[c] + gapBeforePx;
                stopPref[c] = acc; // delimiter c starts here
                // carry delimiter width and right gap to the next boundary
                acc += maxDelimiterWidthPx[c] + gapAfterPx;
            }
        }

        // ---------- Final stops: take the stronger of preferred vs. cumulative EOL clamp ----------
        stopPx.assign(stopsCount, 0);
        for (size_t c = 0; c < stopsCount; ++c) {
            int target = (stopPref[c] < eolClamp[c]) ? eolClamp[c] : stopPref[c];
            // enforce non-decreasing without artificial +1
            if (c > 0 && target < stopPx[c - 1])
                target = stopPx[c - 1];
            stopPx[c] = target;
        }

        return true;
    }


    // Tabstops (Pixel) pro Zeile setzen.
    static void setTabStopsRangePx(HWND hSci, int line0, int line1, const std::vector<int>& stops)
    {
        for (int ln = line0; ln <= line1; ++ln) {
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
            for (size_t i = 0; i < stops.size(); ++i)
                S(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)stops[i]);
        }
    }

    // --- Public: apply tabstops only --------------------------------------------
    bool ApplyElasticTabStops(HWND hSci, const CT_ColumnModelView& model,
        int firstLine, int lastLine, int paddingPx)
    {
        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, firstLine);
        const int line1 = hasVec
            ? ((lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(firstLine, lastLine))
            : lastLine;

        const int gapPx = (paddingPx > 0) ? paddingPx : 0;

        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx)) return false;
        setTabStopsRangePx(hSci, line0, line1, stops);
        return true;
    }

    // --- Public: destructive insert (one '\t' per delimiter) ---------------------
    bool CT_InsertAlignedPadding(HWND hSci, const CT_ColumnModelView& model, const CT_AlignOptions& opt)
    {
        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, opt.firstLine);
        const int line1 = hasVec
            ? ((opt.lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(opt.firstLine, opt.lastLine))
            : opt.lastLine;

        // Tabstops vorab setzen
        const int gapPx = (opt.gapCells > 0) ? (pxOfSpace(hSci) * opt.gapCells) : 0;
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx)) return false;
        setTabStopsRangePx(hSci, line0, line1, stops);

        if (opt.oneElasticTabOnly && model.delimiterIsTab)
            return true;

        RedrawGuard rd(hSci);
        S(hSci, SCI_BEGINUNDOACTION);
        S(hSci, SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
        S(hSci, SCI_INDICSETSTYLE, g_CT_IndicatorId, INDIC_HIDDEN);
        S(hSci, SCI_INDICSETALPHA, g_CT_IndicatorId, 0);

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            Sci_Position delta = 0;

            const size_t nDelims = L.delimiterOffsets.size();
            for (size_t c = 0; c < nDelims; ++c) {
                const Sci_Position delimPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;

                // trim before delimiter
                Sci_Position wsStart = delimPos;
                while (wsStart > base) {
                    const int ch = (int)S(hSci, SCI_GETCHARAT, (uptr_t)(wsStart - 1), 0);
                    if (ch == ' ' || ch == '\t') --wsStart; else break;
                }
                const Sci_Position wsLen = delimPos - wsStart;
                if (wsLen > 0) {
                    S(hSci, SCI_DELETERANGE, wsStart, wsLen);
                    S(hSci, SCI_INDICATORCLEARRANGE, wsStart, wsLen);
                    delta -= wsLen;
                }

                // insert one tab and tag it
                const Sci_Position tabPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                S(hSci, SCI_INSERTTEXT, tabPos, (sptr_t)"\t");
                S(hSci, SCI_INDICATORFILLRANGE, tabPos, 1);
                delta += 1;

                // trim after delimiter (start of next field)
                Sci_Position afterDelim = tabPos + (Sci_Position)model.delimiterLength + 1;
                while (true) {
                    const int ch = (int)S(hSci, SCI_GETCHARAT, (uptr_t)afterDelim, 0);
                    if (ch == ' ' || ch == '\t') {
                        S(hSci, SCI_DELETERANGE, afterDelim, 1);
                        S(hSci, SCI_INDICATORCLEARRANGE, afterDelim, 1);
                    }
                    else break;
                }
            }
        }

        S(hSci, SCI_ENDUNDOACTION);
        return true;
    }

    // --- Public: remove all tagged tabs & clear tabstops -------------------------
    bool CT_RemoveAlignedPadding(HWND hSci)
    {
        RedrawGuard rd(hSci);
        S(hSci, SCI_BEGINUNDOACTION);
        S(hSci, SCI_SETINDICATORCURRENT, g_CT_IndicatorId);

        const Sci_Position len = S(hSci, SCI_GETLENGTH);
        std::vector<std::pair<Sci_Position, Sci_Position>> spans;
        spans.reserve(1024);

        Sci_Position pos = 0;
        while (pos < len) {
            if ((int)S(hSci, SCI_INDICATORVALUEAT, g_CT_IndicatorId, pos) != 0) {
                const Sci_Position start = pos;
                while (pos < len && (int)S(hSci, SCI_INDICATORVALUEAT, g_CT_IndicatorId, pos) != 0) ++pos;
                spans.emplace_back(start, pos);
            }
            else ++pos;
        }
        for (auto it = spans.rbegin(); it != spans.rend(); ++it) {
            S(hSci, SCI_INDICATORCLEARRANGE, it->first, it->second - it->first);
            S(hSci, SCI_DELETERANGE, it->first, it->second - it->first);
        }

        const int total = (int)S(hSci, SCI_GETLINECOUNT);
        for (int ln = 0; ln < total; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

        S(hSci, SCI_ENDUNDOACTION);
        return true;
    }

    bool ClearTabStops(HWND hSci)
    {
        const int total = (int)S(hSci, SCI_GETLINECOUNT);
        for (int ln = 0; ln < total; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
        return true;
    }

    bool CT_HasAlignedPadding(HWND hSci) noexcept
    {
        S(hSci, SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
        const Sci_Position any = S(hSci, SCI_INDICATORSTART, g_CT_IndicatorId, S(hSci, SCI_GETLENGTH));
        return (any >= 0);
    }

} // namespace ColumnTabs
