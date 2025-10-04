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

namespace {
    // Track which lines currently have ETS-owned visual tabstops
    std::vector<uint8_t> g_hasETSLine;                 // 0/1 per line

    // Snapshot of manual tabstops (in px) that existed before ETS took over the line
    std::vector<std::vector<int>> g_savedManualStopsPx;

    inline void ensureCapacity(HWND hSci) {
        const int total = (int)SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
        if ((int)g_hasETSLine.size() < total)          g_hasETSLine.resize((size_t)total, 0);
        if ((int)g_savedManualStopsPx.size() < total)  g_savedManualStopsPx.resize((size_t)total);
    }

    // Collect all current tabstops (px) on a given line (manual or otherwise)
    static std::vector<int> collectTabStopsPx(HWND hSci, int line) {
        std::vector<int> stops;
        int pos = 0;
        for (;;) {
            const int next = (int)SendMessage(hSci, SCI_GETNEXTTABSTOP, (uptr_t)line, (sptr_t)pos);
            if (next <= 0 || next == pos) break;
            stops.push_back(next);
            pos = next;
        }
        return stops;
    }
}

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
    bool ApplyElasticTabStops(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,
        int paddingPx)
    {
        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, firstLine);
        const int line1 = hasVec
            ? ((lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(firstLine, lastLine))
            : lastLine;
        if (line0 > line1) return false;

        const int gapPx = (paddingPx > 0) ? paddingPx : 0;

        // --- ETS ownership: snapshot manual tabstops once per line ---
        ensureCapacity(hSci); // prepares g_hasETSLine / g_savedManualStopsPx
        for (int ln = line0; ln <= line1; ++ln) {
            if ((int)g_hasETSLine.size() > ln && g_hasETSLine[(size_t)ln] == 0u) {
                g_savedManualStopsPx[(size_t)ln] = collectTabStopsPx(hSci, ln);
            }
        }

        // compute + set visual tabstops (unchanged core)
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx))
            return false;

        setTabStopsRangePx(hSci, line0, line1, stops);

        // mark lines as ETS-owned so ClearElasticTabStops can act selectively
        for (int ln = line0; ln <= line1; ++ln)
            g_hasETSLine[(size_t)ln] = 1u;

        return true;
    }

    bool ColumnTabs::ClearElasticTabStops(HWND hSci)
    {
        ensureCapacity(hSci);

        const int total = (int)SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
        const int limit = (std::min)(total, (int)g_hasETSLine.size());

        for (int ln = 0; ln < limit; ++ln) {
            if (!g_hasETSLine[(size_t)ln]) continue;

            // 1) clear all tabstops on this line (only visual editor stops, not text)
            SendMessage(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

            // 2) restore previously saved manual tabstops (if any)
            const auto& manual = g_savedManualStopsPx[(size_t)ln];
            for (int px : manual)
                SendMessage(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)px);

            // 3) reset tracking
            g_hasETSLine[(size_t)ln] = 0u;
            g_savedManualStopsPx[(size_t)ln].clear();
        }
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
        if (line0 > line1) return false;

        // pre-compute editor tabstops (visual aid)
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

        // Clamp helper (prevents Editor.cxx assert)
        auto safeDeleteRange = [&](Sci_Position pos, Sci_Position len) {
            const Sci_Position docLen = (Sci_Position)S(hSci, SCI_GETLENGTH, 0, 0);
            if (pos < 0 || len <= 0) return;
            if (pos > docLen) return;
            if (pos + len > docLen) len = docLen - pos;
            if (len > 0) {
                S(hSci, SCI_DELETERANGE, pos, len);
                S(hSci, SCI_INDICATORCLEARRANGE, pos, len);
            }
            };

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            Sci_Position delta = 0;

            const size_t nDelims = L.delimiterOffsets.size();
            if (nDelims == 0) continue;

            for (size_t c = 0; c < nDelims; ++c) {
                // dynamic line end (runtime, not model)
                const Sci_Position lineEndNow = (Sci_Position)S(hSci, SCI_GETLINEENDPOSITION, ln, 0);

                // compute delimiter position in current text
                const Sci_Position delimPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                if (delimPos < base || delimPos > lineEndNow) {
                    // stale offset or past EOL; skip this delimiter safely
                    continue;
                }

                // scan whitespace run BEFORE delimiter
                Sci_Position wsStart = delimPos;
                while (wsStart > base) {
                    const int ch = (int)S(hSci, SCI_GETCHARAT, (uptr_t)(wsStart - 1), 0);
                    if (ch == ' ' || ch == '\t') --wsStart; else break;
                }

                // existing '\t' exactly before delimiter?
                const bool keepExistingTab =
                    (wsStart < delimPos) &&
                    ((int)S(hSci, SCI_GETCHARAT, (uptr_t)(delimPos - 1), 0) == '\t');

                // trim BEFORE delimiter
                if (wsStart < delimPos) {
                    if (keepExistingTab) {
                        const Sci_Position lastTab = delimPos - 1;

                        if (lastTab > wsStart) {
                            const Sci_Position len1 = lastTab - wsStart;
                            safeDeleteRange(wsStart, len1);
                            delta -= len1;
                        }
                        const Sci_Position afterTab = lastTab + 1;
                        if (afterTab < delimPos) {
                            const Sci_Position len2 = delimPos - afterTab;
                            safeDeleteRange(afterTab, len2);
                            delta -= len2;
                        }
                    }
                    else {
                        const Sci_Position wsLen = delimPos - wsStart;
                        safeDeleteRange(wsStart, wsLen);
                        delta -= wsLen;
                    }
                }

                // ensure exactly one tab right before the delimiter
                Sci_Position tabPos = 0;
                if (keepExistingTab) {
                    tabPos = delimPos - 1; // reuse manual tab; do NOT mark it
                }
                else {
                    // insert new elastic tab at delimiter start and mark it
                    tabPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                    S(hSci, SCI_INSERTTEXT, tabPos, (sptr_t)"\t");
                    S(hSci, SCI_INDICATORFILLRANGE, tabPos, 1);
                    delta += 1;
                }

                // --- trim AFTER delimiter (start of next field) ---
                // Correct start-of-trim:
                //   if we reused a manual tab:  afterDelim = delimPos + delimiterLength
                //   if we inserted a tab:       afterDelim = delimPos + 1 /*tab*/ + delimiterLength
                Sci_Position afterDelim = keepExistingTab
                    ? (delimPos + (Sci_Position)model.delimiterLength)
                    : (delimPos + 1 + (Sci_Position)model.delimiterLength);

                // guard: never run past current doc end
                for (;;) {
                    const Sci_Position docLen = (Sci_Position)S(hSci, SCI_GETLENGTH, 0, 0);
                    if (afterDelim >= docLen) break;

                    const int ch = (int)S(hSci, SCI_GETCHARAT, (uptr_t)afterDelim, 0);
                    if (ch == ' ' || ch == '\t') {
                        safeDeleteRange(afterDelim, 1);
                        // no delta update (we delete after the boundary)
                    }
                    else {
                        break;
                    }
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
