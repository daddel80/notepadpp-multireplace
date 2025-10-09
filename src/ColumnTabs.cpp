#include "ColumnTabs.h"
#include <algorithm>

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

namespace ColumnTabs::detail {

    // Tracks which lines currently have ETS-owned visual tab stops.
    static std::vector<uint8_t> g_hasETSLine;                 // 0/1 per line

    // Snapshot of manual tab stops (in px) that existed before ETS took over the line.
    static std::vector<std::vector<int>> g_savedManualStopsPx;

    // Ensure our global tracking vectors have capacity for the current buffer.
    inline void ensureCapacity(HWND hSci) {
        const int total = (int)SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);

        if ((int)g_hasETSLine.size() != total)
            g_hasETSLine.assign((size_t)total, 0u);

        if ((int)g_savedManualStopsPx.size() != total)
            g_savedManualStopsPx.assign((size_t)total, std::vector<int>{});
    }

    // Collect all current tab stops (px) on a given line (manual or otherwise).
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

    // Pixel width of a space in the current style; stable fallback if renderer reports 0.
    static inline int pxOfSpace(HWND hSci) {
        const int px = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        return (px > 0) ? px : 8;
    }

    // Fetch a model line by absolute document line number.
    static inline CT_ColumnLineInfo fetchLine(const CT_ColumnModelView& model, int line) {
        const size_t idx = (size_t)(line - (int)model.docStartLine);
        return model.getLineInfo ? model.getLineInfo(idx) : model.Lines[idx];
    }

    // Compute pixel tab stops (N fields -> N-1 stops) with uniform, fixed spacing.
 // Guarantees:
 //  • Column width = maximum cell width (px) per column across [line0..line1].
 //  • Visual spacing is constant: fixed LEFT gap (gapPx) before each delimiter; RIGHT gap = 0.
 //  • EOL clamp is capped by preferred stops (cannot push a stop further right).
 //  • A robust safety margin (>1 px; ~half a space) avoids TAB jumping to the next stop.
    static bool computeStopsFromWidthsPx(
        HWND hSci,
        const ColumnTabs::CT_ColumnModelView& model,
        int line0, int line1,
        std::vector<int>& stopPx,   // OUT
        int gapPx)
    {
        using namespace ColumnTabs;

        if (line1 < line0) std::swap(line0, line1);

        // Uniform spacing: left-only gap; right gap must be 0 for visual uniformity.
        const int gapBeforePx = (gapPx >= 0) ? gapPx : 0;
        const int gapAfterPx = 0;

        // Safety margin so currentX is strictly < stopX even with rounding/kerning.
        // Use ~half a space width, but at least 2 px for stability across fonts/DPI.
        const int spacePx = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        const int minAdvancePx = (spacePx > 0) ? ((spacePx + 1) / 2) : 2;

        // Determine maximum field count.
        size_t maxCols = 0;
        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            maxCols = (std::max)(maxCols, L.FieldCount()); // FieldCount = nDelims + 1
        }
        if (maxCols < 2) { stopPx.clear(); return true; }
        const size_t stopsCount = maxCols - 1;

        // Clear tab stops for all affected lines (safety).
        for (int ln = line0; ln <= line1; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

        // Helpers: extract text; measure width in px.
        auto getSubText = [&](Sci_Position pos0, Sci_Position pos1, std::string& out) {
            if (pos1 <= pos0) { out.clear(); return; }
            const ptrdiff_t need = (ptrdiff_t)(pos1 - pos0);
            out.assign((size_t)need + 1u, '\0'); // +1 for '\0'

            Sci_TextRangeFull tr{};
            tr.chrg.cpMin = pos0;
            tr.chrg.cpMax = pos1;
            tr.lpstrText = out.data();
            S(hSci, SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);

            if (!out.empty()) out.pop_back(); // drop trailing '\0'
            };

        auto measurePx = [&](const std::string& s) -> int {
            const int w = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)s.c_str());
            // Robust fallback if the renderer returns 0:
            return (w > 0) ? w : (int)s.size() * ((spacePx > 0) ? spacePx : 8);
            };

        // PASS 1: collect maxima per column and per-line EOL-x under the same layout rules.
        std::vector<int> maxCellWidthPx(maxCols, 0);         // per column
        std::vector<int> maxDelimiterWidthPx(stopsCount, 0); // per delimiter position
        struct LM { std::vector<int> cellW, delimW; int lastIdx = -1; int eolX = 0; };
        std::vector<LM> lines; lines.reserve((size_t)(line1 - line0 + 1));

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            const size_t nDelim = L.delimiterOffsets.size();
            const size_t nField = nDelim + 1;

            LM lm{};
            if (nField == 0) { lines.push_back(std::move(lm)); continue; }

            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, ln);
            lm.cellW.assign(nField, 0);
            lm.delimW.assign(stopsCount, 0);
            lm.lastIdx = (int)nField - 1;

            // Measure cells; update maxima.
            for (size_t k = 0; k < nField; ++k) {
                Sci_Position s = base;
                if (k > 0)
                    s = base + (Sci_Position)L.delimiterOffsets[k - 1]
                    + (Sci_Position)model.delimiterLength;

                Sci_Position e = (k < nDelim)
                    ? (base + (Sci_Position)L.delimiterOffsets[k])
                    : (base + (Sci_Position)L.lineLength);

                std::string cell;
                getSubText(s, e, cell);

                // Tab-invariant measurement: ignore '\t' only for width computation
                if (!cell.empty()) {
                    cell.erase(std::remove(cell.begin(), cell.end(), '\t'), cell.end());
                }

                const int w = measurePx(cell);
                lm.cellW[k] = w;
                if (w > maxCellWidthPx[k]) maxCellWidthPx[k] = w;
            }

            // Measure delimiter widths (non-tab delimiter).
            if (!model.delimiterIsTab && model.delimiterLength > 0) {
                for (size_t d = 0; d < nDelim && d < stopsCount; ++d) {
                    Sci_Position d0 = base + (Sci_Position)L.delimiterOffsets[d];
                    Sci_Position d1 = d0 + (Sci_Position)model.delimiterLength;
                    std::string del;
                    getSubText(d0, d1, del);
                    const int dw = measurePx(del);
                    lm.delimW[d] = dw;
                    if (dw > maxDelimiterWidthPx[d]) maxDelimiterWidthPx[d] = dw;
                }
            }
            else {
                std::fill(lm.delimW.begin(), lm.delimW.end(), 0);
            }

            // Compute this line's EOL-x with the SAME layout model we use for stops.
            int eol = 0;
            for (int k = 0; k < lm.lastIdx; ++k) {
                eol += lm.cellW[(size_t)k] + gapBeforePx;   // left gap for next column
                eol += lm.delimW[(size_t)k] + gapAfterPx;   // right gap (0)
            }
            if (lm.lastIdx >= 0)
                eol += lm.cellW[(size_t)lm.lastIdx];        // last cell: no gap/delimiter
            lm.eolX = eol;

            lines.push_back(std::move(lm));
        }

        // PASS 2: preferred stops from maxima (uniform spacing + robust safety margin).
        std::vector<int> stopPref(stopsCount, 0);
        {
            int acc = 0;
            for (size_t c = 0; c < stopsCount; ++c) {
                // strictly beyond widest cell in column c:
                acc += maxCellWidthPx[c] + gapBeforePx + minAdvancePx;
                stopPref[c] = acc;
                acc += maxDelimiterWidthPx[c] + gapAfterPx; // 0
            }
        }

        // PASS 3: capped EOL clamps — never beyond preferred.
        std::vector<int> clamp(stopsCount, 0);
        for (const auto& lm : lines) {
            if (lm.lastIdx < 0) continue;
            for (size_t c = (size_t)lm.lastIdx; c < stopsCount; ++c) {
                const int capped = (lm.eolX < stopPref[c]) ? lm.eolX : stopPref[c];
                if (capped > clamp[c]) clamp[c] = capped;
            }
        }

        // Final: choose max(preferred, clamp) and enforce monotonicity.
        stopPx.assign(stopsCount, 0);
        for (size_t c = 0; c < stopsCount; ++c) {
            int target = (stopPref[c] < clamp[c]) ? clamp[c] : stopPref[c];
            if (c > 0 && target < stopPx[c - 1]) target = stopPx[c - 1];
            stopPx[c] = target;
        }

        return true;
    }

    // Set tab stops (in pixels) for all lines in [line0..line1].
    static void setTabStopsRangePx(HWND hSci, int line0, int line1, const std::vector<int>& stops)
    {
        for (int ln = line0; ln <= line1; ++ln) {
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
            for (size_t i = 0; i < stops.size(); ++i)
                S(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)stops[i]);
        }
    }

} // namespace ColumnTabs::detail

namespace ColumnTabs
{

    // ----------------------------------------------------------------------------
    // Indicator id (tracks inserted padding)
    // ----------------------------------------------------------------------------
    static int g_CT_IndicatorId = 30;

    void CT_SetIndicatorId(int id) noexcept { g_CT_IndicatorId = id; }
    int  CT_GetIndicatorId() noexcept { return g_CT_IndicatorId; }

    // ----------------------------------------------------------------------------
    // Utilities
    // ----------------------------------------------------------------------------
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

    // ----------------------------------------------------------------------------
    // Visual API (non-destructive)
    // ----------------------------------------------------------------------------
    bool CT_ApplyFlowTabStops(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,
        int paddingPx /*pixels*/)
    {
        using namespace detail;

        ensureCapacity(hSci);

        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, firstLine);
        const int modelCount = hasVec ? (int)model.Lines.size() : 0;

        // -1 means "apply to all lines in model" (if we have a vector-backed model).
        const int effectiveLast = (lastLine < 0 && hasVec)
            ? (line0 + modelCount - 1)
            : (lastLine < 0
                ? firstLine                        // fallback when using getLineInfo only
                : (std::max)(firstLine, lastLine));

        if (line0 > effectiveLast) return false;

        const int gapPx = (paddingPx > 0) ? paddingPx : 0;

        // Snapshot manual tab stops once per line (ETS ownership tracking).
        ensureCapacity(hSci);
        for (int ln = line0; ln <= effectiveLast; ++ln) {
            if ((int)g_hasETSLine.size() > ln && g_hasETSLine[(size_t)ln] == 0u) {
                g_savedManualStopsPx[(size_t)ln] = collectTabStopsPx(hSci, ln);
            }
        }

        // Compute and set visual tab stops.
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, effectiveLast, stops, gapPx))
            return false;

        setTabStopsRangePx(hSci, line0, effectiveLast, stops);

        // Mark lines as ETS-owned so CT_ClearFlowTabStops can act selectively.
        for (int ln = line0; ln <= effectiveLast; ++ln)
            g_hasETSLine[(size_t)ln] = 1u;

        return true;
    }


    bool CT_ApplyFlowTabStopsSpaces(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,
        int gapSpaces /*spaces*/)
    {
        const int px = (gapSpaces > 0) ? (detail::pxOfSpace(hSci) * gapSpaces) : 0;
        return CT_ApplyFlowTabStops(hSci, model, firstLine, lastLine, px);
    }

    bool CT_DisableFlowTabStops(HWND hSci, bool restoreManual)
    {
        using namespace detail;

        ensureCapacity(hSci);

        // Batch repaint while clearing many lines
        RedrawGuard rd(hSci);

        const int total = (int)SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
        const int limit = (std::min)(total, (int)g_hasETSLine.size());

        for (int ln = 0; ln < limit; ++ln) {
            if (!g_hasETSLine[(size_t)ln]) continue;

            // 1) Remove ETS-owned per-line tab stops (visual only)
            SendMessage(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

            // 2) Optionally restore previously saved manual per-line tab stops
            if (restoreManual) {
                const auto& manual = g_savedManualStopsPx[(size_t)ln];
                for (int px : manual)
                    SendMessage(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)px);
            }

            // 3) Reset ETS ownership and discard the snapshot for this line
            g_hasETSLine[(size_t)ln] = 0u;
            g_savedManualStopsPx[(size_t)ln].clear();
        }
        return true;
    }


    bool CT_ClearAllTabStops(HWND hSci)
    {
        const int total = (int)S(hSci, SCI_GETLINECOUNT);
        for (int ln = 0; ln < total; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
        return true;
    }

    void CT_ResetFlowVisualState() noexcept
    {
        detail::g_hasETSLine.clear();
        detail::g_savedManualStopsPx.clear();
    }

    bool CT_HasAlignedPadding(HWND hSci) noexcept
    {
        using namespace detail;

        // Make sure we check the correct indicator
        S(hSci, SCI_SETINDICATORCURRENT, g_CT_IndicatorId);

        const Sci_Position len = S(hSci, SCI_GETLENGTH);
        for (Sci_Position pos = 0; pos < len; ++pos) {
            // Non-zero => indicator present at this position
            if ((int)S(hSci, SCI_INDICATORVALUEAT, g_CT_IndicatorId, pos) != 0)
                return true;
        }
        return false;
    }

    bool ColumnTabs::CT_HasFlowTabStops() noexcept
    {
        using namespace detail;
        // Linear scan in memory; avoids any editor calls.
        for (size_t i = 0, n = g_hasETSLine.size(); i < n; ++i) {
            if (g_hasETSLine[i]) return true;
        }
        return false;
    }


    // ----------------------------------------------------------------------------
    // Destructive API (edits text)
    // ----------------------------------------------------------------------------
    bool CT_InsertAlignedPadding(HWND hSci, const CT_ColumnModelView& model, const CT_AlignOptions& opt)
    {
        using namespace detail;

        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, opt.firstLine);
        const int line1 = hasVec
            ? ((opt.lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(opt.firstLine, opt.lastLine))
            : opt.lastLine;
        if (line0 > line1) return false;

        // Pre-compute editor tab stops (visual aid only).
        const int gapPx = (opt.gapCells > 0) ? (pxOfSpace(hSci) * opt.gapCells) : 0;
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx)) return false;
        setTabStopsRangePx(hSci, line0, line1, stops);

        // If the delimiter itself is a tab and we only want one Flow tab, there is nothing to insert.
        if (opt.oneFlowTabOnly && model.delimiterIsTab)
            return true;

        RedrawGuard rd(hSci);
        S(hSci, SCI_BEGINUNDOACTION);
        S(hSci, SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
        S(hSci, SCI_INDICSETSTYLE, g_CT_IndicatorId, INDIC_HIDDEN);
        S(hSci, SCI_INDICSETALPHA, g_CT_IndicatorId, 0);

        // Safe delete range helper (kept for completeness; currently unused after removing trims)
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
                // Dynamic line end
                const Sci_Position lineEndNow = (Sci_Position)S(hSci, SCI_GETLINEENDPOSITION, ln, 0);

                // Current delimiter position (accounts for previous inserts in this line via 'delta')
                const Sci_Position delimPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                if (delimPos < base || delimPos > lineEndNow) {
                    // Stale offset or past EOL; skip safely
                    continue;
                }

                // Scan whitespace BEFORE delimiter (only to detect if a tab already sits directly before)
                Sci_Position wsStart = delimPos;
                while (wsStart > base) {
                    const int ch = (int)S(hSci, SCI_GETCHARAT, (uptr_t)(wsStart - 1), 0);
                    if (ch == ' ' || ch == '\t') --wsStart; else break;
                }

                // Existing '\t' exactly before the delimiter?
                const bool keepExistingTab =
                    (wsStart < delimPos) &&
                    ((int)S(hSci, SCI_GETCHARAT, (uptr_t)(delimPos - 1), 0) == '\t');

                // Do not trim any whitespace around the delimiter; preserve user's spaces/tabs.
                // Ensure exactly one Flow tab right before the delimiter (insert only if none exists there).
                if (!keepExistingTab) {
                    const Sci_Position tabPos = delimPos;  // insert at delimiter start
                    S(hSci, SCI_INSERTTEXT, tabPos, (sptr_t)"\t");
                    S(hSci, SCI_INDICATORFILLRANGE, tabPos, 1); // mark the inserted Flow tab
                    delta += 1; // keep offsets correct for subsequent delimiters in this line
                }
            }
        }

        S(hSci, SCI_ENDUNDOACTION);
        return true;
    }

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
            else {
                ++pos;
            }
        }
        for (auto it = spans.rbegin(); it != spans.rend(); ++it) {
            S(hSci, SCI_INDICATORCLEARRANGE, it->first, it->second - it->first);
            S(hSci, SCI_DELETERANGE, it->first, it->second - it->first);
        }

        // IMPORTANT: Do NOT clear tabstops here. Visual ETS will be disabled explicitly by caller.
        S(hSci, SCI_ENDUNDOACTION);

        return true;
    }

} // namespace ColumnTabs
