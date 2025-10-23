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

#include "ColumnTabs.h"
#include <algorithm>
#include <unordered_map>
#include "NumericToken.h" 

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

    // docPtr -> has pads (for O(1) gate)
    static std::unordered_map<sptr_t, bool> g_docHasPads;

    // docPtr -> list of (pos,len) we inserted (tabs/spaces), for removal fallback
    static std::unordered_map<sptr_t, std::vector<std::pair<Sci_Position, int>>> g_padRunsByDoc;


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

        // Wrap SendMessage for this hSci
        auto S = [hSci](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t {
            return (sptr_t)::SendMessage(hSci, msg, w, l);
            };

        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        const int line0 = (std::max)(0, opt.firstLine);
        const int line1 = hasVec
            ? ((opt.lastLine < 0) ? (line0 + (int)model.Lines.size() - 1)
                : (std::max)(opt.firstLine, opt.lastLine))
            : opt.lastLine;
        if (line0 > line1) return false;

        // Visual stops
        const int gapPx = (opt.gapCells > 0) ? (pxOfSpace(hSci) * opt.gapCells) : 0;
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx)) return false;
        setTabStopsRangePx(hSci, line0, line1, stops);

        // One-flow-tab-only + delimiter is TAB => nothing to insert
        if (opt.oneFlowTabOnly && model.delimiterIsTab)
            return true;

        RedrawGuard rd(hSci);

        S(SCI_BEGINUNDOACTION);
        S(SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
        S(SCI_INDICSETSTYLE, g_CT_IndicatorId, INDIC_HIDDEN);
        S(SCI_INDICSETALPHA, g_CT_IndicatorId, 0);

        bool madePads = false;

        // We'll record inserted runs for fallback removal
        const sptr_t doc = (sptr_t)S(SCI_GETDOCPOINTER);

        auto fetchLine = [&](int ln)->CT_ColumnLineInfo {
            if (hasVec) return model.Lines[(size_t)(ln - line0)];
            return model.getLineInfo ? model.getLineInfo(ln) : CT_ColumnLineInfo{};
            };

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(ln);
            const Sci_Position base = (Sci_Position)S(SCI_POSITIONFROMLINE, ln);
            Sci_Position delta = 0;

            const size_t nDelims = L.delimiterOffsets.size();
            if (nDelims == 0) continue;

            for (size_t c = 0; c < nDelims; ++c) {
                const Sci_Position lineEndNow = (Sci_Position)S(SCI_GETLINEENDPOSITION, ln, 0);
                const Sci_Position delimPos = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                if (delimPos < base || delimPos > lineEndNow) continue;

                // scan whitespace BEFORE delimiter
                Sci_Position wsStart = delimPos;
                while (wsStart > base) {
                    const int ch = (int)S(SCI_GETCHARAT, (uptr_t)(wsStart - 1), 0);
                    if (ch == ' ' || ch == '\t') --wsStart; else break;
                }

                const bool keepExistingTab =
                    (wsStart < delimPos) &&
                    ((int)S(SCI_GETCHARAT, (uptr_t)(delimPos - 1), 0) == '\t');

                if (!keepExistingTab) {
                    const Sci_Position tabPos = delimPos;
                    S(SCI_INSERTTEXT, (uptr_t)tabPos, (sptr_t)"\t");
                    S(SCI_INDICATORFILLRANGE, (uptr_t)tabPos, (sptr_t)1);
                    delta += 1;
                    madePads = true;

                    // record for fallback removal
                    g_padRunsByDoc[doc].emplace_back(tabPos, 1);
                }
            }
        }

        S(SCI_ENDUNDOACTION);

        if (madePads)
            ColumnTabs::CT_SetDocHasPads(doc, true);

        return true;
    }

    bool ColumnTabs::CT_RemoveAlignedPadding(HWND hSci)
    {
        auto S = [hSci](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t {
            return (sptr_t)::SendMessage(hSci, msg, w, l);
            };

        const int ind = CT_GetIndicatorId();
        S(SCI_SETINDICATORCURRENT, (uptr_t)ind);

        bool removedAny = false;

        // --- Fast path: remove by indicator marks
        S(SCI_BEGINUNDOACTION);
        {
            Sci_Position pos = (Sci_Position)S(SCI_GETLENGTH);
            while (pos > 0) {
                --pos;
                if ((int)S(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)pos) == 0)
                    continue;

                const Sci_Position start = (Sci_Position)S(SCI_INDICATORSTART, (uptr_t)ind, (sptr_t)pos);
                const Sci_Position end = (Sci_Position)S(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)pos);
                if (end > start) {
                    S(SCI_INDICATORCLEARRANGE, (uptr_t)start, (sptr_t)(end - start));
                    S(SCI_DELETERANGE, (uptr_t)start, (sptr_t)(end - start));
                    pos = start;
                    removedAny = true;
                }
            }
        }
        S(SCI_ENDUNDOACTION);

        // --- Fallback: recorded runs (if indicators are gone)
        const sptr_t doc = (sptr_t)S(SCI_GETDOCPOINTER);
        if (!removedAny) {
            auto it = detail::g_padRunsByDoc.find(doc);
            if (it != detail::g_padRunsByDoc.end() && !it->second.empty()) {
                auto& runs = it->second;

                // delete from end to start
                std::sort(runs.begin(), runs.end(),
                    [](const auto& a, const auto& b) { return a.first > b.first; });

                S(SCI_BEGINUNDOACTION);
                const Sci_Position docLen = (Sci_Position)S(SCI_GETLENGTH);
                for (const auto& r : runs) {
                    const Sci_Position p = r.first;
                    const Sci_Position n = (Sci_Position)r.second;
                    if (p < 0 || p + n > docLen) continue;

                    bool ok = true;
                    for (Sci_Position k = 0; k < n; ++k) {
                        const int ch = (int)S(SCI_GETCHARAT, (uptr_t)(p + k), 0);
                        if (!(ch == ' ' || ch == '\t')) { ok = false; break; }
                    }
                    if (!ok) continue;

                    S(SCI_DELETERANGE, (uptr_t)p, (sptr_t)n);
                    removedAny = true;
                }
                S(SCI_ENDUNDOACTION);

                runs.clear(); // consumed
            }
        }

        // --- Keep doc state & fallback-index in sync
        if (removedAny) {
            ColumnTabs::CT_SetCurDocHasPads(hSci, false);
            detail::g_padRunsByDoc.erase(doc);
        }

        return removedAny; // reflect whether anything was actually removed
    }

    // small helpers (file-local)
    static int CT_TextWidthPx(HWND hSci, const std::string& s) {
        const int w = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)s.c_str());
        if (w > 0) return w;
        const int sp = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        return (int)s.size() * (sp > 0 ? sp : 8);
    }

    static void CT_GetSubTextUtf8(HWND hSci, Sci_Position pos0, Sci_Position pos1, std::string& out) {
        out.clear();
        if (pos1 <= pos0) return;
        const ptrdiff_t need = (ptrdiff_t)(pos1 - pos0);
        out.assign((size_t)need + 1u, '\0');
        Sci_TextRangeFull tr{}; tr.chrg.cpMin = pos0; tr.chrg.cpMax = pos1; tr.lpstrText = out.data();
        S(hSci, SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);
        if (!out.empty()) out.pop_back();
    }


    // find number in [fieldPos0, fieldPos1): returns doc pos of first digit/sign and px width of left part
    static bool CT_FindNumericToken(HWND hSci,
        Sci_Position fieldPos0,
        Sci_Position fieldPos1,
        Sci_Position& outStartDoc,
        int& outLeftPx,
        bool& outHasDecimal)
    {
        std::string fld;
        CT_GetSubTextUtf8(hSci, fieldPos0, fieldPos1, fld);
        if (fld.empty()) return false;

        // Trim spaces/tabs around the field slice
        auto is_space = [](unsigned char c) noexcept { return c == ' ' || c == '\t'; };
        std::size_t n = fld.size();
        std::size_t t0 = 0, t1 = n;
        while (t0 < n && is_space((unsigned char)fld[t0])) ++t0;
        while (t1 > t0 && is_space((unsigned char)fld[t1 - 1])) --t1;

        if (t1 <= t0) return false;
        std::string_view trimmed(fld.data() + t0, t1 - t0);

        mr::num::NumericToken tok;
        if (!mr::num::classify_numeric_field(trimmed, tok))
            return false;

        // Compute pixel width of integer part (no sign, up to decimal separator)
        std::size_t intStart = tok.start + (tok.hasSign ? 1u : 0u);
        std::size_t intLen = (tok.intDigits > 0) ? (std::size_t)tok.intDigits : 0u;

        std::string leftDigits;
        if (intLen > 0 && intStart + intLen <= trimmed.size())
            leftDigits.assign(trimmed.data() + intStart, intLen);

        outLeftPx = CT_TextWidthPx(hSci, leftDigits);
        outStartDoc = fieldPos0 + (Sci_Position)(t0 + tok.start);
        outHasDecimal = tok.hasDecimal;
        return true;
    }

    bool ColumnTabs::CT_ApplyNumericPadding(
        HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine)
    {
        const HWND hwnd = hSci;
        auto S = [hwnd](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t { return (sptr_t)::SendMessage(hwnd, msg, w, l); };
        auto ls = [&](int ln)->Sci_Position { return (Sci_Position)S(SCI_POSITIONFROMLINE, (uptr_t)ln); };
        auto le = [&](int ln)->Sci_Position { return (Sci_Position)S(SCI_GETLINEENDPOSITION, (uptr_t)ln); };
        auto gc = [&](Sci_Position p)->int { return (int)S(SCI_GETCHARAT, (uptr_t)p); };

        const bool haveVec = !model.Lines.empty();
        const bool liveInfo = (model.getLineInfo != nullptr);
        if (!haveVec && !liveInfo) return false;

        const int l0 = (firstLine < 0) ? 0 : firstLine;
        int       l1 = (lastLine < 0) ? (haveVec ? (l0 + (int)model.Lines.size() - 1) : l0) : lastLine;
        if (l0 > l1) return false;

        S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
        S(SCI_INDICSETSTYLE, (uptr_t)CT_GetIndicatorId(), INDIC_HIDDEN);
        S(SCI_INDICSETUNDER, (uptr_t)CT_GetIndicatorId(), 1);

        auto fetchLine = [&](int ln)->CT_ColumnLineInfo {
            if (liveInfo) return model.getLineInfo(ln);
            size_t i = (size_t)(ln - l0);
            if (i >= model.Lines.size()) i = model.Lines.size() - 1;
            return model.Lines[i];
            };

        size_t maxCols = 0;
        for (int ln = l0; ln <= l1; ++ln)
            maxCols = (std::max)(maxCols, fetchLine(ln).FieldCount());
        if (maxCols == 0) return true;

        std::vector<int>  maxIntDigits(maxCols, 0);
        std::vector<char> colHasDec(maxCols, 0);

        // Detection: rely solely on CT_FindNumericToken (strict numeric fields)
        auto findNumberUsingCT = [hwnd](Sci_Position s, Sci_Position e,
            Sci_Position& tokStart,
            int& intDigits,
            bool& hasDec) -> bool
            {
                int dummyPx = 0; hasDec = false;
                if (!ColumnTabs::CT_FindNumericToken(hwnd, s, e, tokStart, dummyPx, hasDec))
                    return false;

                // Count integer digits (skip sign) up to separator
                auto gcLoc = [hwnd](Sci_Position p)->int { return (int)::SendMessage(hwnd, SCI_GETCHARAT, (uptr_t)p, 0); };
                intDigits = 0;
                Sci_Position p = tokStart;
                int ch = gcLoc(p);
                if (ch == '+' || ch == '-') ++p;
                for (;;) {
                    ch = gcLoc(p);
                    if (ch >= '0' && ch <= '9') { ++intDigits; ++p; continue; }
                    if (ch == '.' || ch == ',') break;
                    break;
                }
                return (intDigits > 0) || hasDec;
            };

        // PASS 1: collect maxima and decimal flags per column
        for (int ln = l0; ln <= l1; ++ln) {
            const auto L = fetchLine(ln);
            const size_t nField = L.FieldCount();
            Sci_Position dummyDelta = 0;
            for (size_t c = 0; c < nField; ++c) {
                Sci_Position s{}, e{}; // field bounds
                {
                    const Sci_Position base = ls(ln);
                    const size_t nDelim = L.delimiterOffsets.size();
                    const size_t nField2 = nDelim + 1;
                    if (c >= nField2) continue;

                    // end
                    if (c < nDelim) {
                        e = base + (Sci_Position)L.delimiterOffsets[c] + dummyDelta;
                        while (e > base) { int chL = gc(e - 1); if (chL == ' ' || chL == '\t') --e; else break; }
                    }
                    else e = le(ln);

                    // start
                    if (c == 0) s = base;
                    else {
                        s = base + (Sci_Position)L.delimiterOffsets[c - 1] + dummyDelta + (Sci_Position)model.delimiterLength;
                        while (s < e) { int chR = gc(s); if (chR == ' ' || chR == '\t') ++s; else break; }
                    }
                }

                Sci_Position tokStart{}; int intDigits = 0; bool hasDec = false;
                if (findNumberUsingCT(s, e, tokStart, intDigits, hasDec)) {
                    if (intDigits > maxIntDigits[c]) maxIntDigits[c] = intDigits;
                    if (hasDec) colHasDec[c] = 1;
                }
            }
        }

        // PASS 2: apply spaces (only ours, marked by indicator)
        S(SCI_BEGINUNDOACTION);

        for (int ln = l0; ln <= l1; ++ln) {
            const auto L0 = fetchLine(ln);
            const size_t nField = L0.FieldCount();
            Sci_Position delta = 0;

            for (size_t c = 0; c < nField; ++c) {
                if (c >= maxCols || maxIntDigits[c] <= 0) continue;

                const auto L = fetchLine(ln);
                Sci_Position s{}, e{};
                {
                    const Sci_Position base = ls(ln);
                    const size_t nDelim = L.delimiterOffsets.size();
                    const size_t nField2 = nDelim + 1;
                    if (c >= nField2) continue;

                    if (c < nDelim) {
                        e = base + (Sci_Position)L.delimiterOffsets[c] + delta;
                        while (e > base) { int chL = gc(e - 1); if (chL == ' ' || chL == '\t') --e; else break; }
                    }
                    else e = le(ln);

                    if (c == 0) s = base;
                    else {
                        s = base + (Sci_Position)L.delimiterOffsets[c - 1] + delta + (Sci_Position)model.delimiterLength;
                        while (s < e) { int chR = gc(s); if (chR == ' ' || chR == '\t') ++s; else break; }
                    }
                }

                Sci_Position tokStart{}; int intDigits = 0; bool hasDec = false;
                const bool isNum = findNumberUsingCT(s, e, tokStart, intDigits, hasDec);
                if (!isNum) continue;

                const Sci_Position base = ls(ln);
                const Sci_Position fieldStartRaw =
                    (c == 0)
                    ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1] + delta + (Sci_Position)model.delimiterLength;

                // count ASCII spaces immediately left of the token
                int spacesTouching = 0;
                { Sci_Position p = tokStart; while (p > fieldStartRaw && gc(p - 1) == ' ') { --p; ++spacesTouching; } }

                const int baseGapUnits = 1;

                bool hasSign2 = false;
                if (tokStart < e) { const int ch0 = gc(tokStart); hasSign2 = (ch0 == '+' || ch0 == '-'); }

                int needUnits = baseGapUnits + (maxIntDigits[c] - intDigits);
                if (needUnits < 0) needUnits = 0;
                if (hasSign2 && needUnits > 0) --needUnits; // always let the sign "hang"

                const int diff = needUnits - spacesTouching;

                if (diff > 0) {
                    std::string pad((size_t)diff, ' ');
                    S(SCI_INSERTTEXT, (uptr_t)tokStart, (sptr_t)pad.c_str());
                    S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
                    S(SCI_INDICATORFILLRANGE, (uptr_t)tokStart, (sptr_t)diff);
                    delta += (Sci_Position)diff;
                }
                else if (diff < 0) {
                    int toDel = -diff;
                    const int ind = CT_GetIndicatorId();
                    S(SCI_SETINDICATORCURRENT, (uptr_t)ind);

                    int delAvail = 0;
                    if (tokStart > 0 &&
                        (int)S(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)(tokStart - 1)) != 0)
                    {
                        const Sci_Position runBeg = (Sci_Position)S(SCI_INDICATORSTART, (uptr_t)ind, (sptr_t)(tokStart - 1));
                        const Sci_Position lower = (runBeg < fieldStartRaw) ? fieldStartRaw : runBeg;
                        delAvail = (int)(tokStart - lower);
                    }
                    if (delAvail > 0) {
                        if (toDel > delAvail) toDel = delAvail;
                        const Sci_Position posDel = tokStart - (Sci_Position)toDel;
                        S(SCI_INDICATORCLEARRANGE, (uptr_t)posDel, (sptr_t)toDel);
                        S(SCI_DELETERANGE, (uptr_t)posDel, (sptr_t)toDel);
                        delta -= (Sci_Position)toDel;
                    }
                }
            }
        }

        S(SCI_ENDUNDOACTION);
        return true;
    }




    void ColumnTabs::CT_SetDocHasPads(sptr_t docPtr, bool has) noexcept {
        detail::g_docHasPads[docPtr] = has;
    }
    bool ColumnTabs::CT_GetDocHasPads(sptr_t docPtr) noexcept {
        auto it = detail::g_docHasPads.find(docPtr);
        return (it != detail::g_docHasPads.end()) && it->second;
    }
    void ColumnTabs::CT_SetCurDocHasPads(HWND hSci, bool has) noexcept {
        sptr_t doc = (sptr_t)::SendMessage(hSci, SCI_GETDOCPOINTER, 0, 0);
        CT_SetDocHasPads(doc, has);
    }
    bool ColumnTabs::CT_GetCurDocHasPads(HWND hSci) noexcept {
        sptr_t doc = (sptr_t)::SendMessage(hSci, SCI_GETDOCPOINTER, 0, 0);
        return CT_GetDocHasPads(doc);
    }

    bool ColumnTabs::CT_CleanupVisuals(HWND hSci) noexcept
    {
        if (!hSci) return false;
        CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);   // visual only
        CT_ResetFlowVisualState();
        return true;
    }

    bool ColumnTabs::CT_CleanupAllForDoc(HWND hSci) noexcept
    {
        if (!hSci) return false;
        // hard cleanup (text + visuals)
        CT_RemoveAlignedPadding(hSci);                           // marks-only delete (oder Fallback)
        CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);
        CT_ResetFlowVisualState();
        return true;
    }


} // namespace ColumnTabs
