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
#include "NumericToken.h"
#include <algorithm>
#include <unordered_map>
#include <string_view>
#include <cctype> 

// --- Scintilla direct --------------------------------------------------------
static inline sptr_t S(HWND hSci, UINT m, uptr_t w = 0, sptr_t l = 0)
{
    static thread_local HWND         cachedHwnd = nullptr;
    static thread_local SciFnDirect  cachedFn = nullptr;
    static thread_local sptr_t       cachedPtr = 0;

    if (hSci != cachedHwnd || !cachedFn || !cachedPtr) {
        cachedFn = reinterpret_cast<SciFnDirect>(::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
        cachedPtr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
        cachedHwnd = hSci;
    }
    return cachedFn ? cachedFn(cachedPtr, m, w, l) : ::SendMessage(hSci, m, w, l);
}

// -----------------------------------------------------------------------------
// File-local helpers (no linkage outside this TU)
// -----------------------------------------------------------------------------
struct RedrawGuard {
    HWND h{};
    explicit RedrawGuard(HWND hwnd) : h(hwnd) {
        ::SendMessage(h, WM_SETREDRAW, FALSE, 0);
    }
    ~RedrawGuard() {
        ::SendMessage(h, WM_SETREDRAW, TRUE, 0);
        ::InvalidateRect(h, nullptr, TRUE);
    }
};

struct OptionalRedrawGuard {
    HWND  h{};
    bool  active = false;
    OptionalRedrawGuard(HWND hwnd, size_t opCount,
        size_t threshold = 2000) : h(hwnd) {
        // threshold: expected number of editor calls/line operations
        active = (opCount >= threshold);
        if (active) ::SendMessage(h, WM_SETREDRAW, FALSE, 0);
    }
    ~OptionalRedrawGuard() {
        if (active) {
            ::SendMessage(h, WM_SETREDRAW, TRUE, 0);
            ::InvalidateRect(h, nullptr, TRUE);
        }
    }
};

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

// -----------------------------------------------------------------------------
// ColumnTabs::detail – persistent state and helpers  (separater Namespace)
// -----------------------------------------------------------------------------
namespace ColumnTabs::detail {

    // docPtr -> has pads (for O(1) gate)
    static std::unordered_map<sptr_t, bool> g_docHasPads;

    // Tracks which lines currently have ETS-owned visual tab stops.
    static std::vector<uint8_t> g_hasETSLine;                 // 0/1 per line

    // Snapshot of manual tab stops (in px) that existed before ETS took over the line.
    static std::vector<std::vector<int>> g_savedManualStopsPx;

    // Ensure our global tracking vectors have capacity for the current buffer.
    inline void ensureCapacity(HWND hSci) {
        const int total = (int)S(hSci, SCI_GETLINECOUNT, 0, 0);

        if ((int)g_hasETSLine.size() != total)
            g_hasETSLine.resize((size_t)total, 0u);

        if ((int)g_savedManualStopsPx.size() != total)
            g_savedManualStopsPx.resize((size_t)total);
    }

    // Collect all current tab stops (px) on a given line (manual or otherwise).
    static std::vector<int> collectTabStopsPx(HWND hSci, int line) {
        std::vector<int> stops;
        int pos = 0;
        for (;;) {
            const int next = (int)S(hSci, SCI_GETNEXTTABSTOP, (uptr_t)line, (sptr_t)pos);
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

    // Expects doc-absolute line indices [line0..line1].
    static bool computeStopsFromWidthsPx(
        HWND hSci,
        const ColumnTabs::CT_ColumnModelView& model,
        int line0, int line1,
        std::vector<int>& stopPx,   // OUT
        int gapPx)
    {
        using namespace ColumnTabs;

        if (line1 < line0) std::swap(line0, line1);

        // Uniform spacing: left-only gap; right gap must be 0.
        const int gapBeforePx = (gapPx >= 0) ? gapPx : 0;
        const int gapAfterPx = 0;

        // Safety margin to stay strictly left of stop with rounding/kerning.
        const int spacePx = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        const int minAdvancePx = (spacePx > 0) ? ((spacePx + 1) / 2) : 2;

        // Determine maximum field count across the range.
        size_t maxCols = 0;
        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln); // model/doc-absolute
            maxCols = (std::max)(maxCols, L.FieldCount()); // FieldCount = nDelims + 1
        }
        if (maxCols < 2) { stopPx.clear(); return true; }
        const size_t stopsCount = maxCols - 1;

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
            return (w > 0) ? w : (int)s.size() * ((spacePx > 0) ? spacePx : 8);
            };

        // PASS 1: collect maxima per column and per-line EOL-x under same rules.
        std::vector<int> maxCellWidthPx(maxCols, 0);
        std::vector<int> maxDelimiterWidthPx(stopsCount, 0);
        struct LM { std::vector<int> cellW, delimW; int lastIdx = -1; int eolX = 0; };
        std::vector<LM> lines; lines.reserve((size_t)(line1 - line0 + 1));

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetchLine(model, ln);
            const size_t nDel = L.delimiterOffsets.size();
            const size_t nFld = nDel + 1;

            LM lm{};
            if (nFld == 0) { lines.push_back(std::move(lm)); continue; }

            const Sci_Position base = (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, (uptr_t)ln);
            lm.cellW.assign(nFld, 0);
            lm.delimW.assign(stopsCount, 0);
            lm.lastIdx = (int)nFld - 1;

            // Cells
            for (size_t k = 0; k < nFld; ++k) {
                Sci_Position s = base;
                if (k > 0)
                    s = base + (Sci_Position)L.delimiterOffsets[k - 1]
                    + (Sci_Position)model.delimiterLength;

                Sci_Position e = (k < nDel)
                    ? (base + (Sci_Position)L.delimiterOffsets[k])
                    : (base + (Sci_Position)L.lineLength);

                std::string cell;
                getSubText(s, e, cell);

                // Tab-invariant measure: ignore '\t' for width.
                if (!cell.empty())
                    cell.erase(std::remove(cell.begin(), cell.end(), '\t'), cell.end());

                const int w = measurePx(cell);
                lm.cellW[k] = w;
                if (w > maxCellWidthPx[k]) maxCellWidthPx[k] = w;
            }

            // Delimiters (only if not TAB)
            if (!model.delimiterIsTab && model.delimiterLength > 0) {
                for (size_t d = 0; d < nDel && d < stopsCount; ++d) {
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

            // EOL under the same layout model.
            int eol = 0;
            for (int k = 0; k < lm.lastIdx; ++k) {
                eol += lm.cellW[(size_t)k] + gapBeforePx;
                eol += lm.delimW[(size_t)k] + gapAfterPx; // 0
            }
            if (lm.lastIdx >= 0)
                eol += lm.cellW[(size_t)lm.lastIdx];
            lm.eolX = eol;

            lines.push_back(std::move(lm));
        }

        // PASS 2: preferred stops from maxima (+ safety margin).
        std::vector<int> stopPref(stopsCount, 0);
        {
            int acc = 0;
            for (size_t c = 0; c < stopsCount; ++c) {
                acc += maxCellWidthPx[c] + gapBeforePx + minAdvancePx; // strictly beyond widest cell
                stopPref[c] = acc;
                acc += maxDelimiterWidthPx[c] + gapAfterPx;            // 0
            }
        }

        // PASS 3: EOL clamps — never beyond preferred (guard rail).
        std::vector<int> clamp(stopsCount, 0);
        for (const auto& lm : lines) {
            if (lm.lastIdx < 0) continue;
            for (size_t c = (size_t)lm.lastIdx; c < stopsCount; ++c) {
                const int capped = (lm.eolX < stopPref[c]) ? lm.eolX : stopPref[c];
                if (capped > clamp[c]) clamp[c] = capped;
            }
        }

        // Final: max(preferred, clamp) and enforce monotonicity.
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
        const size_t perLine = 1u + stops.size();
        const size_t lines = (line1 >= line0) ? (size_t)(line1 - line0 + 1) : 0;
        OptionalRedrawGuard rg(hSci, perLine * lines); // active from ~2000 operations

        for (int ln = line0; ln <= line1; ++ln) {
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
            for (size_t i = 0; i < stops.size(); ++i)
                S(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)stops[i]);
        }
    }

} // namespace ColumnTabs::detail

// -----------------------------------------------------------------------------
// ColumnTabs – all public APIs (separater Namespace)
// -----------------------------------------------------------------------------
namespace ColumnTabs {

    // -------------------------------------------------------------------------
    // Indicator (tracks inserted padding)
    // -------------------------------------------------------------------------
    static int g_CT_IndicatorId = 30;

    void CT_SetIndicatorId(int id) noexcept { g_CT_IndicatorId = id; }
    int  CT_GetIndicatorId() noexcept { return g_CT_IndicatorId; }

    // -------------------------------------------------------------------------
    // Destructive API (edits text)
    // -------------------------------------------------------------------------
    bool ColumnTabs::CT_InsertAlignedPadding(HWND hSci,
        const CT_ColumnModelView& model,
        const CT_AlignOptions& opt,
        bool* outNothingToAlign /*=nullptr*/)
    {
        using namespace ColumnTabs::detail;

        if (outNothingToAlign) *outNothingToAlign = false;

        auto Sci = [hSci](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t {
            return (sptr_t)::S(hSci, msg, w, l);
            };

        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) {
            return false;
        }

        const int baseDoc = (int)model.docStartLine;
        const int rel0 = (opt.firstLine < 0) ? 0 : opt.firstLine;

        int line0 = baseDoc + rel0;
        int line1;

        if (hasVec) {
            const int rel1 = (opt.lastLine < 0)
                ? ((int)model.Lines.size() - 1)
                : (std::max)(opt.firstLine, opt.lastLine);
            line1 = baseDoc + rel1;

            const int modelFirst = baseDoc;
            const int modelLast = baseDoc + (int)model.Lines.size() - 1;
            if (line0 < modelFirst) line0 = modelFirst;
            if (line1 > modelLast)  line1 = modelLast;
        }
        else {
            if (opt.lastLine < 0) {
                const int docLast = (int)S(hSci, SCI_GETLINECOUNT, 0, 0) - 1;
                line1 = (docLast < line0) ? line0 : docLast;
            }
            else {
                line1 = baseDoc + opt.lastLine;
                if (line1 < line0) line1 = line0;
            }
        }

        if (line1 < line0) {
            if (outNothingToAlign) *outNothingToAlign = true;
            return false;
        }

        ensureCapacity(hSci);

        const int gapPx = (opt.gapCells > 0) ? (pxOfSpace(hSci) * opt.gapCells) : 0;
        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, line1, stops, gapPx)) {
            return false;
        }

        if (stops.empty()) {
            if (outNothingToAlign) *outNothingToAlign = true;
            return false;
        }

        for (int ln = line0; ln <= line1; ++ln) {
            if ((int)g_hasETSLine.size() > ln &&
                (int)g_savedManualStopsPx.size() > ln &&
                g_hasETSLine[(size_t)ln] == 0u)
            {
                g_savedManualStopsPx[(size_t)ln] = collectTabStopsPx(hSci, ln);
            }
        }

        setTabStopsRangePx(hSci, line0, line1, stops);

        for (int ln = line0; ln <= line1; ++ln) {
            if ((int)g_hasETSLine.size() > ln)
                g_hasETSLine[(size_t)ln] = 1u;
        }

        if (opt.oneFlowTabOnly && model.delimiterIsTab)
            return true;

        RedrawGuard rd(hSci);

        // Suppress undo collection: FlowTab padding is cosmetic, the toggle
        // button serves as the undo mechanism.  We keep BEGIN/ENDUNDOACTION
        // as a performance hint so Scintilla optimises its gap buffer.
        const bool undoWasOn = Sci(SCI_GETUNDOCOLLECTION) != 0;
        if (undoWasOn) Sci(SCI_SETUNDOCOLLECTION, 0);

        // Suppress modification notifications during bulk insert to avoid
        // per-edit re-lexing and folding overhead.
        const int savedEventMask = (int)Sci(SCI_GETMODEVENTMASK);
        Sci(SCI_SETMODEVENTMASK, 0);

        Sci(SCI_BEGINUNDOACTION);

        Sci(SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
        Sci(SCI_INDICSETSTYLE, g_CT_IndicatorId, INDIC_HIDDEN);
        Sci(SCI_INDICSETALPHA, g_CT_IndicatorId, 0);

        bool madePads = false;

        // Bulk approach: build modified text per line in memory, then replace
        // the whole line via SCI_REPLACETARGET (one Scintilla call per line
        // instead of one per delimiter — avoids O(n²) gap-buffer shifts).
        auto fetch = [&](int ln)->CT_ColumnLineInfo {
            return detail::fetchLine(model, ln);
            };

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetch(ln);
            const size_t nDelims = L.delimiterOffsets.size();
            if (nDelims == 0) continue;

            const Sci_Position lineStart = (Sci_Position)Sci(SCI_POSITIONFROMLINE, (uptr_t)ln);
            const Sci_Position lineEnd = (Sci_Position)Sci(SCI_GETLINEENDPOSITION, (uptr_t)ln);
            const Sci_Position lineLen = lineEnd - lineStart;
            if (lineLen <= 0) continue;

            // Collect existing indicator runs on this line (e.g. from NumericPadding)
            // as byte offsets relative to lineStart.  REPLACETARGET destroys them,
            // so we must capture and re-apply after the replace.
            std::vector<std::pair<size_t, size_t>> existingIndicRuns; // (offset, len)
            {
                Sci_Position p = lineStart;
                while (p < lineEnd) {
                    int val = (int)Sci(SCI_INDICATORVALUEAT, (uptr_t)g_CT_IndicatorId, (sptr_t)p);
                    Sci_Position runEnd = (Sci_Position)Sci(SCI_INDICATOREND, (uptr_t)g_CT_IndicatorId, (sptr_t)p);
                    if (runEnd <= p) break;
                    if (runEnd > lineEnd) runEnd = lineEnd;
                    if (val != 0) {
                        existingIndicRuns.push_back({
                            (size_t)(p - lineStart),
                            (size_t)(runEnd - p) });
                    }
                    p = runEnd;
                }
            }

            // Read current line text
            std::string origLine((size_t)lineLen, '\0');
            Sci_TextRangeFull tr{};
            tr.chrg.cpMin = lineStart;
            tr.chrg.cpMax = lineEnd;
            tr.lpstrText = origLine.data();
            Sci(SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);

            // Build new line with tabs inserted before delimiters.
            std::string newLine;
            newLine.reserve(origLine.size() + nDelims);
            std::vector<std::pair<size_t, size_t>> indicRanges; // (pos, len) in newLine
            std::vector<size_t> tabInsertOffsets;                // original offsets of insertions

            size_t prevOff = 0;
            int tabsInserted = 0;
            for (size_t c = 0; c < nDelims; ++c) {
                size_t delimOff = (size_t)L.delimiterOffsets[c];
                if (delimOff > origLine.size()) continue;

                size_t wsStart = delimOff;
                while (wsStart > 0 && (origLine[wsStart - 1] == ' ' || origLine[wsStart - 1] == '\t'))
                    --wsStart;

                bool keepExistingTab =
                    (wsStart < delimOff) && (origLine[delimOff - 1] == '\t');

                newLine.append(origLine, prevOff, delimOff - prevOff);

                if (!keepExistingTab) {
                    size_t tabPos = newLine.size();
                    newLine += '\t';
                    indicRanges.push_back({ tabPos, 1 });
                    tabInsertOffsets.push_back(delimOff);
                    ++tabsInserted;
                    madePads = true;
                }

                prevOff = delimOff;
            }
            newLine.append(origLine, prevOff, origLine.size() - prevOff);

            if (tabsInserted == 0) continue;

            // Remap existing indicator runs to new positions (shifted by inserted tabs).
            for (const auto& er : existingIndicRuns) {
                size_t origOff = er.first;
                size_t shift = 0;
                for (size_t tOff : tabInsertOffsets) {
                    if (tOff <= origOff) ++shift;
                    else break;
                }
                indicRanges.push_back({ origOff + shift, er.second });
            }

            // Replace the entire line in one shot
            Sci(SCI_SETTARGETRANGE, (uptr_t)lineStart, (sptr_t)lineEnd);
            Sci(SCI_REPLACETARGET, (uptr_t)newLine.size(), (sptr_t)newLine.c_str());

            // Re-apply all indicators (new tabs + remapped existing)
            for (const auto& ir : indicRanges) {
                Sci(SCI_INDICATORFILLRANGE,
                    (uptr_t)(lineStart + (Sci_Position)ir.first),
                    (sptr_t)ir.second);
            }
        }

        Sci(SCI_ENDUNDOACTION);
        Sci(SCI_SETMODEVENTMASK, (uptr_t)savedEventMask);

        if (undoWasOn) {
            Sci(SCI_SETUNDOCOLLECTION, 1);
            if (madePads) Sci(SCI_EMPTYUNDOBUFFER);
        }

        if (madePads) {
            CT_SetCurDocHasPads(hSci, true);
        }

        return true;
    }

    bool ColumnTabs::CT_RemoveAlignedPadding(HWND hSci)
    {
        if (!ColumnTabs::CT_GetCurDocHasPads(hSci))
            return false;

        auto S = [hSci](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t {
            return (sptr_t)::S(hSci, msg, w, l);
            };

        const int ind = CT_GetIndicatorId();
        S(SCI_SETINDICATORCURRENT, (uptr_t)ind);

        const Sci_Position docLen = (Sci_Position)S(SCI_GETLENGTH);
        if (docLen <= 0) {
            ColumnTabs::CT_SetCurDocHasPads(hSci, false);
            return false;
        }

        // Suppress undo collection and modification notifications.
        const bool undoWasOn = S(SCI_GETUNDOCOLLECTION) != 0;
        if (undoWasOn) S(SCI_SETUNDOCOLLECTION, 0);

        const int savedEventMask = (int)S(SCI_GETMODEVENTMASK);
        S(SCI_SETMODEVENTMASK, 0);

        S(SCI_BEGINUNDOACTION);

        // Bulk removal: for each line that contains indicator-marked bytes,
        // build a new line without them and replace via SCI_REPLACETARGET.
        // Iterating back-to-front keeps earlier line positions stable.
        const int lineCount = (int)S(SCI_GETLINECOUNT);

        for (int ln = lineCount - 1; ln >= 0; --ln) {
            const Sci_Position lineStart = (Sci_Position)S(SCI_POSITIONFROMLINE, (uptr_t)ln);
            const Sci_Position lineEnd = (Sci_Position)S(SCI_GETLINEENDPOSITION, (uptr_t)ln);
            const Sci_Position lineLen = lineEnd - lineStart;
            if (lineLen <= 0) continue;

            // Quick check: any indicator on this line?
            bool hasIndicator = false;
            {
                Sci_Position p = lineStart;
                while (p < lineEnd) {
                    if ((int)S(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)p) != 0) {
                        hasIndicator = true;
                        break;
                    }
                    Sci_Position nextEnd = (Sci_Position)S(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)p);
                    if (nextEnd <= p) break;
                    p = nextEnd;
                }
            }
            if (!hasIndicator) continue;

            // Read line text
            std::string origLine((size_t)lineLen, '\0');
            Sci_TextRangeFull tr{};
            tr.chrg.cpMin = lineStart;
            tr.chrg.cpMax = lineEnd;
            tr.lpstrText = origLine.data();
            S(SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);

            // Build new line without indicator-marked bytes
            std::string newLine;
            newLine.reserve(origLine.size());

            Sci_Position p = lineStart;
            size_t srcIdx = 0;
            while (p < lineEnd) {
                int val = (int)S(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)p);
                Sci_Position runEnd = (Sci_Position)S(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)p);
                if (runEnd <= p) break;
                if (runEnd > lineEnd) runEnd = lineEnd;

                const size_t runLen = (size_t)(runEnd - p);

                if (val == 0) {
                    newLine.append(origLine, srcIdx, runLen);
                }
                // else: skip indicator-marked padding bytes

                srcIdx += runLen;
                p = runEnd;
            }

            if (newLine.size() == (size_t)lineLen) continue;

            S(SCI_INDICATORCLEARRANGE, (uptr_t)lineStart, (sptr_t)lineLen);
            S(SCI_SETTARGETRANGE, (uptr_t)lineStart, (sptr_t)lineEnd);
            S(SCI_REPLACETARGET, (uptr_t)newLine.size(), (sptr_t)newLine.c_str());
        }

        S(SCI_ENDUNDOACTION);
        S(SCI_SETMODEVENTMASK, (uptr_t)savedEventMask);

        if (undoWasOn) {
            S(SCI_SETUNDOCOLLECTION, 1);
            S(SCI_EMPTYUNDOBUFFER);
        }

        // Safety: clear any leftover indicator fragments across entire doc
        S(SCI_SETINDICATORCURRENT, (uptr_t)ind);
        S(SCI_INDICATORCLEARRANGE, 0, (sptr_t)S(SCI_GETLENGTH));

        ColumnTabs::CT_SetCurDocHasPads(hSci, false);
        return true;
    }

    bool ColumnTabs::CT_HasAlignedPadding(HWND hSci) noexcept
    {
        return CT_GetCurDocHasPads(hSci);
    }

    bool ColumnTabs::CT_ApplyNumericPadding(
        HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine)
    {
        const HWND hwnd = hSci;
        auto S = [hwnd](UINT msg, WPARAM w = 0, LPARAM l = 0)->sptr_t { return (sptr_t)::S(hwnd, msg, w, l); };
        auto ls = [&](int ln)->Sci_Position { return (Sci_Position)S(SCI_POSITIONFROMLINE, (uptr_t)ln); };
        auto le = [&](int ln)->Sci_Position { return (Sci_Position)S(SCI_GETLINEENDPOSITION, (uptr_t)ln); };
        auto gc = [&](Sci_Position p)->int { return (int)S(SCI_GETCHARAT, (uptr_t)p);  };

        const bool haveVec = !model.Lines.empty();
        const bool liveInfo = (model.getLineInfo != nullptr);
        if (!haveVec && !liveInfo) return false;

        const int baseDoc = (int)model.docStartLine;
        int l0 = (firstLine < 0) ? baseDoc : firstLine; // non-const
        int l1 = (lastLine < 0)
            ? (haveVec ? (l0 + (int)model.Lines.size() - 1) : l0)
            : ((lastLine < l0) ? l0 : lastLine);         // non-const

        // clamp to model bounds
        if (haveVec) {
            const int modelFirst = baseDoc;
            const int modelLast = baseDoc + (int)model.Lines.size() - 1;
            if (l0 < modelFirst) l0 = modelFirst;
            if (l1 > modelLast)  l1 = modelLast;
        }
        else
        {
            const int modelFirst = baseDoc;
            if (l0 < modelFirst) l0 = modelFirst;
            // (l1 is already normalized relative to l0; no further clamping needed here)
        }

        if (l1 < l0) return false;

        RedrawGuard rg(hSci);

        S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
        S(SCI_INDICSETSTYLE, (uptr_t)CT_GetIndicatorId(), INDIC_HIDDEN);
        S(SCI_INDICSETALPHA, (uptr_t)CT_GetIndicatorId(), 0);

        auto countDigits = [](const std::string& s, int& intD, int& fracD, bool& hasDec) {
            intD = fracD = 0; hasDec = false;
            size_t i = 0;
            if (i < s.size() && (s[i] == '+' || s[i] == '-')) ++i;
            for (; i < s.size() && isdigit((unsigned char)s[i]); ++i) ++intD;
            if (i < s.size() && (s[i] == '.' || s[i] == ',')) { hasDec = true; ++i; }
            for (; i < s.size() && isdigit((unsigned char)s[i]); ++i) ++fracD;
            };

        // PASS 1: collect maxima
        size_t maxCols = 0;
        for (int ln = l0; ln <= l1; ++ln) {
            const auto L = ColumnTabs::detail::fetchLine(model, ln);
            maxCols = (std::max)(maxCols, L.FieldCount());
        }
        if (maxCols == 0) return true;

        std::vector<int>  maxIntDigits(maxCols, 0);
        std::vector<int>  maxFracDigits(maxCols, 0);
        std::vector<bool> colHasDec(maxCols, false);

        for (int ln = l0; ln <= l1; ++ln) {
            const auto L = ColumnTabs::detail::fetchLine(model, ln);
            const size_t nField = L.FieldCount();
            const size_t nDelim = L.delimiterOffsets.size();
            const Sci_Position base = ls(ln);

            for (size_t c = 0; c < nField; ++c) {
                if (c >= nDelim + 1) continue;

                Sci_Position s = (c == 0) ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1]
                    + (Sci_Position)model.delimiterLength;
                Sci_Position e = (c < nDelim)
                    ? base + (Sci_Position)L.delimiterOffsets[c]
                    : le(ln);

                while (s < e) { int ch = gc(s);     if (ch == ' ' || ch == '\t') ++s; else break; }
                while (e > s) { int ch = gc(e - 1); if (ch == ' ' || ch == '\t') --e; else break; }
                if (e <= s) continue;

                Sci_Position tokStart{}; int leftPx = 0; bool hasDec = false;
                if (!CT_FindNumericToken(hSci, s, e, tokStart, leftPx, hasDec))
                    continue;

                std::string tok; tok.assign((size_t)(e - tokStart), '\0');
                Sci_TextRangeFull tr{}; tr.chrg.cpMin = tokStart; tr.chrg.cpMax = e; tr.lpstrText = tok.data();
                S(SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);
                while (!tok.empty() && (tok.back() == ' ' || tok.back() == '\t')) tok.pop_back();

                int intD = 0, fracD = 0; bool hd = false;
                countDigits(tok, intD, fracD, hd);

                if (intD > maxIntDigits[c])   maxIntDigits[c] = intD;
                if (fracD > maxFracDigits[c]) maxFracDigits[c] = fracD;
                if (hasDec || hd) colHasDec[c] = true;
            }
        }

        // PASS 2: apply padding and decimal alignment
        // Suppress undo and modification notifications during bulk padding.
        const bool undoWasOn = S(SCI_GETUNDOCOLLECTION) != 0;
        if (undoWasOn) S(SCI_SETUNDOCOLLECTION, 0);
        const int savedEventMask = (int)S(SCI_GETMODEVENTMASK);
        S(SCI_SETMODEVENTMASK, 0);
        S(SCI_BEGINUNDOACTION);
        bool madeAny = false;

        for (int ln = l0; ln <= l1; ++ln) {
            const auto L = ColumnTabs::detail::fetchLine(model, ln);
            const size_t nField = L.FieldCount();
            const size_t nDelim = L.delimiterOffsets.size();
            const Sci_Position base = ls(ln);
            Sci_Position delta = 0;

            for (size_t c = 0; c < nField; ++c) {
                if (c >= maxCols || maxIntDigits[c] <= 0) continue;
                if (c >= nDelim + 1) continue;

                Sci_Position s = (c == 0) ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1]
                    + (Sci_Position)model.delimiterLength + delta;
                Sci_Position e = (c < nDelim)
                    ? base + (Sci_Position)L.delimiterOffsets[c] + delta
                    : le(ln);

                while (s < e) { int ch = gc(s);     if (ch == ' ' || ch == '\t') ++s; else break; }
                while (e > s) { int ch = gc(e - 1); if (ch == ' ' || ch == '\t') --e; else break; }
                if (e <= s) continue;

                Sci_Position tokStart{}; int leftPx = 0; bool hasDec = false;
                if (!CT_FindNumericToken(hSci, s, e, tokStart, leftPx, hasDec))
                    continue;

                int intD = 0, fracD = 0; bool hd = false;
                {
                    std::string tok; tok.assign((size_t)(e - tokStart), '\0');
                    Sci_TextRangeFull tr{}; tr.chrg.cpMin = tokStart; tr.chrg.cpMax = e; tr.lpstrText = tok.data();
                    S(SCI_GETTEXTRANGEFULL, 0, (sptr_t)&tr);
                    while (!tok.empty() && (tok.back() == ' ' || tok.back() == '\t')) tok.pop_back();
                    countDigits(tok, intD, fracD, hd);
                }
                if (intD <= 0 && !hasDec && !hd) continue;

                const Sci_Position fieldStart =
                    (c == 0)
                    ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1]
                    + (Sci_Position)model.delimiterLength + delta;

                int spacesTouching = 0;
                { Sci_Position p = tokStart; while (p > fieldStart && gc(p - 1) == ' ') { --p; ++spacesTouching; } }

                const bool hasSign = (tokStart < e) && (gc(tokStart) == '+' || gc(tokStart) == '-');
                int needUnits = 1 + (maxIntDigits[c] - intD);
                if (needUnits < 0) needUnits = 0;
                if (hasSign && needUnits > 0) --needUnits;

                int appliedShift = 0;
                const int diff = needUnits - spacesTouching;

                if (diff > 0) {
                    std::string pad((size_t)diff, ' ');
                    S(SCI_INSERTTEXT, (uptr_t)tokStart, (sptr_t)pad.c_str());
                    S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
                    S(SCI_INDICATORFILLRANGE, (uptr_t)tokStart, (sptr_t)diff);
                    delta += (Sci_Position)diff;
                    appliedShift += diff;
                    madeAny = true;
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
                        const Sci_Position lower = (runBeg < fieldStart) ? fieldStart : runBeg;
                        delAvail = (int)(tokStart - lower);
                    }
                    if (delAvail > 0) {
                        if (toDel > delAvail) toDel = delAvail;
                        const Sci_Position posDel = tokStart - (Sci_Position)toDel;
                        S(SCI_INDICATORCLEARRANGE, (uptr_t)posDel, (sptr_t)toDel);
                        S(SCI_DELETERANGE, (uptr_t)posDel, (sptr_t)toDel);
                        delta -= (Sci_Position)toDel;
                        appliedShift -= toDel;
                        madeAny = true;
                    }
                }

                if (colHasDec[c] && (maxFracDigits[c] > 0) && (fracD == 0)) {
                    const Sci_Position tokEnd = tokStart + (Sci_Position)appliedShift
                        + (Sci_Position)((hasSign ? 1 : 0) + intD);
                    std::string pad = "."; pad.append((size_t)maxFracDigits[c], '0');
                    S(SCI_INSERTTEXT, (uptr_t)tokEnd, (sptr_t)pad.c_str());
                    S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
                    S(SCI_INDICATORFILLRANGE, (uptr_t)tokEnd, (sptr_t)pad.size());
                    delta += (Sci_Position)pad.size();
                    madeAny = true;
                }
            }
        }

        S(SCI_ENDUNDOACTION);
        S(SCI_SETMODEVENTMASK, (uptr_t)savedEventMask);
        if (undoWasOn) {
            S(SCI_SETUNDOCOLLECTION, 1);
            if (madeAny) S(SCI_EMPTYUNDOBUFFER);
        }
        if (madeAny) CT_SetCurDocHasPads(hSci, true);
        return true;
    }

    // -------------------------------------------------------------------------
    // Visual API (non-destructive; manages Scintilla tab stops)
    // -------------------------------------------------------------------------
    bool CT_ApplyFlowTabStops(HWND hSci,
        const CT_ColumnModelView& model,
        int firstLine,
        int lastLine,
        int paddingPx /*pixels*/)
    {
        using namespace ColumnTabs::detail;

        ensureCapacity(hSci);

        const bool hasVec = !model.Lines.empty();
        if (!hasVec && !model.getLineInfo) return false;

        int line0 = firstLine;
        int effectiveLast = firstLine;

        if (hasVec) {
            const int modelFirst = static_cast<int>(model.docStartLine);
            const int modelLast = modelFirst + static_cast<int>(model.Lines.size()) - 1;

            if (line0 < modelFirst) line0 = modelFirst;

            if (lastLine < 0) {
                effectiveLast = modelLast;
            }
            else {
                effectiveLast = (lastLine < line0) ? line0 : lastLine;
                if (effectiveLast > modelLast) effectiveLast = modelLast;
            }
        }
        else {
            const int modelFirst = static_cast<int>(model.docStartLine);
            if (line0 < modelFirst) line0 = modelFirst;

            if (lastLine < 0) {
                const int docLast = (int)S(hSci, SCI_GETLINECOUNT, 0, 0) - 1;
                effectiveLast = (docLast < line0) ? line0 : docLast;
            }
            else {
                effectiveLast = (lastLine < line0) ? line0 : lastLine;
            }
        }

        if (line0 > effectiveLast) return false;

        const int gapPx = (paddingPx > 0) ? paddingPx : 0;

        std::vector<int> stops;
        if (!computeStopsFromWidthsPx(hSci, model, line0, effectiveLast, stops, gapPx))
            return false;
        if (stops.empty())
            return false;

        for (int ln = line0; ln <= effectiveLast; ++ln) {
            if ((int)g_hasETSLine.size() > ln &&
                (int)g_savedManualStopsPx.size() > ln &&
                g_hasETSLine[(size_t)ln] == 0u)
            {
                g_savedManualStopsPx[(size_t)ln] = collectTabStopsPx(hSci, ln);
            }
        }

        setTabStopsRangePx(hSci, line0, effectiveLast, stops);

        for (int ln = line0; ln <= effectiveLast; ++ln) {
            if ((int)g_hasETSLine.size() > ln)
                g_hasETSLine[(size_t)ln] = 1u;
        }

        return true;
    }

    bool CT_DisableFlowTabStops(HWND hSci, bool restoreManual)
    {
        using namespace detail;

        if (!ColumnTabs::CT_HasFlowTabStops())
            return true;

        ensureCapacity(hSci);

        RedrawGuard rd(hSci);

        const int total = (int)S(hSci, SCI_GETLINECOUNT, 0, 0);
        const int limit = (std::min)(total, (int)g_hasETSLine.size());

        for (int ln = 0; ln < limit; ++ln) {
            if (!g_hasETSLine[(size_t)ln]) continue;

            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);

            if (restoreManual) {
                const auto& manual = g_savedManualStopsPx[(size_t)ln];
                for (int px : manual)
                    S(hSci, SCI_ADDTABSTOP, (uptr_t)ln, (sptr_t)px);
            }

            g_hasETSLine[(size_t)ln] = 0u;
            g_savedManualStopsPx[(size_t)ln].clear();
        }
        return true;
    }

    bool CT_ClearAllTabStops(HWND hSci)
    {
        const int total = (int)S(hSci, SCI_GETLINECOUNT);
        OptionalRedrawGuard rg(hSci, /*opCount=*/(size_t)total);
        for (int ln = 0; ln < total; ++ln)
            S(hSci, SCI_CLEARTABSTOPS, (uptr_t)ln, 0);
        return true;
    }

    void CT_ResetFlowVisualState() noexcept
    {
        detail::g_hasETSLine.clear();
        detail::g_savedManualStopsPx.clear();
    }

    bool ColumnTabs::CT_HasFlowTabStops() noexcept
    {
        using namespace detail;
        for (size_t i = 0, n = g_hasETSLine.size(); i < n; ++i) {
            if (g_hasETSLine[i]) return true;
        }
        return false;
    }

    // -------------------------------------------------------------------------
    // Utilities (kept for callers)
    // -------------------------------------------------------------------------
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

    // -------------------------------------------------------------------------
    // Per-document state
    // -------------------------------------------------------------------------
    void ColumnTabs::CT_SetDocHasPads(sptr_t docPtr, bool has) noexcept {
        if (has) detail::g_docHasPads[docPtr] = true;
        else     detail::g_docHasPads.erase(docPtr);
    }

    bool ColumnTabs::CT_GetDocHasPads(sptr_t docPtr) noexcept {
        auto it = detail::g_docHasPads.find(docPtr);
        return it != detail::g_docHasPads.end();
    }

    void ColumnTabs::CT_SetCurDocHasPads(HWND hSci, bool has) noexcept {
        const sptr_t doc = (sptr_t)::S(hSci, SCI_GETDOCPOINTER, 0, 0);
        if (!doc) return;
        CT_SetDocHasPads(doc, has);
    }

    bool ColumnTabs::CT_GetCurDocHasPads(HWND hSci) noexcept {
        const sptr_t doc = (sptr_t)::S(hSci, SCI_GETDOCPOINTER, 0, 0);
        if (!doc) return false;
        return CT_GetDocHasPads(doc);
    }

    // -------------------------------------------------------------------------
    // Cleanup
    // -------------------------------------------------------------------------
    bool ColumnTabs::CT_CleanupVisuals(HWND hSci) noexcept
    {
        if (!hSci) return false;
        CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);
        CT_ResetFlowVisualState();
        return true;
    }

    bool ColumnTabs::CT_CleanupAllForDoc(HWND hSci) noexcept
    {
        if (!hSci) return false;

        if (ColumnTabs::CT_GetCurDocHasPads(hSci)) {
            CT_RemoveAlignedPadding(hSci);
        }
        if (ColumnTabs::CT_HasFlowTabStops()) {
            CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);
        }
        CT_ResetFlowVisualState();
        return true;
    }

} // namespace ColumnTabs