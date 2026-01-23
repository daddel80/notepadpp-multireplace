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
#include "SciUndoGuard.h"
#include "NumericToken.h"
#include <algorithm>
#include <unordered_map>
#include <string_view>
#include <cctype> 
#include <chrono>
#include <sstream>

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

    // Compute tab-stop pixel positions from text widths (columnar alignment).
    static bool computeStopsFromWidthsPx(
        HWND hSci,
        const ColumnTabs::CT_ColumnModelView& model,
        int line0, int line1,
        std::vector<int>& stopPx,   // OUT
        int gapPx)
    {
        using namespace ColumnTabs;

        if (line1 < line0) std::swap(line0, line1);
        const int numLines = line1 - line0 + 1;

        const int gapBeforePx = (gapPx >= 0) ? gapPx : 0;
        const int gapAfterPx = 0;

        const int spacePx = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" ");
        const int minAdvancePx = (spacePx > 0) ? ((spacePx + 1) / 2) : 2;

        // Determine maximum field count
        size_t maxCols = 0;
        if (!model.Lines.empty()) {
            for (int ln = line0; ln <= line1; ++ln) {
                const size_t idx = static_cast<size_t>(ln - static_cast<int>(model.docStartLine));
                if (idx < model.Lines.size()) {
                    maxCols = (std::max)(maxCols, model.Lines[idx].FieldCount());
                }
            }
        }
        else if (model.getLineInfo) {
            for (int ln = line0; ln <= line1; ++ln) {
                const auto L = detail::fetchLine(model, ln);
                maxCols = (std::max)(maxCols, L.FieldCount());
            }
        }

        if (maxCols < 2) { stopPx.clear(); return true; }
        const size_t stopsCount = maxCols - 1;

        // Read entire document once
        const Sci_Position docLen = (Sci_Position)S(hSci, SCI_GETLENGTH);
        std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
        S(hSci, SCI_GETTEXT, (uptr_t)(docLen + 1), (sptr_t)fullText.data());
        fullText.resize(static_cast<size_t>(docLen));

        std::vector<Sci_Position> lineStarts(static_cast<size_t>(numLines));
        for (int ln = line0; ln <= line1; ++ln) {
            lineStarts[static_cast<size_t>(ln - line0)] =
                (Sci_Position)S(hSci, SCI_POSITIONFROMLINE, (uptr_t)ln);
        }

        // Width cache
        std::unordered_map<std::string, int> widthCache;
        widthCache.reserve(5000);

        auto measurePxCached = [&](const std::string& s) -> int {
            if (s.empty()) return 0;
            auto it = widthCache.find(s);
            if (it != widthCache.end()) {
                return it->second;
            }
            const int w = (int)S(hSci, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)s.c_str());
            const int result = (w > 0) ? w : (int)s.size() * ((spacePx > 0) ? spacePx : 8);
            widthCache[s] = result;
            return result;
            };

        // PASS 1: collect maxima per column
        std::vector<int> maxCellWidthPx(maxCols, 0);
        std::vector<int> maxDelimiterWidthPx(stopsCount, 0);
        struct LM { std::vector<int> cellW, delimW; int lastIdx = -1; int eolX = 0; };
        std::vector<LM> lines;
        lines.reserve(static_cast<size_t>(numLines));

        for (int ln = line0; ln <= line1; ++ln) {
            const size_t lineIdx = static_cast<size_t>(ln - line0);
            const auto L = detail::fetchLine(model, ln);
            const size_t nDel = L.delimiterOffsets.size();
            const size_t nFld = nDel + 1;

            LM lm{};
            if (nFld == 0) { lines.push_back(std::move(lm)); continue; }

            const Sci_Position base = lineStarts[lineIdx];
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
                if (e > s && static_cast<size_t>(s) < fullText.size()) {
                    size_t len = static_cast<size_t>(e - s);
                    if (static_cast<size_t>(s) + len > fullText.size())
                        len = fullText.size() - static_cast<size_t>(s);
                    cell = fullText.substr(static_cast<size_t>(s), len);
                }

                if (!cell.empty())
                    cell.erase(std::remove(cell.begin(), cell.end(), '\t'), cell.end());

                const int w = measurePxCached(cell);
                lm.cellW[k] = w;
                if (w > maxCellWidthPx[k]) maxCellWidthPx[k] = w;
            }

            // Delimiters
            if (!model.delimiterIsTab && model.delimiterLength > 0) {
                for (size_t d = 0; d < nDel && d < stopsCount; ++d) {
                    Sci_Position d0 = base + (Sci_Position)L.delimiterOffsets[d];
                    Sci_Position d1 = d0 + (Sci_Position)model.delimiterLength;

                    std::string del;
                    if (d1 > d0 && static_cast<size_t>(d0) < fullText.size()) {
                        size_t len = static_cast<size_t>(d1 - d0);
                        if (static_cast<size_t>(d0) + len > fullText.size())
                            len = fullText.size() - static_cast<size_t>(d0);
                        del = fullText.substr(static_cast<size_t>(d0), len);
                    }

                    const int dw = measurePxCached(del);
                    lm.delimW[d] = dw;
                    if (dw > maxDelimiterWidthPx[d]) maxDelimiterWidthPx[d] = dw;
                }
            }
            else {
                std::fill(lm.delimW.begin(), lm.delimW.end(), 0);
            }

            int eol = 0;
            for (int k = 0; k < lm.lastIdx; ++k) {
                eol += lm.cellW[(size_t)k] + gapBeforePx;
                eol += lm.delimW[(size_t)k] + gapAfterPx;
            }
            if (lm.lastIdx >= 0)
                eol += lm.cellW[(size_t)lm.lastIdx];
            lm.eolX = eol;

            lines.push_back(std::move(lm));
        }

        // PASS 2: preferred stops
        std::vector<int> stopPref(stopsCount, 0);
        {
            int acc = 0;
            for (size_t c = 0; c < stopsCount; ++c) {
                acc += maxCellWidthPx[c] + gapBeforePx + minAdvancePx;
                stopPref[c] = acc;
                acc += maxDelimiterWidthPx[c] + gapAfterPx;
            }
        }

        // PASS 3: EOL clamps
        std::vector<int> clamp(stopsCount, 0);
        for (const auto& lm : lines) {
            if (lm.lastIdx < 0) continue;
            for (size_t c = (size_t)lm.lastIdx; c < stopsCount; ++c) {
                const int capped = (lm.eolX < stopPref[c]) ? lm.eolX : stopPref[c];
                if (capped > clamp[c]) clamp[c] = capped;
            }
        }

        // Final stops
        stopPx.resize(stopsCount);
        int prevStop = 0;
        for (size_t c = 0; c < stopsCount; ++c) {
            int candidate = (stopPref[c] > clamp[c]) ? stopPref[c] : clamp[c];
            if (candidate <= prevStop)
                candidate = prevStop + minAdvancePx;
            stopPx[c] = candidate;
            prevStop = candidate;
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
// ColumnTabs – all public APIs
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

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 1: Compute tab stop positions
        // ══════════════════════════════════════════════════════════════════════
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

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 2: Save manual tab stops
        // ══════════════════════════════════════════════════════════════════════
        for (int ln = line0; ln <= line1; ++ln) {
            if ((int)g_hasETSLine.size() > ln &&
                (int)g_savedManualStopsPx.size() > ln &&
                g_hasETSLine[(size_t)ln] == 0u)
            {
                g_savedManualStopsPx[(size_t)ln] = collectTabStopsPx(hSci, ln);
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 3: Apply visual tab stops
        // ══════════════════════════════════════════════════════════════════════
        setTabStopsRangePx(hSci, line0, line1, stops);

        for (int ln = line0; ln <= line1; ++ln) {
            if ((int)g_hasETSLine.size() > ln)
                g_hasETSLine[(size_t)ln] = 1u;
        }

        if (opt.oneFlowTabOnly && model.delimiterIsTab)
            return true;

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 4: Read entire document
        // ══════════════════════════════════════════════════════════════════════
        auto fetch = [&](int ln)->CT_ColumnLineInfo {
            return detail::fetchLine(model, ln);
            };

        const Sci_Position docLen = (Sci_Position)Sci(SCI_GETLENGTH);
        std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
        Sci(SCI_GETTEXT, (uptr_t)(docLen + 1), (sptr_t)fullText.data());
        fullText.resize(static_cast<size_t>(docLen));

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 5: Collect existing indicator ranges
        // ══════════════════════════════════════════════════════════════════════
        const int ind = g_CT_IndicatorId;
        std::vector<std::pair<Sci_Position, Sci_Position>> existingIndicators;
        {
            Sci_Position pos = 0;
            while (pos < docLen) {
                if ((int)Sci(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)pos) != 0) {
                    const Sci_Position start = (Sci_Position)Sci(SCI_INDICATORSTART, (uptr_t)ind, (sptr_t)pos);
                    const Sci_Position end = (Sci_Position)Sci(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)pos);
                    if (end > start) {
                        existingIndicators.emplace_back(start, end - start);
                        pos = end;
                    }
                    else {
                        ++pos;
                    }
                }
                else {
                    const Sci_Position nextEnd = (Sci_Position)Sci(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)pos);
                    if (nextEnd > pos) {
                        pos = nextEnd;
                    }
                    else {
                        ++pos;
                    }
                }
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 6: Collect insertion points
        // ══════════════════════════════════════════════════════════════════════
        std::vector<Sci_Position> insertPositions;
        insertPositions.reserve(static_cast<size_t>((line1 - line0 + 1) * 10));

        for (int ln = line0; ln <= line1; ++ln) {
            const auto L = fetch(ln);
            const Sci_Position base = (Sci_Position)Sci(SCI_POSITIONFROMLINE, (uptr_t)ln);
            const Sci_Position lineEnd = (Sci_Position)Sci(SCI_GETLINEENDPOSITION, (uptr_t)ln);

            const size_t nDelims = L.delimiterOffsets.size();
            if (nDelims == 0) continue;

            for (size_t c = 0; c < nDelims; ++c) {
                const Sci_Position delimPos = base + (Sci_Position)L.delimiterOffsets[c];
                if (delimPos < base || delimPos > lineEnd) continue;
                if (static_cast<size_t>(delimPos) > fullText.size()) continue;

                Sci_Position wsStart = delimPos;
                while (wsStart > base && static_cast<size_t>(wsStart - 1) < fullText.size()) {
                    const char ch = fullText[static_cast<size_t>(wsStart - 1)];
                    if (ch == ' ' || ch == '\t') --wsStart; else break;
                }

                const bool keepExistingTab =
                    (wsStart < delimPos) &&
                    (static_cast<size_t>(delimPos - 1) < fullText.size()) &&
                    (fullText[static_cast<size_t>(delimPos - 1)] == '\t');

                if (!keepExistingTab) {
                    insertPositions.push_back(delimPos);
                }
            }
        }

        if (insertPositions.empty()) {
            if (outNothingToAlign) *outNothingToAlign = true;
            return true;
        }

        std::sort(insertPositions.begin(), insertPositions.end());

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 7: Build new text with inserted tabs
        // ══════════════════════════════════════════════════════════════════════
        std::string newText;
        newText.reserve(fullText.size() + insertPositions.size());

        std::vector<Sci_Position> newTabPositions;
        newTabPositions.reserve(insertPositions.size());

        Sci_Position copyFrom = 0;

        for (Sci_Position insertPos : insertPositions) {
            if (insertPos > copyFrom) {
                newText.append(fullText, static_cast<size_t>(copyFrom),
                    static_cast<size_t>(insertPos - copyFrom));
            }
            newTabPositions.push_back(static_cast<Sci_Position>(newText.size()));
            newText.push_back('\t');
            copyFrom = insertPos;
        }

        if (copyFrom < static_cast<Sci_Position>(fullText.size())) {
            newText.append(fullText, static_cast<size_t>(copyFrom),
                fullText.size() - static_cast<size_t>(copyFrom));
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 8: Adjust existing indicator positions (binary search)
        // ══════════════════════════════════════════════════════════════════════
        std::vector<std::pair<Sci_Position, Sci_Position>> adjustedExistingIndicators;
        adjustedExistingIndicators.reserve(existingIndicators.size());

        for (const auto& existingInd : existingIndicators) {
            const Sci_Position oldPos = existingInd.first;
            const Sci_Position len = existingInd.second;

            auto it = std::upper_bound(insertPositions.begin(), insertPositions.end(), oldPos);
            const Sci_Position offset = static_cast<Sci_Position>(it - insertPositions.begin());

            adjustedExistingIndicators.emplace_back(oldPos + offset, len);
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 9: Replace document and set indicators
        // ══════════════════════════════════════════════════════════════════════
        {
            SciUndoGuard undo(hSci);

            Sci(SCI_SETTEXT, 0, (sptr_t)newText.c_str());

            Sci(SCI_SETINDICATORCURRENT, g_CT_IndicatorId);
            Sci(SCI_INDICSETSTYLE, g_CT_IndicatorId, INDIC_HIDDEN);
            Sci(SCI_INDICSETALPHA, g_CT_IndicatorId, 0);

            for (const auto& adjInd : adjustedExistingIndicators) {
                Sci(SCI_INDICATORFILLRANGE, (uptr_t)adjInd.first, (sptr_t)adjInd.second);
            }

            for (Sci_Position tabPos : newTabPositions) {
                Sci(SCI_INDICATORFILLRANGE, (uptr_t)tabPos, (sptr_t)1);
            }
        }

        CT_SetCurDocHasPads(hSci, true);

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
        if (docLen == 0) return false;

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 1: Collect all indicator ranges (positions to remove)
        // ══════════════════════════════════════════════════════════════════════
        std::vector<std::pair<Sci_Position, Sci_Position>> ranges;
        ranges.reserve(1024);

        Sci_Position pos = 0;
        while (pos < docLen) {
            if ((int)S(SCI_INDICATORVALUEAT, (uptr_t)ind, (sptr_t)pos) != 0) {
                const Sci_Position start = (Sci_Position)S(SCI_INDICATORSTART, (uptr_t)ind, (sptr_t)pos);
                const Sci_Position end = (Sci_Position)S(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)pos);
                if (end > start) {
                    ranges.emplace_back(start, end);
                    pos = end;
                }
                else {
                    ++pos;
                }
            }
            else {
                const Sci_Position nextEnd = (Sci_Position)S(SCI_INDICATOREND, (uptr_t)ind, (sptr_t)pos);
                if (nextEnd > pos) {
                    pos = nextEnd;
                }
                else {
                    ++pos;
                }
            }
        }

        if (ranges.empty()) {
            ColumnTabs::CT_SetCurDocHasPads(hSci, false);
            return false;
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 2: Read entire document and build new text without padding
        // ══════════════════════════════════════════════════════════════════════
        std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
        S(SCI_GETTEXT, (uptr_t)(docLen + 1), (sptr_t)fullText.data());
        fullText.resize(static_cast<size_t>(docLen));

        // Calculate new size (original minus all padding)
        size_t totalPaddingBytes = 0;
        for (const auto& r : ranges) {
            totalPaddingBytes += static_cast<size_t>(r.second - r.first);
        }

        std::string newText;
        newText.reserve(fullText.size() - totalPaddingBytes);

        // Build new text by copying segments between padding ranges
        Sci_Position copyFrom = 0;
        for (const auto& r : ranges) {
            // Copy text before this padding range
            if (r.first > copyFrom) {
                newText.append(fullText, static_cast<size_t>(copyFrom),
                    static_cast<size_t>(r.first - copyFrom));
            }
            // Skip the padding range
            copyFrom = r.second;
        }
        // Copy remaining text after last padding
        if (copyFrom < docLen) {
            newText.append(fullText, static_cast<size_t>(copyFrom),
                static_cast<size_t>(docLen - copyFrom));
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 3: Replace entire document content at once
        // ══════════════════════════════════════════════════════════════════════
        // Replace document (uses SciUndoGuard for nesting support)
        {
            SciUndoGuard undo(hSci);

            // Clear all indicators first
            S(SCI_INDICATORCLEARRANGE, 0, (sptr_t)docLen);

            // Replace entire content
            S(SCI_SETTEXT, 0, (sptr_t)newText.c_str());
        }

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

        const bool haveVec = !model.Lines.empty();
        const bool liveInfo = (model.getLineInfo != nullptr);
        if (!haveVec && !liveInfo) return false;

        const int baseDoc = (int)model.docStartLine;
        int l0 = (firstLine < 0) ? baseDoc : firstLine;
        int l1 = (lastLine < 0)
            ? (haveVec ? (l0 + (int)model.Lines.size() - 1) : l0)
            : ((lastLine < l0) ? l0 : lastLine);

        if (haveVec) {
            const int modelFirst = baseDoc;
            const int modelLast = baseDoc + (int)model.Lines.size() - 1;
            if (l0 < modelFirst) l0 = modelFirst;
            if (l1 > modelLast)  l1 = modelLast;
        }
        else {
            const int modelFirst = baseDoc;
            if (l0 < modelFirst) l0 = modelFirst;
        }

        if (l1 < l0) return false;

        const int numLines = l1 - l0 + 1;

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 1: Read entire document + CACHE all line positions
        // ══════════════════════════════════════════════════════════════════════
        const Sci_Position docLen = (Sci_Position)S(SCI_GETLENGTH);
        std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
        S(SCI_GETTEXT, (uptr_t)(docLen + 1), (sptr_t)fullText.data());
        fullText.resize(static_cast<size_t>(docLen));

        // CACHE line positions
        std::vector<Sci_Position> lineStarts(static_cast<size_t>(numLines));
        std::vector<Sci_Position> lineEnds(static_cast<size_t>(numLines));
        for (int ln = l0; ln <= l1; ++ln) {
            const size_t idx = static_cast<size_t>(ln - l0);
            lineStarts[idx] = (Sci_Position)S(SCI_POSITIONFROMLINE, (uptr_t)ln);
            lineEnds[idx] = (Sci_Position)S(SCI_GETLINEENDPOSITION, (uptr_t)ln);
        }

        auto gc = [&](Sci_Position p) -> char {
            if (p < 0 || static_cast<size_t>(p) >= fullText.size()) return '\0';
            return fullText[static_cast<size_t>(p)];
            };

        auto countDigits = [](const std::string& s, int& intD, int& fracD, bool& hasDec) {
            intD = fracD = 0; hasDec = false;
            size_t i = 0;
            if (i < s.size() && (s[i] == '+' || s[i] == '-')) ++i;
            for (; i < s.size() && isdigit((unsigned char)s[i]); ++i) ++intD;
            if (i < s.size() && (s[i] == '.' || s[i] == ',')) { hasDec = true; ++i; }
            for (; i < s.size() && isdigit((unsigned char)s[i]); ++i) ++fracD;
            };

        auto findNumericInBuffer = [&](Sci_Position start, Sci_Position end,
            Sci_Position& tokStart, bool& hasDec) -> bool {
                hasDec = false;
                tokStart = start;
                while (tokStart < end && (gc(tokStart) == ' ' || gc(tokStart) == '\t'))
                    ++tokStart;
                if (tokStart >= end) return false;
                Sci_Position pos = tokStart;
                char ch = gc(pos);
                if (ch == '+' || ch == '-') ++pos;
                Sci_Position digitStart = pos;
                while (pos < end && isdigit((unsigned char)gc(pos))) ++pos;
                if (pos == digitStart) return false;
                ch = gc(pos);
                if (pos < end && (ch == '.' || ch == ',')) {
                    hasDec = true;
                    ++pos;
                    while (pos < end && isdigit((unsigned char)gc(pos))) ++pos;
                }
                return true;
            };

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 2: Collect maxima per column
        // ══════════════════════════════════════════════════════════════════════
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
            const size_t lineIdx = static_cast<size_t>(ln - l0);
            const auto L = ColumnTabs::detail::fetchLine(model, ln);
            const size_t nField = L.FieldCount();
            const size_t nDelim = L.delimiterOffsets.size();
            const Sci_Position base = lineStarts[lineIdx];
            const Sci_Position lineEnd = lineEnds[lineIdx];

            for (size_t c = 0; c < nField && c < maxCols; ++c) {
                if (c >= nDelim + 1) continue;

                Sci_Position s = (c == 0) ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1] + (Sci_Position)model.delimiterLength;
                Sci_Position e = (c < nDelim)
                    ? base + (Sci_Position)L.delimiterOffsets[c]
                    : lineEnd;

                while (s < e && (gc(s) == ' ' || gc(s) == '\t')) ++s;
                while (e > s && (gc(e - 1) == ' ' || gc(e - 1) == '\t')) --e;
                if (e <= s) continue;

                Sci_Position tokStart;
                bool hasDec;
                if (!findNumericInBuffer(s, e, tokStart, hasDec))
                    continue;

                std::string tok;
                if (e > tokStart && static_cast<size_t>(tokStart) < fullText.size()) {
                    size_t len = static_cast<size_t>(e - tokStart);
                    if (static_cast<size_t>(tokStart) + len > fullText.size())
                        len = fullText.size() - static_cast<size_t>(tokStart);
                    tok = fullText.substr(static_cast<size_t>(tokStart), len);
                }
                while (!tok.empty() && (tok.back() == ' ' || tok.back() == '\t'))
                    tok.pop_back();

                int intD = 0, fracD = 0;
                bool hd = false;
                countDigits(tok, intD, fracD, hd);

                if (intD > maxIntDigits[c])   maxIntDigits[c] = intD;
                if (fracD > maxFracDigits[c]) maxFracDigits[c] = fracD;
                if (hasDec || hd) colHasDec[c] = true;
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 3: Collect all edits
        // ══════════════════════════════════════════════════════════════════════
        struct EditOp {
            Sci_Position pos = 0;
            std::string insert;
        };

        std::vector<EditOp> edits;
        edits.reserve(static_cast<size_t>(numLines * maxCols));

        for (int ln = l0; ln <= l1; ++ln) {
            const size_t lineIdx = static_cast<size_t>(ln - l0);
            const auto L = ColumnTabs::detail::fetchLine(model, ln);
            const size_t nField = L.FieldCount();
            const size_t nDelim = L.delimiterOffsets.size();
            const Sci_Position base = lineStarts[lineIdx];
            const Sci_Position lineEnd = lineEnds[lineIdx];

            for (size_t c = 0; c < nField && c < maxCols; ++c) {
                if (maxIntDigits[c] <= 0) continue;
                if (c >= nDelim + 1) continue;

                Sci_Position fieldStart = (c == 0) ? base
                    : base + (Sci_Position)L.delimiterOffsets[c - 1] + (Sci_Position)model.delimiterLength;
                Sci_Position fieldEnd = (c < nDelim)
                    ? base + (Sci_Position)L.delimiterOffsets[c]
                    : lineEnd;

                Sci_Position s = fieldStart, e = fieldEnd;
                while (s < e && (gc(s) == ' ' || gc(s) == '\t')) ++s;
                while (e > s && (gc(e - 1) == ' ' || gc(e - 1) == '\t')) --e;
                if (e <= s) continue;

                Sci_Position tokStart;
                bool hasDec;
                if (!findNumericInBuffer(s, e, tokStart, hasDec))
                    continue;

                std::string tok;
                if (e > tokStart && static_cast<size_t>(tokStart) < fullText.size()) {
                    size_t len = static_cast<size_t>(e - tokStart);
                    if (static_cast<size_t>(tokStart) + len > fullText.size())
                        len = fullText.size() - static_cast<size_t>(tokStart);
                    tok = fullText.substr(static_cast<size_t>(tokStart), len);
                }
                while (!tok.empty() && (tok.back() == ' ' || tok.back() == '\t'))
                    tok.pop_back();

                int intD = 0, fracD = 0;
                bool hd = false;
                countDigits(tok, intD, fracD, hd);
                if (intD <= 0 && !hasDec && !hd) continue;

                int spacesTouching = 0;
                {
                    Sci_Position p = tokStart;
                    while (p > fieldStart && gc(p - 1) == ' ') {
                        --p;
                        ++spacesTouching;
                    }
                }

                const bool hasSign = (tokStart < e) && (gc(tokStart) == '+' || gc(tokStart) == '-');
                int needUnits = 1 + (maxIntDigits[c] - intD);
                if (needUnits < 0) needUnits = 0;
                if (hasSign && needUnits > 0) --needUnits;

                const int diff = needUnits - spacesTouching;

                if (diff > 0) {
                    EditOp op;
                    op.pos = tokStart;
                    op.insert = std::string(static_cast<size_t>(diff), ' ');
                    edits.push_back(op);
                }

                if (colHasDec[c] && (maxFracDigits[c] > 0) && (fracD == 0)) {
                    Sci_Position tokEnd = tokStart + (hasSign ? 1 : 0) + intD;
                    EditOp op;
                    op.pos = tokEnd;
                    op.insert = ".";
                    op.insert.append(static_cast<size_t>(maxFracDigits[c]), '0');
                    edits.push_back(op);
                }
            }
        }

        if (edits.empty()) return true;

        std::sort(edits.begin(), edits.end(),
            [](const EditOp& a, const EditOp& b) { return a.pos < b.pos; });

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 4: Build new text with all insertions
        // ══════════════════════════════════════════════════════════════════════
        std::string newText;
        newText.reserve(fullText.size() + edits.size() * 4);

        std::vector<std::pair<Sci_Position, Sci_Position>> indicatorRanges;
        indicatorRanges.reserve(edits.size());

        Sci_Position copyFrom = 0;

        for (const auto& edit : edits) {
            if (edit.pos > copyFrom) {
                newText.append(fullText, static_cast<size_t>(copyFrom),
                    static_cast<size_t>(edit.pos - copyFrom));
            }
            indicatorRanges.emplace_back(
                static_cast<Sci_Position>(newText.size()),
                static_cast<Sci_Position>(edit.insert.size()));
            newText.append(edit.insert);
            copyFrom = edit.pos;
        }

        if (copyFrom < static_cast<Sci_Position>(fullText.size())) {
            newText.append(fullText, static_cast<size_t>(copyFrom),
                fullText.size() - static_cast<size_t>(copyFrom));
        }

        // ══════════════════════════════════════════════════════════════════════
        // PHASE 5: Replace document and set indicators
        // ══════════════════════════════════════════════════════════════════════
        // Replace document and set indicators (uses SciUndoGuard for nesting support)
        {
            SciUndoGuard undo(hSci);

            S(SCI_SETTEXT, 0, (sptr_t)newText.c_str());

            S(SCI_SETINDICATORCURRENT, (uptr_t)CT_GetIndicatorId());
            S(SCI_INDICSETSTYLE, (uptr_t)CT_GetIndicatorId(), INDIC_HIDDEN);
            S(SCI_INDICSETALPHA, (uptr_t)CT_GetIndicatorId(), 0);

            for (const auto& range : indicatorRanges) {
                S(SCI_INDICATORFILLRANGE, (uptr_t)range.first, (sptr_t)range.second);
            }
        }

        if (!indicatorRanges.empty()) {
            CT_SetCurDocHasPads(hSci, true);
        }

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
        detail::g_hasETSLine.shrink_to_fit();
        detail::g_savedManualStopsPx.clear();
        detail::g_savedManualStopsPx.shrink_to_fit();
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
