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

#include "TandemDock.h"

#include <algorithm>
#include <climits>
#include <cstdlib>
#include <dwmapi.h>

#pragma comment(lib, "dwmapi.lib")

namespace tandem_dock {

    // =====================================================================
    //  Win32 helpers
    // =====================================================================

    RECT getVisualBounds(HWND hwnd)
    {
        RECT r{};
        if (SUCCEEDED(DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
            &r, sizeof(r)))) {
            return r;
        }
        GetWindowRect(hwnd, &r);
        return r;
    }

    RECT getMonitorWorkArea(HWND hwnd)
    {
        HMONITOR hMon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{};
        mi.cbSize = sizeof(mi);
        if (GetMonitorInfo(hMon, &mi)) return mi.rcWork;
        RECT r{ 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN) };
        return r;
    }

    namespace {
        struct UnionEnumCtx {
            RECT probe;       // in
            RECT unionRc;     // out - initialized empty
            bool any;         // out - at least one hit
        };
        BOOL CALLBACK unionMonitorProc(HMONITOR hMon, HDC, LPRECT, LPARAM lp)
        {
            auto* ctx = reinterpret_cast<UnionEnumCtx*>(lp);
            MONITORINFO mi{}; mi.cbSize = sizeof(mi);
            if (!GetMonitorInfo(hMon, &mi)) return TRUE;
            // Does this monitor intersect the probe rect?
            RECT ix{};
            if (!IntersectRect(&ix, &ctx->probe, &mi.rcMonitor)) return TRUE;
            // Accumulate its work area into the union.
            if (!ctx->any) {
                ctx->unionRc = mi.rcWork;
                ctx->any = true;
            }
            else {
                UnionRect(&ctx->unionRc, &ctx->unionRc, &mi.rcWork);
            }
            return TRUE;
        }
    }

    RECT getWorkAreaUnionForRect(const RECT& rect)
    {
        UnionEnumCtx ctx{};
        ctx.probe = rect;
        EnumDisplayMonitors(nullptr, nullptr, unionMonitorProc,
            reinterpret_cast<LPARAM>(&ctx));
        if (ctx.any) return ctx.unionRc;

        // Fallback: rect is off-screen or enumeration failed. Use the
        // monitor the rect's center is nearest to.
        const POINT c{ (rect.left + rect.right) / 2,
                       (rect.top + rect.bottom) / 2 };
        HMONITOR hMon = MonitorFromPoint(c, MONITOR_DEFAULTTONEAREST);
        MONITORINFO mi{}; mi.cbSize = sizeof(mi);
        if (GetMonitorInfo(hMon, &mi)) return mi.rcWork;
        RECT r{ 0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN) };
        return r;
    }

    ShadowOffsets getShadowOffsets(HWND hwnd)
    {
        RECT full{}; GetWindowRect(hwnd, &full);
        const RECT vis = getVisualBounds(hwnd);
        return { static_cast<int>(vis.left - full.left),
                 static_cast<int>(vis.top - full.top),
                 static_cast<int>(full.right - vis.right),
                 static_cast<int>(full.bottom - vis.bottom) };
    }

    bool mouseButtonDown()
    {
        return (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
    }

    // =====================================================================
    //  Layout solvers (one per dock edge)
    // =====================================================================
    //
    //  All three solvers work in VISIBLE coordinates end-to-end, then
    //  convert to OUTER coords at the very end by expanding each side
    //  by the matching shadow offset. That way the dock seam between
    //  client and host visually lines up (no gap, no overlap), which
    //  is what the user actually sees.
    //
    //  The "primary axis" is the axis along which the client's length
    //  varies:
    //    Bottom : vertical   (client height)
    //    Right  : horizontal (client width)
    //    Left   : horizontal (client width)

    namespace {

        // Minimum final length (outer px) below which we refuse to emit
        // a host shrink (cosmetic sanity - prevents degenerate windows).
        constexpr int kMinHostShrinkPrimary = 50;

        // At vertical dock edges (Right / Left), the visible client edge
        // is placed one pixel INSIDE the host's visible edge. This closes
        // the single-pixel column that a strict flush join would leave
        // open due to Windows' standard exclusive-right RECT convention
        // (the column at x == host.visible.right is "after" the host but
        // "before" the client, so both paint nothing there). Bottom dock
        // does not need this because the corresponding pixel row falls
        // inside the client region either way.
        constexpr int kVerticalEdgeOverlapPx = 1;

        DockLayout solveBottomImpl(const DockInputs& in)
        {
            DockLayout out{};
            out.flushAnchorEdge = DockEdge::Bottom;
            out.flushAnchorVis = in.hostVisible.bottom;

            // Secondary axis (horizontal): client matches host width.
            const int clientLeftVis = in.hostVisible.left;
            const int clientRightVis = in.hostVisible.right;

            // Primary axis (vertical): anchor at host.bottom, grow downward.
            const int anchorTopVis = in.hostVisible.bottom;
            const int bottomLimit = in.workArea.bottom;
            const int shadowV = in.clientShadow.top + in.clientShadow.bottom;
            // Convert desired OUTER primary -> VISIBLE for math against limits.
            const int desiredVisH = (std::max)(in.desiredClientPrimary,
                in.minClientPrimary) - shadowV;
            const int minVisH = in.minClientPrimary - shadowV;

            int topVis = anchorTopVis;
            int visH = desiredVisH;
            if (topVis + visH > bottomLimit) {
                const int available = bottomLimit - topVis;
                if (available >= minVisH) {
                    visH = available;
                }
                else {
                    visH = minVisH;
                    topVis = bottomLimit - visH;
                    out.clientOverlappedHost = true;
                }
            }
            if (topVis < in.workArea.top) topVis = in.workArea.top;

            // VISIBLE -> OUTER.
            out.clientOuter.left = clientLeftVis - in.clientShadow.left;
            out.clientOuter.right = clientRightVis + in.clientShadow.right;
            out.clientOuter.top = topVis - in.clientShadow.top;
            out.clientOuter.bottom = topVis + visH + in.clientShadow.bottom;

            if (out.clientOverlappedHost) {
                // Target: host.visible.bottom = client.visible.top.
                const int targetBottomVis = topVis;
                int newTopVis = in.hostVisible.top;
                int newHeightVis = targetBottomVis - newTopVis;
                if (newHeightVis < in.minHostPrimary) {
                    newTopVis = (std::max)(static_cast<int>(in.workArea.top),
                        targetBottomVis - in.minHostPrimary);
                    newHeightVis = targetBottomVis - newTopVis;
                }
                if (newHeightVis > kMinHostShrinkPrimary &&
                    newTopVis >= in.workArea.top)
                {
                    out.hostNeedsShrink = true;
                    out.hostOuterTarget.left = in.hostFull.left;
                    out.hostOuterTarget.right = in.hostFull.right;
                    out.hostOuterTarget.top = newTopVis - in.hostShadow.top;
                    out.hostOuterTarget.bottom = targetBottomVis + in.hostShadow.bottom;
                }
            }
            return out;
        }

        DockLayout solveRightImpl(const DockInputs& in)
        {
            DockLayout out{};
            out.flushAnchorEdge = DockEdge::Right;
            out.flushAnchorVis = in.hostVisible.right - kVerticalEdgeOverlapPx;

            // Primary axis (horizontal): anchor one pixel INSIDE host's
            // visible right edge (see kVerticalEdgeOverlapPx). Any residual
            // sub-pixel error is corrected by nudgeClientToFlush().
            const int anchorLeftVis = in.hostVisible.right - kVerticalEdgeOverlapPx;
            const int rightLimit = in.workArea.right;
            const int shadowH = in.clientShadow.left + in.clientShadow.right;
            const int desiredVisW = (std::max)(in.desiredClientPrimary,
                in.minClientPrimary) - shadowH;
            const int minVisW = in.minClientPrimary - shadowH;

            int leftVis = anchorLeftVis;
            int visW = desiredVisW;
            if (leftVis + visW > rightLimit) {
                const int available = rightLimit - leftVis;
                if (available >= minVisW) {
                    visW = available;
                }
                else {
                    visW = minVisW;
                    leftVis = rightLimit - visW;
                    out.clientOverlappedHost = true;
                }
            }
            if (leftVis < in.workArea.left) leftVis = in.workArea.left;

            out.clientOuter.top = in.hostVisible.top - in.clientShadow.top;
            out.clientOuter.bottom = in.hostVisible.bottom + in.clientShadow.bottom;
            out.clientOuter.left = leftVis - in.clientShadow.left;
            out.clientOuter.right = leftVis + visW + in.clientShadow.right;

            if (out.clientOverlappedHost) {
                const int targetRightVis = leftVis;
                int newLeftVis = in.hostVisible.left;
                int newWidthVis = targetRightVis - newLeftVis;
                if (newWidthVis < in.minHostPrimary) {
                    newLeftVis = (std::max)(static_cast<int>(in.workArea.left),
                        targetRightVis - in.minHostPrimary);
                    newWidthVis = targetRightVis - newLeftVis;
                }
                if (newWidthVis > kMinHostShrinkPrimary &&
                    newLeftVis >= in.workArea.left)
                {
                    out.hostNeedsShrink = true;
                    out.hostOuterTarget.top = in.hostFull.top;
                    out.hostOuterTarget.bottom = in.hostFull.bottom;
                    out.hostOuterTarget.left = newLeftVis - in.hostShadow.left;
                    out.hostOuterTarget.right = targetRightVis + in.hostShadow.right;
                }
            }
            return out;
        }

        DockLayout solveLeftImpl(const DockInputs& in)
        {
            DockLayout out{};
            out.flushAnchorEdge = DockEdge::Left;
            out.flushAnchorVis = in.hostVisible.left + kVerticalEdgeOverlapPx;

            // Primary axis (horizontal): anchor one pixel INSIDE host's
            // visible left edge (see kVerticalEdgeOverlapPx).
            const int anchorRightVis = in.hostVisible.left + kVerticalEdgeOverlapPx;
            const int leftLimit = in.workArea.left;
            const int shadowH = in.clientShadow.left + in.clientShadow.right;
            const int desiredVisW = (std::max)(in.desiredClientPrimary,
                in.minClientPrimary) - shadowH;
            const int minVisW = in.minClientPrimary - shadowH;

            int rightVis = anchorRightVis;
            int visW = desiredVisW;
            if (rightVis - visW < leftLimit) {
                const int available = rightVis - leftLimit;
                if (available >= minVisW) {
                    visW = available;
                }
                else {
                    visW = minVisW;
                    rightVis = leftLimit + visW;
                    out.clientOverlappedHost = true;
                }
            }
            if (rightVis > in.workArea.right) rightVis = in.workArea.right;

            out.clientOuter.top = in.hostVisible.top - in.clientShadow.top;
            out.clientOuter.bottom = in.hostVisible.bottom + in.clientShadow.bottom;
            out.clientOuter.left = rightVis - visW - in.clientShadow.left;
            out.clientOuter.right = rightVis + in.clientShadow.right;

            if (out.clientOverlappedHost) {
                const int targetLeftVis = rightVis;
                int newRightVis = in.hostVisible.right;
                int newWidthVis = newRightVis - targetLeftVis;
                if (newWidthVis < in.minHostPrimary) {
                    newRightVis = (std::min)(static_cast<int>(in.workArea.right),
                        targetLeftVis + in.minHostPrimary);
                    newWidthVis = newRightVis - targetLeftVis;
                }
                if (newWidthVis > kMinHostShrinkPrimary &&
                    newRightVis <= in.workArea.right)
                {
                    out.hostNeedsShrink = true;
                    out.hostOuterTarget.top = in.hostFull.top;
                    out.hostOuterTarget.bottom = in.hostFull.bottom;
                    out.hostOuterTarget.left = targetLeftVis - in.hostShadow.left;
                    out.hostOuterTarget.right = newRightVis + in.hostShadow.right;
                }
            }
            return out;
        }

    } // anonymous namespace

    DockLayout solveDock(const DockInputs& in)
    {
        switch (in.edge) {
        case DockEdge::Bottom: return solveBottomImpl(in);
        case DockEdge::Right:  return solveRightImpl(in);
        case DockEdge::Left:   return solveLeftImpl(in);
        }
        return {};
    }

    // =====================================================================
    //  Magnetic snap + edge detection
    // =====================================================================

    SnapResult snapMovingRectToHost(RECT* pClientRect,
        const RECT& hostVis,
        const ShadowOffsets& clientShadow,
        int thresholdPx)
    {
        SnapResult res{};
        if (!pClientRect) return res;

        // Each candidate is scored by the |delta| along its primary
        // axis. The best (smallest, within threshold, secondary-axis
        // overlap present) wins.
        //
        // Per edge: OUTER-coord offset that produces a flush VISIBLE
        // join to the host (the solver uses the same math):
        //   Bottom : client.outer.top    = host.vis.bottom - client.shadow.top
        //   Right  : client.outer.left   = host.vis.right - 1 - client.shadow.left
        //   Left   : client.outer.right  = host.vis.left  + 1 + client.shadow.right

        struct Candidate {
            int      dx;
            int      dy;
            int      dist;
            DockEdge edge;
            bool     overlapOk;
        };
        Candidate cands[3];

        // Bottom
        {
            const int targetOuterTop = hostVis.bottom - clientShadow.top;
            const int delta = targetOuterTop - pClientRect->top;
            const int visLeft = pClientRect->left + clientShadow.left;
            const int visRight = pClientRect->right - clientShadow.right;
            cands[0] = { 0, delta, std::abs(delta), DockEdge::Bottom,
                         visRight > hostVis.left && visLeft < hostVis.right };
        }
        // Right
        {
            const int targetOuterLeft = hostVis.right - 1 - clientShadow.left;
            const int delta = targetOuterLeft - pClientRect->left;
            const int visTop = pClientRect->top + clientShadow.top;
            const int visBottom = pClientRect->bottom - clientShadow.bottom;
            cands[1] = { delta, 0, std::abs(delta), DockEdge::Right,
                         visBottom > hostVis.top && visTop < hostVis.bottom };
        }
        // Left
        {
            const int targetOuterRight = hostVis.left + 1 + clientShadow.right;
            const int delta = targetOuterRight - pClientRect->right;
            const int visTop = pClientRect->top + clientShadow.top;
            const int visBottom = pClientRect->bottom - clientShadow.bottom;
            cands[2] = { delta, 0, std::abs(delta), DockEdge::Left,
                         visBottom > hostVis.top && visTop < hostVis.bottom };
        }

        int bestIdx = -1;
        int bestDist = INT_MAX;
        for (int i = 0; i < 3; ++i) {
            if (cands[i].overlapOk && cands[i].dist <= thresholdPx
                && cands[i].dist < bestDist) {
                bestDist = cands[i].dist;
                bestIdx = i;
            }
        }
        if (bestIdx < 0) return res;

        OffsetRect(pClientRect, cands[bestIdx].dx, cands[bestIdx].dy);
        res.applied = true;
        res.edge = cands[bestIdx].edge;
        return res;
    }

    DockEdge pickEdgeByProximity(const RECT& clientVis,
        const RECT& hostVis,
        DockEdge fallback)
    {
        const int dBottom = std::abs(static_cast<int>(hostVis.bottom) - clientVis.top);
        const int dRight = std::abs(static_cast<int>(hostVis.right) - clientVis.left);
        const int dLeft = std::abs(static_cast<int>(hostVis.left) - clientVis.right);

        const bool hOverlap =
            clientVis.right > hostVis.left && clientVis.left < hostVis.right;
        const bool vOverlap =
            clientVis.bottom > hostVis.top && clientVis.top < hostVis.bottom;

        DockEdge bestEdge = fallback;
        int      bestDist = INT_MAX;
        if (hOverlap && dBottom < bestDist) { bestDist = dBottom; bestEdge = DockEdge::Bottom; }
        if (vOverlap && dRight < bestDist) { bestDist = dRight;  bestEdge = DockEdge::Right; }
        if (vOverlap && dLeft < bestDist) { bestDist = dLeft;   bestEdge = DockEdge::Left; }
        return bestEdge;
    }

    DockEdge pickNearestEdge(const RECT& clientVis, const RECT& hostVis)
    {
        // Squared distance from a point to a host edge SEGMENT (not
        // infinite line). Used instead of simple perpendicular distance
        // so a client that's, say, far above N++'s right edge doesn't
        // get credited for a zero-distance Right pick just because it
        // happens to be aligned on X.

        auto clampedSqDist = [](int px, int py,
            int ax, int ay, int bx, int by) -> long long
            {
                // Project (px,py) onto segment (ax,ay)-(bx,by), clamp [0,1].
                const long long ABx = bx - ax;
                const long long ABy = by - ay;
                const long long APx = px - ax;
                const long long APy = py - ay;
                const long long ab2 = ABx * ABx + ABy * ABy;
                long long t_num = APx * ABx + APy * ABy;
                if (t_num < 0) t_num = 0;
                if (ab2 > 0 && t_num > ab2) t_num = ab2;
                const double t = (ab2 > 0) ? static_cast<double>(t_num) / ab2 : 0.0;
                const double cx = ax + t * ABx;
                const double cy = ay + t * ABy;
                const double dx = px - cx;
                const double dy = py - cy;
                return static_cast<long long>(dx * dx + dy * dy);
            };

        const int cx = (clientVis.left + clientVis.right) / 2;
        const int cy = (clientVis.top + clientVis.bottom) / 2;

        const long long dBottom = clampedSqDist(cx, cy,
            hostVis.left, hostVis.bottom,
            hostVis.right, hostVis.bottom);
        const long long dRight = clampedSqDist(cx, cy,
            hostVis.right, hostVis.top,
            hostVis.right, hostVis.bottom);
        const long long dLeft = clampedSqDist(cx, cy,
            hostVis.left, hostVis.top,
            hostVis.left, hostVis.bottom);

        // Pick smallest. Ties go to Bottom (traditional default), then
        // Right, then Left.
        DockEdge best = DockEdge::Bottom;
        long long bestD = dBottom;
        if (dRight < bestD) { bestD = dRight; best = DockEdge::Right; }
        if (dLeft < bestD) { best = DockEdge::Left; }
        return best;
    }

    // =====================================================================
    //  Self-calibrating flush correction
    // =====================================================================
    //
    // After a SetWindowPos call the client's visible anchor edge may
    // have landed a pixel or two off the intended target because:
    //
    //   * DwmGetWindowAttribute(DWMWA_EXTENDED_FRAME_BOUNDS) does not
    //     adjust for DPI, while GetWindowRect does (MSDN). On 125% /
    //     150% scaling the difference between the two is prone to
    //     half-pixel rounding.
    //   * DWM frame-bounds have historically varied by Windows build
    //     (Win10 vs Win11, different updates) in how the bottom-right
    //     edges are reported.
    //
    // Rather than trying to predict these offsets algebraically -
    // which is what a +1/-1 magic constant would do, and which breaks
    // on the next Windows version - we simply measure: after placing
    // the client, read its actual visible bounds and nudge it by
    // whatever correction is needed so its visible anchor lines up
    // exactly. One nudge is enough in practice; the calibration is
    // stable between frames.

    void nudgeClientToFlush(HWND clientHwnd, const DockLayout& layout)
    {
        if (!clientHwnd || !IsWindow(clientHwnd)) return;

        const RECT actualVis = getVisualBounds(clientHwnd);
        int  actualEdge = 0;
        switch (layout.flushAnchorEdge) {
        case DockEdge::Bottom: actualEdge = actualVis.top;   break;
        case DockEdge::Right:  actualEdge = actualVis.left;  break;
        case DockEdge::Left:   actualEdge = actualVis.right; break;
        }
        const int err = layout.flushAnchorVis - actualEdge;
        if (err == 0) return;

        // Apply correction to OUTER rect (that's what SetWindowPos speaks).
        RECT outer{};
        if (!GetWindowRect(clientHwnd, &outer)) return;

        int dx = 0, dy = 0;
        switch (layout.flushAnchorEdge) {
        case DockEdge::Bottom: dy = err; break;
        case DockEdge::Right:  dx = err; break;
        case DockEdge::Left:   dx = err; break;
        }
        if (!dx && !dy) return;

        SetWindowPos(clientHwnd, nullptr,
            outer.left + dx, outer.top + dy,
            outer.right - outer.left,
            outer.bottom - outer.top,
            SWP_NOZORDER | SWP_NOACTIVATE);
    }

} // namespace tandem_dock