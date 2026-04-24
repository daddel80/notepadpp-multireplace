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

// =====================================================================
//  TandemDock
// =====================================================================
//  A small, self-contained Win32 library for "magnetic multi-edge
//  docking" of one window to another. Given two HWNDs and a dock
//  edge (Bottom / Right / Left), the library computes the correct
//  geometry so the docked window sticks to the chosen edge and
//  follows the host as it is moved or resized.
//
//  Features:
//    * Axis-agnostic solver - Bottom / Right / Left share one
//      underlying algorithm via a config struct.
//    * Winamp-style magnetic snap during user drag (WM_MOVING).
//    * Proximity-based edge detection after a drag is released.
//    * DWM-aware shadow handling so docked edges visually flush.
//    * Pure layout functions: no side effects, no hidden state -
//      unit-testable and easy to reuse elsewhere.
//
//  Usage pattern (typical host plugin):
//    1. On enable: snapshot both HWNDs' outer rects for restore.
//    2. In a timer (or a host-move notification): call
//       solveDock() with the current host rect + desired primary
//       length. Place the client HWND at the returned rect.
//    3. In WM_MOVING on the client: call snapMovingRectToHost()
//       and apply the returned edge / modified rect.
//    4. In WM_EXITSIZEMOVE: call pickEdgeByProximity() to decide
//       the final dock edge, update desired length from the
//       user's drag, re-run solveDock().
//
//  The library deliberately avoids owning any state. All state
//  lives with the host (dock edge, desired length, saved rects).
// =====================================================================

#pragma once

#include <windows.h>

namespace tandem_dock {

    // -------------------------------------------------------------------
    //  Constants
    // -------------------------------------------------------------------

    // Default Winamp/winsnap magnet distance: the client snaps when
    // its edge is within this many pixels of a host edge. One
    // threshold is enough because callers using the winsnap "free
    // rect" reconstruction model measure distance against the
    // cursor-driven position, not the held snapped rect - which
    // makes release behaviour speed-independent.
    constexpr int kDefaultSnapThresholdPx = 30;

    // -------------------------------------------------------------------
    //  Types
    // -------------------------------------------------------------------

    // Which side of the host the client docks to.
    enum class DockEdge { Bottom, Right, Left };

    // Offsets between OUTER (GetWindowRect) and VISIBLE (DWM) bounds.
    // Windows 10/11 put an invisible resize border around most windows;
    // these four values let callers convert between the two coord systems.
    struct ShadowOffsets
    {
        int left = 0;
        int top = 0;
        int right = 0;
        int bottom = 0;
    };

    // Inputs for solveDock(). All RECTs in monitor-relative screen pixels.
    struct DockInputs
    {
        DockEdge       edge;            // which host edge to dock to
        RECT           workArea;        // monitor work area (DWM visible)
        RECT           hostVisible;     // host's visible rect
        RECT           hostFull;        // host's outer rect (with shadow)
        ShadowOffsets  hostShadow;      // host's shadow offsets
        ShadowOffsets  clientShadow;    // client's shadow offsets

        // Length the user wants for the client along the edge's primary
        // axis (height for Bottom; width for Right/Left). Outer coords.
        int  desiredClientPrimary = 0;

        // Smallest length the OS would enforce anyway for the client
        // along the primary axis (from WM_GETMINMAXINFO, outer coords).
        int  minClientPrimary = 0;

        // Minimum length the host must retain along the primary axis
        // when the solver has to shrink it (visible pixels). Typical
        // value: 200 px to keep the host usable.
        int  minHostPrimary = 200;
    };

    // Result of a single solveDock() call.
    struct DockLayout
    {
        // Final client outer rect, ready for SetWindowPos.
        RECT clientOuter = {};

        // After placing clientOuter via SetWindowPos, the caller should
        // verify that the client's visible anchor edge ended up at
        // `flushAnchorVis` and nudge the window if it didn't. This is
        // the robust way to get a seamless dock across DPI scaling,
        // DWM version differences, and window-style variations - none
        // of which a purely geometric solver can know about in advance.
        //
        // The member `flushAnchorEdge` names which of the client's
        // VISIBLE edges must equal `flushAnchorVis`:
        //   Bottom dock -> top    (client.visible.top    == flushAnchorVis)
        //   Right  dock -> left   (client.visible.left   == flushAnchorVis)
        //   Left   dock -> right  (client.visible.right  == flushAnchorVis)
        DockEdge flushAnchorEdge = DockEdge::Bottom;
        int      flushAnchorVis = 0;

        // True if the solver has determined the host must be pinched
        // along the primary axis because the client had to overlap it.
        // When true, callers typically apply this only on mouse-up;
        // during live drag a "wall" rect (host shrunk only along the
        // dock edge) is usually preferred to avoid fighting Windows.
        bool hostNeedsShrink = false;
        RECT hostOuterTarget = {};

        // Diagnostic flag: true if the clamp pulled the client over
        // the host because no room was left on the monitor.
        bool clientOverlappedHost = false;
    };

    // Result of snapMovingRectToHost().
    struct SnapResult
    {
        bool      applied = false;   // true if pRect was modified
        DockEdge  edge = DockEdge::Bottom;  // valid iff applied
    };

    // -------------------------------------------------------------------
    //  Win32 helpers (can be used standalone)
    // -------------------------------------------------------------------

    // Visible bounds of hwnd via DWM, falling back to GetWindowRect.
    RECT getVisualBounds(HWND hwnd);

    // Work area (monitor minus taskbar/app bars) of the monitor that
    // hosts the largest portion of hwnd. Falls back to primary screen.
    RECT getMonitorWorkArea(HWND hwnd);

    // Union of the work areas of all monitors the rect intersects.
    // Used when a host straddles a monitor seam - gives a single
    // logical area spanning both monitors rather than truncating.
    // Falls back to the primary monitor's work area if nothing
    // intersects (rect fully off-screen).
    RECT getWorkAreaUnionForRect(const RECT& rect);

    // Shadow offsets of hwnd (outer vs visible).
    ShadowOffsets getShadowOffsets(HWND hwnd);

    // Is the left mouse button currently held down?
    bool mouseButtonDown();

    // -------------------------------------------------------------------
    //  Layout solver
    // -------------------------------------------------------------------

    // Compute the client's final placement given the current host
    // geometry and dock configuration. Pure function; no side effects.
    DockLayout solveDock(const DockInputs& in);

    // -------------------------------------------------------------------
    //  Magnetic snap during user drag
    // -------------------------------------------------------------------

    // WM_MOVING handler helper. Tests whether any of the client's
    // edges is within `thresholdPx` of a matching host edge (with
    // secondary-axis overlap) and if so, shifts the client rect so
    // its VISIBLE edge meets the host VISIBLE edge flush.
    //
    // This is the classic Winamp/winsnap-style single-threshold
    // magnet. The caller is responsible for the "release on cursor
    // pull" behaviour: while currently snapped, pass a reconstructed
    // "free" rect (based on cursor offset) instead of the actual
    // snapped rect, so distance reflects where the user's cursor
    // really wants the window - not where the magnet is holding it.
    SnapResult snapMovingRectToHost(RECT* pClientRect,
        const RECT& hostVisible,
        const ShadowOffsets& clientShadow,
        int thresholdPx = kDefaultSnapThresholdPx);

    // WM_EXITSIZEMOVE helper. Decides the dock edge based on which
    // host edge the client ended up closest to (with secondary-axis
    // overlap required). Returns `fallback` if no edge qualifies -
    // callers typically pass the current dock edge, so tandem never
    // turns itself off by the user dragging too far.
    DockEdge pickEdgeByProximity(const RECT& clientVisible,
        const RECT& hostVisible,
        DockEdge fallback);

    // Initial-dock helper. Picks the host edge whose closest point is
    // nearest to the client's center. No overlap requirement, so it
    // works for any client position - typical use is first-ever
    // tandem toggle when no persisted preference exists. Top edge
    // is never returned (MR only docks Bottom/Right/Left).
    DockEdge pickNearestEdge(const RECT& clientVisible,
        const RECT& hostVisible);

    // After SetWindowPos, nudge the client so its VISIBLE anchor edge
    // matches exactly `layout.flushAnchorVis`. Closes residual 1 px
    // seams that arise from DPI rounding or DWM frame-bounds quirks
    // (DwmGetWindowAttribute is documented to not adjust for DPI,
    // unlike GetWindowRect, so the geometric solver - which mixes the
    // two - can land one pixel off on high-DPI displays).
    //
    // The caller passes its own HWND and a function pointer for
    // SetWindowPos so the library stays free of host-specific globals.
    // Typically called once after the main placement call.
    void nudgeClientToFlush(HWND clientHwnd, const DockLayout& layout);

} // namespace tandem_dock