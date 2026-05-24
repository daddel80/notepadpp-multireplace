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

// FlowTabs: tab-strip UI for the per-tab list model (build, layout,
// scrolling, tooltips, dirty-state, switching). Split out of
// MultiReplacePanel.cpp; these remain MultiReplace members.

#define NOMINMAX
#include "MultiReplacePanel.h"
#include "UndoRedoManager.h"
#include <windows.h>
#include <uxtheme.h>
#include <string>
#include <algorithm>

// Local singleton alias, matching MultiReplacePanel.cpp, so the
// moved function bodies stay byte-for-byte identical.
static UndoRedoManager& URM = UndoRedoManager::instance();

std::wstring MultiReplace::truncateTabName(const std::wstring& name, size_t maxChars)
{
    if (name.length() <= maxChars) return name;
    if (maxChars <= 1) return std::wstring(L"\u2026");
    return name.substr(0, maxChars - 1) + L"\u2026";
}

// Truncated name plus dirty bullet.
std::wstring MultiReplace::buildTabLabel(int tabIndex) const
{
    if (tabIndex < 0 || tabIndex >= static_cast<int>(_tabs.size())) return {};
    std::wstring label = truncateTabName(_tabs[tabIndex]->name,
        static_cast<size_t>(tabMaxLength));
    if (_tabs[tabIndex]->isDirty) {
        label += L" \u25CF";  // U+25CF BLACK CIRCLE
    }
    return label;
}

void MultiReplace::rebuildTabControl()
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    // Subclass the tab control so we can intercept double-click for
    // inline rename. Windows' default tab control does not emit
    // NM_DBLCLK in a usable way here. Idempotent: SetWindowSubclass
    // with the same id+proc just updates the ref data.
    SetWindowSubclass(hTab, &MultiReplace::tabControlSubclassProc, 1,
        reinterpret_cast<DWORD_PTR>(this));

    // Clear existing tab items.
    TabCtrl_DeleteAllItems(hTab);

    // Compact layout tuning.
    TabCtrl_SetPadding(hTab, 6, 3);
    TabCtrl_SetMinTabWidth(hTab, 40);

    // Insert one tab item per TabState. Label = truncated name plus
    // dirty bullet (built by buildTabLabel). The bullet inherits tab
    // text color, so light/dark mode is handled automatically.
    for (size_t i = 0; i < _tabs.size(); ++i) {
        std::wstring label = buildTabLabel(static_cast<int>(i));
        TCITEM item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<LPWSTR>(label.c_str());
        TabCtrl_InsertItem(hTab, static_cast<int>(i), &item);
    }

    // Select the active tab.
    if (!_tabs.empty()) {
        int sel = _activeTabIndex;
        if (sel < 0) sel = 0;
        if (sel >= static_cast<int>(_tabs.size())) sel = static_cast<int>(_tabs.size()) - 1;
        TabCtrl_SetCurSel(hTab, sel);
    }

    // Reposition the "+" button BEFORE forcing the tab control repaint,
    // otherwise the button still occupies its old slot when the new tab
    // gets painted, briefly clipping the left half of the new label.
    repositionNewTabButton();

    // The first reposition pass uses the natural tab widths reported
    // by the common control immediately after insertion. After the
    // tab control has been resized the second pass observes the
    // stable post-layout state, so any width drift caused by the
    // resize itself (e.g. coming back from a previous overflow with
    // a wider tab control) is corrected before the user sees it.
    repositionNewTabButton();

    // After a rebuild the common control places the active tab
    // somewhere visible - often centered when newly inserted at the
    // end. We need two corrections here:
    //
    // 1. If everything fits but Win32 has scrolled the strip left
    //    (because SetCurSel forced the active tab visible), snap the
    //    strip back to tab 0 so earlier tabs are reachable.
    //
    // 2. If the strip genuinely overflows and the active tab is the
    //    last one, scroll right until the new tab is flush against
    //    the right edge.
    //
    // Technique adapted from Notepad++'s TabBar.cpp: scroll the tab
    // control directly with SB_THUMBPOSITION and the desired tab
    // index. This is undocumented but stable - the WM_HSCROLL goes
    // to the tab control itself, not to its internal UpDown.
    if (!_tabs.empty()) {
        const int count = TabCtrl_GetItemCount(hTab);
        RECT rcFirst, rcLast, rcClient;
        if (count > 0
            && GetClientRect(hTab, &rcClient)
            && TabCtrl_GetItemRect(hTab, 0, &rcFirst)
            && TabCtrl_GetItemRect(hTab, count - 1, &rcLast))
        {
            const int viewRight = rcClient.right - rcClient.left;
            const int stripWidth = rcLast.right - rcFirst.left;
            const bool everythingFits = (stripWidth <= viewRight);

            if (everythingFits && rcFirst.left < 0) {
                // Snap strip back to start.
                SendMessage(hTab, WM_HSCROLL,
                    MAKEWPARAM(SB_THUMBPOSITION, 0), 0);
                repositionNewTabButton();
            }
            else if (!everythingFits
                && _activeTabIndex == count - 1
                && rcLast.right > viewRight)
            {
                // Push strip right so the new last tab is visible.
                SendMessage(hTab, WM_HSCROLL,
                    MAKEWPARAM(SB_THUMBPOSITION, count - 1), 0);
                repositionNewTabButton();
            }
        }
    }

    // Honor collapsed-list state; otherwise SW_SHOW leaves a tab-strip
    // remnant on first open.
    const bool listShown = useListEnabled || keepListVisible;
    ShowWindow(hTab, listShown ? SW_SHOW : SW_HIDE);
    if (listShown) {
        InvalidateRect(hTab, nullptr, TRUE);
        UpdateWindow(hTab);
    }
    updateTabTooltip(_activeTabIndex);
}

// Keep "+" adjacent to the last tab; overflow "v" shows only when clipped.
void MultiReplace::repositionNewTabButton()
{
    HWND hPlus = GetDlgItem(_hSelf, IDC_NEW_LIST_BUTTON);
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    HWND hDrop = GetDlgItem(_hSelf, IDC_TAB_LIST_DROPDOWN);
    HWND hLeft = GetDlgItem(_hSelf, IDC_TAB_SCROLL_LEFT);
    HWND hRight = GetDlgItem(_hSelf, IDC_TAB_SCROLL_RIGHT);
    if (!hPlus || !hTab || !hDrop || !hLeft || !hRight) return;

    if (!useListEnabled && !keepListVisible) return;

    RECT panelClient;
    GetClientRect(_hSelf, &panelClient);
    const int W = panelClient.right - panelClient.left;

    const int tabLeftBase = sx(14) - sx(2);
    const int useListLeft = W - sx(40);
    const int btnW = sx(22);                  // "+" / dropdown
    const int dotW = sx(18);                  // ellipsis indicator
    const int btnH = sy(22);
    const int gap = sx(2);
    const int reservedSlot = dotW + gap;       // slot kept free for "..."

    // Anchor positions, derived RIGHT-TO-LEFT from the UseList toggle:
    //   ... [tab control] [...slot] [+] [drop] | UseList
    // Used both for the right-anchored layout and as a clamp for the
    // tabs-hugging layout, so "+" never crosses into the toggle.
    const int dropLeft = useListLeft - gap - btnW;
    const int plusLeftAnchor = dropLeft - gap - btnW;
    const int slotLeftAnchor = plusLeftAnchor - gap - dotW;
    const int tabRightCap = slotLeftAnchor - gap;
    const int maxTabCtrlWidth = std::max(tabRightCap - tabLeftBase, sx(60));

    RECT cur;
    GetWindowRect(hTab, &cur);
    POINT curTL = { cur.left, cur.top };
    ScreenToClient(_hSelf, &curTL);
    const int tabHeight = cur.bottom - cur.top;
    const int curWidth = cur.right - cur.left;

    const int tabCount = TabCtrl_GetItemCount(hTab);

    // Sum of natural tab widths is independent of scroll position.
    int totalTabsWidth = 0;
    int firstLeft = 0;
    for (int i = 0; i < tabCount; ++i) {
        RECT rc;
        if (!TabCtrl_GetItemRect(hTab, i, &rc)) continue;
        if (i == 0) firstLeft = std::max<int>(rc.left, 0);
        totalTabsWidth += (rc.right - rc.left);
    }
    const int naturalStripWidth = firstLeft + totalTabsWidth;

    const bool overflow = (tabCount > 0) &&
        (naturalStripWidth + sx(2) > maxTabCtrlWidth);

    // Which side has tabs scrolled out of view. Only meaningful while
    // overflowing.
    bool clippedLeft = false;
    bool clippedRight = false;
    if (overflow) {
        RECT rc0, rcN;
        if (TabCtrl_GetItemRect(hTab, 0, &rc0)) {
            clippedLeft = (rc0.left < 0);
        }
        if (TabCtrl_GetItemRect(hTab, tabCount - 1, &rcN)) {
            RECT clientRc;
            GetClientRect(hTab, &clientRc);
            clippedRight = (rcN.right > (clientRc.right - clientRc.left));
        }
    }

    // Left "..." sits before the tab control and steals real space
    // from it, shifting the strip right by (dotW + gap).
    const int leftDotPad = (overflow && clippedLeft) ? (dotW + gap) : 0;
    const int tabLeftRel = tabLeftBase + leftDotPad;

    // Hug the tabs unless the strip overflows on the right. The
    // overflow-only-on-the-left case still hugs: any empty space
    // right of the last tab would be dead tab-control area that
    // swallows hover and clicks but visually reads as tab-strip.
    int newTabCtrlWidth;
    if (tabCount == 0) {
        newTabCtrlWidth = sx(60);
    }
    else if (!overflow) {
        newTabCtrlWidth = std::min(naturalStripWidth + sx(2),
            maxTabCtrlWidth - leftDotPad);
    }
    else if (!clippedRight) {
        RECT rcLast;
        const int lastRight = TabCtrl_GetItemRect(hTab, tabCount - 1, &rcLast)
            ? rcLast.right
            : (maxTabCtrlWidth - leftDotPad);
        newTabCtrlWidth = std::min(lastRight + sx(2),
            maxTabCtrlWidth - leftDotPad);
    }
    else {
        newTabCtrlWidth = maxTabCtrlWidth - leftDotPad;
    }

    if (curTL.x != tabLeftRel || curWidth != newTabCtrlWidth) {
        SetWindowPos(hTab, nullptr,
            tabLeftRel, curTL.y, newTabCtrlWidth, tabHeight,
            SWP_NOZORDER | SWP_NOACTIVATE);
        InvalidateRect(hTab, nullptr, TRUE);
        UpdateWindow(hTab);
    }

    // Hide the tab control's built-in UpDown spinner. It gets
    // re-created whenever the strip overflows; we replace its
    // function with our "..." indicators. Move it off-screen rather
    // than SW_HIDE so it stays addressable for the programmatic
    // scroll messages used by ensureTabVisible/scrollTabStrip.
    if (HWND hUpDown = FindWindowExW(hTab, nullptr, UPDOWN_CLASS, nullptr)) {
        RECT udRc;
        GetWindowRect(hUpDown, &udRc);
        if (udRc.left > -1000) {
            SetWindowPos(hUpDown, nullptr, -10000, -10000, 0, 0,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSIZE);
            InvalidateRect(hTab, nullptr, TRUE);
            UpdateWindow(hTab);
        }
    }

    const int yMid = curTL.y + (tabHeight - btnH) / 2;

    // "+" position. The decisive question is "does the last tab end
    // inside the viewport". When yes (including under technical
    // overflow with tabs hidden on the left), glue "+" to it. When
    // no, anchor "+" at the right cap so the "..." slot sits between
    // the last visible tab and "+".
    int plusXRel;
    if (tabCount == 0) {
        plusXRel = tabLeftRel + gap + reservedSlot;
    }
    else if (overflow && clippedRight) {
        plusXRel = plusLeftAnchor;
    }
    else {
        // Read the post-resize right edge so we layout against the
        // settled geometry. rebuildTabControl deliberately calls us
        // twice for the same reason.
        RECT rcLast;
        const int lastRight = TabCtrl_GetItemRect(hTab, tabCount - 1, &rcLast)
            ? rcLast.right
            : naturalStripWidth;
        plusXRel = tabLeftRel + lastRight + gap + reservedSlot;
        if (plusXRel > plusLeftAnchor) plusXRel = plusLeftAnchor;
    }

    // Side indicators. Both only appear under overflow, on the side
    // where tabs are clipped.
    if (overflow && clippedLeft) {
        SetWindowPos(hLeft, HWND_TOP, tabLeftBase, yMid, dotW, btnH, SWP_NOACTIVATE);
        ShowWindow(hLeft, SW_SHOWNA);
    }
    else {
        ShowWindow(hLeft, SW_HIDE);
    }

    if (overflow && clippedRight) {
        SetWindowPos(hRight, HWND_TOP, plusXRel - gap - dotW, yMid, dotW, btnH, SWP_NOACTIVATE);
        ShowWindow(hRight, SW_SHOWNA);
    }
    else {
        ShowWindow(hRight, SW_HIDE);
    }

    SetWindowPos(hPlus, HWND_TOP, plusXRel, yMid, btnW, btnH, SWP_NOACTIVATE);

    // Dropdown is shown whenever there are tabs the user cannot all
    // see at once - this is the same condition as `overflow`.
    if (overflow) {
        SetWindowPos(hDrop, HWND_TOP, dropLeft, yMid, btnW, btnH, SWP_NOACTIVATE);
        ShowWindow(hDrop, SW_SHOWNA);
    }
    else {
        ShowWindow(hDrop, SW_HIDE);
    }
}

// Scroll the strip so the given tab is fully visible; no-op if already on.
void MultiReplace::ensureTabVisible(int tabIndex)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    const int count = TabCtrl_GetItemCount(hTab);
    if (tabIndex < 0 || tabIndex >= count) return;

    RECT clientRc;
    GetClientRect(hTab, &clientRc);
    const int viewW = clientRc.right - clientRc.left;

    // Skip the UpDown scroll pass when the target is already in view
    // (touching the UpDown can reflow the strip), but still drop
    // through to repositionNewTabButton: Win32 may have scrolled the
    // strip itself, flipping clippedRight and shifting the geometry.
    bool alreadyVisible = false;
    RECT rcTarget;
    if (TabCtrl_GetItemRect(hTab, tabIndex, &rcTarget)) {
        alreadyVisible = (rcTarget.left >= 0 && rcTarget.right <= viewW);
    }

    HWND hUpDown = alreadyVisible ? nullptr : FindWindowExW(hTab, nullptr, UPDOWN_CLASS, nullptr);
    if (!alreadyVisible && !hUpDown) return;

    if (!alreadyVisible) {
        // Bring the UpDown back into a valid position so it can accept
        // synthesized clicks. Mirror what scrollTabStrip does.
        RECT udRect;
        GetWindowRect(hUpDown, &udRect);
        const int udW = (udRect.right - udRect.left) > 0 ? (udRect.right - udRect.left) : 38;
        const int udH = (udRect.bottom - udRect.top) > 0 ? (udRect.bottom - udRect.top) : 18;
        const int defX = clientRc.right - udW;

        SetWindowPos(hUpDown, nullptr, defX, 0, udW, udH,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

        RECT udClient;
        GetClientRect(hUpDown, &udClient);
        const int halfW = (udClient.right - udClient.left) / 2;
        const int clickY = (udClient.bottom - udClient.top) / 2;

        const int safetyMax = count + 2;
        for (int step = 0; step < safetyMax; ++step) {
            RECT rc;
            if (!TabCtrl_GetItemRect(hTab, tabIndex, &rc)) break;
            if (rc.left >= 0 && rc.right <= viewW) break;

            const int clickX = (rc.left < 0) ? halfW / 2 : halfW + halfW / 2;
            const LPARAM mouseLP = MAKELPARAM(clickX, clickY);
            SendMessage(hUpDown, WM_LBUTTONDOWN, MK_LBUTTON, mouseLP);
            SendMessage(hUpDown, WM_LBUTTONUP, 0, mouseLP);
        }

        SetWindowPos(hUpDown, nullptr, -10000, -10000, udW, udH,
            SWP_NOZORDER | SWP_NOACTIVATE);
        InvalidateRect(hTab, nullptr, TRUE);
        UpdateWindow(hTab);
    }
    else {
        // Win32 scrolls via ScrollWindow without invalidating, leaving
        // stale pixels (half-clipped tab labels) behind. Force redraw.
        InvalidateRect(hTab, nullptr, TRUE);
        UpdateWindow(hTab);
    }

    repositionNewTabButton();

    // Tab tooltips bind to fixed rects at TTM_ADDTOOL time; re-register
    // since the strip may have scrolled.
    updateTabTooltip(_activeTabIndex);
}

// Scroll the strip by one tab (negative left, positive right).
void MultiReplace::scrollTabStrip(int direction)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    HWND hUpDown = FindWindowExW(hTab, nullptr, UPDOWN_CLASS, nullptr);
    if (!hUpDown) return;

    // Move the UpDown back to a real position in the tab control
    // briefly so it can process synthesized clicks. Off-screen or
    // hidden UpDowns ignore mouse input. We capture and restore its
    // current rect to leave it as we found it.
    RECT udRect;
    GetWindowRect(hUpDown, &udRect);
    POINT udTL = { udRect.left, udRect.top };
    ScreenToClient(hTab, &udTL);
    const int udW = udRect.right - udRect.left;
    const int udH = udRect.bottom - udRect.top;

    // Place at a sensible default if its current position is invalid
    // (off-screen). Top-right corner of the tab strip is what the
    // common control would have used naturally.
    RECT tabClient;
    GetClientRect(hTab, &tabClient);
    const int defW = (udW > 0) ? udW : 38;   // typical width
    const int defH = (udH > 0) ? udH : 18;
    const int defX = tabClient.right - defW;
    const int defY = 0;

    SetWindowPos(hUpDown, nullptr, defX, defY, defW, defH,
        SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

    // The UpDown spinner is a horizontal pair of arrow buttons.
    // Left half = scroll left, right half = scroll right. Synthesize
    // a left-button down/up at the matching half.
    RECT udClient;
    GetClientRect(hUpDown, &udClient);
    const int halfW = (udClient.right - udClient.left) / 2;
    const int clickY = (udClient.bottom - udClient.top) / 2;
    const int clickX = (direction < 0) ? halfW / 2 : halfW + halfW / 2;

    const LPARAM mouseLP = MAKELPARAM(clickX, clickY);
    SendMessage(hUpDown, WM_LBUTTONDOWN, MK_LBUTTON, mouseLP);
    SendMessage(hUpDown, WM_LBUTTONUP, 0, mouseLP);

    // Move the UpDown back off-screen so it stays invisible to the user.
    SetWindowPos(hUpDown, nullptr, -10000, -10000, defW, defH,
        SWP_NOZORDER | SWP_NOACTIVATE);

    InvalidateRect(hTab, nullptr, TRUE);
    UpdateWindow(hTab);

    repositionNewTabButton();

    // Tooltips bind to fixed rects at TTM_ADDTOOL time, so after the
    // strip scrolled the tooltip-to-tab mapping is stale. Re-register.
    updateTabTooltip(_activeTabIndex);
}

// Popup listing all tabs (overflow dropdown).
void MultiReplace::showTabListPopup()
{
    if (_tabs.empty()) return;

    HWND hDrop = GetDlgItem(_hSelf, IDC_TAB_LIST_DROPDOWN);
    if (!hDrop) return;

    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) return;

    // Reserve a range comfortably above the rest of the command IDs
    // and high enough to never collide with menu identifiers used
    // elsewhere. Dynamic per-popup; not declared as a permanent ID.
    constexpr UINT kBaseId = 0xC000;

    const int activeIdx = _activeTabIndex;
    const int count = static_cast<int>(_tabs.size());
    for (int i = 0; i < count; ++i) {
        std::wstring label = _tabs[i]->name;
        if (_tabs[i]->isDirty) label += L" \u2022";

        UINT flags = MF_STRING;
        if (i == activeIdx) flags |= MF_CHECKED;

        AppendMenuW(hMenu, flags, kBaseId + static_cast<UINT>(i),
            label.c_str());
    }

    RECT rc{};
    GetWindowRect(hDrop, &rc);
    const POINT anchor = { rc.left, rc.bottom };

    const UINT cmd = TrackPopupMenu(hMenu,
        TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_NONOTIFY,
        anchor.x, anchor.y, 0, _hSelf, nullptr);

    DestroyMenu(hMenu);

    if (cmd >= kBaseId) {
        const int picked = static_cast<int>(cmd - kBaseId);
        if (picked >= 0 && picked < count && picked != activeIdx) {
            switchToTab(picked);
        }
    }
}

// Dirty-indicator management. markActiveTabDirty diffs the current list
// hash against the saved baseline (so undo back to saved clears it);
// clearTabDirty force-clears after a save. Both rebuild only on change.
void MultiReplace::markActiveTabDirty()
{
    if (_activeTabIndex < 0 ||
        _activeTabIndex >= static_cast<int>(_tabs.size())) return;

    TabState& tab = *_tabs[_activeTabIndex];

    // Content-based dirty detection: the tab is dirty iff the current
    // list content differs from what was last saved or loaded. Undoing
    // back to the saved state therefore clears the flag automatically.
    const std::size_t currentHash = computeListHash(replaceListData);
    const bool shouldBeDirty = (currentHash != tab.originalHash);

    if (shouldBeDirty == tab.isDirty) return;  // no change
    tab.isDirty = shouldBeDirty;

    // Rebuild so the bullet suffix appears or disappears as needed.
    rebuildTabControl();
}

void MultiReplace::clearTabDirty(int tabIndex)
{
    if (tabIndex < 0 ||
        tabIndex >= static_cast<int>(_tabs.size())) return;

    TabState& tab = *_tabs[tabIndex];
    if (!tab.isDirty) return;  // fast path: already clean

    tab.isDirty = false;

    // Rebuild so the bullet suffix is removed from the label.
    rebuildTabControl();
}

// Show/hide the tab bar together with the rest of the list UI.
void MultiReplace::setBottomRowVisible(bool visible)
{
    // The tab strip and its associated indicator widgets share visibility:
    // when the list collapses, the dropdown / side dots / "+" must
    // disappear with the strip, otherwise the dropdown ("v"/"...") sits
    // on top of the wrapper toggle in the empty space.
    const int ids[] = {
        IDC_LIST_TABS,
        IDC_NEW_LIST_BUTTON,
        IDC_TAB_LIST_DROPDOWN,
        IDC_TAB_SCROLL_LEFT,
        IDC_TAB_SCROLL_RIGHT,
    };
    const int cmd = visible ? SW_SHOW : SW_HIDE;
    for (int id : ids) {
        if (HWND h = GetDlgItem(_hSelf, id)) {
            ShowWindow(h, cmd);
        }
    }
    // When showing again, the layout / overflow state may have changed
    // while we were hidden (window resized, tabs added, etc.). Re-run
    // the positioning so the indicator widgets only re-appear where
    // they are actually needed.
    if (visible) {
        repositionNewTabButton();
    }
}

void MultiReplace::updateTabTooltip(int /*tabIndex*/)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    // Create the shared tooltip control on first use. This tooltip is
    // informational (shows the tab's file path) rather than explanatory,
    // so it deliberately ignores the tooltipsEnabled option - same
    // rationale as IDC_FILTER_HELP's "(?)" marker which also stays on.
    if (!_hTabTooltip) {
        _hTabTooltip = CreateWindowEx(
            WS_EX_TOPMOST, TOOLTIPS_CLASS, nullptr,
            WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON | TTS_NOPREFIX,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            _hSelf, nullptr, (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE), nullptr);

        if (_hTabTooltip) {
            // Match the visual style used for the regular control
            // tooltips: dark theme when Notepad++ is in dark mode.
            if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
                SetWindowTheme(_hTabTooltip, L"DarkMode_Explorer", nullptr);
            }
            SendMessage(_hTabTooltip, TTM_SETMAXTIPWIDTH, 0, 600);
        }
    }
    if (!_hTabTooltip) return;

    // Rebuild the tool list so tab ids match the current set. Per-tab
    // tools (vs. one whole-control tool) ensure the tooltip re-triggers
    // when the mouse crosses between tabs.
    TOOLINFO tiDel = {};
    tiDel.cbSize = sizeof(TOOLINFO);
    tiDel.hwnd = hTab;
    for (UINT_PTR i = 0; i < 64; ++i) {
        tiDel.uId = i;
        SendMessage(_hTabTooltip, TTM_DELTOOL, 0, reinterpret_cast<LPARAM>(&tiDel));
    }

    // Register one tool per tab, each covering just that tab's rect.
    for (size_t i = 0; i < _tabs.size(); ++i) {
        RECT rc = {};
        if (!TabCtrl_GetItemRect(hTab, static_cast<int>(i), &rc)) continue;

        TOOLINFO ti = {};
        ti.cbSize = sizeof(TOOLINFO);
        ti.uFlags = TTF_SUBCLASS;     // let the tooltip track mouse itself
        ti.hwnd = hTab;
        ti.uId = static_cast<UINT_PTR>(i);
        ti.rect = rc;
        ti.lpszText = LPSTR_TEXTCALLBACKW;
        SendMessage(_hTabTooltip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
    }
}

// Switch active tab: capture outgoing state, restore incoming into live.
void MultiReplace::switchToTab(int newIndex)
{
    if (newIndex < 0 || newIndex >= static_cast<int>(_tabs.size())) return;
    if (newIndex == _activeTabIndex) return;

    // If a list cell edit is in progress, commit it to the outgoing
    // tab's data first. Otherwise the edit control stays open, hovers
    // over the new tab's content, and the user's change silently
    // lands in the wrong list.
    if (hwndEdit) {
        closeEditField(true);
    }

    // Capture live state into the outgoing tab.
    if (_activeTabIndex >= 0 && _activeTabIndex < static_cast<int>(_tabs.size())) {
        captureStateIntoTab(*_tabs[_activeTabIndex]);
    }

    _activeTabIndex = newIndex;
    URM.setActiveTab(_tabs[newIndex]->id);

    // Programmatic switches (e.g. from the overflow dropdown) bypass
    // the tab control's own click handling, so we mirror what a normal
    // click would do: update its selection state, then nudge the
    // strip if the target tab is currently scrolled out of view.
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (hTab) {
        TabCtrl_SetCurSel(hTab, newIndex);
        ensureTabVisible(newIndex);
    }

    // Restore incoming tab into live working state, then rebuild UI.
    // The guard prevents createListViewColumns from reading the old
    // live widths back into the (already correctly restored) globals.
    SuppressWidthSyncGuard widthGuard(_suppressLiveWidthSync);

    restoreStateFromTab(*_tabs[newIndex]);

    // Rebuild ListView columns with the new tab's widths, visibility
    // and order. Honour the lazy-relayout flag: if the panel was
    // resized while this tab was inactive, its stored widths no
    // longer fit the panel and a Redistribute pass is required;
    // otherwise the stored widths apply directly and the manual
    // sizing survives the switch.
    // Suppress ListView redraws to avoid flicker from the
    // column teardown/rebuild; other controls repaint themselves on
    // normal Win32 messages without visible artifacts.
    if (_replaceListView) {
        const bool needsRelayout = _tabs[newIndex]->needsRelayout;
        const WidthMode mode = needsRelayout
            ? WidthMode::Redistribute : WidthMode::UseStored;

        SendMessage(_replaceListView, WM_SETREDRAW, FALSE, 0);
        createListViewColumns(mode);
        autoShowCommentsColumn();
        ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
        SendMessage(_replaceListView, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        _tabs[newIndex]->needsRelayout = false;
    }

    // UI dependent on scope and visibility state.
    setUIElementVisibility();
    updateHeaderSelection();

    updateTabTooltip(_activeTabIndex);
}