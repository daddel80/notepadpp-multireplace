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
#include "TabStripLayout.h"
#include "UndoRedoManager.h"
#include <windows.h>
#include <uxtheme.h>
#include <string>
#include <vector>
#include <algorithm>

// Local singleton alias, matching MultiReplacePanel.cpp, so the
// moved function bodies stay byte-for-byte identical.
static UndoRedoManager& URM = UndoRedoManager::instance();

namespace {

    // Follows the label of a tab with unsaved changes, in the strip and in the dropdown
    constexpr wchar_t kDirtyMarker[] = L" \u25CF";

    // Tab labels and menu entries read "&" as a mnemonic prefix
    std::wstring escapeAmpersands(const std::wstring& text)
    {
        std::wstring escaped;
        escaped.reserve(text.size());
        for (wchar_t c : text) {
            if (c == L'&') escaped += L'&';
            escaped += c;
        }
        return escaped;
    }

    // Tab edges at the natural layout. While narrower than its tabs (inserts, the generic
    // resize pass) the control may have scrolled itself; that is undone before measuring.
    std::vector<TabStripLayout::Edge> tabEdges(HWND hTab)
    {
        const int count = TabCtrl_GetItemCount(hTab);
        RECT rc{};
        bool scrolledBack = false;
        for (int guard = count; guard > 0 && TabCtrl_GetItemRect(hTab, 0, &rc) && rc.left < 0; --guard) {
            SendMessage(hTab, WM_HSCROLL, MAKEWPARAM(SB_THUMBPOSITION, 0), 0);
            scrolledBack = true;
        }
        // The control scrolls without invalidating
        if (scrolledBack) InvalidateRect(hTab, nullptr, TRUE);

        std::vector<TabStripLayout::Edge> edges;
        edges.reserve(count);
        for (int i = 0; i < count && TabCtrl_GetItemRect(hTab, i, &rc); ++i) {
            edges.push_back({ rc.left, rc.right });
        }
        return edges;
    }

    // Show `width` pixels of the control from client x `left` on; a width of 0 shows all of it
    void clipTabControl(HWND hTab, int left, int width)
    {
        RECT box{};
        const int current = GetWindowRgnBox(hTab, &box);
        if (width <= 0) {
            if (current != ERROR) SetWindowRgn(hTab, nullptr, TRUE);
            return;
        }

        // Window regions count from the window corner, item rects from the client area
        RECT window{};
        GetWindowRect(hTab, &window);
        POINT client{ 0, 0 };
        ClientToScreen(hTab, &client);
        const int x = client.x - window.left + left;
        const RECT wanted = { x, 0, x + width, window.bottom - window.top };
        if (current == SIMPLEREGION && EqualRect(&box, &wanted)) return;
        SetWindowRgn(hTab, CreateRectRgnIndirect(&wanted), TRUE);
    }
}

std::wstring MultiReplace::truncateTabName(const std::wstring& name, size_t maxChars)
{
    if (name.length() <= maxChars) return name;
    if (maxChars <= 1) return std::wstring(L"\u2026");
    return name.substr(0, maxChars - 1) + L"\u2026";
}

// Truncated name plus dirty marker, "&" doubled so the tab shows it
std::wstring MultiReplace::buildTabLabel(int tabIndex) const
{
    if (tabIndex < 0 || tabIndex >= static_cast<int>(_tabs.size())) return {};
    std::wstring label = escapeAmpersands(truncateTabName(_tabs[tabIndex]->name,
        static_cast<size_t>(tabMaxLength)));
    if (_tabs[tabIndex]->isDirty) {
        label += kDirtyMarker;
    }
    return label;
}

void MultiReplace::rebuildTabControl()
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    // Clicks, double-click rename and drag reordering run through the subclass.
    // Idempotent: SetWindowSubclass with the same id+proc just updates the ref data.
    SetWindowSubclass(hTab, &MultiReplace::tabControlSubclassProc, 1,
        reinterpret_cast<DWORD_PTR>(this));

    TabCtrl_DeleteAllItems(hTab);

    // Compact tabs, scaled like Notepad++ scales its tab padding
    TabCtrl_SetPadding(hTab, sx(6), sy(3));
    TabCtrl_SetMinTabWidth(hTab, sx(40));

    for (size_t i = 0; i < _tabs.size(); ++i) {
        std::wstring label = buildTabLabel(static_cast<int>(i));
        TCITEM item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<LPWSTR>(label.c_str());
        TabCtrl_InsertItem(hTab, static_cast<int>(i), &item);
    }

    if (!_tabs.empty()) {
        TabCtrl_SetCurSel(hTab, std::clamp(_activeTabIndex, 0, static_cast<int>(_tabs.size()) - 1));
    }

    // Honor collapsed-list state; otherwise SW_SHOW leaves a tab-strip
    // remnant on first open.
    const bool listShown = useListEnabled || keepListVisible;
    ShowWindow(hTab, listShown ? SW_SHOW : SW_HIDE);

    // Lays the row out with the active tab in view
    ensureTabVisible(_activeTabIndex);

    if (listShown) {
        InvalidateRect(hTab, nullptr, TRUE);
        UpdateWindow(hTab);
    }
    updateTabTooltip();
}

// Tab row geometry, anchored right to left from the Use List toggle: dropdown, "+", "..." slot, strip
TabStripLayout::Metrics MultiReplace::tabStripMetrics()
{
    RECT client{};
    GetClientRect(_hSelf, &client);
    const int gap = sx(2);
    const int button = sx(22);
    const int indicator = sx(18) + gap;
    const int left = sx(14) - sx(2);
    const int plusMax = client.right - sx(40) - 2 * (gap + button);
    const int maxWidth = std::max(plusMax - indicator - gap - left, sx(60));
    return { left, maxWidth, indicator, gap, gap + indicator, plusMax };
}

// Place the tab control and clip it to the visible strip, then "..." where tabs are
// hidden, "+" and the dropdown. The control holds every tab; scrolling moves it.
void MultiReplace::layoutTabStrip()
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    HWND hPlus = GetDlgItem(_hSelf, IDC_NEW_LIST_BUTTON);
    HWND hDrop = GetDlgItem(_hSelf, IDC_TAB_LIST_DROPDOWN);
    HWND hLeft = GetDlgItem(_hSelf, IDC_TAB_SCROLL_LEFT);
    HWND hRight = GetDlgItem(_hSelf, IDC_TAB_SCROLL_RIGHT);
    if (!hPlus || !hTab || !hDrop || !hLeft || !hRight) return;

    if (!useListEnabled && !keepListVisible) return;

    const TabStripLayout::Metrics m = tabStripMetrics();
    const TabStripLayout::Layout lay = TabStripLayout::arrange(tabEdges(hTab), m, _tabFirstVisible);
    _tabFirstVisible = lay.first;

    // Row and height come from the generic layout pass, x and width from the strip
    RECT cur;
    GetWindowRect(hTab, &cur);
    POINT curTL = { cur.left, cur.top };
    ScreenToClient(_hSelf, &curTL);
    const int tabHeight = cur.bottom - cur.top;
    if (curTL.x != lay.controlLeft || cur.right - cur.left != lay.controlWidth) {
        SetWindowPos(hTab, nullptr, lay.controlLeft, curTL.y, lay.controlWidth, tabHeight,
            SWP_NOZORDER | SWP_NOACTIVATE);
    }

    // Holding every tab, the control needs no arrow buttons; "..." and the dropdown replace them
    if (HWND hUpDown = FindWindowExW(hTab, nullptr, UPDOWN_CLASS, nullptr)) {
        ShowWindow(hUpDown, SW_HIDE);
    }
    clipTabControl(hTab, lay.clipLeft, lay.clipped ? lay.viewWidth : 0);

    const int gap = sx(2);
    const int dotW = sx(18);
    const int buttonW = sx(22);
    const int buttonH = sy(22);
    const int yMid = curTL.y + (tabHeight - buttonH) / 2;

    if (lay.hiddenLeft) {
        SetWindowPos(hLeft, HWND_TOP, m.left, yMid, dotW, buttonH, SWP_NOACTIVATE);
        ShowWindow(hLeft, SW_SHOWNA);
    }
    else {
        ShowWindow(hLeft, SW_HIDE);
    }

    if (lay.hiddenRight) {
        SetWindowPos(hRight, HWND_TOP, lay.plusLeft - m.indicator, yMid, dotW, buttonH, SWP_NOACTIVATE);
        ShowWindow(hRight, SW_SHOWNA);
    }
    else {
        ShowWindow(hRight, SW_HIDE);
    }

    SetWindowPos(hPlus, HWND_TOP, lay.plusLeft, yMid, buttonW, buttonH, SWP_NOACTIVATE);

    // The dropdown lists every tab whenever not all of them fit
    if (lay.overflow) {
        SetWindowPos(hDrop, HWND_TOP, m.plusMax + buttonW + gap, yMid, buttonW, buttonH, SWP_NOACTIVATE);
        ShowWindow(hDrop, SW_SHOWNA);
    }
    else {
        ShowWindow(hDrop, SW_HIDE);
    }
}

// New labels for tabs from..to after a rename, move or dirty change; their widths shift the row
void MultiReplace::refreshTabLabels(int from, int to)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    const int last = std::min(to, TabCtrl_GetItemCount(hTab) - 1);
    for (int i = std::max(from, 0); i <= last; ++i) {
        std::wstring label = buildTabLabel(i);
        TCITEM item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<LPWSTR>(label.c_str());
        TabCtrl_SetItem(hTab, i, &item);
    }
    // Lays the row out again; a grown label must not push the active tab out of view
    ensureTabVisible(_activeTabIndex);
    updateTabTooltip();
}

// Scroll the strip as little as possible so the tab shows completely
void MultiReplace::ensureTabVisible(int tabIndex)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    _tabFirstVisible = TabStripLayout::revealing(tabEdges(hTab), tabStripMetrics(), _tabFirstVisible, tabIndex);
    layoutTabStrip();
}

// Scroll the strip by whole tabs (negative left, positive right)
void MultiReplace::scrollTabStrip(int direction)
{
    HWND hTab = GetDlgItem(_hSelf, IDC_LIST_TABS);
    if (!hTab) return;

    _tabFirstVisible = TabStripLayout::scrolled(tabEdges(hTab), tabStripMetrics(), _tabFirstVisible, direction);
    layoutTabStrip();
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
        std::wstring label = escapeAmpersands(_tabs[i]->name);
        if (_tabs[i]->isDirty) label += kDirtyMarker;

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
// clearTabDirty force-clears after a save. Both relabel only on change.
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

    refreshTabLabels(_activeTabIndex, _activeTabIndex);
}

void MultiReplace::clearTabDirty(int tabIndex)
{
    if (tabIndex < 0 ||
        tabIndex >= static_cast<int>(_tabs.size())) return;

    TabState& tab = *_tabs[tabIndex];
    if (!tab.isDirty) return;  // fast path: already clean

    tab.isDirty = false;

    refreshTabLabels(tabIndex, tabIndex);
}

// Show/hide the tab row together with the rest of the list UI. Hidden, the dropdown
// and "..." go with the strip; shown, the layout decides which of them it needs.
void MultiReplace::setBottomRowVisible(bool visible)
{
    if (!visible) {
        for (int id : { IDC_LIST_TABS, IDC_NEW_LIST_BUTTON, IDC_TAB_LIST_DROPDOWN, IDC_TAB_SCROLL_LEFT, IDC_TAB_SCROLL_RIGHT }) {
            if (HWND h = GetDlgItem(_hSelf, id)) ShowWindow(h, SW_HIDE);
        }
        return;
    }
    for (int id : { IDC_LIST_TABS, IDC_NEW_LIST_BUTTON }) {
        if (HWND h = GetDlgItem(_hSelf, id)) ShowWindow(h, SW_SHOW);
    }
    layoutTabStrip();
}

// One tooltip tool per tab, so the tip re-triggers between tabs; the rects move only with the labels
void MultiReplace::updateTabTooltip()
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

    // Drop every tool of the previous pass, however many tabs it had
    for (LRESULT left = SendMessage(_hTabTooltip, TTM_GETTOOLCOUNT, 0, 0); left > 0; --left) {
        TOOLINFO old = {};
        old.cbSize = sizeof(TOOLINFO);
        if (!SendMessage(_hTabTooltip, TTM_ENUMTOOLS, 0, reinterpret_cast<LPARAM>(&old))) break;
        SendMessage(_hTabTooltip, TTM_DELTOOL, 0, reinterpret_cast<LPARAM>(&old));
    }

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
}