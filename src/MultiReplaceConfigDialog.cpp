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

#include "MultiReplaceConfigDialog.h"
#include "PluginDefinition.h"
#include "Notepad_plus_msgs.h"
#include "StaticDialog/resource.h"
#include "LanguageManager.h"
#include "MultiReplacePanel.h"
#include <cstddef>
#include <algorithm>
#include <cmath>
#include <commctrl.h>

extern NppData nppData;

static LanguageManager& LM = LanguageManager::instance();

static int clampInt(int v, int lo, int hi) { return (v < lo) ? lo : (v > hi ? hi : v); }

// ============================================================================
// DATA BINDINGS
// ============================================================================

void MultiReplaceConfigDialog::registerBindingsOnce()
{
    if (_bindingsRegistered) return;
    _bindingsRegistered = true;
    _bindings.clear();

    // List View
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_EDITFIELD_SIZE_COMBO, ControlType::IntEdit, ValueType::Int, offsetof(MultiReplace::Settings, editFieldSize), 2, 20 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, listStatisticsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, groupResultsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_HIGHLIGHT_MATCH, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, highlightMatchEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_DOUBLECLICK_EDITS, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, doubleClickEditsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_HOVER_TEXT_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isHoverTextEnabled), 0, 0 });

    // CSV
    _bindings.push_back(Binding{ &_hCsvFlowTabsPanel, IDC_CFG_HEADERLINES_EDIT, ControlType::IntEdit, ValueType::Int, offsetof(MultiReplace::Settings, csvHeaderLinesCount), 0, 999 });
    _bindings.push_back(Binding{ &_hCsvFlowTabsPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, flowTabsNumericAlignEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hCsvFlowTabsPanel, IDC_CFG_FLOWTABS_INTRO_DONTSHOW, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, flowTabsIntroDontShowEnabled), 0, 0 });  // NEU
    
    // Appearance
    _bindings.push_back(Binding{ &_hAppearancePanel, IDC_CFG_TOOLTIPS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, tooltipsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hAppearancePanel, IDC_CFG_RESULT_DOCK_ENTRY_COLORS, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, resultDockPerEntryColorsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hAppearancePanel, IDC_CFG_USE_LIST_COLORS_MARKING, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, useListColorsForMarking), 0, 0 });

    // Columns
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isFindCountVisible), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isReplaceCountVisible), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isCommentsColumnVisible), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isDeleteButtonVisible), 0, 0 });

    // Search
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, stayAfterReplaceEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, allFromCursorEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_MUTE_SOUNDS, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, muteSounds), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_LIMIT_FILESIZE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, limitFileSizeEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_MAX_FILESIZE_EDIT, ControlType::IntEdit, ValueType::Int, offsetof(MultiReplace::Settings, maxFileSizeMB), 1, 10000 });
}

void MultiReplaceConfigDialog::applyBindingsToUI_Generic(void* settingsPtr)
{
    unsigned char* base = reinterpret_cast<unsigned char*>(settingsPtr);
    for (const auto& b : _bindings) {
        if (!b.panelHandlePtr) continue;
        HWND panel = *b.panelHandlePtr;
        if (!panel) continue;
        if (b.value == ValueType::Bool) {
            bool v = *reinterpret_cast<bool*>(base + b.offset);
            ::CheckDlgButton(panel, b.controlID, v ? BST_CHECKED : BST_UNCHECKED);
        }
        else {
            int v = *reinterpret_cast<int*>(base + b.offset);
            v = clampInt(v, b.minVal, b.maxVal);
            ::SetDlgItemInt(panel, b.controlID, (UINT)v, FALSE);
        }
    }
}

void MultiReplaceConfigDialog::readBindingsFromUI_Generic(void* settingsPtr)
{
    unsigned char* base = reinterpret_cast<unsigned char*>(settingsPtr);
    for (const auto& b : _bindings) {
        if (!b.panelHandlePtr) continue;
        HWND panel = *b.panelHandlePtr;
        if (!panel) continue;

        // Find the control - use FindWindowEx for WC_STATIC panels
        HWND hCtrl = nullptr;
        HWND hChild = ::GetWindow(panel, GW_CHILD);
        while (hChild) {
            if (::GetDlgCtrlID(hChild) == b.controlID) {
                hCtrl = hChild;
                break;
            }
            hChild = ::GetWindow(hChild, GW_HWNDNEXT);
        }
        if (!hCtrl) continue;

        if (b.value == ValueType::Bool) {
            bool v = (::SendMessage(hCtrl, BM_GETCHECK, 0, 0) == BST_CHECKED);
            *reinterpret_cast<bool*>(base + b.offset) = v;
        }
        else {
            wchar_t buf[32] = { 0 };
            ::GetWindowText(hCtrl, buf, 32);
            int vi = _wtoi(buf);
            vi = clampInt(vi, b.minVal, b.maxVal);
            *reinterpret_cast<int*>(base + b.offset) = vi;
        }
    }
}

// ============================================================================
// LIFECYCLE
// ============================================================================

MultiReplaceConfigDialog::~MultiReplaceConfigDialog()
{
    delete dpiMgr;
    dpiMgr = nullptr;
    if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }
}

void MultiReplaceConfigDialog::init(HINSTANCE hInst, HWND hParent)
{
    StaticDialog::init(hInst, hParent);
}

intptr_t CALLBACK MultiReplaceConfigDialog::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);

    switch (message)
    {
    case WM_INITDIALOG:
    {
        dpiMgr = new DPIManager(_hSelf);
        _userScaleFactor = 1.0;

        // Load custom scale
        {
            auto paths = MultiReplace::generateConfigFilePaths();
            ConfigManager::instance().load(paths.first);
            std::wstring sScale = ConfigManager::instance().readString(L"Window", L"ScaleFactor", L"1.0");
            try { _userScaleFactor = std::stod(sScale); }
            catch (...) { _userScaleFactor = 1.0; }
            if (_userScaleFactor < 0.5) _userScaleFactor = 0.5;
            if (_userScaleFactor > 2.0) _userScaleFactor = 2.0;
        }

        dpiMgr->setCustomScaleFactor((float)_userScaleFactor);

        // Calculate dimensions based on scale
        int baseWidth = 810;
        int baseHeight = 380;

        int clientW = scaleX(baseWidth);
        int clientH = scaleY(baseHeight);

        RECT rc = { 0, 0, clientW, clientH };
        DWORD style = GetWindowLong(_hSelf, GWL_STYLE);
        DWORD exStyle = GetWindowLong(_hSelf, GWL_EXSTYLE);
        AdjustWindowRectEx(&rc, style, FALSE, exStyle);
        int finalW = rc.right - rc.left;
        int finalH = rc.bottom - rc.top;

        RECT rcParent;
        if (GetWindowRect(::GetParent(_hSelf), &rcParent)) {
            int x = rcParent.left + (rcParent.right - rcParent.left - finalW) / 2;
            int y = rcParent.top + (rcParent.bottom - rcParent.top - finalH) / 2;
            SetWindowPos(_hSelf, NULL, x, y, finalW, finalH, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        else {
            SetWindowPos(_hSelf, NULL, 0, 0, finalW, finalH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        createUI();
        initUI();
        loadSettingsToConfigUI();

        createFonts();
        applyFonts();
        resizeUI();

        // Initial Dark Mode application
        WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);

        return TRUE;
    }

    case WM_CTLCOLOREDIT:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
        static HBRUSH hEditBrushDark = nullptr;
        static HBRUSH hEditBrushLight = nullptr;

        if (isDark) {
            ::SetTextColor(hdc, RGB(220, 220, 220));
            ::SetBkMode(hdc, OPAQUE);
            if (!hEditBrushDark) hEditBrushDark = ::CreateSolidBrush(RGB(45, 45, 48));
            ::SetBkColor(hdc, RGB(45, 45, 48));
            return reinterpret_cast<INT_PTR>(hEditBrushDark);
        }
        else {
            ::SetTextColor(hdc, RGB(0, 0, 0));
            ::SetBkMode(hdc, OPAQUE);
            if (!hEditBrushLight) hEditBrushLight = ::CreateSolidBrush(::GetSysColor(COLOR_WINDOW));
            ::SetBkColor(hdc, ::GetSysColor(COLOR_WINDOW));
            return reinterpret_cast<INT_PTR>(hEditBrushLight);
        }
    }

    case WM_SIZE:
        resizeUI();
        return TRUE;

    case WM_SHOWWINDOW:
        if (wParam == TRUE) {
            loadSettingsToConfigUI(false);
        }
        return TRUE;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_CONFIG_CATEGORY_LIST:
            if (HIWORD(wParam) == LBN_SELCHANGE) {
                int sel = (int)::SendMessage(_hCategoryList, LB_GETCURSEL, 0, 0);
                showCategory(sel);
            }
            return TRUE;

        case IDC_BTN_RESET:
            resetToDefaults();
            return TRUE;

        case IDCANCEL:
            applyConfigToSettings();
            display(false);
            return TRUE;
        }
        break;

    case WM_DPICHANGED:
    {
        if (dpiMgr) {
            dpiMgr->updateDPI(_hSelf);
            dpiMgr->setCustomScaleFactor((float)_userScaleFactor);
        }
        createFonts();
        applyFonts();
        resizeUI();
        return TRUE;
    }

    case WM_CLOSE:
        applyConfigToSettings();
        display(false);
        return TRUE;
    }
    return FALSE;
}

// ============================================================================
// GRAPHICS & HELPERS
// ============================================================================

int MultiReplaceConfigDialog::scaleX(int value) const {
    int sysScaled = dpiMgr ? dpiMgr->scaleX(value) : value;
    return sysScaled;
}

int MultiReplaceConfigDialog::scaleY(int value) const {
    int sysScaled = dpiMgr ? dpiMgr->scaleY(value) : value;
    return sysScaled;
}

INT_PTR MultiReplaceConfigDialog::handleCtlColorStatic(WPARAM wParam, LPARAM lParam) {
    (void)lParam;
    HDC hdc = reinterpret_cast<HDC>(wParam);
    if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
        ::SetTextColor(hdc, RGB(220, 220, 220));
        ::SetBkMode(hdc, TRANSPARENT);
        return reinterpret_cast<INT_PTR>(::GetStockObject(NULL_BRUSH));
    }
    return 0;
}

void MultiReplaceConfigDialog::showCategory(int index) {
    if (index == _currentCategory) return;
    if (index < 0 || index > 3) return;
    _currentCategory = index;
    ShowWindow(_hSearchReplacePanel, index == 0 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hListViewLayoutPanel, index == 1 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hCsvFlowTabsPanel, index == 2 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hAppearancePanel, index == 3 ? SW_SHOW : SW_HIDE);
    if (_hCategoryFont) applyFonts();
}

void MultiReplaceConfigDialog::createUI() {
    _hCategoryList = ::CreateWindowEx(0, WC_LISTBOX, TEXT(""), WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, _hSelf, (HMENU)IDC_CONFIG_CATEGORY_LIST, _hInst, nullptr);
    _hCloseButton = ::CreateWindowEx(0, WC_BUTTON, LM.getLPCW(L"config_btn_close"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, _hSelf, (HMENU)IDCANCEL, _hInst, nullptr);
    _hResetButton = ::CreateWindowEx(0, WC_BUTTON, LM.getLPCW(L"config_btn_reset"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, _hSelf, (HMENU)IDC_BTN_RESET, _hInst, nullptr);
    _hSearchReplacePanel = createPanel();
    _hListViewLayoutPanel = createPanel();
    _hCsvFlowTabsPanel = createPanel();
    _hAppearancePanel = createPanel();

    createSearchReplacePanelControls();
    createListViewLayoutPanelControls();
    createCsvOptionsPanelControls();
    createAppearancePanelControls();
}

void MultiReplaceConfigDialog::initUI() {
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)LM.getLPCW(L"config_cat_search_replace"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)LM.getLPCW(L"config_cat_list_view"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)LM.getLPCW(L"config_cat_csv"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)LM.getLPCW(L"config_cat_appearance"));

    int catToSelect = (_currentCategory >= 0) ? _currentCategory : 0;
    ::SendMessage(_hCategoryList, LB_SETCURSEL, catToSelect, 0);
    _currentCategory = -1;
    showCategory(catToSelect);
}

void MultiReplaceConfigDialog::createFonts() {
    if (!dpiMgr) return;
    cleanupFonts();

    auto create = [&](int height, int weight, const wchar_t* fontName) -> HFONT {
        HFONT hf = ::CreateFont(
            dpiMgr->scaleY(height), 0, 0, 0, weight, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName
        );
        return hf ? hf : (HFONT)::GetStockObject(DEFAULT_GUI_FONT);
        };

    _hCategoryFont = create(13, FW_NORMAL, L"MS Shell Dlg 2");
}

void MultiReplaceConfigDialog::cleanupFonts() {
    if (_hCategoryFont) {
        ::DeleteObject(_hCategoryFont);
        _hCategoryFont = nullptr;
    }
}

void MultiReplaceConfigDialog::applyFonts() {
    auto applyToHierarchy = [this](HWND parent) {
        if (!parent) return;
        ::SendMessage(parent, WM_SETFONT, (WPARAM)_hCategoryFont, TRUE);
        for (HWND child = ::GetWindow(parent, GW_CHILD); child != nullptr; child = ::GetWindow(child, GW_HWNDNEXT)) {
            ::SendMessage(child, WM_SETFONT, (WPARAM)_hCategoryFont, TRUE);
        }
        };

    applyToHierarchy(_hSearchReplacePanel);
    applyToHierarchy(_hListViewLayoutPanel);
    applyToHierarchy(_hAppearancePanel);
    applyToHierarchy(_hCsvFlowTabsPanel);

    if (_hCategoryList) ::SendMessage(_hCategoryList, WM_SETFONT, (WPARAM)_hCategoryFont, TRUE);
    if (_hCloseButton)  ::SendMessage(_hCloseButton, WM_SETFONT, (WPARAM)_hCategoryFont, TRUE);
    if (_hResetButton)  ::SendMessage(_hResetButton, WM_SETFONT, (WPARAM)_hCategoryFont, TRUE);
}

void MultiReplaceConfigDialog::resizeUI() {
    if (!dpiMgr) return;
    RECT rc{};
    ::GetClientRect(_hSelf, &rc);
    const int clientW = rc.right - rc.left;
    const int clientH = rc.bottom - rc.top;

    const int margin = scaleX(10);
    const int catW = scaleX(140);
    const int buttonHeight = scaleY(24);
    const int bottomMargin = scaleY(10);
    const int closeButtonWidth = scaleX(80);
    const int resetButtonWidth = scaleX(100);
    const int buttonY = clientH - bottomMargin - buttonHeight;

    const int contentTop = margin;
    const int contentBottom = buttonY - margin;
    const int contentHeight = contentBottom - contentTop;
    MoveWindow(_hCategoryList, margin, contentTop, catW, contentHeight, TRUE);

    const int panelLeft = margin + catW + margin;
    const int panelWidth = clientW - panelLeft - margin;
    const int panelHeight = contentHeight;

    const int closeX = panelLeft + (panelWidth - closeButtonWidth) / 2;
    MoveWindow(_hCloseButton, closeX, buttonY, closeButtonWidth, buttonHeight, TRUE);

    const int resetX = margin + (catW - resetButtonWidth) / 2;
    MoveWindow(_hResetButton, resetX, buttonY, resetButtonWidth, buttonHeight, TRUE);

    MoveWindow(_hSearchReplacePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hListViewLayoutPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hCsvFlowTabsPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hAppearancePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
}

LRESULT CALLBACK MultiReplaceConfigDialog::CheckboxSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    (void)uIdSubclass;
    (void)dwRefData;

    // Intercept left button up - this is when checkbox state changes
    if (uMsg == WM_LBUTTONUP) {
        // First let the checkbox handle the click (state changes)
        LRESULT result = ::DefSubclassProc(hWnd, uMsg, wParam, lParam);

        // Now get the new checkbox state and enable/disable the edit field
        HWND hPanel = ::GetParent(hWnd);
        if (hPanel) {
            BOOL checked = (::SendMessage(hWnd, BM_GETCHECK, 0, 0) == BST_CHECKED);
            HWND hEdit = ::GetDlgItem(hPanel, IDC_CFG_MAX_FILESIZE_EDIT);
            if (hEdit) {
                ::EnableWindow(hEdit, checked);
            }
        }
        return result;
    }

    return ::DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK MultiReplaceConfigDialog::PanelSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    (void)uIdSubclass;
    (void)dwRefData;

    if (uMsg == WM_CTLCOLORSTATIC) {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND hCtrl = reinterpret_cast<HWND>(lParam);

        // Check if the control is a Trackbar (Slider)
        TCHAR className[64];
        if (::GetClassName(hCtrl, className, 64)) {
            // FIX: Use lstrcmpi instead of _tcsicmp to avoid missing include error
            if (lstrcmpi(className, TRACKBAR_CLASS) == 0) {

                // TRACKBARS NEED SOLID BACKGROUNDS (Fixes the black box issue)
                bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

                static HBRUSH hBrushDark = ::CreateSolidBrush(RGB(45, 45, 48)); // Dark Mode Background
                static HBRUSH hBrushLight = ::CreateSolidBrush(::GetSysColor(COLOR_WINDOW)); // Light Mode Background (White)

                if (isDark) {
                    ::SetTextColor(hdc, RGB(220, 220, 220));
                    ::SetBkColor(hdc, RGB(45, 45, 48)); // Match the brush
                    return (LRESULT)hBrushDark;
                }
                else {
                    ::SetTextColor(hdc, RGB(0, 0, 0));
                    ::SetBkColor(hdc, ::GetSysColor(COLOR_WINDOW)); // Match the brush
                    return (LRESULT)hBrushLight;
                }
            }
        }

        // Standard handling for Labels (Transparent)
        ::SetBkMode(hdc, TRANSPARENT);
        if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
            ::SetTextColor(hdc, RGB(220, 220, 220));
        }
        else {
            ::SetTextColor(hdc, RGB(0, 0, 0));
        }

        return (LRESULT)::GetStockObject(NULL_BRUSH);
    }

    return ::DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void MultiReplaceConfigDialog::createSearchReplacePanelControls() {
    if (!_hSearchReplacePanel) return;
    const int groupW = 420;
    const int left = 70;
    int y = 20;

    createGroupBox(_hSearchReplacePanel, left, y, groupW, 160, IDC_CFG_GRP_SEARCH_BEHAVIOUR, LM.getLPCW(L"config_grp_search_behaviour"));

    int innerLeft = left + 22;
    int innerTop = y + 35;
    int innerWidth = groupW - 34;

    LayoutBuilder lb(this, _hSearchReplacePanel, innerLeft, innerTop, innerWidth, 26);
    lb.AddCheckbox(IDC_CFG_STAY_AFTER_REPLACE, LM.getLPCW(L"config_chk_stay_after_replace"));
    lb.AddCheckbox(IDC_CFG_ALL_FROM_CURSOR, LM.getLPCW(L"config_chk_all_from_cursor"));
    lb.AddCheckbox(IDC_CFG_MUTE_SOUNDS, LM.getLPCW(L"config_chk_mute_sounds"));

    // File size limit: Checkbox + Edit + "MB" label
    HWND hChkLimit = createCheckBox(_hSearchReplacePanel, innerLeft, lb.y, innerWidth, IDC_CFG_LIMIT_FILESIZE, LM.getLPCW(L"config_chk_limit_filesize"));
    lb.y += lb.stepY;

    // Subclass the checkbox to notify dialog on click
    if (hChkLimit) {
        ::SetWindowSubclass(hChkLimit, CheckboxSubclassProc, 0, reinterpret_cast<DWORD_PTR>(_hSelf));
    }

    lb.AddNumberEdit(IDC_CFG_MAX_FILESIZE_EDIT, 270, -24, 45, 20);
    createStaticText(_hSearchReplacePanel, innerLeft + 325, lb.y - 22, 30, 18, IDC_CFG_FILESIZE_MB_LABEL, L"MB");
}

HWND MultiReplaceConfigDialog::createPanel() {
    HWND hPanel = ::CreateWindowEx(0, WC_STATIC, TEXT(""), WS_CHILD, 0, 0, 0, 0, _hSelf, nullptr, _hInst, nullptr);

    if (hPanel) {
        ::SetWindowSubclass(hPanel, PanelSubclassProc, 0, 0);
    }

    return hPanel;
}
HWND MultiReplaceConfigDialog::createGroupBox(HWND parent, int left, int top, int width, int height, int id, const TCHAR* text) {
    return ::CreateWindowEx(0, WC_BUTTON, text, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createCheckBox(HWND parent, int left, int top, int width, int id, const TCHAR* text) {
    return ::CreateWindowEx(0, WC_BUTTON, text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, scaleX(left), scaleY(top), scaleX(width), scaleY(18), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createStaticText(HWND parent, int left, int top, int width, int height, int id, const TCHAR* text, DWORD extraStyle) {
    return ::CreateWindowEx(0, WC_STATIC, text, WS_CHILD | WS_VISIBLE | extraStyle, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createNumberEdit(HWND parent, int left, int top, int width, int height, int id) {
    return ::CreateWindowEx(WS_EX_CLIENTEDGE, WC_EDIT, TEXT(""), WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL | ES_NUMBER, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createComboDropDownList(HWND parent, int left, int top, int width, int height, int id) {
    return ::CreateWindowEx(0, WC_COMBOBOX, TEXT(""), WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_HASSTRINGS, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createTrackbarHorizontal(HWND parent, int left, int top, int width, int height, int id) {
    return ::CreateWindowEx(0, TRACKBAR_CLASS, TEXT(""), WS_CHILD | WS_VISIBLE | TBS_HORZ | TBS_BOTTOM, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createSlider(HWND parent, int left, int top, int width, int height, int id, int minValue, int maxValue, int tickMark) {
    HWND hTrack = createTrackbarHorizontal(parent, left, top, width, height, id);
    ::SendMessage(hTrack, TBM_SETRANGE, TRUE, MAKELPARAM(minValue, maxValue));
    if (tickMark >= minValue && tickMark <= maxValue) {
        ::SendMessage(hTrack, TBM_SETTIC, 0, (LPARAM)tickMark);
    }
    return hTrack;
}

// ============================================================================
// PANEL LAYOUTS
// ============================================================================

void MultiReplaceConfigDialog::createListViewLayoutPanelControls() {
    if (!_hListViewLayoutPanel) return;

    const int marginX = 30;
    const int marginY = 10;

    // LAYOUT ANPASSUNG:
    // Links schmaler (kurze Texte), Rechts breiter (lange Sätze in DE/RU/FR).
    // Gesamtbreite ca. 540px + Spacing.
    const int leftColWidth = 210;  // Vorher 260
    const int rightColWidth = 330; // Vorher 260 (Verhältnis ~1:1.6 - Goldener Schnitt)
    const int columnSpacing = 20;

    const int leftGroupH = 140;
    const int rightGroupH = 85;
    const int gapAfterTop = 10;
    const int bottomGroupH = 140;
    const int topRowY = marginY;

    // --- Linke Box (Visible Columns) ---
    createGroupBox(_hListViewLayoutPanel, marginX, topRowY, leftColWidth, leftGroupH, IDC_CFG_GRP_LIST_COLUMNS, LM.getLPCW(L"config_grp_list_columns"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 22, topRowY + 25, leftColWidth - 24, 24);
        lb.AddCheckbox(IDC_CFG_FINDCOUNT_VISIBLE, LM.getLPCW(L"config_chk_find_count"));
        lb.AddCheckbox(IDC_CFG_REPLACECOUNT_VISIBLE, LM.getLPCW(L"config_chk_replace_count"));
        lb.AddCheckbox(IDC_CFG_COMMENTS_VISIBLE, LM.getLPCW(L"config_chk_comments"));
        lb.AddCheckbox(IDC_CFG_DELETEBUTTON_VISIBLE, LM.getLPCW(L"config_chk_delete_button"));
    }

    // --- Rechte Box (List Results) ---
    const int rightColX = marginX + leftColWidth + columnSpacing;

    createGroupBox(_hListViewLayoutPanel, rightColX, topRowY, rightColWidth, rightGroupH, IDC_CFG_GRP_LIST_STATS, LM.getLPCW(L"config_grp_list_results"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, rightColX + 22, topRowY + 25, rightColWidth - 24, 24);
        lb.AddCheckbox(IDC_CFG_LISTSTATISTICS_ENABLED, LM.getLPCW(L"config_chk_list_stats"));
        lb.AddCheckbox(IDC_CFG_GROUPRESULTS_ENABLED, LM.getLPCW(L"config_chk_group_results"));
    }

    // --- Untere Box (List Interaction) ---
    const int bottomY = topRowY + (leftGroupH > rightGroupH ? leftGroupH : rightGroupH) + gapAfterTop;

    // Die untere Box spannt sich über die gesamte Breite beider oberer Boxen
    const int totalWidth = leftColWidth + columnSpacing + rightColWidth;

    createGroupBox(_hListViewLayoutPanel, marginX, bottomY, totalWidth, bottomGroupH, IDC_STATIC, LM.getLPCW(L"config_grp_list_interaction"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 22, bottomY + 25, totalWidth - 24, 24);
        lb.AddCheckbox(IDC_CFG_HIGHLIGHT_MATCH, LM.getLPCW(L"config_chk_highlight_match"));
        lb.AddCheckbox(IDC_CFG_DOUBLECLICK_EDITS, LM.getLPCW(L"config_chk_doubleclick"));
        lb.AddCheckbox(IDC_CFG_HOVER_TEXT_ENABLED, LM.getLPCW(L"config_chk_hover_text"));
        lb.AddSpace(6);
        lb.AddLabel(IDC_CFG_EDITFIELD_LABEL, LM.getLPCW(L"config_lbl_edit_height"), 190, 18);
        lb.AddNumberEdit(IDC_CFG_EDITFIELD_SIZE_COMBO, 195, -2, 45, 22);
    }
}

void MultiReplaceConfigDialog::createAppearancePanelControls() {
    if (!_hAppearancePanel) return;

    const int left = 70;
    const int top = 15;
    const int groupW = 460;

    const int groupH_Interface = 135;
    const int groupH_Display = 130;
    const int gap = 25;

    LayoutBuilder root(this, _hAppearancePanel, left, top, groupW, 28);

    auto iface = root.BeginGroup(left, top, groupW, groupH_Interface, 22, 35, IDC_CFG_GRP_INTERFACE, LM.getLPCW(L"config_grp_interface"));

    iface.AddLabeledSlider(IDC_CFG_FOREGROUND_LABEL, LM.getLPCW(L"config_lbl_foreground"), IDC_CFG_FOREGROUND_SLIDER, 190, 160, 0, 255, 34, 170, 18, -4);
    iface.AddLabeledSlider(IDC_CFG_BACKGROUND_LABEL, LM.getLPCW(L"config_lbl_background"), IDC_CFG_BACKGROUND_SLIDER, 190, 160, 0, 255, 34, 170, 18, -4);
    iface.AddLabeledSlider(IDC_CFG_SCALE_LABEL, LM.getLPCW(L"config_lbl_scale_factor"), IDC_CFG_SCALE_SLIDER, 190, 160, 50, 200, 34, 170, 18, -4, 100);

    int nextTop = top + groupH_Interface + gap;

    auto display = root.BeginGroup(left, nextTop, groupW, groupH_Display, 22, 30, IDC_CFG_GRP_DISPLAY_OPTIONS, LM.getLPCW(L"config_grp_display_options"));

    display.AddCheckbox(IDC_CFG_TOOLTIPS_ENABLED, LM.getLPCW(L"config_chk_enable_tooltips"));
    display.AddCheckbox(IDC_CFG_RESULT_DOCK_ENTRY_COLORS, LM.getLPCW(L"config_chk_result_dock_entry_colors"));
    display.AddCheckbox(IDC_CFG_USE_LIST_COLORS_MARKING, LM.getLPCW(L"config_chk_use_list_colors_marking"));
}

void MultiReplaceConfigDialog::createCsvOptionsPanelControls() {
    if (!_hCsvFlowTabsPanel) return;
    const int left = 70;
    int top = 20;
    const int groupW = 440;
    const int groupH = 150;  // Erhöht von 120 auf 150 für die zusätzliche Checkbox

    createGroupBox(_hCsvFlowTabsPanel, left, top, groupW, groupH, IDC_CFG_GRP_CSV_SETTINGS, TEXT("CSV Settings"));
    const int innerLeft = left + 22;
    const int innerTop = top + 35;
    const int innerWidth = groupW - 24;

    LayoutBuilder lb(this, _hCsvFlowTabsPanel, innerLeft, innerTop, innerWidth, 32);
    lb.AddCheckbox(IDC_CFG_FLOWTABS_NUMERIC_ALIGN, LM.getLPCW(L"config_chk_numeric_align"));
    lb.AddCheckbox(IDC_CFG_FLOWTABS_INTRO_DONTSHOW, LM.getLPCW(L"config_chk_flowtabs_intro_dontshow"));  // NEU
    lb.AddLabel(IDC_STATIC, LM.getLPCW(L"config_lbl_csv_sort"), 240);
    lb.AddNumberEdit(IDC_CFG_HEADERLINES_EDIT, 250, -2, 45, 22);

    if (_hCategoryFont) applyFonts();
}

// ============================================================================
// CONFIG LOADING / SAVING
// ============================================================================

void MultiReplaceConfigDialog::loadSettingsToConfigUI(bool reloadFile)
{
    (void)reloadFile;  // Parameter no longer used - cache is always current

    const auto& CFG = ConfigManager::instance();
    MultiReplace::Settings s{};
    s.tooltipsEnabled = CFG.readBool(L"Options", L"Tooltips", true);
    s.exportToBashEnabled = CFG.readBool(L"Options", L"ExportToBash", false);
    s.muteSounds = CFG.readBool(L"Options", L"MuteSounds", false);
    s.doubleClickEditsEnabled = CFG.readBool(L"Options", L"DoubleClickEdits", true);
    s.highlightMatchEnabled = CFG.readBool(L"Options", L"HighlightMatch", true);
    s.flowTabsIntroDontShowEnabled = CFG.readBool(L"Options", L"FlowTabsIntroDontShow", false);
    s.flowTabsNumericAlignEnabled = CFG.readBool(L"Options", L"FlowTabsNumericAlign", true);
    s.isHoverTextEnabled = CFG.readBool(L"Options", L"HoverText", true);
    s.listStatisticsEnabled = CFG.readBool(L"Options", L"ListStatistics", false);
    s.stayAfterReplaceEnabled = CFG.readBool(L"Options", L"StayAfterReplace", false);
    s.groupResultsEnabled = CFG.readBool(L"Options", L"GroupResults", false);
    s.allFromCursorEnabled = CFG.readBool(L"Options", L"AllFromCursor", false);
    s.limitFileSizeEnabled = CFG.readBool(L"ReplaceInFiles", L"LimitFileSize", false);
    s.maxFileSizeMB = CFG.readInt(L"ReplaceInFiles", L"MaxFileSizeMB", 100);
    s.isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    s.isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    s.isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    s.isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);
    s.editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 5);
    s.csvHeaderLinesCount = CFG.readInt(L"Scope", L"HeaderLines", 1);
    s.resultDockPerEntryColorsEnabled = CFG.readBool(L"Options", L"ResultDockPerEntryColors", true);
    s.useListColorsForMarking = CFG.readBool(L"Options", L"UseListColorsForMarking", true);

    registerBindingsOnce();
    applyBindingsToUI_Generic((void*)&s);

    if (_hAppearancePanel) {
        int fg = CFG.readInt(L"Window", L"ForegroundTransparency", 255);
        int bg = CFG.readInt(L"Window", L"BackgroundTransparency", 190);

        if (fg < 0) fg = 0; if (fg > 255) fg = 255;
        if (bg < 0) bg = 0; if (bg > 255) bg = 255;

        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
            ::SendMessage(h, TBM_SETPOS, TRUE, fg);
        }
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
            ::SendMessage(h, TBM_SETPOS, TRUE, bg);
        }

        double sf = 1.0;
        {
            std::wstring sScale = CFG.readString(L"Window", L"ScaleFactor", L"1.0");
            if (!sScale.empty()) {
                try { sf = std::stod(sScale); }
                catch (...) { sf = 1.0; }
            }
        }
        sf = std::clamp(sf, 0.5, 2.0);
        int pos = (int)(sf * 100.0 + 0.5);
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(50, 200));
            ::SendMessage(h, TBM_SETPOS, TRUE, pos);
        }
    }

    // Enable/Disable file size edit based on checkbox state
    if (_hSearchReplacePanel) {
        BOOL checked = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_LIMIT_FILESIZE) == BST_CHECKED);
        ::EnableWindow(::GetDlgItem(_hSearchReplacePanel, IDC_CFG_MAX_FILESIZE_EDIT), checked);
    }
}

void MultiReplaceConfigDialog::applyConfigToSettings()
{
    ::ShowWindow(_hSelf, SW_HIDE);

    auto [iniFilePath, _] = MultiReplace::generateConfigFilePaths();

    // Sync or reload to get latest state
    if (MultiReplace::instance && ::IsWindow(MultiReplace::instance->getDialogHandle())) {
        MultiReplace::instance->syncUIToCache();
    }
    else {
        // MPanel closed - reload INI to get any changes it wrote
        ConfigManager::instance().forceReload(iniFilePath);
    }

    MultiReplace::Settings s = MultiReplace::getSettings();
    readBindingsFromUI_Generic((void*)&s);

    double newScaleFactor = _userScaleFactor;
    if (_hAppearancePanel) {
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"ForegroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"BackgroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));

        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER)) {
            int pos = (int)::SendMessage(h, TBM_GETPOS, 0, 0);
            pos = std::clamp(pos, 50, 200);
            newScaleFactor = (double)pos / 100.0;

            wchar_t buf[32];
            swprintf(buf, 32, L"%.2f", newScaleFactor);
            ConfigManager::instance().writeString(L"Window", L"ScaleFactor", buf);
        }
    }

    MultiReplace::writeStructToConfig(s);
    ConfigManager::instance().save(L"");

    if (MultiReplace::instance) {
        MultiReplace::instance->applyConfigSettingsOnly();
        MultiReplace::instance->loadUIConfigFromIni();
    }

    if (std::abs(newScaleFactor - _userScaleFactor) > 0.001) {
        _userScaleFactor = newScaleFactor;
        dpiMgr->setCustomScaleFactor((float)_userScaleFactor);

        int baseWidth = 810; int baseHeight = 380;
        int clientW = scaleX(baseWidth);
        int clientH = scaleY(baseHeight);

        RECT rc = { 0, 0, clientW, clientH };
        DWORD style = GetWindowLong(_hSelf, GWL_STYLE);
        DWORD exStyle = GetWindowLong(_hSelf, GWL_EXSTYLE);
        AdjustWindowRectEx(&rc, style, FALSE, exStyle);

        int finalW = rc.right - rc.left;
        int finalH = rc.bottom - rc.top;
        SetWindowPos(_hSelf, NULL, 0, 0, finalW, finalH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

        auto safeDestroy = [](HWND& h) { if (h && IsWindow(h)) DestroyWindow(h); h = nullptr; };
        safeDestroy(_hCategoryList); safeDestroy(_hCloseButton); safeDestroy(_hResetButton);
        safeDestroy(_hSearchReplacePanel); safeDestroy(_hListViewLayoutPanel);

        createUI();
        _currentCategory = -1;
        initUI();

        loadSettingsToConfigUI(false);

        createFonts();
        applyFonts();
        resizeUI();

        WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);
    }
}

void MultiReplaceConfigDialog::resetToDefaults()
{
    auto [iniFilePath, _] = MultiReplace::generateConfigFilePaths();

    // Sync or reload to preserve session state
    if (MultiReplace::instance && ::IsWindow(MultiReplace::instance->getDialogHandle())) {
        MultiReplace::instance->syncUIToCache();
    }
    else {
        ConfigManager::instance().forceReload(iniFilePath);
    }

    // Create defaults
    MultiReplace::Settings def{};
    def.tooltipsEnabled = true;
    def.editFieldSize = 5;
    def.listStatisticsEnabled = false;
    def.groupResultsEnabled = false;
    def.stayAfterReplaceEnabled = false;
    def.allFromCursorEnabled = false;
    def.muteSounds = false;
    def.isDeleteButtonVisible = true;
    def.limitFileSizeEnabled = false;
    def.maxFileSizeMB = 100;
    def.highlightMatchEnabled = true;
    def.doubleClickEditsEnabled = true;
    def.isHoverTextEnabled = true;
    def.csvHeaderLinesCount = 1;
    def.flowTabsNumericAlignEnabled = true;
    def.exportToBashEnabled = false;
    def.isFindCountVisible = false;
    def.isReplaceCountVisible = false;
    def.isCommentsColumnVisible = false;
    def.isDeleteButtonVisible = true;
    def.resultDockPerEntryColorsEnabled = true;
    def.useListColorsForMarking = true;

    MultiReplace::writeStructToConfig(def);

    auto& cm = ConfigManager::instance();
    cm.writeInt(L"Window", L"ForegroundTransparency", 255);
    cm.writeInt(L"Window", L"BackgroundTransparency", 190);
    cm.writeString(L"Window", L"ScaleFactor", L"1.0");

    cm.save(L"");

    double oldScale = _userScaleFactor;
    _userScaleFactor = 1.0;
    dpiMgr->setCustomScaleFactor(1.0f);

    if (MultiReplace::instance) {
        MultiReplace::instance->applyConfigSettingsOnly();
        MultiReplace::instance->loadUIConfigFromIni();
    }

    loadSettingsToConfigUI(false);

    if (std::abs(oldScale - 1.0) > 0.001) {
        int baseWidth = 810; int baseHeight = 380;
        int clientW = scaleX(baseWidth);
        int clientH = scaleY(baseHeight);

        RECT rc = { 0, 0, clientW, clientH };
        DWORD style = GetWindowLong(_hSelf, GWL_STYLE);
        DWORD exStyle = GetWindowLong(_hSelf, GWL_EXSTYLE);
        AdjustWindowRectEx(&rc, style, FALSE, exStyle);
        int finalW = rc.right - rc.left;
        int finalH = rc.bottom - rc.top;

        SetWindowPos(_hSelf, NULL, 0, 0, finalW, finalH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

        auto safeDestroy = [](HWND& h) { if (h && IsWindow(h)) DestroyWindow(h); h = nullptr; };
        safeDestroy(_hCategoryList); safeDestroy(_hCloseButton); safeDestroy(_hResetButton);
        safeDestroy(_hSearchReplacePanel); safeDestroy(_hListViewLayoutPanel);

        createUI();
        _currentCategory = -1;
        initUI();
        loadSettingsToConfigUI(false);

        createFonts();
        applyFonts();
        resizeUI();

        WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);
    }
}

void MultiReplaceConfigDialog::applyInternalTheme() {
    WPARAM mode = (WPARAM)NppDarkMode::dmfInit;

    SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);

    HWND panels[] = {
        _hSearchReplacePanel,
        _hListViewLayoutPanel,
        _hCsvFlowTabsPanel,
        _hAppearancePanel
    };

    for (HWND hPanel : panels) {
        if (hPanel) {
            SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)hPanel);
        }
    }

    if (_hCategoryList) SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCategoryList);
    if (_hCloseButton) SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCloseButton);
    if (_hResetButton) SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hResetButton);
}