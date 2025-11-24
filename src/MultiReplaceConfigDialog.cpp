#include "MultiReplaceConfigDialog.h"
#include "PluginDefinition.h"
#include "Notepad_plus_msgs.h"
#include "StaticDialog/resource.h"
#include "NppStyleKit.h"
#include "MultiReplacePanel.h"
#include <cstddef>
#include <algorithm>
#include <cmath>

extern NppData nppData;

// --- Binding helpers ---
static int clampInt(int v, int lo, int hi) { return (v < lo) ? lo : (v > hi ? hi : v); }

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

    // Appearance
    _bindings.push_back(Binding{ &_hAppearancePanel, IDC_CFG_TOOLTIPS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, tooltipsEnabled), 0, 0 });

    // Search
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, stayAfterReplaceEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, allFromCursorEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, alertNotFoundEnabled), 0, 0 });
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
        if (b.value == ValueType::Bool) {
            bool v = (::IsDlgButtonChecked(panel, b.controlID) == BST_CHECKED);
            *reinterpret_cast<bool*>(base + b.offset) = v;
        }
        else {
            BOOL ok = FALSE;
            UINT v = ::GetDlgItemInt(panel, b.controlID, &ok, FALSE);
            if (ok) {
                int vi = clampInt((int)v, b.minVal, b.maxVal);
                *reinterpret_cast<int*>(base + b.offset) = vi;
            }
        }
    }
}

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

        // Load Scale
        {
            auto paths = MultiReplace::generateConfigFilePaths();
            ConfigManager::instance().load(paths.first);
            std::wstring sScale = ConfigManager::instance().readString(L"Window", L"ScaleFactor", L"1.0");
            try { _userScaleFactor = std::stod(sScale); }
            catch (...) { _userScaleFactor = 1.0; }
            if (_userScaleFactor < 0.5) _userScaleFactor = 0.5;
            if (_userScaleFactor > 2.0) _userScaleFactor = 2.0;
        }

        // Custom Scale Factor -> DPIManager
        dpiMgr->setCustomScaleFactor((float)_userScaleFactor);

        // --- BASE DIMENSIONS ---
        int baseWidth = 810;
        int baseHeight = 380;

        int newW = scaleX(baseWidth);
        int newH = scaleY(baseHeight);

        RECT rc = { 0, 0, newW, newH };
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
        loadSettingsFromConfig();
        resizeUI();

        if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }

        // --- FONT ERSTELLUNG ---
        int fontHeight = dpiMgr->scaleY(13);
        _hCategoryFont = ::CreateFont(
            fontHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
            TEXT("MS Shell Dlg 2")
        );

        applyPanelFonts();

        // Dark Mode
        WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hVariablesAutomationPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hImportScopePanel);
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

        if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }

        int fontHeight = dpiMgr->scaleY(13);
        _hCategoryFont = ::CreateFont(
            fontHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
            TEXT("MS Shell Dlg 2")
        );

        applyPanelFonts();
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
    if (index < 0 || index > 5) return;
    _currentCategory = index;
    ShowWindow(_hSearchReplacePanel, index == 0 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hListViewLayoutPanel, index == 1 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hCsvFlowTabsPanel, index == 2 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hAppearancePanel, index == 3 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hVariablesAutomationPanel, index == 4 ? SW_SHOW : SW_HIDE);
    if (_hCategoryFont) applyPanelFonts();
}

void MultiReplaceConfigDialog::createUI() {
    _hCategoryList = ::CreateWindowEx(0, WC_LISTBOX, TEXT(""), WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, _hSelf, (HMENU)IDC_CONFIG_CATEGORY_LIST, _hInst, nullptr);
    _hCloseButton = ::CreateWindowEx(0, WC_BUTTON, TEXT("Close"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, _hSelf, (HMENU)IDCANCEL, _hInst, nullptr);
    _hResetButton = ::CreateWindowEx(0, WC_BUTTON, TEXT("Reset"), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, _hSelf, (HMENU)IDC_BTN_RESET, _hInst, nullptr);

    _hSearchReplacePanel = createPanel();
    _hListViewLayoutPanel = createPanel();
    _hCsvFlowTabsPanel = createPanel();
    _hAppearancePanel = createPanel();
    _hVariablesAutomationPanel = createPanel();

    createSearchReplacePanelControls();
    createListViewLayoutPanelControls();
    createCsvFlowTabsPanelControls();
    createAppearancePanelControls();
    createVariablesAutomationPanelControls();
}

void MultiReplaceConfigDialog::initUI() {
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Search and Replace"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("List View and Layout"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Flow Tabs"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Appearance"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Variables and Automation"));

    int catToSelect = (_currentCategory >= 0) ? _currentCategory : 0;
    ::SendMessage(_hCategoryList, LB_SETCURSEL, catToSelect, 0);
    _currentCategory = -1; // Force refresh
    showCategory(catToSelect);
}

void MultiReplaceConfigDialog::applyPanelFonts() {
    auto applyFontToChildren = [this](HWND parent) {
        if (!parent) return;
        for (HWND child = ::GetWindow(parent, GW_CHILD); child != nullptr; child = ::GetWindow(child, GW_HWNDNEXT)) {
            ::SendMessage(child, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
        }
        };
    applyFontToChildren(_hSearchReplacePanel);
    applyFontToChildren(_hListViewLayoutPanel);
    applyFontToChildren(_hAppearancePanel);
    applyFontToChildren(_hVariablesAutomationPanel);
    applyFontToChildren(_hImportScopePanel);
    applyFontToChildren(_hCsvFlowTabsPanel);
    if (_hCategoryList) ::SendMessage(_hCategoryList, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
    if (_hCloseButton)  ::SendMessage(_hCloseButton, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
    if (_hResetButton)  ::SendMessage(_hResetButton, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
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
    const int buttonWidth = scaleX(80);

    // Close Button (Right)
    const int closeX = clientW - margin - buttonWidth;
    const int buttonY = clientH - bottomMargin - buttonHeight;
    MoveWindow(_hCloseButton, closeX, buttonY, buttonWidth, buttonHeight, TRUE);

    // Reset Button (Left of Close Button)
    const int resetX = closeX - margin - buttonWidth;
    MoveWindow(_hResetButton, resetX, buttonY, buttonWidth, buttonHeight, TRUE);

    const int contentTop = margin;
    const int contentBottom = buttonY - margin;
    const int contentHeight = contentBottom - contentTop;
    MoveWindow(_hCategoryList, margin, contentTop, catW, contentHeight, TRUE);

    const int panelLeft = margin + catW + margin;
    const int panelWidth = clientW - panelLeft - margin;
    const int panelHeight = contentHeight;

    MoveWindow(_hSearchReplacePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hListViewLayoutPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hCsvFlowTabsPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hAppearancePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hVariablesAutomationPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
}

HWND MultiReplaceConfigDialog::createPanel() {
    return ::CreateWindowEx(0, WC_STATIC, TEXT(""), WS_CHILD, 0, 0, 0, 0, _hSelf, nullptr, _hInst, nullptr);
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
    // TBS_NOTICKS removed to ensure TBM_SETTIC works!
    return ::CreateWindowEx(0, TRACKBAR_CLASS, TEXT(""), WS_CHILD | WS_VISIBLE | TBS_HORZ | TBS_BOTTOM | TBS_TOOLTIPS, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}
HWND MultiReplaceConfigDialog::createSlider(HWND parent, int left, int top, int width, int height, int id, int minValue, int maxValue, int tickMark) {
    HWND hTrack = createTrackbarHorizontal(parent, left, top, width, height, id);
    ::SendMessage(hTrack, TBM_SETRANGE, TRUE, MAKELPARAM(minValue, maxValue));
    if (tickMark >= minValue && tickMark <= maxValue) {
        ::SendMessage(hTrack, TBM_SETTIC, 0, (LPARAM)tickMark);
    }
    return hTrack;
}

// ---------------------------------------------------------------------------
// PANEL LAYOUTS
// ---------------------------------------------------------------------------

void MultiReplaceConfigDialog::createSearchReplacePanelControls() {
    if (!_hSearchReplacePanel) return;
    const int groupW = 420;
    const int left = 70;
    int y = 20;

    createGroupBox(_hSearchReplacePanel, left, y, groupW, 130, IDC_CFG_GRP_SEARCH_BEHAVIOUR, TEXT("Search behaviour"));

    int innerLeft = left + 22;
    int innerTop = y + 35;
    int innerWidth = groupW - 34;

    LayoutBuilder lb(this, _hSearchReplacePanel, innerLeft, innerTop, innerWidth, 26);
    lb.AddCheckbox(IDC_CFG_STAY_AFTER_REPLACE, TEXT("Stay open after replace"));
    lb.AddCheckbox(IDC_CFG_ALL_FROM_CURSOR, TEXT("Search from cursor position"));
    lb.AddCheckbox(IDC_CFG_ALERT_NOT_FOUND, TEXT("Alert when not found"));
}

void MultiReplaceConfigDialog::createListViewLayoutPanelControls() {
    if (!_hListViewLayoutPanel) return;

    const int marginX = 30; const int marginY = 10; const int columnWidth = 260; const int columnSpacing = 20;

    // Adjusted heights for symmetrical padding
    const int leftGroupH = 140;
    const int rightGroupH = 85;
    const int gapAfterTop = 10;
    const int bottomGroupH = 140; // Reduced
    const int topRowY = marginY;

    // Left Group
    createGroupBox(_hListViewLayoutPanel, marginX, topRowY, columnWidth, leftGroupH, IDC_CFG_GRP_LIST_COLUMNS, TEXT("List columns"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 22, topRowY + 25, columnWidth - 24, 24); // +25 innerTop
        lb.AddCheckbox(IDC_CFG_FINDCOUNT_VISIBLE, TEXT("Show Find Count"));
        lb.AddCheckbox(IDC_CFG_REPLACECOUNT_VISIBLE, TEXT("Show Replace Count"));
        lb.AddCheckbox(IDC_CFG_COMMENTS_VISIBLE, TEXT("Show Comments column"));
        lb.AddCheckbox(IDC_CFG_DELETEBUTTON_VISIBLE, TEXT("Show Delete button column"));
    }

    // Right Group
    const int rightColX = marginX + columnWidth + columnSpacing;
    createGroupBox(_hListViewLayoutPanel, rightColX, topRowY, columnWidth, rightGroupH, IDC_CFG_GRP_LIST_STATS, TEXT("List Results"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, rightColX + 22, topRowY + 25, columnWidth - 24, 24);
        lb.AddCheckbox(IDC_CFG_LISTSTATISTICS_ENABLED, TEXT("Show footer statistics"));
        lb.AddCheckbox(IDC_CFG_GROUPRESULTS_ENABLED, TEXT("Group 'Find All' results in Result Dock"));
    }

    // Bottom Group
    const int bottomY = topRowY + (leftGroupH > rightGroupH ? leftGroupH : rightGroupH) + gapAfterTop;

    createGroupBox(_hListViewLayoutPanel, marginX, bottomY, columnWidth + columnSpacing + columnWidth, bottomGroupH, IDC_STATIC, TEXT("List interaction & feedback"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 22, bottomY + 25, (columnWidth + columnSpacing + columnWidth) - 24, 24); // +25 innerTop
        lb.AddCheckbox(IDC_CFG_HIGHLIGHT_MATCH, TEXT("Highlight current match in list"));
        lb.AddCheckbox(IDC_CFG_DOUBLECLICK_EDITS, TEXT("Enable double-click editing"));
        lb.AddCheckbox(IDC_CFG_HOVER_TEXT_ENABLED, TEXT("Show full text on hover"));
        lb.AddSpace(6);
        lb.AddLabel(IDC_CFG_EDITFIELD_LABEL, TEXT("Edit field size (lines)"));
        lb.AddNumberEdit(IDC_CFG_EDITFIELD_SIZE_COMBO, 170, -2, 60, 22);
    }
}

void MultiReplaceConfigDialog::createAppearancePanelControls() {
    if (!_hAppearancePanel) return;

    const int left = 70;
    int top = 20;
    const int groupW = 460;

    LayoutBuilder root(this, _hAppearancePanel, left, top, groupW, 20);

    auto trans = root.BeginGroup(left, top, groupW, 115, 22, 30, IDC_CFG_GRP_TRANSPARENCY, TEXT("Transparency"));
    trans.AddLabeledSlider(IDC_CFG_FOREGROUND_LABEL, TEXT("Foreground (0–255)"), IDC_CFG_FOREGROUND_SLIDER, 190, 160, 0, 255, 40, 170, 18, -4);
    trans.AddLabeledSlider(IDC_CFG_BACKGROUND_LABEL, TEXT("Background (0–255)"), IDC_CFG_BACKGROUND_SLIDER, 190, 160, 0, 255, 40, 170, 18, -4);
    top += 125;

    auto scaleGrp = root.BeginGroup(left, top, groupW, 75, 22, 30, IDC_CFG_GRP_SCALE, TEXT("Scale"));
    scaleGrp.AddLabeledSlider(IDC_CFG_SCALE_LABEL, TEXT("Scale factor"), IDC_CFG_SCALE_SLIDER, 190, 160, 50, 200, 40, 170, 18, -4, 100);
    top += 85;

    auto tips = root.BeginGroup(left, top, groupW, 65, 22, 30, IDC_STATIC, TEXT("Tooltips"));
    tips.AddCheckbox(IDC_CFG_TOOLTIPS_ENABLED, TEXT("Enable tooltips"));
}

void MultiReplaceConfigDialog::createImportScopePanelControls() {
    if (!_hImportScopePanel) return;
    const int left = 40; const int top = 20;
    const int groupWidth = 440;
    const int groupHeight = 140;

    createGroupBox(_hImportScopePanel, left, top, groupWidth, groupHeight, IDC_CFG_GRP_IMPORT_SCOPE, TEXT("Import and scope"));

    const int innerLeft = left + 22;
    const int innerTop = top + 30;
    const int innerWidth = groupWidth - 40;

    LayoutBuilder lb(this, _hImportScopePanel, innerLeft, innerTop, innerWidth, 24);
    lb.AddCheckbox(IDC_CFG_SCOPE_USE_LIST, TEXT("Use list entries for import scope"));
    lb.AddCheckbox(IDC_CFG_IMPORT_ON_STARTUP, TEXT("Import list on startup"));
    lb.AddCheckbox(IDC_CFG_REMEMBER_IMPORT_PATH, TEXT("Remember last import path"));
}

void MultiReplaceConfigDialog::createVariablesAutomationPanelControls() {
    if (!_hVariablesAutomationPanel) return;
    const int margin = 70;
    const int top = 20;
    const int groupW = 460;
    const int groupH = 75;

    LayoutBuilder root(this, _hVariablesAutomationPanel, margin, top, groupW, 24);
    auto inner = root.BeginGroup(margin, top, groupW, groupH, 22, 30, IDC_CFG_GRP_LUA, TEXT("Lua"));
    inner.AddCheckbox(IDC_CFG_LUA_SAFEMODE_ENABLED, TEXT("Enable Lua safe mode"));
}

void MultiReplaceConfigDialog::createCsvFlowTabsPanelControls() {
    if (!_hCsvFlowTabsPanel) return;
    const int left = 70;
    int top = 20;
    const int groupW = 440;
    const int groupH = 120;

    createGroupBox(_hCsvFlowTabsPanel, left, top, groupW, groupH, IDC_CFG_GRP_FLOWTABS, TEXT("Flow Tabs"));
    const int innerLeft = left + 22;
    const int innerTop = top + 35;
    const int innerWidth = groupW - 24;

    LayoutBuilder lb(this, _hCsvFlowTabsPanel, innerLeft, innerTop, innerWidth, 32);
    lb.AddCheckbox(IDC_CFG_FLOWTABS_NUMERIC_ALIGN, TEXT("Align numbers to the right"));
    lb.AddLabel(IDC_STATIC, TEXT("Header lines to skip when sorting CSV"), 240);
    lb.AddNumberEdit(IDC_CFG_HEADERLINES_EDIT, 250, -2, 60, 22);

    if (_hCategoryFont) applyPanelFonts();
}

void MultiReplaceConfigDialog::loadSettingsFromConfig(bool reloadFile)
{
    if (reloadFile)
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
        ConfigManager::instance().load(iniFilePath);
    }
    const auto& CFG = ConfigManager::instance();
    const MultiReplace::Settings s = MultiReplace::getSettings();

    registerBindingsOnce();
    applyBindingsToUI_Generic((void*)&s);

    // Search & Replace
    if (_hSearchReplacePanel) {
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE, s.stayAfterReplaceEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR, s.allFromCursorEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND, s.alertNotFoundEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_HIGHLIGHT_MATCH, s.highlightMatchEnabled ? BST_CHECKED : BST_UNCHECKED);
    }
    // List view
    if (_hListViewLayoutPanel) {
        const BOOL vFind = CFG.readBool(L"ListColumns", L"FindCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vReplace = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vComments = CFG.readBool(L"ListColumns", L"CommentsVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vDelete = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", TRUE) ? BST_CHECKED : BST_UNCHECKED;
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE, vFind);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE, vReplace);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE, vComments);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE, vDelete);
    }
    // Appearance
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
}

void MultiReplaceConfigDialog::applyConfigToSettings()
{
    // 1. Load current state from INI to memory (base)
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
        ConfigManager::instance().load(iniFilePath);
    }

    // 2. Populate Settings struct from UI bindings
    MultiReplace::Settings s = MultiReplace::getSettings();
    readBindingsFromUI_Generic((void*)&s);

    // 3. Handle manual controls not in binding system
    if (_hVariablesAutomationPanel)
        s.luaSafeModeEnabled = (::IsDlgButtonChecked(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED) == BST_CHECKED);

    if (_hAppearancePanel) {
        // Transparency
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"ForegroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"BackgroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));

        // Scale
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER)) {
            int pos = (int)::SendMessage(h, TBM_GETPOS, 0, 0);
            pos = std::clamp(pos, 50, 200);
            double sf = (double)pos / 100.0;
            wchar_t buf[32];
            swprintf(buf, 32, L"%.2f", sf);
            ConfigManager::instance().writeString(L"Window", L"ScaleFactor", buf);

            // Resize config dialog itself if scale changed
            if (std::abs(sf - _userScaleFactor) > 0.001) {
                _userScaleFactor = sf;
                dpiMgr->setCustomScaleFactor((float)_userScaleFactor);

                // --- FIX START: Correct Window Size Calculation (Include Borders/TitleBar) ---
                int baseWidth = 810; int baseHeight = 380;
                int clientW = scaleX(baseWidth);
                int clientH = scaleY(baseHeight);

                // Get current window styles to calculate the full frame size needed
                RECT rc = { 0, 0, clientW, clientH };
                DWORD style = GetWindowLong(_hSelf, GWL_STYLE);
                DWORD exStyle = GetWindowLong(_hSelf, GWL_EXSTYLE);

                // Adjust rectangle to include non-client area (borders, caption)
                AdjustWindowRectEx(&rc, style, FALSE, exStyle);

                int finalW = rc.right - rc.left;
                int finalH = rc.bottom - rc.top;

                // Apply correct dimensions
                SetWindowPos(_hSelf, NULL, 0, 0, finalW, finalH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                // --- FIX END ---

                // Rebuild internal UI (Fonts/Panels)
                auto safeDestroy = [](HWND& h) { if (h && IsWindow(h)) DestroyWindow(h); h = nullptr; };
                safeDestroy(_hCategoryList); safeDestroy(_hCloseButton); safeDestroy(_hResetButton);
                safeDestroy(_hSearchReplacePanel); safeDestroy(_hListViewLayoutPanel);
                safeDestroy(_hAppearancePanel); safeDestroy(_hVariablesAutomationPanel);
                safeDestroy(_hImportScopePanel); safeDestroy(_hCsvFlowTabsPanel);

                createUI(); _currentCategory = -1; initUI();
                loadSettingsFromConfig(false);

                if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }
                int fontHeight = dpiMgr->scaleY(13);
                _hCategoryFont = ::CreateFont(fontHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, TEXT("MS Shell Dlg 2"));

                applyPanelFonts();
                resizeUI();

                // Re-apply Dark Mode themes
                WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hVariablesAutomationPanel);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hImportScopePanel);
                SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);
            }
        }
    }

    // 4. Write struct data to ConfigManager (Memory)
    MultiReplace::writeStructToConfig(s);

    // 5. Commit to Disk
    ConfigManager::instance().save(L"");

    // 6. Notify Main Panel to reload from ConfigManager
    if (MultiReplace::instance) {
        MultiReplace::instance->loadSettingsFromIni();   // Logic settings
        MultiReplace::instance->loadUIConfigFromIni();   // Visuals (Hot-Reload)
    }
}

// --- RESET LOGIC ---
void MultiReplaceConfigDialog::resetToDefaults()
{
    // 1. Create Defaults
    MultiReplace::Settings def{};
    def.tooltipsEnabled = true;
    def.editFieldSize = 5;
    def.listStatisticsEnabled = false;
    def.groupResultsEnabled = false;
    def.stayAfterReplaceEnabled = false;
    def.allFromCursorEnabled = false;
    def.alertNotFoundEnabled = true;
    def.highlightMatchEnabled = true;
    def.doubleClickEditsEnabled = true;
    def.isHoverTextEnabled = true;
    def.csvHeaderLinesCount = 1;
    def.flowTabsNumericAlignEnabled = true;
    def.luaSafeModeEnabled = false;
    def.exportToBashEnabled = false; // Ensure all fields are covered

    // 2. Write Defaults to ConfigManager (instead of applying directly)
    MultiReplace::writeStructToConfig(def);

    // 3. Reset other INI values (manually managed ones)
    auto& cm = ConfigManager::instance();
    cm.writeInt(L"Window", L"ForegroundTransparency", 255);
    cm.writeInt(L"Window", L"BackgroundTransparency", 190);
    cm.writeString(L"Window", L"ScaleFactor", L"1.0");

    cm.writeInt(L"ListColumns", L"FindCountVisible", 0);
    cm.writeInt(L"ListColumns", L"ReplaceCountVisible", 0);
    cm.writeInt(L"ListColumns", L"CommentsVisible", 0);
    cm.writeInt(L"ListColumns", L"DeleteButtonVisible", 1);

    // 4. Save to Disk
    cm.save(L"");

    // 5. Update Config Dialog UI (Sliders, Local State)
    double oldScale = _userScaleFactor;
    _userScaleFactor = 1.0;
    dpiMgr->setCustomScaleFactor(1.0f);

    loadSettingsFromConfig(true); // Reload Dialog UI from the INI we just saved

    // 6. Update Main Panel (The Single Point of Truth Flow)
    if (MultiReplace::instance) {
        MultiReplace::instance->loadSettingsFromIni(); // Logic
        MultiReplace::instance->loadUIConfigFromIni(); // Scaling/Visuals
    }

    // 7. Rebuild Config UI if scale changed (Existing Logic)
    if (std::abs(oldScale - 1.0) > 0.001) {
        int baseWidth = 810; int baseHeight = 380;
        int newW = scaleX(baseWidth); int newH = scaleY(baseHeight);
        SetWindowPos(_hSelf, NULL, 0, 0, newW, newH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);

        auto safeDestroy = [](HWND& h) { if (h && IsWindow(h)) DestroyWindow(h); h = nullptr; };
        safeDestroy(_hCategoryList);
        safeDestroy(_hCloseButton);
        safeDestroy(_hResetButton);
        safeDestroy(_hSearchReplacePanel);
        safeDestroy(_hListViewLayoutPanel);
        safeDestroy(_hAppearancePanel);
        safeDestroy(_hVariablesAutomationPanel);
        safeDestroy(_hImportScopePanel);
        safeDestroy(_hCsvFlowTabsPanel);

        createUI();
        _currentCategory = -1;
        initUI();
        loadSettingsFromConfig(false);

        if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }
        int fontHeight = dpiMgr->scaleY(13);
        _hCategoryFont = ::CreateFont(fontHeight, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, TEXT("MS Shell Dlg 2"));

        applyPanelFonts();
        resizeUI();

        WPARAM mode = (WPARAM)NppDarkMode::dmfInit;
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hVariablesAutomationPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hImportScopePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);
    }
}