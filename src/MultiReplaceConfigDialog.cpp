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

// Helper to create a DPI-scaled font, considering the custom user scale factor
static HFONT CreateScaledDialogFontFor(HWND hwnd, double scaleFactor, int basePt = 9)
{
    UINT dpiY = 96;
    HDC hdc = ::GetDC(hwnd);
    if (hdc) {
        dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
        ::ReleaseDC(hwnd, hdc);
    }

    // Scale base point size by user factor (System DPI handled by MulDiv/LOGPIXELSY)
    int scaledPt = (int)(basePt * scaleFactor);

    const int height = -MulDiv(scaledPt, dpiY, 72);
    return ::CreateFont(
        height, 0, 0, 0,
        FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE,
        TEXT("Segoe UI"));
}

// --- Binding helpers ---
static int clampInt(int v, int lo, int hi) { return (v < lo) ? lo : (v > hi ? hi : v); }

void MultiReplaceConfigDialog::registerBindingsOnce()
{
    if (_bindingsRegistered) return;
    _bindingsRegistered = true;
    _bindings.clear();

    // List View / Layout
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_EDITFIELD_SIZE_COMBO, ControlType::IntEdit, ValueType::Int, offsetof(MultiReplace::Settings, editFieldSize), 2, 20 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, listStatisticsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, groupResultsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_HIGHLIGHT_MATCH, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, highlightMatchEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_DOUBLECLICK_EDITS, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, doubleClickEditsEnabled), 0, 0 });
    _bindings.push_back(Binding{ &_hListViewLayoutPanel, IDC_CFG_HOVER_TEXT_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, isHoverTextEnabled), 0, 0 });

    // CSV & Flow Tabs
    _bindings.push_back(Binding{ &_hCsvFlowTabsPanel, IDC_CFG_HEADERLINES_EDIT, ControlType::IntEdit, ValueType::Int, offsetof(MultiReplace::Settings, csvHeaderLinesCount), 0, 999 });
    _bindings.push_back(Binding{ &_hCsvFlowTabsPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, flowTabsNumericAlignEnabled), 0, 0 });

    // Appearance
    _bindings.push_back(Binding{ &_hAppearancePanel, IDC_CFG_TOOLTIPS_ENABLED, ControlType::Checkbox, ValueType::Bool, offsetof(MultiReplace::Settings, tooltipsEnabled), 0, 0 });

    // Search & Replace
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

    if (_hCategoryFont) {
        ::DeleteObject(_hCategoryFont);
        _hCategoryFont = nullptr;
    }
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
        // 1. Create DPI manager
        dpiMgr = new DPIManager(_hSelf);

        // 2. Load config to read ScaleFactor BEFORE building UI
        _userScaleFactor = 1.0;
        {
            auto paths = MultiReplace::generateConfigFilePaths();
            const auto& iniFilePath = paths.first;
            ConfigManager::instance().load(iniFilePath);

            std::wstring sScale = ConfigManager::instance().readString(L"Window", L"ScaleFactor", L"1.0");
            try { _userScaleFactor = std::stod(sScale); }
            catch (...) { _userScaleFactor = 1.0; }

            // Clamp
            if (_userScaleFactor < 0.5) _userScaleFactor = 0.5;
            if (_userScaleFactor > 2.0) _userScaleFactor = 2.0;

            // Note: DPIManager is untouched, so we handle user scale locally in scaleX/Y
        }

        // 3. Resize Main Window frame (uses scaleX/Y which now includes _userScaleFactor)
        int baseWidth = 620;
        int baseHeight = 480;
        int newW = scaleX(baseWidth);
        int newH = scaleY(baseHeight);

        // Adjust for borders
        RECT rc = { 0, 0, newW, newH };
        DWORD style = GetWindowLong(_hSelf, GWL_STYLE);
        DWORD exStyle = GetWindowLong(_hSelf, GWL_EXSTYLE);
        AdjustWindowRectEx(&rc, style, FALSE, exStyle);
        int finalW = rc.right - rc.left;
        int finalH = rc.bottom - rc.top;

        // Center on parent
        RECT rcParent;
        if (GetWindowRect(::GetParent(_hSelf), &rcParent)) {
            int x = rcParent.left + (rcParent.right - rcParent.left - finalW) / 2;
            int y = rcParent.top + (rcParent.bottom - rcParent.top - finalH) / 2;
            SetWindowPos(_hSelf, NULL, x, y, finalW, finalH, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        else {
            SetWindowPos(_hSelf, NULL, 0, 0, finalW, finalH, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        // 4. Create UI (uses scaleX/Y)
        createUI();
        initUI();
        loadSettingsFromConfig();
        resizeUI();

        // 5. Create Scaled Font
        if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }
        _hCategoryFont = CreateScaledDialogFontFor(_hSelf, _userScaleFactor, 9);
        applyPanelFonts();

        // Dark Mode Init
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
        const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
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

        case IDCANCEL:
            applyConfigToSettings();
            display(false);
            return TRUE;
        }
        break;

    case WM_DPICHANGED:
    {
        // Update system DPI
        if (dpiMgr) dpiMgr->updateDPI(_hSelf);

        // Rebuild font using local user scale
        if (_hCategoryFont) { ::DeleteObject(_hCategoryFont); _hCategoryFont = nullptr; }
        _hCategoryFont = CreateScaledDialogFontFor(_hSelf, _userScaleFactor, 9);

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

// Combined scaling: System DPI (via DPIManager) * User Scale (local)
int MultiReplaceConfigDialog::scaleX(int value) const
{
    int sysScaled = dpiMgr ? dpiMgr->scaleX(value) : value;
    return (int)(sysScaled * _userScaleFactor);
}

int MultiReplaceConfigDialog::scaleY(int value) const
{
    int sysScaled = dpiMgr ? dpiMgr->scaleY(value) : value;
    return (int)(sysScaled * _userScaleFactor);
}

INT_PTR MultiReplaceConfigDialog::handleCtlColorStatic(WPARAM wParam, LPARAM lParam)
{
    (void)lParam;
    HDC hdc = reinterpret_cast<HDC>(wParam);
    const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    if (isDark) {
        ::SetTextColor(hdc, RGB(220, 220, 220));
        ::SetBkMode(hdc, TRANSPARENT);
        return reinterpret_cast<INT_PTR>(::GetStockObject(NULL_BRUSH));
    }
    return 0;
}

void MultiReplaceConfigDialog::showCategory(int index)
{
    if (index == _currentCategory) return;
    if (index < 0 || index > 5) return;

    _currentCategory = index;

    ShowWindow(_hSearchReplacePanel, index == 0 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hListViewLayoutPanel, index == 1 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hCsvFlowTabsPanel, index == 2 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hAppearancePanel, index == 3 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hVariablesAutomationPanel, index == 4 ? SW_SHOW : SW_HIDE);

    // ensure fonts applied to newly visible panel
    if (_hCategoryFont) applyPanelFonts();
}

void MultiReplaceConfigDialog::createUI()
{
    _hCategoryList = ::CreateWindowEx(
        0, WC_LISTBOX, TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
        0, 0, 0, 0, _hSelf, (HMENU)IDC_CONFIG_CATEGORY_LIST, _hInst, nullptr);

    _hSearchReplacePanel = createPanel();
    _hListViewLayoutPanel = createPanel();
    _hCsvFlowTabsPanel = createPanel();
    _hAppearancePanel = createPanel();
    _hVariablesAutomationPanel = createPanel();

    _hCloseButton = ::CreateWindowEx(
        0, WC_BUTTON, TEXT("Close"),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        0, 0, 0, 0, _hSelf, (HMENU)IDCANCEL, _hInst, nullptr);

    createSearchReplacePanelControls();
    createListViewLayoutPanelControls();
    createCsvFlowTabsPanelControls();
    createAppearancePanelControls();
    createVariablesAutomationPanelControls();
}

void MultiReplaceConfigDialog::initUI()
{
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Search and Replace"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("List View and Layout"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Flow Tabs"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Appearance"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Variables and Automation"));

    ::SendMessage(_hCategoryList, LB_SETCURSEL, 0, 0);
    showCategory(0);
}

void MultiReplaceConfigDialog::applyPanelFonts()
{
    auto applyFontToChildren = [this](HWND parent)
        {
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
}

void MultiReplaceConfigDialog::resizeUI()
{
    if (!dpiMgr) return;

    RECT rc{};
    ::GetClientRect(_hSelf, &rc);

    const int clientW = rc.right - rc.left;
    const int clientH = rc.bottom - rc.top;

    const int margin = scaleX(10);
    const int catW = scaleX(180);
    const int buttonHeight = scaleY(24);
    const int bottomMargin = scaleY(10);

    const int buttonWidth = scaleX(80);
    const int buttonX = clientW - margin - buttonWidth;
    const int buttonY = clientH - bottomMargin - buttonHeight;
    MoveWindow(_hCloseButton, buttonX, buttonY, buttonWidth, buttonHeight, TRUE);

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

HWND MultiReplaceConfigDialog::createPanel()
{
    return ::CreateWindowEx(0, WC_STATIC, TEXT(""), WS_CHILD, 0, 0, 0, 0, _hSelf, nullptr, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createGroupBox(HWND parent, int left, int top, int width, int height, int id, const TCHAR* text)
{
    return ::CreateWindowEx(0, WC_BUTTON, text, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createCheckBox(HWND parent, int left, int top, int width, int id, const TCHAR* text)
{
    return ::CreateWindowEx(0, WC_BUTTON, text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, scaleX(left), scaleY(top), scaleX(width), scaleY(18), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createStaticText(HWND parent, int left, int top, int width, int height, int id, const TCHAR* text, DWORD extraStyle)
{
    return ::CreateWindowEx(0, WC_STATIC, text, WS_CHILD | WS_VISIBLE | extraStyle, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createNumberEdit(HWND parent, int left, int top, int width, int height, int id)
{
    return ::CreateWindowEx(WS_EX_CLIENTEDGE, WC_EDIT, TEXT(""), WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL | ES_NUMBER, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createComboDropDownList(HWND parent, int left, int top, int width, int height, int id)
{
    return ::CreateWindowEx(0, WC_COMBOBOX, TEXT(""), WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_HASSTRINGS, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createTrackbarHorizontal(HWND parent, int left, int top, int width, int height, int id)
{
    return ::CreateWindowEx(0, TRACKBAR_CLASS, TEXT(""), WS_CHILD | WS_VISIBLE | TBS_HORZ, scaleX(left), scaleY(top), scaleX(width), scaleY(height), parent, (HMENU)(INT_PTR)id, _hInst, nullptr);
}

HWND MultiReplaceConfigDialog::createSlider(HWND parent, int left, int top, int width, int height, int id, int minValue, int maxValue)
{
    HWND hTrack = createTrackbarHorizontal(parent, left, top, width, height, id);
    ::SendMessage(hTrack, TBM_SETRANGE, TRUE, MAKELPARAM(minValue, maxValue));
    return hTrack;
}

void MultiReplaceConfigDialog::createSearchReplacePanelControls()
{
    if (!_hSearchReplacePanel) return;
    const int margin = 10;
    const int groupW = 360;
    int y = 10;

    createGroupBox(_hSearchReplacePanel, margin, y, groupW, 80, IDC_CFG_GRP_SEARCH_BEHAVIOUR, TEXT("Search behaviour"));

    int lbX = margin + 12;
    int lbY = y + 20;
    int innerWidth = groupW - 24;

    LayoutBuilder lb(this, _hSearchReplacePanel, lbX, lbY, innerWidth, 20);
    lb.AddCheckbox(IDC_CFG_STAY_AFTER_REPLACE, TEXT("Stay open after replace"));
    lb.AddCheckbox(IDC_CFG_ALL_FROM_CURSOR, TEXT("Search from cursor position"));
    lb.AddCheckbox(IDC_CFG_ALERT_NOT_FOUND, TEXT("Alert when not found"));
}

void MultiReplaceConfigDialog::createListViewLayoutPanelControls()
{
    if (!_hListViewLayoutPanel) return;

    const int marginX = 20;
    const int marginY = 10;
    const int columnWidth = 260;
    const int columnSpacing = 20;
    const int leftGroupH = 130;
    const int rightGroupH = 100;
    const int gapAfterTop = 10;
    const int bottomGroupH = 140;
    const int topRowY = marginY;

    // Left
    createGroupBox(_hListViewLayoutPanel, marginX, topRowY, columnWidth, leftGroupH, IDC_CFG_GRP_LIST_COLUMNS, TEXT("List columns"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 12, topRowY + 20, columnWidth - 24, 20);
        lb.AddCheckbox(IDC_CFG_FINDCOUNT_VISIBLE, TEXT("Show Find Count"));
        lb.AddCheckbox(IDC_CFG_REPLACECOUNT_VISIBLE, TEXT("Show Replace Count"));
        lb.AddCheckbox(IDC_CFG_COMMENTS_VISIBLE, TEXT("Show Comments column"));
        lb.AddCheckbox(IDC_CFG_DELETEBUTTON_VISIBLE, TEXT("Show Delete button column"));
    }

    // Right
    const int rightColX = marginX + columnWidth + columnSpacing;
    createGroupBox(_hListViewLayoutPanel, rightColX, topRowY, columnWidth, rightGroupH, IDC_CFG_GRP_LIST_STATS, TEXT("List Results"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, rightColX + 12, topRowY + 20, columnWidth - 24, 20);
        lb.AddCheckbox(IDC_CFG_LISTSTATISTICS_ENABLED, TEXT("Show footer statistics"));
        lb.AddCheckbox(IDC_CFG_GROUPRESULTS_ENABLED, TEXT("Group 'Find All' results in Result Dock"));
    }

    // Bottom
    const int bottomY = topRowY + (leftGroupH > rightGroupH ? leftGroupH : rightGroupH) + gapAfterTop;
    createGroupBox(_hListViewLayoutPanel, marginX, bottomY, columnWidth + columnSpacing + columnWidth, bottomGroupH, IDC_STATIC, TEXT("List interaction & feedback"));
    {
        LayoutBuilder lb(this, _hListViewLayoutPanel, marginX + 12, bottomY + 20, (columnWidth + columnSpacing + columnWidth) - 24, 20);
        lb.AddCheckbox(IDC_CFG_HIGHLIGHT_MATCH, TEXT("Highlight current match in list"));
        lb.AddCheckbox(IDC_CFG_DOUBLECLICK_EDITS, TEXT("Enable double-click editing"));
        lb.AddCheckbox(IDC_CFG_HOVER_TEXT_ENABLED, TEXT("Show full text on hover"));
        lb.AddSpace(16);
        lb.AddLabel(IDC_CFG_EDITFIELD_LABEL, TEXT("Edit field size (lines)"));
        lb.AddNumberEdit(IDC_CFG_EDITFIELD_SIZE_COMBO, 170, -2, 60, 22);
    }
}

void MultiReplaceConfigDialog::createAppearancePanelControls()
{
    if (!_hAppearancePanel) return;

    const int left = scaleX(20);
    int top = scaleY(20);
    const int groupW = scaleX(360);

    LayoutBuilder root(this, _hAppearancePanel, left, top, groupW, 20);
    auto trans = root.BeginGroup(left, top, groupW, 120, scaleX(12), scaleY(24), IDC_CFG_GRP_TRANSPARENCY, TEXT("Transparency"));
    trans.AddLabeledSlider(IDC_CFG_FOREGROUND_LABEL, TEXT("Foreground (0–255)"), IDC_CFG_FOREGROUND_SLIDER, 180, 160, 0, 255);
    trans.AddLabeledSlider(IDC_CFG_BACKGROUND_LABEL, TEXT("Background (0–255)"), IDC_CFG_BACKGROUND_SLIDER, 180, 160, 0, 255);
    top += 130;

    auto scaleGrp = root.BeginGroup(left, top, groupW, 80, scaleX(12), scaleY(24), IDC_CFG_GRP_SCALE, TEXT("Scale"));
    scaleGrp.AddLabeledSlider(IDC_CFG_SCALE_LABEL, TEXT("Scale factor"), IDC_CFG_SCALE_SLIDER, 180, 160, 0, 200, 40);
    top += 90;

    auto tips = root.BeginGroup(left, top, groupW, 60, scaleX(12), scaleY(24), IDC_STATIC, TEXT("Tooltips"));
    tips.AddCheckbox(IDC_CFG_TOOLTIPS_ENABLED, TEXT("Enable tooltips"));
}

void MultiReplaceConfigDialog::createImportScopePanelControls()
{
    if (!_hImportScopePanel) return;

    const int left = scaleX(20);
    const int top = scaleY(20);
    const int groupWidth = scaleX(400);
    const int groupHeight = scaleY(130);

    createGroupBox(_hImportScopePanel, left, top, groupWidth, groupHeight, IDC_CFG_GRP_IMPORT_SCOPE, TEXT("Import and scope"));

    const int innerLeft = left + scaleX(12);
    const int innerTop = top + scaleY(24);
    const int innerWidth = groupWidth - scaleX(40);

    {
        LayoutBuilder lb(this, _hImportScopePanel, innerLeft, innerTop, innerWidth, scaleY(20));
        lb.AddCheckbox(IDC_CFG_SCOPE_USE_LIST, TEXT("Use list entries for import scope"));
        lb.AddCheckbox(IDC_CFG_IMPORT_ON_STARTUP, TEXT("Import list on startup"));
        lb.AddCheckbox(IDC_CFG_REMEMBER_IMPORT_PATH, TEXT("Remember last import path"));
    }
}

void MultiReplaceConfigDialog::createVariablesAutomationPanelControls()
{
    if (!_hVariablesAutomationPanel) return;

    const int margin = scaleX(10);
    const int top = scaleY(10);
    const int groupW = scaleX(360);
    const int groupH = scaleY(60);

    LayoutBuilder root(this, _hVariablesAutomationPanel, margin, top, groupW, scaleY(20));
    auto inner = root.BeginGroup(margin, top, groupW, groupH, scaleX(12), scaleY(24), IDC_CFG_GRP_LUA, TEXT("Lua"));
    inner.AddCheckbox(IDC_CFG_LUA_SAFEMODE_ENABLED, TEXT("Enable Lua safe mode"));
}

void MultiReplaceConfigDialog::createCsvFlowTabsPanelControls()
{
    if (!_hCsvFlowTabsPanel) return;

    const int left = scaleX(20);
    int top = scaleY(20);
    const int groupW = scaleX(360);
    const int groupH = scaleY(100);

    createGroupBox(_hCsvFlowTabsPanel, left, top, groupW, groupH, IDC_CFG_GRP_FLOWTABS, TEXT("Flow Tabs"));

    const int innerLeft = left + scaleX(12);
    const int innerTop = top + scaleY(24);
    const int innerWidth = groupW - scaleX(24);

    LayoutBuilder lb(this, _hCsvFlowTabsPanel, innerLeft, innerTop, innerWidth, scaleY(20));
    lb.AddCheckbox(IDC_CFG_FLOWTABS_NUMERIC_ALIGN, TEXT("Align numbers to the right"));
    lb.AddLabel(IDC_STATIC, TEXT("Header lines to skip when sorting CSV"));
    lb.AddNumberEdit(IDC_CFG_HEADERLINES_EDIT, scaleX(200), -2, scaleX(60), scaleY(22));

    if (_hCategoryFont) applyPanelFonts();
}

void MultiReplaceConfigDialog::loadSettingsFromConfig()
{
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
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 200));
            ::SendMessage(h, TBM_SETPOS, TRUE, pos);
        }
    }
}

void MultiReplaceConfigDialog::applyConfigToSettings()
{
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
        ConfigManager::instance().load(iniFilePath);
    }

    MultiReplace::Settings s = MultiReplace::getSettings();
    readBindingsFromUI_Generic((void*)&s);

    if (_hSearchReplacePanel) {
        s.stayAfterReplaceEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE) == BST_CHECKED);
        s.allFromCursorEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR) == BST_CHECKED);
        s.alertNotFoundEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND) == BST_CHECKED);
        s.highlightMatchEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_HIGHLIGHT_MATCH) == BST_CHECKED);
    }

    if (_hListViewLayoutPanel) {
        s.listStatisticsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED) == BST_CHECKED);
        s.groupResultsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED) == BST_CHECKED);

        auto& cm = ConfigManager::instance();
        cm.writeInt(L"ListColumns", L"FindCountVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"ReplaceCountVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"CommentsVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"DeleteButtonVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE) == BST_CHECKED) ? 1 : 0);
    }

    if (_hAppearancePanel) {
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"ForegroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"BackgroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));

        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER)) {
            int pos = (int)::SendMessage(h, TBM_GETPOS, 0, 0);
            if (pos < 0)   pos = 0;
            if (pos > 200) pos = 200;
            double sf = (double)pos / 100.0;
            sf = std::clamp(sf, 0.5, 2.0);
            wchar_t buf[32];
            swprintf(buf, 32, L"%.2f", sf);
            ConfigManager::instance().writeString(L"Window", L"ScaleFactor", buf);
        }
    }

    if (_hVariablesAutomationPanel)
        s.luaSafeModeEnabled = (::IsDlgButtonChecked(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED) == BST_CHECKED);

    MultiReplace::applySettings(s);
    MultiReplace::persistSettings(s);
    ConfigManager::instance().save(L"");

    if (MultiReplace::instance)
        MultiReplace::instance->loadUIConfigFromIni();
}