#include "MultiReplaceConfigDialog.h"
#include "PluginDefinition.h"
#include "Notepad_plus_msgs.h"
#include "StaticDialog/resource.h"
#include "NppStyleKit.h"
#include "MultiReplacePanel.h"


extern NppData nppData;

MultiReplaceConfigDialog::~MultiReplaceConfigDialog()
{
    delete dpiMgr;
    dpiMgr = nullptr;

    if (_hCategoryFont)
    {
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
        dpiMgr = new DPIManager(_hSelf);

        createUI();
        initUI();
        loadSettingsFromConfig();
        resizeUI();

        auto mode = (WPARAM)NppDarkMode::dmfInit;

        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSelf);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hSearchReplacePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hListViewLayoutPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hInterfaceTooltipsPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hAppearancePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hVariablesAutomationPanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hImportScopePanel);
        SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, mode, (LPARAM)_hCsvFlowTabsPanel);


        ::SendMessage(
            nppData._nppHandle,
            NPPM_DARKMODESUBCLASSANDTHEME,
            (WPARAM)NppDarkMode::dmfInit,
            (LPARAM)_hSelf);

        return TRUE;
    }

    case WM_CTLCOLOREDIT:
    {
        HDC hdc = reinterpret_cast<HDC>(wParam);

        const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

        static HBRUSH hEditBrushDark = nullptr;
        static HBRUSH hEditBrushLight = nullptr;

        if (isDark)
        {
            ::SetTextColor(hdc, RGB(220, 220, 220));
            ::SetBkMode(hdc, OPAQUE);

            if (!hEditBrushDark)
                hEditBrushDark = ::CreateSolidBrush(RGB(45, 45, 48));

            ::SetBkColor(hdc, RGB(45, 45, 48));
            return reinterpret_cast<INT_PTR>(hEditBrushDark);
        }
        else
        {
            ::SetTextColor(hdc, RGB(0, 0, 0));
            ::SetBkMode(hdc, OPAQUE);

            if (!hEditBrushLight)
                hEditBrushLight = ::CreateSolidBrush(::GetSysColor(COLOR_WINDOW));

            ::SetBkColor(hdc, ::GetSysColor(COLOR_WINDOW));
            return reinterpret_cast<INT_PTR>(hEditBrushLight);
        }
    }

    case WM_CTLCOLORSTATIC:
        return handleCtlColorStatic(wParam, lParam);

    case WM_SIZE:
        resizeUI();
        return TRUE;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_CONFIG_CATEGORY_LIST:
            if (HIWORD(wParam) == LBN_SELCHANGE)
            {
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

    case WM_CLOSE:
        applyConfigToSettings();
        display(false);
        return TRUE;
    }

    return FALSE;
}

int MultiReplaceConfigDialog::scaleX(int value) const
{
    return dpiMgr ? dpiMgr->scaleX(value) : value;
}

int MultiReplaceConfigDialog::scaleY(int value) const
{
    return dpiMgr ? dpiMgr->scaleY(value) : value;
}

INT_PTR MultiReplaceConfigDialog::handleCtlColorStatic(WPARAM wParam, LPARAM lParam)
{
    (void)lParam;

    HDC hdc = reinterpret_cast<HDC>(wParam);

    const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    if (isDark)
    {
        ::SetTextColor(hdc, RGB(220, 220, 220));
        ::SetBkMode(hdc, TRANSPARENT);
        return reinterpret_cast<INT_PTR>(::GetStockObject(NULL_BRUSH));
    }

    return 0;
}

void MultiReplaceConfigDialog::showCategory(int index)
{
    if (index == _currentCategory)
        return;

    // 0..5: Search/Replace, ListView/Layout, CSV&FlowTabs, Interface/Tooltips, Appearance, Variables
    if (index < 0 || index > 5)
        return;

    _currentCategory = index;

    ShowWindow(_hSearchReplacePanel, index == 0 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hListViewLayoutPanel, index == 1 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hCsvFlowTabsPanel, index == 2 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hAppearancePanel, index == 3 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hVariablesAutomationPanel, index == 4 ? SW_SHOW : SW_HIDE);
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
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("CSV & Flow Tabs"));   // NEW
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Appearance"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Variables and Automation"));

    ::SendMessage(_hCategoryList, LB_SETCURSEL, 0, 0);
    showCategory(0);
}


void MultiReplaceConfigDialog::applyPanelFonts()
{
    if (!_hCategoryFont)
        return;

    auto applyFontToChildren = [this](HWND parent)
        {
            if (!parent)
                return;

            for (HWND child = ::GetWindow(parent, GW_CHILD);
                child != nullptr;
                child = ::GetWindow(child, GW_HWNDNEXT))
            {
                ::SendMessage(child, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
            }
        };

    applyFontToChildren(_hSearchReplacePanel);
    applyFontToChildren(_hListViewLayoutPanel);
    applyFontToChildren(_hInterfaceTooltipsPanel);
    applyFontToChildren(_hAppearancePanel);
    applyFontToChildren(_hVariablesAutomationPanel);
    applyFontToChildren(_hImportScopePanel);
    if (_hCloseButton)
        ::SendMessage(_hCloseButton, WM_SETFONT, (WPARAM)_hCategoryFont, FALSE);
}

void MultiReplaceConfigDialog::resizeUI()
{
    if (!dpiMgr)
        return;

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
    MoveWindow(_hCsvFlowTabsPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE); // NEW
    MoveWindow(_hInterfaceTooltipsPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hAppearancePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hVariablesAutomationPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
}

HWND MultiReplaceConfigDialog::createPanel()
{
    return ::CreateWindowEx(
        0,
        WC_STATIC,
        TEXT(""),
        WS_CHILD,        // first visible via showCategory()
        0, 0, 0, 0,
        _hSelf,
        nullptr,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createGroupBox(HWND parent, int left, int top, int width, int height,
    int id, const TCHAR* text)
{
    return ::CreateWindowEx(
        0,
        WC_BUTTON,
        text,
        WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        scaleY(height),
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createCheckBox(HWND parent, int left, int top,
    int width, int id,
    const TCHAR* text)
{
    const int h = scaleY(18);
    return ::CreateWindowEx(
        0,
        WC_BUTTON,
        text,
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        h,
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createStaticText(HWND parent, int left, int top,
    int width, int height,
    int id, const TCHAR* text,
    DWORD extraStyle)
{
    DWORD style = WS_CHILD | WS_VISIBLE | extraStyle;
    return ::CreateWindowEx(
        0,
        WC_STATIC,
        text,
        style,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        scaleY(height),
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createNumberEdit(HWND parent, int left, int top, int width, int height,
    int id)
{
    return ::CreateWindowEx(
        WS_EX_CLIENTEDGE,
        WC_EDIT,
        TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_LEFT | ES_AUTOHSCROLL | ES_NUMBER,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        scaleY(height),
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createComboDropDownList(HWND parent, int left, int top, int width, int height,
    int id)
{
    return ::CreateWindowEx(
        0,
        WC_COMBOBOX,
        TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | CBS_HASSTRINGS,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        scaleY(height),
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createTrackbarHorizontal(HWND parent, int left, int top, int width, int height,
    int id)
{
    return ::CreateWindowEx(
        0,
        TRACKBAR_CLASS,
        TEXT(""),
        WS_CHILD | WS_VISIBLE | TBS_HORZ,
        scaleX(left),
        scaleY(top),
        scaleX(width),
        scaleY(height),
        parent,
        (HMENU)(INT_PTR)id,
        _hInst,
        nullptr);
}

HWND MultiReplaceConfigDialog::createSlider(HWND parent, int left, int top, int width, int height,
    int id, int minValue, int maxValue)
{
    HWND hTrack = createTrackbarHorizontal(parent, left, top, width, height, id);
    ::SendMessage(hTrack, TBM_SETRANGE, TRUE, MAKELPARAM(minValue, maxValue));
    return hTrack;
}

void MultiReplaceConfigDialog::createSearchReplacePanelControls()
{
    if (!_hSearchReplacePanel)
        return;

    const int margin = 10;
    const int groupW = 360;

    int y = 10;
    int innerLeft = 0;
    int innerWidth = 0;
    int lineY = 0;

    HWND grpSearch = createGroupBox(
        _hSearchReplacePanel, margin, y, groupW, 80,
        IDC_CFG_GRP_SEARCH_BEHAVIOUR,
        TEXT("Search behaviour"));
    (void)grpSearch;

    innerLeft = margin + 12;
    innerWidth = groupW - 24;
    lineY = y + 20;

    createCheckBox(
        _hSearchReplacePanel, innerLeft, lineY, innerWidth,
        IDC_CFG_STAY_AFTER_REPLACE,
        TEXT("Replace: Don't move to the following occurence"));

    lineY += 20;
    createCheckBox(
        _hSearchReplacePanel, innerLeft, lineY, innerWidth,
        IDC_CFG_ALL_FROM_CURSOR,
        TEXT("Search from cursor position"));

    lineY += 20;
    createCheckBox(
        _hSearchReplacePanel, innerLeft, lineY, innerWidth,
        IDC_CFG_ALERT_NOT_FOUND,
        TEXT("Alert when not found"));
}

void MultiReplaceConfigDialog::createListViewLayoutPanelControls()
{
    if (!_hListViewLayoutPanel)
        return;

    const int marginX = 20;
    const int marginY = 10;
    const int columnWidth = 260;
    const int columnSpacing = 20;

    int topRowY = marginY;
    int innerLeft = 0;
    int innerWidth = 0;
    int lineY = 0;

    // Left column: Columns
    HWND grpColumns = createGroupBox(
        _hListViewLayoutPanel, marginX, topRowY, columnWidth, 130,
        IDC_CFG_GRP_LIST_COLUMNS, TEXT("List columns"));
    (void)grpColumns;

    innerLeft = marginX + 12;
    innerWidth = columnWidth - 24;
    lineY = topRowY + 20;

    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_FINDCOUNT_VISIBLE, TEXT("Show Find Count"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_REPLACECOUNT_VISIBLE, TEXT("Show Replace Count"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_COMMENTS_VISIBLE, TEXT("Show Comments column"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_DELETEBUTTON_VISIBLE, TEXT("Show Delete button column"));

    // Right column: List & Results (stats + result dock)
    const int rightColX = marginX + columnWidth + columnSpacing;

    HWND grpStats = createGroupBox(
        _hListViewLayoutPanel, rightColX, topRowY, columnWidth, 100,
        IDC_CFG_GRP_LIST_STATS, TEXT("List & Results"));
    (void)grpStats;

    innerLeft = rightColX + 12;
    innerWidth = columnWidth - 24;
    lineY = topRowY + 20;

    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_LISTSTATISTICS_ENABLED, TEXT("Show footer statistics"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_GROUPRESULTS_ENABLED, TEXT("Group ‘Find All’ results in Result Dock"));

    // Bottom: List interaction & feedback
    int bottomY = topRowY + 140;

    HWND grpInteract = createGroupBox(
        _hListViewLayoutPanel, marginX, bottomY, columnWidth + columnSpacing + columnWidth, 110,
        IDC_STATIC, TEXT("List interaction & feedback"));
    (void)grpInteract;

    innerLeft = marginX + 12;
    innerWidth = (columnWidth + columnSpacing + columnWidth) - 24;
    lineY = bottomY + 20;

    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_HIGHLIGHT_MATCH, TEXT("Highlight current match in list"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_DOUBLECLICK_EDITS, TEXT("Enable double-click editing"));
    lineY += 20;
    createCheckBox(_hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_HOVER_TEXT_ENABLED, TEXT("Show full text on hover"));

    // EditField size (numeric)
    createStaticText(_hListViewLayoutPanel, innerLeft, lineY + 24, 160, 18,
        IDC_CFG_EDITFIELD_LABEL, TEXT("Edit field size (lines)"));
    createNumberEdit(_hListViewLayoutPanel, innerLeft + 170, lineY + 22, 60, 22,
        IDC_CFG_EDITFIELD_SIZE_COMBO);
}

void MultiReplaceConfigDialog::createAppearancePanelControls()
{
    if (!_hAppearancePanel) return;

    const int left = scaleX(20);
    int       top = scaleY(20);
    const int groupW = scaleX(360);

    // --- Transparency ---
    createGroupBox(_hAppearancePanel, left, top, groupW, 120,
        IDC_CFG_GRP_TRANSPARENCY, TEXT("Transparency"));
    createStaticText(_hAppearancePanel, left + 12, top + 24, 150, 18,
        IDC_CFG_FOREGROUND_LABEL, TEXT("Foreground (0-255)"));
    createSlider(_hAppearancePanel, left + 180, top + 20, 160, 26,
        IDC_CFG_FOREGROUND_SLIDER, 0, 255);
    createStaticText(_hAppearancePanel, left + 12, top + 64, 150, 18,
        IDC_CFG_BACKGROUND_LABEL, TEXT("Background (0-255)"));
    createSlider(_hAppearancePanel, left + 180, top + 60, 160, 26,
        IDC_CFG_BACKGROUND_SLIDER, 0, 255);

    HWND hFg = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER);
    HWND hBg = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER);
    if (hFg) ::SendMessage(hFg, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
    if (hBg) ::SendMessage(hBg, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
    top += 130;

    // --- Scale (0.5–2.0) → slider 0..200 (100 == 1.0) ---
    createGroupBox(_hAppearancePanel, left, top, groupW, 80,
        IDC_CFG_GRP_SCALE, TEXT("Scale"));
    createStaticText(_hAppearancePanel, left + 12, top + 24, 150, 18,
        IDC_CFG_SCALE_LABEL, TEXT("Scale factor"));
    HWND hScale = createSlider(_hAppearancePanel, left + 180, top + 20, 160, 26,
        IDC_CFG_SCALE_SLIDER, 0, 200);
    if (hScale) ::SendMessage(hScale, TBM_SETRANGE, TRUE, MAKELPARAM(0, 200));

    // --- Tooltips ---
    top += 90;
    createGroupBox(_hAppearancePanel, left, top, groupW, 60,
        IDC_STATIC, TEXT("Tooltips"));
    createCheckBox(_hAppearancePanel, left + 12, top + 24, groupW - 24,
        IDC_CFG_TOOLTIPS_ENABLED, TEXT("Enable tooltips"));
}

void MultiReplaceConfigDialog::createImportScopePanelControls()
{
    if (!_hImportScopePanel)
        return;

    const int left = scaleX(20);
    int       top = scaleY(20);
    const int groupWidth = scaleX(400);
    const int groupHeight = scaleY(130);

    // GroupBox 
    createGroupBox(
        _hImportScopePanel,
        left,
        top,
        groupWidth,
        groupHeight,
        IDC_CFG_GRP_IMPORT_SCOPE,
        TEXT("Import and scope"));

    const int innerLeft = left + scaleX(12);
    int innerTop = top + scaleY(24);

    createCheckBox(
        _hImportScopePanel, innerLeft, innerTop, groupWidth - scaleX(40),
        IDC_CFG_SCOPE_USE_LIST,
        TEXT("Use list entries for import scope"));

    innerTop += scaleY(20);
    createCheckBox(
        _hImportScopePanel, innerLeft, innerTop, groupWidth - scaleX(40),
        IDC_CFG_IMPORT_ON_STARTUP,
        TEXT("Import list on startup"));

    innerTop += scaleY(20);
    createCheckBox(
        _hImportScopePanel, innerLeft, innerTop, groupWidth - scaleX(40),
        IDC_CFG_REMEMBER_IMPORT_PATH,
        TEXT("Remember last import path"));
}

void MultiReplaceConfigDialog::createVariablesAutomationPanelControls()
{
    if (!_hVariablesAutomationPanel) return;

    const int margin = 10;
    const int groupW = 360;
    const int y = 10;

    createGroupBox(_hVariablesAutomationPanel, margin, y, groupW, 60,
        IDC_CFG_GRP_LUA, TEXT("Lua"));
    createCheckBox(_hVariablesAutomationPanel, margin + 12, y + 24, groupW - 24,
        IDC_CFG_LUA_SAFEMODE_ENABLED, TEXT("Enable Lua safe mode"));
}

void MultiReplaceConfigDialog::createCsvFlowTabsPanelControls()
{
    if (!_hCsvFlowTabsPanel) return;

    const int margin = 20;
    const int groupW = 360;

    HWND grp = createGroupBox(
        _hCsvFlowTabsPanel, margin, 10, groupW, 120,
        IDC_STATIC, TEXT("CSV & Flow Tabs"));
    (void)grp;

    const int innerLeft = margin + 12;
    const int innerWidth = groupW - 24;
    int lineY = 10 + 20;

    createCheckBox(
        _hCsvFlowTabsPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_FLOWTABS_NUMERIC_ALIGN,
        TEXT("Align numeric values in Flow Tabs"));

    lineY += 30;

    createStaticText(
        _hCsvFlowTabsPanel, innerLeft, lineY, 260, 18,
        IDC_CFG_HEADERLINES_LABEL,
        TEXT("Header lines to skip when sorting CSV:"));

    createNumberEdit(
        _hCsvFlowTabsPanel, innerLeft, lineY + 18, 50, 18,
        IDC_CFG_HEADERLINES_EDIT);
}


// --- Bind Settings -> Controls ---
void MultiReplaceConfigDialog::loadSettingsFromConfig()
{
    // Load INI
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
        ConfigManager::instance().load(iniFilePath);
    }

    const auto& CFG = ConfigManager::instance();
    const MultiReplace::Settings s = MultiReplace::getSettings();

    // ---- Search & Replace ----
    if (_hSearchReplacePanel)
    {
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE, s.stayAfterReplaceEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR, s.allFromCursorEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND, s.alertNotFoundEnabled ? BST_CHECKED : BST_UNCHECKED);
    }

    // ---- List view / layout ----
    if (_hListViewLayoutPanel)
    {
        // Bind list/summary toggles from settings
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED, s.listStatisticsEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED, s.groupResultsEnabled ? BST_CHECKED : BST_UNCHECKED);

        // Bind column visibility from INI keys (kept as before)
        const BOOL vFind = CFG.readBool(L"ListColumns", L"FindCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vReplace = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vComments = CFG.readBool(L"ListColumns", L"CommentsVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vDelete = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", TRUE) ? BST_CHECKED : BST_UNCHECKED;

        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE, vFind);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE, vReplace);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE, vComments);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE, vDelete);

        // Edit field size (lines) now lives on the ListView panel
        ::SetDlgItemInt(_hListViewLayoutPanel, IDC_CFG_EDITFIELD_SIZE_COMBO, (UINT)s.editFieldSize, FALSE);

        // List interaction flags from settings
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_HIGHLIGHT_MATCH, s.highlightMatchEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_DOUBLECLICK_EDITS, s.doubleClickEditsEnabled ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_HOVER_TEXT_ENABLED, s.isHoverTextEnabled ? BST_CHECKED : BST_UNCHECKED);
    }

    // ---- CSV & Flow Tabs ----
    if (_hCsvFlowTabsPanel)
    {
        ::CheckDlgButton(_hCsvFlowTabsPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN,
            s.flowTabsNumericAlignEnabled ? BST_CHECKED : BST_UNCHECKED);

        // Header lines field from settings (already loaded from INI)
        int hl = std::clamp(s.csvHeaderLinesCount, 0, 999);
        ::SetDlgItemInt(_hCsvFlowTabsPanel, IDC_CFG_HEADERLINES_EDIT, (UINT)hl, FALSE);
    }

    // ---- Appearance ----
    if (_hAppearancePanel)
    {
        // Transparency sliders from INI
        int fg = std::clamp(CFG.readInt(L"Window", L"ForegroundTransparency", 255), 0, 255);
        int bg = std::clamp(CFG.readInt(L"Window", L"BackgroundTransparency", 190), 0, 255);

        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
            ::SendMessage(h, TBM_SETPOS, TRUE, fg);
        }
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255));
            ::SendMessage(h, TBM_SETPOS, TRUE, bg);
        }

        // Scale slider (0..200, 100 == 1.0) from INI
        double sf = 1.0;
        {
            std::wstring v = CFG.readString(L"Window", L"ScaleFactor", L"1.0");
            try { sf = std::stod(v); }
            catch (...) { sf = 1.0; }
        }
        sf = std::clamp(sf, 0.5, 2.0);
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER)) {
            ::SendMessage(h, TBM_SETRANGE, TRUE, MAKELPARAM(0, 200));
            ::SendMessage(h, TBM_SETPOS, TRUE, (int)std::lround(sf * 100.0));
        }

        // Tooltip master switch from settings
        ::CheckDlgButton(_hAppearancePanel, IDC_CFG_TOOLTIPS_ENABLED, s.tooltipsEnabled ? BST_CHECKED : BST_UNCHECKED);
    }

    // ---- Variables / automation ----
    if (_hVariablesAutomationPanel)
    {
        ::CheckDlgButton(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED,
            s.luaSafeModeEnabled ? BST_CHECKED : BST_UNCHECKED);
    }
}

void MultiReplaceConfigDialog::applyConfigToSettings()
{
    // Ensure INI is loaded for writeback
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
        ConfigManager::instance().load(iniFilePath);
    }

    MultiReplace::Settings s = MultiReplace::getSettings();

    // ---- Search & Replace ----
    if (_hSearchReplacePanel)
    {
        s.stayAfterReplaceEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE) == BST_CHECKED);
        s.allFromCursorEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR) == BST_CHECKED);
        s.alertNotFoundEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND) == BST_CHECKED);
    }

    // ---- List view / layout ----
    if (_hListViewLayoutPanel)
    {
        s.listStatisticsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED) == BST_CHECKED);
        s.groupResultsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED) == BST_CHECKED);

        // Edit field size (lines) from ListView panel
        BOOL okEdit = FALSE;
        UINT ev = ::GetDlgItemInt(_hListViewLayoutPanel, IDC_CFG_EDITFIELD_SIZE_COMBO, &okEdit, FALSE);
        if (okEdit)
        {
            int e = (int)ev;
            s.editFieldSize = (int)std::clamp(e, 2, 20);
        }

        // Interaction flags
        s.highlightMatchEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_HIGHLIGHT_MATCH) == BST_CHECKED);
        s.doubleClickEditsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_DOUBLECLICK_EDITS) == BST_CHECKED);
        s.isHoverTextEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_HOVER_TEXT_ENABLED) == BST_CHECKED);

        // Column visibility -> INI (kept behavior)
        auto& cm = ConfigManager::instance();
        cm.writeInt(L"ListColumns", L"FindCountVisible",
            (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"ReplaceCountVisible",
            (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"CommentsVisible",
            (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"DeleteButtonVisible",
            (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE) == BST_CHECKED) ? 1 : 0);
    }

    // ---- CSV & Flow Tabs ----
    if (_hCsvFlowTabsPanel)
    {
        s.flowTabsNumericAlignEnabled =
            (::IsDlgButtonChecked(_hCsvFlowTabsPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN) == BST_CHECKED);

        BOOL okHL = FALSE;
        UINT hv = ::GetDlgItemInt(_hCsvFlowTabsPanel, IDC_CFG_HEADERLINES_EDIT, &okHL, FALSE);
        if (okHL)
        {
            int h = (int)hv;
            s.csvHeaderLinesCount = (int)std::clamp(h, 0, 999);
        }
    }

    // ---- Appearance ----
    if (_hAppearancePanel)
    {
        s.tooltipsEnabled = (::IsDlgButtonChecked(_hAppearancePanel, IDC_CFG_TOOLTIPS_ENABLED) == BST_CHECKED);

        // Transparency writeback to INI
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"ForegroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"BackgroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));

        // ScaleFactor writeback to INI
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_SCALE_SLIDER))
        {
            int pos = (int)::SendMessage(h, TBM_GETPOS, 0, 0);
            pos = std::clamp(pos, 0, 200);
            double sf = (double)pos / 100.0;
            sf = std::clamp(sf, 0.5, 2.0);

            wchar_t buf[32];
            swprintf(buf, 32, L"%.2f", sf);
            ConfigManager::instance().writeString(L"Window", L"ScaleFactor", buf);
        }
    }

    // ---- Variables / automation ----
    if (_hVariablesAutomationPanel)
    {
        s.luaSafeModeEnabled =
            (::IsDlgButtonChecked(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED) == BST_CHECKED);
    }

    // Apply + persist
    MultiReplace::applySettings(s);
    MultiReplace::persistSettings(s);
    ConfigManager::instance().save(L"");

    if (MultiReplace::instance)
        MultiReplace::instance->loadUIConfigFromIni();
}
