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

    // 0..4: Search/Replace, ListView/Layout, Interface/Tooltips, Appearance, Variables
    if (index < 0 || index > 4)
        return;

    _currentCategory = index;

    ShowWindow(_hSearchReplacePanel, index == 0 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hListViewLayoutPanel, index == 1 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hInterfaceTooltipsPanel, index == 2 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hAppearancePanel, index == 3 ? SW_SHOW : SW_HIDE);
    ShowWindow(_hVariablesAutomationPanel, index == 4 ? SW_SHOW : SW_HIDE);
    // _hImportScopePanel wird nicht mehr verwendet
}

void MultiReplaceConfigDialog::createUI()
{
    // Left categories list
    _hCategoryList = ::CreateWindowEx(
        0, WC_LISTBOX, TEXT(""),
        WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
        0, 0, 0, 0, _hSelf, (HMENU)IDC_CONFIG_CATEGORY_LIST, _hInst, nullptr);

    // Containerpanels
    _hSearchReplacePanel = createPanel();
    _hListViewLayoutPanel = createPanel();
    _hInterfaceTooltipsPanel = createPanel();
    _hAppearancePanel = createPanel();
    _hVariablesAutomationPanel = createPanel();
    // _hImportScopePanel nicht mehr erzeugen

    // Close button
    _hCloseButton = ::CreateWindowEx(
        0, WC_BUTTON, TEXT("Close"),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        0, 0, 0, 0, _hSelf, (HMENU)IDCANCEL, _hInst, nullptr);

    // Controls
    createSearchReplacePanelControls();
    createListViewLayoutPanelControls();
    createInterfaceTooltipsPanelControls();
    createAppearancePanelControls();
    createVariablesAutomationPanelControls();
    // createImportScopePanelControls() entfällt
}

void MultiReplaceConfigDialog::initUI()
{
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Search and Replace"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("List View and Layout"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Interface and Tooltips"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Appearance"));
    SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Variables and Automation"));
    // "Import and Scope" entfällt

    // Startansicht
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

    // Close button at bottom right
    const int buttonWidth = scaleX(80);
    const int buttonX = clientW - margin - buttonWidth;
    const int buttonY = clientH - bottomMargin - buttonHeight;

    MoveWindow(_hCloseButton, buttonX, buttonY, buttonWidth, buttonHeight, TRUE);

    // Main content area (above Close button)
    const int contentTop = margin;
    const int contentBottom = buttonY - margin;
    const int contentHeight = contentBottom - contentTop;

    // Category list
    MoveWindow(_hCategoryList, margin, contentTop, catW, contentHeight, TRUE);

    // Panel area
    const int panelLeft = margin + catW + margin;
    const int panelWidth = clientW - panelLeft - margin;
    const int panelHeight = contentHeight;

    MoveWindow(_hSearchReplacePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hListViewLayoutPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hInterfaceTooltipsPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hAppearancePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hVariablesAutomationPanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
    MoveWindow(_hImportScopePanel, panelLeft, contentTop, panelWidth, panelHeight, TRUE);
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

    // Group: Search behaviour
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
        TEXT("Stay open after replace"));

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

    // Group: Highlighting
    y += groupW / 4 + 12;

    HWND grpHighlight = createGroupBox(
        _hSearchReplacePanel, margin, y, groupW, 60,
        IDC_CFG_GRP_HIGHLIGHT,
        TEXT("Highlighting"));

    (void)grpHighlight;

    innerLeft = margin + 12;
    innerWidth = groupW - 24;
    lineY = y + 20;

    createCheckBox(
        _hSearchReplacePanel, innerLeft, lineY, innerWidth,
        IDC_CFG_HIGHLIGHT_MATCH,
        TEXT("Highlight current match"));
}

void MultiReplaceConfigDialog::createListViewLayoutPanelControls()
{
    if (!_hListViewLayoutPanel)
        return;

    // Two columns on the panel:
    // left:   List columns, Flow tabs
    // right:  Statistics and layout, Sorting

    const int marginX = 20;
    const int marginY = 10;
    const int columnWidth = 260;
    const int columnSpacing = 20;

    int topRowY = marginY;
    int innerLeft = 0;
    int innerWidth = 0;
    int lineY = 0;

    // Left column, top: List columns
    HWND grpColumns = createGroupBox(
        _hListViewLayoutPanel,
        marginX,
        topRowY,
        columnWidth,
        100,
        IDC_CFG_GRP_LIST_COLUMNS,
        TEXT("List columns"));

    (void)grpColumns;

    innerLeft = marginX + 12;
    innerWidth = columnWidth - 24;
    lineY = topRowY + 20;

    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_FINDCOUNT_VISIBLE,
        TEXT("Show Find Count"));

    lineY += 20;
    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_REPLACECOUNT_VISIBLE,
        TEXT("Show Replace Count"));

    lineY += 20;
    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_COMMENTS_VISIBLE,
        TEXT("Show Comments column"));

    lineY += 20;
    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_DELETEBUTTON_VISIBLE,
        TEXT("Show Delete button column"));

    // Right column, top: Statistics and layout
    const int rightColX = marginX + columnWidth + columnSpacing;

    HWND grpStats = createGroupBox(
        _hListViewLayoutPanel,
        rightColX,
        topRowY,
        columnWidth,
        100,
        IDC_CFG_GRP_LIST_STATS,
        TEXT("Statistics and layout"));

    (void)grpStats;

    innerLeft = rightColX + 12;
    innerWidth = columnWidth - 24;
    lineY = topRowY + 20;

    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_LISTSTATISTICS_ENABLED,
        TEXT("Show footer statistics"));

    lineY += 20;
    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_GROUPRESULTS_ENABLED,
        TEXT("Group results in Result Dock"));

    // Left column, bottom: Flow tabs
    int flowTabsY = topRowY + 110;

    HWND grpFlowTabs = createGroupBox(
        _hListViewLayoutPanel,
        marginX,
        flowTabsY,
        columnWidth,
        80,
        IDC_CFG_GRP_FLOWTABS,
        TEXT("Flow tabs"));

    (void)grpFlowTabs;

    innerLeft = marginX + 12;
    innerWidth = columnWidth - 24;
    lineY = flowTabsY + 20;

    createCheckBox(
        _hListViewLayoutPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_FLOWTABS_NUMERIC_ALIGN,
        TEXT("Align numeric values"));

    // Right column, bottom: Sorting
    int sortY = topRowY + 110;

    HWND grpSorting = createGroupBox(
        _hListViewLayoutPanel,
        rightColX,
        sortY,
        columnWidth,
        80,
        IDC_CFG_GRP_SORTING,
        TEXT("Sorting"));

    (void)grpSorting;

    innerLeft = rightColX + 12;
    innerWidth = columnWidth - 24;
    lineY = sortY + 20;

    createStaticText(
        _hListViewLayoutPanel,
        innerLeft,
        lineY,
        200,
        18,
        IDC_CFG_HEADERLINES_LABEL,
        TEXT("Header lines to skip (not used yet):"));

    createNumberEdit(
        _hListViewLayoutPanel,
        innerLeft,
        lineY + 18,
        40,
        18,
        IDC_CFG_HEADERLINES_EDIT);
}

void MultiReplaceConfigDialog::createInterfaceTooltipsPanelControls()
{
    if (!_hInterfaceTooltipsPanel) return;

    const int margin = 10;
    const int groupW = 360;

    int y = 10;
    int innerLeft = 0;
    int innerWidth = 0;
    int lineY = 0;

    // Interaction
    HWND grpInteraction = createGroupBox(_hInterfaceTooltipsPanel, margin, y, groupW, 80,
        IDC_CFG_GRP_INTERACTION, TEXT("Interaction"));
    (void)grpInteraction;

    innerLeft = margin + 12;
    innerWidth = groupW - 24;
    lineY = y + 20;

    createCheckBox(_hInterfaceTooltipsPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_DOUBLECLICK_EDITS, TEXT("Enable double-click editing"));
    lineY += 20;

    createCheckBox(_hInterfaceTooltipsPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_HOVER_TEXT_ENABLED, TEXT("Enable per-entry hover text"));

    // Tooltips
    y += 90;
    HWND grpTooltips = createGroupBox(_hInterfaceTooltipsPanel, margin, y, groupW, 60,
        IDC_CFG_GRP_TOOLTIPS, TEXT("Tooltips"));
    (void)grpTooltips;

    innerLeft = margin + 12;
    lineY = y + 20;

    createCheckBox(_hInterfaceTooltipsPanel, innerLeft, lineY, innerWidth,
        IDC_CFG_TOOLTIPS_ENABLED, TEXT("Enable tooltips"));
}

void MultiReplaceConfigDialog::createAppearancePanelControls()
{
    if (!_hAppearancePanel) return;

    const int left = 20;
    const int top = 10;
    const int groupW = 360;

    // Edit field (numeric)
    createGroupBox(_hAppearancePanel, left, top, groupW, 60,
        IDC_CFG_GRP_EDITFIELD, TEXT("Edit field"));

    createStaticText(_hAppearancePanel, left + 12, top + 24, 140, 18,
        IDC_CFG_EDITFIELD_LABEL, TEXT("Edit field size (lines)"));

    createNumberEdit(_hAppearancePanel, left + 160, top + 22, 60, 22,
        IDC_CFG_EDITFIELD_SIZE_COMBO);

    // Transparency
    int transTop = top + 70;
    createGroupBox(_hAppearancePanel, left, transTop, groupW, 110,
        IDC_CFG_GRP_TRANSPARENCY, TEXT("Transparency"));

    createStaticText(_hAppearancePanel, left + 12, transTop + 24, 150, 18,
        IDC_CFG_FOREGROUND_LABEL, TEXT("Focused transparency"));
    HWND hFg = createSlider(_hAppearancePanel, left + 170, transTop + 20, 160, 26,
        IDC_CFG_FOREGROUND_SLIDER, 0, 255);

    createStaticText(_hAppearancePanel, left + 12, transTop + 64, 150, 18,
        IDC_CFG_BACKGROUND_LABEL, TEXT("Background transparency"));
    HWND hBg = createSlider(_hAppearancePanel, left + 170, transTop + 60, 160, 26,
        IDC_CFG_BACKGROUND_SLIDER, 0, 255);

    if (hFg) { ::SendMessage(hFg, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255)); }
    if (hBg) { ::SendMessage(hBg, TBM_SETRANGE, TRUE, MAKELPARAM(0, 255)); }
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

// --- Bind Settings -> Controls ---
void MultiReplaceConfigDialog::loadSettingsFromConfig()
{
    // load INI once, independent of panel lifetime
    auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths();
    ConfigManager::instance().load(iniFilePath);
    auto& CFG = ConfigManager::instance();

    // ---- Search & Replace ----
    if (_hSearchReplacePanel)
    {
        const bool stayAfterReplace = CFG.readBool(L"Options", L"StayAfterReplace", false);
        const bool allFromCursor = CFG.readBool(L"Options", L"AllFromCursor", false);
        const bool alertNotFound = CFG.readBool(L"Options", L"AlertNotFound", true);
        const bool highlightMatch = CFG.readBool(L"Options", L"HighlightMatch", true);

        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE, stayAfterReplace ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR, allFromCursor ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND, alertNotFound ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hSearchReplacePanel, IDC_CFG_HIGHLIGHT_MATCH, highlightMatch ? BST_CHECKED : BST_UNCHECKED);
    }

    // ---- List view / layout ----
    if (_hListViewLayoutPanel)
    {
        const bool listStats = CFG.readBool(L"Options", L"ListStatistics", false);
        const bool groupResults = CFG.readBool(L"Options", L"GroupResults", false);
        const bool flowNumericAlign = CFG.readBool(L"Options", L"FlowTabsNumericAlign", true);
        const int  headerLines = CFG.readInt(L"Scope", L"HeaderLines", 1);

        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED, listStats ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED, groupResults ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN, flowNumericAlign ? BST_CHECKED : BST_UNCHECKED);
        ::SetDlgItemInt(_hListViewLayoutPanel, IDC_CFG_HEADERLINES_EDIT, (UINT)headerLines, FALSE);

        // columns from INI (unchanged)
        const BOOL vFind = CFG.readBool(L"ListColumns", L"FindCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vReplace = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vComments = CFG.readBool(L"ListColumns", L"CommentsVisible", FALSE) ? BST_CHECKED : BST_UNCHECKED;
        const BOOL vDelete = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", TRUE) ? BST_CHECKED : BST_UNCHECKED;

        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE, vFind);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE, vReplace);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE, vComments);
        ::CheckDlgButton(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE, vDelete);
    }

    // ---- Interface / tooltips ----
    if (_hInterfaceTooltipsPanel)
    {
        const bool tooltips = CFG.readBool(L"Options", L"Tooltips", true);
        const bool dblClickEdits = CFG.readBool(L"Options", L"DoubleClickEdits", true);
        const bool hoverText = CFG.readBool(L"Options", L"HoverText", true);

        ::CheckDlgButton(_hInterfaceTooltipsPanel, IDC_CFG_TOOLTIPS_ENABLED, tooltips ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hInterfaceTooltipsPanel, IDC_CFG_DOUBLECLICK_EDITS, dblClickEdits ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(_hInterfaceTooltipsPanel, IDC_CFG_HOVER_TEXT_ENABLED, hoverText ? BST_CHECKED : BST_UNCHECKED);
    }

    // ---- Appearance ----
    if (_hAppearancePanel)
    {
        const int editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 12);
        ::SetDlgItemInt(_hAppearancePanel, IDC_CFG_EDITFIELD_SIZE_COMBO, (UINT)editFieldSize, FALSE);

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
    }

    // ---- Variables / automation ----
    if (_hVariablesAutomationPanel)
    {
        const bool luaSafe = CFG.readBool(L"Lua", L"SafeMode", false);
        ::CheckDlgButton(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED, luaSafe ? BST_CHECKED : BST_UNCHECKED);
    }
}

void MultiReplaceConfigDialog::applyConfigToSettings()
{
    // 1) Ensure INI is loaded (protect existing cache like history/pos)
    {
        auto [iniFilePath, _csv] = MultiReplace::generateConfigFilePaths(); // public helper
        ConfigManager::instance().load(iniFilePath); // fills cache
    } // :contentReference[oaicite:0]{index=0} :contentReference[oaicite:1]{index=1}

    MultiReplace::Settings s = MultiReplace::getSettings();

    // Search & Replace
    if (_hSearchReplacePanel)
    {
        s.stayAfterReplaceEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_STAY_AFTER_REPLACE) == BST_CHECKED);
        s.allFromCursorEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALL_FROM_CURSOR) == BST_CHECKED);
        s.alertNotFoundEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_ALERT_NOT_FOUND) == BST_CHECKED);
        s.highlightMatchEnabled = (::IsDlgButtonChecked(_hSearchReplacePanel, IDC_CFG_HIGHLIGHT_MATCH) == BST_CHECKED);
    }

    // List view / layout (+ header lines)
    if (_hListViewLayoutPanel)
    {
        s.listStatisticsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_LISTSTATISTICS_ENABLED) == BST_CHECKED);
        s.groupResultsEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_GROUPRESULTS_ENABLED) == BST_CHECKED);
        s.flowTabsNumericAlignEnabled = (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_FLOWTABS_NUMERIC_ALIGN) == BST_CHECKED);

        BOOL ok = FALSE;
        UINT v = ::GetDlgItemInt(_hListViewLayoutPanel, IDC_CFG_HEADERLINES_EDIT, &ok, FALSE);
        if (ok) { int hl = (int)v; if (hl < 0) hl = 0; if (hl > 999) hl = 999; s.csvHeaderLinesCount = hl; }

        // columns -> INI (cache only; save happens once at the end)
        auto& cm = ConfigManager::instance();
        cm.writeInt(L"ListColumns", L"FindCountVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_FINDCOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"ReplaceCountVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_REPLACECOUNT_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"CommentsVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_COMMENTS_VISIBLE) == BST_CHECKED) ? 1 : 0);
        cm.writeInt(L"ListColumns", L"DeleteButtonVisible", (::IsDlgButtonChecked(_hListViewLayoutPanel, IDC_CFG_DELETEBUTTON_VISIBLE) == BST_CHECKED) ? 1 : 0);
    } // :contentReference[oaicite:2]{index=2}

    // Interface / tooltips
    if (_hInterfaceTooltipsPanel)
    {
        s.tooltipsEnabled = (::IsDlgButtonChecked(_hInterfaceTooltipsPanel, IDC_CFG_TOOLTIPS_ENABLED) == BST_CHECKED);
        s.doubleClickEditsEnabled = (::IsDlgButtonChecked(_hInterfaceTooltipsPanel, IDC_CFG_DOUBLECLICK_EDITS) == BST_CHECKED);
        s.isHoverTextEnabled = (::IsDlgButtonChecked(_hInterfaceTooltipsPanel, IDC_CFG_HOVER_TEXT_ENABLED) == BST_CHECKED);
    }

    // Appearance (edit size + sliders)
    if (_hAppearancePanel)
    {
        BOOL ok = FALSE;
        UINT v = ::GetDlgItemInt(_hAppearancePanel, IDC_CFG_EDITFIELD_SIZE_COMBO, &ok, FALSE);
        if (ok) { int e = (int)v; if (e < 2) e = 2; if (e > 20) e = 20; s.editFieldSize = e; }

        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_FOREGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"ForegroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
        if (HWND h = ::GetDlgItem(_hAppearancePanel, IDC_CFG_BACKGROUND_SLIDER))
            ConfigManager::instance().writeInt(L"Window", L"BackgroundTransparency", (int)::SendMessage(h, TBM_GETPOS, 0, 0));
    }

    // Variables / automation
    if (_hVariablesAutomationPanel)
        s.luaSafeModeEnabled = (::IsDlgButtonChecked(_hVariablesAutomationPanel, IDC_CFG_LUA_SAFEMODE_ENABLED) == BST_CHECKED);

    // 2) Apply (live) + persist (cache)
    MultiReplace::applySettings(s);     // live toggles + hover INFOTIP etc. :contentReference[oaicite:3]{index=3}
    MultiReplace::persistSettings(s);   // writes keys into CFG cache only      :contentReference[oaicite:4]{index=4}

    // 3) Single save flush at the end (keeps untouched keys)
    ConfigManager::instance().save(L""); // writes full cache once               :contentReference[oaicite:5]{index=5}

    // 4) Live pull UI config (safe if panel exists)
    if (MultiReplace::instance)
        MultiReplace::instance->loadUIConfigFromIni(); // list columns/tooltips    :contentReference[oaicite:6]{index=6}
}
