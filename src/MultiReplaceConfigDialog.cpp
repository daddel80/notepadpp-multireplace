#include "MultiReplaceConfigDialog.h"
#include "PluginDefinition.h"
#include "Notepad_plus_msgs.h"
#include "StaticDialog/resource.h"
#include "NppStyleKit.h"

extern NppData nppData;

MultiReplaceConfigDialog::~MultiReplaceConfigDialog()
{
    delete dpiMgr;
    dpiMgr = nullptr;
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
        resizeUI();

        ::SendMessage(
            nppData._nppHandle,
            NPPM_DARKMODESUBCLASSANDTHEME,
            (WPARAM)NppDarkMode::dmfInit,
            (LPARAM)_hSelf);

        return TRUE;
    }

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
            display(false);
            return TRUE;
        }
        break;

    case WM_CLOSE:
        display(false);
        return TRUE;
    }

    return FALSE;
}

void MultiReplaceConfigDialog::createUI()
{
    _hCategoryList = ::CreateWindowEx(
        0, WC_LISTBOX, TEXT(""),
        WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_BORDER | WS_TABSTOP,
        0, 0, 0, 0,
        _hSelf,
        (HMENU)IDC_CONFIG_CATEGORY_LIST,
        _hInst,
        nullptr);

    _hGeneralPanel = ::CreateWindowEx(
        0, WC_STATIC, TEXT("General settings."),
        WS_CHILD,
        0, 0, 0, 0,
        _hSelf,
        nullptr,
        _hInst,
        nullptr);

    _hResultPanel = ::CreateWindowEx(
        0, WC_STATIC, TEXT("Result dock settings."),
        WS_CHILD,
        0, 0, 0, 0,
        _hSelf,
        nullptr,
        _hInst,
        nullptr);

    _hCloseButton = ::CreateWindowEx(
        0,
        WC_BUTTON,
        TEXT("Close"),
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        0, 0, 0, 0,
        _hSelf,
        (HMENU)IDCANCEL,
        _hInst,
        nullptr);
}



void MultiReplaceConfigDialog::initUI()
{
    ::SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("General"));
    ::SendMessage(_hCategoryList, LB_ADDSTRING, 0, (LPARAM)TEXT("Result Dock"));
    ::SendMessage(_hCategoryList, LB_SETCURSEL, 0, 0);

    showCategory(0);
}

void MultiReplaceConfigDialog::showCategory(int index)
{
    ::ShowWindow(_hGeneralPanel, index == 0 ? SW_SHOW : SW_HIDE);
    ::ShowWindow(_hResultPanel, index == 1 ? SW_SHOW : SW_HIDE);
}

void MultiReplaceConfigDialog::resizeUI()
{
    if (!dpiMgr)
        return;

    RECT rc;
    ::GetClientRect(_hSelf, &rc);

    const int margin = dpiMgr->scaleX(8);
    const int catW = dpiMgr->scaleX(120);
    const int btnW = dpiMgr->scaleX(80);
    const int btnH = dpiMgr->scaleY(23);

    const int clientW = rc.right - rc.left;
    const int clientH = rc.bottom - rc.top;

    // Close-Button unten mittig
    const int btnX = (clientW - btnW) / 2;
    const int btnY = clientH - margin - btnH;

    if (_hCloseButton)
    {
        ::MoveWindow(_hCloseButton, btnX, btnY, btnW, btnH, TRUE);
    }

    // Verbleibender Bereich oberhalb des Buttons
    const int contentTop = margin;
    const int contentBottom = btnY - margin;
    int contentHeight = contentBottom - contentTop;
    if (contentHeight < 0)
        contentHeight = 0;

    // Kategorienliste links
    ::MoveWindow(
        _hCategoryList,
        margin,
        contentTop,
        catW,
        contentHeight,
        TRUE);

    // Panels rechts
    const int panelLeft = margin + catW + margin;
    int panelWidth = clientW - panelLeft - margin;
    if (panelWidth < 0)
        panelWidth = 0;

    int panelHeight = contentHeight;

    ::MoveWindow(
        _hGeneralPanel,
        panelLeft,
        contentTop,
        panelWidth,
        panelHeight,
        TRUE);

    ::MoveWindow(
        _hResultPanel,
        panelLeft,
        contentTop,
        panelWidth,
        panelHeight,
        TRUE);
}

