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

#include "AboutDialog.h"
#include "StaticDialog/resource.h"
#include "EditorPalette.h"
#include "LanguageManager.h"
#include "MultiReplacePanel.h"
#include "NppStyleKit.h"
#include "Version.h"
#include <commctrl.h>
#include <shellapi.h>
#include <vector>

namespace {
    // UI Constants
    constexpr int kBaseFontSize = 13;
    constexpr COLORREF kLinkColorLight = RGB(0, 102, 204);
    constexpr COLORREF kLinkColorDark = RGB(100, 149, 237);
    constexpr UINT_PTR kNameTimerId = 1;

    constexpr unsigned char kTextKey = 0xA7;
    const unsigned char eT[] = { 0xC1, 0xC8, 0xD5, 0x87, 0xE6, 0xC3, 0xD5, 0xCE, 0xC6, 0xC9, 0x87, 0xC6, 0xC9, 0xC3, 0x87, 0xED, 0xD2, 0xCB, 0xCE, 0xC6, 0xC9 };
    const unsigned char eN[] = { 0xE0, 0xC6, 0xC5, 0xD5, 0xCE, 0xC2, 0xCB, 0xC6 };
    constexpr int kFirstHoldTicks = 34;   // the text in between stands about 1.2 seconds

    template <size_t N>
    std::wstring decodeText(const unsigned char (&encoded)[N])
    {
        std::wstring text;
        for (const unsigned char c : encoded) {
            text += static_cast<wchar_t>(c ^ kTextKey);
        }
        return text;
    }

    std::wstring windowText(HWND hwnd)
    {
        std::wstring text(static_cast<size_t>(GetWindowTextLength(hwnd)), L'\0');
        GetWindowText(hwnd, text.data(), static_cast<int>(text.size()) + 1);
        return text;
    }

    int textWidth(HDC hdc, std::wstring_view text)
    {
        SIZE size{};
        if (!text.empty()) {
            GetTextExtentPoint32(hdc, text.data(), static_cast<int>(text.size()), &size);
        }
        return size.cx;
    }

    // Background and text color as the dialog gives a plain static, light or dark, plus the control's font;
    // returns the previous font
    HGDIOBJ beginOwnerDraw(const DRAWITEMSTRUCT& dis)
    {
        const auto hBrush = reinterpret_cast<HBRUSH>(SendMessage(GetParent(dis.hwndItem), WM_CTLCOLORSTATIC,
            reinterpret_cast<WPARAM>(dis.hDC), reinterpret_cast<LPARAM>(dis.hwndItem)));
        FillRect(dis.hDC, &dis.rcItem, hBrush ? hBrush : GetSysColorBrush(COLOR_3DFACE));
        SetBkMode(dis.hDC, TRANSPARENT);
        return SelectObject(dis.hDC, reinterpret_cast<HFONT>(SendMessage(dis.hwndItem, WM_GETFONT, 0, 0)));
    }

    // Draws into a bitmap and copies it in one go, so a control redrawn many times a second does not flicker;
    // without a bitmap it draws directly
    template <typename Draw>
    void drawBuffered(const DRAWITEMSTRUCT& dis, Draw draw)
    {
        const int width = dis.rcItem.right - dis.rcItem.left;
        const int height = dis.rcItem.bottom - dis.rcItem.top;
        const HDC hdc = CreateCompatibleDC(dis.hDC);
        const HBITMAP hBitmap = hdc ? CreateCompatibleBitmap(dis.hDC, width, height) : nullptr;
        if (!hBitmap) {
            if (hdc) DeleteDC(hdc);
            draw(dis);
            return;
        }
        const HGDIOBJ hOldBitmap = SelectObject(hdc, hBitmap);
        DRAWITEMSTRUCT buffered = dis;
        buffered.hDC = hdc;
        buffered.rcItem = { 0, 0, width, height };
        draw(buffered);
        BitBlt(dis.hDC, dis.rcItem.left, dis.rcItem.top, width, height, hdc, 0, 0, SRCCOPY);
        SelectObject(hdc, hOldBitmap);
        DeleteObject(hBitmap);
        DeleteDC(hdc);
    }
}

AboutDialog::~AboutDialog()
{
    if (_hFont) {
        DeleteObject(_hFont);
        _hFont = nullptr;
    }
    if (_hUnderlineFont) {
        DeleteObject(_hUnderlineFont);
        _hUnderlineFont = nullptr;
    }
}

void AboutDialog::doDialog()
{
    if (!isCreated()) {
        create(IDD_ABOUT_DIALOG);
    }
    goToCenter();
    display(true);
}

LRESULT CALLBACK AboutDialog::WebsiteLinkProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    (void)dwRefData;

    switch (uMsg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
        COLORREF linkColor = isDark ? kLinkColorDark : kLinkColorLight;

        SetTextColor(hdc, linkColor);
        SetBkMode(hdc, TRANSPARENT);

        HFONT hFont = reinterpret_cast<HFONT>(SendMessage(hwnd, WM_GETFONT, 0, 0));
        HFONT hOldFont = nullptr;
        if (hFont) {
            hOldFont = static_cast<HFONT>(SelectObject(hdc, hFont));
        }

        TCHAR szText[MAX_PATH];
        GetWindowText(hwnd, szText, _countof(szText));
        RECT rect;
        GetClientRect(hwnd, &rect);

        DrawText(hdc, szText, -1, &rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER);

        if (hOldFont) {
            SelectObject(hdc, hOldFont);
        }

        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_SETCURSOR:
        SetCursor(LoadCursor(NULL, IDC_HAND));
        return TRUE;

    case WM_LBUTTONUP:
        ShellExecute(NULL, TEXT("open"), IDC_WEBSITE_LINK_VALUE, NULL, NULL, SW_SHOWNORMAL);
        return TRUE;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, WebsiteLinkProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK AboutDialog::NameStaticProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (uMsg)
    {
    case WM_LBUTTONDBLCLK:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            reinterpret_cast<AboutDialog*>(dwRefData)->swapName();
        }
        return TRUE;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, NameStaticProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

void AboutDialog::createFonts()
{
    // Cleanup old fonts
    if (_hFont) {
        DeleteObject(_hFont);
        _hFont = nullptr;
    }
    if (_hUnderlineFont) {
        DeleteObject(_hUnderlineFont);
        _hUnderlineFont = nullptr;
    }

    // Get current DPI
    HDC hdc = GetDC(_hSelf);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(_hSelf, hdc);

    int fontSize = -MulDiv(kBaseFontSize, dpi, 96);

    _hFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, VARIABLE_PITCH, L"MS Shell Dlg");

    _hUnderlineFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, FALSE, TRUE, FALSE,
        DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, VARIABLE_PITCH, L"MS Shell Dlg");
}

void AboutDialog::applyFonts()
{
    // Apply underline font to website link
    HWND hwndWebsiteLink = GetDlgItem(_hSelf, IDC_WEBSITE_LINK);
    if (hwndWebsiteLink) {
        SendMessage(hwndWebsiteLink, WM_SETFONT, reinterpret_cast<WPARAM>(_hUnderlineFont), TRUE);
    }

    // Apply regular font to static controls
    const int controlIDs[] = {
        IDC_VERSION_STATIC,
        IDC_LICENSE_STATIC,
        IDC_NAME_STATIC,
        IDC_STATUS_STATIC,
        IDC_VERSION_LABEL,
        IDC_AUTHOR_LABEL,
        IDC_LICENSE_LABEL
    };

    for (int id : controlIDs) {
        HWND hControl = GetDlgItem(_hSelf, id);
        if (hControl) {
            SendMessage(hControl, WM_SETFONT, reinterpret_cast<WPARAM>(_hFont), TRUE);
        }
    }
}

void AboutDialog::swapName()
{
    if (_nameSwap.isActive()) {
        return;
    }

    LanguageManager& LM = LanguageManager::instance();
    HWND hwndName = GetDlgItem(_hSelf, IDC_NAME_STATIC);
    const std::wstring text = windowText(hwndName);
    std::vector<TextReplaceAnimation::Step> steps;
    if (_nameSwapped) {
        steps.push_back({ text, _name, LM.get(L"status_replace_one") });
    }
    else {
        // Two list entries, the second works on what the first left
        const std::wstring firstName = _name.substr(0, _name.find(L' '));
        const std::wstring otherName = decodeText(eN);
        steps.push_back({ firstName, otherName, LM.get(L"status_replace_one"), kFirstHoldTicks });
        steps.push_back({ otherName + _name.substr(firstName.size()), decodeText(eT), LM.get(L"status_occurrences_replaced", { L"2" }) });
    }
    _nameSwapped = !_nameSwapped;
    _nameSwap.start(text, std::move(steps));

    BOOL animate = TRUE;
    SystemParametersInfo(SPI_GETCLIENTAREAANIMATION, 0, &animate, 0);
    if (!animate || !_nameSwap.isActive()) {
        _nameSwap.stop();
        SetWindowText(hwndName, _nameSwap.result().c_str());
        return;
    }

    SetTimer(_hSelf, kNameTimerId, TextReplaceAnimation::TICK_MS, nullptr);
    InvalidateRect(hwndName, nullptr, FALSE);
}

void AboutDialog::advanceNameSwap()
{
    switch (_nameSwap.tick()) {
    case TextReplaceAnimation::Tick::Over:
        endNameSwap();
        return;
    case TextReplaceAnimation::Tick::StepDone:
    {
        // The step's result stands, its status takes the place of the website link
        HWND hwndStatus = GetDlgItem(_hSelf, IDC_STATUS_STATIC);
        SetWindowText(hwndStatus, _nameSwap.step().status.c_str());
        InvalidateRect(hwndStatus, nullptr, FALSE);
        showSwapStatus(true);
        break;
    }
    case TextReplaceAnimation::Tick::Running:
        break;
    }
    InvalidateRect(GetDlgItem(_hSelf, IDC_NAME_STATIC), nullptr, FALSE);
}

void AboutDialog::endNameSwap()
{
    KillTimer(_hSelf, kNameTimerId);
    _nameSwap.stop();
    HWND hwndName = GetDlgItem(_hSelf, IDC_NAME_STATIC);
    SetWindowText(hwndName, _nameSwap.result().c_str());
    showSwapStatus(false);
    InvalidateRect(hwndName, nullptr, FALSE);
}

void AboutDialog::showSwapStatus(bool show)
{
    ShowWindow(GetDlgItem(_hSelf, IDC_WEBSITE_LINK), show ? SW_HIDE : SW_SHOW);
    ShowWindow(GetDlgItem(_hSelf, IDC_STATUS_STATIC), show ? SW_SHOW : SW_HIDE);
}

void AboutDialog::drawName(const DRAWITEMSTRUCT& dis) const
{
    const HGDIOBJ hOldFont = beginOwnerDraw(dis);

    // The background just painted, read before any text covers it
    COLORREF background = GetPixel(dis.hDC, dis.rcItem.left, dis.rcItem.top);
    if (background == CLR_INVALID) {
        background = GetBkColor(dis.hDC);
    }

    const std::wstring text = _nameSwap.isActive() ? std::wstring() : windowText(dis.hwndItem);
    const std::wstring_view marked = _nameSwap.marked();
    const int top = dis.rcItem.top;
    int x = dis.rcItem.left;
    const auto draw = [&dis, top, &x](std::wstring_view part) {
        if (!part.empty()) {
            TextOut(dis.hDC, x, top, part.data(), static_cast<int>(part.size()));
        }
        x += textWidth(dis.hDC, part);
    };

    draw(_nameSwap.isActive() ? _nameSwap.before() : std::wstring_view(text));
    if (!marked.empty()) {
        // Marked like a Mark Matches hit, in the list color of the step's find text
        const EditorPalette::Theme theme = EditorPalette::makeTheme(background, GetTextColor(dis.hDC));
        const COLORREF mark = EditorPalette::blend(MultiReplace::matchMarkColor(_nameSwap.step().find, theme), background,
            theme.dark ? MultiReplace::EDITOR_MARK_ALPHA_DARK : MultiReplace::EDITOR_MARK_ALPHA_LIGHT);

        TEXTMETRIC tm{};
        GetTextMetrics(dis.hDC, &tm);
        SetDCBrushColor(dis.hDC, mark);
        SetDCPenColor(dis.hDC, mark);
        const HGDIOBJ hOldBrush = SelectObject(dis.hDC, GetStockObject(DC_BRUSH));
        const HGDIOBJ hOldPen = SelectObject(dis.hDC, GetStockObject(DC_PEN));
        const int corner = tm.tmHeight / 4;
        RoundRect(dis.hDC, x, top, x + textWidth(dis.hDC, marked), top + tm.tmHeight, corner, corner);
        SelectObject(dis.hDC, hOldPen);
        SelectObject(dis.hDC, hOldBrush);
    }
    draw(marked);
    draw(_nameSwap.after());

    SelectObject(dis.hDC, hOldFont);
}

void AboutDialog::drawStatus(const DRAWITEMSTRUCT& dis) const
{
    const HGDIOBJ hOldFont = beginOwnerDraw(dis);
    const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
    SetTextColor(dis.hDC, isDark ? MultiReplace::DMODE_SUCCESS : MultiReplace::LMODE_SUCCESS);

    RECT rect = dis.rcItem;
    DrawText(dis.hDC, windowText(dis.hwndItem).c_str(), -1, &rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX);

    SelectObject(dis.hDC, hOldFont);
}


intptr_t CALLBACK AboutDialog::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        LanguageManager& LM = LanguageManager::instance();

        SetWindowText(_hSelf, LM.getLPCW(L"about_title"));
        SetDlgItemText(_hSelf, IDC_VERSION_LABEL, LM.getLPCW(L"about_version"));
        SetDlgItemText(_hSelf, IDC_AUTHOR_LABEL, LM.getLPCW(L"about_author"));
        SetDlgItemText(_hSelf, IDC_LICENSE_LABEL, LM.getLPCW(L"about_license"));
        SetDlgItemText(_hSelf, IDC_WEBSITE_LINK, LM.getLPCW(L"about_help_support"));
        SetDlgItemText(_hSelf, IDOK, LM.getLPCW(L"about_ok"));
        SetDlgItemText(_hSelf, IDC_VERSION_STATIC, VER_WSTRING);

        createFonts();
        applyFonts();

        // Setup website link subclass
        HWND hwndWebsiteLink = GetDlgItem(_hSelf, IDC_WEBSITE_LINK);
        if (hwndWebsiteLink) {
            SetWindowSubclass(hwndWebsiteLink, WebsiteLinkProc, 0, 0);
        }

        // Setup name static for special interaction
        HWND hwndName = GetDlgItem(_hSelf, IDC_NAME_STATIC);
        if (hwndName) {
            _name = windowText(hwndName);
            SetWindowSubclass(hwndName, NameStaticProc, 0, reinterpret_cast<DWORD_PTR>(this));
        }

        // Enable Dark Mode support
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME,
            static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(_hSelf));

        return TRUE;
    }

    case WM_DPICHANGED:
    {
        createFonts();
        applyFonts();

        RECT* pRect = reinterpret_cast<RECT*>(lParam);
        if (pRect) {
            SetWindowPos(_hSelf, nullptr, pRect->left, pRect->top,
                pRect->right - pRect->left, pRect->bottom - pRect->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return TRUE;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_WEBSITE_LINK)) {
            bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
            COLORREF linkColor = isDark ? kLinkColorDark : kLinkColorLight;
            SetTextColor(hdcStatic, linkColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
        }
        break;
    }

    case WM_DRAWITEM:
    {
        const auto* dis = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam);
        if (dis->CtlID == IDC_NAME_STATIC) {
            drawBuffered(*dis, [this](const DRAWITEMSTRUCT& buffered) { drawName(buffered); });   // the animation redraws it every tick
            return TRUE;
        }
        if (dis->CtlID == IDC_STATUS_STATIC) {
            drawStatus(*dis);
            return TRUE;
        }
        break;
    }

    case WM_TIMER:
        if (wParam == kNameTimerId) {
            advanceNameSwap();
            return TRUE;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            if (_nameSwap.isActive()) {
                endNameSwap();
            }
            display(false);
            return TRUE;
        }
        break;
    }

    return FALSE;
}