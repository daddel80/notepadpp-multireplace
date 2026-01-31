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
#include "LanguageManager.h"
#include "NppStyleKit.h"
#include <commctrl.h>
#include <shellapi.h>

namespace {
    // UI Constants
    constexpr int kBaseFontSize = 13;
    constexpr COLORREF kLinkColorLight = RGB(0, 102, 204);
    constexpr COLORREF kLinkColorDark = RGB(100, 149, 237);

    const char eT[] = { 102, 111, 114, 32, 65, 100, 114, 105, 97, 110, 32, 97, 110, 100, 32, 74, 117, 108, 105, 97, 110, 0 };
    bool isDT = false;
    TCHAR oT[MAX_PATH] = { 0 };
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
    (void)dwRefData;
    (void)wParam;
    (void)lParam;

    switch (uMsg)
    {
    case WM_LBUTTONDBLCLK:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            if (!isDT)
            {
                wchar_t dT[MAX_PATH];
                int i = 0;
                for (; eT[i] != '\0'; ++i) {
                    dT[i] = static_cast<wchar_t>(eT[i]);
                }
                dT[i] = L'\0';

                if (oT[0] == 0) {
                    GetWindowText(hwnd, oT, _countof(oT));
                }

                SetWindowText(hwnd, dT);
                isDT = true;
            }
            else
            {
                SetWindowText(hwnd, oT);
                isDT = false;
            }
        }
        return TRUE;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, NameStaticProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
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

        // Font Setup with DPI awareness
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

        // Setup website link
        HWND hwndWebsiteLink = GetDlgItem(_hSelf, IDC_WEBSITE_LINK);
        if (hwndWebsiteLink) {
            SendMessage(hwndWebsiteLink, WM_SETFONT, reinterpret_cast<WPARAM>(_hUnderlineFont), TRUE);
            SetWindowSubclass(hwndWebsiteLink, WebsiteLinkProc, 0, 0);
        }

        // Apply font to static controls
        const int controlIDs[] = {
            IDC_VERSION_STATIC,
            IDC_AUTHOR_STATIC,
            IDC_LICENSE_STATIC,
            IDC_NAME_STATIC,
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

        // Setup name static for special interaction
        HWND hwndName = GetDlgItem(_hSelf, IDC_NAME_STATIC);
        if (hwndName) {
            SetWindowSubclass(hwndName, NameStaticProc, 0, 0);
        }

        // Enable Dark Mode support
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME,
            static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(_hSelf));

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

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            display(false);
            return TRUE;
        }
        break;
    }

    return FALSE;
}