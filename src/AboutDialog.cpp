//this file is part of notepad++
//Copyright (C)2023 Thomas Knoefel
//
//This program is free software; you can redistribute it and/or
//modify it under the terms of the GNU General Public License
//as published by the Free Software Foundation; either
//version 2 of the License, or (at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.
//
//You should have received a copy of the GNU General Public License
//along with this program; if not, write to the Free Software
//Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include "StaticDialog/resource.h"
#include "AboutDialog.h"

const char eT[] = { 102, 111, 114, 32, 65, 100, 114, 105, 97, 110, 32, 97, 110, 100, 32, 74, 117, 108, 105, 97, 110, 0 };
bool isDT = false;

LRESULT CALLBACK WebsiteLinkProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    UNREFERENCED_PARAMETER(dwRefData);

    PAINTSTRUCT ps;
    HDC hdc = NULL;
    HFONT hOldFont = NULL;
    HFONT hUnderlineFont = NULL;
    LOGFONT lf = { 0 };
    TCHAR szWinText[MAX_PATH];
    RECT rect;

    switch (uMsg)
    {
    case WM_PAINT:
    {
        hdc = BeginPaint(hwnd, &ps);
        SetTextColor(hdc, RGB(0, 0, 255));
        SetBkMode(hdc, TRANSPARENT);

        hOldFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
        if (hOldFont)
        {
            GetObject(hOldFont, sizeof(LOGFONT), &lf);
            hUnderlineFont = CreateFontIndirect(&lf);
            SelectObject(hdc, hUnderlineFont);
        }

        GetWindowText(hwnd, szWinText, _countof(szWinText));
        GetClientRect(hwnd, &rect);
        DrawText(hdc, szWinText, -1, &rect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);

        hOldFont&& SelectObject(hdc, hOldFont);
        hUnderlineFont&& DeleteObject(hUnderlineFont);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(0, 0, 255));
        SetBkMode(hdcStatic, TRANSPARENT);
        return (LRESULT)GetStockObject(NULL_BRUSH);
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

LRESULT CALLBACK NameStaticProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    UNREFERENCED_PARAMETER(dwRefData); 

    static TCHAR oT[MAX_PATH] = { 0 };
    switch (uMsg)
    {
    case WM_LBUTTONDBLCLK:
        if (GetKeyState(VK_CONTROL) & 0x8000)
        {
            if (!isDT)
            {
                wchar_t dT[MAX_PATH];
                int i = 0;
                for (; eT[i] != '\0'; ++i)
                {
                    dT[i] = static_cast<wchar_t>(eT[i]);
                }
                dT[i] = L'\0';

                if (oT[0] == 0)
                {
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

INT_PTR CALLBACK AboutDialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM /*lParam*/)
{
    switch (uMsg)
    {
    case WM_INITDIALOG:
    {
        HDC hdc = GetDC(hwnd);
        int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hwnd, hdc);

        int baseSize = 15;
        int fontSize = -MulDiv(baseSize, dpi, 96);

        static HFONT hFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, VARIABLE_PITCH, L"MS Shell Dlg");

        static HFONT hUnderlineFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, FALSE, TRUE, FALSE,
            DEFAULT_CHARSET, OUT_TT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, VARIABLE_PITCH, L"MS Shell Dlg");

        HWND hwndWebsiteLink = GetDlgItem(hwnd, IDC_WEBSITE_LINK);
        SendMessage(hwndWebsiteLink, WM_SETFONT, (WPARAM)hUnderlineFont, TRUE);
        RedrawWindow(hwndWebsiteLink, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);

        SetWindowSubclass(hwndWebsiteLink, WebsiteLinkProc, 0, reinterpret_cast<DWORD_PTR>(IDC_WEBSITE_LINK_VALUE));

        HWND controls[] = {
            GetDlgItem(hwnd, IDC_VERSION_STATIC),
            GetDlgItem(hwnd, IDC_AUTHOR_STATIC),
            GetDlgItem(hwnd, IDC_LICENSE_STATIC),
            GetDlgItem(hwnd, IDC_NAME_STATIC),
            GetDlgItem(hwnd, IDC_VERSION_LABEL),
            GetDlgItem(hwnd, IDC_AUTHOR_LABEL),
            GetDlgItem(hwnd, IDC_LICENSE_LABEL)
        };

        for (HWND hControl : controls)
        {
            if (hControl)
            {
                SendMessage(hControl, WM_SETFONT, (WPARAM)hFont, TRUE);
                RedrawWindow(hControl, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW);
            }
        }

        SetWindowSubclass(GetDlgItem(hwnd, IDC_NAME_STATIC), NameStaticProc, 0, 0);

        return TRUE;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK:
        case IDCANCEL:
            EndDialog(hwnd, 0);
            return TRUE;
        }
        break;

    case WM_CLOSE:
        EndDialog(hwnd, 0);
        return TRUE;
    }

    return FALSE;
}

void ShowAboutDialog(HWND hwndParent)
{
    DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUT_DIALOG), hwndParent, AboutDialogProc);
}