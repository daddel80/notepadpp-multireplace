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

LRESULT CALLBACK WebsiteLinkProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (uMsg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = ::BeginPaint(hwnd, &ps);

        ::SetTextColor(hdc, RGB(0, 0, 255)); 
        ::SetBkMode(hdc, TRANSPARENT);

        TCHAR szWinText[MAX_PATH];
        ::GetWindowText(hwnd, szWinText, _countof(szWinText));
        RECT rect;
        ::GetClientRect(hwnd, &rect);
        ::DrawText(hdc, szWinText, -1, &rect, DT_SINGLELINE | DT_VCENTER);

        ::EndPaint(hwnd, &ps);

        return 0;
    }
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetBkMode(hdcStatic, TRANSPARENT);
        return (LRESULT)GetStockObject(NULL_BRUSH);
    }
    case WM_SETCURSOR:
        SetCursor(LoadCursor(NULL, IDC_HAND));
        return TRUE;
    case WM_LBUTTONUP:
    {
        RECT rect;
        GetClientRect(hwnd, &rect);
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        if (PtInRect(&rect, pt))
        {
            ShellExecute(NULL, TEXT("open"), reinterpret_cast<const TCHAR*>(dwRefData), NULL, NULL, SW_SHOWNORMAL);
        }
        return TRUE;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, WebsiteLinkProc, uIdSubclass);
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
        HFONT hUnderlineFont = CreateFont(0, 0, 0, 0, FW_NORMAL, FALSE, TRUE, FALSE, 0, 0, 0, 0, 0, NULL);
        HWND hwndWebsiteLink = GetDlgItem(hwnd, IDC_WEBSITE_LINK);
        SendMessage(hwndWebsiteLink, WM_SETFONT, (WPARAM)hUnderlineFont, TRUE);
        SetWindowSubclass(hwndWebsiteLink, WebsiteLinkProc, 0, reinterpret_cast<DWORD_PTR>(IDC_WEBSITE_LINK_VALUE));

        DeleteObject(hUnderlineFont);
        HFONT hFont = CreateFont(8, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"MS Shell Dlg");
        SendMessage(GetDlgItem(hwnd, IDC_VERSION_STATIC), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, IDC_AUTHOR_STATIC), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, IDC_LICENSE_STATIC), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, IDC_NAME_STATIC), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, IDC_WEBSITE_LINK), WM_SETFONT, (WPARAM)hFont, TRUE);
        DeleteObject(hFont);

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


// Show the About dialog
void ShowAboutDialog(HWND hwndParent)
{
    DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUT_DIALOG), hwndParent, AboutDialogProc);
}
