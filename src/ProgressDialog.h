#pragma once
#ifndef PROGRESSDIALOG_H
#define PROGRESSDIALOG_H

#include "PluginDefinition.h"
#include <sstream>
#include <windows.h>
#include <CommCtrl.h>

class ProgressBarWindow
{
private:
    HWND hwnd;
    HWND progressbar;
    HWND parentWindow; 
    static const int PROGRESSBAR_ID = 1;

public:
    ProgressBarWindow(HWND parent)
        : parentWindow(parent)
    {

        WNDCLASS wc = { 0 };
        wc.lpfnWndProc = WndProc;
        wc.hInstance = GetModuleHandle(NULL);
        wc.lpszClassName = L"ProgressBarWindowClass";
        RegisterClass(&wc);
        hwnd = CreateWindowEx(WS_EX_TOPMOST, L"ProgressBarWindowClass", L"Progress", WS_OVERLAPPEDWINDOW | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 200, 100, NULL, NULL, GetModuleHandle(NULL), NULL);

        InitCommonControls();
        progressbar = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE, 10, 10, 180, 30, hwnd, reinterpret_cast<HMENU>(static_cast<size_t>(PROGRESSBAR_ID)), GetModuleHandle(NULL), NULL);
    }

    void CenterWindow()
    {
        RECT rectParent;
        GetWindowRect(parentWindow, &rectParent);

        int width = rectParent.right - rectParent.left;
        int height = rectParent.bottom - rectParent.top;

        int x = rectParent.left + (width - 200) / 2;
        int y = rectParent.top + (height - 100) / 2;

        MoveWindow(hwnd, x, y, 200, 100, TRUE);
    }

    void Show(int nCmdShow)
    {
        ShowWindow(hwnd, nCmdShow);
    }

    void UpdateProgress(int progress)
    {
        SendMessage(progressbar, PBM_SETPOS, (WPARAM)progress, 0);
    }

    void RunMessageLoop()
    {
        MSG msg = { 0 };
        while (GetMessage(&msg, NULL, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        switch (uMsg)
        {
        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
        }
    }
};

#endif // PROGRESSDIALOG_H
