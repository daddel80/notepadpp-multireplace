// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel

#include "DPIManager.h"

// Constructor: Initializes DPI values.
DPIManager::DPIManager(HWND hwnd)
    : _hwnd(hwnd), _dpiX(120), _dpiY(120), _customScaleFactor(1.0f)
{
    init();
}

// Destructor: Currently not used, but can be expanded for resource management.
DPIManager::~DPIManager()
{
    // No dynamic resources to release at this time.
}

// Initializes the DPI values using Win32 APIs.
void DPIManager::init()
{
    UINT dpiX = 96, dpiY = 96; // Default DPI.

    // Attempt to load Shcore.dll for modern DPI functions.
    HMODULE hShcore = LoadLibrary(TEXT("Shcore.dll"));
    if (hShcore)
    {
        typedef HRESULT(WINAPI* GetDpiForMonitorFunc)(HMONITOR, MONITOR_DPI_TYPE, UINT*, UINT*);
        GetDpiForMonitorFunc pGetDpiForMonitor = (GetDpiForMonitorFunc)GetProcAddress(hShcore, "GetDpiForMonitor");

        if (pGetDpiForMonitor)
        {
            HMONITOR hMonitor = MonitorFromWindow(_hwnd, MONITOR_DEFAULTTONEAREST);
            if (SUCCEEDED(pGetDpiForMonitor(hMonitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY)))
            {
                // Successfully retrieved DPI values.
            }
        }

        FreeLibrary(hShcore);
    }
    else
    {
        // Fallback for older Windows versions.
        HDC hdc = GetDC(_hwnd);
        if (hdc)
        {
            dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(_hwnd, hdc);
        }
    }

    _dpiX = dpiX;
    _dpiY = dpiY;
}

// Updates the DPI values, typically called when DPI changes.
void DPIManager::updateDPI(HWND hwnd)
{
    _hwnd = hwnd; // Update window handle in case it has changed.
    init();       // Reinitialize DPI values.
}

// Scales a RECT structure.
void DPIManager::scaleRect(RECT* pRect) const
{
    if (pRect)
    {
        pRect->left = scaleX(pRect->left);
        pRect->right = scaleX(pRect->right);
        pRect->top = scaleY(pRect->top);
        pRect->bottom = scaleY(pRect->bottom);
    }
}

// Scales a POINT structure.
void DPIManager::scalePoint(POINT* pPoint) const
{
    if (pPoint)
    {
        pPoint->x = scaleX(pPoint->x);
        pPoint->y = scaleY(pPoint->y);
    }
}

// Scales a SIZE structure.
void DPIManager::scaleSize(SIZE* pSize) const
{
    if (pSize)
    {
        pSize->cx = scaleX(pSize->cx);
        pSize->cy = scaleY(pSize->cy);
    }
}
