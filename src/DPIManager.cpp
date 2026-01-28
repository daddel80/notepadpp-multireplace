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

#include "DPIManager.h"

// Constructor: Initializes DPI values.
DPIManager::DPIManager(HWND hwnd)
    : _hwnd(hwnd)
    , _dpiX(96)
    , _dpiY(96)
    , _customScaleFactor(1.0f)
    , _isSystemMetricsForDpiSupported(false)
    , _pGetSystemMetricsForDpi(nullptr)
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
    UINT dpiX = 96, dpiY = 96;  // Default DPI (96 is the standard base DPI)

    // Step 1: Try to load Shcore.dll for modern DPI API (Windows 8.1+)
    HMODULE hShcore = LoadLibrary(TEXT("Shcore.dll"));
    if (hShcore)
    {
        typedef HRESULT(WINAPI* GetDpiForMonitorFunc)(HMONITOR, MONITOR_DPI_TYPE, UINT*, UINT*);
        GetDpiForMonitorFunc pGetDpiForMonitor =
            (GetDpiForMonitorFunc)GetProcAddress(hShcore, "GetDpiForMonitor");

        if (pGetDpiForMonitor)
        {
            HMONITOR hMonitor = MonitorFromWindow(_hwnd, MONITOR_DEFAULTTONEAREST);
            if (SUCCEEDED(pGetDpiForMonitor(hMonitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY)))
            {
                // Successfully retrieved DPI values from modern API
            }
            else
            {
                // If GetDpiForMonitor fails, fallback to GetDeviceCaps
                HDC hdc = GetDC(_hwnd);
                if (hdc)
                {
                    dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                    dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                    ReleaseDC(_hwnd, hdc);
                }
            }
        }

        FreeLibrary(hShcore);
    }
    else
    {
        // If Shcore.dll isn't available (Windows 7), use GetDeviceCaps directly
        HDC hdc = GetDC(_hwnd);
        if (hdc)
        {
            dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(_hwnd, hdc);
        }
    }

    // Step 2: Check for GetSystemMetricsForDpi support (Windows 10 1607+)
    // Use GetModuleHandle since User32.dll is always loaded
    HMODULE hUser32 = GetModuleHandle(TEXT("User32.dll"));
    if (hUser32)
    {
        _pGetSystemMetricsForDpi =
            (decltype(GetSystemMetricsForDpi)*)GetProcAddress(hUser32, "GetSystemMetricsForDpi");
        _isSystemMetricsForDpiSupported = (_pGetSystemMetricsForDpi != nullptr);
        // No FreeLibrary needed for GetModuleHandle
    }

    // Store the DPI values
    _dpiX = dpiX;
    _dpiY = dpiY;
}

// Function for custom metrics with fallback.
int DPIManager::getCustomMetricOrFallback(int nIndex, UINT dpi, int fallbackValue) const
{
    if (_isSystemMetricsForDpiSupported && _pGetSystemMetricsForDpi)
    {
        return _pGetSystemMetricsForDpi(nIndex, dpi);
    }

    // Return the fallback value if not supported
    return fallbackValue;
}

// Updates the DPI values, typically called when DPI changes.
void DPIManager::updateDPI(HWND hwnd)
{
    _hwnd = hwnd;  // Update window handle in case it has changed.
    init();        // Reinitialize DPI values.
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