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

#pragma once

#include <windows.h>
#include <shellscalingapi.h>
#include <algorithm>

// DPIManager class handles DPI awareness and scaling for UI elements.
class DPIManager
{
public:
    // Constructor initializes DPI values based on the provided window handle.
    DPIManager(HWND hwnd);

    // Destructor (if needed for resource management).
    ~DPIManager();

    // Getter for custom scale factor
    float getCustomScaleFactor() const { return _customScaleFactor; }

    // Setter for custom scale factor (clamped to 50%-200%)
    void setCustomScaleFactor(float scaleFactor) {
        _customScaleFactor = std::clamp(scaleFactor, 0.5f, 2.0f);
    }

    // Retrieves the current DPI values.
    int getDPIX() const { return _dpiX; }
    int getDPIY() const { return _dpiY; }

    // Converts raw pixels to scaled pixels, includes the custom scale factor.
    int scaleX(int x) const { return static_cast<int>(MulDiv(x, _dpiX, 96) * _customScaleFactor); }
    int scaleY(int y) const { return static_cast<int>(MulDiv(y, _dpiY, 96) * _customScaleFactor); }

    // Converts scaled pixels back to raw pixels, includes the custom scale factor.
    int unscaleX(int x) const { return static_cast<int>(MulDiv(x, 96, _dpiX) / _customScaleFactor); }
    int unscaleY(int y) const { return static_cast<int>(MulDiv(y, 96, _dpiY) / _customScaleFactor); }

    // Scales a RECT structure.
    void scaleRect(RECT* pRect) const;

    // Scales a POINT structure.
    void scalePoint(POINT* pPoint) const;

    // Scales a SIZE structure.
    void scaleSize(SIZE* pSize) const;

    // Updates DPI values (e.g., after a DPI change event).
    void updateDPI(HWND hwnd);

    // Function for custom metrics with fallback.
    int getCustomMetricOrFallback(int nIndex, UINT dpi, int fallbackValue) const;

private:
    HWND _hwnd;   // Handle to the window.
    int _dpiX;    // Horizontal DPI.
    int _dpiY;    // Vertical DPI.
    float _customScaleFactor;  // Custom scaling factor stored in INI file.
    bool _isSystemMetricsForDpiSupported;  // To check if GetSystemMetricsForDpi is supported
    decltype(GetSystemMetricsForDpi)* _pGetSystemMetricsForDpi;  // Pointer to the function

    // Initializes the DPI values.
    void init();
};
