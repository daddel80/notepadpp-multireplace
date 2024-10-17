// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel

#ifndef DPIMANAGER_H
#define DPIMANAGER_H

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

    // Setter for custom scale factor
    void DPIManager::setCustomScaleFactor(float scaleFactor) {
        _customScaleFactor = std::clamp(scaleFactor, 0.5f, 2.0f);
    }

    // Retrieves the current DPI values.
    int getDPIX() const { return _dpiX; }
    int getDPIY() const { return _dpiY; }

    // Converts raw pixels to scaled pixels, includes the custom scale factor.
    int scaleX(int x) const { return static_cast<int>(MulDiv(x, _dpiX, 96) * _customScaleFactor); }
    int scaleY(int y) const { return static_cast<int>(MulDiv(y, _dpiY, 96) * _customScaleFactor); }

    // Converts scaled pixels back to raw pixels.
    int unscaleX(int x) const { return MulDiv(x, 96, _dpiX); }
    int unscaleY(int y) const { return MulDiv(y, 96, _dpiY); }

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
    HWND _hwnd;  // Handle to the window.
    int _dpiX;   // Horizontal DPI.
    int _dpiY;   // Vertical DPI.
    float _customScaleFactor; // Custom scaling factor stored INI file.
    bool _isSystemMetricsForDpiSupported;  // To check if GetSystemMetricsForDpi is supported
    decltype(GetSystemMetricsForDpi)* _pGetSystemMetricsForDpi; // Pointer to the function

    // Initializes the DPI values.
    void init();
};

#endif // DPIMANAGER_H
