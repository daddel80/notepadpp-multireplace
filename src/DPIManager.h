// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel

#ifndef DPIMANAGER_H
#define DPIMANAGER_H

#include <windows.h>
#include <shellscalingapi.h>

// DPIManager class handles DPI awareness and scaling for UI elements.
class DPIManager
{
public:
    // Constructor initializes DPI values based on the provided window handle.
    DPIManager(HWND hwnd);

    // Destructor (if needed for resource management).
    ~DPIManager();

    // Retrieves the current DPI values.
    int getDPIX() const { return _dpiX; }
    int getDPIY() const { return _dpiY; }

    // Converts raw pixels to scaled pixels.
    int scaleX(int x) const { return MulDiv(x, _dpiX, 96); }
    int scaleY(int y) const { return MulDiv(y, _dpiY, 96); }

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

    int _dpiX;   // Horizontal DPI.
    int _dpiY;   // Vertical DPI.

private:
    HWND _hwnd;  // Handle to the window.


    // Initializes the DPI values.
    void init();
};

#endif // DPIMANAGER_H
