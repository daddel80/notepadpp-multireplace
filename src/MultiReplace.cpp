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

#include "PluginDefinition.h"
#include "MultiReplacePanel.h"
#include "MultiReplaceConfigDialog.h"
#include "image_data.h"
#include <algorithm>
#include <vector>
#include <windows.h>
#include "ResultDock.h"

extern FuncItem funcItem[nbFunc];
extern NppData nppData;
extern MultiReplaceConfigDialog _MultiReplaceConfig;

HINSTANCE g_inst;

void CalculateIconSize(UINT dpi, int& iconWidth, int& iconHeight) {
    constexpr int refDpi = 96;  // Standard DPI
    constexpr int refSize = 16; // Icon size at 96 DPI
    float scaleFactor = static_cast<float>(dpi) / refDpi;
    iconWidth = static_cast<int>(std::round(refSize * scaleFactor));
    iconHeight = static_cast<int>(std::round(refSize * scaleFactor));
    iconWidth = (std::max)(16, iconWidth);
    iconHeight = (std::max)(16, iconHeight);
}

void ScaleBitmapData(const uint8_t* src, int srcWidth, int srcHeight,
    std::vector<uint8_t>& dst, int dstWidth, int dstHeight) {
    for (int y = 0; y < dstHeight; ++y) {
        for (int x = 0; x < dstWidth; ++x) {
            int srcX = static_cast<int>(std::round(static_cast<float>(x) / dstWidth * (srcWidth - 1)));
            int srcY = static_cast<int>(std::round(static_cast<float>(y) / dstHeight * (srcHeight - 1)));
            srcX = (std::min)(srcX, srcWidth - 1);
            srcY = (std::min)(srcY, srcHeight - 1);
            const uint8_t* pixel = &src[(srcY * srcWidth + srcX) * 4];
            int dstIndex = (y * dstWidth + x) * 4;

            // Swap Red & Blue channels
            dst[dstIndex] = pixel[2]; // R <== Blue
            dst[dstIndex + 1] = pixel[1]; // G stays same
            dst[dstIndex + 2] = pixel[0]; // B <== Red
            dst[dstIndex + 3] = pixel[3]; // A stays same
        }
    }
}

HBITMAP CreateBitmapFromArray(UINT dpi) {
    const uint8_t* imageData = gimp_image.pixel_data;
    int srcWidth = gimp_image.width;
    int srcHeight = gimp_image.height;
    int width, height;
    CalculateIconSize(dpi, width, height);
    size_t pixelCount = width * height;
    size_t expectedSize = pixelCount * 4;
    std::vector<uint8_t> resizedPixelData(expectedSize);
    ScaleBitmapData(imageData, srcWidth, srcHeight, resizedPixelData, width, height);
    BITMAPINFO bmi = {};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = width;
    bmi.bmiHeader.biHeight = -height;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;
    void* pBits = nullptr;
    HDC hdc = GetDC(NULL);
    HBITMAP hBitmap = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &pBits, NULL, 0);
    ReleaseDC(NULL, hdc);
    memcpy(pBits, resizedPixelData.data(), expectedSize);
    return hBitmap;
}


BOOL APIENTRY DllMain(HINSTANCE hModule, DWORD  reasonForCall, LPVOID /*lpReserved*/)
{
	try {

		switch (reasonForCall)
		{
			case DLL_PROCESS_ATTACH:
				pluginInit(hModule);
                g_inst = HINSTANCE(hModule);
				break;

			case DLL_PROCESS_DETACH:
				pluginCleanUp();
				break;

			case DLL_THREAD_ATTACH:
				break;

			case DLL_THREAD_DETACH:
				break;
		}
	}
	catch (...) { return FALSE; }

    return TRUE;
}


extern "C" __declspec(dllexport) void setInfo(NppData notpadPlusData)
{
	nppData = notpadPlusData;
    MultiReplace::loadLanguageGlobal();  // loading language Information for Plugin
    MultiReplace::loadConfigOnce();   // Load INI into cache once at plugin start
	commandMenuInit();
}

extern "C" __declspec(dllexport) const TCHAR * getName()
{
	return NPP_PLUGIN_NAME;
}

extern "C" __declspec(dllexport) FuncItem * getFuncsArray(int *nbF)
{
	*nbF = nbFunc;
	return funcItem;
}


extern "C" __declspec(dllexport) void beNotified(SCNotification * notifyCode)
{
    // forward notifications for deferred jump handling
    ResultDock::instance().onNppNotification(notifyCode);

    switch (notifyCode->nmhdr.code)
    {
    case NPPN_TBMODIFICATION:
    {
        toolbarIconsWithDarkMode tbIcons;

        // --- Start of the optimized DPI detection block-- -
        static bool s_isInitialized = false;
        static decltype(&GetDpiForWindow) s_pfnGetDpiForWindow = nullptr;

        if (!s_isInitialized)
        {
            // One-time check for the modern DPI API on first run.
            HMODULE hUser32 = ::GetModuleHandle(TEXT("User32.dll"));
            if (hUser32) {
                s_pfnGetDpiForWindow = reinterpret_cast<decltype(s_pfnGetDpiForWindow)>(
                    ::GetProcAddress(hUser32, "GetDpiForWindow"));
            }
            s_isInitialized = true;
        }

        UINT dpi = 96; // Default DPI

        if (s_pfnGetDpiForWindow) {
            // Use the modern API if available.
            dpi = s_pfnGetDpiForWindow(nppData._nppHandle);
        }
        else {
            // Fallback to the legacy method on older systems.
            HDC hdc = ::GetDC(nppData._nppHandle);
            if (hdc) {
                dpi = ::GetDeviceCaps(hdc, LOGPIXELSX);
                ::ReleaseDC(nppData._nppHandle, hdc);
            }
        }
        // --- End of the optimized DPI detection block-- -

        // Generate the bitmap with proper DPI scaling
        tbIcons.hToolbarBmp = CreateBitmapFromArray(dpi);

        // Load icons for normal and dark mode
        tbIcons.hToolbarIcon = ::LoadIcon(g_inst, MAKEINTRESOURCE(IDI_MR_ICON));
        tbIcons.hToolbarIconDarkMode = ::LoadIcon(g_inst, MAKEINTRESOURCE(IDI_MR_DM_ICON));

        // Send updated toolbar icons to Notepad++
        ::SendMessage(nppData._nppHandle, NPPM_ADDTOOLBARICON_FORDARKMODE, funcItem[0]._cmdID, (LPARAM)&tbIcons);
    }
    break;


    case NPPN_SHUTDOWN:
    {
        MultiReplace::signalShutdown();
        commandMenuCleanUp();
    }
    break;
    case SCN_UPDATEUI:
    {
        if (notifyCode->updated & SC_UPDATE_SELECTION)
        {
            MultiReplace::onSelectionChanged();
        }
        MultiReplace::onCaretPositionChanged();
    }
    break;
    case SCN_MODIFIED:
    {
        if (notifyCode->modificationType & (SC_MOD_INSERTTEXT | SC_MOD_DELETETEXT)) {
            MultiReplace::onTextChanged();
            MultiReplace::processTextChange(notifyCode);
            MultiReplace::processLog();
        }
    }
    break;

    case NPPN_BUFFERACTIVATED:
    {
        MultiReplace::onDocumentSwitched();
        MultiReplace::pointerToScintilla();
    }
    break;

    case NPPN_DARKMODECHANGED:
    {
        MultiReplace::onThemeChanged();
        ResultDock::instance().onThemeChanged();
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfHandleChange), reinterpret_cast<LPARAM>(_MultiReplace.getHSelf()));
        ::SetWindowPos(_MultiReplace.getHSelf(), nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED); // to redraw titlebar and window
        break;
    }
    case NPPN_NATIVELANGCHANGED:
    {
        MultiReplace::refreshUILanguage();
        _MultiReplaceConfig.refreshUILanguage();
        break;
    }
    default:
        return;
    }
}


// Here you can process the Npp Messages 
// I will make the messages accessible little by little, according to the need of plugin development.
// Please let me know if you need to access to some messages :
// http://sourceforge.net/forum/forum.php?forum_id=482781
//
extern "C" __declspec(dllexport) LRESULT messageProc(UINT /*Message*/, WPARAM /*wParam*/, LPARAM /*lParam*/)
{/*
	if (Message == WM_MOVE)
	{
		::MessageBox(NULL, "move", "", MB_OK);
	}
*/
	return TRUE;
}

#ifdef UNICODE
extern "C" __declspec(dllexport) BOOL isUnicode()
{
    return TRUE;
}
#endif //UNICODE
