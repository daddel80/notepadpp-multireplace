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

// Formula Debug Window: modal window that shows per-match variable
// and capture snapshots during Replace, for both the Lua and ExprTk
// engines. Split out of MultiReplacePanel.cpp; these remain
// MultiReplace members.

#define NOMINMAX
#include "MultiReplacePanel.h"
#include "menuCmdID.h"
#include <windows.h>
#include <commctrl.h>
#include <string>

// Local singleton alias, matching MultiReplacePanel.cpp, so the
// moved function bodies stay byte-for-byte identical.
static LanguageManager& LM = LanguageManager::instance();

int MultiReplace::ShowDebugWindow(const std::string& message) {
    debugWindowResponse = -1;

    // Convert the message from UTF-8 to UTF-16 (dynamic size)
    std::wstring wMessage = Encoding::utf8ToWString(message);
    if (wMessage.empty() && !message.empty()) {
        MessageBox(nppData._nppHandle, L"Error converting message", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    static bool isClassRegistered = false;

    if (!isClassRegistered) {
        WNDCLASS wc = {};

        wc.lpfnWndProc = DebugWindowProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = L"DebugWindowClass";
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

        if (!RegisterClass(&wc)) {
            MessageBox(nppData._nppHandle, L"Error registering class", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
            return -1;
        }

        isClassRegistered = true;
    }

    // Format the message for ListView
    std::wstringstream formattedMessage;
    std::wstringstream inputMessageStream(wMessage);
    std::wstring line;

    while (std::getline(inputMessageStream, line)) {
        if (line.empty()) continue;

        std::wistringstream iss(line);
        std::wstring variable, type, value;

        if (!std::getline(iss, variable, L'\t')) continue;
        if (!std::getline(iss, type, L'\t')) continue;
        std::getline(iss, value);  // Value may be empty

        if (type == L"Number") {
            try {
                double num = std::stod(value);
                if (num == std::floor(num)) {
                    value = std::to_wstring(static_cast<long long>(num));
                }
                else {
                    std::wstringstream numStream;
                    numStream << std::fixed << std::setprecision(8) << num;
                    std::wstring formatted = numStream.str();
                    size_t pos = formatted.find_last_not_of(L'0');
                    if (pos != std::wstring::npos) {
                        formatted.erase(pos + 1);
                    }
                    if (!formatted.empty() && formatted.back() == L'.') {
                        formatted.pop_back();
                    }
                    value = formatted;
                }
            }
            catch (...) {
                // If conversion fails, retain the original value
            }
        }
        else if (type == L"String" && value.empty()) {
            value = L"<empty>";
        }

        formattedMessage << variable << L"\t" << type << L"\t" << value << L"\n";
    }

    std::wstring finalMessage = formattedMessage.str();

    std::wstring windowTitle = LM.get(L"debug_title");

    // Check if window already exists - if so, just update content (PERSISTENT WINDOW)
    if (IsWindow(hDebugWnd) && hDebugListView != nullptr) {
        // Update window title
        SetWindowTextW(hDebugWnd, windowTitle.c_str());

        // Clear existing ListView items
        ListView_DeleteAllItems(hDebugListView);

        // Repopulate ListView with new data
        std::wstringstream ss(finalMessage);
        std::wstring dataLine;
        int itemIndex = 0;
        while (std::getline(ss, dataLine)) {
            if (dataLine.empty()) continue;

            std::wstring variable, type, value;
            std::wistringstream iss(dataLine);

            if (!std::getline(iss, variable, L'\t')) continue;
            if (!std::getline(iss, type, L'\t')) continue;
            std::getline(iss, value);

            LVITEM lvItem = {};
            lvItem.mask = LVIF_TEXT;
            lvItem.iItem = itemIndex;
            lvItem.iSubItem = 0;
            lvItem.pszText = const_cast<LPWSTR>(variable.c_str());
            ListView_InsertItem(hDebugListView, &lvItem);
            ListView_SetItemText(hDebugListView, itemIndex, 1, const_cast<LPWSTR>(type.c_str()));
            ListView_SetItemText(hDebugListView, itemIndex, 2, const_cast<LPWSTR>(value.c_str()));
            ++itemIndex;
        }

        // Wait for user response (window stays open)
        debugWindowResponse = -1;
        MSG msg = {};
        while (IsWindow(hDebugWnd) && debugWindowResponse == -1)
        {
            if (_isShuttingDown) {
                DestroyWindow(hDebugWnd);
                debugWindowResponse = 3;
                hDebugWnd = nullptr;
                hDebugListView = nullptr;
                continue;
            }

            if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == WM_QUIT) break;

                // Keyboard shortcuts when the debug window itself is the
                // foreground - intercepted before IsDialogMessage so they
                // fire regardless of which child control has focus. Plain
                // Enter is left to IsDialogMessage so the default-button
                // (Next) activation keeps working.
                if (msg.message == WM_KEYDOWN
                    && msg.hwnd != nullptr
                    && GetForegroundWindow() == hDebugWnd) {
                    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                    if (msg.wParam == VK_ESCAPE) {
                        SendMessage(hDebugWnd, WM_COMMAND, MAKEWPARAM(3, 0), 0); // Stop
                        continue;
                    }
                    if (ctrl && msg.wParam == 'C') {
                        SendMessage(hDebugWnd, WM_COMMAND, MAKEWPARAM(4, 0), 0); // Copy
                        continue;
                    }
                    if (ctrl && msg.wParam == VK_RETURN) {
                        SendMessage(hDebugWnd, WM_COMMAND, MAKEWPARAM(5, 0), 0); // Run to End
                        continue;
                    }
                }

                if (!IsDialogMessage(hDebugWnd, &msg)) {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
            else {
                WaitMessage();
            }
        }

        // If Stop was pressed or window closed, cleanup
        if (debugWindowResponse != 2) {
            hDebugWnd = nullptr;
            hDebugListView = nullptr;
        }

        return debugWindowResponse;
    }

    // Window doesn't exist - create it. Compute the default size
    // exactly from the same layout constants the button cluster uses,
    // then convert client-area to outer-window size via the same
    // mechanism WM_GETMINMAXINFO uses, so the default opens with the
    // toolbar fitting tightly and no wasted space to the right of
    // Stop. The user's saved size (held in memory for the current
    // plugin session only - not persisted to the INI) takes precedence.
    int width;
    int height;
    if (debugWindowSizeSet) {
        width = debugWindowSize.cx;
        height = debugWindowSize.cy;
    }
    else {
        // Mirror the WM_GETMINMAXINFO computation: client = 4 buttons
        // + 2 inner gaps + 1 group gap + 2 side margins.
        int clientW = 4 * 120 + 2 * 10 + 25 + 2 * 10;   // = 545
        int clientH = 400;                              // generous default
        if (dpiMgr) {
            clientW = dpiMgr->scaleX(clientW);
            clientH = dpiMgr->scaleY(clientH);
        }
        RECT r = { 0, 0, clientW, clientH };
        const DWORD style = WS_OVERLAPPEDWINDOW;
        const DWORD exStyle = WS_EX_TOPMOST;
        if (AdjustWindowRectEx(&r, style, FALSE, exStyle)) {
            width = r.right - r.left;
            height = r.bottom - r.top;
        }
        else {
            // Conservative fallback if the call fails - generous
            // enough to fit any reasonable Windows frame.
            width = clientW + 20;
            height = clientH + 50;
        }
    }
    int x = debugWindowPositionSet ? debugWindowPosition.x : (GetSystemMetrics(SM_CXSCREEN) - width) / 2;
    int y = debugWindowPositionSet ? debugWindowPosition.y : (GetSystemMetrics(SM_CYSCREEN) - height) / 2;

    HWND hwnd = CreateWindowEx(
        WS_EX_TOPMOST,
        L"DebugWindowClass",
        windowTitle.c_str(),
        WS_OVERLAPPEDWINDOW,
        x, y, width, height,
        nppData._nppHandle, nullptr, hInstance, reinterpret_cast<LPVOID>(const_cast<wchar_t*>(finalMessage.c_str()))
    );

    if (hwnd == nullptr) {
        MessageBoxW(nppData._nppHandle, L"Error creating window", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    // Activate Dark Mode
    ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(hwnd));

    // Save the handle of the debug window
    hDebugWnd = hwnd;

    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg = {};

    // Message loop: wait until a response is set (not until the window closes)
    while (IsWindow(hwnd) && debugWindowResponse == -1)
    {
        if (_isShuttingDown) {
            DestroyWindow(hwnd);
            debugWindowResponse = 3;
            hDebugWnd = nullptr;
            hDebugListView = nullptr;
            continue;
        }

        if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT) {
                break;
            }

            // Keyboard shortcuts when the debug window itself is the
            // foreground - intercepted before IsDialogMessage so they
            // fire regardless of which child control has focus. Plain
            // Enter is left to IsDialogMessage so the default-button
            // (Next) activation keeps working.
            if (msg.message == WM_KEYDOWN
                && msg.hwnd != nullptr
                && GetForegroundWindow() == hwnd) {
                const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                if (msg.wParam == VK_ESCAPE) {
                    SendMessage(hwnd, WM_COMMAND, MAKEWPARAM(3, 0), 0); // Stop
                    continue;
                }
                if (ctrl && msg.wParam == 'C') {
                    SendMessage(hwnd, WM_COMMAND, MAKEWPARAM(4, 0), 0); // Copy
                    continue;
                }
                if (ctrl && msg.wParam == VK_RETURN) {
                    SendMessage(hwnd, WM_COMMAND, MAKEWPARAM(5, 0), 0); // Run to End
                    continue;
                }
            }

            if (!IsDialogMessage(hwnd, &msg)) {
                if (GetForegroundWindow() != hwnd &&
                    msg.message == WM_KEYDOWN &&
                    (GetKeyState(VK_CONTROL) & 0x8000))
                {
                    bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                    bool handled = true;

                    switch (msg.wParam) {
                    case 'C': SendMessage(nppData._scintillaMainHandle, SCI_COPY, 0, 0); break;
                    case 'V': SendMessage(nppData._scintillaMainHandle, SCI_PASTE, 0, 0); break;
                    case 'X': SendMessage(nppData._scintillaMainHandle, SCI_CUT, 0, 0); break;
                    case 'U': SendMessage(nppData._scintillaMainHandle,
                        shiftPressed ? SCI_UPPERCASE : SCI_LOWERCASE, 0, 0); break;
                    case 'S': SendMessage(nppData._nppHandle,
                        shiftPressed ? NPPM_SAVEALLFILES : NPPM_SAVECURRENTFILE, 0, 0); break;
                    case 'G': SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_SEARCH_GOTOLINE); break;
                    case 'F': SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_SEARCH_FIND); break;
                    default: handled = false; break;
                    }

                    if (handled) continue;
                }

                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
        else
        {
            WaitMessage();
        }
    }

    // If Stop was pressed or window closed (not Next), cleanup
    if (debugWindowResponse != 2) {
        hDebugWnd = nullptr;
        hDebugListView = nullptr;
    }

    if (debugWindowResponse == 2)
    {
        MSG m;
        while (PeekMessage(&m, nullptr, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE)) {}
    }

    return debugWindowResponse;
}

LRESULT CALLBACK MultiReplace::DebugWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hListView;
    static HWND hNextButton;
    static HWND hStopButton;
    static HWND hCopyButton;
    static HWND hRunToEndButton;

    switch (msg) {
    case WM_CREATE: {
        hListView = CreateWindowW(WC_LISTVIEW, L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            10, 10, 360, 140,
            hwnd, reinterpret_cast<HMENU>(1), nullptr, nullptr);
        hDebugListView = hListView;  // Save handle for content updates

        // Initialize columns
        LVCOLUMN lvCol = {};
        lvCol.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        lvCol.pszText = LM.getW(L"debug_col_variable");
        lvCol.cx = 120;
        ListView_InsertColumn(hListView, 0, &lvCol);
        lvCol.pszText = LM.getW(L"debug_col_type");
        lvCol.cx = 80;
        ListView_InsertColumn(hListView, 1, &lvCol);
        lvCol.pszText = LM.getW(L"debug_col_value");
        lvCol.cx = 180;
        ListView_InsertColumn(hListView, 2, &lvCol);

        // Button layout: a single fixed-spacing group anchored to the
        // left edge. Order is Next, Copy, Run to End, Stop - Next sits
        // closest to the variable list (left side of the window) so the
        // mouse stays near the data the user is reading, and Stop sits
        // at the far right of the cluster (out of the primary click
        // path). A wider gap between Copy and Run to End separates the
        // inspection actions (Next/Copy) from the run-control actions
        // (Run to End/Stop) without breaking them into two clusters.
        // The group does not stretch with the window: WM_SIZE keeps the
        // inter-button gaps constant. Real positions are set in WM_SIZE;
        // values here are provisional for the first paint.
        // Layout constants in unscaled pixels - all values are scaled
        // through sx()/sy() before use so the cluster grows with the
        // user's DPI setting (consistent with the rest of the plugin).
        // Without scaling, WM_GETMINMAXINFO would compute a DPI-aware
        // minimum window width while the buttons themselves stayed at
        // 96-DPI sizes, leaving wasted space to the right of Stop on
        // high-DPI displays.
        auto sxv = [](int v) {
            return (instance && instance->dpiMgr) ? instance->sx(v) : v;
            };
        auto syv = [](int v) {
            return (instance && instance->dpiMgr) ? instance->sy(v) : v;
            };

        const int btnWidth = sxv(120);
        const int btnHeight = syv(30);
        const int btnGap = sxv(10);
        const int groupGap = sxv(25);
        const int btnY = syv(160);
        const int margin = sxv(10);

        // x positions: Next at margin; Copy directly after; Run to End
        // shifted by groupGap to widen the visual gap; Stop after Run
        // to End at the standard btnGap.
        const int xNext = margin;
        const int xCopy = xNext + (btnWidth + btnGap);
        const int xRun = xCopy + (btnWidth + groupGap);
        const int xStop = xRun + (btnWidth + btnGap);

        hNextButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_next"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            xNext, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(2), nullptr, nullptr);

        hCopyButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_copy"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            xCopy, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(4), nullptr, nullptr);

        hRunToEndButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_run_to_end"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            xRun, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(5), nullptr, nullptr);

        hStopButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_stop"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            xStop, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(3), nullptr, nullptr);

        // Extract the debug message from lParam
        LPCWSTR debugMessage = reinterpret_cast<LPCWSTR>(reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams);

        // Populate ListView with the debug message
        std::wstringstream ss(debugMessage);
        std::wstring line;
        int itemIndex = 0;
        while (std::getline(ss, line)) {
            if (line.empty()) continue;

            std::wstring variable, type, value;
            std::wistringstream iss(line);

            if (!std::getline(iss, variable, L'\t')) continue;
            if (!std::getline(iss, type, L'\t')) continue;
            std::getline(iss, value);

            LVITEM lvItem = {};
            lvItem.mask = LVIF_TEXT;
            lvItem.iItem = itemIndex;
            lvItem.iSubItem = 0;
            lvItem.pszText = const_cast<LPWSTR>(variable.c_str());
            ListView_InsertItem(hListView, &lvItem);
            ListView_SetItemText(hListView, itemIndex, 1, const_cast<LPWSTR>(type.c_str()));
            ListView_SetItemText(hListView, itemIndex, 2, const_cast<LPWSTR>(value.c_str()));
            ++itemIndex;
        }

        // Initial focus on Next so Enter/Space activates it immediately
        // and the user does not have to Tab through the buttons first.
        SetFocus(hNextButton);

        break;
    }

    case WM_SIZE: {
        RECT rect;
        GetClientRect(hwnd, &rect);

        // Layout constants in unscaled pixels - all values are scaled
        // through sx()/sy() before use so the cluster stays consistent
        // with the rest of the plugin's DPI handling.
        auto sxv = [](int v) {
            return (instance && instance->dpiMgr) ? instance->sx(v) : v;
            };
        auto syv = [](int v) {
            return (instance && instance->dpiMgr) ? instance->sy(v) : v;
            };

        const int btnWidth = sxv(120);
        const int btnHeight = syv(30);
        const int btnGap = sxv(10);
        const int groupGap = sxv(25);
        const int margin = sxv(10);
        const int btnY = rect.bottom - btnHeight - margin;
        const int listHeight = rect.bottom - btnHeight - syv(40);

        SetWindowPos(hListView, nullptr, margin, margin,
            rect.right - 2 * margin, listHeight, SWP_NOZORDER);

        // Single fixed-spacing button cluster, left-anchored. Order:
        // Next, Copy, [groupGap], Run to End, Stop. The wider gap
        // between Copy and Run to End separates the inspection actions
        // from the run-control actions while keeping the cluster
        // visually unified. The cluster keeps its layout when the
        // window grows - it does not stretch.
        const int xNext = margin;
        const int xCopy = xNext + (btnWidth + btnGap);
        const int xRun = xCopy + (btnWidth + groupGap);
        const int xStop = xRun + (btnWidth + btnGap);

        SetWindowPos(hNextButton, nullptr,
            xNext, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hCopyButton, nullptr,
            xCopy, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hRunToEndButton, nullptr,
            xRun, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hStopButton, nullptr,
            xStop, btnY, btnWidth, btnHeight, SWP_NOZORDER);

        // ListView columns: Variable and Type get fixed widths sized
        // for their typical content (no need to grow them when the
        // window is resized), while Value takes whatever horizontal
        // space remains. This keeps long values (file paths, regex
        // matches) readable instead of clipping at the column header.
        const int colVariable = sxv(120);
        const int colType = sxv(80);
        // ListView width = client width minus the two side margins;
        // GetClientRect already reflects scrollbar presence so the
        // remaining-pixels math stays valid even if a vertical
        // scrollbar appears.
        const int listInnerWidth = (rect.right - 2 * margin);
        const int colValue = std::max(listInnerWidth - colVariable - colType,
            colVariable);  // fall back to a sensible minimum
        ListView_SetColumnWidth(hListView, 0, colVariable);
        ListView_SetColumnWidth(hListView, 1, colType);
        ListView_SetColumnWidth(hListView, 2, colValue);
        break;
    }

    case WM_GETMINMAXINFO: {
        // Enforce a minimum window size so the four-button cluster
        // (Next, Copy, [groupGap], Run to End, Stop at 120px each plus
        // 10px / 25px gaps and 10px side margins) just fits with no
        // wasted space to the right of Stop, and the ListView keeps
        // enough vertical room to be readable.
        //
        // The cluster's required client-area width is computed from
        // the same constants that WM_CREATE / WM_SIZE use. Window
        // size includes the non-client frame, so AdjustWindowRectEx
        // is used to compute the exact frame contribution from the
        // window's own style flags (works even before WM_CREATE has
        // run and is independent of theme / Windows version quirks).
        MINMAXINFO* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
        if (mmi) {
            const int btnWidth = 120;
            const int btnGap = 10;
            const int groupGap = 25;
            const int margin = 10;

            // Required client-area dimensions.
            int minClientW = 4 * btnWidth + 2 * btnGap + groupGap + 2 * margin;
            int minClientH = 160;

            if (instance && instance->dpiMgr) {
                minClientW = instance->dpiMgr->scaleX(minClientW);
                minClientH = instance->dpiMgr->scaleY(minClientH);
            }

            // Translate the client-area minimum into a window-area
            // minimum by adding the non-client frame thickness.
            const DWORD style = static_cast<DWORD>(GetWindowLong(hwnd, GWL_STYLE));
            const DWORD exStyle = static_cast<DWORD>(GetWindowLong(hwnd, GWL_EXSTYLE));
            RECT r = { 0, 0, minClientW, minClientH };
            if (AdjustWindowRectEx(&r, style, FALSE, exStyle)) {
                mmi->ptMinTrackSize.x = r.right - r.left;
                mmi->ptMinTrackSize.y = r.bottom - r.top;
            }
            else {
                // Conservative fallback if the call fails - generous
                // enough to fit any reasonable Windows frame.
                mmi->ptMinTrackSize.x = minClientW + 20;
                mmi->ptMinTrackSize.y = minClientH + 50;
            }
        }
        return 0;
    }

    case WM_DPICHANGED: {
        if (instance && instance->dpiMgr) {
            instance->dpiMgr->updateDPI(hwnd);
        }
        RECT* pRect = reinterpret_cast<RECT*>(lParam);
        if (pRect) {
            SetWindowPos(hwnd, nullptr, pRect->left, pRect->top,
                pRect->right - pRect->left, pRect->bottom - pRect->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return 0;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == 2) {  // Next button
            debugWindowResponse = 2;
            // Don't destroy window - just set response and let message loop exit
        }
        else if (LOWORD(wParam) == 3) {  // Stop/Close button
            debugWindowResponse = 3;

            // Save the window position and size before closing
            RECT rect;
            if (GetWindowRect(hwnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }

            DestroyWindow(hwnd);
            hDebugListView = nullptr;
        }
        else if (LOWORD(wParam) == 4) {  // Copy button was pressed
            CopyListViewToClipboard(hListView);
        }
        else if (LOWORD(wParam) == 5) {  // Run to End: silence dialog for the rest of this run
            // Flip the per-run skip flag on the panel and tear the window
            // down. The next debug-write call will see isDebugModeEnabled()
            // return false and proceed without showing the dialog. The
            // flag is cleared at the start of the next run, so the
            // persistent Debug-Mode toggle is unaffected.
            if (instance) {
                instance->_debugSkipForRun = true;
            }
            debugWindowResponse = 2;  // continue (same semantics as Next)

            RECT rect;
            if (GetWindowRect(hwnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }

            DestroyWindow(hwnd);
            hDebugListView = nullptr;
        }
        break;

    case WM_CLOSE:
        // Only set debugWindowResponse to -1 if not already set by a button
        if (debugWindowResponse == -1) {
            // Closed by X, Alt+F4, or CloseDebugWindow()
            debugWindowResponse = -1;
        }

        // Save the window position and size before closing
        RECT rect;
        if (GetWindowRect(hwnd, &rect)) {
            debugWindowPosition.x = rect.left;
            debugWindowPosition.y = rect.top;
            debugWindowPositionSet = true;
            debugWindowSize.cx = rect.right - rect.left;
            debugWindowSize.cy = rect.bottom - rect.top;
            debugWindowSizeSet = true;
        }

        DestroyWindow(hwnd);
        break;

    case WM_DESTROY:
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void MultiReplace::CopyListViewToClipboard(HWND hListView) {
    const int itemCount = ListView_GetItemCount(hListView);
    if (itemCount <= 0) {
        return;
    }

    const int columnCount = Header_GetItemCount(ListView_GetHeader(hListView));
    std::wstring clipboardText;
    clipboardText.reserve(static_cast<size_t>(itemCount) * columnCount * 64);

    wchar_t buffer[512];

    for (int i = 0; i < itemCount; ++i) {
        for (int j = 0; j < columnCount; ++j) {
            LVITEMW li{};
            li.iSubItem = j;
            li.cchTextMax = std::size(buffer);
            li.pszText = buffer;

            buffer[0] = L'\0';
            SendMessageW(hListView, LVM_GETITEMTEXTW, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(&li));
            buffer[std::size(buffer) - 1] = L'\0';

            int len = lstrlenW(buffer);
            if (len > 0) {
                clipboardText.append(buffer, len);
            }

            if (j < columnCount - 1) {
                clipboardText.push_back(L'\t');
            }
        }
        clipboardText.push_back(L'\n');
    }

    if (clipboardText.empty()) {
        return;
    }

    if (OpenClipboard(nullptr)) {
        EmptyClipboard();
        size_t bytes = (clipboardText.size() + 1) * sizeof(wchar_t);
        HGLOBAL hData = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (hData) {
            void* ptr = GlobalLock(hData);
            if (ptr) {
                memcpy(ptr, clipboardText.c_str(), bytes);
                GlobalUnlock(hData);
                SetClipboardData(CF_UNICODETEXT, hData);
            }
            else {
                GlobalFree(hData);
            }
        }
        CloseClipboard();
    }
}

void MultiReplace::CloseDebugWindow()
{
    // Triggers the WM_CLOSE message for the debug window, handled in DebugWindowProc
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // Save the window position and size before closing
    RECT rc;
    if (GetWindowRect(hDebugWnd, &rc)) {
        debugWindowPosition = { rc.left, rc.top };
        debugWindowSize = { rc.right - rc.left, rc.bottom - rc.top };
        debugWindowPositionSet = debugWindowSizeSet = true;
    }

    SendMessage(hDebugWnd, WM_CLOSE, 0, 0);  // synchronous -> window really gone
}

void MultiReplace::SetDebugComplete()
{
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // Update window title to show completion
    SetWindowTextW(hDebugWnd, LM.getW(L"debug_title_complete"));

    // Change Stop button to Close
    HWND hStopButton = GetDlgItem(hDebugWnd, 3);
    if (hStopButton) {
        SetWindowTextW(hStopButton, LM.getW(L"debug_btn_close"));
    }

    // Disable Next button
    HWND hNextButton = GetDlgItem(hDebugWnd, 2);
    if (hNextButton) {
        EnableWindow(hNextButton, FALSE);
    }

    // Disable Run-to-End - the run is already over, nothing to skip to.
    HWND hRunToEnd = GetDlgItem(hDebugWnd, 5);
    if (hRunToEnd) {
        EnableWindow(hRunToEnd, FALSE);
    }

    // The default-focus target (Next) is now disabled. Hand focus to
    // the Close button so Enter / Space close the window without the
    // user having to Tab past the disabled controls.
    if (hStopButton) {
        SetFocus(hStopButton);
    }
}

void MultiReplace::WaitForDebugWindowClose(bool autoClose)
{
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // For single actions: close window immediately without waiting
    if (autoClose) {
        CloseDebugWindow();
        return;
    }

    SetDebugComplete();

    // Wait for user to close the window
    MSG msg = {};
    while (IsWindow(hDebugWnd))
    {
        if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) break;
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        else {
            WaitMessage();
        }
    }
    hDebugWnd = nullptr;
    hDebugListView = nullptr;
}