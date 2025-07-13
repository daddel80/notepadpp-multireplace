#pragma once
// ResultDock – manages the dockable “Search results” pane.
// Implemented as a singleton, accessed via ResultDock::instance().

#include "PluginInterface.h"
#include <string>
#include <windows.h>

class ResultDock
{
public:
    // Singleton accessor
    static ResultDock& instance();

    // --- Public Interface ---

    // This is the only function needed to interact with the dock.
    // It creates the window on the first call, and ensures it's visible on every call.
    void ensureCreatedAndVisible(const NppData& npp);

    // Overwrites the entire buffer with new text.
    void setText(const std::wstring& wText);

    // propagate N++ dark-mode change
    void onThemeChanged();

    void _applyTheme();

    // --- Helpers & Callbacks ---


private:
    // The subclass procedure is still needed to handle the close button.
    static LRESULT CALLBACK sciSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
    static inline WNDPROC s_prevSciProc = nullptr;

    // Construction is private for the singleton pattern.
    explicit ResultDock(HINSTANCE hInst) : _hInst(hInst) {}
    ResultDock(const ResultDock&) = delete;
    ResultDock& operator=(const ResultDock&) = delete;

    // Internal implementation functions
    void _create(const NppData& npp);
    void _initFolding() const;
    void _rebuildFolding() const;

    // Member variables
    HINSTANCE _hInst = nullptr;   // DLL instance handle for resources.
    HWND      _hSci = nullptr;    // Handle to the Scintilla control inside the dock.
    HWND      _hDock = nullptr;   // Handle to the dockable container window returned by Notepad++.
};

