#pragma once
// ResultDock – manages the dockable “Search results” pane.
// Implemented as a singleton, accessed via ResultDock::instance().

#include "PluginInterface.h"
#include <string>
#include <windows.h>

class ResultDock
{
private:
    struct Hit {
        int           fileLine;    // 1-based file line number
        std::string   findUtf8;    // exact UTF-8 search pattern
    };
    std::vector<Hit> _hits;        // one entry per visible “Line N: …” in the dock

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

    void prependText(const std::wstring& wText);

    void appendText(const std::wstring& wText);

    void _applyTheme();

    // Clear all recorded hits (call before running a new Find All)
    void clearHits();

    // Record one hit (fileLine is 1-based, findUtf8 is the exact pattern)
    void recordHit(int fileLine, const std::string& findUtf8);

    // Expose hits for your double-click handler
    const std::vector<Hit>& hits() const { return _hits; }

};

