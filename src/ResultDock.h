#pragma once
// ResultDock – manages the dockable “Search results” pane.
// Implemented as a singleton, accessed via ResultDock::instance().

#include "PluginInterface.h"
#include <string>
#include <windows.h>

class ResultDock
{
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

public:
    struct Hit {
        std::string fullPathUtf8;
        Sci_Position pos;
        Sci_Position length;
    };
    std::vector<Hit> _hits;        // one entry per visible “Line N: …” in the dock

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

    // Replace recordHit with a straight push_back if you ever need it:
    void recordHit(const std::string& fullPathUtf8, Sci_Position pos, Sci_Position length);

    // New prependHits: insert _and_ update the view in one go.
    void prependHits(const std::vector<Hit>& newHits, const std::wstring& text);

    // Clear everything (if you ever need a full reset):
    void clearAll();

    void rebuildFolding() const { _rebuildFolding(); }

    // Expose hits for your double-click handler
    const std::vector<Hit>& hits() const { return _hits; }

};

