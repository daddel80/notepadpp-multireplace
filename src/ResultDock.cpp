
#include "PluginDefinition.h"
#include <windows.h>
#include "Scintilla.h"
#include "StaticDialog/resource.h"
#include "StaticDialog/Docking.h"
#include "StaticDialog/DockingDlgInterface.h"
#include "ResultDock.h"

// Make the global NppData object available to this file.
extern NppData nppData;

// --- Singleton Accessor ---
ResultDock& ResultDock::instance()
{
    // This assumes a global g_inst is defined and set in DllMain.
    extern HINSTANCE g_inst;
    static ResultDock s{ g_inst };
    return s;
}

// --- Main public function ---
void ResultDock::ensureCreatedAndVisible(const NppData& npp)
{
    // 1) first-time creation
    if (!_hSci)                // _hSci wird im _create() gesetzt
        _create(npp);

    // 2) show again – MUST use the client handle!
    if (_hSci)
        ::SendMessage(npp._nppHandle, NPPM_DMMSHOW, 0,
            reinterpret_cast<LPARAM>(_hSci));
}

// --- Subclass Procedure ---
// This static function intercepts messages for our Scintilla control.
LRESULT CALLBACK ResultDock::sciSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == WM_NOTIFY)
    {
        auto* pNM = reinterpret_cast<NMHDR*>(lp);
        // Check if it's the specific notification for closing a dock.
        if (pNM && pNM->code == DMN_CLOSE)
        {
            // Tell Notepad++ to hide the container of this window.
            // Passing the client handle (hwnd) works to identify the dock.
            ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, (LPARAM)hwnd);
            return TRUE; // Message handled.
        }
    }

    // For all other messages, pass them on to the original Scintilla procedure.
    return ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp);
}

void ResultDock::onThemeChanged()
{
    _applyTheme();
}

static constexpr uint32_t argb(BYTE a, COLORREF c)
{
    return (uint32_t(a) << 24) |
        (uint32_t(GetRValue(c)) << 16) |
        (uint32_t(GetGValue(c)) << 8) |
        uint32_t(GetBValue(c));
}

void ResultDock::_applyTheme()
{
    if (!_hSci) return;

    const bool dark = ::SendMessage(nppData._nppHandle,
        NPPM_ISDARKMODEENABLED, 0, 0);

    COLORREF bg = (COLORREF)::SendMessage(
        nppData._nppHandle, NPPM_GETEDITORDEFAULTBACKGROUNDCOLOR, 0, 0);
    COLORREF fg = (COLORREF)::SendMessage(
        nppData._nppHandle, NPPM_GETEDITORDEFAULTFOREGROUNDCOLOR, 0, 0);

    COLORREF lnBg = dark ? RGB(0, 0, 0) : RGB(255, 255, 255);
    COLORREF lnFg = dark ? RGB(200, 200, 200) : RGB(80, 80, 80);

    COLORREF selBg = dark ? RGB(96, 96, 96) : ::GetSysColor(COLOR_HIGHLIGHT);
    COLORREF selFg = dark ? RGB(255, 255, 255) : ::GetSysColor(COLOR_HIGHLIGHTTEXT);

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { ::SendMessage(_hSci, m, w, l); };

    /* base */
    S(SCI_STYLESETBACK, STYLE_DEFAULT, bg);
    S(SCI_STYLESETFORE, STYLE_DEFAULT, fg);
    S(SCI_STYLECLEARALL);

    /* margins */
    S(SCI_SETMARGINBACKN, 0, lnBg);
    S(SCI_STYLESETBACK, STYLE_LINENUMBER, lnBg);
    S(SCI_STYLESETFORE, STYLE_LINENUMBER, lnFg);

    S(SCI_SETMARGINBACKN, 1, bg);
    S(SCI_SETMARGINBACKN, 2, bg);
    S(SCI_SETFOLDMARGINCOLOUR, 1, bg);
    S(SCI_SETFOLDMARGINHICOLOUR, 1, bg);

    /* active selection */
    S(SCI_SETSELFORE, 1, selFg);
    S(SCI_SETSELBACK, 1, selBg);
    S(SCI_SETSELALPHA, 256);

    /* inactive selection (keep same colour) */
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_BACK,
        argb(0xFF, selBg));
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_TEXT,
        argb(0xFF, selFg));

    /* additional selection (multi-sel) */
    S(SCI_SETADDITIONALSELFORE, selFg);
    S(SCI_SETADDITIONALSELBACK, selBg);
    S(SCI_SETADDITIONALSELALPHA, 256);
}
// --- Public Methods ---


void ResultDock::setText(const std::wstring& wText)
{
    if (!_hSci) return;

    // Helper to convert UTF-16 wstring to UTF-8 string for Scintilla.
    auto wstringToUtf8 = [](const std::wstring& w) -> std::string {
        if (w.empty()) return {};
        int len = ::WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
        std::string out(len, '\0');
        ::WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), out.data(), len, nullptr, nullptr);
        return out;
        };

    std::string utf8 = wstringToUtf8(wText);
    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);
    ::SendMessage(_hSci, SCI_CLEARALL, 0, 0);
    ::SendMessage(_hSci, SCI_ADDTEXT, (WPARAM)utf8.size(), (LPARAM)utf8.c_str());
    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);

    // Call the new function to build the fold map for the new text.
    _rebuildFolding();
}


// --- Private Methods ---

void ResultDock::_create(const NppData& npp)
{
    // 1) Create the Scintilla control.
    _hSci = ::CreateWindowExW(0, L"Scintilla", L"", WS_CHILD,
        0, 0, 100, 100,
        npp._nppHandle, nullptr, _hInst, nullptr);

    if (!_hSci)
    {
        MessageBoxW(npp._nppHandle, L"FATAL: CreateWindowExW for Scintilla failed!", L"ResultDock Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Subclass the Scintilla window to intercept its messages.
    s_prevSciProc = reinterpret_cast<WNDPROC>(
        ::SetWindowLongPtr(_hSci, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(sciSubclassProc))
        );

    ::SendMessage(_hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0);

    // 2) Prepare the docking descriptor.
    static tTbData dock{};
    dock.hClient = _hSci;
    dock.pszName = L"MultiReplace – Search results";
    dock.dlgID = IDD_MULTIREPLACE_RESULT_DOCK;
    dock.uMask = DWS_DF_CONT_BOTTOM | DWS_ICONTAB;
    dock.hIconTab = nullptr;
    dock.pszAddInfo = L"";
    dock.pszModuleName = NPP_PLUGIN_NAME;

    dock.iPrevCont = -1;          // *** critical ***
    dock.rcFloat = { 0,0,0,0 };   // optional

    // Register the dock and capture the container handle.
    _hDock = (HWND)::SendMessage(npp._nppHandle, NPPM_DMMREGASDCKDLG, 0, reinterpret_cast<LPARAM>(&dock));

    if (!_hDock)
    {
        MessageBoxW(npp._nppHandle, L"ERROR: NPPM_DMMREGASDCKDLG failed. Docking was rejected by Notepad++.", L"ResultDock Error", MB_OK | MB_ICONERROR);
    }

    // Ask Notepad++ to theme the newly created dock container and the Scintilla control
    ::SendMessage(npp._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME,
        static_cast<WPARAM>(NppDarkMode::dmfInit),
        reinterpret_cast<LPARAM>(_hDock));

    // Ask Notepad++ to theme the newly created dock container
    ::SendMessage(npp._nppHandle,
        NPPM_DARKMODESUBCLASSANDTHEME,
        static_cast<WPARAM>(NppDarkMode::dmfInit),
        reinterpret_cast<LPARAM>(_hDock));

    _initFolding();

    // Set the initial styles to match the current N++ theme upon creation.
    _applyTheme();
}

void ResultDock::_initFolding() const
{
    if (!_hSci) return;
    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0) { ::SendMessage(_hSci, m, w, l); };

    constexpr int MARGIN_FOLD = 2;
    S(SCI_SETMARGINTYPEN, MARGIN_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN, MARGIN_FOLD, SC_MASK_FOLDERS);
    S(SCI_SETMARGINWIDTHN, MARGIN_FOLD, 16);

    // Define the markers for folding (+, -, etc.)
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_EMPTY);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER);

    S(SCI_SETFOLDFLAGS,
        SC_FOLDFLAG_LINEAFTER_CONTRACTED |
        SC_FOLDFLAG_LINEBEFORE_CONTRACTED |
        SC_FOLDFLAG_LINEBEFORE_EXPANDED |
        SC_FOLDFLAG_LINEAFTER_EXPANDED);

    // Enable folding in the control
    S(SCI_SETPROPERTY, (sptr_t)"fold", (sptr_t)"1");
    S(SCI_SETPROPERTY, (sptr_t)"fold.compact", (sptr_t)"1");
}