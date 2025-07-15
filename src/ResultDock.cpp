
#include "PluginDefinition.h"
#include <windows.h>
#include "Scintilla.h"
#include "StaticDialog/resource.h"
#include "StaticDialog/Docking.h"
#include "StaticDialog/DockingDlgInterface.h"
#include "ResultDock.h"
#include <algorithm>
#include <string>
#include <commctrl.h>


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
    switch (msg)
    {
    case WM_LBUTTONDBLCLK:
    {
        // 0) Save the current scroll position
        LRESULT firstVisibleLine = ::SendMessage(
            hwnd, SCI_GETFIRSTVISIBLELINE, 0, 0);

        // 1) Hit-test to find which buffer line was clicked
        int x = static_cast<short>(LOWORD(lp));
        int y = static_cast<short>(HIWORD(lp));
        Sci_Position pos = static_cast<Sci_Position>(
            ::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y));
        int displayLine = static_cast<int>(
            ::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0));

        // 2) Convert to hit index (skip the two header lines)
        constexpr int headerLines = 2;
        int hitIndex = displayLine - headerLines;
        const auto& hits = ResultDock::instance().hits();
        if (hitIndex < 0 || hitIndex >= static_cast<int>(hits.size()))
        {
            // Click was outside of actual hits: restore scroll and consume
            ::SendMessage(hwnd,
                SCI_SETFIRSTVISIBLELINE,
                firstVisibleLine,
                0);
            return 0;
        }
        const auto& hit = hits[hitIndex];

        // 3) Jump main editor to the hit’s line and select the pattern
        {
            HWND mainEditor = nppData._scintillaMainHandle;
            ::SendMessage(mainEditor,
                SCI_GOTOLINE,
                hit.fileLine - 1,
                0);
            ::SendMessage(mainEditor,
                SCI_ENSUREVISIBLEENFORCEPOLICY,
                hit.fileLine - 1,
                0);

            // Restrict search to that line
            Sci_Position lineStart = static_cast<Sci_Position>(
                ::SendMessage(mainEditor,
                    SCI_POSITIONFROMLINE,
                    hit.fileLine - 1,
                    0));
            Sci_Position lineEnd = static_cast<Sci_Position>(
                ::SendMessage(mainEditor,
                    SCI_GETLINEENDPOSITION,
                    hit.fileLine - 1,
                    0));
            ::SendMessage(mainEditor,
                SCI_SETTARGETRANGE,
                lineStart,
                lineEnd);

            // Search for exact UTF-8 pattern
            ::SendMessage(mainEditor,
                SCI_SEARCHINTARGET,
                static_cast<WPARAM>(hit.findUtf8.length()),
                reinterpret_cast<LPARAM>(hit.findUtf8.c_str()));
            // Highlight it
            ::SendMessage(mainEditor,
                SCI_SETSEL,
                ::SendMessage(mainEditor, SCI_GETTARGETSTART, 0, 0),
                ::SendMessage(mainEditor, SCI_GETTARGETEND, 0, 0));
        }

        // --- START OF CORRECTIONS ---

        // 4) Place caret in the dock at the start of the clicked line
        Sci_Position dockLineStartPos = static_cast<Sci_Position>(
            ::SendMessage(hwnd,
                SCI_POSITIONFROMLINE,
                displayLine,
                0));
        ::SendMessage(hwnd,
            SCI_SETEMPTYSELECTION,
            dockLineStartPos,
            0);

        // 5) Restore the dock’s original scroll position
        ::SendMessage(hwnd,
            SCI_SETFIRSTVISIBLELINE,
            firstVisibleLine,
            0);

        // 6) Focus the dock so its caret-line highlight remains visible
        ::SetFocus(hwnd);

        // --- END OF CORRECTIONS ---

        // 7) Consume the double-click so Scintilla’s default handler won’t scroll
        return 0;
    }

    case WM_NCDESTROY:
    {
        // Restore the original Scintilla WNDPROC
        ::SetWindowLongPtr(hwnd,
            GWLP_WNDPROC,
            reinterpret_cast<LONG_PTR>(s_prevSciProc));
        s_prevSciProc = nullptr;
        break;
    }

    case DMN_CLOSE:
    {
        // Hide the dock when its close button is clicked
        ::SendMessage(nppData._nppHandle,
            NPPM_DMMHIDE,
            0,
            reinterpret_cast<LPARAM>(hwnd));
        break;
    }
    }

    // Forward all other messages to the previous proc
    if (s_prevSciProc)
        return ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp);

    return ::DefWindowProc(hwnd, msg, wp, lp);
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
        NPPM_ISDARKMODEENABLED,
        0, 0) != 0;

    // Editor default colours
    COLORREF bg = (COLORREF)::SendMessage(
        nppData._nppHandle,
        NPPM_GETEDITORDEFAULTBACKGROUNDCOLOR,
        0, 0);
    COLORREF fg = (COLORREF)::SendMessage(
        nppData._nppHandle,
        NPPM_GETEDITORDEFAULTFOREGROUNDCOLOR,
        0, 0);

    // Margin colours (unchanged) …
    COLORREF lnBg = dark ? RGB(0, 0, 0) : RGB(255, 255, 255);
    COLORREF lnFg = dark ? RGB(200, 200, 200) : RGB(80, 80, 80);

    // **Use your selection background** here for the caret‐line
    COLORREF selBg = dark ? RGB(96, 96, 96)
        : ::GetSysColor(COLOR_HIGHLIGHT);
    COLORREF selFg = dark ? RGB(255, 255, 255)
        : ::GetSysColor(COLOR_HIGHLIGHTTEXT);

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { ::SendMessage(_hSci, m, w, l); };

    // 2) Base styles
    S(SCI_STYLESETBACK, STYLE_DEFAULT, bg);
    S(SCI_STYLESETFORE, STYLE_DEFAULT, fg);
    S(SCI_STYLECLEARALL);

    // 3) Margins and fold margin
    S(SCI_SETMARGINBACKN, 0, lnBg);
    S(SCI_STYLESETBACK, STYLE_LINENUMBER, lnBg);
    S(SCI_STYLESETFORE, STYLE_LINENUMBER, lnFg);
    S(SCI_SETMARGINBACKN, 1, bg);

    const COLORREF foldBg = RGB(0, 0, 0);
    S(SCI_SETMARGINBACKN, 2, foldBg);
    S(SCI_SETFOLDMARGINCOLOUR, 1, foldBg);
    S(SCI_SETFOLDMARGINHICOLOUR, 1, foldBg);

    // 4) Selection colors
    S(SCI_SETSELFORE, 1, selFg);
    S(SCI_SETSELBACK, 1, selBg);
    S(SCI_SETSELALPHA, 256, 0);
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_BACK,
        argb(0xFF, selBg));
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_TEXT,
        argb(0xFF, selFg));
    S(SCI_SETADDITIONALSELFORE, selFg);
    S(SCI_SETADDITIONALSELBACK, selBg);
    S(SCI_SETADDITIONALSELALPHA, 256, 0);

    // 5) Fold‐marker colors
    const COLORREF boxFill = foldBg;
    const COLORREF lineClr = RGB(230, 230, 210);

    for (int id : { SC_MARKNUM_FOLDER, SC_MARKNUM_FOLDEROPEN })
    {
        S(SCI_MARKERSETFORE, id, lineClr);
        S(SCI_MARKERSETBACK, id, boxFill);
    }
    for (int id : { SC_MARKNUM_FOLDERSUB,
        SC_MARKNUM_FOLDERMIDTAIL,
        SC_MARKNUM_FOLDERTAIL,
        SC_MARKNUM_FOLDEREND })
    {
        S(SCI_MARKERSETFORE, id, lineClr);
        S(SCI_MARKERSETBACK, id, lineClr);
    }

    // 6) Configure indicator #0 as a full‐line highlight
    S(SCI_INDICSETSTYLE, 0, INDIC_ROUNDBOX);
    S(SCI_INDICSETFORE, 0, selBg);
    S(SCI_INDICSETUNDER, 0, true);
    S(SCI_INDICSETALPHA, 0, 128);

    // 7) Give focus so the indicator (caret) shows immediately
    S(SCI_SETCARETLINEVISIBLE, 1, 0);       // enable current‐line highlight
    S(SCI_SETCARETLINEBACK, selBg, 0);   // paint with your selection BG
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

    auto S = [this](UINT msg, WPARAM w = 0, LPARAM l = 0)
             { ::SendMessage(_hSci, msg, w, l); };

    // ── 1) Set up fold margin (#2) as symbol margin ────────────────────────
    constexpr int M_FOLD = 2;
    S(SCI_SETMARGINTYPEN,      M_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN,      M_FOLD, SC_MASK_FOLDERS);

    // ── 2) Compute a width that fits a 16px box plus its little stem ───────
    //    (vector shapes are 16px wide, but the connected versions need a bit more)
    LRESULT lr = ::SendMessage(_hSci, SCI_TEXTHEIGHT, 0, 0);
    int     h  = static_cast<int>(lr);
    int     w  = h + 4;    // give 4px extra so the stem is never clipped
    S(SCI_SETMARGINWIDTHN,   M_FOLD, w);

    S(SCI_SETMARGINSENSITIVEN, M_FOLD, TRUE);
    S(SCI_SETAUTOMATICFOLD,
        SC_AUTOMATICFOLD_SHOW  |
        SC_AUTOMATICFOLD_CLICK |
        SC_AUTOMATICFOLD_CHANGE, 0);

    // ── 3) Hide margins 0 & 1 ───────────────────────────────────────────────
    S(SCI_SETMARGINWIDTHN, 0, 0);
    S(SCI_SETMARGINWIDTHN, 1, 0);

    // ── 4) Define **connected** box shapes for closed/open headers ──────────
    //    so the vertical guide-line runs through the box edge
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER,     SC_MARK_ARROW);   // closed box with upward stem
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_ARROWDOWN);  // open   box with downward stem

    // ── 5) Define the guide-line and corner shapes ─────────────────────────
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB,     SC_MARK_VLINE);    // │
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);  // ├
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL,    SC_MARK_LCORNER);  // └
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND,     SC_MARK_LCORNER);  // └ (end)

    // ── 6) Give them a neutral placeholder colour (real colours in _applyTheme) ─
    const COLORREF placeholder = RGB(200,200,200);
    for (int id : { SC_MARKNUM_FOLDER,
                    SC_MARKNUM_FOLDEROPEN,
                    SC_MARKNUM_FOLDERSUB,
                    SC_MARKNUM_FOLDERMIDTAIL,
                    SC_MARKNUM_FOLDERTAIL,
                    SC_MARKNUM_FOLDEREND })
    {
        S(SCI_MARKERSETFORE, id, placeholder);
        S(SCI_MARKERSETBACK, id, placeholder);
    }
}


void ResultDock::_rebuildFolding() const
{
    if (!_hSci) return;

    int lines = (int)::SendMessage(_hSci, SCI_GETLINECOUNT, 0, 0);

    for (int i = 0, level; i < lines; ++i)
    {
        Sci_Position pos = ::SendMessage(_hSci, SCI_POSITIONFROMLINE, i, 0);
        int ch = (int)::SendMessage(_hSci, SCI_GETCHARAT, pos, 0);

        bool blank = (ch == '\r' || ch == '\n');
        bool header = !blank && (ch != ' ' && ch != '\t');   // indent == 0

        level = header ? (SC_FOLDLEVELBASE | SC_FOLDLEVELHEADERFLAG)
            : (SC_FOLDLEVELBASE + 1);

        if (blank) level |= SC_FOLDLEVELWHITEFLAG;

        ::SendMessage(_hSci, SCI_SETFOLDLEVEL, i, level);
    }
}

// ResultDock.cpp
void ResultDock::prependText(const std::wstring& wText)
{
    if (!_hSci) return;

    // UTF-16 → UTF-8 helper (same as the one used in setText)
    auto w2u = [](const std::wstring& ws)->std::string
    {
        if (ws.empty()) return {};
        int len = ::WideCharToMultiByte(CP_UTF8,0,ws.data(),
                                        (int)ws.size(),nullptr,0,nullptr,nullptr);
        std::string out(len, '\0');
        ::WideCharToMultiByte(CP_UTF8,0,ws.data(),(int)ws.size(),
                              out.data(),len,nullptr,nullptr);
        return out;
    };

    std::string utf8 = w2u(wText);

    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);

    // add CR/LF before existing text if buffer is not empty
    if (::SendMessage(_hSci, SCI_GETLENGTH, 0, 0) > 0)
        ::SendMessage(_hSci, SCI_INSERTTEXT, 0, (LPARAM)"\r\n");

    // insert new block at buffer start
    ::SendMessage(_hSci, SCI_INSERTTEXT, 0, (LPARAM)utf8.c_str());

    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);

    _rebuildFolding();   // rebuild fold levels incl. new header
}


void ResultDock::appendText(const std::wstring& wText)
{
    if (!_hSci) return;

    // UTF-16 → UTF-8 helper
    auto w2u = [](const std::wstring& ws) -> std::string
        {
            if (ws.empty()) return {};
            int len = ::WideCharToMultiByte(CP_UTF8, 0,
                ws.data(), static_cast<int>(ws.size()),
                nullptr, 0, nullptr, nullptr);
            std::string out(len, '\0');
            ::WideCharToMultiByte(CP_UTF8, 0,
                ws.data(), static_cast<int>(ws.size()),
                out.data(), len,
                nullptr, nullptr);
            return out;
        };

    std::string utf8 = w2u(wText);

    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);

    // ensure there is exactly ONE blank line between blocks
    Sci_Position len = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);
    if (len > 0) {
        const char* eol = "\r\n";
        ::SendMessage(_hSci, SCI_ADDTEXT, 2, (LPARAM)eol);
    }

    ::SendMessage(_hSci, SCI_ADDTEXT, (WPARAM)utf8.size(), (LPARAM)utf8.c_str());
    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);

    _rebuildFolding();          // refresh fold map incl. new header
}

void ResultDock::clearHits() {
    _hits.clear();
}

void ResultDock::recordHit(int fileLine, const std::string& findUtf8) {
    _hits.push_back({ fileLine, findUtf8 });
}
