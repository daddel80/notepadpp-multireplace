
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
    extern NppData nppData;
    const std::string marker = "Line ";

    switch (msg)
    {
    case WM_NOTIFY:
    {
        NMHDR* hdr = reinterpret_cast<NMHDR*>(lp);
        if (hdr->code == DMN_CLOSE)
        {
            ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(hwnd));
            return TRUE;  
        }

        SCNotification* scn = reinterpret_cast<SCNotification*>(lp);
        if (scn->nmhdr.code == SCN_MARGINCLICK && scn->margin == 2)
        {
            int line = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, scn->position, 0);
            ::SendMessage(hwnd, SCI_TOGGLEFOLD, line, 0);
            return 0;
        }
        break;
    }

    case WM_LBUTTONDBLCLK:
    {
        // 1) Remember scroll
        LRESULT firstVisible = ::SendMessage(hwnd, SCI_GETFIRSTVISIBLELINE, 0, 0);

        // 2) Which display‐line was clicked?
        int x = (short)LOWORD(lp), y = (short)HIWORD(lp);
        Sci_Position pos = ::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
        int dispLine = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0);

        // 3) Read the raw text of that line
        int lineLen = (int)::SendMessage(hwnd, SCI_LINELENGTH, dispLine, 0);
        std::string raw(lineLen, '\0');
        ::SendMessage(hwnd, SCI_GETLINE, dispLine, (LPARAM)&raw[0]);
        raw.resize(strnlen(raw.c_str(), lineLen));

        // 4) Bail if this line does *not* contain "Line "
        if (raw.find(marker) == std::string::npos)
            return 0;

        // 5) Count *all* previous "Line " occurrences to get our hitIndex
        int hitIndex = -1;
        for (int i = 0; i <= dispLine; ++i)
        {
            int len = (int)::SendMessage(hwnd, SCI_LINELENGTH, i, 0);
            std::string buf(len, '\0');
            ::SendMessage(hwnd, SCI_GETLINE, i, (LPARAM)&buf[0]);
            buf.resize(strnlen(buf.c_str(), len));
            if (buf.find(marker) != std::string::npos)
                ++hitIndex;
        }

        const auto& hits = ResultDock::instance().hits();
        if (hitIndex < 0 || hitIndex >= (int)hits.size())
            return 0;
        const auto& hit = hits[hitIndex];

        // 6) Convert UTF-8 path → wide
        std::wstring wpath;
        int wlen = ::MultiByteToWideChar(CP_UTF8, 0,
            hit.fullPathUtf8.c_str(), -1,
            nullptr, 0);
        if (wlen > 0)
        {
            wpath.resize(wlen - 1);
            ::MultiByteToWideChar(CP_UTF8, 0,
                hit.fullPathUtf8.c_str(), -1,
                &wpath[0], wlen);
        }

        // 7) Switch file
        HWND hEd = nppData._scintillaMainHandle;
        int targetLine = (int)::SendMessage(hEd, SCI_LINEFROMPOSITION, hit.pos, 0);

        // 8) Jump + select in the editor
        int visible = (int)::SendMessage(hEd, SCI_LINESONSCREEN, 0, 0);
        int topLine = targetLine - visible / 2;
        if (topLine < 0)                          topLine = 0;
        int totalLines = (int)::SendMessage(hEd, SCI_GETLINECOUNT, 0, 0);
        if (topLine > totalLines - visible)       topLine = totalLines - visible;
        ::SendMessage(hEd, SCI_SETFIRSTVISIBLELINE, topLine, 0);

        ::SendMessage(hEd, SCI_GOTOPOS, hit.pos, 0);
        ::SendMessage(hEd, SCI_SETSEL, hit.pos, hit.pos + hit.length);

        // 9) Restore dock scroll
        ::SendMessage(hwnd, SCI_SETFIRSTVISIBLELINE, firstVisible, 0);

        // 10) Clear selection in the dock
        Sci_Position dockPos = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
        ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, dockPos, 0);

        // 11) Give focus back to the dock control so its caret‐line highlight stays visible
        ::SetFocus(hwnd);
        return 0;
    }

    case DMN_CLOSE:
        ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, (LPARAM)ResultDock::instance()._hDock);
        return TRUE;

    case WM_NCDESTROY:
        s_prevSciProc = nullptr;
        break;
    }

    return s_prevSciProc
        ? ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp)
        : ::DefWindowProc(hwnd, msg, wp, lp);
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
    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0) {
        return ::SendMessage(_hSci, m, w, l);
        };

    const int lines = (int)S(SCI_GETLINECOUNT);
    std::string txt;
    for (int i = 0; i < lines; ++i) {
        // get raw line
        int len = (int)S(SCI_LINELENGTH, i, 0);
        txt.resize(len);
        if (len > 0) S(SCI_GETLINE, i, reinterpret_cast<LPARAM>(&txt[0]));

        bool blank = (len == 0);
        // match either no indent or 4-space indent
        bool header =
            !blank
            && (txt.rfind("Search ", 0) == 0
                || txt.rfind("    Search ", 0) == 0);

        // header → BASE+HEADERFLAG, others → BASE+1
        int level = header
            ? (SC_FOLDLEVELBASE | SC_FOLDLEVELHEADERFLAG)
            : (SC_FOLDLEVELBASE + 1);

        if (blank) level |= SC_FOLDLEVELWHITEFLAG;
        S(SCI_SETFOLDLEVEL, i, level);
    }
}

void ResultDock::prependText(const std::wstring& wText)
{
    if (!_hSci || wText.empty()) {
        return;
    }

    // Convert the incoming wide string (UTF-16) to a UTF-8 encoded std::string.
    std::string utf8Text;
    int len = ::WideCharToMultiByte(CP_UTF8, 0, wText.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len > 1) {
        utf8Text.resize(len - 1);
        ::WideCharToMultiByte(CP_UTF8, 0, wText.c_str(), -1, &utf8Text[0], len, nullptr, nullptr);
    }
    else {
        return; // Nothing to insert.
    }

    // --- FIX: Add a separating newline if the dock is not empty ---
    LRESULT currentLength = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);
    if (currentLength > 0) {
        // Prepend a newline to separate this block from the previous one.
        utf8Text.insert(0, "\r\n");
    }
    // --- End of FIX ---

    // Temporarily make the control writable if it's in read-only mode.
    bool isReadOnly = (::SendMessage(_hSci, SCI_GETREADONLY, 0, 0) != 0);
    if (isReadOnly) {
        ::SendMessage(_hSci, SCI_SETREADONLY, 0, 0);
    }

    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);

    // Insert the correctly converted UTF-8 text at the beginning of the document.
    ::SendMessage(_hSci, SCI_INSERTTEXT, 0, reinterpret_cast<LPARAM>(utf8Text.c_str()));

    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);

    // Restore read-only state if it was set.
    if (isReadOnly) {
        ::SendMessage(_hSci, SCI_SETREADONLY, 1, 0);
    }

    // After prepending, the entire view needs to be restyled and folded again.
    _rebuildFolding();
    _applyTheme();
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

void ResultDock::recordHit(const std::string& fullPathUtf8, Sci_Position pos, Sci_Position length) {
    _hits.push_back({ fullPathUtf8, pos, length });
}

void ResultDock::clearAll() {
    _hits.clear();
    if (_hSci) {
        ::SendMessage(_hSci, SCI_CLEARALL, 0, 0);
    }
}

void ResultDock::prependHits(const std::vector<Hit>& newHits, const std::wstring& text)
{
    // 1) Prepend the new data to the internal data model.
    _hits.insert(_hits.begin(), newHits.begin(), newHits.end());

    if (!_hSci || text.empty()) {
        return;
    }

    // 2) Convert the incoming wstring (UTF-16) to a UTF-8 encoded std::string.
    std::string utf8PrependedText;
    int len = ::WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len > 1) {
        utf8PrependedText.resize(len - 1);
        ::WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, &utf8PrependedText[0], len, nullptr, nullptr);
    }
    else {
        return; // Nothing to insert.
    }

    ::SendMessage(_hSci, SCI_SETREADONLY, FALSE, 0);
    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);

    // --- START OF FIX ---

    // 3) Get the current document length BEFORE any insertion.
    LRESULT currentLength = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);

    // 4) Insert the new text block at the very beginning.
    ::SendMessage(_hSci, SCI_INSERTTEXT, 0, reinterpret_cast<LPARAM>(utf8PrependedText.c_str()));

    // 5) If the document was NOT empty before, insert a newline separator AFTER the new block.
    if (currentLength > 0) {
        const char* separator = "\r\n";
        ::SendMessage(_hSci, SCI_INSERTTEXT, utf8PrependedText.length(), reinterpret_cast<LPARAM>(separator));
    }

    // --- END OF FIX ---

    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);
    ::SendMessage(_hSci, SCI_SETREADONLY, TRUE, 0);

    // 6) After prepending, rebuild folding and apply themes.
    _rebuildFolding();
    _applyTheme();

    // 7) Explicitly set the scroll position and caret to the top of the document.
    // This prevents the view from jumping and keeps the caret at a predictable location.
    ::SendMessage(_hSci, SCI_SETFIRSTVISIBLELINE, 0, 0);
    ::SendMessage(_hSci, SCI_GOTOPOS, 0, 0);
}