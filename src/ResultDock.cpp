
#include "PluginDefinition.h"
#include <windows.h>
#include "Scintilla.h"
#include "StaticDialog/resource.h"
#include "StaticDialog/Docking.h"
#include "StaticDialog/DockingDlgInterface.h"
#include "ResultDock.h"
#include "image_data.h"
#include "LanguageManager.h"
#include <algorithm>
#include <string>
#include <commctrl.h>
#include <functional>
#include "Encoding.h"
#include <unordered_set>
#include <sstream>
#include <windowsx.h> 

extern NppData nppData;

namespace {
    LanguageManager& LM = LanguageManager::instance();

    static constexpr uint32_t argb(BYTE a, COLORREF c)
    {
        return (uint32_t(a) << 24) |
            (uint32_t(GetRValue(c)) << 16) |
            (uint32_t(GetGValue(c)) << 8) |
            uint32_t(GetBValue(c));
    }
}

// --- Singleton & Public Methods ------------------------------------------

ResultDock& ResultDock::instance()
{
    // This assumes a global g_inst is defined and set in DllMain.
    extern HINSTANCE g_inst;
    static ResultDock s{ g_inst };
    return s;
}

void ResultDock::ensureCreatedAndVisible(const NppData& npp)
{
    // 1) first-time creation
    if (!_hSci)                // _hSci is initialized in create()
        create(npp);

    // 2) show again – MUST use the client handle!
    if (_hSci)
        ::SendMessage(npp._nppHandle, NPPM_DMMSHOW, 0,
            reinterpret_cast<LPARAM>(_hSci));
}

// --- Public API Block (Multiple Search) ----------------------------------

// ---------------------------------------------------------------------
//  1. open a new block
// ---------------------------------------------------------------------
void ResultDock::startSearchBlock(const std::wstring& header,bool groupView,bool purge)
{
    if (purge)
        clear();                        // honour “Purge” before we start

    _pendingText.clear();
    _pendingHits.clear();
    _utf8LenPending = 0;
    _groupViewPending = groupView;
    _blockOpen = true;

    // one leading indent so classify() marks it as SearchHdr
    _pendingText = getIndentString(LineLevel::SearchHdr) + header + L"\r\n";
    _utf8LenPending = Encoding::wstringToUtf8(_pendingText).size();
}

// ---------------------------------------------------------------------
//  2. append results for ONE file (can be called many times)
// ---------------------------------------------------------------------
void ResultDock::appendFileBlock(const FileMap& fm, const SciSendFn& sciSend)
{
    if (!_blockOpen) return;

    std::wstring  partText;
    std::vector<Hit> partHits;

    // build WITHOUT another search header
    buildListText(fm, _groupViewPending, L"", sciSend,
        partText, partHits);

    shiftHits(partHits, _utf8LenPending);          // adjust offsets

    _pendingText += partText;
    _utf8LenPending += Encoding::wstringToUtf8(partText).size();

    _pendingHits.insert(_pendingHits.end(),
        std::make_move_iterator(partHits.begin()),
        std::make_move_iterator(partHits.end()));
}

// ---------------------------------------------------------------------
//  3. finalise header, insert block, restyle
// ---------------------------------------------------------------------
void ResultDock::closeSearchBlock(int totalHits, int totalFiles)
{
    if (!_blockOpen) return;

    // -- Replace the 2 numeric placeholders in the header line --------
    std::wstring hitsStr = std::to_wstring(totalHits);
    std::wstring filesStr = std::to_wstring(totalFiles);

    // search starts AFTER the single leading indent
    size_t pos = _pendingText.find_first_of(L"0123456789", /*start*/1);
    if (pos != std::wstring::npos)
    {
        size_t endPos = _pendingText.find_first_not_of(L"0123456789", pos);
        _pendingText.replace(pos, endPos - pos, hitsStr);

        size_t pos2 = _pendingText.find_first_of(
            L"0123456789", pos + hitsStr.size());
        if (pos2 != std::wstring::npos) {
            size_t endPos2 = _pendingText.find_first_not_of(L"0123456789", pos2);
            _pendingText.replace(pos2, endPos2 - pos2, filesStr);
        }
    }

    // -- Byte-offset correction for stored hits -----------------------
    size_t newUtf8 = Encoding::wstringToUtf8(_pendingText).size();
    ptrdiff_t delta = static_cast<ptrdiff_t>(newUtf8) -
        static_cast<ptrdiff_t>(_utf8LenPending);
    if (delta != 0) {
        for (auto& h : _pendingHits)
            h.displayLineStart += static_cast<int>(delta);
    }

    // -- Commit to the dock ------------------------------------------
    prependHits(_pendingHits, _pendingText);
    rebuildFolding();
    applyStyling();
    _blockOpen = false;
}

// --- Other Public Methods ------------------------------------------------

void ResultDock::clear()
{
    // ----- Data structures -------------------------------------------------
    _hits.clear();


    if (!_hSci)
        return;

    // ----- reset Scintilla‑Puffer completly --------------------------
    ::SendMessage(_hSci, SCI_SETREADONLY, FALSE, 0);
    ::SendMessage(_hSci, SCI_CLEARALL, 0, 0);             // Text & styles
    ::SendMessage(_hSci, SCI_SETREADONLY, TRUE, 0);

    // remove all folding levels → BASE
    const int lineCount = (int)::SendMessage(_hSci,
        SCI_GETLINECOUNT, 0, 0);
    for (int l = 0; l < lineCount; ++l)
        ::SendMessage(_hSci, SCI_SETFOLDLEVEL, l, SC_FOLDLEVELBASE);

    // Clear every indicator used by the dock
    for (int ind : { INDIC_LINE_BACKGROUND, INDIC_LINENUMBER_FORE,
        INDIC_MATCH_BG, INDIC_MATCH_FORE })
    {
        ::SendMessage(_hSci, SCI_SETINDICATORCURRENT, ind, 0);
        ::SendMessage(_hSci, SCI_INDICATORCLEARRANGE, 0,
            ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0));
    }

    // reset caret/scroll position
    ::SendMessage(_hSci, SCI_GOTOPOS, 0, 0);
    ::SendMessage(_hSci, SCI_SETFIRSTVISIBLELINE, 0, 0);

    // ----- Styles & Folding neu aufbauen -----------------------------------
    rebuildFolding();
    applyStyling();
}

void ResultDock::rebuildFolding()
{
    if (!_hSci) return;

    auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT
        { return ::SendMessage(_hSci, m, w, l); };

    const int lineCount = static_cast<int>(sciSend(SCI_GETLINECOUNT));

    // Enable folding (redundant if already set)
    sciSend(SCI_SETPROPERTY, (WPARAM)"fold", (LPARAM)"1");
    sciSend(SCI_SETPROPERTY, (WPARAM)"fold.compact", (LPARAM)"1");

    const int BASE = SC_FOLDLEVELBASE;

    // Pass 1: mark every line as plain content
    for (int l = 0; l < lineCount; ++l)
        sciSend(SCI_SETFOLDLEVEL, l, BASE);

    // Pass 2: detect semantic levels purely by leading spaces via classify()
    for (int l = 0; l < lineCount; ++l)
    {
        // read raw line
        const int rawLen = static_cast<int>(sciSend(SCI_LINELENGTH, l));
        std::string buf(rawLen, '\0');
        sciSend(SCI_GETLINE, l, reinterpret_cast<LPARAM>(buf.data()));
        buf.resize(strnlen(buf.c_str(), rawLen));

        LineKind kind = classify(buf);
        if (kind == LineKind::Blank) {                     // blank → plain content
            sciSend(SCI_SETFOLDLEVEL, l, BASE);
            continue;
        }

        // count leading spaces
        int spaces = 0;
        while (spaces < static_cast<int>(buf.size()) && buf[spaces] == ' ')
            ++spaces;

        bool isHeader = false;
        int  level = BASE;

        switch (kind) {
        case LineKind::SearchHdr:
            isHeader = true;  level = BASE + static_cast<int>(LineLevel::SearchHdr); break;
        case LineKind::FileHdr:
            isHeader = true;  level = BASE + static_cast<int>(LineLevel::FileHdr);   break;
        case LineKind::CritHdr:
            isHeader = true;  level = BASE + static_cast<int>(LineLevel::CritHdr);   break;
        default: break;       // HitLine or unknown stays content
        }

        if (isHeader) {
            sciSend(SCI_SETFOLDLEVEL, l, level | SC_FOLDLEVELHEADERFLAG);
        }
        else {
            int depth = 1;  // min 1 Level
            if (spaces >= INDENT_SPACES[static_cast<int>(LineLevel::HitLine)])
                depth = static_cast<int>(LineLevel::HitLine);   // 3
            else if (spaces >= INDENT_SPACES[static_cast<int>(LineLevel::CritHdr)])
                depth = static_cast<int>(LineLevel::CritHdr);   // 2
            else if (spaces >= INDENT_SPACES[static_cast<int>(LineLevel::FileHdr)])
                depth = static_cast<int>(LineLevel::FileHdr);   // 1

            sciSend(SCI_SETFOLDLEVEL, l, BASE + depth);
        }
    }

}

void ResultDock::applyStyling() const
{
    if (!_hSci)
        return;

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT
        { return ::SendMessage(_hSci, m, w, l); };

    // Step 1: Clear previous styling indicators
    const std::vector<int> indicatorsToClear = { INDIC_LINE_BACKGROUND, INDIC_LINENUMBER_FORE, INDIC_MATCH_FORE, INDIC_MATCH_BG };
    for (int indicator : indicatorsToClear) {
        S(SCI_SETINDICATORCURRENT, indicator);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // Step 2: Apply base style for each line (Header, File Path, Default)
    S(SCI_STARTSTYLING, 0, 0);

    const int lineCount = static_cast<int>(S(SCI_GETLINECOUNT));
    for (int line = 0; line < lineCount; ++line) {
        const Sci_Position lineStart = S(SCI_POSITIONFROMLINE, line, 0);
        const int lineLength = static_cast<int>(S(SCI_LINELENGTH, line, 0));

        int style = STYLE_DEFAULT;
        if (lineLength > 0)
        {
            std::string buf(lineLength, '\0');
            S(SCI_GETLINE, line, reinterpret_cast<LPARAM>(buf.data()));
            buf.resize(strnlen(buf.c_str(), lineLength));   // trim trailing NULs

            switch (classify(buf))
            {
            case LineKind::SearchHdr:   style = STYLE_HEADER;     break;
            case LineKind::CritHdr:     style = STYLE_CRITHDR;    break;
            case LineKind::FileHdr:     style = STYLE_FILEPATH;   break;
            case LineKind::HitLine:     /* keep default */        break;
            case LineKind::Blank:       /* keep default */        break;
            }
        }

        // Apply determined style to line content
        if (lineLength > 0) {
            S(SCI_SETSTYLING, lineLength, style);
        }

        // Apply default style to EOL characters
        const Sci_Position lineEnd = S(SCI_GETLINEENDPOSITION, line, 0);
        const int eolLength = static_cast<int>(lineEnd - (lineStart + lineLength));
        if (eolLength > 0) {
            S(SCI_SETSTYLING, eolLength, STYLE_DEFAULT);
        }
    }

    // Step 3: Overlay indicators for hits

    // 3a. Background for hit lines
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    for (const auto& hit : _hits) {
        if (hit.displayLineStart < 0)
            continue;

        int line = static_cast<int>(S(SCI_LINEFROMPOSITION, hit.displayLineStart, 0));
        Sci_Position startPos = S(SCI_POSITIONFROMLINE, line, 0);
        Sci_Position length = S(SCI_LINELENGTH, line, 0);

        if (length > 0)
            S(SCI_INDICATORFILLRANGE, startPos, length);
    }

    // 3b. Line-number digits
    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& hit : _hits) {
        if (hit.displayLineStart >= 0)
            S(SCI_INDICATORFILLRANGE, hit.displayLineStart + hit.numberStart, hit.numberLen);
    }

    // 3c. Match substrings (background and foreground)
    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_BG);
    for (const auto& hit : _hits) {
        if (hit.displayLineStart < 0)
            continue;

        for (size_t i = 0; i < hit.matchStarts.size(); ++i) {
            S(SCI_INDICATORFILLRANGE, hit.displayLineStart + hit.matchStarts[i], hit.matchLens[i]);
        }
    }

    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
    for (const auto& hit : _hits) {
        if (hit.displayLineStart < 0)
            continue;

        for (size_t i = 0; i < hit.matchStarts.size(); ++i) {
            S(SCI_INDICATORFILLRANGE, hit.displayLineStart + hit.matchStarts[i], hit.matchLens[i]);
        }
    }
}

void ResultDock::onThemeChanged() {
    applyTheme();
}

// --- Private Methods -----------------------------------------------------

ResultDock::ResultDock(HINSTANCE hInst)
    : _hInst(hInst)      // member‑initialiser‑list
    , _hSci(nullptr)
    , _hDock(nullptr)
{
    // no further code needed here – everything else happens in create()
}

void ResultDock::create(const NppData& npp)
{
    // 1) Create the Scintilla window that will become our dock client
    _hSci = ::CreateWindowExW(
        WS_EX_CLIENTEDGE,           // extended style (adds a thin 3‑D border)
        L"Scintilla",               // window class
        L"",                        // caption – not shown
        WS_CHILD | WS_CLIPSIBLINGS, // style – child because N++ owns the dock
        0, 0, 100, 100,             // initial size – N++ resizes later
        npp._nppHandle,             // parent – the N++ main frame
        nullptr,                    // no menu
        _hInst,                     // our DLL instance
        nullptr);                   // no creation data

    if (!_hSci) {
        ::MessageBoxW(npp._nppHandle,
            L"FATAL: Failed to create Scintilla window!",
            L"ResultDock Error",
            MB_OK | MB_ICONERROR);
        return;
    }

    // 2) Subclass Scintilla so we receive double‑clicks and other messages
    s_prevSciProc = reinterpret_cast<WNDPROC>(
        ::SetWindowLongPtrW(_hSci,
            GWLP_WNDPROC,
            reinterpret_cast<LONG_PTR>(sciSubclassProc)));

    // 3) Make Scintilla use UTF‑8 internally (matches our Hit list)
    ::SendMessageW(_hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0);

    // 4) Define size of horizontal scrollbar
    ::SendMessageW(_hSci, SCI_SETSCROLLWIDTHTRACKING, TRUE, 0);

    // 5) Fill docking descriptor (persists in _dockData)
    static UINT  s_cachedDpi = 0;
    static HICON s_cachedLightIcon = nullptr;

    ::ZeroMemory(&_dockData, sizeof(_dockData));
    _dockData.hClient = _hSci;
    _dockData.pszName = L"MultiReplace – Search results";
    _dockData.dlgID = IDD_MULTIREPLACE_RESULT_DOCK;
    _dockData.uMask = DWS_DF_CONT_BOTTOM | DWS_ICONTAB;
    _dockData.pszAddInfo = L"";
    _dockData.pszModuleName = NPP_PLUGIN_NAME;
    _dockData.iPrevCont = -1;              // no previous container
    _dockData.rcFloat = { 0, 0, 0, 0 };    // sensible default when undocked
    const bool darkMode = ::SendMessage(npp._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0;

    if (!darkMode) {
        // Light mode: generate or reuse colored icon
        UINT dpi = 96;
        HDC hdc = ::GetDC(npp._nppHandle);
        if (hdc) {
            dpi = ::GetDeviceCaps(hdc, LOGPIXELSX);
            ::ReleaseDC(npp._nppHandle, hdc);
        }

        if (dpi != s_cachedDpi) {
            // cleanup old icon
            if (s_cachedLightIcon)::DestroyIcon(s_cachedLightIcon);

            // create new bitmap & icon
            extern HBITMAP CreateBitmapFromArray(UINT dpi);
            HBITMAP bmp = CreateBitmapFromArray(dpi);
            if (bmp) {
                ICONINFO ii = {};
                ii.fIcon = TRUE;
                ii.hbmMask = bmp;
                ii.hbmColor = bmp;
                s_cachedLightIcon = ::CreateIconIndirect(&ii);
                ::DeleteObject(bmp);
            }
            s_cachedDpi = dpi;
        }

        // assign icon or fallback
        if (s_cachedLightIcon)
            _dockData.hIconTab = s_cachedLightIcon;
        else
            _dockData.hIconTab = static_cast<HICON>(::LoadImage(_hInst, MAKEINTRESOURCE(IDI_MR_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
    }
    else {
        // Dark mode: fallback to monochrome resource icon
        _dockData.hIconTab = static_cast<HICON>(::LoadImage(_hInst, MAKEINTRESOURCE(IDI_MR_DM_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
    }


    // 6) Register the dock window with Notepad++
    _hDock = reinterpret_cast<HWND>(::SendMessageW(npp._nppHandle, NPPM_DMMREGASDCKDLG, 0, reinterpret_cast<LPARAM>(&_dockData)));

    if (!_hDock) {
        ::MessageBoxW(npp._nppHandle,
            L"ERROR: Docking registration failed.",
            L"ResultDock Error",
            MB_OK | MB_ICONERROR);
        return;
    }

    // 7) Let Notepad++ apply its current light/dark theme to dock and Scintilla
    ::SendMessageW(npp._nppHandle,
        NPPM_DARKMODESUBCLASSANDTHEME,
        static_cast<WPARAM>(NppDarkMode::dmfInit),
        reinterpret_cast<LPARAM>(_hDock));

    ::SendMessage(_hSci, SCI_SETWRAPMODE, wrapEnabled() ? SC_WRAP_WORD : SC_WRAP_NONE, 0);

    // 8) Initialise folding and apply syntax colours that match current N++ theme
    initFolding();
    applyTheme();
}

void ResultDock::initFolding() const
{
    if (!_hSci) return;

    auto S = [this](UINT msg, WPARAM w = 0, LPARAM l = 0)
        { return ::SendMessage(_hSci, msg, w, l); };

    // 1) configure margin #2 for folding symbols -----------------
    constexpr int M_FOLD = 2;
    S(SCI_SETMARGINTYPEN, M_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN, M_FOLD, SC_MASK_FOLDERS);

    // 2) width: 16‑px box + 4 px stem so nothing is clipped ------
    const int h = static_cast<int>(S(SCI_TEXTHEIGHT));
    S(SCI_SETMARGINWIDTHN, M_FOLD, h + 4);

    // 3) hide margins 0/1 (line‑numbers & bookmarks) -------------
    S(SCI_SETMARGINWIDTHN, 0, 0);
    S(SCI_SETMARGINWIDTHN, 1, 0);

    // 4) header markers
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS);

    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED);

    // 5) guide‑line markers (│ ├ └) ------------------------------
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER);

    // 6) placeholder colour – themed colours set in applyTheme()
    constexpr COLORREF placeholder = RGB(200, 200, 200);
    for (int id : {
        SC_MARKNUM_FOLDER,
            SC_MARKNUM_FOLDEROPEN,
            SC_MARKNUM_FOLDERSUB,
            SC_MARKNUM_FOLDERMIDTAIL,
            SC_MARKNUM_FOLDERTAIL,
            SC_MARKNUM_FOLDEREND })
    {
        S(SCI_MARKERSETFORE, id, placeholder);
        S(SCI_MARKERSETBACK, id, placeholder);
    }

    // 7) enable mouse interaction & auto‑fold updates ------------
    S(SCI_SETMARGINSENSITIVEN, M_FOLD, TRUE);
    S(SCI_SETAUTOMATICFOLD, SC_AUTOMATICFOLD_CLICK);

    S(SCI_SETFOLDFLAGS, SC_FOLDFLAG_LINEAFTER_CONTRACTED);

    S(SCI_MARKERENABLEHIGHLIGHT, TRUE);
}

void ResultDock::applyTheme()
{
    if (!_hSci)
        return;

    // Helper lambda for Scintilla calls
    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT {
        return ::SendMessage(_hSci, m, w, l);
        };

    // Determine if dark mode is active
    const bool dark = ::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0;
    const DockThemeColors& theme = currentColors(dark);

    // Base editor colors from Notepad++
    const COLORREF editorBg = (COLORREF)::SendMessage(nppData._nppHandle, NPPM_GETEDITORDEFAULTBACKGROUNDCOLOR, 0, 0);
    const COLORREF editorFg = (COLORREF)::SendMessage(nppData._nppHandle, NPPM_GETEDITORDEFAULTFOREGROUNDCOLOR, 0, 0);

    // Reset styles and set default font
    S(SCI_STYLESETFORE, STYLE_DEFAULT, editorFg);
    S(SCI_STYLESETBACK, STYLE_DEFAULT, editorBg);
    S(SCI_STYLECLEARALL);
    S(SCI_STYLESETFONT, STYLE_DEFAULT, (LPARAM)"Consolas");
    S(SCI_STYLESETSIZE, STYLE_DEFAULT, 10);

    // Margin colors
    const COLORREF marginBg = dark ? RGB(0, 0, 0) : editorBg;
    const COLORREF marginFg = dark ? RGB(200, 200, 200) : RGB(80, 80, 80);

    for (int m = 0; m <= 2; ++m)
        S(SCI_SETMARGINBACKN, m, marginBg);

    S(SCI_STYLESETBACK, STYLE_LINENUMBER, marginBg);
    S(SCI_STYLESETFORE, STYLE_LINENUMBER, marginFg);
    S(SCI_SETFOLDMARGINCOLOUR, TRUE, marginBg);
    S(SCI_SETFOLDMARGINHICOLOUR, TRUE, marginBg);

    // Selection colors
    const COLORREF selBg = dark ? RGB(96, 96, 96) : RGB(224, 224, 224);
    const COLORREF selFg = dark ? RGB(255, 255, 255) : editorFg;

    S(SCI_SETSELFORE, TRUE, selFg);
    S(SCI_SETSELBACK, TRUE, selBg);
    S(SCI_SETSELALPHA, 256, 0);
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_BACK, argb(0xFF, selBg));
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_TEXT, argb(0xFF, selFg));
    S(SCI_SETADDITIONALSELFORE, selFg);
    S(SCI_SETADDITIONALSELBACK, selBg);
    S(SCI_SETADDITIONALSELALPHA, 256, 0);

    // Fold markers
    for (int id : {
        SC_MARKNUM_FOLDER,
            SC_MARKNUM_FOLDEREND,
            SC_MARKNUM_FOLDEROPEN,
            SC_MARKNUM_FOLDEROPENMID,
            SC_MARKNUM_FOLDERSUB,
            SC_MARKNUM_FOLDERMIDTAIL,
            SC_MARKNUM_FOLDERTAIL
    }) {
        S(SCI_MARKERSETBACK, id, theme.foldGlyph);
        S(SCI_MARKERSETFORE, id, marginBg);
        S(SCI_MARKERSETBACKSELECTED, id, theme.foldHighlight);
    }

    // Caret line
    S(SCI_SETCARETLINEVISIBLE, TRUE, 0);
    S(SCI_SETCARETLINEBACK, theme.caretLineBg, 0);
    S(SCI_SETCARETLINEBACKALPHA, theme.caretLineAlpha);


    // Line background indicator
    S(SCI_INDICSETSTYLE, INDIC_LINE_BACKGROUND, INDIC_HIDDEN);

    // Line number indicator
    S(SCI_INDICSETSTYLE, INDIC_LINENUMBER_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_LINENUMBER_FORE, theme.lineNr);

    // Match background indicator
    if (dark) {
        // Dark Mode: hide background highlight
        S(SCI_INDICSETSTYLE, INDIC_MATCH_BG, INDIC_HIDDEN);
    }
    else {
        // Light Mode: use theme.matchBg for background
        S(SCI_INDICSETSTYLE, INDIC_MATCH_BG, INDIC_STRAIGHTBOX);
        S(SCI_INDICSETFORE, INDIC_MATCH_BG, theme.matchBg);
        S(SCI_INDICSETALPHA, INDIC_MATCH_BG, 100);
        S(SCI_INDICSETUNDER, INDIC_MATCH_BG, TRUE);
    }

    // Rote Match-Farbe
    S(SCI_INDICSETSTYLE, INDIC_MATCH_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_MATCH_FORE, theme.matchFg);
    S(SCI_INDICSETUNDER, INDIC_MATCH_FORE, TRUE);

    // Header style
    S(SCI_STYLESETFORE, STYLE_HEADER, theme.headerFg);
    S(SCI_STYLESETBACK, STYLE_HEADER, theme.headerBg);
    S(SCI_STYLESETBOLD, STYLE_HEADER, TRUE);
    S(SCI_STYLESETEOLFILLED, STYLE_HEADER, TRUE);

    // CritHdr style
    S(SCI_STYLESETFORE, STYLE_CRITHDR, theme.critHdrFg);
    S(SCI_STYLESETBACK, STYLE_CRITHDR, theme.critHdrBg);
    S(SCI_STYLESETBOLD, STYLE_CRITHDR, TRUE);
    S(SCI_STYLESETEOLFILLED, STYLE_CRITHDR, TRUE);

    // File path style
    S(SCI_STYLESETFORE, STYLE_FILEPATH, theme.filePathFg);
    S(SCI_STYLESETBACK, STYLE_FILEPATH, editorBg);
    S(SCI_STYLESETBOLD, STYLE_FILEPATH, TRUE);
    S(SCI_STYLESETITALIC, STYLE_FILEPATH, TRUE);
    S(SCI_STYLESETEOLFILLED, STYLE_FILEPATH, TRUE);
}

void ResultDock::prependHits(const std::vector<Hit>& newHits, const std::wstring& text)
{
    // 1) prepend data objects
    _hits.insert(_hits.begin(), newHits.begin(), newHits.end());

    if (!_hSci || text.empty())
        return;

    // 2) convert wide → UTF‑8 7
    std::string utf8PrependedText;
    int len = ::WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1,
        nullptr, 0, nullptr, nullptr);
    if (len <= 1)
        return;                       // nothing to insert

    utf8PrependedText.resize(len - 1);
    ::WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1,
        &utf8PrependedText[0], len, nullptr, nullptr);

    LRESULT currentLen = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);

    // bytes that will be inserted at position 0 (text + optional “\r\n”)
    const int delta = static_cast<int>(utf8PrependedText.length()) +
        (currentLen > 0 ? 2 /* “\r\n” */ : 0);

    // shift every *existing* hit so its dock offset stays valid
    const size_t startOld = newHits.size();          // first “old” hit
    for (size_t i = startOld; i < _hits.size(); ++i)
        _hits[i].displayLineStart += delta;
    // numberStart & matchStarts are relative to displayLineStart,
    // so no update needed for them

    ::SendMessage(_hSci, SCI_SETREADONLY, FALSE, 0);
    ::SendMessage(_hSci, SCI_BEGINUNDOACTION, 0, 0);

    // 3) remember old length, insert new block + optional separator
    LRESULT oldLen = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);
    ::SendMessage(_hSci, SCI_INSERTTEXT, 0,
        reinterpret_cast<LPARAM>(utf8PrependedText.c_str()));
    if (oldLen > 0)
        ::SendMessage(_hSci, SCI_INSERTTEXT, utf8PrependedText.length(),
            reinterpret_cast<LPARAM>("\r\n"));

    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);
    ::SendMessage(_hSci, SCI_SETREADONLY, TRUE, 0);

    rebuildFolding();
    applyTheme();
    applyStyling();
    collapseOldSearches();

    ::SendMessage(_hSci, SCI_SETFIRSTVISIBLELINE, 0, 0);
    ::SendMessage(_hSci, SCI_GOTOPOS, 0, 0);
}

void ResultDock::collapseOldSearches()
{
    if (!_hSci) return;

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { return ::SendMessage(_hSci, m, w, l); };

    bool newestSeen = false;
    const int lines = (int)S(SCI_GETLINECOUNT);

    for (int l = 0; l < lines; ++l) {
        const int lvl = (int)S(SCI_GETFOLDLEVEL, l);

        // root header?  (= Level 0 + Header‑Flag)
        const bool isRootHeader =
            (lvl & SC_FOLDLEVELHEADERFLAG) &&
            ((lvl & SC_FOLDLEVELNUMBERMASK) == SC_FOLDLEVELBASE);

        if (!isRootHeader)
            continue;

        if (!newestSeen) {
            newestSeen = true;               // keep newest block open
        }
        else {
            // only collapse when it’s still expanded
            if (S(SCI_GETFOLDEXPANDED, l))
                S(SCI_FOLDLINE, l, SC_FOLDACTION_CONTRACT);
        }
    }
}

// --- Private Formatting Helpers ------------------------------------------

void ResultDock::buildListText(
    const FileMap& files,
    bool groupView,
    const std::wstring& header,
    const SciSendFn& sciSend,
    std::wstring& outText,
    std::vector<Hit>& outHits) const
{
    std::wstring body;
    size_t utf8Len = 0;

    auto appendIndented = [&](LineLevel lvl, const std::wstring& txt)
        {
            std::wstring line = getIndentString(lvl) + txt + L"\r\n";
            body += line;
            utf8Len += getIndentUtf8Length(lvl)
                + Encoding::wstringToUtf8(txt + L"\r\n").size();
        };

    if (header.empty())
    {
        // Called from multi-document search: we do NOT want a second SearchHdr.
        // Simply skip the SearchHdr generation and start with the file loop.
    }
    else
    {
        appendIndented(LineLevel::SearchHdr, header);
    }

    for (auto& [path, f] : files)
    {
        appendIndented(LineLevel::FileHdr, f.wPath + L" " + LM.get(L"dock_hits_suffix", { std::to_wstring(f.hitCount) }));

        if (groupView)
        {
            // grouped: first the file header, then each criterion block
            for (auto& c : f.crits)
            {
                appendIndented(LineLevel::CritHdr, LM.get(L"dock_crit_header", { c.text, std::to_wstring(c.hits.size()) }));

                // make a mutable copy of c.hits
                auto hitsCopy = c.hits;
                // pass the copy to formatHitsLines instead of const c.hits
                formatHitsLines(sciSend, hitsCopy, body, utf8Len);
                // move results from hitsCopy into outHits
                outHits.insert(outHits.end(),
                    std::make_move_iterator(hitsCopy.begin()),
                    std::make_move_iterator(hitsCopy.end()));
            }
        }
        else
        {
            // flat: collect all hits per file and sort by position
            std::vector<Hit> merged;
            for (auto& c : f.crits)
                merged.insert(merged.end(),
                    std::make_move_iterator(c.hits.begin()),
                    std::make_move_iterator(c.hits.end()));
            std::sort(merged.begin(), merged.end(),
                [](auto const& a, auto const& b) { return a.pos < b.pos; });

            formatHitsLines(sciSend, merged, body, utf8Len);
            outHits.insert(outHits.end(),
                std::make_move_iterator(merged.begin()),
                std::make_move_iterator(merged.end()));
        }
    }

    outText = body;

}

void ResultDock::formatHitsLines(const SciSendFn& sciSend,
    std::vector<Hit>& hits,
    std::wstring& out,
    size_t& utf8Pos) const
{
    // Get the document code page (ANSI, UTF-8, etc.)
    const UINT docCp = static_cast<UINT>(sciSend(SCI_GETCODEPAGE, 0, 0));

    // Cache indentation for hit lines once
    const std::wstring indentHitW = getIndentString(LineLevel::HitLine);
    const size_t       indentHitU8 = getIndentUtf8Length(LineLevel::HitLine);

    // Pre-converted constant pieces of the prefix
    const std::wstring kLineW = LM.get(L"dock_line") + L" ";
    static const size_t       kLineU8 = Encoding::wstringToUtf8(kLineW).size();
    constexpr size_t          kColonSpaceU8 = 2;    // ": "

    // STEP 1: determine line numbers and max digit width for right alignment
    std::vector<int> lineNumbers;
    lineNumbers.reserve(hits.size());

    size_t maxDigits = 0;
    for (const Hit& h : hits) {
        int line1 = static_cast<int>(sciSend(SCI_LINEFROMPOSITION, h.pos, 0)) + 1;
        lineNumbers.push_back(line1);
        maxDigits = (std::max)(maxDigits, std::to_wstring(line1).size());
    }

    // Helper: pad a line number to given width (right-aligned)
    auto padNumber = [](int number, size_t width) -> std::wstring {
        std::wstring numStr = std::to_wstring(number);
        size_t pad = (width > numStr.size() ? width - numStr.size() : 0);
        return std::wstring(pad, L' ') + numStr;
        };

    int   prevDocLine = -1;
    Hit* firstHitOnRow = nullptr;
    size_t hitIdx = 0;

    // STEP 2: iterate hits and build output lines
    for (Hit& h : hits) {
        int line1 = lineNumbers[hitIdx++];
        int line0 = line1 - 1;

        // Read raw bytes of this line from Scintilla
        int rawLen = static_cast<int>(sciSend(SCI_LINELENGTH, line0, 0));
        std::string raw(rawLen, '\0');
        sciSend(SCI_GETLINE, line0, reinterpret_cast<LPARAM>(raw.data()));
        while (!raw.empty() && (raw.back() == '\r' || raw.back() == '\n'))
            raw.pop_back();

        const std::string_view sliceBytes(raw);

        // 1) UTF-8 bytes for accurate byte-offset highlighting
        const std::string sliceU8 = Encoding::bytesToUtf8(std::string(sliceBytes), docCp);
        // 2) Wide string for display in the dock
        const std::wstring sliceW = Encoding::utf8ToWString(sliceU8);

        // Build the line-number prefix (with right alignment)
        std::wstring paddedNumW = padNumber(line1, maxDigits);
        std::wstring prefixW = indentHitW + kLineW + paddedNumW + L": ";

        // Prefix length in UTF-8 (computed arithmetically)
        const size_t prefixU8Len = indentHitU8 + kLineU8 + paddedNumW.size() + kColonSpaceU8;

        // Compute byte-offsets for the match start within sliceU8
        Sci_Position absLineStart = sciSend(SCI_POSITIONFROMLINE, line0, 0);
        size_t relBytes = static_cast<size_t>(h.pos - absLineStart);

        size_t hitStartU8 = Encoding::bytesToUtf8(
            std::string(sliceBytes.substr(0, relBytes)), docCp).size();
        size_t hitLenU8 = Encoding::bytesToUtf8(
            std::string(sliceBytes.substr(relBytes, h.length)), docCp).size();

        if (line0 != prevDocLine) {
            // First hit on this visible row
            out += prefixW + sliceW + L"\r\n";

            h.displayLineStart = static_cast<int>(utf8Pos);
            h.numberStart = static_cast<int>(indentHitU8 + kLineU8 +
                paddedNumW.size() - std::to_wstring(line1).size());
            h.numberLen = static_cast<int>(std::to_wstring(line1).size());

            h.matchStarts = { static_cast<int>(prefixU8Len + hitStartU8) };
            h.matchLens = { static_cast<int>(hitLenU8) };

            utf8Pos += prefixU8Len + sliceU8.size() + 2;   // include CR/LF

            firstHitOnRow = &h;
            prevDocLine = line0;
        }
        else {
            // Additional hit on the same visual row
            assert(firstHitOnRow);
            firstHitOnRow->matchStarts.push_back(static_cast<int>(prefixU8Len + hitStartU8));
            firstHitOnRow->matchLens.push_back(static_cast<int>(hitLenU8));
            h.displayLineStart = -1;   // mark as dummy
        }
    }

    // Remove dummy entries (all but the first hit per visual line)
    hits.erase(std::remove_if(hits.begin(), hits.end(),
        [](const Hit& e) { return e.displayLineStart < 0; }),
        hits.end());
}

// --- Private Static Helpers ----------------------------------------------

ResultDock::LineKind ResultDock::classify(const std::string& raw) {
    size_t spaces = 0;
    while (spaces < raw.size() && raw[spaces] == ' ') ++spaces;

    std::string_view trimmed(raw.data() + spaces, raw.size() - spaces);
    if (trimmed.empty()) return LineKind::Blank;

    // Hier vollqualifizieren:
    for (int lvl = 0; lvl <= static_cast<int>(ResultDock::LineLevel::HitLine); ++lvl) {
        if ((int)spaces == ResultDock::INDENT_SPACES[lvl]) {
            switch (static_cast<ResultDock::LineLevel>(lvl)) {
            case ResultDock::LineLevel::SearchHdr: return LineKind::SearchHdr;
            case ResultDock::LineLevel::FileHdr:   return LineKind::FileHdr;
            case ResultDock::LineLevel::CritHdr:   return LineKind::CritHdr;
            case ResultDock::LineLevel::HitLine:   return LineKind::HitLine;
            }
        }
    }
    return LineKind::Blank;
}

std::wstring ResultDock::getIndentString(LineLevel lvl) {
        return std::wstring( INDENT_SPACES[static_cast<int>(lvl)], L' ' );
    }

size_t ResultDock::getIndentUtf8Length(ResultDock::LineLevel lvl) {
        return Encoding::wstringToUtf8(getIndentString(lvl)).size();
    }

void ResultDock::shiftHits(std::vector<ResultDock::Hit>& v, size_t delta)
    {
        for (auto& h : v)
            h.displayLineStart += static_cast<int>(delta);
    }

std::wstring ResultDock::stripHitPrefix(const std::wstring& w)
    {
        const int indentLen = ResultDock::INDENT_SPACES[static_cast<int>(ResultDock::LineLevel::HitLine)];
        size_t i = (std::min)(static_cast<size_t>(indentLen), w.size()); // skip indent

        while (i < w.size() && iswspace(w[i])) ++i;   // extra padding
        while (i < w.size() && iswalpha(w[i])) ++i;   // “Line”, “Zeile”, …
        if (i < w.size() && w[i] == L' ') ++i;        // single space

        size_t digitStart = i;
        while (i < w.size() && iswdigit(w[i])) ++i;   // line number
        if (i == digitStart || i >= w.size() || w[i] != L':') return w;
        ++i;                                          // skip ':'

        while (i < w.size() && iswspace(w[i])) ++i;   // space after ':'
        return w.substr(i);                           // pure hit text
    }

std::wstring ResultDock::pathFromFileHdr(const std::wstring& w)
    {
        const int indentLen = ResultDock::INDENT_SPACES[static_cast<int>(ResultDock::LineLevel::FileHdr)];

        // skip exact FileHdr indent
        size_t pos = 0;
        while (pos < w.size() && pos < static_cast<size_t>(indentLen) && w[pos] == L' ')
            ++pos;

        std::wstring s = w.substr(pos);           // path + " (…)"

        // remove trailing " (…)"
        size_t r = s.find_last_of(L')');
        if (r != std::wstring::npos)
        {
            size_t l = s.rfind(L'(', r);
            if (l != std::wstring::npos && l > 0 && s[l - 1] == L' ')
                s.erase(l - 1);                   // also drop preceding space
        }
        return s;
    }

int ResultDock::ancestorFileLine(HWND hSci, int startLine) {
        for (int l = startLine; l >= 0; --l) {
            int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
            std::string raw(len, '\0');
            ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
            raw.resize(strnlen(raw.c_str(), len));
            if (classify(raw) == LineKind::FileHdr) return l;
        }
        return -1;
    }

std::wstring ResultDock::getLineW(HWND hSci, int line) {
        int len = (int)::SendMessage(hSci, SCI_LINELENGTH, line, 0);
        std::string raw(len, '\0');
        ::SendMessage(hSci, SCI_GETLINE, line, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), len));
        return Encoding::utf8ToWString(raw);
    }

int ResultDock::leadingSpaces(const char* line, int len)
{
    int s = 0;
    while (s < len && line[s] == ' ') ++s;
    return s;
}

// --- Private Context Menu Handlers ---------------------------------------

void ResultDock::copySelectedLines(HWND hSci) {

    auto indentOf = [](ResultDock::LineLevel lvl) -> int {                                                                  // NEW
        return ResultDock::INDENT_SPACES[static_cast<int>(lvl)];
        };

    // determine selection range
    Sci_Position a = ::SendMessage(hSci, SCI_GETSELECTIONNANCHOR, 0, 0);
    Sci_Position c = ::SendMessage(hSci, SCI_GETCURRENTPOS, 0, 0);
    if (a > c) std::swap(a, c);
    int lineStart = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, a, 0);
    int lineEnd = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, c, 0);
    bool hasSel = (a != c);

    std::vector<std::wstring> out;

    if (hasSel) {                               // ----- selection mode
        for (int l = lineStart; l <= lineEnd; ++l) {
            int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
            std::string raw(len, '\0');
            ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
            raw.resize(strnlen(raw.c_str(), len));
            if (classify(raw) == LineKind::HitLine)
                out.emplace_back(stripHitPrefix(Encoding::utf8ToWString(raw)));
        }
    }
    else {                                    // ----- caret hierarchy walk
        int caretLine = lineStart;
        int len = (int)::SendMessage(hSci, SCI_LINELENGTH, caretLine, 0);
        std::string rawCaret(len, '\0');
        ::SendMessage(hSci, SCI_GETLINE, caretLine, (LPARAM)rawCaret.data());
        rawCaret.resize(strnlen(rawCaret.c_str(), len));
        LineKind kind = classify(rawCaret);

        auto pushHitsBelow = [&](int fromLine, int minIndent) {
            int total = (int)::SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
            for (int l = fromLine; l < total; ++l) {
                int lLen = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
                std::string raw(lLen, '\0');
                ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
                raw.resize(strnlen(raw.c_str(), lLen));

                // stop when leaving the subtree
                if ((int)ResultDock::leadingSpaces(raw.c_str(), (int)raw.size()) <= minIndent &&
                    classify(raw) != LineKind::HitLine) break;

                if (classify(raw) == LineKind::HitLine)
                    out.emplace_back(stripHitPrefix(Encoding::utf8ToWString(raw)));
            }
            };

        switch (kind) {
        case LineKind::HitLine:
            out.emplace_back(stripHitPrefix( Encoding::utf8ToWString(rawCaret)));
            break;

        case LineKind::CritHdr:
            pushHitsBelow(caretLine + 1, indentOf(LineLevel::CritHdr));
            break;

        case LineKind::FileHdr:
            pushHitsBelow(caretLine + 1, indentOf(LineLevel::FileHdr));
            break;

        case LineKind::SearchHdr:
            pushHitsBelow(caretLine + 1, indentOf(LineLevel::SearchHdr));
            break;
        default: break;
        }
    }

    if (!out.empty()) {
        std::wstring joined;
        for (size_t i = 0; i < out.size(); ++i) {
            joined += out[i];                   // << keeps original behaviour
        }
        copyTextToClipboard(hSci, joined);
    }
}

void ResultDock::copyTextToClipboard(HWND owner, const std::wstring& wText) {
    if (!::OpenClipboard(owner)) return;
    ::EmptyClipboard();
    HGLOBAL hMem = ::GlobalAlloc(GMEM_MOVEABLE, (wText.size() + 1) * sizeof(wchar_t));
    if (hMem) {
        auto* buf = static_cast<wchar_t*>(::GlobalLock(hMem));
        std::copy(wText.c_str(), wText.c_str() + wText.size() + 1, buf);
        ::GlobalUnlock(hMem);
        ::SetClipboardData(CF_UNICODETEXT, hMem);
    }
    ::CloseClipboard();
}

void ResultDock::copySelectedPaths(HWND hSci) {
    Sci_Position a = ::SendMessage(hSci, SCI_GETSELECTIONNANCHOR, 0, 0);
    Sci_Position c = ::SendMessage(hSci, SCI_GETCURRENTPOS, 0, 0);
    if (a > c) std::swap(a, c);
    int lineStart = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, a, 0);
    int lineEnd = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, c, 0);
    bool hasSel = (a != c);

    std::vector<std::wstring> paths;
    std::unordered_set<std::wstring> seen;
    auto addUnique = [&](const std::wstring& p) { if (seen.insert(p).second) paths.push_back(p); };

    auto pushFileHdrsBelow = [&](int fromLine, int minIndent) {
        int total = (int)::SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
        for (int l = fromLine; l < total; ++l) {
            int lLen = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
            std::string raw(lLen, '\0');
            ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
            raw.resize(strnlen(raw.c_str(), lLen));

            int indent = ResultDock::leadingSpaces(raw.c_str(), (int)raw.size());
            if (indent <= minIndent && classify(raw) != LineKind::HitLine) break;

            if (classify(raw) == LineKind::FileHdr)
                addUnique(pathFromFileHdr(Encoding::utf8ToWString(raw)));
        }
        };

    if (hasSel) {
        for (int l = lineStart; l <= lineEnd; ++l) {
            int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
            std::string raw(len, '\0');
            ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
            raw.resize(strnlen(raw.c_str(), len));
            LineKind k = classify(raw);

            if (k == LineKind::FileHdr)
                addUnique(pathFromFileHdr(Encoding::utf8ToWString(raw)));
            else if (k == LineKind::HitLine || k == LineKind::CritHdr) {
                int fLine = ancestorFileLine(hSci, l);
                if (fLine != -1)
                    addUnique(pathFromFileHdr(getLineW(hSci, fLine)));
            }
        }
    }
    else {
        int caretLine = lineStart;
        int len = (int)::SendMessage(hSci, SCI_LINELENGTH, caretLine, 0);
        std::string raw(len, '\0');
        ::SendMessage(hSci, SCI_GETLINE, caretLine, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), len));
        LineKind kind = classify(raw);

        switch (kind) {
        case LineKind::HitLine:
        case LineKind::CritHdr: {
            int fLine = ancestorFileLine(hSci, caretLine);
            if (fLine != -1)
                addUnique(pathFromFileHdr(getLineW(hSci, fLine)));
            break;
        }
        case LineKind::FileHdr:
            addUnique(pathFromFileHdr(Encoding::utf8ToWString(raw)));
            break;
        case LineKind::SearchHdr:
            pushFileHdrsBelow(caretLine + 1, 0);
            break;
        default: break;
        }
    }

    if (!paths.empty()) {
        std::wstring joined;
        for (size_t i = 0; i < paths.size(); ++i) {
            if (i) joined += L"\r\n";
            joined += paths[i];
        }
        copyTextToClipboard(hSci, joined);
    }
}

void ResultDock::openSelectedPaths(HWND hSci) {
    copySelectedPaths(hSci);   // copy first (UX)

    Sci_Position a = ::SendMessage(hSci, SCI_GETSELECTIONNANCHOR, 0, 0);
    Sci_Position c = ::SendMessage(hSci, SCI_GETCURRENTPOS, 0, 0);
    if (a > c) std::swap(a, c);
    int lineStart = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, a, 0);
    int lineEnd = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, c, 0);

    std::unordered_set<std::wstring> opened;

    for (int l = lineStart; l <= lineEnd; ++l) {
        int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
        std::string raw(len, '\0');
        ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), len));
        LineKind k = classify(raw);

        if (k == LineKind::FileHdr) {
            std::wstring p = pathFromFileHdr(Encoding::utf8ToWString(raw));
            if (opened.insert(p).second)
                ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)p.c_str());
        }
        else if (k == LineKind::HitLine || k == LineKind::CritHdr) {
            int fLine = ancestorFileLine(hSci, l);
            if (fLine != -1) {
                std::wstring p = pathFromFileHdr(getLineW(hSci, fLine));
                if (opened.insert(p).second)
                    ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)p.c_str());
            }
        }
    }
}

void ResultDock::deleteSelectedItems(HWND hSci)
{
    auto& dock = ResultDock::instance();

    // ------------------------------------------------------------------
    //  Determine anchor / caret lines
    // ------------------------------------------------------------------
    Sci_Position a = ::SendMessage(hSci, SCI_GETSELECTIONNANCHOR, 0, 0);
    Sci_Position c = ::SendMessage(hSci, SCI_GETCURRENTPOS, 0, 0);
    if (a > c) std::swap(a, c);

    int firstLine = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, a, 0);
    int lastLine = (int)::SendMessage(hSci, SCI_LINEFROMPOSITION, c, 0);
    bool hasSel = (a != c);

    // ------------------------------------------------------------------
    //  Build list of display‑line ranges that must be deleted
    // ------------------------------------------------------------------
    struct DelRange { int first; int last; };
    std::vector<DelRange> ranges;

    auto subtreeEnd = [&](int fromLine, int minIndent) -> int
        {
            int total = (int)::SendMessage(hSci, SCI_GETLINECOUNT, 0, 0);
            for (int l = fromLine; l < total; ++l)
            {
                int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
                std::string raw(len, '\0');
                ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
                raw.resize(strnlen(raw.c_str(), len));

                int indent = ResultDock::leadingSpaces(raw.c_str(), (int)raw.size());
                if (indent <= minIndent && classify(raw) != LineKind::HitLine)
                    return l - 1;                           // previous line ends subtree
            }
            return (int)::SendMessage(hSci, SCI_GETLINECOUNT, 0, 0) - 1; // EOF
        };

    auto pushRange = [&](int first, int last)
        {
            if (ranges.empty() || first > ranges.back().last + 1)
                ranges.push_back({ first, last });
            else
                ranges.back().last = (std::max)(ranges.back().last, last);
        };

    // ------------------------------------------------------------------
    //  Selection mode: collect every header / hit and expand hierarchy
    // ------------------------------------------------------------------
    if (hasSel) {
        for (int l = firstLine; l <= lastLine; ++l)
        {
            // skip lines already included by previous pushRange
            if (!ranges.empty() && l <= ranges.back().last) continue;

            int len = (int)::SendMessage(hSci, SCI_LINELENGTH, l, 0);
            std::string raw(len, '\0');
            ::SendMessage(hSci, SCI_GETLINE, l, (LPARAM)raw.data());
            raw.resize(strnlen(raw.c_str(), len));

            LineKind kind = classify(raw);
            int endLine = l;   // default: only this line

            switch (kind) {
            case LineKind::HitLine: endLine = l;
                break;
            case LineKind::CritHdr:
                endLine = subtreeEnd( l + 1, INDENT_SPACES[static_cast<int>(LineLevel::CritHdr)]);
                break;
            case LineKind::FileHdr:
                endLine = subtreeEnd( l + 1, INDENT_SPACES[static_cast<int>(LineLevel::FileHdr)]);
                break;
            case LineKind::SearchHdr:
                endLine = subtreeEnd( l + 1, INDENT_SPACES[static_cast<int>(LineLevel::SearchHdr)]);
                break;
            default:                  continue;                       // ignore blanks
            }
            pushRange(l, endLine);
        }
    }
    else {
        // ------------------------------------------------------------------
        //  Caret‑only mode: single logical block
        // ------------------------------------------------------------------
        int len = (int)::SendMessage(hSci, SCI_LINELENGTH, firstLine, 0);
        std::string raw(len, '\0');
        ::SendMessage(hSci, SCI_GETLINE, firstLine, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), len));

        LineKind kind = classify(raw);
        int endLine = firstLine;

        switch (kind)
        {
        case LineKind::HitLine: endLine = firstLine; break;
        case LineKind::CritHdr: endLine = subtreeEnd( firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::CritHdr)]);
            break;
        case LineKind::FileHdr: endLine = subtreeEnd( firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::FileHdr)]); 
            break;
        case LineKind::SearchHdr: endLine = subtreeEnd( firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::SearchHdr)]); 
            break;
        default: return; // nothing deletable on this line
        }
        ranges.push_back({ firstLine, endLine });
    }

    if (ranges.empty())
        return;

    // ------------------------------------------------------------------
    //  Delete ranges bottom‑up and update _hits offsets
    // ------------------------------------------------------------------
    ::SendMessage(hSci, SCI_SETREADONLY, FALSE, 0);

    for (auto it = ranges.rbegin(); it != ranges.rend(); ++it) {
        int fLine = it->first;
        int lLine = it->last;

        Sci_Position start = ::SendMessage(hSci, SCI_POSITIONFROMLINE, fLine, 0);
        Sci_Position end = (lLine + 1 < ::SendMessage(hSci, SCI_GETLINECOUNT, 0, 0))
            ? ::SendMessage(hSci, SCI_POSITIONFROMLINE, lLine + 1, 0)
            : ::SendMessage(hSci, SCI_GETLENGTH, 0, 0);

        Sci_Position delta = end - start;
        if (delta <= 0) continue;

        // ---- remove from Scintilla buffer
        ::SendMessage(hSci, SCI_DELETERANGE, start, delta);

        // ---- sync internal hit list
        for (auto h = dock._hits.begin(); h != dock._hits.end(); )
        {
            if (h->displayLineStart >= start && h->displayLineStart < end)
            {
                h = dock._hits.erase(h);                  // entry deleted
            }
            else
            {
                if (h->displayLineStart >= end)
                    h->displayLineStart -= (int)delta;    // shift left
                ++h;
            }
        }
    }

    ::SendMessage(hSci, SCI_SETREADONLY, TRUE, 0);

    // ------------------------------------------------------------------
    //  Re‑fold and re‑style remaining lines
    // ------------------------------------------------------------------
    dock.rebuildFolding();
    dock.applyStyling();
}

// --- Private Callbacks ---------------------------------------------------

LRESULT CALLBACK ResultDock::sciSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    extern NppData nppData;

    switch (msg) {
    case WM_NOTIFY: {
        NMHDR* hdr = reinterpret_cast<NMHDR*>(lp);
        if (hdr->code == DMN_CLOSE) {
            ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(hwnd));
            return TRUE;
        }

        SCNotification* scn = reinterpret_cast<SCNotification*>(lp);
        if (scn->nmhdr.code == SCN_MARGINCLICK && scn->margin == 2) {
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
        int x = LOWORD(lp);
        int y = HIWORD(lp);
        Sci_Position pos = ::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
        if (pos < 0)                         // guard for empty area
            return 0;
        int dispLine = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0);

        // --- Toggle fold if the clicked line is a header ----------
        int level = (int)::SendMessage(hwnd, SCI_GETFOLDLEVEL, dispLine, 0);
        if (level & SC_FOLDLEVELHEADERFLAG) {
            ::SendMessage(hwnd, SCI_TOGGLEFOLD, dispLine, 0);
            // --- remove double‑click word selection & keep scroll ----------
            Sci_Position linePos = (Sci_Position)::SendMessage(
                hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
            ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, linePos, 0);
            ::SendMessage(hwnd, SCI_SETFIRSTVISIBLELINE, firstVisible, 0);

            return 0;
        }

        // 3) Read the raw text of that line
        int lineLen = (int)::SendMessage(hwnd, SCI_LINELENGTH, dispLine, 0);
        std::string raw(lineLen, '\0');
        ::SendMessage(hwnd, SCI_GETLINE, dispLine, (LPARAM)&raw[0]);
        raw.resize(strnlen(raw.c_str(), lineLen));

        // 4) Bail if this line does *not* contain "Line "
        if (classify(raw) != LineKind::HitLine)
            return 0;

        // 5) Count *all* previous "Line " occurrences to get our hitIndex
        int hitIndex = -1;
        for (int i = 0; i <= dispLine; ++i) {
            int len = (int)::SendMessage(hwnd, SCI_LINELENGTH, i, 0);
            std::string buf(len, '\0');
            ::SendMessage(hwnd, SCI_GETLINE, i, (LPARAM)&buf[0]);
            buf.resize(strnlen(buf.c_str(), len));
            if (classify(buf) == LineKind::HitLine)
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
        if (wlen > 0) {
            wpath.resize(wlen - 1);
            ::MultiByteToWideChar(CP_UTF8, 0, hit.fullPathUtf8.c_str(), -1, &wpath[0], wlen);
        }

        // 7) Switch file
        ::SendMessage(nppData._nppHandle, NPPM_SWITCHTOFILE, 0, (LPARAM)wpath.c_str());

        // 8) Jump, select and centre the hit in the visible editor area
        HWND hEd = nppData._scintillaMainHandle;
        Sci_Position targetPos = hit.pos;

        // 8‑a) ensure target is unfolded and visible
        ::SendMessage(hEd, SCI_ENSUREVISIBLE,
            ::SendMessage(hEd, SCI_LINEFROMPOSITION, targetPos, 0), 0);

        // 8‑b) move caret to target position and select the range
        ::SendMessage(hEd, SCI_GOTOPOS, targetPos, 0);
        ::SendMessage(hEd, SCI_SETSEL, targetPos, targetPos + hit.length);

        // 8‑c) re‑query display geometry
        size_t firstVisibleDocLine =
            (size_t)::SendMessage(hEd, SCI_DOCLINEFROMVISIBLE, (WPARAM)::SendMessage(hEd, SCI_GETFIRSTVISIBLELINE, 0, 0), 0);
        size_t linesOnScreen =
            (size_t)::SendMessage(hEd, SCI_LINESONSCREEN, (WPARAM)firstVisibleDocLine, 0);
        if (linesOnScreen == 0) linesOnScreen = 1;

        // 8‑d) compute centred target visible line
        size_t caretLine =
            (size_t)::SendMessage(hEd, SCI_LINEFROMPOSITION, targetPos, 0);
        size_t midDisplay = firstVisibleDocLine + linesOnScreen / 2;
        ptrdiff_t deltaLines = (ptrdiff_t)caretLine - (ptrdiff_t)midDisplay;

        // 8‑e) scroll by delta without further auto‑scroll
        ::SendMessage(hEd, SCI_LINESCROLL, 0, deltaLines);
        ::SendMessage(hEd, SCI_ENSUREVISIBLEENFORCEPOLICY, (WPARAM)::SendMessage(hEd, SCI_LINEFROMPOSITION, targetPos, 0), 0);


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

    case WM_KEYDOWN: {
        if (wp == VK_DELETE) {
            deleteSelectedItems(hwnd);
            return 0;
        }
        if (wp == VK_SPACE || wp == VK_RETURN) {
            Sci_Position pos = ::SendMessage(hwnd, SCI_GETCURRENTPOS, 0, 0);
            int line = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0);
            int level = (int)::SendMessage(hwnd, SCI_GETFOLDLEVEL, line, 0);
            if (level & SC_FOLDLEVELHEADERFLAG) {
                ::SendMessage(hwnd, SCI_TOGGLEFOLD, line, 0);
                return 0;
            }
        }
        break;
    }

    case WM_NCDESTROY:
        s_prevSciProc = nullptr;
        break;

    case WM_CONTEXTMENU: {
        // ignore clicks on Scintilla margins
        if (wp != (WPARAM)hwnd) return 0;

        HMENU hMenu = ::CreatePopupMenu();
        if (!hMenu) return 0;

        // helper: append menu entry with localisation
        auto add = [&](UINT id, const wchar_t* langId, UINT flags = MF_STRING) {
            std::wstring txt = LM.get(langId).empty()
                ? std::wstring(langId)
                : LM.get(langId);
            ::AppendMenuW(hMenu, flags, id, txt.c_str());
            };

        const bool hasSel =
            ::SendMessage(hwnd, SCI_GETSELECTIONSTART, 0, 0) !=
            ::SendMessage(hwnd, SCI_GETSELECTIONEND, 0, 0);

        // -------- exact order & separators (matches screenshot) ----------------

        add(IDM_RD_FOLD_ALL, L"rdmenu_fold_all");
        add(IDM_RD_UNFOLD_ALL, L"rdmenu_unfold_all");
        ::AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
        UINT copyFlags = MF_STRING | (hasSel ? 0 : MF_GRAYED);
        add(IDM_RD_COPY_STD, L"rdmenu_copy_std", copyFlags);
        add(IDM_RD_COPY_LINES, L"rdmenu_copy_lines");
        add(IDM_RD_COPY_PATHS, L"rdmenu_copy_paths");
        add(IDM_RD_SELECT_ALL, L"rdmenu_select_all");
        add(IDM_RD_CLEAR_ALL, L"rdmenu_clear_all");
        ::AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);

        add(IDM_RD_OPEN_PATHS, L"rdmenu_open_paths");
        ::AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);

        add(IDM_RD_TOGGLE_WRAP, L"rdmenu_wrap",
            MF_STRING | (ResultDock::_wrapEnabled ? MF_CHECKED : 0));
        add(IDM_RD_TOGGLE_PURGE, L"rdmenu_purge",
            MF_STRING | (ResultDock::_purgeOnNextSearch ? MF_CHECKED : 0));

        // show the popup
        POINT pt{ LOWORD(lp), HIWORD(lp) };
        ::TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
        ::DestroyMenu(hMenu);
        return 0;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wp))
        {
            // ── fold / unfold ─────────────────────────────
        case IDM_RD_FOLD_ALL:
            ::SendMessage(hwnd, SCI_FOLDALL, SC_FOLDACTION_CONTRACT, 0);
            break;

        case IDM_RD_UNFOLD_ALL:
            ::SendMessage(hwnd, SCI_FOLDALL, SC_FOLDACTION_EXPAND, 0);
            break;

        case IDM_RD_COPY_STD:
            ::SendMessage(hwnd, SCI_COPY, 0, 0);
            break;

            // ── select / clear ────────────────────────────
        case IDM_RD_SELECT_ALL:
            ::SendMessage(hwnd, SCI_SELECTALL, 0, 0);
            break;

        case IDM_RD_CLEAR_ALL:
            ResultDock::instance().clear();
            break;

            // ── clipboard actions ─────────────────────────
        case IDM_RD_COPY_LINES:  copySelectedLines(hwnd);  break;
        case IDM_RD_COPY_PATHS:  copySelectedPaths(hwnd);  break;

            // ── open paths in N++ ─────────────────────────
        case IDM_RD_OPEN_PATHS:  openSelectedPaths(hwnd);  break;

            // ── toggle word‑wrap ──────────────────────────
        case IDM_RD_TOGGLE_WRAP:
            ResultDock::_wrapEnabled = !ResultDock::_wrapEnabled;
            ::SendMessage(hwnd, SCI_SETWRAPMODE,
                ResultDock::_wrapEnabled ? SC_WRAP_WORD : SC_WRAP_NONE, 0);
            break;

            // ── toggle purge‑flag ─────────────────────────
        case IDM_RD_TOGGLE_PURGE:
            ResultDock::_purgeOnNextSearch = !ResultDock::_purgeOnNextSearch;
            break;
        }
        return 0;          // WM_COMMAND handled
    }

    }

    return s_prevSciProc
        ? ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp)
        : ::DefWindowProc(hwnd, msg, wp, lp);
}