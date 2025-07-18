
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
#include <functional>
#include "Encoding.h"
#include <unordered_set>

// --- Public Methods ---

// Make the global NppData object available to this file.
extern NppData nppData;

ResultDock::ResultDock(HINSTANCE hInst)
    : _hInst(hInst)      // member‑initialiser‑list
    , _hSci(nullptr)
    , _hDock(nullptr)
{
    // hier kein weiterer Code nötig – alles andere geschieht in create()
}

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
    if (!_hSci)                // _hSci wird im _create() gesetzt
        create(npp);

    // 2) show again – MUST use the client handle!
    if (_hSci)
        ::SendMessage(npp._nppHandle, NPPM_DMMSHOW, 0,
            reinterpret_cast<LPARAM>(_hSci));
}

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
    rebuildFolding();
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
    rebuildFolding();
    applyTheme();
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

    rebuildFolding();          // refresh fold map incl. new header
}

void ResultDock::prependHits(const std::vector<Hit>& newHits, const std::wstring& text)
{
    /* 1) prepend data objects ---------------------------------------- */
    _hits.insert(_hits.begin(), newHits.begin(), newHits.end());

    if (!_hSci || text.empty())
        return;

    /* 2) convert wide → UTF‑8 ---------------------------------------- */
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

    /* 3) remember old length, insert new block + optional separator -- */
    LRESULT oldLen = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);
    ::SendMessage(_hSci, SCI_INSERTTEXT, 0,
        reinterpret_cast<LPARAM>(utf8PrependedText.c_str()));
    if (oldLen > 0)                                    /* ★ CHANGED ★ */
        ::SendMessage(_hSci, SCI_INSERTTEXT, utf8PrependedText.length(),
            reinterpret_cast<LPARAM>("\r\n"));

    ::SendMessage(_hSci, SCI_ENDUNDOACTION, 0, 0);
    ::SendMessage(_hSci, SCI_SETREADONLY, TRUE, 0);

    rebuildFolding();
    applyTheme();
    applyStyling();

    ::SendMessage(_hSci, SCI_SETFIRSTVISIBLELINE, 0, 0);
    ::SendMessage(_hSci, SCI_GOTOPOS, 0, 0);
}

void ResultDock::recordHit(const std::string& fullPathUtf8, Sci_Position pos, Sci_Position length) {
    _hits.push_back({ fullPathUtf8, pos, length });
}

void ResultDock::clear() {
    _hits.clear();
    if (_hSci) {
        ::SendMessage(_hSci, SCI_CLEARALL, 0, 0);
    }
}

int ResultDock::leadingSpaces(const char* line, int len)
{
    int s = 0;
    while (s < len && line[s] == ' ') ++s;
    return s;
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

    // Pass 2: detect headers via leading spaces & keywords
    auto isSearchHeader = [](const std::string& s) -> bool
        {
            return s.rfind("Search ", 0) == 0; // starts with "Search "
        };

    for (int l = 0; l < lineCount; ++l)
    {
        const int rawLen = static_cast<int>(sciSend(SCI_LINELENGTH, l, 0));
        if (rawLen <= 1) continue; // blank line

        std::string buf(rawLen + 1, '\0');
        sciSend(SCI_GETLINE, l, (LPARAM)buf.data());
        buf.erase(std::find_if(buf.begin(), buf.end(), [](char c)
            { return c == '\r' || c == '\n' || c == '\0'; }), buf.end());

        // Count leading spaces
        int spaces = 0;
        while (spaces < (int)buf.size() && buf[spaces] == ' ') ++spaces;

        // Remove indent for easier tests
        const std::string_view trimmed(buf.data() + spaces, buf.size() - spaces);

        bool header = false;
        int  level = BASE;

        if (spaces == 0 && (trimmed.rfind("Search in List", 0) == 0 || isSearchHeader(std::string(trimmed))))
        {
            header = true;
            level = BASE;               // 0
        }
        else if (spaces == 4)
        {
            if (isSearchHeader(std::string(trimmed))) { header = true; level = BASE + 2; } // Old list entry (rare)
            else { header = true; level = BASE + 1; } // File path
        }
        else if (spaces == 8 && isSearchHeader(std::string(trimmed)))
        {
            header = true;
            level = BASE + 2;           // Criteria
        }
        else {
            // Content line → level = BASE + 3 (or deeper) handled later
        }

        if (header)
        {
            sciSend(SCI_SETFOLDLEVEL, l, level | SC_FOLDLEVELHEADERFLAG);
        }
        else
        {
            int contentLevel = BASE + (spaces / 4);
            if (contentLevel < BASE) contentLevel = BASE;
            sciSend(SCI_SETFOLDLEVEL, l, contentLevel);
        }
    }
}

void ResultDock::applyStyling() const
{
    if (!_hSci) return;
    auto S = [this](UINT msg, WPARAM w = 0, LPARAM l = 0) {
        return ::SendMessage(_hSci, msg, w, l);
        };

    // Clear all three indicators over the entire doc
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
    S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));

    // 1) Fill full‑line background on each hit line
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    for (const auto& hit : _hits) {
        // compute end of line
        int dispLine = (int)S(SCI_LINEFROMPOSITION, hit.displayLineStart, 0);
        int lineEnd = (int)S(SCI_GETLINEENDPOSITION, dispLine, 0);
        int length = lineEnd - hit.displayLineStart;
        if (length > 0)
            S(SCI_INDICATORFILLRANGE, hit.displayLineStart, length);
    }

    // 2) Fill digits of the line number
    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& hit : _hits) {
        S(SCI_INDICATORFILLRANGE,
            hit.displayLineStart + hit.numberStart,
            hit.numberLen);
    }

    // 3) Fill each match substring
    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
    for (const auto& hit : _hits) {
        for (size_t i = 0; i < hit.matchStarts.size(); ++i) {
            S(SCI_INDICATORFILLRANGE,
                hit.displayLineStart + hit.matchStarts[i],
                hit.matchLens[i]);
        }
    }
}

void ResultDock::onThemeChanged()
{
    applyTheme();
}

void ResultDock::formatHitsForFile(const std::wstring& wFilePath,
    const SciSendFn& sciSend,
    std::vector<Hit>& hitsInOut,
    std::wstring& outBlock,
    size_t& ioUtf8Len) const
{
    const int hitsHere = static_cast<int>(hitsInOut.size());
    if (hitsHere == 0) return;

    /* 1) file header -------------------------------------------------- */
    std::wstring fileHdr = L"    " + wFilePath                // ★ 4 blanks
        + L" (" + std::to_wstring(hitsHere)
        + L" hits)\r\n";
    outBlock += fileHdr;
    ioUtf8Len += Encoding::wstringToUtf8(fileHdr).size();

    /* 2) loop over matches ------------------------------------------- */
    int  lastPrintedLine = -1;
    Hit* lastHit = nullptr;

    for (auto& hit : hitsInOut)
    {
        const int lineN = static_cast<int>(
            sciSend(SCI_LINEFROMPOSITION,
                static_cast<WPARAM>(hit.pos), 0)) + 1;

        const Sci_Position lineStartPos =
            sciSend(SCI_POSITIONFROMLINE,
                static_cast<WPARAM>(lineN - 1), 0);

        LRESULT rawLen = sciSend(SCI_LINELENGTH,
            static_cast<WPARAM>(lineN - 1), 0);

        std::string raw(static_cast<size_t>(rawLen), '\0');
        sciSend(SCI_GETLINE,
            static_cast<WPARAM>(lineN - 1),
            reinterpret_cast<LPARAM>(raw.data()));
        raw.resize(strnlen(raw.c_str(), rawLen));

        size_t lead = raw.find_first_not_of(" \t");
        if (lead == std::string::npos) lead = 0;
        std::string trim = raw.substr(lead);
        while (!trim.empty() &&
            (trim.back() == '\r' || trim.back() == '\n'))
            trim.pop_back();

        const Sci_Position absLineStart =
            lineStartPos + static_cast<Sci_Position>(lead);

        /* merge duplicates ------------------------------------------ */
        if (lineN == lastPrintedLine && lastHit)
        {
            int rel = static_cast<int>(hit.pos - absLineStart);
            lastHit->matchStarts.push_back(
                lastHit->numberStart + lastHit->numberLen + 2 + rel);
            lastHit->matchLens.push_back(static_cast<int>(hit.length));

            hit.matchStarts.clear();
            hit.displayLineStart = -1;
            continue;
        }
        lastPrintedLine = lineN;
        lastHit = &hit;

        hit.displayLineStart = static_cast<int>(ioUtf8Len);
        hit.numberStart =
            static_cast<int>(std::string("            Line ").size());
        hit.numberLen =
            static_cast<int>(std::to_string(lineN).size());

        int rel = static_cast<int>(hit.pos - absLineStart);
        hit.matchStarts = { hit.numberStart + hit.numberLen + 2 + rel };
        hit.matchLens = { static_cast<int>(hit.length) };

        std::wstring prefixW = L"            Line " +
            std::to_wstring(lineN) + L": ";
        outBlock += prefixW + Encoding::stringToWString(trim) + L"\r\n";
        ioUtf8Len += Encoding::wstringToUtf8(prefixW).size()
            + trim.size() + 2;
    }

    /* 3) remove hidden duplicates ----------------------------------- */
    hitsInOut.erase(
        std::remove_if(hitsInOut.begin(), hitsInOut.end(),
            [](const Hit& h) { return h.displayLineStart == -1; }),
        hitsInOut.end());
}


// --- Private Methods ---

static constexpr uint32_t argb(BYTE a, COLORREF c)
{
    return (uint32_t(a) << 24) |
        (uint32_t(GetRValue(c)) << 16) |
        (uint32_t(GetGValue(c)) << 8) |
        uint32_t(GetBValue(c));
}

void ResultDock::create(const NppData& npp)
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

    initFolding();

    // Set the initial styles to match the current N++ theme upon creation.
    applyTheme();
}

void ResultDock::initFolding() const
{
    if (!_hSci) return;

    auto S = [this](UINT msg, WPARAM w = 0, LPARAM l = 0)
        { return ::SendMessage(_hSci, msg, w, l); };

    // 1) Configure margin #2 as the folding symbol margin
    constexpr int M_FOLD = 2;
    S(SCI_SETMARGINTYPEN, M_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN, M_FOLD, SC_MASK_FOLDERS);

    // 2) Make the margin wide enough for a 16‑px box plus the vertical stem
    LRESULT lr = ::SendMessage(_hSci, SCI_TEXTHEIGHT, 0, 0);
    int h = static_cast<int>(lr);
    int w = h + 4;               // add 4 px so the connected stem is never clipped
    S(SCI_SETMARGINWIDTHN, M_FOLD, w);

    // Enable mouse interaction and automatic folding updates
    S(SCI_SETMARGINSENSITIVEN, M_FOLD, TRUE);
    S(SCI_SETAUTOMATICFOLD,
        SC_AUTOMATICFOLD_SHOW |
        SC_AUTOMATICFOLD_CLICK |
        SC_AUTOMATICFOLD_CHANGE, 0);

    // 3) Hide the first two margins (line numbers / bookmarks)
    S(SCI_SETMARGINWIDTHN, 0, 0);
    S(SCI_SETMARGINWIDTHN, 1, 0);

    // 4) Define header markers (plus/minus boxes, connected style)
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUSCONNECTED);    // closed block
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUSCONNECTED);   // open block
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED);   // open (nested)
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED);    // closed (last child)

    // 5) Define guide‑line and corner markers
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE);    // │ vertical guide
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);  // ├ tee junction
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER);  // └ corner
    // Do NOT redefine SC_MARKNUM_FOLDEREND here – keeps the “+” glyph intact

    // 6) Apply a neutral placeholder colour; themed colours are set later in _applyTheme()
    const COLORREF placeholder = RGB(200, 200, 200);
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

void ResultDock::applyTheme()
{
    if (!_hSci) return;

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { ::SendMessage(_hSci, m, w, l); };

    /* 0) Base colours ---------------------------------------------------- */
    const bool dark = ::SendMessage(
        nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0;

    COLORREF bg = (COLORREF)::SendMessage(
        nppData._nppHandle, NPPM_GETEDITORDEFAULTBACKGROUNDCOLOR, 0, 0);
    COLORREF fg = (COLORREF)::SendMessage(
        nppData._nppHandle, NPPM_GETEDITORDEFAULTFOREGROUNDCOLOR, 0, 0);

    COLORREF lnBg = dark ? RGB(0, 0, 0) : RGB(255, 255, 255);   // margin fill
    COLORREF lnFg = dark ? RGB(200, 200, 200) : RGB(80, 80, 80);   // margin text

    COLORREF selBg = dark ? RGB(96, 96, 96)
        : ::GetSysColor(COLOR_HIGHLIGHT);
    COLORREF selFg = dark ? RGB(255, 255, 255)
        : ::GetSysColor(COLOR_HIGHLIGHTTEXT);

    /* 1) Reset Scintilla styles ----------------------------------------- */
    S(SCI_STYLESETBACK, STYLE_DEFAULT, bg);
    S(SCI_STYLESETFORE, STYLE_DEFAULT, fg);
    S(SCI_STYLECLEARALL);

    /* 2) Margin colours -------------------------------------------------- */
    S(SCI_SETMARGINBACKN, 0, lnBg);
    S(SCI_STYLESETBACK, STYLE_LINENUMBER, lnBg);
    S(SCI_STYLESETFORE, STYLE_LINENUMBER, lnFg);

    S(SCI_SETMARGINBACKN, 1, bg);     // spacer margin
    S(SCI_SETMARGINBACKN, 2, lnBg);   // fold margin

    S(SCI_SETFOLDMARGINCOLOUR, 1, lnBg);
    S(SCI_SETFOLDMARGINHICOLOUR, 1, lnBg);

    /* 3) Selection / caret‑line ----------------------------------------- */
    S(SCI_SETSELFORE, 1, selFg);
    S(SCI_SETSELBACK, 1, selBg);
    S(SCI_SETSELALPHA, 256, 0);

    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_BACK, argb(0xFF, selBg));
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_TEXT, argb(0xFF, selFg));

    S(SCI_SETADDITIONALSELFORE, selFg);
    S(SCI_SETADDITIONALSELBACK, selBg);
    S(SCI_SETADDITIONALSELALPHA, 256, 0);

    /* 4) Fold‑marker colours (dark theme → black box, white glyph) -------- */
    const COLORREF markerFill = lnBg;                        // box outline / frame
    const COLORREF markerGlyph = dark ? RGB(255, 255, 255)      // glyph + guide lines
        : RGB(80, 80, 80);    // light theme: dark grey

    // Header markers: +  −
    for (int id : { SC_MARKNUM_FOLDER,          // +
        SC_MARKNUM_FOLDEREND,       // + (last child)
        SC_MARKNUM_FOLDEROPEN,      // −
        SC_MARKNUM_FOLDEROPENMID }) // − (nested)
    {
        S(SCI_MARKERSETBACK, id, markerGlyph);  // glyph colour  (plus/minus)
        S(SCI_MARKERSETFORE, id, markerFill);   // box outline   (frame)
    }

    // Guide‑line markers: │ ├ └   → durchgehender Strich
    for (int id : { SC_MARKNUM_FOLDERSUB,
        SC_MARKNUM_FOLDERMIDTAIL,
        SC_MARKNUM_FOLDERTAIL })
    {
        S(SCI_MARKERSETBACK, id, markerGlyph);  // solid line
        S(SCI_MARKERSETFORE, id, markerGlyph);
    }

    /* 5) Caret‑line highlight (indicator 0) ----------------------------- */
    S(SCI_INDICSETSTYLE, 0, INDIC_ROUNDBOX);
    S(SCI_INDICSETFORE, 0, selBg);
    S(SCI_INDICSETUNDER, 0, TRUE);
    S(SCI_INDICSETALPHA, 0, 128);

    /* 6) Custom indicators ---------------------------------------------- */
    COLORREF lineBgColor = dark ? RGB(0x3A, 0x3D, 0x33) : RGB(0xE7, 0xF2, 0xFF);
    COLORREF lineNumberClr = dark ? RGB(0xAE, 0x81, 0xFF) : RGB(0xFD, 0x97, 0x1F);
    COLORREF matchClr = dark ? RGB(0xE6, 0xDB, 0x74) : RGB(0xFF, 0x00, 0x00);

    S(SCI_INDICSETSTYLE, INDIC_LINE_BACKGROUND, INDIC_STRAIGHTBOX);
    S(SCI_INDICSETFORE, INDIC_LINE_BACKGROUND, lineBgColor);
    S(SCI_INDICSETALPHA, INDIC_LINE_BACKGROUND, 100);
    S(SCI_INDICSETUNDER, INDIC_LINE_BACKGROUND, TRUE);

    S(SCI_INDICSETSTYLE, INDIC_LINENUMBER_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_LINENUMBER_FORE, lineNumberClr);

    S(SCI_INDICSETSTYLE, INDIC_MATCH_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_MATCH_FORE, matchClr);

    /* 7) Caret‑line visibility ----------------------------------------- */
    S(SCI_SETCARETLINEVISIBLE, 1, 0);
    S(SCI_SETCARETLINEBACK, selBg, 0);
}

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
        ::SendMessage(nppData._nppHandle,
            NPPM_SWITCHTOFILE, 0,
            (LPARAM)wpath.c_str());

        // 8) Jump + select in the editor
        HWND hEd = nppData._scintillaMainHandle;
        int targetLine = (int)::SendMessage(hEd, SCI_LINEFROMPOSITION, hit.pos, 0);
        ::SendMessage(hEd, SCI_GOTOLINE, targetLine, 0);
        ::SendMessage(hEd, SCI_ENSUREVISIBLEENFORCEPOLICY, targetLine, 0);
        ::SendMessage(hEd, SCI_SETSEL, hit.pos, hit.pos + hit.length);
        // no SetFocus(hEd)

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

void ResultDock::formatHitsLines(
    const SciSendFn& sciSend,
    std::vector<Hit>& hits,
    std::wstring& out,
    size_t& utf8Pos) const
{
    constexpr wchar_t INDENT_W[] = L"            "; // 12 spaces
    constexpr char   INDENT_A[] = "            "; // 12 spaces (ASCII)
    const size_t indentUtf8Len = sizeof(INDENT_A) - 1;            // 12

    int  lastPrintedLine = -1;
    Hit* lastHit = nullptr;

    for (auto& h : hits)
    {
        const int lineZero = static_cast<int>(
            sciSend(SCI_LINEFROMPOSITION, (WPARAM)h.pos, 0));
        const int lineOne = lineZero + 1;

        // ─── Merge multiple hits on the same physical line ────────────────
        if (lineZero == lastPrintedLine && lastHit)
        {
            Sci_Position lineStartPos = sciSend(SCI_POSITIONFROMLINE, lineZero, 0);
            int rel = static_cast<int>(h.pos - lineStartPos);
            lastHit->matchStarts.push_back(
                lastHit->numberStart + lastHit->numberLen + 2 + rel);
            lastHit->matchLens.push_back(static_cast<int>(h.length));
            h.displayLineStart = -1; // placeholder, not styled
            continue;
        }
        lastPrintedLine = lineZero;
        lastHit = &h;

        // ─── Read line text (UTF‑8) and trim leading WS ─────────────────––
        const int bufLen = static_cast<int>(
            sciSend(SCI_LINELENGTH, (WPARAM)lineZero, 0)) + 1; // incl. NUL
        std::string utf8(bufLen, '\0');
        sciSend(SCI_GETLINE, (WPARAM)lineZero, (LPARAM)utf8.data());

        utf8.erase(std::find_if(utf8.begin(), utf8.end(), [](char c)
            { return c == '\r' || c == '\n' || c == '\0'; }), utf8.end());

        size_t lead = utf8.find_first_not_of(" \t");
        if (lead == std::string::npos) lead = 0;
        std::string trim = utf8.substr(lead);
        Sci_Position absLineStart = sciSend(SCI_POSITIONFROMLINE, lineZero, 0) + static_cast<Sci_Position>(lead);

        // ─── Record styling offsets for this primary hit ─────────────────––
        h.displayLineStart = static_cast<int>(utf8Pos);
        h.numberStart = static_cast<int>(indentUtf8Len + 5);   // after "Line "
        h.numberLen = static_cast<int>(std::to_string(lineOne).size());
        int rel = static_cast<int>(h.pos - absLineStart);
        h.matchStarts = { h.numberStart + h.numberLen + 2 + rel };
        h.matchLens = { static_cast<int>(h.length) };

        // ─── Compose output line ─────────────────────────────────────────––
        std::wstring prefix = INDENT_W;
        prefix += L"Line " + std::to_wstring(lineOne) + L": ";

        out += prefix + Encoding::stringToWString(trim) + L"\r\n";
        utf8Pos += Encoding::wstringToUtf8(prefix).size() + trim.size() + 2; // +CRLF
    }

    // Remove placeholder hits created by duplicate‑merge
    hits.erase(std::remove_if(hits.begin(), hits.end(), [](const Hit& h)
        { return h.displayLineStart < 0; }), hits.end());
}



