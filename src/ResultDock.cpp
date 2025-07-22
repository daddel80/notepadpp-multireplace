
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

#ifndef SC_LINEACTION_JOIN
#define SC_LINEACTION_JOIN 0x1000000
#define SC_LINEACTION_JOIN 0x1000000
#endif

extern NppData nppData;

ResultDock::ResultDock(HINSTANCE hInst)
    : _hInst(hInst)      // member‑initialiser‑list
    , _hSci(nullptr)
    , _hDock(nullptr)
{
    // no further code needed here – everything else happens in create()
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
    if (!_hSci)                // _hSci is initialized in create()
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

    // Add a separating newline if the dock is not empty
    LRESULT currentLength = ::SendMessage(_hSci, SCI_GETLENGTH, 0, 0);
    if (currentLength > 0) {
        // Prepend a newline to separate this block from the previous one.
        utf8Text.insert(0, "\r\n");
    }

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

void ResultDock::onThemeChanged()
{
    applyTheme();
}

void ResultDock::formatHitsForFile(const std::wstring& wPath,
    const SciSendFn& sciSend,
    std::vector<Hit>& hits,
    std::wstring& out,
    size_t& utf8Pos) const
{
    std::wstring hdr = L"    " + wPath +
        L" (" + std::to_wstring(hits.size()) + L" hits)\r\n";
    out += hdr;
    utf8Pos += Encoding::wstringToUtf8(hdr).size();

    // per‑line formatting (handles encoding, styling offsets, merging)
    formatHitsLines(sciSend, hits, out, utf8Pos);
}

void ResultDock::formatHitsLines(const SciSendFn& sciSend,
    std::vector<Hit>& hits,
    std::wstring& out,
    size_t& utf8Pos) const
{
    constexpr wchar_t INDENT_W[] = L"            ";
    constexpr size_t  INDENT_U8 = 12;

    const UINT docCp = static_cast<UINT>(sciSend(SCI_GETCODEPAGE, 0, 0));

    int  prevLine = -1;
    Hit* firstHitOnLine = nullptr;

    for (Hit& h : hits)
    {
        int lineZero = (int)sciSend(SCI_LINEFROMPOSITION, h.pos, 0);
        int lineOne = lineZero + 1;

        // --- read raw bytes ------------------------------------
        int rawLen = (int)sciSend(SCI_LINELENGTH, lineZero, 0);
        std::string raw(rawLen, '\0');
        sciSend(SCI_GETLINE, lineZero, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), rawLen));
        while (!raw.empty() && (raw.back() == '\r' || raw.back() == '\n'))
            raw.pop_back();

        // --- trim leading blanks -------------------------------
        size_t lead = raw.find_first_not_of(" \t");
        std::string_view sliceBytes =
            (lead == std::string::npos) ? std::string_view{}
        : std::string_view(raw).substr(lead);

        // --- FULL slice as UTF‑8 (★) ---------------------------
        const std::string sliceU8 = Encoding::bytesToUtf8(sliceBytes, docCp);
        const std::wstring sliceW = Encoding::utf8ToWString(sliceU8);

        // --- prefix --------------------------------------------
        std::wstring prefixW = std::wstring(INDENT_W)
            + L"Line " + std::to_wstring(lineOne) + L": ";
        size_t prefixU8Len = Encoding::wstringToUtf8(prefixW).size();

        // --- byte offset of hit inside slice -------------------
        Sci_Position absTrimmed = sciSend(SCI_POSITIONFROMLINE, lineZero, 0) + (Sci_Position)lead;
        size_t relBytes = (size_t)(h.pos - absTrimmed);

        // --- locate hit in UTF‑8 text  (★) ---------------------
        size_t hitStartInSlice = Encoding::bytesToUtf8(
            sliceBytes.substr(0, relBytes), docCp).size();
        size_t hitLenU8 = Encoding::bytesToUtf8(
            sliceBytes.substr(relBytes, h.length), docCp).size();

        // --- first hit on this visual line? --------------------
        if (lineZero != prevLine)
        {
            out += prefixW + sliceW + L"\r\n";

            h.displayLineStart = (int)utf8Pos;
            h.numberStart = (int)(INDENT_U8 + 5);
            h.numberLen = (int)std::to_string(lineOne).size();
            h.matchStarts = { (int)(prefixU8Len + hitStartInSlice) };
            h.matchLens = { (int)hitLenU8 };

            utf8Pos += prefixU8Len + sliceU8.size() + 2;
            firstHitOnLine = &h;
            prevLine = lineZero;
        }
        else // additional hit on same line
        {
            assert(firstHitOnLine != nullptr);
            firstHitOnLine->matchStarts.push_back((int)(prefixU8Len + hitStartInSlice));
            firstHitOnLine->matchLens.push_back((int)hitLenU8);
            h.displayLineStart = -1;       // dummy
        }
    }

    hits.erase(std::remove_if(hits.begin(), hits.end(),
        [](const Hit& h) { return h.displayLineStart < 0; }),
        hits.end());
}

void ResultDock::buildListText(
    const FileMap& files,
    bool flatView,
    const std::wstring& header,
    const SciSendFn& sciSend,
    std::wstring& outText,
    std::vector<Hit>& outHits) const
{
    size_t utf8Len = Encoding::wstringToUtf8(header).size();
    std::wstring body;

    for (auto& [path, f] : files)
    {
        std::wstring fileHdr = L"    " + f.wPath + L" (" +
            std::to_wstring(f.hitCount) + L" hits)\r\n";
        body += fileHdr;
        utf8Len += Encoding::wstringToUtf8(fileHdr).size();

        if (flatView)
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
        else
        {
            // grouped: first the file header, then each criterion block
            for (auto& c : f.crits)
            {
                std::wstring critHdr = L"        Search \"" + c.text + L"\" (" +
                    std::to_wstring(c.hits.size()) + L" hits)\r\n";
                body += critHdr;
                utf8Len += Encoding::wstringToUtf8(critHdr).size();

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
    }

    outText = header + body;
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

    if (!_hSci)
    {
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
    ::ZeroMemory(&_dockData, sizeof(_dockData));
    _dockData.hClient = _hSci;
    _dockData.pszName = L"MultiReplace – Search results";
    _dockData.dlgID = IDD_MULTIREPLACE_RESULT_DOCK;
    _dockData.uMask = DWS_DF_CONT_BOTTOM | DWS_ICONTAB;
    _dockData.hIconTab = nullptr;
    _dockData.pszAddInfo = L"";
    _dockData.pszModuleName = NPP_PLUGIN_NAME;
    _dockData.iPrevCont = -1;                        // no previous container
    _dockData.rcFloat = { 0, 0, 0, 0 };    // sensible default when undocked

    // 6) Register the dock window with Notepad++
    _hDock = reinterpret_cast<HWND>(
        ::SendMessageW(npp._nppHandle,
            NPPM_DMMREGASDCKDLG,
            0,
            reinterpret_cast<LPARAM>(&_dockData)));

    if (!_hDock)
    {
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

    // 8) Initialise folding and apply syntax colours that match current N++ theme
    initFolding();
    applyTheme();
}

void ResultDock::initFolding() const
{
    if (!_hSci) return;

    auto S = [this](UINT msg, WPARAM w = 0, LPARAM l = 0)
        { return ::SendMessage(_hSci, msg, w, l); };

    /* 1) configure margin #2 for folding symbols ----------------- */
    constexpr int M_FOLD = 2;
    S(SCI_SETMARGINTYPEN, M_FOLD, SC_MARGIN_SYMBOL);
    S(SCI_SETMARGINMASKN, M_FOLD, SC_MASK_FOLDERS);

    /* 2) width: 16‑px box + 4 px stem so nothing is clipped ------ */
    const int h = static_cast<int>(S(SCI_TEXTHEIGHT));
    S(SCI_SETMARGINWIDTHN, M_FOLD, h + 4);

    /* 3) hide margins 0/1 (line‑numbers & bookmarks) ------------- */
    S(SCI_SETMARGINWIDTHN, 0, 0);
    S(SCI_SETMARGINWIDTHN, 1, 0);

    /* -------------------------------------------------------------
     * 4) header markers
     *    – root header boxes:  BOXPLUS / BOXMINUS          (no top stem)
     *    – nested headers:     BOXPLUSCONNECTED / …MINUS…  (with stem)
     * ---------------------------------------------------------- */
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDER, SC_MARK_BOXPLUS);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPEN, SC_MARK_BOXMINUS);

    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEROPENMID, SC_MARK_BOXMINUSCONNECTED);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDEREND, SC_MARK_BOXPLUSCONNECTED);

    /* 5) guide‑line markers (│ ├ └) ------------------------------ */
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERSUB, SC_MARK_VLINE);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERMIDTAIL, SC_MARK_TCORNER);
    S(SCI_MARKERDEFINE, SC_MARKNUM_FOLDERTAIL, SC_MARK_LCORNER);

    /* 6) placeholder colour – themed colours set in applyTheme() */
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

    /* 7) enable mouse interaction & auto‑fold updates ------------ */
    S(SCI_SETMARGINSENSITIVEN, M_FOLD, TRUE);
    S(SCI_SETAUTOMATICFOLD,
        SC_AUTOMATICFOLD_SHOW |
        SC_AUTOMATICFOLD_CLICK |
        SC_AUTOMATICFOLD_CHANGE);

    S(SCI_SETFOLDFLAGS, SC_FOLDFLAG_LINEAFTER_CONTRACTED);

    S(SCI_MARKERENABLEHIGHLIGHT, TRUE);
}

void ResultDock::applyTheme()
{
    if (!_hSci)
        return;

    // Helper: send message to the Scintilla dock
    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT
        { return ::SendMessage(_hSci, m, w, l); };

    // 0. Retrieve base editor colors from Notepad++
    const bool dark = ::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0;

    COLORREF editorBg = (COLORREF)::SendMessage(nppData._nppHandle, NPPM_GETEDITORDEFAULTBACKGROUNDCOLOR, 0, 0);
    COLORREF editorFg = (COLORREF)::SendMessage(nppData._nppHandle, NPPM_GETEDITORDEFAULTFOREGROUNDCOLOR, 0, 0);

    // 1. Reset styles and set global font
    S(SCI_STYLESETBACK, STYLE_DEFAULT, editorBg);
    S(SCI_STYLESETFORE, STYLE_DEFAULT, editorFg);
    S(SCI_STYLECLEARALL);
    S(SCI_STYLESETFONT, STYLE_DEFAULT, (LPARAM)"Consolas");
    S(SCI_STYLESETSIZE, STYLE_DEFAULT, 10);

    // 2. Configure margins (0=line number, 1=symbol, 2=fold)
    COLORREF marginBg = dark ? RGB(0, 0, 0) : editorBg;
    COLORREF marginFg = dark ? RGB(200, 200, 200) : RGB(80, 80, 80);

    for (int m = 0; m <= 2; ++m)
        S(SCI_SETMARGINBACKN, m, marginBg);

    S(SCI_STYLESETBACK, STYLE_LINENUMBER, marginBg);
    S(SCI_STYLESETFORE, STYLE_LINENUMBER, marginFg);

    S(SCI_SETFOLDMARGINCOLOUR, TRUE, marginBg);
    S(SCI_SETFOLDMARGINHICOLOUR, TRUE, marginBg);

    // 3. Selection colors
    COLORREF selBg = dark ? RGB(96, 96, 96) : RGB(0xE0, 0xE0, 0xE0);
    COLORREF selFg = dark ? RGB(255, 255, 255) : editorFg;

    S(SCI_SETSELFORE, TRUE, selFg);
    S(SCI_SETSELBACK, TRUE, selBg);
    S(SCI_SETSELALPHA, 256, 0);

    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_BACK, argb(0xFF, selBg));
    S(SCI_SETELEMENTCOLOUR, SC_ELEMENT_SELECTION_INACTIVE_TEXT, argb(0xFF, selFg));

    S(SCI_SETADDITIONALSELFORE, selFg);
    S(SCI_SETADDITIONALSELBACK, selBg);
    S(SCI_SETADDITIONALSELALPHA, 256, 0);

    // 4. Fold marker colors
    const COLORREF markerGlyph = dark ? RDColors::FoldGlyphDark : RDColors::FoldGlyphLight;

    for (int id : {SC_MARKNUM_FOLDER, SC_MARKNUM_FOLDEREND, SC_MARKNUM_FOLDEROPEN,
        SC_MARKNUM_FOLDEROPENMID, SC_MARKNUM_FOLDERSUB,
        SC_MARKNUM_FOLDERMIDTAIL, SC_MARKNUM_FOLDERTAIL})
    {
        S(SCI_MARKERSETBACK, id, markerGlyph);
        S(SCI_MARKERSETFORE, id, marginBg);
        S(SCI_MARKERSETBACKSELECTED, id, dark ? RDColors::FoldHiDark : RDColors::FoldHiLight);
    }

    // 5. Caret line indicator (indicator 0)
    S(SCI_INDICSETSTYLE, 0, INDIC_ROUNDBOX);
    S(SCI_INDICSETFORE, 0, selBg);
    S(SCI_INDICSETALPHA, 0, dark ? RDColors::CaretLineAlphaDark : RDColors::CaretLineAlphaLight);
    S(SCI_INDICSETUNDER, 0, TRUE);

    S(SCI_SETCARETLINEVISIBLE, TRUE, 0);
    S(SCI_SETCARETLINEBACK, dark ? selBg : RDColors::CaretLineBackLight, 0);
    S(SCI_SETCARETLINEBACKALPHA, dark ? RDColors::CaretLineAlphaDark : RDColors::CaretLineAlphaLight);

    // 6. Custom indicators and styles
    COLORREF hitLineBg = dark ? RDColors::LineBgDark : RDColors::LineBgLight;
    COLORREF lineNrFg = dark ? RDColors::LineNrDark : RDColors::LineNrLight;
    COLORREF matchFg = dark ? RDColors::MatchDark : RDColors::MatchLight;
    COLORREF matchBg = RDColors::MatchBgLight;
    COLORREF headerBg = dark ? RDColors::HeaderBgDark : RDColors::HeaderBgLight;
    COLORREF filePathFg = dark ? RDColors::FilePathFgDark : RDColors::FilePathFgLight;

    // 6-a Hit line background
    S(SCI_INDICSETSTYLE, INDIC_LINE_BACKGROUND, INDIC_STRAIGHTBOX);
    S(SCI_INDICSETFORE, INDIC_LINE_BACKGROUND, hitLineBg);
    S(SCI_INDICSETALPHA, INDIC_LINE_BACKGROUND, 100);
    S(SCI_INDICSETUNDER, INDIC_LINE_BACKGROUND, TRUE);

    // 6-b Line number color
    S(SCI_INDICSETSTYLE, INDIC_LINENUMBER_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_LINENUMBER_FORE, lineNrFg);

    // 6-c Match indicators
    S(SCI_INDICSETSTYLE, INDIC_MATCH_BG, INDIC_FULLBOX);
    S(SCI_INDICSETFORE, INDIC_MATCH_BG, matchBg);
    S(SCI_INDICSETALPHA, INDIC_MATCH_BG, dark ? 0 : 255);
    S(SCI_INDICSETUNDER, INDIC_MATCH_BG, TRUE);

    S(SCI_INDICSETSTYLE, INDIC_MATCH_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_MATCH_FORE, matchFg);

    // 6-d Header line style
    S(SCI_STYLESETFORE, STYLE_HEADER, RDColors::HeaderFg);
    S(SCI_STYLESETBACK, STYLE_HEADER, headerBg);
    S(SCI_STYLESETBOLD, STYLE_HEADER, TRUE);
    S(SCI_STYLESETEOLFILLED, STYLE_HEADER, TRUE);

    // 6-e File path style
    S(SCI_STYLESETFORE, STYLE_FILEPATH, filePathFg);
    S(SCI_STYLESETBACK, STYLE_FILEPATH, -1);
    S(SCI_STYLESETBOLD, STYLE_FILEPATH, TRUE);
    S(SCI_STYLESETITALIC, STYLE_FILEPATH, TRUE);
    S(SCI_STYLESETEOLFILLED, STYLE_FILEPATH, TRUE);
}


void ResultDock::applyStyling() const
{
    if (!_hSci) return;

    auto S = [this](UINT m, WPARAM w = 0, LPARAM l = 0)
        { return ::SendMessage(_hSci, m, w, l); };


    // 0)  Clear all previous styling indicators
    for (int ind : { INDIC_LINE_BACKGROUND,
        INDIC_LINENUMBER_FORE,
        INDIC_MATCH_FORE,
        INDIC_MATCH_BG}) // Note: Header/FilePath indicators are no longer needed here
    {
        S(SCI_SETINDICATORCURRENT, ind);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }



    // 1)  Apply base style for each line (Header, File Path, or Default)
    // This sets the font (bold/regular) and base colors.
    S(SCI_STARTSTYLING, 0, 0);

    const int lineCount = static_cast<int>(S(SCI_GETLINECOUNT));
    for (int ln = 0; ln < lineCount; ++ln)
    {
        const Sci_Position lineStartPos = S(SCI_POSITIONFROMLINE, ln, 0);
        const int lineRawLen = static_cast<int>(S(SCI_LINELENGTH, ln, 0));

        // Determine the style for the current line
        int lineStyle = STYLE_DEFAULT;
        if (lineRawLen > 0)
        {
            std::string buf(lineRawLen, '\0');
            S(SCI_GETLINE, ln, (LPARAM)buf.data());

            size_t lead = buf.find_first_not_of(' ');
            if (lead == std::string::npos) lead = 0; // line is all spaces

            std::string_view trimmed(buf.data() + lead, buf.size() - lead);

            if (trimmed.rfind("Search ", 0) == 0)
            {
                lineStyle = STYLE_HEADER;
            }
            // A file path has exactly 4 leading spaces and is not empty.
            else if (lead == 4 && !trimmed.empty())
            {
                lineStyle = STYLE_FILEPATH;
            }
        }

        // Apply styling for the line content
        if (lineRawLen > 0) {
            S(SCI_SETSTYLING, lineRawLen, lineStyle);
        }

        // Apply default style to the EOL characters (\r\n)
        const Sci_Position lineEndPos = S(SCI_GETLINEENDPOSITION, ln, 0);
        const int eolLen = static_cast<int>(lineEndPos - (lineStartPos + lineRawLen));
        if (eolLen > 0) {
            S(SCI_SETSTYLING, eolLen, STYLE_DEFAULT);
        }
    }

    // 2)  Apply overlay indicators for hit details.
    // These are drawn on top of the base styles set above.

    // 2-a) Full-line background for each hit
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    for (const auto& h : _hits)
    {
        if (h.displayLineStart < 0) continue; // Skip merged hits
        int ln = (int)S(SCI_LINEFROMPOSITION, h.displayLineStart, 0);
        Sci_Position start = S(SCI_POSITIONFROMLINE, ln, 0);
        Sci_Position len = S(SCI_LINELENGTH, ln, 0);
        if (len > 0)
            S(SCI_INDICATORFILLRANGE, start, len);
    }

    // 2-b) Line-number digits
    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& h : _hits) {
        if (h.displayLineStart >= 0) // Ensure hit is valid
            S(SCI_INDICATORFILLRANGE,
                h.displayLineStart + h.numberStart,
                h.numberLen);
    }

    // 2-c) Match substrings (background and foreground)
    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_BG);
    for (const auto& h : _hits) {
        if (h.displayLineStart < 0) continue;
        for (size_t i = 0; i < h.matchStarts.size(); ++i)
            S(SCI_INDICATORFILLRANGE,
                h.displayLineStart + h.matchStarts[i],
                h.matchLens[i]);
    }

    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
    for (const auto& h : _hits) {
        if (h.displayLineStart < 0) continue;
        for (size_t i = 0; i < h.matchStarts.size(); ++i)
            S(SCI_INDICATORFILLRANGE,
                h.displayLineStart + h.matchStarts[i],
                h.matchLens[i]);
    }
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
        int x = LOWORD(lp);
        int y = HIWORD(lp);
        Sci_Position pos = ::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
        if (pos < 0)                         // guard for empty area
            return 0;
        int dispLine = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0);

        // --- Toggle fold if the clicked line is a header ----------
        int level = (int)::SendMessage(hwnd, SCI_GETFOLDLEVEL, dispLine, 0);
        if (level & SC_FOLDLEVELHEADERFLAG)
        {
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

    case WM_NCDESTROY:
        s_prevSciProc = nullptr;
        break;
    }

    return s_prevSciProc
        ? ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp)
        : ::DefWindowProc(hwnd, msg, wp, lp);
}





