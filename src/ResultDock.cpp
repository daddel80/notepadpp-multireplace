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

#include "PluginDefinition.h"
#include "Notepad_plus_msgs.h"
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
void ResultDock::startSearchBlock(const std::wstring& header, bool groupView, bool purge)
{
    if (purge) {
        clear();
        _searchHeaderLines.clear();
    }

    _pendingText.clear();
    _pendingHits.clear();
    _utf8LenPending = 0;
    _groupViewPending = groupView;
    _blockOpen = true;

    // Einleitungszeile (Search-Header). Zahlen werden in closeSearchBlock ersetzt.
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


void ResultDock::appendUtf8Chunked(HWND hSci, const std::string& u8)
{
    constexpr size_t CHUNK = 1u << 20; // 1 MB
    size_t sent = 0;
    const size_t total = u8.size();
    while (sent < total) {
        const size_t n = (std::min)(CHUNK, total - sent);
        ::SendMessageA(hSci, SCI_ADDTEXT, (WPARAM)n, (LPARAM)(u8.data() + sent));
        sent += n;
    }
}

// ---------------------------------------------------------------------
//  3. finalise header, insert block, restyle
// ---------------------------------------------------------------------
void ResultDock::closeSearchBlock(int totalHits, int totalFiles)
{
    if (!_blockOpen) return;

    // Replace placeholders in the pending header line
    const std::wstring hitsStr = std::to_wstring(totalHits);
    const std::wstring filesStr = std::to_wstring(totalFiles);

    size_t p = _pendingText.find_first_of(L"0123456789", /*start*/1);
    if (p != std::wstring::npos) {
        size_t e = _pendingText.find_first_not_of(L"0123456789", p);
        _pendingText.replace(p, (e == std::wstring::npos ? _pendingText.size() - p : e - p), hitsStr);

        size_t p2 = _pendingText.find_first_of(L"0123456789", p + hitsStr.size());
        if (p2 != std::wstring::npos) {
            size_t e2 = _pendingText.find_first_not_of(L"0123456789", p2);
            _pendingText.replace(p2, (e2 == std::wstring::npos ? _pendingText.size() - p2 : e2 - p2), filesStr);
        }
    }

    // Adjust byte offsets due to the replaced numbers
    const size_t newU8 = Encoding::wstringToUtf8(_pendingText).size();
    const ptrdiff_t deltaBytes = static_cast<ptrdiff_t>(newU8) - static_cast<ptrdiff_t>(_utf8LenPending);
    if (deltaBytes != 0) {
        for (auto& h : _pendingHits)
            h.displayLineStart += static_cast<int>(deltaBytes);
    }

    // Prepend the whole block (with a blank separator line between searches)
    prependBlock(_pendingText, _pendingHits);

    _blockOpen = false;
}

// --- Other Public Methods ------------------------------------------------

void ResultDock::clear()
{
    // ----- Data structures -------------------------------------------------
    _hits.clear();
    _lineStartToHitIndex.clear();


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

void ResultDock::prependBlock(const std::wstring& dockText, std::vector<Hit>& newHits)
{
    if (!_hSci || dockText.empty())
        return;

    // Convert once to UTF-8
    const std::string u8 = Encoding::wstringToUtf8(dockText);

    // Current buffer length (in bytes)
    const Sci_Position oldLen = (Sci_Position)S(SCI_GETLENGTH);

    // Bytes we will insert at position 0. We ensure a blank separator line if there is old content.
    const int sepBytes = (oldLen > 0 ? 2 : 0);          // "\r\n" between searches
    const int deltaBytes = (int)u8.size() + sepBytes;

    // Shift all existing hit offsets by the inserted byte count
    for (auto& h : _hits)
        h.displayLineStart += deltaBytes;

    // Insert new block at top (minimize UI work)
    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);
    S(SCI_SETREADONLY, FALSE);
    S(SCI_ALLOCATE, oldLen + (Sci_Position)deltaBytes + 65536, 0);

    // Insert block text at position 0
    S(SCI_INSERTTEXT, 0, (LPARAM)u8.c_str());
    // Insert blank separator line between searches (exactly one CRLF)
    if (sepBytes)
        S(SCI_INSERTTEXT, (WPARAM)u8.size(), (LPARAM)("\r\n"));

    S(SCI_SETREADONLY, TRUE);
    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);

    // Range-only styling/folding for the newly inserted block
    const Sci_Position pos0 = 0;
    const Sci_Position len = (Sci_Position)u8.size();
    // The newly inserted block spans from line 0 to lastLine, plus 1 separator line if any
    int newBlockLines = 0; for (char c : u8) if (c == '\n') ++newBlockLines;
    const int firstLine = 0;
    const int lastLine = (newBlockLines > 0 ? newBlockLines - 1 : 0);

    rebuildFoldingRange(firstLine, lastLine);
    applyStylingRange(pos0, len, newHits);

    // Collapse the previous (now second) top search block in O(1)
    if (oldLen > 0) {
        const int firstLineOfOldBlock = newBlockLines + 1; // + separator line
        const int level = (int)S(SCI_GETFOLDLEVEL, firstLineOfOldBlock);
        if ((level & SC_FOLDLEVELHEADERFLAG) && S(SCI_GETFOLDEXPANDED, firstLineOfOldBlock))
            S(SCI_FOLDLINE, firstLineOfOldBlock, SC_FOLDACTION_CONTRACT);
    }

    // Add the new hits to our master list (already at correct absolute positions)
    _hits.insert(_hits.begin(),
        std::make_move_iterator(newHits.begin()),
        std::make_move_iterator(newHits.end()));

    // Rebuild O(1) index after structural change
    rebuildHitLineIndex();

    // Keep the view at the top so the user sees the newest block immediately
    ::InvalidateRect(_hSci, nullptr, FALSE);
    S(SCI_SETFIRSTVISIBLELINE, 0);
    S(SCI_GOTOPOS, 0);
}

void ResultDock::rebuildFoldingRange(int firstLine, int lastLine) const
{
    if (!_hSci || lastLine < firstLine) return;

    const int BASE = SC_FOLDLEVELBASE;

    for (int l = firstLine; l <= lastLine; ++l)
        S(SCI_SETFOLDLEVEL, l, BASE);

    for (int l = firstLine; l <= lastLine; ++l) {
        const int indent = (int)S(SCI_GETLINEINDENTATION, l);
        bool isHeader = false;
        int  level = BASE;

        if (indent == INDENT_SPACES[(int)LineLevel::SearchHdr]) { isHeader = true; level = BASE + (int)LineLevel::SearchHdr; }
        else if (indent == INDENT_SPACES[(int)LineLevel::FileHdr]) { isHeader = true; level = BASE + (int)LineLevel::FileHdr; }
        else if (indent == INDENT_SPACES[(int)LineLevel::CritHdr]) { isHeader = true; level = BASE + (int)LineLevel::CritHdr; }

        if (isHeader)
            S(SCI_SETFOLDLEVEL, l, level | SC_FOLDLEVELHEADERFLAG);
        else
            S(SCI_SETFOLDLEVEL, l, BASE + (int)LineLevel::HitLine);
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

void ResultDock::applyStylingRange(Sci_Position pos0, Sci_Position len, const std::vector<Hit>& newHits) const
{
    if (!_hSci || len <= 0) return;

    const Sci_Position endPos = pos0 + len;
    const int firstLine = (int)S(SCI_LINEFROMPOSITION, pos0);
    const int lastLine = (int)S(SCI_LINEFROMPOSITION, endPos);

    // Base styling for affected lines, using indentation only (no full GETLINE)
    Sci_Position lineStartPos = S(SCI_POSITIONFROMLINE, firstLine);
    S(SCI_STARTSTYLING, lineStartPos, 0);

    for (int line = firstLine; line <= lastLine; ++line) {
        const Sci_Position ls = S(SCI_POSITIONFROMLINE, line);
        const int ll = (int)S(SCI_LINELENGTH, line);

        int style = STYLE_DEFAULT;
        if (ll > 0) {
            const int indent = (int)S(SCI_GETLINEINDENTATION, line); // spaces (tabs expanded)
            if (indent == INDENT_SPACES[(int)LineLevel::SearchHdr]) style = STYLE_HEADER;
            else if (indent == INDENT_SPACES[(int)LineLevel::CritHdr])   style = STYLE_CRITHDR;
            else if (indent == INDENT_SPACES[(int)LineLevel::FileHdr])   style = STYLE_FILEPATH;
        }
        if (ll > 0) S(SCI_SETSTYLING, ll, style);

        // EOL default (optional; keeps visuals consistent)
        const Sci_Position lineEnd = S(SCI_GETLINEENDPOSITION, line);
        const int eolLen = (int)(lineEnd - (ls + ll));
        if (eolLen > 0) S(SCI_SETSTYLING, eolLen, STYLE_DEFAULT);
    }

    // Indicators only for the new hits (minimal work: no global clears)
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    for (const auto& h : newHits) {
        if (h.displayLineStart < 0) continue;
        const int line = (int)S(SCI_LINEFROMPOSITION, h.displayLineStart);
        const Sci_Position ls = S(SCI_POSITIONFROMLINE, line);
        const Sci_Position ll = S(SCI_LINELENGTH, line);
        if (ll > 0) S(SCI_INDICATORFILLRANGE, ls, ll);
    }

    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& h : newHits)
        if (h.displayLineStart >= 0)
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.numberStart, h.numberLen);

    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_BG);
    for (const auto& h : newHits) {
        if (h.displayLineStart < 0) continue;
        for (size_t i = 0; i < h.matchStarts.size(); ++i)
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.matchStarts[i], h.matchLens[i]);
    }

    S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
    for (const auto& h : newHits) {
        if (h.displayLineStart < 0) continue;
        for (size_t i = 0; i < h.matchStarts.size(); ++i)
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.matchStarts[i], h.matchLens[i]);
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
        WS_EX_CLIENTEDGE,
        L"Scintilla",
        L"",
        WS_CHILD | WS_CLIPSIBLINGS,
        0, 0, 100, 100,
        npp._nppHandle,
        nullptr,
        _hInst,
        nullptr);

    if (!_hSci) {
        ::MessageBoxW(npp._nppHandle,
            L"FATAL: Failed to create Scintilla window!",
            L"ResultDock Error",
            MB_OK | MB_ICONERROR);
        return;
    }

    // 2) Subclass Scintilla so we receive double-clicks and other messages
    s_prevSciProc = reinterpret_cast<WNDPROC>(
        ::SetWindowLongPtrW(_hSci, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(sciSubclassProc)));

    // 3) Fast buffer config for huge search results
    ::SendMessageW(_hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
    ::SendMessageW(_hSci, SCI_SETUNDOCOLLECTION, FALSE, 0);   // no undo stack
    ::SendMessageW(_hSci, SCI_SETMODEVENTMASK, 0, 0);         // no mod events
    ::SendMessageW(_hSci, SCI_SETSCROLLWIDTHTRACKING, TRUE, 0);
    ::SendMessageW(_hSci, SCI_SETWRAPMODE, wrapEnabled() ? SC_WRAP_WORD : SC_WRAP_NONE, 0);

    // 3b) DirectFunction (hot path)
    _sciFn = reinterpret_cast<SciFnDirect_t>(::SendMessage(_hSci, SCI_GETDIRECTFUNCTION, 0, 0));
    _sciPtr = static_cast<sptr_t>(::SendMessage(_hSci, SCI_GETDIRECTPOINTER, 0, 0));

    // 4) Fill docking descriptor (kept from your version)
    static UINT  s_cachedDpi = 0;
    static HICON s_cachedLightIcon = nullptr;

    ::ZeroMemory(&_dockData, sizeof(_dockData));
    _dockData.hClient = _hSci;
    _dockData.pszName = L"MultiReplace – Search results";
    _dockData.dlgID = IDD_MULTIREPLACE_RESULT_DOCK;
    _dockData.uMask = DWS_DF_CONT_BOTTOM | DWS_ICONTAB;
    _dockData.pszAddInfo = L"";
    _dockData.pszModuleName = NPP_PLUGIN_NAME;
    _dockData.iPrevCont = -1;
    _dockData.rcFloat = { 0, 0, 0, 0 };

    const bool darkMode = (::SendMessage(npp._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0);
    if (!darkMode) {
        UINT dpi = 96;
        if (HDC hdc = ::GetDC(npp._nppHandle)) {
            dpi = ::GetDeviceCaps(hdc, LOGPIXELSX);
            ::ReleaseDC(npp._nppHandle, hdc);
        }
        if (dpi != s_cachedDpi) {
            if (s_cachedLightIcon) ::DestroyIcon(s_cachedLightIcon);
            extern HBITMAP CreateBitmapFromArray(UINT dpi);
            if (HBITMAP bmp = CreateBitmapFromArray(dpi)) {
                ICONINFO ii = {}; ii.fIcon = TRUE; ii.hbmMask = bmp; ii.hbmColor = bmp;
                s_cachedLightIcon = ::CreateIconIndirect(&ii);
                ::DeleteObject(bmp);
            }
            s_cachedDpi = dpi;
        }
        _dockData.hIconTab = s_cachedLightIcon
            ? s_cachedLightIcon
            : static_cast<HICON>(::LoadImage(_hInst, MAKEINTRESOURCE(IDI_MR_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
    }
    else {
        _dockData.hIconTab = static_cast<HICON>(::LoadImage(_hInst, MAKEINTRESOURCE(IDI_MR_DM_ICON), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
    }

    // 5) Register dock & theme it
    _hDock = reinterpret_cast<HWND>(::SendMessageW(npp._nppHandle, NPPM_DMMREGASDCKDLG, 0, reinterpret_cast<LPARAM>(&_dockData)));
    if (!_hDock) {
        ::MessageBoxW(npp._nppHandle, L"ERROR: Docking registration failed.", L"ResultDock Error", MB_OK | MB_ICONERROR);
        return;
    }
    ::SendMessageW(npp._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME,
        static_cast<WPARAM>(NppDarkMode::dmfInit),
        reinterpret_cast<LPARAM>(_hDock));

    // 6) Initial folding & theme
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
    S(SCI_INDICSETSTYLE, INDIC_MATCH_BG, INDIC_HIDDEN);
    S(SCI_INDICSETUNDER, INDIC_MATCH_BG, TRUE);

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

void ResultDock::collapseOldSearches()
{
    if (!_hSci || _searchHeaderLines.size() < 2)
        return;

    // Alle außer dem letzten (aktuellen) zusammenfalten
    const size_t lastIdx = _searchHeaderLines.size() - 1;
    for (size_t i = 0; i < lastIdx; ++i) {
        const int headerLine = _searchHeaderLines[i];
        // Sicherheitsprüfung: innerhalb des aktuellen LineCounts?
        const int lineCount = (int)::SendMessage(_hSci, SCI_GETLINECOUNT, 0, 0);
        if (headerLine >= 0 && headerLine < lineCount) {
            ::SendMessage(_hSci, SCI_SETFOLDEXPANDED, headerLine, FALSE);
            ::SendMessage(_hSci, SCI_FOLDCHILDREN, headerLine, SC_FOLDACTION_CONTRACT);
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
    const UINT docCp = (UINT)sciSend(SCI_GETCODEPAGE, 0, 0);

    const std::wstring indentHitW = getIndentString(LineLevel::HitLine);
    const size_t       indentHitU8 = getIndentUtf8Length(LineLevel::HitLine);

    const std::wstring kLineW = LM.get(L"dock_line") + L" ";
    static const size_t kLineU8 = Encoding::wstringToUtf8(kLineW).size();
    constexpr size_t    kColonSpaceU8 = 2; // ": "

    // 1) Precompute line numbers + max digits
    std::vector<int> lineNumbers; lineNumbers.reserve(hits.size());
    size_t maxDigits = 0;
    for (const Hit& h : hits) {
        int line1 = (int)sciSend(SCI_LINEFROMPOSITION, h.pos, 0) + 1;
        lineNumbers.push_back(line1);
        maxDigits = (std::max)(maxDigits, std::to_wstring(line1).size());
    }

    auto padNumber = [](int number, size_t width) -> std::wstring {
        std::wstring numStr = std::to_wstring(number);
        size_t pad = (width > numStr.size() ? width - numStr.size() : 0);
        return std::wstring(pad, L' ') + numStr;
        };

    int   prevDocLine = -1;
    Hit* firstHitOnRow = nullptr;
    size_t hitIdx = 0;

    // ---- Per line cache (reduces conversions for multi-hit lines) ----
    std::string cachedRaw;         // original bytes (line without CRLF)
    std::string cachedU8;          // same line as UTF-8
    std::wstring cachedW;          // same line as W
    Sci_Position cachedAbsLineStart = 0;

    // byte→utf8 prefix length cache for this line (only for needed offsets)
    std::unordered_map<size_t, size_t> u8PrefixLenByByte;

    auto loadLineIfNeeded = [&](int line0)
        {
            if (line0 == prevDocLine) return;

            int rawLen = (int)sciSend(SCI_LINELENGTH, line0, 0);
            cachedRaw.assign(rawLen, '\0');
            sciSend(SCI_GETLINE, line0, (LPARAM)cachedRaw.data());
            while (!cachedRaw.empty() && (cachedRaw.back() == '\r' || cachedRaw.back() == '\n'))
                cachedRaw.pop_back();

            cachedAbsLineStart = sciSend(SCI_POSITIONFROMLINE, line0, 0);
            cachedU8 = Encoding::bytesToUtf8(cachedRaw, docCp);
            cachedW = Encoding::utf8ToWString(cachedU8);

            u8PrefixLenByByte.clear(); // reset cache for this new line
        };

    for (Hit& h : hits) {
        int line1 = lineNumbers[hitIdx++];
        int line0 = line1 - 1;

        loadLineIfNeeded(line0);

        // Compute offsets relative to the raw line bytes
        size_t relBytes = (size_t)(h.pos - cachedAbsLineStart);

        auto u8Prefix = [&](size_t byteCount)->size_t {
            auto it = u8PrefixLenByByte.find(byteCount);
            if (it != u8PrefixLenByByte.end()) return it->second;
            size_t val = Encoding::bytesToUtf8(cachedRaw.substr(0, byteCount), docCp).size();
            u8PrefixLenByByte.emplace(byteCount, val);
            return val;
            };

        size_t hitStartU8 = u8Prefix(relBytes);
        size_t hitLenU8 = Encoding::bytesToUtf8(cachedRaw.substr(relBytes, h.length), docCp).size();

        if (line0 != prevDocLine) {
            std::wstring paddedNumW = padNumber(line1, maxDigits);
            std::wstring prefixW = indentHitW + kLineW + paddedNumW + L": ";

            const size_t prefixU8Len = indentHitU8 + kLineU8 + paddedNumW.size() + kColonSpaceU8;

            out += prefixW + cachedW + L"\r\n";

            h.displayLineStart = (int)utf8Pos;
            h.numberStart = (int)(indentHitU8 + kLineU8 + paddedNumW.size() - std::to_wstring(line1).size());
            h.numberLen = (int)std::to_wstring(line1).size();
            h.matchStarts = { (int)(prefixU8Len + hitStartU8) };
            h.matchLens = { (int)hitLenU8 };

            utf8Pos += prefixU8Len + cachedU8.size() + 2; // + CRLF

            firstHitOnRow = &h;
            prevDocLine = line0;
        }
        else {
            // Additional hit on same visible row: only extend match ranges
            if (firstHitOnRow) {
                const size_t prefixU8Len = indentHitU8 + kLineU8
                    + std::to_wstring(line1).size() + (maxDigits - std::to_wstring(line1).size()) + kColonSpaceU8;
                firstHitOnRow->matchStarts.push_back((int)(prefixU8Len + hitStartU8));
                firstHitOnRow->matchLens.push_back((int)hitLenU8);
            }
            h.displayLineStart = -1; // dummy
        }
    }

    // Drop dummy entries
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
    return std::wstring(INDENT_SPACES[static_cast<int>(lvl)], L' ');
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
            out.emplace_back(stripHitPrefix(Encoding::utf8ToWString(rawCaret)));
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
                endLine = subtreeEnd(l + 1, INDENT_SPACES[static_cast<int>(LineLevel::CritHdr)]);
                break;
            case LineKind::FileHdr:
                endLine = subtreeEnd(l + 1, INDENT_SPACES[static_cast<int>(LineLevel::FileHdr)]);
                break;
            case LineKind::SearchHdr:
                endLine = subtreeEnd(l + 1, INDENT_SPACES[static_cast<int>(LineLevel::SearchHdr)]);
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
        case LineKind::CritHdr: endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::CritHdr)]);
            break;
        case LineKind::FileHdr: endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::FileHdr)]);
            break;
        case LineKind::SearchHdr: endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[static_cast<int>(LineLevel::SearchHdr)]);
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
    dock.rebuildHitLineIndex();
}

bool ResultDock::IsPseudoPath(const std::wstring& p)
{
    // Pseudo when no drive/dir separators -> e.g. "new 1"
    return (p.find(L'\\') == std::wstring::npos) &&
        (p.find(L'/') == std::wstring::npos) &&
        (p.find(L':') == std::wstring::npos);
}

bool ResultDock::FileExistsW(const std::wstring& fullPath)
{
    const DWORD a = ::GetFileAttributesW(fullPath.c_str());
    return (a != INVALID_FILE_ATTRIBUTES) && !(a & FILE_ATTRIBUTE_DIRECTORY);
}

std::wstring ResultDock::GetNppProgramDir()
{
    wchar_t buf[MAX_PATH] = { 0 };
    DWORD n = ::GetModuleFileNameW(nullptr, buf, MAX_PATH);
    std::wstring exePath = (n ? std::wstring(buf, n) : std::wstring());
    const size_t pos = exePath.find_last_of(L"\\/");
    return (pos == std::wstring::npos) ? L"." : exePath.substr(0, pos);
}

bool ResultDock::IsCurrentDocByFullPath(const std::wstring& fullPath)
{
    wchar_t cur[MAX_PATH] = { 0 };
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, (WPARAM)MAX_PATH, (LPARAM)cur);
    return (_wcsicmp(cur, fullPath.c_str()) == 0);
}

bool ResultDock::IsCurrentDocByTitle(const std::wstring& titleOnly)
{
    // Compare against current tab title (filename only)
    wchar_t name[MAX_PATH] = { 0 };
    ::SendMessage(nppData._nppHandle, NPPM_GETFILENAME, (WPARAM)MAX_PATH, (LPARAM)name);
    return (_wcsicmp(name, titleOnly.c_str()) == 0);
}

void ResultDock::SwitchToFileIfOpenByFullPath(const std::wstring& fullPath)
{
    // NPP reuses an existing tab if the file is open; otherwise: no effect.
    ::SendMessage(nppData._nppHandle, NPPM_SWITCHTOFILE, 0, (LPARAM)fullPath.c_str());
}

std::wstring ResultDock::BuildDefaultPathForPseudo(const std::wstring& label)
{
    // Map "new 1" -> "<NppProgramDir>\new 1"
    const std::wstring dir = GetNppProgramDir();
    if (dir.empty()) return label;
    return dir + L"\\" + label;
}

bool ResultDock::EnsureFileOpenOrOfferCreate(const std::wstring& desiredPath, std::wstring& outOpenedPath)
{
    // Detect pseudo (e.g., "new 1") and normalize to a real path if needed.
    const bool isPseudo = IsPseudoPath(desiredPath);

    // Fast path: if it's a pseudo-name and that tab is currently open, just return.
    if (isPseudo && IsCurrentDocByTitle(desiredPath)) {
        outOpenedPath = desiredPath;
        return true;
    }

    // Normalize pseudo names to a full path upfront (e.g., "<NppDir>\new 1")
    const std::wstring targetPath = isPseudo ? BuildDefaultPathForPseudo(desiredPath)
        : desiredPath;

    // Try to activate if already open
    SwitchToFileIfOpenByFullPath(desiredPath);
    if (IsCurrentDocByFullPath(desiredPath)) {
        outOpenedPath = desiredPath;
        return true;
    }

    // Not open: try to open if it exists
    if (FileExistsW(desiredPath)) {
        ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)desiredPath.c_str());
        outOpenedPath = desiredPath;
        return true;
    }

    // Missing: prompt to create at the desiredPath
    const std::wstring title = LM.get(L"msgbox_title_create_file");
    const std::wstring prompt = LM.get(L"msgbox_prompt_create_file", { desiredPath });

    const int res = ::MessageBoxW(nppData._nppHandle, prompt.c_str(), title.c_str(), MB_YESNO | MB_APPLMODAL);
    if (res != IDYES) return false;

    // Try to create an empty file; if it fails, show error
    HANDLE hFile = ::CreateFileW(desiredPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr,
        CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        const std::wstring err = LM.get(L"msgbox_error_create_file", { desiredPath });
        ::MessageBoxW(nppData._nppHandle, err.c_str(), title.c_str(), MB_OK | MB_APPLMODAL);
        return false;
    }
    ::CloseHandle(hFile);

    // Open the just created file
    const LRESULT ok = ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, (LPARAM)desiredPath.c_str());
    if (ok) {
        outOpenedPath = desiredPath;
        return true;
    }

    // Should not happen, but keep safe fallback
    const std::wstring err = LM.get(L"msgbox_error_create_file", { desiredPath });
    ::MessageBoxW(nppData._nppHandle, err.c_str(), title.c_str(), MB_OK | MB_APPLMODAL);
    return true;
}

void ResultDock::JumpSelectCenterActiveEditor(Sci_Position pos, Sci_Position len)
{
    // Always use the currently active Scintilla to avoid focusing the wrong view.
    int whichView = 0; // 0 main, 1 secondary
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&whichView);
    HWND hEd = (whichView == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
    if (!hEd) return;

    ::SetFocus(hEd);
    ::SendMessage(hEd, SCI_SETFOCUS, TRUE, 0);

    // Ensure line visible, select range, and center
    const Sci_Position targetLine = (Sci_Position)::SendMessage(hEd, SCI_LINEFROMPOSITION, pos, 0);
    ::SendMessage(hEd, SCI_ENSUREVISIBLE, targetLine, 0);
    ::SendMessage(hEd, SCI_GOTOPOS, pos, 0);
    ::SendMessage(hEd, SCI_SETSEL, pos, pos + len);
    ::SendMessage(hEd, SCI_SCROLLCARET, 0, 0);

    // Center the caret line without extra auto-scrolling
    const size_t firstVisibleDocLine =
        (size_t)::SendMessage(hEd, SCI_DOCLINEFROMVISIBLE, (WPARAM)::SendMessage(hEd, SCI_GETFIRSTVISIBLELINE, 0, 0), 0);
    size_t linesOnScreen = (size_t)::SendMessage(hEd, SCI_LINESONSCREEN, (WPARAM)firstVisibleDocLine, 0);
    if (linesOnScreen == 0) linesOnScreen = 1;
    const size_t caretLine = (size_t)::SendMessage(hEd, SCI_LINEFROMPOSITION, pos, 0);
    const size_t midDisplay = firstVisibleDocLine + linesOnScreen / 2;
    const ptrdiff_t deltaLines = (ptrdiff_t)caretLine - (ptrdiff_t)midDisplay;
    ::SendMessage(hEd, SCI_LINESCROLL, 0, deltaLines);
    ::SendMessage(hEd, SCI_ENSUREVISIBLEENFORCEPOLICY, (WPARAM)caretLine, 0);
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

                  // ResultDock.cpp  (innerhalb LRESULT CALLBACK ResultDock::sciSubclassProc(...))
    case WM_LBUTTONDBLCLK:
    {
        // 1) Remember current dock scroll state
        const LRESULT firstVisible = ::SendMessage(hwnd, SCI_GETFIRSTVISIBLELINE, 0, 0);

        // 2) Determine clicked line
        const int x = LOWORD(lp);
        const int y = HIWORD(lp);
        const Sci_Position posInDock = ::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
        if (posInDock < 0) return 0;
        const int dispLine = (int)::SendMessage(hwnd, SCI_LINEFROMPOSITION, posInDock, 0);

        // 2a) Header fold toggle unchanged
        const int level = (int)::SendMessage(hwnd, SCI_GETFOLDLEVEL, dispLine, 0);
        if (level & SC_FOLDLEVELHEADERFLAG) {
            ::SendMessage(hwnd, SCI_TOGGLEFOLD, dispLine, 0);
            const Sci_Position lineStartPos = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
            ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, lineStartPos, 0);
            ::SendMessage(hwnd, SCI_SETFIRSTVISIBLELINE, firstVisible, 0);
            return 0;
        }

        // 3) Read raw to classify as HitLine
        const int lineLen = (int)::SendMessage(hwnd, SCI_LINELENGTH, dispLine, 0);
        std::string raw(lineLen, '\0');
        ::SendMessage(hwnd, SCI_GETLINE, dispLine, (LPARAM)raw.data());
        raw.resize(strnlen(raw.c_str(), lineLen));
        if (ResultDock::classify(raw) != LineKind::HitLine) return 0;

        // 4) Compute hitIndex by counting previous HitLines (unchanged)
        int hitIndex = -1;
        for (int i = 0; i <= dispLine; ++i) {
            const int len = (int)::SendMessage(hwnd, SCI_LINELENGTH, i, 0);
            std::string buf(len, '\0');
            ::SendMessage(hwnd, SCI_GETLINE, i, (LPARAM)buf.data());
            buf.resize(strnlen(buf.c_str(), len));
            if (ResultDock::classify(buf) == LineKind::HitLine)
                ++hitIndex;
        }
        const std::vector<Hit>& hits = ResultDock::instance().hits();
        if (hitIndex < 0 || hitIndex >= (int)hits.size()) return 0;
        const Hit& hit = hits[(size_t)hitIndex];

        // 5) Convert UTF-8 path to wide
        std::wstring wPath;
        const int wlen = ::MultiByteToWideChar(CP_UTF8, 0, hit.fullPathUtf8.c_str(), -1, nullptr, 0);
        if (wlen > 0) {
            wPath.resize(wlen - 1);
            ::MultiByteToWideChar(CP_UTF8, 0, hit.fullPathUtf8.c_str(), -1, &wPath[0], wlen);
        }

        // 6) Decide pseudo vs real, ensure open, then select
        const bool isPseudo = ResultDock::IsPseudoPath(wPath);
        std::wstring pathToOpen = wPath;

        if (isPseudo) {
            // First try: if current tab title matches this pseudo label -> just jump
            if (ResultDock::IsCurrentDocByTitle(wPath)) {
                ResultDock::JumpSelectCenterActiveEditor(hit.pos, hit.length);
            }
            else {
                // Not open: use generated default path, then same logic as real files
                pathToOpen = ResultDock::BuildDefaultPathForPseudo(wPath);
                std::wstring opened;
                const bool ok = ResultDock::EnsureFileOpenOrOfferCreate(pathToOpen, opened);
                if (!ok) return 0;
                ResultDock::JumpSelectCenterActiveEditor(hit.pos, hit.length);
            }
        }
        else {
            // Real path: try switch, open if needed, or ask to create
            std::wstring opened;
            const bool ok = ResultDock::EnsureFileOpenOrOfferCreate(pathToOpen,  opened);
            if (!ok) return 0;
            ResultDock::JumpSelectCenterActiveEditor(hit.pos, hit.length);
        }

        // 7) Restore dock scroll & clear its selection; keep focus on dock (as before)
        ::SendMessage(hwnd, SCI_SETFIRSTVISIBLELINE, firstVisible, 0);
        const Sci_Position lineStartPos = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
        ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, lineStartPos, 0);
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

void ResultDock::rebuildHitLineIndex()
{
    _lineStartToHitIndex.clear();
    _lineStartToHitIndex.reserve(_hits.size());
    for (int i = 0; i < (int)_hits.size(); ++i)
    {
        const int pos = _hits[i].displayLineStart;
        if (pos >= 0) _lineStartToHitIndex[pos] = i;
    }
}
