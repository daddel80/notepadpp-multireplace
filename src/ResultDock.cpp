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
#include "ColumnTabs.h"
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
#include "NppStyleKit.h"

extern NppData nppData;

// #define MR_DEBUG_JUMP 0

namespace {
    LanguageManager& LM = LanguageManager::instance();

    // Case-insensitive UTF-8 path comparison (Windows paths are case-insensitive)
    inline bool pathsEqualUtf8(const std::string& a, const std::string& b) {
        return _stricmp(a.c_str(), b.c_str()) == 0;
    }

    // Consolidated pending jump state - minimal data for robust cross-file navigation
    struct PendingJumpState {
        bool active = false;
        std::wstring path;              // Target file path for validation
        std::string fullPathUtf8;       // UTF-8 path for file matching

        // Minimal fields needed for NavigateToHit-style re-search
        int docLine = -1;               // 0-based line number in target document
        int searchFlags = 0;            // SCFIND_MATCHCASE | SCFIND_WHOLEWORD | SCFIND_REGEXP
        std::wstring findTextW;         // Search text for re-search on line

        // Fallback position if re-search fails
        Sci_Position fallbackPos = 0;
        Sci_Position fallbackLen = 0;

        // Navigation state machine
        HWND targetEditor = nullptr;    // Active editor to watch
        int phase = 0;                  // 0=idle, 1=buffer-activated, 2=update-seen

        void clear() {
            active = false;
            path.clear();
            fullPathUtf8.clear();
            docLine = -1;
            searchFlags = 0;
            findTextW.clear();
            fallbackPos = 0;
            fallbackLen = 0;
            targetEditor = nullptr;
            phase = 0;
        }

        void setFromHit(const std::wstring& targetPath, const ResultDock::Hit& hit) {
            active = !targetPath.empty();
            path = targetPath;
            fullPathUtf8 = hit.fullPathUtf8;
            docLine = hit.docLine;
            searchFlags = hit.searchFlags;
            findTextW = hit.findTextW;
            fallbackPos = hit.pos;
            fallbackLen = hit.length;
            targetEditor = nullptr;
            phase = 0;
        }
    };

    static PendingJumpState s_pending;
    static const UINT s_timerId = 1001;

    // Legacy function signature for compatibility with SwitchAndJump
    static void SetPendingJump(const std::wstring& path, Sci_Position pos, Sci_Position len)
    {
        s_pending.clear();
        s_pending.active = !path.empty();
        s_pending.path = path;
        s_pending.fallbackPos = pos;
        s_pending.fallbackLen = len;
        // docLine = -1 means fallback-only mode (no re-search)
    }

    static constexpr uint32_t argb(BYTE a, COLORREF c)
    {
        return (uint32_t(a) << 24) |
            (uint32_t(GetRValue(c)) << 16) |
            (uint32_t(GetGValue(c)) << 8) |
            uint32_t(GetBValue(c));
    }
}

// --------------- Singleton & API methods ------------------

ResultDock& ResultDock::instance()
{
    // This assumes a global g_inst is defined and set in DllMain.
    extern HINSTANCE g_inst;
    static ResultDock s{ g_inst };
    return s;
}

void ResultDock::ensureCreated(const NppData& npp)
{
    // Create the dock if it doesn't exist yet (but don't show it)
    if (!_hSci)
        create(npp);
}

void ResultDock::ensureCreatedAndVisible(const NppData& npp)
{
    // 1) first-time creation
    if (!_hSci)                // _hSci is initialized in create()
        create(npp);

    // 2) show – MUST use the client handle!
    if (_hSci) {
        ::SendMessage(npp._nppHandle, NPPM_DMMSHOW, 0, reinterpret_cast<LPARAM>(_hSci));
    }
}

void ResultDock::hide(const NppData& npp)
{
    if (_hSci) {
        ::SendMessage(npp._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(_hSci));

        // Force editor to repaint the area where dock was
        ::InvalidateRect(npp._scintillaMainHandle, nullptr, TRUE);
        ::InvalidateRect(npp._scintillaSecondHandle, nullptr, TRUE);
        ::UpdateWindow(npp._nppHandle);
    }
}

bool ResultDock::hasHitsForFile(const std::string& fullPathUtf8) const
{
    for (const auto& hit : _hits) {
        if (pathsEqualUtf8(hit.fullPathUtf8, fullPathUtf8))
            return true;
    }
    return false;
}

void ResultDock::clear()
{
    // ----- Data structures -------------------------------------------------
    _hits.clear();
    _lineStartToHitIndex.clear();

    _searchHeaderLines.clear();
    _slotToColor.clear();

    if (!_hSci)
        return;

    // ----- Reset Scintilla buffer completely --------------------------
    S(SCI_SETREADONLY, FALSE);
    S(SCI_CLEARALL);                                      // Text & styles
    S(SCI_SETREADONLY, TRUE);

    // remove all folding levels → BASE
    const int lineCount = static_cast<int>(S(SCI_GETLINECOUNT));
    for (int l = 0; l < lineCount; ++l)
        S(SCI_SETFOLDLEVEL, l, SC_FOLDLEVELBASE);;

    // Clear every indicator used by the dock
    for (int ind : { INDIC_LINE_BACKGROUND, INDIC_LINENUMBER_FORE,
        INDIC_MATCH_BG, INDIC_MATCH_FORE })
    {
        S(SCI_SETINDICATORCURRENT, ind);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // Clear per-entry background color indicators
    for (int i = 0; i < MAX_ENTRY_COLORS; ++i) {
        S(SCI_SETINDICATORCURRENT, INDIC_ENTRY_BG_BASE + i);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // reset caret/scroll position
    S(SCI_GOTOPOS, 0);
    S(SCI_SETFIRSTVISIBLELINE, 0);

    // ----- Rebuild Styles & Folding -----------------------------------
    rebuildFolding();
    applyStyling();

    ::RedrawWindow(_hSci, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);
}

bool ResultDock::hasContentBeyondIndent(HWND hSci, int line)
{
    int lineEnd = static_cast<int>(::SendMessage(hSci, SCI_GETLINEENDPOSITION, line, 0));
    int indentPos = static_cast<int>(::SendMessage(hSci, SCI_GETLINEINDENTPOSITION, line, 0));

    return indentPos < lineEnd; // true if there is content after indentation
}

void ResultDock::rebuildFolding()
{
    if (!_hSci) return;

    const int lineCount = static_cast<int>(S(SCI_GETLINECOUNT));
    if (lineCount <= 0) return;

    // Ensure folding is enabled; harmless if already set
    S(SCI_SETPROPERTY, reinterpret_cast<uptr_t>("fold"), reinterpret_cast<sptr_t>("1"));
    S(SCI_SETPROPERTY, reinterpret_cast<uptr_t>("fold.compact"), reinterpret_cast<sptr_t>("1"));

    const int BASE = SC_FOLDLEVELBASE;

    // Cache indent thresholds and level constants locally to avoid array lookups in the loop
    const int IND_SEARCH = INDENT_SPACES[(int)LineLevel::SearchHdr];
    const int IND_FILE = INDENT_SPACES[(int)LineLevel::FileHdr];
    const int IND_CRIT = INDENT_SPACES[(int)LineLevel::CritHdr];
    const int IND_HIT = INDENT_SPACES[(int)LineLevel::HitLine];

    auto calcLevel = [&](int indent, bool& isHeader) -> int {
        // 1) Headers: exact indent match for clear detection
        if (indent == IND_SEARCH) { isHeader = true;  return BASE + (int)LineLevel::SearchHdr; }
        if (indent == IND_FILE) { isHeader = true;  return BASE + (int)LineLevel::FileHdr; }
        if (indent == IND_CRIT) { isHeader = true;  return BASE + (int)LineLevel::CritHdr; }

        // 2) Non-headers: bucket depth by >= thresholds (restores fine depth levels)
        isHeader = false;
        if (indent >= IND_HIT)    return BASE + (int)LineLevel::HitLine;
        if (indent >= IND_CRIT)   return BASE + (int)LineLevel::CritHdr;
        if (indent >= IND_FILE)   return BASE + (int)LineLevel::FileHdr;
        return BASE + 1; // minimal content depth
        };

    // Batch update to minimize repaints during rebuild
    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);

    for (int l = 0; l < lineCount; ++l)
    {
        const int lineLen = static_cast<int>(S(SCI_LINELENGTH, l));

        // True blank or whitespace-only → BASE (non-foldable)
        if (lineLen == 0 || !hasContentBeyondIndent(_hSci, l)) {
            const int cur = static_cast<int>(S(SCI_GETFOLDLEVEL, l));
            if (cur != BASE) S(SCI_SETFOLDLEVEL, l, BASE);
            continue;
        }

        const int indent = static_cast<int>(S(SCI_GETLINEINDENTATION, l));
        bool isHdr = false;
        const int level = calcLevel(indent, isHdr);
        const int target = isHdr ? (level | SC_FOLDLEVELHEADERFLAG) : level;

        // Only set if value changed to reduce unnecessary messages
        const int cur = static_cast<int>(S(SCI_GETFOLDLEVEL, l));
        if (cur != target) {
            S(SCI_SETFOLDLEVEL, l, target);
        }
    }

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);
}

void ResultDock::applyStyling() const
{
    if (!_hSci) return;

    // 1) Clear fixed indicators (New IDs at end of range)
    const std::vector<int> indicatorsToClear = {
        INDIC_LINE_BACKGROUND, INDIC_LINENUMBER_FORE, INDIC_MATCH_FORE, INDIC_MATCH_BG
    };
    for (int id : indicatorsToClear) {
        S(SCI_SETINDICATORCURRENT, id);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // 2) Clear dynamic entry indicators (Range 0 to MAX_ENTRY_COLORS)
    for (int i = 0; i < MAX_ENTRY_COLORS; ++i) {
        S(SCI_SETINDICATORCURRENT, INDIC_ENTRY_BG_BASE + i);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // 3) Configure Dynamic Indicators based on _slotToColor map
    const bool dark = ::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0;
    const int bgAlpha = dark ? ENTRY_BG_ALPHA_DARK : ENTRY_BG_ALPHA_LIGHT;
    const int outlineAlpha = dark ? ENTRY_OUTLINE_ALPHA_DARK : ENTRY_OUTLINE_ALPHA_LIGHT;

    if (_perEntryColorsEnabled) {
        // Configure only slots that are actually used
        for (const auto& [slotIdx, rgbColor] : _slotToColor) {
            // Safety check
            if (slotIdx < 0 || slotIdx >= MAX_ENTRY_COLORS) continue;

            const int indicId = INDIC_ENTRY_BG_BASE + slotIdx;
            S(SCI_INDICSETSTYLE, indicId, INDIC_ROUNDBOX);
            S(SCI_INDICSETFORE, indicId, static_cast<COLORREF>(rgbColor));
            S(SCI_INDICSETALPHA, indicId, bgAlpha);
            S(SCI_INDICSETOUTLINEALPHA, indicId, outlineAlpha);
            S(SCI_INDICSETUNDER, indicId, TRUE);
        }

        // Apply Colors to hits
        for (const auto& hit : _hits) {
            if (hit.displayLineStart < 0) continue;

            for (size_t i = 0; i < hit.matchStarts.size(); ++i) {
                if (i >= hit.matchColors.size()) break;

                // Retrieve the slot index we stored in matchColors
                int slotIdx = hit.matchColors[i];

                if (slotIdx >= 0 && slotIdx < MAX_ENTRY_COLORS) {
                    const int indicId = INDIC_ENTRY_BG_BASE + slotIdx;
                    S(SCI_SETINDICATORCURRENT, indicId);
                    S(SCI_INDICATORFILLRANGE, hit.displayLineStart + hit.matchStarts[i], hit.matchLens[i]);
                }
            }
        }
    }
    else {
        // Standard Mode (Single Color)
        // Red Match Color
        S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
        for (const auto& hit : _hits) {
            if (hit.displayLineStart < 0) continue;
            for (size_t i = 0; i < hit.matchStarts.size(); ++i) {
                S(SCI_INDICATORFILLRANGE, hit.displayLineStart + hit.matchStarts[i], hit.matchLens[i]);
            }
        }
    }
}
void ResultDock::onThemeChanged() {
    applyTheme();
    updateTabIcon();
}

void ResultDock::updateTabIcon() {
    if (!_hSci) return;

    const bool darkMode = (::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0);

    UINT dpi = 96;
    if (HDC hdc = ::GetDC(nppData._nppHandle)) {
        dpi = ::GetDeviceCaps(hdc, LOGPIXELSX);
        ::ReleaseDC(nppData._nppHandle, hdc);
    }

    // Cache icons to avoid recreation and memory leaks
    static HICON s_cachedTabIconLight = nullptr;
    static HICON s_cachedTabIconDark = nullptr;
    static UINT s_cachedDpi = 0;

    // Recreate icons only if DPI changed
    if (dpi != s_cachedDpi) {
        if (s_cachedTabIconLight) ::DestroyIcon(s_cachedTabIconLight);
        if (s_cachedTabIconDark) ::DestroyIcon(s_cachedTabIconDark);
        s_cachedTabIconLight = CreateIconFromImageData(gimp_image_tab_light, dpi);
        s_cachedTabIconDark = CreateIconFromImageData(gimp_image_tab_dark, dpi);
        s_cachedDpi = dpi;
    }

    HICON hNewIcon = darkMode ? s_cachedTabIconDark : s_cachedTabIconLight;
    if (!hNewIcon) return;

    // Update our local copy
    _dockData.hIconTab = hNewIcon;

    // Find the docking container and its tab control
    HWND hParent = ::GetParent(_hSci);
    while (hParent) {
        HWND hTab = ::FindWindowEx(hParent, NULL, WC_TABCONTROL, NULL);
        if (hTab) {
            int tabCount = static_cast<int>(::SendMessage(hTab, TCM_GETITEMCOUNT, 0, 0));
            for (int i = 0; i < tabCount; ++i) {
                TCITEM tcItem = {};
                tcItem.mask = TCIF_PARAM;
                if (::SendMessage(hTab, TCM_GETITEM, i, reinterpret_cast<LPARAM>(&tcItem)) && tcItem.lParam) {
                    tTbData* pTbData = reinterpret_cast<tTbData*>(tcItem.lParam);
                    if (pTbData->hClient == _hSci) {
                        pTbData->hIconTab = hNewIcon;
                        ::InvalidateRect(hTab, NULL, TRUE);
                        return;
                    }
                }
            }
        }
        hParent = ::GetParent(hParent);
    }
}

// ------------------- Search Block API ---------------------

// -----------------------------------------------------------
//  1. open a new block
// -----------------------------------------------------------
void ResultDock::startSearchBlock(const std::wstring& header, bool groupView, bool purge)
{
    if (purge) {
        clear();
        _searchHeaderLines.clear();
        _slotToColor.clear();
    }

    _pendingText.clear();
    _pendingHits.clear();
    _utf8LenPending = 0;
    _groupViewPending = groupView;
    _blockOpen = true;

    // Intro line for the search header; hit/file numbers are patched in closeSearchBlock().
    _pendingText = getIndentString(LineLevel::SearchHdr) + header + L"\r\n";
    _utf8LenPending = Encoding::wstringToUtf8(_pendingText).size();
}

// -----------------------------------------------------------
//  2. append results for ONE file (can be called many times)
// -----------------------------------------------------------
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

// -----------------------------------------------------------
//  3. finalise header, insert block, restyle
// -----------------------------------------------------------
void ResultDock::closeSearchBlock(int totalHits, int totalFiles)
{
    if (!_blockOpen) return;

    // Replace placeholders in the pending header line ONLY (first line)
    const std::wstring hitsStr = std::to_wstring(totalHits);
    const std::wstring filesStr = std::to_wstring(totalFiles);

    // Limit search to the first header line to avoid matching digits in
    // user search patterns or file content that follows
    const size_t headerEnd = _pendingText.find(L"\r\n");
    const size_t searchLimit = (headerEnd != std::wstring::npos) ? headerEnd : _pendingText.size();

    // Search backwards from the end of the header to find the placeholder numbers
    // which are always the last two number sequences before the closing parenthesis.
    // Pattern: "... (0 hits in 0 files)\r\n"
    //                ^hits       ^files
    size_t p2End = std::wstring::npos;
    size_t p2Start = std::wstring::npos;
    {
        // Find last digit sequence before searchLimit
        size_t pos = searchLimit;
        while (pos > 0) {
            --pos;
            if (_pendingText[pos] >= L'0' && _pendingText[pos] <= L'9') {
                p2End = pos + 1;
                while (pos > 0 && _pendingText[pos - 1] >= L'0' && _pendingText[pos - 1] <= L'9')
                    --pos;
                p2Start = pos;
                break;
            }
        }
    }

    // Find second-to-last digit sequence (hits count)
    size_t p1Start = std::wstring::npos;
    size_t p1End = std::wstring::npos;
    if (p2Start != std::wstring::npos && p2Start > 0) {
        size_t pos = p2Start;
        while (pos > 0) {
            --pos;
            if (_pendingText[pos] >= L'0' && _pendingText[pos] <= L'9') {
                p1End = pos + 1;
                while (pos > 0 && _pendingText[pos - 1] >= L'0' && _pendingText[pos - 1] <= L'9')
                    --pos;
                p1Start = pos;
                break;
            }
        }
    }

    // Replace files count first (higher position), then hits count
    if (p2Start != std::wstring::npos) {
        _pendingText.replace(p2Start, p2End - p2Start, filesStr);
    }
    if (p1Start != std::wstring::npos) {
        _pendingText.replace(p1Start, p1End - p1Start, hitsStr);
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

    ::RedrawWindow(_hSci, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);
}


// -----------------------------------------------------------------------------
// Insert one formatted file block *immediately* into the dock.
// Builds text for exactly one file (no SearchHdr) and prepends it as a block.
// After insertion we re-expand the previously-top block to neutralize the
// "collapse previous block" behaviour inside prependBlock().
// -----------------------------------------------------------------------------
void ResultDock::insertFileBlockNow(const FileMap& fm, const SciSendFn& sciSend)
{
    if (!_hSci) return;

    // Build per-file text & hits (no SearchHdr)
    std::wstring  partText;
    std::vector<Hit> partHits;
    buildListText(fm, _groupViewPending, L"", sciSend, partText, partHits);
    if (partText.empty()) return;

    // Count lines of the new block; check whether there was old content
    const bool hadOld = S(SCI_GETLENGTH) > 0;
    int newBlockLines = 0;
    for (wchar_t c : partText) if (c == L'\n') ++newBlockLines;

    // Prepend the block (will also style/fold this fragment and collapse previous block)
    prependBlock(partText, partHits);

    // If there was a previous block, prependBlock() collapsed it by design.
    // For per-file incremental commits we prefer keeping it expanded → re-expand it.
    if (hadOld) {
        const int firstLineOfOldBlock = newBlockLines + 1; // + separator line
        const int level = static_cast<int>(S(SCI_GETFOLDLEVEL, firstLineOfOldBlock));
        if (level & SC_FOLDLEVELHEADERFLAG) {
            S(SCI_SETFOLDEXPANDED, firstLineOfOldBlock, TRUE);
            S(SCI_FOLDCHILDREN, firstLineOfOldBlock, SC_FOLDACTION_EXPAND);
        }
    }
}

// -----------------------------------------------------------------------------
// Insert only the search header at the very top (final numbers), used after
// incremental per-file inserts. We also re-expand the now-second block so the
// last file block does not end up collapsed just because we added the header.
// -----------------------------------------------------------------------------
void ResultDock::insertSearchHeader(const std::wstring& header)
{
    if (!_hSci) return;

    // Build single header line (as in startSearchBlock)
    std::wstring hdr = getIndentString(LineLevel::SearchHdr) + header + L"\r\n";

    // Count header lines (always 1) and whether there was content before
    const bool hadOld = S(SCI_GETLENGTH) > 0;
    int hdrLines = 0; for (wchar_t c : hdr) if (c == L'\n') ++hdrLines;

    std::vector<Hit> none;
    prependBlock(hdr, none);

    // If there was existing content, the previous top block (a file block)
    // was just collapsed by prependBlock(). Immediately re-expand it.
    if (hadOld) {
        const int firstLineOfOldBlock = hdrLines + 1; // typically 1
        const int level = static_cast<int>(S(SCI_GETFOLDLEVEL, firstLineOfOldBlock));
        if (level & SC_FOLDLEVELHEADERFLAG) {
            S(SCI_SETFOLDEXPANDED, firstLineOfOldBlock, TRUE);
            S(SCI_FOLDCHILDREN, firstLineOfOldBlock, SC_FOLDACTION_EXPAND);
        }
    }

    // Optional: remember header line index if you use collapseOldSearches elsewhere
    _searchHeaderLines.insert(_searchHeaderLines.begin(), 0);
}

// ---------------- Construction & Core State ---------------

ResultDock::ResultDock(HINSTANCE hInst)
    : _hInst(hInst)      // member‑initialiser‑list
    , _hSci(nullptr)
    , _hDock(nullptr)
{
    // no further code needed here – everything else happens in create()
}

// Map control chars to a visible blank in our dock (display only).
static void applyControlRepresentations(HWND hSci)
{
    auto setRep = [&](const char* key, const char* rep)
        {
            ::SendMessage(hSci, SCI_SETREPRESENTATION, reinterpret_cast<WPARAM>(key), reinterpret_cast<LPARAM>(rep));
        };

    // C0 controls except \t, \n, \r
    for (int c = 0x00; c <= 0x1F; ++c)
    {
        if (c == '\t' || c == '\n' || c == '\r')
            continue;

        char key[2] = { (char)c, 0 };
        setRep(key, " ");
    }

    // DEL
    setRep("\x7F", " ");

    // C1 controls U+0080..U+009F -> UTF-8: C2 80 .. C2 9F
    for (int c = 0x80; c <= 0x9F; ++c)
    {
        unsigned char key[3] = { 0xC2u, (unsigned char)c, 0u };
        setRep(reinterpret_cast<const char*>(key), " ");
    }
}


void ResultDock::create(const NppData& npp)
{
    // 1) Create the Scintilla window that will become our dock client
    _hSci = ::CreateWindowExW(
        WS_EX_CLIENTEDGE,
        L"Scintilla", L"",
        WS_CHILD | WS_CLIPSIBLINGS,
        0, 0, 100, 100,
        npp._nppHandle, nullptr, _hInst, nullptr);

    if (!_hSci) {
        ::MessageBoxW(npp._nppHandle,
            L"FATAL: Failed to create Scintilla window!",
            L"ResultDock Error",
            MB_OK | MB_ICONERROR);
        return;
    }

    // This ensures colors and transparency match the main editor (DirectWrite vs GDI)
    int tech = static_cast<int>(::SendMessage(npp._scintillaMainHandle, SCI_GETTECHNOLOGY, 0, 0));
    ::SendMessage(_hSci, SCI_SETTECHNOLOGY, tech, 0);

    // Optional: Synchronize bidirectional text settings
    int bidi = static_cast<int>(::SendMessage(npp._scintillaMainHandle, SCI_GETBIDIRECTIONAL, 0, 0));
    ::SendMessage(_hSci, SCI_SETBIDIRECTIONAL, bidi, 0);

    // 2) Subclass Scintilla so we receive double-clicks and other messages
    s_prevSciProc = reinterpret_cast<WNDPROC>(
        ::SetWindowLongPtrW(_hSci, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(sciSubclassProc)));

    // 3) Fast buffer config for huge search results
    ::SendMessageW(_hSci, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
    ::SendMessageW(_hSci, SCI_SETUNDOCOLLECTION, FALSE, 0);   // no undo stack
    ::SendMessageW(_hSci, SCI_SETMODEVENTMASK, 0, 0);         // no mod events
    ::SendMessageW(_hSci, SCI_SETSCROLLWIDTHTRACKING, TRUE, 0);
    ::SendMessageW(_hSci, SCI_SETWRAPMODE, wrapEnabled() ? SC_WRAP_WORD : SC_WRAP_NONE, 0);

    // 3a) **Performance/caching**: keep layout cached for whole doc + buffered drawing
    ::SendMessageW(_hSci, SCI_SETBUFFEREDDRAW, TRUE, 0);
    ::SendMessageW(_hSci, SCI_SETLAYOUTCACHE, SC_CACHE_DOCUMENT, 0); // critical for fast fold open on older blocks
    applyControlRepresentations(_hSci);

    // 3b) DirectFunction (hot path)
    _sciFn = reinterpret_cast<SciFnDirect_t>(::SendMessage(_hSci, SCI_GETDIRECTFUNCTION, 0, 0));
    _sciPtr = static_cast<sptr_t>(::SendMessage(_hSci, SCI_GETDIRECTPOINTER, 0, 0));

    // 4) Fill docking descriptor
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

    UINT dpi = 96;
    if (HDC hdc = ::GetDC(npp._nppHandle)) {
        dpi = ::GetDeviceCaps(hdc, LOGPIXELSX);
        ::ReleaseDC(npp._nppHandle, hdc);
    }

    // Use monochrome tab icons
    _dockData.hIconTab = darkMode
        ? CreateIconFromImageData(gimp_image_tab_dark, dpi)
        : CreateIconFromImageData(gimp_image_tab_light, dpi);

    // 5) Register dock & theme it
    _hDock = reinterpret_cast<HWND>(::SendMessageW(npp._nppHandle, NPPM_DMMREGASDCKDLG, 0, reinterpret_cast<LPARAM>(&_dockData)));
    if (!_hDock) {
        ::MessageBoxW(npp._nppHandle, L"ERROR: Docking registration failed.", L"ResultDock Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Hide immediately after registration to prevent empty panel flash.
    // N++ does the same: display(false) right after NPPM_DMMREGASDCKDLG.
    // The caller shows the panel later via ensureCreatedAndVisible().
    ::SendMessage(npp._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(_hSci));

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

    // 7) enable mouse interaction - but disable auto-fold so we can handle it ourselves
    //    This allows us to control horizontal scroll position after fold toggle
    S(SCI_SETMARGINSENSITIVEN, M_FOLD, TRUE);
    S(SCI_SETAUTOMATICFOLD, 0);  // Disable automatic fold on click - handled in sciSubclassProc

    S(SCI_SETFOLDFLAGS, SC_FOLDFLAG_LINEAFTER_CONTRACTED);

    S(SCI_MARKERENABLEHIGHLIGHT, TRUE);
}

void ResultDock::applyTheme()
{
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
    S(SCI_STYLESETFONT, STYLE_DEFAULT, reinterpret_cast<LPARAM>("Consolas"));
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

    // Caret line (always visible, even when dock loses focus after navigation)
    S(SCI_SETCARETLINEVISIBLEALWAYS, TRUE, 0);
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

    // Red match color
    S(SCI_INDICSETSTYLE, INDIC_MATCH_FORE, INDIC_TEXTFORE);
    S(SCI_INDICSETFORE, INDIC_MATCH_FORE, theme.matchFg);
    S(SCI_INDICSETUNDER, INDIC_MATCH_FORE, TRUE);

    // Per-entry background color indicators (10 distinct colors)
    // Text color remains standard (matchFg), only background varies per entry
    // Determine alpha based on mode
    const int bgAlpha = dark ? ENTRY_BG_ALPHA_DARK : ENTRY_BG_ALPHA_LIGHT;
    const int outlineAlpha = dark ? ENTRY_OUTLINE_ALPHA_DARK : ENTRY_OUTLINE_ALPHA_LIGHT;

    for (int i = 0; i < MAX_ENTRY_COLORS; ++i) {
        const int indicId = INDIC_ENTRY_BG_BASE + i;
        S(SCI_INDICSETSTYLE, indicId, INDIC_ROUNDBOX);
        S(SCI_INDICSETALPHA, indicId, bgAlpha);
        S(SCI_INDICSETOUTLINEALPHA, indicId, outlineAlpha);
        S(SCI_INDICSETUNDER, indicId, TRUE);
    }

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

// -------- Range styling / folding (partial updates) -------

void ResultDock::applyStylingRange(Sci_Position pos0, Sci_Position len, const std::vector<Hit>& newHits) const
{
    if (!_hSci || len <= 0) return;

    const Sci_Position endPos = pos0 + len;
    const int firstLine = static_cast<int>(S(SCI_LINEFROMPOSITION, pos0));
    const int lastLine = static_cast<int>(S(SCI_LINEFROMPOSITION, endPos));

    // Base styling using indentation (no GETLINE)
    Sci_Position lineStartPos = S(SCI_POSITIONFROMLINE, firstLine);
    S(SCI_STARTSTYLING, lineStartPos, 0);

    const int IND_SRCH = INDENT_SPACES[(int)LineLevel::SearchHdr];
    const int IND_FILE = INDENT_SPACES[(int)LineLevel::FileHdr];
    const int IND_CRIT = INDENT_SPACES[(int)LineLevel::CritHdr];

    for (int line = firstLine; line <= lastLine; ++line) {
        const Sci_Position ls = S(SCI_POSITIONFROMLINE, line);
        const int          ll = static_cast<int>(S(SCI_LINELENGTH, line));

        int style = STYLE_DEFAULT;
        if (ll > 0) {
            const int indent = static_cast<int>(S(SCI_GETLINEINDENTATION, line));
            if (indent == IND_SRCH) style = STYLE_HEADER;
            else if (indent == IND_CRIT) style = STYLE_CRITHDR;
            else if (indent == IND_FILE) style = STYLE_FILEPATH;
        }
        if (ll > 0) S(SCI_SETSTYLING, ll, style);

        // Keep EOL visuals aligned with default
        const Sci_Position lineEnd = S(SCI_GETLINEENDPOSITION, line);
        const int          eolLen = (int)(lineEnd - (ls + ll));
        if (eolLen > 0) S(SCI_SETSTYLING, eolLen, STYLE_DEFAULT);
    }

    // Indicators only on the freshly added hits
    S(SCI_SETINDICATORCURRENT, INDIC_LINE_BACKGROUND);
    for (const auto& h : newHits) {
        if (h.displayLineStart < 0) continue;
        const int          line = static_cast<int>(S(SCI_LINEFROMPOSITION, h.displayLineStart));
        const Sci_Position ls = S(SCI_POSITIONFROMLINE, line);
        const Sci_Position ll = S(SCI_LINELENGTH, line);
        if (ll > 0) S(SCI_INDICATORFILLRANGE, ls, ll);
    }

    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& h : newHits)
        if (h.displayLineStart >= 0)
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.numberStart, h.numberLen);

    // 3c/3d) Match Highlighting (Exclusive Logic for Partial Updates)
    if (_perEntryColorsEnabled) {
        // CASE A: Colorful Backgrounds -> Apply ONLY background indicators (Text remains standard/white)
        for (const auto& h : newHits) {
            if (h.displayLineStart < 0) continue;
            for (size_t i = 0; i < h.matchStarts.size(); ++i) {
                const int colorIdx = (i < h.matchColors.size())
                    ? h.matchColors[i]
                    : h.colorIndex;  // fallback

                if (colorIdx >= 0 && colorIdx < MAX_ENTRY_COLORS) {
                    S(SCI_SETINDICATORCURRENT, INDIC_ENTRY_BG_BASE + colorIdx);
                    S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.matchStarts[i], h.matchLens[i]);
                }
            }
        }
    }
    else {
        // CASE B: Standard Mode -> Apply ONLY text color indicator (e.g. Orange/Green)
        S(SCI_SETINDICATORCURRENT, INDIC_MATCH_FORE);
        for (const auto& h : newHits) {
            if (h.displayLineStart < 0) continue;
            for (size_t i = 0; i < h.matchStarts.size(); ++i) {
                S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.matchStarts[i], h.matchLens[i]);
            }
        }
    }
}

void ResultDock::rebuildFoldingRange(int firstLine, int lastLine) const
{
    if (!_hSci || lastLine < firstLine) return;

    const int BASE = SC_FOLDLEVELBASE;
    const int IND_SRCH = INDENT_SPACES[(int)LineLevel::SearchHdr];
    const int IND_FILE = INDENT_SPACES[(int)LineLevel::FileHdr];
    const int IND_CRIT = INDENT_SPACES[(int)LineLevel::CritHdr];
    const int IND_HIT = INDENT_SPACES[(int)LineLevel::HitLine];

    auto calcLevel = [&](int indent, bool& isHeader) -> int {
        if (indent == IND_SRCH) { isHeader = true; return BASE + (int)LineLevel::SearchHdr; }
        if (indent == IND_FILE) { isHeader = true; return BASE + (int)LineLevel::FileHdr; }
        if (indent == IND_CRIT) { isHeader = true; return BASE + (int)LineLevel::CritHdr; }
        isHeader = false;
        if (indent >= IND_HIT)  return BASE + (int)LineLevel::HitLine;
        if (indent >= IND_CRIT) return BASE + (int)LineLevel::CritHdr;
        if (indent >= IND_FILE) return BASE + (int)LineLevel::FileHdr;
        return BASE + 1;
        };

    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);

    for (int l = firstLine; l <= lastLine; ++l) {
        // Treat whitespace-only like blank
        const int lineLen = static_cast<int>(S(SCI_LINELENGTH, l));
        int target = BASE;
        if (lineLen > 0 && hasContentBeyondIndent(_hSci, l)) {
            const int indent = static_cast<int>(S(SCI_GETLINEINDENTATION, l));
            bool isHdr = false;
            const int lvl = calcLevel(indent, isHdr);
            target = isHdr ? (lvl | SC_FOLDLEVELHEADERFLAG) : lvl;
        }

        const int cur = static_cast<int>(S(SCI_GETFOLDLEVEL, l));
        if (cur != target) S(SCI_SETFOLDLEVEL, l, target);
    }

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);
}

// ---------------- Block building / insertion --------------

void ResultDock::prependBlock(const std::wstring& dockText, std::vector<Hit>& newHits)
{
    if (!_hSci || dockText.empty())
        return;

    const std::string u8 = Encoding::wstringToUtf8(dockText);
    const Sci_Position oldLen = (Sci_Position)S(SCI_GETLENGTH);

    const int sepBytes = (oldLen > 0 ? 2 : 0);
    const int deltaBytes = (int)u8.size() + sepBytes;

    // Shift existing hits
    for (auto& h : _hits)
        h.displayLineStart += deltaBytes;

    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);
    S(SCI_SETREADONLY, FALSE);
    S(SCI_ALLOCATE, oldLen + (Sci_Position)deltaBytes + 65536, 0);

    // Move cursor to position 0 before inserting to prevent auto-scroll to old cursor position
    S(SCI_SETEMPTYSELECTION, 0, 0);

    // Insert new block at start
    S(SCI_INSERTTEXT, 0, (sptr_t)u8.c_str());

    // Defensive: recompute length before inserting separator
    if (sepBytes) {
        const Sci_Position lenAfterBlock = (Sci_Position)S(SCI_GETLENGTH);
        const Sci_Position sepPos = (Sci_Position)u8.size();
        if (sepPos <= lenAfterBlock) {
            S(SCI_INSERTTEXT, (uptr_t)sepPos, (sptr_t)"\r\n");
        }
        // If that doesn't fit, no insert -> avoids invalid position.
    }

    S(SCI_SETREADONLY, TRUE);

    // Reset horizontal scroll to prevent scrolling to old cursor position on long lines
    S(SCI_SETXOFFSET, 0, 0);

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);

    const Sci_Position pos0 = 0;
    const Sci_Position len = (Sci_Position)u8.size();

    int newBlockLines = 0;
    for (char c : u8) if (c == '\n') ++newBlockLines;
    const int firstLine = 0;
    const int lastLine = (newBlockLines > 0 ? newBlockLines - 1 : 0);

    rebuildFoldingRange(firstLine, lastLine);
    applyStylingRange(pos0, len, newHits);

    if (oldLen > 0) {
        const int firstLineOfOldBlock = newBlockLines + 1;
        const int level = static_cast<int>(S(SCI_GETFOLDLEVEL, firstLineOfOldBlock));
        if ((level & SC_FOLDLEVELHEADERFLAG) && S(SCI_GETFOLDEXPANDED, firstLineOfOldBlock))
            S(SCI_FOLDLINE, firstLineOfOldBlock, SC_FOLDACTION_CONTRACT);
    }

    _hits.insert(_hits.begin(),
        std::make_move_iterator(newHits.begin()),
        std::make_move_iterator(newHits.end()));

    rebuildHitLineIndex();

    // Ensure view starts at top-left after new block is inserted
    S(SCI_SETFIRSTVISIBLELINE, 0, 0);
    S(SCI_SETXOFFSET, 0, 0);
}

void ResultDock::collapseOldSearches()
{
    if (!_hSci || _searchHeaderLines.size() < 2)
        return;

    /// Collapse all previous search blocks; keep only the most recent expanded.
    const size_t lastIdx = _searchHeaderLines.size() - 1;
    for (size_t i = 0; i < lastIdx; ++i) {
        const int headerLine = _searchHeaderLines[i];
        /// Safety check: ensure headerLine is within the current line count.
        const int lineCount = static_cast<int>(S(SCI_GETLINECOUNT));
        if (headerLine >= 0 && headerLine < lineCount) {
            S(SCI_SETFOLDEXPANDED, headerLine, FALSE);
            S(SCI_FOLDCHILDREN, headerLine, SC_FOLDACTION_CONTRACT);
        }
    }
}

// ---------------------- Formatting ------------------------

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

                // make a mutable copy of c.hits (colorIndex already set per-hit)
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
            for (const auto& c : f.crits) {
                // Make a copy and assign colorIndex from CritAgg
                for (const auto& hit : c.hits) {
                    Hit hitCopy = hit;
                    // colorIndex is already set per-hit based on matched text
                    merged.push_back(std::move(hitCopy));
                }
            }
            // Sort by position
            std::sort(merged.begin(), merged.end(),
                [](auto const& a, auto const& b) {
                    if (a.pos != b.pos) return a.pos < b.pos;
                    if (a.length != b.length) return a.length < b.length;

                    return a.findTextW < b.findTextW;
                });

            // Remove duplicate hits at same position (keep first occurrence)
            merged.erase(
                std::unique(merged.begin(), merged.end(),
                    [](const Hit& a, const Hit& b) { return a.pos == b.pos && a.length == b.length; }),
                merged.end());

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
    const bool isUtf8Doc = (docCp == SC_CP_UTF8);

    const std::wstring indentHitW = getIndentString(LineLevel::HitLine);
    const size_t       indentHitU8 = getIndentUtf8Length(LineLevel::HitLine);

    const std::wstring kLineW = LM.get(L"dock_line") + L" ";
    const size_t kLineU8 = Encoding::wstringToUtf8(kLineW).size();  // No static: supports runtime language change
    constexpr size_t    kColonSpaceU8 = 2; // ": "

    constexpr size_t kMaxHitTextUtf8 = 2048; // cap display text to limit memory/rendering cost on very long lines

    auto capUtf8WithEllipsis = [](std::string& s, size_t maxBytes) {
        if (s.size() <= maxBytes) return;
        if (maxBytes <= 3) { s.assign("..."); return; }
        size_t cut = maxBytes - 3;
        while (cut > 0 && (static_cast<unsigned char>(s[cut]) & 0xC0) == 0x80) --cut;
        s.resize(cut);
        s.append("...");
        };

    std::vector<int> lineNumbers; lineNumbers.reserve(hits.size());
    size_t maxDigits = 0;
    for (const Hit& h : hits) {
        int line1 = static_cast<int>(sciSend(SCI_LINEFROMPOSITION, h.pos, 0)) + 1;
        lineNumbers.push_back(line1);
        // Count digits without creating temporary string
        int temp = line1;
        size_t digits = 0;
        do { ++digits; temp /= 10; } while (temp > 0);
        if (digits > maxDigits) maxDigits = digits;
    }

    // Helper to count digits in a number (avoids std::to_wstring allocation)
    auto countDigits = [](int n) -> size_t {
        if (n <= 0) return 1;
        size_t count = 0;
        while (n > 0) { ++count; n /= 10; }
        return count;
        };

    auto padNumber = [](int number, size_t width) -> std::wstring {
        std::wstring numStr = std::to_wstring(number);
        size_t pad = (width > numStr.size() ? width - numStr.size() : 0);
        return std::wstring(pad, L' ') + numStr;
        };

    int   prevDocLine = -1;
    Hit* firstHitOnRow = nullptr;
    size_t hitIdx = 0;
    size_t currentRowLenU8 = 0;

    std::string  cachedRaw;
    std::string  cachedRawFiltered;  // Only populated when FlowTabs padding exists
    std::string  origU8;
    std::string  displayU8;
    std::wstring displayW;
    Sci_Position cachedAbsLineStart = 0;

    // Pointer to the effective raw bytes (either cachedRaw or cachedRawFiltered)
    const std::string* effectiveRaw = &cachedRaw;

    // FlowTabs filtering: mapping from original raw byte index to filtered byte index
    // Only populated when the document actually has FlowTabs padding.
    std::vector<size_t> mapRawToFiltered;

    // FlowTabs filtering: only scan when the document actually has padding.
    const int flowTabsIndicatorId = ColumnTabs::CT_GetIndicatorId();
    bool docHasFlowTabPadding = false;
    if (flowTabsIndicatorId >= 0) {
        // Quick check: does the document have any indicator at all?
        const Sci_Position docLen = static_cast<Sci_Position>(sciSend(SCI_GETLENGTH, 0, 0));
        if (docLen > 0) {
            const Sci_Position firstEnd = static_cast<Sci_Position>(
                sciSend(SCI_INDICATOREND, static_cast<WPARAM>(flowTabsIndicatorId), 0));
            const int firstVal = static_cast<int>(
                sciSend(SCI_INDICATORVALUEAT, static_cast<WPARAM>(flowTabsIndicatorId), 0));
            // firstVal != 0: padding starts at pos 0.
            // firstEnd check: padding may start after pos 0, so detect any range boundary.
            docHasFlowTabPadding = (firstVal != 0) || (firstEnd > 0 && firstEnd < docLen);
        }
    }

    std::vector<size_t> mapOrigToDisp;
    std::unordered_map<size_t, size_t> u8PrefixLenByByte;

    auto loadLineIfNeeded = [&](int line0)
        {
            if (line0 == prevDocLine)
                return;

            // raw line
            const int rawLen = static_cast<int>(sciSend(SCI_LINELENGTH, line0, 0));
            cachedRaw.assign(rawLen + 1, '\0');
            sciSend(SCI_GETLINE, line0, reinterpret_cast<LPARAM>(cachedRaw.data()));
            cachedRaw.resize(rawLen);
            while (!cachedRaw.empty() && (cachedRaw.back() == '\r' || cachedRaw.back() == '\n'))
                cachedRaw.pop_back();

            cachedAbsLineStart = sciSend(SCI_POSITIONFROMLINE, line0, 0);

            // Filter out FlowTabs padding using range-based indicator scan
            mapRawToFiltered.clear();
            if (docHasFlowTabPadding && !cachedRaw.empty()) {
                cachedRawFiltered.clear();
                cachedRawFiltered.reserve(cachedRaw.size());
                mapRawToFiltered.reserve(cachedRaw.size() + 1);

                // Build a set of padding ranges on this line using SCI_INDICATOREND
                // (O(ranges) Scintilla calls instead of O(bytes))
                const Sci_Position lineEndDoc = cachedAbsLineStart + static_cast<Sci_Position>(cachedRaw.size());
                std::vector<std::pair<Sci_Position, Sci_Position>> padRanges;
                {
                    Sci_Position scanPos = cachedAbsLineStart;
                    while (scanPos < lineEndDoc) {
                        const Sci_Position rangeEnd = static_cast<Sci_Position>(
                            sciSend(SCI_INDICATOREND, static_cast<WPARAM>(flowTabsIndicatorId),
                                static_cast<LPARAM>(scanPos)));
                        if (rangeEnd <= scanPos) break;

                        const int val = static_cast<int>(
                            sciSend(SCI_INDICATORVALUEAT, static_cast<WPARAM>(flowTabsIndicatorId),
                                static_cast<LPARAM>(scanPos)));
                        if (val != 0) {
                            // This range [scanPos..rangeEnd) is padding
                            const Sci_Position clampedEnd = (rangeEnd < lineEndDoc) ? rangeEnd : lineEndDoc;
                            padRanges.emplace_back(scanPos, clampedEnd);
                        }
                        scanPos = rangeEnd;
                    }
                }

                if (padRanges.empty()) {
                    // No padding on this line - use cachedRaw directly
                    effectiveRaw = &cachedRaw;
                }
                else {
                    // Filter using ranges (no per-byte Scintilla calls)
                    size_t padIdx = 0;
                    for (size_t i = 0; i < cachedRaw.size(); ++i) {
                        const Sci_Position docPos = cachedAbsLineStart + static_cast<Sci_Position>(i);

                        // Advance past ranges that end before this position
                        while (padIdx < padRanges.size() && padRanges[padIdx].second <= docPos)
                            ++padIdx;

                        const bool isPadding = (padIdx < padRanges.size()
                            && docPos >= padRanges[padIdx].first
                            && docPos < padRanges[padIdx].second);

                        mapRawToFiltered.push_back(cachedRawFiltered.size());
                        if (!isPadding) {
                            cachedRawFiltered.push_back(cachedRaw[i]);
                        }
                    }
                    mapRawToFiltered.push_back(cachedRawFiltered.size());
                    effectiveRaw = &cachedRawFiltered;
                }
            }
            else {
                // No FlowTabs padding - use raw directly, no mapping needed
                effectiveRaw = &cachedRaw;
            }

            // Build origU8 — for UTF-8 docs the raw bytes are already UTF-8
            if (isUtf8Doc)
                origU8 = *effectiveRaw;
            else
                origU8 = Encoding::bytesToUtf8(*effectiveRaw, docCp);

            // classify helpers
            auto isCtlWide = [](wchar_t ch)->bool {
                const unsigned u = (unsigned)ch;
                if (u == 0x0000) return true;                            // NUL
                if ((u <= 0x1Fu) && ch != L'\t' && ch != L'\n' && ch != L'\r') return true; // C0
                if (u == 0x007Fu) return true;                            // DEL
                if (u >= 0x0080u && u <= 0x009Fu) return true;            // C1 (incl. NEL)
                if (u == 0x2028 || u == 0x2029) return true;              // LS/PS
                return false;
                };
            auto isCtlCp = [](uint32_t cp)->bool {
                if (cp == 0x0000) return true;
                if ((cp <= 0x1Fu) && cp != 0x09 && cp != 0x0A && cp != 0x0D) return true;
                if (cp == 0x007Fu) return true;
                if (cp >= 0x80u && cp <= 0x9Fu) return true;
                if (cp == 0x2028 || cp == 0x2029) return true;
                return false;
                };
            auto isNbspWide = [](wchar_t ch)->bool { return (unsigned)ch == 0x00A0; };
            auto isNbspCp = [](uint32_t cp)->bool { return cp == 0x00A0; };

            // build display (safe replacement on Wide)
            std::wstring wide;
            {
                const UINT cp = (UINT)docCp;
                const UINT wcp = (cp == 0 ? CP_ACP : cp);
                if (cp == SC_CP_UTF8) {
                    wide = Encoding::utf8ToWString(*effectiveRaw);
                }
                else {
                    int wlen = ::MultiByteToWideChar(wcp, MB_ERR_INVALID_CHARS,
                        effectiveRaw->data(), (int)effectiveRaw->size(), nullptr, 0);
                    if (wlen <= 0)
                        wlen = ::MultiByteToWideChar(wcp, 0,
                            effectiveRaw->data(), (int)effectiveRaw->size(), nullptr, 0);
                    wide.resize((size_t)wlen);
                    if (wlen > 0)
                        ::MultiByteToWideChar(wcp, 0,
                            effectiveRaw->data(), (int)effectiveRaw->size(),
                            &wide[0], wlen);
                }
            }

            std::wstring wideClean;
            wideClean.reserve(wide.size());
            for (wchar_t ch : wide) {
                const unsigned u = (unsigned)ch;
                if (u == 0x00AD) {
                    // SHY → drop
                    continue;
                }
                if (isNbspWide(ch) || isCtlWide(ch)) {
                    wideClean.push_back(L' ');   // one ASCII space per codepoint
                }
                else {
                    wideClean.push_back(ch);
                }
            }
            displayU8 = Encoding::wstringToUtf8(wideClean);

            // mapping origU8 → displayU8 (exact same rule)
            mapOrigToDisp.assign(origU8.size() + 1, 0);

            size_t o = 0, d = 0;
            while (o < origU8.size())
            {
                unsigned char c = (unsigned char)origU8[o];
                size_t clen;
                if ((c & 0x80) == 0x00) clen = 1;
                else if ((c & 0xE0) == 0xC0 && o + 1 < origU8.size()) clen = 2;
                else if ((c & 0xF0) == 0xE0 && o + 2 < origU8.size()) clen = 3;
                else if ((c & 0xF8) == 0xF0 && o + 3 < origU8.size()) clen = 4;
                else clen = 1;

                uint32_t cp = 0;
                if (clen == 1) cp = c;
                else if (clen == 2) cp = ((c & 0x1Fu) << 6) | (uint32_t(origU8[o + 1]) & 0x3Fu);
                else if (clen == 3) cp = ((c & 0x0Fu) << 12) |
                    ((uint32_t(origU8[o + 1]) & 0x3Fu) << 6) |
                    (uint32_t(origU8[o + 2]) & 0x3Fu);
                else cp = ((c & 0x07u) << 18) |
                    ((uint32_t(origU8[o + 1]) & 0x3Fu) << 12) |
                    ((uint32_t(origU8[o + 2]) & 0x3Fu) << 6) |
                    (uint32_t(origU8[o + 3]) & 0x3Fu);

                if (cp == 0x00AD) {
                    // SHY drop
                    for (size_t k = 0; k < clen && (o + k) < origU8.size(); ++k)
                        mapOrigToDisp[o + k] = d;
                    // no d++
                }
                else if (isNbspCp(cp) || isCtlCp(cp)) {
                    for (size_t k = 0; k < clen && (o + k) < origU8.size(); ++k)
                        mapOrigToDisp[o + k] = d;
                    ++d; // one visible ASCII space per replaced codepoint
                }
                else {
                    for (size_t k = 0; k < clen && (o + k) < origU8.size(); ++k)
                        mapOrigToDisp[o + k] = d + k;
                    d += clen;
                }

                o += clen;
            }
            if (!mapOrigToDisp.empty()) mapOrigToDisp.back() = d;

            // finalize
            capUtf8WithEllipsis(displayU8, kMaxHitTextUtf8);
            displayW = Encoding::utf8ToWString(displayU8);

            u8PrefixLenByByte.clear();
            // prevDocLine is set by the caller when the row is actually written
        };

    // Convert raw byte position to filtered byte position
    // When no padding is present (effectiveRaw == &cachedRaw), this is identity.
    auto rawToFilteredPos = [&](size_t rawPos) -> size_t {
        if (effectiveRaw == &cachedRaw)
            return rawPos;  // No padding: identity mapping
        if (rawPos >= mapRawToFiltered.size())
            return mapRawToFiltered.empty() ? 0 : mapRawToFiltered.back();
        return mapRawToFiltered[rawPos];
        };

    auto u8PrefixFromRawBytes = [&](size_t byteCount)->size_t {
        if (isUtf8Doc) return rawToFilteredPos(byteCount); // byte pos == UTF-8 pos
        auto it = u8PrefixLenByByte.find(byteCount);
        if (it != u8PrefixLenByByte.end()) return it->second;
        const size_t filteredByteCount = rawToFilteredPos(byteCount);
        size_t val = Encoding::bytesToUtf8(effectiveRaw->data(), filteredByteCount, docCp).size();
        u8PrefixLenByByte.emplace(byteCount, val);
        return val;
        };

    for (Hit& h : hits) {
        int line1 = lineNumbers[hitIdx++];
        int line0 = line1 - 1;

        loadLineIfNeeded(line0);

        size_t relBytes = (size_t)(h.pos - cachedAbsLineStart);
        size_t hitStartU8_orig = u8PrefixFromRawBytes(relBytes);

        // Calculate hit length in filtered bytes
        const size_t filteredStart = rawToFilteredPos(relBytes);
        const size_t filteredEnd = rawToFilteredPos(relBytes + h.length);
        const size_t filteredLen = (filteredEnd > filteredStart) ? (filteredEnd - filteredStart) : 0;
        const size_t hitLenU8_orig = isUtf8Doc
            ? filteredLen
            : Encoding::bytesToUtf8(effectiveRaw->data() + filteredStart, filteredLen, docCp).size();

        auto mapToDisp = [&](size_t offU8) -> size_t {
            if (offU8 >= mapOrigToDisp.size()) return mapOrigToDisp.empty() ? 0 : mapOrigToDisp.back();
            return mapOrigToDisp[offU8];
            };
        size_t dispStart = mapToDisp(hitStartU8_orig);
        size_t dispEnd = mapToDisp(hitStartU8_orig + hitLenU8_orig);
        size_t dispLen = (dispEnd > dispStart ? dispEnd - dispStart : 0);


        if (line0 != prevDocLine) {
            std::wstring paddedNumW = padNumber(line1, maxDigits);
            std::wstring prefixW = indentHitW + kLineW + paddedNumW + L": ";
            const size_t prefixU8Len = indentHitU8 + kLineU8 + paddedNumW.size() + kColonSpaceU8;

            currentRowLenU8 = displayU8.size();
            out += prefixW + displayW + L"\r\n";

            const size_t line1Digits = countDigits(line1);
            h.displayLineStart = (int)utf8Pos;
            h.numberStart = (int)(indentHitU8 + kLineU8 + paddedNumW.size() - line1Digits);
            h.numberLen = (int)line1Digits;

            h.matchStarts.clear();
            h.matchLens.clear();
            h.matchColors.clear();

            firstHitOnRow = &h;
            firstHitOnRow->allFindTexts.clear();
            firstHitOnRow->allFindTexts.push_back(h.findTextW);
            firstHitOnRow->allPositions.clear();
            firstHitOnRow->allPositions.push_back(h.pos);
            firstHitOnRow->allLengths.clear();
            firstHitOnRow->allLengths.push_back(h.length);

            if (dispStart < displayU8.size()) {
                size_t safeLen = dispLen;
                if (dispStart + safeLen > displayU8.size()) safeLen = displayU8.size() - dispStart;
                if (safeLen > 0) {
                    h.matchStarts.push_back((int)(prefixU8Len + dispStart));
                    h.matchLens.push_back((int)safeLen);

                    h.matchColors.push_back(h.colorIndex);
                }
            }

            utf8Pos += prefixU8Len + currentRowLenU8 + 2;
            prevDocLine = line0;
        }
        else {
            if (firstHitOnRow) {
                const size_t line1Digits = countDigits(line1);
                const size_t prefixU8Len = indentHitU8 + kLineU8
                    + line1Digits
                    + (maxDigits - line1Digits)
                    + kColonSpaceU8;

                if (dispStart < displayU8.size()) {
                    size_t safeLen = dispLen;
                    if (dispStart + safeLen > displayU8.size()) safeLen = displayU8.size() - dispStart;
                    if (safeLen > 0) {
                        firstHitOnRow->matchStarts.push_back((int)(prefixU8Len + dispStart));
                        firstHitOnRow->matchLens.push_back((int)safeLen);

                        firstHitOnRow->matchColors.push_back(h.colorIndex);
                    }
                }
                firstHitOnRow->allFindTexts.push_back(h.findTextW);
                firstHitOnRow->allPositions.push_back(h.pos);
                firstHitOnRow->allLengths.push_back(h.length);
            }
            h.displayLineStart = -1;
        }
    }

    hits.erase(std::remove_if(hits.begin(), hits.end(),
        [](const Hit& e) { return e.displayLineStart < 0; }),
        hits.end());

}



// --------------------- Line helpers -----------------------

ResultDock::LineKind ResultDock::classify(const std::string& raw) {
    size_t spaces = 0;
    while (spaces < raw.size() && raw[spaces] == ' ') ++spaces;

    std::string_view trimmed(raw.data() + spaces, raw.size() - spaces);
    if (trimmed.empty()) return LineKind::Blank;

    // Fully qualify level by matching exact indent widths.
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
    size_t i = (std::min)(static_cast<size_t>(indentLen), w.size()); // skip leading indent

    // optional extra padding right after indent
    while (i < w.size() && iswspace(w[i])) ++i;

    // skip the localized token ("Line", "Zeile", …)
    while (i < w.size() && iswalpha(w[i])) ++i;

    // allow *multiple* spaces between token and number
    while (i < w.size() && iswspace(w[i])) ++i;

    // read the line number
    size_t digitStart = i;
    while (i < w.size() && iswdigit(w[i])) ++i;

    // allow *multiple* spaces between number and ':'
    while (i < w.size() && iswspace(w[i])) ++i;

    if (i == digitStart || i >= w.size() || w[i] != L':')
        return w; // not a formatted hit line → return as-is

    ++i; // skip ':'

    // Keep any original leading spaces/tabs of the source line intact.
    if (i < w.size() && w[i] == L' ') ++i;

    // return the original line content (incl. original EOLs), prefix removed
    return w.substr(i);
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
    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };

    for (int l = startLine; l >= 0; --l) {
        int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
        std::string raw(len, '\0');
        Sx(SCI_GETLINE, l, (sptr_t)raw.data());
        raw.resize(strnlen(raw.c_str(), len));
        if (classify(raw) == LineKind::FileHdr) return l;
    }
    return -1;
}

std::wstring ResultDock::getLineW(HWND hSci, int line) {
    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };
    int len = static_cast<int>(Sx(SCI_LINELENGTH, line));
    std::string raw(len, '\0');
    Sx(SCI_GETLINE, line, (sptr_t)raw.data());
    raw.resize(strnlen(raw.c_str(), len));
    return Encoding::utf8ToWString(raw);
}

int ResultDock::leadingSpaces(const char* line, int len)
{
    int s = 0;
    while (s < len && line[s] == ' ') ++s;
    return s;
}

// -------------------- Path/File helpers -------------------

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
    wchar_t buf[MAX_PATH] = {};
    DWORD n = ::GetModuleFileNameW(nullptr, buf, MAX_PATH);
    std::wstring exePath = (n ? std::wstring(buf, n) : std::wstring());
    const size_t pos = exePath.find_last_of(L"\\/");
    return (pos == std::wstring::npos) ? L"." : exePath.substr(0, pos);
}

bool ResultDock::IsCurrentDocByFullPath(const std::wstring& fullPath)
{
    wchar_t cur[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, static_cast<WPARAM>(MAX_PATH), reinterpret_cast<LPARAM>(cur));
    return (_wcsicmp(cur, fullPath.c_str()) == 0);
}

bool ResultDock::IsCurrentDocByTitle(const std::wstring& titleOnly)
{
    // Compare against current tab title (filename only)
    wchar_t name[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFILENAME, static_cast<WPARAM>(MAX_PATH), reinterpret_cast<LPARAM>(name));
    return (_wcsicmp(name, titleOnly.c_str()) == 0);
}

void ResultDock::SwitchToFileIfOpenByFullPath(const std::wstring& fullPath)
{
    // NPP reuses an existing tab if the file is open; otherwise: no effect.
    ::SendMessage(nppData._nppHandle, NPPM_SWITCHTOFILE, 0, reinterpret_cast<LPARAM>(fullPath.c_str()));
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
        ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(desiredPath.c_str()));
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
    const LRESULT ok = ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(desiredPath.c_str()));
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
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, reinterpret_cast<LPARAM>(&whichView));
    HWND hEd = (whichView == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
    if (!hEd)
        return;

    ::SetFocus(hEd);
    ::SendMessage(hEd, SCI_SETFOCUS, TRUE, 0);

    // Normalize match range
    Sci_Position startPos = pos;
    Sci_Position endPos = (len > 0) ? (pos + len) : pos;
    if (endPos < startPos)
        std::swap(startPos, endPos);

    // Make sure target lines are visible (unfold / scroll into view)
    const Sci_Position startLine = (Sci_Position)::SendMessage(hEd, SCI_LINEFROMPOSITION, startPos, 0);
    const Sci_Position endLine = (Sci_Position)::SendMessage(hEd, SCI_LINEFROMPOSITION, endPos, 0);
    ::SendMessage(hEd, SCI_ENSUREVISIBLE, startLine, 0);
    ::SendMessage(hEd, SCI_ENSUREVISIBLE, endLine, 0);

    // Set selection on the match
    ::SendMessage(hEd, SCI_GOTOPOS, startPos, 0);
    ::SendMessage(hEd, SCI_SETSEL, startPos, endPos);

    // Ensure full range is visible horizontally as well
    ::SendMessage(hEd, SCI_SCROLLRANGE, static_cast<WPARAM>(startPos), static_cast<LPARAM>(endPos));

    // --- Vertical centering (wrap-aware) ---
    // Use y-position of the match relative to the current viewport.
    const int yCaret = static_cast<int>(::SendMessage(hEd, SCI_POINTYFROMPOSITION, 0, static_cast<LPARAM>(startPos)));
    int lineHeight = static_cast<int>(::SendMessage(hEd, SCI_TEXTHEIGHT, 0, 0));
    int linesOnScr = static_cast<int>(::SendMessage(hEd, SCI_LINESONSCREEN, 0, 0));

    if (lineHeight <= 0) lineHeight = 1;
    if (linesOnScr <= 0) linesOnScr = 1;

    // Current visible subline index of the match
    const int sublineFromTop = yCaret / lineHeight;
    const int middleSubline = linesOnScr / 2;
    const int deltaSubLines = sublineFromTop - middleSubline;

    if (deltaSubLines != 0)
    {
        // Scroll by sublines; Scintilla treats the second arg as display lines (wrap-aware).
        ::SendMessage(hEd, SCI_LINESCROLL, 0, static_cast<WPARAM>(deltaSubLines));
    }

    // Keep caret selection consistent after scroll
    ::SendMessage(hEd, SCI_GOTOPOS, endPos, 0);
    ::SendMessage(hEd, SCI_SETANCHOR, startPos, 0);
    ::SendMessage(hEd, SCI_CHOOSECARETX, 0, 0);
}

void ResultDock::NavigateToHit(const Hit& hit)
{
    // Get active Scintilla
    int whichView = 0;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, reinterpret_cast<LPARAM>(&whichView));
    HWND hEd = (whichView == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
    if (!hEd)
        return;

    // ---- Line-based re-search: find the match closest to stored position ----
    if (hit.docLine < 0) {
        JumpSelectCenterActiveEditor(hit.pos, hit.length);
        return;
    }

    const int lineCount = static_cast<int>(::SendMessage(hEd, SCI_GETLINECOUNT, 0, 0));
    if (hit.docLine >= lineCount) {
        JumpSelectCenterActiveEditor(hit.pos, hit.length);
        return;
    }

    const Sci_Position lineStart = ::SendMessage(hEd, SCI_POSITIONFROMLINE, hit.docLine, 0);
    const Sci_Position lineEnd = ::SendMessage(hEd, SCI_GETLINEENDPOSITION, hit.docLine, 0);

    if (hit.findTextW.empty()) {
        JumpSelectCenterActiveEditor(hit.pos, hit.length);
        return;
    }

    const int docCp = static_cast<int>(::SendMessage(hEd, SCI_GETCODEPAGE, 0, 0));
    const std::string findBytes = Encoding::wstringToBytes(hit.findTextW, docCp);
    if (findBytes.empty()) {
        JumpSelectCenterActiveEditor(hit.pos, hit.length);
        return;
    }

    // Find ALL matches on the line and pick the one closest to stored position
    ::SendMessage(hEd, SCI_SETSEARCHFLAGS, hit.searchFlags, 0);

    Sci_Position bestPos = -1;
    Sci_Position bestLen = 0;
    Sci_Position bestDistance = 0x7FFFFFFF;

    Sci_Position searchFrom = lineStart;
    while (searchFrom < lineEnd) {
        ::SendMessage(hEd, SCI_SETTARGETRANGE, searchFrom, lineEnd);
        const Sci_Position found = ::SendMessage(hEd, SCI_SEARCHINTARGET,
            findBytes.size(), reinterpret_cast<LPARAM>(findBytes.c_str()));

        if (found < 0) break;

        const Sci_Position foundEnd = ::SendMessage(hEd, SCI_GETTARGETEND, 0, 0);
        const Sci_Position foundLen = foundEnd - found;

        Sci_Position distance = (found > hit.pos) ? (found - hit.pos) : (hit.pos - found);
        if (distance < bestDistance) {
            bestDistance = distance;
            bestPos = found;
            bestLen = foundLen;
        }

        searchFrom = foundEnd;
        if (searchFrom <= found) break;
    }

    if (bestPos >= 0) {
        JumpSelectCenterActiveEditor(bestPos, bestLen);
    }
    else {
        JumpSelectCenterActiveEditor(hit.pos, hit.length);
    }
}

// ------------------- FlowTab Position Adjustment -----------

void ResultDock::adjustHitPositionsForFlowTab(
    const std::string& filePathUtf8,
    const std::vector<std::pair<Sci_Position, Sci_Position>>& paddingRanges,
    bool added)
{
    if (paddingRanges.empty() || filePathUtf8.empty())
        return;

    // paddingRanges: sorted by position, pairs of (start, end) for each padding region.
    //
    // REMOVAL (added=false):
    //   paddingRanges are in pre-removal coordinates (same as hit.pos).
    //   For each hit, count total padding bytes before hit.pos, then subtract.
    //
    // INSERTION (added=true):
    //   paddingRanges are in post-insertion coordinates.
    //   hit.pos is in pre-insertion coordinates (no pads yet).
    //   We track cumulative delta: a padding at post-pos P corresponds to
    //   original pos (P - cumulativeDeltaSoFar). If that original pos <= hit.pos,
    //   this padding shifts the hit forward.

    auto computeDelta = [&](Sci_Position hitPos) -> Sci_Position {
        Sci_Position delta = 0;
        if (added) {
            // Post-insertion ranges: need to find how many pad bytes map to original positions <= hitPos
            for (const auto& [padStart, padEnd] : paddingRanges)
            {
                const Sci_Position padLen = padEnd - padStart;
                // The original position of this padding (before it was inserted) is padStart - delta
                const Sci_Position origPadPos = padStart - delta;
                if (origPadPos <= hitPos)
                    delta += padLen;
                else
                    break; // Sorted: remaining pads are after hitPos in original coords
            }
        }
        else {
            // Pre-removal ranges: same coordinate system as hitPos
            for (const auto& [padStart, padEnd] : paddingRanges)
            {
                if (padStart < hitPos)
                    delta += (padEnd - padStart);
                else
                    break; // Sorted: remaining pads are at or after hitPos
            }
        }
        return delta;
        };

    const Sci_Position sign = added ? 1 : -1;

    for (auto& hit : _hits)
    {
        if (!pathsEqualUtf8(hit.fullPathUtf8, filePathUtf8))
            continue;

        // Adjust primary position
        Sci_Position d = computeDelta(hit.pos);
        hit.pos += sign * d;
        if (hit.pos < 0) hit.pos = 0;

        // Adjust merged positions (multiple hits on same display row)
        for (auto& aPos : hit.allPositions)
        {
            d = computeDelta(aPos);
            aPos += sign * d;
            if (aPos < 0) aPos = 0;
        }
    }
}

void ResultDock::SwitchAndJump(const std::wstring& fullPath, Sci_Position pos, Sci_Position len)
{
    // Check for pseudo-paths (like <untitled>)
    const bool isPseudo = IsPseudoPath(fullPath);

    // Check if already in the correct document
    wchar_t currentPath[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(currentPath));

    if (!isPseudo && _wcsicmp(currentPath, fullPath.c_str()) == 0) {
        // Already in correct document - jump directly
        JumpSelectCenterActiveEditor(pos, len);
        return;
    }

    if (isPseudo && IsCurrentDocByTitle(fullPath)) {
        // Already in correct pseudo document - jump directly
        JumpSelectCenterActiveEditor(pos, len);
        return;
    }

    // Need to switch/open document
    std::wstring pathToOpen = fullPath;
    if (isPseudo) {
        pathToOpen = BuildDefaultPathForPseudo(fullPath);
        if (pathToOpen.empty()) return;
    }

    SetPendingJump(pathToOpen, pos, len);

    // Use NPPM_DOOPEN instead to open files that aren't already open
    ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(pathToOpen.c_str()));
}

void ResultDock::scrollToHitAndHighlight(int displayLineStart)
{
    if (!_hSci || displayLineStart < 0)
        return;

    // Get the document line from the byte position
    int docLine = static_cast<int>(S(SCI_LINEFROMPOSITION, displayLineStart, 0));

    // Ensure the line is visible (expand any collapsed folds containing this line)
    S(SCI_ENSUREVISIBLE, docLine, 0);

    // Also expand parent folds if necessary
    int parentLine = static_cast<int>(S(SCI_GETFOLDPARENT, docLine, 0));
    while (parentLine >= 0) {
        if (!S(SCI_GETFOLDEXPANDED, parentLine, 0)) {
            S(SCI_SETFOLDEXPANDED, parentLine, TRUE);
            S(SCI_FOLDCHILDREN, parentLine, SC_FOLDACTION_EXPAND);
        }
        parentLine = static_cast<int>(S(SCI_GETFOLDPARENT, parentLine, 0));
    }

    // Convert document line to visible (display) line for proper centering
    // This accounts for collapsed folds above the target line
    int visibleLine = static_cast<int>(S(SCI_VISIBLEFROMDOCLINE, docLine, 0));

    // Calculate target scroll position BEFORE setting selection
    // (SCI_SETSEL can trigger auto-scroll which would corrupt our calculation)
    int linesOnScreen = static_cast<int>(S(SCI_LINESONSCREEN, 0, 0));
    int targetFirstLine = visibleLine - (linesOnScreen / 2);
    if (targetFirstLine < 0) targetFirstLine = 0;

    // Suppress redraw during scroll operations to prevent flicker
    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);

    // First scroll to the target position
    S(SCI_SETFIRSTVISIBLELINE, targetFirstLine, 0);

    // Now set the selection (this won't auto-scroll since we're already at the right position)
    Sci_Position lineStartPos = S(SCI_POSITIONFROMLINE, docLine, 0);
    Sci_Position lineEndPos = S(SCI_GETLINEENDPOSITION, docLine, 0);
    S(SCI_SETSEL, lineStartPos, lineEndPos);

    // Reset horizontal scroll to show line beginning
    S(SCI_SETXOFFSET, 0, 0);

    // Re-enable redraw and force repaint
    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);
    ::RedrawWindow(_hSci, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
}

// ------------------- Global Shortcut Actions -----------------

void ResultDock::focusDock()
{
    extern NppData nppData;
    if (!_hSci) return;

    // Show dock if hidden, then set keyboard focus to the Scintilla control
    ensureCreatedAndVisible(nppData);
    ::SetFocus(_hSci);
}

void ResultDock::gotoNextHit()
{
    if (!_hSci || _hits.empty()) return;

    // Determine current position: if dock has focus, use dock cursor; otherwise start from 0
    Sci_Position curPos = S(SCI_GETCURRENTPOS, 0, 0);
    int curLine = static_cast<int>(S(SCI_LINEFROMPOSITION, curPos, 0));
    Sci_Position curLineStart = S(SCI_POSITIONFROMLINE, curLine, 0);

    // Find the current hit index (if cursor is on a hit line)
    auto it = _lineStartToHitIndex.find(static_cast<int>(curLineStart));
    size_t nextIdx;

    if (it != _lineStartToHitIndex.end()) {
        // On a hit line: advance to the next hit (wrap around)
        nextIdx = (it->second + 1) % _hits.size();
    }
    else {
        // Not on a hit line: find the first hit after the current line
        nextIdx = 0;
        for (size_t i = 0; i < _hits.size(); ++i) {
            if (_hits[i].displayLineStart > static_cast<int>(curLineStart)) {
                nextIdx = i;
                break;
            }
        }
    }

    // Navigate to the hit in the editor
    navigateFromDockLine(_hSci,
        static_cast<int>(S(SCI_LINEFROMPOSITION, _hits[nextIdx].displayLineStart, 0)));

    // Scroll dock to show the navigated hit
    scrollToHitAndHighlight(_hits[nextIdx].displayLineStart);
}

void ResultDock::gotoPrevHit()
{
    if (!_hSci || _hits.empty()) return;

    Sci_Position curPos = S(SCI_GETCURRENTPOS, 0, 0);
    int curLine = static_cast<int>(S(SCI_LINEFROMPOSITION, curPos, 0));
    Sci_Position curLineStart = S(SCI_POSITIONFROMLINE, curLine, 0);

    auto it = _lineStartToHitIndex.find(static_cast<int>(curLineStart));
    size_t prevIdx;

    if (it != _lineStartToHitIndex.end()) {
        // On a hit line: go to previous (wrap around)
        prevIdx = (it->second == 0) ? _hits.size() - 1 : it->second - 1;
    }
    else {
        // Not on a hit line: find the last hit before the current line
        prevIdx = _hits.size() - 1;
        for (size_t i = _hits.size(); i-- > 0;) {
            if (_hits[i].displayLineStart < static_cast<int>(curLineStart)) {
                prevIdx = i;
                break;
            }
        }
    }

    navigateFromDockLine(_hSci,
        static_cast<int>(S(SCI_LINEFROMPOSITION, _hits[prevIdx].displayLineStart, 0)));

    scrollToHitAndHighlight(_hits[prevIdx].displayLineStart);
}

ResultDock::CursorHitInfo ResultDock::getCurrentCursorHitInfo() const {
    CursorHitInfo info;
    if (!_hSci) return info;

    // Get current cursor position and line
    Sci_Position curPos = S(SCI_GETCURRENTPOS, 0, 0);
    int line = static_cast<int>(S(SCI_LINEFROMPOSITION, curPos, 0));
    Sci_Position lineStart = S(SCI_POSITIONFROMLINE, line, 0);

    // Lookup hit index from line start position
    auto it = _lineStartToHitIndex.find(static_cast<int>(lineStart));
    if (it == _lineStartToHitIndex.end()) return info;

    size_t hitIdx = it->second;
    if (hitIdx >= _hits.size()) return info;

    info.valid = true;
    info.hitIndex = hitIdx;

    return info;
}

size_t ResultDock::getHitIndexAtLineStart(int lineStartPos) const
{
    auto it = _lineStartToHitIndex.find(lineStartPos);
    if (it == _lineStartToHitIndex.end()) return SIZE_MAX;

    size_t hitIdx = it->second;
    if (hitIdx >= _hits.size()) return SIZE_MAX;

    return hitIdx;
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

// --------------- Context Menu Command Handlers ------------

void ResultDock::copySelectedLines(HWND hSci) {

    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };

    auto indentOf = [](ResultDock::LineLevel lvl) -> int {
        return ResultDock::INDENT_SPACES[static_cast<int>(lvl)];
        };

    // determine selection range
    Sci_Position a = (Sci_Position)Sx(SCI_GETSELECTIONNANCHOR);
    Sci_Position c = (Sci_Position)Sx(SCI_GETCURRENTPOS);
    if (a > c) std::swap(a, c);
    int lineStart = static_cast<int>(Sx(SCI_LINEFROMPOSITION, a));
    int lineEnd = static_cast<int>(Sx(SCI_LINEFROMPOSITION, c));
    bool hasSel = (a != c);

    std::vector<std::wstring> out;

    if (hasSel) {                               // ----- selection mode
        for (int l = lineStart; l <= lineEnd; ++l) {
            int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
            std::string raw(len, '\0');
            Sx(SCI_GETLINE, l, (sptr_t)raw.data());
            raw.resize(strnlen(raw.c_str(), len));
            if (classify(raw) == LineKind::HitLine)
                out.emplace_back(stripHitPrefix(Encoding::utf8ToWString(raw)));
        }
    }
    else {                                    // ----- caret hierarchy walk
        int caretLine = lineStart;
        int len = static_cast<int>(Sx(SCI_LINELENGTH, caretLine));
        std::string rawCaret(len, '\0');
        Sx(SCI_GETLINE, caretLine, (sptr_t)rawCaret.data());
        rawCaret.resize(strnlen(rawCaret.c_str(), len));
        LineKind kind = classify(rawCaret);

        auto pushHitsBelow = [&](int fromLine, int minIndent) {
            int total = static_cast<int>(Sx(SCI_GETLINECOUNT));
            for (int l = fromLine; l < total; ++l) {
                int lLen = static_cast<int>(Sx(SCI_LINELENGTH, l));
                std::string raw(lLen, '\0');
                Sx(SCI_GETLINE, l, (sptr_t)raw.data());
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
    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };

    Sci_Position a = (Sci_Position)Sx(SCI_GETSELECTIONNANCHOR);
    Sci_Position c = (Sci_Position)Sx(SCI_GETCURRENTPOS);
    if (a > c) std::swap(a, c);
    int lineStart = static_cast<int>(Sx(SCI_LINEFROMPOSITION, a));
    int lineEnd = static_cast<int>(Sx(SCI_LINEFROMPOSITION, c));
    bool hasSel = (a != c);

    std::vector<std::wstring> paths;
    std::unordered_set<std::wstring> seen;
    auto addUnique = [&](const std::wstring& p) { if (seen.insert(p).second) paths.push_back(p); };

    auto pushFileHdrsBelow = [&](int fromLine, int minIndent) {
        int total = static_cast<int>(Sx(SCI_GETLINECOUNT));
        for (int l = fromLine; l < total; ++l) {
            int lLen = static_cast<int>(Sx(SCI_LINELENGTH, l));
            std::string raw(lLen, '\0');
            Sx(SCI_GETLINE, l, (sptr_t)raw.data());
            raw.resize(strnlen(raw.c_str(), lLen));

            int indent = ResultDock::leadingSpaces(raw.c_str(), (int)raw.size());
            if (indent <= minIndent && classify(raw) != LineKind::HitLine) break;

            if (classify(raw) == LineKind::FileHdr)
                addUnique(pathFromFileHdr(Encoding::utf8ToWString(raw)));
        }
        };

    if (hasSel) {
        for (int l = lineStart; l <= lineEnd; ++l) {
            int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
            std::string raw(len, '\0');
            Sx(SCI_GETLINE, l, (sptr_t)raw.data());
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
        int len = static_cast<int>(Sx(SCI_LINELENGTH, caretLine));
        std::string raw(len, '\0');
        Sx(SCI_GETLINE, caretLine, (sptr_t)raw.data());
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

    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };

    Sci_Position a = (Sci_Position)Sx(SCI_GETSELECTIONNANCHOR);
    Sci_Position c = (Sci_Position)Sx(SCI_GETCURRENTPOS);
    if (a > c) std::swap(a, c);
    int lineStart = static_cast<int>(Sx(SCI_LINEFROMPOSITION, a));
    int lineEnd = static_cast<int>(Sx(SCI_LINEFROMPOSITION, c));

    std::unordered_set<std::wstring> opened;

    for (int l = lineStart; l <= lineEnd; ++l) {
        int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
        std::string raw(len, '\0');
        Sx(SCI_GETLINE, l, (sptr_t)raw.data());
        raw.resize(strnlen(raw.c_str(), len));
        LineKind k = classify(raw);

        if (k == LineKind::FileHdr) {
            std::wstring p = pathFromFileHdr(Encoding::utf8ToWString(raw));
            if (opened.insert(p).second)
                ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(p.c_str()));
        }
        else if (k == LineKind::HitLine || k == LineKind::CritHdr) {
            int fLine = ancestorFileLine(hSci, l);
            if (fLine != -1) {
                std::wstring p = pathFromFileHdr(getLineW(hSci, fLine));
                if (opened.insert(p).second)
                    ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(p.c_str()));
            }
        }
    }
}

void ResultDock::deleteSelectedItems(HWND hSci)
{
    auto& dock = ResultDock::instance();

    auto fn = (SciFnDirect_t)::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0);
    auto ptr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
    auto Sx = [&](UINT m, uptr_t w = 0, sptr_t l = 0)->sptr_t {
        return fn ? fn(ptr, m, w, l) : ::SendMessage(hSci, m, w, l);
        };

    // Determine anchor / caret lines
    Sci_Position a = (Sci_Position)Sx(SCI_GETSELECTIONNANCHOR);
    Sci_Position c = (Sci_Position)Sx(SCI_GETCURRENTPOS);
    if (a > c) std::swap(a, c);

    int firstLine = static_cast<int>(Sx(SCI_LINEFROMPOSITION, a));
    int lastLine = static_cast<int>(Sx(SCI_LINEFROMPOSITION, c));
    bool hasSel = (a != c);

    // Build list of display-line ranges to delete (unchanged logic)
    struct DelRange { int first; int last; };
    std::vector<DelRange> ranges;

    auto subtreeEnd = [&](int fromLine, int minIndent) -> int {
        int total = static_cast<int>(Sx(SCI_GETLINECOUNT));
        for (int l = fromLine; l < total; ++l) {
            int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
            std::string raw(len, '\0');
            Sx(SCI_GETLINE, l, (sptr_t)raw.data());
            raw.resize(strnlen(raw.c_str(), len));

            int indent = ResultDock::leadingSpaces(raw.c_str(), (int)raw.size());
            // stop when same-or-less indent and not a HitLine
            if (indent <= minIndent && classify(raw) != LineKind::HitLine)
                return l - 1;
        }
        return static_cast<int>(Sx(SCI_GETLINECOUNT) - 1);
        };

    auto pushRange = [&](int f, int l) {
        if (ranges.empty() || f > ranges.back().last + 1)
            ranges.push_back({ f, l });
        else
            ranges.back().last = (std::max)(ranges.back().last, l);
        };

    if (hasSel) {
        for (int l = firstLine; l <= lastLine; ++l) {
            if (!ranges.empty() && l <= ranges.back().last) continue;

            int len = static_cast<int>(Sx(SCI_LINELENGTH, l));
            std::string raw(len, '\0');
            Sx(SCI_GETLINE, l, (sptr_t)raw.data());
            raw.resize(strnlen(raw.c_str(), len));

            LineKind kind = classify(raw);
            int endLine = l;
            switch (kind) {
            case LineKind::HitLine:   endLine = l; break;
            case LineKind::CritHdr:   endLine = subtreeEnd(l + 1, INDENT_SPACES[(int)LineLevel::CritHdr]);  break;
            case LineKind::FileHdr:   endLine = subtreeEnd(l + 1, INDENT_SPACES[(int)LineLevel::FileHdr]);  break;
            case LineKind::SearchHdr: endLine = subtreeEnd(l + 1, INDENT_SPACES[(int)LineLevel::SearchHdr]); break;
            default:                  continue;
            }
            pushRange(l, endLine);
        }
    }
    else {
        int len = static_cast<int>(Sx(SCI_LINELENGTH, firstLine));
        std::string raw(len, '\0');
        Sx(SCI_GETLINE, firstLine, (sptr_t)raw.data());
        raw.resize(strnlen(raw.c_str(), len));

        LineKind kind = classify(raw);
        int endLine = firstLine;
        switch (kind) {
        case LineKind::HitLine:   endLine = firstLine; break;
        case LineKind::CritHdr:   endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[(int)LineLevel::CritHdr]);  break;
        case LineKind::FileHdr:   endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[(int)LineLevel::FileHdr]);  break;
        case LineKind::SearchHdr: endLine = subtreeEnd(firstLine + 1, INDENT_SPACES[(int)LineLevel::SearchHdr]); break;
        default: break;
        }
        pushRange(firstLine, endLine);
    }

    if (ranges.empty())
        return;

    // Delete ranges bottom-up (keep your original structure)
    std::sort(ranges.begin(), ranges.end(), [](auto& A, auto& B) {
        return A.first != B.first ? A.first < B.first : A.last < B.last;
        });

    for (auto it = ranges.rbegin(); it != ranges.rend(); ++it) {
        const int l0 = it->first;
        const int l1 = it->last;

        Sci_Position p0 = (Sci_Position)Sx(SCI_POSITIONFROMLINE, l0);
        Sci_Position p1 = (Sci_Position)Sx(SCI_GETLINEENDPOSITION, l1);

        // include trailing CRLF if not EOF
        const int totalLines = static_cast<int>(Sx(SCI_GETLINECOUNT));
        if (l1 < totalLines - 1)
            p1 += 2;

        const int delta = (int)(p1 - p0);

        // remove hits inside [p0, p1)
        dock._hits.erase(
            std::remove_if(dock._hits.begin(), dock._hits.end(),
                [&](const Hit& h) { return h.displayLineStart >= (int)p0 && h.displayLineStart < (int)p1; }),
            dock._hits.end());
        // shift hits at/after p1 back by delta
        for (auto& h : dock._hits)
            if (h.displayLineStart >= (int)p1)
                h.displayLineStart -= delta;

        // per-range redraw/read-only toggling
        ::SendMessage(hSci, WM_SETREDRAW, FALSE, 0);
        Sx(SCI_SETREADONLY, FALSE);
        Sx(SCI_DELETERANGE, (uptr_t)p0, (sptr_t)delta);
        Sx(SCI_SETREADONLY, TRUE);
        ::SendMessage(hSci, WM_SETREDRAW, TRUE, 0);
    }

    // Rebuild dock caches as before
    dock.rebuildFolding();
    dock.applyStyling();
    dock.rebuildHitLineIndex();

    // force a synchronous repaint (matches prependBlock's explicit refresh idea) ---
    ::RedrawWindow(hSci, nullptr, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    ::UpdateWindow(hSci);
}

// -------------------- Callbacks/Subclassing ---------------

// Toggle fold on a header line: move cursor to start, toggle, reset scroll.
void ResultDock::toggleFoldAtLine(HWND hwnd, int line)
{
    const Sci_Position lineStart = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, line, 0);
    ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, lineStart, 0);
    ::SendMessage(hwnd, SCI_SETXOFFSET, 0, 0);
    ::SendMessage(hwnd, SCI_TOGGLEFOLD, line, 0);
    ::SendMessage(hwnd, SCI_SETXOFFSET, 0, 0);
}

// Navigate from a result dock line to the corresponding hit in the editor.
// Returns true if the line was a hit and navigation was initiated.
bool ResultDock::navigateFromDockLine(HWND hwnd, int dispLine)
{
    extern NppData nppData;

    const int lineLen = static_cast<int>(::SendMessage(hwnd, SCI_LINELENGTH, dispLine, 0));
    if (lineLen <= 0) return false;

    std::string raw(lineLen, '\0');
    ::SendMessage(hwnd, SCI_GETLINE, dispLine, reinterpret_cast<LPARAM>(raw.data()));
    raw.resize(strnlen(raw.c_str(), lineLen));
    if (classify(raw) != LineKind::HitLine) return false;

    const Sci_Position lineStartPos = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
    ResultDock& dock = instance();
    const size_t hitIndex = dock.getHitIndexAtLineStart(static_cast<int>(lineStartPos));
    if (hitIndex == SIZE_MAX) return false;

    const auto& allHits = dock.hits();
    if (hitIndex >= allHits.size()) return false;
    const Hit& hit = allHits[hitIndex];

    std::wstring wPath;
    if (!hit.fullPathUtf8.empty()) {
        const int needed = ::MultiByteToWideChar(CP_UTF8, 0,
            hit.fullPathUtf8.c_str(), static_cast<int>(hit.fullPathUtf8.size()),
            nullptr, 0);
        if (needed > 0) {
            wPath.resize(needed);
            ::MultiByteToWideChar(CP_UTF8, 0,
                hit.fullPathUtf8.c_str(), static_cast<int>(hit.fullPathUtf8.size()),
                &wPath[0], needed);
        }
    }
    if (wPath.empty()) return false;

    // Preserve dock scroll and cursor position across navigation
    const int firstVisible = static_cast<int>(::SendMessage(hwnd, SCI_GETFIRSTVISIBLELINE, 0, 0));

    auto restoreDockView = [&]() {
        ::SendMessage(hwnd, SCI_SETFIRSTVISIBLELINE, firstVisible, 0);
        const Sci_Position dockLineStart = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMLINE, dispLine, 0);
        ::SendMessage(hwnd, SCI_SETEMPTYSELECTION, dockLineStart, 0);
        };

    const bool isPseudo = IsPseudoPath(wPath);
    std::wstring pathToOpen = wPath;

    if (!isPseudo && IsCurrentDocByFullPath(pathToOpen)) {
        NavigateToHit(hit);
        restoreDockView();
        return true;
    }
    if (isPseudo && IsCurrentDocByTitle(wPath)) {
        NavigateToHit(hit);
        restoreDockView();
        return true;
    }
    if (isPseudo) {
        pathToOpen = BuildDefaultPathForPseudo(wPath);
        if (pathToOpen.empty()) return false;
    }

    s_pending.setFromHit(pathToOpen, hit);

    const LRESULT ok = ::SendMessage(nppData._nppHandle, NPPM_DOOPEN, 0, reinterpret_cast<LPARAM>(pathToOpen.c_str()));
    if (!ok) {
        s_pending.clear();
        return false;
    }

    restoreDockView();
    return true;
}

LRESULT CALLBACK ResultDock::sciSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    extern NppData nppData;

    switch (msg) {
    case WM_LBUTTONDOWN: {
        // Check if click is in the fold margin (margin 2)
        int x = GET_X_LPARAM(lp);
        int y = GET_Y_LPARAM(lp);

        int margin0Width = static_cast<int>(::SendMessage(hwnd, SCI_GETMARGINWIDTHN, 0, 0));
        int margin1Width = static_cast<int>(::SendMessage(hwnd, SCI_GETMARGINWIDTHN, 1, 0));
        int margin2Width = static_cast<int>(::SendMessage(hwnd, SCI_GETMARGINWIDTHN, 2, 0));

        int foldMarginStart = margin0Width + margin1Width;
        int foldMarginEnd = foldMarginStart + margin2Width;

        if (x >= foldMarginStart && x < foldMarginEnd) {
            Sci_Position pos = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
            int line = static_cast<int>(::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0));
            int level = static_cast<int>(::SendMessage(hwnd, SCI_GETFOLDLEVEL, line, 0));

            if (level & SC_FOLDLEVELHEADERFLAG) {
                toggleFoldAtLine(hwnd, line);
                return 0;
            }
        }
        break;
    }

    case WM_NOTIFY: {
        NMHDR* hdr = reinterpret_cast<NMHDR*>(lp);
        if (hdr->code == DMN_CLOSE) {
            ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(hwnd));
            return TRUE;
        }
        break;
    }

    case WM_LBUTTONDBLCLK:
    {
        const int x = LOWORD(lp);
        const int y = HIWORD(lp);
        const Sci_Position posInDock = (Sci_Position)::SendMessage(hwnd, SCI_POSITIONFROMPOINT, x, y);
        if (posInDock < 0) return 0;

        const int dispLine = static_cast<int>(::SendMessage(hwnd, SCI_LINEFROMPOSITION, posInDock, 0));

        const int level = static_cast<int>(::SendMessage(hwnd, SCI_GETFOLDLEVEL, dispLine, 0));
        if (level & SC_FOLDLEVELHEADERFLAG) {
            toggleFoldAtLine(hwnd, dispLine);
            return 0;
        }

        navigateFromDockLine(hwnd, dispLine);
        return 0;
    }

    case WM_TIMER:
    {
        if (wp == s_timerId)
        {
            ::KillTimer(hwnd, s_timerId);
            if (s_pending.active && s_pending.targetEditor && s_pending.phase == 2)
            {
                HWND hEd = s_pending.targetEditor;
                bool jumped = false;

                // Line-based re-search: find match closest to stored position
                if (!jumped && s_pending.docLine >= 0 && !s_pending.findTextW.empty())
                {
                    const int lineCount = static_cast<int>(::SendMessage(hEd, SCI_GETLINECOUNT, 0, 0));
                    if (s_pending.docLine < lineCount)
                    {
                        const Sci_Position lineStart = ::SendMessage(hEd, SCI_POSITIONFROMLINE, s_pending.docLine, 0);
                        const Sci_Position lineEnd = ::SendMessage(hEd, SCI_GETLINEENDPOSITION, s_pending.docLine, 0);

                        const int docCp = static_cast<int>(::SendMessage(hEd, SCI_GETCODEPAGE, 0, 0));
                        const std::string findBytes = Encoding::wstringToBytes(s_pending.findTextW, docCp);

                        if (!findBytes.empty())
                        {
                            ::SendMessage(hEd, SCI_SETSEARCHFLAGS, s_pending.searchFlags, 0);

                            Sci_Position bestPos = -1;
                            Sci_Position bestLen = 0;
                            Sci_Position bestDistance = 0x7FFFFFFF;

                            Sci_Position searchFrom = lineStart;
                            while (searchFrom < lineEnd) {
                                ::SendMessage(hEd, SCI_SETTARGETRANGE, searchFrom, lineEnd);
                                const Sci_Position found = ::SendMessage(hEd, SCI_SEARCHINTARGET,
                                    findBytes.size(), reinterpret_cast<LPARAM>(findBytes.c_str()));

                                if (found < 0) break;

                                const Sci_Position foundEnd = ::SendMessage(hEd, SCI_GETTARGETEND, 0, 0);
                                const Sci_Position foundLen = foundEnd - found;

                                Sci_Position distance = (found > s_pending.fallbackPos)
                                    ? (found - s_pending.fallbackPos)
                                    : (s_pending.fallbackPos - found);
                                if (distance < bestDistance) {
                                    bestDistance = distance;
                                    bestPos = found;
                                    bestLen = foundLen;
                                }

                                searchFrom = foundEnd;
                                if (searchFrom <= found) break;
                            }

                            if (bestPos >= 0)
                            {
                                JumpSelectCenterActiveEditor(bestPos, bestLen);
                                jumped = true;
                            }
                        }
                    }
                }

                // Direct position fallback
                if (!jumped)
                {
                    JumpSelectCenterActiveEditor(s_pending.fallbackPos, s_pending.fallbackLen);
                }

                s_pending.clear();
            }
            return 0;
        }
        break;
    }



    case DMN_CLOSE:
        ::SendMessage(nppData._nppHandle, NPPM_DMMHIDE, 0, reinterpret_cast<LPARAM>(ResultDock::instance()._hDock));
        return TRUE;

    case WM_KEYDOWN: {
        if (wp == VK_ESCAPE) {
            // Match N++ behavior: ESC in search results panel hides it
            ResultDock::instance().hide(nppData);
            return 0;
        }
        if (wp == VK_DELETE) {
            deleteSelectedItems(hwnd);
            return 0;
        }

        if (wp == VK_RETURN || wp == VK_SPACE) {
            Sci_Position pos = ::SendMessage(hwnd, SCI_GETCURRENTPOS, 0, 0);
            int line = static_cast<int>(::SendMessage(hwnd, SCI_LINEFROMPOSITION, pos, 0));
            int level = static_cast<int>(::SendMessage(hwnd, SCI_GETFOLDLEVEL, line, 0));

            if (level & SC_FOLDLEVELHEADERFLAG) {
                toggleFoldAtLine(hwnd, line);
                return 0;
            }

            // Enter on a hit line navigates to the match in the editor
            if (wp == VK_RETURN) {
                navigateFromDockLine(hwnd, line);
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
        if (wp != reinterpret_cast<WPARAM>(hwnd)) return 0;

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
        {
            const int lineCount = static_cast<int>(::SendMessage(hwnd, SCI_GETLINECOUNT, 0, 0));

            for (int line = lineCount - 1; line >= 0; --line)
            {
                const int level = static_cast<int>(::SendMessage(hwnd, SCI_GETFOLDLEVEL, line, 0));
                if (level & SC_FOLDLEVELHEADERFLAG)
                {
                    if (::SendMessage(hwnd, SCI_GETFOLDEXPANDED, line, 0))
                    {
                        ::SendMessage(hwnd, SCI_SETFOLDEXPANDED, line, FALSE);
                        ::SendMessage(hwnd, SCI_FOLDLINE, line, SC_FOLDACTION_CONTRACT);
                    }
                }
            }
            return 0;
        }

        case IDM_RD_UNFOLD_ALL:
            ::SendMessage(hwnd, SCI_FOLDALL, SC_FOLDACTION_EXPAND, 0);
            return 0;

            // ── copy ──────────────────────────────────────
        case IDM_RD_COPY_STD:
            ::SendMessage(hwnd, SCI_COPY, 0, 0);
            return 0;

        case IDM_RD_COPY_LINES:
            copySelectedLines(hwnd);
            return 0;

        case IDM_RD_COPY_PATHS:
            copySelectedPaths(hwnd);
            return 0;

            // ── select / clear ────────────────────────────
        case IDM_RD_SELECT_ALL:
            ::SendMessage(hwnd, SCI_SELECTALL, 0, 0);
            return 0;

        case IDM_RD_CLEAR_ALL:
            ResultDock::instance().clear();
            return 0;

            // ── open paths in N++ ─────────────────────────
        case IDM_RD_OPEN_PATHS:
            openSelectedPaths(hwnd);
            return 0;

            // ── toggle word-wrap ──────────────────────────
        case IDM_RD_TOGGLE_WRAP:
            ResultDock::_wrapEnabled = !ResultDock::_wrapEnabled;
            ::SendMessage(hwnd, SCI_SETWRAPMODE,
                ResultDock::_wrapEnabled ? SC_WRAP_WORD : SC_WRAP_NONE, 0);
            return 0;

            // ── toggle purge-flag ─────────────────────────
        case IDM_RD_TOGGLE_PURGE:
            ResultDock::_purgeOnNextSearch = !ResultDock::_purgeOnNextSearch;
            return 0;
        }

        // Unknown command: not ours → let Scintilla / default proc handle it.
        break;
    }

    }

    return s_prevSciProc
        ? ::CallWindowProc(s_prevSciProc, hwnd, msg, wp, lp)
        : ::DefWindowProc(hwnd, msg, wp, lp);
}

void ResultDock::onNppNotification(const SCNotification* notify)
{
    if (!notify)
        return;

    if (!s_pending.active || s_pending.path.empty())
        return;

    if (notify->nmhdr.code == NPPN_BUFFERACTIVATED)
    {
        wchar_t cur[MAX_PATH] = {};
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(cur));
        if (!cur[0] || _wcsicmp(cur, s_pending.path.c_str()) != 0)
            return;

        int whichView = 0;
        ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, reinterpret_cast<LPARAM>(&whichView));
        s_pending.targetEditor = (whichView == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
        s_pending.phase = 1;
        return;
    }

    if (notify->nmhdr.code == SCN_UPDATEUI)
    {
        if (!s_pending.targetEditor || notify->nmhdr.hwndFrom != s_pending.targetEditor || s_pending.phase != 1)
            return;

        s_pending.phase = 2;
        if (_hSci)
            ::SetTimer(_hSci, s_timerId, 1, nullptr);
        return;
    }
}

// ------------------- Color Utilities ----------------------

COLORREF ResultDock::hslToRgb(double hue01, double s, double l)
{
    // Clamp saturation and lightness
    s = (std::max)(TUNE_SAT_MIN, (std::min)(TUNE_SAT_MAX, s));
    l = (std::max)(TUNE_LIT_MIN, (std::min)(TUNE_LIT_MAX, l));

    auto hueToRgb = [](double p, double q, double t) -> double {
        if (t < 0.0) t += 1.0;
        if (t > 1.0) t -= 1.0;
        if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
        if (t < 1.0 / 2.0) return q;
        if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        return p;
        };

    double r, g, b;
    if (s < 0.001) {
        r = g = b = l;
    }
    else {
        double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
        double p = 2.0 * l - q;
        r = hueToRgb(p, q, hue01 + 1.0 / 3.0);
        g = hueToRgb(p, q, hue01);
        b = hueToRgb(p, q, hue01 - 1.0 / 3.0);
    }

    return RGB(
        static_cast<BYTE>(r * 255.0 + 0.5),
        static_cast<BYTE>(g * 255.0 + 0.5),
        static_cast<BYTE>(b * 255.0 + 0.5)
    );
}

COLORREF ResultDock::generateColorFromText(const std::wstring& text, bool darkMode)
{
    // MurmurHash3-inspired mixing for excellent bit distribution
    unsigned long h = 0;
    for (wchar_t c : text) {
        h ^= static_cast<unsigned long>(c);
        h *= 0x5bd1e995UL;
        h ^= h >> 15;
    }
    // Final mixing
    h ^= h >> 16;
    h *= 0x85ebca6bUL;
    h ^= h >> 13;
    h *= 0xc2b2ae35UL;
    h ^= h >> 16;

    // Golden Ratio for optimal hue distribution
    // This guarantees that consecutive hashes produce well-spaced hues
    constexpr double GOLDEN_RATIO = 0.618033988749895;
    double hue01 = fmod(h * GOLDEN_RATIO, 1.0);

    // Base saturation and lightness
    double saturation = darkMode ? TUNE_SAT_DARK : TUNE_SAT_LIGHT;
    double lightness = darkMode ? TUNE_LIT_DARK : TUNE_LIT_LIGHT;

    // Add slight variation based on different hash bits (makes colors more unique)
    double satVar = (((h >> 8) & 0xFF) / 255.0 - 0.5) * 2.0 * TUNE_SAT_VAR;
    double litVar = (((h >> 16) & 0xFF) / 255.0 - 0.5) * 2.0 * TUNE_LIT_VAR;

    saturation += satVar;
    lightness += litVar;

    return hslToRgb(hue01, saturation, lightness);
}

void ResultDock::defineSlotColor(int slotIndex, COLORREF color)
{
    // 1. Save to internal map
    if (slotIndex < 0 || slotIndex >= MAX_ENTRY_COLORS) return;  // Bounds check
    _slotToColor[slotIndex] = static_cast<int>(color);

    // 2. Apply configuration immediately to Scintilla
    // This is crucial because clear() might have reset the indicators,
    // and we need them ready BEFORE we start adding text.
    if (!_hSci) return;

    const int indicId = INDIC_ENTRY_BG_BASE + slotIndex;

    // Check dark mode for alpha values (reusing logic from applyTheme)
    const bool dark = (::SendMessage(nppData._nppHandle, NPPM_ISDARKMODEENABLED, 0, 0) != 0);
    const int bgAlpha = dark ? ENTRY_BG_ALPHA_DARK : ENTRY_BG_ALPHA_LIGHT;
    const int outlineAlpha = dark ? ENTRY_OUTLINE_ALPHA_DARK : ENTRY_OUTLINE_ALPHA_LIGHT;

    S(SCI_INDICSETSTYLE, indicId, INDIC_ROUNDBOX);
    S(SCI_INDICSETFORE, indicId, color);
    S(SCI_INDICSETALPHA, indicId, bgAlpha);
    S(SCI_INDICSETOUTLINEALPHA, indicId, outlineAlpha);
    S(SCI_INDICSETUNDER, indicId, TRUE);
}