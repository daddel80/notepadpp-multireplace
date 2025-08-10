#pragma once

// ResultDock: dockable Scintilla view that shows search results

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


// ResultDock: dockable Scintilla view that shows search results

#pragma once

#include <windows.h>
#include <string>
#include <vector>
#include <functional>
#include <unordered_map>

#include "Encoding.h"
#include "Sci_Position.h"
#include "PluginDefinition.h"
#include "StaticDialog/DockingDlgInterface.h"

// Helper for Scintilla message calls
using SciSendFn = std::function<LRESULT(UINT, WPARAM, LPARAM)>;

// ============================================================================
// ResultDock
// ============================================================================
class ResultDock final
{
public:
    // --------------------- Hit definition ---------------------
    // Represents one visible "Line N:" entry in the dock
    struct Hit
    {
        std::string  fullPathUtf8;  // File path (UTF-8)
        Sci_Position pos{};         // Match start in original file
        Sci_Position length{};      // Match length

        // Styling offsets in dock buffer
        int displayLineStart{ -1 }; // Absolute char pos of "Line N:"
        int numberStart{ 0 };       // Offset of digits
        int numberLen{ 0 };         // Length of digits
        std::vector<int> matchStarts; // Offsets of match substrings
        std::vector<int> matchLens;   // Lengths of match substrings
    };

    struct CritAgg { std::wstring text; std::vector<Hit> hits; };
    struct FileAgg { std::wstring wPath; int hitCount = 0; std::vector<CritAgg> crits; };
    using FileMap = std::unordered_map<std::string, FileAgg>;

    // --------------- Singleton & API methods ------------------
    static ResultDock& instance();

    void ensureCreatedAndVisible(const NppData& npp);

    void clear();              // Remove all text and hits
    void rebuildFolding();     // Re-calculate folding markers
    void applyStyling() const; // Update indicators/colors
    void onThemeChanged();     // Called on N++ dark mode toggle

    const std::vector<Hit>& hits() const { return _hits; }

    static bool  wrapEnabled() { return _wrapEnabled; }
    static bool  purgeEnabled() { return _purgeOnNextSearch; }
    static void  setWrapEnabled(bool v) { _wrapEnabled = v; }
    static void  setPurgeEnabled(bool v) { _purgeOnNextSearch = v; }

    // ------------------- Search Block API ---------------------
    void startSearchBlock(const std::wstring& header, bool  groupView, bool purge);
    void appendFileBlock(const FileMap& fm, const SciSendFn& sciSend);
    void closeSearchBlock(int totalHits, int totalFiles);

private:
    // ---------------- Construction & Core State ---------------
    explicit ResultDock(HINSTANCE hInst);
    ResultDock(const ResultDock&) = delete;
    ResultDock& operator=(const ResultDock&) = delete;

    void create(const NppData& npp);
    void initFolding() const;
    void applyTheme();

    // -------- Range styling / folding (partial updates) -------
    void applyStylingRange(Sci_Position pos0, Sci_Position len, const std::vector<Hit>& newHits) const;
    void rebuildFoldingRange(int firstLine, int lastLine) const;

    // ---------------- Block building / insertion --------------
    void prependBlock(const std::wstring& dockText, std::vector<Hit>& newHits);
    static void appendUtf8Chunked(HWND hSci, const std::string& u8);
    void collapseOldSearches();

    // ---------------------- Formatting ------------------------
    void buildListText(const FileMap& files,
        bool groupView,
        const std::wstring& header,
        const SciSendFn& sciSend,
        std::wstring& outText,
        std::vector<Hit>& outHits) const;

    void formatHitsLines(const SciSendFn& sciSend,
        std::vector<Hit>& hitsInOut,
        std::wstring& outBlock,
        size_t& ioUtf8Len) const;

    // --------------------- Line helpers -----------------------
    enum class LineLevel : int { SearchHdr = 0, FileHdr = 1, CritHdr = 2, HitLine = 3 };
    static constexpr int INDENT_SPACES[] = { 1, 2, 3, 4 };

    enum class LineKind { Blank, SearchHdr, FileHdr, CritHdr, HitLine };
    static LineKind classify(const std::string& raw);

    static std::wstring getIndentString(LineLevel lvl);
    static size_t getIndentUtf8Length(LineLevel lvl);
    static void shiftHits(std::vector<Hit>& v, size_t delta);
    static std::wstring stripHitPrefix(const std::wstring& w);
    static std::wstring pathFromFileHdr(const std::wstring& w);
    static int ancestorFileLine(HWND hSci, int startLine);
    static std::wstring getLineW(HWND hSci, int line);
    static int leadingSpaces(const char* line, int len);

    // -------------------- Path/File helpers -------------------
    static bool IsPseudoPath(const std::wstring& pathOrLabel);
    static bool FileExistsW(const std::wstring& fullPath);
    static std::wstring GetNppProgramDir();
    static bool IsCurrentDocByFullPath(const std::wstring& fullPath);
    static bool IsCurrentDocByTitle(const std::wstring& titleOnly);
    static void SwitchToFileIfOpenByFullPath(const std::wstring& fullPath);
    static std::wstring BuildDefaultPathForPseudo(const std::wstring& label);
    static bool EnsureFileOpenOrOfferCreate(const std::wstring& desiredPath, std::wstring& outOpenedPath);
    static void JumpSelectCenterActiveEditor(Sci_Position pos, Sci_Position len);

    // --------------- Context Menu Command Handlers ------------
    static void copySelectedLines(HWND hSci);
    static void copySelectedPaths(HWND hSci);
    static void openSelectedPaths(HWND hSci);
    static void copyTextToClipboard(HWND owner, const std::wstring& w);
    static void deleteSelectedItems(HWND hSci);

    // -------------------- Callbacks/Subclassing ---------------
    static LRESULT CALLBACK sciSubclassProc(HWND, UINT, WPARAM, LPARAM);
    static inline WNDPROC s_prevSciProc = nullptr;

    // ----------------------- Theme Colors ---------------------
    struct DockThemeColors {
        COLORREF lineBg;
        COLORREF lineNr;
        COLORREF matchFg;
        COLORREF matchBg;
        COLORREF headerBg;
        COLORREF headerFg;
        COLORREF critHdrBg;
        COLORREF critHdrFg;
        COLORREF filePathFg;
        COLORREF foldGlyph;
        COLORREF foldHighlight;
        COLORREF caretLineBg;
        int      caretLineAlpha;
    };

    static constexpr DockThemeColors LightDockTheme = {
        RGB(0xEE, 0xEE, 0xEE), // lineBg
        RGB(0x40, 0x80, 0xBF), // lineNr
        RGB(0xFA, 0x3F, 0x34), // matchFg
        RGB(0xFF, 0xEB, 0x5A), // matchBg
        RGB(0xD5, 0xFF, 0xD5), // headerBg
        RGB(0x00, 0x00, 0x00), // headerFg
        RGB(0xC4, 0xEB, 0xC4), // critHdrBg
        RGB(0x00, 0x00, 0x00), // critHdrFg
        RGB(0xA0, 0x80, 0x50), // filePathFg
        RGB(0x50, 0x50, 0x50), // foldGlyph
        RGB(0xFF, 0x00, 0x00), // foldHighlight
        RGB(0x87, 0x78, 0xCD), // caretLineBg
        45                     // caretLineAlpha
    };

    static constexpr DockThemeColors DarkDockTheme = {
        RGB(0x3A, 0x3D, 0x33), // lineBg
        RGB(0x80, 0xC0, 0xFF), // lineNr
        RGB(0xA6, 0xE2, 0x2E), // matchFg
        RGB(0x3A, 0x3D, 0x33), // matchBg
        RGB(0x8F, 0xAF, 0x9F), // headerBg
        RGB(0x00, 0x00, 0x00), // headerFg
        RGB(0x78, 0x94, 0x84), // critHdrBg
        RGB(0x00, 0x00, 0x00), // critHdrFg
        RGB(0xEB, 0xCB, 0x8B), // filePathFg
        RGB(0x80, 0x80, 0x80), // foldGlyph
        RGB(0x79, 0x94, 0x86), // foldHighlight
        RGB(0xAA, 0xAA, 0xAA), // caretLineBg
        64                     // caretLineAlpha
    };

    static constexpr const DockThemeColors& currentColors(bool darkMode)
    {
        return darkMode ? DarkDockTheme : LightDockTheme;
    }

    // ------------------- Command & style IDs ------------------
    enum : UINT {
        IDM_RD_FOLD_ALL = 60001,
        IDM_RD_UNFOLD_ALL = 60002,
        IDM_RD_COPY_STD = 60003,
        IDM_RD_SELECT_ALL = 60004,
        IDM_RD_CLEAR_ALL = 60005,
        IDM_RD_COPY_LINES = 60006,
        IDM_RD_COPY_PATHS = 60007,
        IDM_RD_OPEN_PATHS = 60008,
        IDM_RD_TOGGLE_WRAP = 60009,
        IDM_RD_TOGGLE_PURGE = 60010
    };

    static constexpr int INDIC_LINE_BACKGROUND = 8;
    static constexpr int INDIC_LINENUMBER_FORE = 9;
    static constexpr int INDIC_MATCH_FORE = 10;
    static constexpr int INDIC_MATCH_BG = 14;
    static constexpr int STYLE_HEADER = 33;
    static constexpr int STYLE_CRITHDR = 34;
    static constexpr int STYLE_FILEPATH = 35;

    // -------------------------- Members -----------------------
    // Handles
    HINSTANCE _hInst{ nullptr };
    HWND      _hSci{ nullptr };
    HWND      _hDock{ nullptr };

    // Core data
    std::vector<Hit> _hits;
    tTbData _dockData{};

    // Pending block build state
    std::wstring     _pendingText;
    std::vector<Hit> _pendingHits;
    size_t           _utf8LenPending = 0;
    bool             _groupViewPending = false;
    bool             _blockOpen = false;

    // Track header line indices for collapse logic
    std::vector<int> _searchHeaderLines;

    // UI Option Flags
    inline static bool _wrapEnabled = false;
    inline static bool _purgeOnNextSearch = false;

    // --- Call acceleration (Scintilla DirectFunction) ---
    using SciFnDirect_t = sptr_t(__cdecl*)(sptr_t, unsigned int, uptr_t, sptr_t);
    SciFnDirect_t _sciFn = nullptr;
    sptr_t        _sciPtr = 0;

    inline LRESULT S(UINT m, WPARAM w = 0, LPARAM l = 0) const {
        return _sciFn ? _sciFn(_sciPtr, m, w, l) : ::SendMessage(_hSci, m, w, l);
    }

    // O(1) mapping from absolute line start to hit index
    void rebuildHitLineIndex();
    std::unordered_map<int, int> _lineStartToHitIndex;
};