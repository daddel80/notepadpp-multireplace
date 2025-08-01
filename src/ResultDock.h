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

#pragma once

// ResultDock: dockable Scintilla view that shows search results

#include <windows.h>
#include <string>
#include <vector>
#include <functional>
#include "Encoding.h"
#include "Sci_Position.h"
#include "PluginDefinition.h"
#include "StaticDialog/DockingDlgInterface.h"

// Helper for Scintilla message calls
using SciSendFn = std::function<LRESULT(UINT, WPARAM, LPARAM)>;

// -------------------------------------------------------------
// ResultDock (singleton)
// -------------------------------------------------------------
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

    bool isPurgeEnabled() const { return _purgeOnNextSearch; }

    // --------------- Singleton & API methods ------------------
    static ResultDock& instance();

    void ensureCreatedAndVisible(const NppData& npp);

    void setText(const std::wstring& text);
    void prependText(const std::wstring& text);
    void appendText(const std::wstring& text);

    void prependHits(const std::vector<Hit>& newHits, const std::wstring& dockText);
    void recordHit(const std::string& fullPathUtf8, Sci_Position pos, Sci_Position length);

    void clear();              // Remove all text and hits
    void rebuildFolding();     // Re-calculate folding markers
    void applyStyling() const; // Update indicators/colors
    void onThemeChanged();     // Called on N++ dark mode toggle

    const std::vector<Hit>& hits() const { return _hits; }

    void formatHitsForFile(const std::wstring& wFilePath,
        const SciSendFn& sciSend,
        std::vector<Hit>& hitsInOut,
        std::wstring& outBlock,
        size_t& ioUtf8Len) const;

    void formatHitsLines(const SciSendFn& sciSend,
        std::vector<Hit>& hitsInOut,
        std::wstring& outBlock,
        size_t& ioUtf8Len) const;

    // Used for buildListText() (multi-file/grouped display)
    struct CritAgg { std::wstring text; std::vector<Hit> hits; };
    struct FileAgg { std::wstring wPath; int hitCount = 0; std::vector<CritAgg> crits; };
    using FileMap = std::unordered_map<std::string, FileAgg>;

    void buildListText(const FileMap& files,
        bool groupView,
        const std::wstring& header,
        const SciSendFn& sciSend,
        std::wstring& outText,
        std::vector<Hit>& outHits) const;

    static bool  wrapEnabled() { return _wrapEnabled; }
    static bool  purgeEnabled() { return _purgeOnNextSearch; }
    static void  setWrapEnabled(bool v) { _wrapEnabled = v; }
    static void  setPurgeEnabled(bool v) { _purgeOnNextSearch = v; }

    /// Semantic levels in the result dock
    enum class LineLevel : int {
        SearchHdr = 0,  // top‑level search header
        FileHdr = 1,  // file header (only in grouped view)
        CritHdr = 2,  // criterion header (only in grouped view)
        HitLine = 3   // actual hit line
    };

    /// Number of spaces to indent for each LineLevel
    static constexpr int INDENT_SPACES[] = {
        /* SearchHdr */ 1,
        /* FileHdr   */ 2,
        /* CritHdr   */ 3,
        /* HitLine   */ 4
    };

private:
    // --------------------- Theme Colors ------------------------
    // Holds all relevant colors for the dock panel (light/dark)
    struct DockThemeColors {
        COLORREF lineBg;        // Background for result line
        COLORREF lineNr;        // Line number color
        COLORREF matchFg;       // Match text color
        COLORREF matchBg;       // Match background color
        COLORREF headerBg;      // Header line background (SearchHdr)
        COLORREF headerFg;      // Header text color (SearchHdr)
        COLORREF critHdrBg;     // Header line background (CritHdr)
        COLORREF critHdrFg;     // Header text color (CritHdr)
        COLORREF filePathFg;    // File path color
        COLORREF foldGlyph;     // Fold glyph color
        COLORREF foldHighlight; // Fold marker highlight
        COLORREF caretLineBg;   // Caret line background
        int caretLineAlpha;     // Caret line transparency
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

    // Returns theme colors for current dark/light mode
    static constexpr const DockThemeColors& currentColors(bool darkMode) {
        return darkMode ? DarkDockTheme : LightDockTheme;
    }

    // Context‑menu command IDs (reserve a private range) ‑‑‑
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

    // ------------------- Construction & State -----------------
    explicit ResultDock(HINSTANCE hInst);
    ResultDock(const ResultDock&) = delete;
    ResultDock& operator=(const ResultDock&) = delete;

    void create(const NppData& npp);
    void initFolding() const;
    void applyTheme();

    static void copySelectedLines(HWND hSci);
    static void copySelectedPaths(HWND hSci);
    static void openSelectedPaths(HWND hSci);
    static void copyTextToClipboard(HWND owner, const std::wstring& w);
    static void deleteSelectedItems(HWND hSci);
    void collapseOldSearches();

    // ‑‑‑ persistent UI flags ‑‑‑

    inline static bool _wrapEnabled = false;
    inline static bool _purgeOnNextSearch = false;

    HWND                 _hScintilla = nullptr; // handle for Scintilla control
    std::wstring         _content;              // dock display text
    std::vector<int>     _foldHeaders;          // lines with folding headers

    static int  leadingSpaces(const char* line, int len);
    static LRESULT CALLBACK sciSubclassProc(HWND, UINT, WPARAM, LPARAM);
    static inline WNDPROC   s_prevSciProc = nullptr;

    tTbData _dockData{}; // Persistent docking info

    HINSTANCE   _hInst{ nullptr };     // DLL instance
    HWND        _hSci{ nullptr };      // Scintilla handle (client)
    HWND        _hDock{ nullptr };     // N++ dock container
    std::vector<Hit> _hits;            // all visible hits

    // ----------------- Scintilla Style/Indicator IDs ----------
    static constexpr int INDIC_LINE_BACKGROUND   = 8;   // grey hit background
    static constexpr int INDIC_LINENUMBER_FORE   = 9;   // colored line number
    static constexpr int INDIC_MATCH_FORE        = 10;  // match text color
    static constexpr int INDIC_MATCH_BG          = 14;  // (not used in dark)
    static constexpr int INDIC_HEADER_BACKGROUND = 11;  // header bg
    static constexpr int INDIC_FILEPATH_FORE     = 12;  // file path color
    static constexpr int INDIC_HEADER_FORE       = 13;  // header fg
    static constexpr int MARKER_HEADER_BACKGROUND = 24;
    static constexpr int STYLE_HEADER            = 33;
    static constexpr int STYLE_CRITHDR = 34;
    static constexpr int STYLE_FILEPATH          = 35;
};
