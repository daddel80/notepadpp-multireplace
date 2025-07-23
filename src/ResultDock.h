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
        bool flatView,
        const std::wstring& header,
        const SciSendFn& sciSend,
        std::wstring& outText,
        std::vector<Hit>& outHits) const;

private:
    // --------------------- Theme Colors ------------------------
    // Holds all relevant colors for the dock panel (light/dark)
    struct DockThemeColors {
        COLORREF lineBg;        // Background for result line
        COLORREF lineNr;        // Line number color
        COLORREF matchFg;       // Match text color
        COLORREF matchBg;       // Match background color
        COLORREF headerBg;      // Header line background
        COLORREF headerFg;      // Header text color
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

    // ------------------- Construction & State -----------------
    explicit ResultDock(HINSTANCE hInst);
    ResultDock(const ResultDock&) = delete;
    ResultDock& operator=(const ResultDock&) = delete;

    void create(const NppData& npp);
    void initFolding() const;
    void applyTheme();

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
    static constexpr int STYLE_FILEPATH          = 34;
};
