#pragma once
/* ------------------------------------------------------------------
 * ResultDock – dockable Scintilla view that shows search results
 * ------------------------------------------------------------------ */

#include <windows.h>
#include <string>
#include <vector>
#include <functional>
#include "Encoding.h"        // UTF‑16/UTF‑8 helpers
#include "Sci_Position.h"
#include "PluginDefinition.h"
#include "StaticDialog/DockingDlgInterface.h"

 /* ------------------------------------------------------------------
  *  Helper types
  * ------------------------------------------------------------------ */
using SciSendFn = std::function<LRESULT(UINT, WPARAM, LPARAM)>;

/* ------------------------------------------------------------------
 *  ResultDock  (singleton)
 * ------------------------------------------------------------------ */
class ResultDock final
{
public:
    /* === public data ================================================= */

    struct Hit          // one visible “Line N:” entry in the dock
    {
        /* navigation */
        std::string  fullPathUtf8;     // UTF‑8 file path
        Sci_Position pos{};            // match start in source buffer
        Sci_Position length{};         // match length

        /* styling offsets (within dock buffer) */
        int displayLineStart{ -1 };    // absolute char pos of "Line N:"
        int numberStart{ 0 };          // offset of digits
        int numberLen{ 0 };            // length of digits
        std::vector<int> matchStarts;  // offsets of each match substring
        std::vector<int> matchLens;    // lengths of each match substring
    };

    /* === singleton =================================================== */
    static ResultDock& instance();

    /* === public API ================================================== */
    void ensureCreatedAndVisible(const NppData& npp);

    void setText(const std::wstring& text);
    void prependText(const std::wstring& text);
    void appendText(const std::wstring& text);

    void prependHits(const std::vector<Hit>& newHits, const std::wstring& dockText);
    void recordHit(const std::string& fullPathUtf8, Sci_Position pos, Sci_Position length);

    void clear();                           // Clears all hits and text
    void rebuildFolding();       // refresh folding markers
    void applyStyling()        const;       // colour indicators
    void onThemeChanged();                  // react to N++ dark‑mode switch

    const std::vector<Hit>& hits() const { return _hits; }

    /* build dock text for exactly one file                               */
    void formatHitsForFile(const std::wstring& wFilePath,
        const SciSendFn& sciSend,
        std::vector<Hit>& hitsInOut,
        std::wstring& outBlock,
        size_t& ioUtf8Len) const;

    void formatHitsLines(const SciSendFn& sciSend,
        std::vector<Hit>& hitsInOut,
        std::wstring& outBlock,
        size_t& ioUtf8Len) const;

    // ─── for buildListText() ───────────────────────────────────────
    struct CritAgg { std::wstring text; std::vector<Hit> hits; };
    struct FileAgg { std::wstring wPath; int hitCount = 0; std::vector<CritAgg> crits; };
    using FileMap = std::unordered_map<std::string, FileAgg>;

    /// Build list‑view text (grouped OR flat) from pre‑aggregated files
    void buildListText(const FileMap& files,
        bool flatView,
        const std::wstring& header,
        const SciSendFn& sciSend,
        std::wstring& outText,
        std::vector<Hit>& outHits) const;

private:
    struct DockThemeColors {
        COLORREF lineBg;
        COLORREF lineNr;
        COLORREF matchFg;
        COLORREF matchBg;
        COLORREF headerBg;
        COLORREF headerFg;
        COLORREF filePathFg;
        COLORREF foldGlyph;
        COLORREF foldHighlight;
        COLORREF caretLineBg;
        int caretLineAlpha;
    };

    static constexpr DockThemeColors LightDockTheme = {
        RGB(0xEE, 0xEE, 0xEE), // lineBg
        RGB(0x40, 0x80, 0xBF), // lineNr
        RGB(0xFA, 0x3F, 0x34), // matchFg
        RGB(0xFF, 0xFF, 0xBF), // matchBg
        RGB(0xD5, 0xFF, 0xD5), // headerBg
        RGB(0x00, 0x00, 0x00), // headerFg
        RGB(0xA0, 0x80, 0x50), // filePathFg
        RGB(0x50, 0x50, 0x50), // foldGlyph
        RGB(0xFF, 0x00, 0x00), // foldHighlight
        RGB(0xEE, 0xEE, 0xFF), // caretLineBg
        64                     // caretLineAlpha
    };

    static constexpr DockThemeColors DarkDockTheme = {
        RGB(0x3A, 0x3D, 0x33), // lineBg
        RGB(0x80, 0xC0, 0xFF), // lineNr
        RGB(0xA6, 0xE2, 0x2E), // matchFg
        RGB(0x00, 0x00, 0x00), // matchBg (unused in dark mode)
        RGB(0x8F, 0xAF, 0x9F), // headerBg
        RGB(0x00, 0x00, 0x00), // headerFg
        RGB(0xEB, 0xCB, 0x8B), // filePathFg
        RGB(0x80, 0x80, 0x80), // foldGlyph
        RGB(0x79, 0x94, 0x86), // foldHighlight
        RGB(0x44, 0x44, 0x44), // caretLineBg
        64                    // caretLineAlpha
    };

    static constexpr const DockThemeColors& currentColors(bool darkMode) {
        return darkMode ? DarkDockTheme : LightDockTheme;
    }

    /* === construction ================================================= */
    explicit ResultDock(HINSTANCE hInst);
    ResultDock(const ResultDock&) = delete;
    ResultDock& operator=(const ResultDock&) = delete;

    /* === private helpers ============================================= */
    void create(const NppData& npp);
    void initFolding() const;   
    void applyTheme();

    HWND                 _hScintilla = nullptr;
    std::wstring         _content;          // full display text
    std::vector<int>     _foldHeaders;      // line numbers that start a fold block

    static int  leadingSpaces(const char* line, int len);

    static LRESULT CALLBACK sciSubclassProc(HWND, UINT, WPARAM, LPARAM);

    static inline WNDPROC   s_prevSciProc = nullptr;

    tTbData _dockData{}; // Holds docking info persistently

    /* === data members ================================================= */
    HINSTANCE   _hInst{ nullptr };         // DLL instance
    HWND        _hSci{ nullptr };         // Scintilla handle (client)
    HWND        _hDock{ nullptr };         // N++ dock container

    std::vector<Hit> _hits;                 // visible hits

    /* indicator IDs (Scintilla) */
    static constexpr int INDIC_LINE_BACKGROUND = 8;
    static constexpr int INDIC_LINENUMBER_FORE = 9;
    static constexpr int INDIC_MATCH_FORE = 10;
    static constexpr int INDIC_MATCH_BG = 14;
    static constexpr int INDIC_HEADER_BACKGROUND = 11;   // headline back‑colour
    static constexpr int INDIC_FILEPATH_FORE = 12;   // file‑path foreground
    static constexpr int INDIC_HEADER_FORE = 13;        // NEW – text colour
    static constexpr int MARKER_HEADER_BACKGROUND = 24;
    static constexpr int STYLE_HEADER = 33;
    static constexpr int STYLE_FILEPATH = 34;

};
