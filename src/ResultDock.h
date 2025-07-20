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

    /// Build list‑view text (grouped OR flat) from pre‑aggregated files
    void buildListText(const std::unordered_map<std::string, FileAgg>& files,
        bool flatView,
        const std::wstring& header,
        const SciSendFn& sciSend,
        std::wstring& outText,
        std::vector<Hit>& outHits) const;

private:
    /* Central colour palette for ResultDock*/
    struct RDColors
    {
        /* Hit‑line background (already used) */
        static constexpr COLORREF LineBgLight = RGB(0xE7, 0xF2, 0xFF);
        static constexpr COLORREF LineBgDark = RGB(0x3A, 0x3D, 0x33);

        /* Line‑number digits (already used) */
        static constexpr COLORREF LineNrLight = RGB(0x80, 0xC0, 0xFF);
        static constexpr COLORREF LineNrDark = RGB(0x80, 0xC0, 0xFF);

        /* Match substring (already used) */
        static constexpr COLORREF MatchLight = RGB(0xA6, 0xE2, 0x2E);
        static constexpr COLORREF MatchDark = RGB(0xA6, 0xE2, 0x2E);

        /* NEW – first headline:  Search "..." (…) */
        static constexpr COLORREF HeaderBgLight = RGB(0x79, 0x94, 0x86);   // pale teal
        static constexpr COLORREF HeaderBgDark = RGB(0x2E, 0x3D, 0x36);   // deep teal

        /* NEW – file path line (4‑space indent) */
        static constexpr COLORREF FilePathFgLight = RGB(0xC8, 0xAE, 0x6F); // khaki
        static constexpr COLORREF FilePathFgDark = RGB(0xEB, 0xCB, 0x8B); // sand

        static constexpr COLORREF HeaderFg = RGB(0, 0, 0); //black

        /* Fold markers: glyph (“+”/“-”/lines) */
        static constexpr COLORREF FoldGlyphLight = RGB(80, 80, 80);
        static constexpr COLORREF FoldGlyphDark = RGB(128, 128, 128);    // dark‑mode glyph

        static constexpr COLORREF FoldBoxLight = FoldGlyphLight;      // light theme
        static constexpr COLORREF FoldBoxDark = FoldGlyphDark;

        static constexpr COLORREF FoldHiMint = RGB(121, 148, 134);
    };

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
    static constexpr int INDIC_HEADER_BACKGROUND = 11;   // headline back‑colour
    static constexpr int INDIC_FILEPATH_FORE = 12;   // file‑path foreground
    static constexpr int INDIC_HEADER_FORE = 13;        // NEW – text colour
    static constexpr int MARKER_HEADER_BACKGROUND = 24;

};
