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


private:
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
};
