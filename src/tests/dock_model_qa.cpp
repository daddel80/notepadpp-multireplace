// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
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

// QA harness for the result dock's hit model: search blocks built by the new code (RawHit and
// CritAgg per search, one RowMatch list per dock line, line numbers and matches as styles) must
// give the same dock as the code of a43e162 (a full Hit per match, seven vectors per dock line,
// line numbers and matches as indicators). Both run on the real Scintilla document of
// Notepad++ 697f46b; the window calls around it are stand-ins.
// Compared after every search block and after a full restyle:
//   text     dock text and fold levels
//   rows     per dock line: path, position, length, line, flags, find text, line start, number
//            range; every match of the line with find text, position, length, flags; the shown
//            ones with their range in the dock line and color
//   color    per byte: header styles, line numbers, matches, per-entry backgrounds
//   index    line start to hit index; positions after FlowTabs padding is added and removed
// At the end of every dock:
//   wrap     line breaks of every dock line by Scintilla's LineLayout::WrapLine at four widths,
//            a43e162 with SC_WRAP_WORD, the new dock with the wrap mode ../ResultDock.cpp sets
//   delete   the dock's Delete key on random lines, blocks and selections, each followed by the
//            comparisons above
// Then long lines in code page 1252, a hit every 32 bytes, compared as above and timed: a43e162
// converts a line up to every hit again, the new dock must be at least 20 times as fast.
// Searches as Find in Files runs them: files of the given folders and generated texts (UTF-8 and
// code page 1252, lines over the 2048-byte display cap, control characters, soft hyphens, no-break
// spaces), one to four criteria with match case, whole word and regex, merged and grouped,
// standard and per-entry colors, FlowTabs padding in the searched document, several blocks.
// The new functions are verbatim copies of ../ResultDock.cpp and ../MultiReplacePanel.cpp, checked
// at the start; the old ones are those of a43e162.
//
// Needs the Scintilla and Boost regex objects of find_in_files_bench.sh (<work> is its folder).
// Build (from src/tests, w = <work>):
//   g++ -std=c++17 -O1 -Wall -Wextra -fsanitize=address,undefined -I.. -I$w/npp/scintilla/include
//       -I$w/npp/scintilla/src dock_model_qa.cpp $w/npp/scintilla/src/PositionCache.cxx
//       $w/npp/scintilla/src/Selection.cxx $w/obj/*.o -o dock_model_qa
// Run:   ./dock_model_qa [folder ...]
#include <algorithm>
#include <cctype>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <forward_list>
#include <fstream>
#include <functional>
#include <iterator>
#include <map>
#include <memory>
#include <mutex>
#include <optional>
#include <random>
#include <set>
#include <stdexcept>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "ScintillaTypes.h"
#include "ScintillaMessages.h"
#include "ILoader.h"
#include "ILexer.h"
#include "Debugging.h"
#include "Geometry.h"
#include "Platform.h"
#include "CharacterType.h"
#include "CharacterCategoryMap.h"
#include "Position.h"
#include "UniqueString.h"
#include "SplitVector.h"
#include "Partitioning.h"
#include "RunStyles.h"
#include "ContractionState.h"
#include "CellBuffer.h"
#include "PerLine.h"
#include "KeyMap.h"
#include "Indicator.h"
#include "LineMarker.h"
#include "Style.h"
#include "ViewStyle.h"
#include "CharClassify.h"
#include "Decoration.h"
#include "CaseFolder.h"
#include "Document.h"
#include "UniConversion.h"
#include "DBCS.h"
#include "Selection.h"
#include "PositionCache.h"

#include "Scintilla.h"   // MultiReplace's: SCI_*, STYLE_*, INDIC_*, SC_FOLD*

// Regex execution flags of Notepad++'s Boost bridge, as MultiReplacePanel.h defines them
constexpr int SCFIND_REGEXP_SKIPCRLFASONE = 0x08000000;
constexpr int SCFIND_REGEXP_EMPTYMATCH_ALL = 0x40000000;

namespace Scintilla::Internal {
void Platform::Assert(const char* c, const char* file, int line) noexcept { std::fprintf(stderr, "Assertion [%s] failed at %s %d\n", c, file, line); std::abort(); }
void Platform::DebugPrintf(const char*, ...) noexcept {}
void Platform::DebugDisplay(const char*) noexcept {}
bool Platform::ShowAssertionPopUps(bool) noexcept { return false; }
ColourRGBA Platform::Chrome() { return ColourRGBA(0xe0, 0xe0, 0xe0); }
ColourRGBA Platform::ChromeHighlight() { return ColourRGBA(0xff, 0xff, 0xff); }
const char* Platform::DefaultFont() { return "monospace"; }
int Platform::DefaultFontSize() { return 10; }
unsigned int Platform::DoubleClickTime() { return 500; }
// PositionCache.cxx's BreakFinder needs it at link time; LineLayout::WrapLine never calls it
bool ViewStyle::DrawCaretInsideSelection(bool, bool) const noexcept { return false; }
}

// ---------------------------------------------------------------------------------------------
// Stand-ins for Win32 and the plugin parts the copied functions use
// ---------------------------------------------------------------------------------------------
using HWND = void*;
using UINT = unsigned int;
using WPARAM = uintptr_t;
using LPARAM = intptr_t;
using LRESULT = intptr_t;
using COLORREF = uint32_t;
using BOOL = int;
using BYTE = unsigned char;
constexpr BOOL TRUE = 1;
constexpr BOOL FALSE = 0;
constexpr UINT WM_SETREDRAW = 0x000B;
constexpr UINT RDW_INVALIDATE = 0x0001, RDW_ERASE = 0x0004, RDW_ALLCHILDREN = 0x0080, RDW_UPDATENOW = 0x0100;
constexpr UINT NPPM_ISDARKMODEENABLED = 0x0400 + 1000 + 107;
LRESULT sendToDockWindow(UINT m, WPARAM w, LPARAM l);
inline LRESULT SendMessage(HWND h, UINT m, WPARAM w, LPARAM l) {   // other windows: redraw switch, dark mode: light
    return h == reinterpret_cast<HWND>(1) ? sendToDockWindow(m, w, l) : 0;
}
inline BOOL RedrawWindow(HWND, const void*, void*, UINT) { return TRUE; }
inline BOOL UpdateWindow(HWND) { return TRUE; }
struct { HWND _nppHandle = nullptr; } nppData;
using SciSendFn = std::function<LRESULT(UINT, WPARAM, LPARAM)>;

struct LanguageManagerStandIn {
    std::wstring get(const std::wstring& key, const std::vector<std::wstring>& args = {}) const {
        if (key == L"dock_line") return L"Line";
        if (key == L"dock_hits_suffix") return L"(" + args.at(0) + L" hits)";
        if (key == L"dock_crit_header") return L"Search \"" + args.at(0) + L"\" (" + args.at(1) + L" hits)";
        return key;
    }
} LM;

namespace StringUtils {
    std::wstring formatNumber(size_t number) {
        std::wstring s = std::to_wstring(number);
        for (int i = static_cast<int>(s.size()) - 3; i > 0; i -= 3) s.insert(static_cast<size_t>(i), L".");
        return s;
    }
}

namespace Encoding {
    std::string wstringToUtf8(const std::wstring& w) {
        std::string out;
        for (wchar_t wc : w) {
            const uint32_t c = static_cast<uint32_t>(wc);
            if (c < 0x80) out += static_cast<char>(c);
            else if (c < 0x800) { out += static_cast<char>(0xC0 | (c >> 6)); out += static_cast<char>(0x80 | (c & 0x3F)); }
            else if (c < 0x10000) { out += static_cast<char>(0xE0 | (c >> 12)); out += static_cast<char>(0x80 | ((c >> 6) & 0x3F)); out += static_cast<char>(0x80 | (c & 0x3F)); }
            else { out += static_cast<char>(0xF0 | (c >> 18)); out += static_cast<char>(0x80 | ((c >> 12) & 0x3F)); out += static_cast<char>(0x80 | ((c >> 6) & 0x3F)); out += static_cast<char>(0x80 | (c & 0x3F)); }
        }
        return out;
    }
    // Code page 1252 read as Latin-1, enough for display text; UTF-8 passes through
    std::string bytesToUtf8(const char* data, size_t len, UINT cp) {
        if (cp == SC_CP_UTF8) return std::string(data, len);
        std::wstring w;
        for (size_t i = 0; i < len; ++i) w += static_cast<wchar_t>(static_cast<unsigned char>(data[i]));
        return wstringToUtf8(w);
    }
}

namespace ColumnTabs {
    int g_indicatorId = -1;
    int CT_GetIndicatorId() { return g_indicatorId; }
}

namespace {
    inline bool pathsEqualUtf8(const std::string& a, const std::string& b) { return strcasecmp(a.c_str(), b.c_str()) == 0; }
}

// A Scintilla document answering the SCI_* messages these functions send, as Editor does
class SciEmu {
public:
    SciEmu(bool styles, int codepage) {
        _d = new Scintilla::Internal::Document(static_cast<Scintilla::DocumentOption>(styles ? 0 : 1));
        _d->AddRef();
        _d->SetUndoCollection(false);
        _d->SetDBCSCodePage(codepage);
        if (codepage == SC_CP_UTF8) _d->SetCaseFolder(std::make_unique<Scintilla::Internal::CaseFolderUnicode>());
        else _d->SetCaseFolder(std::make_unique<Scintilla::Internal::CaseFolderTable>());
    }
    ~SciEmu() { _d->Release(); }
    SciEmu(const SciEmu&) = delete;
    SciEmu& operator=(const SciEmu&) = delete;
    Scintilla::Internal::Document& doc() { return *_d; }

    LRESULT send(UINT m, WPARAM w = 0, LPARAM l = 0) {
        using namespace Scintilla::Internal;
        switch (m) {
        case SCI_GETLENGTH: return _d->Length();
        case SCI_GETCODEPAGE: return _d->dbcsCodePage;
        case SCI_GETTABWIDTH: return _d->tabInChars;
        case SCI_LINEFROMPOSITION: return static_cast<intptr_t>(w) < 0 ? 0 : _d->SciLineFromPosition(static_cast<Sci::Position>(w));
        case SCI_POSITIONFROMLINE: return static_cast<Sci::Line>(w) > _d->LinesTotal() ? -1 : _d->LineStart(static_cast<Sci::Line>(w));
        case SCI_GETLINEENDPOSITION: return _d->LineEnd(static_cast<Sci::Line>(w));
        case SCI_LINELENGTH: return _d->LineStart(static_cast<Sci::Line>(w) + 1) - _d->LineStart(static_cast<Sci::Line>(w));
        case SCI_GETLINE: {
            const Sci::Position a = _d->LineStart(static_cast<Sci::Line>(w)), b = _d->LineStart(static_cast<Sci::Line>(w) + 1);
            if (l) _d->GetCharRange(reinterpret_cast<char*>(l), a, b - a);
            return b - a;
        }
        case SCI_GETLINEINDENTATION: return _d->GetLineIndentation(static_cast<Sci::Line>(w));
        case SCI_GETLINEINDENTPOSITION: return _d->GetLineIndentPosition(static_cast<Sci::Line>(w));
        case SCI_GETLINECOUNT: return _d->LinesTotal();
        case SCI_DELETERANGE: _d->DeleteChars(static_cast<Sci::Position>(w), static_cast<Sci::Position>(l)); return 0;
        case SCI_GETSELECTIONNANCHOR: return selAnchor;
        case SCI_GETCURRENTPOS: return selCaret;
        case SCI_POSITIONAFTER: return _d->MovePositionOutsideChar(static_cast<Sci::Position>(w) + 1, 1, true);
        case SCI_INSERTTEXT: {
            const char* t = reinterpret_cast<const char*>(l);
            _d->InsertString(static_cast<Sci::Position>(w), t, static_cast<Sci::Position>(std::strlen(t)));
            return 0;
        }
        case SCI_STARTSTYLING: _d->StartStyling(static_cast<Sci::Position>(w)); return 0;
        case SCI_SETSTYLING: _d->SetStyleFor(static_cast<Sci::Position>(w), static_cast<char>(l)); return 0;
        case SCI_SETSTYLINGEX: _d->SetStyles(static_cast<Sci::Position>(w), reinterpret_cast<const char*>(l)); return 0;
        case SCI_SETINDICATORCURRENT: _d->DecorationSetCurrentIndicator(static_cast<int>(w)); return 0;
        case SCI_INDICATORFILLRANGE: _d->DecorationFillRange(static_cast<Sci::Position>(w), _d->decorations->GetCurrentValue(), static_cast<Sci::Position>(l)); return 0;
        case SCI_INDICATORCLEARRANGE: _d->DecorationFillRange(static_cast<Sci::Position>(w), 0, static_cast<Sci::Position>(l)); return 0;
        case SCI_INDICATORVALUEAT: return _d->decorations->ValueAt(static_cast<int>(w), static_cast<Sci::Position>(l));
        case SCI_INDICATOREND: return _d->decorations->End(static_cast<int>(w), static_cast<Sci::Position>(l));
        case SCI_GETCHARACTERPOINTER: return reinterpret_cast<LRESULT>(_d->BufferPointer());
        case SCI_SETFOLDLEVEL: _d->SetLevel(static_cast<Sci::Line>(w), static_cast<int>(l)); return 0;
        case SCI_GETFOLDLEVEL: return static_cast<LRESULT>(_d->GetFoldLevel(static_cast<Sci::Line>(w)));
        case SCI_GETFOLDEXPANDED: return 1;
        default: return 0;   // view state: read-only flag, allocation, scrolling, folding display
        }
    }

    Sci::Position selAnchor = 0, selCaret = 0;   // the selection the Delete key acts on

private:
    Scintilla::Internal::Document* _d = nullptr;
};

// The dock window ::SendMessage reaches while a dock's static window functions run
static SciEmu* g_dockWindow = nullptr;
LRESULT sendToDockWindow(UINT m, WPARAM w, LPARAM l) { return g_dockWindow ? g_dockWindow->send(m, w, l) : 0; }

// ---------------------------------------------------------------------------------------------
// a43e162
// ---------------------------------------------------------------------------------------------
namespace olddock {

class ResultDock final
{
public:
    // --------------------- Hit definition ---------------------
    // Represents one visible "Line N:" entry in the dock
    struct Hit
    {
        std::string  fullPathUtf8;
        Sci_Position pos{};
        Sci_Position length{};

        // For robust line-based navigation (FlowTabs-proof)
        int          docLine{ -1 };      // 0-based line number
        int          searchFlags{ 0 };   // Full flags as built by
        // MultiReplace::buildSearchFlags() (user options + regex execution
        // flags), stored verbatim so re-search reproduces match semantics.

        std::wstring findTextW;

        std::vector<std::wstring> allFindTexts;
        std::vector<Sci_Position> allPositions;
        std::vector<Sci_Position> allLengths;
        std::vector<int>          allSearchFlags;

        int displayLineStart{ -1 };
        int numberStart{ 0 };
        int numberLen{ 0 };

        std::vector<int> matchStarts;
        std::vector<int> matchLens;

        std::vector<int> matchColors;

        int colorIndex{ -1 };
    };

    struct CritAgg { std::wstring text; std::vector<Hit> hits; };
    struct FileAgg {
        std::wstring wPath;
        int hitCount = 0;
        std::vector<CritAgg> crits;
    };
    using FileMap = std::unordered_map<std::string, FileAgg>;

    void startSearchBlock(const std::wstring& header, bool  groupView, bool purge);
    void appendFileBlock(FileMap& fm, const SciSendFn& sciSend);
    void closeSearchBlock(int totalHits, int totalFiles, const std::wstring& scanSuffix = L"");
    void applyStyling() const;
    void adjustHitPositionsForFlowTab(const std::string& filePathUtf8,
        const std::vector<std::pair<Sci_Position, Sci_Position>>& paddingRanges, bool added);
    void clear() {}

    void applyStylingRange(Sci_Position pos0, Sci_Position len, const std::vector<Hit>& newHits) const;
    void applyHitIndicators(const std::vector<Hit>& hits) const;
    void markAllStyled() const;
    void rebuildFoldingRange(int firstLine, int lastLine, const std::string& dockTextU8) const;
    void prependBlock(const std::string& dockTextU8, std::vector<Hit>& newHits);
    void buildListText(FileMap& files, bool groupView, const std::wstring& header, const SciSendFn& sciSend,
        std::string& outTextU8, std::vector<Hit>& outHits) const;
    void formatHitsLines(const SciSendFn& sciSend, std::vector<Hit>& hitsInOut, std::string& outBlockU8) const;
    enum class LineLevel : int { SearchHdr = 0, FileHdr = 1, CritHdr = 2, HitLine = 3 };
    static constexpr int INDENT_SPACES[] = { 1, 2, 3, 4 };
    static std::string  getIndentStringU8(LineLevel lvl);
    static size_t getIndentUtf8Length(LineLevel lvl);
    static void shiftHits(std::vector<Hit>& v, size_t delta);
    void rebuildHitLineIndex();
    void rebuildFolding();
    enum class LineKind { Blank, SearchHdr, FileHdr, CritHdr, HitLine };
    static LineKind classify(const std::string& raw);
    static int leadingSpaces(const char* line, int len);
    static bool hasContentBeyondIndent(HWND hSci, int line);
    static void deleteSelectedItems(HWND hSci);
    using SciFnDirect_t = sptr_t (*)(sptr_t, unsigned int, uptr_t, sptr_t);
    inline static ResultDock* s_instance = nullptr;
    static ResultDock& instance() { return *s_instance; }

    static constexpr int ENTRY_BG_ALPHA_LIGHT = 150;
    static constexpr int ENTRY_OUTLINE_ALPHA_LIGHT = 0;
    static constexpr int ENTRY_BG_ALPHA_DARK = 130;
    static constexpr int ENTRY_OUTLINE_ALPHA_DARK = 0;
    static constexpr int MAX_ENTRY_COLORS = 28;
    static constexpr int INDIC_ENTRY_BG_BASE = 0;
    static constexpr int INDIC_LINENUMBER_FORE = 29;
    static constexpr int INDIC_MATCH_FORE = 30;
    static constexpr int STYLE_HEADER = 33;
    static constexpr int STYLE_CRITHDR = 34;
    static constexpr int STYLE_FILEPATH = 35;

    SciEmu* emu = nullptr;
    HWND _hSci = reinterpret_cast<HWND>(1);
    std::unordered_map<int, int> _slotToColor;
    std::vector<Hit> _hits;
    std::string      _pendingText;
    std::vector<Hit> _pendingHits;
    bool             _groupViewPending = false;
    bool             _blockOpen = false;
    inline static bool _perEntryColorsEnabled = false;
    std::unordered_map<int, int> _lineStartToHitIndex;
    LRESULT S(UINT m, WPARAM w = 0, LPARAM l = 0) const { return emu->send(m, w, l); }
};

class MultiReplace {
public:
    void trimHitToFirstLine(const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend, ResultDock::Hit& h);
};

void ResultDock::startSearchBlock(const std::wstring& header, bool groupView, bool purge)
{
    if (purge) {
        clear();
        _slotToColor.clear();
    }

    _pendingText.clear();
    _pendingHits.clear();
    _groupViewPending = groupView;
    _blockOpen = true;

    // Intro line for the search header; hit/file numbers are patched in closeSearchBlock().
    _pendingText = getIndentStringU8(LineLevel::SearchHdr);
    _pendingText += Encoding::wstringToUtf8(header);
    _pendingText += "\r\n";
}

void ResultDock::appendFileBlock(FileMap& fm, const SciSendFn& sciSend)
{
    if (!_blockOpen) return;

    std::string  partText;
    std::vector<Hit> partHits;

    // build WITHOUT another search header
    buildListText(fm, _groupViewPending, L"", sciSend, partText, partHits);

    shiftHits(partHits, _pendingText.size());      // adjust offsets

    _pendingText += partText;

    _pendingHits.insert(_pendingHits.end(),
        std::make_move_iterator(partHits.begin()),
        std::make_move_iterator(partHits.end()));
}

void ResultDock::closeSearchBlock(int totalHits, int totalFiles, const std::wstring& scanSuffix)
{
    if (!_blockOpen) return;

    // Replace placeholders in the pending header line ONLY (first line)
    const std::string hitsStr = Encoding::wstringToUtf8(StringUtils::formatNumber(static_cast<size_t>(totalHits)));
    const std::string filesStr = Encoding::wstringToUtf8(StringUtils::formatNumber(static_cast<size_t>(totalFiles)));

    // Limit search to the first header line to avoid matching digits in
    // user search patterns or file content that follows
    const size_t headerEnd = _pendingText.find("\r\n");
    const size_t searchLimit = (headerEnd != std::string::npos) ? headerEnd : _pendingText.size();

    // Search backwards from the end of the header to find the placeholder numbers
    // which are always the last two number sequences before the closing parenthesis.
    // Pattern: "... (0 hits in 0 files)\r\n"
    //                ^hits       ^files
    // Note: digits '0'..'9' are ASCII (single-byte in UTF-8), and UTF-8 continuation
    // bytes never fall into that range, so byte-wise scanning is safe even when the
    // header contains non-ASCII translations.
    size_t p2End = std::string::npos;
    size_t p2Start = std::string::npos;
    {
        size_t pos = searchLimit;
        while (pos > 0) {
            --pos;
            const unsigned char c = static_cast<unsigned char>(_pendingText[pos]);
            if (c >= '0' && c <= '9') {
                p2End = pos + 1;
                while (pos > 0) {
                    const unsigned char p = static_cast<unsigned char>(_pendingText[pos - 1]);
                    if (p < '0' || p > '9') break;
                    --pos;
                }
                p2Start = pos;
                break;
            }
        }
    }

    // Find second-to-last digit sequence (hits count)
    size_t p1Start = std::string::npos;
    size_t p1End = std::string::npos;
    if (p2Start != std::string::npos && p2Start > 0) {
        size_t pos = p2Start;
        while (pos > 0) {
            --pos;
            const unsigned char c = static_cast<unsigned char>(_pendingText[pos]);
            if (c >= '0' && c <= '9') {
                p1End = pos + 1;
                while (pos > 0) {
                    const unsigned char p = static_cast<unsigned char>(_pendingText[pos - 1]);
                    if (p < '0' || p > '9') break;
                    --pos;
                }
                p1Start = pos;
                break;
            }
        }
    }

    // Replace files count first (higher position), then hits count.
    // delta tracked from replace lengths - no re-conversion of the buffer.
    ptrdiff_t deltaBytes = 0;
    if (p2Start != std::string::npos) {
        const ptrdiff_t oldLen = static_cast<ptrdiff_t>(p2End - p2Start);
        const ptrdiff_t newLen = static_cast<ptrdiff_t>(filesStr.size());
        _pendingText.replace(p2Start, p2End - p2Start, filesStr);
        deltaBytes += (newLen - oldLen);
    }
    if (p1Start != std::string::npos) {
        const ptrdiff_t oldLen = static_cast<ptrdiff_t>(p1End - p1Start);
        const ptrdiff_t newLen = static_cast<ptrdiff_t>(hitsStr.size());
        _pendingText.replace(p1Start, p1End - p1Start, hitsStr);
        deltaBytes += (newLen - oldLen);
    }

    // Append scan statistics AFTER the digit patching so the placeholder
    // scan above never picks up numbers from the suffix.
    if (!scanSuffix.empty()) {
        const std::string sufU8 = Encoding::wstringToUtf8(scanSuffix);
        const size_t hdrEnd = _pendingText.find("\r\n");
        const size_t insertAt = (hdrEnd != std::string::npos) ? hdrEnd : _pendingText.size();
        _pendingText.insert(insertAt, sufU8);
        deltaBytes += static_cast<ptrdiff_t>(sufU8.size());
    }

    // Adjust hit offsets if the patched header changed length.
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

void ResultDock::applyStyling() const
{
    if (!_hSci) return;

    // 1) Clear fixed indicators
    const std::vector<int> indicatorsToClear = { INDIC_LINENUMBER_FORE, INDIC_MATCH_FORE };
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
    }

    // 4) Line numbers and matches of all hits
    applyHitIndicators(_hits);
    markAllStyled();
}

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
        const int ll = static_cast<int>(S(SCI_LINELENGTH, line));

        int style = STYLE_DEFAULT;
        if (ll > 0) {
            const int indent = static_cast<int>(S(SCI_GETLINEINDENTATION, line));
            if (indent == IND_SRCH) style = STYLE_HEADER;
            else if (indent == IND_CRIT) style = STYLE_CRITHDR;
            else if (indent == IND_FILE) style = STYLE_FILEPATH;
        }
        if (ll > 0) S(SCI_SETSTYLING, ll, style);
    }

    // Indicators only on the freshly added hits
    applyHitIndicators(newHits);
    markAllStyled();
}

void ResultDock::applyHitIndicators(const std::vector<Hit>& hits) const
{
    S(SCI_SETINDICATORCURRENT, INDIC_LINENUMBER_FORE);
    for (const auto& h : hits)
        if (h.displayLineStart >= 0)
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.numberStart, h.numberLen);

    // Match Highlighting (Exclusive Logic)
    if (_perEntryColorsEnabled) {
        // CASE A: Colorful Backgrounds -> Apply ONLY background indicators (Text remains standard/white)
        for (const auto& h : hits) {
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
        for (const auto& h : hits) {
            if (h.displayLineStart < 0) continue;
            for (size_t i = 0; i < h.matchStarts.size(); ++i) {
                S(SCI_INDICATORFILLRANGE, h.displayLineStart + h.matchStarts[i], h.matchLens[i]);
            }
        }
    }
}

void ResultDock::markAllStyled() const
{
    S(SCI_STARTSTYLING, S(SCI_GETLENGTH), 0);
}

void ResultDock::rebuildFoldingRange(int firstLine, int lastLine, const std::string& dockTextU8) const
{
    // Note: dockTextU8 is expected to start at the document position of `firstLine`,
    // i.e. its byte 0 corresponds to the first character of `firstLine`.
    if (!_hSci || lastLine < firstLine || dockTextU8.empty()) return;

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

    // Walk dockTextU8 ourselves: indent = leading ASCII spaces, line ends at '\n' (or buffer end).
    // hasContent = anything non-space, non-CR, non-LF after indent before the line break.
    // This avoids per-line SCI_LINELENGTH / SCI_GETLINEINDENTPOSITION / SCI_GETLINEENDPOSITION /
    // SCI_GETLINEINDENTATION / SCI_GETFOLDLEVEL round-trips.
    const size_t N = dockTextU8.size();
    size_t pos = 0;
    int line = firstLine;

    while (pos < N && line <= lastLine) {
        // Count leading ASCII spaces
        size_t indent = 0;
        while (pos < N && dockTextU8[pos] == ' ') { ++indent; ++pos; }

        // Determine whether anything beyond indent exists before line end
        bool hasBeyond = (pos < N) && (dockTextU8[pos] != '\r') && (dockTextU8[pos] != '\n');

        // Skip to end of line
        while (pos < N && dockTextU8[pos] != '\n') ++pos;
        if (pos < N) ++pos;  // skip the '\n'

        int target = BASE;
        if (hasBeyond) {
            bool isHdr = false;
            const int lvl = calcLevel((int)indent, isHdr);
            target = isHdr ? (lvl | SC_FOLDLEVELHEADERFLAG) : lvl;
        }

        S(SCI_SETFOLDLEVEL, line, target);

        ++line;
    }

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);
}

void ResultDock::prependBlock(const std::string& dockTextU8, std::vector<Hit>& newHits)
{
    if (!_hSci || dockTextU8.empty())
        return;

    const Sci_Position oldLen = (Sci_Position)S(SCI_GETLENGTH);

    const int sepBytes = (oldLen > 0 ? 2 : 0);
    const int deltaBytes = (int)dockTextU8.size() + sepBytes;

    // Shift existing hits
    for (auto& h : _hits)
        h.displayLineStart += deltaBytes;

    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);
    S(SCI_SETREADONLY, FALSE);
    S(SCI_ALLOCATE, oldLen + (Sci_Position)deltaBytes + 65536, 0);

    // Move cursor to position 0 before inserting to prevent auto-scroll to old cursor position
    S(SCI_SETEMPTYSELECTION, 0, 0);

    // Insert new block at start
    S(SCI_INSERTTEXT, 0, (sptr_t)dockTextU8.c_str());

    // Defensive: recompute length before inserting separator
    if (sepBytes) {
        const Sci_Position lenAfterBlock = (Sci_Position)S(SCI_GETLENGTH);
        const Sci_Position sepPos = (Sci_Position)dockTextU8.size();
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
    const Sci_Position len = (Sci_Position)dockTextU8.size();

    int newBlockLines = 0;
    for (char c : dockTextU8) if (c == '\n') ++newBlockLines;
    const int firstLine = 0;
    const int lastLine = (newBlockLines > 0 ? newBlockLines - 1 : 0);

    rebuildFoldingRange(firstLine, lastLine, dockTextU8);


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

void ResultDock::buildListText(
    FileMap& files,
    bool groupView,
    const std::wstring& header,
    const SciSendFn& sciSend,
    std::string& outTextU8,
    std::vector<Hit>& outHits) const
{
    std::string body;

    auto appendIndented = [&](LineLevel lvl, const std::wstring& txt)
        {
            // Header text comes from translations and is converted once per call.
            const std::string txtU8 = Encoding::wstringToUtf8(txt);
            body.append(getIndentUtf8Length(lvl), ' ');
            body.append(txtU8);
            body.append("\r\n", 2);
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
        appendIndented(LineLevel::FileHdr, f.wPath + L" " + LM.get(L"dock_hits_suffix", { StringUtils::formatNumber(static_cast<size_t>(f.hitCount)) }));

        if (groupView)
        {
            // grouped: first the file header, then each criterion block
            for (auto& c : f.crits)
            {
                appendIndented(LineLevel::CritHdr, LM.get(L"dock_crit_header", { c.text, StringUtils::formatNumber(c.hits.size()) }));

                // Move c.hits into hitsCopy: each crit is processed once,
                // and c.hits is not used afterwards.
                std::vector<Hit> hitsCopy = std::move(c.hits);

                // pass the copy to formatHitsLines instead of const c.hits
                formatHitsLines(sciSend, hitsCopy, body);

                // move results from hitsCopy into outHits
                outHits.insert(outHits.end(),
                    std::make_move_iterator(hitsCopy.begin()),
                    std::make_move_iterator(hitsCopy.end()));
            }
        }
        else
        {
            // Fast path: a single criterion means hits are already in pos-order
            // (sequential forward search) with no possible duplicates. Move the
            // hits vector to avoid the copy of 200k+ elements.
            // The FileMap is a local in the caller and is not used after this call.
            std::vector<Hit> merged;
            if (f.crits.size() == 1) {
                merged = std::move(f.crits.front().hits);
            }
            else {
                // Multi-criterion: collect, sort by position, dedup overlaps
                size_t total = 0;
                for (const auto& c : f.crits) total += c.hits.size();
                merged.reserve(total);
                for (auto& c : f.crits) {
                    // Move from each crit (still safe: iterating once, c.hits not used again)
                    for (auto& hit : c.hits) {
                        merged.push_back(std::move(hit));
                    }
                }
                std::sort(merged.begin(), merged.end(),
                    [](auto const& a, auto const& b) {
                        if (a.pos != b.pos) return a.pos < b.pos;
                        if (a.length != b.length) return a.length < b.length;

                        return a.findTextW < b.findTextW;
                    });
                merged.erase(
                    std::unique(merged.begin(), merged.end(),
                        [](const Hit& a, const Hit& b) { return a.pos == b.pos && a.length == b.length; }),
                    merged.end());
            }

            formatHitsLines(sciSend, merged, body);

            outHits.insert(outHits.end(),
                std::make_move_iterator(merged.begin()),
                std::make_move_iterator(merged.end()));
        }
    }

    outTextU8 = std::move(body);
}

void ResultDock::formatHitsLines(const SciSendFn& sciSend,
    std::vector<Hit>& hits,
    std::string& out) const
{
    const UINT docCp = (UINT)sciSend(SCI_GETCODEPAGE, 0, 0);
    const bool isUtf8Doc = (docCp == SC_CP_UTF8);

    const std::string  indentHitU8_str = getIndentStringU8(LineLevel::HitLine);
    const size_t       indentHitU8 = indentHitU8_str.size();

    const std::wstring kLineW = LM.get(L"dock_line") + L" ";
    const std::string  kLineU8_str = Encoding::wstringToUtf8(kLineW);  // converted once per call
    const size_t       kLineU8 = kLineU8_str.size();
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

    // Helper to count digits in a number (avoids std::to_string allocation)
    auto countDigits = [](int n) -> size_t {
        if (n <= 0) return 1;
        size_t count = 0;
        while (n > 0) { ++count; n /= 10; }
        return count;
        };

    auto appendIntU8 = [](std::string& dst, int n) {
        if (n == 0) { dst.push_back('0'); return; }
        char buf[16];
        int len = 0;
        bool neg = (n < 0);
        unsigned int v = neg ? static_cast<unsigned int>(-(n + 1)) + 1u : static_cast<unsigned int>(n);
        while (v > 0) {
            buf[len++] = static_cast<char>('0' + (v % 10u));
            v /= 10u;
        }
        if (neg) buf[len++] = '-';
        for (int i = 0, j = len - 1; i < j; ++i, --j) std::swap(buf[i], buf[j]);
        dst.append(buf, len);
        };

    int   prevDocLine = -1;
    Hit* firstHitOnRow = nullptr;
    size_t hitIdx = 0;

    std::string  cachedRaw;
    std::string  cachedRawFiltered;  // Only populated when FlowTabs padding exists
    std::string  origU8Owned;          // populated only for non-UTF-8 docs
    const std::string* origU8Ptr = nullptr; // points at effectiveRaw (UTF-8) or origU8Owned
    std::string  displayU8;
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
            cachedRaw.resize(rawLen + 1);
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

            // Point origU8Ptr at effectiveRaw (UTF-8 doc) or own a converted copy.
            if (isUtf8Doc) {
                origU8Ptr = effectiveRaw;
            }
            else {
                origU8Owned = Encoding::bytesToUtf8(effectiveRaw->data(), effectiveRaw->size(), docCp);
                origU8Ptr = &origU8Owned;
            }
            const std::string& origU8 = *origU8Ptr;

            // classify helpers (codepoint-based; both old paths used the same rules)
            auto isCtlCp = [](uint32_t cp)->bool {
                if (cp == 0x0000) return true;
                if ((cp <= 0x1Fu) && cp != 0x09 && cp != 0x0A && cp != 0x0D) return true;
                if (cp == 0x007Fu) return true;
                if (cp >= 0x80u && cp <= 0x9Fu) return true;
                if (cp == 0x2028 || cp == 0x2029) return true;
                return false;
                };
            auto isNbspCp = [](uint32_t cp)->bool { return cp == 0x00A0; };

            // Single-pass transform on UTF-8: build displayU8 and mapOrigToDisp
            // in one walk over origU8. Same cleanup rules as before
            // (SHY drop / NBSP+Ctl -> ASCII space / else verbatim).
            // The walk stops past the display cap; later offsets map beyond the capped text.
            displayU8.clear();
            mapOrigToDisp.clear();

            size_t o = 0, d = 0;
            while (o < origU8.size() && d <= kMaxHitTextUtf8)
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
                    mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d);
                    // no d++, no append
                }
                else if (isNbspCp(cp) || isCtlCp(cp)) {
                    mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d);
                    displayU8.push_back(' ');
                    ++d; // one visible ASCII space per replaced codepoint
                }
                else {
                    for (size_t k = 0; k < clen; ++k)
                        mapOrigToDisp.push_back(d + k);
                    displayU8.append(origU8, o, clen);
                    d += clen;
                }

                o += clen;
            }
            mapOrigToDisp.push_back(d);

            // finalize
            capUtf8WithEllipsis(displayU8, kMaxHitTextUtf8);

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
            const size_t line1Digits = countDigits(line1);
            const size_t padCount = (maxDigits > line1Digits) ? (maxDigits - line1Digits) : 0;
            const size_t prefixU8Len = indentHitU8 + kLineU8 + maxDigits + kColonSpaceU8;

            // out.size() before appending is the UTF-8 byte position where this row starts.
            const size_t rowStartPos = out.size();

            // Build the prefix + body directly into out, no temp strings.
            // Layout: indent + "Line " + (pad spaces) + (digits of line1) + ": " + display + "\r\n"
            out.append(indentHitU8_str);
            out.append(kLineU8_str);
            out.append(padCount, ' ');
            appendIntU8(out, line1);
            out.append(": ", 2);
            out.append(displayU8);
            out.append("\r\n", 2);

            // Look ahead: how many subsequent hits share this line1?
            // We know lineNumbers[] is in pos-order, and pos-order matches
            // by-line aggregation. Reserve exact capacity to avoid reallocations
            // in the per-hit push_back chain below.
            size_t lineHitCount = 1;
            for (size_t j = hitIdx; j < lineNumbers.size() && lineNumbers[j] == line1; ++j)
                ++lineHitCount;

            h.displayLineStart = (int)rowStartPos;
            h.numberStart = (int)(indentHitU8 + kLineU8 + maxDigits - line1Digits);
            h.numberLen = (int)line1Digits;

            h.matchStarts.clear();
            h.matchLens.clear();
            h.matchColors.clear();
            h.matchStarts.reserve(lineHitCount);
            h.matchLens.reserve(lineHitCount);
            h.matchColors.reserve(lineHitCount);

            firstHitOnRow = &h;
            firstHitOnRow->allFindTexts.clear();
            firstHitOnRow->allPositions.clear();
            firstHitOnRow->allLengths.clear();
            firstHitOnRow->allSearchFlags.clear();
            firstHitOnRow->allFindTexts.reserve(lineHitCount);
            firstHitOnRow->allPositions.reserve(lineHitCount);
            firstHitOnRow->allLengths.reserve(lineHitCount);
            firstHitOnRow->allSearchFlags.reserve(lineHitCount);
            firstHitOnRow->allFindTexts.push_back(h.findTextW);
            firstHitOnRow->allPositions.push_back(h.pos);
            firstHitOnRow->allLengths.push_back(h.length);
            firstHitOnRow->allSearchFlags.push_back(h.searchFlags);

            if (dispStart < displayU8.size()) {
                size_t safeLen = dispLen;
                if (dispStart + safeLen > displayU8.size()) safeLen = displayU8.size() - dispStart;
                if (safeLen > 0) {
                    h.matchStarts.push_back((int)(prefixU8Len + dispStart));
                    h.matchLens.push_back((int)safeLen);

                    h.matchColors.push_back(h.colorIndex);
                }
            }

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
                firstHitOnRow->allSearchFlags.push_back(h.searchFlags);
            }
            h.displayLineStart = -1;
        }
    }

    hits.erase(std::remove_if(hits.begin(), hits.end(),
        [](const Hit& e) { return e.displayLineStart < 0; }),
        hits.end());

}

std::string ResultDock::getIndentStringU8(LineLevel lvl) {
    return std::string(INDENT_SPACES[static_cast<int>(lvl)], ' ');
}

size_t ResultDock::getIndentUtf8Length(ResultDock::LineLevel lvl) {
    // Indent is ASCII spaces only, so the UTF-8 byte count equals the
    // wide-character count - no conversion needed.
    return static_cast<size_t>(INDENT_SPACES[static_cast<int>(lvl)]);
}

void ResultDock::shiftHits(std::vector<ResultDock::Hit>& v, size_t delta)
{
    for (auto& h : v)
        h.displayLineStart += static_cast<int>(delta);
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

int ResultDock::leadingSpaces(const char* line, int len)
{
    int s = 0;
    while (s < len && line[s] == ' ') ++s;
    return s;
}

bool ResultDock::hasContentBeyondIndent(HWND hSci, int line)
{
    int lineEnd = static_cast<int>(::SendMessage(hSci, SCI_GETLINEENDPOSITION, line, 0));
    int indentPos = static_cast<int>(::SendMessage(hSci, SCI_GETLINEINDENTPOSITION, line, 0));

    return indentPos < lineEnd; // true if there is content after indentation
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

void MultiReplace::trimHitToFirstLine(
    const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend,
    ResultDock::Hit& h)
{
    // Only regex matches can span line boundaries.
    // Normal, Extended, and WholeWord searches always stay within a single line,
    // so trimming is unnecessary and we can skip 3 Scintilla calls per hit.
    if (!(h.searchFlags & SCFIND_REGEXP))
        return;

    // Quick check: if hit doesn't extend beyond line end, no trimming needed.
    // This avoids further work for regex matches that stay within one line.
    const int lineZero = (h.docLine >= 0)
        ? h.docLine
        : static_cast<int>(sciSend(SCI_LINEFROMPOSITION, h.pos, 0));
    const Sci_Position lineStart = sciSend(SCI_POSITIONFROMLINE, lineZero, 0);
    const Sci_Position lineEnd = sciSend(SCI_GETLINEENDPOSITION, lineZero, 0);

    // If match fits entirely within the line content (before EOL), no trim needed
    if (h.pos >= lineStart && (h.pos + h.length) <= lineEnd)
        return;

    // Match extends beyond line end - need to trim
    if (h.pos >= lineEnd) {
        // Match starts at or after EOL
        h.length = std::min<Sci_Position>(h.length, 1);
        return;
    }

    // Match spans from within line content past EOL - trim to line end
    h.length = std::max<Sci_Position>(0, lineEnd - h.pos);
}

}

// ---------------------------------------------------------------------------------------------
// Now: hit model and styling as in ../ResultDock.h, ../ResultDock.cpp, ../MultiReplacePanel.cpp
// ---------------------------------------------------------------------------------------------
namespace newdock {

class ResultDock final
{
public:
    // --------------------- Hit definition ---------------------
    // One match as a search collects it. Find text, flags and color are those of its CritAgg.
    struct RawHit
    {
        Sci_Position pos{};
        Sci_Position length{};
        int          docLine{ -1 };      // 0-based line number
    };

    // One match on a dock line
    struct RowMatch
    {
        std::shared_ptr<const std::wstring> findTextW;   // shared by all matches of a search
        Sci_Position pos{};
        Sci_Position length{};
        int searchFlags{ 0 };
        int colorIndex{ -1 };
        int dispStart{ -1 };   // offset in the dock line, -1 when beyond the shown text
        int dispLen{ 0 };
    };

    // Represents one visible "Line N:" entry in the dock. Its first match is the jump target.
    struct Hit
    {
        std::string  fullPathUtf8;
        Sci_Position pos{};
        Sci_Position length{};

        // For robust line-based navigation (FlowTabs-proof)
        int          docLine{ -1 };      // 0-based line number
        int          searchFlags{ 0 };   // Full flags as built by
        // MultiReplace::buildSearchFlags() (user options + regex execution
        // flags), stored verbatim so re-search reproduces match semantics.

        std::wstring findTextW;

        std::vector<RowMatch> matches;   // every match on this line, the first one included

        int displayLineStart{ -1 };
        int numberStart{ 0 };
        int numberLen{ 0 };
    };

    // Cursor position info: not needed here
    struct CritAgg {
        std::wstring text;                               // label in the dock
        std::shared_ptr<const std::wstring> findTextW;   // as searched
        int searchFlags = 0;
        int colorIndex = -1;
        std::vector<RawHit> hits;
    };
    struct FileAgg {
        std::wstring wPath;
        int hitCount = 0;
        std::vector<CritAgg> crits;
    };
    using FileMap = std::unordered_map<std::string, FileAgg>;

    void startSearchBlock(const std::wstring& header, bool  groupView, bool purge);
    void appendFileBlock(FileMap& fm, const SciSendFn& sciSend);
    void closeSearchBlock(int totalHits, int totalFiles, const std::wstring& scanSuffix = L"");
    void applyStyling() const;
    void adjustHitPositionsForFlowTab(const std::string& filePathUtf8,
        const std::vector<std::pair<Sci_Position, Sci_Position>>& paddingRanges, bool added);
    void clear() {}

    void applyStylingRange(Sci_Position pos0, std::string_view text, const std::vector<Hit>& hits) const;
    void applyHitIndicators(const std::vector<Hit>& hits) const;
    void markAllStyled() const;
    void rebuildFoldingRange(int firstLine, int lastLine, const std::string& dockTextU8) const;
    void prependBlock(const std::string& dockTextU8, std::vector<Hit>& newHits);
    void buildListText(FileMap& files, bool groupView, const std::wstring& header, const SciSendFn& sciSend,
        std::string& outTextU8, std::vector<Hit>& outHits) const;
    struct SourceHit { const RawHit* hit; const CritAgg* crit; };
    void formatHitsLines(const SciSendFn& sciSend, const std::string& pathUtf8, const std::vector<SourceHit>& hits,
        std::string& outBlockU8, std::vector<Hit>& outRows) const;
    enum class LineLevel : int { SearchHdr = 0, FileHdr = 1, CritHdr = 2, HitLine = 3 };
    static constexpr int INDENT_SPACES[] = { 1, 2, 3, 4 };
    static std::string  getIndentStringU8(LineLevel lvl);
    static size_t getIndentUtf8Length(LineLevel lvl);
    static void shiftHits(std::vector<Hit>& v, size_t delta);
    void rebuildHitLineIndex();
    void rebuildFolding();
    enum class LineKind { Blank, SearchHdr, FileHdr, CritHdr, HitLine };
    static LineKind classify(const std::string& raw);
    static int leadingSpaces(const char* line, int len);
    static bool hasContentBeyondIndent(HWND hSci, int line);
    static void deleteSelectedItems(HWND hSci);
    using SciFnDirect_t = sptr_t (*)(sptr_t, unsigned int, uptr_t, sptr_t);
    inline static ResultDock* s_instance = nullptr;
    static ResultDock& instance() { return *s_instance; }

    static constexpr int ENTRY_BG_ALPHA_LIGHT = 150;
    static constexpr int ENTRY_OUTLINE_ALPHA_LIGHT = 0;
    static constexpr int ENTRY_BG_ALPHA_DARK = 130;
    static constexpr int ENTRY_OUTLINE_ALPHA_DARK = 0;
    static constexpr int MAX_ENTRY_COLORS = 28;
    static constexpr int INDIC_ENTRY_BG_BASE = 0;
    static constexpr int STYLE_HEADER = 33;
    static constexpr int STYLE_CRITHDR = 34;
    static constexpr int STYLE_FILEPATH = 35;
    static constexpr int STYLE_HIT_NUMBER = 40;
    static constexpr int STYLE_HIT_MATCH = 41;

    SciEmu* emu = nullptr;
    HWND _hSci = reinterpret_cast<HWND>(1);
    std::unordered_map<int, int> _slotToColor;
    std::vector<Hit> _hits;
    std::string      _pendingText;
    std::vector<Hit> _pendingHits;
    bool             _groupViewPending = false;
    bool             _blockOpen = false;
    inline static bool _perEntryColorsEnabled = false;
    std::unordered_map<int, int> _lineStartToHitIndex;
    LRESULT S(UINT m, WPARAM w = 0, LPARAM l = 0) const { return emu->send(m, w, l); }
};

class MultiReplace {
public:
    void trimHitToFirstLine(const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend, ResultDock::RawHit& h, int searchFlags);
};

// ===== VERBATIM from ../ResultDock.cpp and ../MultiReplacePanel.cpp
void ResultDock::startSearchBlock(const std::wstring& header, bool groupView, bool purge)
{
    if (purge) {
        clear();
        _slotToColor.clear();
    }

    _pendingText.clear();
    _pendingHits.clear();
    _groupViewPending = groupView;
    _blockOpen = true;

    // Intro line for the search header; hit/file numbers are patched in closeSearchBlock().
    _pendingText = getIndentStringU8(LineLevel::SearchHdr);
    _pendingText += Encoding::wstringToUtf8(header);
    _pendingText += "\r\n";
}

void ResultDock::appendFileBlock(FileMap& fm, const SciSendFn& sciSend)
{
    if (!_blockOpen) return;

    std::string  partText;
    std::vector<Hit> partHits;

    // build WITHOUT another search header
    buildListText(fm, _groupViewPending, L"", sciSend, partText, partHits);

    shiftHits(partHits, _pendingText.size());      // adjust offsets

    _pendingText += partText;

    _pendingHits.insert(_pendingHits.end(),
        std::make_move_iterator(partHits.begin()),
        std::make_move_iterator(partHits.end()));
}

void ResultDock::closeSearchBlock(int totalHits, int totalFiles, const std::wstring& scanSuffix)
{
    if (!_blockOpen) return;

    // Replace placeholders in the pending header line ONLY (first line)
    const std::string hitsStr = Encoding::wstringToUtf8(StringUtils::formatNumber(static_cast<size_t>(totalHits)));
    const std::string filesStr = Encoding::wstringToUtf8(StringUtils::formatNumber(static_cast<size_t>(totalFiles)));

    // Limit search to the first header line to avoid matching digits in
    // user search patterns or file content that follows
    const size_t headerEnd = _pendingText.find("\r\n");
    const size_t searchLimit = (headerEnd != std::string::npos) ? headerEnd : _pendingText.size();

    // Search backwards from the end of the header to find the placeholder numbers
    // which are always the last two number sequences before the closing parenthesis.
    // Pattern: "... (0 hits in 0 files)\r\n"
    //                ^hits       ^files
    // Note: digits '0'..'9' are ASCII (single-byte in UTF-8), and UTF-8 continuation
    // bytes never fall into that range, so byte-wise scanning is safe even when the
    // header contains non-ASCII translations.
    size_t p2End = std::string::npos;
    size_t p2Start = std::string::npos;
    {
        size_t pos = searchLimit;
        while (pos > 0) {
            --pos;
            const unsigned char c = static_cast<unsigned char>(_pendingText[pos]);
            if (c >= '0' && c <= '9') {
                p2End = pos + 1;
                while (pos > 0) {
                    const unsigned char p = static_cast<unsigned char>(_pendingText[pos - 1]);
                    if (p < '0' || p > '9') break;
                    --pos;
                }
                p2Start = pos;
                break;
            }
        }
    }

    // Find second-to-last digit sequence (hits count)
    size_t p1Start = std::string::npos;
    size_t p1End = std::string::npos;
    if (p2Start != std::string::npos && p2Start > 0) {
        size_t pos = p2Start;
        while (pos > 0) {
            --pos;
            const unsigned char c = static_cast<unsigned char>(_pendingText[pos]);
            if (c >= '0' && c <= '9') {
                p1End = pos + 1;
                while (pos > 0) {
                    const unsigned char p = static_cast<unsigned char>(_pendingText[pos - 1]);
                    if (p < '0' || p > '9') break;
                    --pos;
                }
                p1Start = pos;
                break;
            }
        }
    }

    // Replace files count first (higher position), then hits count.
    // delta tracked from replace lengths - no re-conversion of the buffer.
    ptrdiff_t deltaBytes = 0;
    if (p2Start != std::string::npos) {
        const ptrdiff_t oldLen = static_cast<ptrdiff_t>(p2End - p2Start);
        const ptrdiff_t newLen = static_cast<ptrdiff_t>(filesStr.size());
        _pendingText.replace(p2Start, p2End - p2Start, filesStr);
        deltaBytes += (newLen - oldLen);
    }
    if (p1Start != std::string::npos) {
        const ptrdiff_t oldLen = static_cast<ptrdiff_t>(p1End - p1Start);
        const ptrdiff_t newLen = static_cast<ptrdiff_t>(hitsStr.size());
        _pendingText.replace(p1Start, p1End - p1Start, hitsStr);
        deltaBytes += (newLen - oldLen);
    }

    // Append scan statistics AFTER the digit patching so the placeholder
    // scan above never picks up numbers from the suffix.
    if (!scanSuffix.empty()) {
        const std::string sufU8 = Encoding::wstringToUtf8(scanSuffix);
        const size_t hdrEnd = _pendingText.find("\r\n");
        const size_t insertAt = (hdrEnd != std::string::npos) ? hdrEnd : _pendingText.size();
        _pendingText.insert(insertAt, sufU8);
        deltaBytes += static_cast<ptrdiff_t>(sufU8.size());
    }

    // Adjust hit offsets if the patched header changed length.
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

void ResultDock::applyStyling() const
{
    if (!_hSci) return;

    // 1) Clear dynamic entry indicators (Range 0 to MAX_ENTRY_COLORS)
    for (int i = 0; i < MAX_ENTRY_COLORS; ++i) {
        S(SCI_SETINDICATORCURRENT, INDIC_ENTRY_BG_BASE + i);
        S(SCI_INDICATORCLEARRANGE, 0, S(SCI_GETLENGTH));
    }

    // 2) Configure Dynamic Indicators based on _slotToColor map
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
    }

    // 3) Styles of the whole text: headers, line numbers, matches
    const Sci_Position length = S(SCI_GETLENGTH);
    const char* text = reinterpret_cast<const char*>(S(SCI_GETCHARACTERPOINTER));
    if (text && length > 0)
        applyStylingRange(0, std::string_view(text, static_cast<size_t>(length)), _hits);
    markAllStyled();
}

void ResultDock::applyStylingRange(Sci_Position pos0, std::string_view text, const std::vector<Hit>& hits) const
{
    if (!_hSci || text.empty()) return;

    const int IND_SRCH = INDENT_SPACES[(int)LineLevel::SearchHdr];
    const int IND_FILE = INDENT_SPACES[(int)LineLevel::FileHdr];
    const int IND_CRIT = INDENT_SPACES[(int)LineLevel::CritHdr];
    const int tabWidth = (std::max)(1, static_cast<int>(S(SCI_GETTABWIDTH)));

    std::string styles(text.size(), static_cast<char>(STYLE_DEFAULT));
    for (size_t lineStart = 0; lineStart < text.size();) {
        const size_t eol = text.find_first_of("\r\n", lineStart);
        const size_t next = (eol == std::string_view::npos) ? text.size()
            : (text[eol] == '\r' && eol + 1 < text.size() && text[eol + 1] == '\n') ? eol + 2 : eol + 1;
        int indent = 0;
        for (size_t i = lineStart; i < text.size() && (text[i] == ' ' || text[i] == '\t'); ++i)
            indent = (text[i] == ' ') ? indent + 1 : (indent / tabWidth + 1) * tabWidth;
        int style = STYLE_DEFAULT;
        if (indent == IND_SRCH) style = STYLE_HEADER;
        else if (indent == IND_CRIT) style = STYLE_CRITHDR;
        else if (indent == IND_FILE) style = STYLE_FILEPATH;
        if (style != STYLE_DEFAULT)
            std::fill(styles.begin() + static_cast<ptrdiff_t>(lineStart), styles.begin() + static_cast<ptrdiff_t>(next), static_cast<char>(style));
        lineStart = next;
    }

    auto paint = [&](Sci_Position docPos, Sci_Position len, int style) {
        const Sci_Position from = docPos - pos0;
        if (from < 0 || len <= 0 || static_cast<size_t>(from) >= styles.size()) return;
        const size_t n = (std::min)(static_cast<size_t>(len), styles.size() - static_cast<size_t>(from));
        std::fill_n(styles.begin() + from, n, static_cast<char>(style));
    };
    for (const Hit& h : hits) {
        if (h.displayLineStart < 0) continue;
        paint(h.displayLineStart + h.numberStart, h.numberLen, STYLE_HIT_NUMBER);
        if (_perEntryColorsEnabled) continue;
        for (const RowMatch& m : h.matches)
            if (m.dispStart >= 0) paint(h.displayLineStart + m.dispStart, m.dispLen, STYLE_HIT_MATCH);
    }

    S(SCI_STARTSTYLING, pos0, 0);
    S(SCI_SETSTYLINGEX, styles.size(), reinterpret_cast<LPARAM>(styles.data()));

    applyHitIndicators(hits);
}

void ResultDock::applyHitIndicators(const std::vector<Hit>& hits) const
{
    if (!_perEntryColorsEnabled) return;
    for (const auto& h : hits) {
        if (h.displayLineStart < 0) continue;
        for (const RowMatch& m : h.matches) {
            if (m.dispStart < 0 || m.colorIndex < 0 || m.colorIndex >= MAX_ENTRY_COLORS) continue;
            S(SCI_SETINDICATORCURRENT, INDIC_ENTRY_BG_BASE + m.colorIndex);
            S(SCI_INDICATORFILLRANGE, h.displayLineStart + m.dispStart, m.dispLen);
        }
    }
}

void ResultDock::markAllStyled() const
{
    S(SCI_STARTSTYLING, S(SCI_GETLENGTH), 0);
}

void ResultDock::rebuildFoldingRange(int firstLine, int lastLine, const std::string& dockTextU8) const
{
    // Note: dockTextU8 is expected to start at the document position of `firstLine`,
    // i.e. its byte 0 corresponds to the first character of `firstLine`.
    if (!_hSci || lastLine < firstLine || dockTextU8.empty()) return;

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

    // Walk dockTextU8 ourselves: indent = leading ASCII spaces, line ends at '\n' (or buffer end).
    // hasContent = anything non-space, non-CR, non-LF after indent before the line break.
    // This avoids per-line SCI_LINELENGTH / SCI_GETLINEINDENTPOSITION / SCI_GETLINEENDPOSITION /
    // SCI_GETLINEINDENTATION / SCI_GETFOLDLEVEL round-trips.
    const size_t N = dockTextU8.size();
    size_t pos = 0;
    int line = firstLine;

    while (pos < N && line <= lastLine) {
        // Count leading ASCII spaces
        size_t indent = 0;
        while (pos < N && dockTextU8[pos] == ' ') { ++indent; ++pos; }

        // Determine whether anything beyond indent exists before line end
        bool hasBeyond = (pos < N) && (dockTextU8[pos] != '\r') && (dockTextU8[pos] != '\n');

        // Skip to end of line
        while (pos < N && dockTextU8[pos] != '\n') ++pos;
        if (pos < N) ++pos;  // skip the '\n'

        int target = BASE;
        if (hasBeyond) {
            bool isHdr = false;
            const int lvl = calcLevel((int)indent, isHdr);
            target = isHdr ? (lvl | SC_FOLDLEVELHEADERFLAG) : lvl;
        }

        S(SCI_SETFOLDLEVEL, line, target);

        ++line;
    }

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);
}

void ResultDock::prependBlock(const std::string& dockTextU8, std::vector<Hit>& newHits)
{
    if (!_hSci || dockTextU8.empty())
        return;

    const Sci_Position oldLen = (Sci_Position)S(SCI_GETLENGTH);

    const int sepBytes = (oldLen > 0 ? 2 : 0);
    const int deltaBytes = (int)dockTextU8.size() + sepBytes;

    // Shift existing hits
    for (auto& h : _hits)
        h.displayLineStart += deltaBytes;

    ::SendMessage(_hSci, WM_SETREDRAW, FALSE, 0);
    S(SCI_SETREADONLY, FALSE);
    S(SCI_ALLOCATE, oldLen + (Sci_Position)deltaBytes + 65536, 0);

    // Move cursor to position 0 before inserting to prevent auto-scroll to old cursor position
    S(SCI_SETEMPTYSELECTION, 0, 0);

    // Insert new block at start
    S(SCI_INSERTTEXT, 0, (sptr_t)dockTextU8.c_str());

    // Defensive: recompute length before inserting separator
    if (sepBytes) {
        const Sci_Position lenAfterBlock = (Sci_Position)S(SCI_GETLENGTH);
        const Sci_Position sepPos = (Sci_Position)dockTextU8.size();
        if (sepPos <= lenAfterBlock) {
            S(SCI_INSERTTEXT, (uptr_t)sepPos, (sptr_t)"\r\n");
        }
        // If that doesn't fit, no insert -> avoids invalid position.
    }


    S(SCI_SETREADONLY, TRUE);

    // Reset horizontal scroll to prevent scrolling to old cursor position on long lines
    S(SCI_SETXOFFSET, 0, 0);

    ::SendMessage(_hSci, WM_SETREDRAW, TRUE, 0);


    int newBlockLines = 0;
    for (char c : dockTextU8) if (c == '\n') ++newBlockLines;
    const int firstLine = 0;
    const int lastLine = (newBlockLines > 0 ? newBlockLines - 1 : 0);

    rebuildFoldingRange(firstLine, lastLine, dockTextU8);

    applyStylingRange(0, dockTextU8, newHits);
    if (sepBytes && S(SCI_GETLENGTH) >= (Sci_Position)dockTextU8.size() + sepBytes) {
        S(SCI_STARTSTYLING, (Sci_Position)dockTextU8.size(), 0);
        S(SCI_SETSTYLING, sepBytes, STYLE_DEFAULT);
    }
    markAllStyled();


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

void ResultDock::buildListText(
    FileMap& files,
    bool groupView,
    const std::wstring& header,
    const SciSendFn& sciSend,
    std::string& outTextU8,
    std::vector<Hit>& outHits) const
{
    std::string body;

    auto appendIndented = [&](LineLevel lvl, const std::wstring& txt)
        {
            // Header text comes from translations and is converted once per call.
            const std::string txtU8 = Encoding::wstringToUtf8(txt);
            body.append(getIndentUtf8Length(lvl), ' ');
            body.append(txtU8);
            body.append("\r\n", 2);
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

    std::vector<SourceHit> source;
    for (auto& [path, f] : files)
    {
        appendIndented(LineLevel::FileHdr, f.wPath + L" " + LM.get(L"dock_hits_suffix", { StringUtils::formatNumber(static_cast<size_t>(f.hitCount)) }));

        if (groupView)
        {
            // grouped: first the file header, then each criterion block
            for (const auto& c : f.crits)
            {
                appendIndented(LineLevel::CritHdr, LM.get(L"dock_crit_header", { c.text, StringUtils::formatNumber(c.hits.size()) }));
                source.clear();
                source.reserve(c.hits.size());
                for (const RawHit& h : c.hits) source.push_back({ &h, &c });
                formatHitsLines(sciSend, path, source, body, outHits);
            }
        }
        else
        {
            // A single criterion is in position order already (sequential forward search).
            // Several: by position, a match found by more than one criterion shown once.
            size_t total = 0;
            for (const auto& c : f.crits) total += c.hits.size();
            source.clear();
            source.reserve(total);
            for (const auto& c : f.crits)
                for (const RawHit& h : c.hits) source.push_back({ &h, &c });
            if (f.crits.size() > 1) {
                std::sort(source.begin(), source.end(),
                    [](const SourceHit& a, const SourceHit& b) {
                        if (a.hit->pos != b.hit->pos) return a.hit->pos < b.hit->pos;
                        if (a.hit->length != b.hit->length) return a.hit->length < b.hit->length;
                        return *a.crit->findTextW < *b.crit->findTextW;
                    });
                source.erase(
                    std::unique(source.begin(), source.end(),
                        [](const SourceHit& a, const SourceHit& b) { return a.hit->pos == b.hit->pos && a.hit->length == b.hit->length; }),
                    source.end());
            }
            formatHitsLines(sciSend, path, source, body, outHits);
        }
    }

    outTextU8 = std::move(body);
}

void ResultDock::formatHitsLines(const SciSendFn& sciSend,
    const std::string& pathUtf8,
    const std::vector<SourceHit>& hits,
    std::string& out,
    std::vector<Hit>& rows) const
{
    const UINT docCp = (UINT)sciSend(SCI_GETCODEPAGE, 0, 0);
    const bool isUtf8Doc = (docCp == SC_CP_UTF8);

    const std::string  indentHitU8_str = getIndentStringU8(LineLevel::HitLine);
    const size_t       indentHitU8 = indentHitU8_str.size();

    const std::wstring kLineW = LM.get(L"dock_line") + L" ";
    const std::string  kLineU8_str = Encoding::wstringToUtf8(kLineW);  // converted once per call
    const size_t       kLineU8 = kLineU8_str.size();
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
    for (const SourceHit& h : hits) {
        const int line0 = (h.hit->docLine >= 0) ? h.hit->docLine
            : static_cast<int>(sciSend(SCI_LINEFROMPOSITION, h.hit->pos, 0));
        const int line1 = line0 + 1;
        lineNumbers.push_back(line1);
        // Count digits without creating temporary string
        int temp = line1;
        size_t digits = 0;
        do { ++digits; temp /= 10; } while (temp > 0);
        if (digits > maxDigits) maxDigits = digits;
    }

    // Helper to count digits in a number (avoids std::to_string allocation)
    auto countDigits = [](int n) -> size_t {
        if (n <= 0) return 1;
        size_t count = 0;
        while (n > 0) { ++count; n /= 10; }
        return count;
        };

    auto appendIntU8 = [](std::string& dst, int n) {
        if (n == 0) { dst.push_back('0'); return; }
        char buf[16];
        int len = 0;
        bool neg = (n < 0);
        unsigned int v = neg ? static_cast<unsigned int>(-(n + 1)) + 1u : static_cast<unsigned int>(n);
        while (v > 0) {
            buf[len++] = static_cast<char>('0' + (v % 10u));
            v /= 10u;
        }
        if (neg) buf[len++] = '-';
        for (int i = 0, j = len - 1; i < j; ++i, --j) std::swap(buf[i], buf[j]);
        dst.append(buf, len);
        };

    int   prevDocLine = -1;
    size_t hitIdx = 0;

    std::string  cachedRaw;
    std::string  cachedRawFiltered;  // Only populated when FlowTabs padding exists
    std::string  origU8Owned;          // populated only for non-UTF-8 docs
    const std::string* origU8Ptr = nullptr; // points at effectiveRaw (UTF-8) or origU8Owned
    std::string  displayU8;
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
    size_t u8PrefixBytes = 0;   // non-UTF-8 docs: filtered bytes of the line converted so far
    size_t u8PrefixLen = 0;     // and their UTF-8 length

    auto loadLineIfNeeded = [&](int line0)
        {
            if (line0 == prevDocLine)
                return;

            // raw line
            const int rawLen = static_cast<int>(sciSend(SCI_LINELENGTH, line0, 0));
            cachedRaw.resize(rawLen + 1);
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

            // Point origU8Ptr at effectiveRaw (UTF-8 doc) or own a converted copy.
            if (isUtf8Doc) {
                origU8Ptr = effectiveRaw;
            }
            else {
                origU8Owned = Encoding::bytesToUtf8(effectiveRaw->data(), effectiveRaw->size(), docCp);
                origU8Ptr = &origU8Owned;
            }
            const std::string& origU8 = *origU8Ptr;

            // classify helpers (codepoint-based; both old paths used the same rules)
            auto isCtlCp = [](uint32_t cp)->bool {
                if (cp == 0x0000) return true;
                if ((cp <= 0x1Fu) && cp != 0x09 && cp != 0x0A && cp != 0x0D) return true;
                if (cp == 0x007Fu) return true;
                if (cp >= 0x80u && cp <= 0x9Fu) return true;
                if (cp == 0x2028 || cp == 0x2029) return true;
                return false;
                };
            auto isNbspCp = [](uint32_t cp)->bool { return cp == 0x00A0; };

            // Single-pass transform on UTF-8: build displayU8 and mapOrigToDisp
            // in one walk over origU8. Same cleanup rules as before
            // (SHY drop / NBSP+Ctl -> ASCII space / else verbatim).
            // The walk stops past the display cap; later offsets map beyond the capped text.
            displayU8.clear();
            mapOrigToDisp.clear();

            size_t o = 0, d = 0;
            while (o < origU8.size() && d <= kMaxHitTextUtf8)
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
                    mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d);
                    // no d++, no append
                }
                else if (isNbspCp(cp) || isCtlCp(cp)) {
                    mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d);
                    displayU8.push_back(' ');
                    ++d; // one visible ASCII space per replaced codepoint
                }
                else {
                    for (size_t k = 0; k < clen; ++k)
                        mapOrigToDisp.push_back(d + k);
                    displayU8.append(origU8, o, clen);
                    d += clen;
                }

                o += clen;
            }
            mapOrigToDisp.push_back(d);

            // finalize
            capUtf8WithEllipsis(displayU8, kMaxHitTextUtf8);

            u8PrefixBytes = 0;
            u8PrefixLen = 0;
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

    // Hits of a line come in position order: only the bytes since the previous hit are converted
    auto u8PrefixFromRawBytes = [&](size_t byteCount)->size_t {
        if (isUtf8Doc) return rawToFilteredPos(byteCount); // byte pos == UTF-8 pos
        const size_t filteredByteCount = rawToFilteredPos(byteCount);
        if (filteredByteCount < u8PrefixBytes) {
            u8PrefixBytes = 0;
            u8PrefixLen = 0;
        }
        u8PrefixLen += Encoding::bytesToUtf8(effectiveRaw->data() + u8PrefixBytes, filteredByteCount - u8PrefixBytes, docCp).size();
        u8PrefixBytes = filteredByteCount;
        return u8PrefixLen;
        };

    for (const SourceHit& source : hits) {
        const RawHit& h = *source.hit;
        const CritAgg& crit = *source.crit;
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

        const size_t prefixU8Len = indentHitU8 + kLineU8 + maxDigits + kColonSpaceU8;
        if (line0 != prevDocLine) {
            const size_t line1Digits = countDigits(line1);
            const size_t padCount = (maxDigits > line1Digits) ? (maxDigits - line1Digits) : 0;

            // out.size() before appending is the UTF-8 byte position where this row starts.
            const size_t rowStartPos = out.size();

            // Build the prefix + body directly into out, no temp strings.
            // Layout: indent + "Line " + (pad spaces) + (digits of line1) + ": " + display + "\r\n"
            out.append(indentHitU8_str);
            out.append(kLineU8_str);
            out.append(padCount, ' ');
            appendIntU8(out, line1);
            out.append(": ", 2);
            out.append(displayU8);
            out.append("\r\n", 2);

            // Look ahead: how many subsequent hits share this line1?
            // lineNumbers[] is in pos-order, so the matches of a row are adjacent.
            size_t lineHitCount = 1;
            for (size_t j = hitIdx; j < lineNumbers.size() && lineNumbers[j] == line1; ++j)
                ++lineHitCount;

            Hit& row = rows.emplace_back();
            row.fullPathUtf8 = pathUtf8;
            row.pos = h.pos;
            row.length = h.length;
            row.docLine = h.docLine;
            row.searchFlags = crit.searchFlags;
            row.findTextW = *crit.findTextW;
            row.displayLineStart = (int)rowStartPos;
            row.numberStart = (int)(indentHitU8 + kLineU8 + maxDigits - line1Digits);
            row.numberLen = (int)line1Digits;
            row.matches.reserve(lineHitCount);

            prevDocLine = line0;
        }

        // The match on the current row, shown when it starts within the displayed text
        RowMatch& m = rows.back().matches.emplace_back();
        m.findTextW = crit.findTextW;
        m.pos = h.pos;
        m.length = h.length;
        m.searchFlags = crit.searchFlags;
        m.colorIndex = crit.colorIndex;
        if (dispStart < displayU8.size()) {
            size_t safeLen = dispLen;
            if (dispStart + safeLen > displayU8.size()) safeLen = displayU8.size() - dispStart;
            if (safeLen > 0) {
                m.dispStart = (int)(prefixU8Len + dispStart);
                m.dispLen = (int)safeLen;
            }
        }
    }
}

std::string ResultDock::getIndentStringU8(LineLevel lvl) {
    return std::string(INDENT_SPACES[static_cast<int>(lvl)], ' ');
}

size_t ResultDock::getIndentUtf8Length(ResultDock::LineLevel lvl) {
    // Indent is ASCII spaces only, so the UTF-8 byte count equals the
    // wide-character count - no conversion needed.
    return static_cast<size_t>(INDENT_SPACES[static_cast<int>(lvl)]);
}

void ResultDock::shiftHits(std::vector<ResultDock::Hit>& v, size_t delta)
{
    for (auto& h : v)
        h.displayLineStart += static_cast<int>(delta);
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
        for (RowMatch& m : hit.matches)
        {
            d = computeDelta(m.pos);
            m.pos += sign * d;
            if (m.pos < 0) m.pos = 0;
        }
    }
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

int ResultDock::leadingSpaces(const char* line, int len)
{
    int s = 0;
    while (s < len && line[s] == ' ') ++s;
    return s;
}

bool ResultDock::hasContentBeyondIndent(HWND hSci, int line)
{
    int lineEnd = static_cast<int>(::SendMessage(hSci, SCI_GETLINEENDPOSITION, line, 0));
    int indentPos = static_cast<int>(::SendMessage(hSci, SCI_GETLINEINDENTPOSITION, line, 0));

    return indentPos < lineEnd; // true if there is content after indentation
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

void MultiReplace::trimHitToFirstLine(
    const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend,
    ResultDock::RawHit& h, int searchFlags)
{
    // Only regex matches can span line boundaries.
    // Normal, Extended, and WholeWord searches always stay within a single line,
    // so trimming is unnecessary and we can skip 3 Scintilla calls per hit.
    if (!(searchFlags & SCFIND_REGEXP))
        return;

    // Quick check: if hit doesn't extend beyond line end, no trimming needed.
    // This avoids further work for regex matches that stay within one line.
    const int lineZero = (h.docLine >= 0)
        ? h.docLine
        : static_cast<int>(sciSend(SCI_LINEFROMPOSITION, h.pos, 0));
    const Sci_Position lineStart = sciSend(SCI_POSITIONFROMLINE, lineZero, 0);
    const Sci_Position lineEnd = sciSend(SCI_GETLINEENDPOSITION, lineZero, 0);

    // If match fits entirely within the line content (before EOL), no trim needed
    if (h.pos >= lineStart && (h.pos + h.length) <= lineEnd)
        return;

    // Match extends beyond line end - need to trim
    if (h.pos >= lineEnd) {
        // Match starts at or after EOL
        h.length = std::min<Sci_Position>(h.length, 1);
        return;
    }

    // Match spans from within line content past EOL - trim to line end
    h.length = std::max<Sci_Position>(0, lineEnd - h.pos);
}

}

// ---------------------------------------------------------------------------------------------
// Copies against the sources
// ---------------------------------------------------------------------------------------------
static std::string readText(const std::filesystem::path& p) {
    std::ifstream in(p, std::ios::binary);
    std::string s((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    s.erase(std::remove(s.begin(), s.end(), '\r'), s.end());
    if (s.compare(0, 3, "\xEF\xBB\xBF") == 0) s.erase(0, 3);
    return s;
}

static std::string functionText(const std::string& text, const std::string& head, size_t from) {
    size_t p = text.find("\n" + head, from);
    if (p == std::string::npos) return {};
    ++p;
    const size_t e = text.find("\n}\n", p);
    return e == std::string::npos ? std::string() : text.substr(p, e + 2 - p);
}

static int verbatimCheck() {
    const std::filesystem::path here = std::filesystem::path(__FILE__).parent_path();
    const std::string self = readText(here / "dock_model_qa.cpp");
    const std::string dock = readText(here / ".." / "ResultDock.cpp");
    const std::string panel = readText(here / ".." / "MultiReplacePanel.cpp");
    const std::string header = readText(here / ".." / "ResultDock.h");
    if (self.empty() || dock.empty() || panel.empty() || header.empty()) { std::printf("V. copies vs sources: SKIP (sources not found)\n"); return 0; }
    const size_t mark = self.find("// ===== VERBATIM from ../ResultDock.cpp");
    static const char* heads[] = {
        "void ResultDock::startSearchBlock(",
        "void ResultDock::appendFileBlock(",
        "void ResultDock::closeSearchBlock(",
        "void ResultDock::applyStyling() const",
        "void ResultDock::applyStylingRange(",
        "void ResultDock::applyHitIndicators(",
        "void ResultDock::markAllStyled(",
        "void ResultDock::rebuildFoldingRange(",
        "void ResultDock::prependBlock(",
        "void ResultDock::buildListText(",
        "void ResultDock::formatHitsLines(",
        "std::string ResultDock::getIndentStringU8(",
        "size_t ResultDock::getIndentUtf8Length(",
        "void ResultDock::shiftHits(",
        "void ResultDock::rebuildHitLineIndex()",
        "void ResultDock::adjustHitPositionsForFlowTab(",
        "void ResultDock::rebuildFolding()",
        "ResultDock::LineKind ResultDock::classify(",
        "int ResultDock::leadingSpaces(",
        "bool ResultDock::hasContentBeyondIndent(",
        "void ResultDock::deleteSelectedItems("
    };
    int different = 0, compared = 0;
    for (const char* h : heads) {
        ++compared;
        if (functionText(self, h, mark) != functionText(dock, h, 0) || functionText(dock, h, 0).empty()) {
            ++different;
            std::printf("  differs: %s\n", h);
        }
    }
    ++compared;
    const char* trim = "void MultiReplace::trimHitToFirstLine(";
    if (functionText(self, trim, mark) != functionText(panel, trim, 0)) { ++different; std::printf("  differs: %s\n", trim); }
    // The types: from the hit definitions to the cursor info, and CritAgg to FileMap
    auto between = [](const std::string& s, const std::string& a, const std::string& b, size_t from) {
        const size_t i = s.find(a, from), j = (i == std::string::npos) ? std::string::npos : s.find(b, i);
        return (i == std::string::npos || j == std::string::npos) ? std::string() : s.substr(i, j - i);
    };
    const size_t newClass = self.find("namespace newdock {");
    for (auto [a, b] : { std::pair<std::string, std::string>{ "    // --------------------- Hit definition", "    // Cursor position info" },
                         std::pair<std::string, std::string>{ "    struct CritAgg", "    using FileMap" } }) {
        ++compared;
        const std::string mine = between(self, a, b, newClass), theirs = between(header, a, b, 0);
        if (mine.empty() || mine != theirs) { ++different; std::printf("  differs: types from %s\n", a.c_str() + 4); }
    }
    std::printf("V. copies vs sources: %d of %d identical\n", compared - different, compared);
    return different;
}

// The wrap mode ../ResultDock.cpp sets with SCI_SETWRAPMODE (create() and the wrap toggle); all must agree
static std::optional<Scintilla::Wrap> wrapModeOfSource(std::string& name) {
    const std::string dock = readText(std::filesystem::path(__FILE__).parent_path() / ".." / "ResultDock.cpp");
    std::set<std::string> modes;
    for (size_t p = dock.find("SCI_SETWRAPMODE,"); p != std::string::npos; p = dock.find("SCI_SETWRAPMODE,", p + 1)) {
        const size_t q = dock.find("? ", p);
        if (q == std::string::npos || q - p > 120) continue;
        modes.insert(dock.substr(q + 2, dock.find(' ', q + 2) - (q + 2)));
    }
    for (const std::string& m : modes) name += (name.empty() ? "" : " / ") + m;
    if (modes.size() != 1) return std::nullopt;
    if (name == "SC_WRAP_WORD") return Scintilla::Wrap::Word;
    if (name == "SC_WRAP_CHAR") return Scintilla::Wrap::Char;
    if (name == "SC_WRAP_WHITESPACE") return Scintilla::Wrap::WhiteSpace;
    return std::nullopt;
}

// ---------------------------------------------------------------------------------------------
// Searches as Find in Files runs them, into both docks
// ---------------------------------------------------------------------------------------------
struct Criterion {
    std::wstring findW;
    std::string bytes;
    int flags = 0;
    int color = 0;
};

struct SourceFile {
    std::string pathUtf8;
    std::wstring pathW;
    std::string text;
    int codepage = SC_CP_UTF8;
    std::vector<std::pair<Sci::Position, Sci::Position>> padding;   // FlowTabs padding (start, length)
};

static std::wstring widen(const std::string& u8) {
    std::wstring w;
    for (size_t i = 0; i < u8.size();) {
        const unsigned char c = static_cast<unsigned char>(u8[i]);
        uint32_t cp = c; size_t n = 1;
        if (c >= 0xC0 && c < 0xE0 && i + 1 < u8.size()) { cp = ((c & 0x1F) << 6) | (u8[i + 1] & 0x3F); n = 2; }
        else if (c >= 0xE0 && c < 0xF0 && i + 2 < u8.size()) { cp = ((c & 0x0F) << 12) | ((u8[i + 1] & 0x3F) << 6) | (u8[i + 2] & 0x3F); n = 3; }
        else if (c >= 0xF0 && i + 3 < u8.size()) { cp = ((c & 0x07) << 18) | ((u8[i + 1] & 0x3F) << 12) | ((u8[i + 2] & 0x3F) << 6) | (u8[i + 3] & 0x3F); n = 4; }
        w += static_cast<wchar_t>(cp);
        i += n;
    }
    return w;
}

// MultiReplace::sanitizeSearchPattern
static std::wstring label(const std::wstring& raw) {
    std::wstring s = raw;
    for (size_t p = 0; (p = s.find(L"\r", p)) != std::wstring::npos; p += 2) s.replace(p, 1, L"\\r");
    for (size_t p = 0; (p = s.find(L"\n", p)) != std::wstring::npos; p += 2) s.replace(p, 1, L"\\n");
    return s;
}

struct Docks {
    SciEmu oldEmu{ true, SC_CP_UTF8 }, newEmu{ true, SC_CP_UTF8 };
    olddock::ResultDock oldDock;
    newdock::ResultDock newDock;
    Docks() { oldDock.emu = &oldEmu; newDock.emu = &newEmu; }
};

static long long g_checks = 0, g_failures = 0, g_wrapLines = 0, g_wrapWordDiffers = 0, g_deletions = 0;
static double g_oldAppendSeconds = 0, g_newAppendSeconds = 0;
static void fail(const char* what, const std::string& detail) {
    if (++g_failures <= 15) std::printf("  FAIL %s: %s\n", what, detail.c_str());
}

// One search block over the files, the same hits fed into both docks
static void searchBlock(Docks& docks, std::vector<SourceFile>& files, const std::vector<Criterion>& crits, bool listMode, bool groupView) {
    const std::wstring header = listMode ? L"List search (0 hits in 0 files)" : L"Search \"" + label(crits.front().findW) + L"\" (0 hits in 0 files)";
    docks.oldDock.startSearchBlock(header, listMode ? groupView : false, false);
    docks.newDock.startSearchBlock(header, listMode ? groupView : false, false);
    int totalHits = 0, filesWithHits = 0;
    olddock::MultiReplace oldMr;
    newdock::MultiReplace newMr;
    std::vector<std::shared_ptr<const std::wstring>> shared;
    for (const Criterion& c : crits) shared.push_back(std::make_shared<const std::wstring>(c.findW));

    for (SourceFile& f : files) {
        SciEmu src(false, f.codepage);
        src.doc().InsertString(0, f.text.data(), static_cast<Sci::Position>(f.text.size()));
        if (ColumnTabs::g_indicatorId >= 0) {
            src.doc().DecorationSetCurrentIndicator(ColumnTabs::g_indicatorId);
            for (auto [p, n] : f.padding) src.doc().DecorationFillRange(p, 1, n);
        }
        const SciSendFn sciSend = [&src](UINT m, WPARAM w, LPARAM l) { return src.send(m, w, l); };
        const Sci::Position docLength = src.doc().Length();
        olddock::ResultDock::FileMap oldMap;
        newdock::ResultDock::FileMap newMap;
        int hitsInFile = 0;
        for (size_t ci = 0; ci < crits.size(); ++ci) {
            const Criterion& c = crits[ci];
            std::vector<olddock::ResultDock::Hit> oldRaw;
            std::vector<newdock::ResultDock::RawHit> newRaw;
            Sci::Position pos = 0;
            while (!c.bytes.empty() && oldRaw.size() < 20000) {
                Sci::Position len = static_cast<Sci::Position>(c.bytes.size());
                Sci::Position p = -1;
                try { p = src.doc().FindText(pos, docLength, c.bytes.c_str(), static_cast<Scintilla::FindOption>(c.flags), &len); }
                catch (...) { p = -1; }
                if (p < 0 || p + len < p || p + len > docLength) break;   // MR's performSingleSearch
                if (len > 0) pos = p + len;                               // MR's advanceAfterMatch
                else { const Sci::Position after = src.send(SCI_POSITIONAFTER, static_cast<WPARAM>(p), 0); pos = std::min(after > p ? after : p + 1, docLength); }
                const int line = static_cast<int>(src.send(SCI_LINEFROMPOSITION, static_cast<WPARAM>(p), 0));

                olddock::ResultDock::Hit h{};
                h.fullPathUtf8 = f.pathUtf8;
                h.pos = p;
                h.length = len;
                h.docLine = line;
                h.searchFlags = c.flags;
                oldMr.trimHitToFirstLine(sciSend, h);
                h.findTextW = c.findW;
                h.colorIndex = c.color;
                oldRaw.push_back(std::move(h));

                newdock::ResultDock::RawHit r{ p, len, line };
                newMr.trimHitToFirstLine(sciSend, r, c.flags);
                newRaw.push_back(r);
                if (p >= docLength) break;
            }
            const int n = static_cast<int>(oldRaw.size());
            if (n == 0) continue;
            auto& oa = oldMap[f.pathUtf8];
            oa.wPath = f.pathW; oa.hitCount += n;
            oa.crits.push_back({ label(c.findW), std::move(oldRaw) });
            auto& na = newMap[f.pathUtf8];
            na.wPath = f.pathW; na.hitCount += n;
            na.crits.push_back({ label(c.findW), shared[ci], c.flags, c.color, std::move(newRaw) });
            hitsInFile += n;
        }
        if (hitsInFile > 0) {
            const auto t0 = std::chrono::steady_clock::now();
            docks.oldDock.appendFileBlock(oldMap, sciSend);
            const auto t1 = std::chrono::steady_clock::now();
            docks.newDock.appendFileBlock(newMap, sciSend);
            g_oldAppendSeconds += std::chrono::duration<double>(t1 - t0).count();
            g_newAppendSeconds += std::chrono::duration<double>(std::chrono::steady_clock::now() - t1).count();
            totalHits += hitsInFile;
            ++filesWithHits;
        }
    }
    ++g_checks;
    if (docks.oldDock._pendingText != docks.newDock._pendingText) fail("block text", "differs before closing");
    docks.oldDock.closeSearchBlock(totalHits, filesWithHits, L" [3 files searched]");
    docks.newDock.closeSearchBlock(totalHits, filesWithHits, L" [3 files searched]");
}

static std::string describeHit(const char* side, size_t i, Sci::Position dls) {
    return std::string(side) + " row " + std::to_string(i) + " at " + std::to_string(dls);
}

static void compareRows(const Docks& docks) {
    const auto& o = docks.oldDock._hits;
    const auto& n = docks.newDock._hits;
    ++g_checks;
    if (o.size() != n.size()) { fail("rows", std::to_string(o.size()) + " vs " + std::to_string(n.size())); return; }
    for (size_t i = 0; i < o.size(); ++i) {
        const auto& a = o[i];
        const auto& b = n[i];
        ++g_checks;
        if (a.fullPathUtf8 != b.fullPathUtf8 || a.pos != b.pos || a.length != b.length || a.docLine != b.docLine ||
            a.searchFlags != b.searchFlags || a.findTextW != b.findTextW || a.displayLineStart != b.displayLineStart ||
            a.numberStart != b.numberStart || a.numberLen != b.numberLen) { fail("row fields", describeHit("", i, a.displayLineStart)
                + " old " + std::to_string(a.pos) + "/" + std::to_string(a.length) + "/" + std::to_string(a.docLine) + "/" + std::to_string(a.searchFlags) + "/" + Encoding::wstringToUtf8(a.findTextW) + "/" + std::to_string(a.displayLineStart)
                + " new " + std::to_string(b.pos) + "/" + std::to_string(b.length) + "/" + std::to_string(b.docLine) + "/" + std::to_string(b.searchFlags) + "/" + Encoding::wstringToUtf8(b.findTextW) + "/" + std::to_string(b.displayLineStart)); continue; }
        if (a.allFindTexts.size() != b.matches.size()) { fail("row matches", describeHit("", i, a.displayLineStart)); continue; }
        std::vector<std::tuple<int, int, int>> shown;
        for (size_t k = 0; k < b.matches.size(); ++k) {
            const auto& m = b.matches[k];
            if (!m.findTextW || *m.findTextW != a.allFindTexts[k] || m.pos != a.allPositions[k] || m.length != a.allLengths[k] ||
                m.searchFlags != a.allSearchFlags[k]) fail("match fields", describeHit("", i, a.displayLineStart) + " k=" + std::to_string(k)
                    + " old pos " + std::to_string(a.allPositions[k]) + " len " + std::to_string(a.allLengths[k]) + " flags " + std::to_string(a.allSearchFlags[k])
                    + " new pos " + std::to_string(m.pos) + " len " + std::to_string(m.length) + " flags " + std::to_string(m.searchFlags)
                    + " texts " + Encoding::wstringToUtf8(a.allFindTexts[k]) + "|" + (m.findTextW ? Encoding::wstringToUtf8(*m.findTextW) : std::string("null")));
            if (m.dispStart >= 0) shown.emplace_back(m.dispStart, m.dispLen, m.colorIndex);
        }
        if (b.matches.empty() || b.matches.front().colorIndex != a.colorIndex) fail("row color", describeHit("", i, a.displayLineStart));
        std::vector<std::tuple<int, int, int>> oldShown;
        for (size_t k = 0; k < a.matchStarts.size(); ++k) oldShown.emplace_back(a.matchStarts[k], a.matchLens[k], a.matchColors[k]);
        if (shown != oldShown) fail("shown matches", describeHit("", i, a.displayLineStart));
    }
    ++g_checks;
    if (docks.oldDock._lineStartToHitIndex != docks.newDock._lineStartToHitIndex) fail("line index", "differs");
}

// Text, fold levels and the color of every byte
static void compareDock(Docks& docks) {
    auto& od = docks.oldEmu.doc();
    auto& nd = docks.newEmu.doc();
    ++g_checks;
    const Sci::Position len = od.Length();
    if (len != nd.Length()) { fail("dock length", std::to_string(len) + " vs " + std::to_string(nd.Length())); return; }
    std::string a(static_cast<size_t>(len), '\0'), b(static_cast<size_t>(len), '\0');
    od.GetCharRange(a.data(), 0, len);
    nd.GetCharRange(b.data(), 0, len);
    if (a != b) { fail("dock text", "differs"); return; }
    for (Sci::Line l = 0; l < od.LinesTotal(); ++l)
        if (od.GetFoldLevel(l) != nd.GetFoldLevel(l)) { fail("fold level", "line " + std::to_string(l)); break; }
    using O = olddock::ResultDock;
    using N = newdock::ResultDock;
    long long bad = 0;
    Sci::Position firstBad = -1;
    for (Sci::Position p = 0; p < len; ++p) {
        int oldFore = static_cast<unsigned char>(od.StyleAt(p));
        if (od.decorations->ValueAt(O::INDIC_LINENUMBER_FORE, p)) oldFore = 1000;
        else if (od.decorations->ValueAt(O::INDIC_MATCH_FORE, p)) oldFore = 1001;
        int newFore = static_cast<unsigned char>(nd.StyleAt(p));
        if (newFore == N::STYLE_HIT_NUMBER) newFore = 1000;
        else if (newFore == N::STYLE_HIT_MATCH) newFore = 1001;
        uint32_t oldBack = 0, newBack = 0;
        for (int i = 0; i < O::MAX_ENTRY_COLORS; ++i) {
            if (od.decorations->ValueAt(O::INDIC_ENTRY_BG_BASE + i, p)) oldBack |= 1u << i;
            if (nd.decorations->ValueAt(N::INDIC_ENTRY_BG_BASE + i, p)) newBack |= 1u << i;
        }
        if (oldFore != newFore || oldBack != newBack) { if (firstBad < 0) firstBad = p; ++bad; }
    }
    ++g_checks;
    if (bad) fail("colors", std::to_string(bad) + " bytes, first at " + std::to_string(firstBad));
}

// FlowTabs padding added and removed in a file: stored positions follow in both models
static void compareFlowTabs(Docks& docks, const std::string& path, std::mt19937_64& rng) {
    std::vector<std::pair<Sci_Position, Sci_Position>> pads;
    Sci_Position at = 0;
    for (int k = 0; k < 6; ++k) { at += static_cast<Sci_Position>(rng() % 400); const Sci_Position n = 1 + static_cast<Sci_Position>(rng() % 8); pads.push_back({ at, at + n }); at += n; }
    const bool added = (rng() & 1) != 0;
    docks.oldDock.adjustHitPositionsForFlowTab(path, pads, added);
    docks.newDock.adjustHitPositionsForFlowTab(path, pads, added);
    compareRows(docks);
}

// Line breaks of every dock line by Scintilla's LineLayout::WrapLine, one monospace cell per
// character, at four widths. Also counts the lines SC_WRAP_WORD would break apart from a43e162.
static void compareWrap(Docks& docks, Scintilla::Wrap newMode) {
    using namespace Scintilla::Internal;
    const Document& od = docks.oldEmu.doc();
    const Document& nd = docks.newEmu.doc();
    Sci::Position longest = 0;
    for (Sci::Line l = 0; l < od.LinesTotal(); ++l) longest = std::max(longest, od.LineEnd(l) - od.LineStart(l));
    LineLayout lo(-1, static_cast<int>(longest) + 1), ln(-1, static_cast<int>(longest) + 1);
    auto layout = [](LineLayout& ll, const Document& d, Sci::Position start, int n) {
        ll.numCharsInLine = ll.numCharsBeforeEOL = n;
        ll.positions[0] = 0;
        int cells = 0;
        for (int i = 0; i < n; ++i) {
            ll.chars[i] = d.CharAt(start + i);
            ll.styles[i] = static_cast<unsigned char>(d.StyleAt(start + i));
            if ((static_cast<unsigned char>(ll.chars[i]) & 0xC0) != 0x80) ++cells;
            ll.positions[i + 1] = 8.0 * cells;
        }
    };
    auto breaks = [](LineLayout& ll, const Document& d, Sci::Position start, Scintilla::Wrap mode, int width) {
        ll.WrapLine(&d, start, mode, 8.0 * width);
        std::vector<int> v;
        for (int s = 0; s < ll.lines; ++s) v.push_back(ll.LineStart(s));
        return v;
    };
    for (Sci::Line l = 0; l < od.LinesTotal(); ++l) {
        const Sci::Position s = od.LineStart(l);
        const int n = static_cast<int>(od.LineEnd(l) - s);
        if (n == 0) continue;
        layout(lo, od, s, n);
        layout(ln, nd, s, n);
        ++g_wrapLines;
        bool wordDiffers = false;
        for (int width : { 16, 40, 72, 120 }) {
            ++g_checks;
            const std::vector<int> a = breaks(lo, od, s, Scintilla::Wrap::Word, width);
            if (breaks(ln, nd, s, newMode, width) != a) fail("wrap", "line " + std::to_string(l) + ", width " + std::to_string(width));
            if (breaks(ln, nd, s, Scintilla::Wrap::Word, width) != a) wordDiffers = true;
        }
        if (wordDiffers) ++g_wrapWordDiffers;
    }
}

// The dock's Delete key on a random line, block or selection, the same in both docks
static void deleteRandomly(Docks& docks, std::mt19937_64& rng) {
    const int deletions = static_cast<int>(rng() % 6);
    for (int d = 0; d < deletions; ++d) {
        const Sci::Position len = docks.oldEmu.doc().Length();
        if (len <= 0) break;
        const Sci::Position a = static_cast<Sci::Position>(rng() % static_cast<uint64_t>(len));
        const Sci::Position c = (rng() % 2) ? a : static_cast<Sci::Position>(rng() % static_cast<uint64_t>(len));
        docks.oldEmu.selAnchor = docks.newEmu.selAnchor = a;
        docks.oldEmu.selCaret = docks.newEmu.selCaret = c;
        if (rng() % 3 == 0) olddock::ResultDock::_perEntryColorsEnabled = newdock::ResultDock::_perEntryColorsEnabled = !newdock::ResultDock::_perEntryColorsEnabled;
        olddock::ResultDock::s_instance = &docks.oldDock;
        g_dockWindow = &docks.oldEmu;
        olddock::ResultDock::deleteSelectedItems(docks.oldDock._hSci);
        newdock::ResultDock::s_instance = &docks.newDock;
        g_dockWindow = &docks.newEmu;
        newdock::ResultDock::deleteSelectedItems(docks.newDock._hSci);
        g_dockWindow = nullptr;
        ++g_deletions;
        compareRows(docks);
        compareDock(docks);
    }
}

// ---------------------------------------------------------------------------------------------
// Inputs
// ---------------------------------------------------------------------------------------------
static std::string generatedText(std::mt19937_64& rng, bool ansi) {
    static const char* words[] = { "in", "int", "include", "Find", "find", "Fi", "file", "Straße", "über", "Übung", "naïve",
        "\xC2\xA0", "\xC2\xAD", "\x01", "\x7F", "\t", "  ", "x", "_", "1024", "Grüße", "im", "IN", "Inhalt", "sin", "pin" };
    std::string t;
    const int lines = 1 + static_cast<int>(rng() % 60);
    for (int l = 0; l < lines; ++l) {
        const bool longLine = rng() % 12 == 0;
        const int n = longLine ? 600 : static_cast<int>(rng() % 14);
        for (int k = 0; k < n; ++k) { if (k) t += ' '; t += words[rng() % (sizeof(words) / sizeof(words[0]))]; }
        t += (rng() % 3 == 0) ? "\n" : "\r\n";
    }
    if (!ansi) return t;
    std::string a;   // UTF-8 to code page 1252 (Latin-1 range only in this text)
    for (size_t i = 0; i < t.size(); ++i) {
        const unsigned char c = static_cast<unsigned char>(t[i]);
        if (c == 0xC2 || c == 0xC3) { a += static_cast<char>(c == 0xC2 ? t[i + 1] : (t[i + 1] | 0x40)); ++i; }
        else a += t[i];
    }
    return a;
}

static std::vector<Criterion> makeCriteria(std::mt19937_64& rng, const std::vector<SourceFile>& files, bool listMode) {
    static const char* fixed[] = { "in", "Fi", "Straße", "IN", "\xC3\xBC" "ber", "zqxjvqk", "\t", "e" };
    static const char* regexes[] = { "\\w+ing", "^\\s+\\w", "[A-Z]{2,}", "i.", "\\bF\\w*", "e$", "n\\r?\\n.", "." };
    std::vector<Criterion> crits;
    const int count = listMode ? 1 + static_cast<int>(rng() % 4) : 1;
    for (int i = 0; i < count; ++i) {
        Criterion c;
        std::string u8;
        const int kind = static_cast<int>(rng() % 10);
        if (kind < 2) { u8 = regexes[rng() % (sizeof(regexes) / sizeof(regexes[0]))]; c.flags = SCFIND_REGEXP | SCFIND_POSIX | SCFIND_REGEXP_SKIPCRLFASONE | SCFIND_REGEXP_EMPTYMATCH_ALL; }
        else if (kind < 5) u8 = fixed[rng() % (sizeof(fixed) / sizeof(fixed[0]))];
        else {
            const SourceFile& f = files[rng() % files.size()];
            if (f.text.empty()) u8 = "in";
            else {
                size_t a = rng() % f.text.size();
                while (a > 0 && (static_cast<unsigned char>(f.text[a]) & 0xC0) == 0x80) --a;
                size_t b = std::min(f.text.size(), a + 1 + rng() % 6);
                while (b < f.text.size() && (static_cast<unsigned char>(f.text[b]) & 0xC0) == 0x80) ++b;
                u8 = f.text.substr(a, b - a);
                if (f.codepage != SC_CP_UTF8) { std::string w; for (unsigned char ch : u8) { if (ch < 0x80) w += static_cast<char>(ch); } u8 = w.empty() ? "in" : w; }
            }
        }
        if (rng() % 2) c.flags |= SCFIND_MATCHCASE;
        if (!(c.flags & SCFIND_REGEXP) && rng() % 4 == 0) c.flags |= SCFIND_WHOLEWORD;
        c.findW = widen(u8);
        c.bytes = u8;
        c.color = listMode ? i : 0;
        crits.push_back(std::move(c));
    }
    return crits;
}

static std::vector<SourceFile> loadFolder(const std::vector<std::string>& dirs) {
    std::vector<SourceFile> files;
    for (const std::string& d : dirs) {
        std::error_code ec;
        for (auto it = std::filesystem::recursive_directory_iterator(d, ec); !ec && it != std::filesystem::recursive_directory_iterator(); it.increment(ec)) {
            if (!it->is_regular_file(ec)) continue;
            SourceFile f;
            f.text = readText(it->path());
            if (f.text.find('\0') != std::string::npos || f.text.size() > 400000) continue;
            f.pathUtf8 = it->path().string();
            f.pathW = widen(f.pathUtf8);
            files.push_back(std::move(f));
        }
    }
    std::sort(files.begin(), files.end(), [](const SourceFile& a, const SourceFile& b) { return a.pathUtf8 < b.pathUtf8; });
    return files;
}

// One line of 32 to 128 KB in code page 1252 with a hit every 32 bytes, after the cap mostly unshown
static void longLines() {
    for (const size_t kb : { 32, 64, 128 }) {
        SourceFile f;
        f.codepage = 0;
        f.pathUtf8 = "C:\\gen\\long" + std::to_string(kb) + ".txt";
        f.pathW = widen(f.pathUtf8);
        while (f.text.size() < kb * 1024) f.text += "Fi \xE4\xF6\xFC\xDF \xAD\xA0 abcdefghijklmnopqrst ";
        f.text += "\r\nFi\r\n";
        size_t hits = 0;
        for (size_t p = f.text.find("Fi"); p != std::string::npos; p = f.text.find("Fi", p + 2)) ++hits;
        Criterion c;
        c.findW = L"Fi";
        c.bytes = "Fi";
        c.flags = SCFIND_MATCHCASE;
        std::vector<SourceFile> files{ f };
        Docks docks;
        g_oldAppendSeconds = g_newAppendSeconds = 0;
        searchBlock(docks, files, { c }, false, false);
        compareRows(docks);
        compareDock(docks);
        if (kb == 128) {
            ++g_checks;
            if (g_newAppendSeconds * 20 > g_oldAppendSeconds) fail("long line", "not 20 times as fast as a43e162");
        }
        std::printf("L. %3zu KB line, %zu hits: a43e162 %.3f s, now %.3f s\n", kb, hits, g_oldAppendSeconds, g_newAppendSeconds);
    }
}

int main(int argc, char** argv) {
    std::printf("Result dock: new hit model against a43e162\n\n");
    const int drift = verbatimCheck();
    std::string wrapName;
    const std::optional<Scintilla::Wrap> wrapMode = wrapModeOfSource(wrapName);
    if (!wrapMode) fail("wrap mode", "../ResultDock.cpp: " + (wrapName.empty() ? std::string("not found") : wrapName));
    std::vector<SourceFile> corpus = loadFolder(std::vector<std::string>(argv + 1, argv + argc));
    std::mt19937_64 rng(20261010), deleteRng(20261011);
    int blocks = 0;
    long long rows = 0;
    for (int session = 0; session < 300; ++session) {
        Docks docks;
        const int searches = 1 + static_cast<int>(rng() % 4);
        for (int s = 0; s < searches; ++s) {
            std::vector<SourceFile> files;
            const int nFiles = 1 + static_cast<int>(rng() % 3);
            for (int k = 0; k < nFiles; ++k) {
                if (!corpus.empty() && rng() % 3) files.push_back(corpus[rng() % corpus.size()]);
                else {
                    SourceFile f;
                    const bool ansi = rng() % 4 == 0;
                    f.text = generatedText(rng, ansi);
                    f.codepage = ansi ? 0 : SC_CP_UTF8;
                    f.pathUtf8 = "C:\\gen\\file" + std::to_string(rng() % 1000) + ".txt";
                    f.pathW = widen(f.pathUtf8);
                    files.push_back(std::move(f));
                }
            }
            ColumnTabs::g_indicatorId = (rng() % 4 == 0) ? 8 : -1;
            for (SourceFile& f : files) {
                f.padding.clear();
                if (ColumnTabs::g_indicatorId >= 0 && !f.text.empty())
                    for (int k = 0; k < 20; ++k) {
                        const Sci::Position p = static_cast<Sci::Position>(rng() % f.text.size());
                        f.padding.push_back({ p, std::min<Sci::Position>(1 + static_cast<Sci::Position>(rng() % 4), static_cast<Sci::Position>(f.text.size()) - p) });
                    }
            }
            const bool listMode = rng() % 2;
            const bool groupView = rng() % 2;
            olddock::ResultDock::_perEntryColorsEnabled = newdock::ResultDock::_perEntryColorsEnabled = (rng() % 3 == 0);
            searchBlock(docks, files, makeCriteria(rng, files, listMode), listMode, groupView);
            ++blocks;
            compareRows(docks);
            compareDock(docks);
            if (!files.empty()) compareFlowTabs(docks, files.front().pathUtf8, rng);
        }
        olddock::ResultDock::_perEntryColorsEnabled = newdock::ResultDock::_perEntryColorsEnabled = (rng() % 2);
        docks.oldDock.applyStyling();
        docks.newDock.applyStyling();
        compareDock(docks);
        rows += static_cast<long long>(docks.newDock._hits.size());
        if (wrapMode) compareWrap(docks, *wrapMode);
        deleteRandomly(docks, deleteRng);
    }
    std::printf("S. %d search blocks in 300 docks, %lld dock lines\n", blocks, rows);
    longLines();
    std::printf("W. wrap of %lld dock lines at 4 widths, a43e162 with SC_WRAP_WORD, now with %s; SC_WRAP_WORD would break %lld of them differently\n",
        g_wrapLines, wrapName.c_str(), g_wrapWordDiffers);
    std::printf("D. Delete key: %lld deletions of lines, blocks and selections\n", g_deletions);
    std::printf("C. %lld checks, %lld failed\n", g_checks, g_failures);
    const bool ok = g_failures == 0 && drift == 0;
    std::printf("\n%s\n", ok ? "PASS" : "FAIL");
    return ok ? 0 : 1;
}
