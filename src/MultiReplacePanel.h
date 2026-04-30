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

#ifndef MULTI_REPLACE_H
#define MULTI_REPLACE_H

// Project headers
#include "ColumnTabs.h"
#include "ConfigManager.h"
#include "DPIManager.h"
#include "DropTarget.h"
#include "Encoding.h"
#include "LanguageManager.h"
#include "MultiReplaceConfigDialog.h"
#include "NppStyleKit.h"
#include "PluginInterface.h"
#include "ResultDock.h"
#include "SciUndoGuard.h"
#include "StaticDialog/resource.h"

// Standard library
#include <array>
#include <filesystem>
#include <functional>
#include <map>
#include <memory>
#include <regex>
#include <set>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

// Third-party

// Formula engine (Lua / ExprTk dispatcher)
#include "engine/IFormulaEngine.h"
#include "engine/EngineFactory.h"
#include "engine/ILuaEngineHost.h"

// Windows
#include <commctrl.h>

extern NppData nppData;

enum class DelimiterOperation { LoadAll, Update };
enum class Direction { Up, Down };

struct ReplaceItemData
{
    size_t id = 0;
    int findCount = -1;
    int replaceCount = -1;
    bool isEnabled = true;
    std::wstring findText;
    std::wstring replaceText;
    bool wholeWord = false;
    bool matchCase = false;
    bool useVariables = false;
    bool extended = false;
    bool regex = false;
    std::wstring comments = L"";
    bool isDirty = false;           // Session-only: marks row as modified since last save/load
    std::wstring lastModified;      // Persistent timestamp: set on content change, saved in CSV

    bool operator==(const ReplaceItemData& rhs) const {
        return
            isEnabled == rhs.isEnabled &&
            findText == rhs.findText &&
            replaceText == rhs.replaceText &&
            wholeWord == rhs.wholeWord &&
            matchCase == rhs.matchCase &&
            extended == rhs.extended &&
            regex == rhs.regex;
    }

    bool operator!=(const ReplaceItemData& rhs) const {
        return !(*this == rhs);
    }
};

// Hash function for ReplaceItemData
struct ReplaceItemDataHasher {
    std::size_t operator()(const ReplaceItemData& item) const {
        std::size_t hash = std::hash<bool>{}(item.isEnabled);
        hash ^= std::hash<std::wstring>{}(item.findText) << 1;
        hash ^= std::hash<std::wstring>{}(item.replaceText) << 1;
        hash ^= std::hash<bool>{}(item.wholeWord) << 1;
        hash ^= std::hash<bool>{}(item.matchCase) << 1;
        hash ^= std::hash<bool>{}(item.useVariables) << 1;
        hash ^= std::hash<bool>{}(item.extended) << 1;
        hash ^= std::hash<bool>{}(item.regex) << 1;
        hash ^= std::hash<std::wstring>{}(item.comments) << 1;
        return hash;
    }
};

struct UndoRedoAction {
    std::function<void()> undoAction;
    std::function<void()> redoAction;
};



enum class FontRole {
    Standard = 0,
    Normal1, // 14px MS Shell Dlg 2
    Normal2, // 12px Courier New
    Normal3, // 14px Courier New
    Normal4, // 16px Courier New
    Normal5, // 18px Courier New
    Normal6, // 22px Courier New
    Normal7, // 26px Courier New
    Bold1,   // 22px Courier New Bold
    Bold2,   // 12px MS Shell Dlg 2 Bold
    Count    // Keep this last to determine array size
};

struct ControlInfo
{
    int x = 0;
    int y = 0;
    int cx = 0;
    int cy = 0;
    LPCWSTR className = nullptr;
    LPCWSTR windowName = nullptr;
    DWORD style = 0;
    LPCWSTR tooltipText = nullptr;
    bool isStatic = false;
    FontRole fontRole = FontRole::Standard;
};

struct SearchContext {
    std::string findText = "";      // search string in Scintilla encoding
    int searchFlags = 0;            // Search flags (e.g., SCFIND_MATCHCASE, SCFIND_WHOLEWORD, etc.)
    LRESULT docLength = 0;          // Cached document length
    int cachedCodepage = -1;        // Cached document codepage (-1 = not set, use SCI_GETCODEPAGE)
    bool isColumnMode = false;      // Cached state: true if Column Mode is active
    bool isSelectionMode = false;   // Cached state: true if Selection Mode is active
    bool retrieveFoundText = false; // If true, retrieve the found text from Scintilla
    bool highlightMatch = false;    // If true, highlight the found match
    bool useStoredSelections = false; // If true, use m_selectionScope instead of current selection
};

struct SearchResult {
    LRESULT pos = -1;
    LRESULT length = 0;
    std::string foundText = "";
};

struct SelectionInfo {
    Sci_Position startPos;
    Sci_Position endPos;
    Sci_Position length;
};

struct SelectionRange {
    LRESULT start = 0;
    LRESULT end = 0;
};

struct ColumnDelimiterData {
    std::vector<int> inputColumns; // original order of the columns
    std::set<int> columns;
    std::string extendedDelimiter;
    std::string quoteChar;
    SIZE_T delimiterLength = 0;
    bool delimiterChanged = false;
    bool quoteCharChanged = false;
    bool columnChanged = false;

    bool isValid() const {
        bool isQuoteCharValid = quoteChar.empty() ||
            (quoteChar.length() == 1 && (quoteChar[0] == '"' || quoteChar[0] == '\''));
        return !columns.empty() && !extendedDelimiter.empty() && isQuoteCharValid;
    }
};

struct ColumnValue {
    bool        isNumeric = false;
    double      numericValue = 0.0;
    std::string text;
    std::wstring textW;  // Cached wide string for fast comparison
};

struct CombinedColumns {
    std::vector<ColumnValue> columns;
};

struct DelimiterPosition {
    LRESULT offsetInLine = 0;  // where the delimiter is within this line
};

struct LineInfo {
    std::vector<DelimiterPosition> positions;
    LRESULT lineLength = 0; // how many chars total in this line
};

struct ColumnInfo {
    LRESULT totalLines;
    LRESULT startLine;
    SIZE_T startColumnIndex;
};



struct HighlightedTabs {
    std::unordered_set<int> bufferIDs;  // Stores buffer IDs with active highlighting

    bool isHighlighted(int bufferID) const {
        return bufferIDs.find(bufferID) != bufferIDs.end();
    }

    void mark(int bufferID) {
        bufferIDs.insert(bufferID);
    }

    void clear(int bufferID) {
        bufferIDs.erase(bufferID);
    }
};

struct MenuState {
    bool listNotEmpty = false;
    bool canEdit = false;
    bool canPaste = false;
    bool hasSelection = false;
    bool clickedOnItem = false;
    bool allEnabled = false;
    bool allDisabled = false;
    bool canUndo = false;
    bool canRedo = false;
};

struct MonitorEnumData {
    std::wstring monitorInfo;
    int monitorCount;
    int currentMonitor;
    int primaryMonitorIndex;
};

enum class ItemAction {
    Search,
    Edit,
    Paste,
    Copy,
    Cut,
    Delete,
    Undo,
    Redo,
    Add
};

enum class SearchOption {
    WholeWord,
    MatchCase,
    Variables,
    Extended,
    Regex
};

enum class SortDirection {
    Unsorted,
    Ascending,
    Descending
};

enum ColumnID {
    INVALID = -1,
    FIND_COUNT,         // 0
    REPLACE_COUNT,      // 1
    SELECTION,          // 2
    FIND_TEXT,          // 3
    REPLACE_TEXT,       // 4
    WHOLE_WORD,         // 5
    MATCH_CASE,         // 6
    USE_VARIABLES,      // 7
    EXTENDED,           // 8
    REGEX,              // 9
    COMMENTS,           // 10
    LAST_MODIFIED,      // 11
    DELETE_BUTTON       // 12
};

// Per-tab snapshot of all state that must survive a tab switch.
// Panel-level settings (scope, search options, column layout, input
// field contents) live here. Settings-dialog options and UI
// preferences stay global on MultiReplace.
struct TabState {
    // Identity
    int          id = 0;
    std::wstring name;
    std::wstring filePath;            // real user-chosen path; empty for untitled
    std::wstring snapshotPath;        // plugin-internal cache path

    // List data
    std::vector<ReplaceItemData> data;
    std::size_t                  originalHash = 0;

    // Dirty indicator (runtime-only, drives the tab's dirty marker).
    // Set on any edit, cleared only by save/load - a manual revert
    // does not flip it back.
    bool         isDirty = false;

    // Search options
    int          searchMode = 0;     // 0=Normal, 1=Extended, 2=Regex
    bool         wholeWord = false;
    bool         matchCase = false;
    bool         useVariables = false;
    bool         wrapAround = false;
    bool         replaceAtMatches = false;
    std::wstring replaceAtMatchesEdit = L"1";

    // Formula engine choice for this tab. Default Lua for backward
    // compatibility. Persisted as a string ("Lua" / "ExprTk").
    MultiReplaceEngine::EngineType engine = MultiReplaceEngine::EngineType::Lua;

    // Owned engine instance, created lazily on first execute() call.
    // Not persisted - recreated from `engine` after a workspace load.
    MultiReplaceEngine::FormulaEnginePtr engineInstance;

    // Scope
    int          scope = 0;     // 0=AllText, 1=Selection, 2=CSV
    std::wstring csvCols = L"1-50";
    std::wstring csvDelim = L",";
    std::wstring csvQuote = L"\"";

    // Input fields
    std::wstring findText;
    std::wstring replaceText;

    // Files / OpenDocs panel
    std::wstring filesFilter = L"*.*";
    std::wstring docsFilter = L"*.*";
    std::wstring directory;
    bool         inSubfolders = false;
    bool         inHiddenFolders = false;
    bool         allDocuments = false;

    // Column layout
    int          findCountWidth = 50;
    int          replaceCountWidth = 50;
    int          findWidth = 50;
    int          replaceWidth = 50;
    int          commentsWidth = 50;
    bool         findCountVisible = false;
    bool         replaceCountVisible = false;
    bool         commentsVisible = false;
    bool         timestampVisible = false;
    bool         deleteButtonVisible = true;
    bool         findLocked = true;
    bool         replaceLocked = false;
    bool         commentsLocked = true;
    // Lazy-relayout flag. Set when the window was resized or column
    // visibility changed while this tab was inactive. Consumed by the
    // next switchToTab: if true, the tab is redistributed to fit the
    // current window; if false, the stored widths are applied as-is.
    // Not persisted — purely runtime state.
    bool         needsRelayout = false;
    std::vector<ColumnID>         columnOrder;
    std::map<int, SortDirection>  columnSortOrder;

    // Default-constructible (all members have member-initializers).
    TabState() = default;

    // Custom copy constructor: copies all data members but NOT the
    // engineInstance. Required because std::unique_ptr<IFormulaEngine>
    // is non-copyable, but we still want to allow tab duplication.
    // The duplicated tab keeps the same engine *type* (Lua / ExprTk)
    // and re-creates its engine lazily on the first execute() call.
    TabState(const TabState& other)
        : id(other.id)
        , name(other.name)
        , filePath(other.filePath)
        , snapshotPath(other.snapshotPath)
        , data(other.data)
        , originalHash(other.originalHash)
        , isDirty(other.isDirty)
        , searchMode(other.searchMode)
        , wholeWord(other.wholeWord)
        , matchCase(other.matchCase)
        , useVariables(other.useVariables)
        , wrapAround(other.wrapAround)
        , replaceAtMatches(other.replaceAtMatches)
        , replaceAtMatchesEdit(other.replaceAtMatchesEdit)
        , engine(other.engine)
        // engineInstance intentionally NOT copied; left default (null).
        , scope(other.scope)
        , csvCols(other.csvCols)
        , csvDelim(other.csvDelim)
        , csvQuote(other.csvQuote)
        , findText(other.findText)
        , replaceText(other.replaceText)
        , filesFilter(other.filesFilter)
        , docsFilter(other.docsFilter)
        , directory(other.directory)
        , inSubfolders(other.inSubfolders)
        , inHiddenFolders(other.inHiddenFolders)
        , allDocuments(other.allDocuments)
        , findCountWidth(other.findCountWidth)
        , replaceCountWidth(other.replaceCountWidth)
        , findWidth(other.findWidth)
        , replaceWidth(other.replaceWidth)
        , commentsWidth(other.commentsWidth)
        , findCountVisible(other.findCountVisible)
        , replaceCountVisible(other.replaceCountVisible)
        , commentsVisible(other.commentsVisible)
        , timestampVisible(other.timestampVisible)
        , deleteButtonVisible(other.deleteButtonVisible)
        , findLocked(other.findLocked)
        , replaceLocked(other.replaceLocked)
        , commentsLocked(other.commentsLocked)
        , needsRelayout(other.needsRelayout)
        , columnOrder(other.columnOrder)
        , columnSortOrder(other.columnSortOrder)
    {
    }

    // Copy-assignment, move-construct, move-assign: not currently used,
    // and explicit definitions would have to mirror the field list. Keep
    // them deleted/defaulted via the rule of five so a future addition
    // doesn't silently leak engineInstance copies through assignment.
    TabState& operator=(const TabState&) = delete;
    TabState(TabState&&) = default;
    TabState& operator=(TabState&&) = default;
};

enum class SearchDirection {
    Forward,
    Backward
};

struct ResizableColWidths {
    HWND listView;
    int listViewWidth;
    int findCountWidth;
    int replaceCountWidth;
    int findWidth;
    int replaceWidth;
    int commentsWidth;
    int timestampWidth;
    int deleteWidth;
    int margin;
};

// Column-width strategy for createListViewColumns().
//   Redistribute  — even split across unlocked columns to fit the
//                   current window (resize, visibility, init, …).
//   UseStored     — apply the stored per-column widths unchanged,
//                   used on tab switches when the target tab's
//                   layout still matches the current window size.
enum class WidthMode {
    Redistribute,
    UseStored
};

struct ViewState {
    int firstVisibleLine = 0;
    int xOffset = 0;
    Sci_Position caret = 0;
    Sci_Position anchor = 0;
    int wrapMode = 0; // not strictly required for restore
};

struct EncodingInfo {
    int sc_codepage = 0;      // The codepage value for Scintilla (e.g., SC_CP_UTF8)
    size_t bom_length = 0;    // The length of the BOM in bytes (0 if no BOM)
};

class CsvLoadException : public std::exception {
public:
    explicit CsvLoadException(const std::string& message) : message_(message) {}
    const char* what() const noexcept override {
        return message_.c_str();
    }
private:
    std::string message_;
};

struct EditControlContext
{
    MultiReplace* pThis;
    HWND hwndExpandBtn;
};

class MultiReplace : public StaticDialog, public MultiReplaceEngine::ILuaEngineHost
{
public:

    // Tandem-mode dock edge. Public so free helpers outside the class
    // can bridge to tandem_dock::DockEdge without friend declarations.
    enum class TandemDockEdge { Bottom, Right, Left };

    /// RAII guard for Scintilla undo grouping.
    /// Uses shared nesting counter from SciUndoGuard.h to ensure that
    /// nested operations (e.g., ColumnTabs functions called from here)
    /// all become part of ONE undo step.
    class ScopedUndoAction {
        MultiReplace& _owner;
        bool _ownsAction;  // true if this instance sent SCI_BEGINUNDOACTION
    public:
        explicit ScopedUndoAction(MultiReplace& owner)
            : _owner(owner), _ownsAction(false)
        {
            // Only the outermost guard (nesting depth 0 → 1) sends BEGIN
            if (SciUndo::detail::g_nestingDepth == 0) {
                if (_owner.send(SCI_GETUNDOCOLLECTION, 0, 0)) {
                    _owner.send(SCI_BEGINUNDOACTION, 0, 0);
                    _ownsAction = true;
                }
            }
            ++SciUndo::detail::g_nestingDepth;
        }

        ~ScopedUndoAction() {
            if (SciUndo::detail::g_nestingDepth > 0) {
                --SciUndo::detail::g_nestingDepth;
            }
            if (_ownsAction) {
                _owner.send(SCI_ENDUNDOACTION, 0, 0);
            }
        }

        ScopedUndoAction(const ScopedUndoAction&) = delete;
        ScopedUndoAction& operator=(const ScopedUndoAction&) = delete;
    };

    class ScopedRedrawLock {
        HWND _hwnd;
    public:
        explicit ScopedRedrawLock(HWND hwnd) : _hwnd(hwnd) {
            if (_hwnd) {
                ::SendMessage(_hwnd, WM_SETREDRAW, FALSE, 0);
            }
        }
        ~ScopedRedrawLock() {
            if (_hwnd) {
                ::SendMessage(_hwnd, WM_SETREDRAW, TRUE, 0);
                ::RedrawWindow(_hwnd, NULL, NULL, RDW_ERASE | RDW_FRAME | RDW_INVALIDATE | RDW_ALLCHILDREN);
            }
        }
        ScopedRedrawLock(const ScopedRedrawLock&) = delete;
        ScopedRedrawLock& operator=(const ScopedRedrawLock&) = delete;
    };

    MultiReplace() :
        hInstance(nullptr),
        _hScintilla(nullptr),
        _replaceListView(nullptr),
        _statusMessageColor(RGB(0, 0, 0))
    {
        setInstance(this);
    };

    inline static MultiReplace* instance = nullptr; // Static instance of the class
    static HHOOK _hMsgFilterHook;                   // Thread-local hook for Alt+Up/Down
    static LRESULT CALLBACK MsgFilterHookProc(int nCode, WPARAM wParam, LPARAM lParam);

    // Helper functions for scaling
    inline int sx(int value) { return dpiMgr->scaleX(value); }
    inline int sy(int value) { return dpiMgr->scaleY(value); }

    static inline void setInstance(MultiReplace* inst) {
        instance = inst;
    }

    virtual inline void display(bool toShow = true) const override {
        StaticDialog::display(toShow);
    };

    inline void setParent(HWND parent2set) {
        _hParent = parent2set;
        _hParent = parent2set;
    };

    static inline HWND getScintillaHandle() {
        return s_hScintilla;
    }

    static inline HWND getDialogHandle() {
        return s_hDlg;
    }

    struct DMDARKMODECOLORS {
        COLORREF bkgColor;
        COLORREF foreColor;
        COLORREF hotColor;
        COLORREF darkColor;
    };

    enum class MessageStatus {
        Info,
        Success,
        Error
    };

    struct Settings
    {
        bool tooltipsEnabled;
        bool exportToBashEnabled;
        bool muteSounds;
        bool doubleClickEditsEnabled;
        bool highlightMatchEnabled;
        bool flowTabsIntroDontShowEnabled;
        bool flowTabsNumericAlignEnabled;
        bool isHoverTextEnabled;
        bool listStatisticsEnabled;
        bool stayAfterReplaceEnabled;
        bool groupResultsEnabled;
        bool allFromCursorEnabled;
        bool keepListVisible;
        int  listDimIntensity;
        int  tabMaxLength;
        bool limitFileSizeEnabled;
        int  maxFileSizeMB;
        int  editFieldSize;
        int  csvHeaderLinesCount;
        bool resultDockPerEntryColorsEnabled;
        bool useListColorsForMarking;
        bool duplicateBookmarksEnabled;
    };

    static Settings getSettings();
    static void writeStructToConfig(const Settings& settings);
    // Getters for ConfigDialog
    bool isUseListEnabled() const;
    bool isTwoButtonsModeEnabled() const;

    void loadUIConfigFromIni();
    void loadSettingsToPanelUI();
    void syncUIToCache();
    void applyConfigSettingsOnly();
    static  std::pair<std::wstring, std::wstring> generateConfigFilePaths();

    // Snapshot directory holds per-tab CSV shutdown caches (one file per tab).
    // User-chosen CSVs stay at the user's own paths and are never moved here.
    static std::wstring getSnapshotsDir();
    static bool         snapshotsDirExists();
    static bool         ensureSnapshotsDir();  // returns false only on filesystem error
    static void         migrateSnapshotsDir(); // one-shot move from legacy nested path

    // Light Mode Colors for Message
    static constexpr COLORREF LMODE_SUCCESS = RGB(0, 128, 0);
    static constexpr COLORREF LMODE_ERROR = RGB(200, 0, 0);
    static constexpr COLORREF LMODE_INFO = RGB(0, 0, 128);
    static constexpr COLORREF LMODE_FILTER_HELP = RGB(0, 0, 255);
    static constexpr COLORREF LMODE_ENGINE_LINK = RGB(0, 0, 255);

    // Dark Mode Colors for Message
    static constexpr COLORREF DMODE_SUCCESS = RGB(120, 220, 120);
    static constexpr COLORREF DMODE_ERROR = RGB(255, 110, 110);
    static constexpr COLORREF DMODE_INFO = RGB(180, 180, 255);
    static constexpr COLORREF DMODE_FILTER_HELP = RGB(255, 235, 59);
    static constexpr COLORREF DMODE_ENGINE_LINK = RGB(100, 180, 255);

    inline static bool isWindowOpen = false;
    inline static bool textModified = true;
    inline static bool documentSwitched = false;
    inline static int  scannedDelimiterBufferID = -1;
    inline static bool _delimiterPositionsStale = false;
    inline static bool isLoggingEnabled = true;
    inline static bool isCaretPositionEnabled = false;
    inline static bool _isLuaErrorDialogEnabled = true;

    inline static std::vector<size_t> originalLineOrder{}; // Stores the order of lines before sorting
    inline static std::vector<size_t> _markedDuplicateLines{};  // Stores line indices of marked duplicates
    inline static size_t _duplicateGroupCount = 0;              // Number of unique duplicate groups
    inline static bool _duplicateMatchCase = false;             // Match case setting used when marking
    inline static bool _duplicateBookmarksEnabled = false;      // Tracks bookmark setting for rescan detection

    // Stored scan criteria for validation rescan before delete
    inline static std::set<int> _duplicateScanColumns{};        // Columns used for scan
    inline static std::string _duplicateScanDelimiter{};        // Delimiter used for scan

    inline static SortDirection currentSortState = SortDirection::Unsorted; // Status of column sort
    inline static bool isSortedColumn = false; // Indicates if a column is sorted

    // Static methods for Event Handling
    static void onSelectionChanged();
    static void onTextChanged();
    static void onDocumentSwitched();
    static void pointerToScintilla();
    static void processLog();
    static void processTextChange(SCNotification* notifyCode);
    static void onCaretPositionChanged();
    static void onThemeChanged();
    static void signalShutdown();
    static bool shouldReopenPanelOnStartup();
    static bool isReopenOnStartupEnabled();
    static void toggleReopenOnStartup();
    static void loadConfigOnce();
    static void migrateLegacyStartupKeys();

    enum class ChangeType { Insert, Delete, Modify };
    enum class ReplaceMode { Normal, Extended, Regex };

    struct LogEntry {
        ChangeType changeType = ChangeType::Modify;
        Sci_Position lineNumber = 0;
        Sci_Position blockSize = 1;
    };

    inline static std::vector<LogEntry> logChanges{};

    // Drag-and-Drop functionality
    DropTarget* dropTarget = nullptr;  // Pointer to DropTarget instance
    DropTarget* tabBarDropTarget = nullptr;  // Drop target on the tab bar
    void loadListFromCsvIntoNewTab(const std::wstring& filePath);  // used by DropTarget for all drops
    void initializeDragAndDrop();

    inline static HWND       hwndExpandBtn = nullptr;
    bool _keepOnTopDuringBatch = false;
    void setBatchUIState(HWND hDlg, bool inProgress);

    static void loadLanguageGlobal();
    static void refreshUILanguage();

    // Tandem mode (experimental): docks MR to a N++ edge and follows it.
    // Toggled from the N++ plugin menu; state is synced via the wrapper.
    void toggleTandemMode();
    bool isTandemEnabled() const { return _tandemEnabled; }
    static bool isTandemPersistedEnabled();  // for menu-init BEFORE the panel exists

protected:
    virtual INT_PTR CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam) override;

private:
    static constexpr int MARKER_COLOR_LIGHT = 0x00A5FF; // Color for non-list Marker (Orange)
    static constexpr int MARKER_COLOR_DARK = 0x0050D2;  // Color for non-list Marker (Orange)
    static constexpr int DUPLICATE_MARKER_COLOR_LIGHT = 0xE6E0B0; // Color for Duplicate Marker (Light Ice Blue)
    static constexpr int DUPLICATE_MARKER_COLOR_DARK = 0x909060;  // Color for Duplicate Marker (Muted Teal)
    static constexpr int REPLACE_FILES_PANEL_HEIGHT = 88;
    bool isReplaceAllInDocs = false; // True if replacing in all open documents, false for current document only.
    bool isReplaceInFiles = false;   // True if replacing in files, false for current document only.
    bool _debugModeEnabled = false;  // Debug Mode checkbox state
    bool isFindAllInDocs = false;
    bool isFindAllInFiles = false;
    bool _listSearchBarVisible = false;

    // Separate filter strings for each panel mode (swapped into IDC_FILTER_EDIT on mode switch)
    std::wstring _filesFilter = L"*.*";
    std::wstring _docsFilter = L"*.*";

    static constexpr int MIN_GENERAL_WIDTH = 40;
    static constexpr int DEFAULT_COLUMN_WIDTH_FIND = 150;   // Default size for Find Column
    static constexpr int DEFAULT_COLUMN_WIDTH_REPLACE = 150; // Default size for Replace Column
    static constexpr int DEFAULT_COLUMN_WIDTH_COMMENTS = 120; // Default size for Comments Column
    static constexpr int DEFAULT_COLUMN_WIDTH_FIND_COUNT = 50; // Default size for Find Count Column
    static constexpr int DEFAULT_COLUMN_WIDTH_REPLACE_COUNT = 50; // Default size for Replace Count Column

    static constexpr BYTE MIN_TRANSPARENCY = 50;  // Minimum visible transparency
    static constexpr BYTE MAX_TRANSPARENCY = 255; // Fully opaque
    static constexpr BYTE DEFAULT_FOREGROUND_TRANSPARENCY = 255; // Default foreground transparency
    static constexpr BYTE DEFAULT_BACKGROUND_TRANSPARENCY = 190; // Default background transparency

    static constexpr int MIN_EDIT_FIELD_SIZE = 2; // Minimum size for Multiline Editor
    static constexpr int MAX_EDIT_FIELD_SIZE = 20; // Maximum size for Multiline Editor

    static constexpr wchar_t* symbolSortAsc = L"▼";
    static constexpr wchar_t* symbolSortDesc = L"▲";
    static constexpr wchar_t* symbolSortAscUnsorted = L"▽";
    static constexpr wchar_t* symbolSortDescUnsorted = L"△";
    static constexpr int MAX_CAP_GROUPS = 9; // Maximum number of capture groups supported by Notepad++

    DPIManager* dpiMgr; // Pointer to DPIManager instance

    // Static variables related to GUI 
    inline static HWND s_hScintilla = nullptr;
    inline static HWND s_hDlg = nullptr;
    HWND hwndEdit = nullptr;
    WNDPROC originalListViewProc;
    inline static std::map<int, ControlInfo> ctrlMap{};

    // Instance-specific GUI-related variables 
    HINSTANCE hInstance;
    HWND _hScintilla;
    HWND _replaceListView;
    std::array<HFONT, static_cast<size_t>(FontRole::Count)> _fontHandles{};
    COLORREF COLOR_SUCCESS;
    COLORREF COLOR_ERROR;
    COLORREF COLOR_INFO;
    COLORREF _statusMessageColor = LMODE_INFO; // Holds the actual color to be drawn. Initialized for light mode.
    COLORREF _filterHelpColor;
    // Engine selector marker colour (theme-aware). Updated by
    // applyThemePalette so the (L)/(E) marker renders in link blue
    // in both light and dark mode.
    COLORREF _engineLinkColor = LMODE_ENGINE_LINK;
    MessageStatus _lastMessageStatus = MessageStatus::Info; // Holds the TYPE of the last message.
    HWND _hHeaderTooltip;        // Handle to the tooltip for the ListView header
    HWND _hUseListButtonTooltip; // Handle to the tooltip for the Use List Button
    // Handle to the tooltip for the (L)/(E) engine selector marker.
    // Stored so syncEngineSelectorLabel can update the text via
    // TTM_UPDATETIPTEXT when the active engine changes.
    HWND _hEngineSelectorTooltip = nullptr;

    // All per-control tooltip windows created in createControls().
    // The handles are tracked here so applyThemePalette() can refresh
    // their dark/light theme when Notepad++ switches modes at runtime
    // - SetWindowTheme is otherwise only applied at tooltip creation
    // and would leave stale tooltips in the wrong theme until restart.
    std::vector<HWND> _ctrlTooltips;

    POINT _contextMenuClickPoint;

    // Style-related variables and constants
    /*
       Available styles (self-tested):
       { 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 }
       Note: Gaps in the list are intentional.

       Styles 0 - 7 are reserved for syntax style.
       Indicator 30 is reserved for MultiReplace ColumnTabs
    */

    const std::vector<int> hColumnStyles = { STYLE1, STYLE2, STYLE3, STYLE4, STYLE5, STYLE6, STYLE7, STYLE8, STYLE9, STYLE10 };

    const std::vector<int> lightModeColumnColors = { 0xFFE0E0, 0xC0E0FF, 0x80FF80, 0xFFE0FF,  0xB0E0E0, 0xFFFF80, 0xE0C0C0, 0x80FFFF, 0xFFB0FF, 0xC0FFC0 };
    const std::vector<int> darkModeColumnColors = { 0x553333, 0x335577, 0x225522, 0x553355, 0x335555, 0x555522, 0x774444, 0x225555, 0x553366, 0x336633 };

    // Alpha values for Marking values
    static constexpr int EDITOR_MARK_ALPHA_LIGHT = 150;
    static constexpr int EDITOR_OUTLINE_ALPHA_LIGHT = 0;
    static constexpr int EDITOR_MARK_ALPHA_DARK = 130;
    static constexpr int EDITOR_OUTLINE_ALPHA_DARK = 0;

    // Preferences (input)
    int preferredColumnTabsStyleId = 30; // preferred ColumnTabs id; -1 = auto


    // Assigned (output)
    // Pools (runtime)
    std::vector<int> textStyles;           // highlight pool (excludes ColumnTabs + standard)

    // Coordinator config
    inline static constexpr int kPreferredIds[] = {
        9,10,11,12,13,14,15,16,17,18,19,20,30,
        32,33,34,35,36,37,38,39,40,41,42,43
    };

    inline static constexpr int kReservedIds[] = {
        0,1,2,3,4,5,6,7,       // lexer
        21,22,23,24,25,26,27,28,29,
        31                     // N++ mark
    };

    // Data-related variables 
    size_t markedStringsCount = 0;
    bool allSelected = true;
    std::unordered_map<int, int> colorToStyleMap;
    std::map<int, SortDirection> columnSortOrder;
    ColumnDelimiterData columnDelimiterData;
    std::vector<ReplaceItemData> replaceListData;
    std::vector<LineInfo> lineDelimiterPositions;
    std::vector<char> lineBuffer; // reusable Buffer for findDelimitersInLine()
    std::vector<char> styleBuffer; // reusable Buffer for highlightColumnsInLine()
    std::vector<char> tagBuffer;  // reusable Buffer for SCI_GETTAG in fillCapturesForEngine()
    bool isColumnHighlighted = false;
    SIZE_T CSVheaderLinesCount = 1; // Number of header lines not included in CSV sorting
    inline static POINT debugWindowPosition{ CW_USEDEFAULT, CW_USEDEFAULT };
    inline static bool  debugWindowPositionSet = false;
    inline static int   debugWindowResponse = -1;
    inline static SIZE  debugWindowSize{ 400, 250 };
    inline static bool  debugWindowSizeSet = false;
    int _editingItemIndex;
    int _editingColumnIndex;
    int _editingColumnID;
    std::string cachedFilePath;
    std::string cachedFileName;

    // Debugging and logging related 
    std::wstring findNextButtonText;        // member variable to ensure persists for button label throughout the object's lifetime.

    // Scintilla related 
    SciFnDirect pSciMsg = nullptr;
    sptr_t pSciWndData = 0;
    HighlightedTabs highlightedTabs;

    // Multi-tab storage. Active tab is _tabs[_activeTabIndex].
    std::vector<std::unique_ptr<TabState>> _tabs;
    int                                    _activeTabIndex = 0;
    int                                    _nextTabId = 0;
    int                                    _tabMenuTargetIndex = -1;   // set during context menu invocation
    HWND                                   _hTabTooltip = nullptr;
    std::wstring                           _tabTooltipText;   // owned storage; TTM_ADDTOOL keeps a pointer

    // Inline rename state for the tab control.
    HWND                                   _hTabRenameEdit = nullptr;
    int                                    _tabRenameIndex = -1;
    WNDPROC                                _prevTabEditProc = nullptr;
    HHOOK                                  _hTabRenameMouseHook = nullptr;
    static MultiReplace* _tabRenameHookOwner;

    // Drag-reorder state for the tab control.
    bool                                   _tabDragActive = false;
    int                                    _tabDragSourceIdx = -1;   // set on WM_LBUTTONDOWN
    POINT                                  _tabDragStartPt = { 0, 0 };

    // List related
    bool useListEnabled; // status for List enabled
    bool _altBypassActive = false; // momentary Ctrl+Shift bypass: use input fields while list is open
    // Suppresses live column-width read during tab switches. Without
    // this, createListViewColumns would sample the outgoing tab's
    // widths and leak them into the incoming tab.
    bool _suppressLiveWidthSync = false;

    // Tandem mode (experimental). Layout and snap math live in the
    // tandem_dock library; this class is orchestration only: a 50 ms
    // polling timer drives MR to follow N++, WM_MOVING drives the
    // drag-time magnet. INI persistence is under [Tandem] section.
    bool            _tandemEnabled = false; // feature switched on by user
    bool            _tandemDocked = false; // currently snapped to a host edge
    bool            _tandemHasSnapshot = false; // pre-dock geometry captured below is valid
    TandemDockEdge  _tandemDockEdge = TandemDockEdge::Bottom;
    UINT_PTR        _tandemTimerId = 0;
    RECT            _tandemLastNppRect = {};
    bool            _tandemUserDragging = false; // true between MR's WM_ENTERSIZEMOVE / WM_EXITSIZEMOVE

    // Sticky-magnet state. While engaged, the on-screen rect is held
    // at the edge; the original cursor-to-rect offset lets
    // tandemHandleMoving reconstruct the "real" free rect so release
    // is speed-independent.
    bool  _tandemMagnetEngaged = false;
    POINT _tandemMagnetOriginalCursorOffset = {};

    // Pre-dock geometry snapshot (valid when _tandemHasSnapshot).
    // Restored on menu-off; discarded when the user drags MR free.
    RECT _tandemSavedNppRect = {};
    RECT _tandemSavedMrRect = {};

    // User-desired OUTER size along the free axis of the current
    // dock (height for Bottom, width for Right/Left).
    int _tandemDesiredMrHeight = 0;
    int _tandemDesiredMrWidth = 0;

    bool _tandemUserResize = false; // latch: one layout pass after WM_EXITSIZEMOVE keeps user size
    bool _tandemPendingShrinkNpp = false; // latch: mouse-down tick asked for host shrink; retry next tick

    std::wstring listFilePath = L"";
    const std::size_t golden_ratio_constant = 0x9e3779b9; // 2^32 / φ, used for hashing
    std::size_t originalListHash = 0;
    int useListOnHeight = MIN_HEIGHT;      // Default height when "Use List" is on
    int checkMarkWidth_scaled;
    int crossWidth_scaled;
    int boxWidth_scaled;
    int timestampWidth_scaled;

    std::map<ColumnID, int> columnIndices;  // Mapping of ColumnID to ColumnIndex due to dynamic Columns
    std::vector<ColumnID> columnOrder;      // User-configurable column order
    static const std::vector<ColumnID> defaultColumnOrder;
    int lastTooltipRow;
    int lastTooltipSubItem;
    int lastMouseX;
    int lastMouseY;

    inline static bool tooltipsEnabled = true;            // Status for showing Tooltips on Panel
    inline static bool muteSounds = false;                // Status for Bell if String hasn't been found
    inline static bool doubleClickEditsEnabled = true;    // Double click to Edit List entries
    inline static bool highlightMatchEnabled = true;      // HighlightMatch during Find in List
    inline static bool exportToBashEnabled = false;       // shows/hides the "Export to Bash" button
    inline static bool isHoverTextEnabled = true;         // Important to set on false as TIMER will be triggered at startup.
    inline static int  editFieldSize = 5;                 // Size of the edit field for find/replace input
    inline static bool listStatisticsEnabled = false;     // Status for showing list statistics
    inline static bool stayAfterReplaceEnabled = false;   // Status for keeping panel open after replace
    inline static bool groupResultsEnabled = false;       // Status for flat list view
    inline static bool _luaSafeModeEnabled = false;        // Safer Lua mode: disables system/file/debug libs; common libs stay enabled
    inline static bool allFromCursorEnabled = false;      // Controls the starting point for Replace All, Find All and Mark when wrap is OFF.
    inline static bool keepListVisible = false;            // Library mode: list stays visible when toggled off, only dims
    inline static int  listDimIntensity = 50;               // INI setting [Options] DimIntensity (0-100): how strongly the inactive list is dimmed. Persisted, but not surfaced in the config dialog.
    inline static int  tabMaxLength = 14;                    // INI setting [Options] TabMaxLength (4-60): how many characters of a tab label are shown. Persisted, but not surfaced in the config dialog.
    inline static MultiReplaceEngine::EngineType _defaultEngine = MultiReplaceEngine::EngineType::Lua;  // INI setting [Options] DefaultEngine (Lua/ExprTk): the engine new tabs start with. Updated only on a deliberate engine selection click; loading a list with a different engine does not change it. Persisted, but not surfaced in the config dialog.
    inline static bool flowTabsIntroDontShowEnabled = false;
    inline static bool flowTabsNumericAlignEnabled = true;
    inline static bool limitFileSizeEnabled = false;
    inline static bool resultDockPerEntryColorsEnabled = true;  // Per-entry background colors in ResultDock
    inline static bool useListColorsForMarking = true;          // Use different colors per list entry when marking
    inline static size_t maxFileSizeMB = 100;

    inline static std::vector<int> _textMarkerIds;  // Fixed IDs for text marking (0-9 = list, last = single)
    inline static bool _textMarkersInitialized = false;
    inline static int _duplicateIndicatorId = -1;    // Separate indicator for duplicate marking

    // Selection Scope Management for interactive search
    std::vector<SelectionRange> m_selectionScope;
    SelectionRange m_lastFindResult = { -1, -1 };
    int m_lastTotalReplaceCount = 0;
    void adjustSelectionScope(Sci_Position replacePos, Sci_Position oldLen, Sci_Position newLen);

    inline static HWND  hDebugWnd = nullptr; // Handle for the debug window
    inline static HWND  hDebugListView = nullptr;  // ListView handle for content updates

    bool isHoverTextSuppressed = false;    // Temporarily suppress HoverText to avoid flickering when Edit in list is open

    bool _editIsExpanded = false; // track expand state
    bool _isCancelRequested = false; // Flag to signal cancellation in Replace Files
    inline static bool  _isShuttingDown = false; // Flag to signal app shutdown
    HBRUSH _hDlgBrush = nullptr; // Handle for the dialog's background brush
    bool _flowTabsActive = false;   // current visual state (editor tabstops)
    int  _flowPaddingPx = 8;       // min pixel gap after each column

    enum class CsvOp {
        Sort,
        DeleteColumns,
        CopyColumns,
        Replace,        // single or all
        Mark,           // mark/highlight
        BatchReplace,   // replace in files or multi-doc
    };

    // ETabs presentation mode used by the layer (derived internally by default)
    enum class EtabsMode { Off, Visual, Padding };

    // GUI control-related constants
    const int maxHistoryItems = 10;  // Maximum number of history items to be saved for Find/Replace

    // Coloring Slot for marking
    std::unordered_map<std::wstring, int> textToSlot;
    int nextSlot = 0;

    // Window related settings
    RECT windowRect; // Structure to store window position and size
    BYTE foregroundTransparency;     // Default Foreground transparency
    BYTE backgroundTransparency;     // Default Background transparency
    int findCountColumnWidth;        // Width of the "Find Count" column
    int replaceCountColumnWidth;     // Width of the "Replace Count" column
    int findColumnWidth;             // Width of the "Find what" column
    int replaceColumnWidth;          // Width of the "Replace" column
    int commentsColumnWidth;         // Width of the "Comments" column
    int deleteButtonColumnWidth;     // Width of the "Delete" column
    bool isFindCountVisible;         // Visibility of the "Find Count" column
    bool isReplaceCountVisible;      // Visibility of the "Replace Count" column
    bool isCommentsColumnVisible;    // Visibility of the "Comments" column
    bool isTimestampColumnVisible;   // Visibility of the "Timestamp" column
    bool isDeleteButtonVisible;      // Visibility of the "Delete" column
    bool findColumnLockedEnabled;    // Indicates if the "Find what" column is locked
    bool replaceColumnLockedEnabled; // Indicates if the "Replace" column is locked
    bool commentsColumnLockedEnabled;// Indicates if the "Comments" column is locked

    // Window DPI scaled size 
    int MIN_WIDTH_scaled;
    int MIN_HEIGHT_scaled;
    int SHRUNK_HEIGHT_scaled;
    int MIN_GENERAL_WIDTH_scaled;
    int DEFAULT_COLUMN_WIDTH_FIND_scaled;
    int DEFAULT_COLUMN_WIDTH_REPLACE_scaled;
    int DEFAULT_COLUMN_WIDTH_COMMENTS_scaled;
    int DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled;
    int DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled;

    // Initialization
    void initializeWindowSize();
    RECT calculateMinWindowFrame(HWND hwnd);
    void createFonts();
    void cleanupFonts();
    void applyFonts();
    HFONT font(FontRole role) const;
    void positionAndResizeControls(int windowWidth, int windowHeight);
    void initializeCtrlMap();
    bool createAndShowWindows();
    void ensureIndicatorContext();
    void initializeListView();
    void moveAndResizeControls(bool moveStatic);
    void updateTwoButtonsVisibility();
    void updateListViewFrame();
    void repaintPanelContents(HWND hGrp, const std::wstring& title);
    void updateFilesPanel();
    void updateAllDocsCheckboxState();
    bool isFilesPanelNeeded() const;
    bool matchesDocFilter(const std::wstring& fileName, const std::wstring& filter) const;
    void setUIElementVisibility();
    void drawGripper();
    void SetWindowTransparency(HWND hwnd, BYTE alpha);
    void adjustWindowSize();
    void updateUseListState(bool isUpdate);

    //List Data Operations
    void addItemsToReplaceList(const std::vector<ReplaceItemData>& items, size_t insertPosition);
    void removeItemsFromReplaceList(const std::vector<size_t>& indicesToRemove);
    void modifyItemInReplaceList(size_t index, const ReplaceItemData& newData);
    bool moveItemsInReplaceList(std::vector<size_t>& indices, Direction direction);
    void sortItemsInReplaceList(const std::vector<size_t>& originalOrder, const std::vector<size_t>& newOrder, const std::map<int, SortDirection>& previousColumnSortOrder, int columnID, SortDirection direction);
    void scrollToIndices(size_t firstIndex, size_t lastIndex);
    void exportDataToClipboard();


    //ListView
    HWND CreateHeaderTooltip(HWND hwndParent);
    void AddHeaderTooltip(HWND hwndTT, HWND hwndHeader, int columnIndex, LPCTSTR pszText);
    void createListViewColumns(WidthMode mode = WidthMode::Redistribute);
    // Mark every inactive tab as needing a full redistribute the next
    // time it is activated. Called on WM_SIZE so tabs sized before
    // the change still fit the panel when the user revisits them.
    void markAllTabsNeedRelayout();
    void insertSingleColumn(ColumnID id, int& currentIndex, int perColumnWidth, LVCOLUMN& lvc);
    bool isColumnVisible(ColumnID id) const;
    bool validateColumnOrder(const std::vector<ColumnID>& order) const;
    void syncColumnOrderFromHeader();
    void initColumnOrder();
    void insertReplaceListItem(const ReplaceItemData& itemData);
    int  getColumnWidth(ColumnID columnID);
    int  calcDynamicColWidth(const ResizableColWidths& widths);
    void updateListViewAndColumns();
    void updateListViewItem(size_t index);
    void updateListViewTooltips();
    void updateHeaderSelection();
    void updateHeaderSortDirection();
    void showColumnVisibilityMenu(HWND hWnd, POINT pt);
    // Column that was right-clicked in the header. Stashed by
    // showColumnVisibilityMenu so the Lock/Unlock command handler
    // knows which column to toggle.
    ColumnID _tagetedLockColumn = ColumnID::INVALID;

    // ----- Engine selector UI -----------------------------------------
    // Pop up the small engine-chooser menu next to the (L)/(E) marker.
    // The marker itself is IDC_USE_VARIABLES_ENGINE; this opens the menu
    // that lets the user switch between Lua and ExprTk for the active tab.
    void showEngineSelectorMenu();

    // Apply a new engine selection to the active tab. Updates the tab's
    // EngineType, drops the cached engine instance so getActiveEngine()
    // creates a fresh one on next use, and refreshes the (L)/(E) marker.
    void applyEngineSelection(MultiReplaceEngine::EngineType type);

    // Refresh the (L)/(E) marker text from the active tab's engine. Also
    // enables/disables the marker according to the "Use Variables"
    // checkbox state.
    void syncEngineSelectorLabel();

    void handleCopyBack(NMITEMACTIVATE* pnmia);
    void handleUpdateFromFields();
    static std::wstring getCurrentTimestamp();
    void shiftListItem(const Direction& direction);
    void handleDeletion(NMITEMACTIVATE* pnmia);
    void deleteSelectedLines();
    void sortReplaceListData(int columnID);
    size_t generateUniqueID();
    std::vector<size_t> getSelectedRows();
    void selectRows(const std::vector<size_t>& selectedIDs);
    void handleCopyToListButton();
    void resetCountColumns();
    void updateCountColumns(const size_t itemIndex, const int findCount, int replaceCount = -1);
    void refreshUIListView();
    void handleColumnVisibilityToggle(UINT menuId);
    ColumnID getColumnIDFromIndex(int columnIndex) const;
    int getColumnIndexFromID(ColumnID columnID) const;

    //UI Settings
    void onTooltipsToggled(bool enable);
    void destroyAllTooltipWindows();
    void rebuildAllTooltips();

    //Contextmenu List
    void toggleBooleanAt(int itemIndex, ColumnID columnID);
    void editTextAt(int itemIndex, ColumnID columnID);
    void closeEditField(bool commitChanges);
    static LRESULT CALLBACK ListViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    static LRESULT CALLBACK EditControlSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    void toggleEditExpand();
    void createContextMenu(HWND hwnd, POINT ptScreen, MenuState state);
    MenuState checkMenuConditions(POINT ptScreen);
    void performItemAction(POINT pt, ItemAction action);
    void copySelectedItemsToClipboard();
    bool canPasteFromClipboard();
    void pasteItemsIntoList();
    void handleEditOnDoubleClick(int itemIndex, ColumnID columnID);
    void toggleListSearchBar();
    void showListSearchBar();
    void hideListSearchBar();
    void findInList(bool forward);
    int searchInListData(int startIdx, const std::wstring& searchText, bool forward);
    void jumpToNextMatchInEditor(size_t listIndex);

    //Replace
    void replaceAllInOpenedDocs();
    bool handleReplaceAllButton(bool showCompletionMessage = true, const std::filesystem::path* explicitPath = nullptr);
    void handleReplaceButton();
    bool replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount, const size_t itemIndex = SIZE_MAX);
    bool replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos, size_t itemIndex, const SearchContext& context);
    Sci_Position performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    Sci_Position performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    void updateLineDelimiterAfterReplace(Sci_Position pos);
    bool preProcessListForReplace(bool highlight);
    SelectionInfo getSelectionInfo(bool isBackward);
    bool hasAnyNonEmptySelection();
    bool hasAnySelectedEntry() const;
    Sci_Position computeAllStartPos(const SearchContext& context, bool wrapEnabled, bool fromCursorEnabled);

    void updateFilePathCache(const std::filesystem::path* explicitPath = nullptr);

    // ----- ILuaEngineHost implementation ------------------------------
    // Bridge implementations forwarded to existing MR helpers; declared
    // here so the formula engine can call back without including the
    // whole panel header. See engine/LuaEngine.h for the contract.
    std::string escapeForRegex(const std::string& input) override;
    int          showDebugWindow(const std::string& message) override;
    void         refreshUiListView() override;
    void         showErrorMessage(MultiReplaceEngine::ILuaEngineHost::ErrorCategory category,
        const std::string& engineName,
        const std::string& details) override;
    bool         isLuaErrorDialogEnabled() const override;
    bool         isLuaSafeModeEnabled()    const override;
    bool         isDebugModeEnabled()      const override;

    // Replace-pipeline helpers that work in terms of the engine-agnostic
    // FormulaVars structure. Used by the replace pipeline to populate
    // the per-match data the active engine consumes.
    void fillFormulaVars(MultiReplaceEngine::FormulaVars& vars,
        Sci_Position matchPos,
        const std::string& foundText,
        int cnt, int lcnt,
        bool isColumnMode,
        int documentCodepage);
    void fillCapturesForEngine(MultiReplaceEngine::FormulaVars& vars,
        int documentCodepage);

    // Returns the active tab's engine, creating it on first call.
    // Returns nullptr if engine creation fails.
    MultiReplaceEngine::IFormulaEngine* getActiveEngine();

    //Lua Debug Window
    int ShowDebugWindow(const std::string& message);
    static LRESULT CALLBACK DebugWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static void CopyListViewToClipboard(HWND hListView);
    static void CloseDebugWindow();
    void SetDebugComplete();
    void WaitForDebugWindowClose(bool autoClose = false);

    //Replace in Files
    bool handleBrowseDirectoryButton();
    bool selectDirectoryDialog(HWND owner, std::wstring& outPath);
    void handleReplaceInFiles();

    //Find All
    std::wstring sanitizeSearchPattern(const std::wstring& raw);
    void trimHitToFirstLine(const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend, ResultDock::Hit& h);
    void handleFindAllButton();
    void handleFindAllInDocsButton();
    void handleFindInFiles();

    //Find
    void handleFindNextButton();
    void handleFindPrevButton();
    SearchResult performSingleSearch(const SearchContext& context, SelectionRange range);
    SearchResult performSearchForward(const SearchContext& context, LRESULT start);
    SearchResult performSearchBackward(const SearchContext& context, LRESULT start);
    SearchResult performSearchColumn(const SearchContext& context, LRESULT start, bool isBackward);
    SearchResult performSearchSelection(const SearchContext& context, LRESULT start, bool isBackward);
    SearchResult performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context);
    SearchResult performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context);
    void displayResultCentered(size_t posStart, size_t posEnd, bool isDownwards);
    void selectListItem(size_t matchIndex);
    void updateSelectionScope();
    void captureCurrentSelectionAsScope();
    std::wstring getSelectionScopeSuffix();

    //Mark
    void handleMarkMatchesButton();
    // markString runs the search-and-mark loop. If bookmarkMarkerId is
    // non-negative, every match line additionally gets a bookmark via
    // SCI_MARKERADD — driven by the "+ Bookmarks" checkbox next to the
    // Mark Matches button.
    int markString(const SearchContext& context, Sci_Position initialStart,
        const std::wstring& findText = L"",
        int bookmarkMarkerId = -1);
    int calcMaxListSlots() const;
    int resolveIndicatorForText(const std::wstring& findText);
    void handleClearTextMarksButton();
    void handleCopyMarkedTextToClipboardButton();
    void copyTextToClipboard(const std::wstring& text, int textCount);
    void initTextMarkerIndicators();
    void updateTextMarkerStyles();
    std::vector<size_t> getIndicesOfUniqueEnabledItems(bool removeDuplicates) const;

    //CSV
    void handleCopyColumnsToClipboard();
    bool confirmColumnDeletion();
    void handleDeleteColumns();
    void handleColumnGridTabsButton();
    void handleDuplicatesButton();
    void findAndMarkDuplicates(bool showDialog = true);
    bool scanForDuplicates();
    bool validateAndRescanIfNeeded();
    void applyDuplicateMarks();
    void clearDuplicateMarks();
    void showDeleteDuplicatesDialog();
    void deleteDuplicateLines();
    void clearFlowTabsIfAny();
    bool buildCTModelFromMatrix(ColumnTabs::CT_ColumnModelView& outModel) const;
    bool applyFlowTabStops(const ColumnTabs::CT_ColumnModelView* existingModel = nullptr);
    bool runCsvWithFlowTabs(CsvOp op, const std::function<bool()>& body);
    bool showFlowTabsIntroDialog(bool& dontShowFlag) const;
    ViewState saveViewState() const;
    void restoreViewStateExact(const ViewState& s);

    //CSV Sort
    std::vector<CombinedColumns> extractColumnData(SIZE_T startLine, SIZE_T lineCount);
    void sortRowsByColumn(SortDirection sortDirection);
    void reorderLinesInScintilla(const std::vector<size_t>& sortedIndex);
    void restoreOriginalLineOrder(const std::vector<size_t>& originalOrder);
    void UpdateSortButtonSymbols();
    void handleSortStateAndSort(SortDirection direction);
    void updateUnsortedDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType);
    void detectNumericColumns(std::vector<CombinedColumns>& data);
    int compareColumnValue(const ColumnValue& left, const ColumnValue& right);

    //Scope
    bool parseColumnAndDelimiterData();
    bool validateDelimiterData();
    void findAllDelimitersInDocument();
    void findDelimitersInLine(LRESULT line);
    ColumnInfo getColumnInfo(LRESULT startPosition);
    LRESULT adjustForegroundForDarkMode(LRESULT textColor, LRESULT backgroundColor);
    void initializeColumnStyles();
    void handleHighlightColumnsInDocument();
    void reapplyColumnHighlighting();
    void highlightColumnsInLine(LRESULT line);
    void fixHighlightAtDocumentEnd();
    void handleClearColumnMarks();
    std::wstring addLineAndColumnMessage(LRESULT pos);
    void updateDelimitersInDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType);
    void processLogForDelimiters();
    void handleDelimiterPositions(DelimiterOperation operation);
    void handleClearDelimiterState();

    //Utilities
    std::string convertAndExtendW(const std::wstring& input, bool extended, UINT cp) const;
    std::string convertAndExtendW(const std::wstring& input, bool extended);
    static void addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems = 100);
    std::wstring getTextFromDialogItem(HWND hwnd, int itemID) const;
    void setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text);
    void setSelections(bool select, bool onlySelected = false);
    void setOptionForSelection(SearchOption option, bool value);
    void showStatusMessage(const std::wstring& messageText, MessageStatus status, bool isNotFound = false, bool isTransient = false);
    void drainMessageQueue();
    void applyThemePalette();
    void refreshColumnStylesIfNeeded();
    std::wstring getShortenedFilePath(const std::wstring& path, int maxLength, HDC hDC = nullptr);
    std::wstring getSelectedText();
    LRESULT getEOLLengthForLine(LRESULT line);
    std::string getEOLStyle();
    sptr_t send(unsigned int iMessage, uptr_t wParam = 0, sptr_t lParam = 0, bool useDirect = true) const;
    bool normalizeAndValidateNumber(std::string& str);
    std::vector<WCHAR> createFilterString(const std::vector<std::pair<std::wstring, std::wstring>>& filters);
    int getFontHeight(HWND hwnd, HFONT hFont);
    std::vector<int> parseNumberRanges(const std::wstring& input, const std::wstring& errorMessage);
    UINT getCurrentDocCodePage();
    std::size_t computeListHash(const std::vector<ReplaceItemData>& list);
    Sci_Position advanceAfterMatch(const SearchResult& r);
    Sci_Position ensureForwardProgress(Sci_Position nextPos, const SearchResult& r);
    void MultiReplace::forceWrapRecalculation();

    //FileOperations
    std::wstring promptSaveListToCsv(const TabState* tabHint = nullptr);
    std::wstring openFileDialog(bool saveFile, const std::vector<std::pair<std::wstring, std::wstring>>& filters, const WCHAR* title, DWORD flags, const std::wstring& fileExtension, const std::wstring& defaultFilePath);
    bool saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    bool saveListToCsvWithSettings(const std::wstring& filePath, const std::vector<ReplaceItemData>& list, const TabState& tab);
    bool saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list, TabState* tabForSettings = nullptr);
    void autoShowCommentsColumn();
    void checkForFileChangesAtStartup();
    void exportToBashScript(const std::wstring& fileName);

    // Save-All: writes every dirty tab. Tabs with a path are saved
    // silently; tabs without a path show a Save As dialog per tab
    // (user can cancel individually, like Notepad++).
    void saveAllTabs();

    //INI
    void saveSettings();
    // 3-button Save/Don't save/Cancel prompt. tabIndex < 0 = active
    // tab via globals; concrete index queries that tab directly.
    int checkForUnsavedChanges(int tabIndex = -1);
    void loadSettings();
    void syncHistoryToCache(HWND hComboBox, const std::wstring& keyPrefix);

    // Multi-tab helpers. Live working state stays in the legacy
    // members (replaceListData, listFilePath, ...); a TabState is
    // a snapshot that is swapped in/out on tab switch.
    void ensurePrimaryTabExists();
    void captureStateIntoTab(TabState& tab);
    void restoreStateFromTab(const TabState& tab);

    // Syncs live ListView column widths back into the global width
    // members so subsequent captures see the user's latest resize.
    void readColumnWidthsFromListView();

    // Multi-tab persistence.
    void writeTabsToConfig();
    void loadTabsFromConfig();
    void migrateLegacyList();
    void dropLegacyConfigEntries();
    void saveAllTabSnapshots();
    void cleanupOrphanSnapshots();
    void checkSingleTabForFileChange(int tabIndex);

    // Tab control rendering.
    void rebuildTabControl();
    // Lay out the tab strip's right-hand controls so they always sit
    // directly adjacent to the last visible tab. The "+" button sticks
    // immediately to the tab control; the overflow dropdown ("v") sits
    // next to "+" but is only shown when tabs are clipped.
    void repositionNewTabButton();

    // Open a popup listing every tab; clicking an entry switches to it.
    // Used by the overflow dropdown.
    void showTabListPopup();

    // Scroll the tab strip far enough that the indicated tab is fully
    // visible. No-op if it is already on screen.
    void ensureTabVisible(int tabIndex);

    // Scroll the strip by one tab in the given direction (negative
    // left, positive right). Used by the "..." indicators.
    void scrollTabStrip(int direction);

    void updateTabTooltip(int tabIndex);
    static std::wstring truncateTabName(const std::wstring& name, size_t maxChars = 14);

    // Tab dirty-indicator management. markActiveTabDirty compares the
    // current list hash against the baseline from the last save/load
    // and toggles the flag; undoing back to the saved state therefore
    // clears it automatically. clearTabDirty force-clears after a
    // successful save. Both rebuild the tab control only on real
    // state changes.
    void markActiveTabDirty();
    void clearTabDirty(int tabIndex);

    // Shows or hides the tab bar together with the rest of the list UI.
    void setBottomRowVisible(bool visible);

    // Switches active tab, capturing current state into the outgoing
    // tab and restoring the incoming tab into the live working state.
    void switchToTab(int newIndex);

    // Creates a new empty untitled tab and makes it active.
    void addNewTab();

    // Copies column layout (widths, visibility, locks, order) from
    // the active tab into the destination tab so new tabs inherit
    // the user's workspace preferences. Sort order and list data
    // are intentionally not copied. No-op if no active tab exists.
    void inheritLayoutFromActiveTab(TabState& dst) const;

    // Reorders tabs: moves the tab at fromIdx so it sits at toIdx in
    // the tab vector. Updates _activeTabIndex consistently and
    // rebuilds the tab control. Idempotent if fromIdx == toIdx.
    void moveTab(int fromIdx, int toIdx);

    // Tab context menu actions.
    void showTabContextMenu(int tabIndex, int screenX, int screenY);
    void onTabRename(int tabIndex);
    void onTabDuplicate(int tabIndex);
    void onTabClose(int tabIndex);
    void onTabCloseOthers(int keepTabIndex);
    void onTabCloseAll();
    void onTabSave(int tabIndex);
    void onTabSaveAs(int tabIndex);
    void onTabOpenFileLocation(int tabIndex);

    // Find an open tab by its file path (case-insensitive on Windows).
    // Returns the tab index, or -1 if no tab has this file open.
    int findTabByFilePath(const std::wstring& filePath) const;

    // Inline rename editing on a tab.
    void beginInlineTabRename(int tabIndex);
    void commitInlineTabRename();
    void cancelInlineTabRename();

    // File-backed tabs rename their .mrl on disk via Save-As + delete.
    void renameTabFile(int tabIndex);
    static LRESULT CALLBACK tabRenameEditProc(HWND, UINT, WPARAM, LPARAM);
    static LRESULT CALLBACK tabControlSubclassProc(HWND, UINT, WPARAM, LPARAM,
        UINT_PTR, DWORD_PTR);
    static LRESULT CALLBACK tabRenameMouseHookProc(int, WPARAM, LPARAM);

    //Debug DPI Information
    void showDPIAndFontInfo();

    static void displayLogChangesInMessageBox();

    // Tandem mode internals. Timer-driven follow, drag-time magnet,
    // and the free <-> docked state-transition helpers.
    void onTandemTick();                    // docked tick: follow N++
    void onTandemFreeTick();                // free tick: re-engage when N++ approaches
    void applyTandemLayout(const RECT& nppRect);
    void tandemHandleMoving(RECT* pTargetRect);     // WM_MOVING snap
    void tandemHandleExitSizeMove();                // WM_EXITSIZEMOVE state transition
    void tandemDockToCurrentEdge();         // free -> docked (captures pre-dock snapshot)
    void tandemUndockAndRestore();          // docked -> free (restores snapshot)
    void tandemPersistEdgeToIni() const;    // write last dock edge to INI cache
    bool tandemLoadEdgeFromIni();           // read last dock edge; false if no persisted value
    bool tandemRestoreFromIniIfEnabled();   // silent startup restore; returns whether state changed
};

extern MultiReplace _MultiReplace;

#endif // MULTI_REPLACE_H