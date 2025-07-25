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

#include "StaticDialog/StaticDialog.h"
#include "StaticDialog/resource.h"
#include "PluginInterface.h"
#include "DropTarget.h"
#include "DPIManager.h"
#include "ResultDock.h"
#include "Encoding.h"

#include <string>
#include <vector>
#include <map>
#include <functional>
#include <regex>
#include <algorithm>
#include <unordered_map>
#include <unordered_set>
#include <set>
#include <filesystem>

#include <commctrl.h>

#include <lua.hpp>

extern NppData nppData;

enum class DelimiterOperation { LoadAll, Update };
enum class Direction { Up, Down };

struct ReplaceItemData
{
    size_t id = 0;
    std::wstring findCount = L"";
    std::wstring replaceCount = L"";
    bool isEnabled = true;
    std::wstring findText;
    std::wstring replaceText;
    bool wholeWord = false;
    bool matchCase = false;
    bool useVariables = false;
    bool extended = false;
    bool regex = false;
    std::wstring comments = L"";

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

struct WindowSettings {
    int posX;
    int posY;
    int width;
    int height;
};

struct ControlInfo
{
    int x, y, cx, cy;
    LPCWSTR className;
    LPCWSTR windowName;
    DWORD style;
    LPCWSTR tooltipText;
};

struct SearchContext {
    std::string findText = "";      // search string in Scintilla encoding
    int searchFlags = 0;            // Search flags (e.g., SCFIND_MATCHCASE, SCFIND_WHOLEWORD, etc.)
    LRESULT docLength = 0;          // Cached document length
    bool isColumnMode = false;      // Cached state: true if Column Mode is active
    bool isSelectionMode = false;   // Cached state: true if Selection Mode is active
    bool retrieveFoundText = false; // If true, retrieve the found text from Scintilla
    bool highlightMatch = false;    // If true, highlight the found match
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
    double      numericValue = 0;
    std::string text = "";
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
    size_t lineIndex = 0;   // which line it is (for Fenwicksum queries)
};

struct ColumnInfo {
    LRESULT totalLines;
    LRESULT startLine;
    SIZE_T startColumnIndex;
};

struct ContextMenuInfo {
    int hitItem = -1;
    int clickedColumn = -1;
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
    bool canCopy = false;
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

enum class SortDirection {
    Unsorted,
    Ascending,
    Descending
};

enum class SearchDirection {
    Forward,
    Backward
};

// Lua Engine
struct LuaVariables {
    int CNT = 0;
    int LINE = 0;
    int LPOS = 0;
    int LCNT = 0;
    int APOS = 0;
    int COL = 0;
    std::string MATCH = "";
    std::string FPATH = "";
    std::string FNAME = "";
};

enum class LuaVariableType {
    String,
    Number,
    Boolean,
    None
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
    DELETE_BUTTON       // 11
};

struct ResizableColWidths {
    HWND listView;
    int listViewWidth;
    int findCountWidth;
    int replaceCountWidth;
    int findWidth;
    int replaceWidth;
    int commentsWidth;
    int deleteWidth;
    int margin;
};

struct LuaVariable {
    std::string name;
    LuaVariableType type;
    std::string stringValue;
    double numberValue;
    bool booleanValue;

    LuaVariable() : name(""), type(LuaVariableType::None), numberValue(0.0), booleanValue(false) {}
};

struct EncodingInfo {
    int sc_codepage = 0;      // The codepage value for Scintilla (e.g., SC_CP_UTF8)
    size_t bom_length = 0;    // The length of the BOM in bytes (0 if no BOM)
};

using LuaVariablesMap = std::map<std::string, LuaVariable>;

class CsvLoadException : public std::exception {
public:
    explicit CsvLoadException(const std::string& message) : message_(message) {}
    const char* what() const noexcept override {
        return message_.c_str();
    }
private:
    std::string message_;
};

class LuaSyntaxException : public std::exception {
};

struct EditControlContext
{
    MultiReplace* pThis;
    HWND hwndExpandBtn;
};

// each new Vector has to be delared outside of the class due to unresolved memory behaviours, 
// possible initial limited stack size in N++ for Plugins
inline std::vector<UndoRedoAction> undoStack;
inline std::vector<UndoRedoAction> redoStack;
inline HWND hwndExpandBtn = nullptr;
inline HFONT _hBoldFont2;
inline lua_State* _luaState = nullptr;    // Reused Lua state

inline bool _editIsExpanded = false; // track expand state
using IniData = std::map<std::wstring, std::map<std::wstring, std::wstring>>;
inline IniData iniCache;

class MultiReplace : public StaticDialog
{
public:
    MultiReplace() :
        hInstance(NULL),
        _hScintilla(0),
        _hClearMarksButton(nullptr),
        _hCopyMarkedTextButton(nullptr),
        _hInListCheckbox(nullptr),
        _hMarkMatchesButton(nullptr),
        _hReplaceAllButton(nullptr),
        _replaceListView(NULL),
        _hStandardFont(nullptr),
        _hBoldFont1(nullptr),
        _hNormalFont1(nullptr),
        _hNormalFont2(nullptr),
        _hNormalFont3(nullptr),
        _hNormalFont4(nullptr),
        _hNormalFont5(nullptr),
        _hNormalFont6(nullptr),
        _hStatusMessage(nullptr),
        _statusMessageColor(RGB(0, 0, 0))
    {
        setInstance(this);
    };
    static MultiReplace* instance; // Static instance of the class

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

    // Light Mode Colors for Message
    static constexpr COLORREF LMODE_SUCCESS = RGB(0, 128, 0);
    static constexpr COLORREF LMODE_ERROR = RGB(200, 0, 0);
    static constexpr COLORREF LMODE_INFO = RGB(0, 0, 128);
    static constexpr COLORREF LMODE_FILTER_HELP = RGB(0, 0, 255);

    // Dark Mode Colors for Message
    static constexpr COLORREF DMODE_SUCCESS = RGB(120, 220, 120); 
    static constexpr COLORREF DMODE_ERROR = RGB(255, 110, 110);
    static constexpr COLORREF DMODE_INFO = RGB(147, 147, 255);
    static constexpr COLORREF DMODE_FILTER_HELP = RGB(255, 235, 59);

    static bool isWindowOpen;
    static bool textModified;
    static bool documentSwitched;
    static int scannedDelimiterBufferID;
    static bool isLoggingEnabled;
    static bool isCaretPositionEnabled;
    static bool isLuaErrorDialogEnabled;

    static std::vector<size_t> originalLineOrder; // Stores the order of lines before sorting
    static SortDirection currentSortState; // Status of column sort
    static bool isSortedColumn; // Indicates if a column is sorted

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

    enum class ChangeType { Insert, Delete, Modify };
    enum class ReplaceMode { Normal, Extended, Regex };

    struct LogEntry {
        ChangeType changeType = ChangeType::Modify;
        Sci_Position lineNumber = 0;
        Sci_Position blockSize = 1;
    };

    static std::vector<LogEntry> logChanges;

    // Drag-and-Drop functionality
    DropTarget* dropTarget;  // Pointer to DropTarget instance
    void loadListFromCsv(const std::wstring& filePath); // used in DropTarget.cpp
    void showListFilePath();
    void initializeDragAndDrop();
    std::wstring getLangStr(const std::wstring& id, const std::vector<std::wstring>& replacements = {});

protected:
    virtual INT_PTR CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam) override;

private:
    static constexpr int MAX_TEXT_LENGTH = 4096; // Maximum text length for Find and Replace String
    static constexpr int MARKER_COLOR = 0x007F00; // Color for non-list Marker
    static constexpr LRESULT PROGRESS_THRESHOLD = 50000; // Will show progress bar if total exceeds defined threshold
    static constexpr int REPLACE_FILES_PANEL_HEIGHT = 88;
    bool isReplaceAllInDocs = false;   // True if replacing in all open documents, false for current document only.
    bool isReplaceInFiles = false;   // True if replacing in files, false for current document only.

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
    static constexpr int MAX_EDIT_FIELD_SIZE = 20; // MAximum size for Multiline Editor

    static constexpr int STEP_SIZE = 5; // Speed for opening and closing Count Columns
    static constexpr wchar_t* symbolSortAsc = L"▼";
    static constexpr wchar_t* symbolSortDesc = L"▲";
    static constexpr wchar_t* symbolSortAscUnsorted = L"▽";
    static constexpr wchar_t* symbolSortDescUnsorted = L"△";
    static constexpr int MAX_CAP_GROUPS = 9; // Maximum number of capture groups supported by Notepad++

    DPIManager* dpiMgr; // Pointer to DPIManager instance

    // Static variables related to GUI 
    static HWND s_hScintilla;
    static HWND s_hDlg;
    HWND hwndEdit = NULL;
    WNDPROC originalListViewProc;
    static std::map<int, ControlInfo> ctrlMap;

    // Instance-specific GUI-related variables 
    HINSTANCE hInstance;
    HWND _hScintilla;
    HWND _hClearMarksButton;
    HWND _hCopyMarkedTextButton;
    HWND _hInListCheckbox;
    HWND _hMarkMatchesButton;
    HWND _hReplaceAllButton;
    HWND _replaceListView;
    HWND _hStatusMessage;
    HFONT _hStandardFont;
    HFONT _hBoldFont1;
    HFONT _hNormalFont1;
    HFONT _hNormalFont2;
    HFONT _hNormalFont3;
    HFONT _hNormalFont4;
    HFONT _hNormalFont5;
    HFONT _hNormalFont6;
    COLORREF COLOR_SUCCESS;
    COLORREF COLOR_ERROR;
    COLORREF COLOR_INFO;
    COLORREF _statusMessageColor = LMODE_INFO; // Holds the actual color to be drawn. Initialized for light mode.
    COLORREF _filterHelpColor;
    MessageStatus _lastMessageStatus = MessageStatus::Info; // Holds the TYPE of the last message.
    HWND _hHeaderTooltip;        // Handle to the tooltip for the ListView header
    HWND _hUseListButtonTooltip; // Handle to the tooltip for the Use List Button

    // ContextMenuInfo structure instance
    POINT _contextMenuClickPoint;

    // Style-related variables and constants
    /*
       Available styles (self-tested):
       { 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 }
       Note: Gaps in the list are intentional.

       Styles 0 - 7 are reserved for syntax style.
       Styles 21 - 29, 31 are reserved by N++ (see SciLexer.h).
    */
    const std::vector<int> textStyles = { 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 30, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 };
    const std::vector<int> hColumnStyles = { STYLE1, STYLE2, STYLE3, STYLE4, STYLE5, STYLE6, STYLE7, STYLE8, STYLE9, STYLE10 };
    const std::vector<int> lightModeColumnColors = { 0xFFE0E0, 0xC0E0FF, 0x80FF80, 0xFFE0FF,  0xB0E0E0, 0xFFFF80, 0xE0C0C0, 0x80FFFF, 0xFFB0FF, 0xC0FFC0 };
    const std::vector<int> darkModeColumnColors  = { 0x553333, 0x335577, 0x225522, 0x553355, 0x335555, 0x555522, 0x774444, 0x225555, 0x553366, 0x336633 };

    // Data-related variables 
    size_t markedStringsCount = 0;
    bool allSelected = true;
    std::vector<char> stylesBuffer;
    std::unordered_map<int, int> colorToStyleMap;
    std::map<int, SortDirection> columnSortOrder;
    ColumnDelimiterData columnDelimiterData;
    std::vector<ReplaceItemData> replaceListData;
    std::vector<LineInfo> lineDelimiterPositions;
    std::vector<char> lineBuffer; // reusable Buffer for findDelimitersInLine()
    bool isColumnHighlighted = false;
    std::map<int, bool> stateSnapshot; // stores the state of the Elements
    LuaVariablesMap globalLuaVariablesMap; // stores Lua Global Variables
    SIZE_T CSVheaderLinesCount = 1; // Number of header lines not included in CSV sorting
    bool isStatisticsColumnsExpanded = false;
    static POINT debugWindowPosition;
    static bool debugWindowPositionSet;
    static int debugWindowResponse;
    static SIZE debugWindowSize;
    static bool debugWindowSizeSet;
    static HWND hDebugWnd; // Handle for the debug window
    int _editingItemIndex;
    int _editingColumnIndex;
    int _editingColumnID;
    std::string cachedFilePath;
    std::string cachedFileName;
    int _luaCompiledReplaceRef = LUA_NOREF;       // Reference to compiled Lua code
    std::string _lastCompiledLuaCode;             // Cached Lua code for reuse

    // Debugging and logging related 
    std::string messageBoxContent;  // just for temporary debugging usage
    std::wstring findNextButtonText;        // member variable to ensure persists for button label throughout the object's lifetime.

    // Scintilla related 
    SciFnDirect pSciMsg = nullptr;
    sptr_t pSciWndData = 0;
    HighlightedTabs highlightedTabs;

    // List related
    bool useListEnabled; // status for List enabled
    std::wstring listFilePath = L""; //to store the file path of loaded list
    const std::size_t golden_ratio_constant = 0x9e3779b9; // 2^32 / φ /uused for Hashing
    std::size_t originalListHash = 0;
    int useListOnHeight = MIN_HEIGHT;      // Default height when "Use List" is on
    const int useListOffHeight = SHRUNK_HEIGHT; // Height when "Use List" is off (constant)
    int checkMarkWidth_scaled;
    int crossWidth_scaled;
    int boxWidth_scaled;

    std::map<ColumnID, int> columnIndices;  // Mapping of ColumnID to ColumnIndex due to dynamic Columns
    int lastTooltipRow;
    int lastTooltipSubItem;
    int lastMouseX;
    int lastMouseY;

    inline static bool tooltipsEnabled = true;            // Status for showing Tooltips on Panel
    inline static bool alertNotFoundEnabled = true;       // Status for Bell if String hasn't been found
    inline static bool doubleClickEditsEnabled = true;    // Double click to Edit List entries
    inline static bool highlightMatchEnabled = true;      // HighlightMatch during Find in List
    inline static bool exportToBashEnabled = false;      // shows/hides the "Export to Bash" button
    inline static bool isHoverTextEnabled = true;        // Important to set on false as TIMER will be triggered at startup.
    inline static int  editFieldSize = 5;                // Size of the edit field for find/replace input
    inline static bool listStatisticsEnabled = false;    // Status for showing list statistics
    inline static bool stayAfterReplaceEnabled = false; // Status for keeping panel open after replace
    inline static bool flatListEnabled = false;          // Status for flat list view

    bool isHoverTextSuppressed = false;    // Temporarily suppress HoverText to avoid flickering when Edit in list is open


    bool _isCancelRequested = false; // Flag to signal cancellation in Replace Files
    static bool _isShuttingDown; // Flag to signal app shutdown
    HBRUSH _hDlgBrush = nullptr; // Handle for the dialog's background brush

    // GUI control-related constants
    const int maxHistoryItems = 10;  // Maximum number of history items to be saved for Find/Replace

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
    bool isDeleteButtonVisible;      // Visibility of the "Delete" column
    bool findColumnLockedEnabled;    // Indicates if the "Find what" column is locked
    bool replaceColumnLockedEnabled; // Indicates if the "Replace" column is locked
    bool commentsColumnLockedEnabled;// Indicates if the "Comments" column is locked

    // Window DPI scaled size 
    int MIN_WIDTH_scaled;
    int MIN_HEIGHT_scaled;
    int SHRUNK_HEIGHT_scaled;
    int COUNT_COLUMN_WIDTH_scaled;
    int MIN_FIND_REPLACE_WIDTH_scaled;
    int MIN_GENERAL_WIDTH_scaled;
    int COMMENTS_COLUMN_WIDTH_scaled;
    int DEFAULT_COLUMN_WIDTH_FIND_scaled;
    int DEFAULT_COLUMN_WIDTH_REPLACE_scaled;
    int DEFAULT_COLUMN_WIDTH_COMMENTS_scaled;
    int DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled;
    int DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled;


    //Initialization
    void initializeWindowSize();
    RECT calculateMinWindowFrame(HWND hwnd);
    void initializeFontStyles();
    void positionAndResizeControls(int windowWidth, int windowHeight);
    void initializeCtrlMap();
    bool createAndShowWindows();
    void initializeMarkerStyle();
    void initializeListView();
    void moveAndResizeControls();
    void updateTwoButtonsVisibility();
    void updateListViewFrame();
    void updateReplaceInFilesVisibility();
    void setUIElementVisibility();
    void drawGripper();
    void SetWindowTransparency(HWND hwnd, BYTE alpha);
    void adjustWindowSize();
    void updateUseListState(bool isUpdate);

    // Undo
    void undo();
    void redo();
    void addItemsToReplaceList(const std::vector<ReplaceItemData>& items, size_t insertPosition);
    void removeItemsFromReplaceList(const std::vector<size_t>& indicesToRemove);
    void modifyItemInReplaceList(size_t index, const ReplaceItemData& newData);
    bool moveItemsInReplaceList(std::vector<size_t>& indices, Direction direction);
    void sortItemsInReplaceList(const std::vector<size_t>& originalOrder, const std::vector<size_t>& newOrder, const std::map<int, SortDirection>& previousColumnSortOrder, int columnID, SortDirection direction);
    void scrollToIndices(size_t firstIndex, size_t lastIndex);

    //ListView
    HWND CreateHeaderTooltip(HWND hwndParent);
    void AddHeaderTooltip(HWND hwndTT, HWND hwndHeader, int columnIndex, LPCTSTR pszText);
    void createListViewColumns();
    void insertReplaceListItem(const ReplaceItemData& itemData);
    int  getColumnWidth(ColumnID columnID);
    int  calcDynamicColWidth(const ResizableColWidths& widths);
    void updateListViewAndColumns();
    void updateListViewTooltips();
    void handleCopyBack(NMITEMACTIVATE* pnmia);
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
    void clearList();
    std::size_t computeListHash(const std::vector<ReplaceItemData>& list);
    void refreshUIListView();
    void handleColumnVisibilityToggle(UINT menuId);
    ColumnID getColumnIDFromIndex(int columnIndex) const;
    int getColumnIndexFromID(ColumnID columnID) const;
    void updateListViewItem(size_t index);
    void onPathDisplayDoubleClick();

    //Contextmenu Display Columns
    void showColumnVisibilityMenu(HWND hWnd, POINT pt);

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
    void performSearchInList();
    int searchInListData(int startIdx, const std::wstring& findText, const std::wstring& replaceText);
    void handleEditOnDoubleClick(int itemIndex, ColumnID columnID);

    //Replace
    bool handleReplaceAllButton(bool showCompletionMessage = true, const std::filesystem::path* explicitPath = nullptr);
    void handleReplaceButton();
    bool replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount, const size_t itemIndex = SIZE_MAX);
    bool replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos, size_t itemIndex, const SearchContext& context);
    Sci_Position performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    Sci_Position performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    bool preProcessListForReplace(bool highlight);
    SelectionInfo getSelectionInfo(bool isBackward);
    void captureLuaGlobals(lua_State* L);
    std::string escapeForRegex(const std::string& input);
    bool resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex);
    void setLuaVariable(lua_State* L, const std::string& varName, std::string value);
    void updateFilePathCache(const std::filesystem::path* explicitPath = nullptr);
    void setLuaFileVars(LuaVariables& vars);
    bool initLuaState();
    bool compileLuaReplaceCode(const std::string& luaCode);
    static int safeLoadFileSandbox(lua_State* L);

    //Replace in files
    bool handleBrowseDirectoryButton();
    bool selectDirectoryDialog(HWND owner, std::wstring& outPath);
    void handleReplaceInFiles();
    bool convertUtf8ToOriginal(const std::string& utf8_input, const EncodingInfo& original_enc_info, const std::string& original_buf_with_bom, std::string& final_output_with_bom);
    bool convertBufferToUtf8(const std::string& original_buf, const EncodingInfo& enc_info, std::string& utf8_output);
    EncodingInfo detectEncoding(const std::string& buffer);

    //DebugWindow
    int ShowDebugWindow(const std::string& message);
    static LRESULT CALLBACK DebugWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static void CopyListViewToClipboard(HWND hListView);
    static void CloseDebugWindow();

    //Find All
    std::wstring sanitizeSearchPattern(const std::wstring& raw);
    void trimHitToFirstLine(const std::function<LRESULT(UINT, WPARAM, LPARAM)>& sciSend, ResultDock::Hit& h);
    void handleFindAllButton();

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
    void selectListItem(size_t matchIndex);

    //Mark
    void handleMarkMatchesButton();
    int markString(const SearchContext& context);
    void highlightTextRange(LRESULT pos, LRESULT len, const std::string& findTextUtf8);
    int generateColorValue(const std::string& str);
    void handleClearTextMarksButton();
    void handleCopyMarkedTextToClipboardButton();
    void copyTextToClipboard(const std::wstring& text, int textCount);

    //CSV
    void handleCopyColumnsToClipboard();
    bool confirmColumnDeletion();
    void handleDeleteColumns();

    //CSV Sort
    std::vector<CombinedColumns> extractColumnData(SIZE_T startLine, SIZE_T lineCount);
    void sortRowsByColumn(SortDirection sortDirection);
    void reorderLinesInScintilla(const std::vector<size_t>& sortedIndex);
    void restoreOriginalLineOrder(const std::vector<size_t>& originalOrder);
    void extractLineContent(size_t idx, std::string& content, const std::string& lineBreak);
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
    void highlightColumnsInLine(LRESULT line);
    void handleClearColumnMarks();
    std::wstring addLineAndColumnMessage(LRESULT pos);
    void updateDelimitersInDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType);
    void processLogForDelimiters();
    void handleDelimiterPositions(DelimiterOperation operation);
    void handleClearDelimiterState();

    //Utilities
    std::string convertAndExtendW(const std::wstring& input, bool extended, UINT cp /*docCp or CP_UTF8*/) const;
    std::string convertAndExtendW(const std::wstring& input, bool extended);
    static void addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems = 100);
    std::wstring getTextFromDialogItem(HWND hwnd, int itemID);
    void setSelections(bool select, bool onlySelected = false);
    void updateHeaderSelection();
    void updateHeaderSortDirection();
    void showStatusMessage(const std::wstring& messageText, MessageStatus status, bool isNotFound = false);
    void updateThemeAndColors();
    std::wstring getShortenedFilePath(const std::wstring& path, int maxLength, HDC hDC = nullptr);
    void displayResultCentered(size_t posStart, size_t posEnd, bool isDownwards);
    std::wstring getSelectedText();
    LRESULT getEOLLengthForLine(LRESULT line);
    std::string getEOLStyle();
    sptr_t send(unsigned int iMessage, uptr_t wParam = 0, sptr_t lParam = 0, bool useDirect = true) const;
    bool normalizeAndValidateNumber(std::string& str);
    std::vector<WCHAR> createFilterString(const std::vector<std::pair<std::wstring, std::wstring>>& filters);
    int getCharacterWidth(int elementID, const wchar_t* character);
    int getFontHeight(HWND hwnd, HFONT hFont);
    std::vector<int> MultiReplace::parseNumberRanges(const std::wstring& input, const std::wstring& errorMessage);
    UINT getCurrentDocCodePage();

    //FileOperations
    std::wstring promptSaveListToCsv();
    std::wstring openFileDialog(bool saveFile, const std::vector<std::pair<std::wstring, std::wstring>>& filters, const WCHAR* title, DWORD flags, const std::wstring& fileExtension, const std::wstring& defaultFilePath);
    bool saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list);
    void checkForFileChangesAtStartup();
    std::wstring escapeCsvValue(const std::wstring& value);
    std::wstring unescapeCsvValue(const std::wstring& value);
    std::vector<std::wstring> parseCsvLine(const std::wstring& line);

    //Export
    void exportToBashScript(const std::wstring& fileName);
    std::string escapeSpecialChars(const std::string& input, bool extended);
    void handleEscapeSequence(const std::regex& regex, const std::string& input, std::string& output, std::function<char(const std::string&)> converter);
    std::string translateEscapes(const std::string& input);
    std::string replaceNewline(const std::string& input, ReplaceMode mode);

    //INI
    std::pair<std::wstring, std::wstring> generateConfigFilePaths();
    void saveSettingsToIni(const std::wstring& iniFilePath);
    void saveSettings();
    int checkForUnsavedChanges();
    void loadSettingsFromIni();
    void loadSettings();
    void loadUIConfigFromIni();
    void setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text);
    bool parseIniFile(const std::wstring& iniFilePath);
    std::wstring readStringFromIniCache(const std::wstring& section, const std::wstring& key, const std::wstring& defaultValue);
    bool readBoolFromIniCache(const std::wstring& section, const std::wstring& key, bool defaultValue);
    int readIntFromIniCache(const std::wstring& section, const std::wstring& key, int defaultValue);
    BYTE readByteFromIniCache(const std::wstring& section, const std::wstring& key, BYTE defaultValue);
    float readFloatFromIniCache(const std::wstring& section, const std::wstring& key, float defaultValue);
    std::size_t readSizeTFromIniCache(const std::wstring& section, const std::wstring& key, std::size_t defaultValue);

    // Language
    void loadLanguage();
    void loadLanguageFromIni(const std::wstring& iniFilePath, const std::wstring& languageCode);
    std::wstring getLanguageFromNativeLangXML();
    
    LPCWSTR getLangStrLPCWSTR(const std::wstring& id);
    LPWSTR getLangStrLPWSTR(const std::wstring& id);

    // Debug DPI Information
    void showDPIAndFontInfo();

};

extern std::unordered_map<std::wstring, std::wstring> languageMap;

extern MultiReplace _MultiReplace;

#endif // MULTI_REPLACE_H