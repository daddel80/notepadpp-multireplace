// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#ifndef MULTI_REPLACE_H
#define MULTI_REPLACE_H

#include "StaticDialog/StaticDialog.h"
#include "StaticDialog/resource.h"
#include "PluginInterface.h"

#include <string>
#include <vector>
#include <map>
#include <functional>
#include <regex>
#include <algorithm>
#include <unordered_map>
#include <set>
#include <commctrl.h>
#include <lua.hpp>

extern NppData nppData;

enum class DelimiterOperation { LoadAll, Update };
enum class Direction { Up, Down };

struct ReplaceItemData
{
    bool isSelected = true;
    std::wstring findText;
    std::wstring replaceText;
    bool wholeWord = false;
    bool matchCase = false;
    bool useVariables = false;
    bool extended = false;
    bool regex = false;

    bool operator==(const ReplaceItemData& rhs) const {
        return
            isSelected == rhs.isSelected &&
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

struct ControlInfo
{
    int x, y, cx, cy;
    LPCWSTR className;
    LPCWSTR windowName;
    DWORD style;
    LPCWSTR tooltipText;
};

struct SearchResult {
    LRESULT pos = -1;
    LRESULT length = 0;
    std::string foundText = "";
};

struct SelectionInfo {
    std::string text;
    Sci_Position startPos;
    Sci_Position length;
};

struct SelectionRange {
    LRESULT start = 0;
    LRESULT end = 0;
};

struct ColumnDelimiterData {
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

struct DelimiterPosition {
    LRESULT position;
};

struct LineInfo {
    std::vector<DelimiterPosition> positions;
    LRESULT startPosition = 0;
    LRESULT endPosition = 0;
};

struct ColumnInfo {
    LRESULT totalLines;
    LRESULT startLine;
    SIZE_T startColumnIndex;
};

struct LuaVariables {
    int CNT =  0;
    int LINE = 0;
    int LPOS = 0;
    int LCNT = 0;
    int APOS = 0;
    int COL =  1;
    std::string MATCH;
};

// Exceptions
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
        _hFont(nullptr),
        _hStatusMessage(nullptr),
        _statusMessageColor(RGB(0, 0, 0))
    {
        setInstance(this);
    };

    static MultiReplace* instance; // Static instance of the class

    static void setInstance(MultiReplace* inst) {
        instance = inst;
    }

    virtual void display(bool toShow = true) const override {
        StaticDialog::display(toShow);
    };

    void setParent(HWND parent2set) {
        _hParent = parent2set;
    };

    static HWND getScintillaHandle() {
        return s_hScintilla;
    }

    static HWND getDialogHandle() {
        return s_hDlg;
    }

    static bool isWindowOpen;
    static bool textModified;
    static bool documentSwitched;
    static int scannedDelimiterBufferID;
    static bool isLoggingEnabled;
    static bool isCaretPositionEnabled;
    static bool isLuaErrorDialogEnabled;

    // Static methods for Event Handling
    static void onSelectionChanged();
    static void onTextChanged();
    static void onDocumentSwitched();
    static void processLog();
    static void processTextChange(SCNotification* notifyCode);
    static void onCaretPositionChanged();

    enum class ChangeType { Insert, Delete, Modify };

    struct LogEntry {
        ChangeType changeType;
        Sci_Position lineNumber;
    };

    static std::vector<LogEntry> logChanges;


protected:
    virtual INT_PTR CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam) override;

private:
    static constexpr int MAX_TEXT_LENGTH = 4096; // Maximum Textlength for Find and Replace String
    static constexpr const TCHAR* FONT_NAME = TEXT("MS Shell Dlg");
    static constexpr int FONT_SIZE = 16;
    static constexpr long MARKER_COLOR = 0x007F00; // Color for non-list Marker
    static constexpr LRESULT PROGRESS_THRESHOLD = 50000; // Will show progress bar if total exceeds defined threshold
    bool isReplaceOnceInList = false;            // When set, replacement stops after the first match in list and the next list entry gets activated.

    // Static variables related to GUI 
    static HWND s_hScintilla;
    static HWND s_hDlg;
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
    HFONT _hFont;
    COLORREF _statusMessageColor;

    // Style-related variables and constants
    /*
       Available styles (self-tested):
       { 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 }
       Note: Gaps in the list are intentional.

       Styles 0 - 7 are reserved for syntax style.
       Styles 21 - 29, 31 are reserved by N++ (see SciLexer.h).
    */
    std::vector<int> textStyles = { 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 30, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 };
    std::vector<int> hColumnStyles = { STYLE1, STYLE2, STYLE3, STYLE4, STYLE5, STYLE6, STYLE7, STYLE8, STYLE9, STYLE10 };
    std::vector<long> columnColors = { 0xFFE0E0, 0xC0E0FF, 0x80FF80, 0xFFE0FF,  0xB0E0E0, 0xFFFF80, 0xE0C0C0, 0x80FFFF, 0xFFB0FF, 0xC0FFC0 };

    // Data-related variables 
    size_t markedStringsCount = 0;
    bool allSelected = true;
    std::unordered_map<long, int> colorToStyleMap;
    int lastColumn = -1;
    bool ascending = true;
    ColumnDelimiterData columnDelimiterData;
    LRESULT eolLength = -1; // Stores the length of the EOL character sequence
    std::vector<ReplaceItemData> replaceListData;
    std::vector<LineInfo> lineDelimiterPositions;
    bool isColumnHighlighted = false;
    std::map<int, bool> stateSnapshot; // stores the state of the Elements

    // Debugging and logging related 
    std::string messageBoxContent;  // just for temporary debugging usage

    // Scintilla related 
    SciFnDirect pSciMsg = nullptr;
    sptr_t pSciWndData = 0;

    // GUI control-related constants
    const std::vector<int> selectionRadioDisabledButtons = {
        IDC_FIND_BUTTON, IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON, IDC_REPLACE_BUTTON
    };
    const std::vector<int> columnRadioDependentElements = {
        IDC_COLUMN_NUM_EDIT, IDC_DELIMITER_EDIT, IDC_QUOTECHAR_EDIT, IDC_COLUMN_HIGHLIGHT_BUTTON
    };

    //Initialization
    void initializeWindowSize();
    RECT calculateMinWindowFrame(HWND hwnd);
    void positionAndResizeControls(int windowWidth, int windowHeight);
    void initializeCtrlMap();
    bool createAndShowWindows();
    void setupScintilla();
    void initializePluginStyle();
    void initializeListView();
    void moveAndResizeControls();
    void updateButtonVisibilityBasedOnMode();

    //ListView
    void createListViewColumns(HWND listView);
    void insertReplaceListItem(const ReplaceItemData& itemData);
    void updateListViewAndColumns(HWND listView, LPARAM lParam);
    void handleCopyBack(NMITEMACTIVATE* pnmia);
    void shiftListItem(HWND listView, const Direction& direction);
    void handleDeletion(NMITEMACTIVATE* pnmia);
    void deleteSelectedLines(HWND listView);
    void sortReplaceListData(int column);
    std::vector<ReplaceItemData> getSelectedRows();
    void selectRows(const std::vector<ReplaceItemData>& rowsToSelect);
    void handleCopyToListButton();

    //Replace
    void handleReplaceAllButton();
    void handleReplaceButton();
    int replaceAll(const ReplaceItemData& itemData);
    bool replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos);
    Sci_Position performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    Sci_Position performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    SelectionInfo getSelectionInfo();
    bool resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex);
    void setLuaVariable(lua_State* L, const std::string& varName, std::string value);

    //Find
    void handleFindNextButton();
    void handleFindPrevButton();
    SearchResult performSingleSearch(const std::string& findTextUtf8, int searchFlags, bool selectMatch, SelectionRange range);
    SearchResult performSearchForward(const std::string& findTextUtf8, int searchFlags, bool selectMatch, LRESULT start);
    SearchResult performSearchBackward(const std::string& findTextUtf8, int searchFlags, LRESULT start);
    SearchResult performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos);
    SearchResult performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos);

    //Mark
    void handleMarkMatchesButton();
    int markString(const std::string& findTextUtf8, int searchFlags);
    void highlightTextRange(LRESULT pos, LRESULT len, const std::string& findTextUtf8);
    long generateColorValue(const std::string& str);
    void handleClearTextMarksButton();
    void handleCopyMarkedTextToClipboardButton();

    //Scope
    bool parseColumnAndDelimiterData();
    void findAllDelimitersInDocument();
    void findDelimitersInLine(LRESULT line);
    ColumnInfo getColumnInfo(LRESULT startPosition);
    void initializeColumnStyles();
    void handleHighlightColumnsInDocument();
    void highlightColumnsInLine(LRESULT line);
    void handleClearColumnMarks();
    std::wstring addLineAndColumnMessage(LRESULT pos);
    void updateDelimitersInDocument(SIZE_T lineNumber, ChangeType changeType);
    void processLogForDelimiters();
    void handleDelimiterPositions(DelimiterOperation operation);
    void handleClearDelimiterState();
    //void displayLogChangesInMessageBox();

    //Utilities
    int convertExtendedToString(const std::string& query, std::string& result);
    std::string convertAndExtend(const std::wstring& input, bool extended);
    static void addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems = 10);
    std::wstring getTextFromDialogItem(HWND hwnd, int itemID);
    void setSelections(bool select, bool onlySelected = false);
    void updateHeader();
    void showStatusMessage(const std::wstring& messageText, COLORREF color);
    void displayResultCentered(size_t posStart, size_t posEnd, bool isDownwards);
    std::wstring getSelectedText();
    LRESULT updateEOLLength();
    void setElementsState(const std::vector<int>& elements, bool enable);
    sptr_t send(unsigned int iMessage, uptr_t wParam = 0, sptr_t lParam = 0, bool useDirect = true);
    bool MultiReplace::normalizeAndValidateNumber(std::string& str);

    //StringHandling
    std::wstring stringToWString(const std::string& encodedInput) const;
    std::string wstringToString(const std::wstring& input) const;
    std::wstring MultiReplace::utf8ToWString(const char* cstr) const;
    std::string utf8ToCodepage(const std::string& utf8Str, int codepage) const;

    //FileOperations
    std::wstring openFileDialog(bool saveFile, const WCHAR* filter, const WCHAR* title, DWORD flags, const std::wstring& fileExtension);
    bool saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list);
    void loadListFromCsv(const std::wstring& filePath);
    std::wstring escapeCsvValue(const std::wstring& value);
    std::wstring unescapeCsvValue(const std::wstring& value);

    //Export
    void exportToBashScript(const std::wstring& fileName);
    std::string escapeSpecialChars(const std::string& input, bool extended);
    void handleEscapeSequence(const std::regex& regex, const std::string& input, std::string& output, std::function<char(const std::string&)> converter);
    std::string translateEscapes(const std::string& input);

    //INI
    void saveSettingsToIni(const std::wstring& iniFilePath);
    void saveSettings();
    void loadSettingsFromIni(const std::wstring& iniFilePath);
    void loadSettings();
    std::wstring readStringFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, const std::wstring& defaultValue);
    bool readBoolFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, bool defaultValue);
    int readIntFromIniFile(const std::wstring& iniFilePath, const std::wstring& section, const std::wstring& key, int defaultValue);
    void setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text);

};

extern MultiReplace _MultiReplace;

#endif // MULTI_REPLACE_H
