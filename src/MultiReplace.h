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

#include "DockingFeature\DockingDlgInterface.h"
#include "DockingFeature\resource.h"
#include <string>
#include <vector>
#include <map>
#include <commctrl.h>
#include "PluginInterface.h"
#include <functional>
#include <regex>
#include <algorithm>
#include <unordered_map>


extern NppData nppData;

struct ReplaceItemData
{
    bool isSelected = true;
    std::wstring findText;
    std::wstring replaceText;
    bool wholeWord = false;
    bool matchCase = false;
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
    std::string foundText;
};

struct SelectionInfo {
    std::string text;
    Sci_Position startPos;
    Sci_Position length;
};

enum class Direction { Up, Down };

typedef std::basic_string<TCHAR> generic_string;

class MultiReplace : public DockingDlgInterface
{
public:
    MultiReplace() :
        hInstance(NULL),
        _hScintilla(0),
        _hClearMarksButton(nullptr),
        _hCopyBackIcon(nullptr),
        _hCopyMarkedTextButton(nullptr),
        _hInListCheckbox(nullptr),
        _hMarkMatchesButton(nullptr),
        _hReplaceAllButton(nullptr),
        DockingDlgInterface(IDD_REPLACE_DIALOG),
        _replaceListView(NULL),
        _hDeleteIcon(NULL),
        _hEnabledIcon(NULL),
        copyBackIconIndex(-1),
        deleteIconIndex(-1),
        enabledIconIndex(-1),
        _himl(NULL),
        hFont(nullptr),
        _hStatusMessage(nullptr),
        _statusMessageColor(RGB(0, 0, 0))
    {};

    virtual void display(bool toShow = true) const {
        DockingDlgInterface::display(toShow);
    };

    void setParent(HWND parent2set) {
        _hParent = parent2set;
    };

	bool isFloating() const {
		return _isFloating;
	};

protected:
    virtual INT_PTR CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
    static const int RESIZE_TIMER_ID = 1;
    HINSTANCE hInstance;
    HWND _hScintilla;
    HWND _replaceListView;

    HWND _hInListCheckbox;
    HWND _hReplaceAllButton;
    HWND _hMarkMatchesButton;
    HWND _hClearMarksButton;
    HWND _hCopyMarkedTextButton;

    HICON _hCopyBackIcon;
    HICON _hDeleteIcon;
    HICON _hEnabledIcon;
    HWND _hStatusMessage;
    COLORREF _statusMessageColor;
    HFONT hFont;
    int copyBackIconIndex;
    int deleteIconIndex;
    int enabledIconIndex;
    static constexpr const TCHAR* FONT_NAME = TEXT("MS Shell Dlg");
    static constexpr int FONT_SIZE = 16;
    size_t markedStringsCount = 0;
    int lastClickedComboBoxId = 0;    // for Combobox workaround
    static const int MAX_TEXT_LENGTH = 4096; // Set maximum Textlength for Find and Replace String
    bool allSelected = true;
    std::unordered_map<long, int> colorToStyleMap;
    static const long MARKER_COLOR = 0x007F00; // Color for non-list Marker
    int lastColumn = -1;
    bool ascending = true;
    
    /*
       Available styles (self-tested):
       { 0, 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
         22, 23, 24, 25, 28, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 }
       Note: Gaps in the list are intentional. 

       Styles 0 - 7 are reserved for syntax style.
       Styles 21 - 29, 31 are reserved bei N++ (see SciLexer.h).
    */
    std::vector<int> validStyles = { 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 
                                    30, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43 };

    HIMAGELIST _himl;
    std::vector<ReplaceItemData> replaceListData;
    static std::map<int, ControlInfo> ctrlMap;

    //Initialization
    void positionAndResizeControls(int windowWidth, int windowHeight);
    void initializeCtrlMap();
    bool createAndShowWindows();
    void initializeScintilla();
    void initializePluginStyle();
    void initializeListView();
    void moveAndResizeControls();
    void updateButtonVisibilityBasedOnMode();
    void updateUIVisibility();

    //ListView
    void createListViewColumns(HWND listView);
    void insertReplaceListItem(const ReplaceItemData& itemData);
    void updateListViewAndColumns(HWND listView, LPARAM lParam);
    void handleSelection(NMITEMACTIVATE* pnmia);
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
    int replaceString(const std::wstring& findText, const std::wstring& replaceText, bool wholeWord, bool matchCase, bool regex, bool extended);
    Sci_Position performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length);
    SelectionInfo getSelectionInfo();

    //Find
    void handleFindNextButton();
    void handleFindPrevButton();
    SearchResult performSearchForward(const std::string& findTextUtf8, int searchFlags, LRESULT start, bool selectMatch);
    SearchResult performSearchBackward(const std::string& findTextUtf8, int searchFlags, LRESULT start);
    SearchResult performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos);
    SearchResult performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos);

    //Mark
    void handleMarkMatchesButton();
    int markString(const std::wstring& findText, bool wholeWord, bool matchCase, bool regex, bool extended);
    void highlightTextRange(LRESULT pos, LRESULT len, const std::string& findTextUtf8);
    long generateColorValue(const std::string& str);
    void handleClearAllMarksButton();
    void handleCopyMarkedTextToClipboardButton();

    //Utilities
    int convertExtendedToString(const std::string& query, std::string& result);
    std::string convertAndExtend(const std::wstring& input, bool extended);
    static void addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems = 10);
    std::wstring getTextFromDialogItem(HWND hwnd, int itemID);
    void setSelections(bool select, bool onlySelected = false);
    void updateHeader();
    void showStatusMessage(size_t count, const wchar_t* messageFormat, COLORREF color);

    //StringHandling
    std::wstring stringToWString(const std::string& encodedInput);
    std::string wstringToString(const std::wstring& input);

    //FileOperations
    std::wstring openFileDialog(bool saveFile, const WCHAR* filter, const WCHAR* title, DWORD flags, const std::wstring& fileExtension);
    bool saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    void saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list);
    bool loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list);
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
