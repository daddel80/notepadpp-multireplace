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

#define NOMINMAX
#define WM_UPDATE_FOCUS (WM_APP + 2)

// Own header
#include "MultiReplacePanel.h"

// Project headers
#include "BatchUIGuard.h"
#include "ColumnTabs.h"
#include "ConfigManager.h"
#include "DPIManager.h"
#include "Encoding.h"
#include "HiddenSciGuard.h"
#include "LanguageManager.h"
#include "language_mapping.h"
#include "luaEmbedded.h"
#include "menuCmdID.h"
#include "MultiReplaceConfigDialog.h"
#include "Notepad_plus_msgs.h"
#include "NppStyleKit.h"
#include "NumericToken.h"
#include "PluginDefinition.h"
#include "ResultDock.h"
#include "Scintilla.h"
#include "StaticDialog/StaticDialog.h"
#include "StringUtils.h"
#include "UndoRedoManager.h"

// Standard library
#include <algorithm>
#include <bitset>
#include <cctype>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <map>
#include <numeric>
#include <regex>
#include <set>
#include <sstream>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

// Third-party
#include <lua.hpp>

// Windows
#include <sdkddkver.h>
#include <windows.h>
#include <Commctrl.h>
#include <uxtheme.h>
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "comctl32.lib")

static LanguageManager& LM = LanguageManager::instance();
static ConfigManager& CFG = ConfigManager::instance();
static UndoRedoManager& URM = UndoRedoManager::instance();
namespace SU = StringUtils;
extern MultiReplaceConfigDialog _MultiReplaceConfig;

// Case-insensitive UTF-8 path comparison (Windows paths are case-insensitive)
static inline bool pathsEqualUtf8(const std::string& a, const std::string& b) {
    return _stricmp(a.c_str(), b.c_str()) == 0;
}

// Pointer-sized BufferID
using BufferId = UINT_PTR;

// Async leave-clean state machine
static BufferId g_prevBufId = 0;    // last active buffer (source)
static BufferId g_returnBufId = 0;    // where we must end up
static BufferId g_pendingCleanId = 0;    // buffer we still need to clean
static bool     g_cleanInProgress = false;

// O(1) gate: which buffers currently have flow-pads
static std::unordered_set<BufferId> g_padBufs;

#pragma region Initialization

void MultiReplace::initializeWindowSize()
{
    loadUIConfigFromIni();

    HMONITOR hMonitor = MonitorFromRect(&windowRect, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo = { sizeof(monitorInfo) };

    if (GetMonitorInfo(hMonitor, &monitorInfo))
    {
        int monitorLeft = monitorInfo.rcWork.left;
        int monitorTop = monitorInfo.rcWork.top;
        int monitorRight = monitorInfo.rcWork.right;
        int monitorBottom = monitorInfo.rcWork.bottom;

        int windowWidth = windowRect.right - windowRect.left;
        int windowHeight = windowRect.bottom - windowRect.top;

        const int visibilityMargin = 10;

        bool isCompletelyOffScreen =
            (windowRect.right <= monitorLeft || windowRect.left >= monitorRight ||
                windowRect.bottom <= monitorTop || windowRect.top >= monitorBottom);

        if (isCompletelyOffScreen)
        {
            windowRect.left = monitorLeft + visibilityMargin;
            windowRect.top = monitorTop + visibilityMargin;
        }
        else
        {
            if (windowRect.left < monitorLeft + visibilityMargin)
                windowRect.left = monitorLeft + visibilityMargin;
            if (windowRect.top < monitorTop + visibilityMargin)
                windowRect.top = monitorTop + visibilityMargin;
            if (windowRect.left + windowWidth > monitorRight - visibilityMargin)
                windowRect.left = monitorRight - windowWidth - visibilityMargin;
            if (windowRect.top + windowHeight > monitorBottom - visibilityMargin)
                windowRect.top = monitorBottom - windowHeight - visibilityMargin;
        }

        windowRect.right = windowRect.left + windowWidth;
        windowRect.bottom = windowRect.top + windowHeight;
    }

    SetWindowPos(
        _hSelf,
        nullptr,
        windowRect.left,
        windowRect.top,
        windowRect.right - windowRect.left,
        windowRect.bottom - windowRect.top,
        SWP_NOZORDER
    );
}

void MultiReplace::createFonts() {
    if (!dpiMgr) return;
    cleanupFonts(); // Clean up before recreation

    // Helper lambda with fallback
    auto create = [&](int height, int weight, const wchar_t* fontName) -> HFONT {
        HFONT hf = ::CreateFont(
            dpiMgr->scaleY(height), 0, 0, 0, weight, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName
        );
        return hf ? hf : reinterpret_cast<HFONT>(::GetStockObject(DEFAULT_GUI_FONT));
        };

    // Populate Registry
    using FR = FontRole;
    _fontHandles[(size_t)FR::Standard] = create(13, FW_NORMAL, L"MS Shell Dlg 2");
    _fontHandles[(size_t)FR::Normal1] = create(14, FW_NORMAL, L"MS Shell Dlg 2");
    _fontHandles[(size_t)FR::Normal2] = create(12, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Normal3] = create(14, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Normal4] = create(16, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Normal5] = create(18, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Normal6] = create(22, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Normal7] = create(26, FW_NORMAL, L"Courier New");
    _fontHandles[(size_t)FR::Bold1] = create(22, FW_BOLD, L"Courier New");
    _fontHandles[(size_t)FR::Bold2] = create(12, FW_BOLD, L"MS Shell Dlg 2");

    // Calculate metrics (Using Screen DC to allow running before window creation)
    HDC hdc = GetDC(NULL);
    if (hdc) {
        auto measure = [&](const wchar_t* text) -> int {
            SIZE size;
            HGDIOBJ oldFont = SelectObject(hdc, font(FontRole::Standard));
            GetTextExtentPoint32W(hdc, text, 1, &size);
            SelectObject(hdc, oldFont);
            return size.cx;
            };

        checkMarkWidth_scaled = measure(L"\u2714") + 15;
        crossWidth_scaled = measure(L"\u2716") + 15;
        boxWidth_scaled = measure(L"\u2610") + 15;

        ReleaseDC(nullptr, hdc);
    }
    deleteButtonColumnWidth = crossWidth_scaled;
}

void MultiReplace::cleanupFonts() {
    for (auto& hFont : _fontHandles) {
        if (hFont) {
            DeleteObject(hFont);
            hFont = nullptr;
        }
    }
}

void MultiReplace::applyFonts() {
    for (const auto& pair : ctrlMap) {
        HWND hCtrl = GetDlgItem(_hSelf, pair.first);
        if (hCtrl) {
            FontRole role = pair.second.fontRole;
            HFONT hFont = _fontHandles[static_cast<size_t>(role)];
            if (hFont) {
                SendMessage(hCtrl, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
            }
        }
    }
}

RECT MultiReplace::calculateMinWindowFrame(HWND hwnd) {
    // Use local variables to avoid modifying windowRect
    RECT tempWindowRect;
    GetWindowRect(hwnd, &tempWindowRect);

    // Measure the window's borders and title bar
    RECT clientRect;
    GetClientRect(hwnd, &clientRect);

    int borderWidth = ((tempWindowRect.right - tempWindowRect.left) - (clientRect.right - clientRect.left)) / 2;
    int titleBarHeight = (tempWindowRect.bottom - tempWindowRect.top) - (clientRect.bottom - clientRect.top) - borderWidth;

    // Base minimum content height (list on/off)
    int minContentHeight = useListEnabled ? MIN_HEIGHT_scaled : SHRUNK_HEIGHT_scaled;

    // Add extra room if “Replace in Files” panel is visible AND we're not in Two-Buttons-Mode
    bool twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;
    int panelExtra = ((isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode) ? sy(REPLACE_FILES_PANEL_HEIGHT) : 0;

    int minHeight = minContentHeight + panelExtra;
    int minWidth = MIN_WIDTH_scaled;

    // Adjust for window borders and title bar
    minHeight += borderWidth + titleBarHeight;
    minWidth += 2 * borderWidth;

    RECT minSize = { 0, 0, minWidth, minHeight };

    return minSize;
}

void MultiReplace::positionAndResizeControls(int windowWidth, int windowHeight)
{
    // Helper functions sx() and sy() are used to scale values for DPI awareness.
    // sx(value): scales the x-axis value based on DPI settings.
    // sy(value): scales the y-axis value based on DPI settings.

    // DPI Aware System Metrics for Checkboxes and Radio Buttons
    UINT dpi = dpiMgr->getDPIX(); // Get the DPI from DPIManager

    // Calculate appropriate height for checkboxes and radiobuttons
    int checkboxBaseHeight = dpiMgr->getCustomMetricOrFallback(SM_CYMENUCHECK, dpi, 14);
    int radioButtonBaseHeight = dpiMgr->getCustomMetricOrFallback(SM_CYMENUCHECK, dpi, 14);

    // Get the font height from the standard font
    int fontHeight = getFontHeight(_hSelf, font(FontRole::Standard));
    fontHeight = fontHeight + sy(8); // Padding Font

    // Choose the larger value between the font height and the base height
    int checkboxHeight = std::max(checkboxBaseHeight, fontHeight);
    int radioButtonHeight = std::max(radioButtonBaseHeight, fontHeight);

    // Calculate dimensions without scaling
    BOOL twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;
    int filesOffsetY = ((isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode) ? sy(REPLACE_FILES_PANEL_HEIGHT) : 0;
    int buttonX = windowWidth - sx(33 + 128);
    int checkbox2X = buttonX + sx(134);
    int useListButtonX = buttonX + sx(133);
    int swapButtonX = windowWidth - sx(33 + 128 + 26);
    int comboWidth = windowWidth - sx(289);
    int listWidth = windowWidth - sx(207);
    int pathDisplayY = windowHeight - sy(22);
    int searchBarHeight = sy(22);
    int searchBarGap = sy(2);
    int searchBarY = pathDisplayY - searchBarHeight - searchBarGap;
    int listStartY = sy(227) + filesOffsetY;
    int listGap = sy(2);
    int listEndY = _listSearchBarVisible ? (searchBarY - listGap) : (pathDisplayY - listGap);
    int listHeight = std::max(listEndY - listStartY, sy(20));
    int useListButtonY = windowHeight - sy(34);

    // Apply scaling only when assigning to ctrlMap
   // --- STATIC CONTROLS ---

    // Default Font (Standard)
    ctrlMap[IDC_STATIC_FIND] = { sx(11), sy(18), sx(80), sy(19), WC_STATIC, LM.getLPCW(L"panel_find_what"), SS_RIGHT, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_STATIC_REPLACE] = { sx(11), sy(47), sx(80), sy(19), WC_STATIC, LM.getLPCW(L"panel_replace_with"), SS_RIGHT, nullptr, true, FontRole::Standard };

    ctrlMap[IDC_WHOLE_WORD_CHECKBOX] = { sx(16), sy(76), sx(155), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_match_whole_word_only"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_MATCH_CASE_CHECKBOX] = { sx(16), sy(101), sx(155), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_match_case"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_USE_VARIABLES_CHECKBOX] = { sx(16), sy(126), sx(133), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_use_variables"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_USE_VARIABLES_HELP] = { sx(152), sy(126), sx(20), sy(20), WC_BUTTON, LM.getLPCW(L"panel_help"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_WRAP_AROUND_CHECKBOX] = { sx(16), sy(151), sx(155), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_wrap_around"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_REPLACE_AT_MATCHES_CHECKBOX] = { sx(16), sy(176), sx(110), checkboxHeight, WC_BUTTON, LM.getLPCW(L"panel_replace_at_matches"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, true, FontRole::Standard };

    // Note: Replace Hit Edit uses Standard Font implicitly
    ctrlMap[IDC_REPLACE_HIT_EDIT] = { sx(130), sy(176), sx(41), sy(16), WC_EDIT, nullptr, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL,  LM.getLPCW(L"tooltip_replace_at_matches"), true, FontRole::Standard };

    ctrlMap[IDC_SEARCH_MODE_GROUP] = { sx(180), sy(79), sx(173), sy(104), WC_BUTTON, LM.getLPCW(L"panel_search_mode"), BS_GROUPBOX, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_NORMAL_RADIO] = { sx(188), sy(101), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_normal"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_EXTENDED_RADIO] = { sx(188), sy(126), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_extended"), BS_AUTORADIOBUTTON | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_REGEX_RADIO] = { sx(188), sy(150), sx(162), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_regular_expression"), BS_AUTORADIOBUTTON | WS_TABSTOP, nullptr, true, FontRole::Standard };

    ctrlMap[IDC_SCOPE_GROUP] = { sx(367), sy(79), sx(252), sy(125), WC_BUTTON, LM.getLPCW(L"panel_scope"), BS_GROUPBOX, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_ALL_TEXT_RADIO] = { sx(375), sy(101), sx(189), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_all_text"), BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_SELECTION_RADIO] = { sx(375), sy(126), sx(189), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_selection"), BS_AUTORADIOBUTTON | WS_TABSTOP, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_COLUMN_MODE_RADIO] = { sx(375), sy(150), sx(45), radioButtonHeight, WC_BUTTON, LM.getLPCW(L"panel_csv"), BS_AUTORADIOBUTTON | WS_TABSTOP, nullptr, true, FontRole::Standard };

    ctrlMap[IDC_COLUMN_NUM_STATIC] = { sx(412), sy(151), sx(30), sy(20), WC_STATIC, LM.getLPCW(L"panel_cols"), SS_RIGHT, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_COLUMN_NUM_EDIT] = { sx(443), sy(151), sx(41), sy(16), WC_EDIT, nullptr, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_columns"), true, FontRole::Standard };
    ctrlMap[IDC_DELIMITER_STATIC] = { sx(485), sy(151), sx(38), sy(20), WC_STATIC, LM.getLPCW(L"panel_delim"), SS_RIGHT, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_DELIMITER_EDIT] = { sx(524), sy(151), sx(25), sy(16), WC_EDIT, nullptr, ES_LEFT | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_delimiter"), true, FontRole::Standard };
    ctrlMap[IDC_QUOTECHAR_STATIC] = { sx(549), sy(151), sx(37), sy(20), WC_STATIC, LM.getLPCW(L"panel_quote"), SS_RIGHT, nullptr, true, FontRole::Standard };
    ctrlMap[IDC_QUOTECHAR_EDIT] = { sx(587), sy(151), sx(15), sy(16), WC_EDIT, nullptr, ES_CENTER | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, LM.getLPCW(L"tooltip_quote"), true, FontRole::Standard };

    // CSV Tool Buttons - grouped visually: SORT | EDIT | VIEW | ANALYZE
    // Group 1: SORT (tight spacing)
    ctrlMap[IDC_COLUMN_SORT_DESC_BUTTON] = { sx(373), sy(176), sx(30), sy(20), WC_BUTTON, symbolSortDesc, BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_sort_descending"), true, FontRole::Standard };
    ctrlMap[IDC_COLUMN_SORT_ASC_BUTTON] = { sx(404), sy(176), sx(30), sy(20), WC_BUTTON, symbolSortAsc, BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_sort_ascending"), true, FontRole::Standard };
    // Group 2: EDIT (gap before, tight spacing)
    ctrlMap[IDC_COLUMN_DROP_BUTTON] = { sx(441), sy(176), sx(30), sy(20), WC_BUTTON, L"✖", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_drop_columns"), true, FontRole::Normal2 };
    ctrlMap[IDC_COLUMN_COPY_BUTTON] = { sx(472), sy(176), sx(30), sy(20), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_copy_columns"), true, FontRole::Normal3 };
    // Group 3: VIEW (gap before, tight spacing)
    ctrlMap[IDC_COLUMN_HIGHLIGHT_BUTTON] = { sx(509), sy(176), sx(30), sy(20), WC_BUTTON, L"🖍", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_column_highlight"), true, FontRole::Normal2 };
    ctrlMap[IDC_COLUMN_GRIDTABS_BUTTON] = { sx(540), sy(176), sx(30), sy(20), WC_BUTTON, L"⇥", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_column_tabs"), true, FontRole::Normal7 };
    // Group 4: ANALYZE (gap before)
    ctrlMap[IDC_COLUMN_DUPLICATES_BUTTON] = { sx(577), sy(176), sx(30), sy(20), WC_BUTTON, L"☰", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_find_duplicates"), true, FontRole::Standard };

    // --- DYNAMIC CONTROLS ---

    // Edit Controls -> Normal1
    ctrlMap[IDC_FIND_EDIT] = { sx(96), sy(14), comboWidth, sy(160), WC_COMBOBOX, nullptr, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, nullptr, false, FontRole::Normal1 };
    ctrlMap[IDC_REPLACE_EDIT] = { sx(96), sy(44), comboWidth, sy(160), WC_COMBOBOX, nullptr, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, nullptr, false, FontRole::Normal1 };

    // Swap Button -> Bold1
    ctrlMap[IDC_SWAP_BUTTON] = { swapButtonX, sy(26), sx(22), sy(27), WC_BUTTON, L"⇅", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Bold1 };

    ctrlMap[IDC_COPY_TO_LIST_BUTTON] = { buttonX, sy(14), sx(128), sy(52), WC_BUTTON, LM.getLPCW(L"panel_add_into_list"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_REPLACE_ALL_BUTTON] = { buttonX, sy(91), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_replace_all"), BS_SPLITBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_REPLACE_BUTTON] = { buttonX, sy(91), sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_replace"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Replace All Small -> Normal6
    ctrlMap[IDC_REPLACE_ALL_SMALL_BUTTON] = { buttonX + sx(100), sy(91), sx(28), sy(24), WC_BUTTON, L"↻", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_replace_all"), false, FontRole::Normal6 };

    ctrlMap[IDC_2_BUTTONS_MODE] = { checkbox2X, sy(91), sx(20), sy(20), WC_BUTTON, L"", BS_AUTOCHECKBOX | WS_TABSTOP, LM.getLPCW(L"tooltip_2_buttons_mode"), false, FontRole::Standard };

    ctrlMap[IDC_FIND_ALL_BUTTON] = { buttonX, sy(119), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_find_all"), BS_SPLITBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    findNextButtonText = L"▼ " + LM.get(L"panel_find_next");
    ctrlMap[IDC_FIND_NEXT_BUTTON] = ControlInfo{ buttonX + sx(32), sy(119), sx(96), sy(24), WC_BUTTON, findNextButtonText.c_str(), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    ctrlMap[IDC_FIND_PREV_BUTTON] = { buttonX, sy(119), sx(28), sy(24), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_MARK_BUTTON] = { buttonX, sy(147), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_mark_matches"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_MARK_MATCHES_BUTTON] = { buttonX, sy(147), sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_mark_matches_small"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Copy Marked -> Normal4
    ctrlMap[IDC_COPY_MARKED_TEXT_BUTTON] = { buttonX + sx(100), sy(147), sx(28), sy(24), WC_BUTTON, L"⧉", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_copy_marked_text"), false, FontRole::Normal4 };

    ctrlMap[IDC_CLEAR_MARKS_BUTTON] = { buttonX, sy(175), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_clear_all_marks"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Status Message -> Normal1
    ctrlMap[IDC_STATUS_MESSAGE] = { sx(19), sy(205) + filesOffsetY, listWidth - sx(5), sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_ENDELLIPSIS | SS_NOPREFIX | SS_OWNERDRAW, nullptr, false, FontRole::Normal1 };

    ctrlMap[IDC_LOAD_FROM_CSV_BUTTON] = { buttonX, sy(227) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_LOAD_LIST_BUTTON] = { buttonX, sy(227) + filesOffsetY, sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_load_list"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_NEW_LIST_BUTTON] = { buttonX + sx(100), sy(227) + filesOffsetY, sx(28), sy(24), WC_BUTTON, L"➕", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_new_list"), false, FontRole::Standard };
    ctrlMap[IDC_SAVE_TO_CSV_BUTTON] = { buttonX, sy(255) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_save_list"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Save -> Normal3
    ctrlMap[IDC_SAVE_BUTTON] = { buttonX, sy(255) + filesOffsetY, sx(28), sy(24), WC_BUTTON, L"💾", BS_PUSHBUTTON | WS_TABSTOP, LM.getLPCW(L"tooltip_save"), false, FontRole::Normal3 };

    ctrlMap[IDC_SAVE_AS_BUTTON] = { buttonX + sx(32), sy(255) + filesOffsetY, sx(96), sy(24), WC_BUTTON, LM.getLPCW(L"panel_save_as"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_EXPORT_BASH_BUTTON] = { buttonX, sy(283) + filesOffsetY, sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_export_to_bash"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Move Buttons - positioned at right edge of list, vertically fixed at list start
    int moveButtonX = sx(14) + listWidth + sx(4);  // 4px gap to list
    int moveButtonY = sy(227) + filesOffsetY;       // Same Y as list start
    ctrlMap[IDC_UP_BUTTON] = { moveButtonX, moveButtonY, sx(20), sy(20), WC_BUTTON, L"▲", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, LM.getLPCW(L"tooltip_move_up"), false, FontRole::Standard };
    ctrlMap[IDC_DOWN_BUTTON] = { moveButtonX, moveButtonY + sy(28), sx(20), sy(20), WC_BUTTON, L"▼", BS_PUSHBUTTON | WS_TABSTOP | BS_CENTER, LM.getLPCW(L"tooltip_move_down"), false, FontRole::Standard };
    ctrlMap[IDC_REPLACE_LIST] = { sx(14), sy(227) + filesOffsetY, listWidth, listHeight, WC_LISTVIEW, nullptr, LVS_REPORT | LVS_OWNERDATA | WS_BORDER | WS_TABSTOP | WS_VSCROLL | LVS_SHOWSELALWAYS, nullptr, false, FontRole::Standard };

    // List Search Bar (between list and path display)
    int searchComboWidth = listWidth - sx(24 + 24 + 4);  // 2 buttons + spacing
    ctrlMap[IDC_LIST_SEARCH_COMBO] = { sx(14), searchBarY, searchComboWidth, sy(100), WC_COMBOBOX, nullptr, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, nullptr, false, FontRole::Normal1 };
    ctrlMap[IDC_LIST_SEARCH_BUTTON] = { sx(14) + searchComboWidth + sx(2), searchBarY, sx(24), sy(22), WC_BUTTON, L"▶", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_LIST_SEARCH_CLOSE] = { sx(14) + searchComboWidth + sx(28), searchBarY, sx(24), sy(22), WC_BUTTON, L"×", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };

    // Path/Stats -> Normal1
    ctrlMap[IDC_PATH_DISPLAY] = { sx(14), pathDisplayY, listWidth, sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_NOTIFY, nullptr, false, FontRole::Normal1 };
    ctrlMap[IDC_STATS_DISPLAY] = { sx(14) + listWidth, pathDisplayY, 0, sy(19), WC_STATIC, L"", WS_VISIBLE | SS_LEFT | SS_NOTIFY, nullptr, false, FontRole::Normal1 };

    // Use List -> Normal5
    ctrlMap[IDC_USE_LIST_BUTTON] = { useListButtonX, useListButtonY , sx(22), sy(22), WC_BUTTON, useListEnabled ? L"˄" : L"˅", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Normal5 };

    ctrlMap[IDC_CANCEL_REPLACE_BUTTON] = { buttonX, sy(260), sx(128), sy(24), WC_BUTTON, LM.getLPCW(L"panel_cancel_replace"), BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_FILE_OPS_GROUP] = { sx(14), sy(210), listWidth, sy(80), WC_BUTTON,LM.getLPCW(L"panel_replace_in_files"), BS_GROUPBOX, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_FILTER_STATIC] = { sx(15),  sy(230), sx(75),  sy(19), WC_STATIC, LM.getLPCW(L"panel_filter"), SS_RIGHT, nullptr, false, FontRole::Standard };

    // Filter/Dir Edits -> Normal1
    ctrlMap[IDC_FILTER_EDIT] = { sx(96),  sy(230), comboWidth - sx(170),  sy(160), WC_COMBOBOX, nullptr, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, nullptr, false, FontRole::Normal1 };
    ctrlMap[IDC_FILTER_HELP] = { sx(96) + comboWidth - sx(170) + sx(5), sy(228), sx(20), sy(20), WC_STATIC, L"(?)", SS_CENTER | SS_OWNERDRAW | SS_NOTIFY, LM.getLPCW(L"tooltip_filter_help"), false, FontRole::Standard };
    ctrlMap[IDC_DIR_STATIC] = { sx(15),  sy(257), sx(75),  sy(19), WC_STATIC, LM.getLPCW(L"panel_directory"), SS_RIGHT, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_DIR_EDIT] = { sx(96),  sy(257), comboWidth - sx(170),  sy(160), WC_COMBOBOX, nullptr, CBS_DROPDOWN | CBS_AUTOHSCROLL | WS_VSCROLL | WS_TABSTOP, nullptr, false, FontRole::Normal1 };

    ctrlMap[IDC_BROWSE_DIR_BUTTON] = { comboWidth - sx(70), sy(257), sx(20),  sy(20), WC_BUTTON, L"...", BS_PUSHBUTTON | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_SUBFOLDERS_CHECKBOX] = { comboWidth - sx(27), sy(230), sx(120), sy(13), WC_BUTTON, LM.getLPCW(L"panel_in_subfolders"), BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, false, FontRole::Standard };
    ctrlMap[IDC_HIDDENFILES_CHECKBOX] = { comboWidth - sx(27), sy(257), sx(120), sy(13), WC_BUTTON, LM.getLPCW(L"panel_in_hidden_folders"),BS_AUTOCHECKBOX | WS_TABSTOP, nullptr, false, FontRole::Standard };
}

HFONT MultiReplace::font(FontRole role) const {
    return _fontHandles[static_cast<size_t>(role)];
}

void MultiReplace::initializeCtrlMap() {

    hInstance = (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE);
    s_hDlg = _hSelf;

    // Get the client rectangle
    RECT rcClient;
    GetClientRect(_hSelf, &rcClient);
    // Extract width and height from the RECT
    int windowWidth = rcClient.right - rcClient.left;
    int windowHeight = rcClient.bottom - rcClient.top;

    // Define Position for all Elements
    positionAndResizeControls(windowWidth, windowHeight);

    // Now iterate over the controls and create each one.
    if (!createAndShowWindows()) {
        return;
    }

    // Show or hide the "Export to Bash" button according to the ExportToBash option
    ShowWindow(GetDlgItem(_hSelf, IDC_EXPORT_BASH_BUTTON), exportToBashEnabled ? SW_SHOW : SW_HIDE);

    // immediately hide or show the files sub-panel
    updateFilesPanel();

    // Initialize the tooltip for the "Use List" button with dynamic text
    updateUseListState(false);

    // Hide path and stats display if list is not enabled
    if (!useListEnabled) {
        ShowWindow(GetDlgItem(_hSelf, IDC_PATH_DISPLAY), SW_HIDE);
        ShowWindow(GetDlgItem(_hSelf, IDC_STATS_DISPLAY), SW_HIDE);
    }

    // Limit the input for IDC_QUOTECHAR_EDIT to one character
    SendMessage(GetDlgItem(_hSelf, IDC_QUOTECHAR_EDIT), EM_SETLIMITTEXT, static_cast<WPARAM>(1), 0);

    isWindowOpen = true;
}

bool MultiReplace::createAndShowWindows() {
    // IDs of all controls in the "Replace/Find in Files" panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    // IDs of List Search Bar controls (initially hidden)
    static const std::vector<int> listSearchBarIds = {
        IDC_LIST_SEARCH_COMBO, IDC_LIST_SEARCH_BUTTON, IDC_LIST_SEARCH_CLOSE
    };

    auto isRepInFilesId = [&](int id) {
        return std::find(repInFilesIds.begin(), repInFilesIds.end(), id) != repInFilesIds.end();
        };

    auto isListSearchBarId = [&](int id) {
        return std::find(listSearchBarIds.begin(), listSearchBarIds.end(), id) != listSearchBarIds.end();
        };

    const bool twoButtonsMode = (IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED);
    const bool initialShow = (isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode;

    for (auto& pair : ctrlMap)
    {
        const bool isFilesCtrl = isRepInFilesId(pair.first);
        const bool isSearchBarCtrl = isListSearchBarId(pair.first);

        // Create all controls as children, but only set WS_VISIBLE if needed
        DWORD style = pair.second.style | WS_CHILD;

        // Determine visibility:
        // - Files panel controls: visible only if initialShow is true
        // - Search bar controls: always start hidden
        // - All other controls: always visible
        if (isSearchBarCtrl) {
            // Search bar: always start hidden (no WS_VISIBLE)
        }
        else if (isFilesCtrl) {
            // Files panel: visible only if initialShow
            if (initialShow) {
                style |= WS_VISIBLE;
            }
        }
        else {
            // All other controls: always visible
            style |= WS_VISIBLE;
        }

        HWND hwndControl = CreateWindowEx(
            0,
            pair.second.className,
            pair.second.windowName,
            style,
            pair.second.x, pair.second.y, pair.second.cx, pair.second.cy,
            _hSelf,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(pair.first)),
            hInstance,
            NULL
        );
        if (!hwndControl) {
            return false; // Handle creation error
        }

        // Only create tooltips if enabled and tooltip text is available
        if ((tooltipsEnabled || pair.first == IDC_FILTER_HELP)
            && pair.second.tooltipText != nullptr
            && pair.second.tooltipText[0] != L'\0')
        {
            // Create a tooltip window for this control
            HWND hwndTooltip = CreateWindowEx(
                0,                                   // no extended styles
                TOOLTIPS_CLASS,                      // tooltip class
                nullptr,
                WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON | TTS_NOPREFIX,
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                _hSelf,                              // parent = our panel/dialog
                nullptr,
                hInstance,
                nullptr
            );

            if (hwndTooltip)
            {
                // Apply dark mode theme to tooltip
                if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
                    SetWindowTheme(hwndTooltip, L"DarkMode_Explorer", nullptr);
                }

                // Limit width only for the "?" help tooltip; 0 = unlimited
                DWORD maxWidth = (pair.first == IDC_FILTER_HELP) ? 200 : 0;
                SendMessage(hwndTooltip, TTM_SETMAXTIPWIDTH, 0, maxWidth);

                // Bind the tooltip to the specific child control (by HWND)
                TOOLINFO ti = {};
                ti.cbSize = sizeof(ti);
                ti.hwnd = _hSelf;                         // parent window
                ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;  // subclass the control to show tips
                ti.uId = (UINT_PTR)hwndControl;           // identify by child HWND
                ti.lpszText = const_cast<LPWSTR>(pair.second.tooltipText);
                SendMessage(hwndTooltip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
            }
        }
    }
    return true;
}

void MultiReplace::ensureIndicatorContext()
{
    HWND hSci0 = nppData._scintillaMainHandle;
    HWND hSci1 = nppData._scintillaSecondHandle;
    if (!hSci0) return;

    std::vector<int> preferred(std::begin(kPreferredIds), std::end(kPreferredIds));
    std::vector<int> reserved(std::begin(kReservedIds), std::end(kReservedIds));

    // (Re-)init coordinator only if handles changed
    NppStyleKit::gIndicatorCoord.ensureIndicatorsInitialized(hSci0, hSci1, preferred, reserved);

    // ColumnTabs: reserve preferred or first free
    const bool colIdValid =
        (NppStyleKit::gColumnTabsIndicatorId >= 0) &&
        NppStyleKit::gIndicatorCoord.isIndicatorReserved(NppStyleKit::gColumnTabsIndicatorId);

    if (!colIdValid) {
        const int wantCol = preferredColumnTabsStyleId; // -1 = auto
        NppStyleKit::gColumnTabsIndicatorId =
            NppStyleKit::gIndicatorCoord.reservePreferredOrFirstIndicator("ColumnTabs", wantCol);
    }
    ColumnTabs::CT_SetIndicatorId(NppStyleKit::gColumnTabsIndicatorId);

    // Registry gets the full remaining pool (no dedicated standard id)
    auto remaining = NppStyleKit::gIndicatorCoord.availableIndicatorPool();
    NppStyleKit::gIndicatorReg.init(hSci0, hSci1, remaining, /*capacityHint*/100);

    textStyles = remaining;
    textStylesList = remaining;

}

void MultiReplace::initializeListView() {
    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    SetWindowSubclass(_replaceListView, &MultiReplace::ListViewSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));
    createListViewColumns();
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Set tooltips based on isHoverTextEnabled
    DWORD extendedStyle = LVS_EX_FULLROWSELECT;
    if (isHoverTextEnabled) {
        extendedStyle |= LVS_EX_INFOTIP; // Enable tooltips
    }
    ListView_SetExtendedListViewStyle(_replaceListView, extendedStyle);

    // Initialize columnSortOrder with Unsorted state for all sortable columns
    columnSortOrder[ColumnID::FIND_COUNT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::REPLACE_COUNT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::FIND_TEXT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::REPLACE_TEXT] = SortDirection::Unsorted;
    columnSortOrder[ColumnID::COMMENTS] = SortDirection::Unsorted;

}

void MultiReplace::initializeDragAndDrop() {
    dropTarget = new DropTarget(this);
    HRESULT hr = ::RegisterDragDrop(_replaceListView, dropTarget);

    if (FAILED(hr)) {
        // COM-correct cleanup: Release() will delete the object when refCount reaches 0
        dropTarget->Release();
        dropTarget = nullptr;
    }
}

void MultiReplace::moveAndResizeControls(bool moveStatic) {
    int moveCount = 0;
    for (const auto& pair : ctrlMap) {
        if (GetDlgItem(_hSelf, pair.first)) moveCount++;
    }
    if (moveCount == 0) return;

    HDWP hdwp = BeginDeferWindowPos(moveCount);
    if (!hdwp) return;

    bool anyLayoutChanged = false;

    for (const auto& pair : ctrlMap) {
        int ctrlId = pair.first;
        const ControlInfo& ctrlInfo = pair.second;

        HWND resizeHwnd = GetDlgItem(_hSelf, ctrlId);
        if (!resizeHwnd) continue;

        // Skip static controls during resize operations to preserve layout stability
        if (!moveStatic && ctrlInfo.isStatic) {
            continue;
        }

        // 1. Target Geometry
        int targetX = ctrlInfo.x;
        int targetY = ctrlInfo.y;
        int targetW = ctrlInfo.cx;
        int targetH = ctrlInfo.cy;

        // Special height for ComboBoxes
        bool isDynamicHeightCombo = (ctrlId == IDC_FIND_EDIT || ctrlId == IDC_REPLACE_EDIT ||
            ctrlId == IDC_DIR_EDIT || ctrlId == IDC_FILTER_EDIT);

        if (isDynamicHeightCombo) {
            COMBOBOXINFO cbi = { sizeof(COMBOBOXINFO) };
            if (GetComboBoxInfo(resizeHwnd, &cbi)) {
                targetH = cbi.rcItem.bottom - cbi.rcItem.top;
            }
        }

        // Save selection
        bool isComboBox = (ctrlInfo.className && wcscmp(ctrlInfo.className, WC_COMBOBOX) == 0);
        bool isSelectionSensitive = isComboBox || ctrlId == IDC_REPLACE_HIT_EDIT ||
            ctrlId == IDC_COLUMN_NUM_EDIT || ctrlId == IDC_DELIMITER_EDIT ||
            ctrlId == IDC_QUOTECHAR_EDIT;

        DWORD startSelection = 0, endSelection = 0;
        if (isSelectionSensitive) {
            SendMessage(resizeHwnd, CB_GETEDITSEL, reinterpret_cast<WPARAM>(&startSelection), reinterpret_cast<LPARAM>(&endSelection));
        }

        // Queue Move
        hdwp = DeferWindowPos(hdwp, resizeHwnd, nullptr, targetX, targetY, targetW, targetH,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);

        // Restore selection
        if (isSelectionSensitive) {
            SendMessage(resizeHwnd, CB_SETEDITSEL, 0, MAKELPARAM(startSelection, endSelection));
        }

        anyLayoutChanged = true;
    }

    EndDeferWindowPos(hdwp);

    if (anyLayoutChanged) {
        showListFilePath();
    }
}

void MultiReplace::updateTwoButtonsVisibility() {
    BOOL twoButtonsMode = IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED;

    auto setVisibility = [this](const std::vector<int>& elements, bool condition) {
        for (int id : elements) {
            ShowWindow(GetDlgItem(_hSelf, id), condition ? SW_SHOW : SW_HIDE);
        }
        };

    // Replace-Buttons
    setVisibility({ IDC_REPLACE_ALL_SMALL_BUTTON, IDC_REPLACE_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_REPLACE_ALL_BUTTON }, !twoButtonsMode);

    // Find-Buttons
    setVisibility({ IDC_FIND_NEXT_BUTTON, IDC_FIND_PREV_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_FIND_ALL_BUTTON }, !twoButtonsMode);

    // Mark-Buttons
    setVisibility({ IDC_MARK_MATCHES_BUTTON, IDC_COPY_MARKED_TEXT_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_MARK_BUTTON }, !twoButtonsMode);

    // Load-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_LOAD_FROM_CSV_BUTTON }, !twoButtonsMode);

    // Save-Buttons (only depend on twoButtonsMode now)
    setVisibility({ IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON }, twoButtonsMode);
    setVisibility({ IDC_SAVE_TO_CSV_BUTTON }, !twoButtonsMode);

    updateFilesPanel();
}

void MultiReplace::updateListViewFrame()
{
    HWND lv = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    if (!lv) return;

    const ControlInfo& ci = ctrlMap[IDC_REPLACE_LIST];
    MoveWindow(lv, ci.x, ci.y, ci.cx, ci.cy, TRUE);
}

void MultiReplace::repaintPanelContents(HWND hGrp, const std::wstring& title)
{
    // All controls inside the "Replace/Find in Files" panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    // Keep the caption in sync
    SetDlgItemText(_hSelf, IDC_FILE_OPS_GROUP, title.c_str());

    // Never repaint the frame if the group is hidden. This avoids "ghost frames".
    if (!IsWindowVisible(hGrp)) {
        return; // Title updated for next show; no drawing while hidden.
    }

    // Map group rect to parent coordinates
    RECT rcGrp;
    GetWindowRect(hGrp, &rcGrp);
    MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcGrp), 2);

    // Clean the parent background in the group area
    RedrawWindow(_hSelf, &rcGrp, nullptr,
        RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_NOCHILDREN);

    // Redraw the groupbox (including non-client frame)
    RedrawWindow(hGrp, nullptr, nullptr,
        RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE | RDW_NOCHILDREN | RDW_FRAME);

    // Redraw visible children (no erase to prevent flicker)
    for (int id : repInFilesIds) {
        if (id == IDC_FILE_OPS_GROUP) continue;
        HWND hChild = GetDlgItem(_hSelf, id);
        if (IsWindow(hChild) && IsWindowVisible(hChild)) {
            RedrawWindow(hChild, nullptr, nullptr,
                RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
        }
    }
}

void MultiReplace::updateFilesPanel()
{
    // All controls inside the "Replace/Find in Files" panel
    static const std::vector<int> repInFilesIds = {
        IDC_FILE_OPS_GROUP,
        IDC_FILTER_STATIC,  IDC_FILTER_EDIT,  IDC_FILTER_HELP,
        IDC_DIR_STATIC,     IDC_DIR_EDIT,     IDC_BROWSE_DIR_BUTTON,
        IDC_SUBFOLDERS_CHECKBOX, IDC_HIDDENFILES_CHECKBOX,
        IDC_CANCEL_REPLACE_BUTTON
    };

    // Persisted UI state to avoid unnecessary work
    static bool lastShow = false;
    static std::wstring lastTitleKey;

    const bool twoButtonsMode = (IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED);
    const bool show = (isReplaceInFiles || isFindAllInFiles) && !twoButtonsMode;

    // Determine title
    std::wstring titleKey;
    std::wstring titleText;
    if (isReplaceInFiles && isFindAllInFiles) {
        titleKey = L"panel_find_replace_in_files";
        titleText = LM.get(titleKey);
    }
    else if (isFindAllInFiles) {
        titleKey = L"panel_find_in_files";
        titleText = LM.get(titleKey);
    }
    else {
        titleKey = L"panel_replace_in_files";
        titleText = LM.get(titleKey);
    }

    HWND hGrp = GetDlgItem(_hSelf, IDC_FILE_OPS_GROUP);
    HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);

    // Capture geometry BEFORE layout changes (used to clean freed area when hiding)
    RECT rcGrpBefore{};
    if (IsWindow(hGrp)) {
        GetWindowRect(hGrp, &rcGrpBefore);
        MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcGrpBefore), 2);
    }
    RECT rcStatusBefore{};
    if (IsWindow(hStatus)) {
        GetWindowRect(hStatus, &rcStatusBefore);
        MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcStatusBefore), 2);
    }

    // Case A) Visibility toggles: full show/hide + relayout + localized repaint
    if (show != lastShow)
    {

        // Toggle visibility of all panel controls
        for (int id : repInFilesIds) {
            ShowWindow(GetDlgItem(_hSelf, id), show ? SW_SHOW : SW_HIDE);
        }

        // Layout after visibility change
        RECT rcClient; GetClientRect(_hSelf, &rcClient);
        positionAndResizeControls(rcClient.right - rcClient.left, rcClient.bottom - rcClient.top);
        updateListViewFrame();
        moveAndResizeControls(false);
        adjustWindowSize();
        onSelectionChanged();

        // Thaw parent painting
        SendMessage(_hSelf, WM_SETREDRAW, TRUE, 0);

        // Compute "after" rect for status (for union repaint when hiding)
        RECT rcStatusAfter{};
        if (IsWindow(hStatus)) {
            GetWindowRect(hStatus, &rcStatusAfter);
            MapWindowPoints(HWND_DESKTOP, _hSelf, reinterpret_cast<LPPOINT>(&rcStatusAfter), 2);
        }

        if (show) {
            EnableWindow(GetDlgItem(_hSelf, IDC_CANCEL_REPLACE_BUTTON), FALSE);
            // Showing: repaint panel (title + frame + children) safely
            repaintPanelContents(hGrp, titleText);
            // Clear selection in both edit fields after opening the panel
            SendMessage(GetDlgItem(_hSelf, IDC_FILTER_EDIT), CB_SETEDITSEL, 0, 0);
            SendMessage(GetDlgItem(_hSelf, IDC_DIR_EDIT), CB_SETEDITSEL, 0, 0);
        }
        else {

            // Force controls that moved up to repaint (no erase to avoid flicker)
            const int idsShiftedUp[] = {
                IDC_REPLACE_LIST, IDC_STATUS_MESSAGE, IDC_PATH_DISPLAY, IDC_STATS_DISPLAY,
                IDC_LOAD_FROM_CSV_BUTTON, IDC_LOAD_LIST_BUTTON, IDC_NEW_LIST_BUTTON,
                IDC_SAVE_TO_CSV_BUTTON,  IDC_SAVE_BUTTON, IDC_SAVE_AS_BUTTON,
                IDC_EXPORT_BASH_BUTTON, IDC_UP_BUTTON, IDC_DOWN_BUTTON
            };
            for (int id : idsShiftedUp) {
                HWND hCtrl = GetDlgItem(_hSelf, id);
                if (IsWindow(hCtrl) && IsWindowVisible(hCtrl)) {
                    RedrawWindow(hCtrl, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_NOERASE);
                }
            }

            // IMPORTANT: Do NOT change the group title while hidden.
            // No SetDlgItemText() here — prevents any accidental frame repaint.
            lastTitleKey.clear();

        }
        RedrawWindow(_hSelf, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW);

        lastShow = show;
        if (show) lastTitleKey = titleKey;
        return;
    }

    // Case B) Still hidden: nothing to do (keep state clean)
    if (!show) {
        lastShow = false;
        lastTitleKey.clear();
        return;
    }

    // Case C) Visible and only the title changed
    if (titleKey != lastTitleKey)
    {
        // Repaint only the panel contents/title without relayout
        repaintPanelContents(hGrp, titleText);
        lastTitleKey = titleKey;
    }
}

void MultiReplace::setUIElementVisibility() {
    // Determine the state of mode radio buttons
    bool regexChecked = SendMessage(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_GETCHECK, 0, 0) == BST_CHECKED;
    bool columnModeChecked = SendMessage(GetDlgItem(_hSelf, IDC_COLUMN_MODE_RADIO), BM_GETCHECK, 0, 0) == BST_CHECKED;

    // Whole Word must be disabled in regex mode (and unchecked to avoid stale state)
    HWND hWholeWord = GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX);
    if (regexChecked) {
        EnableWindow(hWholeWord, FALSE);
        SendMessageW(hWholeWord, BM_SETCHECK, BST_UNCHECKED, 0);
    }
    else {
        EnableWindow(hWholeWord, TRUE);
    }

    // Column-mode dependent controls
    const std::vector<int> columnRadioDependentElements = {
        IDC_COLUMN_SORT_DESC_BUTTON, IDC_COLUMN_SORT_ASC_BUTTON, IDC_COLUMN_DROP_BUTTON,
        IDC_COLUMN_COPY_BUTTON, IDC_COLUMN_HIGHLIGHT_BUTTON, IDC_COLUMN_GRIDTABS_BUTTON,
        IDC_COLUMN_DUPLICATES_BUTTON
    };
    for (int id : columnRadioDependentElements) {
        EnableWindow(GetDlgItem(_hSelf, id), columnModeChecked);
    }

    // Update the FIND_PREV_BUTTON state based on Regex and Selection mode
    // EnableWindow(GetDlgItem(_hSelf, IDC_FIND_PREV_BUTTON), !regexChecked);
}

void MultiReplace::drawGripper() {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(_hSelf, &ps);

    RECT rect;
    GetClientRect(_hSelf, &rect);

    // Gripper position and dimensions
    constexpr int GRIPPER_BASE_SIZE = 11;
    int gripperSize = sx(GRIPPER_BASE_SIZE);
    POINT startPoint = { rect.right - gripperSize, rect.bottom - gripperSize };

    int dotSize = sx(2);
    int gap = std::max(sx(1), 1);

    // Dark Mode aware color
    bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
    COLORREF dotColor = isDark ? RGB(100, 100, 100) : RGB(200, 200, 200);
    HBRUSH hBrush = CreateSolidBrush(dotColor);

    // Triangle pattern: 6 dots
    static constexpr int positions[3][3] = {
        {0, 0, 1},
        {0, 1, 1},
        {1, 1, 1}
    };

    for (int row = 0; row < 3; ++row)
    {
        for (int col = 0; col < 3; ++col)
        {
            if (positions[row][col] == 0) continue;

            int x = startPoint.x + col * (dotSize + gap);
            int y = startPoint.y + row * (dotSize + gap);
            RECT dotRect = { x, y, x + dotSize, y + dotSize };
            FillRect(hdc, &dotRect, hBrush);
        }
    }

    DeleteObject(hBrush);
    EndPaint(_hSelf, &ps);
}

void MultiReplace::SetWindowTransparency(HWND hwnd, BYTE alpha) {
    LONG style = GetWindowLong(hwnd, GWL_EXSTYLE);
    SetWindowLong(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED);
    SetLayeredWindowAttributes(hwnd, 0, alpha, LWA_ALPHA);
}

void MultiReplace::adjustWindowSize() {

    // Get the minimum allowed window size
    RECT minSize = calculateMinWindowFrame(_hSelf);
    int minHeight = minSize.bottom; // minSize.bottom is minHeight

    // Get the current window position and size
    RECT currentRect;
    GetWindowRect(_hSelf, &currentRect);
    int currentWidth = currentRect.right - currentRect.left;
    int currentX = currentRect.left;
    int currentY = currentRect.top;

    int newHeight = useListEnabled ? std::max(useListOnHeight, minHeight) : useListOffHeight;

    // Adjust the window size while keeping the current position and width
    SetWindowPos(_hSelf, nullptr, currentX, currentY, currentWidth, newHeight, SWP_NOZORDER);
}

void MultiReplace::updateUseListState(bool isUpdate)
{   // isUpdate - specifies whether to update the existing tooltip (true) or create a new one (false)

    // Set the button text based on the current state
    SetDlgItemText(_hSelf, IDC_USE_LIST_BUTTON, useListEnabled ? L"˄" : L"˅");

    // Set the status message based on the list state
    showStatusMessage(useListEnabled ? LM.get(L"status_enable_list") : LM.get(L"status_disable_list"), MessageStatus::Info);

    // Return early if tooltips are disabled
    if (!tooltipsEnabled) {
        return;
    }

    // Create the tooltip window if it doesn't exist yet
    if (!isUpdate && !_hUseListButtonTooltip)
    {
        _hUseListButtonTooltip = CreateWindowEx(
            0, TOOLTIPS_CLASS, nullptr,
            WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON,  // Use the same styles as in createAndShowWindows
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            _hSelf, nullptr, hInstance, nullptr);

        if (!_hUseListButtonTooltip)
        {
            // Handle error if tooltip creation fails
            return;
        }

        // Activate the tooltip
        SendMessage(_hUseListButtonTooltip, TTM_ACTIVATE, TRUE, 0);

        // Apply dark mode theme to tooltip
        if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
            SetWindowTheme(_hUseListButtonTooltip, L"DarkMode_Explorer", nullptr);
        }
    }

    // Prepare the TOOLINFO structure
    TOOLINFO ti = {};
    ti.cbSize = sizeof(TOOLINFO);
    ti.hwnd = _hSelf;  // Parent window handle
    ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    ti.uId = (UINT_PTR)GetDlgItem(_hSelf, IDC_USE_LIST_BUTTON);  // Associate with the specific control

    // Get tooltip text based on the list state
    LPCWSTR tooltipText = useListEnabled ? LM.getLPCW(L"tooltip_disable_list") : LM.getLPCW(L"tooltip_enable_list");
    ti.lpszText = const_cast<LPWSTR>(tooltipText);  // Assign the tooltip text

    // If it's an update, delete the old tooltip first
    if (isUpdate)
    {
        SendMessage(_hUseListButtonTooltip, TTM_DELTOOL, 0, reinterpret_cast<LPARAM>(&ti));
    }

    // Add or update the tooltip
    SendMessage(_hUseListButtonTooltip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
}

void MultiReplace::loadLanguageGlobal() {
    if (!nppData._nppHandle) return;

    wchar_t pluginDir[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, MAX_PATH, reinterpret_cast<LPARAM>(pluginDir));

    wchar_t langXml[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, reinterpret_cast<LPARAM>(langXml));
    wcscat_s(langXml, L"\\..\\..\\nativeLang.xml");

    LanguageManager::instance().load(pluginDir, langXml);
}

void MultiReplace::refreshUILanguage()
{
    // Reload language strings
    loadLanguageGlobal();

    if (!_MultiReplace._hSelf || !IsWindow(_MultiReplace._hSelf))
        return;

    // Rebuild ctrlMap with new language strings
    RECT rc;
    GetClientRect(_MultiReplace._hSelf, &rc);
    _MultiReplace.positionAndResizeControls(rc.right, rc.bottom);

    // Update all controls - SetWindowPos forces Windows to recalculate text extent
    for (auto& pair : _MultiReplace.ctrlMap) {
        HWND hCtrl = GetDlgItem(_MultiReplace._hSelf, pair.first);
        if (!hCtrl) continue;

        // Only set text for controls with defined windowName (skip edit fields/comboboxes)
        if (pair.second.windowName && pair.second.windowName[0] != L'\0') {
            SetWindowTextW(hCtrl, pair.second.windowName);
        }

        SetWindowPos(hCtrl, nullptr,
            pair.second.x, pair.second.y,
            pair.second.cx, pair.second.cy,
            SWP_NOZORDER | SWP_NOACTIVATE);
    }

    // Update ListView headers
    if (_MultiReplace._replaceListView) {
        auto& colIndices = _MultiReplace.columnIndices;
        LVCOLUMN lvc = {};
        lvc.mask = LVCF_TEXT;

        for (size_t i = 0; i < kHeaderTextMappingsCount; ++i) {
            ColumnID colId = static_cast<ColumnID>(kHeaderTextMappings[i].columnId);
            auto it = colIndices.find(colId);
            if (it != colIndices.end() && it->second >= 0) {
                lvc.pszText = LM.getW(kHeaderTextMappings[i].langKey);
                ListView_SetColumn(_MultiReplace._replaceListView, it->second, &lvc);
            }
        }

        // Update header tooltips
        if (_MultiReplace._hHeaderTooltip) {
            HWND hwndHeader = ListView_GetHeader(_MultiReplace._replaceListView);
            if (hwndHeader) {
                for (size_t i = 0; i < kHeaderTooltipMappingsCount; ++i) {
                    ColumnID colId = static_cast<ColumnID>(kHeaderTooltipMappings[i].columnId);
                    auto it = colIndices.find(colId);
                    if (it != colIndices.end() && it->second >= 0) {
                        TOOLINFO ti = { sizeof(TOOLINFO) };
                        ti.hwnd = hwndHeader;
                        ti.uId = static_cast<UINT_PTR>(it->second);
                        ti.lpszText = LM.getW(kHeaderTooltipMappings[i].langKey);
                        SendMessage(_MultiReplace._hHeaderTooltip, TTM_UPDATETIPTEXT, 0, reinterpret_cast<LPARAM>(&ti));
                    }
                }
            }
        }
    }

    // Update ConfigDialog
    _MultiReplaceConfig.refreshUILanguage();

    // Update Debug Window
    if (hDebugWnd && IsWindow(hDebugWnd)) {
        SetWindowTextW(hDebugWnd, LM.getLPCW(L"debug_title"));

        // Update buttons (IDs: 2=Next, 3=Stop, 4=Copy)
        HWND hBtn = GetDlgItem(hDebugWnd, 2);
        if (hBtn) SetWindowTextW(hBtn, LM.getLPCW(L"debug_btn_next"));
        hBtn = GetDlgItem(hDebugWnd, 3);
        if (hBtn) SetWindowTextW(hBtn, LM.getLPCW(L"debug_btn_stop"));
        hBtn = GetDlgItem(hDebugWnd, 4);
        if (hBtn) SetWindowTextW(hBtn, LM.getLPCW(L"debug_btn_copy"));

        // Update ListView columns
        if (hDebugListView && IsWindow(hDebugListView)) {
            LVCOLUMN lvc = {};
            lvc.mask = LVCF_TEXT;
            lvc.pszText = LM.getW(L"debug_col_variable");
            ListView_SetColumn(hDebugListView, 0, &lvc);
            lvc.pszText = LM.getW(L"debug_col_type");
            ListView_SetColumn(hDebugListView, 1, &lvc);
            lvc.pszText = LM.getW(L"debug_col_value");
            ListView_SetColumn(hDebugListView, 2, &lvc);
        }
    }
    RedrawWindow(_MultiReplace._hSelf, nullptr, nullptr, RDW_ERASE | RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
}

#pragma endregion


#pragma region List Data Operations

void MultiReplace::addItemsToReplaceList(const std::vector<ReplaceItemData>& items, size_t insertPosition = std::numeric_limits<size_t>::max()) {
    // Determine the insert position
    if (insertPosition > replaceListData.size()) {
        insertPosition = replaceListData.size();
    }

    size_t startIndex = insertPosition;
    size_t endIndex = startIndex + items.size() - 1;

    // Insert items at the specified position
    replaceListData.insert(replaceListData.begin() + insertPosition, items.begin(), items.end());

    // Update the ListView
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);

    // Create undo and redo actions
    UndoRedoAction action;

    // Undo action: Remove the added items
    action.undoAction = [this, startIndex, endIndex]() {
        // Remove items from the replace list
        replaceListData.erase(replaceListData.begin() + startIndex, replaceListData.begin() + endIndex + 1);

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Scroll to the position where items were removed
        scrollToIndices(startIndex, startIndex);
        };

    // Redo action: Re-insert the items
    std::vector<ReplaceItemData> itemsToRedo(items); // Copy of items to redo
    action.redoAction = [this, itemsToRedo, startIndex, endIndex]() {
        replaceListData.insert(replaceListData.begin() + startIndex, itemsToRedo.begin(), itemsToRedo.end());

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the re-inserted items
        for (size_t i = startIndex; i <= endIndex; ++i) {
            ListView_SetItemState(_replaceListView, static_cast<int>(i), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the inserted items
        scrollToIndices(startIndex, endIndex);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Add items");
}

void MultiReplace::removeItemsFromReplaceList(const std::vector<size_t>& indicesToRemove) {
    // Sort indices in descending order to handle shifting issues during removal
    std::vector<size_t> sortedIndices = indicesToRemove;
    std::sort(sortedIndices.rbegin(), sortedIndices.rend());

    // Store removed items along with their indices for undo purposes
    std::vector<std::pair<size_t, ReplaceItemData>> removedItemsWithIndices;
    for (size_t index : sortedIndices) {
        if (index < replaceListData.size()) {
            removedItemsWithIndices.emplace_back(index, replaceListData[index]);
            replaceListData.erase(replaceListData.begin() + index);
        }
    }

    // Update the ListView to reflect changes
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);

    // Create undo and redo actions
    UndoRedoAction action;

    // Undo action: Re-insert the removed items at their original positions
    action.undoAction = [this, removedItemsWithIndices]() {
        for (auto it = removedItemsWithIndices.rbegin(); it != removedItemsWithIndices.rend(); ++it) {
            const auto& [index, item] = *it;
            if (index <= replaceListData.size()) {
                replaceListData.insert(replaceListData.begin() + index, item);
            }
        }

        // Update the ListView to include restored items
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select restored rows
        for (const auto& [index, _] : removedItemsWithIndices) {
            if (index < replaceListData.size()) {
                ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
            }
        }

        // Calculate visible range and ensure restored rows are visible
        std::vector<size_t> indices;
        for (const auto& [index, _] : removedItemsWithIndices) {
            indices.push_back(index);
        }
        size_t firstIndex = *std::min_element(indices.begin(), indices.end());
        size_t lastIndex = *std::max_element(indices.begin(), indices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    // Redo action: Remove the same items again
    action.redoAction = [this, removedItemsWithIndices]() {
        // Sort indices in descending order to avoid shifting issues
        std::vector<size_t> sortedIndices;
        for (const auto& [index, _] : removedItemsWithIndices) {
            sortedIndices.push_back(index);
        }
        std::sort(sortedIndices.rbegin(), sortedIndices.rend());

        for (size_t index : sortedIndices) {
            if (index < replaceListData.size()) {
                replaceListData.erase(replaceListData.begin() + index);
            }
        }

        // Update the ListView to reflect removal
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Scroll to the position where deletion occurred
        if (!sortedIndices.empty()) {
            size_t minIndex = sortedIndices.back(); // Since sortedIndices is descending
            scrollToIndices(minIndex, minIndex);
        }
        };

    // Push the action to the undoStack
    URM.push(action.undoAction, action.redoAction, L"Remove items");
}

void MultiReplace::modifyItemInReplaceList(size_t index, const ReplaceItemData& newData) {
    // Store the original data
    ReplaceItemData originalData = replaceListData[index];

    // Modify the item
    replaceListData[index] = newData;

    // Update the ListView item
    updateListViewItem(index);

    // Create Undo/Redo actions
    UndoRedoAction action;

    // Undo action: Restore the original data
    action.undoAction = [this, index, originalData]() {
        replaceListData[index] = originalData;
        updateListViewItem(index);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select and ensure visibility of the modified item
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
        scrollToIndices(index, index);

        // Set focus to the ListView
        SetFocus(_replaceListView);
        };

    // Redo action: Apply the new data again
    action.redoAction = [this, index, newData]() {
        replaceListData[index] = newData;
        updateListViewItem(index);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select and ensure visibility of the modified item
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED, LVIS_SELECTED);
        scrollToIndices(index, index);

        // Set focus to the ListView
        SetFocus(_replaceListView);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Modify item");
}

bool MultiReplace::moveItemsInReplaceList(std::vector<size_t>& indices, Direction direction) {
    if (indices.empty()) {
        return false; // No items to move
    }

    // Check the bounds
    if ((direction == Direction::Up && indices.front() == 0) ||
        (direction == Direction::Down && indices.back() == replaceListData.size() - 1)) {
        return false; // Out of bounds, do nothing
    }

    // Store pre-move indices for undo
    std::vector<size_t> preMoveIndices = indices;

    // Adjust indices for the move
    if (direction == Direction::Up) {
        for (size_t& idx : indices) {
            idx -= 1; // Adjust indices upwards
        }
    }
    else { // Direction::Down
        for (size_t& idx : indices) {
            idx += 1; // Adjust indices downwards
        }
    }

    // Perform the move
    for (size_t i = 0; i < preMoveIndices.size(); ++i) {
        std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
    }

    // Update the ListView
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);

    // Create Undo/Redo actions
    UndoRedoAction action;

    action.undoAction = [this, preMoveIndices, indices]() {
        // Swap back the moved items
        for (size_t i = 0; i < preMoveIndices.size(); ++i) {
            std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
        }

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the original positions
        for (size_t idx : preMoveIndices) {
            ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the moved items
        size_t firstIndex = *std::min_element(preMoveIndices.begin(), preMoveIndices.end());
        size_t lastIndex = *std::max_element(preMoveIndices.begin(), preMoveIndices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    action.redoAction = [this, preMoveIndices, indices]() {
        // Swap items to their new positions again
        for (size_t i = 0; i < preMoveIndices.size(); ++i) {
            std::swap(replaceListData[preMoveIndices[i]], replaceListData[indices[i]]);
        }

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Deselect all items
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

        // Select the moved positions
        for (size_t idx : indices) {
            ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
        }

        // Ensure visibility of the moved items
        size_t firstIndex = *std::min_element(indices.begin(), indices.end());
        size_t lastIndex = *std::max_element(indices.begin(), indices.end());
        scrollToIndices(firstIndex, lastIndex);
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Move items");

    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select the moved rows in their new positions
    for (size_t idx : indices) {
        ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
    }

    // Ensure visibility of the moved items
    size_t firstIndex = *std::min_element(indices.begin(), indices.end());
    size_t lastIndex = *std::max_element(indices.begin(), indices.end());
    scrollToIndices(firstIndex, lastIndex);

    return true;
}

void MultiReplace::sortItemsInReplaceList(const std::vector<size_t>& originalOrder, const std::vector<size_t>& newOrder, const std::map<int, SortDirection>& previousColumnSortOrder, int columnID, SortDirection direction) {
    UndoRedoAction action;

    // Undo action: Restore original order and sort state
    action.undoAction = [this, originalOrder, previousColumnSortOrder]() {
        // Restore original order
        std::unordered_map<size_t, ReplaceItemData> idToItemMap;
        for (const auto& item : replaceListData) {
            idToItemMap[item.id] = item;
        }

        replaceListData.clear();
        for (size_t id : originalOrder) {
            replaceListData.push_back(idToItemMap[id]);
        }

        // Restore previous sort order
        columnSortOrder = previousColumnSortOrder;

        // Update UI
        updateHeaderSortDirection();
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);
        };

    // Redo action: Restore sorted order and current sort state
    action.redoAction = [this, newOrder, columnID, direction]() {
        // Restore sorted order
        std::unordered_map<size_t, ReplaceItemData> idToItemMap;
        for (const auto& item : replaceListData) {
            idToItemMap[item.id] = item;
        }

        replaceListData.clear();
        for (size_t id : newOrder) {
            replaceListData.push_back(idToItemMap[id]);
        }

        // Update sort state
        columnSortOrder.clear();
        columnSortOrder[columnID] = direction;

        // Update UI
        updateHeaderSortDirection();
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);
        };

    // Push the action onto the undo stack
    URM.push(action.undoAction, action.redoAction, L"Sort items");
}

void MultiReplace::scrollToIndices(size_t firstIndex, size_t lastIndex) {
    if (firstIndex > lastIndex) {
        std::swap(firstIndex, lastIndex);
    }
    if (lastIndex >= replaceListData.size()) {
        lastIndex = replaceListData.empty() ? 0 : replaceListData.size() - 1;
    }

    // Get the client area of the ListView
    RECT rcClient;
    GetClientRect(_replaceListView, &rcClient);

    // Calculate the height of a single item in the ListView
    int itemHeight = ListView_GetItemSpacing(_replaceListView, TRUE) >> 16;

    // Calculate the number of fully visible items in the client area
    int visibleItemCount = rcClient.bottom / itemHeight;

    // Calculate the middle index of the range
    size_t middleIndex = firstIndex + (lastIndex - firstIndex) / 2;

    // Determine the desired top index so that the middle of the range is centered
    int desiredTopIndex = static_cast<int>(middleIndex) - (visibleItemCount / 2);

    // Ensure the desired top index is within valid bounds
    if (desiredTopIndex < 0) {
        desiredTopIndex = 0;
    }
    else if (desiredTopIndex + visibleItemCount > static_cast<int>(replaceListData.size())) {
        desiredTopIndex = static_cast<int>(replaceListData.size()) - visibleItemCount;
        if (desiredTopIndex < 0) {
            desiredTopIndex = 0;
        }
    }

    // Get the current top index of the ListView
    int currentTopIndex = ListView_GetTopIndex(_replaceListView);

    // Calculate the scroll offset in items
    int scrollItems = desiredTopIndex - currentTopIndex;

    // Scroll the ListView if necessary
    if (scrollItems != 0) {
        int scrollPixels = scrollItems * itemHeight;
        ListView_Scroll(_replaceListView, 0, scrollPixels);
    }
}

void MultiReplace::exportDataToClipboard() {
    // Get selected items (or all if none selected)
    std::vector<size_t> selectedIndices;

    int itemCount = ListView_GetItemCount(_replaceListView);
    int selCount = ListView_GetSelectedCount(_replaceListView);

    if (selCount > 0) {
        // Get selected items
        int idx = -1;
        while ((idx = ListView_GetNextItem(_replaceListView, idx, LVNI_SELECTED)) != -1) {
            selectedIndices.push_back(static_cast<size_t>(idx));
        }
    }
    else {
        // No selection - use all items
        for (int i = 0; i < itemCount; ++i) {
            selectedIndices.push_back(static_cast<size_t>(i));
        }
    }

    if (selectedIndices.empty()) {
        showStatusMessage(LM.get(L"status_no_items_to_export"), MessageStatus::Error);
        return;
    }

    // Load settings from config
    const auto& cfg = ConfigManager::instance();
    std::wstring templateStr = cfg.readString(L"ExportData", L"Template",
        L"%FIND%\\t%REPLACE%\\t%FCOUNT%\\t%RCOUNT%\\t%COMMENT%");
    bool escapeChars = cfg.readBool(L"ExportData", L"Escape", false);
    bool includeHeader = cfg.readBool(L"ExportData", L"Header", false);

    // This converts \t to real tabs in the TEMPLATE, not in the data
    std::wstring processedTemplate = processTemplateEscapes(templateStr);

    // Build output
    std::wstring output;

    // Add header row if enabled
    if (includeHeader) {
        std::wstring headerLine = processedTemplate;

        // Replace all variables with header names
        headerLine = replaceTemplateVar(headerLine, L"%FIND%", L"Find");
        headerLine = replaceTemplateVar(headerLine, L"%REPLACE%", L"Replace");
        headerLine = replaceTemplateVar(headerLine, L"%FCOUNT%", L"FindCount");
        headerLine = replaceTemplateVar(headerLine, L"%RCOUNT%", L"ReplaceCount");
        headerLine = replaceTemplateVar(headerLine, L"%COMMENT%", L"Comment");
        headerLine = replaceTemplateVar(headerLine, L"%SEL%", L"Selected");
        headerLine = replaceTemplateVar(headerLine, L"%ROW%", L"Row");
        // Option flags
        headerLine = replaceTemplateVar(headerLine, L"%REGEX%", L"Regex");
        headerLine = replaceTemplateVar(headerLine, L"%CASE%", L"MatchCase");
        headerLine = replaceTemplateVar(headerLine, L"%WORD%", L"WholeWord");
        headerLine = replaceTemplateVar(headerLine, L"%EXT%", L"Extended");
        headerLine = replaceTemplateVar(headerLine, L"%VAR%", L"Variables");

        // NO processTemplateEscapes here - already done above
        output += headerLine;
        if (!headerLine.empty() && headerLine.back() != L'\n') {
            output += L"\r\n";
        }
    }

    // Process each item
    for (size_t idx : selectedIndices) {
        if (idx >= replaceListData.size()) continue;

        const ReplaceItemData& item = replaceListData[idx];
        std::wstring line = processedTemplate;  // Use already-processed template

        // Get field values
        std::wstring findText = item.findText;
        std::wstring replaceText = item.replaceText;
        std::wstring comment = item.comments;

        // Escape if checkbox is enabled
        // This converts real newlines to \n (two chars), tabs to \t, etc.
        if (escapeChars) {
            findText = StringUtils::escapeCsvValue(item.findText);
            replaceText = StringUtils::escapeCsvValue(item.replaceText);
            comment = StringUtils::escapeCsvValue(item.comments);
        }
        else {
            findText = StringUtils::quoteField(item.findText);
            replaceText = StringUtils::quoteField(item.replaceText);
            comment = StringUtils::quoteField(item.comments);
        }

        // Replace template variables - Main data
        line = replaceTemplateVar(line, L"%FIND%", findText);
        line = replaceTemplateVar(line, L"%REPLACE%", replaceText);
        line = replaceTemplateVar(line, L"%FCOUNT%", std::to_wstring(item.findCount >= 0 ? item.findCount : 0));
        line = replaceTemplateVar(line, L"%RCOUNT%", std::to_wstring(item.replaceCount >= 0 ? item.replaceCount : 0));
        line = replaceTemplateVar(line, L"%COMMENT%", comment);
        line = replaceTemplateVar(line, L"%SEL%", item.isEnabled ? L"1" : L"0");
        line = replaceTemplateVar(line, L"%ROW%", std::to_wstring(idx + 1));

        // Replace template variables - Option flags
        line = replaceTemplateVar(line, L"%REGEX%", item.regex ? L"1" : L"0");
        line = replaceTemplateVar(line, L"%CASE%", item.matchCase ? L"1" : L"0");
        line = replaceTemplateVar(line, L"%WORD%", item.wholeWord ? L"1" : L"0");
        line = replaceTemplateVar(line, L"%EXT%", item.extended ? L"1" : L"0");
        line = replaceTemplateVar(line, L"%VAR%", item.useVariables ? L"1" : L"0");

        output += line;

        // Add newline if template doesn't end with one
        if (!line.empty() && line.back() != L'\n') {
            output += L"\r\n";
        }
    }

    // Copy to clipboard (reuse existing clipboard logic pattern)
    if (output.empty()) return;

    if (OpenClipboard(_hSelf)) {
        EmptyClipboard();
        const SIZE_T size = (output.size() + 1) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
        if (hMem) {
            if (void* p = GlobalLock(hMem)) {
                memcpy(p, output.c_str(), size);
                GlobalUnlock(hMem);
                if (SetClipboardData(CF_UNICODETEXT, hMem)) {
                    CloseClipboard();
                    showStatusMessage(LM.get(L"status_exported_to_clipboard",
                        { std::to_wstring(selectedIndices.size()) }),
                        MessageStatus::Info);
                    return;
                }
            }
            GlobalFree(hMem);
        }
        CloseClipboard();
    }
    showStatusMessage(LM.get(L"status_export_failed"), MessageStatus::Error);
}

std::wstring MultiReplace::replaceTemplateVar(const std::wstring& tmpl,
    const std::wstring& var,
    const std::wstring& value) {
    std::wstring result = tmpl;
    size_t pos = 0;
    while ((pos = result.find(var, pos)) != std::wstring::npos) {
        result.replace(pos, var.length(), value);
        pos += value.length();
    }
    return result;
}

std::wstring MultiReplace::processTemplateEscapes(const std::wstring& tmpl) {
    std::wstring result;
    result.reserve(tmpl.size());

    for (size_t i = 0; i < tmpl.size(); ++i) {
        if (tmpl[i] == L'\\' && i + 1 < tmpl.size()) {
            switch (tmpl[i + 1]) {
            case L't': result += L'\t'; ++i; break;
            case L'n': result += L'\n'; ++i; break;
            case L'r': result += L'\r'; ++i; break;
            case L'\\': result += L'\\'; ++i; break;
            default: result += tmpl[i]; break;
            }
        }
        else {
            result += tmpl[i];
        }
    }
    return result;
}

#pragma endregion


#pragma region ListView

HWND MultiReplace::CreateHeaderTooltip(HWND hwndParent)
{
    HWND hwndTT = CreateWindowEx(
        WS_EX_TOPMOST,
        TOOLTIPS_CLASS,
        nullptr,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        hwndParent,
        nullptr,
        GetModuleHandle(NULL),
        NULL
    );

    if (hwndTT)
    {
        SendMessage(hwndTT, TTM_ACTIVATE, TRUE, 0);
    }

    return hwndTT;
}

void MultiReplace::AddHeaderTooltip(HWND hwndTT, HWND hwndHeader, int columnIndex, LPCTSTR pszText) {
    RECT rect;
    Header_GetItemRect(hwndHeader, columnIndex, &rect);

    TOOLINFO ti = {};
    ti.cbSize = sizeof(TOOLINFO);
    ti.uFlags = TTF_SUBCLASS;
    ti.hwnd = hwndHeader;
    ti.hinst = GetModuleHandle(NULL);
    ti.uId = columnIndex;
    ti.lpszText = const_cast<LPWSTR>(pszText);
    ti.rect = rect;

    SendMessage(hwndTT, TTM_DELTOOL, 0, reinterpret_cast<LPARAM>(&ti));
    SendMessage(hwndTT, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
}

void MultiReplace::createListViewColumns() {
    LVCOLUMN lvc;
    ZeroMemory(&lvc, sizeof(lvc));

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;

    // Retrieve column widths
    findCountColumnWidth = getColumnWidth(ColumnID::FIND_COUNT);
    replaceCountColumnWidth = getColumnWidth(ColumnID::REPLACE_COUNT);
    findColumnWidth = getColumnWidth(ColumnID::FIND_TEXT);
    replaceColumnWidth = getColumnWidth(ColumnID::REPLACE_TEXT);
    commentsColumnWidth = getColumnWidth(ColumnID::COMMENTS);
    deleteButtonColumnWidth = getColumnWidth(ColumnID::DELETE_BUTTON);

    // Delete all existing columns
    int columnCount = Header_GetItemCount(ListView_GetHeader(_replaceListView));
    for (int i = columnCount - 1; i >= 0; --i) {
        ListView_DeleteColumn(_replaceListView, i);
    }
    columnIndices.clear();  // Clear previous indices after deleting

    const ControlInfo& listCtrlInfo = ctrlMap[IDC_REPLACE_LIST];

    // Define the ResizableColWidths struct to pass to calcDynamicColWidth
    ResizableColWidths widths = {
        _replaceListView,
        listCtrlInfo.cx,
        (isFindCountVisible) ? findCountColumnWidth : 0,
        (isReplaceCountVisible) ? replaceCountColumnWidth : 0,
        findColumnWidth,
        replaceColumnWidth,
        (isCommentsColumnVisible) ? commentsColumnWidth : 0,
        (isDeleteButtonVisible) ? deleteButtonColumnWidth : 0,
        GetSystemMetrics(SM_CXVSCROLL)
    };

    // Calculate dynamic width for Find, Replace, and Comments columns
    int perColumnWidth = calcDynamicColWidth(widths);

    int currentIndex = 0;  // Start tracking columns

    // Column 1: Find Count
    if (isFindCountVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getW(L"header_find_count");
        lvc.cx = findCountColumnWidth;
        lvc.fmt = LVCFMT_LEFT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::FIND_COUNT] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::FIND_COUNT] = -1;  // Mark as not visible
    }

    // Column 2: Replace Count
    if (isReplaceCountVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getW(L"header_replace_count");
        lvc.cx = replaceCountColumnWidth;
        lvc.fmt = LVCFMT_LEFT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::REPLACE_COUNT] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::REPLACE_COUNT] = -1;  // Mark as not visible
    }

    // Column 3: Selection Checkbox
    lvc.iSubItem = currentIndex;
    lvc.pszText = L"\u2610";
    lvc.cx = getColumnWidth(ColumnID::SELECTION);
    lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::SELECTION] = currentIndex;
    ++currentIndex;

    // Column 4: Find Text (dynamic width)
    lvc.iSubItem = currentIndex;
    lvc.pszText = LM.getW(L"header_find");
    lvc.cx = (findColumnLockedEnabled ? findColumnWidth : perColumnWidth);
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::FIND_TEXT] = currentIndex;
    ++currentIndex;

    // Column 5: Replace Text (dynamic width)
    lvc.iSubItem = currentIndex;
    lvc.pszText = LM.getW(L"header_replace");
    lvc.cx = (replaceColumnLockedEnabled ? replaceColumnWidth : perColumnWidth);
    lvc.fmt = LVCFMT_LEFT;
    ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
    columnIndices[ColumnID::REPLACE_TEXT] = currentIndex;
    ++currentIndex;

    // Columns 6 to 10: Options (fixed width)
    const std::wstring options[] = {
        L"header_whole_word",
        L"header_match_case",
        L"header_use_variables",
        L"header_extended",
        L"header_regex"
    };
    for (int i = 0; i < 5; ++i) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getW(options[i]);
        lvc.cx = checkMarkWidth_scaled;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[static_cast<ColumnID>(static_cast<int>(ColumnID::WHOLE_WORD) + i)] = currentIndex;
        ++currentIndex;
    }

    // Column 11: Comments (dynamic width)
    if (isCommentsColumnVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = LM.getW(L"header_comments");
        lvc.cx = (commentsColumnLockedEnabled ? commentsColumnWidth : perColumnWidth);
        lvc.fmt = LVCFMT_LEFT;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::COMMENTS] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::COMMENTS] = -1;  // Mark as not visible
    }

    // Column 12: Delete Button (fixed width)
    if (isDeleteButtonVisible) {
        lvc.iSubItem = currentIndex;
        lvc.pszText = L"";
        lvc.cx = crossWidth_scaled;
        lvc.fmt = LVCFMT_CENTER | LVCFMT_FIXED_WIDTH;
        ListView_InsertColumn(_replaceListView, currentIndex, &lvc);
        columnIndices[ColumnID::DELETE_BUTTON] = currentIndex;
        ++currentIndex;
    }
    else {
        columnIndices[ColumnID::DELETE_BUTTON] = -1;
    }

    updateHeaderSortDirection();
    updateHeaderSelection();
    updateListViewTooltips();
}

void MultiReplace::insertReplaceListItem(const ReplaceItemData& itemData) {
    int useVariables = IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED ? 1 : 0;

    // Return early if findText is empty and "Use Variables" is not enabled
    if (itemData.findText.empty() && useVariables == 0) {
        showStatusMessage(LM.get(L"status_no_find_string"), MessageStatus::Error);
        return;
    }

    _replaceListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    // Check if itemData is already in the vector
    bool isDuplicate = false;
    for (const auto& existingItem : replaceListData) {
        if (existingItem == itemData) {
            isDuplicate = true;
            break;
        }
    }

    std::vector<ReplaceItemData> itemsToAdd = { itemData };
    addItemsToReplaceList(itemsToAdd);

    // Show a status message indicating the value added to the list
    std::wstring message;
    if (isDuplicate) {
        message = LM.get(L"status_duplicate_entry") + itemData.findText;
    }
    else {
        message = LM.get(L"status_value_added");
    }
    showStatusMessage(message, MessageStatus::Success);

    // Update the item count in the ListView
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);

    // Select newly added line
    size_t newItemIndex = replaceListData.size() - 1;
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED); // Delete previous selection
    ListView_SetItemState(_replaceListView, static_cast<int>(newItemIndex), LVIS_SELECTED, LVIS_SELECTED);
    ListView_EnsureVisible(_replaceListView, static_cast<int>(newItemIndex), FALSE);

    // Update Header if there might be any changes
    updateHeaderSelection();
}

std::wstring getColumnIDText(ColumnID columnID) {
    switch (columnID) {
    case ColumnID::INVALID:          return L"INVALID";
    case ColumnID::FIND_COUNT:       return L"FIND_COUNT";
    case ColumnID::REPLACE_COUNT:    return L"REPLACE_COUNT";
    case ColumnID::SELECTION:        return L"SELECTION";
    case ColumnID::FIND_TEXT:        return L"FIND_TEXT";
    case ColumnID::REPLACE_TEXT:     return L"REPLACE_TEXT";
    case ColumnID::WHOLE_WORD:       return L"WHOLE_WORD";
    case ColumnID::MATCH_CASE:       return L"MATCH_CASE";
    case ColumnID::USE_VARIABLES:    return L"USE_VARIABLES";
    case ColumnID::EXTENDED:         return L"EXTENDED";
    case ColumnID::REGEX:            return L"REGEX";
    case ColumnID::COMMENTS:         return L"COMMENTS";
    case ColumnID::DELETE_BUTTON:    return L"DELETE_BUTTON";
    default:                         return L"UNKNOWN";
    }
}

int MultiReplace::getColumnWidth(ColumnID columnID) {

    int width = 0;

    switch (columnID) {
    case ColumnID::DELETE_BUTTON:
        width = crossWidth_scaled;
        break;
    case ColumnID::SELECTION:
        width = boxWidth_scaled;
        break;
    case ColumnID::WHOLE_WORD:
    case ColumnID::MATCH_CASE:
    case ColumnID::USE_VARIABLES:
    case ColumnID::EXTENDED:
    case ColumnID::REGEX:
        width = checkMarkWidth_scaled;
        break;
    default:
        // For dynamic columns (Find, Replace, Comments, Find Count, Replace Count)
        auto it = columnIndices.find(columnID);
        if (it != columnIndices.end() && it->second != -1) {
            // Column exists, get width from ListView
            int index = it->second;
            width = ListView_GetColumnWidth(_replaceListView, index);
        }
        else {
            // Column does not exist, use default column width
            switch (columnID) {
            case ColumnID::FIND_COUNT:
                width = findCountColumnWidth;
                break;
            case ColumnID::REPLACE_COUNT:
                width = replaceCountColumnWidth;
                break;
            case ColumnID::FIND_TEXT:
                width = findColumnWidth;
                break;
            case ColumnID::REPLACE_TEXT:
                width = replaceColumnWidth;
                break;
            case ColumnID::COMMENTS:
                width = commentsColumnWidth;
                break;
            default:
                width = MIN_GENERAL_WIDTH_scaled; // Fallback to minimum width
                break;
            }
        }

        // Ensure the width is at least MIN_GENERAL_WIDTH_scaled
        width = std::max(width, MIN_GENERAL_WIDTH_scaled);

        break;
    }

    return width;
}

int MultiReplace::calcDynamicColWidth(const ResizableColWidths& widths) {
    // Calculate the total width of fixed columns
    int totalWidthFixedColumns = boxWidth_scaled + (checkMarkWidth_scaled * 5)
        + widths.findCountWidth
        + widths.replaceCountWidth
        + widths.deleteWidth;

    // Calculate the remaining width by subtracting fixed widths and visible static column widths
    int totalRemainingWidth = widths.listViewWidth
        - widths.margin  // Adjust for the vertical scrollbar margin
        - totalWidthFixedColumns
        - (findColumnLockedEnabled ? widths.findWidth : 0)
        - (replaceColumnLockedEnabled ? widths.replaceWidth : 0)
        - (commentsColumnLockedEnabled ? widths.commentsWidth : 0);

    // Calculate the number of dynamic columns
    int dynamicColumnCount = (!findColumnLockedEnabled)
        + (!replaceColumnLockedEnabled)
        + (!commentsColumnLockedEnabled && isCommentsColumnVisible);

    // Ensure at least one dynamic column is present for safety
    dynamicColumnCount = std::max(dynamicColumnCount, 1);

    // Calculate the width for each dynamic column
    int perColumnWidth = std::max(totalRemainingWidth / dynamicColumnCount, MIN_GENERAL_WIDTH_scaled);

    return perColumnWidth; // Return the calculated width for each dynamic column
}

void MultiReplace::updateListViewAndColumns() {
    // Retrieve control information
    const ControlInfo& listCtrlInfo = ctrlMap[IDC_REPLACE_LIST];
    HWND listView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    // Retrieve column widths
    findCountColumnWidth = getColumnWidth(ColumnID::FIND_COUNT);
    replaceCountColumnWidth = getColumnWidth(ColumnID::REPLACE_COUNT);
    commentsColumnWidth = getColumnWidth(ColumnID::COMMENTS);
    deleteButtonColumnWidth = getColumnWidth(ColumnID::DELETE_BUTTON);
    findColumnWidth = getColumnWidth(ColumnID::FIND_TEXT);
    replaceColumnWidth = getColumnWidth(ColumnID::REPLACE_TEXT);

    // Prepare the ResizableColWidths struct for dynamic width calculations
    ResizableColWidths widths = {
        listView,
        listCtrlInfo.cx,
        (isFindCountVisible) ? findCountColumnWidth : 0,
        (isReplaceCountVisible) ? replaceCountColumnWidth : 0,
        findColumnWidth,
        replaceColumnWidth,
        (isCommentsColumnVisible) ? commentsColumnWidth : 0,
        (isDeleteButtonVisible) ? deleteButtonColumnWidth : 0,
        GetSystemMetrics(SM_CXVSCROLL)
    };

    // Calculate dynamic column width
    int perColumnWidth = calcDynamicColWidth(widths);

    // Disable redraw to prevent flickering during updates
    SendMessage(listView, WM_SETREDRAW, FALSE, 0);

    // Update widths of dynamic columns
    if (columnIndices[ColumnID::FIND_TEXT] != -1) {
        int width = findColumnLockedEnabled ? findColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::FIND_TEXT], width);
    }
    if (columnIndices[ColumnID::REPLACE_TEXT] != -1) {
        int width = replaceColumnLockedEnabled ? replaceColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::REPLACE_TEXT], width);
    }
    if (columnIndices[ColumnID::COMMENTS] != -1) {
        int width = commentsColumnLockedEnabled ? commentsColumnWidth : perColumnWidth;
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::COMMENTS], width);
    }

    if (columnIndices[ColumnID::FIND_COUNT] != -1) {
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::FIND_COUNT], findCountColumnWidth);
    }
    if (columnIndices[ColumnID::REPLACE_COUNT] != -1) {
        ListView_SetColumnWidth(listView, columnIndices[ColumnID::REPLACE_COUNT], replaceCountColumnWidth);
    }

    // Resize the ListView control to the updated width and height
    updateListViewFrame();

    // Re-enable redraw after updates
    SendMessage(listView, WM_SETREDRAW, TRUE, 0);
}

void MultiReplace::updateListViewItem(size_t index) {
    if (index >= replaceListData.size()) return;

    const ReplaceItemData& item = replaceListData[index];

    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::FIND_TEXT], const_cast<LPWSTR>(item.findText.c_str()));
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::REPLACE_TEXT], const_cast<LPWSTR>(item.replaceText.c_str()));
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::COMMENTS], const_cast<LPWSTR>(item.comments.c_str()));

    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::WHOLE_WORD], item.wholeWord ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::MATCH_CASE], item.matchCase ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::USE_VARIABLES], item.useVariables ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::EXTENDED], item.extended ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::REGEX], item.regex ? L"\u2714" : L"");
    ListView_SetItemText(_replaceListView, static_cast<int>(index), columnIndices[ColumnID::SELECTION], item.isEnabled ? L"\u25A0" : L"\u2610");

    ListView_RedrawItems(_replaceListView, static_cast<int>(index), static_cast<int>(index));
}

void MultiReplace::updateListViewTooltips() {
    if (!_replaceListView) return;

    // when disabled: ensure header tooltip is gone
    if (!tooltipsEnabled) {
        if (_hHeaderTooltip) { DestroyWindow(_hHeaderTooltip); _hHeaderTooltip = nullptr; }
        return;
    }

    HWND hwndHeader = ListView_GetHeader(_replaceListView);
    if (!hwndHeader) return;

    if (_hHeaderTooltip) { DestroyWindow(_hHeaderTooltip); _hHeaderTooltip = nullptr; }

    _hHeaderTooltip = CreateHeaderTooltip(hwndHeader);

    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::WHOLE_WORD], LM.getW(L"tooltip_header_whole_word"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::MATCH_CASE], LM.getW(L"tooltip_header_match_case"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::USE_VARIABLES], LM.getW(L"tooltip_header_use_variables"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::EXTENDED], LM.getW(L"tooltip_header_extended"));
    AddHeaderTooltip(_hHeaderTooltip, hwndHeader, columnIndices[ColumnID::REGEX], LM.getW(L"tooltip_header_regex"));
}

void MultiReplace::handleCopyBack(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }

    // Copy the data from the selected item back to the source interfaces
    ReplaceItemData& itemData = replaceListData[pnmia->iItem];

    // Update the controls directly
    SetWindowTextW(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText.c_str());
    SetWindowTextW(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText.c_str());
    SendMessageW(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, itemData.wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, itemData.matchCase ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, itemData.useVariables ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_NORMAL_RADIO), BM_SETCHECK, (!itemData.regex && !itemData.extended) ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_EXTENDED_RADIO), BM_SETCHECK, itemData.extended ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(_hSelf, IDC_REGEX_RADIO), BM_SETCHECK, itemData.regex ? BST_CHECKED : BST_UNCHECKED, 0);

    // Enforce UI consistency: when regex is on, whole-word must be off in UI
    if (itemData.regex) {
        SendMessageW(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, BST_UNCHECKED, 0);
    }
    // Apply enable/disable rules (disables the checkbox if regex is on)
    setUIElementVisibility();
}

void MultiReplace::shiftListItem(const Direction& direction) {
    std::vector<size_t> selectedIndices;
    int i = -1;

    // Collect selected indices
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(i);
    }

    if (selectedIndices.empty()) {
        showStatusMessage(LM.get(L"status_no_rows_selected_to_shift"), MessageStatus::Error);
        return;
    }

    // Pass the selected indices to moveItemsInReplaceList
    if (!moveItemsInReplaceList(selectedIndices, direction)) {
        return;
    }

    // Deselect all items
    for (int j = 0; j < ListView_GetItemCount(_replaceListView); ++j) {
        ListView_SetItemState(_replaceListView, j, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Re-select the shifted items
    for (size_t index : selectedIndices) {
        ListView_SetItemState(_replaceListView, static_cast<int>(index), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Show status message when rows are successfully shifted
    showStatusMessage(LM.get(L"status_rows_shifted", { std::to_wstring(selectedIndices.size()) }), MessageStatus::Success);

    // Show Statistics
    showListFilePath();
}

void MultiReplace::handleDeletion(NMITEMACTIVATE* pnmia) {

    if (pnmia == nullptr || static_cast<size_t>(pnmia->iItem) >= replaceListData.size()) {
        return;
    }
    // Store the index of the item to be deleted
    size_t itemIndex = static_cast<size_t>(pnmia->iItem);

    // Prepare a vector containing the index of the item to remove
    std::vector<size_t> indicesToRemove = { itemIndex };

    // Delegate the deletion and undo/redo logic to removeItemsFromReplaceList
    removeItemsFromReplaceList(indicesToRemove);

    // Update Header if there might be any changes
    updateHeaderSelection();

    InvalidateRect(_replaceListView, nullptr, TRUE);

    showStatusMessage(LM.get(L"status_one_line_deleted"), MessageStatus::Success);
}

void MultiReplace::deleteSelectedLines() {
    // Collect selected indices
    std::vector<size_t> selectedIndices;
    int i = -1;
    while ((i = ListView_GetNextItem(_replaceListView, i, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(static_cast<size_t>(i));
    }

    if (selectedIndices.empty()) {
        showStatusMessage(LM.get(L"status_no_rows_selected_to_delete"), MessageStatus::Error);
        return;
    }

    // Ensure indices are valid before calling removeItemsFromReplaceList
    for (size_t index : selectedIndices) {
        if (index >= replaceListData.size()) {
            showStatusMessage(LM.get(L"status_invalid_indices"), MessageStatus::Error);
            return;
        }
    }

    // Call the removeItemsFromReplaceList function
    removeItemsFromReplaceList(selectedIndices);

    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select the next available line
    size_t nextIndexToSelect = selectedIndices.back() < replaceListData.size()
        ? selectedIndices.back()
        : replaceListData.size() - 1;
    if (nextIndexToSelect < replaceListData.size()) {
        ListView_SetItemState(_replaceListView, static_cast<int>(nextIndexToSelect), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    }

    // Update Header if there might be any changes
    updateHeaderSelection();

    // Show status message
    showStatusMessage(LM.get(L"status_lines_deleted", { std::to_wstring(selectedIndices.size()) }), MessageStatus::Success);
}

void MultiReplace::sortReplaceListData(int columnID) {
    // Check if the column is sortable
    if (columnID != ColumnID::FIND_COUNT &&
        columnID != ColumnID::REPLACE_COUNT &&
        columnID != ColumnID::FIND_TEXT &&
        columnID != ColumnID::REPLACE_TEXT &&
        columnID != ColumnID::COMMENTS) {
        return;
    }

    // Preserve the selection
    auto selectedIDs = getSelectedRows();

    // Ensure IDs are assigned
    for (auto& item : replaceListData) {
        if (item.id == 0) {
            item.id = generateUniqueID();
        }
    }

    // Capture the original order and previous sort state
    std::vector<size_t> originalOrder;
    for (const auto& item : replaceListData) {
        originalOrder.push_back(item.id);
    }
    auto previousColumnSortOrder = columnSortOrder;

    // Determine the new sort direction (3-state cycle)
    SortDirection direction = SortDirection::Ascending;
    auto it = columnSortOrder.find(columnID);
    if (it != columnSortOrder.end()) {
        if (it->second == SortDirection::Ascending) {
            direction = SortDirection::Descending;
        }
        else if (it->second == SortDirection::Descending) {
            direction = SortDirection::Unsorted;
        }
    }

    // Update the column sort order
    columnSortOrder.clear();
    if (direction != SortDirection::Unsorted) {
        columnSortOrder[columnID] = direction;
    }

    // Perform the sorting
    std::sort(replaceListData.begin(), replaceListData.end(),
        [this, columnID, direction]
        (const ReplaceItemData& a, const ReplaceItemData& b) -> bool
        {
            // Unsorted = sort by ID (insertion order)
            if (direction == SortDirection::Unsorted) {
                return a.id < b.id;
            }

            switch (columnID)
            {
            case ColumnID::FIND_COUNT:
                return direction == SortDirection::Ascending ? a.findCount < b.findCount : a.findCount > b.findCount;
            case ColumnID::REPLACE_COUNT:
                return direction == SortDirection::Ascending ? a.replaceCount < b.replaceCount : a.replaceCount > b.replaceCount;
            case ColumnID::FIND_TEXT:
            {
                int cmp = lstrcmpiW(a.findText.c_str(), b.findText.c_str());
                return direction == SortDirection::Ascending ? (cmp < 0) : (cmp > 0);
            }
            case ColumnID::REPLACE_TEXT:
            {
                int cmp = lstrcmpiW(a.replaceText.c_str(), b.replaceText.c_str());
                return direction == SortDirection::Ascending ? (cmp < 0) : (cmp > 0);
            }
            case ColumnID::COMMENTS:
            {
                int cmp = lstrcmpiW(a.comments.c_str(), b.comments.c_str());
                return direction == SortDirection::Ascending ? (cmp < 0) : (cmp > 0);
            }
            default:
                return false;
            }
        });

    // Capture the new order after sorting
    std::vector<size_t> newOrder;
    for (const auto& item : replaceListData) {
        newOrder.push_back(item.id);
    }

    // Update the UI and restore the selection
    updateHeaderSortDirection();
    ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);
    selectRows(selectedIDs);

    // Delegate undo/redo creation
    sortItemsInReplaceList(originalOrder, newOrder, previousColumnSortOrder, columnID, direction);
}

std::vector<size_t> MultiReplace::getSelectedRows() {
    std::vector<size_t> selectedIDs;
    int index = -1; // Use int to properly handle -1 case
    while ((index = ListView_GetNextItem(_replaceListView, index, LVNI_SELECTED)) != -1) {
        if (index >= 0 && static_cast<size_t>(index) < replaceListData.size()) {
            selectedIDs.push_back(replaceListData[index].id);
        }
    }
    return selectedIDs;
}

size_t MultiReplace::generateUniqueID() {
    static size_t currentID = 0;
    return ++currentID;
}

void MultiReplace::selectRows(const std::vector<size_t>& selectedIDs) {
    // Deselect all items
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Reselect rows based on IDs
    for (size_t i = 0; i < replaceListData.size(); ++i) {
        if (std::find(selectedIDs.begin(), selectedIDs.end(), replaceListData[i].id) != selectedIDs.end()) {
            ListView_SetItemState(_replaceListView, i, LVIS_SELECTED, LVIS_SELECTED);
        }
    }
}

void MultiReplace::handleCopyToListButton() {
    ReplaceItemData itemData;

    itemData.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
    itemData.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

    itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
    itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
    itemData.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
    itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
    itemData.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);

    insertReplaceListItem(itemData);

    // Add the entered text to the combo box history
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText);
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText);

    // Enable the ListView accordingly
    useListEnabled = true;
    updateUseListState(true);
    adjustWindowSize();
    ShowWindow(GetDlgItem(_hSelf, IDC_PATH_DISPLAY), SW_SHOW);
    ShowWindow(GetDlgItem(_hSelf, IDC_STATS_DISPLAY), SW_SHOW);
}

void MultiReplace::resetCountColumns() {
    // Reset the find and replace count columns in the list data
    for (auto& itemData : replaceListData) {
        itemData.findCount = -1;
        itemData.replaceCount = -1;
    }

    // Update the list view to immediately reflect the changes
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);
}

void MultiReplace::updateCountColumns(const size_t itemIndex, const int findCount, int replaceCount)
{
    if (itemIndex >= replaceListData.size()) return;
    ReplaceItemData& itemData = replaceListData[itemIndex];

    if (findCount == -2) {
        itemData.findCount = -1;
    }
    else if (findCount != -1) {
        itemData.findCount = findCount;
    }

    if (replaceCount != -1) {
        itemData.replaceCount = replaceCount;
    }
}

void MultiReplace::clearList() {
    // Check for unsaved changes before clearing the list
    if (checkForUnsavedChanges() == IDCANCEL) {
        return;
    }

    // Clear the replace list data vector
    replaceListData.clear();

    // Update the ListView to reflect the cleared list
    ListView_SetItemCountEx(_replaceListView, 0, LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);

    // Reset listFilePath to an empty string
    listFilePath.clear();

    // Call showListFilePath to update the UI with the cleared file path
    showListFilePath();

    // Set the original list hash to a default value when list is cleared
    originalListHash = 0;
}

void MultiReplace::refreshUIListView()
{
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);
}

void MultiReplace::handleColumnVisibilityToggle(UINT menuId) {
    // Toggle the corresponding visibility flag AND update ConfigManager
    // This ensures the Config Dialog reflects changes made here.
    switch (menuId) {
    case IDM_TOGGLE_FIND_COUNT:
        isFindCountVisible = !isFindCountVisible;
        CFG.writeInt(L"ListColumns", L"FindCountVisible", isFindCountVisible ? 1 : 0);
        break;
    case IDM_TOGGLE_REPLACE_COUNT:
        isReplaceCountVisible = !isReplaceCountVisible;
        CFG.writeInt(L"ListColumns", L"ReplaceCountVisible", isReplaceCountVisible ? 1 : 0);
        break;
    case IDM_TOGGLE_COMMENTS:
        isCommentsColumnVisible = !isCommentsColumnVisible;
        CFG.writeInt(L"ListColumns", L"CommentsVisible", isCommentsColumnVisible ? 1 : 0);
        break;
    case IDM_TOGGLE_DELETE:
        isDeleteButtonVisible = !isDeleteButtonVisible;
        CFG.writeInt(L"ListColumns", L"DeleteButtonVisible", isDeleteButtonVisible ? 1 : 0);
        break;
    default:
        return; // Unhandled menu ID
    }

    // IMPORTANT: Save to disk immediately.
    // The ConfigDialog reloads from disk when opening. If we don't save here,
    // opening the ConfigDialog would revert these changes to the old disk state.
    CFG.save(L"");

    // Recreate the ListView columns to reflect the changes
    HWND listView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    createListViewColumns();

    // Refresh the ListView (if necessary)
    InvalidateRect(listView, nullptr, TRUE);
}

ColumnID MultiReplace::getColumnIDFromIndex(int columnIndex) const {
    auto it = std::find_if(
        columnIndices.begin(),
        columnIndices.end(),
        [columnIndex](const auto& pair) { return pair.second == columnIndex; });

    return (it != columnIndices.end()) ? it->first : ColumnID::INVALID;
}

int MultiReplace::getColumnIndexFromID(ColumnID columnID) const {
    auto it = std::find_if(
        columnIndices.begin(),
        columnIndices.end(),
        [columnID](const auto& pair) { return pair.first == columnID; });

    return (it != columnIndices.end()) ? it->second : -1; // Return -1 if not found
}

void MultiReplace::onPathDisplayDoubleClick()
{
    if (listFilePath.empty())
        return;

    std::wstring cmdParam = L"/select,\"" + listFilePath + L"\"";
    ShellExecuteW(nullptr, L"open", L"explorer.exe", cmdParam.c_str(), nullptr, SW_SHOWNORMAL);
}

void MultiReplace::updateHeaderSelection() {
    bool anySelected = false;
    allSelected = !replaceListData.empty();

    // Check if any or all items in the replaceListData vector are selected
    for (const auto& itemData : replaceListData) {
        if (itemData.isEnabled) {
            anySelected = true;
        }
        else {
            allSelected = false;
        }
    }

    // Initialize the LVCOLUMN structure
    LVCOLUMN lvc = {};
    lvc.mask = LVCF_TEXT;

    // Determine the symbol to show in the header
    if (allSelected) {
        lvc.pszText = L"\u25A0"; // Ballot box with check
    }
    else if (anySelected) {
        lvc.pszText = L"\u25A3"; // Black square containing small white square
    }
    else {
        lvc.pszText = L"\u2610"; // Ballot box without check
    }

    // Update the Selection column header dynamically, if it's enabled
    if (columnIndices[ColumnID::SELECTION] != -1) {
        ListView_SetColumn(_replaceListView, columnIndices[ColumnID::SELECTION], &lvc);
    }

}

void MultiReplace::updateHeaderSortDirection() {
    const wchar_t* ascendingSymbol = L" ▲";
    const wchar_t* descendingSymbol = L" ▼";
    const wchar_t* lockedSymbol = L" 🔒"; // Lock symbol for locked columns

    // Iterate through all columns in columnIndices
    for (const auto& [columnID, columnIndex] : columnIndices) {
        // Skip columns that are not visible (columnIndex == -1)
        if (columnIndex == -1) {
            continue;
        }

        // Only update headers for sortable columns
        if (columnID != ColumnID::FIND_COUNT &&
            columnID != ColumnID::REPLACE_COUNT &&
            columnID != ColumnID::FIND_TEXT &&
            columnID != ColumnID::REPLACE_TEXT &&
            columnID != ColumnID::COMMENTS) {
            continue;
        }

        // Get the base header text
        std::wstring headerText;
        switch (columnID) {
        case ColumnID::FIND_COUNT:
            headerText = LM.get(L"header_find_count");
            break;
        case ColumnID::REPLACE_COUNT:
            headerText = LM.get(L"header_replace_count");
            break;
        case ColumnID::FIND_TEXT:
            headerText = LM.get(L"header_find");
            // Append lock symbol if the FIND_TEXT column is locked
            if (findColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::REPLACE_TEXT:
            headerText = LM.get(L"header_replace");
            // Append lock symbol if the REPLACE_TEXT column is locked
            if (replaceColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        case ColumnID::COMMENTS:
            headerText = LM.get(L"header_comments");
            if (commentsColumnLockedEnabled) {
                headerText += lockedSymbol;
            }
            break;
        default:
            continue; // Skip if it's not a sortable column
        }

        // Append sort symbol if the column is currently sorted
        auto sortIt = columnSortOrder.find(columnID);
        if (sortIt != columnSortOrder.end()) {
            SortDirection direction = sortIt->second;
            if (direction == SortDirection::Ascending) {
                headerText += ascendingSymbol;
            }
            else if (direction == SortDirection::Descending) {
                headerText += descendingSymbol;
            }
            // Do not append any symbol if direction is Unsorted
        }

        // Prepare the LVCOLUMN structure for updating the header
        LVCOLUMN lvc = {};
        lvc.mask = LVCF_TEXT;
        lvc.pszText = const_cast<LPWSTR>(headerText.c_str());

        // Update the column header using the correct index
        ListView_SetColumn(_replaceListView, columnIndex, &lvc);
    }
}

void MultiReplace::showListFilePath()
{
    HWND hPathDisplay = GetDlgItem(_hSelf, IDC_PATH_DISPLAY);
    HWND hStatsDisplay = GetDlgItem(_hSelf, IDC_STATS_DISPLAY);
    HWND hListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);

    if (!hPathDisplay || !hListView)
        return;

    HDC hDC = GetDC(hPathDisplay);
    HFONT hFont = reinterpret_cast<HFONT>(SendMessage(hPathDisplay, WM_GETFONT, 0, 0));
    SelectObject(hDC, hFont);

    // Get ListView width (X-Position bleibt relativ zur Liste)
    RECT rcListView;
    GetWindowRect(hListView, &rcListView);
    MapWindowPoints(nullptr, _hSelf, reinterpret_cast<LPPOINT>(&rcListView), 2);
    int listWidth = rcListView.right - rcListView.left;
    int listX = rcListView.left;

    const int spacing = sx(10);

    const ControlInfo& pathInfo = ctrlMap[IDC_PATH_DISPLAY];
    int fieldY = pathInfo.y;
    int fieldHeight = pathInfo.cy;

    int statsWidth = 0;

    if (listStatisticsEnabled && hStatsDisplay)
    {
        // Gather statistics
        int totalItems = static_cast<int>(replaceListData.size());
        int selectedItems = ListView_GetSelectedCount(_replaceListView);
        int activatedCount = static_cast<int>(std::count_if(replaceListData.begin(), replaceListData.end(),
            [](const ReplaceItemData& item) { return item.isEnabled; }));
        int focusedRow = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
        int displayRow = (selectedItems > 0 && focusedRow != -1) ? (focusedRow + 1) : 0;

        // Compose stats string
        std::wstring statsString = L"A:" + std::to_wstring(activatedCount)
            + L"  L:" + std::to_wstring(totalItems)
            + L"  |  R:" + std::to_wstring(displayRow)
            + L"  S:" + std::to_wstring(selectedItems);

        // Measure stats string width
        SIZE sz;
        GetTextExtentPoint32(hDC, statsString.c_str(), static_cast<int>(statsString.size()), &sz);
        statsWidth = sz.cx + sx(5); // padding

        // Position stats field
        int statsX = listX + listWidth - statsWidth;
        MoveWindow(hStatsDisplay, statsX, fieldY, statsWidth, fieldHeight, TRUE);
        SetWindowTextW(hStatsDisplay, statsString.c_str());
        ShowWindow(hStatsDisplay, SW_SHOW);
    }
    else if (hStatsDisplay)
    {
        // Pragmatic solution: Set width to zero instead of hiding
        statsWidth = 0;
        MoveWindow(hStatsDisplay, listX + listWidth, fieldY, 0, fieldHeight, TRUE);
        SetWindowTextW(hStatsDisplay, L"");
        ShowWindow(hStatsDisplay, SW_HIDE);
    }

    // Adjust path field to use remaining space
    int pathWidth = listWidth - statsWidth - (listStatisticsEnabled ? spacing : 0);
    pathWidth = std::max(pathWidth, 0);
    MoveWindow(hPathDisplay, listX, fieldY, pathWidth, fieldHeight, TRUE);

    // Update path display text
    std::wstring shortenedPath = getShortenedFilePath(listFilePath, pathWidth, hDC);
    SetWindowTextW(hPathDisplay, shortenedPath.c_str());

    ReleaseDC(hPathDisplay, hDC);

    // Immediate redraw
    InvalidateRect(hPathDisplay, nullptr, TRUE);
    UpdateWindow(hPathDisplay);
    if (hStatsDisplay) {
        InvalidateRect(hStatsDisplay, nullptr, TRUE);
        UpdateWindow(hStatsDisplay);
    }
}

#pragma endregion


#pragma region ListView Dialog

void MultiReplace::showColumnVisibilityMenu(HWND hWnd, POINT pt) {
    // Create a popup menu
    HMENU hMenu = CreatePopupMenu();

    // Add menu items with checkmarks based on current visibility
    AppendMenu(hMenu, MF_STRING | (isFindCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_FIND_COUNT, LM.getLPCW(L"header_find_count"));
    AppendMenu(hMenu, MF_STRING | (isReplaceCountVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_REPLACE_COUNT, LM.getLPCW(L"header_replace_count"));
    AppendMenu(hMenu, MF_STRING | (isCommentsColumnVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_COMMENTS, LM.getLPCW(L"header_comments"));
    AppendMenu(hMenu, MF_STRING | (isDeleteButtonVisible ? MF_CHECKED : MF_UNCHECKED), IDM_TOGGLE_DELETE, LM.getLPCW(L"header_delete_button"));

    // Display the menu at the specified location
    TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, hWnd, NULL);

    // Destroy the menu after use
    DestroyMenu(hMenu);
}

#pragma endregion


#pragma region UI Settings

void MultiReplace::onTooltipsToggled(bool enable)
{
    if (!instance) return;

    if (!enable)
    {
        destroyAllTooltipWindows();
        return;
    }

    rebuildAllTooltips();
}

void MultiReplace::destroyAllTooltipWindows()
{
    if (!_hSelf) return;

    // destroy known singletons
    if (_hHeaderTooltip) { DestroyWindow(_hHeaderTooltip);        _hHeaderTooltip = nullptr; }
    if (_hUseListButtonTooltip) { DestroyWindow(_hUseListButtonTooltip); _hUseListButtonTooltip = nullptr; }

    const DWORD tid = GetCurrentThreadId();
    EnumThreadWindows(tid,
        [](HWND hwnd, LPARAM lParam)->BOOL
        {
            auto pThis = reinterpret_cast<MultiReplace*>(lParam);
            if (!pThis || !pThis->_hSelf) return TRUE;

            wchar_t cls[64]{};
            GetClassNameW(hwnd, cls, 64);
            if (_wcsicmp(cls, TOOLTIPS_CLASS) != 0)
                return TRUE;

            // only tooltips owned by this panel (or its children)
            HWND hOwner = GetWindow(hwnd, GW_OWNER);
            if (hOwner != pThis->_hSelf && !IsChild(pThis->_hSelf, hOwner))
                return TRUE;

            // keep the "(?)" filter help tooltip if present among tools
            HWND hFilterHelp = GetDlgItem(pThis->_hSelf, IDC_FILTER_HELP);
            if (hFilterHelp) {
                TOOLINFO ti{}; ti.cbSize = sizeof(ti);
                for (int i = 0; SendMessage(hwnd, TTM_ENUMTOOLS, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(&ti)); ++i) {
                    if ((ti.uFlags & TTF_IDISHWND) && (HWND)ti.uId == hFilterHelp) {
                        return TRUE; // keep this tooltip window
                    }
                }
            }

            DestroyWindow(hwnd);
            return TRUE;
        },
        reinterpret_cast<LPARAM>(this));
}

void MultiReplace::rebuildAllTooltips()
{
    if (!_hSelf) return;

    // always remove leftovers first
    destroyAllTooltipWindows();

    if (!tooltipsEnabled)
        return;

    // Rebuild control tooltips from ctrlMap
    for (const auto& pair : ctrlMap)
    {
        const int ctrlId = pair.first;
        const auto& info = pair.second;
        if (!info.tooltipText || !info.tooltipText[0]) continue;

        HWND hCtrl = GetDlgItem(_hSelf, ctrlId);
        if (!hCtrl) continue;

        HWND hwndTooltip = CreateWindowEx(
            0,
            TOOLTIPS_CLASS,
            nullptr,
            WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON | TTS_NOPREFIX,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            _hSelf,
            nullptr,
            hInstance,
            nullptr
        );
        if (!hwndTooltip) continue;

        // Apply dark mode theme to tooltip
        if (NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle)) {
            SetWindowTheme(hwndTooltip, L"DarkMode_Explorer", nullptr);
        }

        DWORD maxWidth = (ctrlId == IDC_FILTER_HELP) ? 200 : 0;
        SendMessage(hwndTooltip, TTM_SETMAXTIPWIDTH, 0, maxWidth);

        TOOLINFO ti{};
        ti.cbSize = sizeof(TOOLINFO);
        ti.hwnd = _hSelf;
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        ti.uId = (UINT_PTR)hCtrl;
        ti.lpszText = const_cast<LPWSTR>(info.tooltipText);

        SendMessage(hwndTooltip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));
        SendMessage(hwndTooltip, TTM_ACTIVATE, TRUE, 0);
    }

    // Update special tooltips
    updateUseListState(false);
    updateListViewTooltips();
}

bool MultiReplace::isUseListEnabled() const {
    return useListEnabled;
}

bool MultiReplace::isTwoButtonsModeEnabled() const {
    // Check if window handle is valid
    if (!_hSelf || !::IsWindow(_hSelf)) return false;
    return (::IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED);
}

#pragma endregion


#pragma region Contextmenu List

void MultiReplace::toggleBooleanAt(int itemIndex, ColumnID columnID) {
    if (itemIndex < 0 || itemIndex >= static_cast<int>(replaceListData.size())) {
        return; // invalid row
    }

    ReplaceItemData originalData = replaceListData[itemIndex];
    ReplaceItemData newData = originalData;

    switch (columnID) {
    case ColumnID::SELECTION:
        newData.isEnabled = !newData.isEnabled;
        break;

    case ColumnID::WHOLE_WORD:
        // Ping-pong: if regex is on, turn it off, then toggle whole-word
        if (originalData.regex) {
            newData.regex = false;
        }
        newData.wholeWord = !originalData.wholeWord;
        break;

    case ColumnID::MATCH_CASE:
        newData.matchCase = !newData.matchCase;
        break;

    case ColumnID::USE_VARIABLES:
        newData.useVariables = !newData.useVariables;
        break;

    case ColumnID::EXTENDED:
        newData.extended = !newData.extended;
        if (newData.extended) {
            // Extended and Regex are mutually exclusive
            newData.regex = false;
        }
        break;

    case ColumnID::REGEX:
        newData.regex = !newData.regex;
        if (newData.regex) {
            // Regex disables Whole Word and excludes Extended
            newData.wholeWord = false;
            newData.extended = false;
        }
        break;

    default:
        return;
    }

    // Apply change (keeps Undo/Redo behavior)
    modifyItemInReplaceList(static_cast<size_t>(itemIndex), newData);

    if (columnID == ColumnID::SELECTION)
        updateHeaderSelection();
}

void MultiReplace::editTextAt(int itemIndex, ColumnID columnID) {
    // Determine column index and validate
    int column = getColumnIndexFromID(columnID);
    if (column == -1)
        return;

    // Suppress hover text and disable tooltips
    isHoverTextSuppressed = true;
    DWORD extendedStyle = ListView_GetExtendedListViewStyle(_replaceListView);
    extendedStyle &= ~LVS_EX_INFOTIP;
    ListView_SetExtendedListViewStyle(_replaceListView, extendedStyle);

    // Get subitem rectangle directly from ListView (handles scroll position automatically)
    // Note: ListView_GetSubItemRect works correctly for column >= 1
    // Editable columns (FIND_TEXT, REPLACE_TEXT, COMMENTS) are always at index >= 1
    RECT subItemRect;
    ListView_GetSubItemRect(_replaceListView, itemIndex, column, LVIR_BOUNDS, &subItemRect);

    int correctedX = subItemRect.left;
    int correctedY = subItemRect.top;
    int columnWidth = subItemRect.right - subItemRect.left;
    int editHeight = subItemRect.bottom - subItemRect.top;

    // Button dimensions
    const int EXPAND_BTN_WIDTH = 20;
    const int EXPAND_BTN_Y_OFFSET = -1;
    const int EXPAND_BTN_HEIGHT_EXTRA = 2;

    int buttonWidth = sx(EXPAND_BTN_WIDTH);
    int buttonHeight = editHeight + EXPAND_BTN_HEIGHT_EXTRA;

    // Adjust edit control width to reserve space for the button
    int editWidth = columnWidth - buttonWidth;

    // Create multi-line edit control with vertical scrollbar
    hwndEdit = CreateWindowEx(
        0,
        L"EDIT",
        L"",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN,
        correctedX,
        correctedY,
        editWidth,
        editHeight,
        _replaceListView,
        nullptr,
        (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE),
        NULL
    );

    // Create an expand/collapse toggle button next to the edit control
    hwndExpandBtn = CreateWindowEx(
        0,
        L"BUTTON",
        L"↓", // Indicator for expand/collapse
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        correctedX + editWidth,
        correctedY + EXPAND_BTN_Y_OFFSET,
        buttonWidth,
        buttonHeight,
        _replaceListView,
        reinterpret_cast<HMENU>(ID_EDIT_EXPAND_BUTTON),
        (HINSTANCE)GetWindowLongPtr(_hSelf, GWLP_HINSTANCE),
        NULL
    );

    // Set the _hBoldFont2 to the expand/collapse button
    if (font(FontRole::Bold2)) {
        SendMessage(hwndExpandBtn, WM_SETFONT, reinterpret_cast<WPARAM>(font(FontRole::Bold2)), TRUE);
    }

    // Set the initial text of the edit control
    const std::wstring* src = nullptr;
    if (itemIndex >= 0 && itemIndex < static_cast<int>(replaceListData.size())) {
        switch (columnID) {
        case ColumnID::FIND_TEXT:    src = &replaceListData[itemIndex].findText;    break;
        case ColumnID::REPLACE_TEXT: src = &replaceListData[itemIndex].replaceText; break;
        case ColumnID::COMMENTS:     src = &replaceListData[itemIndex].comments;    break;
        default: break;
        }
    }
    SetWindowTextW(hwndEdit, (src ? src->c_str() : L""));

    // Get the ListView font and set it for the Edit control
    HFONT hListViewFont = reinterpret_cast<HFONT>(SendMessage(_replaceListView, WM_GETFONT, 0, 0));
    if (hListViewFont) {
        SendMessage(hwndEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hListViewFont), TRUE);
    }

    // Focus the edit control and select all its text
    SetFocus(hwndEdit);
    SendMessage(hwndEdit, EM_SETSEL, 0, -1);

    // Subclass the edit control for custom keyboard handling
    SetWindowSubclass(hwndEdit, EditControlSubclassProc, 1, (DWORD_PTR)this);

    // Store the editing state
    _editingItemIndex = itemIndex;
    _editingColumnIndex = column;
    _editingColumnID = columnID;
    _editIsExpanded = false; // Initial state is collapsed
}

void MultiReplace::closeEditField(bool commitChanges)
{
    if (!hwndEdit) return;

    if (commitChanges &&
        _editingColumnID != ColumnID::INVALID &&
        _editingItemIndex >= 0 &&
        _editingItemIndex < static_cast<int>(replaceListData.size()))
    {
        // Read edited text dynamically
        int len = GetWindowTextLengthW(hwndEdit);
        std::wstring newText;
        if (len > 0) {
            newText.resize(static_cast<size_t>(len) + 1, L'\0');
            int written = GetWindowTextW(hwndEdit, &newText[0], len + 1);
            if (written < 0) written = 0;
            if (!newText.empty() && newText.back() == L'\0') newText.pop_back();
            newText.resize(static_cast<size_t>(written));
        }
        else {
            newText.clear();
        }

        // Prepare old/new data
        ReplaceItemData originalData = replaceListData[_editingItemIndex];
        ReplaceItemData newData = originalData;
        bool hasChanged = false;

        switch (_editingColumnID) {
        case ColumnID::FIND_TEXT:
            if (originalData.findText != newText) {
                newData.findText = std::move(newText);
                hasChanged = true;
            }
            break;

        case ColumnID::REPLACE_TEXT:
            if (originalData.replaceText != newText) {
                newData.replaceText = std::move(newText);
                hasChanged = true;
            }
            break;

        case ColumnID::COMMENTS:
            if (originalData.comments != newText) {
                newData.comments = std::move(newText);
                hasChanged = true;
            }
            break;

        default:
            break;
        }

        if (hasChanged) {
            modifyItemInReplaceList(_editingItemIndex, newData);
        }
    }

    // Tear down edit controls (unchanged)
    DestroyWindow(hwndEdit);
    hwndEdit = nullptr;
    if (hwndExpandBtn && IsWindow(hwndExpandBtn)) {
        DestroyWindow(hwndExpandBtn);
        hwndExpandBtn = nullptr;
    }
    _editIsExpanded = false;
    _editingItemIndex = -1;
    _editingColumnIndex = -1;
    _editingColumnID = ColumnID::INVALID;

    if (isHoverTextSuppressed) {
        isHoverTextSuppressed = false;
        if (isHoverTextEnabled) {
            DWORD es = ListView_GetExtendedListViewStyle(_replaceListView);
            es |= LVS_EX_INFOTIP;
            ListView_SetExtendedListViewStyle(_replaceListView, es);
        }
    }
}

LRESULT CALLBACK MultiReplace::EditControlSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    // Access the instance of MultiReplace
    MultiReplace* pThis = reinterpret_cast<MultiReplace*>(dwRefData);

    switch (msg) {
    case WM_KEYDOWN:
        if (wParam == VK_ESCAPE || wParam == VK_TAB) {
            pThis->closeEditField(true);
            RemoveWindowSubclass(hwnd, EditControlSubclassProc, uIdSubclass);
            return 0;
        }
        break;

    case WM_KILLFOCUS:
    {
        HWND newFocus = GetFocus();
        if (newFocus == hwndExpandBtn) {
            return 0;
        }

        pThis->closeEditField(true);
        RemoveWindowSubclass(hwnd, EditControlSubclassProc, uIdSubclass);
        return 0;
    }

    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK MultiReplace::ListViewSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    // Retrieve the 'this' pointer from the safe reference data, not from GWLP_USERDATA.
    MultiReplace* pThis = reinterpret_cast<MultiReplace*>(dwRefData);

    // Defensive check: If the pointer is somehow invalid, fall back to the default procedure.
    if (!pThis || !IsWindow(pThis->_hSelf)) {
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }

    // Handle WM_NCDESTROY to remove the subclass before the window is finally destroyed.
    if (msg == WM_NCDESTROY) {
        RemoveWindowSubclass(hwnd, &MultiReplace::ListViewSubclassProc, uIdSubclass);
        return DefSubclassProc(hwnd, msg, wParam, lParam);
    }
    switch (msg) {
    case WM_VSCROLL:
    case WM_MOUSEWHEEL:
    case WM_HSCROLL: {
        // Close the edit control if it exists and is visible
        if (pThis->hwndEdit && IsWindow(pThis->hwndEdit)) {
            DestroyWindow(pThis->hwndEdit);
            pThis->hwndEdit = nullptr;
        }
        // Allow the default list view procedure to handle standard scrolling
        break;
    }
    case WM_NOTIFY: {
        NMHDR* pnmhdr = reinterpret_cast<NMHDR*>(lParam);
        // Check notifications from header
        if (pnmhdr->hwndFrom == ListView_GetHeader(hwnd)) {
            int code = static_cast<int>(static_cast<short>(pnmhdr->code));

            if (code == HDN_ITEMCHANGED) {
                pThis->updateListViewTooltips(); //Updating Positions of Column Tooltips
            }

            // Handle right-click (NM_RCLICK) on the header
            if (code == NM_RCLICK) {
                // Get the current cursor position
                POINT ptScreen;
                GetCursorPos(&ptScreen);

                // Show the column visibility menu at the cursor position
                pThis->showColumnVisibilityMenu(pThis->_hSelf, ptScreen);

                return TRUE; // Indicate that the message has been handled
            }

            if (code == HDN_DIVIDERDBLCLICK) {
                NMHEADER* phdn = reinterpret_cast<NMHEADER*>(lParam);

                // Identify the clicked column based on phdn->iItem
                int clickedColumn = phdn->iItem;

                // Lock or unlock the column based on the clicked divider
                if (clickedColumn == pThis->columnIndices[ColumnID::FIND_TEXT]) {
                    pThis->findColumnLockedEnabled = !pThis->findColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }
                else if (clickedColumn == pThis->columnIndices[ColumnID::REPLACE_TEXT]) {
                    pThis->replaceColumnLockedEnabled = !pThis->replaceColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }
                else if (clickedColumn == pThis->columnIndices[ColumnID::COMMENTS]) {
                    pThis->commentsColumnLockedEnabled = !pThis->commentsColumnLockedEnabled;
                    pThis->updateHeaderSortDirection();
                }

                return TRUE; // Indicate that the message has been handled
            }

            // These values are derived from the HDN_ITEMCHANGEDW and HDN_ITEMCHANGEDA constants:
            // HDN_ITEMCHANGEDW = 0U - 300U - 21 = -321
            // HDN_ITEMCHANGEDA = 0U - 300U - 1  = -301
            // The constants are not used directly to avoid unsigned arithmetic overflow warnings.
            if (code == (int(0) - 300 - 21) || code == (int(0) - 300 - 1)) {
                // If there is an active edit control, destroy it when the header is changed
                if (pThis->hwndEdit && IsWindow(pThis->hwndEdit)) {
                    DestroyWindow(pThis->hwndEdit);
                    pThis->hwndEdit = nullptr;
                }

                // Set a timer to defer the tooltip update, ensuring the column resize is complete
                SetTimer(hwnd, 1, 100, nullptr);  // Timer ID 1, 100ms delay

            }
        }
        break;
    }

    case WM_MOUSEMOVE: {

        if (!pThis->isHoverTextEnabled || pThis->isHoverTextSuppressed) {
            return CallWindowProc(pThis->originalListViewProc, hwnd, msg, wParam, lParam);
        }

        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);

        LVHITTESTINFO hitTestInfo = {};
        hitTestInfo.pt = pt;

        int hitResult = ListView_HitTest(hwnd, &hitTestInfo);
        if (hitResult != -1) {
            int currentRow = hitTestInfo.iItem;
            int currentSubItem = hitTestInfo.iSubItem;

            // Update only if the row, sub-item, or mouse position has significantly changed
            if (currentRow != pThis->lastTooltipRow || currentSubItem != pThis->lastTooltipSubItem ||
                abs(pt.x - pThis->lastMouseX) > 5 || abs(pt.y - pThis->lastMouseY) > 5) {

                pThis->lastTooltipRow = currentRow;
                pThis->lastTooltipSubItem = currentSubItem;
                pThis->lastMouseX = pt.x;
                pThis->lastMouseY = pt.y;

                // Temporarily disable LVS_EX_INFOTIP to force tooltip update
                DWORD extendedStyle = ListView_GetExtendedListViewStyle(hwnd);
                ListView_SetExtendedListViewStyle(hwnd, extendedStyle & ~LVS_EX_INFOTIP);

                // Set a timer to re-enable LVS_EX_INFOTIP after 10ms
                SetTimer(hwnd, 1, 10, NULL);
            }
        }
        break;
    }

    case WM_SYSKEYDOWN:
    {
        if ((GetKeyState(VK_MENU) & 0x8000) && wParam == VK_UP) {
            int iItem = ListView_GetNextItem(hwnd, -1, LVNI_SELECTED);
            if (iItem >= 0) { NMITEMACTIVATE nmia{}; nmia.iItem = iItem; pThis->handleCopyBack(&nmia); }
            return 0;
        }
        // Optional fallback for Alt+E / Alt+D
        if (GetKeyState(VK_MENU) & 0x8000) {
            if (wParam == 'E') { pThis->setSelections(true, ListView_GetSelectedCount(hwnd) > 0);  return 0; }
            if (wParam == 'D') { pThis->setSelections(false, ListView_GetSelectedCount(hwnd) > 0);  return 0; }
        }
        break;
    }

    case WM_TIMER: {
        if (wParam == 1) { // Tooltip re-enable timer   
            KillTimer(hwnd, 1); // Kill the timer first to prevent it from firing again
            if (!pThis->isHoverTextEnabled || pThis->isHoverTextSuppressed) {
                return 0;
            }
            // Re-enable LVS_EX_INFOTIP
            DWORD extendedStyle = ListView_GetExtendedListViewStyle(hwnd);
            ListView_SetExtendedListViewStyle(hwnd, extendedStyle | LVS_EX_INFOTIP);
        }
        break;
    }

    case WM_COMMAND: {
        // Check if the source is our expand button
        WORD wId = LOWORD(wParam);
        WORD wEvent = HIWORD(wParam);

        if (wId == ID_EDIT_EXPAND_BUTTON && wEvent == BN_CLICKED) {
            // user clicked the expand button
            pThis->toggleEditExpand();
            return 0;
        }
        break;
    }

    case WM_UPDATE_FOCUS:
    {
        pThis->showListFilePath();
        return 0;
    }

    default:
        break;
    }

    // For all unhandled messages, call the default subclass procedure.
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

void MultiReplace::toggleEditExpand()
{
    if (!hwndEdit || !hwndExpandBtn)
        return;

    // Button dimensions (must match editTextAt)
    const int EXPAND_BTN_WIDTH = 20;
    const int EXPAND_BTN_Y_OFFSET = -1;
    const int EXPAND_BTN_HEIGHT_EXTRA = 2;

    // Get current position and size of the edit control in ListView coordinates
    RECT rc;
    GetWindowRect(hwndEdit, &rc);
    POINT ptLT = { rc.left, rc.top };
    POINT ptRB = { rc.right, rc.bottom };
    MapWindowPoints(nullptr, _replaceListView, &ptLT, 1);
    MapWindowPoints(nullptr, _replaceListView, &ptRB, 1);

    int curWidth = ptRB.x - ptLT.x;
    int curHeight = ptRB.y - ptLT.y;

    // Calculate new height (expand or collapse)
    int newHeight;
    if (_editIsExpanded) {
        newHeight = curHeight / editFieldSize; // Collapse
        SetWindowText(hwndExpandBtn, L"↓");
    }
    else {
        newHeight = curHeight * editFieldSize; // Expand
        SetWindowText(hwndExpandBtn, L"↑");
    }
    _editIsExpanded = !_editIsExpanded;

    // Set the _hBoldFont2 to the expand/collapse button
    SendMessage(hwndExpandBtn, WM_SETFONT, reinterpret_cast<WPARAM>(font(FontRole::Bold2)), TRUE);

    // Update position and size of edit control and button
    MoveWindow(hwndEdit, ptLT.x, ptLT.y, curWidth, newHeight, TRUE);
    MoveWindow(hwndExpandBtn, ptLT.x + curWidth, ptLT.y + EXPAND_BTN_Y_OFFSET, sx(EXPAND_BTN_WIDTH), newHeight + EXPAND_BTN_HEIGHT_EXTRA, TRUE);

    // Bring controls to the top and ensure edit has focus
    SetWindowPos(hwndEdit, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    SetWindowPos(hwndExpandBtn, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    SetFocus(hwndEdit);
}

void MultiReplace::createContextMenu(HWND hwnd, POINT ptScreen, MenuState state) {
    HMENU hMenu = CreatePopupMenu();
    if (hMenu) {
        AppendMenu(hMenu, MF_STRING | (state.canUndo ? MF_ENABLED : MF_GRAYED), IDM_UNDO, LM.get(L"ctxmenu_undo").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canRedo ? MF_ENABLED : MF_GRAYED), IDM_REDO, LM.get(L"ctxmenu_redo").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_CUT_LINES_TO_CLIPBOARD, LM.get(L"ctxmenu_cut").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_COPY_LINES_TO_CLIPBOARD, LM.get(L"ctxmenu_copy").c_str());
        AppendMenu(hMenu, MF_STRING | (state.canPaste ? MF_ENABLED : MF_GRAYED), IDM_PASTE_LINES_FROM_CLIPBOARD, LM.get(L"ctxmenu_paste").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_SELECT_ALL, LM.get(L"ctxmenu_select_all").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.canEdit ? MF_ENABLED : MF_GRAYED), IDM_EDIT_VALUE, LM.get(L"ctxmenu_edit").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection ? MF_ENABLED : MF_GRAYED), IDM_DELETE_LINES, LM.get(L"ctxmenu_delete").c_str());
        AppendMenu(hMenu, MF_STRING, IDM_ADD_NEW_LINE, LM.get(L"ctxmenu_add_new_line").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.clickedOnItem ? MF_ENABLED : MF_GRAYED), IDM_COPY_DATA_TO_FIELDS, LM.get(L"ctxmenu_transfer_to_input_fields").c_str());
        AppendMenu(hMenu, MF_STRING | (state.listNotEmpty ? MF_ENABLED : MF_GRAYED), IDM_EXPORT_DATA, LM.get(L"ctxmenu_export_data").c_str());
        AppendMenu(hMenu, MF_STRING | (state.listNotEmpty ? MF_ENABLED : MF_GRAYED), IDM_SEARCH_IN_LIST, LM.get(L"ctxmenu_search_in_list").c_str());
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allEnabled ? MF_ENABLED : MF_GRAYED), IDM_ENABLE_LINES, LM.get(L"ctxmenu_enable").c_str());
        AppendMenu(hMenu, MF_STRING | (state.hasSelection && !state.allDisabled ? MF_ENABLED : MF_GRAYED), IDM_DISABLE_LINES, LM.get(L"ctxmenu_disable").c_str());

        // Set Options Submenu
        HMENU hSetMenu = CreatePopupMenu();
        if (hSetMenu) {
            AppendMenu(hSetMenu, MF_STRING, IDM_SET_WHOLEWORD, LM.get(L"ctxmenu_opt_wholeword").c_str());
            AppendMenu(hSetMenu, MF_STRING, IDM_SET_MATCHCASE, LM.get(L"ctxmenu_opt_matchcase").c_str());
            AppendMenu(hSetMenu, MF_STRING, IDM_SET_VARIABLES, LM.get(L"ctxmenu_opt_variables").c_str());
            AppendMenu(hSetMenu, MF_STRING, IDM_SET_EXTENDED, LM.get(L"ctxmenu_opt_extended").c_str());
            AppendMenu(hSetMenu, MF_STRING, IDM_SET_REGEX, LM.get(L"ctxmenu_opt_regex").c_str());
            AppendMenu(hMenu, MF_POPUP | (state.hasSelection ? MF_ENABLED : MF_GRAYED),
                reinterpret_cast<UINT_PTR>(hSetMenu), LM.get(L"ctxmenu_set_options").c_str());
        }

        // Clear Options Submenu
        HMENU hClearMenu = CreatePopupMenu();
        if (hClearMenu) {
            AppendMenu(hClearMenu, MF_STRING, IDM_CLEAR_WHOLEWORD, LM.get(L"ctxmenu_opt_wholeword").c_str());
            AppendMenu(hClearMenu, MF_STRING, IDM_CLEAR_MATCHCASE, LM.get(L"ctxmenu_opt_matchcase").c_str());
            AppendMenu(hClearMenu, MF_STRING, IDM_CLEAR_VARIABLES, LM.get(L"ctxmenu_opt_variables").c_str());
            AppendMenu(hClearMenu, MF_STRING, IDM_CLEAR_EXTENDED, LM.get(L"ctxmenu_opt_extended").c_str());
            AppendMenu(hClearMenu, MF_STRING, IDM_CLEAR_REGEX, LM.get(L"ctxmenu_opt_regex").c_str());
            AppendMenu(hMenu, MF_POPUP | (state.hasSelection ? MF_ENABLED : MF_GRAYED),
                reinterpret_cast<UINT_PTR>(hClearMenu), LM.get(L"ctxmenu_clear_options").c_str());
        }

        TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hwnd, NULL);
        DestroyMenu(hMenu); // Cleans up submenus too
    }
}

MenuState MultiReplace::checkMenuConditions(POINT ptScreen) {
    MenuState state;

    // Convert screen coordinates to client coordinates
    POINT ptClient = ptScreen;
    ScreenToClient(_replaceListView, &ptClient);

    // Perform hit testing to determine the exact location of the click within the ListView
    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = ptClient;
    int hitTestResult = ListView_HitTest(_replaceListView, &hitInfo);
    state.clickedOnItem = (hitTestResult != -1);

    // Determine the clicked column
    int clickedColumn = -1;
    int totalWidth = 0;
    HWND header = ListView_GetHeader(_replaceListView);
    int columnCount = Header_GetItemCount(header);

    // Iterate through the columns and calculate the total width
    for (int i = 0; i < columnCount; ++i) {
        totalWidth += ListView_GetColumnWidth(_replaceListView, i);
        if (ptClient.x < totalWidth) {
            clickedColumn = i;
            break;
        }
    }

    // Map the clicked column to ColumnID, using a direct lookup in columnIndices
    int columnID = getColumnIDFromIndex(clickedColumn);

    // Enable editing if the clicked column is editable and currently visible
    state.canEdit = state.clickedOnItem && (
        columnID == ColumnID::SELECTION ||
        columnID == ColumnID::WHOLE_WORD ||
        columnID == ColumnID::MATCH_CASE ||
        columnID == ColumnID::USE_VARIABLES ||
        columnID == ColumnID::EXTENDED ||
        columnID == ColumnID::REGEX ||
        columnID == ColumnID::FIND_TEXT ||
        columnID == ColumnID::REPLACE_TEXT ||
        columnID == ColumnID::COMMENTS
        );

    // Rest of the checks remain the same
    state.listNotEmpty = (ListView_GetItemCount(_replaceListView) > 0);
    state.canPaste = canPasteFromClipboard();
    state.hasSelection = (ListView_GetSelectedCount(_replaceListView) > 0);

    unsigned int enabledCount = 0, disabledCount = 0;

    int itemIndex = -1;
    while ((itemIndex = ListView_GetNextItem(_replaceListView, itemIndex, LVNI_SELECTED)) != -1) {
        auto& itemData = replaceListData[itemIndex];
        if (itemData.isEnabled) {
            ++enabledCount;
        }
        else {
            ++disabledCount;
        }
    }

    state.allEnabled = (enabledCount == ListView_GetSelectedCount(_replaceListView));
    state.allDisabled = (disabledCount == ListView_GetSelectedCount(_replaceListView));

    state.canUndo = URM.canUndo();
    state.canRedo = URM.canRedo();

    return state;
}

void MultiReplace::performItemAction(POINT pt, ItemAction action) {
    // Get the horizontal scroll offset
    SCROLLINFO si = {};
    si.cbSize = sizeof(si);
    si.fMask = SIF_POS;
    GetScrollInfo(_replaceListView, SB_HORZ, &si);
    int scrollX = si.nPos;

    // Adjust point based on horizontal scroll
    POINT ptAdjusted = pt;
    ptAdjusted.x += scrollX;

    LVHITTESTINFO hitInfo = {};
    hitInfo.pt = ptAdjusted; // Use adjusted coordinates
    int hitTestResult = ListView_HitTest(_replaceListView, &hitInfo);

    // Calculate the clicked column based on the adjusted X coordinate
    int clickedColumn = -1;
    int totalWidth = 0;
    HWND header = ListView_GetHeader(_replaceListView);
    int columnCount = Header_GetItemCount(header);

    for (int i = 0; i < columnCount; ++i) {
        totalWidth += ListView_GetColumnWidth(_replaceListView, i);
        if (ptAdjusted.x < totalWidth) {
            clickedColumn = i;
            break; // Column found
        }
    }

    // Correct mapping from clickedColumn to ColumnID
    ColumnID columnID = getColumnIDFromIndex(clickedColumn);

    switch (action) {
    case ItemAction::Undo:
        URM.undo();
        showListFilePath();
        break;
    case ItemAction::Redo:
        URM.redo();
        showListFilePath();
        break;
    case ItemAction::Search:
        toggleListSearchBar();
        break;
    case ItemAction::Cut:
        copySelectedItemsToClipboard();
        deleteSelectedLines();
        showListFilePath();
        break;
    case ItemAction::Copy:
        copySelectedItemsToClipboard();
        break;
    case ItemAction::Paste:
        pasteItemsIntoList();
        showListFilePath();
        break;
    case ItemAction::Edit:
        if (columnID == ColumnID::FIND_TEXT ||
            columnID == ColumnID::REPLACE_TEXT ||
            columnID == ColumnID::COMMENTS) {
            editTextAt(hitTestResult, columnID);
        }
        else if (columnID == ColumnID::SELECTION ||
            columnID == ColumnID::WHOLE_WORD ||
            columnID == ColumnID::MATCH_CASE ||
            columnID == ColumnID::USE_VARIABLES ||
            columnID == ColumnID::EXTENDED ||
            columnID == ColumnID::REGEX) {
            toggleBooleanAt(hitTestResult, columnID);
            showListFilePath();
        }
        break;
    case ItemAction::Delete: {
        int selectedCount = ListView_GetSelectedCount(_replaceListView);
        std::wstring confirmationMessage;

        if (selectedCount == 1) {
            confirmationMessage = LM.get(L"msgbox_confirm_delete_single");
        }
        else if (selectedCount > 1) {
            confirmationMessage = LM.get(L"msgbox_confirm_delete_multiple", { std::to_wstring(selectedCount) });
        }

        int msgBoxID = MessageBox(nppData._nppHandle, confirmationMessage.c_str(), LM.get(L"msgbox_title_confirm").c_str(), MB_ICONWARNING | MB_YESNO);
        if (msgBoxID == IDYES) {
            deleteSelectedLines();
            showListFilePath();
        }
        break;
    }
    case ItemAction::Add: {
        int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
        if (insertPosition != -1) {
            ++insertPosition;
        }
        else {
            insertPosition = ListView_GetItemCount(_replaceListView);
        }
        ReplaceItemData newItem; // Default-initialized
        std::vector<ReplaceItemData> itemsToAdd = { newItem };
        addItemsToReplaceList(itemsToAdd, static_cast<size_t>(insertPosition));

        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);
        ListView_SetItemState(_replaceListView, insertPosition, LVIS_SELECTED, LVIS_SELECTED);
        ListView_EnsureVisible(_replaceListView, insertPosition, FALSE);
        showListFilePath();
        break;
    }
    }
}

void MultiReplace::copySelectedItemsToClipboard() {
    const int itemCount = ListView_GetItemCount(_replaceListView);
    const int columnCount = Header_GetItemCount(ListView_GetHeader(_replaceListView));
    std::wstring csvData;
    csvData.reserve(itemCount * columnCount * 16);

    for (int i = 0; i < itemCount; ++i) {
        if (ListView_GetItemState(_replaceListView, i, LVIS_SELECTED) & LVIS_SELECTED) {
            const ReplaceItemData& item = replaceListData[i];
            csvData += std::to_wstring(item.isEnabled) + L",";
            csvData += StringUtils::escapeCsvValue(item.findText) + L",";
            csvData += StringUtils::escapeCsvValue(item.replaceText) + L",";
            csvData += std::to_wstring(item.wholeWord) + L",";
            csvData += std::to_wstring(item.matchCase) + L",";
            csvData += std::to_wstring(item.useVariables) + L",";
            csvData += std::to_wstring(item.extended) + L",";
            csvData += std::to_wstring(item.regex) + L",";
            csvData += StringUtils::escapeCsvValue(item.comments);
            csvData += L'\n';
        }
    }

    if (csvData.empty()) return;

    if (OpenClipboard(nullptr)) {
        EmptyClipboard();
        const SIZE_T size = (csvData.size() + 1) * sizeof(wchar_t);
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
        if (hMem) {
            if (void* p = GlobalLock(hMem)) {
                memcpy(p, csvData.c_str(), size);
                GlobalUnlock(hMem);
                if (!SetClipboardData(CF_UNICODETEXT, hMem)) {
                    GlobalFree(hMem);
                }
            }
            else {
                GlobalFree(hMem);
            }
        }
        CloseClipboard();
    }
}

bool MultiReplace::canPasteFromClipboard() {
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT) || !OpenClipboard(nullptr)) {
        return false;
    }
    struct ClipboardGuard { ~ClipboardGuard() { CloseClipboard(); } } guard;

    HGLOBAL hData = GetClipboardData(CF_UNICODETEXT);
    if (!hData) {
        return false;
    }
    wchar_t* raw = static_cast<wchar_t*>(GlobalLock(hData));
    if (!raw) {
        return false;
    }
    std::wstring content{ raw };
    GlobalUnlock(hData);

    std::wistringstream stream(content);
    std::wstring line;
    while (std::getline(stream, line)) {
        if (line.empty()) {
            continue;
        }
        auto columns = StringUtils::parseCsvLine(line);
        if (columns.size() == 8 || columns.size() == 9) {
            return true;
        }
    }
    return false;
}

void MultiReplace::pasteItemsIntoList() {
    if (!OpenClipboard(nullptr)) return;

    struct ClipboardGuard { ~ClipboardGuard() { CloseClipboard(); } } guard;

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (!hData) return;

    auto pData = static_cast<const wchar_t*>(GlobalLock(hData));
    if (!pData) return;

    std::wstring content(pData);
    GlobalUnlock(hData);

    std::wstringstream contentStream(content);
    std::wstring line;
    std::vector<ReplaceItemData> itemsToInsert; // Collect items to insert

    // Determine insert position based on focused item or default to end if none is focused
    int insertPosition = ListView_GetNextItem(_replaceListView, -1, LVNI_FOCUSED);
    if (insertPosition != -1) {
        // Increase by one to insert after the focused item
        ++insertPosition;
    }
    else {
        // If no item is focused, consider the paste position as the end of the list
        insertPosition = ListView_GetItemCount(_replaceListView);
    }

    while (std::getline(contentStream, line)) {
        if (line.empty()) continue; // Skip empty lines

        std::vector<std::wstring> columns = StringUtils::parseCsvLine(line);

        // Check for proper column count before adding to the list
        if ((columns.size() != 8 && columns.size() != 9)) continue;

        ReplaceItemData item;
        try {
            item.isEnabled = std::stoi(columns[0]) != 0;
            item.findText = columns[1];
            item.replaceText = columns[2];
            item.wholeWord = std::stoi(columns[3]) != 0;
            item.matchCase = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended = std::stoi(columns[6]) != 0;
            item.regex = std::stoi(columns[7]) != 0;
            item.comments = (columns.size() == 9 ? columns[8] : L"");
        }
        catch (const std::exception&) {
            continue; // Silently ignore lines with conversion errors
        }

        itemsToInsert.push_back(item);
    }

    // No items were collected, so nothing to insert
    if (itemsToInsert.empty()) return;

    // Use addItemsToReplaceList to insert items at the specified position
    addItemsToReplaceList(itemsToInsert, static_cast<size_t>(insertPosition));

    // Deselect all items first
    ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED);

    // Select newly inserted items in the list view
    for (size_t i = 0; i < itemsToInsert.size(); ++i) {
        size_t idx = static_cast<size_t>(insertPosition) + i;
        ListView_SetItemState(_replaceListView, static_cast<int>(idx), LVIS_SELECTED, LVIS_SELECTED);
    }

    // Ensure the first inserted item is visible
    ListView_EnsureVisible(_replaceListView, insertPosition, FALSE);
}

int MultiReplace::searchInListData(int startIdx, const std::wstring& searchText, bool forward) {
    if (searchText.empty()) return -1;

    int listSize = static_cast<int>(replaceListData.size());
    if (listSize == 0) return -1;

    // Case-insensitive search text
    std::wstring searchLower = searchText;
    std::transform(searchLower.begin(), searchLower.end(), searchLower.begin(), ::towlower);

    int step = forward ? 1 : -1;
    int i = (startIdx < 0) ? (forward ? 0 : listSize - 1) : startIdx + step;

    for (int count = 0; count < listSize; ++count) {
        // Wrap around
        if (i >= listSize) i = 0;
        if (i < 0) i = listSize - 1;

        const auto& item = replaceListData[i];

        // Search in Find, Replace, and Comments columns (case-insensitive)
        auto containsSearch = [&searchLower](const std::wstring& text) {
            std::wstring textLower = text;
            std::transform(textLower.begin(), textLower.end(), textLower.begin(), ::towlower);
            return textLower.find(searchLower) != std::wstring::npos;
            };

        if (containsSearch(item.findText) || containsSearch(item.replaceText) || containsSearch(item.comments)) {
            return i;
        }

        i += step;
    }
    return -1;
}

void MultiReplace::toggleListSearchBar() {
    if (!useListEnabled) {
        return;
    }

    if (_listSearchBarVisible) {
        hideListSearchBar();
    }
    else {
        showListSearchBar();
    }
}

void MultiReplace::showListSearchBar() {
    if (_listSearchBarVisible) {
        SetFocus(GetDlgItem(_hSelf, IDC_LIST_SEARCH_COMBO));
        return;
    }

    _listSearchBarVisible = true;

    // Recalculate layout
    RECT rc;
    GetClientRect(_hSelf, &rc);
    positionAndResizeControls(rc.right, rc.bottom);

    // Explicitly shrink ListView
    const ControlInfo& listInfo = ctrlMap[IDC_REPLACE_LIST];
    SetWindowPos(_replaceListView, nullptr,
        listInfo.x, listInfo.y, listInfo.cx, listInfo.cy,
        SWP_NOZORDER | SWP_NOACTIVATE);

    // Adjust other controls
    moveAndResizeControls(false);

    // Show search bar controls
    HWND hCombo = GetDlgItem(_hSelf, IDC_LIST_SEARCH_COMBO);
    HWND hButton = GetDlgItem(_hSelf, IDC_LIST_SEARCH_BUTTON);
    HWND hClose = GetDlgItem(_hSelf, IDC_LIST_SEARCH_CLOSE);

    ShowWindow(hCombo, SW_SHOW);
    ShowWindow(hButton, SW_SHOW);
    ShowWindow(hClose, SW_SHOW);

    // Focus the combo
    SetFocus(hCombo);

    InvalidateRect(_hSelf, NULL, TRUE);
}

void MultiReplace::hideListSearchBar() {
    if (!_listSearchBarVisible) return;

    _listSearchBarVisible = false;

    // Hide search bar controls
    ShowWindow(GetDlgItem(_hSelf, IDC_LIST_SEARCH_COMBO), SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_LIST_SEARCH_BUTTON), SW_HIDE);
    ShowWindow(GetDlgItem(_hSelf, IDC_LIST_SEARCH_CLOSE), SW_HIDE);

    // Recalculate layout
    RECT rc;
    GetClientRect(_hSelf, &rc);
    positionAndResizeControls(rc.right, rc.bottom);

    // Explicitly enlarge ListView
    const ControlInfo& listInfo = ctrlMap[IDC_REPLACE_LIST];
    SetWindowPos(_replaceListView, nullptr,
        listInfo.x, listInfo.y, listInfo.cx, listInfo.cy,
        SWP_NOZORDER | SWP_NOACTIVATE);

    // Adjust other controls
    moveAndResizeControls(false);

    SetFocus(_replaceListView);
    InvalidateRect(_hSelf, NULL, TRUE);
}

void MultiReplace::findInList(bool forward) {
    HWND hCombo = GetDlgItem(_hSelf, IDC_LIST_SEARCH_COMBO);
    if (!hCombo) return;

    int len = GetWindowTextLength(hCombo);
    if (len == 0) return;

    std::wstring searchText(len + 1, L'\0');
    GetWindowText(hCombo, &searchText[0], len + 1);
    searchText.resize(len);

    // Add to history (avoid duplicates)
    int existingIndex = static_cast<int>(SendMessage(hCombo, CB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(searchText.c_str())));
    if (existingIndex != CB_ERR) {
        SendMessage(hCombo, CB_DELETESTRING, existingIndex, 0);
    }
    SendMessage(hCombo, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(searchText.c_str()));
    SetWindowText(hCombo, searchText.c_str());

    // Get current selection as start index
    int startIdx = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);

    int matchIdx = searchInListData(startIdx, searchText, forward);

    if (matchIdx != -1) {
        ListView_SetItemState(_replaceListView, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_SetItemState(_replaceListView, matchIdx, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(_replaceListView, matchIdx, FALSE);
        showStatusMessage(LM.get(L"status_found_in_list"), MessageStatus::Success);
    }
    else {
        showStatusMessage(LM.get(L"status_not_found_in_list"), MessageStatus::Error);
    }
}

void MultiReplace::jumpToNextMatchInEditor(size_t listIndex) {
    // 1. Basic validation
    if (listIndex >= replaceListData.size()) return;

    const auto& item = replaceListData[listIndex];

    // 2. Collect match positions from ResultDock hits that match this list entry
    //    This ensures Column-Scope filtering is respected (only actual search hits are used)
    struct MatchRange { Sci_Position start; Sci_Position length; int docLine; size_t hitIdx; };
    std::vector<MatchRange> ranges;

    ResultDock& dock = ResultDock::instance();
    const auto& allHits = dock.hits();

    // Get current document path for filtering
    wchar_t curPath[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(curPath));
    std::string curPathUtf8 = Encoding::wstringToUtf8(curPath);

    // Collect all matching hits from ResultDock (stored positions)
    for (size_t i = 0; i < allHits.size(); ++i)
    {
        const auto& hit = allHits[i];
        if (!pathsEqualUtf8(hit.fullPathUtf8, curPathUtf8))
            continue;

        // Check if this hit belongs to the requested list entry (by findText match)
        bool textMatch = (hit.findTextW == item.findText);
        if (!textMatch) {
            for (const auto& ft : hit.allFindTexts) {
                if (ft == item.findText) { textMatch = true; break; }
            }
        }
        if (!textMatch)
            continue;

        // Add primary position
        ranges.push_back({ hit.pos, hit.length, hit.docLine, i });

        // Add merged positions (if any)
        for (size_t j = 0; j < hit.allPositions.size(); ++j)
        {
            Sci_Position len = (j < hit.allLengths.size()) ? hit.allLengths[j] : hit.length;
            int line = (hit.docLine >= 0) ? hit.docLine : static_cast<int>(send(SCI_LINEFROMPOSITION, hit.allPositions[j], 0));
            ranges.push_back({ hit.allPositions[j], len, line, i });
        }
    }

    // Sort by position
    std::sort(ranges.begin(), ranges.end(),
        [](const MatchRange& a, const MatchRange& b) { return a.start < b.start; });

    // No live search fallback - navigation only works with validated Find All data
    if (ranges.empty()) {
        showStatusMessage(LM.get(L"status_no_results_linked"), MessageStatus::Error);
        return;
    }

    // 3. Determine anchor position
    Sci_Position anchorPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));

    if (anchorPos == 0) {
        auto cursorInfo = dock.getCurrentCursorHitInfo();
        if (cursorInfo.valid && cursorInfo.hitIndex < allHits.size()) {
            int anchorLine = allHits[cursorInfo.hitIndex].docLine;
            Sci_Position lineStart = static_cast<Sci_Position>(send(SCI_POSITIONFROMLINE, anchorLine, 0));
            if (lineStart > 0) {
                anchorPos = lineStart;
            }
        }
    }

    // 4. Find next range at or after anchor position
    size_t foundIdx = SIZE_MAX;
    bool wrapped = false;

    for (size_t i = 0; i < ranges.size(); ++i) {
        if (ranges[i].start >= anchorPos) {
            foundIdx = i;
            break;
        }
    }

    // 5. Wrap-around
    if (foundIdx == SIZE_MAX) {
        foundIdx = 0;
        wrapped = true;
    }

    // 6. Jump to the found range
    Sci_Position jumpPos = ranges[foundIdx].start;
    Sci_Position jumpLen = ranges[foundIdx].length;

    displayResultCentered(jumpPos, jumpPos + jumpLen, true);

    // 7. Sync ResultDock to the matching hit
    size_t hitIdx = ranges[foundIdx].hitIdx;
    if (hitIdx < allHits.size() && allHits[hitIdx].displayLineStart >= 0) {
        dock.scrollToHitAndHighlight(allHits[hitIdx].displayLineStart);
    }
    else if (!allHits.empty()) {
        // Fallback: find hit by line
        int jumpLine = ranges[foundIdx].docLine;
        for (size_t i = 0; i < allHits.size(); ++i) {
            const auto& hit = allHits[i];
            if (hit.docLine == jumpLine && hit.displayLineStart >= 0) {
                bool textMatch = (hit.findTextW == item.findText);
                if (!textMatch) {
                    for (const auto& ft : hit.allFindTexts) {
                        if (ft == item.findText) { textMatch = true; break; }
                    }
                }
                if (textMatch) {
                    dock.scrollToHitAndHighlight(hit.displayLineStart);
                    break;
                }
            }
        }
    }

    // 8. Status message
    size_t total = ranges.size();
    size_t current = foundIdx + 1;

    if (wrapped) {
        showStatusMessage(LM.get(L"status_wrapped_to_first_of",
            { std::to_wstring(current), std::to_wstring(total) }),
            MessageStatus::Info);
    }
    else {
        showStatusMessage(LM.get(L"status_match_position",
            { std::to_wstring(current), std::to_wstring(total) }),
            MessageStatus::Success);
    }
}

void MultiReplace::handleEditOnDoubleClick(int itemIndex, ColumnID columnID) {
    if (columnID == ColumnID::FIND_TEXT || columnID == ColumnID::REPLACE_TEXT || columnID == ColumnID::COMMENTS) {
        editTextAt(itemIndex, columnID);
    }
    else if (columnID == ColumnID::FIND_COUNT) {
        jumpToNextMatchInEditor(static_cast<size_t>(itemIndex));
    }
    else if (columnID == ColumnID::SELECTION ||
        columnID == ColumnID::WHOLE_WORD ||
        columnID == ColumnID::MATCH_CASE ||
        columnID == ColumnID::USE_VARIABLES ||
        columnID == ColumnID::EXTENDED ||
        columnID == ColumnID::REGEX) {
        toggleBooleanAt(itemIndex, columnID);
    }
}

#pragma endregion


#pragma region Dialog

INT_PTR CALLBACK MultiReplace::run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
    {
        ResultDock::setPerEntryColorsEnabled(true);  // Enable per-entry coloring for testing

        dpiMgr = new DPIManager(_hSelf);
        initializeWindowSize();

        pointerToScintilla();
        if (_hScintilla) {
            g_prevBufId = static_cast<int>(static_cast<UINT_PTR>(
                ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0)
                ));
        }

        ensureIndicatorContext();
        initTextMarkerIndicators();
        createFonts();
        initializeCtrlMap();
        applyFonts();
        applyThemePalette();
        loadSettings();
        updateTwoButtonsVisibility();
        initializeListView();
        initializeDragAndDrop();
        adjustWindowSize();

        // Activate Dark Mode
        ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(_hSelf));

        // Post a custom message to perform post-initialization tasks after the dialog is shown
        PostMessage(_hSelf, WM_POST_INIT, 0, 0);

        return TRUE;
    }

    case WM_POST_INIT:
    {
        checkForFileChangesAtStartup();
        return TRUE;
    }

    case WM_GETMINMAXINFO:
    {
        MINMAXINFO* pMMI = reinterpret_cast<MINMAXINFO*>(lParam);
        RECT adjustedSize = calculateMinWindowFrame(_hSelf);

        // Set minimum window width
        pMMI->ptMinTrackSize.x = adjustedSize.right;
        // Allow horizontal resizing up to a maximum value
        pMMI->ptMaxTrackSize.x = MAXLONG;

        if (useListEnabled) {
            // Set minimum window height
            pMMI->ptMinTrackSize.y = adjustedSize.bottom;
            // Allow vertical resizing up to a maximum value
            pMMI->ptMaxTrackSize.y = MAXLONG;
        }
        else {
            // Set fixed window height
            pMMI->ptMinTrackSize.y = adjustedSize.bottom;
            pMMI->ptMaxTrackSize.y = adjustedSize.bottom; // Fix the height when Use List is unchecked
        }

        return 0;
    }

    case WM_ACTIVATE:
    {
        if (_keepOnTopDuringBatch) {
            // Force foreground alpha during batch
            SetWindowTransparency(_hSelf, foregroundTransparency);
            return 0;
        }

        if (LOWORD(wParam) == WA_INACTIVE) {
            // The window loses focus
            SetWindowTransparency(_hSelf, backgroundTransparency); // Use the loaded value
        }
        else {
            // The window gains focus
            SetWindowTransparency(_hSelf, foregroundTransparency); // Use the loaded value
        }
        return 0;
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = reinterpret_cast<HDC>(wParam);
        HWND hwndStatic = reinterpret_cast<HWND>(lParam);

        if (hwndStatic == GetDlgItem(_hSelf, IDC_STATUS_MESSAGE)) {
            SetTextColor(hdcStatic, _statusMessageColor);
            SetBkMode(hdcStatic, TRANSPARENT);
            return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH)); // Return a brush handle
        }

        return FALSE;
    }

    case WM_DESTROY:
    {

        if (_replaceListView && originalListViewProc) {
            SetWindowLongPtr(_replaceListView, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalListViewProc));
        }

        saveSettings(); // Save any settings before destroying

        // Unregister Drag-and-Drop (COM-correct cleanup)
        if (dropTarget) {
            RevokeDragDrop(_replaceListView);  // Calls Release(): refCount 2→1
            dropTarget->Release();              // Calls Release(): refCount 1→0, deletes itself
            dropTarget = nullptr;
        }

        if (hwndEdit) {
            DestroyWindow(hwndEdit);
        }

        cleanupFonts();

        // Close the debug window if open
        if (hDebugWnd != nullptr) {
            RECT rect;
            if (GetWindowRect(hDebugWnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }
            PostMessage(hDebugWnd, WM_CLOSE, 0, 0);
            hDebugWnd = nullptr; // Reset the handle after closing
        }

        // Clean up DPIManager
        if (dpiMgr) {
            delete dpiMgr;
            dpiMgr = nullptr;
        }

        DestroyWindow(_hSelf); // Destroy the main window

        // Post a quit message to ensure the application terminates cleanly
        PostQuitMessage(0);

        return 0;
    }

    case WM_DPICHANGED:
    {
        if (dpiMgr) {
            dpiMgr->updateDPI(_hSelf);
        }
        return 0;
    }

    case WM_SIZE:
    {
        if (isWindowOpen) {
            // Force the edit control of the right mouse click to lose focus by setting focus to the main window
            if (isWindowOpen && hwndEdit && GetFocus() == hwndEdit) {
                HWND hwndListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
                SetFocus(hwndListView);
            }

            int newWidth = LOWORD(lParam);
            int newHeight = HIWORD(lParam);

            // Calculate Position for all Elements
            positionAndResizeControls(newWidth, newHeight);

            // Move and resize the List
            updateListViewAndColumns();

            // Move all Elements
            moveAndResizeControls(false);

            // Refresh UI and gripper by invalidating window
            InvalidateRect(_hSelf, nullptr, TRUE);

            if (useListEnabled) {
                // Update useListOnHeight with the new height
                RECT currentRect;
                GetWindowRect(_hSelf, &currentRect);
                int currentHeight = currentRect.bottom - currentRect.top;

                // Get the minimum allowed window height
                RECT minSize = calculateMinWindowFrame(_hSelf);
                int minHeight = minSize.bottom;

                // Ensure useListOnHeight is at least the minimum height
                useListOnHeight = std::max(currentHeight, minHeight);
            }
        }
        return 0;
    }

    case WM_NOTIFY:
    {
        NMHDR* pnmh = reinterpret_cast<NMHDR*>(lParam);

        if (pnmh->code == BCN_DROPDOWN && pnmh->hwndFrom == GetDlgItem(_hSelf, IDC_REPLACE_ALL_BUTTON))
        {
            // split-button menu for Replace All
            RECT rc; ::GetWindowRect(pnmh->hwndFrom, &rc);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_ALL_OPTION, LM.getW(L"split_menu_replace_all"));
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_IN_ALL_DOCS_OPTION, LM.getW(L"split_menu_replace_all_in_docs"));
            AppendMenu(hMenu, MF_STRING, ID_REPLACE_IN_FILES_OPTION, LM.getW(L"split_menu_replace_all_in_files"));
            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
            AppendMenu(hMenu, MF_STRING | (_debugModeEnabled ? MF_CHECKED : MF_UNCHECKED), ID_DEBUG_MODE_OPTION, LM.getW(L"split_menu_debug_mode"));
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, rc.left, rc.bottom, 0, _hSelf, NULL);
            DestroyMenu(hMenu);
            return TRUE;
        }

        if (pnmh->code == BCN_DROPDOWN && pnmh->hwndFrom == GetDlgItem(_hSelf, IDC_FIND_ALL_BUTTON))
        {
            // split-button menu for Find All
            RECT rc; ::GetWindowRect(pnmh->hwndFrom, &rc);
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_OPTION, LM.getW(L"split_menu_find_all"));
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_IN_ALL_DOCS_OPTION, LM.getW(L"split_menu_find_all_in_docs"));
            AppendMenu(hMenu, MF_STRING, ID_FIND_ALL_IN_FILES_OPTION, LM.getW(L"split_menu_find_all_in_files"));
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, rc.left, rc.bottom, 0, _hSelf, nullptr);
            DestroyMenu(hMenu);
            return TRUE;
        }

        if (pnmh->idFrom == IDC_REPLACE_LIST) {
            switch (pnmh->code) {
            case NM_CLICK:
            {
                NMITEMACTIVATE* pnmia = reinterpret_cast<NMITEMACTIVATE*>(lParam);
                int subItem = pnmia->iSubItem;
                int itemIndex = pnmia->iItem;

                // Ensure valid item index, as subItem 0 could refer to the first column or an invalid click area.
                // This prevents accidental actions when clicking outside valid list items.
                if (itemIndex >= 0 && itemIndex < static_cast<int>(replaceListData.size())) {
                    ColumnID columnID = getColumnIDFromIndex(subItem);

                    switch (columnID) {
                    case ColumnID::DELETE_BUTTON:
                        handleDeletion(pnmia);
                        break;

                    case ColumnID::SELECTION:
                    {
                        bool currentSelectionStatus = replaceListData[itemIndex].isEnabled;
                        setSelections(!currentSelectionStatus, true);
                    }
                    break;

                    default:
                        break;
                    }

                    showListFilePath();
                }
                return TRUE;
            }

            case NM_DBLCLK:
            {
                NMITEMACTIVATE* pnmia = reinterpret_cast<NMITEMACTIVATE*>(lParam);
                int itemIndex = pnmia->iItem;
                int clickedColumn = pnmia->iSubItem;

                if (itemIndex != -1 && clickedColumn != -1) {
                    if (doubleClickEditsEnabled) {
                        ColumnID columnID = getColumnIDFromIndex(clickedColumn);
                        handleEditOnDoubleClick(itemIndex, columnID);
                    }
                    else {
                        handleCopyBack(pnmia);
                    }
                }
                return TRUE;
            }

            case LVN_GETDISPINFO:
            {
                NMLVDISPINFO* plvdi = reinterpret_cast<NMLVDISPINFO*>(lParam);
                const int itemIndex = plvdi->item.iItem;
                const int subItem = plvdi->item.iSubItem;

                // Provide text only when requested
                if ((plvdi->item.mask & LVIF_TEXT) == 0)
                    return TRUE;

                if (itemIndex < 0 || itemIndex >= static_cast<int>(replaceListData.size()))
                {
                    if (plvdi->item.pszText && plvdi->item.cchTextMax > 0)
                        plvdi->item.pszText[0] = L'\0';
                    return TRUE;
                }

                const ReplaceItemData& d = replaceListData[itemIndex];

                // Fill text depending on which column is asking
                if (columnIndices[ColumnID::FIND_COUNT] != -1 && subItem == columnIndices[ColumnID::FIND_COUNT]) {
                    if (d.findCount >= 0)
                        (void)lstrcpynW(plvdi->item.pszText, std::to_wstring(d.findCount).c_str(), plvdi->item.cchTextMax);
                    else
                        plvdi->item.pszText[0] = L'\0';
                }
                else if (columnIndices[ColumnID::REPLACE_COUNT] != -1 && subItem == columnIndices[ColumnID::REPLACE_COUNT]) {
                    if (d.replaceCount >= 0)
                        (void)lstrcpynW(plvdi->item.pszText, std::to_wstring(d.replaceCount).c_str(), plvdi->item.cchTextMax);
                    else
                        plvdi->item.pszText[0] = L'\0';
                }
                else if (columnIndices[ColumnID::SELECTION] != -1 && subItem == columnIndices[ColumnID::SELECTION]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.isEnabled ? L"\u25A0" : L"\u2610", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::FIND_TEXT] != -1 && subItem == columnIndices[ColumnID::FIND_TEXT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.findText.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::REPLACE_TEXT] != -1 && subItem == columnIndices[ColumnID::REPLACE_TEXT]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.replaceText.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::WHOLE_WORD] != -1 && subItem == columnIndices[ColumnID::WHOLE_WORD]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.wholeWord ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::MATCH_CASE] != -1 && subItem == columnIndices[ColumnID::MATCH_CASE]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.matchCase ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::USE_VARIABLES] != -1 && subItem == columnIndices[ColumnID::USE_VARIABLES]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.useVariables ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::EXTENDED] != -1 && subItem == columnIndices[ColumnID::EXTENDED]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.extended ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::REGEX] != -1 && subItem == columnIndices[ColumnID::REGEX]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.regex ? L"\u2714" : L"", plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::COMMENTS] != -1 && subItem == columnIndices[ColumnID::COMMENTS]) {
                    (void)lstrcpynW(plvdi->item.pszText, d.comments.c_str(), plvdi->item.cchTextMax);
                }
                else if (columnIndices[ColumnID::DELETE_BUTTON] != -1 && subItem == columnIndices[ColumnID::DELETE_BUTTON]) {
                    (void)lstrcpynW(plvdi->item.pszText, L"\u2716", plvdi->item.cchTextMax);
                }
                else {
                    if (plvdi->item.pszText && plvdi->item.cchTextMax > 0)
                        plvdi->item.pszText[0] = L'\0';
                }
                return TRUE;
            }

            case LVN_COLUMNCLICK:
            {
                NMLISTVIEW* pnmv = reinterpret_cast<NMLISTVIEW*>(lParam);

                int clickedColumn = pnmv->iSubItem;

                // Map the clicked column index to ColumnID
                int columnID = getColumnIDFromIndex(clickedColumn);

                if (columnID == ColumnID::INVALID) {
                    // If no valid column is found, ignore the click
                    return TRUE;
                }

                if (columnID == ColumnID::SELECTION) {
                    setSelections(!allSelected);
                }
                else {
                    sortReplaceListData(columnID);
                }
                return TRUE;
            }

            case LVN_KEYDOWN:
            {
                LPNMLVKEYDOWN pnkd = reinterpret_cast<LPNMLVKEYDOWN>(pnmh);
                int iItem = -1;

                // For arrow keys and Page Up/Page Down, post a custom message to update the focus
                if (pnkd->wVKey == VK_UP || pnkd->wVKey == VK_DOWN ||
                    pnkd->wVKey == VK_PRIOR || pnkd->wVKey == VK_NEXT)
                {
                    PostMessage(_replaceListView, WM_UPDATE_FOCUS, 0, 0);
                    return TRUE;
                }

                // For non-arrow keys, proceed with existing focus update.
                PostMessage(_replaceListView, WM_SETFOCUS, 0, 0);

                // Handling keyboard shortcuts for menu actions
                if (GetKeyState(VK_CONTROL) & 0x8000) { // If Ctrl is pressed
                    switch (pnkd->wVKey) {
                    case 'Z': // Ctrl+Z for Undo
                        URM.undo();
                        break;
                    case 'Y': // Ctrl+Y for Redo
                        URM.redo();
                        break;
                    case 'F': // Ctrl+F for Search in List
                        performItemAction(_contextMenuClickPoint, ItemAction::Search);
                        break;
                    case 'X': // Ctrl+X for Cut
                        performItemAction(_contextMenuClickPoint, ItemAction::Cut);
                        break;
                    case 'C': // Ctrl+C for Copy
                        performItemAction(_contextMenuClickPoint, ItemAction::Copy);
                        break;
                    case 'V': // Ctrl+V for Paste
                        performItemAction(_contextMenuClickPoint, ItemAction::Paste);
                        break;
                    case 'A': // Ctrl+A for Select All
                        ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
                        showListFilePath();
                        break;
                    case 'I': // Ctrl+I for Adding new Line
                        performItemAction(_contextMenuClickPoint, ItemAction::Add);
                        break;
                    }
                }
                else {
                    switch (pnkd->wVKey) {
                    case VK_DELETE: // Delete key for deleting selected lines
                        performItemAction(_contextMenuClickPoint, ItemAction::Delete);
                        break;
                    case VK_F12: // F12 key
                    {
                        showDPIAndFontInfo();  // Show the DPI and font information
                        break;
                    }
                    case VK_SPACE: // Spacebar key
                        iItem = ListView_GetNextItem(_replaceListView, -1, LVNI_SELECTED);
                        if (iItem >= 0) {
                            // Get current selection status of the item
                            bool currentSelectionStatus = replaceListData[iItem].isEnabled;
                            // Set the selection status to its opposite
                            setSelections(!currentSelectionStatus, true);
                        }
                        break;
                    }
                }
                return TRUE;
            }
            }
        }
        return FALSE;
    }

    case WM_CONTEXTMENU:
    {
        // Check if the right-click is on the ListView (not the header)
        if ((HWND)wParam == _replaceListView) {
            POINT ptScreen;
            ptScreen.x = LOWORD(lParam);
            ptScreen.y = HIWORD(lParam);

            // Get the size of the virtual screen (the total width and height of all monitors)
            int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            // Check if the x-coordinates are outside the expected range and adjust
            if (ptScreen.x > virtualWidth) {
                ptScreen.x -= 65536; // Adjust for potential overflow
            }

            // Check if the y-coordinates are outside the expected range and adjust
            if (ptScreen.y > virtualHeight) {
                ptScreen.y -= 65536; // Adjust for potential overflow
            }

            _contextMenuClickPoint = ptScreen; // Store initial click point for later action determination
            ScreenToClient(_replaceListView, &_contextMenuClickPoint); // Convert to client coordinates for hit testing
            MenuState state = checkMenuConditions(ptScreen);
            createContextMenu(_hSelf, ptScreen, state); // Show context menu
            return TRUE;
        }
        return FALSE;
    }

    case WM_SHOWWINDOW:
    {
        if (wParam == TRUE) {
            std::wstring wstr = getSelectedText();

            // Set selected text in IDC_FIND_EDIT
            if (!wstr.empty()) {
                SetWindowTextW(GetDlgItem(_hSelf, IDC_FIND_EDIT), wstr.c_str());
            }
        }
        else {
            handleClearTextMarksButton();
            handleClearDelimiterState();
        }
        return 0;
    }

    case WM_PAINT:
    {
        drawGripper();
        return 0;
    }

    case WM_NCHITTEST:
    {
        // Enable resize via gripper area
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        ScreenToClient(_hSelf, &pt);
        RECT rc;
        GetClientRect(_hSelf, &rc);
        int gripperSize = sx(11);
        if (pt.x >= rc.right - gripperSize && pt.y >= rc.bottom - gripperSize)
        {
            SetWindowLongPtr(_hSelf, DWLP_MSGRESULT, HTBOTTOMRIGHT);
            return TRUE;
        }
        return FALSE;
    }

    case WM_DRAWITEM:
    {
        DRAWITEMSTRUCT* pdis = (DRAWITEMSTRUCT*)lParam;

        if (pdis->CtlID == IDC_STATUS_MESSAGE) {
            wchar_t buffer[256];
            GetWindowTextW(pdis->hwndItem, buffer, 256);

            SetTextColor(pdis->hDC, _statusMessageColor);
            SetBkMode(pdis->hDC, TRANSPARENT);

            RECT textRect = pdis->rcItem;
            DrawTextW(pdis->hDC, buffer, -1, &textRect, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            return TRUE;
        }
        else if (pdis->CtlID == IDC_FILTER_HELP) {

            wchar_t buffer[16];
            GetWindowTextW(pdis->hwndItem, buffer, 16);

            SetTextColor(pdis->hDC, _filterHelpColor);
            SetBkMode(pdis->hDC, TRANSPARENT);

            DrawTextW(pdis->hDC, buffer, -1, &pdis->rcItem, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            return TRUE;
        }

        return FALSE;
    }

    case WM_CTLCOLORDLG:
    {
        // Store the handle to the brush that the dialog will use to paint its background
        _hDlgBrush = reinterpret_cast<HBRUSH>(wParam);
        return (INT_PTR)_hDlgBrush;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {

        case IDC_PATH_DISPLAY:
        {
            if (HIWORD(wParam) == STN_DBLCLK)
            {
                onPathDisplayDoubleClick();
                return TRUE;
            }
            return FALSE;
        }

        case IDC_USE_VARIABLES_HELP:
        {
            auto n = SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, 0, 0);
            std::wstring path(n, 0);
            SendMessage(nppData._nppHandle, NPPM_GETPLUGINHOMEPATH, n + 1, reinterpret_cast<LPARAM>(path.data()));
            path += L"\\MultiReplace";
            const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
            std::wstring filename = isDark ? L"\\help_use_variables_dark.html" : L"\\help_use_variables_light.html";
            path += filename;
            ShellExecute(nullptr, L"open", path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            return TRUE;
        }

        case IDCANCEL:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            EndDialog(_hSelf, 0);
            _MultiReplace.display(false);
            return TRUE;
        }

        case IDC_2_BUTTONS_MODE:
        {
            // Check if the Find checkbox has been clicked
            if (HIWORD(wParam) == BN_CLICKED)
            {
                updateTwoButtonsVisibility();
                return TRUE;
            }
            return FALSE;
        }

        case IDC_REGEX_RADIO:
        {
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_NORMAL_RADIO:
        case IDC_EXTENDED_RADIO:
        {
            EnableWindow(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), TRUE);
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_ALL_TEXT_RADIO:
        {
            m_selectionScope.clear();
            setUIElementVisibility();
            handleClearDelimiterState();
            return TRUE;
        }

        case IDC_SELECTION_RADIO:
        {
            m_selectionScope.clear();
            setUIElementVisibility();
            handleClearDelimiterState();
            return TRUE;
        }

        case IDC_COLUMN_NUM_EDIT:
        case IDC_DELIMITER_EDIT:
        case IDC_QUOTECHAR_EDIT:
        case IDC_COLUMN_MODE_RADIO:
        {
            m_selectionScope.clear();
            CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_COLUMN_MODE_RADIO);
            setUIElementVisibility();
            return TRUE;
        }

        case IDC_COLUMN_SORT_ASC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortStateAndSort(SortDirection::Ascending);
                UpdateSortButtonSymbols();
            }
            return TRUE;
        }

        case IDC_COLUMN_SORT_DESC_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleSortStateAndSort(SortDirection::Descending);
                UpdateSortButtonSymbols();
            }
            return TRUE;
        }

        case IDC_COLUMN_DROP_BUTTON:
        {
            if (confirmColumnDeletion()) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleDeleteColumns();
                }
            }
            return TRUE;
        }

        case IDC_COLUMN_COPY_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleCopyColumnsToClipboard();
            }
            return TRUE;
        }

        case IDC_COLUMN_HIGHLIGHT_BUTTON:
        {
            int currentBufferID = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));

            if (!highlightedTabs.isHighlighted(currentBufferID)) {
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                if (columnDelimiterData.isValid()) {
                    handleHighlightColumnsInDocument();
                }
            }
            else {
                handleClearColumnMarks();
                showStatusMessage(LanguageManager::instance().get(L"status_column_marks_cleared"), MessageStatus::Success);
            }
            return TRUE;
        }

        case IDC_COLUMN_GRIDTABS_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleColumnGridTabsButton();
            }
            return TRUE;
        }

        case IDC_COLUMN_DUPLICATES_BUTTON:
        {
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            if (columnDelimiterData.isValid()) {
                handleDuplicatesButton();
            }
            return TRUE;
        }

        case IDC_USE_LIST_BUTTON:
        {
            useListEnabled = !useListEnabled;

            if (!useListEnabled) {
                // Closing: First hide controls, then resize
                if (_listSearchBarVisible) {
                    hideListSearchBar();
                }
                ShowWindow(GetDlgItem(_hSelf, IDC_PATH_DISPLAY), SW_HIDE);
                ShowWindow(GetDlgItem(_hSelf, IDC_STATS_DISPLAY), SW_HIDE);
                updateUseListState(true);
                adjustWindowSize();
            }
            else {
                // Opening: First resize, then show controls
                updateUseListState(true);
                adjustWindowSize();
                ShowWindow(GetDlgItem(_hSelf, IDC_PATH_DISPLAY), SW_SHOW);
                ShowWindow(GetDlgItem(_hSelf, IDC_STATS_DISPLAY), SW_SHOW);
            }
            return TRUE;
        }

        case IDC_SWAP_BUTTON:
        {
            std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            std::wstring replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);

            // Swap the content of the two text fields
            SetDlgItemTextW(_hSelf, IDC_FIND_EDIT, replaceText.c_str());
            SetDlgItemTextW(_hSelf, IDC_REPLACE_EDIT, findText.c_str());
            return TRUE;
        }

        case IDC_COPY_TO_LIST_BUTTON:
        {
            handleCopyToListButton();
            return TRUE;
        }

        case IDC_FIND_ALL_BUTTON:
        {
            CloseDebugWindow();
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);

            if (isFindAllInFiles)
                handleFindInFiles();
            else if (isFindAllInDocs)
                handleFindAllInDocsButton();
            else
                handleFindAllButton();

            return TRUE;
        }

        case ID_FIND_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getW(L"split_button_find_all"));
            isFindAllInDocs = false;
            isFindAllInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_FIND_ALL_IN_ALL_DOCS_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getW(L"split_button_find_all_in_docs"));
            isFindAllInDocs = true;
            isFindAllInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_FIND_ALL_IN_FILES_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_FIND_ALL_BUTTON, LM.getW(L"split_button_find_all_in_files"));
            isFindAllInDocs = false;
            isFindAllInFiles = true;
            updateFilesPanel();
            return TRUE;
        }

        case IDC_FIND_NEXT_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindNextButton();
            return TRUE;
        }

        case IDC_FIND_PREV_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleFindPrevButton();
            return TRUE;
        }

        case IDC_REPLACE_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceButton();
            return TRUE;
        }

        case IDC_REPLACE_ALL_SMALL_BUTTON:
        {
            CloseDebugWindow(); // Close the Lua debug window if it is open
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleReplaceAllButton();
            return TRUE;
        }

        case IDC_REPLACE_ALL_BUTTON:
        {
            CloseDebugWindow();

            if (isReplaceAllInDocs)
            {
                replaceAllInOpenedDocs();
            }
            else if (isReplaceInFiles) {
                std::wstring filter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
                std::wstring dir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);

                handleReplaceInFiles();
            }
            else
            {
                resetCountColumns();
                handleDelimiterPositions(DelimiterOperation::LoadAll);
                handleReplaceAllButton();
            }
            return TRUE;
        }

        case IDC_BROWSE_DIR_BUTTON:
        {
            handleBrowseDirectoryButton();
            return TRUE;
        }

        case IDC_MARK_MATCHES_BUTTON:
        case IDC_MARK_BUTTON:
        {
            resetCountColumns();
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            handleClearTextMarksButton();
            handleMarkMatchesButton();
            return TRUE;
        }

        case IDC_CLEAR_MARKS_BUTTON:
        {
            handleClearTextMarksButton();
            clearDuplicateMarks();
            showStatusMessage(LM.get(L"status_all_marks_cleared"), MessageStatus::Success);
            return TRUE;
        }

        case IDC_COPY_MARKED_TEXT_BUTTON:
        {
            handleCopyMarkedTextToClipboardButton();
            return TRUE;
        }

        case IDC_SAVE_AS_BUTTON:
        case IDC_SAVE_TO_CSV_BUTTON:
        {
            std::wstring filePath = promptSaveListToCsv();

            if (!filePath.empty()) {
                saveListToCsv(filePath, replaceListData);
            }

            return TRUE;
        }

        case IDC_SAVE_BUTTON:
        {
            if (!listFilePath.empty()) {
                saveListToCsv(listFilePath, replaceListData);
            }
            else {
                // If no file path is set, prompt the user to select a file path
                std::wstring filePath = promptSaveListToCsv();

                if (!filePath.empty()) {
                    saveListToCsv(filePath, replaceListData);
                }
            }
            return TRUE;
        }

        case IDC_LOAD_LIST_BUTTON:
        case IDC_LOAD_FROM_CSV_BUTTON:
        {
            std::wstring csvDescription = LM.get(L"filetype_csv");  // "CSV Files (*.csv)"
            std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {csvDescription, L"*.csv"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = LM.get(L"panel_load_list");
            std::wstring filePath = openFileDialog(false, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST, L"csv", L"");

            if (!filePath.empty()) {
                loadListFromCsv(filePath);
            }
            return TRUE;
        }

        case IDC_NEW_LIST_BUTTON:
        {
            clearList();
            showStatusMessage(LM.get(L"status_new_list_created"), MessageStatus::Success);
            return TRUE;
        }

        case IDC_UP_BUTTON:
        {
            shiftListItem(Direction::Up);
            return TRUE;
        }

        case IDC_DOWN_BUTTON:
        {
            shiftListItem(Direction::Down);
            return TRUE;
        }

        case IDC_EXPORT_BASH_BUTTON:
        {
            std::wstring bashDescription = LM.get(L"filetype_bash");  // "Bash Scripts (*.sh)"
            std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

            std::vector<std::pair<std::wstring, std::wstring>> filters = {
                {bashDescription, L"*.sh"},
                {allFilesDescription, L"*.*"}
            };

            std::wstring dialogTitle = LM.get(L"panel_export_to_bash");

            // Set a default filename if none is provided
            static int scriptCounter = 1;
            std::wstring defaultFileName = L"Replace_Script_" + std::to_wstring(++scriptCounter) + L".sh";

            // Open the file dialog with the default filename for bash scripts
            std::wstring filePath = openFileDialog(true, filters, dialogTitle.c_str(), OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, L"sh", defaultFileName);

            if (!filePath.empty()) {
                exportToBashScript(filePath);
            }
            return TRUE;
        }

        case ID_REPLACE_ALL_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getW(L"split_button_replace_all"));
            isReplaceAllInDocs = false;
            isReplaceInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_REPLACE_IN_ALL_DOCS_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getW(L"split_button_replace_all_in_docs"));
            isReplaceAllInDocs = true;
            isReplaceInFiles = false;
            updateFilesPanel();
            return TRUE;
        }

        case ID_REPLACE_IN_FILES_OPTION:
        {
            SetDlgItemText(_hSelf, IDC_REPLACE_ALL_BUTTON, LM.getW(L"split_button_replace_all_in_files"));
            isReplaceAllInDocs = false;
            isReplaceInFiles = true;
            updateFilesPanel();
            return TRUE;
        }

        case ID_DEBUG_MODE_OPTION:
        {
            _debugModeEnabled = !_debugModeEnabled;
            return TRUE;
        }

        case IDC_CANCEL_REPLACE_BUTTON:
        {
            _isCancelRequested = true;
            return TRUE;
        }

        case IDM_SEARCH_IN_LIST:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Search);
            return TRUE;
        }

        case IDM_UNDO: {
            performItemAction(_contextMenuClickPoint, ItemAction::Undo);
            return TRUE;
        }

        case IDM_REDO: {
            performItemAction(_contextMenuClickPoint, ItemAction::Redo);
            return TRUE;
        }

        case IDM_COPY_DATA_TO_FIELDS:
        {
            NMITEMACTIVATE nmia = {};
            nmia.iItem = ListView_HitTest(_replaceListView, &_contextMenuClickPoint);
            handleCopyBack(&nmia);
            return TRUE;
        }

        case IDM_EXPORT_DATA:
        {
            exportDataToClipboard();
            return TRUE;
        }

        case IDM_CUT_LINES_TO_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Cut);
            return TRUE;
        }

        case IDM_COPY_LINES_TO_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Copy);
            return TRUE;
        }

        case IDM_PASTE_LINES_FROM_CLIPBOARD:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Paste);
            return TRUE;
        }

        case IDM_EDIT_VALUE:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Edit);
            return TRUE;
        }

        case IDM_DELETE_LINES:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Delete);
            return TRUE;
        }

        case IDM_SELECT_ALL:
        {
            ListView_SetItemState(_replaceListView, -1, LVIS_SELECTED, LVIS_SELECTED);
            return TRUE;
        }

        case IDM_ENABLE_LINES:
        {
            setSelections(true, ListView_GetSelectedCount(_replaceListView) > 0);
            return TRUE;
        }

        case IDM_DISABLE_LINES:
        {
            setSelections(false, ListView_GetSelectedCount(_replaceListView) > 0);
            return TRUE;
        }

        // Set Options
        case IDM_SET_WHOLEWORD:
        {
            setOptionForSelection(SearchOption::WholeWord, true);
            return TRUE;
        }
        case IDM_SET_MATCHCASE:
        {
            setOptionForSelection(SearchOption::MatchCase, true);
            return TRUE;
        }
        case IDM_SET_VARIABLES:
        {
            setOptionForSelection(SearchOption::Variables, true);
            return TRUE;
        }
        case IDM_SET_EXTENDED:
        {
            setOptionForSelection(SearchOption::Extended, true);
            return TRUE;
        }
        case IDM_SET_REGEX:
        {
            setOptionForSelection(SearchOption::Regex, true);
            return TRUE;
        }

        // Clear Options
        case IDM_CLEAR_WHOLEWORD:
        {
            setOptionForSelection(SearchOption::WholeWord, false);
            return TRUE;
        }
        case IDM_CLEAR_MATCHCASE:
        {
            setOptionForSelection(SearchOption::MatchCase, false);
            return TRUE;
        }
        case IDM_CLEAR_VARIABLES:
        {
            setOptionForSelection(SearchOption::Variables, false);
            return TRUE;
        }
        case IDM_CLEAR_EXTENDED:
        {
            setOptionForSelection(SearchOption::Extended, false);
            return TRUE;
        }
        case IDM_CLEAR_REGEX:
        {
            setOptionForSelection(SearchOption::Regex, false);
            return TRUE;
        }

        case IDC_LIST_SEARCH_BUTTON:
            findInList(true);
            return TRUE;

        case IDC_LIST_SEARCH_CLOSE:
            hideListSearchBar();
            return TRUE;

        case IDM_TOGGLE_FIND_COUNT:
        case IDM_TOGGLE_REPLACE_COUNT:
        case IDM_TOGGLE_COMMENTS:
        case IDM_TOGGLE_DELETE:
        {
            // Toggle the visibility based on the menu selection
            handleColumnVisibilityToggle(LOWORD(wParam));
            return TRUE;
        }

        case IDM_ADD_NEW_LINE:
        {
            performItemAction(_contextMenuClickPoint, ItemAction::Add);
            return TRUE;
        }

        default:
            return FALSE;
        }
    }

    default:
        return FALSE;
    }
}

#pragma endregion


#pragma region Replace

void MultiReplace::replaceAllInOpenedDocs()
{
    if (MessageBox(
        nppData._nppHandle,
        LM.get(L"msgbox_confirm_replace_all").c_str(),
        LM.get(L"msgbox_title_confirm").c_str(),
        MB_ICONWARNING | MB_OKCANCEL
    ) != IDOK)
    {
        return;
    }

    // Reset the UI counters before starting
    resetCountColumns();
    std::vector<int> listFindTotals(replaceListData.size(), 0);
    std::vector<int> listReplaceTotals(replaceListData.size(), 0);

    // How many docs are open in each view?
    LRESULT docCountMain = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, PRIMARY_VIEW);
    LRESULT docCountSecondary = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, SECOND_VIEW);

    bool visibleMain = IsWindowVisible(nppData._scintillaMainHandle);
    bool visibleSecond = IsWindowVisible(nppData._scintillaSecondHandle);

    // Remember which doc was active
    LRESULT currentDocIndex = ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTDOCINDEX, 0, MAIN_VIEW);

    // Local lambda to process one document
    auto processOneDoc = [&](int view, LRESULT idx) -> bool {
        ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, view, idx);
        handleDelimiterPositions(DelimiterOperation::LoadAll);
        if (!handleReplaceAllButton()) return false; // aborted via Stop/Error

        // Accumulate the per-list-entry counters from the UI
        for (size_t j = 0; j < replaceListData.size(); ++j) {
            int f = (replaceListData[j].findCount > -1) ? replaceListData[j].findCount : 0;
            int r = (replaceListData[j].replaceCount > -1) ? replaceListData[j].replaceCount : 0;
            listFindTotals[j] += f;
            listReplaceTotals[j] += r;
        }
        resetCountColumns();
        return true;
        };

    // Iterate main view
    if (visibleMain) {
        for (LRESULT i = 0; i < docCountMain; ++i) {
            if (!processOneDoc(PRIMARY_VIEW, i)) break;
        }
    }
    // Iterate secondary view
    if (visibleSecond) {
        for (LRESULT i = 0; i < docCountSecondary; ++i) {
            if (!processOneDoc(SECOND_VIEW, i)) break;
        }
    }

    // Restore the originally active document
    ::SendMessage(nppData._nppHandle,
        NPPM_ACTIVATEDOC,
        visibleMain ? PRIMARY_VIEW : SECOND_VIEW,
        currentDocIndex);

    // Write back only the enabled entries
    for (size_t j = 0; j < replaceListData.size(); ++j) {
        if (!replaceListData[j].isEnabled) continue;
        updateCountColumns(j, listFindTotals[j], listReplaceTotals[j]);
    }
    refreshUIListView();

}

bool MultiReplace::handleReplaceAllButton(bool showCompletionMessage, const std::filesystem::path* explicitPath) {

    if (!validateDelimiterData()) {
        return false;
    }

    // Selection mode with no selection: different behavior for single-doc vs. multi-doc
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED &&
        getSelectionInfo(false).length == 0)
    {
        if (isReplaceAllInDocs) {
            return true;  // Multi-doc mode: skip to next document silently
        }
        else {
            showStatusMessage(LM.get(L"status_no_selection"), MessageStatus::Error, true);
            return false; // Single-doc mode: show error
        }
    }

    // First check if the document is read-only
    LRESULT isReadOnly = send(SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(LM.get(L"status_cannot_replace_read_only"), MessageStatus::Error);
        return false;
    }

    if (!initLuaState()) {
        // Fallback: The _luaInitialized flag remains false, 
        // so resolveLuaSyntax(...) calls will fail safely inside
        // and effectively do nothing. We just continue.
    }

    // Read Filename and Path for LUA
    updateFilePathCache(explicitPath);

    int totalReplaceCount = 0;
    bool replaceSuccess = true;

    if (useListEnabled)
    {
        // Check if the replaceListData is empty and warn the user if so
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_instructions"), MessageStatus::Error);
            return false;
        }

        // Check status for initial call if stopped by DEBUG, don't highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/false)) {
            return false;
        }

        // snapshot a stable start and current selection (only computed once)
        const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
        const bool allFromCursor = allFromCursorEnabled;

        SearchContext startCtx{};
        startCtx.docLength = send(SCI_GETLENGTH, 0, 0);
        startCtx.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
        startCtx.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
        startCtx.retrieveFoundText = false;
        startCtx.highlightMatch = false;

        const SelectionInfo fixedSel = getSelectionInfo(false);
        const Sci_Position  fixedStart = computeAllStartPos(startCtx, wrapAroundEnabled, allFromCursor);

        {
            ScopedUndoAction undo(*this);
            for (size_t i = 0; i < replaceListData.size(); ++i)
            {
                if (replaceListData[i].isEnabled)
                {
                    if (!wrapAroundEnabled && allFromCursor)
                    {
                        const Sci_Position docLenNow = send(SCI_GETLENGTH, 0, 0);
                        auto clamp = [&](Sci_Position p) {
                            return (p < 0) ? 0 : (p > docLenNow ? docLenNow : p);
                            };

                        if (startCtx.isSelectionMode) {
                            Sci_Position s = clamp(fixedSel.startPos);
                            Sci_Position e = clamp(fixedSel.endPos);
                            if (e < s) std::swap(s, e);
                            send(SCI_SETSEL, s, e);
                        }
                        else {
                            const Sci_Position s = clamp(fixedStart);
                            send(SCI_GOTOPOS, s, 0);
                        }
                    }

                    int findCount = 0;
                    int replaceCount = 0;

                    // Call replaceAll and break out if there is an error or a Debug Stop
                    replaceSuccess = replaceAll(replaceListData[i], findCount, replaceCount, i);

                    // Refresh ListView to show updated statistics
                    refreshUIListView();

                    // Accumulate total replacements
                    totalReplaceCount += replaceCount;

                    // cosmetic caret restore only for single-document run
                    if (!replaceSuccess) {
                        break;
                    }
                }
            }
        }

    }
    else
    {
        ReplaceItemData itemData;
        itemData.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        itemData.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        itemData.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        itemData.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        itemData.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
        itemData.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        itemData.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), itemData.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), itemData.replaceText);

        {
            ScopedUndoAction undo(*this);
            int findCount = 0;
            replaceSuccess = replaceAll(itemData, findCount, totalReplaceCount);
        }

    }

    // If debug window is still open after all matches, wait for user to close it
    WaitForDebugWindowClose();

    // Display status message
    if (replaceSuccess && showCompletionMessage) {
        showStatusMessage(LM.get(L"status_occurrences_replaced", { std::to_wstring(totalReplaceCount) }), MessageStatus::Success);
    }

    return replaceSuccess;

}

void MultiReplace::handleReplaceButton() {

    if (!validateDelimiterData()) {
        return;
    }

    // Safety: Selection mode with no selection → do nothing
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        SelectionInfo sel = getSelectionInfo(false);
        if (sel.length == 0) {
            showStatusMessage(LM.get(L"status_no_selection"), MessageStatus::Error, true);
            return;
        }
    }

    updateSelectionScope();

    // First check if the document is read-only
    LRESULT isReadOnly = send(SCI_GETREADONLY, 0, 0);
    if (isReadOnly) {
        showStatusMessage(LM.getW(L"status_cannot_replace_read_only"), MessageStatus::Error);
        return;
    }

    if (!initLuaState()) {
        // Fallback: The _luaInitialized flag remains false, 
        // so resolveLuaSyntax(...) calls will fail safely inside
        // and effectively do nothing. We just continue.
    }

    // Read Filename and Path for LUA
    updateFilePathCache();

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SearchResult searchResult;
    searchResult.pos = -1;
    searchResult.length = 0;
    searchResult.foundText = "";

    SelectionInfo selection = getSelectionInfo(false);
    Sci_Position newPos = (selection.length > 0) ? selection.startPos : send(SCI_GETCURRENTPOS, 0, 0);

    size_t matchIndex = std::numeric_limits<size_t>::max();

    SearchContext context;
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.useStoredSelections = context.isSelectionMode;
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_uncheck"), MessageStatus::Error);
            return;
        }

        // Check status of initial if stopped by DEBUG, highlight entry
        if (!preProcessListForReplace(/*highlightEntry=*/true)) {
            return;
        }

        bool wasReplaced = false;  // Detection for eplacements
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled) {
                context.findText = convertAndExtendW(replaceListData[i].findText, replaceListData[i].extended);
                context.searchFlags = (replaceListData[i].wholeWord * SCFIND_WHOLEWORD) |
                    (replaceListData[i].matchCase * SCFIND_MATCHCASE) |
                    (replaceListData[i].regex * SCFIND_REGEXP);

                // Set search flags before calling replaceOne
                send(SCI_SETSEARCHFLAGS, context.searchFlags);

                wasReplaced = replaceOne(replaceListData[i], selection, searchResult, newPos, i, context);
                if (wasReplaced) {
                    refreshUIListView(); // Refresh the ListView to show updated statistic
                    break;
                }
            }
        }

        if (!(wasReplaced && stayAfterReplaceEnabled)) {
            searchResult = performListSearchForward(replaceListData, newPos, matchIndex, context);

            if (searchResult.pos < 0 && wrapAroundEnabled) {
                searchResult = performListSearchForward(replaceListData, 0, matchIndex, context);
            }
        }

        // Build and show message based on results
        if (wasReplaced) {
            if (stayAfterReplaceEnabled) {
                refreshUIListView();
                showStatusMessage(LM.get(L"status_replace_one"), MessageStatus::Success);
            }
            else if (searchResult.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                showStatusMessage(LM.get(L"status_replace_one_next_found"), MessageStatus::Info);
            }
            else {
                showStatusMessage(LM.get(L"status_replace_one_none_left"), MessageStatus::Info);
            }
        }

        else {
            if (searchResult.pos < 0) {
                showStatusMessage(LM.get(L"status_no_occurrence_found"), MessageStatus::Error, true);
            }
            else {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_found_text_not_replaced"), MessageStatus::Info);
            }
        }
    }
    else {
        ReplaceItemData replaceItem;
        replaceItem.findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        replaceItem.replaceText = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        replaceItem.wholeWord = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED);
        replaceItem.matchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);
        replaceItem.useVariables = (IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED);
        replaceItem.regex = (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED);
        replaceItem.extended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);

        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), replaceItem.findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceItem.replaceText);

        context.findText = convertAndExtendW(replaceItem.findText, replaceItem.extended);
        context.searchFlags = (replaceItem.wholeWord * SCFIND_WHOLEWORD) |
            (replaceItem.matchCase * SCFIND_MATCHCASE) |
            (replaceItem.regex * SCFIND_REGEXP);

        // Set search flags before calling `performSearchForward` and 'replaceOne' which contains 'performSearchForward'
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        bool wasReplaced = replaceOne(replaceItem, selection, searchResult, newPos, SIZE_MAX, context);

        if (!(wasReplaced && stayAfterReplaceEnabled)) {
            if (searchResult.pos < 0 && wrapAroundEnabled) {
                searchResult = performSearchForward(context, 0);
            }
            else if (searchResult.pos >= 0) {
                searchResult = performSearchForward(context, newPos);
            }
        }

        if (wasReplaced) {
            if (stayAfterReplaceEnabled) {
                showStatusMessage(LM.get(L"status_replace_one"), MessageStatus::Success);
            }
            else if (searchResult.pos >= 0) {
                showStatusMessage(LM.get(L"status_replace_one_next_found"), MessageStatus::Success);
            }
            else {
                showStatusMessage(LM.get(L"status_replace_one_none_left"), MessageStatus::Info);
            }
        }
        else {
            if (searchResult.pos < 0) {
                showStatusMessage(LM.get(L"status_no_occurrence_found"), MessageStatus::Error, true);
            }
            else {
                showStatusMessage(LM.get(L"status_found_text_not_replaced"), MessageStatus::Info);
            }
        }
    }

    // If debug window is still open after single replace, close it automatically
    WaitForDebugWindowClose(true);
}

bool MultiReplace::replaceOne(const ReplaceItemData& itemData, const SelectionInfo& selection, SearchResult& searchResult, Sci_Position& newPos, size_t itemIndex, const SearchContext& context)
{
    // Get the document's codepage once at the beginning.
    int documentCodepage = getCurrentDocCodePage();

    searchResult = performSearchForward(context, selection.startPos);

    // Only proceed if the found match is exactly the same as the initial selection.
    if (searchResult.pos == selection.startPos && searchResult.length == selection.length) {
        bool skipReplace = false;

        if (itemIndex != SIZE_MAX) {
            updateCountColumns(itemIndex, 1);
            selectListItem(itemIndex);
        }

        {
            std::string finalReplaceText; // Will hold the final text in the document's native encoding.

            // --- Lua Variable Expansion ---
            if (itemData.useVariables) {
                std::string luaTemplateUtf8 = Encoding::wstringToUtf8(itemData.replaceText);
                if (!ensureLuaCodeCompiled(luaTemplateUtf8)) {
                    return false;
                }

                LuaVariables vars;
                fillLuaMatchVars(vars, searchResult.pos, searchResult.foundText, 1, 1, context.isColumnMode, documentCodepage);

                if (!resolveLuaSyntax(luaTemplateUtf8, vars, skipReplace, itemData.regex, true, documentCodepage)) {
                    return false;
                }

                // Convert the result from Lua (UTF-8) to the final document codepage.
                finalReplaceText = convertAndExtendW(Encoding::utf8ToWString(luaTemplateUtf8), itemData.extended, documentCodepage);
            }
            else {
                // Case without variables: convert once using the safe helper.
                finalReplaceText = convertAndExtendW(itemData.replaceText, itemData.extended, documentCodepage);
            }

            // --- Final Replacement Execution ---
            if (!skipReplace) {
                newPos = itemData.regex
                    ? performRegexReplace(finalReplaceText, searchResult.pos, searchResult.length)
                    : performReplace(finalReplaceText, searchResult.pos, searchResult.length);

                // Only skip advancing when we deleted a non-empty match:
                if (searchResult.length == 0 || newPos != searchResult.pos) {
                    newPos = ensureForwardProgress(newPos, searchResult);
                }

                send(SCI_SETSEL, newPos, newPos);
                if (itemIndex != SIZE_MAX) {
                    updateCountColumns(itemIndex, -2, 1);
                }
                return true; // A replacement was made.
            }
        }

        // Replacement was skipped by Lua logic.
        newPos = ensureForwardProgress(searchResult.pos + searchResult.length, searchResult);
        send(SCI_SETSEL, newPos, newPos);
    }
    return false; // No replacement was made.
}

bool MultiReplace::replaceAll(const ReplaceItemData& itemData, int& findCount, int& replaceCount, size_t itemIndex)
{
    if (itemData.findText.empty() && !itemData.useVariables) {
        findCount = replaceCount = 0;
        return true;
    }

    // Get the document's codepage once at the beginning.
    const int documentCodepage = getCurrentDocCodePage();

    // --- Search bytes MUST match document codepage---
    SearchContext context;
    context.findText = convertAndExtendW(itemData.findText, itemData.extended);
    context.searchFlags = (itemData.wholeWord * SCFIND_WHOLEWORD) | (itemData.matchCase * SCFIND_MATCHCASE) | (itemData.regex * SCFIND_REGEXP);
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.cachedCodepage = documentCodepage;
    context.isColumnMode = IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED;
    context.isSelectionMode = IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED;
    context.retrieveFoundText = itemData.useVariables;
    context.highlightMatch = false;

    send(SCI_SETSEARCHFLAGS, context.searchFlags);

    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    Sci_Position startPos = computeAllStartPos(context, wrapAroundEnabled, allFromCursorEnabled);

    SearchResult searchResult = performSearchForward(context, startPos);

    // --- Replace at matches---
    bool useMatchList = IsDlgButtonChecked(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX) == BST_CHECKED;
    std::unordered_set<int> matchSet;
    if (useMatchList) {
        std::wstring sel = getTextFromDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT);
        if (sel.empty()) {
            showStatusMessage(LM.get(L"status_missing_match_selection"), MessageStatus::Error);
            return false;
        }
        std::vector<int> matchList = parseNumberRanges(sel, LM.get(L"status_invalid_range_in_match_data"));
        if (matchList.empty()) return false;
        matchSet.insert(matchList.begin(), matchList.end());
    }

    // --- Prepare replacement text template (only if needed for Lua) ---
    std::string luaTemplateUtf8;
    if (itemData.useVariables) {
        luaTemplateUtf8 = Encoding::wstringToUtf8(itemData.replaceText);
        if (!ensureLuaCodeCompiled(luaTemplateUtf8)) {
            return false;
        }
    }

    std::string fixedReplace;
    if (!itemData.useVariables) {
        fixedReplace = convertAndExtendW(itemData.replaceText, itemData.extended, documentCodepage);
    }

    int prevLineIdx = -1;
    int lineFindCount = 0;

    // --- Main replacement loop ---
    while (searchResult.pos >= 0)
    {
        bool skipReplace = false;
        ++findCount;
        if (itemIndex != SIZE_MAX) updateCountColumns(itemIndex, findCount);

        const bool replaceThisHit =
            !useMatchList || (matchSet.count(findCount) != 0);

        Sci_Position nextPos; // declared before both branches use it

        {
            std::string finalReplaceText; // Will hold the final text in the document's native encoding.

            // --- Lua Variable Expansion ---
            if (itemData.useVariables) {
                // Track line position for LCNT (needed even for skipped hits to maintain correct count)
                int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(searchResult.pos), 0));
                if (currentLineIndex != prevLineIdx) { lineFindCount = 0; prevLineIdx = currentLineIndex; }
                ++lineFindCount;

                // Only run Lua if we actually intend to replace this hit
                if (replaceThisHit) {
                    std::string luaWorkingUtf8 = luaTemplateUtf8;
                    LuaVariables vars;
                    fillLuaMatchVars(vars, searchResult.pos, searchResult.foundText, findCount, lineFindCount, context.isColumnMode, documentCodepage);

                    if (!resolveLuaSyntax(luaWorkingUtf8, vars, skipReplace, itemData.regex, true, documentCodepage)) {
                        return false;
                    }

                    if (!skipReplace) {
                        finalReplaceText = convertAndExtendW(Encoding::utf8ToWString(luaWorkingUtf8), itemData.extended, documentCodepage);
                    }
                }
            }

            if (replaceThisHit && !skipReplace) {
                if (!itemData.useVariables) {
                    nextPos = itemData.regex
                        ? performRegexReplace(fixedReplace, searchResult.pos, searchResult.length)
                        : performReplace(fixedReplace, searchResult.pos, searchResult.length);
                }
                else {
                    nextPos = itemData.regex
                        ? performRegexReplace(finalReplaceText, searchResult.pos, searchResult.length)
                        : performReplace(finalReplaceText, searchResult.pos, searchResult.length);
                }

                ++replaceCount;
                if (itemIndex != SIZE_MAX) updateCountColumns(itemIndex, -1, replaceCount);
                context.docLength = send(SCI_GETLENGTH); // Update doc length after modification.
            }
            else {
                nextPos = searchResult.pos + searchResult.length;
            }
            // Only skip advancing when we deleted a non-empty match:
            if (searchResult.length == 0 || nextPos != searchResult.pos) {
                nextPos = ensureForwardProgress(nextPos, searchResult);
            }
            searchResult = performSearchForward(context, nextPos);
        }
    }
    return true;
}

Sci_Position MultiReplace::performReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    send(SCI_SETTARGETRANGE, pos, pos + length);

    return pos + send(
        SCI_REPLACETARGET,
        replaceTextUtf8.size(),
        reinterpret_cast<sptr_t>(replaceTextUtf8.c_str())
    );
}

Sci_Position MultiReplace::performRegexReplace(const std::string& replaceTextUtf8, Sci_Position pos, Sci_Position length)
{
    send(SCI_SETTARGETRANGE, pos, pos + length);

    sptr_t replacedLen = send(
        SCI_REPLACETARGETRE,
        static_cast<WPARAM>(-1),
        reinterpret_cast<sptr_t>(replaceTextUtf8.c_str())
    );

    Sci_Position newPos = pos + static_cast<Sci_Position>(replacedLen);
    return newPos;
}

bool MultiReplace::preProcessListForReplace(bool highlight) {
    if (!replaceListData.empty()) {
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (replaceListData[i].isEnabled && replaceListData[i].useVariables) {
                if (replaceListData[i].findText.empty()) {
                    if (highlight) {
                        selectListItem(i);  // Highlight the list entry
                    }
                    std::string localReplaceTextUtf8 = Encoding::wstringToUtf8(replaceListData[i].replaceText);

                    // Compile the Lua code once and cache it
                    if (!ensureLuaCodeCompiled(localReplaceTextUtf8)) {
                        return false;
                    }

                    bool skipReplace = false;
                    LuaVariables vars;
                    setLuaFileVars(vars);   // Setting FNAME and FPATH
                    if (!resolveLuaSyntax(localReplaceTextUtf8, vars, skipReplace, replaceListData[i].regex, false)) {
                        return false;
                    }
                }
            }
        }
    }
    return true;
}

SelectionInfo MultiReplace::getSelectionInfo(bool isBackward) {
    // Get the number of selections
    LRESULT selectionCount = send(SCI_GETSELECTIONS, 0, 0);
    Sci_Position selectedStart = 0;
    Sci_Position correspondingEnd = 0;
    std::vector<SelectionRange> selections; // Store all selections for sorting

    if (selectionCount > 0) {
        selections.resize(selectionCount);

        // Retrieve all selections
        for (LRESULT i = 0; i < selectionCount; ++i) {
            selections[i].start = send(SCI_GETSELECTIONNSTART, i, 0);
            selections[i].end = send(SCI_GETSELECTIONNEND, i, 0);
        }

        // Sort selections based on direction
        if (isBackward) {
            std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) {
                return a.start > b.start;
                });
        }
        else {
            std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) {
                return a.start < b.start;
                });
        }

        // Set the start and end based on the sorted selection
        selectedStart = selections[0].start;
        correspondingEnd = selections[0].end;
    }
    else {
        // No selection case: use current cursor position
        selectedStart = send(SCI_GETCURRENTPOS, 0, 0);
        correspondingEnd = selectedStart;
    }

    // Calculate the selection length
    Sci_Position selectionLength = correspondingEnd - selectedStart;
    return SelectionInfo{ selectedStart, correspondingEnd, selectionLength };
}

Sci_Position MultiReplace::computeAllStartPos(const SearchContext& context, bool wrapEnabled, bool fromCursorEnabled)
{
    SelectionInfo selInfo = getSelectionInfo(false);
    if (context.isSelectionMode) {
        return selInfo.startPos;
    }
    const Sci_Position caretPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
    return wrapEnabled ? 0 : (fromCursorEnabled ? caretPos : 0);
}

#pragma endregion


#pragma region Lua Engine

void MultiReplace::captureLuaGlobals(lua_State* L) {
    lua_pushglobaltable(L);
    lua_pushnil(L);
    while (lua_next(L, -2) != 0) {
        // Skip non-string keys: lua_tostring on a numeric key would convert it
        // in-place and break lua_next traversal (Lua 5.4 reference §4.6).
        if (lua_type(L, -2) != LUA_TSTRING) {
            lua_pop(L, 1);
            continue;
        }
        const char* key = lua_tostring(L, -2);
        LuaVariable luaVar;
        luaVar.name = key;
        int type = lua_type(L, -1);
        if (type == LUA_TNUMBER) {
            luaVar.type = LuaVariableType::Number;
            luaVar.numberValue = lua_tonumber(L, -1);
        }
        else if (type == LUA_TSTRING) {
            luaVar.type = LuaVariableType::String;
            luaVar.stringValue = lua_tostring(L, -1);
        }
        else if (type == LUA_TBOOLEAN) {
            luaVar.type = LuaVariableType::Boolean;
            luaVar.booleanValue = lua_toboolean(L, -1);
        }
        else {
            // Skipping unknown types
            lua_pop(L, 1);
            continue;
        }
        globalLuaVariablesMap[key] = luaVar;
        lua_pop(L, 1);
    }
    lua_pop(L, 1);
}

std::string MultiReplace::escapeForRegex(const std::string& input) {
    std::string escaped;
    for (char c : input) {
        switch (c) {
        case '\\':
        case '^': case '$': case '.': case '|':
        case '?': case '*': case '+': case '(': case ')':
        case '[': case ']': case '{': case '}':
            escaped.push_back('\\'); // Prepend a backslash to escape regex special chars
            escaped.push_back(c);
            break;
        default:
            escaped.push_back(c);
        }
    }
    return escaped;
}

void MultiReplace::setLuaVariable(lua_State* L, const std::string& varName, const std::string& value) {
    lua_pushstring(L, value.c_str());
    lua_setglobal(L, varName.c_str());
}

void MultiReplace::updateFilePathCache(const std::filesystem::path* explicitPath) {
    if (explicitPath) {
        // A specific path was provided (during "Replace in Files"). Use it.
        // Convert reliably to UTF-8 for Lua.
        cachedFilePath = Encoding::wstringToUtf8(explicitPath->wstring());
        cachedFileName = Encoding::wstringToUtf8(explicitPath->filename().wstring());
    }
    else {
        // No specific path was provided. Default to the active Notepad++ document.
        wchar_t filePathBuffer[MAX_PATH] = {};
        wchar_t fileNameBuffer[MAX_PATH] = {};
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(filePathBuffer));
        ::SendMessage(nppData._nppHandle, NPPM_GETFILENAME, MAX_PATH, reinterpret_cast<LPARAM>(fileNameBuffer));

        // Convert reliably to UTF-8 for Lua.
        cachedFilePath = Encoding::wstringToUtf8(std::wstring(filePathBuffer));
        cachedFileName = Encoding::wstringToUtf8(std::wstring(fileNameBuffer));
    }
}

void MultiReplace::setLuaFileVars(LuaVariables& vars) {
    // Use cached file path if valid
    if (!cachedFilePath.empty() &&
        (cachedFilePath.find('\\') != std::string::npos || cachedFilePath.find('/') != std::string::npos)) {
        vars.FPATH = cachedFilePath;
    }
    else {
        vars.FPATH.clear();
    }

    // Use cached file name
    vars.FNAME = cachedFileName;
}

void MultiReplace::fillLuaMatchVars(
    LuaVariables& vars,
    Sci_Position matchPos,
    const std::string& foundText,
    int cnt,
    int lcnt,
    bool isColumnMode,
    int documentCodepage)
{
    int currentLineIndex = static_cast<int>(send(SCI_LINEFROMPOSITION, static_cast<uptr_t>(matchPos), 0));
    int lineStartPos = (currentLineIndex == 0) ? 0
        : static_cast<int>(send(SCI_POSITIONFROMLINE, static_cast<uptr_t>(currentLineIndex), 0));

    setLuaFileVars(vars);

    if (isColumnMode) {
        ColumnInfo columnInfo = getColumnInfo(matchPos);
        vars.COL = static_cast<int>(columnInfo.startColumnIndex);
    }

    vars.CNT = cnt;
    vars.LCNT = lcnt;
    vars.APOS = static_cast<int>(matchPos) + 1;
    vars.LINE = currentLineIndex + 1;
    vars.LPOS = static_cast<int>(matchPos) - lineStartPos + 1;

    vars.MATCH = foundText;
    if (documentCodepage != SC_CP_UTF8) {
        vars.MATCH = Encoding::wstringToUtf8(Encoding::bytesToWString(vars.MATCH, documentCodepage));
    }
}

bool MultiReplace::initLuaState()
{
    // Reset the Lua state if it already exists.
    if (_luaState) {
        lua_close(_luaState);
        _luaState = nullptr;
    }

    // Invalidate Lua code cache
    _lastCompiledLuaCode.clear();
    _luaCompiledReplaceRef = LUA_NOREF;

    // Create a new Lua state.
    _luaState = luaL_newstate();
    if (!_luaState) {
        MessageBox(nppData._nppHandle, L"Failed to create Lua state", L"Lua Error", MB_OK | MB_ICONERROR);
        return false;
    }

    luaL_openlibs(_luaState);

    if (luaSafeModeEnabled) {
        applyLuaSafeMode(_luaState);
    }

    // Load Lua source code directly (fallback mode)
    if (luaL_loadstring(_luaState, luaSourceCode) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        MessageBox(nppData._nppHandle, Encoding::utf8ToWString(errMsg).c_str(), L"Lua Script Load Error", MB_OK | MB_ICONERROR);
        lua_pop(_luaState, 1);
        lua_close(_luaState);
        _luaState = nullptr;
        return false;
    }

    // Execute the loaded Lua code.
    if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        MessageBox(nppData._nppHandle, Encoding::utf8ToWString(errMsg).c_str(), L"Lua Execution Error", MB_OK | MB_ICONERROR);
        lua_pop(_luaState, 1);
        lua_close(_luaState);
        _luaState = nullptr;
        return false;
    }

    // Any Lua call to safeLoadFileSandbox(...) invokes the C++ code
    lua_pushcfunction(_luaState, &MultiReplace::safeLoadFileSandbox);
    lua_setglobal(_luaState, "safeLoadFileSandbox");

    return true;
}

bool MultiReplace::ensureLuaCodeCompiled(const std::string& luaCode)
{
    if (!_luaState) { return false; }

    // Compile only if changed or not compiled yet
    if (luaCode == _lastCompiledLuaCode && _luaCompiledReplaceRef != LUA_NOREF) {
        return true; // already compiled, reuse
    }

    // Remove old reference if exists
    if (_luaCompiledReplaceRef != LUA_NOREF) {
        luaL_unref(_luaState, LUA_REGISTRYINDEX, _luaCompiledReplaceRef);
        _luaCompiledReplaceRef = LUA_NOREF;
    }

    // Compile Lua code
    if (luaL_loadstring(_luaState, luaCode.c_str()) != LUA_OK) {
        const char* errMsg = lua_tostring(_luaState, -1);
        if (errMsg && isLuaErrorDialogEnabled) {
            std::wstring error_message = Encoding::utf8ToWString(errMsg);
            MessageBox(nppData._nppHandle,
                error_message.c_str(),
                LM.get(L"msgbox_title_use_variables_syntax_error").c_str(),
                MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        }
        lua_pop(_luaState, 1); // pop error message
        return false;
    }

    // Store compiled Lua chunk
    _luaCompiledReplaceRef = luaL_ref(_luaState, LUA_REGISTRYINDEX);
    _lastCompiledLuaCode = luaCode; // Cache the compiled Lua code
    return true;
}

bool MultiReplace::resolveLuaSyntax(std::string& inputString, const LuaVariables& vars, bool& skip, bool regex, bool showDebugWindow, int docCodepage)
{
    // 1) Stack-checkpoint
    const int stackBase = lua_gettop(_luaState);
    auto restoreStack = [this, stackBase]() { lua_settop(_luaState, stackBase); };

    // 2) Ensure Lua state exists
    if (!_luaState) { return false; }

    // 3) Numeric globals
    lua_pushinteger(_luaState, vars.CNT);  lua_setglobal(_luaState, "CNT");
    lua_pushinteger(_luaState, vars.LCNT); lua_setglobal(_luaState, "LCNT");
    lua_pushinteger(_luaState, vars.LINE); lua_setglobal(_luaState, "LINE");
    lua_pushinteger(_luaState, vars.LPOS); lua_setglobal(_luaState, "LPOS");
    lua_pushinteger(_luaState, vars.APOS); lua_setglobal(_luaState, "APOS");
    lua_pushinteger(_luaState, vars.COL);  lua_setglobal(_luaState, "COL");

    // 4) String globals
    setLuaVariable(_luaState, "FPATH", vars.FPATH);
    setLuaVariable(_luaState, "FNAME", vars.FNAME);
    setLuaVariable(_luaState, "MATCH", vars.MATCH);

    // 5) REGEX flag
    lua_pushboolean(_luaState, regex);
    lua_setglobal(_luaState, "REGEX");

    // 6) CAP# globals (regex only)
    std::vector<std::string> capNames;
    if (regex) {
        // Use passed codepage or fall back to Scintilla query
        const int docCp = (docCodepage >= 0) ? docCodepage : static_cast<int>(send(SCI_GETCODEPAGE));
        for (int i = 1; i <= MAX_CAP_GROUPS; ++i) {
            sptr_t len = send(SCI_GETTAG, i, 0, true);
            if (len < 0) { break; }
            std::string capVal;
            if (len > 0) {
                if (tagBuffer.size() < static_cast<size_t>(len + 1)) {
                    tagBuffer.resize(len + 1);
                }
                tagBuffer[0] = '\0';  // Ensure null-termination on short reads
                if (send(SCI_GETTAG, i,
                    reinterpret_cast<sptr_t>(tagBuffer.data()), false) >= 0) {
                    capVal.assign(tagBuffer.data());
                    if (docCp != SC_CP_UTF8) {
                        capVal = Encoding::wstringToUtf8(
                            Encoding::bytesToWString(capVal, docCp));
                    }
                }
            }
            std::string capName = "CAP" + std::to_string(i);
            if (!capVal.empty()) {
                setLuaVariable(_luaState, capName, capVal);
            }
            capNames.push_back(capName);  // Always track for cleanup
        }
    }

    // 7) Run pre-compiled chunk
    lua_rawgeti(_luaState, LUA_REGISTRYINDEX, _luaCompiledReplaceRef);
    if (lua_pcall(_luaState, 0, LUA_MULTRET, 0) != LUA_OK) {
        const char* err = lua_tostring(_luaState, -1);
        if (err && isLuaErrorDialogEnabled) {
            MessageBox(nppData._nppHandle,
                Encoding::utf8ToWString(err).c_str(),
                LM.get(L"msgbox_title_use_variables_syntax_error").c_str(),
                MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        }
        restoreStack();
        return false;
    }

    // 8) resultTable
    lua_getglobal(_luaState, "resultTable");
    if (!lua_istable(_luaState, -1)) {
        if (isLuaErrorDialogEnabled) {
            std::wstring msg =
                LM.get(L"msgbox_use_variables_execution_error",
                    { Encoding::utf8ToWString(inputString.c_str()) });
            MessageBox(nppData._nppHandle, msg.c_str(),
                LM.get(L"msgbox_title_use_variables_execution_error").c_str(),
                MB_OK);
        }
        restoreStack();
        return false;
    }

    // 9) result & skip
    lua_getfield(_luaState, -1, "result");                 // push result
    if (lua_isnil(_luaState, -1)) {
        inputString.clear();
    }
    else if (lua_isstring(_luaState, -1) || lua_isnumber(_luaState, -1)) {
        std::string res = lua_tostring(_luaState, -1);
        if (regex) { res = escapeForRegex(res); }
        inputString = res;
    }
    lua_pop(_luaState, 1);                                 // pop result

    lua_getfield(_luaState, -1, "skip");                   // push skip
    skip = lua_isboolean(_luaState, -1) && lua_toboolean(_luaState, -1);
    lua_pop(_luaState, 1);                                 // pop skip

    // (resultTable left on stack until the very end, then dropped by restoreStack)

    // 10) CAP variable dump & cleanup
    std::string capVariablesStr;
    for (const auto& capName : capNames) {
        lua_getglobal(_luaState, capName.c_str());         // push CAP value

        if (lua_isnumber(_luaState, -1)) {
            double n = lua_tonumber(_luaState, -1);
            std::ostringstream os; os << std::fixed << std::setprecision(8) << n;
            capVariablesStr += capName + "\tNumber\t" + os.str() + "\n\n";
        }
        else if (lua_isboolean(_luaState, -1)) {
            bool b = lua_toboolean(_luaState, -1);
            capVariablesStr += capName + "\tBoolean\t" + (b ? "true" : "false") + "\n\n";
        }
        else if (lua_isstring(_luaState, -1)) {
            capVariablesStr += capName + "\tString\t" + SU::escapeControlChars(lua_tostring(_luaState, -1)) + "\n\n";
        }

        lua_pop(_luaState, 1);                             // pop CAP value

        lua_pushnil(_luaState);                            // clear global
        lua_setglobal(_luaState, capName.c_str());
    }

    // 11) DEBUG flag & window
    lua_getglobal(_luaState, "DEBUG");
    bool luaDebugExists = !lua_isnil(_luaState, -1);
    bool luaDebug = luaDebugExists && lua_toboolean(_luaState, -1);
    lua_pop(_luaState, 1);
    bool debugOn = luaDebugExists ? luaDebug : _debugModeEnabled;

    if (debugOn && showDebugWindow) {
        globalLuaVariablesMap.clear();
        captureLuaGlobals(_luaState);

        std::string globalsStr = "Global Lua variables:\n\n";
        for (const auto& p : globalLuaVariablesMap) {
            const LuaVariable& v = p.second;
            if (v.type == LuaVariableType::String) {
                globalsStr += v.name + "\tString\t" + SU::escapeControlChars(v.stringValue) + "\n\n";
            }
            else if (v.type == LuaVariableType::Number) {
                std::ostringstream os; os << std::fixed << std::setprecision(8) << v.numberValue;
                globalsStr += v.name + "\tNumber\t" + os.str() + "\n\n";
            }
            else if (v.type == LuaVariableType::Boolean) {
                globalsStr += v.name + "\tBoolean\t" + (v.booleanValue ? "true" : "false") + "\n\n";
            }
        }

        refreshUIListView();

        int resp = ShowDebugWindow(capVariablesStr + globalsStr);

        if (resp == 3) { restoreStack(); return false; }   // "Stop"
        if (resp == -1) { restoreStack(); return false; }  // window closed
    }

    // 12) Success
    restoreStack();
    return true;
}

int MultiReplace::safeLoadFileSandbox(lua_State* L)
{
    const char* path = luaL_checkstring(L, 1);

    // Read file (your existing UTF-8/ANSI logic)
    std::wstring wpath = Encoding::utf8ToWString(path);
    std::ifstream in(std::filesystem::path(wpath), std::ios::binary);
    if (!in) {
        lua_pushboolean(L, false);
        lua_pushfstring(L, "Cannot open file: %s", path);
        return 2;
    }
    std::string raw((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
    in.close();

    bool rawIsUtf8 = Encoding::isValidUtf8(raw.data(), raw.size());
    std::string utf8_buf;
    if (rawIsUtf8) {
        utf8_buf = std::move(raw);
    }
    else {
        int wlen = MultiByteToWideChar(CP_ACP, 0, raw.data(), static_cast<int>(raw.size()), nullptr, 0);
        std::wstring wide(wlen, L'\0');
        MultiByteToWideChar(CP_ACP, 0, raw.data(), static_cast<int>(raw.size()), &wide[0], wlen);

        int u8len = WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, nullptr, 0, nullptr, nullptr);
        utf8_buf.resize(u8len);
        WideCharToMultiByte(CP_UTF8, 0, wide.data(), wlen, &utf8_buf[0], u8len, nullptr, nullptr);
    }

    // Strip UTF-8 BOM if present
    if (utf8_buf.size() >= 3 &&
        (unsigned char)utf8_buf[0] == 0xEF &&
        (unsigned char)utf8_buf[1] == 0xBB &&
        (unsigned char)utf8_buf[2] == 0xBF)
    {
        utf8_buf.erase(0, 3);
    }

    // Load as text chunk only (avoid binary precompiled chunks)
#if LUA_VERSION_NUM >= 503
    if (luaL_loadbufferx(L, utf8_buf.data(), utf8_buf.size(), path, "t") != LUA_OK)
#else
    if (luaL_loadbuffer(L, utf8_buf.data(), utf8_buf.size(), path) != LUA_OK)
#endif
    {
        lua_pushboolean(L, false);
        lua_pushstring(L, lua_tostring(L, -2)); // copy error
        lua_remove(L, -3); // remove original error under the two pushes
        return 2;
    }

    if (lua_pcall(L, 0, 1, 0) != LUA_OK) {
        lua_pushboolean(L, false);
        lua_pushstring(L, lua_tostring(L, -2));
        lua_remove(L, -3);
        return 2;
    }

    // Success: stack has [ table ]
    lua_pushboolean(L, true);
    lua_insert(L, -2); // [ true, table ]
    return 2;
}

void MultiReplace::applyLuaSafeMode(lua_State* L)
{
    auto removeGlobal = [&](const char* name) {
        lua_pushnil(L);
        lua_setglobal(L, name);
        };

    // Remove dangerous base functions
    removeGlobal("dofile");
    removeGlobal("load");
    removeGlobal("loadfile");
    removeGlobal("require");
    removeGlobal("collectgarbage");

    // Remove whole libraries
    removeGlobal("os");
    removeGlobal("io");
    removeGlobal("package");
    removeGlobal("debug");

    // Keep string/table/math/utf8/base intact.
}

#pragma endregion


#pragma region Lua Debug Window

int MultiReplace::ShowDebugWindow(const std::string& message) {
    debugWindowResponse = -1;

    // Convert the message from UTF-8 to UTF-16 (dynamic size)
    std::wstring wMessage = Encoding::utf8ToWString(message);
    if (wMessage.empty() && !message.empty()) {
        MessageBox(nppData._nppHandle, L"Error converting message", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    static bool isClassRegistered = false;

    if (!isClassRegistered) {
        WNDCLASS wc = {};

        wc.lpfnWndProc = DebugWindowProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = L"DebugWindowClass";
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);

        if (!RegisterClass(&wc)) {
            MessageBox(nppData._nppHandle, L"Error registering class", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
            return -1;
        }

        isClassRegistered = true;
    }

    // Format the message for ListView
    std::wstringstream formattedMessage;
    std::wstringstream inputMessageStream(wMessage);
    std::wstring line;

    while (std::getline(inputMessageStream, line)) {
        if (line.empty()) continue;

        std::wistringstream iss(line);
        std::wstring variable, type, value;

        if (!std::getline(iss, variable, L'\t')) continue;
        if (!std::getline(iss, type, L'\t')) continue;
        std::getline(iss, value);  // Value may be empty

        if (type == L"Number") {
            try {
                double num = std::stod(value);
                if (num == std::floor(num)) {
                    value = std::to_wstring(static_cast<long long>(num));
                }
                else {
                    std::wstringstream numStream;
                    numStream << std::fixed << std::setprecision(8) << num;
                    std::wstring formatted = numStream.str();
                    size_t pos = formatted.find_last_not_of(L'0');
                    if (pos != std::wstring::npos) {
                        formatted.erase(pos + 1);
                    }
                    if (!formatted.empty() && formatted.back() == L'.') {
                        formatted.pop_back();
                    }
                    value = formatted;
                }
            }
            catch (...) {
                // If conversion fails, retain the original value
            }
        }
        else if (type == L"String" && value.empty()) {
            value = L"<empty>";
        }

        formattedMessage << variable << L"\t" << type << L"\t" << value << L"\n";
    }

    std::wstring finalMessage = formattedMessage.str();

    std::wstring windowTitle = LM.get(L"debug_title");

    // Check if window already exists - if so, just update content (PERSISTENT WINDOW)
    if (IsWindow(hDebugWnd) && hDebugListView != nullptr) {
        // Update window title
        SetWindowTextW(hDebugWnd, windowTitle.c_str());

        // Clear existing ListView items
        ListView_DeleteAllItems(hDebugListView);

        // Repopulate ListView with new data
        std::wstringstream ss(finalMessage);
        std::wstring dataLine;
        int itemIndex = 0;
        while (std::getline(ss, dataLine)) {
            if (dataLine.empty()) continue;

            std::wstring variable, type, value;
            std::wistringstream iss(dataLine);

            if (!std::getline(iss, variable, L'\t')) continue;
            if (!std::getline(iss, type, L'\t')) continue;
            std::getline(iss, value);

            LVITEM lvItem = {};
            lvItem.mask = LVIF_TEXT;
            lvItem.iItem = itemIndex;
            lvItem.iSubItem = 0;
            lvItem.pszText = const_cast<LPWSTR>(variable.c_str());
            ListView_InsertItem(hDebugListView, &lvItem);
            ListView_SetItemText(hDebugListView, itemIndex, 1, const_cast<LPWSTR>(type.c_str()));
            ListView_SetItemText(hDebugListView, itemIndex, 2, const_cast<LPWSTR>(value.c_str()));
            ++itemIndex;
        }

        // Wait for user response (window stays open)
        debugWindowResponse = -1;
        MSG msg = {};
        while (IsWindow(hDebugWnd) && debugWindowResponse == -1)
        {
            if (_isShuttingDown) {
                DestroyWindow(hDebugWnd);
                debugWindowResponse = 3;
                hDebugWnd = nullptr;
                hDebugListView = nullptr;
                continue;
            }

            if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == WM_QUIT) break;
                if (!IsDialogMessage(hDebugWnd, &msg)) {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
            else {
                WaitMessage();
            }
        }

        // If Stop was pressed or window closed, cleanup
        if (debugWindowResponse != 2) {
            hDebugWnd = nullptr;
            hDebugListView = nullptr;
        }

        return debugWindowResponse;
    }

    // Window doesn't exist - create it
    int width = debugWindowSizeSet ? debugWindowSize.cx : sx(334);
    int height = debugWindowSizeSet ? debugWindowSize.cy : sy(400);
    int x = debugWindowPositionSet ? debugWindowPosition.x : (GetSystemMetrics(SM_CXSCREEN) - width) / 2;
    int y = debugWindowPositionSet ? debugWindowPosition.y : (GetSystemMetrics(SM_CYSCREEN) - height) / 2;

    HWND hwnd = CreateWindowEx(
        WS_EX_TOPMOST,
        L"DebugWindowClass",
        windowTitle.c_str(),
        WS_OVERLAPPEDWINDOW,
        x, y, width, height,
        nppData._nppHandle, nullptr, hInstance, reinterpret_cast<LPVOID>(const_cast<wchar_t*>(finalMessage.c_str()))
    );

    if (hwnd == nullptr) {
        MessageBoxW(nppData._nppHandle, L"Error creating window", L"Error", MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
        return -1;
    }

    // Activate Dark Mode
    ::SendMessage(nppData._nppHandle, NPPM_DARKMODESUBCLASSANDTHEME, static_cast<WPARAM>(NppDarkMode::dmfInit), reinterpret_cast<LPARAM>(hwnd));

    // Save the handle of the debug window
    hDebugWnd = hwnd;

    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg = {};

    // Message loop - wait for user response (exit when response is set, NOT when window closes)
    while (IsWindow(hwnd) && debugWindowResponse == -1)
    {
        if (_isShuttingDown) {
            DestroyWindow(hwnd);
            debugWindowResponse = 3;
            hDebugWnd = nullptr;
            hDebugListView = nullptr;
            continue;
        }

        if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
        {
            if (msg.message == WM_QUIT) {
                break;
            }

            if (!IsDialogMessage(hwnd, &msg)) {
                if (GetForegroundWindow() != hwnd &&
                    msg.message == WM_KEYDOWN &&
                    (GetKeyState(VK_CONTROL) & 0x8000))
                {
                    bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                    bool handled = true;

                    switch (msg.wParam) {
                    case 'C': SendMessage(nppData._scintillaMainHandle, SCI_COPY, 0, 0); break;
                    case 'V': SendMessage(nppData._scintillaMainHandle, SCI_PASTE, 0, 0); break;
                    case 'X': SendMessage(nppData._scintillaMainHandle, SCI_CUT, 0, 0); break;
                    case 'U': SendMessage(nppData._scintillaMainHandle,
                        shiftPressed ? SCI_UPPERCASE : SCI_LOWERCASE, 0, 0); break;
                    case 'S': SendMessage(nppData._nppHandle,
                        shiftPressed ? NPPM_SAVEALLFILES : NPPM_SAVECURRENTFILE, 0, 0); break;
                    case 'G': SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_SEARCH_GOTOLINE); break;
                    case 'F': SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_SEARCH_FIND); break;
                    default: handled = false; break;
                    }

                    if (handled) continue;
                }

                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
        else
        {
            WaitMessage();
        }
    }

    // If Stop was pressed or window closed (not Next), cleanup
    if (debugWindowResponse != 2) {
        hDebugWnd = nullptr;
        hDebugListView = nullptr;
    }

    if (debugWindowResponse == 2)
    {
        MSG m;
        while (PeekMessage(&m, nullptr, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE)) {}
    }

    return debugWindowResponse;
}

LRESULT CALLBACK MultiReplace::DebugWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HWND hListView;
    static HWND hNextButton;
    static HWND hStopButton;
    static HWND hCopyButton;

    switch (msg) {
    case WM_CREATE: {
        hListView = CreateWindowW(WC_LISTVIEW, L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            10, 10, 360, 140,
            hwnd, reinterpret_cast<HMENU>(1), nullptr, nullptr);
        hDebugListView = hListView;  // Save handle for content updates

        // Initialize columns
        LVCOLUMN lvCol = {};
        lvCol.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        lvCol.pszText = LM.getW(L"debug_col_variable");
        lvCol.cx = 120;
        ListView_InsertColumn(hListView, 0, &lvCol);
        lvCol.pszText = LM.getW(L"debug_col_type");
        lvCol.cx = 80;
        ListView_InsertColumn(hListView, 1, &lvCol);
        lvCol.pszText = LM.getW(L"debug_col_value");
        lvCol.cx = 180;
        ListView_InsertColumn(hListView, 2, &lvCol);

        // Button dimensions
        const int btnWidth = 120;
        const int btnHeight = 30;
        const int btnGap = 10;
        const int btnStartX = 10;
        const int btnY = 160;

        hNextButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_next"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            btnStartX, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(2), nullptr, nullptr);

        hStopButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_stop"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            btnStartX + btnWidth + btnGap, btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(3), nullptr, nullptr);

        hCopyButton = CreateWindowW(L"BUTTON", LM.getW(L"debug_btn_copy"),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            btnStartX + 2 * (btnWidth + btnGap), btnY, btnWidth, btnHeight,
            hwnd, reinterpret_cast<HMENU>(4), nullptr, nullptr);

        // Extract the debug message from lParam
        LPCWSTR debugMessage = reinterpret_cast<LPCWSTR>(reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams);

        // Populate ListView with the debug message
        std::wstringstream ss(debugMessage);
        std::wstring line;
        int itemIndex = 0;
        while (std::getline(ss, line)) {
            if (line.empty()) continue;

            std::wstring variable, type, value;
            std::wistringstream iss(line);

            if (!std::getline(iss, variable, L'\t')) continue;
            if (!std::getline(iss, type, L'\t')) continue;
            std::getline(iss, value);

            LVITEM lvItem = {};
            lvItem.mask = LVIF_TEXT;
            lvItem.iItem = itemIndex;
            lvItem.iSubItem = 0;
            lvItem.pszText = const_cast<LPWSTR>(variable.c_str());
            ListView_InsertItem(hListView, &lvItem);
            ListView_SetItemText(hListView, itemIndex, 1, const_cast<LPWSTR>(type.c_str()));
            ListView_SetItemText(hListView, itemIndex, 2, const_cast<LPWSTR>(value.c_str()));
            ++itemIndex;
        }

        break;
    }

    case WM_SIZE: {
        RECT rect;
        GetClientRect(hwnd, &rect);

        // Button dimensions (fixed left position)
        const int btnWidth = 120;
        const int btnHeight = 30;
        const int btnGap = 10;
        const int btnStartX = 10;
        const int btnY = rect.bottom - btnHeight - 10;
        const int listHeight = rect.bottom - btnHeight - 40;

        SetWindowPos(hListView, nullptr, 10, 10, rect.right - 20, listHeight, SWP_NOZORDER);
        SetWindowPos(hNextButton, nullptr, btnStartX, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hStopButton, nullptr, btnStartX + btnWidth + btnGap, btnY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(hCopyButton, nullptr, btnStartX + 2 * (btnWidth + btnGap), btnY, btnWidth, btnHeight, SWP_NOZORDER);

        ListView_SetColumnWidth(hListView, 0, LVSCW_AUTOSIZE_USEHEADER);
        ListView_SetColumnWidth(hListView, 1, LVSCW_AUTOSIZE_USEHEADER);
        ListView_SetColumnWidth(hListView, 2, LVSCW_AUTOSIZE_USEHEADER);
        break;
    }

    case WM_DPICHANGED: {
        if (instance && instance->dpiMgr) {
            instance->dpiMgr->updateDPI(hwnd);
        }
        RECT* pRect = reinterpret_cast<RECT*>(lParam);
        if (pRect) {
            SetWindowPos(hwnd, nullptr, pRect->left, pRect->top,
                pRect->right - pRect->left, pRect->bottom - pRect->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return 0;
    }

    case WM_COMMAND:
        if (LOWORD(wParam) == 2) {  // Next button
            debugWindowResponse = 2;
            // Don't destroy window - just set response and let message loop exit
        }
        else if (LOWORD(wParam) == 3) {  // Stop/Close button
            debugWindowResponse = 3;

            // Save the window position and size before closing
            RECT rect;
            if (GetWindowRect(hwnd, &rect)) {
                debugWindowPosition.x = rect.left;
                debugWindowPosition.y = rect.top;
                debugWindowPositionSet = true;
                debugWindowSize.cx = rect.right - rect.left;
                debugWindowSize.cy = rect.bottom - rect.top;
                debugWindowSizeSet = true;
            }

            DestroyWindow(hwnd);
            hDebugListView = nullptr;
        }
        else if (LOWORD(wParam) == 4) {  // Copy button was pressed
            CopyListViewToClipboard(hListView);
        }
        break;

    case WM_CLOSE:
        // Only set debugWindowResponse to -1 if not already set by a button
        if (debugWindowResponse == -1) {
            // Closed by X, Alt+F4, or CloseDebugWindow()
            debugWindowResponse = -1;
        }

        // Save the window position and size before closing
        RECT rect;
        if (GetWindowRect(hwnd, &rect)) {
            debugWindowPosition.x = rect.left;
            debugWindowPosition.y = rect.top;
            debugWindowPositionSet = true;
            debugWindowSize.cx = rect.right - rect.left;
            debugWindowSize.cy = rect.bottom - rect.top;
            debugWindowSizeSet = true;
        }

        DestroyWindow(hwnd);
        break;

    case WM_DESTROY:
        break;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void MultiReplace::CopyListViewToClipboard(HWND hListView) {
    const int itemCount = ListView_GetItemCount(hListView);
    if (itemCount <= 0) {
        return;
    }

    const int columnCount = Header_GetItemCount(ListView_GetHeader(hListView));
    std::wstring clipboardText;
    clipboardText.reserve(static_cast<size_t>(itemCount) * columnCount * 64);

    wchar_t buffer[512];

    for (int i = 0; i < itemCount; ++i) {
        for (int j = 0; j < columnCount; ++j) {
            LVITEMW li{};
            li.iSubItem = j;
            li.cchTextMax = std::size(buffer);
            li.pszText = buffer;

            buffer[0] = L'\0';
            SendMessageW(hListView, LVM_GETITEMTEXTW, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(&li));
            buffer[std::size(buffer) - 1] = L'\0';

            int len = lstrlenW(buffer);
            if (len > 0) {
                clipboardText.append(buffer, len);
            }

            if (j < columnCount - 1) {
                clipboardText.push_back(L'\t');
            }
        }
        clipboardText.push_back(L'\n');
    }

    if (clipboardText.empty()) {
        return;
    }

    if (OpenClipboard(nullptr)) {
        EmptyClipboard();
        size_t bytes = (clipboardText.size() + 1) * sizeof(wchar_t);
        HGLOBAL hData = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (hData) {
            void* ptr = GlobalLock(hData);
            if (ptr) {
                memcpy(ptr, clipboardText.c_str(), bytes);
                GlobalUnlock(hData);
                SetClipboardData(CF_UNICODETEXT, hData);
            }
            else {
                GlobalFree(hData);
            }
        }
        CloseClipboard();
    }
}

void MultiReplace::CloseDebugWindow()
{
    // Triggers the WM_CLOSE message for the debug window, handled in DebugWindowProc
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // Save the window position and size before closing
    RECT rc;
    if (GetWindowRect(hDebugWnd, &rc)) {
        debugWindowPosition = { rc.left, rc.top };
        debugWindowSize = { rc.right - rc.left, rc.bottom - rc.top };
        debugWindowPositionSet = debugWindowSizeSet = true;
    }

    SendMessage(hDebugWnd, WM_CLOSE, 0, 0);  // synchronous → window really gone
}

void MultiReplace::SetDebugComplete()
{
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // Update window title to show completion
    SetWindowTextW(hDebugWnd, LM.getW(L"debug_title_complete"));

    // Change Stop button to Close
    HWND hStopButton = GetDlgItem(hDebugWnd, 3);
    if (hStopButton) {
        SetWindowTextW(hStopButton, LM.getW(L"debug_btn_close"));
    }

    // Disable Next button
    HWND hNextButton = GetDlgItem(hDebugWnd, 2);
    if (hNextButton) {
        EnableWindow(hNextButton, FALSE);
    }
}

void MultiReplace::WaitForDebugWindowClose(bool autoClose)
{
    if (!IsWindow(hDebugWnd)) {
        return;
    }

    // For single actions: close window immediately without waiting
    if (autoClose) {
        CloseDebugWindow();
        return;
    }

    SetDebugComplete();

    // Wait for user to close the window
    MSG msg = {};
    while (IsWindow(hDebugWnd))
    {
        if (::PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) break;
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        else {
            WaitMessage();
        }
    }
    hDebugWnd = nullptr;
    hDebugListView = nullptr;
}
#pragma endregion


#pragma region Replace in Files

bool MultiReplace::selectDirectoryDialog(HWND owner, std::wstring& outPath)
{
    // Initialize COM and store the result.
    HRESULT hrInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hrInit)) {
        // If critical COM error, we cannot proceed at all.
        return false;
    }

    IFileDialog* pfd = nullptr;
    // Use a different HRESULT variable for the instance creation.
    HRESULT hrCreate = CoCreateInstance(
        CLSID_FileOpenDialog, nullptr,
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd)
    );

    // Only proceed if the dialog instance was created successfully.
    if (SUCCEEDED(hrCreate)) {
        // Tell it to pick folders only.
        FILEOPENDIALOGOPTIONS opts;
        pfd->GetOptions(&opts);
        pfd->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

        // Show the dialog.
        if (SUCCEEDED(pfd->Show(owner)))
        {
            IShellItem* psi = nullptr;
            if (SUCCEEDED(pfd->GetResult(&psi)) && psi)
            {
                // Retrieve the selected folder’s file system path.
                PWSTR pszPath = nullptr;
                if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath)))
                {
                    outPath = pszPath;
                    CoTaskMemFree(pszPath);
                }
                psi->Release();
            }
        }
        pfd->Release();
    }

    // Uninitialize COM only if our initial call was the one that actually
    // initialized it (returned S_OK).
    if (hrInit == S_OK) {
        CoUninitialize();
    }

    return !outPath.empty();
}

bool MultiReplace::handleBrowseDirectoryButton()
{
    std::wstring dir;
    if (selectDirectoryDialog(_hSelf, dir))
    {
        // User picked one — set the combo-box text & keep history
        SetDlgItemTextW(_hSelf, IDC_DIR_EDIT, dir.c_str());
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), dir);
    }
    return true;  // always return TRUE so the dialog proc knows we handled it
}

void MultiReplace::handleReplaceInFiles() {
    HiddenSciGuard guard;
    if (!guard.create()) {
        showStatusMessage(LM.get(L"status_error_hidden_buffer"), MessageStatus::Error);
        return;
    }

    // Read all inputs once at the beginning.
    auto wDir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);
    auto wFilter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
    const bool recurse = (IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED);
    const bool hide = (IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED);

    if (wFilter.empty()) {
        wFilter = L"*.*";
        SetDlgItemTextW(_hSelf, IDC_FILTER_EDIT, wFilter.c_str());
    }

    // history first (always)
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FILTER_EDIT), wFilter);
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), wDir);

    if (!useListEnabled) {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        std::wstring replW = getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findW);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replW);
    }

    // validate directory
    if (wDir.empty() || !std::filesystem::exists(wDir)) {
        showStatusMessage(LM.get(L"status_error_invalid_directory"), MessageStatus::Error);
        return;
    }

    // CSV config (no-op if column mode is off)
    if (!validateDelimiterData()) return;

    // parse filter after UI/defaults/checks
    guard.parseFilter(wFilter);

    // Apply file size limit settings
    guard.setFileSizeLimitEnabled(limitFileSizeEnabled);
    guard.setMaxFileSizeMB(maxFileSizeMB);

    std::vector<std::filesystem::path> files;
    try {
        namespace fs = std::filesystem;
        if (recurse) {
            for (auto& e : fs::recursive_directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
        else {
            for (auto& e : fs::directory_iterator(wDir, fs::directory_options::skip_permission_denied)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) { files.push_back(e.path()); }
            }
        }
    }
    catch (const std::exception& ex) {
        std::wstring wideReason = Encoding::utf8ToWString(ex.what());
        showStatusMessage(LM.get(L"status_error_scanning_directory", { wideReason }), MessageStatus::Error);
        return;
    }

    if (files.empty()) {
        MessageBox(_hSelf, LM.getW(L"msgbox_no_files"), LM.getW(L"msgbox_title_confirm"), MB_OK);
        return;
    }

    // --- Confirmation Dialog Setup ---
    // Manually shorten the directory path to prevent ugly wrapping.
    HDC dialogHdc = GetDC(_hSelf);
    HFONT dialogHFont = reinterpret_cast<HFONT>(SendMessage(_hSelf, WM_GETFONT, 0, 0));
    SelectObject(dialogHdc, dialogHFont);
    std::wstring shortenedDirectory = getShortenedFilePath(wDir, 400, dialogHdc);
    ReleaseDC(_hSelf, dialogHdc);

    std::wstring message = LM.get(L"msgbox_confirm_replace_in_files", { std::to_wstring(files.size()), shortenedDirectory, wFilter });
    if (MessageBox(_hSelf, message.c_str(), LM.getW(L"msgbox_title_confirm"), MB_OKCANCEL | MB_SETFOREGROUND) != IDOK)
        return;

    // RAII-based UI State Management
    BatchUIGuard uiGuard(this, _hSelf);
    _isCancelRequested = false;

    if (useListEnabled) {
        resetCountColumns();
    }

    std::vector<int> listFindTotals;
    std::vector<int> listReplaceTotals;
    if (useListEnabled) {
        listFindTotals.assign(replaceListData.size(), 0);
        listReplaceTotals.assign(replaceListData.size(), 0);
    }

    int total = static_cast<int>(files.size()), idx = 0, changed = 0;
    showStatusMessage(L"Progress: [  0%]", MessageStatus::Info);

    // Per-file binding guard
    struct SciBindingGuard {
        MultiReplace* self;
        HWND oldSci; SciFnDirect oldFn; sptr_t oldData;
        HiddenSciGuard& g;
        SciBindingGuard(MultiReplace* s, HiddenSciGuard& guard) : self(s), g(guard) {
            oldSci = s->_hScintilla; oldFn = s->pSciMsg; oldData = s->pSciWndData;
            s->_hScintilla = g.hSci; s->pSciMsg = g.fn; s->pSciWndData = g.pData;
        }
        ~SciBindingGuard() {
            self->_hScintilla = oldSci; self->pSciMsg = oldFn; self->pSciWndData = oldData;
        }
    };

    bool aborted = false;

    for (const auto& fp : files) {
        MSG m = {};
        while (::PeekMessage(&m, nullptr, 0, 0, PM_REMOVE)) {
            ::TranslateMessage(&m);
            ::DispatchMessage(&m);
        }

        if (_isShuttingDown) { aborted = true; break; }
        if (_isCancelRequested) { aborted = true; break; }

        ++idx;

        int percent = static_cast<int>((static_cast<double>(idx) / (std::max)(1, total)) * 100.0);
        std::wstring prefix = L"Progress: [" + std::to_wstring(percent) + L"%] ";
        HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
        HDC hdc = GetDC(hStatus);
        HFONT hFont = reinterpret_cast<HFONT>(SendMessage(hStatus, WM_GETFONT, 0, 0));
        SelectObject(hdc, hFont);
        SIZE sz{}; GetTextExtentPoint32W(hdc, prefix.c_str(), static_cast<int>(prefix.length()), &sz);
        RECT rc{}; GetClientRect(hStatus, &rc);
        int avail = (rc.right - rc.left) - sz.cx;
        std::wstring shortPath = getShortenedFilePath(fp.wstring(), avail, hdc);
        ReleaseDC(hStatus, hdc);
        showStatusMessage(prefix + shortPath, MessageStatus::Info);

        std::string original;
        if (!guard.loadFile(fp, original)) { continue; }

        DWORD attrs = GetFileAttributesW(fp.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) { continue; }

        Encoding::DetectOptions dopts;
        const std::vector<char> raw(original.begin(), original.end());
        const Encoding::EncodingInfo enc = Encoding::detectEncoding(raw.data(), raw.size(), dopts);
        std::string u8in;
        if (!Encoding::convertBufferToUtf8(raw, enc, u8in)) { continue; }

        // Bind hidden buffer for the file scope
        {
            SciBindingGuard bind(this, guard);

            send(SCI_CLEARALL, 0, 0);
            send(SCI_SETCODEPAGE, SC_CP_UTF8, 0);
            send(SCI_ADDTEXT, (WPARAM)u8in.length(), reinterpret_cast<sptr_t>(u8in.data()));

            handleDelimiterPositions(DelimiterOperation::LoadAll);

            if (!handleReplaceAllButton(false, &fp)) { _isCancelRequested = true; aborted = true; }

            // Keep list counters in sync (per-item find/replace tallies)
            if (useListEnabled) {
                for (size_t i = 0; i < replaceListData.size(); ++i) {
                    if (!replaceListData[i].isEnabled) continue;
                    const int f = (replaceListData[i].findCount > -1) ? replaceListData[i].findCount : 0;
                    const int r = (replaceListData[i].replaceCount > -1) ? replaceListData[i].replaceCount : 0;
                    listFindTotals[i] += f;
                    listReplaceTotals[i] += r;
                }
            }

            // Write back only if content changed
            std::string u8out = guard.getText();
            if (u8out != u8in) {
                std::vector<char> outBytes;
                if (Encoding::convertUtf8ToOriginal(u8out, enc, outBytes)) {
                    std::string finalBuf(outBytes.begin(), outBytes.end());
                    if (guard.writeFile(fp, finalBuf)) {
                        ++changed;
                    }
                }
            }
        }

        if (aborted) break; // ensures RAII restored before leaving loop
    }

    if (useListEnabled) {
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;
            updateCountColumns(i, listFindTotals[i], listReplaceTotals[i]);
        }
        refreshUIListView();
    }

    // status line
    if (!_isShuttingDown) {
        const bool wasCanceled = (_isCancelRequested || aborted);

        std::wstring msg = LM.get(L"status_replace_summary",
            { std::to_wstring(changed), std::to_wstring(files.size()) });
        if (wasCanceled) {
            msg += L" - " + LM.get(L"status_canceled");
        }

        MessageStatus ms = MessageStatus::Success;
        if (wasCanceled || changed == 0) {
            ms = MessageStatus::Info;
        }

        showStatusMessage(msg, ms);
    }
    _isCancelRequested = false;
}

#pragma endregion


#pragma region Find All

std::wstring MultiReplace::sanitizeSearchPattern(const std::wstring& raw) {
    // Escape newlines directly in the wstring (no encoding conversion needed)
    std::wstring escaped = raw;
    // Replace CR/LF as before, but for wstring
    size_t pos = 0;
    while ((pos = escaped.find(L"\r", pos)) != std::wstring::npos)
        escaped.replace(pos, 1, L"\\r"), pos += 2;
    pos = 0;
    while ((pos = escaped.find(L"\n", pos)) != std::wstring::npos)
        escaped.replace(pos, 1, L"\\n"), pos += 2;
    return escaped;
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

void MultiReplace::handleFindAllButton()
{
    if (!validateDelimiterData()) return;

    // Close any open autocomplete/calltip windows to prevent visual artifacts during search
    ::SendMessage(_hScintilla, SCI_AUTOCCANCEL, 0, 0);
    ::SendMessage(_hScintilla, SCI_CALLTIPCANCEL, 0, 0);

    if (!useListEnabled) {
        std::wstring earlyFind = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), earlyFind);
    }

    ResultDock& dock = ResultDock::instance();
    // Create dock and immediately hide it (like Notepad++ does)
    // This prevents visual artifacts during search
    dock.ensureCreated(nppData);
    dock.hide(nppData);

    auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0) -> LRESULT {
        return ::SendMessage(_hScintilla, m, w, l);
        };

    wchar_t buf[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(buf));
    std::wstring wFilePath = *buf ? buf : L"<untitled>";
    std::string  utf8FilePath = Encoding::wstringToUtf8(wFilePath);

    SearchContext context;
    context.docLength = sciSend(SCI_GETLENGTH);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.retrieveFoundText = false;
    context.highlightMatch = false;

    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
    Sci_Position scanStart = computeAllStartPos(context, wrapAroundEnabled, allFromCursorEnabled);

    ResultDock::FileMap fileMap;
    int totalHits = 0;

    if (useListEnabled)
    {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_uncheck"), MessageStatus::Error);
            return;
        }
        resetCountColumns();

        std::vector<size_t> workIndices = getIndicesOfUniqueEnabledItems(true);

        // Synchronized Limit Calculation
        int editorLimit = static_cast<int>(_textMarkerIds.size());
        int dockLimit = ResultDock::MAX_ENTRY_COLORS;
        int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
        if (effectiveLimit < 1) effectiveLimit = 1;
        int maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;
        bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

        for (size_t idx : workIndices)
        {
            auto& item = replaceListData[idx];

            // Limit Handling
            int slotIndex = static_cast<int>(idx);
            if (slotIndex >= maxListSlots) slotIndex = maxListSlots - 1;

            COLORREF c = ResultDock::generateColorFromText(item.findText, isDark);
            dock.defineSlotColor(slotIndex, c);

            std::wstring sanitizedPattern = this->sanitizeSearchPattern(item.findText);
            context.findText = convertAndExtendW(item.findText, item.extended);
            context.searchFlags = (item.wholeWord ? SCFIND_WHOLEWORD : 0) | (item.matchCase ? SCFIND_MATCHCASE : 0) | (item.regex ? SCFIND_REGEXP : 0);
            sciSend(SCI_SETSEARCHFLAGS, context.searchFlags);

            std::vector<ResultDock::Hit> rawHits;
            LRESULT pos = scanStart;
            while (true) {
                SearchResult r = performSearchForward(context, pos);
                if (r.pos < 0) break;
                pos = advanceAfterMatch(r);

                ResultDock::Hit h{};
                h.fullPathUtf8 = utf8FilePath;
                h.pos = (Sci_Position)r.pos;
                h.length = (Sci_Position)r.length;
                h.docLine = static_cast<int>(sciSend(SCI_LINEFROMPOSITION, r.pos, 0));
                h.searchFlags = context.searchFlags;
                this->trimHitToFirstLine(sciSend, h);
                if (h.length > 0) {
                    h.findTextW = item.findText;
                    h.colorIndex = slotIndex;
                    rawHits.push_back(std::move(h));
                }
            }

            const int hitCnt = static_cast<int>(rawHits.size());
            item.findCount = hitCnt;
            updateCountColumns(idx, hitCnt);

            if (hitCnt > 0) {
                auto& agg = fileMap[utf8FilePath];
                agg.wPath = wFilePath;
                agg.hitCount += hitCnt;
                agg.crits.push_back({ sanitizedPattern, std::move(rawHits) });
                totalHits += hitCnt;
            }
        }
        refreshUIListView();
    }
    else {
        // Single Mode
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        std::wstring headerPattern = this->sanitizeSearchPattern(findW);

        bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
        COLORREF c = isDark ? MARKER_COLOR_DARK : MARKER_COLOR_LIGHT;
        dock.defineSlotColor(0, c);

        context.findText = convertAndExtendW(findW, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        context.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) ? SCFIND_WHOLEWORD : 0) | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) ? SCFIND_MATCHCASE : 0) | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) ? SCFIND_REGEXP : 0);
        sciSend(SCI_SETSEARCHFLAGS, context.searchFlags);

        std::vector<ResultDock::Hit> rawHits;
        LRESULT pos = scanStart;
        while (true) {
            SearchResult r = performSearchForward(context, pos);
            if (r.pos < 0) break;
            pos = advanceAfterMatch(r);

            ResultDock::Hit h{};
            h.fullPathUtf8 = utf8FilePath;
            h.pos = r.pos;
            h.length = r.length;
            h.docLine = static_cast<int>(sciSend(SCI_LINEFROMPOSITION, r.pos, 0));
            h.searchFlags = context.searchFlags;
            this->trimHitToFirstLine(sciSend, h);
            if (h.length > 0) {
                h.findTextW = findW;
                h.colorIndex = 0;
                rawHits.push_back(std::move(h));
            }
        }

        if (!rawHits.empty()) {
            auto& agg = fileMap[utf8FilePath];
            agg.wPath = wFilePath;
            agg.hitCount = static_cast<int>(rawHits.size());
            agg.crits.push_back({ headerPattern, std::move(rawHits) });
            totalHits += agg.hitCount;
        }
    }

    const size_t fileCount = fileMap.size();
    const std::wstring header = useListEnabled
        ? LM.get(L"dock_list_header", { std::to_wstring(totalHits), std::to_wstring(fileCount) })
        : LM.get(L"dock_single_header", { this->sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)), std::to_wstring(totalHits), std::to_wstring(fileCount) });

    // NOW show the dock (after search is complete, like Notepad++ does)
    dock.ensureCreatedAndVisible(nppData);
    if (ResultDock::purgeEnabled()) dock.clear();

    dock.startSearchBlock(header, useListEnabled ? groupResultsEnabled : false, false);
    if (fileCount > 0) dock.appendFileBlock(fileMap, sciSend);
    dock.closeSearchBlock(totalHits, static_cast<int>(fileCount));

    showStatusMessage((totalHits == 0) ? LM.get(L"status_no_matches_found") : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) }), (totalHits == 0) ? MessageStatus::Error : MessageStatus::Success);
}

void MultiReplace::handleFindAllInDocsButton()
{
    if (!validateDelimiterData()) return;

    // Close any open autocomplete/calltip windows to prevent visual artifacts during search
    ::SendMessage(_hScintilla, SCI_AUTOCCANCEL, 0, 0);
    ::SendMessage(_hScintilla, SCI_CALLTIPCANCEL, 0, 0);

    if (!useListEnabled) {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findW);
    }

    ResultDock& dock = ResultDock::instance();
    // Create dock and immediately hide it (like Notepad++ does)
    dock.ensureCreated(nppData);
    dock.hide(nppData);

    int totalHits = 0;
    std::unordered_set<std::string> uniqueFiles;
    if (useListEnabled) resetCountColumns();
    std::vector<int> listHitTotals(useListEnabled ? replaceListData.size() : 0, 0);

    std::vector<size_t> workIndices;
    if (useListEnabled) {
        workIndices = getIndicesOfUniqueEnabledItems(true);
    }

    int maxListSlots = 1;
    bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    if (useListEnabled) {
        int editorLimit = static_cast<int>(_textMarkerIds.size());
        int dockLimit = ResultDock::MAX_ENTRY_COLORS;
        int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
        if (effectiveLimit < 1) effectiveLimit = 1;
        maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;

        // Define Colors Loop (using clean list)
        for (size_t idx : workIndices) {
            int slot = static_cast<int>(idx);
            if (slot >= maxListSlots) slot = maxListSlots - 1;
            COLORREF c = ResultDock::generateColorFromText(replaceListData[idx].findText, isDark);
            dock.defineSlotColor(slot, c);
        }
    }
    else {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        COLORREF c = isDark ? MARKER_COLOR_DARK : MARKER_COLOR_LIGHT;
        dock.defineSlotColor(0, c);
    }

    std::wstring placeholder = useListEnabled
        ? LM.get(L"dock_list_header", { L"0", L"0" })
        : LM.get(L"dock_single_header", { sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)), L"0", L"0" });

    dock.startSearchBlock(placeholder, useListEnabled ? groupResultsEnabled : false, false);

    auto processCurrentBuffer = [&]() {
        pointerToScintilla();
        auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0)->LRESULT { return ::SendMessage(_hScintilla, m, w, l); };

        wchar_t wBuf[MAX_PATH] = {};
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(wBuf));
        std::wstring wPath = *wBuf ? wBuf : L"<untitled>";
        std::string  u8Path = Encoding::wstringToUtf8(wPath);

        const bool selMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
        SelectionInfo sel = getSelectionInfo(false);
        if (selMode && sel.length == 0) return;
        Sci_Position scanStart = selMode ? sel.startPos : 0;
        const bool columnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);

        ResultDock::FileMap fileMap;
        int hitsInFile = 0;

        auto collect = [&](size_t critIdx, const std::wstring& patt, SearchContext& ctx) {
            std::vector<ResultDock::Hit> raw;
            LRESULT pos = scanStart;
            while (true) {
                SearchResult r = performSearchForward(ctx, pos);
                if (r.pos < 0) break;
                pos = advanceAfterMatch(r);

                ResultDock::Hit h{};
                h.fullPathUtf8 = u8Path;
                h.pos = r.pos;
                h.length = r.length;
                h.docLine = static_cast<int>(sciSend(SCI_LINEFROMPOSITION, r.pos, 0));
                h.searchFlags = ctx.searchFlags;
                this->trimHitToFirstLine(sciSend, h);
                if (h.length > 0) {
                    h.findTextW = replaceListData[critIdx].findText;
                    if (useListEnabled) {
                        int slot = static_cast<int>(critIdx);
                        if (slot >= maxListSlots) slot = maxListSlots - 1;
                        h.colorIndex = slot;
                    }
                    else { h.colorIndex = 0; }
                    raw.push_back(std::move(h));
                }
            }
            const int hitCnt = static_cast<int>(raw.size());
            if (useListEnabled && critIdx < listHitTotals.size())
                listHitTotals[critIdx] += hitCnt;
            if (hitCnt == 0) return;

            auto& agg = fileMap[u8Path];
            agg.wPath = wPath;
            agg.hitCount += hitCnt;
            agg.crits.push_back({ sanitizeSearchPattern(patt), std::move(raw) });
            hitsInFile += hitCnt;
            };

        if (useListEnabled) {
            // Optimized: Re-use clean vector
            for (size_t idx : workIndices) {
                const auto& it = replaceListData[idx];
                SearchContext ctx;
                ctx.docLength = sciSend(SCI_GETLENGTH);
                ctx.isColumnMode = columnMode; ctx.isSelectionMode = selMode;
                ctx.findText = convertAndExtendW(it.findText, it.extended);
                ctx.searchFlags = (it.wholeWord ? SCFIND_WHOLEWORD : 0) | (it.matchCase ? SCFIND_MATCHCASE : 0) | (it.regex ? SCFIND_REGEXP : 0);
                sciSend(SCI_SETSEARCHFLAGS, ctx.searchFlags);
                collect(idx, it.findText, ctx);
            }
        }
        else {
            // Single Mode
            std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            if (!findW.empty()) {
                SearchContext ctx;
                ctx.docLength = sciSend(SCI_GETLENGTH);
                ctx.isColumnMode = columnMode; ctx.isSelectionMode = selMode;
                ctx.findText = convertAndExtendW(findW, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
                ctx.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) ? SCFIND_WHOLEWORD : 0) | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) ? SCFIND_MATCHCASE : 0) | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) ? SCFIND_REGEXP : 0);
                sciSend(SCI_SETSEARCHFLAGS, ctx.searchFlags);
                collect(0, sanitizeSearchPattern(findW), ctx);
            }
        }

        if (hitsInFile > 0) {
            dock.appendFileBlock(fileMap, sciSend);
            totalHits += hitsInFile;
            uniqueFiles.insert(u8Path);
        }
        };

    LRESULT savedIdx = ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTDOCINDEX, 0, MAIN_VIEW);
    const bool mainVis = !!::IsWindowVisible(nppData._scintillaMainHandle);
    const bool subVis = !!::IsWindowVisible(nppData._scintillaSecondHandle);

    if (mainVis) {
        LRESULT nbMain = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, PRIMARY_VIEW);
        for (LRESULT i = 0; i < nbMain; ++i) {
            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, MAIN_VIEW, i);
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            processCurrentBuffer();
        }
    }
    if (subVis) {
        LRESULT nbSub = ::SendMessage(nppData._nppHandle, NPPM_GETNBOPENFILES, 0, SECOND_VIEW);
        for (LRESULT i = 0; i < nbSub; ++i) {
            ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, SUB_VIEW, i);
            handleDelimiterPositions(DelimiterOperation::LoadAll);
            processCurrentBuffer();
        }
    }
    ::SendMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, mainVis ? MAIN_VIEW : SUB_VIEW, savedIdx);

    if (useListEnabled) {
        for (size_t i = 0; i < listHitTotals.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;
            replaceListData[i].findCount = listHitTotals[i];
            updateCountColumns(i, listHitTotals[i]);
        }
        refreshUIListView();
    }

    // NOW show the dock (after search is complete, like Notepad++ does)
    dock.ensureCreatedAndVisible(nppData);
    if (ResultDock::purgeEnabled()) dock.clear();

    dock.closeSearchBlock(totalHits, static_cast<int>(uniqueFiles.size()));

    showStatusMessage((totalHits == 0) ? LM.get(L"status_no_matches_found") : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) }), (totalHits == 0) ? MessageStatus::Error : MessageStatus::Success);
}

void MultiReplace::handleFindInFiles() {
    // Close any open autocomplete/calltip windows to prevent visual artifacts during search
    ::SendMessage(_hScintilla, SCI_AUTOCCANCEL, 0, 0);
    ::SendMessage(_hScintilla, SCI_CALLTIPCANCEL, 0, 0);

    HiddenSciGuard guard;
    if (!guard.create()) {
        showStatusMessage(LM.get(L"status_error_hidden_buffer"), MessageStatus::Error);
        return;
    }

    // (Standard File-Parsing block omitted for brevity - no changes until dock clear)
    auto wDir = getTextFromDialogItem(_hSelf, IDC_DIR_EDIT);
    auto wFilter = getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT);
    const bool recurse = (IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED);
    const bool hide = (IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED);
    if (wFilter.empty()) { wFilter = L"*.*"; SetDlgItemTextW(_hSelf, IDC_FILTER_EDIT, wFilter.c_str()); }
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FILTER_EDIT), wFilter);
    addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), wDir);
    if (!useListEnabled) {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findW);
    }
    if (wDir.empty() || !std::filesystem::exists(wDir)) {
        showStatusMessage(LM.get(L"status_error_invalid_directory"), MessageStatus::Error);
        return;
    }
    if (!validateDelimiterData()) return;
    guard.parseFilter(wFilter);
    guard.setFileSizeLimitEnabled(limitFileSizeEnabled);
    guard.setMaxFileSizeMB(maxFileSizeMB);

    std::vector<std::filesystem::path> files;
    try {
        namespace fs = std::filesystem;
        auto iterOpts = fs::directory_options::skip_permission_denied;
        if (recurse) {
            for (auto& e : fs::recursive_directory_iterator(wDir, iterOpts)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) files.push_back(e.path());
            }
        }
        else {
            for (auto& e : fs::directory_iterator(wDir, iterOpts)) {
                if (_isShuttingDown) return;
                if (e.is_regular_file() && guard.matchPath(e.path(), hide)) files.push_back(e.path());
            }
        }
    }
    catch (...) { return; }

    if (files.empty()) {
        MessageBox(_hSelf, LM.getW(L"msgbox_no_files"), LM.getW(L"msgbox_title_confirm"), MB_OK);
        return;
    }

    ResultDock& dock = ResultDock::instance();
    // Create dock and immediately hide it (like Notepad++ does)
    dock.ensureCreated(nppData);
    dock.hide(nppData);

    int totalHits = 0;
    std::unordered_set<std::string> uniqueFiles;
    if (useListEnabled) resetCountColumns();
    std::vector<int> listHitTotals(useListEnabled ? replaceListData.size() : 0, 0);

    std::vector<size_t> workIndices;
    if (useListEnabled) {
        workIndices = getIndicesOfUniqueEnabledItems(true);
    }

    int maxListSlots = 1;
    bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    if (useListEnabled) {
        int editorLimit = static_cast<int>(_textMarkerIds.size());
        int dockLimit = ResultDock::MAX_ENTRY_COLORS;
        int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
        if (effectiveLimit < 1) effectiveLimit = 1;
        maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;

        // Define Colors (Clean loop)
        for (size_t idx : workIndices) {
            int slot = static_cast<int>(idx);
            if (slot >= maxListSlots) slot = maxListSlots - 1;
            COLORREF c = ResultDock::generateColorFromText(replaceListData[idx].findText, isDark);
            dock.defineSlotColor(slot, c);
        }
    }
    else {
        std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        COLORREF c = isDark ? MARKER_COLOR_DARK : MARKER_COLOR_LIGHT;
        dock.defineSlotColor(0, c);
    }

    std::wstring placeholder = useListEnabled
        ? LM.get(L"dock_list_header", { L"0", L"0" })
        : LM.get(L"dock_single_header", { sanitizeSearchPattern(getTextFromDialogItem(_hSelf, IDC_FIND_EDIT)), L"0", L"0" });

    dock.startSearchBlock(placeholder, useListEnabled ? groupResultsEnabled : false, false);

    BatchUIGuard uiGuard(this, _hSelf);
    _isCancelRequested = false;
    int idx = 0;
    const int total = static_cast<int>(files.size());
    showStatusMessage(L"Progress: [  0%]", MessageStatus::Info);

    struct SciBindingGuard {
        MultiReplace* self; HWND oldSci; SciFnDirect oldFn; sptr_t oldData; HiddenSciGuard& g;
        SciBindingGuard(MultiReplace* s, HiddenSciGuard& guard) : self(s), g(guard) {
            oldSci = s->_hScintilla; oldFn = s->pSciMsg; oldData = s->pSciWndData;
            s->_hScintilla = g.hSci; s->pSciMsg = g.fn; s->pSciWndData = g.pData;
        }
        ~SciBindingGuard() { self->_hScintilla = oldSci; self->pSciMsg = oldFn; self->pSciWndData = oldData; }
    };

    bool aborted = false;

    for (const auto& fp : files) {
        MSG m; while (::PeekMessage(&m, nullptr, 0, 0, PM_REMOVE)) { ::TranslateMessage(&m); ::DispatchMessage(&m); }
        if (_isShuttingDown || _isCancelRequested) { aborted = true; break; }

        ++idx;

        const int percent = static_cast<int>((static_cast<double>(idx) / (std::max)(1, total)) * 100.0);
        const std::wstring prefix = L"Progress: [" + std::to_wstring(percent) + L"%] ";
        HWND hStatus = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
        HDC hdc = GetDC(hStatus);
        HFONT hFont = reinterpret_cast<HFONT>(SendMessage(hStatus, WM_GETFONT, 0, 0));
        SelectObject(hdc, hFont);
        SIZE sz{}; GetTextExtentPoint32W(hdc, prefix.c_str(), static_cast<int>(prefix.length()), &sz);
        RECT rc{}; GetClientRect(hStatus, &rc);
        const int avail = (rc.right - rc.left) - sz.cx;
        std::wstring shortPath = getShortenedFilePath(fp.wstring(), avail, hdc);
        ReleaseDC(hStatus, hdc);
        showStatusMessage(prefix + shortPath, MessageStatus::Info);

        std::string original;
        if (!guard.loadFile(fp, original)) continue;

        auto isLikelyBinary = [](const std::string& s) -> bool {
            if (s.find('\0') != std::string::npos) return true;
            size_t ctrl = 0;
            for (unsigned char c : s) if ((c < 0x20 && c != '\r' && c != '\n' && c != '\t') || c == 0x7F) ++ctrl;
            return (s.size() >= 1024 && ctrl > s.size() / 16);
            };

        SciBindingGuard bind(this, guard);
        send(SCI_CLEARALL, 0, 0);

        if (isLikelyBinary(original)) {
            send(SCI_SETCODEPAGE, 0, 0); // ANSI
            send(SCI_ADDTEXT, (WPARAM)original.size(), reinterpret_cast<sptr_t>(original.data()));
        }
        else {
            Encoding::DetectOptions dopts;
            const std::vector<char> raw(original.begin(), original.end());
            const Encoding::EncodingInfo enc = Encoding::detectEncoding(raw.data(), raw.size(), dopts);
            std::string u8;
            if (!Encoding::convertBufferToUtf8(raw, enc, u8)) continue;
            send(SCI_SETCODEPAGE, SC_CP_UTF8, 0);
            send(SCI_ADDTEXT, (WPARAM)u8.length(), reinterpret_cast<sptr_t>(u8.data()));
        }

        handleDelimiterPositions(DelimiterOperation::LoadAll);

        const std::wstring wPath = fp.wstring();
        const std::string  u8Path = Encoding::wstringToUtf8(wPath);
        bool columnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);

        ResultDock::FileMap fileMap;
        int hitsInFile = 0;

        auto collect = [&](size_t critIdx, const std::wstring& pattW, SearchContext& ctx) {
            std::vector<ResultDock::Hit> raw;
            LRESULT pos = 0;
            while (true) {
                SearchResult r = performSearchForward(ctx, pos);
                if (r.pos < 0) break;
                pos = advanceAfterMatch(r);
                ResultDock::Hit h{};
                h.fullPathUtf8 = u8Path;
                h.pos = r.pos;
                h.length = r.length;
                h.docLine = static_cast<int>(send(SCI_LINEFROMPOSITION, r.pos, 0));
                h.searchFlags = ctx.searchFlags;
                this->trimHitToFirstLine([this](UINT m, WPARAM w, LPARAM l)->LRESULT { return send(m, w, l); }, h);
                if (h.length > 0) {
                    h.findTextW = pattW;
                    if (useListEnabled) {
                        int slot = static_cast<int>(critIdx);
                        if (slot >= maxListSlots) slot = maxListSlots - 1;
                        h.colorIndex = slot;
                    }
                    else { h.colorIndex = 0; }
                    raw.push_back(std::move(h));
                }
            }
            const int n = static_cast<int>(raw.size());
            if (n == 0) return;
            auto& agg = fileMap[u8Path];
            agg.wPath = wPath;
            agg.hitCount += n;
            agg.crits.push_back({ sanitizeSearchPattern(pattW), std::move(raw) });
            hitsInFile += n;
            totalHits += n;
            if (useListEnabled && critIdx < listHitTotals.size()) listHitTotals[critIdx] += n;
            };

        if (useListEnabled) {
            for (size_t entryIdx : workIndices) {
                const auto& it = replaceListData[entryIdx];
                SearchContext ctx{};
                ctx.docLength = send(SCI_GETLENGTH); ctx.isColumnMode = columnMode; ctx.isSelectionMode = false;
                ctx.findText = convertAndExtendW(it.findText, it.extended);
                ctx.searchFlags = (it.wholeWord ? SCFIND_WHOLEWORD : 0) | (it.matchCase ? SCFIND_MATCHCASE : 0) | (it.regex ? SCFIND_REGEXP : 0);
                send(SCI_SETSEARCHFLAGS, ctx.searchFlags, 0);
                collect(entryIdx, it.findText, ctx);
            }
        }
        else {
            // Single Mode
            std::wstring findW = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
            if (!findW.empty()) {
                SearchContext ctx{};
                ctx.docLength = send(SCI_GETLENGTH); ctx.isColumnMode = columnMode; ctx.isSelectionMode = false;
                ctx.findText = convertAndExtendW(findW, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
                ctx.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) ? SCFIND_WHOLEWORD : 0) | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) ? SCFIND_MATCHCASE : 0) | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) ? SCFIND_REGEXP : 0);
                send(SCI_SETSEARCHFLAGS, ctx.searchFlags, 0);
                collect(0, findW, ctx);
            }
        }

        if (hitsInFile > 0) {
            auto sciSend = [this](UINT m, WPARAM w = 0, LPARAM l = 0)->LRESULT { return send(m, w, l); };

            // Note: For Find in Files, we do NOT place tracking indicators here.
            // - Files not open in editor: Lazy placement when user navigates to them
            // - Files already open: Also lazy placement (to avoid tab switching/flicker)
            // This is acceptable because Find in Files is primarily for discovery,
            // not for editing. If the user edits before navigating, the position
            // may be off - same behavior as Notepad++'s built-in search.

            dock.appendFileBlock(fileMap, sciSend);
            uniqueFiles.insert(u8Path);
        }
    }

    // NOW show the dock (after search is complete, like Notepad++ does)
    dock.ensureCreatedAndVisible(nppData);
    if (ResultDock::purgeEnabled()) dock.clear();

    dock.closeSearchBlock(totalHits, static_cast<int>(uniqueFiles.size()));

    if (useListEnabled) {
        for (size_t i = 0; i < listHitTotals.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;
            replaceListData[i].findCount = listHitTotals[i];
            updateCountColumns(i, listHitTotals[i]);
        }
        refreshUIListView();
    }
    const bool wasCanceled = (_isCancelRequested || aborted);
    const std::wstring canceledSuffix = wasCanceled ? (L" - " + LM.get(L"status_canceled")) : L"";
    std::wstring msg = (totalHits == 0) ? LM.get(L"status_no_matches_found") : LM.get(L"status_occurrences_found", { std::to_wstring(totalHits) });
    MessageStatus ms = wasCanceled ? MessageStatus::Info : (totalHits == 0 ? MessageStatus::Error : MessageStatus::Success);

    showStatusMessage(msg + canceledSuffix, ms);
    _isCancelRequested = false;
}
#pragma endregion


#pragma region Find

void MultiReplace::handleFindNextButton() {

    if (!validateDelimiterData()) {
        return;
    }

    // Safety: Selection mode with no selection → do nothing
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        SelectionInfo sel = getSelectionInfo(false);
        if (sel.length == 0) {
            showStatusMessage(LM.get(L"status_no_selection"), MessageStatus::Error, true);
            return;
        }
    }

    updateSelectionScope();

    // Selection mode requires valid scope
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED && m_selectionScope.empty()) {
        showStatusMessage(LM.get(L"status_select_area_first"), MessageStatus::Error, true);
        return;
    }

    size_t matchIndex = std::numeric_limits<size_t>::max();
    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);
    SelectionInfo selection = getSelectionInfo(false);

    // Determine the starting search position for forward search
    Sci_Position searchPos;
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED && selection.length > 0) {
        bool isFollowUpSearch = (static_cast<LRESULT>(selection.startPos) == m_lastFindResult.start &&
            static_cast<LRESULT>(selection.endPos) == m_lastFindResult.end);
        // Follow-up search: start AFTER last find. New selection: start at BEGINNING.
        searchPos = isFollowUpSearch ? selection.endPos : selection.startPos;
    }
    else {
        searchPos = send(SCI_GETCURRENTPOS, 0, 0);
    }

    // Initialize SearchContext
    SearchContext context;
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.useStoredSelections = context.isSelectionMode;
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_find_directly"), MessageStatus::Error);
            return;
        }

        // Call performListSearchForward with the complete SearchContext
        SearchResult result = performListSearchForward(replaceListData, searchPos, matchIndex, context);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performListSearchForward(replaceListData, 0, matchIndex, context);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_wrapped_to_first") + getSelectionScopeSuffix(), MessageStatus::Success);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            updateCountColumns(matchIndex, 1);
            refreshUIListView();
            selectListItem(matchIndex);
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found"), MessageStatus::Error, true);
        }
    }
    else {
        // Read search text and check search options
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);

        // Determine if extended mode is active
        bool isExtended = (IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        context.findText = convertAndExtendW(findText, isExtended);

        // Set search flags
        context.searchFlags =
            (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);

        // Set search flags before calling `performSearchForward`
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        // Perform forward search
        SearchResult result = performSearchForward(context, searchPos);
        if (result.pos < 0 && wrapAroundEnabled) {
            result = performSearchForward(context, 0);
            if (result.pos >= 0) {
                showStatusMessage(LM.get(L"status_wrapped_to_first") + getSelectionScopeSuffix(), MessageStatus::Success);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found_for", { findText }), MessageStatus::Error, true);
        }
    }
}

void MultiReplace::handleFindPrevButton() {

    if (!validateDelimiterData()) {
        return;
    }

    // Safety: Selection mode with no selection → do nothing
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED) {
        SelectionInfo sel = getSelectionInfo(false);
        if (sel.length == 0) {
            showStatusMessage(LM.get(L"status_no_selection"), MessageStatus::Error, true);
            return;
        }
    }

    updateSelectionScope();

    // Selection mode requires valid scope
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED && m_selectionScope.empty()) {
        showStatusMessage(LM.get(L"status_select_area_first"), MessageStatus::Error, true);
        return;
    }

    bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    SelectionInfo selection = getSelectionInfo(true);
    Sci_Position searchPos;

    // Determine starting position for backward search:
    // If there is a selection and the selection radio is checked, start from the end of the selection.
    // Otherwise, use the current cursor position.
    if (selection.length > 0 && (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED)) {
        searchPos = selection.endPos;
    }
    else {
        searchPos = send(SCI_GETCURRENTPOS, 0, 0);
    }
    // Move one position backward if possible
    searchPos = (searchPos > 0) ? send(SCI_POSITIONBEFORE, searchPos, 0) : searchPos;

    // Create a fully initialized `SearchContext`
    SearchContext context;
    context.docLength = send(SCI_GETLENGTH, 0, 0);
    context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
    context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
    context.useStoredSelections = context.isSelectionMode;
    context.retrieveFoundText = true;
    context.highlightMatch = true;

    if (useListEnabled) {
        size_t matchIndex = std::numeric_limits<size_t>::max();
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_find_directly"), MessageStatus::Error);
            return;
        }
        SearchResult result = performListSearchBackward(replaceListData, searchPos, matchIndex, context);
        if (result.pos < 0 && wrapAroundEnabled) {
            searchPos = send(SCI_GETLENGTH, 0, 0);
            result = performListSearchBackward(replaceListData, searchPos, matchIndex, context);
            if (result.pos >= 0) {
                updateCountColumns(matchIndex, 1);
                refreshUIListView();
                selectListItem(matchIndex);
                showStatusMessage(LM.get(L"status_wrapped_to_last") + getSelectionScopeSuffix(), MessageStatus::Success);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
            updateCountColumns(matchIndex, 1);
            refreshUIListView();
            selectListItem(matchIndex);
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found"), MessageStatus::Error, true);
        }
    }
    else {
        // Single-Mode: lesen → History → Flags → Suche
        std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);

        context.findText = convertAndExtendW(findText, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        context.searchFlags =
            (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0) |
            (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);

        // Set search flags before calling 'performSearchBackward'
        send(SCI_SETSEARCHFLAGS, context.searchFlags);

        SearchResult result = performSearchBackward(context, searchPos);
        if (result.pos < 0 && wrapAroundEnabled) {
            searchPos = context.isSelectionMode ? selection.endPos : send(SCI_GETLENGTH, 0, 0);
            result = performSearchBackward(context, searchPos);
            if (result.pos >= 0) {
                showStatusMessage(LM.get(L"status_wrapped_to_last") + getSelectionScopeSuffix(), MessageStatus::Success);
                return;
            }
        }
        if (result.pos >= 0) {
            showStatusMessage(L"", MessageStatus::Success);
        }
        else {
            showStatusMessage(LM.get(L"status_no_matches_found_for", { findText }),
                MessageStatus::Error, true);
        }
    }
}

SearchResult MultiReplace::performSingleSearch(const SearchContext& context, SelectionRange range)
{
    // Early exit if the search string is empty
    if (context.findText.empty()) {
        return {};  // Return default-initialized SearchResult
    }

    // Set the target range and search flags via Scintilla
    send(SCI_SETTARGETRANGE, range.start, range.end);

    // Perform the search using SCI_SEARCHINTARGET
    Sci_Position pos = send(SCI_SEARCHINTARGET, context.findText.size(), reinterpret_cast<sptr_t>(context.findText.c_str()));
    Sci_Position matchEnd = send(SCI_GETTARGETEND);

    // Accept zero-length matches (matchEnd == pos) so anchors/lookarounds can be replaced.
    if (pos < 0 || matchEnd < pos || matchEnd > context.docLength) {
        return {};  // Return empty result if no match is found
    }

    // Construct the search result
    SearchResult result;
    result.pos = pos;
    result.length = matchEnd - pos;

    // Retrieve found text only when requested
    if (context.retrieveFoundText) {
        const int codepage = (context.cachedCodepage >= 0)
            ? context.cachedCodepage
            : static_cast<int>(send(SCI_GETCODEPAGE));
        const size_t bytesPerChar = (codepage == SC_CP_UTF8) ? 4u : 1u;

        const size_t charsInTarget = static_cast<size_t>(result.length);
        const size_t cap = charsInTarget * bytesPerChar + 1;

        std::string buf(cap, '\0');
        LRESULT textLength = send(SCI_GETTARGETTEXT, 0, reinterpret_cast<LPARAM>(buf.data()));

        if (textLength < 0) textLength = 0;
        if (static_cast<size_t>(textLength) >= cap)
            textLength = static_cast<LRESULT>(cap - 1);

        buf.resize(static_cast<size_t>(textLength));
        result.foundText = std::move(buf);
    }

    // Highlight the match if required by context
    if (context.highlightMatch) {
        displayResultCentered(pos, matchEnd, true);
    }

    return result;
}

SearchResult MultiReplace::performSearchForward(const SearchContext& context, LRESULT start)
{
    bool isBackward = false;
    SelectionRange targetRange;
    SearchResult result;

    // Use cached mode states from the context to decide the search method
    if (context.isColumnMode && columnDelimiterData.isValid()) {
        result = performSearchColumn(context, start, isBackward);
    }
    else if (context.isSelectionMode) {
        result = performSearchSelection(context, start, isBackward);
    }
    else {
        targetRange.start = start;
        targetRange.end = context.docLength;
        result = performSingleSearch(context, targetRange);
    }
    return result;
}

SearchResult MultiReplace::performSearchBackward(const SearchContext& context, LRESULT start)
{
    bool isBackward = true;
    SelectionRange targetRange;
    SearchResult result;

    if (context.isSelectionMode) {
        result = performSearchSelection(context, start, isBackward);
    }
    else if (context.isColumnMode && columnDelimiterData.isValid()) {
        result = performSearchColumn(context, start, isBackward);
    }
    else {
        targetRange.start = start;
        targetRange.end = 0; // Beginning of the document for backward search
        result = performSingleSearch(context, targetRange);
    }
    return result;
}

SearchResult MultiReplace::performSearchSelection(const SearchContext& context, LRESULT start, bool isBackward) {
    SearchResult result;
    SelectionRange targetRange;
    std::vector<SelectionRange> selections;

    if (context.useStoredSelections && !m_selectionScope.empty()) {
        selections = m_selectionScope;
    }
    else {
        LRESULT selectionCount = send(SCI_GETSELECTIONS, 0, 0);
        if (selectionCount == 0) {
            return SearchResult();
        }
        selections.resize(selectionCount);
        for (int i = 0; i < selectionCount; ++i) {
            selections[i].start = send(SCI_GETSELECTIONNSTART, i, 0);
            selections[i].end = send(SCI_GETSELECTIONNEND, i, 0);
        }
    }

    // Sort selections based on the search direction
    if (isBackward) {
        std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) { return a.start > b.start; });
    }
    else {
        std::sort(selections.begin(), selections.end(), [](const SelectionRange& a, const SelectionRange& b) { return a.start < b.start; });
    }

    // Iterate through each selection range based on the search direction
    for (const auto& selection : selections) {
        if ((isBackward && start < selection.start) || (!isBackward && start > selection.end)) {
            continue; // Skip selections outside the current search range
        }

        // Set the target range within the selection
        targetRange.start = isBackward ? std::min(start, selection.end) : std::max(start, selection.start);
        targetRange.end = isBackward ? selection.start : selection.end;

        // Skip if the range is invalid
        if (targetRange.start == targetRange.end) continue;

        // Perform the search within the target range
        result = performSingleSearch(context, targetRange);
        if (result.pos >= 0) return result;

        // Update the starting position for the next iteration
        start = isBackward ? selection.start - 1 : selection.end + 1;
    }

    return result; // No match found
}

SearchResult MultiReplace::performSearchColumn(const SearchContext& context, LRESULT start, bool isBackward)
{
    SearchResult result;
    SelectionRange targetRange;

    // Identify column start information based on the starting position
    ColumnInfo columnInfo = getColumnInfo(start);
    LRESULT startLine = columnInfo.startLine;
    SIZE_T startColumnIndex = columnInfo.startColumnIndex;
    LRESULT totalLines = columnInfo.totalLines;

    // Set line iteration based on search direction
    LRESULT line = startLine;
    while (isBackward ? (line >= 0) : (line < totalLines)) {
        // Avoid out-of-bounds access in lineDelimiterPositions
        if (line >= static_cast<LRESULT>(lineDelimiterPositions.size())) {
            break;
        }

        // Retrieve local info for the line
        const auto& lineInfo = lineDelimiterPositions[line];
        const auto& linePositions = lineInfo.positions;
        SIZE_T totalColumns = linePositions.size() + 1;

        // Calculate absolute line start and end
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, line, 0);
        LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        // Set column iteration range and step based on direction
        SIZE_T column = isBackward ? (line == startLine ? startColumnIndex : totalColumns) : startColumnIndex;
        SIZE_T endColumnIdx = isBackward ? 1 : totalColumns;
        int columnStep = isBackward ? -1 : 1;

        // Iterate over columns in the specified direction
        for (; (isBackward ? (column >= endColumnIdx) : (column <= endColumnIdx)); column += columnStep) {
            // Skip columns not specified in columnDelimiterData (check BEFORE bounds calculation)
            if (columnDelimiterData.columns.find(static_cast<int>(column)) == columnDelimiterData.columns.end()) {
                continue;
            }

            LRESULT startColumn = 0;
            LRESULT endColumn = 0;

            // Define absolute start of this column
            if (column == 1) {
                startColumn = lineStartPos;
            }
            else {
                startColumn = lineStartPos
                    + lineInfo.positions[column - 2].offsetInLine
                    + columnDelimiterData.delimiterLength;
            }

            // Define absolute end of this column
            if (column == totalColumns) {
                endColumn = lineEndPos;
            }
            else {
                endColumn = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
            }

            // Adjust the target range based on start position and search direction
            if (isBackward && start >= startColumn && start <= endColumn) {
                endColumn = start;
            }
            else if (!isBackward && start >= startColumn && start <= endColumn) {
                startColumn = start;
            }

            // Define target range for the search
            targetRange.start = isBackward ? endColumn : startColumn;
            targetRange.end = isBackward ? startColumn : endColumn;

            // Perform search within the target range
            result = performSingleSearch(context, targetRange);

            // Return immediately if a match is found
            if (result.pos >= 0) {
                return result;
            }
        }

        // Move to the next line based on search direction
        line += (isBackward ? -1 : 1);

        // Reset column index for subsequent lines
        startColumnIndex = isBackward ? totalColumns : 1;
    }

    return result; // No match found in column mode
}

SearchResult MultiReplace::performListSearchBackward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max();

    for (size_t i = 0; i < list.size(); ++i) {
        if (!list[i].isEnabled) {
            continue;
        }

        SearchContext localContext = context;
        localContext.findText = convertAndExtendW(list[i].findText, list[i].extended);
        localContext.searchFlags = (list[i].wholeWord ? SCFIND_WHOLEWORD : 0) |
            (list[i].matchCase ? SCFIND_MATCHCASE : 0) |
            (list[i].regex ? SCFIND_REGEXP : 0);
        localContext.retrieveFoundText = false;
        localContext.highlightMatch = false;

        send(SCI_SETSEARCHFLAGS, localContext.searchFlags);

        SearchResult result = performSearchBackward(localContext, cursorPos);

        if (result.pos >= 0) {
            if (closestMatch.pos < 0 || result.pos > closestMatch.pos) {
                closestMatch.pos = result.pos;
                closestMatch.length = result.length;
                closestMatchIndex = i;
            }
        }
    }

    // Retrieve text and highlight only for the final closest match
    if (closestMatch.pos >= 0) {
        if (context.retrieveFoundText) {
            send(SCI_SETTARGETRANGE, closestMatch.pos, closestMatch.pos + closestMatch.length);
            const int codepage = static_cast<int>(send(SCI_GETCODEPAGE));
            const size_t bytesPerChar = (codepage == SC_CP_UTF8) ? 4u : 1u;
            const size_t cap = static_cast<size_t>(closestMatch.length) * bytesPerChar + 1;
            std::string buf(cap, '\0');
            LRESULT textLength = send(SCI_GETTARGETTEXT, 0, reinterpret_cast<LPARAM>(buf.data()));
            if (textLength < 0) textLength = 0;
            if (static_cast<size_t>(textLength) >= cap) textLength = static_cast<LRESULT>(cap - 1);
            buf.resize(static_cast<size_t>(textLength));
            closestMatch.foundText = std::move(buf);
        }

        if (context.highlightMatch) {
            displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, false);
        }
    }

    return closestMatch;
}

SearchResult MultiReplace::performListSearchForward(const std::vector<ReplaceItemData>& list, LRESULT cursorPos, size_t& closestMatchIndex, const SearchContext& context) {
    SearchResult closestMatch;
    closestMatch.pos = -1;
    closestMatch.length = 0;
    closestMatch.foundText = "";

    closestMatchIndex = std::numeric_limits<size_t>::max();

    for (size_t i = 0; i < list.size(); ++i) {
        if (!list[i].isEnabled) {
            continue;
        }

        SearchContext localContext = context;
        localContext.findText = convertAndExtendW(list[i].findText, list[i].extended);
        localContext.searchFlags = (list[i].wholeWord ? SCFIND_WHOLEWORD : 0) |
            (list[i].matchCase ? SCFIND_MATCHCASE : 0) |
            (list[i].regex ? SCFIND_REGEXP : 0);

        // Disable text retrieval during search - we'll get it only for the final result
        localContext.retrieveFoundText = false;
        localContext.highlightMatch = false;

        send(SCI_SETSEARCHFLAGS, localContext.searchFlags);

        SearchResult result = performSearchForward(localContext, cursorPos);

        if (result.pos >= 0) {
            if (closestMatch.pos < 0 || result.pos < closestMatch.pos) {
                closestMatch.pos = result.pos;
                closestMatch.length = result.length;
                closestMatchIndex = i;
            }
        }
    }

    // retrieve text and highlight only for the final closest match
    if (closestMatch.pos >= 0) {
        if (context.retrieveFoundText) {
            send(SCI_SETTARGETRANGE, closestMatch.pos, closestMatch.pos + closestMatch.length);
            const int codepage = static_cast<int>(send(SCI_GETCODEPAGE));
            const size_t bytesPerChar = (codepage == SC_CP_UTF8) ? 4u : 1u;
            const size_t cap = static_cast<size_t>(closestMatch.length) * bytesPerChar + 1;
            std::string buf(cap, '\0');
            LRESULT textLength = send(SCI_GETTARGETTEXT, 0, reinterpret_cast<LPARAM>(buf.data()));
            if (textLength < 0) textLength = 0;
            if (static_cast<size_t>(textLength) >= cap) textLength = static_cast<LRESULT>(cap - 1);
            buf.resize(static_cast<size_t>(textLength));
            closestMatch.foundText = std::move(buf);
        }

        if (context.highlightMatch) {
            displayResultCentered(closestMatch.pos, closestMatch.pos + closestMatch.length, true);
        }
    }

    return closestMatch;
}

void MultiReplace::displayResultCentered(size_t posStart, size_t posEnd, bool isDownwards)
{
    // Make sure target lines are unfolded
    send(SCI_ENSUREVISIBLE, send(SCI_LINEFROMPOSITION, posStart, 0), 0);
    send(SCI_ENSUREVISIBLE, send(SCI_LINEFROMPOSITION, posEnd, 0), 0);

    // Jump-scroll to center, if current position is out of view
    send(SCI_SETVISIBLEPOLICY, CARET_JUMPS | CARET_EVEN, 0);
    send(SCI_ENSUREVISIBLEENFORCEPOLICY, send(SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // When searching up, the beginning of the (possible multiline) result is important, when scrolling down the end
    send(SCI_GOTOPOS, isDownwards ? posEnd : posStart, 0);
    send(SCI_SETVISIBLEPOLICY, CARET_EVEN, 0);
    send(SCI_ENSUREVISIBLEENFORCEPOLICY, send(SCI_LINEFROMPOSITION, isDownwards ? posEnd : posStart, 0), 0);

    // Adjust so that we see the entire match; primarily horizontally
    send(SCI_SCROLLRANGE, posStart, posEnd);

    // Move cursor to end of result and select result
    send(SCI_GOTOPOS, posEnd, 0);
    send(SCI_SETANCHOR, posStart, 0);

    // Update Scintilla's knowledge about what column the caret is in, so that if user
    // does up/down arrow as first navigation after the search result is selected,
    // the caret doesn't jump to an unexpected column
    send(SCI_CHOOSECARETX, 0, 0);

    m_lastFindResult = { static_cast<LRESULT>(posStart), static_cast<LRESULT>(posEnd) };

}

void MultiReplace::selectListItem(size_t matchIndex) {
    if (!highlightMatchEnabled) {
        return;
    }

    HWND hListView = GetDlgItem(_hSelf, IDC_REPLACE_LIST);
    if (hListView && matchIndex != std::numeric_limits<size_t>::max()) {
        // Deselect all items
        ListView_SetItemState(hListView, -1, 0, LVIS_SELECTED);

        // Select the item at matchIndex
        ListView_SetItemState(hListView, static_cast<int>(matchIndex), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

        // Ensure the item is visible, but only scroll if absolutely necessary
        ListView_EnsureVisible(hListView, static_cast<int>(matchIndex), TRUE);
    }
}

void MultiReplace::updateSelectionScope() {
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) != BST_CHECKED) {
        return;
    }

    Sci_Position selStart = static_cast<Sci_Position>(send(SCI_GETSELECTIONSTART, 0, 0));
    Sci_Position selEnd = static_cast<Sci_Position>(send(SCI_GETSELECTIONEND, 0, 0));

    bool isStaleSelection = (static_cast<LRESULT>(selStart) == m_lastFindResult.start &&
        static_cast<LRESULT>(selEnd) == m_lastFindResult.end);
    bool hasUserSelection = (selEnd > selStart) && !isStaleSelection;

    // Case 1: Scope already exists (follow-up search)
    if (!m_selectionScope.empty()) {
        if (isStaleSelection) {
            return;  // Normal follow-up, keep scope
        }
        if (hasUserSelection) {
            captureCurrentSelectionAsScope();  // User selected new area
            return;
        }
        // Click only (no selection) - check if inside scope
        Sci_Position currentPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
        Sci_Position scopeStart = static_cast<Sci_Position>(m_selectionScope.front().start);
        Sci_Position scopeEnd = static_cast<Sci_Position>(m_selectionScope.back().end);
        if (currentPos < scopeStart || currentPos > scopeEnd) {
            m_selectionScope.clear();  // Click outside - discard scope
        }
        return;
    }

    // Case 2: No scope yet - need new user selection
    if (hasUserSelection) {
        captureCurrentSelectionAsScope();
    }
    // Case 3: Stale selection or just click - no valid scope (m_selectionScope stays empty)
}

void MultiReplace::captureCurrentSelectionAsScope() {
    m_selectionScope.clear();
    LRESULT count = send(SCI_GETSELECTIONS, 0, 0);
    for (LRESULT i = 0; i < count; ++i) {
        SelectionRange range;
        range.start = send(SCI_GETSELECTIONNSTART, i, 0);
        range.end = send(SCI_GETSELECTIONNEND, i, 0);
        if (range.end > range.start) {
            m_selectionScope.push_back(range);
        }
    }
}

std::wstring MultiReplace::getSelectionScopeSuffix() {
    if (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED && !m_selectionScope.empty()) {
        return LM.get(L"status_scope_in_selection");
    }
    return L"";
}

#pragma endregion


#pragma region Mark

void MultiReplace::handleMarkMatchesButton() {
    ensureIndicatorContext();
    if (!validateDelimiterData()) return;

    int totalMatchCount = 0;
    markedStringsCount = 0;
    textToSlot.clear();
    nextSlot = 0;

    const bool wrapAroundEnabled = (IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED);

    if (useListEnabled) {
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_add_values_or_mark_directly"), MessageStatus::Error);
            return;
        }

        std::vector<size_t> workIndices = getIndicesOfUniqueEnabledItems(true);

        // Synchronized Limit Calculation (from QA Fix)
        int editorLimit = static_cast<int>(_textMarkerIds.size());
        int dockLimit = ResultDock::MAX_ENTRY_COLORS;
        int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
        if (effectiveLimit < 1) effectiveLimit = 1;
        int maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;

        // Clean Loop over validated unique items
        for (size_t i : workIndices) {
            const auto& item = replaceListData[i];

            // Use original index 'i' for consistent coloring
            int slot = static_cast<int>(i);
            if (slot >= maxListSlots) slot = maxListSlots - 1;

            textToSlot[item.findText] = slot;

            SearchContext context;
            context.findText = convertAndExtendW(item.findText, item.extended);
            context.searchFlags = (item.wholeWord * SCFIND_WHOLEWORD)
                | (item.matchCase * SCFIND_MATCHCASE)
                | (item.regex * SCFIND_REGEXP);
            context.docLength = send(SCI_GETLENGTH, 0, 0);
            context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
            context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);
            context.retrieveFoundText = false;
            context.highlightMatch = false;

            Sci_Position startPos = 0;
            if (context.isSelectionMode) {
                const SelectionInfo selInfo = getSelectionInfo(false);
                startPos = selInfo.startPos;
            }
            else if (wrapAroundEnabled) {
                startPos = 0;
            }
            else {
                const Sci_Position caretPos = static_cast<Sci_Position>(send(SCI_GETCURRENTPOS, 0, 0));
                startPos = allFromCursorEnabled ? caretPos : 0;
            }

            const int matchCount = markString(context, startPos, item.findText);
            if (matchCount > 0) {
                totalMatchCount += matchCount;
                updateCountColumns(i, matchCount);
            }
        }
        refreshUIListView();
    }
    else {
        // Single Mode (unchanged)
        const std::wstring findText = getTextFromDialogItem(_hSelf, IDC_FIND_EDIT);
        SearchContext context;
        context.findText = convertAndExtendW(findText, IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED);
        context.searchFlags = (IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? SCFIND_WHOLEWORD : 0)
            | (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? SCFIND_MATCHCASE : 0)
            | (IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? SCFIND_REGEXP : 0);
        context.docLength = send(SCI_GETLENGTH, 0, 0);
        context.isColumnMode = (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED);
        context.isSelectionMode = (IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED);

        Sci_Position startPos = 0;
        if (context.isSelectionMode) startPos = getSelectionInfo(false).startPos;
        else if (!wrapAroundEnabled) startPos = allFromCursorEnabled ? (Sci_Position)send(SCI_GETCURRENTPOS) : 0;

        totalMatchCount = markString(context, startPos, findText);
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findText);
    }
    showStatusMessage(LM.get(L"status_occurrences_marked", { std::to_wstring(totalMatchCount) }), MessageStatus::Info);
}

int MultiReplace::markString(const SearchContext& context, Sci_Position initialStart, const std::wstring& findText)
{
    if (context.findText.empty()) return 0;

    // Resolve indicator once before the loop
    const int indicId = resolveIndicatorForText(findText);
    if (indicId < 0) return 0;

    // Set current indicator once — stays valid for all INDICATORFILLRANGE calls
    ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, indicId, 0);

    int markCount = 0;
    LRESULT pos = initialStart;
    send(SCI_SETSEARCHFLAGS, context.searchFlags);

    for (SearchResult r = performSearchForward(context, pos);
        r.pos >= 0;
        r = performSearchForward(context, pos))
    {
        if (r.length > 0) {
            ::SendMessage(_hScintilla, SCI_INDICATORFILLRANGE, r.pos, r.length);
            ++markCount;
        }
        pos = advanceAfterMatch(r);

        if (pos >= context.docLength) break;
    }

    if (useListEnabled && markCount > 0) ++markedStringsCount;
    return markCount;
}

int MultiReplace::resolveIndicatorForText(const std::wstring& findText)
{
    if (_textMarkerIds.empty()) return -1;

    int indicId = -1;

    int editorLimit = static_cast<int>(_textMarkerIds.size());
    int dockLimit = ResultDock::MAX_ENTRY_COLORS;
    int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
    int maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;

    if (useListEnabled && useListColorsForMarking && !findText.empty()) {

        auto it = textToSlot.find(findText);

        if (it != textToSlot.end()) {
            int slot = it->second;
            if (slot < static_cast<int>(_textMarkerIds.size())) {
                indicId = _textMarkerIds[slot];
            }
        }
        else {
            // New text — allocate slot
            int currentSlot = nextSlot;
            if (currentSlot >= maxListSlots) {
                currentSlot = maxListSlots - 1;
            }
            else {
                nextSlot++;
            }
            textToSlot[findText] = currentSlot;
            indicId = _textMarkerIds[currentSlot];
        }

        // Configure indicator style once for this findText
        if (indicId >= 0) {
            bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
            int alpha = isDark ? EDITOR_MARK_ALPHA_DARK : EDITOR_MARK_ALPHA_LIGHT;
            int outlineAlpha = isDark ? EDITOR_OUTLINE_ALPHA_DARK : EDITOR_OUTLINE_ALPHA_LIGHT;
            COLORREF color = ResultDock::generateColorFromText(findText, isDark);

            ::SendMessage(_hScintilla, SCI_INDICSETSTYLE, indicId, INDIC_ROUNDBOX);
            ::SendMessage(_hScintilla, SCI_INDICSETFORE, indicId, color);
            ::SendMessage(_hScintilla, SCI_INDICSETALPHA, indicId, alpha);
            ::SendMessage(_hScintilla, SCI_INDICSETOUTLINEALPHA, indicId, outlineAlpha);
            ::SendMessage(_hScintilla, SCI_INDICSETUNDER, indicId, TRUE);
        }
    }
    else {
        // Single mode — use the last marker
        int singleIndex = static_cast<int>(_textMarkerIds.size()) - 1;
        if (singleIndex >= 0) {
            indicId = _textMarkerIds[singleIndex];
        }
    }

    return indicId;
}

void MultiReplace::handleClearTextMarksButton()
{
    // Clear our fixed marker indicators (but NOT the duplicate indicator)
    for (HWND hSci : { nppData._scintillaMainHandle, nppData._scintillaSecondHandle }) {
        if (!hSci) continue;
        Sci_Position docLen = static_cast<Sci_Position>(::SendMessage(hSci, SCI_GETLENGTH, 0, 0));

        for (int id : _textMarkerIds) {
            if (id >= 0) {
                ::SendMessage(hSci, SCI_SETINDICATORCURRENT, id, 0);
                ::SendMessage(hSci, SCI_INDICATORCLEARRANGE, 0, docLen);
            }
        }
    }

    // Reset color mapping used by resolveIndicatorForText
    // (This is handled by the static map being cleared when no indicators are used)

    markedStringsCount = 0;
    colorToStyleMap.clear();

    textToSlot.clear();
    nextSlot = 0;
}

void MultiReplace::handleCopyMarkedTextToClipboardButton()
{

    if (!validateDelimiterData()) {
        return;
    }

    bool wasLastCharMarked = false;
    size_t markedTextCount = 0;

    std::string markedText;
    std::string styleText;
    std::string eol = getEOLStyle();

    for (int style : textStyles)
    {
        send(SCI_SETINDICATORCURRENT, style, 0);
        LRESULT pos = 0;
        LRESULT nextPos = send(SCI_INDICATOREND, style, pos);

        while (nextPos > pos) // check if nextPos has advanced
        {
            bool atEndOfIndic = send(SCI_INDICATORVALUEAT, style, pos) != 0;

            if (atEndOfIndic)
            {
                if (!wasLastCharMarked)
                {
                    ++markedTextCount;
                }

                wasLastCharMarked = true;

                for (LRESULT i = pos; i < nextPos; ++i)
                {
                    char ch = static_cast<char>(send(SCI_GETCHARAT, i, 0));
                    styleText += ch;
                }

                markedText += styleText + eol; // Append marked text and EOL;
                styleText.clear();
            }
            else
            {
                wasLastCharMarked = false;
            }

            pos = nextPos;
            nextPos = send(SCI_INDICATOREND, style, pos);
        }
    }

    // Remove the last EOL if necessary
    if (!markedText.empty() && markedText.length() >= eol.length()) {
        markedText.erase(markedText.length() - eol.length());
    }

    // Convert encoding to wide string
    std::wstring wstr = Encoding::bytesToWString(markedText, getCurrentDocCodePage());;

    copyTextToClipboard(wstr, static_cast<int>(markedTextCount));
}

void MultiReplace::copyTextToClipboard(const std::wstring& text, int textCount)
{
    if (text.empty()) {
        showStatusMessage(LM.get(L"status_no_text_to_copy"), MessageStatus::Error);
        return;
    }

    if (!OpenClipboard(nullptr))
        return;
    EmptyClipboard();

    const size_t byteSize = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, byteSize);
    if (!hMem) {
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_allocate_memory"), MessageStatus::Error);
        return;
    }

    void* p = GlobalLock(hMem);
    if (!p) {
        GlobalFree(hMem);
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_allocate_memory"), MessageStatus::Error);
        return;
    }

    memcpy(p, text.c_str(), byteSize);
    GlobalUnlock(hMem);

    if (SetClipboardData(CF_UNICODETEXT, hMem) == nullptr) {
        GlobalFree(hMem);
        CloseClipboard();
        showStatusMessage(LM.get(L"status_failed_to_copy"), MessageStatus::Error);
        return;
    }

    CloseClipboard();
    showStatusMessage(LM.get(L"status_items_copied_to_clipboard", { std::to_wstring(textCount) }), MessageStatus::Success);
}

void MultiReplace::initTextMarkerIndicators()
{
    if (_textMarkersInitialized) return;

    HWND hSci0 = nppData._scintillaMainHandle;
    if (!hSci0) return;

    _textMarkerIds.clear();
    _duplicateIndicatorId = -1;

    std::vector<int> available = NppStyleKit::gIndicatorCoord.availableIndicatorPool();

    if (available.empty()) return;

    // Reserve first indicator for duplicate marking (separate from text markers)
    _duplicateIndicatorId = available.front();
    _textMarkerIds.assign(available.begin() + 1, available.end());

    _textMarkersInitialized = true;
    updateTextMarkerStyles();
}

void MultiReplace::updateTextMarkerStyles()
{
    if (_textMarkerIds.empty()) return;

    bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
    int alpha = isDark ? EDITOR_MARK_ALPHA_DARK : EDITOR_MARK_ALPHA_LIGHT;
    int outlineAlpha = isDark ? EDITOR_OUTLINE_ALPHA_DARK : EDITOR_OUTLINE_ALPHA_LIGHT;

    auto applyStyle = [&](HWND hSci, int id, COLORREF color) {
        if (!hSci || id < 0) return;
        ::SendMessage(hSci, SCI_INDICSETSTYLE, id, INDIC_ROUNDBOX);
        ::SendMessage(hSci, SCI_INDICSETFORE, id, color);
        ::SendMessage(hSci, SCI_INDICSETALPHA, id, alpha);
        ::SendMessage(hSci, SCI_INDICSETOUTLINEALPHA, id, outlineAlpha);
        ::SendMessage(hSci, SCI_INDICSETUNDER, id, TRUE);
        };

    // 1. Update List Mode colors (ResultDock & Editor Indicators)
    if (useListEnabled) {
        ResultDock& dock = ResultDock::instance();

        int editorLimit = static_cast<int>(_textMarkerIds.size());
        int dockLimit = ResultDock::MAX_ENTRY_COLORS;
        int effectiveLimit = (editorLimit < dockLimit) ? editorLimit : dockLimit;
        if (effectiveLimit < 1) effectiveLimit = 1;
        int maxListSlots = (effectiveLimit > 1) ? effectiveLimit - 1 : 1;

        // Update ResultDock slots
        for (size_t i = 0; i < replaceListData.size(); ++i) {
            if (!replaceListData[i].isEnabled) continue;

            int slot = static_cast<int>(i);
            if (slot >= maxListSlots) slot = maxListSlots - 1;

            COLORREF c = ResultDock::generateColorFromText(replaceListData[i].findText, isDark);
            dock.defineSlotColor(slot, c);
        }
    }

    // 2. Update active Editor indicators based on map
    for (const auto& [text, slot] : textToSlot) {
        if (slot >= 0 && slot < static_cast<int>(_textMarkerIds.size())) {
            COLORREF c = ResultDock::generateColorFromText(text, isDark);
            int indicId = _textMarkerIds[slot];

            for (HWND hSci : { nppData._scintillaMainHandle, nppData._scintillaSecondHandle }) {
                applyStyle(hSci, indicId, c);
            }
        }
    }

    // 3. Update Duplicate Marker (separate indicator)
    for (HWND hSci : { nppData._scintillaMainHandle, nppData._scintillaSecondHandle }) {
        if (!hSci) continue;

        if (_duplicateIndicatorId >= 0) {
            COLORREF duplicateColor = isDark ? DUPLICATE_MARKER_COLOR_DARK : DUPLICATE_MARKER_COLOR_LIGHT;
            applyStyle(hSci, _duplicateIndicatorId, duplicateColor);
        }
    }

    // 4. Update Single/Standard Marker (last available ID)
    for (HWND hSci : { nppData._scintillaMainHandle, nppData._scintillaSecondHandle }) {
        if (!hSci) continue;

        if (!_textMarkerIds.empty()) {
            int singleMarkerIndex = static_cast<int>(_textMarkerIds.size()) - 1;
            COLORREF markerColor = isDark ? MARKER_COLOR_DARK : MARKER_COLOR_LIGHT;
            applyStyle(hSci, _textMarkerIds[singleMarkerIndex], markerColor);
        }
    }
}

std::vector<size_t> MultiReplace::getIndicesOfUniqueEnabledItems(bool removeDuplicates) const
{
    std::vector<size_t> validIndices;
    validIndices.reserve(replaceListData.size());

    std::unordered_set<std::wstring> seenSignatures;

    for (size_t i = 0; i < replaceListData.size(); ++i) {
        const auto& item = replaceListData[i];

        // 1. Basic Check: Enabled & Not Empty?
        if (!item.isEnabled || item.findText.empty()) continue;

        // 2. Smart Deduplication
        if (removeDuplicates) {
            // Build unique signature: Text + Options
            std::wstring signature;
            signature.reserve(item.findText.size() + 8);

            signature += item.findText;
            signature += L"|"; signature += (item.regex ? L"1" : L"0");
            signature += L"|"; signature += (item.extended ? L"1" : L"0");
            signature += L"|"; signature += (item.matchCase ? L"1" : L"0");
            signature += L"|"; signature += (item.wholeWord ? L"1" : L"0");

            if (seenSignatures.find(signature) != seenSignatures.end()) {
                continue; // Skip exact duplicate
            }
            seenSignatures.insert(signature);
        }

        validIndices.push_back(i);
    }
    return validIndices;
}

#pragma endregion


#pragma region CSV

bool MultiReplace::confirmColumnDeletion() {
    // Attempt to parse columns and delimiters
    if (!parseColumnAndDelimiterData()) {
        return false;  // Parsing failed, exit with false indicating no confirmation
    }

    // Now columnDelimiterData should be populated with the parsed column data
    size_t columnCount = columnDelimiterData.columns.size();
    std::wstring confirmMessage = LM.get(L"msgbox_confirm_delete_columns", { std::to_wstring(columnCount) });

    // Display a message box with Yes/No options and a question mark icon
    int msgboxID = MessageBox(
        nppData._nppHandle,
        confirmMessage.c_str(),
        LM.get(L"msgbox_title_confirm").c_str(),
        MB_ICONWARNING | MB_YESNO
    );

    return (msgboxID == IDYES);  // Return true if user confirmed, else false
}

void MultiReplace::handleDeleteColumns()
{
    if (!validateDelimiterData()) {
        return;
    }

    int deletedFieldsCount = 0;

    {
        ScopedUndoAction undo(*this);

        SIZE_T lineCount = lineDelimiterPositions.size();

        // Iterate over lines from last to first.
        for (SIZE_T i = lineCount; i-- > 0; ) {
            const auto& lineInfo = lineDelimiterPositions[i];

            // Get absolute start and end positions for this line.
            LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, i, 0);
            LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

            // Get the actual EOL length for this line
            LRESULT eolLength = getEOLLengthForLine(i);

            // Vector to collect deletion ranges for the current line.
            std::vector<std::pair<LRESULT, LRESULT>> deletionRanges;

            // Iterate over columns using the set's reverse_iterator.
            for (auto it = columnDelimiterData.columns.rbegin(); it != columnDelimiterData.columns.rend(); ++it) {
                SIZE_T column = *it;

                // Only process columns that exist in this line.
                if (column > lineInfo.positions.size() + 1)
                    continue;

                LRESULT startPos = 0, endPos = 0;
                // Calculate the absolute start position:
                if (column == 1) {
                    startPos = lineStartPos;
                }
                else if (column - 2 < lineInfo.positions.size()) {
                    startPos = lineStartPos + lineInfo.positions[column - 2].offsetInLine;
                }
                else {
                    continue;  // Skip if the column index is invalid for this line.
                }

                // Calculate the absolute end position:
                if (column - 1 < lineInfo.positions.size()) {
                    if (column == 1) {
                        // When deleting the first column, also delete the following delimiter.
                        endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine
                            + columnDelimiterData.delimiterLength;
                    }
                    else {
                        endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
                    }
                }
                else {
                    // Last column: delete until the end of the line.
                    // For non-last lines, subtract actual EOL length to preserve the line break.
                    if (i < lineCount - 1)
                        endPos = lineEndPos - eolLength;
                    else
                        endPos = lineEndPos;
                }

                deletionRanges.push_back({ startPos, endPos });
            }

            // If deletion ranges were computed for this line, merge contiguous or overlapping ranges.
            if (!deletionRanges.empty()) {
                // Sort ranges by start position.
                std::sort(deletionRanges.begin(), deletionRanges.end(),
                    [](const std::pair<LRESULT, LRESULT>& a, const std::pair<LRESULT, LRESULT>& b) {
                        return a.first < b.first;
                    });
                std::vector<std::pair<LRESULT, LRESULT>> mergedRanges;
                auto currentRange = deletionRanges[0];
                for (size_t j = 1; j < deletionRanges.size(); ++j) {
                    const auto& nextRange = deletionRanges[j];
                    // If the next range starts before or exactly when the current range ends, merge them.
                    if (nextRange.first <= currentRange.second) {
                        currentRange.second = std::max(currentRange.second, nextRange.second);
                    }
                    else {
                        mergedRanges.push_back(currentRange);
                        currentRange = nextRange;
                    }
                }
                mergedRanges.push_back(currentRange);

                // Execute the deletions for all merged ranges in reverse order.
                // (Deleting from rightmost to leftmost prevents shifting of absolute positions.)
                for (auto it = mergedRanges.rbegin(); it != mergedRanges.rend(); ++it) {
                    LRESULT lengthToDelete = it->second - it->first;
                    if (lengthToDelete > 0) {
                        send(SCI_DELETERANGE, it->first, lengthToDelete, false);
                        ++deletedFieldsCount;
                    }
                }
            }
        }

    }
    // Display a status message with the number of deleted fields.
    showStatusMessage(LM.get(L"status_deleted_fields_count", { std::to_wstring(deletedFieldsCount) }), MessageStatus::Success);
}

void MultiReplace::handleCopyColumnsToClipboard()
{
    if (!validateDelimiterData()) {
        return;
    }

    // Check once whether Flow-Tab padding is active.
    const bool hasAlignedPadding = ColumnTabs::CT_HasAlignedPadding(_hScintilla);

    // Helper to read [startPos, endPos) from Scintilla for one field.
    // If Flow Tabs padding is active, we skip characters that are marked as padding
    // (indicator range). Otherwise we just read the raw slice.
    auto readFieldText = [this, hasAlignedPadding](LRESULT startPos, LRESULT endPos) -> std::string {
        std::string out;
        out.reserve((size_t)(endPos - startPos));

        if (hasAlignedPadding) {
            const int indicId = ColumnTabs::CT_GetIndicatorId();
            send(SCI_SETINDICATORCURRENT, indicId, 0);

            for (LRESULT p = startPos; p < endPos; ++p) {
                if (static_cast<int>(send(SCI_INDICATORVALUEAT, indicId, p)) != 0)
                    continue; // skip artificial alignment padding
                const int ch = static_cast<int>(send(SCI_GETCHARAT, p, 0));
                if (ch == 0)
                    break;
                out.push_back((char)ch);
            }
        }
        else {
            std::vector<char> buffer(static_cast<size_t>(endPos - startPos) + 1, '\0');
            Sci_TextRangeFull tr;
            tr.chrg.cpMin = startPos;
            tr.chrg.cpMax = endPos;
            tr.lpstrText = buffer.data();
            send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
            out.assign(buffer.data());
        }

        // Trim trailing CR/LF that may belong to the physical line ending.
        // This avoids injecting a line break in the middle of the reordered row
        // (e.g. when copying columns in order "2,1").
        while (!out.empty()) {
            char c = out.back();
            if (c == '\r' || c == '\n') {
                out.pop_back();
            }
            else {
                break;
            }
        }

        return out;
        };

    std::string combinedText;
    int copiedFieldsCount = 0;
    const size_t lineCount = lineDelimiterPositions.size();

    // Walk each known CSV line
    for (size_t i = 0; i < lineCount; ++i) {
        const auto& lineInfo = lineDelimiterPositions[i];

        // Absolute start/end positions of this physical line in Scintilla
        const LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, i, 0);
        const LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        std::string lineText;
        bool firstOutField = true;

        // Use the exact column order the user entered (inputColumns),
        // not the sorted unique set. This preserves e.g. "1,3,2".
        for (SIZE_T colIdx = 0; colIdx < columnDelimiterData.inputColumns.size(); ++colIdx) {
            const int column = columnDelimiterData.inputColumns[colIdx]; // 1-based column index

            // Guard: skip invalid or out-of-range columns.
            // Number of fields in this line = delimiterCount + 1.
            if (column <= 0 || column > static_cast<int>(lineInfo.positions.size() + 1)) {
                continue;
            }

            // Compute [start,end) for the pure field content (no leading delimiter).
            LRESULT startPos = 0;
            LRESULT endPos = 0;

            if (column == 1) {
                // First field starts at line start
                startPos = lineStartPos;
            }
            else {
                // Field N>1 starts right AFTER the delimiter that precedes it
                const SIZE_T delimBeforeIndex = static_cast<SIZE_T>(column - 2);
                if (delimBeforeIndex < lineInfo.positions.size()) {
                    startPos = lineStartPos
                        + lineInfo.positions[delimBeforeIndex].offsetInLine
                        + static_cast<LRESULT>(columnDelimiterData.delimiterLength);
                }
                else {
                    // Safety: malformed position info, skip this column
                    continue;
                }
            }

            // Field ends right BEFORE the next delimiter, or at line end
            if (static_cast<SIZE_T>(column - 1) < lineInfo.positions.size()) {
                endPos = lineStartPos + lineInfo.positions[column - 1].offsetInLine;
            }
            else {
                endPos = lineEndPos;
            }

            // Extract the field text (optionally skipping Flow-Tab padding)
            const std::string fieldText = readFieldText(startPos, endPos);

            // Build output: insert delimiter before every field after the first
            if (!firstOutField) {
                lineText += columnDelimiterData.extendedDelimiter;
            }
            lineText += fieldText;
            firstOutField = false;

            ++copiedFieldsCount;
        }

        combinedText += lineText;

        // Append the document's EOL style to separate rows,
        // but avoid adding duplicate EOL if the last char already is \n or \r.
        if (i < lineCount - 1 &&
            (lineText.empty() || (combinedText.back() != '\n' && combinedText.back() != '\r')))
        {
            combinedText += getEOLStyle();
        }
    }

    // Convert to wide string using the current document code page and copy to clipboard
    std::wstring wstr = Encoding::bytesToWString(combinedText, getCurrentDocCodePage());
    copyTextToClipboard(wstr, copiedFieldsCount);
}

bool MultiReplace::buildCTModelFromMatrix(ColumnTabs::CT_ColumnModelView& outModel) const
{
    if (lineDelimiterPositions.empty() || !columnDelimiterData.isValid())
        return false;

    outModel = {};
    outModel.docStartLine = 0;
    outModel.delimiterIsTab = (columnDelimiterData.extendedDelimiter.size() == 1
        && columnDelimiterData.extendedDelimiter[0] == '\t');
    outModel.delimiterLength = static_cast<int>(columnDelimiterData.delimiterLength);
    outModel.collapseTabRuns = outModel.delimiterIsTab;
    outModel.Lines.reserve(lineDelimiterPositions.size());

    size_t validLines = 0;

    // Copy only what we need (offsets + line length)
    for (size_t i = 0; i < lineDelimiterPositions.size(); ++i) {
        const LineInfo& src = lineDelimiterPositions[i];

        ColumnTabs::CT_ColumnLineInfo li{};
        li.lineLength = static_cast<int>(src.lineLength);

        li.delimiterOffsets.reserve(src.positions.size());
        for (const auto& dp : src.positions) {
            // guard against invalid offsets (must be within line)
            if (dp.offsetInLine >= 0 && dp.offsetInLine < static_cast<LRESULT>(src.lineLength))
                li.delimiterOffsets.push_back(static_cast<int>(dp.offsetInLine));
        }

        // only accept consistent lines
        if (li.lineLength >= 0) {
            outModel.Lines.emplace_back(std::move(li));
            ++validLines;
        }
    }

    // fail if nothing valid was collected
    return validLines > 0;
}

bool MultiReplace::applyFlowTabStops(const ColumnTabs::CT_ColumnModelView* existingModel)
{
    if (existingModel) {
        return ColumnTabs::CT_ApplyFlowTabStopsAll(_hScintilla, *existingModel, _flowPaddingPx);
    }
    ColumnTabs::CT_ColumnModelView model;
    if (!buildCTModelFromMatrix(model))
        return false;
    return ColumnTabs::CT_ApplyFlowTabStopsAll(_hScintilla, model, _flowPaddingPx);
}

void MultiReplace::handleColumnGridTabsButton()
{
    pointerToScintilla();
    if (!_hScintilla) return;

    // Current buffer id
    const BufferId bufId = (BufferId)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

    ColumnTabs::CT_SetIndicatorId(NppStyleKit::gColumnTabsIndicatorId);

    // Seed prev id so the very first switch can clean the source document
    if (g_prevBufId == 0) g_prevBufId = bufId;

    // Ground truth (Padding TEXT state only)
    const bool hasPadNow = ColumnTabs::CT_GetCurDocHasPads(_hScintilla);

    // 1) Desync handling:
    //    If UI state (_flowTabsActive) and document TEXT state (hasPadNow) disagree,
    //    just resync UI/housekeeping without touching document text.
    if (hasPadNow != _flowTabsActive)
    {
        if (hasPadNow)
        {
            // Document still has padding but UI thought it was off.
            // -> Mark as active, restore visuals, update button to "⇤".
            ColumnTabs::CT_SetCurDocHasPads(_hScintilla, true);
            g_padBufs.insert(bufId);

            _flowTabsActive = true;
            if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
                ::SetWindowText(h, L"⇤");

            findAllDelimitersInDocument();
            applyFlowTabStops();

            showStatusMessage(LM.get(L"status_tabs_inserted"), MessageStatus::Success);
        }
        else
        {
            // Document has no padding but UI thought it was on (visuals-only mode).
            // -> Mark as inactive, drop visuals, update button to "⇥".
            ColumnTabs::CT_DisableFlowTabStops(_hScintilla, /*restoreManual=*/false);
            ColumnTabs::CT_ResetFlowVisualState();

            ColumnTabs::CT_SetCurDocHasPads(_hScintilla, false);
            g_padBufs.erase(bufId);

            findAllDelimitersInDocument();

            _flowTabsActive = false;
            if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
                ::SetWindowText(h, L"⇥");

            fixHighlightAtDocumentEnd();

            // Force Scintilla to recalculate word wrapping after visual changes
            forceWrapRecalculation();

            showStatusMessage(LM.get(L"status_tabs_removed"), MessageStatus::Info);
        }

        g_prevBufId = bufId;
        return;
    }

    // From here on: states are in sync.
    // 2) Normal TOGGLE LOGIC

    // ══════════════════════════════════════════════════════════════════════════
    // CASE A: Padding present (and _flowTabsActive == true) -> TURN OFF fully
    // ══════════════════════════════════════════════════════════════════════════
    if (hasPadNow)
    {
        // Disable logging during bulk text modification
        bool wasLoggingEnabled = isLoggingEnabled;
        isLoggingEnabled = false;
        logChanges.clear();

        // Collect padding ranges BEFORE removal (for ResultDock position adjustment)
        std::vector<std::pair<Sci_Position, Sci_Position>> paddingRanges;
        std::string curFileUtf8;
        {
            wchar_t pathBuf[MAX_PATH] = {};
            ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(pathBuf));
            curFileUtf8 = Encoding::wstringToUtf8(pathBuf);
        }

        ResultDock& dock = ResultDock::instance();
        const bool hasResultHits = dock.hasHitsForFile(curFileUtf8);

        if (hasResultHits)
        {
            // Scan for ColumnTabs indicator ranges (same pattern as CT_RemoveAlignedPadding Phase 1)
            const int ctInd = ColumnTabs::CT_GetIndicatorId();
            const Sci_Position docLen = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
            Sci_Position scanPos = 0;
            while (scanPos < docLen) {
                const Sci_Position end = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATOREND, ctInd, scanPos));
                if (end <= scanPos) break;
                if (static_cast<int>(::SendMessage(_hScintilla, SCI_INDICATORVALUEAT, ctInd, scanPos)) != 0) {
                    const Sci_Position start = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATORSTART, ctInd, scanPos));
                    if (end > start)
                        paddingRanges.emplace_back(start, end);
                }
                scanPos = end;
            }
        }

        {
            ScopedRedrawLock redrawLock(_hScintilla);

            ColumnTabs::CT_RemoveAlignedPadding(_hScintilla);
            ColumnTabs::CT_DisableFlowTabStops(_hScintilla, /*restoreManual=*/false);
            ColumnTabs::CT_ResetFlowVisualState();

            ColumnTabs::CT_SetCurDocHasPads(_hScintilla, false);
            g_padBufs.erase(bufId);

            findAllDelimitersInDocument();
        }

        // Adjust ResultDock hit positions: padding was removed, positions shift backward
        if (hasResultHits && !paddingRanges.empty()) {
            dock.adjustHitPositionsForFlowTab(curFileUtf8, paddingRanges, /*added=*/false);
        }

        _flowTabsActive = false;
        if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
            ::SetWindowText(h, L"⇥");

        isLoggingEnabled = wasLoggingEnabled;

        if (isColumnHighlighted) {
            handleHighlightColumnsInDocument();
        }
        else {
            fixHighlightAtDocumentEnd();
        }

        forceWrapRecalculation();

        showStatusMessage(LM.get(L"status_tabs_removed"), MessageStatus::Info);
        g_prevBufId = bufId;
        return;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CASE B: Padding not present (and _flowTabsActive == false) -> TURN ON
    // ══════════════════════════════════════════════════════════════════════════

    if (!flowTabsIntroDontShowEnabled) {
        bool dontShow = false;
        if (!showFlowTabsIntroDialog(dontShow))
            return; // user cancelled
        if (dontShow) {
            flowTabsIntroDontShowEnabled = true;
            saveSettings();
        }
    }

    // Need delimiter data
    if (lineDelimiterPositions.empty()) {
        showStatusMessage(LM.get(L"status_no_delimiters"), MessageStatus::Error);
        return;
    }

    ColumnTabs::CT_ColumnModelView model{};
    if (!buildCTModelFromMatrix(model)) {
        showStatusMessage(LM.get(L"status_model_build_failed"), MessageStatus::Error);
        return;
    }

    // Disable screen updates and logging during bulk operation
    bool wasLoggingEnabled = isLoggingEnabled;
    bool wasHighlighted = isColumnHighlighted;
    bool operationSuccess = false;
    bool earlySuccessPath = false;

    isLoggingEnabled = false;
    logChanges.clear();

    {
        // Lock screen updates during bulk text modification
        ScopedRedrawLock redrawLock(_hScintilla);

        {
            // We want NumericPadding + InsertAlignedPadding to be one single undo step.
            ScopedUndoAction undo(*this);

            // Step 1: numeric alignment (optional; modifies text, inserts numeric left padding)
            if (flowTabsNumericAlignEnabled) {
                ColumnTabs::CT_ApplyNumericPadding(_hScintilla, model, 0, static_cast<int>(model.Lines.size()) - 1);
                findAllDelimitersInDocument();
                if (!buildCTModelFromMatrix(model)) {
                    isLoggingEnabled = wasLoggingEnabled;
                    showStatusMessage(L"Numeric align: model rebuild failed", MessageStatus::Error);
                    return;
                }
            }

            // Step 2: insert Flow-Tab padding
            ColumnTabs::CT_AlignOptions opt{};
            opt.firstLine = 0;
            opt.lastLine = static_cast<int>(model.Lines.size()) - 1;

            const int spacePx = static_cast<int>(SendMessage(_hScintilla, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" "));
            opt.gapCells = 2;
            _flowPaddingPx = spacePx * opt.gapCells;
            opt.oneFlowTabOnly = true;

            bool nothingToAlign = false;
            if (!ColumnTabs::CT_InsertAlignedPadding(_hScintilla, model, opt, &nothingToAlign)) {
                if (ColumnTabs::CT_GetCurDocHasPads(_hScintilla)) {
                    // Numeric padding succeeded, InsertAlignedPadding failed but we have pads
                    _flowTabsActive = true;
                    if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
                        ::SetWindowText(h, L"⇤");

                    findAllDelimitersInDocument();
                    if (buildCTModelFromMatrix(model)) {
                        applyFlowTabStops(&model);
                    }

                    g_padBufs.insert(bufId);
                    g_prevBufId = bufId;
                    earlySuccessPath = true;
                    operationSuccess = true;
                }
                else {
                    isLoggingEnabled = wasLoggingEnabled;
                    showStatusMessage(LM.get(nothingToAlign ? L"status_nothing_to_align"
                        : L"status_padding_insert_failed"),
                        nothingToAlign ? MessageStatus::Info : MessageStatus::Error);
                    return;
                }
            }
            else {
                operationSuccess = true;
            }
        }

        // Rescan after edit (only if not early success path which already did it)
        if (operationSuccess && !earlySuccessPath) {
            findAllDelimitersInDocument();
        }
    }
    // ScopedRedrawLock released here - screen redraws once

    // Adjust ResultDock hit positions if padding was successfully inserted
    if (operationSuccess)
    {
        wchar_t pathBuf[MAX_PATH] = {};
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(pathBuf));
        std::string curFileUtf8 = Encoding::wstringToUtf8(pathBuf);

        ResultDock& dock = ResultDock::instance();
        if (dock.hasHitsForFile(curFileUtf8))
        {
            // Scan for newly inserted padding ranges (ColumnTabs indicator)
            std::vector<std::pair<Sci_Position, Sci_Position>> paddingRanges;
            const int ctInd = ColumnTabs::CT_GetIndicatorId();
            const Sci_Position docLen = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
            Sci_Position scanPos = 0;
            while (scanPos < docLen) {
                const Sci_Position end = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATOREND, ctInd, scanPos));
                if (end <= scanPos) break;
                if (static_cast<int>(::SendMessage(_hScintilla, SCI_INDICATORVALUEAT, ctInd, scanPos)) != 0) {
                    const Sci_Position start = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATORSTART, ctInd, scanPos));
                    if (end > start)
                        paddingRanges.emplace_back(start, end);
                }
                scanPos = end;
            }

            if (!paddingRanges.empty()) {
                dock.adjustHitPositionsForFlowTab(curFileUtf8, paddingRanges, /*added=*/true);
            }
        }
    }

    // Restore logging
    isLoggingEnabled = wasLoggingEnabled;

    // Handle early success path
    if (earlySuccessPath) {
        if (wasHighlighted) {
            handleHighlightColumnsInDocument();
        }
        showStatusMessage(LM.get(L"status_tabs_inserted"), MessageStatus::Success);
        return;
    }

    // Normal success path continues here
    const bool nowHasPads = ColumnTabs::CT_GetCurDocHasPads(_hScintilla);
    if (!nowHasPads && !ColumnTabs::CT_HasFlowTabStops()) {
        if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
            ::SetWindowText(h, L"⇥");
        _flowTabsActive = false;
        showStatusMessage(LM.get(L"status_nothing_to_align"), MessageStatus::Info);
        g_prevBufId = bufId;
        return;
    }

    if (!buildCTModelFromMatrix(model)) {
        showStatusMessage(LM.get(L"status_visual_fail"), MessageStatus::Error);
        return;
    }

    if (!applyFlowTabStops(&model)) {
        showStatusMessage(LM.get(L"status_visual_fail"), MessageStatus::Error);
    }

    _flowTabsActive = true;
    if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
        ::SetWindowText(h, L"⇤");

    // Re-highlight AFTER screen updates are enabled
    if (wasHighlighted) {
        handleHighlightColumnsInDocument();
    }

    showStatusMessage(LM.get(nowHasPads ? L"status_tabs_inserted"
        : L"status_tabs_aligned"),
        MessageStatus::Success);

    if (nowHasPads) g_padBufs.insert(bufId); else g_padBufs.erase(bufId);
    g_prevBufId = bufId;
}

void MultiReplace::handleDuplicatesButton()
{
    if (!validateDelimiterData()) {
        return;
    }

    // Always perform a fresh scan - user may have edited the document
    findAndMarkDuplicates();
}

void MultiReplace::findAndMarkDuplicates(bool showDialog)
{
    // Ensure we have the correct Scintilla handle
    pointerToScintilla();

    // Clear any previous marks and bookmarks first
    clearDuplicateMarks();

    // Re-parse column and delimiter data to ensure fresh data
    if (!parseColumnAndDelimiterData()) {
        showStatusMessage(LM.get(L"status_invalid_column_or_delimiter"), MessageStatus::Error);
        return;
    }

    // Get Match Case setting from panel and store it
    _duplicateMatchCase = (IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED);

    // Store scan criteria for validation rescan before delete
    _duplicateScanColumns = columnDelimiterData.columns;
    _duplicateScanDelimiter = columnDelimiterData.extendedDelimiter;

    // Store bookmark setting at scan time (to know if we created bookmarks)
    Settings settings = getSettings();
    _duplicateBookmarksEnabled = settings.duplicateBookmarksEnabled;

    // Perform the scan (this calls findAllDelimitersInDocument internally)
    if (!scanForDuplicates()) {
        // No duplicates found or no data - message already shown by scanForDuplicates
        return;
    }

    // Apply visual marking
    applyDuplicateMarks();

    // Show dialog asking whether to delete (only if requested)
    if (showDialog) {
        showDeleteDuplicatesDialog();
    }
}

bool MultiReplace::scanForDuplicates()
{
    // Ensure delimiter positions are current
    findAllDelimitersInDocument();

    const size_t lineCount = lineDelimiterPositions.size();
    if (lineCount <= CSVheaderLinesCount) {
        showStatusMessage(LM.get(L"status_no_data_for_duplicates"), MessageStatus::Info);
        return false;
    }

    // Extract column data for comparison (reuse existing infrastructure)
    std::vector<CombinedColumns> columnData = extractColumnData(CSVheaderLinesCount, lineCount);

    // Build comparison keys and find duplicates in single pass
    // Track count per key to identify groups
    std::unordered_map<std::string, std::pair<size_t, size_t>> keyInfo;  // key -> (firstLine, count)
    _markedDuplicateLines.clear();
    _markedDuplicateLines.reserve(lineCount / 4);

    const char KEY_SEP = '\x01';

    for (size_t dataIdx = 0; dataIdx < columnData.size(); ++dataIdx) {
        const size_t lineIdx = dataIdx + CSVheaderLinesCount;
        const auto& row = columnData[dataIdx];

        // Build comparison key
        std::string key;
        key.reserve(256);

        for (size_t colIdx = 0; colIdx < row.columns.size(); ++colIdx) {
            if (colIdx > 0) key += KEY_SEP;

            if (_duplicateMatchCase) {
                key += row.columns[colIdx].text;
            }
            else {
                key += StringUtils::toLowerUtf8(row.columns[colIdx].text);
            }
        }

        auto it = keyInfo.find(key);
        if (it == keyInfo.end()) {
            keyInfo[key] = { lineIdx, 1 };
        }
        else {
            it->second.second++;  // Increment count
            _markedDuplicateLines.push_back(lineIdx);
        }
    }

    // Count groups (keys with count > 1)
    _duplicateGroupCount = 0;
    for (const auto& kv : keyInfo) {
        if (kv.second.second > 1) {
            _duplicateGroupCount++;
        }
    }

    if (_markedDuplicateLines.empty()) {
        MessageBox(
            nppData._nppHandle,
            LM.getLPCW(L"status_no_duplicates_found"),
            LM.getLPCW(L"msgbox_title_delete_duplicates"),
            MB_OK | MB_ICONINFORMATION
        );
        return false;
    }

    return true;
}

bool MultiReplace::validateAndRescanIfNeeded()
{
    if (_markedDuplicateLines.empty() || _duplicateScanColumns.empty()) {
        return false;  // Nothing to validate
    }

    // Save original duplicate indices for comparison
    std::vector<size_t> originalIndices = _markedDuplicateLines;

    // Save current columnDelimiterData (user may have changed UI settings)
    auto savedColumns = columnDelimiterData.columns;
    auto savedDelimiter = columnDelimiterData.extendedDelimiter;
    auto savedInputColumns = columnDelimiterData.inputColumns;
    auto savedDelimiterLength = columnDelimiterData.delimiterLength;

    // Use stored scan criteria (from original scan)
    // Note: _duplicateMatchCase already contains the original Match Case setting
    columnDelimiterData.columns = _duplicateScanColumns;
    columnDelimiterData.extendedDelimiter = _duplicateScanDelimiter;
    columnDelimiterData.delimiterLength = _duplicateScanDelimiter.length();
    columnDelimiterData.inputColumns.clear();
    columnDelimiterData.inputColumns.assign(_duplicateScanColumns.begin(), _duplicateScanColumns.end());

    // Rescan delimiter positions with stored delimiter
    findAllDelimitersInDocument();

    // Rescan duplicates with stored criteria
    _markedDuplicateLines.clear();
    bool foundDuplicates = scanForDuplicates();

    // Restore original columnDelimiterData
    columnDelimiterData.columns = savedColumns;
    columnDelimiterData.extendedDelimiter = savedDelimiter;
    columnDelimiterData.inputColumns = savedInputColumns;
    columnDelimiterData.delimiterLength = savedDelimiterLength;

    // Check if rescan found same duplicates
    if (!foundDuplicates || _markedDuplicateLines != originalIndices) {
        // Document was modified - duplicates changed
        showStatusMessage(LM.get(L"status_document_modified_delete_cancelled"), MessageStatus::Error);

        // Clear all duplicate marks since they're no longer valid
        clearDuplicateMarks();

        return false;
    }

    return true;  // Duplicates unchanged, proceed with delete
}

void MultiReplace::applyDuplicateMarks()
{
    if (_markedDuplicateLines.empty() || _duplicateIndicatorId < 0) return;

    const int indicId = _duplicateIndicatorId;

    bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);
    int alpha = isDark ? EDITOR_MARK_ALPHA_DARK : EDITOR_MARK_ALPHA_LIGHT;
    int outlineAlpha = isDark ? EDITOR_OUTLINE_ALPHA_DARK : EDITOR_OUTLINE_ALPHA_LIGHT;
    COLORREF color = isDark ? DUPLICATE_MARKER_COLOR_DARK : DUPLICATE_MARKER_COLOR_LIGHT;

    send(SCI_INDICSETSTYLE, indicId, INDIC_FULLBOX);
    send(SCI_INDICSETFORE, indicId, color);
    send(SCI_INDICSETALPHA, indicId, alpha);
    send(SCI_INDICSETOUTLINEALPHA, indicId, outlineAlpha);
    send(SCI_INDICSETUNDER, indicId, TRUE);
    send(SCI_SETINDICATORCURRENT, indicId, 0);

    // Get bookmark marker ID from N++ (falls back to 20 for newer versions)
    constexpr UINT LOCAL_NPPM_GETBOOKMARKID = (WM_USER + 1000) + 113;
    LRESULT nppBookmarkId = ::SendMessage(nppData._nppHandle, LOCAL_NPPM_GETBOOKMARKID, 0, 0);
    int markerId = (nppBookmarkId > 0 && nppBookmarkId < 32) ? static_cast<int>(nppBookmarkId) : 20;

    // If bookmarks enabled: clear ALL bookmarks first, then set new ones on duplicates
    if (_duplicateBookmarksEnabled) {
        send(SCI_MARKERDELETEALL, markerId, 0);
    }

    // Apply visual indicators and bookmarks
    for (size_t lineIdx : _markedDuplicateLines) {
        LRESULT lineStart = send(SCI_POSITIONFROMLINE, lineIdx, 0);
        LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, lineIdx, 0);
        LRESULT lineLen = lineEnd - lineStart;

        if (lineLen > 0) {
            send(SCI_INDICATORFILLRANGE, lineStart, lineLen);
        }

        // Add bookmark if enabled
        if (_duplicateBookmarksEnabled) {
            send(SCI_MARKERADD, static_cast<uptr_t>(lineIdx), static_cast<sptr_t>(markerId));
        }
    }

    // Force Scintilla to redraw/refresh the editor view
    ::InvalidateRect(_hScintilla, nullptr, FALSE);
}

void MultiReplace::clearDuplicateMarks()
{
    // Always clear indicators (prophylactically, even if list is empty)
    if (_duplicateIndicatorId >= 0) {
        const int indicId = _duplicateIndicatorId;
        send(SCI_SETINDICATORCURRENT, indicId, 0);

        // Clear indicators on known duplicate lines
        const size_t currentLineCount = static_cast<size_t>(send(SCI_GETLINECOUNT, 0, 0));
        for (size_t lineIdx : _markedDuplicateLines) {
            if (lineIdx >= currentLineCount) continue;

            LRESULT lineStart = send(SCI_POSITIONFROMLINE, lineIdx, 0);
            LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, lineIdx, 0);
            LRESULT lineLen = lineEnd - lineStart;

            if (lineLen > 0) {
                send(SCI_INDICATORCLEARRANGE, lineStart, lineLen);
            }
        }
    }

    // If bookmark feature was enabled during scan, clear ALL bookmarks
    if (_duplicateBookmarksEnabled) {
        constexpr UINT LOCAL_NPPM_GETBOOKMARKID = (WM_USER + 1000) + 113;
        LRESULT nppBookmarkId = ::SendMessage(nppData._nppHandle, LOCAL_NPPM_GETBOOKMARKID, 0, 0);
        int markerId = (nppBookmarkId > 0 && nppBookmarkId < 32) ? static_cast<int>(nppBookmarkId) : 20;
        send(SCI_MARKERDELETEALL, markerId, 0);
    }

    _markedDuplicateLines.clear();
    _duplicateGroupCount = 0;

    // Reset stored scan criteria
    _duplicateScanColumns.clear();
    _duplicateScanDelimiter.clear();

    // Force Scintilla to redraw/refresh the editor view
    ::InvalidateRect(_hScintilla, nullptr, FALSE);
}

void MultiReplace::showDeleteDuplicatesDialog()
{
    const size_t duplicateCount = _markedDuplicateLines.size();
    const size_t groupCount = _duplicateGroupCount;

    // Get language strings
    std::wstring titleText = LM.get(L"msgbox_title_delete_duplicates");
    std::wstring questionText = LM.get(L"msgbox_duplicates_question");
    std::wstring statsText = LM.get(L"msgbox_duplicates_stats", { StringUtils::formatNumber(duplicateCount), StringUtils::formatNumber(groupCount) });
    std::wstring modeStr = _duplicateMatchCase ? LM.get(L"msgbox_duplicates_exact") : LM.get(L"msgbox_duplicates_ignoring");
    std::wstring modeText = LM.get(L"msgbox_duplicates_mode", { modeStr });
    std::wstring undoText = LM.get(L"msgbox_duplicates_undo");

    // Combine content: stats + mode + blank line + undo hint
    std::wstring contentText = statsText + L"\n" + modeText + L"\n\n" + undoText;

    std::wstring btnDeleteText = LM.get(L"msgbox_btn_delete_duplicates");
    std::wstring btnKeepText = LM.get(L"msgbox_btn_keep_marked");

    // Button IDs
    const int ID_BTN_DELETE = 100;
    const int ID_BTN_KEEP = 101;

    // Custom buttons with dynamic text
    TASKDIALOG_BUTTON buttons[] = {
        { ID_BTN_DELETE, btnDeleteText.c_str() },
        { ID_BTN_KEEP, btnKeepText.c_str() }
    };

    // Configure TaskDialog
    TASKDIALOGCONFIG tdc = { sizeof(TASKDIALOGCONFIG) };
    tdc.hwndParent = nppData._nppHandle;
    tdc.dwFlags = TDF_ALLOW_DIALOG_CANCELLATION | TDF_POSITION_RELATIVE_TO_WINDOW | TDF_SIZE_TO_CONTENT;
    tdc.pszWindowTitle = titleText.c_str();
    tdc.pszMainInstruction = questionText.c_str();
    tdc.pszContent = contentText.c_str();
    tdc.pszMainIcon = TD_WARNING_ICON;
    tdc.pButtons = buttons;
    tdc.cButtons = ARRAYSIZE(buttons);
    tdc.nDefaultButton = ID_BTN_KEEP;  // Safety: default to non-destructive action

    int nButtonPressed = 0;
    HRESULT hr = TaskDialogIndirect(&tdc, &nButtonPressed, nullptr, nullptr);

    if (SUCCEEDED(hr) && nButtonPressed == ID_BTN_DELETE) {
        deleteDuplicateLines();
    }
    // If Keep Marked, Cancel, or ESC: marks remain visible for user inspection
}

void MultiReplace::deleteDuplicateLines()
{
    if (_markedDuplicateLines.empty()) return;

    // Validate document state and rescan if needed (user may have edited while dialog was open)
    if (!validateAndRescanIfNeeded()) {
        return;  // No duplicates found after rescan or validation failed
    }

    const size_t deleteCount = _markedDuplicateLines.size();

    // Clear visual marks first
    if (_duplicateIndicatorId >= 0) {
        const int indicId = _duplicateIndicatorId;
        ::SendMessage(_hScintilla, SCI_SETINDICATORCURRENT, indicId, 0);
        const size_t currentLineCount = static_cast<size_t>(send(SCI_GETLINECOUNT, 0, 0));
        for (size_t lineIdx : _markedDuplicateLines) {
            if (lineIdx >= currentLineCount) continue;
            LRESULT lineStart = send(SCI_POSITIONFROMLINE, lineIdx, 0);
            LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, lineIdx, 0);
            if (lineEnd > lineStart) {
                ::SendMessage(_hScintilla, SCI_INDICATORCLEARRANGE, lineStart, lineEnd - lineStart);
            }
        }
    }

    // Delete the marked duplicate lines
    runCsvWithFlowTabs(CsvOp::DeleteColumns, [&]() -> bool {
        ScopedUndoAction undo(*this);

        // Sort descending - delete from bottom to top to keep indices valid
        std::vector<size_t> linesToDelete = _markedDuplicateLines;
        std::sort(linesToDelete.begin(), linesToDelete.end(), std::greater<size_t>());

        const size_t totalLines = static_cast<size_t>(send(SCI_GETLINECOUNT, 0, 0));

        for (size_t lineIdx : linesToDelete) {
            if (lineIdx >= totalLines) continue;

            LRESULT lineStart = send(SCI_POSITIONFROMLINE, lineIdx, 0);
            LRESULT lineEnd;

            if (lineIdx + 1 < totalLines) {
                lineEnd = send(SCI_POSITIONFROMLINE, lineIdx + 1, 0);
            }
            else {
                lineEnd = send(SCI_GETLINEENDPOSITION, lineIdx, 0);
                if (lineIdx > 0) {
                    lineStart = send(SCI_GETLINEENDPOSITION, lineIdx - 1, 0);
                }
            }

            LRESULT deleteLen = lineEnd - lineStart;
            if (deleteLen > 0) {
                send(SCI_DELETERANGE, lineStart, deleteLen, false);
            }

            updateUnsortedDocument(lineIdx, 1, ChangeType::Delete);
        }

        return true;
        });

    // Clear all bookmarks if bookmark feature was enabled
    if (_duplicateBookmarksEnabled) {
        constexpr UINT LOCAL_NPPM_GETBOOKMARKID = (WM_USER + 1000) + 113;
        LRESULT nppBookmarkId = ::SendMessage(nppData._nppHandle, LOCAL_NPPM_GETBOOKMARKID, 0, 0);
        int markerId = (nppBookmarkId > 0 && nppBookmarkId < 32) ? static_cast<int>(nppBookmarkId) : 20;
        send(SCI_MARKERDELETEALL, markerId, 0);
    }

    // Clear state
    _markedDuplicateLines.clear();
    _duplicateGroupCount = 0;

    // Refresh delimiter positions
    findAllDelimitersInDocument();

    // Show success message
    showStatusMessage(LM.get(L"status_duplicates_deleted", { std::to_wstring(deleteCount) }), MessageStatus::Success);
}

void MultiReplace::clearFlowTabsIfAny()
{
    pointerToScintilla();
    if (!_hScintilla) return;

    ColumnTabs::CT_SetIndicatorId(30);

    const bool hadPad = ColumnTabs::CT_HasAlignedPadding(_hScintilla);
    const bool hadVis = ColumnTabs::CT_HasFlowTabStops();

    // Collect padding ranges BEFORE removal for ResultDock position adjustment
    std::vector<std::pair<Sci_Position, Sci_Position>> paddingRanges;
    std::string curFileUtf8;
    bool hasResultHits = false;

    if (hadPad)
    {
        wchar_t pathBuf[MAX_PATH] = {};
        ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(pathBuf));
        curFileUtf8 = Encoding::wstringToUtf8(pathBuf);
        hasResultHits = ResultDock::instance().hasHitsForFile(curFileUtf8);

        if (hasResultHits) {
            const int ctInd = ColumnTabs::CT_GetIndicatorId();
            const Sci_Position docLen = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_GETLENGTH, 0, 0));
            Sci_Position scanPos = 0;
            while (scanPos < docLen) {
                const Sci_Position end = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATOREND, ctInd, scanPos));
                if (end <= scanPos) break;
                if (static_cast<int>(::SendMessage(_hScintilla, SCI_INDICATORVALUEAT, ctInd, scanPos)) != 0) {
                    const Sci_Position start = static_cast<Sci_Position>(::SendMessage(_hScintilla, SCI_INDICATORSTART, ctInd, scanPos));
                    if (end > start)
                        paddingRanges.emplace_back(start, end);
                }
                scanPos = end;
            }
        }
    }

    if (hadPad) ColumnTabs::CT_RemoveAlignedPadding(_hScintilla);
    if (hadVis) ColumnTabs::CT_DisableFlowTabStops(_hScintilla, /*restoreManual=*/false);
    if (hadVis) ColumnTabs::CT_ResetFlowVisualState();

    if (hadPad) {
        // Keep both gates in sync: column-lib doc flag + fast buffer gate
        ColumnTabs::CT_SetCurDocHasPads(_hScintilla, false);

        const BufferId bufId = (BufferId)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);
        g_padBufs.erase(bufId);
        findAllDelimitersInDocument(); // offsets changed

        // Adjust ResultDock hit positions: padding was removed, positions shift backward
        if (hasResultHits && !paddingRanges.empty()) {
            ResultDock::instance().adjustHitPositionsForFlowTab(curFileUtf8, paddingRanges, /*added=*/false);
        }
    }

    if (hadPad || hadVis) {
        _flowTabsActive = false;
        if (HWND h = ::GetDlgItem(_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
            ::SetWindowText(h, L"⇥");
        forceWrapRecalculation();
    }
}

bool MultiReplace::runCsvWithFlowTabs(CsvOp op, const std::function<bool()>& body)
{
    if (lineDelimiterPositions.empty())
        findAllDelimitersInDocument();

    enum class EtabsMode { Off, Visual, Padding };

    const EtabsMode mode =
        (!_flowTabsActive) ? EtabsMode::Off
        : (ColumnTabs::CT_GetCurDocHasPads(_hScintilla) ? EtabsMode::Padding
            : EtabsMode::Visual);

    const bool modifiesText =
        (op == CsvOp::Sort) || (op == CsvOp::DeleteColumns);

    // ════════════════════════════════════════════════════════════════════════
    // Batch rendering: suspend redraw during multiple text modifications
    // ════════════════════════════════════════════════════════════════════════
    const bool needsRedrawLock = (mode == EtabsMode::Padding && modifiesText);

    if (needsRedrawLock) {
        ::SendMessage(_hScintilla, WM_SETREDRAW, FALSE, 0);
    }

    // If Padding + modifying op → unpad to get canonical text
    if (mode == EtabsMode::Padding && modifiesText && ColumnTabs::CT_GetCurDocHasPads(_hScintilla)) {
        ColumnTabs::CT_RemoveAlignedPadding(_hScintilla);
        findAllDelimitersInDocument();
    }

    // Run the actual operation
    const bool ok = body ? body() : true;

    findAllDelimitersInDocument();

    // Reflow presentation
    switch (mode) {
    case EtabsMode::Visual: {
        if (!_flowTabsActive) break;

        ColumnTabs::CT_ColumnModelView model{};
        if (buildCTModelFromMatrix(model)) {
            if (!model.Lines.empty()) {
                const int gapPx = _flowPaddingPx;
                ColumnTabs::CT_ApplyFlowTabStopsAll(_hScintilla, model, gapPx);
            }
        }
        break;
    }
    case EtabsMode::Padding: {
        if (!_flowTabsActive) break;

        ColumnTabs::CT_ColumnModelView model{};
        if (buildCTModelFromMatrix(model) && !model.Lines.empty()) {

            if (flowTabsNumericAlignEnabled) {
                ColumnTabs::CT_ApplyNumericPadding(_hScintilla, model, 0, static_cast<int>(model.Lines.size()) - 1);
                findAllDelimitersInDocument();
                if (!buildCTModelFromMatrix(model)) {
                    break;
                }
            }

            ColumnTabs::CT_AlignOptions a{};
            a.firstLine = 0;
            a.lastLine = static_cast<int>(model.Lines.size()) - 1;
            const int spacePx = static_cast<int>(SendMessage(_hScintilla, SCI_TEXTWIDTH, STYLE_DEFAULT, (sptr_t)" "));
            a.gapCells = (spacePx > 0) ? (_flowPaddingPx / spacePx) : 2;
            a.oneFlowTabOnly = true;

            bool nothingToAlign = false;
            (void)ColumnTabs::CT_InsertAlignedPadding(_hScintilla, model, a, &nothingToAlign);

            if (ColumnTabs::CT_GetCurDocHasPads(_hScintilla)) {
                findAllDelimitersInDocument();
            }
        }
        break;
    }
    case EtabsMode::Off:
    default:
        break;
    }

    // ════════════════════════════════════════════════════════════════════════
    // Resume rendering and repaint once
    // ════════════════════════════════════════════════════════════════════════
    if (needsRedrawLock) {
        ::SendMessage(_hScintilla, WM_SETREDRAW, TRUE, 0);
        ::InvalidateRect(_hScintilla, nullptr, TRUE);
    }

    return ok;
}

bool MultiReplace::showFlowTabsIntroDialog(bool& dontShowFlag) const
{
    // Localized strings (single body + checkbox; buttons from INI with fallback)
    const std::wstring body = LM.get(L"msgbox_flowtabs_intro_body");
    const std::wstring chkLabel = LM.get(L"msgbox_flowtabs_intro_checkbox");
    std::wstring okTxt = LM.get(L"msgbox_button_ok");     if (okTxt.empty())     okTxt = L"OK";
    std::wstring cancelTxt = LM.get(L"msgbox_button_cancel"); if (cancelTxt.empty()) cancelTxt = L"Cancel";
    const std::wstring title = LM.get(L"msgbox_title_info");

    if (body.empty())
        return true; // nothing to show

    // --- Build an in-memory DLGTEMPLATE (no .rc, USER32 only) ---
    // Layout (DLUs)
    const short W = 320, H = 112, Mx = 7, My = 7;
    const short Tw = W - 2 * Mx;
    const short Ch = 12, Bw = 60, Bh = 14, Bp = 6;
    const short Gt = 4, Gb = 4;
    const short ThMax = H - (2 * My + Bh + Ch + Gt + Gb);
    const short Th = (short)std::max<short>(40, ThMax);

    // Size the buffer safely (text length dominates)
    const size_t needBytes =
        1024 + // header/overhead
        (title.size() + 1 +
            body.size() + 1 +
            chkLabel.size() + 1 +
            okTxt.size() + 1 +
            cancelTxt.size() + 1) * sizeof(wchar_t);

    std::vector<BYTE> buf((std::max)(needBytes, (size_t)4096)); // min 4KB
    DLGTEMPLATE* dlg = reinterpret_cast<DLGTEMPLATE*>(buf.data());
    dlg->style = DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME;
    dlg->cdit = 4;  // STATIC, CHECKBOX, OK, Cancel
    dlg->x = 0; dlg->y = 0; dlg->cx = W; dlg->cy = H;

    BYTE* p = buf.data() + sizeof(DLGTEMPLATE);
    auto W16 = [&](WORD v) { *reinterpret_cast<WORD*>(p) = v; p += sizeof(WORD); };
    auto WSTR = [&](const std::wstring& s) {
        const size_t n = (s.size() + 1) * sizeof(wchar_t);
        memcpy(p, s.c_str(), n);
        p += n;
        };
    auto AL = [&]() {                 // correct DWORD align on x86/x64
        uintptr_t a = reinterpret_cast<uintptr_t>(p);
        a = (a + 3) & ~(uintptr_t)3;
        p = reinterpret_cast<BYTE*>(a);
        };

    // Menu=0, Class=0, Title
    W16(0); W16(0); WSTR(title.empty() ? L"" : title);
    // DS_SETFONT (point size + face name)
    W16(9); WSTR(L"Segoe UI");

    // Item 1: STATIC (multiline body)
    AL(); auto* it = reinterpret_cast<DLGITEMTEMPLATE*>(p);
    it->style = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    it->dwExtendedStyle = 0;
    it->x = Mx; it->y = My; it->cx = Tw; it->cy = Th; it->id = 1001;
    p += sizeof(DLGITEMTEMPLATE);
    W16(0xFFFF); W16(0x0082); // STATIC class
    WSTR(body);
    W16(0); // no creation data

    // Item 2: CHECKBOX
    AL(); it = reinterpret_cast<DLGITEMTEMPLATE*>(p);
    it->style = WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX;
    it->dwExtendedStyle = 0;
    it->x = Mx; it->y = My + Th + 4; it->cx = Tw - 2; it->cy = Ch; it->id = 1002;
    p += sizeof(DLGITEMTEMPLATE);
    W16(0xFFFF); W16(0x0080); // BUTTON class
    WSTR(chkLabel.empty() ? L"" : chkLabel);
    W16(0);

    // Buttons (right aligned)
    const short By = H - My - Bh, BxC = W - Mx - Bw, BxO = BxC - Bw - Bp;

    // Item 3: OK
    AL(); it = reinterpret_cast<DLGITEMTEMPLATE*>(p);
    it->style = WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON;
    it->dwExtendedStyle = 0;
    it->x = BxO; it->y = By; it->cx = Bw; it->cy = Bh; it->id = IDOK;
    p += sizeof(DLGITEMTEMPLATE);
    W16(0xFFFF); W16(0x0080); // BUTTON
    WSTR(okTxt);
    W16(0);

    // Item 4: Cancel
    AL(); it = reinterpret_cast<DLGITEMTEMPLATE*>(p);
    it->style = WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON;
    it->dwExtendedStyle = 0;
    it->x = BxC; it->y = By; it->cx = Bw; it->cy = Bh; it->id = IDCANCEL;
    p += sizeof(DLGITEMTEMPLATE);
    W16(0xFFFF); W16(0x0080); // BUTTON
    WSTR(cancelTxt);
    W16(0);

    struct Payload { bool* pDont; } payload{ const_cast<bool*>(&dontShowFlag) };

    auto DlgProc = [](HWND h, UINT m, WPARAM w, LPARAM l) -> INT_PTR {
        auto* pay = reinterpret_cast<Payload*>(::GetWindowLongPtrW(h, GWLP_USERDATA));
        switch (m) {
        case WM_INITDIALOG:
            ::SetWindowLongPtrW(h, GWLP_USERDATA, l);
            return TRUE;
        case WM_COMMAND:
            if (LOWORD(w) == IDOK || LOWORD(w) == IDCANCEL) {
                if (LOWORD(w) == IDOK && pay && pay->pDont) {
                    *pay->pDont = (BST_CHECKED == ::IsDlgButtonChecked(h, 1002));
                }
                ::EndDialog(h, LOWORD(w));
                return TRUE;
            }
            break;
        }
        return FALSE;
        };

    const INT_PTR rc = ::DialogBoxIndirectParamW(_hInst, dlg, _hSelf, (DLGPROC)DlgProc, (LPARAM)&payload);
    return (rc == IDOK);
}

ViewState MultiReplace::saveViewState() const {
    ViewState s{};
    s.firstVisibleLine = static_cast<int>(send(SCI_GETFIRSTVISIBLELINE));
    s.xOffset = static_cast<int>(send(SCI_GETXOFFSET));
    s.caret = (Sci_Position)send(SCI_GETCURRENTPOS);
    s.anchor = (Sci_Position)send(SCI_GETANCHOR);
    s.wrapMode = static_cast<int>(send(SCI_GETWRAPMODE));
    return s;
}

void MultiReplace::restoreViewStateExact(const ViewState& s) {
    // restore selection/caret first to avoid re-centering side effects
    send(SCI_SETSEL, s.anchor, s.caret);
    send(SCI_SETXOFFSET, s.xOffset, 0);

    const int curFirst = static_cast<int>(send(SCI_GETFIRSTVISIBLELINE));
    if (curFirst != s.firstVisibleLine) {
        send(SCI_LINESCROLL, 0, s.firstVisibleLine - curFirst);
    }
}

#pragma endregion


#pragma region CSV Sort

std::vector<CombinedColumns> MultiReplace::extractColumnData(SIZE_T startLine, SIZE_T endLine)
{
    std::vector<CombinedColumns> combinedData;
    const size_t numLines = endLine - startLine;
    combinedData.reserve(numLines);

    // ══════════════════════════════════════════════════════════════════════
    // OPTIMIZATION: Pre-cache all line positions
    // ══════════════════════════════════════════════════════════════════════
    std::vector<LRESULT> lineStarts(numLines);
    for (SIZE_T i = startLine; i < endLine; ++i) {
        lineStarts[i - startLine] = send(SCI_POSITIONFROMLINE, i, 0);
    }

    const size_t numColumns = columnDelimiterData.inputColumns.size();

    for (SIZE_T i = startLine; i < endLine; ++i) {
        const size_t lineIdx = i - startLine;
        const auto& lineInfo = lineDelimiterPositions[i];

        // Use cached line start position
        const LRESULT lineStartPos = lineStarts[lineIdx];
        const LRESULT lineEndPos = lineStartPos + lineInfo.lineLength;

        // Read the entire line content in one call
        const size_t currentLineLength = static_cast<size_t>(lineInfo.lineLength);
        lineBuffer.resize(currentLineLength + 1, '\0');

        {
            Sci_TextRangeFull tr;
            tr.chrg.cpMin = lineStartPos;
            tr.chrg.cpMax = lineEndPos;
            tr.lpstrText = lineBuffer.data();
            send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
        }

        CombinedColumns rowData;
        rowData.columns.resize(numColumns);

        for (size_t columnIndex = 0; columnIndex < numColumns; ++columnIndex) {
            const SIZE_T columnNumber = columnDelimiterData.inputColumns[columnIndex];

            // Determine start position of this column
            LRESULT startPos;
            if (columnNumber == 1) {
                startPos = lineStartPos;
            }
            else if (columnNumber - 2 < lineInfo.positions.size()) {
                startPos = lineStartPos + lineInfo.positions[columnNumber - 2].offsetInLine
                    + columnDelimiterData.delimiterLength;
            }
            else {
                continue; // Skip invalid columns
            }

            // Determine end position
            LRESULT endPos;
            if (columnNumber - 1 < lineInfo.positions.size()) {
                endPos = lineStartPos + lineInfo.positions[columnNumber - 1].offsetInLine;
            }
            else {
                endPos = lineEndPos;
            }

            // Map to local offsets in line buffer
            const size_t localStart = static_cast<size_t>(startPos - lineStartPos);
            const size_t localEnd = static_cast<size_t>(endPos - lineStartPos);

            // Extract column text
            if (localStart < localEnd && localEnd <= currentLineLength) {
                const size_t colSize = localEnd - localStart;
                std::string columnText(lineBuffer.data() + localStart, colSize);

                // Remove trailing newlines
                while (!columnText.empty() && (columnText.back() == '\n' || columnText.back() == '\r')) {
                    columnText.pop_back();
                }

                rowData.columns[columnIndex].text = std::move(columnText);
            }
        }

        combinedData.push_back(std::move(rowData));
    }

    // NOTE: detectNumericColumns is now called by sortRowsByColumn after sanitization
    // This avoids duplicate processing

    return combinedData;
}

void MultiReplace::sortRowsByColumn(SortDirection sortDirection)
{
    // Validate CSV delimiter config
    if (!columnDelimiterData.isValid()) {
        showStatusMessage(LM.get(L"status_invalid_column_or_delimiter"), MessageStatus::Error);
        return;
    }

    // Run under the ETabs layer (handles PAD remove/reinsert or VIS re-apply)
    runCsvWithFlowTabs(CsvOp::Sort, [&]() -> bool {
        ScopedUndoAction undo(*this);

        const size_t lineCount = lineDelimiterPositions.size();
        if (lineCount <= CSVheaderLinesCount) {
            return true; // Nothing to sort
        }

        // Build initial order
        std::vector<size_t> tempOrder(lineCount);
        std::iota(tempOrder.begin(), tempOrder.end(), 0);

        // Extract columns after headers (detectNumericColumns is called inside)
        std::vector<CombinedColumns> combinedData =
            extractColumnData(CSVheaderLinesCount, lineCount);

        // Sanitize in-place (no copy!) and prepare for sorting
        auto sanitize = [](std::string& s) {
            if (s.empty()) return;
            size_t b = 0, e = s.size();
            while (b < e && (s[b] == ' ' || s[b] == '\t')) ++b;
            while (e > b && (s[e - 1] == ' ' || s[e - 1] == '\t')) --e;
            if (e < s.size()) s.erase(e);
            if (b > 0) s.erase(0, b);
            };

        // Sanitize and re-detect numeric (single pass, in-place)
        for (auto& row : combinedData) {
            for (auto& col : row.columns) {
                sanitize(col.text);
                col.isNumeric = false;
                col.numericValue = 0.0;
            }
        }
        detectNumericColumns(combinedData);  // Single call with wstring caching

        // Stable sort preserves original document order for rows with equal keys
        std::stable_sort(tempOrder.begin() + CSVheaderLinesCount, tempOrder.end(),
            [&](size_t a, size_t b) {
                const size_t ia = a - CSVheaderLinesCount;
                const size_t ib = b - CSVheaderLinesCount;
                const auto& ra = combinedData[ia];
                const auto& rb = combinedData[ib];

                for (size_t colIndex = 0; colIndex < columnDelimiterData.inputColumns.size(); ++colIndex) {
                    const int cmp = compareColumnValue(ra.columns[colIndex], rb.columns[colIndex]);
                    if (cmp != 0)
                        return (sortDirection == SortDirection::Ascending) ? (cmp < 0) : (cmp > 0);
                }
                return false; // Keep original order
            }
        );

        // Update originalLineOrder if tracking is used
        if (!originalLineOrder.empty()) {
            std::vector<size_t> newOrder(originalLineOrder.size());
            for (size_t i = 0; i < tempOrder.size(); ++i)
                newOrder[i] = originalLineOrder[tempOrder[i]];
            originalLineOrder = std::move(newOrder);
        }
        else {
            originalLineOrder = tempOrder;
        }

        // Apply new order in editor
        reorderLinesInScintilla(tempOrder);

        return true;
        });
}

void MultiReplace::reorderLinesInScintilla(const std::vector<size_t>& sortedIndex) {
    const std::string lineBreak = getEOLStyle();
    const size_t lineCount = sortedIndex.size();

    isSortedColumn = false; // Stop logging changes

    // ══════════════════════════════════════════════════════════════════════
    // OPTIMIZATION: Read entire document once, then extract lines
    // ══════════════════════════════════════════════════════════════════════
    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
    send(SCI_GETTEXT, static_cast<WPARAM>(docLen + 1), reinterpret_cast<sptr_t>(fullText.data()));
    fullText.resize(static_cast<size_t>(docLen));

    // Pre-cache all line positions
    std::vector<Sci_Position> lineStarts(lineCount);
    std::vector<Sci_Position> lineEnds(lineCount);
    for (size_t i = 0; i < lineCount; ++i) {
        lineStarts[i] = static_cast<Sci_Position>(send(SCI_POSITIONFROMLINE, i, 0));
        lineEnds[i] = static_cast<Sci_Position>(send(SCI_GETLINEENDPOSITION, i, 0));
    }

    // Calculate total size needed
    size_t totalSize = 0;
    for (size_t i = 0; i < lineCount; ++i) {
        const size_t idx = sortedIndex[i];
        totalSize += static_cast<size_t>(lineEnds[idx] - lineStarts[idx]);
        if (i < lineCount - 1) {
            totalSize += lineBreak.size();
        }
    }

    // Build combined lines using pre-read document
    std::string combinedLines;
    combinedLines.reserve(totalSize);

    for (size_t i = 0; i < lineCount; ++i) {
        const size_t idx = sortedIndex[i];
        const Sci_Position start = lineStarts[idx];
        const Sci_Position end = lineEnds[idx];

        if (end > start && static_cast<size_t>(start) < fullText.size()) {
            size_t len = static_cast<size_t>(end - start);
            if (static_cast<size_t>(start) + len > fullText.size()) {
                len = fullText.size() - static_cast<size_t>(start);
            }
            combinedLines.append(fullText, static_cast<size_t>(start), len);
        }

        if (i < lineCount - 1) {
            combinedLines += lineBreak;
        }
    }

    // Clear and re-insert
    send(SCI_CLEARALL);
    send(SCI_APPENDTEXT, combinedLines.length(), reinterpret_cast<sptr_t>(combinedLines.c_str()));

    isSortedColumn = true; // Ready for logging changes
}

void MultiReplace::restoreOriginalLineOrder(const std::vector<size_t>& originalOrder) {

    size_t totalLineCount = static_cast<size_t>(send(SCI_GETLINECOUNT));

    if (originalOrder.empty() || originalOrder.size() != totalLineCount) {
        return;
    }

    // Normalize indices to [0..n-1] range (closes gaps from deleted lines)
    std::vector<size_t> normalizedOrder = originalOrder;

    auto maxElementIt = std::max_element(originalOrder.begin(), originalOrder.end());
    if (maxElementIt != originalOrder.end() && *maxElementIt >= totalLineCount) {
        // Gaps exist - normalization needed
        std::vector<size_t> sortedIndices = originalOrder;
        std::sort(sortedIndices.begin(), sortedIndices.end());

        std::unordered_map<size_t, size_t> indexMapping;
        for (size_t i = 0; i < sortedIndices.size(); ++i) {
            indexMapping[sortedIndices[i]] = i;
        }

        for (size_t i = 0; i < normalizedOrder.size(); ++i) {
            normalizedOrder[i] = indexMapping[normalizedOrder[i]];
        }
    }

    // Validate
    maxElementIt = std::max_element(normalizedOrder.begin(), normalizedOrder.end());
    if (maxElementIt == normalizedOrder.end() || *maxElementIt != totalLineCount - 1) {
        return;
    }

    std::vector<bool> seen(totalLineCount, false);
    for (size_t idx : normalizedOrder) {
        if (idx >= totalLineCount || seen[idx]) {
            return;
        }
        seen[idx] = true;
    }

    // Perform restore using bulk read (same pattern as reorderLinesInScintilla)
    const std::string lineBreak = getEOLStyle();

    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    std::string fullText(static_cast<size_t>(docLen) + 1, '\0');
    send(SCI_GETTEXT, static_cast<WPARAM>(docLen + 1), reinterpret_cast<sptr_t>(fullText.data()));
    fullText.resize(static_cast<size_t>(docLen));

    // Pre-cache all line positions
    std::vector<Sci_Position> lineStarts(totalLineCount);
    std::vector<Sci_Position> lineEnds(totalLineCount);
    for (size_t i = 0; i < totalLineCount; ++i) {
        lineStarts[i] = static_cast<Sci_Position>(send(SCI_POSITIONFROMLINE, i, 0));
        lineEnds[i] = static_cast<Sci_Position>(send(SCI_GETLINEENDPOSITION, i, 0));
    }

    // Build reordered document: line at position i goes to normalizedOrder[i]
    // We need to build output in order 0..n-1, so find which source line maps to each target
    std::vector<size_t> inverseOrder(totalLineCount);
    for (size_t i = 0; i < totalLineCount; ++i) {
        inverseOrder[normalizedOrder[i]] = i;
    }

    // Calculate total size
    size_t totalSize = 0;
    for (size_t i = 0; i < totalLineCount; ++i) {
        const size_t srcIdx = inverseOrder[i];
        totalSize += static_cast<size_t>(lineEnds[srcIdx] - lineStarts[srcIdx]);
        if (i < totalLineCount - 1) {
            totalSize += lineBreak.size();
        }
    }

    // Build combined string
    std::string combinedLines;
    combinedLines.reserve(totalSize);

    for (size_t i = 0; i < totalLineCount; ++i) {
        const size_t srcIdx = inverseOrder[i];
        const Sci_Position start = lineStarts[srcIdx];
        const Sci_Position end = lineEnds[srcIdx];

        if (end > start && static_cast<size_t>(start) < fullText.size()) {
            size_t len = static_cast<size_t>(end - start);
            if (static_cast<size_t>(start) + len > fullText.size()) {
                len = fullText.size() - static_cast<size_t>(start);
            }
            combinedLines.append(fullText, static_cast<size_t>(start), len);
        }

        if (i < totalLineCount - 1) {
            combinedLines += lineBreak;
        }
    }

    send(SCI_CLEARALL);
    send(SCI_APPENDTEXT, combinedLines.length(), reinterpret_cast<sptr_t>(combinedLines.c_str()));
}

void MultiReplace::extractLineContent(size_t idx, std::string& content, const std::string& lineBreak) {
    LRESULT lineStart = send(SCI_POSITIONFROMLINE, idx);
    LRESULT lineEnd = send(SCI_GETLINEENDPOSITION, idx);
    std::vector<char> buffer(static_cast<size_t>(lineEnd - lineStart) + 1);
    Sci_TextRangeFull tr{ lineStart, lineEnd, buffer.data() };
    send(SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
    content.assign(buffer.begin(), buffer.end() - 1); // Remove the null terminator
    content += lineBreak;
}

void MultiReplace::UpdateSortButtonSymbols() {
    HWND hwndAscButton = GetDlgItem(_hSelf, IDC_COLUMN_SORT_ASC_BUTTON);
    HWND hwndDescButton = GetDlgItem(_hSelf, IDC_COLUMN_SORT_DESC_BUTTON);

    switch (currentSortState) {
    case SortDirection::Unsorted:
        SetWindowText(hwndAscButton, symbolSortAsc);
        SetWindowText(hwndDescButton, symbolSortDesc);
        break;
    case SortDirection::Ascending:
        SetWindowText(hwndAscButton, symbolSortAscUnsorted);
        SetWindowText(hwndDescButton, symbolSortDesc);
        break;
    case SortDirection::Descending:
        SetWindowText(hwndAscButton, symbolSortAsc);
        SetWindowText(hwndDescButton, symbolSortDescUnsorted);
        break;
    }
}

void MultiReplace::handleSortStateAndSort(SortDirection direction) {

    if (!validateDelimiterData()) {
        return;
    }

    // Snapshot view state once for this sort/unsort toggle
    const ViewState vs = saveViewState();

    if ((direction == SortDirection::Ascending && currentSortState == SortDirection::Ascending) ||
        (direction == SortDirection::Descending && currentSortState == SortDirection::Descending)) {
        isSortedColumn = false; //Disable logging of changes
        if (!originalLineOrder.empty()) {
            runCsvWithFlowTabs(CsvOp::Sort, [&]() -> bool {
                ScopedUndoAction undo(*this);
                restoreOriginalLineOrder(originalLineOrder);
                return true;
                });
        }
        currentSortState = SortDirection::Unsorted;
        originalLineOrder.clear();
    }
    else {
        currentSortState = (direction == SortDirection::Ascending) ? SortDirection::Ascending : SortDirection::Descending;
        if (columnDelimiterData.isValid()) {
            sortRowsByColumn(direction);
        }
    }
    // Restore view state so the viewport stays on the same top line
    restoreViewStateExact(vs);
}

void MultiReplace::updateUnsortedDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType) {
    if (!isSortedColumn || lineNumber > originalLineOrder.size()) {
        return;
    }

    switch (changeType) {
    case ChangeType::Insert: {
        // New lines get indices > all existing ones (will end up at the end on restore)
        size_t maxIndex = originalLineOrder.empty()
            ? 0
            : (*std::max_element(originalLineOrder.begin(), originalLineOrder.end())) + 1;

        std::vector<size_t> newIndices;
        newIndices.reserve(blockCount);

        for (SIZE_T i = 0; i < blockCount; ++i) {
            newIndices.push_back(maxIndex + i);
        }

        originalLineOrder.insert(
            originalLineOrder.begin() + lineNumber,
            newIndices.begin(),
            newIndices.end()
        );
        break;
    }

    case ChangeType::Delete: {
        // Simply remove entries - do NOT adjust remaining indices!
        // Gaps will be closed by normalization during restore.
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > originalLineOrder.size()) {
            endPos = originalLineOrder.size();
        }

        if (lineNumber < originalLineOrder.size()) {
            originalLineOrder.erase(
                originalLineOrder.begin() + lineNumber,
                originalLineOrder.begin() + endPos
            );
        }
        break;
    }

    case ChangeType::Modify:
    default:
        break;
    }
}

void MultiReplace::detectNumericColumns(std::vector<CombinedColumns>& data)
{
    if (data.empty()) return;
    const size_t colCount = data[0].columns.size();

    for (size_t col = 0; col < colCount; ++col) {
        for (auto& row : data) {
            ColumnValue& cv = row.columns[col];

            // Skip empty fields
            if (cv.text.empty()) {
                cv.textW.clear();
                continue;
            }

            // Try numeric classification (modifies cv.text in-place only if numeric)
            if (normalizeAndValidateNumber(cv.text)) {
                cv.isNumeric = true;
                cv.numericValue = std::stod(cv.text);
            }

            // Cache wide string for string comparison (done once, not per-compare)
            cv.textW = Encoding::utf8ToWString(cv.text);
        }
    }
}

int MultiReplace::compareColumnValue(const ColumnValue& left, const ColumnValue& right)
{
    // If one is numeric and the other is not, numeric values come first
    if (left.isNumeric != right.isNumeric) {
        return left.isNumeric ? -1 : 1;
    }

    // If both are numeric, compare numerical values
    if (left.isNumeric) {
        if (left.numericValue < right.numericValue) return -1;
        if (left.numericValue > right.numericValue) return 1;
        return 0;
    }

    // Use cached wide strings (no conversion during sort!)
    return lstrcmpiW(left.textW.c_str(), right.textW.c_str());
}

#pragma endregion


#pragma region Scope

bool MultiReplace::parseColumnAndDelimiterData() {
    // Retrieve data from dialog items
    std::wstring columnDataString = getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT);
    std::wstring delimiterData = getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT);
    std::wstring quoteCharString = getTextFromDialogItem(_hSelf, IDC_QUOTECHAR_EDIT);

    // Remove invalid delimiter characters (\n, \r)
    std::vector<std::wstring> stringsToRemove = { L"\\n", L"\\r" };
    for (const auto& strToRemove : stringsToRemove) {
        std::wstring::size_type pos = 0;
        while ((pos = delimiterData.find(strToRemove, pos)) != std::wstring::npos) {
            delimiterData.erase(pos, strToRemove.length());
        }
    }

    // Convert delimiter and quote character to standard strings
    std::string extendedDelimiter = convertAndExtendW(delimiterData, true);
    std::string quoteCharConverted = convertAndExtendW(quoteCharString, false);

    // Check for changes BEFORE modifying existing values
    // Cannot be used in LUA as it is Scintilla encoded and not UTF8
    bool delimiterChanged = (columnDelimiterData.extendedDelimiter != extendedDelimiter);
    bool quoteCharChanged = (columnDelimiterData.quoteChar != quoteCharConverted);

    // Trim leading and trailing commas from column data
    columnDataString.erase(0, columnDataString.find_first_not_of(L','));
    columnDataString.erase(columnDataString.find_last_not_of(L',') + 1);

    // Ensure that columnDataString and delimiter are not empty
    if (columnDataString.empty() || delimiterData.empty()) {
        showStatusMessage(LM.get(L"status_missing_column_or_delimiter_data"), MessageStatus::Error);
        return false;
    }

    // Use parseNumberRanges() to process column data
    std::vector<int> parsedColumns = parseNumberRanges(columnDataString, LM.get(L"status_invalid_range_in_column_data"));
    if (parsedColumns.empty()) return false; // Abort if parsing failed

    // Convert parsedColumns to set for uniqueness
    std::set<int> uniqueColumns(parsedColumns.begin(), parsedColumns.end());

    // Validate delimiter
    if (extendedDelimiter.empty()) {
        showStatusMessage(LM.get(L"status_extended_delimiter_empty"), MessageStatus::Error);
        return false;
    }

    // Validate quote character
    if (!quoteCharString.empty() && (quoteCharString.length() != 1 ||
        !(quoteCharString[0] == L'"' || quoteCharString[0] == L'\''))) {
        showStatusMessage(LM.get(L"status_invalid_quote_character"), MessageStatus::Error);
        return false;
    }

    // Check if columns have changed BEFORE updating stored data
    bool columnChanged = (columnDelimiterData.columns != uniqueColumns);

    // Update columnDelimiterData values
    columnDelimiterData.delimiterChanged = delimiterChanged;
    columnDelimiterData.quoteCharChanged = quoteCharChanged;
    columnDelimiterData.columnChanged = columnChanged;

    columnDelimiterData.inputColumns = std::move(parsedColumns);
    columnDelimiterData.columns = std::move(uniqueColumns);
    columnDelimiterData.extendedDelimiter = std::move(extendedDelimiter);
    columnDelimiterData.delimiterLength = columnDelimiterData.extendedDelimiter.length();
    columnDelimiterData.quoteChar = std::move(quoteCharConverted);

    return true;
}

bool MultiReplace::validateDelimiterData() {
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED) {
        return parseColumnAndDelimiterData();
    }

    return true;
}

void MultiReplace::findAllDelimitersInDocument() {

    lineDelimiterPositions.clear();

    textModified = false;
    logChanges.clear();
    isLoggingEnabled = true;

    LRESULT totalLines = send(SCI_GETLINECOUNT, 0, 0);

    lineDelimiterPositions.reserve(totalLines);

    // Find and store delimiter positions for each line
    for (LRESULT line = 0; line < totalLines; ++line) {
        findDelimitersInLine(line);
    }

    lineBuffer.shrink_to_fit();
    logChanges.clear();

}

void MultiReplace::findDelimitersInLine(LRESULT line) {
    // Initialize line information
    LineInfo lineInfo;

    // Pre-reserve based on first line's column count (reduces reallocations for uniform CSVs)
    if (!lineDelimiterPositions.empty() && !lineDelimiterPositions[0].positions.empty()) {
        lineInfo.positions.reserve(lineDelimiterPositions[0].positions.size());
    }

    // Get line length
    LRESULT lineLength = send(SCI_LINELENGTH, line, 0);
    lineInfo.lineLength = lineLength;

    // Ensure buffer is large enough
    if (lineBuffer.size() < static_cast<size_t>(lineLength + 1))
        lineBuffer.resize(lineLength + 1);

    // Read line into buffer and create a string view
    send(SCI_GETLINE, line, reinterpret_cast<sptr_t>(lineBuffer.data()));
    std::string_view lineContent(lineBuffer.data(), static_cast<size_t>(lineLength));

    // Cache delimiter and quote character
    size_t delimLen = columnDelimiterData.delimiterLength;
    std::string_view delimiter(columnDelimiterData.extendedDelimiter);
    bool hasQuoteChar = !columnDelimiterData.quoteChar.empty();
    char currentQuoteChar = hasQuoteChar ? columnDelimiterData.quoteChar[0] : '\0';
    char delimChar = delimiter[0];  // Cache first char for single-char delimiter path

    size_t pos = 0;
    bool inQuotes = false;

    while (pos < lineContent.size()) {
        // Toggle quote status if a quote character is found
        if (hasQuoteChar && lineContent[pos] == currentQuoteChar) {
            inQuotes = !inQuotes;
            ++pos;
            continue;
        }

        // Process delimiter if outside quotes
        if (!inQuotes) {
            if (delimLen == 1) {
                // Single-character delimiter check
                if (lineContent[pos] == delimChar) {
                    lineInfo.positions.push_back({ static_cast<LRESULT>(pos) });
                    ++pos;
                    continue;
                }
            }
            else {
                // Multi-character delimiter search
                size_t foundPos = lineContent.find(delimiter, pos);
                if (foundPos != std::string_view::npos) {
                    if (hasQuoteChar) {
                        size_t nextQuote = lineContent.find(currentQuoteChar, pos);
                        if (nextQuote != std::string_view::npos && nextQuote < foundPos) {
                            pos = nextQuote;
                            continue;
                        }
                    }
                    lineInfo.positions.push_back({ static_cast<LRESULT>(foundPos) });
                    pos = foundPos + delimLen;
                    continue;
                }
                else {
                    break; // No more delimiters found
                }
            }
        }
        ++pos;
    }

    // Store line info in vector
    if (line < static_cast<LRESULT>(lineDelimiterPositions.size()))
        lineDelimiterPositions[line] = std::move(lineInfo);
    else {
        lineDelimiterPositions.resize(line + 1);
        lineDelimiterPositions[line] = std::move(lineInfo);
    }
}

ColumnInfo MultiReplace::getColumnInfo(LRESULT startPosition) {
    if (columnDelimiterData.columns.empty() ||
        columnDelimiterData.extendedDelimiter.empty() ||
        lineDelimiterPositions.empty())
    {
        return { 0, 0, 0 };
    }

    // Determine how many lines are in Scintilla
    LRESULT totalLines = send(SCI_GETLINECOUNT, 0, 0);

    // Determine which line in Scintilla corresponds to startPosition
    LRESULT startLine = send(SCI_LINEFROMPOSITION, startPosition, 0);
    SIZE_T  startColumnIndex = 1;

    // Check if the line index is valid in lineDelimiterPositions
    LRESULT listSize = static_cast<LRESULT>(lineDelimiterPositions.size());
    if (startLine < totalLines && startLine < listSize) {
        const auto& lineInfo = lineDelimiterPositions[startLine];
        const auto& linePositions = lineInfo.positions;

        // Calculate absolute start of this line
        LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, startLine, 0);

        SIZE_T i = 0;
        for (; i < linePositions.size(); ++i) {
            // Absolute position of the i-th delimiter
            LRESULT delimAbsPos = lineStartPos + linePositions[i].offsetInLine;

            // If the caret is before or exactly on this delimiter, we've found our column
            if (startPosition <= delimAbsPos) {
                startColumnIndex = i + 1; // 1-based column
                break;
            }
        }

        // If we did not break, the caret is beyond the last delimiter => last column
        if (i == linePositions.size()) {
            startColumnIndex = linePositions.size() + 1;
        }
    }

    return { totalLines, startLine, startColumnIndex };
}

LRESULT MultiReplace::adjustForegroundForDarkMode(LRESULT textColor, LRESULT backgroundColor) {
    // Extract RGB components from text color
    int redText = (textColor & 0xFF);
    int greenText = ((textColor >> 8) & 0xFF);
    int blueText = ((textColor >> 16) & 0xFF);

    // Extract RGB components from background color
    int redBg = (backgroundColor & 0xFF);
    int greenBg = ((backgroundColor >> 8) & 0xFF);
    int blueBg = ((backgroundColor >> 16) & 0xFF);

    // Blend with background color to improve contrast
    float blendFactor = 0.8f;  // Higher value = stronger tint from background
    redText = static_cast<int>(redText * (1.0f - blendFactor) + redBg * blendFactor);
    greenText = static_cast<int>(greenText * (1.0f - blendFactor) + greenBg * blendFactor);
    blueText = static_cast<int>(blueText * (1.0f - blendFactor) + blueBg * blendFactor);

    // Increase brightness for better readability
    float brightnessBoost = 1.9f;  // Higher value = lighter text
    redText = std::min(255, static_cast<int>(redText * brightnessBoost));
    greenText = std::min(255, static_cast<int>(greenText * brightnessBoost));
    blueText = std::min(255, static_cast<int>(blueText * brightnessBoost));

    // Reconstruct adjusted color
    return (blueText << 16) | (greenText << 8) | redText;
}

void MultiReplace::initializeColumnStyles() {

    ::SendMessage(nppData._nppHandle, NPPM_MENUCOMMAND, 0, IDM_LANG_TEXT);

    // Get default text color from Notepad++
    LRESULT fgColor = send(SCI_STYLEGETFORE, STYLE_DEFAULT);

    // Check if Notepad++ is in dark mode
    const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    // Select color scheme based on mode
    const auto& columnColors = isDark ? darkModeColumnColors : lightModeColumnColors;

    for (SIZE_T column = 0; column < hColumnStyles.size(); ++column) {
        long bgColor = columnColors[column % columnColors.size()];
        LRESULT adjustedFgColor = isDark ? adjustForegroundForDarkMode(fgColor, bgColor) : fgColor;

        send(SCI_STYLESETBACK, hColumnStyles[column], bgColor);
        send(SCI_STYLESETFORE, hColumnStyles[column], adjustedFgColor);
    }
}

void MultiReplace::handleHighlightColumnsInDocument() {
    if (!validateDelimiterData()) {
        return;
    }

    // --- Viewport snapshot (vertical + horizontal + selection)
    const ViewState vs = saveViewState();

    int currentBufferID = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));
    highlightedTabs.mark(currentBufferID);
    initializeColumnStyles();
    LRESULT totalLines = static_cast<LRESULT>(lineDelimiterPositions.size());

    for (LRESULT line = 0; line < totalLines; ++line) {
        highlightColumnsInLine(line);
    }

    if (!lineDelimiterPositions.empty()) {
        LRESULT startPosition = send(SCI_GETCURRENTPOS, 0, 0);
        showStatusMessage(LM.get(L"status_actual_position", { addLineAndColumnMessage(startPosition) }), MessageStatus::Success);
    }

    isColumnHighlighted = true;
    isCaretPositionEnabled = true;

    // --- Viewport restore
    restoreViewStateExact(vs);
}

void MultiReplace::fixHighlightAtDocumentEnd() {
    if (!isColumnHighlighted) return;

    // Workaround: Highlight last lines to fix N++ bug causing loss of styling
    size_t lastLine = lineDelimiterPositions.size();
    LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);

    if (lastLine >= 2) {
        size_t highlightLine1 = lastLine - 2;
        if (highlightLine1 < (size_t)docLineCount) {
            highlightColumnsInLine(static_cast<LRESULT>(highlightLine1));
        }
        size_t highlightLine2 = lastLine - 1;
        if (highlightLine2 < (size_t)docLineCount) {
            highlightColumnsInLine(static_cast<LRESULT>(highlightLine2));
        }
    }
}

void MultiReplace::highlightColumnsInLine(LRESULT line) {
    // Retrieve the pre-parsed line information
    const auto& lineInfo = lineDelimiterPositions[line];

    // Skip empty lines
    if (lineInfo.lineLength == 0) {
        return;
    }

    // Cache frequently accessed values to avoid repeated member access
    const size_t lineLen = static_cast<size_t>(lineInfo.lineLength);
    const size_t styleCount = hColumnStyles.size();
    const SIZE_T delimLen = columnDelimiterData.delimiterLength;
    const size_t delimCount = lineInfo.positions.size();

    // Reuse style buffer - only grow if needed, then zero only the used portion
    if (styleBuffer.size() < lineLen) {
        styleBuffer.resize(lineLen);
    }
    std::fill_n(styleBuffer.data(), lineLen, char(0));

    // If no delimiters are found and column 1 is defined, fill the entire line with column 1's style
    if (delimCount == 0 && columnDelimiterData.columns.count(1) > 0)
    {
        char style = static_cast<char>(hColumnStyles[0] & 0xFF);
        std::fill_n(styleBuffer.data(), lineLen, style);
    }
    else {
        // For each defined column, calculate the start and end offsets within the line
        for (SIZE_T column : columnDelimiterData.columns) {
            if (column <= delimCount + 1) {
                size_t start = 0;
                size_t end = 0;

                if (column == 1) {
                    start = 0;
                }
                else {
                    start = static_cast<size_t>(lineInfo.positions[column - 2].offsetInLine) + delimLen;
                }

                if (column == delimCount + 1) {
                    end = lineLen;
                }
                else {
                    end = static_cast<size_t>(lineInfo.positions[column - 1].offsetInLine);
                }

                // Apply the style if the range is valid
                if (start < end && end <= lineLen) {
                    char style = static_cast<char>(hColumnStyles[(column - 1) % styleCount] & 0xFF);
                    std::fill_n(styleBuffer.data() + start, end - start, style);
                }
            }
        }
    }

    // Determine the absolute starting position of the line in the document
    LRESULT lineStartPos = send(SCI_POSITIONFROMLINE, line, 0);

    // Apply the computed styles to the document via Scintilla's API
    send(SCI_STARTSTYLING, lineStartPos, 0);
    send(SCI_SETSTYLINGEX, lineLen, reinterpret_cast<sptr_t>(styleBuffer.data()));
}

void MultiReplace::handleClearColumnMarks() {
    // Get the current buffer ID
    int currentBufferID = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));

    // If the tab was not highlighted, exit early
    if (!highlightedTabs.isHighlighted(currentBufferID)) {
        return;
    }

    // --- Save viewport
    const ViewState vs = saveViewState();

    // Get total document length
    LRESULT textLength = send(SCI_GETLENGTH);

    // Reset all styling to default
    send(SCI_STARTSTYLING, 0);
    send(SCI_SETSTYLING, textLength, STYLE_DEFAULT);

    isColumnHighlighted = false;
    isCaretPositionEnabled = false;

    // Force Scintilla to recalculate word wrapping if highlighting affected layout
    forceWrapRecalculation();

    // Remove tab from tracked highlighted tabs
    highlightedTabs.clear(currentBufferID);

    // --- Restore viewport
    restoreViewStateExact(vs);
}

std::wstring MultiReplace::addLineAndColumnMessage(LRESULT pos) {
    if (!columnDelimiterData.isValid()) {
        return L"";
    }
    ColumnInfo startInfo = getColumnInfo(pos);
    std::wstring lineAndColumnMessage = LM.get(L"status_line_and_column_position",
        { std::to_wstring(startInfo.startLine + 1),
          std::to_wstring(startInfo.startColumnIndex) });

    return lineAndColumnMessage;
}

void MultiReplace::processLogForDelimiters() {
    // Check if logChanges is accessible
    if (!textModified || logChanges.empty()) {
        return;
    }

    std::vector<LogEntry> modifyLogEntries;
    modifyLogEntries.reserve(logChanges.size());

    // Loop through the log entries in chronological order
    for (auto& logEntry : logChanges) {
        switch (logEntry.changeType) {
        case ChangeType::Insert:
        {
            Sci_Position insertPos = logEntry.lineNumber;
            Sci_Position blockCount = logEntry.blockSize;

            // Shift existing modifies if needed
            for (auto& m : modifyLogEntries) {
                if (m.lineNumber >= insertPos) {
                    m.lineNumber += blockCount;
                }
            }

            // Insert new lines
            updateDelimitersInDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Insert
            );

            // Immediately parse them so they're recognized
            updateDelimitersInDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Modify
            );

            updateUnsortedDocument(
                static_cast<SIZE_T>(insertPos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Insert
            );

            // Optionally highlight all new lines
            if (isColumnHighlighted) {
                LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);
                for (Sci_Position offset = 0; offset < blockCount; ++offset) {
                    Sci_Position lineToHighlight = insertPos + offset;
                    if (lineToHighlight >= 0
                        && static_cast<size_t>(lineToHighlight) < lineDelimiterPositions.size()
                        && lineToHighlight < docLineCount)
                    {
                        highlightColumnsInLine(lineToHighlight);
                    }
                }
            }
            break;
        }

        case ChangeType::Delete:
        {
            Sci_Position deletePos = logEntry.lineNumber;
            Sci_Position blockCount = logEntry.blockSize;

            // remove modifies in [deletePos..deletePos+blockCount)
            for (auto& m : modifyLogEntries) {
                if (m.lineNumber >= deletePos && m.lineNumber < (deletePos + blockCount)) {
                    m.lineNumber = -1; // Mark for removal
                }
                else if (m.lineNumber >= (deletePos + blockCount)) {
                    m.lineNumber -= blockCount;
                }
            }

            updateDelimitersInDocument(
                static_cast<SIZE_T>(deletePos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Delete
            );

            updateUnsortedDocument(
                static_cast<SIZE_T>(deletePos),
                static_cast<SIZE_T>(blockCount),
                ChangeType::Delete
            );

            // Re-parse ONLY the merged line at deletePos
            if (deletePos < (Sci_Position)lineDelimiterPositions.size()) {
                findDelimitersInLine(deletePos);

                if (isColumnHighlighted) {
                    LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);
                    if (deletePos >= 0
                        && deletePos < docLineCount
                        && static_cast<size_t>(deletePos) < lineDelimiterPositions.size())
                    {
                        highlightColumnsInLine(deletePos);
                    }
                }
            }
            break;
        }

        case ChangeType::Modify:
        {
            modifyLogEntries.push_back(logEntry);
            break;
        }

        default:
            break;
        }
    }

    // Apply the saved "Modify" entries
    for (const auto& m : modifyLogEntries)
    {
        if (m.lineNumber == -1) {
            continue; // skip removed lines
        }

        if (static_cast<size_t>(m.lineNumber) < lineDelimiterPositions.size()) {
            updateDelimitersInDocument(
                static_cast<size_t>(m.lineNumber),
                1,
                ChangeType::Modify
            );

            if (isColumnHighlighted) {
                LRESULT docLineCount = send(SCI_GETLINECOUNT, 0, 0);
                if (m.lineNumber >= 0
                    && m.lineNumber < docLineCount
                    && static_cast<size_t>(m.lineNumber) < lineDelimiterPositions.size())
                {
                    highlightColumnsInLine(m.lineNumber);
                }
            }
        }
    }

    // Workaround: Highlight last lines to fix N++ bug causing loss of styling 
    fixHighlightAtDocumentEnd();

    // Clear Log queue
    logChanges.clear();
    textModified = false;
}

void MultiReplace::updateDelimitersInDocument(SIZE_T lineNumber, SIZE_T blockCount, ChangeType changeType) {
    // Return early if the line number is invalid
    if (lineNumber > lineDelimiterPositions.size()) {
        return;
    }

    switch (changeType) {
    case ChangeType::Insert:
    {
        // Insert multiple empty lines instead of just one
        std::vector<LineInfo> newLines;
        newLines.reserve(blockCount);

        for (SIZE_T i = 0; i < blockCount; ++i) {
            LineInfo newLine;
            newLine.lineLength = 0; // New line has no content initially
            newLine.positions.clear();
            newLines.push_back(newLine);
        }

        // Add them all at the specified position
        lineDelimiterPositions.insert(
            lineDelimiterPositions.begin() + lineNumber,
            newLines.begin(),
            newLines.end()
        );

    }
    break;

    case ChangeType::Delete:
    {
        // Remove the specified number of lines if they exist
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > lineDelimiterPositions.size()) {
            endPos = lineDelimiterPositions.size();
        }

        if (lineNumber < lineDelimiterPositions.size()) {
            lineDelimiterPositions.erase(
                lineDelimiterPositions.begin() + lineNumber,
                lineDelimiterPositions.begin() + endPos
            );

        }
    }
    break;

    case ChangeType::Modify:
    {
        // Reanalyze multiple lines if they exist
        SIZE_T endPos = lineNumber + blockCount;
        if (endPos > lineDelimiterPositions.size()) {
            endPos = lineDelimiterPositions.size();
        }

        for (SIZE_T i = lineNumber; i < endPos; ++i) {
            findDelimitersInLine(i);
        }
    }
    break;

    default:
        break;
    }
}

void MultiReplace::handleDelimiterPositions(DelimiterOperation operation) {
    // Check if IDC_COLUMN_MODE_RADIO checkbox is not checked, and if so, return early
    if (IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) != BST_CHECKED) {
        return;
    }

    // If there's been a change in the active window within Notepad++, reset all delimiter settings
    if (documentSwitched) {
        handleClearDelimiterState();
        documentSwitched = false;
    }

    if (operation == DelimiterOperation::LoadAll) {
        // Parse column and delimiter data; exit if parsing fails or if delimiter is empty
        if (!parseColumnAndDelimiterData()) {
            return;
        }

        // Check current highlight status (global or per-tab)
        const int currentBufferID = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));
        const bool highlightActive = (isColumnHighlighted != FALSE) ||
            highlightedTabs.isHighlighted(currentBufferID);

        // Trigger scan only if necessary
        if (columnDelimiterData.isValid() &&
            (columnDelimiterData.delimiterChanged ||
                columnDelimiterData.quoteCharChanged ||
                lineDelimiterPositions.empty()))
        {
            findAllDelimitersInDocument();

            // Re-apply highlight when structure changed and highlight is active
            if (highlightActive) {
                handleHighlightColumnsInDocument();
            }
        }

        // Re-apply highlight when only the selected column changed (no rescan required)
        if (columnDelimiterData.isValid() &&
            columnDelimiterData.columnChanged &&
            highlightActive)
        {
            handleHighlightColumnsInDocument();
        }
    }
    else if (operation == DelimiterOperation::Update) {
        // Ensure the columns and delimiter data is present before processing updates
        if (columnDelimiterData.isValid()) {
            processLogForDelimiters();
        }
    }
}

void MultiReplace::handleClearDelimiterState() {
    pointerToScintilla();
    if (!_hScintilla) return;
    lineDelimiterPositions.clear();
    isLoggingEnabled = false;
    textModified = false;
    logChanges.clear();
    if (isColumnHighlighted) {
        handleClearColumnMarks();
    }
    clearFlowTabsIfAny();
    isCaretPositionEnabled = false;
}

// For testing purposes - uncomment to debug log changes
void MultiReplace::displayLogChangesInMessageBox() {
    // Helper function to convert std::string to std::wstring
    auto utf8ToWString = [](const std::string& input) -> std::wstring {
        if (input.empty()) return std::wstring();
        int size = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, 0, 0);
        std::wstring result(size, 0);
        MultiByteToWideChar(CP_UTF8, 0, input.c_str(), -1, &result[0], size);
        return result;
        };

    std::ostringstream oss;
    oss << "logChanges.size() = " << logChanges.size() << "\n\n";

    int idx = 0;
    for (const auto& entry : logChanges) {
        oss << "[" << idx++ << "] ";
        switch (entry.changeType) {
        case ChangeType::Insert:
            oss << "INSERT line=" << entry.lineNumber << " blockSize=" << entry.blockSize;
            break;
        case ChangeType::Modify:
            oss << "MODIFY line=" << entry.lineNumber;
            break;
        case ChangeType::Delete:
            oss << "DELETE line=" << entry.lineNumber << " blockSize=" << entry.blockSize;
            break;
        }
        oss << "\n";
    }

    std::wstring msg = utf8ToWString(oss.str());
    MessageBox(nullptr, msg.c_str(), L"Log Changes Debug", MB_OK);
}

#pragma endregion


#pragma region Utilities

static inline bool decodeNumericEscape(const std::wstring& src, size_t pos, int  base, int digits, wchar_t& out)
{
    if (pos + digits > src.size())           // not enough characters
        return false;

    unsigned value = 0;
    for (int k = 0; k < digits; ++k)
    {
        wchar_t ch = src[pos + k];
        unsigned v;

        if (ch >= L'0' && ch <= L'9') v = ch - L'0';
        else if (ch >= L'A' && ch <= L'F') v = (ch - L'A') + 10;
        else if (ch >= L'a' && ch <= L'f') v = (ch - L'a') + 10;
        else return false;                  // invalid digit

        if (v >= static_cast<unsigned>(base))
            return false;                   // digit not allowed in this base

        value = value * base + v;
    }

    if (value > 0xFFFF)                     // outside BMP range
        return false;

    out = static_cast<wchar_t>(value);
    return true;
}

std::string MultiReplace::convertAndExtendW(const std::wstring& input, bool extended, UINT targetCodepage) const
{
    if (!extended)                          // fast path – no escapes
        return Encoding::wstringToBytes(input, targetCodepage);

    std::wstring out;
    out.reserve(input.size());

    for (size_t i = 0; i < input.size(); ++i)
    {
        wchar_t ch = input[i];

        if (ch != L'\\' || i + 1 >= input.size())
        {
            out.push_back(ch);
            continue;
        }

        wchar_t esc = input[++i];           // escape designator
        wchar_t decoded;

        switch (esc)
        {
        case L'r': out.push_back(L'\r');                    break;
        case L'n': out.push_back(L'\n');                    break;
        case L't': out.push_back(L'\t');                    break;
        case L'\\':out.push_back(L'\\');                    break;
        case L'0': out.push_back(L'\0');                    break;

        case L'o': // \oNNN  (octal)
            if (decodeNumericEscape(input, i + 1, 8, 3, decoded))
            {
                out.push_back(decoded); i += 3; break;
            }
            [[fallthrough]];                                // literal fallback

        case L'b': // \bNNNNNNNN  (binary, 8 digits)
            if (decodeNumericEscape(input, i + 1, 2, 8, decoded))
            {
                out.push_back(decoded); i += 8; break;
            }
            [[fallthrough]];

        case L'd': // \dNNN  (decimal)
            if (decodeNumericEscape(input, i + 1, 10, 3, decoded))
            {
                out.push_back(decoded); i += 3; break;
            }
            [[fallthrough]];

        case L'x': // \xHH   (hex-byte)
            if (decodeNumericEscape(input, i + 1, 16, 2, decoded))
            {
                out.push_back(decoded); i += 2; break;
            }
            [[fallthrough]];

        case L'u': // \uXXXX (Unicode BMP)
            if (decodeNumericEscape(input, i + 1, 16, 4, decoded))
            {
                out.push_back(decoded); i += 4; break;
            }
            [[fallthrough]];

        default:  // unknown or invalid sequence → keep literally
            out.append({ L'\\', esc });
            break;
        }
    }

    return Encoding::wstringToBytes(out, targetCodepage);
}

std::string MultiReplace::convertAndExtendW(const std::wstring& input, bool extended)
{
    UINT cp = getCurrentDocCodePage();
    return convertAndExtendW(input, extended, cp);
}

void MultiReplace::addStringToComboBoxHistory(HWND hComboBox, const std::wstring& str, int maxItems)
{
    if (str.length() == 0)
    {
        // Skip adding empty strings to the combo box history
        return;
    }

    // Check if the string is already in the combo box
    int index = static_cast<int>(SendMessage(hComboBox, CB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(str.c_str())));

    // If the string is not found, insert it at the beginning
    if (index == CB_ERR)
    {
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str.c_str()));

        // Remove the last item if the list exceeds maxItems
        if (SendMessage(hComboBox, CB_GETCOUNT, 0, 0) > maxItems)
        {
            SendMessage(hComboBox, CB_DELETESTRING, maxItems, 0);
        }
    }
    else
    {
        // If the string is found, move it to the beginning
        SendMessage(hComboBox, CB_DELETESTRING, index, 0);
        SendMessage(hComboBox, CB_INSERTSTRING, 0, reinterpret_cast<LPARAM>(str.c_str()));
    }

    // Select the newly added string
    SendMessage(hComboBox, CB_SETCURSEL, 0, 0);
}

std::wstring MultiReplace::getTextFromDialogItem(HWND hwnd, int itemID)
{
    HWND hCtrl = GetDlgItem(hwnd, itemID);
    if (!hCtrl) return L"";

    int len = GetWindowTextLengthW(hCtrl);
    if (len <= 0) return L"";

    std::wstring wbuf(static_cast<size_t>(len) + 1, L'\0');
    int written = GetDlgItemTextW(hwnd, itemID, &wbuf[0], len + 1);
    if (written < 0) written = 0;
    if (!wbuf.empty() && wbuf.back() == L'\0') wbuf.pop_back();
    wbuf.resize(static_cast<size_t>(written));
    return wbuf;
}

void MultiReplace::setSelections(bool select, bool onlySelected) {
    if (replaceListData.empty()) {
        return;
    }

    // Store the original data for all affected items
    std::vector<std::pair<size_t, ReplaceItemData>> originalDataList;
    for (size_t i = 0; i < replaceListData.size(); ++i) {
        if (!onlySelected || ListView_GetItemState(_replaceListView, static_cast<int>(i), LVIS_SELECTED)) {
            originalDataList.emplace_back(i, replaceListData[i]);
            replaceListData[i].isEnabled = select;
        }
    }

    // Update the allSelected flag if all items were selected/deselected
    if (!onlySelected) {
        allSelected = select;
    }

    // Update the header after changing the selection status of the items
    updateHeaderSelection();

    // Update the ListView
    for (const auto& [index, _] : originalDataList) {
        updateListViewItem(index);
    }

    // Create Undo/Redo actions
    UndoRedoAction action;

    // Undo action: Restore the original data
    action.undoAction = [this, originalDataList]() {
        for (const auto& [index, data] : originalDataList) {
            replaceListData[index] = data;
            updateListViewItem(index);
        }
        updateHeaderSelection();
        };

    // Redo action: Apply the new selection status
    action.redoAction = [this, originalDataList]() {
        for (const auto& [index, data] : originalDataList) {
            replaceListData[index].isEnabled = !data.isEnabled;
            updateListViewItem(index);
        }
        updateHeaderSelection();
        };

    // Push the action onto the undoStack
    URM.push(action.undoAction, action.redoAction, L"Set enabled");

    // Show Select Statisics
    showListFilePath();
}

void MultiReplace::setOptionForSelection(SearchOption option, bool value) {
    if (replaceListData.empty()) return;

    std::vector<std::pair<size_t, ReplaceItemData>> originalDataList;
    for (size_t i = 0; i < replaceListData.size(); ++i) {
        if (ListView_GetItemState(_replaceListView, static_cast<int>(i), LVIS_SELECTED)) {
            originalDataList.emplace_back(i, replaceListData[i]);
            switch (option) {
            case SearchOption::WholeWord:  replaceListData[i].wholeWord = value; break;
            case SearchOption::MatchCase:  replaceListData[i].matchCase = value; break;
            case SearchOption::Variables:  replaceListData[i].useVariables = value; break;
            case SearchOption::Extended:
                replaceListData[i].extended = value;
                if (value) replaceListData[i].regex = false;
                break;
            case SearchOption::Regex:
                replaceListData[i].regex = value;
                if (value) replaceListData[i].extended = false;
                break;
            }
        }
    }
    if (originalDataList.empty()) return;

    for (const auto& [index, _] : originalDataList) updateListViewItem(index);

    std::wstring optionName;
    switch (option) {
    case SearchOption::WholeWord: optionName = L"Whole Word"; break;
    case SearchOption::MatchCase: optionName = L"Match Case"; break;
    case SearchOption::Variables: optionName = L"Variables"; break;
    case SearchOption::Extended:  optionName = L"Extended"; break;
    case SearchOption::Regex:     optionName = L"Regex"; break;
    }

    UndoRedoAction action;
    action.undoAction = [this, originalDataList]() {
        for (const auto& [index, data] : originalDataList) {
            replaceListData[index] = data;
            updateListViewItem(index);
        }
        };
    action.redoAction = [this, originalDataList, option, value]() {
        for (const auto& [index, _] : originalDataList) {
            switch (option) {
            case SearchOption::WholeWord:  replaceListData[index].wholeWord = value; break;
            case SearchOption::MatchCase:  replaceListData[index].matchCase = value; break;
            case SearchOption::Variables:  replaceListData[index].useVariables = value; break;
            case SearchOption::Extended:
                replaceListData[index].extended = value;
                if (value) replaceListData[index].regex = false;
                break;
            case SearchOption::Regex:
                replaceListData[index].regex = value;
                if (value) replaceListData[index].extended = false;
                break;
            }
            updateListViewItem(index);
        }
        };
    URM.push(action.undoAction, action.redoAction, (value ? L"Set " : L"Clear ") + optionName);
}

void MultiReplace::showStatusMessage(const std::wstring& messageText, MessageStatus status, bool isNotFound, bool isTransient)
{
    const size_t MAX_DISPLAY_LENGTH = 150;

    // Any non-transient status message disables "Actual Position" tracking
    // User must toggle Highlight off/on to re-enable position display
    if (!isTransient && isCaretPositionEnabled) {
        isCaretPositionEnabled = false;
    }

    std::wstring strMessage;
    strMessage.reserve(messageText.length());
    for (wchar_t ch : messageText) {
        if (iswprint(ch)) {
            strMessage += ch;
        }
    }

    if (strMessage.size() > MAX_DISPLAY_LENGTH) {
        strMessage = strMessage.substr(0, MAX_DISPLAY_LENGTH - 3) + L"...";
    }

    // Store the TYPE of the message for later theme switches
    _lastMessageStatus = status;

    // Set the active color based on the new status
    switch (status) {
    case MessageStatus::Success:
        _statusMessageColor = COLOR_SUCCESS;
        break;
    case MessageStatus::Error:
        _statusMessageColor = COLOR_ERROR;
        break;
    case MessageStatus::Info:
    default:
        _statusMessageColor = COLOR_INFO;
        break;
    }

    HWND hStatusMessage = GetDlgItem(_hSelf, IDC_STATUS_MESSAGE);
    SetWindowText(hStatusMessage, strMessage.c_str());

    // --- KEY CHANGE: Invalidate the PARENT window's area BEHIND the control ---
    // This forces the N++ themed dialog to erase the background for us with the correct color.
    RECT rect;
    GetWindowRect(hStatusMessage, &rect);
    MapWindowPoints(HWND_DESKTOP, GetParent(hStatusMessage), (LPPOINT)&rect, 2);
    InvalidateRect(GetParent(hStatusMessage), &rect, TRUE); // TRUE means "erase background"
    UpdateWindow(GetParent(hStatusMessage));

    if (isNotFound)
    {
        FLASHWINFO fwInfo = {};
        fwInfo.cbSize = sizeof(FLASHWINFO);
        fwInfo.hwnd = _hSelf;
        fwInfo.dwFlags = FLASHW_ALL;
        fwInfo.uCount = 2;
        fwInfo.dwTimeout = 100;
        FlashWindowEx(&fwInfo);
        if (!muteSounds) {
            ::MessageBeep(MB_ICONASTERISK);
        }
    }
}

void MultiReplace::applyThemePalette()
{
    // Check if Notepad++ is currently in dark mode
    const bool isDark = NppStyleKit::ThemeUtils::isDarkMode(nppData._nppHandle);

    // Assign colours from the predefined palettes in the header file
    if (isDark) {
        COLOR_SUCCESS = DMODE_SUCCESS;
        COLOR_ERROR = DMODE_ERROR;
        COLOR_INFO = DMODE_INFO;
        _filterHelpColor = DMODE_FILTER_HELP;
    }
    else {
        COLOR_SUCCESS = LMODE_SUCCESS;
        COLOR_ERROR = LMODE_ERROR;
        COLOR_INFO = LMODE_INFO;
        _filterHelpColor = LMODE_FILTER_HELP;
    }

    // Update the active colour based on the last message status
    switch (_lastMessageStatus) {
    case MessageStatus::Success: _statusMessageColor = COLOR_SUCCESS; break;
    case MessageStatus::Error:   _statusMessageColor = COLOR_ERROR;   break;
    default:                     _statusMessageColor = COLOR_INFO;    break;
    }

    // Repaint status control and "(?)" tooltip so they pick up the new palette
    InvalidateRect(GetDlgItem(_hSelf, IDC_STATUS_MESSAGE), nullptr, TRUE);
    InvalidateRect(GetDlgItem(_hSelf, IDC_FILTER_HELP), nullptr, TRUE);
}

void MultiReplace::refreshColumnStylesIfNeeded()
{
    if (isColumnHighlighted) {          // flag is managed by highlight/clear handlers
        initializeColumnStyles();
    }
}

std::wstring MultiReplace::getShortenedFilePath(const std::wstring& path, int maxLength, HDC hDC) {
    bool hdcProvided = true;

    // If no HDC is provided, get the one for the main window (_hSelf)
    if (hDC == nullptr) {
        hDC = GetDC(_hSelf);  // Get the device context for _hSelf (less accurate)
        hdcProvided = false;  // Mark that HDC was not provided externally
    }

    double dotWidth = 0.0;
    SIZE charSize;
    std::vector<double> characterWidths;

    // Calculate the width of each character in the path and the width of the dots ("...")
    for (wchar_t ch : path) {
        GetTextExtentPoint32(hDC, &ch, 1, &charSize);
        characterWidths.push_back(static_cast<double>(charSize.cx));
        if (ch == L'.') {
            dotWidth = static_cast<double>(charSize.cx);  // Store width of '.' separately
        }
    }

    double totalDotsWidth = dotWidth * 3;  // Width for "..."

    // Split the directory and filename
    size_t lastSlashPos = path.find_last_of(L"\\/");
    std::wstring directoryPart = (lastSlashPos != std::wstring::npos) ? path.substr(0, lastSlashPos + 1) : L"";
    std::wstring fileName = (lastSlashPos != std::wstring::npos) ? path.substr(lastSlashPos + 1) : path;

    // Calculate widths for directory and file name
    double directoryWidth = 0.0, fileNameWidth = 0.0;

    for (size_t i = 0; i < directoryPart.size(); ++i) {
        directoryWidth += characterWidths[i];
    }

    for (size_t i = directoryPart.size(); i < path.size(); ++i) {
        fileNameWidth += characterWidths[i];
    }

    std::wstring displayPath;
    double currentWidth = 0.0;

    // Shorten the file name if necessary
    if (fileNameWidth + totalDotsWidth > maxLength) {
        for (size_t i = directoryPart.size(); i < path.size(); ++i) {
            if (currentWidth + characterWidths[i] + totalDotsWidth > maxLength) {
                break;
            }
            displayPath += path[i];
            currentWidth += characterWidths[i];
        }
        displayPath += L"...";
    }
    // Shorten the directory part if necessary
    else if (directoryWidth + fileNameWidth > maxLength) {
        for (size_t i = 0; i < directoryPart.size(); ++i) {
            if (currentWidth + characterWidths[i] + totalDotsWidth + fileNameWidth > maxLength) {
                break;
            }
            displayPath += directoryPart[i];
            currentWidth += characterWidths[i];
        }
        displayPath += L"...";
        displayPath += fileName;
    }
    else {
        displayPath = path; // No shortening needed
    }

    // If we obtained HDC ourselves, release it
    if (!hdcProvided) {
        ReleaseDC(_hSelf, hDC);
    }

    return displayPath;
}

std::wstring MultiReplace::getSelectedText()
{
    Sci_Position lengthWithNul = static_cast<Sci_Position>(SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0, 0));
    if (lengthWithNul <= 0) return L"";

    std::string buffer(static_cast<size_t>(lengthWithNul), '\0');
    SendMessage(nppData._scintillaMainHandle, SCI_GETSELTEXT, 0,
        reinterpret_cast<LPARAM>(buffer.data()));

    if (!buffer.empty() && buffer.back() == '\0') buffer.pop_back();

    UINT sciCp = getCurrentDocCodePage();
    return Encoding::bytesToWString(buffer, sciCp);
}

LRESULT MultiReplace::getEOLLengthForLine(LRESULT line) {
    LRESULT lineLen = send(SCI_LINELENGTH, line, 0);
    if (lineLen == 0) {
        return 0; // Empty line, no EOL
    }

    // We'll check up to the last 2 chars only:
    LRESULT lineStart = send(SCI_POSITIONFROMLINE, line, 0);
    LRESULT checkCount = (lineLen >= 2) ? 2 : 1;

    // Read those chars via SCI_GETCHARAT
    char last[2] = { 0, 0 };
    for (LRESULT i = 0; i < checkCount; ++i) {
        last[i] = static_cast<char>(
            send(SCI_GETCHARAT, lineStart + lineLen - checkCount + i, 0)
            );
    }

    // If we have 2 chars and they are '\r\n', that's EOL=2
    if (checkCount == 2 && last[0] == '\r' && last[1] == '\n') {
        return 2;
    }
    // Otherwise check if the last char is '\r' or '\n'
    else if (last[checkCount - 1] == '\r' || last[checkCount - 1] == '\n') {
        return 1;
    }
    return 0;
}

std::string MultiReplace::getEOLStyle() {
    LRESULT eolMode = static_cast<LRESULT>(send(SCI_GETEOLMODE));
    switch (eolMode) {
    case SC_EOL_CRLF:
        return "\r\n";
    case SC_EOL_CR:
        return "\r";
    case SC_EOL_LF:
        return "\n";
    default:
        return "\n";  // Defaulting to LF
    }
}

sptr_t MultiReplace::send(unsigned int iMessage, uptr_t wParam, sptr_t lParam, bool useDirect) const {
    if (useDirect && pSciMsg) {
        return pSciMsg(pSciWndData, iMessage, wParam, lParam);
    }
    else {
        return ::SendMessage(_hScintilla, iMessage, wParam, lParam);
    }
}

bool MultiReplace::normalizeAndValidateNumber(std::string& str) {
    // Trim spaces/tabs before classifying
    auto is_space = [](unsigned char c) noexcept { return c == ' ' || c == '\t'; };
    std::size_t n = str.size(), t0 = 0, t1 = n;
    while (t0 < n && is_space((unsigned char)str[t0])) ++t0;
    while (t1 > t0 && is_space((unsigned char)str[t1 - 1])) --t1;
    if (t1 <= t0) return false;

    std::string_view trimmed(str.data() + t0, t1 - t0);

    mr::num::NumericToken tok;
    if (!mr::num::classify_numeric_field(trimmed, tok))
        return false;

    // Keep only the normalized numeric token ('.' normalized etc.)
    str = tok.normalized;
    return true;
}

std::vector<WCHAR> MultiReplace::createFilterString(const std::vector<std::pair<std::wstring, std::wstring>>& filters) {
    // Calculate the required size for the filter string
    size_t totalSize = 0;
    for (const auto& filter : filters) {
        totalSize += filter.first.size() + 1; // Description + null terminator
        totalSize += filter.second.size() + 1; // Pattern + null terminator
    }
    totalSize += 1; // Double null terminator at the end

    // Create the array
    std::vector<WCHAR> filterString;
    filterString.reserve(totalSize);

    // Fill the array
    for (const auto& filter : filters) {
        filterString.insert(filterString.end(), filter.first.begin(), filter.first.end());
        filterString.push_back(L'\0');
        filterString.insert(filterString.end(), filter.second.begin(), filter.second.end());
        filterString.push_back(L'\0');
    }

    // End with an additional null terminator
    filterString.push_back(L'\0');

    return filterString;
}

int MultiReplace::getFontHeight(HWND hwnd, HFONT hFont) {
    // Get the font size from the specified font and window
    TEXTMETRIC tm;
    HDC hdc = GetDC(hwnd);  // Get the device context for the specified window
    SelectObject(hdc, hFont);  // Select the specified font into the DC
    GetTextMetrics(hdc, &tm);  // Retrieve the text metrics for the font
    int fontHeight = tm.tmHeight;  // Extract the font height
    ReleaseDC(hwnd, hdc);  // Release the device context
    return fontHeight;  // Return the font height
}

int MultiReplace::getCharacterWidth(int elementID, const wchar_t* character) {
    // Get the HWND of the element by its ID
    HWND hwndElement = GetDlgItem(_hSelf, elementID);

    // Get the font used by the element
    HFONT hFont = reinterpret_cast<HFONT>(SendMessage(hwndElement, WM_GETFONT, 0, 0));

    // Get the device context for measuring text
    HDC hdc = GetDC(hwndElement);
    SelectObject(hdc, hFont);  // Use the font of the given element

    // Measure the width of the character
    SIZE size;
    GetTextExtentPoint32W(hdc, character, 1, &size);

    // Release the device context
    ReleaseDC(hwndElement, hdc);

    // Return the width of the character
    return size.cx;
}

std::vector<int> MultiReplace::parseNumberRanges(const std::wstring& input, const std::wstring& errorMessage)
{
    std::vector<int> result;
    if (input.empty()) return result;  // nothing to parse

    // use a hash set to filter out duplicates, but preserve insertion order in 'result'
    std::unordered_set<int> seen;

    std::wistringstream stream(input);
    std::wstring token;

    // helper to add a number only once, in the order encountered
    auto pushUnique = [&](int n) {
        if (seen.insert(n).second)      // if n was not already present
            result.push_back(n);
        };

    // process each comma-separated token, either a single number or a range "start-end"
    auto processToken = [&](const std::wstring& token) -> bool
        {
            if (token.empty()) return true;  // skip empty entries

            try {
                size_t dashPos = token.find(L'-');
                if (dashPos != std::wstring::npos) {
                    // RANGE: parse start and end values
                    int startRange = std::stoi(token.substr(0, dashPos));
                    int endRange = std::stoi(token.substr(dashPos + 1));
                    if (startRange < 1 || endRange < 1)
                        return false;

                    // push each number in the range, preserving order
                    if (endRange >= startRange) {
                        // ascending range
                        for (int i = startRange; i <= endRange; ++i)
                            pushUnique(i);
                    }
                    else {
                        // descending range
                        for (int i = startRange; i >= endRange; --i)
                            pushUnique(i);
                    }
                }
                else {
                    // SINGLE NUMBER
                    int number = std::stoi(token);
                    if (number < 1)
                        return false;  // invalid value

                    // add the single number
                    pushUnique(number);
                }
            }
            catch (const std::exception&) {
                return false;  // parsing failure
            }
            return true;
        };

    // split the input string on commas and feed each token to our processor
    while (std::getline(stream, token, L',')) {
        if (!processToken(token)) {
            showStatusMessage(errorMessage, MessageStatus::Error);
            return {};  // abort on invalid syntax
        }
    }

    // 'result' contains the right values in the correct order
    return result;
}

UINT MultiReplace::getCurrentDocCodePage()
{
    // Ask Scintilla for the current code page (always non‑negative)
    LRESULT cp = send(SCI_GETCODEPAGE, 0, 0);

    // If Scintilla ever returned 0 (undefined), fall back to ANSI
    return static_cast<UINT>(cp != 0 ? cp : CP_ACP);
}

Sci_Position MultiReplace::advanceAfterMatch(const SearchResult& r) {
    if (r.length > 0) return r.pos + r.length;
    const Sci_Position after = static_cast<Sci_Position>(send(SCI_POSITIONAFTER, r.pos, 0));
    const Sci_Position next = (after > r.pos) ? after : (r.pos + 1);
    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    return (next > docLen) ? docLen : next;
}

Sci_Position MultiReplace::ensureForwardProgress(Sci_Position candidate, const SearchResult& last) {
    if (candidate > last.pos) return candidate;
    const Sci_Position after = static_cast<Sci_Position>(send(SCI_POSITIONAFTER, last.pos, 0));
    const Sci_Position next = (after > last.pos) ? after : (last.pos + 1);
    const Sci_Position docLen = static_cast<Sci_Position>(send(SCI_GETLENGTH, 0, 0));
    return (next > docLen) ? docLen : next;
}

std::size_t MultiReplace::computeListHash(const std::vector<ReplaceItemData>& list) {
    std::size_t combinedHash = 0;
    ReplaceItemDataHasher hasher;

    for (const auto& item : list) {
        combinedHash ^= hasher(item) + golden_ratio_constant + (combinedHash << 6) + (combinedHash >> 2);
    }

    return combinedHash;
}

void MultiReplace::setTextInDialogItem(HWND hDlg, int itemID, const std::wstring& text) {
    ::SetDlgItemTextW(hDlg, itemID, text.c_str());
}

void MultiReplace::forceWrapRecalculation() {
    int originalWrapMode = static_cast<int>(send(SCI_GETWRAPMODE));
    if (originalWrapMode != SC_WRAP_NONE) {
        send(SCI_SETWRAPMODE, SC_WRAP_NONE, 0);
        send(SCI_SETWRAPMODE, originalWrapMode, 0);
    }
}

#pragma endregion


#pragma region FileOperations

std::wstring MultiReplace::openFileDialog(bool saveFile, const std::vector<std::pair<std::wstring, std::wstring>>& filters, const WCHAR* title, DWORD flags, const std::wstring& fileExtension, const std::wstring& defaultFilePath) {
    flags |= OFN_NOCHANGEDIR;
    OPENFILENAME ofn = {};
    WCHAR szFile[MAX_PATH] = {};

    // Safely copy the default file path into the buffer and ensure null-termination
    if (!defaultFilePath.empty()) {
        wcsncpy_s(szFile, defaultFilePath.c_str(), MAX_PATH);
    }

    std::vector<WCHAR> filter = createFilterString(filters);

    ofn.lStructSize = sizeof(OPENFILENAME);
    ofn.hwndOwner = _hSelf;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile) / sizeof(WCHAR);
    ofn.lpstrFilter = filter.data();
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = title;
    ofn.Flags = flags;

    if (saveFile ? GetSaveFileName(&ofn) : GetOpenFileName(&ofn)) {
        std::wstring filePath(szFile);

        // If no extension is provided, append the default one
        if (filePath.find_last_of(L".") == std::wstring::npos) {
            filePath += L"." + fileExtension;
        }

        return filePath;
    }
    else {
        return std::wstring();
    }
}

std::wstring MultiReplace::promptSaveListToCsv() {
    std::wstring csvDescription = LM.get(L"filetype_csv");  // "CSV Files (*.csv)"
    std::wstring allFilesDescription = LM.get(L"filetype_all_files");  // "All Files (*.*)"

    std::vector<std::pair<std::wstring, std::wstring>> filters = {
        {csvDescription, L"*.csv"},
        {allFilesDescription, L"*.*"}
    };

    std::wstring dialogTitle = LM.get(L"panel_save_list");
    std::wstring defaultFileName;

    if (!listFilePath.empty()) {
        // If a file path already exists, use its directory and filename
        defaultFileName = listFilePath;
    }
    else {
        // If no file path is set, provide a default file name with a sequential number
        static int fileCounter = 1;
        defaultFileName = L"Replace_List_" + std::to_wstring(++fileCounter) + L".csv";
    }

    // Call openFileDialog with the default file path and name
    std::wstring filePath = openFileDialog(
        true,
        filters,
        dialogTitle.c_str(),
        OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT,
        L"csv",
        defaultFileName
    );

    return filePath;
}

bool MultiReplace::saveListToCsvSilent(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    std::ofstream outFile(std::filesystem::path(filePath), std::ios::binary);

    if (!outFile.is_open()) {
        return false;
    }

    // [BOM] Write UTF-8 BOM for CSV
    outFile.write("\xEF\xBB\xBF", 3);

    // Convert and Write CSV header
    std::string utf8Header = Encoding::wstringToUtf8(L"Selected,Find,Replace,WholeWord,MatchCase,UseVariables,Extended,Regex,Comments\n");
    outFile << utf8Header;

    // Write list items to CSV file
    for (const ReplaceItemData& item : list) {
        std::wstring line = std::to_wstring(item.isEnabled) + L"," +
            StringUtils::escapeCsvValue(item.findText) + L"," +
            StringUtils::escapeCsvValue(item.replaceText) + L"," +
            std::to_wstring(item.wholeWord) + L"," +
            std::to_wstring(item.matchCase) + L"," +
            std::to_wstring(item.useVariables) + L"," +
            std::to_wstring(item.extended) + L"," +
            std::to_wstring(item.regex) + L"," +
            StringUtils::escapeCsvValue(item.comments) + L"\n";
        std::string utf8Line = Encoding::wstringToUtf8(line);
        outFile << utf8Line;
    }

    outFile.close();

    return !outFile.fail();
}

void MultiReplace::saveListToCsv(const std::wstring& filePath, const std::vector<ReplaceItemData>& list) {
    if (!saveListToCsvSilent(filePath, list)) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    showStatusMessage(LM.get(L"status_saved_items_to_csv", { std::to_wstring(list.size()) }), MessageStatus::Success);

    // Update the file path and original hash after a successful save
    listFilePath = filePath;
    originalListHash = computeListHash(list);

    // Update the displayed file path below the list
    showListFilePath();
}

int MultiReplace::checkForUnsavedChanges() {
    std::size_t currentListHash = computeListHash(replaceListData);

    if (currentListHash != originalListHash) {

        std::wstring message;
        if (!listFilePath.empty()) {
            // Get the shortened file path and build the message
            std::wstring shortenedFilePath = getShortenedFilePath(listFilePath, 500);
            message = LM.get(L"msgbox_unsaved_changes_file", { shortenedFilePath });
        }
        else {
            // If no file is associated, use the alternative message
            message = LM.get(L"msgbox_unsaved_changes");
        }

        // Show the MessageBox with the appropriate message
        int result = MessageBox(
            nppData._nppHandle,
            message.c_str(),
            LM.get(L"msgbox_title_save_list").c_str(),
            MB_ICONWARNING | MB_YESNOCANCEL
        );

        if (result == IDYES) {
            if (!listFilePath.empty()) {
                saveListToCsv(listFilePath, replaceListData);
                return IDYES;  // Proceed if saved successfully
            }
            else {
                std::wstring filePath = promptSaveListToCsv();

                if (!filePath.empty()) {
                    saveListToCsv(filePath, replaceListData);
                    return IDYES;  // Proceed with clear after save
                }
                else {
                    return IDCANCEL;  // Cancel if no file was selected
                }
            }
        }
        else if (result == IDNO) {
            return IDNO;  // Allow proceeding without saving
        }
        else if (result == IDCANCEL) {
            return IDCANCEL;  // Cancel the action
        }
    }

    return IDYES;  // No unsaved changes, allow proceeding
}

void MultiReplace::loadListFromCsvSilent(const std::wstring& filePath, std::vector<ReplaceItemData>& list) {
    // Open the CSV file
    std::ifstream inFile(std::filesystem::path(filePath), std::ios::binary);
    if (!inFile.is_open()) {
        std::wstring shortenedFilePathW = getShortenedFilePath(filePath, 500);
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_unable_to_open_file", { shortenedFilePathW })));
    }

    // Read raw bytes
    std::string raw((std::istreambuf_iterator<char>(inFile)), std::istreambuf_iterator<char>());
    inFile.close();

    // [BOM REMOVAL] strip BOM bytes if present
    size_t offset = 0;
    UINT cp = CP_UTF8;                                    // assume UTF-8
    if (raw.size() >= 3
        && static_cast<unsigned char>(raw[0]) == 0xEF
        && static_cast<unsigned char>(raw[1]) == 0xBB
        && static_cast<unsigned char>(raw[2]) == 0xBF)
    {
        offset = 3;                                       // skip BOM
    }
    else if (!Encoding::isValidUtf8(raw.data(), raw.size())) {
        cp = CP_ACP;                                      // fallback to ANSI
    }

    // Convert UTF-8 or ANSI bytes to std::wstring
    int wlen = MultiByteToWideChar(cp, 0, raw.c_str() + offset, static_cast<int>(raw.size() - offset), nullptr, 0);
    std::wstring content(wlen, L'\0');
    MultiByteToWideChar(cp, 0, raw.c_str() + offset, static_cast<int>(raw.size() - offset), &content[0], wlen);

    std::wstringstream contentStream(content);

    // Read the header line
    std::wstring headerLine;
    if (!std::getline(contentStream, headerLine)) {
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_column_count")));
    }

    // Temporary list to hold parsed items
    std::vector<ReplaceItemData> tempList;
    std::wstring line;

    // Read and process each line in the file
    while (std::getline(contentStream, line)) {
        std::vector<std::wstring> columns = StringUtils::parseCsvLine(line);

        // Check if column count is valid
        if (columns.size() < 8 || columns.size() > 9) {
            throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_column_count")));
        }

        try {
            ReplaceItemData item;
            item.isEnabled = std::stoi(columns[0]) != 0;
            item.findText = columns[1];
            item.replaceText = columns[2];
            item.wholeWord = std::stoi(columns[3]) != 0;
            item.matchCase = std::stoi(columns[4]) != 0;
            item.useVariables = std::stoi(columns[5]) != 0;
            item.extended = std::stoi(columns[6]) != 0;
            item.regex = std::stoi(columns[7]) != 0;
            item.comments = (columns.size() == 9) ? columns[8] : L"";
            tempList.push_back(item);
        }
        catch (const std::exception&) {
            throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_invalid_data_in_columns")));
        }
    }

    // Check if the file contains valid data rows
    if (tempList.empty()) {
        throw CsvLoadException(Encoding::wstringToUtf8(LM.get(L"status_no_valid_items_in_csv")));
    }

    // Transfer parsed data to the target list
    list = std::move(tempList);
}

void MultiReplace::loadListFromCsv(const std::wstring& filePath) {
    // Check for unsaved changes before loading new list
    if (checkForUnsavedChanges() == IDCANCEL) {
        return;
    }

    try {
        loadListFromCsvSilent(filePath, replaceListData);

        // Update the file path and display it
        listFilePath = filePath;
        showListFilePath();

        // Compute the hash of the loaded list
        originalListHash = computeListHash(replaceListData);

        // Clear Undo and Redo stacks
        URM.clear();

        // Update the ListView
        ListView_SetItemCountEx(_replaceListView, static_cast<int>(replaceListData.size()), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        // Display success or error message based on the loaded data
        if (replaceListData.empty()) {
            showStatusMessage(LM.get(L"status_no_valid_items_in_csv"), MessageStatus::Error);
        }
        else {
            showStatusMessage(LM.get(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), MessageStatus::Success);
        }
    }
    catch (const CsvLoadException& ex) {
        showStatusMessage(Encoding::utf8ToWString(ex.what()), MessageStatus::Error);
    }
}

void MultiReplace::checkForFileChangesAtStartup() {
    if (listFilePath.empty()) {
        return;
    }

    try {
        std::vector<ReplaceItemData> tempListFromFile;
        loadListFromCsvSilent(listFilePath, tempListFromFile);

        std::size_t newFileHash = computeListHash(tempListFromFile);
        if (newFileHash != originalListHash) {
            std::wstring shortenedFilePath = getShortenedFilePath(listFilePath, 500);
            std::wstring message = LM.get(L"msgbox_file_modified_prompt", { shortenedFilePath });

            int response = MessageBox(nppData._nppHandle, message.c_str(), LM.get(L"msgbox_title_reload").c_str(), MB_YESNO | MB_ICONWARNING | MB_SETFOREGROUND);
            if (response == IDYES) {
                replaceListData = tempListFromFile;
                originalListHash = newFileHash;
                ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
                InvalidateRect(_replaceListView, nullptr, TRUE);
            }
        }
    }
    catch (const CsvLoadException&) {
        // File no longer accessible - silently clear reference, keep temp data
        listFilePath.clear();
        originalListHash = computeListHash(replaceListData);
        showListFilePath();
    }

    // Status message for loaded list (from temp storage)
    if (replaceListData.empty()) {
        showStatusMessage(LM.get(L"status_no_valid_items_in_csv"), MessageStatus::Error);
    }
    else {
        showStatusMessage(LM.get(L"status_items_loaded_from_csv", { std::to_wstring(replaceListData.size()) }), MessageStatus::Success);
    }
}

void MultiReplace::exportToBashScript(const std::wstring& fileName) {
    std::ofstream file(fileName);
    if (!file.is_open()) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    // Create date
    time_t currentTime = time(nullptr);
    struct tm* localTime = localtime(&currentTime);
    char dateBuffer[80];
    strftime(dateBuffer, sizeof(dateBuffer), "%Y-%m-%d", localTime);

    file << "#!/bin/bash\n";
    file << "# Auto-generated by MultiReplace Notepad++\n";
    file << "# Created on: " << dateBuffer << "\n\n";
    file << "inputFile=\"$1\"\n";
    file << "outputFile=\"$2\"\n\n";

    file << "processLine() {\n";
    file << "    local findString=\"$1\"\n";
    file << "    local replaceString=\"$2\"\n";
    file << "    local wholeWord=\"$3\"\n";
    file << "    local matchCase=\"$4\"\n";
    file << "    local normal=\"$5\"\n";
    file << "    local extended=\"$6\"\n";
    file << "    local regex=\"$7\"\n";
    file << R"(
    if [[ "$wholeWord" -eq 1 ]]; then
        findString='\b'${findString}'\b'
    fi
    if [[ "$matchCase" -eq 1 ]]; then
        template='s|'${findString}'|'${replaceString}'|g'
    else
        template='s|'${findString}'|'${replaceString}'|gi'
    fi
    case 1 in
        $normal)
            sed -i "${template}" "$outputFile"
            ;;
        $extended)
            sed -i -e ':a' -e 'N' -e '$!ba' -e 's/\n/__NEWLINE__/g' -e 's/\r/__CARRIAGERETURN__/g' "$outputFile"
            sed -i "${template}" "$outputFile"
            sed -i 's/__NEWLINE__/\n/g; s/__CARRIAGERETURN__/\r/g' "$outputFile"
            ;;
        $regex)
            sed -i -r "${template}" "$outputFile"
            ;;
    esac
)";
    file << "}\n\n";

    file << "cp $inputFile $outputFile\n\n";

    file << "# processLine arguments: \"findString\" \"replaceString\" wholeWord matchCase normal extended regex\n";

    bool hasExcludedItems = false;
    for (const auto& itemData : replaceListData) {
        if (!itemData.isEnabled) continue; // Skip if this item is not selected
        if (itemData.useVariables) {
            hasExcludedItems = true; // Mark as excluded
            continue;
        }

        std::string find;
        std::string replace;
        if (itemData.extended) {
            find = SU::replaceNewline(SU::translateEscapes(SU::escapeSpecialChars(Encoding::wstringToUtf8(itemData.findText), true)), SU::ReplaceMode::Extended);
            replace = SU::replaceNewline(SU::translateEscapes(SU::escapeSpecialChars(Encoding::wstringToUtf8(itemData.replaceText), true)), SU::ReplaceMode::Extended);
        }
        else if (itemData.regex) {
            find = SU::replaceNewline(Encoding::wstringToUtf8(itemData.findText), SU::ReplaceMode::Regex);
            replace = SU::replaceNewline(Encoding::wstringToUtf8(itemData.replaceText), SU::ReplaceMode::Regex);
        }
        else {
            find = SU::replaceNewline(SU::escapeSpecialChars(Encoding::wstringToUtf8(itemData.findText), false), SU::ReplaceMode::Normal);
            replace = SU::replaceNewline(SU::escapeSpecialChars(Encoding::wstringToUtf8(itemData.replaceText), false), SU::ReplaceMode::Normal);
        }

        std::string wholeWord = itemData.wholeWord ? "1" : "0";
        std::string matchCase = itemData.matchCase ? "1" : "0";
        std::string normal = (!itemData.regex && !itemData.extended) ? "1" : "0";
        std::string extended = itemData.extended ? "1" : "0";
        std::string regex = itemData.regex ? "1" : "0";

        file << "processLine \"" << find << "\" \"" << replace << "\" " << wholeWord << " " << matchCase << " " << normal << " " << extended << " " << regex << "\n";
    }

    file.close();

    if (file.fail()) {
        showStatusMessage(LM.get(L"status_unable_to_save_file"), MessageStatus::Error);
        return;
    }

    showStatusMessage(LM.get(L"status_list_exported_to_bash"), MessageStatus::Success);

    // Show message box if excluded items were found
    if (hasExcludedItems) {
        MessageBox(_hSelf,
            LM.get(L"msgbox_use_variables_not_exported").c_str(),
            LM.get(L"msgbox_title_warning").c_str(),
            MB_OK | MB_ICONWARNING);
    }

}

#pragma endregion


#pragma region INI

std::pair<std::wstring, std::wstring> MultiReplace::generateConfigFilePaths() {
    wchar_t configDir[MAX_PATH] = {};
    ::SendMessage(nppData._nppHandle, NPPM_GETPLUGINSCONFIGDIR, MAX_PATH, reinterpret_cast<LPARAM>(configDir));
    configDir[MAX_PATH - 1] = '\0'; // Ensure null-termination

    std::wstring iniFilePath = std::wstring(configDir) + L"\\MultiReplace.ini";
    std::wstring csvFilePath = std::wstring(configDir) + L"\\MultiReplaceList.ini";
    return { iniFilePath, csvFilePath };
}

void MultiReplace::saveSettings() {
    static bool settingsSaved = false;
    if (settingsSaved) {
        return;
    }

    auto [iniFilePath, csvFilePath] = generateConfigFilePaths();

    try {
        syncUIToCache();
        CFG.save(iniFilePath);
        saveListToCsvSilent(csvFilePath, replaceListData);
    }
    catch (const std::exception& ex) {
        std::wstring errorMessage = LM.get(L"msgbox_error_saving_settings",
            { std::wstring(ex.what(), ex.what() + strlen(ex.what())) });
        MessageBox(nppData._nppHandle, errorMessage.c_str(),
            LM.get(L"msgbox_title_error").c_str(), MB_OK | MB_ICONERROR | MB_SETFOREGROUND);
    }
    settingsSaved = true;
}

void MultiReplace::loadSettingsToPanelUI() {

    // Loading the history for the Find text field in reverse order
    int findHistoryCount = CFG.readInt(L"History", L"FindTextHistoryCount", 0);
    for (int i = findHistoryCount - 1; i >= 0; --i) {
        std::wstring findHistoryItem = CFG.readString(L"History", L"FindTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FIND_EDIT), findHistoryItem);
    }

    // Loading the history for the Replace text field in reverse order
    int replaceHistoryCount = CFG.readInt(L"History", L"ReplaceTextHistoryCount", 0);
    for (int i = replaceHistoryCount - 1; i >= 0; --i) {
        std::wstring replaceHistoryItem = CFG.readString(L"History", L"ReplaceTextHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), replaceHistoryItem);
    }

    // Load Filter history
    int filterHistoryCount = CFG.readInt(L"History", L"FilterHistoryCount", 0);
    for (int i = filterHistoryCount - 1; i >= 0; --i) {
        std::wstring item = CFG.readString(L"History", L"FilterHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_FILTER_EDIT), item);
    }

    // Load Dir history
    int dirHistoryCount = CFG.readInt(L"History", L"DirHistoryCount", 0);
    for (int i = dirHistoryCount - 1; i >= 0; --i) {
        std::wstring item = CFG.readString(L"History", L"DirHistory" + std::to_wstring(i), L"");
        addStringToComboBoxHistory(GetDlgItem(_hSelf, IDC_DIR_EDIT), item);
    }

    // Reading and setting current "Find what" and "Replace with" texts
    std::wstring findText = CFG.readString(L"Current", L"FindText", L"");
    std::wstring replaceText = CFG.readString(L"Current", L"ReplaceText", L"");
    setTextInDialogItem(_hSelf, IDC_FIND_EDIT, findText);
    setTextInDialogItem(_hSelf, IDC_REPLACE_EDIT, replaceText);

    // Setting options based on the INI file
    bool wholeWord = CFG.readBool(L"Options", L"WholeWord", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WHOLE_WORD_CHECKBOX), BM_SETCHECK, wholeWord ? BST_CHECKED : BST_UNCHECKED, 0);

    bool matchCase = CFG.readBool(L"Options", L"MatchCase", false);
    SendMessage(GetDlgItem(_hSelf, IDC_MATCH_CASE_CHECKBOX), BM_SETCHECK, matchCase ? BST_CHECKED : BST_UNCHECKED, 0);

    bool useVariables = CFG.readBool(L"Options", L"UseVariables", false);
    SendMessage(GetDlgItem(_hSelf, IDC_USE_VARIABLES_CHECKBOX), BM_SETCHECK, useVariables ? BST_CHECKED : BST_UNCHECKED, 0);

    // Selecting the appropriate search mode radio button based on the settings
    bool extended = CFG.readBool(L"Options", L"Extended", false);
    bool regex = CFG.readBool(L"Options", L"Regex", false);
    if (regex) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_REGEX_RADIO);
    }
    else if (extended) {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_EXTENDED_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_NORMAL_RADIO, IDC_REGEX_RADIO, IDC_NORMAL_RADIO);
    }

    // Setting additional options
    bool wrapAround = CFG.readBool(L"Options", L"WrapAround", false);
    SendMessage(GetDlgItem(_hSelf, IDC_WRAP_AROUND_CHECKBOX), BM_SETCHECK, wrapAround ? BST_CHECKED : BST_UNCHECKED, 0);

    bool replaceAtMatches = CFG.readBool(L"Options", L"ReplaceAtMatches", false);
    SendMessage(GetDlgItem(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX), BM_SETCHECK, replaceAtMatches ? BST_CHECKED : BST_UNCHECKED, 0);

    std::wstring editAtMatchesText = CFG.readString(L"Options", L"EditAtMatches", L"1");
    setTextInDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT, editAtMatchesText);

    bool replaceButtonsMode = CFG.readBool(L"Options", L"ButtonsMode", false);
    SendMessage(GetDlgItem(_hSelf, IDC_2_BUTTONS_MODE), BM_SETCHECK, replaceButtonsMode ? BST_CHECKED : BST_UNCHECKED, 0);

    useListEnabled = CFG.readBool(L"Options", L"UseList", true);
    updateUseListState(false);

    ResultDock::setWrapEnabled(CFG.readBool(L"Options", L"DockWrap", false));
    ResultDock::setPurgeEnabled(CFG.readBool(L"Options", L"DockPurge", false));

    highlightMatchEnabled = CFG.readBool(L"Options", L"HighlightMatch", true);
    flowTabsIntroDontShowEnabled = CFG.readBool(L"Options", L"FlowTabsIntroDontShow", false);
    flowTabsNumericAlignEnabled = CFG.readBool(L"Options", L"FlowTabsNumericAlign", true);

    exportToBashEnabled = CFG.readBool(L"Options", L"ExportToBash", false);
    muteSounds = CFG.readBool(L"Options", L"MuteSounds", false);
    doubleClickEditsEnabled = CFG.readBool(L"Options", L"DoubleClickEdits", true);

    // Side Effect: Update Hover Text Logic
    bool newHover = CFG.readBool(L"Options", L"HoverText", true);
    if (isHoverTextEnabled != newHover) {
        isHoverTextEnabled = newHover;
        if (instance && instance->_replaceListView) {
            DWORD ex = ListView_GetExtendedListViewStyle(instance->_replaceListView);
            if (isHoverTextEnabled) ex |= LVS_EX_INFOTIP; else ex &= ~LVS_EX_INFOTIP;
            ListView_SetExtendedListViewStyle(instance->_replaceListView, ex);
        }
    }

    // CFG.readInt for editFieldSize
    editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 5);
    editFieldSize = std::clamp(editFieldSize, MIN_EDIT_FIELD_SIZE, MAX_EDIT_FIELD_SIZE);

    listStatisticsEnabled = CFG.readBool(L"Options", L"ListStatistics", false);
    stayAfterReplaceEnabled = CFG.readBool(L"Options", L"StayAfterReplace", false);
    allFromCursorEnabled = CFG.readBool(L"Options", L"AllFromCursor", false);
    groupResultsEnabled = CFG.readBool(L"Options", L"GroupResults", false);

    // Lua runtime options
    luaSafeModeEnabled = CFG.readBool(L"Lua", L"SafeMode", false);

    // Loading and setting the scope
    int selection = CFG.readInt(L"Scope", L"Selection", 0);
    int columnMode = CFG.readInt(L"Scope", L"ColumnMode", 0);

    // Reading and setting specific scope settings
    std::wstring columnNum = CFG.readString(L"Scope", L"ColumnNum", L"1-50");
    setTextInDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT, columnNum);

    std::wstring delimiter = CFG.readString(L"Scope", L"Delimiter", L",");
    setTextInDialogItem(_hSelf, IDC_DELIMITER_EDIT, delimiter);

    std::wstring quoteChar = CFG.readString(L"Scope", L"QuoteChar", L"\"");
    setTextInDialogItem(_hSelf, IDC_QUOTECHAR_EDIT, quoteChar);

    CSVheaderLinesCount = CFG.readInt(L"Scope", L"HeaderLines", 1);

    // --- Load “Replace in Files” settings --------------------------
    std::wstring filter = CFG.readString(L"ReplaceInFiles", L"Filter", L"*.*");
    setTextInDialogItem(_hSelf, IDC_FILTER_EDIT, filter);

    std::wstring dir = CFG.readString(L"ReplaceInFiles", L"Directory", L"");
    setTextInDialogItem(_hSelf, IDC_DIR_EDIT, dir);

    bool inSub = CFG.readBool(L"ReplaceInFiles", L"InSubfolders", false);
    SendMessage(GetDlgItem(_hSelf, IDC_SUBFOLDERS_CHECKBOX), BM_SETCHECK, inSub ? BST_CHECKED : BST_UNCHECKED, 0);

    bool inHidden = CFG.readBool(L"ReplaceInFiles", L"InHiddenFolders", false);
    SendMessage(GetDlgItem(_hSelf, IDC_HIDDENFILES_CHECKBOX), BM_SETCHECK, inHidden ? BST_CHECKED : BST_UNCHECKED, 0);

    limitFileSizeEnabled = CFG.readBool(L"ReplaceInFiles", L"LimitFileSize", false);
    maxFileSizeMB = static_cast<size_t>(CFG.readInt(L"ReplaceInFiles", L"MaxFileSizeMB", 100));

    // --- Load Column Settings (Widths & Visibility) ---
    // NOTE: We read them here, and apply them in the 'instance' block below
    findCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindCountWidth", DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceCountWidth", DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);
    findColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindWidth", DEFAULT_COLUMN_WIDTH_FIND_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceWidth", DEFAULT_COLUMN_WIDTH_REPLACE_scaled), MIN_GENERAL_WIDTH_scaled);
    commentsColumnWidth = std::max(CFG.readInt(L"ListColumns", L"CommentsWidth", DEFAULT_COLUMN_WIDTH_COMMENTS_scaled), MIN_GENERAL_WIDTH_scaled);

    isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);

    findColumnLockedEnabled = CFG.readBool(L"ListColumns", L"FindColumnLocked", true);
    replaceColumnLockedEnabled = CFG.readBool(L"ListColumns", L"ReplaceColumnLocked", false);
    commentsColumnLockedEnabled = CFG.readBool(L"ListColumns", L"CommentsColumnLocked", true);

    // Load file path and original hash
    listFilePath = CFG.readString(L"File", L"ListFilePath", L"");
    originalListHash = CFG.readSizeT(L"File", L"OriginalListHash", 0);

    if (selection) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_SELECTION_RADIO);
        onSelectionChanged();
    }
    else if (columnMode) {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_COLUMN_MODE_RADIO);
    }
    else {
        CheckRadioButton(_hSelf, IDC_ALL_TEXT_RADIO, IDC_COLUMN_MODE_RADIO, IDC_ALL_TEXT_RADIO);
    }

    setUIElementVisibility();

    // --- FINAL UI REFRESH / SIDE EFFECTS ---
    if (instance) {
        // 1. Apply Tooltips state
        bool currentTooltips = CFG.readBool(L"Options", L"Tooltips", true);
        if (tooltipsEnabled != currentTooltips) {
            tooltipsEnabled = currentTooltips;
            instance->onTooltipsToggled(tooltipsEnabled);
        }
        else {
            tooltipsEnabled = currentTooltips; // Ensure sync
        }

        // 2. Update Stats/Path UI
        instance->updateUseListState(false);
        instance->showListFilePath();

        // 3. Update Export Button visibility
        HWND hBash = GetDlgItem(instance->_hSelf, IDC_EXPORT_BASH_BUTTON);
        if (hBash) ShowWindow(hBash, exportToBashEnabled ? SW_SHOW : SW_HIDE);

        // 4. Rebuild ListView columns to reflect visibility changes
        if (instance->_replaceListView) {
            instance->createListViewColumns(); // Applies isFindCountVisible etc.

            // Refresh data and redraw
            ListView_SetItemCountEx(instance->_replaceListView, instance->replaceListData.size(), LVSICF_NOINVALIDATEALL);
            InvalidateRect(instance->_replaceListView, nullptr, TRUE);

            // Update Header Icons (Checkmarks)
            instance->updateHeaderSelection();
        }
    }
}

void MultiReplace::loadSettings() {
    auto [_, csvFilePath] = generateConfigFilePaths();

    // Read all settings from the cache
    loadSettingsToPanelUI();

    try {
        loadListFromCsvSilent(csvFilePath, replaceListData);
    }
    catch (const CsvLoadException& ex) {
        std::wstring errorMessage = L"An error occurred while loading the settings: ";
        errorMessage += std::wstring(ex.what(), ex.what() + strlen(ex.what()));
        // MessageBox(nullptr, errorMessage.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
    updateHeaderSelection();
    ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(_replaceListView, nullptr, TRUE);

    showListFilePath();
}

void MultiReplace::loadUIConfigFromIni()
{
    if (!dpiMgr) return;

    float oldScale = dpiMgr->getCustomScaleFactor();

    float customScaleFactor = CFG.readFloat(L"Window", L"ScaleFactor", 1.0f);
    dpiMgr->setCustomScaleFactor(customScaleFactor);

    // --- scaled constants --------------------------------------
    MIN_WIDTH_scaled = sx(MIN_WIDTH);
    MIN_HEIGHT_scaled = sy(MIN_HEIGHT);
    SHRUNK_HEIGHT_scaled = sy(SHRUNK_HEIGHT);
    DEFAULT_COLUMN_WIDTH_FIND_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND);
    DEFAULT_COLUMN_WIDTH_REPLACE_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE);
    DEFAULT_COLUMN_WIDTH_COMMENTS_scaled = sx(DEFAULT_COLUMN_WIDTH_COMMENTS);
    DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_FIND_COUNT);
    DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled = sx(DEFAULT_COLUMN_WIDTH_REPLACE_COUNT);
    MIN_GENERAL_WIDTH_scaled = sx(MIN_GENERAL_WIDTH);

    // --- Hot-Reload Logic for Scaling --------------------------
    if (std::abs(oldScale - customScaleFactor) > 0.001f)
    {
        createFonts();
        applyFonts();

        RECT rc;
        if (GetWindowRect(_hSelf, &rc))
        {
            int currentW = rc.right - rc.left;
            int currentH = rc.bottom - rc.top;

            float ratio = customScaleFactor / oldScale;

            int newW = static_cast<int>(currentW * ratio);
            int newH = static_cast<int>(currentH * ratio);

            // This triggers WM_SIZE -> moveAndResizeControls(false)
            // Static controls will NOT move yet.
            SetWindowPos(_hSelf, nullptr, 0, 0, newW, newH,
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);

            adjustWindowSize();

            // We pass 'true' to override the static filter.
            moveAndResizeControls(true);
        }
    }

    int savedLeft = CFG.readInt(L"Window", L"PosX", CENTER_ON_NPP);
    int savedTop = CFG.readInt(L"Window", L"PosY", CENTER_ON_NPP);

    useListEnabled = CFG.readBool(L"Options", L"UseList", true);
    updateUseListState(false);

    int savedWidth = CFG.readInt(L"Window", L"Width", sx(INIT_WIDTH));
    int width = (savedWidth < MIN_WIDTH_scaled) ? MIN_WIDTH_scaled : savedWidth;

    useListOnHeight = CFG.readInt(L"Window", L"Height", sy(INIT_HEIGHT));
    if (useListOnHeight < MIN_HEIGHT_scaled) useListOnHeight = MIN_HEIGHT_scaled;

    int height = useListEnabled ? useListOnHeight : useListOffHeight;

    // Center over Notepad++ on first run (when no position saved in INI)
    if (savedLeft == CENTER_ON_NPP || savedTop == CENTER_ON_NPP) {
        RECT rcNpp;
        if (::GetWindowRect(nppData._nppHandle, &rcNpp)) {
            int nppW = rcNpp.right - rcNpp.left;
            int nppH = rcNpp.bottom - rcNpp.top;
            windowRect.left = rcNpp.left + (nppW - width) / 2;
            windowRect.top = rcNpp.top + (nppH - height) / 2;
        }
        else {
            // Fallback if GetWindowRect fails
            windowRect.left = 100;
            windowRect.top = 100;
        }
    }
    else {
        windowRect.left = savedLeft;
        windowRect.top = savedTop;
    }

    windowRect.right = windowRect.left + width;
    windowRect.bottom = windowRect.top + height;

    findColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindWidth", DEFAULT_COLUMN_WIDTH_FIND_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceWidth", DEFAULT_COLUMN_WIDTH_REPLACE_scaled), MIN_GENERAL_WIDTH_scaled);
    commentsColumnWidth = std::max(CFG.readInt(L"ListColumns", L"CommentsWidth", DEFAULT_COLUMN_WIDTH_COMMENTS_scaled), MIN_GENERAL_WIDTH_scaled);
    findCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"FindCountWidth", DEFAULT_COLUMN_WIDTH_FIND_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);
    replaceCountColumnWidth = std::max(CFG.readInt(L"ListColumns", L"ReplaceCountWidth", DEFAULT_COLUMN_WIDTH_REPLACE_COUNT_scaled), MIN_GENERAL_WIDTH_scaled);

    isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);

    findColumnLockedEnabled = CFG.readBool(L"ListColumns", L"FindColumnLocked", true);
    replaceColumnLockedEnabled = CFG.readBool(L"ListColumns", L"ReplaceColumnLocked", false);
    commentsColumnLockedEnabled = CFG.readBool(L"ListColumns", L"CommentsColumnLocked", true);

    int fg = static_cast<int>(CFG.readByte(L"Window", L"ForegroundTransparency", 255));
    int bg = static_cast<int>(CFG.readByte(L"Window", L"BackgroundTransparency", 190));
    if (fg < 0) fg = 0; if (fg > 255) fg = 255;
    if (bg < 0) bg = 0; if (bg > 255) bg = 255;
    foregroundTransparency = static_cast<BYTE>(fg);
    backgroundTransparency = static_cast<BYTE>(bg);

    tooltipsEnabled = CFG.readBool(L"Options", L"Tooltips", true);
    isHoverTextEnabled = CFG.readBool(L"Options", L"HoverText", true);

    resultDockPerEntryColorsEnabled = CFG.readBool(L"Options", L"ResultDockPerEntryColors", true);
    useListColorsForMarking = CFG.readBool(L"Options", L"UseListColorsForMarking", true);
    ResultDock::setPerEntryColorsEnabled(resultDockPerEntryColorsEnabled);

    if (_replaceListView)
    {
        createListViewColumns();
        ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);

        DWORD ex = ListView_GetExtendedListViewStyle(_replaceListView);
        if (isHoverTextEnabled) ex |= LVS_EX_INFOTIP; else ex &= ~LVS_EX_INFOTIP;
        ListView_SetExtendedListViewStyle(_replaceListView, ex);

        updateHeaderSelection();
    }

    if (_hSelf)
        SetWindowTransparency(_hSelf, foregroundTransparency);
}

MultiReplace::Settings MultiReplace::getSettings()
{
    // Always read from INI cache (single source of truth)
    Settings s{};
    s.tooltipsEnabled = CFG.readBool(L"Options", L"Tooltips", true);
    s.exportToBashEnabled = CFG.readBool(L"Options", L"ExportToBash", false);
    s.muteSounds = CFG.readBool(L"Options", L"MuteSounds", false);
    s.doubleClickEditsEnabled = CFG.readBool(L"Options", L"DoubleClickEdits", true);
    s.highlightMatchEnabled = CFG.readBool(L"Options", L"HighlightMatch", true);
    s.flowTabsIntroDontShowEnabled = CFG.readBool(L"Options", L"FlowTabsIntroDontShow", false);
    s.flowTabsNumericAlignEnabled = CFG.readBool(L"Options", L"FlowTabsNumericAlign", true);
    s.isHoverTextEnabled = CFG.readBool(L"Options", L"HoverText", true);
    s.listStatisticsEnabled = CFG.readBool(L"Options", L"ListStatistics", false);
    s.stayAfterReplaceEnabled = CFG.readBool(L"Options", L"StayAfterReplace", false);
    s.groupResultsEnabled = CFG.readBool(L"Options", L"GroupResults", false);
    s.allFromCursorEnabled = CFG.readBool(L"Options", L"AllFromCursor", false);
    s.limitFileSizeEnabled = CFG.readBool(L"ReplaceInFiles", L"LimitFileSize", false);
    s.maxFileSizeMB = CFG.readInt(L"ReplaceInFiles", L"MaxFileSizeMB", 100);
    s.isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    s.isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    s.isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    s.isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);
    s.editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 5);
    s.csvHeaderLinesCount = CFG.readInt(L"Scope", L"HeaderLines", 1);
    s.resultDockPerEntryColorsEnabled = CFG.readBool(L"Options", L"ResultDockPerEntryColors", true);
    s.useListColorsForMarking = CFG.readBool(L"Options", L"UseListColorsForMarking", true);
    s.duplicateBookmarksEnabled = CFG.readBool(L"Options", L"DuplicateBookmarks", false);
    return s;
}

void MultiReplace::writeStructToConfig(const Settings& s)
{
    // Write all logic options to ConfigManager
    CFG.writeInt(L"Options", L"Tooltips", s.tooltipsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"ExportToBash", s.exportToBashEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"MuteSounds", s.muteSounds ? 1 : 0);
    CFG.writeInt(L"Options", L"DoubleClickEdits", s.doubleClickEditsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"HighlightMatch", s.highlightMatchEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"FlowTabsIntroDontShow", s.flowTabsIntroDontShowEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"FlowTabsNumericAlign", s.flowTabsNumericAlignEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"HoverText", s.isHoverTextEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"ListStatistics", s.listStatisticsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"StayAfterReplace", s.stayAfterReplaceEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"GroupResults", s.groupResultsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"AllFromCursor", s.allFromCursorEnabled ? 1 : 0);
    CFG.writeInt(L"ReplaceInFiles", L"LimitFileSize", s.limitFileSizeEnabled ? 1 : 0);
    CFG.writeInt(L"ReplaceInFiles", L"MaxFileSizeMB", s.maxFileSizeMB);
    CFG.writeInt(L"ListColumns", L"FindCountVisible", s.isFindCountVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"ReplaceCountVisible", s.isReplaceCountVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"CommentsVisible", s.isCommentsColumnVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"DeleteButtonVisible", s.isDeleteButtonVisible ? 1 : 0);
    CFG.writeInt(L"Options", L"EditFieldSize", s.editFieldSize);
    CFG.writeInt(L"Scope", L"HeaderLines", s.csvHeaderLinesCount);
    CFG.writeInt(L"Options", L"ResultDockPerEntryColors", s.resultDockPerEntryColorsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"UseListColorsForMarking", s.useListColorsForMarking ? 1 : 0);
    CFG.writeInt(L"Options", L"DuplicateBookmarks", s.duplicateBookmarksEnabled ? 1 : 0);
}

void MultiReplace::loadConfigOnce()
{
    auto [iniFilePath, _] = generateConfigFilePaths();
    CFG.load(iniFilePath);
}

void MultiReplace::syncUIToCache()
{
    // Window Position & Size
    RECT currentRect;
    GetWindowRect(_hSelf, &currentRect);
    CFG.writeInt(L"Window", L"PosX", currentRect.left);
    CFG.writeInt(L"Window", L"PosY", currentRect.top);
    CFG.writeInt(L"Window", L"Width", currentRect.right - currentRect.left);

    int height = currentRect.bottom - currentRect.top;
    if (useListEnabled) {
        useListOnHeight = height;
    }
    CFG.writeInt(L"Window", L"Height", useListOnHeight);

    // ScaleFactor: truncate to 1 decimal place like original code
    std::wstring scaleStr = std::to_wstring(dpiMgr->getCustomScaleFactor());
    size_t dotPos = scaleStr.find(L'.');
    if (dotPos != std::wstring::npos && dotPos + 2 < scaleStr.length()) {
        scaleStr = scaleStr.substr(0, dotPos + 2);
    }
    CFG.writeString(L"Window", L"ScaleFactor", scaleStr);

    CFG.writeInt(L"Window", L"ForegroundTransparency", foregroundTransparency);
    CFG.writeInt(L"Window", L"BackgroundTransparency", backgroundTransparency);

    // Column Widths
    if (_replaceListView) {
        if (columnIndices[ColumnID::FIND_COUNT] != -1)
            findCountColumnWidth = ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_COUNT]);
        if (columnIndices[ColumnID::REPLACE_COUNT] != -1)
            replaceCountColumnWidth = ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_COUNT]);
        if (columnIndices[ColumnID::FIND_TEXT] != -1)
            findColumnWidth = ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::FIND_TEXT]);
        if (columnIndices[ColumnID::REPLACE_TEXT] != -1)
            replaceColumnWidth = ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::REPLACE_TEXT]);
        if (columnIndices[ColumnID::COMMENTS] != -1)
            commentsColumnWidth = ListView_GetColumnWidth(_replaceListView, columnIndices[ColumnID::COMMENTS]);
    }

    CFG.writeInt(L"ListColumns", L"FindCountWidth", findCountColumnWidth);
    CFG.writeInt(L"ListColumns", L"ReplaceCountWidth", replaceCountColumnWidth);
    CFG.writeInt(L"ListColumns", L"FindWidth", findColumnWidth);
    CFG.writeInt(L"ListColumns", L"ReplaceWidth", replaceColumnWidth);
    CFG.writeInt(L"ListColumns", L"CommentsWidth", commentsColumnWidth);

    CFG.writeInt(L"ListColumns", L"FindCountVisible", isFindCountVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"ReplaceCountVisible", isReplaceCountVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"CommentsVisible", isCommentsColumnVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"DeleteButtonVisible", isDeleteButtonVisible ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"FindColumnLocked", findColumnLockedEnabled ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"ReplaceColumnLocked", replaceColumnLockedEnabled ? 1 : 0);
    CFG.writeInt(L"ListColumns", L"CommentsColumnLocked", commentsColumnLockedEnabled ? 1 : 0);

    // Current Find/Replace Text
    CFG.writeString(L"Current", L"FindText", getTextFromDialogItem(_hSelf, IDC_FIND_EDIT));
    CFG.writeString(L"Current", L"ReplaceText", getTextFromDialogItem(_hSelf, IDC_REPLACE_EDIT));

    // Search Options
    CFG.writeInt(L"Options", L"WholeWord", IsDlgButtonChecked(_hSelf, IDC_WHOLE_WORD_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"MatchCase", IsDlgButtonChecked(_hSelf, IDC_MATCH_CASE_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"Extended", IsDlgButtonChecked(_hSelf, IDC_EXTENDED_RADIO) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"Regex", IsDlgButtonChecked(_hSelf, IDC_REGEX_RADIO) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"WrapAround", IsDlgButtonChecked(_hSelf, IDC_WRAP_AROUND_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"UseVariables", IsDlgButtonChecked(_hSelf, IDC_USE_VARIABLES_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"ReplaceAtMatches", IsDlgButtonChecked(_hSelf, IDC_REPLACE_AT_MATCHES_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"ButtonsMode", IsDlgButtonChecked(_hSelf, IDC_2_BUTTONS_MODE) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Options", L"UseList", useListEnabled ? 1 : 0);
    CFG.writeString(L"Options", L"EditAtMatches", getTextFromDialogItem(_hSelf, IDC_REPLACE_HIT_EDIT));

    // Config-managed Options
    CFG.writeInt(L"Options", L"Tooltips", tooltipsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"HighlightMatch", highlightMatchEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"FlowTabsIntroDontShow", flowTabsIntroDontShowEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"FlowTabsNumericAlign", flowTabsNumericAlignEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"ExportToBash", exportToBashEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"MuteSounds", muteSounds ? 1 : 0);
    CFG.writeInt(L"Options", L"DoubleClickEdits", doubleClickEditsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"HoverText", isHoverTextEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"EditFieldSize", editFieldSize);
    CFG.writeInt(L"Options", L"ListStatistics", listStatisticsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"StayAfterReplace", stayAfterReplaceEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"AllFromCursor", allFromCursorEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"GroupResults", groupResultsEnabled ? 1 : 0);
    CFG.writeInt(L"Options", L"DockWrap", ResultDock::wrapEnabled() ? 1 : 0);
    CFG.writeInt(L"Options", L"DockPurge", ResultDock::purgeEnabled() ? 1 : 0);

    // Lua Options
    CFG.writeInt(L"Lua", L"SafeMode", luaSafeModeEnabled ? 1 : 0);

    // Scope Settings
    CFG.writeInt(L"Scope", L"Selection", IsDlgButtonChecked(_hSelf, IDC_SELECTION_RADIO) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"Scope", L"ColumnMode", IsDlgButtonChecked(_hSelf, IDC_COLUMN_MODE_RADIO) == BST_CHECKED ? 1 : 0);
    CFG.writeString(L"Scope", L"ColumnNum", getTextFromDialogItem(_hSelf, IDC_COLUMN_NUM_EDIT));
    CFG.writeString(L"Scope", L"Delimiter", getTextFromDialogItem(_hSelf, IDC_DELIMITER_EDIT));
    CFG.writeString(L"Scope", L"QuoteChar", getTextFromDialogItem(_hSelf, IDC_QUOTECHAR_EDIT));
    CFG.writeInt(L"Scope", L"HeaderLines", static_cast<int>(CSVheaderLinesCount));

    // Replace in Files Settings
    CFG.writeString(L"ReplaceInFiles", L"Filter", getTextFromDialogItem(_hSelf, IDC_FILTER_EDIT));
    CFG.writeString(L"ReplaceInFiles", L"Directory", getTextFromDialogItem(_hSelf, IDC_DIR_EDIT));
    CFG.writeInt(L"ReplaceInFiles", L"InSubfolders", IsDlgButtonChecked(_hSelf, IDC_SUBFOLDERS_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"ReplaceInFiles", L"InHiddenFolders", IsDlgButtonChecked(_hSelf, IDC_HIDDENFILES_CHECKBOX) == BST_CHECKED ? 1 : 0);
    CFG.writeInt(L"ReplaceInFiles", L"LimitFileSize", limitFileSizeEnabled ? 1 : 0);
    CFG.writeInt(L"ReplaceInFiles", L"MaxFileSizeMB", static_cast<int>(maxFileSizeMB));

    // File Info
    CFG.writeString(L"File", L"ListFilePath", listFilePath);
    CFG.writeSizeT(L"File", L"OriginalListHash", originalListHash);

    // History 
    syncHistoryToCache(GetDlgItem(_hSelf, IDC_FIND_EDIT), L"FindTextHistory");
    syncHistoryToCache(GetDlgItem(_hSelf, IDC_REPLACE_EDIT), L"ReplaceTextHistory");
    syncHistoryToCache(GetDlgItem(_hSelf, IDC_FILTER_EDIT), L"FilterHistory");
    syncHistoryToCache(GetDlgItem(_hSelf, IDC_DIR_EDIT), L"DirHistory");
}

void MultiReplace::syncHistoryToCache(HWND hComboBox, const std::wstring& keyPrefix)
{
    LRESULT count = SendMessage(hComboBox, CB_GETCOUNT, 0, 0);
    int itemsToSave = std::min(static_cast<int>(count), maxHistoryItems);

    CFG.writeInt(L"History", keyPrefix + L"Count", itemsToSave);

    for (int i = 0; i < itemsToSave; ++i) {
        LRESULT len = SendMessage(hComboBox, CB_GETLBTEXTLEN, i, 0);
        std::vector<wchar_t> buffer(static_cast<size_t>(len + 1));
        SendMessage(hComboBox, CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer.data()));
        std::wstring text(buffer.data());
        CFG.writeString(L"History", keyPrefix + std::to_wstring(i), text);
    }
}

void MultiReplace::applyConfigSettingsOnly()
{
    // Tooltip
    bool newTooltips = CFG.readBool(L"Options", L"Tooltips", true);
    if (tooltipsEnabled != newTooltips) {
        tooltipsEnabled = newTooltips;
        onTooltipsToggled(tooltipsEnabled);
    }

    // Boolean Options
    muteSounds = CFG.readBool(L"Options", L"MuteSounds", false);
    doubleClickEditsEnabled = CFG.readBool(L"Options", L"DoubleClickEdits", true);
    highlightMatchEnabled = CFG.readBool(L"Options", L"HighlightMatch", true);
    listStatisticsEnabled = CFG.readBool(L"Options", L"ListStatistics", false);
    stayAfterReplaceEnabled = CFG.readBool(L"Options", L"StayAfterReplace", false);
    allFromCursorEnabled = CFG.readBool(L"Options", L"AllFromCursor", false);
    groupResultsEnabled = CFG.readBool(L"Options", L"GroupResults", false);
    flowTabsIntroDontShowEnabled = CFG.readBool(L"Options", L"FlowTabsIntroDontShow", false);
    flowTabsNumericAlignEnabled = CFG.readBool(L"Options", L"FlowTabsNumericAlign", true);
    exportToBashEnabled = CFG.readBool(L"Options", L"ExportToBash", false);
    luaSafeModeEnabled = CFG.readBool(L"Lua", L"SafeMode", false);
    limitFileSizeEnabled = CFG.readBool(L"ReplaceInFiles", L"LimitFileSize", false);
    maxFileSizeMB = CFG.readInt(L"ReplaceInFiles", L"MaxFileSizeMB", 100);

    resultDockPerEntryColorsEnabled = CFG.readBool(L"Options", L"ResultDockPerEntryColors", true);
    useListColorsForMarking = CFG.readBool(L"Options", L"UseListColorsForMarking", true);
    ResultDock::setPerEntryColorsEnabled(resultDockPerEntryColorsEnabled);

    // Just update the bookmark setting - no action needed
    // Bookmarks will be handled correctly on next duplicate scan
    _duplicateBookmarksEnabled = CFG.readBool(L"Options", L"DuplicateBookmarks", false);

    // Hover Text
    bool newHover = CFG.readBool(L"Options", L"HoverText", true);
    if (isHoverTextEnabled != newHover) {
        isHoverTextEnabled = newHover;
        if (_replaceListView) {
            DWORD ex = ListView_GetExtendedListViewStyle(_replaceListView);
            if (isHoverTextEnabled) ex |= LVS_EX_INFOTIP; else ex &= ~LVS_EX_INFOTIP;
            ListView_SetExtendedListViewStyle(_replaceListView, ex);
        }
    }

    // Integer Options
    editFieldSize = CFG.readInt(L"Options", L"EditFieldSize", 5);
    editFieldSize = std::clamp(editFieldSize, MIN_EDIT_FIELD_SIZE, MAX_EDIT_FIELD_SIZE);

    CSVheaderLinesCount = CFG.readInt(L"Scope", L"HeaderLines", 1);

    // Column Visibility
    isFindCountVisible = CFG.readBool(L"ListColumns", L"FindCountVisible", false);
    isReplaceCountVisible = CFG.readBool(L"ListColumns", L"ReplaceCountVisible", false);
    isCommentsColumnVisible = CFG.readBool(L"ListColumns", L"CommentsVisible", false);
    isDeleteButtonVisible = CFG.readBool(L"ListColumns", L"DeleteButtonVisible", true);

    // Transparency
    int fg = CFG.readInt(L"Window", L"ForegroundTransparency", 255);
    int bg = CFG.readInt(L"Window", L"BackgroundTransparency", 190);
    fg = std::clamp(fg, 0, 255);
    bg = std::clamp(bg, 0, 255);
    foregroundTransparency = static_cast<BYTE>(fg);
    backgroundTransparency = static_cast<BYTE>(bg);
    if (_hSelf) SetWindowTransparency(_hSelf, foregroundTransparency);

    // UI Updates
    HWND hBash = GetDlgItem(_hSelf, IDC_EXPORT_BASH_BUTTON);
    if (hBash) ShowWindow(hBash, exportToBashEnabled ? SW_SHOW : SW_HIDE);

    updateUseListState(false);
    showListFilePath();

    if (_replaceListView) {
        createListViewColumns();
        ListView_SetItemCountEx(_replaceListView, replaceListData.size(), LVSICF_NOINVALIDATEALL);
        InvalidateRect(_replaceListView, nullptr, TRUE);
        updateHeaderSelection();
    }
}

#pragma endregion


#pragma region Event Handling -- triggered in beNotified() in MultiReplace.cpp

void MultiReplace::processTextChange(SCNotification* notifyCode) {
    if (!isLoggingEnabled) {
        return;
    }

    Sci_Position cursorPosition = notifyCode->position;
    Sci_Position addedLines = notifyCode->linesAdded;
    Sci_Position notifyLength = notifyCode->length;

    Sci_Position lineNumber = ::SendMessage(getScintillaHandle(), SCI_LINEFROMPOSITION, cursorPosition, 0);
    if (notifyCode->modificationType & SC_MOD_INSERTTEXT) {
        if (addedLines != 0) {
            logChanges.push_back({ ChangeType::Modify, lineNumber });
            Sci_Position lineStart = ::SendMessage(getScintillaHandle(), SCI_POSITIONFROMLINE, lineNumber, 0);
            Sci_Position insertLineNumber = (cursorPosition == lineStart) ? lineNumber : lineNumber + 1;
            LogEntry insertEntry;
            insertEntry.changeType = ChangeType::Insert;
            insertEntry.lineNumber = insertLineNumber;  // ← FIXED: Use adjusted position
            insertEntry.blockSize = static_cast<Sci_Position>(std::abs(addedLines));
            logChanges.push_back(insertEntry);
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (logChanges.empty() || logChanges.back().changeType != ChangeType::Modify || logChanges.back().lineNumber != lineNumber) {
                logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
    else if (notifyCode->modificationType & SC_MOD_DELETETEXT) {
        if (addedLines != 0) {
            // Special handling for deletions at position 0
            if (cursorPosition == 0 && notifyLength == 0) {
                logChanges.push_back({ ChangeType::Delete, 0 });
                return;
            }

            // Determine where deletion actually occurred
            Sci_Position lineStart = ::SendMessage(getScintillaHandle(), SCI_POSITIONFROMLINE, lineNumber, 0);
            Sci_Position deletePos = (cursorPosition == lineStart) ? lineNumber : lineNumber + 1;

            logChanges.push_back({ ChangeType::Modify, lineNumber });

            LogEntry deleteEntry;
            deleteEntry.changeType = ChangeType::Delete;
            deleteEntry.lineNumber = deletePos;  // ← FIXED
            deleteEntry.blockSize = static_cast<Sci_Position>(std::abs(addedLines));
            logChanges.push_back(deleteEntry);
        }
        else {
            // Check if the last entry is a Modify on the same line
            if (logChanges.empty() || logChanges.back().changeType != ChangeType::Modify || logChanges.back().lineNumber != lineNumber) {
                logChanges.push_back({ ChangeType::Modify, lineNumber });
            }
        }
    }
}

void MultiReplace::processLog() {
    // displayLogChangesInMessageBox();
    instance->handleDelimiterPositions(DelimiterOperation::Update);
}

// Resolve (view, index) from BufferID. Returns false if not found in any view.
static bool GetViewIndexFromBufferId(BufferId bufId, int& viewOut, int& indexOut)
{
    // Try MAIN view first, then SUB view.
    int pos = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETPOSFROMBUFFERID, (WPARAM)bufId, MAIN_VIEW));
    if (pos < 0)
        pos = static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETPOSFROMBUFFERID, (WPARAM)bufId, SUB_VIEW));
    if (pos < 0) return false;

    viewOut = (pos >> 30) & 0x3;         // top 2 bits
    indexOut = (pos & 0x3FFFFFFF);        // low 30 bits
    return true;
}

// Post (not Send) an activation to avoid reentrancy freezes inside onDocumentSwitched().
static void PostActivateBufferId(BufferId bufId)
{
    int view, index;
    if (!GetViewIndexFromBufferId(bufId, view, index)) return;
    ::PostMessage(nppData._nppHandle, NPPM_ACTIVATEDOC, static_cast<WPARAM>(view), static_cast<LPARAM>(index));
}

void MultiReplace::onDocumentSwitched()
{
    MultiReplace* self = instance;
    if (!self || !self->isWindowOpen) return;

    self->pointerToScintilla();
    if (!self->_hScintilla) return;
    HWND hSci = self->_hScintilla;

    const BufferId currBufId =
        (BufferId)::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

    // Seed prev at first opportunity (e.g., first call after panel shows)
    if (g_prevBufId == 0)
        g_prevBufId = currBufId;

    // PHASE A: We posted an activation to the *previous* buffer, and it is now active.
    // Clean *here*, because now the leaving document is really active.
    if (g_cleanInProgress && g_pendingCleanId == currBufId)
    {
        ColumnTabs::CT_SetIndicatorId(30);

        const BOOL ro = (BOOL)::SendMessage(hSci, SCI_GETREADONLY, 0, 0);
        if (!ro) {
            // Collect padding ranges BEFORE removal for ResultDock position adjustment
            std::vector<std::pair<Sci_Position, Sci_Position>> paddingRanges;
            std::string cleanFileUtf8;
            bool hasResultHits = false;

            if (ColumnTabs::CT_HasAlignedPadding(hSci)) {
                wchar_t pathBuf[MAX_PATH] = {};
                ::SendMessage(nppData._nppHandle, NPPM_GETFULLCURRENTPATH, MAX_PATH, reinterpret_cast<LPARAM>(pathBuf));
                cleanFileUtf8 = Encoding::wstringToUtf8(pathBuf);
                hasResultHits = ResultDock::instance().hasHitsForFile(cleanFileUtf8);

                if (hasResultHits) {
                    const int ctInd = ColumnTabs::CT_GetIndicatorId();
                    const Sci_Position docLen = static_cast<Sci_Position>(::SendMessage(hSci, SCI_GETLENGTH, 0, 0));
                    Sci_Position scanPos = 0;
                    while (scanPos < docLen) {
                        const Sci_Position end = static_cast<Sci_Position>(::SendMessage(hSci, SCI_INDICATOREND, ctInd, scanPos));
                        if (end <= scanPos) break;
                        if (static_cast<int>(::SendMessage(hSci, SCI_INDICATORVALUEAT, ctInd, scanPos)) != 0) {
                            const Sci_Position start = static_cast<Sci_Position>(::SendMessage(hSci, SCI_INDICATORSTART, ctInd, scanPos));
                            if (end > start)
                                paddingRanges.emplace_back(start, end);
                        }
                        scanPos = end;
                    }
                }
            }

            ColumnTabs::CT_RemoveAlignedPadding(hSci); // removes only our marked ranges (or fallback)
            ColumnTabs::CT_SetCurDocHasPads(hSci, false);

            ColumnTabs::CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);

            // Adjust ResultDock hit positions after padding removal
            if (hasResultHits && !paddingRanges.empty()) {
                ResultDock::instance().adjustHitPositionsForFlowTab(cleanFileUtf8, paddingRanges, /*added=*/false);
            }
        }
        // Update O(1) gate: cleaned
        g_padBufs.erase(currBufId);

        // Now go back to the originally requested target buffer, asynchronously.
        const BufferId backId = g_returnBufId;
        g_pendingCleanId = 0;
        g_cleanInProgress = false;
        PostActivateBufferId(backId);
        return; // stop here; arrival on target continues in the next call
    }

    // PHASE B: Natural user switch (we just arrived on a target buffer).
    // If the previous buffer had pads, schedule an async hop back to clean it.
    const bool differentDoc = (g_prevBufId != 0 && g_prevBufId != currBufId);
    const bool prevHasPads = (g_padBufs.find(g_prevBufId) != g_padBufs.end());

    if (!g_cleanInProgress && differentDoc && prevHasPads)
    {
        // Remember where to return after cleaning; then hop back asynchronously.
        g_returnBufId = currBufId;
        g_pendingCleanId = g_prevBufId;
        g_cleanInProgress = true;
        PostActivateBufferId(g_prevBufId);
        return; // we will clean in the next onDocumentSwitched() (when prev becomes active)
    }

    // From here: normal "arrival" housekeeping for the active doc (cheap; no text scans)
    ColumnTabs::CT_DisableFlowTabStops(hSci, /*restoreManual=*/false);
    ColumnTabs::CT_ResetFlowVisualState();

    const int currentBufferID =
        static_cast<int>(::SendMessage(nppData._nppHandle, NPPM_GETCURRENTBUFFERID, 0, 0));

    if (currentBufferID == self->scannedDelimiterBufferID) {
        g_prevBufId = currBufId; // remember for next natural switch
        return;
    }

    self->documentSwitched = true;
    self->isCaretPositionEnabled = false;
    self->scannedDelimiterBufferID = currentBufferID;

    self->handleClearColumnMarks();
    self->isColumnHighlighted = false;

    self->_flowTabsActive = false;
    if (HWND h = ::GetDlgItem(self->_hSelf, IDC_COLUMN_GRIDTABS_BUTTON))
        ::SetWindowText(h, L"⇥");

    self->showStatusMessage(L"", MessageStatus::Info);

    self->m_selectionScope.clear();

    self->originalLineOrder.clear();
    self->currentSortState = SortDirection::Unsorted;
    self->isSortedColumn = false;
    self->UpdateSortButtonSymbols();

    // This buffer is now the "previous" for the *next* user switch
    g_prevBufId = currBufId;
}

void MultiReplace::pointerToScintilla() {

    int which = -1;
    ::SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, reinterpret_cast<LPARAM>(&which));
    if (which != -1) {
        instance->_hScintilla = (which == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
        s_hScintilla = instance->_hScintilla;
    }

    if (instance->_hScintilla) { // just to supress Warning
        instance->pSciMsg = (SciFnDirect)::SendMessage(instance->_hScintilla, SCI_GETDIRECTFUNCTION, 0, 0);
        instance->pSciWndData = (sptr_t)::SendMessage(instance->_hScintilla, SCI_GETDIRECTPOINTER, 0, 0);
    }
}

void MultiReplace::onSelectionChanged()
{
    static bool wasTextSelected = false;

    HWND hDlg = getDialogHandle();

    // -----------------------------------------------------------------------
    // 1) "Replace in Files" mode:
    //    - Selection-Radio is useless here → always disable and switch
    // -----------------------------------------------------------------------
    if (instance && (instance->isReplaceInFiles || instance->isFindAllInFiles))
    {
        HWND hSel = ::GetDlgItem(hDlg, IDC_SELECTION_RADIO);
        ::EnableWindow(hSel, FALSE);

        if (::SendMessage(hSel, BM_GETCHECK, 0, 0) == BST_CHECKED)
        {
            ::CheckRadioButton(
                hDlg,
                IDC_ALL_TEXT_RADIO,
                IDC_COLUMN_MODE_RADIO,
                IDC_ALL_TEXT_RADIO
            );
        }
        return;
    }

    // -----------------------------------------------------------------------
    // 2) Normal modes: No auto-disable, no auto-switch
    // -----------------------------------------------------------------------
    Sci_Position start = ::SendMessage(getScintillaHandle(), SCI_GETSELECTIONSTART, 0, 0);
    Sci_Position end = ::SendMessage(getScintillaHandle(), SCI_GETSELECTIONEND, 0, 0);
    bool isTextSelected = (start != end);

    // Inform other UI parts when selection is lost
    if (wasTextSelected && !isTextSelected)
    {
        if (instance)
            instance->setUIElementVisibility();
    }
    wasTextSelected = isTextSelected;
}

void MultiReplace::onTextChanged() {
    textModified = true;
}

void MultiReplace::onCaretPositionChanged()
{
    if (!isCaretPositionEnabled) {
        return;
    }

    LRESULT startPosition = ::SendMessage(getScintillaHandle(), SCI_GETCURRENTPOS, 0, 0);
    if (instance != nullptr) {
        // isTransient=true: This is a transient position update, don't disable position tracking
        instance->showStatusMessage(LanguageManager::instance().get(L"status_actual_position", { instance->addLineAndColumnMessage(startPosition) }), MessageStatus::Success, false, true);
    }

}

void MultiReplace::onThemeChanged()
{
    if (instance) {
        instance->applyThemePalette();
        instance->refreshColumnStylesIfNeeded();
        ResultDock::instance().onThemeChanged();
        instance->updateTextMarkerStyles();
    }
}

void MultiReplace::signalShutdown() {
    // This static method can be called from the global beNotified function
    if (instance) {
        instance->_isShuttingDown = true;
        instance->_isCancelRequested = true; // Also trigger the cancel flag
    }
}

#pragma endregion


#pragma region Debug DPI Information

BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData) {
    (void)hdcMonitor;  // Mark hdcMonitor as unused
    (void)lprcMonitor; // Mark lprcMonitor as unused

    MonitorEnumData* pData = reinterpret_cast<MonitorEnumData*>(dwData);

    MONITORINFOEX monitorInfoEx;
    monitorInfoEx.cbSize = sizeof(MONITORINFOEX);
    GetMonitorInfo(hMonitor, &monitorInfoEx);

    int screenWidth = monitorInfoEx.rcMonitor.right - monitorInfoEx.rcMonitor.left;
    int screenHeight = monitorInfoEx.rcMonitor.bottom - monitorInfoEx.rcMonitor.top;
    BOOL isPrimary = (monitorInfoEx.dwFlags & MONITORINFOF_PRIMARY) != 0;

    // Add monitor info to the buffer
    pData->monitorInfo += L"Monitor " + std::to_wstring(pData->monitorCount + 1) + L": "
        + (isPrimary ? L"Primary, " : L"Secondary, ")
        + std::to_wstring(screenWidth) + L"x" + std::to_wstring(screenHeight) + L"\n";

    // Increment the monitor counter
    pData->monitorCount++;

    // Check if this is the primary monitor
    if (isPrimary) {
        pData->primaryMonitorIndex = pData->monitorCount;
    }

    // Check if this is the current monitor where the window is
    if (MonitorFromWindow(GetForegroundWindow(), MONITOR_DEFAULTTONEAREST) == hMonitor) {
        pData->currentMonitor = pData->monitorCount;
    }

    return TRUE;  // Continue enumerating monitors
}

void MultiReplace::showDPIAndFontInfo()
{
    RECT sizeWindowRect;
    GetClientRect(_hSelf, &sizeWindowRect);  // Get window dimensions

    HDC hDC = GetDC(_hSelf);  // Get the device context (DC) for the current window
    if (hDC)
    {
        // Get current font of the window
        HFONT currentFont = reinterpret_cast<HFONT>(SendMessage(_hSelf, WM_GETFONT, 0, 0));
        SelectObject(hDC, currentFont);  // Select current font into the DC

        TEXTMETRIC tmCurrent;
        GetTextMetrics(hDC, &tmCurrent);  // Retrieve text metrics for current font

        SelectObject(hDC, font(FontRole::Standard));  // Select standard font into the DC
        TEXTMETRIC tmStandard;
        GetTextMetrics(hDC, &tmStandard);  // Retrieve text metrics for standard font

        SIZE size;
        GetTextExtentPoint32W(hDC, L"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", 52, &size);
        int baseUnitX = (size.cx / 26 + 1) / 2;
        int baseUnitY = tmCurrent.tmHeight;
        int duWidth = MulDiv(sizeWindowRect.right, 4, baseUnitX);
        int duHeight = MulDiv(sizeWindowRect.bottom, 8, baseUnitY);

        // Get DPI values from the DPIManager
        int dpix = dpiMgr->getDPIX();
        int dpiy = dpiMgr->getDPIY();
        float customScaleFactor = dpiMgr->getCustomScaleFactor();
        int scaledDpiX = dpiMgr->scaleX(96);
        int scaledDpiY = dpiMgr->scaleY(96);

        wchar_t scaleBuffer[10];
        swprintf(scaleBuffer, 10, L"%.1f", customScaleFactor);

        MonitorEnumData monitorData = {};
        monitorData.monitorCount = 0;
        monitorData.currentMonitor = 0;
        monitorData.primaryMonitorIndex = 0;

        EnumDisplayMonitors(nullptr, nullptr, MonitorEnumProc, reinterpret_cast<LPARAM>(&monitorData));

        std::wstring message =
            L"On Monitor " + std::to_wstring(monitorData.currentMonitor) + L"\n"

            + monitorData.monitorInfo + L"\n"

            L"Window Size DUs: " + std::to_wstring(duWidth) + L"x" + std::to_wstring(duHeight) + L"\n"
            L"Scaled DPI: " + std::to_wstring(dpix) + L"x" + std::to_wstring(dpiy) + L" * " + scaleBuffer + L" = "
            + std::to_wstring(scaledDpiX) + L"x" + std::to_wstring(scaledDpiY) + L"\n\n"

            L"Windows Font: Height=" + std::to_wstring(tmCurrent.tmHeight) + L", Ascent=" + std::to_wstring(tmCurrent.tmAscent) +
            L", Descent=" + std::to_wstring(tmCurrent.tmDescent) + L", Weight=" + std::to_wstring(tmCurrent.tmWeight) + L"\n"
            L"Plugin Font: Height=" + std::to_wstring(tmStandard.tmHeight) + L", Ascent=" + std::to_wstring(tmStandard.tmAscent) +
            L", Descent=" + std::to_wstring(tmStandard.tmDescent) + L", Weight=" + std::to_wstring(tmStandard.tmWeight);

        MessageBox(_hSelf, message.c_str(), L"Window, Monitor, DPI, and Font Info", MB_ICONINFORMATION | MB_OK);

        ReleaseDC(_hSelf, hDC);
    }
    else
    {
        MessageBox(_hSelf, L"Failed to retrieve device context (HDC).", L"Error", MB_OK);
    }
}

#pragma endregion