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

#ifndef RESOURCE_H
#define RESOURCE_H

// -------------------------------------------------------------------
//  Generic helper for Win32 controls
// -------------------------------------------------------------------
#ifndef IDC_STATIC
#define IDC_STATIC (-1)
#endif



// ===================================================================
//  5000-block - Dialogs & base widgets
// ===================================================================
#define IDD_REPLACE_DIALOG              5000   // main modeless dialog


// -------------------------------------------------------------------
//  5001-5005 - Edit controls & small helpers
// -------------------------------------------------------------------
#define IDC_FIND_EDIT                   5001
#define IDC_REPLACE_EDIT                5002
#define IDC_SWAP_BUTTON                 5003
#define IDC_COPY_TO_LIST_BUTTON         5004
#define IDC_USE_LIST_CHECKBOX           5005



// ===================================================================
//  5006-5018 - Primary action buttons
// ===================================================================
#define IDC_REPLACE_ALL_BUTTON          5006   // split button (Replace ...)
#define IDC_REPLACE_DROPDOWN_BUTTON     5007   // part of the split button
#define IDC_REPLACE_BUTTON              5008   // single "Replace" action
#define IDC_REPLACE_ALL_SMALL_BUTTON    5009   // icon button
#define IDC_2_BUTTONS_MODE              5010   // checkbox: two-button UI

#define IDC_FIND_ALL_BUTTON             5011
#define IDC_FIND_PREV_BUTTON            5012
#define IDC_FIND_NEXT_BUTTON            5013

#define IDC_MARK_BUTTON                 5014
#define IDC_MARK_MATCHES_BUTTON         5015
#define IDC_COPY_MARKED_TEXT_BUTTON     5016
#define IDC_CLEAR_MARKS_BUTTON          5017
#define IDC_CANCEL_REPLACE_BUTTON       5018



// ===================================================================
//  5020-5028 - File & list management
// ===================================================================
#define IDC_LOAD_FROM_CSV_BUTTON        5020
#define IDC_LOAD_LIST_BUTTON            5021
#define IDC_NEW_LIST_BUTTON             5022
#define IDC_SAVE_TO_CSV_BUTTON          5023
#define IDC_SAVE_BUTTON                 5024
#define IDC_SAVE_AS_BUTTON              5025
#define IDC_EXPORT_BASH_BUTTON          5026

#define IDC_UP_BUTTON                   5027
#define IDC_DOWN_BUTTON                 5028



// ===================================================================
//  5031-5040 - Split-button option command IDs
// ===================================================================
//  Replace-All split options
#define ID_REPLACE_ALL_OPTION           5031
#define ID_REPLACE_IN_ALL_DOCS_OPTION   5032
#define ID_REPLACE_IN_FILES_OPTION      5033
#define ID_DEBUG_MODE_OPTION            5034

//  Stand-alone list toggle button
#define IDC_USE_LIST_BUTTON             5040



// ===================================================================
//  5100-5104 - Static texts & status
// ===================================================================
#define IDC_STATIC_FIND                 5100
#define IDC_STATIC_REPLACE              5101
#define IDC_STATUS_MESSAGE              5102
#define IDC_PATH_DISPLAY                5103
#define IDC_STATS_DISPLAY               5104



// ===================================================================
//  5200-5206 - Checkboxes
// ===================================================================
#define IDC_WHOLE_WORD_CHECKBOX         5200
#define IDC_MATCH_CASE_CHECKBOX         5201
#define IDC_USE_VARIABLES_CHECKBOX      5202
#define IDC_USE_VARIABLES_HELP          5203
#define IDC_WRAP_AROUND_CHECKBOX        5204
#define IDC_REPLACE_AT_MATCHES_CHECKBOX 5205
#define IDC_REPLACE_HIT_EDIT            5206



// ===================================================================
//  5300-5303 - Search-mode radio buttons
// ===================================================================
#define IDC_SEARCH_MODE_GROUP           5300
#define IDC_NORMAL_RADIO                5301
#define IDC_EXTENDED_RADIO              5302
#define IDC_REGEX_RADIO                 5303



// ===================================================================
//  5450-5479 - Column / scope group & Replace-in-Files
// ===================================================================

// Scope group & radio buttons (5450-5454)
#define IDC_SCOPE_GROUP                 5450
#define IDC_ALL_TEXT_RADIO              5451
#define IDC_SELECTION_RADIO             5452
#define IDC_COLUMN_MODE_RADIO           5453

// CSV toolbar buttons (5455-5461)
#define IDC_COLUMN_SORT_DESC_BUTTON     5455
#define IDC_COLUMN_SORT_ASC_BUTTON      5456
#define IDC_COLUMN_DROP_BUTTON          5457
#define IDC_COLUMN_COPY_BUTTON          5458
#define IDC_COLUMN_HIGHLIGHT_BUTTON     5459
#define IDC_COLUMN_GRIDTABS_BUTTON      5460
#define IDC_COLUMN_DUPLICATES_BUTTON    5461

// Column edit controls (5462-5467)
#define IDC_COLUMN_NUM_EDIT             5462
#define IDC_DELIMITER_EDIT              5463
#define IDC_QUOTECHAR_EDIT              5464
#define IDC_COLUMN_NUM_STATIC           5465
#define IDC_DELIMITER_STATIC            5466
#define IDC_QUOTECHAR_STATIC            5467

// Replace-in-Files frame (5469-5479)
#define IDC_FILE_OPS_GROUP              5469
#define IDC_FILTER_EDIT                 5470
#define IDC_FILTER_HELP                 5471
#define IDC_FILTER_STATIC               5472
#define IDC_DIR_STATIC                  5473
#define IDC_DIR_EDIT                    5474
#define IDC_BROWSE_DIR_BUTTON           5475
#define IDC_SUBFOLDERS_CHECKBOX         5476
#define IDC_HIDDENFILES_CHECKBOX        5477
#define IDC_DIR_PROGRESS_STATIC         5478
#define IDC_DIR_PROGRESS_BAR            5479



// ===================================================================
//  5500-5502 - List view & frame
// ===================================================================
#define IDC_STATIC_FRAME                5501
#define IDC_REPLACE_LIST                5502



// ===================================================================
//  5600-5608 - About dialog
// ===================================================================
#define IDD_ABOUT_DIALOG                5600
#define IDC_NAME_STATIC                 5601
#define IDC_WEBSITE_LINK                5602
#define IDC_VERSION_STATIC              5603
#define IDC_LICENSE_STATIC              5604
#define IDC_AUTHOR_STATIC               5605
#define IDC_VERSION_LABEL               5606
#define IDC_AUTHOR_LABEL                5607
#define IDC_LICENSE_LABEL               5608



// ===================================================================
//  5700-5728 - Context-menu command IDs & List Search
// ===================================================================
#define IDM_UNDO                        5701
#define IDM_REDO                        5702
#define IDM_CUT_LINES_TO_CLIPBOARD      5703
#define IDM_COPY_LINES_TO_CLIPBOARD     5704
#define IDM_PASTE_LINES_FROM_CLIPBOARD  5705
#define IDM_SELECT_ALL                  5706
#define IDM_EDIT_VALUE                  5707
#define IDM_DELETE_LINES                5708
#define IDM_ADD_NEW_LINE                5709
#define IDM_COPY_DATA_TO_FIELDS         5710
#define IDM_EXPORT_DATA                 5711
#define IDM_SEARCH_IN_LIST              5712
#define IDM_ENABLE_LINES                5713
#define IDM_DISABLE_LINES               5714

// Set/Clear Options Submenus
#define IDM_SET_WHOLEWORD               5715
#define IDM_SET_MATCHCASE               5716
#define IDM_SET_VARIABLES               5717
#define IDM_SET_EXTENDED                5718
#define IDM_SET_REGEX                   5719
#define IDM_CLEAR_WHOLEWORD             5720
#define IDM_CLEAR_MATCHCASE             5721
#define IDM_CLEAR_VARIABLES             5722
#define IDM_CLEAR_EXTENDED              5723
#define IDM_CLEAR_REGEX                 5724

// List Search Bar
#define IDC_LIST_SEARCH_COMBO           5726
#define IDC_LIST_SEARCH_BUTTON          5727
#define IDC_LIST_SEARCH_CLOSE           5728



// ===================================================================
//  5800-5806 - View-toggle menu items & icons
// ===================================================================
#define IDM_TOGGLE_FIND_COUNT           5801
#define IDM_TOGGLE_REPLACE_COUNT        5802
#define IDM_TOGGLE_COMMENTS             5803
#define IDM_TOGGLE_DELETE               5804

// ===================================================================
//  6000-6120 - Miscellaneous & Result Dock
// ===================================================================
#define ID_EDIT_EXPAND_BUTTON           6000

#define IDD_MULTIREPLACE_RESULT_DOCK    6100    // dock panel template
#define IDC_FINDALL_BUTTON              6105    // big split-button
#define ID_FIND_ALL_OPTION              6110
#define ID_FIND_ALL_IN_ALL_DOCS_OPTION  6115
#define ID_FIND_ALL_IN_FILES_OPTION     6120



// ===================================================================
//  Scintilla style IDs  (60 - 69)
// ===================================================================
#define STYLE1                          60
#define STYLE2                          61
#define STYLE3                          62
#define STYLE4                          63
#define STYLE5                          64
#define STYLE6                          65
#define STYLE7                          66
#define STYLE8                          67
#define STYLE9                          68
#define STYLE10                         69



// ===================================================================
//  7000-7010 - Config Dialog Base
// ===================================================================
#define IDD_MULTIREPLACE_CONFIG           7000
#define IDC_CONFIG_CATEGORY_LIST          7010



// ===================================================================
//  Config Dialog Control IDs (7900-7999 block)
//  Grouped by panel, 10-step increments between groups
//  NOTE: Windows Control IDs must be < 65536 (16-bit)
// ===================================================================

// -------------------------------------------------------------------
//  Search & Replace Panel (7900-7909)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_SEARCH_BEHAVIOUR      7900
#define IDC_CFG_STAY_AFTER_REPLACE        7901
#define IDC_CFG_ALL_FROM_CURSOR           7902
#define IDC_CFG_MUTE_SOUNDS               7903
#define IDC_CFG_LIMIT_FILESIZE            7904
#define IDC_CFG_MAX_FILESIZE_EDIT         7905
#define IDC_CFG_FILESIZE_MB_LABEL         7906

// -------------------------------------------------------------------
//  List View Panel - Columns Group (7910-7919)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_LIST_COLUMNS          7910
#define IDC_CFG_FINDCOUNT_VISIBLE         7911
#define IDC_CFG_REPLACECOUNT_VISIBLE      7912
#define IDC_CFG_COMMENTS_VISIBLE          7913
#define IDC_CFG_DELETEBUTTON_VISIBLE      7914

// -------------------------------------------------------------------
//  List View Panel - Results Group (7920-7929)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_LIST_STATS            7920
#define IDC_CFG_LISTSTATISTICS_ENABLED    7921
#define IDC_CFG_GROUPRESULTS_ENABLED      7922

// -------------------------------------------------------------------
//  List View Panel - Interaction Group (7930-7939)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_LIST_INTERACTION      7930
#define IDC_CFG_HIGHLIGHT_MATCH           7931
#define IDC_CFG_DOUBLECLICK_EDITS         7932
#define IDC_CFG_HOVER_TEXT_ENABLED        7933
#define IDC_CFG_EDITFIELD_LABEL           7934
#define IDC_CFG_EDITFIELD_SIZE_COMBO      7935

// -------------------------------------------------------------------
//  CSV Panel (7940-7949)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_CSV_SETTINGS          7940
#define IDC_CFG_FLOWTABS_NUMERIC_ALIGN    7941
#define IDC_CFG_FLOWTABS_INTRO_DONTSHOW   7942
#define IDC_CFG_CSV_SORT_LABEL            7943
#define IDC_CFG_HEADERLINES_EDIT          7944
#define IDC_CFG_DUPLICATE_BOOKMARKS       7945

// -------------------------------------------------------------------
//  Export Data Panel (7950-7959)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_EXPORT_DATA           7950
#define IDC_CFG_EXPORT_FORMAT_LABEL       7951
#define IDC_CFG_EXPORT_FORMAT_COMBO       7952
#define IDC_CFG_EXPORT_TEMPLATE_LABEL     7953
#define IDC_CFG_EXPORT_TEMPLATE_EDIT      7954
#define IDC_CFG_EXPORT_ESCAPE_CHECK       7955
#define IDC_CFG_EXPORT_HEADER_CHECK       7956
#define IDC_CFG_EXPORT_TEMPLATE_HELP      7957

// -------------------------------------------------------------------
//  Appearance Panel - Interface Group (7960-7969)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_INTERFACE             7960
#define IDC_CFG_FOREGROUND_LABEL          7961
#define IDC_CFG_FOREGROUND_SLIDER         7962
#define IDC_CFG_BACKGROUND_LABEL          7963
#define IDC_CFG_BACKGROUND_SLIDER         7964
#define IDC_CFG_SCALE_LABEL               7965
#define IDC_CFG_SCALE_SLIDER              7966

// -------------------------------------------------------------------
//  Appearance Panel - Display Options Group (7970-7979)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_DISPLAY_OPTIONS       7970
#define IDC_CFG_TOOLTIPS_ENABLED          7971
#define IDC_CFG_RESULT_DOCK_ENTRY_COLORS  7972
#define IDC_CFG_USE_LIST_COLORS_MARKING   7973

// -------------------------------------------------------------------
//  Import and Scope Panel (7980-7989)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_IMPORT_SCOPE          7980
#define IDC_CFG_SCOPE_USE_LIST            7981
#define IDC_CFG_IMPORT_ON_STARTUP         7982
#define IDC_CFG_REMEMBER_IMPORT_PATH      7983
#define IDC_CFG_IMPORT_SCOPE_PLACEHOLDER  7984

// -------------------------------------------------------------------
//  Lua Settings (7990-7999)
// -------------------------------------------------------------------
#define IDC_CFG_GRP_LUA                   7990
#define IDC_CFG_LUA_SAFEMODE_ENABLED      7991
#define IDC_CFG_LUA_PLACEHOLDER_LABEL     7992



// ===================================================================
//  Hard-coded defaults & custom messages
// ===================================================================
// --- Panel Position ---
#define CENTER_ON_NPP               -9999   // Sentinel: center over N++ on first run

// --- Main Panel: Window Size (resizable) ---
#define MIN_WIDTH                   811     // Minimum width (resize limit)
#define MIN_HEIGHT                  370     // Minimum height with list (resize limit)
#define SHRUNK_HEIGHT               224     // Minimum height without list
#define INIT_WIDTH                  MIN_WIDTH   // Initial client-area width on first run
#define INIT_HEIGHT                 MIN_HEIGHT  // Initial client-area height on first runn

// --- Config Dialog: Window Size (fixed) ---
#define CONFIG_DLG_WIDTH            810     // Fixed width
#define CONFIG_DLG_HEIGHT           380     // Fixed height

#define WM_POST_INIT        (WM_APP + 1)    // posted after the dialog is shown

#define IDC_WEBSITE_LINK_VALUE TEXT("https://github.com/daddel80/notepadpp-multireplace/issues")

#endif // RESOURCE_H
