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


#ifndef IDC_STATIC
#define IDC_STATIC	-1
#endif

#define IDD_REPLACE_DIALOG              5000
#define IDC_FIND_EDIT                   5001
#define IDC_REPLACE_EDIT                5002
#define IDC_SWAP_BUTTON					5003
#define IDC_COPY_TO_LIST_BUTTON         5004
#define IDC_USE_LIST_CHECKBOX			5005
#define IDC_REPLACE_ALL_BUTTON          5006
#define IDC_REPLACE_DROPDOWN_BUTTON     5007
#define IDC_REPLACE_BUTTON              5008
#define IDC_REPLACE_ALL_SMALL_BUTTON    5009
#define IDC_2_BUTTONS_MODE              5010
#define IDC_FIND_BUTTON                 5011
#define IDC_FIND_PREV_BUTTON            5012
#define IDC_FIND_NEXT_BUTTON            5013
#define IDC_MARK_BUTTON                 5014
#define IDC_MARK_MATCHES_BUTTON         5015
#define IDC_COPY_MARKED_TEXT_BUTTON     5016
#define IDC_CLEAR_MARKS_BUTTON          5017
#define IDC_CANCEL_REPLACE_BUTTON       5018
#define IDC_LOAD_FROM_CSV_BUTTON        5020
#define IDC_LOAD_LIST_BUTTON            5021
#define IDC_NEW_LIST_BUTTON             5022
#define IDC_SAVE_TO_CSV_BUTTON          5023
#define IDC_SAVE_BUTTON                 5024
#define IDC_SAVE_AS_BUTTON              5025
#define IDC_EXPORT_BASH_BUTTON			5026
#define IDC_UP_BUTTON					5027
#define IDC_DOWN_BUTTON					5028
#define IDC_SHIFT_FRAME					5029
#define IDC_SHIFT_TEXT					5030
#define ID_REPLACE_ALL_OPTION           5031
#define ID_REPLACE_IN_ALL_DOCS_OPTION   5032
#define ID_REPLACE_IN_FILES_OPTION      5033
#define IDC_USE_LIST_BUTTON				5034

#define IDC_STATIC_FIND                 5100
#define IDC_STATIC_REPLACE              5101
#define IDC_STATUS_MESSAGE				5102
#define IDC_PATH_DISPLAY                5103
#define IDC_STATS_DISPLAY               5104

#define IDC_WHOLE_WORD_CHECKBOX         5200
#define IDC_MATCH_CASE_CHECKBOX         5201
#define IDC_USE_VARIABLES_CHECKBOX      5202
#define IDC_USE_VARIABLES_HELP          5203
#define IDC_WRAP_AROUND_CHECKBOX        5204
#define IDC_REPLACE_AT_MATCHES_CHECKBOX 5205
#define IDC_REPLACE_HIT_EDIT            5206

#define IDC_SEARCH_MODE_GROUP           5300
#define IDC_NORMAL_RADIO                5301
#define IDC_EXTENDED_RADIO              5302
#define IDC_REGEX_RADIO                 5303

#define IDC_SCOPE_GROUP                 5451
#define IDC_ALL_TEXT_RADIO              5452
#define IDC_SELECTION_RADIO             5453
#define IDC_COLUMN_MODE_RADIO           5454
#define IDC_COLUMN_SORT_DESC_BUTTON     5455
#define IDC_COLUMN_SORT_ASC_BUTTON      5456
#define IDC_COLUMN_DROP_BUTTON          5457
#define IDC_COLUMN_COPY_BUTTON          5458
#define IDC_COLUMN_HIGHLIGHT_BUTTON     5459
#define IDC_COLUMN_NUM_EDIT             5460
#define IDC_DELIMITER_EDIT              5461
#define IDC_QUOTECHAR_EDIT              5462
#define IDC_DELIMITER_STATIC            5463
#define IDC_COLUMN_NUM_STATIC           5464
#define IDC_QUOTECHAR_STATIC            5465

#define IDC_REPLACE_IN_FILES_GROUP      5469
#define IDC_DIR_STATIC                  5470
#define IDC_DIR_EDIT                    5471
#define IDC_BROWSE_DIR_BUTTON           5472
#define IDC_FILTER_STATIC               5473
#define IDC_FILTER_EDIT                 5474
#define IDC_FILTER_HELP                 5475
#define IDC_SUBFOLDERS_CHECKBOX         5476
#define IDC_HIDDENFILES_CHECKBOX        5477
#define IDC_DIR_PROGRESS_STATIC         5478
#define IDC_DIR_PROGRESS_BAR            5479

#define IDC_STATIC_FRAME                5501
#define IDC_REPLACE_LIST                5502

#define IDD_ABOUT_DIALOG                5600
#define IDC_NAME_STATIC	                5601
#define IDC_WEBSITE_LINK                5602
#define IDC_VERSION_STATIC              5603
#define IDC_LICENSE_STATIC              5604
#define IDC_AUTHOR_STATIC               5605
#define IDC_VERSION_LABEL               5606
#define IDC_AUTHOR_LABEL                5607
#define IDC_LICENSE_LABEL               5608

#define IDM_UNDO                        5701
#define IDM_REDO                        5702
#define IDM_COPY_DATA_TO_FIELDS         5703
#define IDM_SEARCH_IN_LIST              5704
#define IDM_CUT_LINES_TO_CLIPBOARD      5705
#define IDM_ADD_NEW_LINE                5706
#define IDM_COPY_LINES_TO_CLIPBOARD     5707
#define IDM_PASTE_LINES_FROM_CLIPBOARD  5708
#define IDM_EDIT_VALUE                  5709
#define IDM_DELETE_LINES                5710
#define IDM_SELECT_ALL                  5711
#define IDM_ENABLE_LINES                5712
#define IDM_DISABLE_LINES               5713

#define IDM_TOGGLE_FIND_COUNT           5801
#define IDM_TOGGLE_REPLACE_COUNT        5802
#define IDM_TOGGLE_COMMENTS             5803
#define IDM_TOGGLE_DELETE               5804

#define ID_EDIT_EXPAND_BUTTON           6000

#define STYLE1							60
#define STYLE2							61
#define STYLE3							62
#define STYLE4							63
#define STYLE5							64
#define STYLE6							65
#define STYLE7							66
#define STYLE8							67
#define STYLE9							68
#define STYLE10							69

#define IDI_MR_ICON 5801
#define IDI_MR_DM_ICON 5802

// Default window position and dimensions
#define POS_X 92
#define POS_Y 40
#define MIN_WIDTH 753
#define MIN_HEIGHT 400
#define SHRUNK_HEIGHT 224

// Custom message used to perform initial actions after the window has been fully opened
#define WM_POST_INIT (WM_APP + 1)

#define IDC_WEBSITE_LINK_VALUE TEXT("https://github.com/daddel80/notepadpp-multireplace/issues")

#endif // RESOURCE_H

