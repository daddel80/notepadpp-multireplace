//this file is part of notepad++
//Copyright (C)2023 Thomas Knoefel
//
//This program is free software; you can redistribute it and/or
//modify it under the terms of the GNU General Public License
//as published by the Free Software Foundation; either
//version 2 of the License, or (at your option) any later version.
//
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License for more details.
//
//You should have received a copy of the GNU General Public License
//along with this program; if not, write to the Free Software
//Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.

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
#define IDC_REPLACE_ALL_BUTTON          5005
#define IDC_REPLACE_DROPDOWN_BUTTON     5006
#define IDC_REPLACE_BUTTON              5007
#define IDC_REPLACE_ALL_SMALL_BUTTON    5008
#define IDC_2_BUTTONS_MODE              5009
#define IDC_FIND_BUTTON                 5010
#define IDC_FIND_NEXT_BUTTON            5011
#define IDC_FIND_PREV_BUTTON            5012
#define IDC_MARK_BUTTON                 5013
#define IDC_MARK_MATCHES_BUTTON         5014
#define IDC_CLEAR_MARKS_BUTTON          5015
#define IDC_COPY_MARKED_TEXT_BUTTON     5017
#define IDC_REPLACE_ALL_IN_LIST_BUTTON  5018
#define IDC_DELETE_REPLACE_ITEM_BUTTON  5019
#define IDC_LOAD_FROM_CSV_BUTTON        5020
#define IDC_SAVE_TO_CSV_BUTTON          5021
#define IDC_EXPORT_BASH_BUTTON			5022
#define IDC_UP_BUTTON					5023
#define IDC_DOWN_BUTTON					5024
#define IDC_SHIFT_FRAME					5025
#define IDC_SHIFT_TEXT					5026
#define ID_REPLACE_ALL_OPTION           5027
#define ID_REPLACE_IN_ALL_DOCS_OPTION   5028
#define IDC_USE_VARIABLES_HELP          5029

#define IDC_STATIC_FIND                 5100
#define IDC_STATIC_REPLACE              5101
#define IDC_STATUS_MESSAGE				5102

#define IDC_WHOLE_WORD_CHECKBOX         5200
#define IDC_MATCH_CASE_CHECKBOX         5201
#define IDC_USE_VARIABLES_CHECKBOX      5202
#define IDC_REPLACE_FIRST_CHECKBOX      5203
#define IDC_WRAP_AROUND_CHECKBOX        5204

#define IDC_SEARCH_MODE_GROUP           5300
#define IDC_NORMAL_RADIO                5301
#define IDC_EXTENDED_RADIO              5302
#define IDC_REGEX_RADIO                 5303

#define IDC_SCOPE_GROUP                 5451
#define IDC_ALL_TEXT_RADIO              5452
#define IDC_SELECTION_RADIO             5453
#define IDC_COLUMN_MODE_RADIO           5454
#define IDC_COLUMN_NUM_EDIT             5455
#define IDC_DELIMITER_EDIT              5456
#define IDC_QUOTECHAR_EDIT              5457
#define IDC_COLUMN_HIGHLIGHT_BUTTON     5458
#define IDC_DELIMITER_STATIC            5459
#define IDC_COLUMN_NUM_STATIC           5460
#define IDC_QUOTECHAR_STATIC            5461

#define IDC_STATIC_FRAME                5501
#define IDC_USE_LIST_CHECKBOX			5502
#define IDC_REPLACE_LIST                5503

#define IDD_ABOUT_DIALOG                5600
#define IDC_NAME_STATIC	                5601
#define IDC_WEBSITE_LINK                5602
#define IDC_VERSION_STATIC              5603
#define IDC_LICENSE_STATIC              5604
#define IDC_AUTHOR_STATIC               5605

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

// Minimum window dimensions
#define MIN_WIDTH 993
#define MIN_HEIGHT 485

#define IDC_WEBSITE_LINK_VALUE TEXT("https://github.com/daddel80/notepadpp-multireplace")

#endif // RESOURCE_H

