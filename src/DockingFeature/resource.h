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
#define IDC_REPLACE_ALL_BUTTON          5003
#define IDC_MARK_MATCHES_BUTTON         5004
#define IDC_CLEAR_MARKS_BUTTON          5005
#define IDC_COPY_MARKED_TEXT_BUTTON     5006
#define IDC_COPY_TO_LIST_BUTTON         5007
#define IDC_REPLACE_ALL_IN_LIST_BUTTON  5008
#define IDC_DELETE_REPLACE_ITEM_BUTTON  5009
#define IDC_SAVE_TO_CSV_BUTTON          5010
#define IDC_LOAD_FROM_CSV_BUTTON        5011
#define IDC_UP_BUTTON					5012
#define IDC_DOWN_BUTTON					5013
#define IDC_SHIFT_FRAME					5014
#define IDC_SHIFT_TEXT					5015
#define IDC_EXPORT_BASH_BUTTON			5016

#define IDC_STATIC_FIND                 5100
#define IDC_STATIC_REPLACE              5101
#define IDC_STATIC_HINT                 5102
#define IDC_STATUS_MESSAGE				5103

#define IDC_WHOLE_WORD_CHECKBOX         5200
#define IDC_MATCH_CASE_CHECKBOX         5201
#define IDC_EXTENDED_CHECKBOX           5202
#define IDC_REGEX_CHECKBOX              5203
#define IDC_USE_LIST_CHECKBOX			5204
#define IDC_STATIC_FRAME                5205
#define IDC_TRANSPARENT_PANEL           5206

#define IDC_SEARCH_MODE_GROUP           5300
#define IDC_NORMAL_RADIO                5301
#define IDC_REGEX_RADIO                 5302
#define IDC_EXTENDED_RADIO              5303

#define IDC_REPLACE_LIST                5400

#define DELETE_ICON                     5500
#define ENABLED_ICON                    5501
#define COPYBACK_ICON                   5502

#define IDD_ABOUT_DIALOG                5600
#define IDC_MAILTO_LINK                 5601
#define IDC_WEBSITE_LINK                5602
#define IDC_VERSION_STATIC              5603
#define IDC_LICENSE_STATIC              5604
#define IDC_AUTHOR_STATIC               5605

#define VERSION_VALUE "1.0.0"
#define IDC_WEBSITE_LINK_VALUE TEXT("https://github.com/daddel80/notepadpp-multireplace")
#define IDC_MAILTO_LINK_VALUE  TEXT("mailto:tknoefel@gmail.com")

#endif // RESOURCE_H

