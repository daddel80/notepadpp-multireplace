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

#include "language_mapping.h"
#include "MultiReplacePanel.h"
#include "StaticDialog/resource.h"

const LangKV kEnglishPairs[] = {
{ L"panel_find_what", L"Find what: "},
{ L"panel_replace_with", L"Replace with: "},
{ L"panel_match_whole_word_only", L"Match whole word only" },
{ L"panel_match_case", L"Match case" },
{ L"panel_use_variables", L"Use Variables" },
{ L"panel_replace_at_matches", L"Replace matches:" },
{ L"panel_wrap_around", L"Wrap around" },
{ L"panel_search_mode", L"Search Mode" },
{ L"panel_normal", L"Normal" },
{ L"panel_extended", L"Extended (\\n, \\r, \\t, \\0, \\x...)" },
{ L"panel_regular_expression", L"Regular expression" },
{ L"panel_scope", L"Scope" },
{ L"panel_all_text", L"All Text" },
{ L"panel_selection", L"Selection" },
{ L"panel_csv", L"CSV" },
{ L"panel_cols", L"Cols:" },
{ L"panel_delim", L"Delim:" },
{ L"panel_quote", L"Quote:" },
{ L"panel_add_into_list", L"Add into List" },
{ L"panel_replace_all", L"Replace All" },
{ L"panel_replace", L"Replace" },
{ L"panel_find_all", L"Find All" },
{ L"panel_find_next", L"Find Next"},
{ L"panel_mark_matches", L"Mark Matches" },
{ L"panel_mark_matches_small", L"Mark Matches"},
{ L"panel_clear_all_marks", L"Clear all marks" },
{ L"panel_load_list", L"Load List" },
{ L"panel_save_list", L"Save List" },
{ L"panel_save_as", L"Save As..." },
{ L"panel_export_to_bash", L"Export to Bash" },
{ L"panel_help", L"?" },
{ L"panel_replace_in_files", L"Replace in Files" },
{ L"panel_find_in_files", L"Find in Files" },
{ L"panel_find_replace_in_files", L"Find/Replace in Files" },
{ L"panel_directory", L"Directory: " },
{ L"panel_filter", L"Filter: " },
{ L"panel_in_subfolders", L"In all sub-folders" },
{ L"panel_in_hidden_folders", L"In hidden folders" },
{ L"panel_cancel_replace", L"Cancel" },

// File Dialog
{ L"filetype_all_files", L"All Files (*.*)" },
{ L"filetype_csv", L"CSV Files (*.csv)" },
{ L"filetype_bash", L"Bash Files (*.sh)" },

// Tooltips
{ L"tooltip_replace_all", L"Replace All" },
{ L"tooltip_2_buttons_mode", L"2 buttons mode" },
{ L"tooltip_columns", L"Columns: '1,3,5-12'" },
{ L"tooltip_delimiter", L"Delimiter: Single/combined chars, \\t for Tab" },
{ L"tooltip_quote", L"Quote: ', or \" or none" },
{ L"tooltip_replace_at_matches", L"Replace at matches (Replace All only): '1,3,5-12'" },
{ L"tooltip_sort_descending", L"Sort Descending" },
{ L"tooltip_sort_ascending", L"Sort Ascending" },
{ L"tooltip_drop_columns", L"Drop Columns" },
{ L"tooltip_copy_columns", L"Copy Columns to Clipboard" },
{ L"tooltip_column_highlight", L"Column highlight: On/Off" },
{ L"tooltip_column_tabs", L"Column Alignment: On/Off" },
{ L"tooltip_copy_marked_text", L"Copy Marked Text" },
{ L"tooltip_new_list", L"New List" },
{ L"tooltip_save", L"Save List" },
{ L"tooltip_enable_list", L"Enable list" },
{ L"tooltip_disable_list", L"Disable list" },
{ L"tooltip_filter_help", L"Find in cpp, cxx, h, hxx & hpp:\n*.cpp *.cxx *.h *.hxx *.hpp\n\nFind in all files except exe, obj & log:\n*.* !*.exe !*.obj !*.log\n\nFind in all files but exclude folders\ntests, bin & bin64:\n*.* !\\tests\\ !\\bin*\n\nFind in all files but exclude all folders\nlog or logs recursively:\n*.* !+\\log*" },
{ L"tooltip_export_template_help", L"Available placeholders:\n\nData fields:\n  %FIND%      - Find pattern\n  %REPLACE%   - Replace text\n  %COMMENT%   - Comment\n  %FCOUNT%    - Find count\n  %RCOUNT%    - Replace count\n\nRow info:\n  %ROW%       - Row number\n  %SEL%       - Selected (1/0)\n\nOptions:\n  %REGEX%     - Regex enabled (1/0)\n  %CASE%      - Match case (1/0)\n  %WORD%      - Whole word (1/0)\n  %EXT%       - Extended (1/0)\n  %VAR%       - Variables (1/0)\n\nUse \\t for tab delimiter" },
{ L"tooltip_move_up", L"Move selected lines up" },
{ L"tooltip_move_down", L"Move selected lines down" },

// header entries
{ L"header_find_count", L"Find Count" },
{ L"header_replace_count", L"Replace Count" },
{ L"header_find", L"Find" },
{ L"header_replace", L"Replace" },
{ L"header_comments", L"Comments" },
{ L"header_delete_button", L"Delete Entry" },
{ L"header_whole_word", L"W" },
{ L"header_match_case", L"C" },
{ L"header_use_variables", L"V" },
{ L"header_extended", L"E" },
{ L"header_regex", L"R" },

// tooltip entries
{ L"tooltip_header_whole_word", L"Whole Word" },
{ L"tooltip_header_match_case", L"Match Case" },
{ L"tooltip_header_use_variables", L"Use Variables" },
{ L"tooltip_header_extended", L"Extended" },
{ L"tooltip_header_regex", L"Regex" },

// SplitButton entries
{ L"split_menu_replace_all", L"Replace All" },
{ L"split_menu_replace_all_in_docs", L"Replace All in All opened Documents" },
{ L"split_menu_replace_all_in_files", L"Replace All in Files" },
{ L"split_menu_debug_mode", L"Debug Mode" },
{ L"split_button_replace_all", L"Replace All" },
{ L"split_button_replace_all_in_docs", L"Replace All in Docs" },
{ L"split_button_replace_all_in_files", L"Replace in Files" },
{ L"split_menu_find_all", L"Find All" },
{ L"split_menu_find_all_in_docs", L"Find All in All opened Documents" },
{ L"split_menu_find_all_in_files", L"Find All in Files" },
{ L"split_button_find_all", L"Find All" },
{ L"split_button_find_all_in_docs", L"Find All in Docs" },
{ L"split_button_find_all_in_files", L"Find in Files" },

// Static Status message entries
{ L"status_duplicate_entry", L"Duplicate entry: " },
{ L"status_value_added", L"Value added to the list." },
{ L"status_no_rows_selected", L"No rows selected to shift." },
{ L"status_one_line_deleted", L"1 line deleted." },
{ L"status_column_marks_cleared", L"Column marks cleared." },
{ L"status_all_marks_cleared", L"All marks cleared." },
{ L"status_cannot_replace_read_only", L"Cannot replace. Document is read-only." },
{ L"status_add_values_instructions", L"Add values into the list. Or disable the list to replace directly." },
{ L"status_no_find_string", L"No 'Find String' entered. Please provide a value to add to the list." },
{ L"status_no_rows_selected_to_shift", L"No rows selected to shift." },
{ L"status_add_values_or_uncheck", L"Add values into the list or disable the list." },
{ L"status_no_occurrence_found", L"No occurrence found." },
{ L"status_found_text_not_replaced", L"Found text was not replaced." },
{ L"status_replace_one_next_found", L"Replace: 1 occurrence replaced. Next found." },
{ L"status_replace_one_none_left", L"Replace: 1 occurrence replaced. None left." },
{ L"status_replace_one", L"Replace: 1 occurrence replaced." },
{ L"status_add_values_or_find_directly", L"Add values into the list. Or disable the list to find directly." },
{ L"status_wrapped", L"Wrapped" },
{ L"status_no_matches_found", L"No matches found." },
{ L"status_add_values_or_mark_directly", L"Add values into the list. Or disable the list to mark directly." },
{ L"status_no_text_to_copy", L"No text to copy." },
{ L"status_failed_to_copy", L"Failed to copy to Clipboard." },
{ L"status_failed_allocate_memory", L"Failed to allocate memory for Clipboard." },
{ L"status_invalid_column_or_delimiter", L"Invalid column or delimiter data." },
{ L"status_missing_column_or_delimiter_data", L"Column data or delimiter data is missing" },
{ L"status_invalid_range_in_column_data", L"Invalid range in column data" },
{ L"status_missing_match_selection", L"Match selection data is missing" },
{ L"status_invalid_range_in_match_data", L"Invalid range in match selection" },
{ L"status_extended_delimiter_empty", L"Extended delimiter is empty" },
{ L"status_invalid_quote_character", L"Invalid quote character. Use \" or ' or leave it empty." },
{ L"status_unable_to_save_file", L"Error: Unable to open or write to file." },
{ L"status_saved_items_to_csv", L"$REPLACE_STRING items saved to CSV." },
{ L"status_no_valid_items_in_csv", L"No valid items found in the CSV file." },
{ L"status_list_exported_to_bash", L"List exported to BASH script." },
{ L"status_invalid_column_count", L"File not loaded! Invalid number of columns in CSV file." },
{ L"status_invalid_data_in_columns", L"File not loaded! Invalid data found in CSV columns." },
{ L"status_no_find_replace_list_input", L"No 'Find what' or 'Replace with' string provided. Please enter a value." },
{ L"status_found_in_list", L"Entry found in the list." },
{ L"status_not_found_in_list", L"No entry found in the list based on input fields." },
{ L"status_enable_list", L"List mode enabled. Actions will use list entries." },
{ L"status_disable_list", L"List mode disabled. Actions will use 'Find what' and 'Replace with' fields." },
{ L"status_new_list_created", L"New list created." },
{ L"status_no_rows_selected_to_delete", L"No rows selected to delete." },
{ L"status_invalid_indices", L"Invalid row indices." },
{ L"status_error_hidden_buffer", L"Error creating hidden buffer for processing." },
{ L"status_error_invalid_directory", L"The specified directory is invalid or does not exist." },
{ L"status_error_scanning_directory", L"Error scanning directory: $REPLACE_STRING" },
{ L"status_operation_cancelled", L"Replacement cancelled." },
{ L"status_replace_summary", L"Replace in files: $REPLACE_STRING1 of $REPLACE_STRING2 file(s) modified." },
{ L"status_occurrences_found", L"$REPLACE_STRING occurrences found." },
{ L"status_canceled", L"Canceled" },
{ L"status_no_delimiters", L"No delimiters found." },
{ L"status_model_build_failed", L"Flow Tabs: model build failed." },
{ L"status_padding_insert_failed", L"Flow Tabs: insert failed." },
{ L"status_visual_fail", L"Flow Tabs: visual tabstops failed." },
{ L"status_tabs_inserted", L"Flow Tabs: INSERTED." },
{ L"status_tabs_removed", L"Flow Tabs: REMOVED." },
{ L"status_tabs_aligned", L"Flow Tabs: ALIGNED." },
{ L"status_nothing_to_align", L"Flow Tabs: nothing to align." },
{ L"status_no_selection", L"No text selected." },
{ L"status_export_failed", L"Export to clipboard failed." },
{ L"status_no_find_all_results", L"No Find All results available. Run Find All first." },
{ L"status_no_matches_for_entry", L"No matches for this entry." },
{ L"status_results_cleared", L"Results have been cleared." },
{ L"status_matches_no_longer_available", L"Matches no longer available." },
{ L"status_match_position", L"Match $REPLACE_STRING1/$REPLACE_STRING2" },
{ L"status_wrapped_to_first", L"Wrapped to first match" },
{ L"status_wrapped_to_last", L"Wrapped to last match" },

// Dynamic Status Messages
{ L"status_rows_shifted", L"$REPLACE_STRING rows successfully shifted." },
{ L"status_lines_deleted", L"$REPLACE_STRING lines deleted." },
{ L"status_occurrences_replaced", L"$REPLACE_STRING occurrences were replaced." },
{ L"status_no_matches_found_for", L"No matches found for '$REPLACE_STRING'." },
{ L"status_actual_position", L"Actual Position $REPLACE_STRING" },
{ L"status_items_loaded_from_csv", L"$REPLACE_STRING items loaded from CSV." },
{ L"status_occurrences_marked", L"$REPLACE_STRING occurrences were marked." },
{ L"status_items_copied_to_clipboard", L"$REPLACE_STRING items copied to Clipboard." },
{ L"status_no_matches_after_wrap_for", L"No matches found for '$REPLACE_STRING' after wrap." },
{ L"status_deleted_fields_count", L"Deleted $REPLACE_STRING fields." },
{ L"status_line_and_column_position", L" (Line: $REPLACE_STRING1, Column: $REPLACE_STRING2)" },
{ L"status_unable_to_open_file", L"Failed to open the file: $REPLACE_STRING" },

// MessageBox Titles
{ L"msgbox_title_error", L"Error" },
{ L"msgbox_title_confirm", L"Confirm" },
{ L"msgbox_title_use_variables_syntax_error", L"Use Variables: Syntax Error" },
{ L"msgbox_title_use_variables_execution_error", L"Use Variables: Execution Error" },
{ L"msgbox_title_save_list", L"Save list" },
{ L"msgbox_title_reload", L"Reload" },
{ L"msgbox_title_warning", L"Warning" },
{ L"msgbox_title_info", L"Info" },

{ L"msgbox_title_create_file", L"Create file" },
{ L"msgbox_prompt_create_file", L"$REPLACE_STRING doesn't exist. Create it?" },
{ L"msgbox_error_create_file", L"Cannot create the file $REPLACE_STRING" },

// MessageBox Messages
{ L"msgbox_failed_create_control", L"Failed to create control with ID: $REPLACE_STRING1, GetLastError returned: $REPLACE_STRING2" },
{ L"msgbox_confirm_replace_all", L"Are you sure you want to replace all occurrences in all open documents?" },
{ L"msgbox_confirm_delete_columns", L"Are you sure you want to delete $REPLACE_STRING column(s)?" },
{ L"msgbox_error_saving_settings", L"An error occurred while saving the settings:<br/>$REPLACE_STRING" },
{ L"msgbox_use_variables_execution_error", L"Execution halted due to execution failure in:<br/>$REPLACE_STRING" },
{ L"msgbox_confirm_delete_single", L"Are you sure you want to delete this line?" },
{ L"msgbox_confirm_delete_multiple", L"Are you sure you want to delete $REPLACE_STRING lines?" },
{ L"msgbox_unsaved_changes_file", L"You have unsaved changes in the list: '$REPLACE_STRING'.<br/>Would you like to save them?" },
{ L"msgbox_unsaved_changes", L"You have unsaved changes.<br/>Would you like to save them?" },
{ L"msgbox_file_modified_prompt", L"'$REPLACE_STRING'<br/><br/>The file has been modified by another program.<br/>Do you want to load the changes and lose unsaved modifications?" },
{ L"msgbox_use_variables_not_exported", L"Some items with 'Use Variables' enabled were not exported." },
{ L"msgbox_no_files", L"No files matched the specified filter." },
{ L"msgbox_confirm_replace_in_files", L"Do you want to perform replacements in $REPLACE_STRING1 files?<br/><br/>In the directory:<br/>  $REPLACE_STRING2<br/><br/>For file type:<br/>  $REPLACE_STRING3" },
{ L"msgbox_flowtabs_intro_body",
  L"Activating Flow-Tabs (Column Alignment)\n\n"
  L"Please note the following before proceeding:\n\n"
  L"•  Tab-delimited files: Existing tabs are reused for alignment.\n"
  L"•  Non-tab delimited files: Real Tabs (\\t) are inserted. Note these may affect Find/Replace.\n"
  L"•  Reversal: Click the Flow-Tabs button again to remove the inserted tabs (before saving)." },
{ L"msgbox_flowtabs_intro_checkbox", L"Do not show this message again" },
{ L"msgbox_button_ok",     L"OK" },
{ L"msgbox_button_cancel", L"Cancel" },

// Context Menu List
{ L"ctxmenu_transfer_to_input_fields", L"&Transfer to Input Fields\tAlt+Up" },
{ L"ctxmenu_search_in_list", L"&Search in List\tCtrl+F" },
{ L"ctxmenu_cut", L"Cu&t\tCtrl+X" },
{ L"ctxmenu_copy", L"&Copy\tCtrl+C" },
{ L"ctxmenu_paste", L"&Paste\tCtrl+V" },
{ L"ctxmenu_edit", L"&Edit Field\t" },
{ L"ctxmenu_delete", L"&Delete\tDel" },
{ L"ctxmenu_select_all", L"Select &All\tCtrl+A" },
{ L"ctxmenu_enable", L"E&nable\tAlt+E" },
{ L"ctxmenu_disable", L"D&isable\tAlt+D" },
{ L"ctxmenu_undo", L"U&ndo\tCtrl+Z" },
{ L"ctxmenu_redo", L"R&edo\tCtrl+Y" },
{ L"ctxmenu_add_new_line", L"&Add New Line\tCtrl+I" },
{ L"ctxmenu_set_options", L"Set Options" },
{ L"ctxmenu_clear_options", L"Clear Options" },
{ L"ctxmenu_opt_wholeword", L"Whole Word" },
{ L"ctxmenu_opt_matchcase", L"Match Case" },
{ L"ctxmenu_opt_variables", L"Variables" },
{ L"ctxmenu_opt_regex", L"Regex" },
{ L"ctxmenu_opt_extended", L"Extended" },

// Result Dock Menu
{ L"rdmenu_fold_all",            L"Fold all" },
{ L"rdmenu_unfold_all",          L"Unfold all" },
{ L"rdmenu_copy_std",            L"&Copy\tCtrl+C" },
{ L"rdmenu_copy_lines",          L"Copy selected line(s)" },
{ L"rdmenu_copy_paths",          L"Copy selected pathname(s)" },
{ L"rdmenu_select_all",          L"Select all\tCtrl+A" },
{ L"rdmenu_clear_all",           L"Clear all" },
{ L"rdmenu_open_paths",          L"Open selected pathname(s)" },
{ L"rdmenu_wrap",                L"Word wrap long lines" },
{ L"rdmenu_purge",               L"Purge for every search" },

{ L"dock_list_header", L"Search in List ($REPLACE_STRING1 hits in $REPLACE_STRING2 file(s))" },
{ L"dock_single_header", L"Search \"$REPLACE_STRING1\" ($REPLACE_STRING2 hits in $REPLACE_STRING3 file(s))" },
{ L"dock_crit_header",L"Search \"$REPLACE_STRING1\" ($REPLACE_STRING2 hits)" },
{ L"dock_hits_suffix", L"($REPLACE_STRING hits)" },
{ L"dock_line", L"Line" },

// Configuration Dialog
{ L"config_btn_close", L"Close" },
{ L"config_btn_reset", L"Reset All Settings" },

// Config Categories
{ L"config_cat_search_replace", L"Search and Replace" },
{ L"config_cat_list_view", L"List View and Layout" },
{ L"config_cat_csv", L"CSV Options" },
{ L"config_cat_export", L"Export" },
{ L"config_cat_appearance", L"Appearance" },

// Search & Replace Settings
{ L"config_grp_search_behaviour", L"Search behaviour" },
{ L"config_chk_stay_after_replace", L"Replace: Don't move to the following occurrence" },
{ L"config_chk_all_from_cursor", L"Find: Search from cursor position" },
{ L"config_chk_mute_sounds", L"Mute all sounds" },

// File size limit
{ L"config_chk_limit_filesize", L"File Search: Skip files larger than" },
{ L"config_lbl_max_filesize_mb", L"MB" },

// List View Settings
{ L"config_grp_list_columns", L"Visible Columns" },
{ L"config_chk_find_count", L"Matches Count" },
{ L"config_chk_replace_count", L"Replaced Count" },
{ L"config_chk_comments", L"Comments" },
{ L"config_chk_delete_button", L"Delete Button" },
{ L"config_grp_list_results", L"List Results" },
{ L"config_chk_list_stats", L"Show list statistics next to file path" },
{ L"config_chk_group_results", L"Find All: Group hits by list entry" },
{ L"config_grp_list_interaction", L"List Interaction && View" },
{ L"config_chk_highlight_match", L"Select list entry on Find Next/Prev" },
{ L"config_chk_doubleclick", L"Edit in-place on double-click" },
{ L"config_chk_hover_text", L"Show full text on hover" },
{ L"config_lbl_edit_height", L"Expanded edit height (lines):" },

// CSV Settings
{ L"config_grp_csv_settings", L"CSV Settings" },
{ L"config_chk_numeric_align", L"Flow Tabs: Right-align numeric columns" },
{ L"config_chk_flowtabs_intro_dontshow", L"Flow Tabs: Don't show intro message" },
{ L"config_lbl_csv_sort", L"CSV Sort: Header lines to exclude:" },

// Export
{ L"config_grp_export_data", L"Copy List Data" },
{ L"config_lbl_export_template", L"Template:" },
{ L"config_chk_export_escape", L"Escape special characters" },
{ L"config_chk_export_header", L"Add header row" },
{ L"ctxmenu_export_data", L"Copy List Data" },
{ L"status_exported_to_clipboard", L"Exported $REPLACE_STRING entries to clipboard" },
{ L"status_no_items_to_export", L"No items to export" },

// Appearance Settings
{ L"config_grp_interface", L"Interface" },
{ L"config_lbl_foreground", L"Foreground transparency" },
{ L"config_lbl_background", L"Background transparency" },
{ L"config_lbl_scale_factor", L"Scale factor" },
{ L"config_grp_display_options", L"Display Options" },
{ L"config_chk_enable_tooltips", L"Enable tooltips" },
{ L"config_chk_result_dock_entry_colors", L"Color matches by list entry in Result Dock" },
{ L"config_chk_use_list_colors_marking", L"Use different colors for each list entry when marking" },

// Plugin Menu
{ L"menu_multiple_replacement", L"&Multiple Replacement..." },
{ L"menu_settings", L"&Settings..." },
{ L"menu_documentation", L"&Documentation" },
{ L"menu_about", L"&About MultiReplace" },

// About Dialog
{ L"about_title", L"MultiReplace Plugin" },
{ L"about_version", L"Version:" },
{ L"about_author", L"Author:" },
{ L"about_license", L"License:" },
{ L"about_help_support", L"Help and Support" },
{ L"about_ok", L"OK" },

// Debug Window
{ L"debug_title", L"Debug Information" },
{ L"debug_title_complete", L"Debug Information (Complete)" },
{ L"debug_btn_next", L"Next" },
{ L"debug_btn_stop", L"Stop" },
{ L"debug_btn_close", L"Close" },
{ L"debug_btn_copy", L"Copy" },
{ L"debug_col_variable", L"Variable" },
{ L"debug_col_type", L"Type" },
{ L"debug_col_value", L"Value" },

};

const size_t kEnglishPairsCount = sizeof(kEnglishPairs) / sizeof(kEnglishPairs[0]);

// ============================================================
// Main Panel: Control Text Mappings
// ============================================================

const UITextMapping kControlTextMappings[] = {
    // Labels
    { IDC_STATIC_FIND,                 L"panel_find_what" },
    { IDC_STATIC_REPLACE,              L"panel_replace_with" },
    // Checkboxes
    { IDC_WHOLE_WORD_CHECKBOX,         L"panel_match_whole_word_only" },
    { IDC_MATCH_CASE_CHECKBOX,         L"panel_match_case" },
    { IDC_USE_VARIABLES_CHECKBOX,      L"panel_use_variables" },
    { IDC_USE_VARIABLES_HELP,          L"panel_help" },
    { IDC_WRAP_AROUND_CHECKBOX,        L"panel_wrap_around" },
    { IDC_REPLACE_AT_MATCHES_CHECKBOX, L"panel_replace_at_matches" },
    // Group boxes
    { IDC_SEARCH_MODE_GROUP,           L"panel_search_mode" },
    { IDC_SCOPE_GROUP,                 L"panel_scope" },
    // Radio buttons
    { IDC_NORMAL_RADIO,                L"panel_normal" },
    { IDC_EXTENDED_RADIO,              L"panel_extended" },
    { IDC_REGEX_RADIO,                 L"panel_regular_expression" },
    { IDC_ALL_TEXT_RADIO,              L"panel_all_text" },
    { IDC_SELECTION_RADIO,             L"panel_selection" },
    { IDC_COLUMN_MODE_RADIO,           L"panel_csv" },
    // Column mode labels
    { IDC_COLUMN_NUM_STATIC,           L"panel_cols" },
    { IDC_DELIMITER_STATIC,            L"panel_delim" },
    { IDC_QUOTECHAR_STATIC,            L"panel_quote" },
    // Buttons
    { IDC_COPY_TO_LIST_BUTTON,         L"panel_add_into_list" },
    { IDC_REPLACE_ALL_BUTTON,          L"panel_replace_all" },
    { IDC_REPLACE_BUTTON,              L"panel_replace" },
    { IDC_FIND_ALL_BUTTON,             L"panel_find_all" },
    { IDC_FIND_NEXT_BUTTON,            L"panel_find_next" },
    { IDC_MARK_BUTTON,                 L"panel_mark_matches" },
    { IDC_MARK_MATCHES_BUTTON,         L"panel_mark_matches_small" },
    { IDC_CLEAR_MARKS_BUTTON,          L"panel_clear_all_marks" },
    { IDC_LOAD_FROM_CSV_BUTTON,        L"panel_load_list" },
    { IDC_LOAD_LIST_BUTTON,            L"panel_load_list" },
    { IDC_SAVE_TO_CSV_BUTTON,          L"panel_save_list" },
    { IDC_SAVE_AS_BUTTON,              L"panel_save_as" },
    { IDC_EXPORT_BASH_BUTTON,          L"panel_export_to_bash" },
    // Replace in Files Panel - NEU HINZUGEFÜGT
    { IDC_DIR_STATIC,                  L"panel_directory" },
    { IDC_FILTER_STATIC,               L"panel_filter" },
    { IDC_SUBFOLDERS_CHECKBOX,         L"panel_in_subfolders" },
    { IDC_HIDDENFILES_CHECKBOX,        L"panel_in_hidden_folders" },
    { IDC_CANCEL_REPLACE_BUTTON,       L"panel_cancel_replace" },
};

const size_t kControlTextMappingsCount = sizeof(kControlTextMappings) / sizeof(kControlTextMappings[0]);

// ============================================================
// Main Panel: Tooltip Mappings (for future use when tooltips are centralized)
// ============================================================

const UITooltipMapping kTooltipMappings[] = {
    { IDC_COLUMN_NUM_EDIT,             L"tooltip_columns" },
    { IDC_DELIMITER_EDIT,              L"tooltip_delimiter" },
    { IDC_QUOTECHAR_EDIT,              L"tooltip_quote" },
    { IDC_REPLACE_HIT_EDIT,            L"tooltip_replace_at_matches" },
    { IDC_COLUMN_SORT_DESC_BUTTON,     L"tooltip_sort_descending" },
    { IDC_COLUMN_SORT_ASC_BUTTON,      L"tooltip_sort_ascending" },
    { IDC_COLUMN_DROP_BUTTON,          L"tooltip_drop_columns" },
    { IDC_COLUMN_COPY_BUTTON,          L"tooltip_copy_columns" },
    { IDC_COLUMN_HIGHLIGHT_BUTTON,     L"tooltip_column_highlight" },
    { IDC_COLUMN_GRIDTABS_BUTTON,      L"tooltip_column_tabs" },
    { IDC_COPY_MARKED_TEXT_BUTTON,     L"tooltip_copy_marked_text" },
    { IDC_REPLACE_ALL_SMALL_BUTTON,    L"tooltip_replace_all" },
    { IDC_2_BUTTONS_MODE,              L"tooltip_2_buttons_mode" },
    { IDC_NEW_LIST_BUTTON,             L"tooltip_new_list" },
    { IDC_SAVE_BUTTON,                 L"tooltip_save" },
};

const size_t kTooltipMappingsCount = sizeof(kTooltipMappings) / sizeof(kTooltipMappings[0]);

// ============================================================
// ListView Header Text Mappings
// ============================================================

const UIHeaderMapping kHeaderTextMappings[] = {
    { ColumnID::FIND_COUNT,    L"header_find_count" },
    { ColumnID::REPLACE_COUNT, L"header_replace_count" },
    { ColumnID::FIND_TEXT,     L"header_find" },
    { ColumnID::REPLACE_TEXT,  L"header_replace" },
    { ColumnID::WHOLE_WORD,    L"header_whole_word" },
    { ColumnID::MATCH_CASE,    L"header_match_case" },
    { ColumnID::USE_VARIABLES, L"header_use_variables" },
    { ColumnID::EXTENDED,      L"header_extended" },
    { ColumnID::REGEX,         L"header_regex" },
    { ColumnID::COMMENTS,      L"header_comments" },
};

const size_t kHeaderTextMappingsCount = sizeof(kHeaderTextMappings) / sizeof(kHeaderTextMappings[0]);

// ============================================================
// ListView Header Tooltip Mappings
// ============================================================

const UIHeaderMapping kHeaderTooltipMappings[] = {
    { ColumnID::WHOLE_WORD,    L"tooltip_header_whole_word" },
    { ColumnID::MATCH_CASE,    L"tooltip_header_match_case" },
    { ColumnID::USE_VARIABLES, L"tooltip_header_use_variables" },
    { ColumnID::EXTENDED,      L"tooltip_header_extended" },
    { ColumnID::REGEX,         L"tooltip_header_regex" },
};

const size_t kHeaderTooltipMappingsCount = sizeof(kHeaderTooltipMappings) / sizeof(kHeaderTooltipMappings[0]);
