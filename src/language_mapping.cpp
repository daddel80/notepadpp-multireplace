﻿#include "MultiReplacePanel.h"


std::unordered_map<std::wstring, std::wstring> languageMap = {
{ L"panel_find_what", L"Find what: "},
{ L"panel_replace_with", L"Replace with: "},
{ L"panel_match_whole_word_only", L"Match whole word only" },
{ L"panel_replace_with", L"Replace with:"},
{ L"panel_match_whole_word_only", L"Match whole word only" },
{ L"panel_match_case", L"Match case" },
{ L"panel_use_variables", L"Use Variables" },
{ L"panel_replace_first_match_only", L"Replace first match only" },
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
{ L"panel_2_buttons_mode", L"2 buttons mode" },
{ L"panel_find_next", L"Find Next" },
{ L"panel_find_next_small", L"Find Next"},
{ L"panel_mark_matches", L"Mark Matches" },
{ L"panel_mark_matches_small", L"Mark Matches"},
{ L"panel_clear_all_marks", L"Clear all marks" },
{ L"panel_load_list", L"Load List" },
{ L"panel_save_list", L"Save List" },
{ L"panel_export_to_bash", L"Export to Bash" },
{ L"panel_shift_lines", L"Shift Lines" },
{ L"panel_use_list", L"Use List" },
{ L"panel_show", L"Show" },
{ L"panel_hide", L"Hide" },
{ L"panel_help", L"?" },

// Tooltips
{ L"tooltip_replace_all", L"Replace All" },
{ L"tooltip_2_buttons_mode", L"2 buttons mode" },
{ L"tooltip_columns", L"Columns: '1,3,5-12' (individuals, ranges)" },
{ L"tooltip_delimiter", L"Delimiter: Single/combined chars, \t for Tab" },
{ L"tooltip_quote", L"Quote: ', \", or empty" },
{ L"tooltip_sort_descending", L"Sort Descending" },
{ L"tooltip_sort_ascending", L"Sort Ascending" },
{ L"tooltip_drop_columns", L"Drop Columns" },
{ L"tooltip_copy_columns", L"Copy Columns to Clipboard" },
{ L"tooltip_column_highlight", L"Column highlight: On/Off" },

// header entries
{ L"header_find_count", L"Find Count" },
{ L"header_replace_count", L"Replace Count" },
{ L"header_find", L"Find" },
{ L"header_replace", L"Replace" },
{ L"header_whole_word", L"W" },
{ L"header_match_case", L"C" },
{ L"header_use_variables", L"V" },
{ L"header_extended", L"E" },
{ L"header_regex", L"R" },

// tooltip entries
{ L"tooltip_header_whole_word", L"Whole Word" },
{ L"tooltip_header_match_case", L"Case Sensitive" },
{ L"tooltip_header_use_variables", L"Use Variables" },
{ L"tooltip_header_extended", L"Extended" },
{ L"tooltip_header_regex", L"Regex" },

// SplitButton entries
{ L"split_menu_replace_all", L"Replace All" },
{ L"split_menu_replace_all_in_docs", L"Replace All in All opened Documents" },
{ L"split_button_replace_all", L"Replace All" },
{ L"split_button_replace_all_in_docs", L"Replace All in Docs" },

// Static Status message entries
{ L"status_duplicate_entry", L"Duplicate entry: " },
{ L"status_value_added", L"Value added to the list." },
{ L"status_no_rows_selected", L"No rows selected to shift." },
{ L"status_one_line_deleted", L"1 line deleted." },
{ L"status_column_marks_cleared", L"Column marks cleared." },
{ L"status_all_marks_cleared", L"All marks cleared." },
{ L"status_cannot_replace_read_only", L"Cannot replace. Document is read-only." },
{ L"status_add_values_instructions", L"Add values into the list. Or uncheck 'Use in List' to replace directly." },
{ L"status_no_find_string", L"No 'Find String' entered. Please provide a value to add to the list." },
{ L"status_no_rows_selected_to_shift", L"No rows selected to shift." },
{ L"status_add_values_or_uncheck", L"Add values into the list or uncheck 'Use in List'." },
{ L"status_no_occurrence_found", L"No occurrence found." },
{ L"status_replace_one_next_found", L"Replace: 1 occurrence replaced. Next found." },
{ L"status_replace_one_none_left", L"Replace: 1 occurrence replaced. None left." },
{ L"status_add_values_or_find_directly", L"Add values into the list. Or uncheck 'Use in List' to find directly." },
{ L"status_wrapped", L"Wrapped" },
{ L"status_no_matches_found", L"No matches found." },
{ L"status_no_matches_after_wrap", L"No matches found after wrap." },
{ L"status_add_values_or_mark_directly", L"Add values into the list. Or uncheck 'Use in List' to mark directly." },
{ L"status_no_text_to_copy", L"No text to copy." },
{ L"status_failed_to_copy", L"Failed to copy to Clipboard." },
{ L"status_failed_allocate_memory", L"Failed to allocate memory for Clipboard." },
{ L"status_invalid_column_or_delimiter", L"Invalid column or delimiter data." },
{ L"status_missing_column_or_delimiter_data", L"Column data or delimiter data is missing" },
{ L"status_invalid_range_in_column_data", L"Invalid range in column data" },
{ L"status_syntax_error_in_column_data", L"Syntax error in column data" },
{ L"status_invalid_column_number", L"Invalid column number" },
{ L"status_extended_delimiter_empty", L"Extended delimiter is empty" },
{ L"status_invalid_quote_character", L"Invalid quote character. Use \" or ' or leave it empty." },
{ L"status_unable_to_open_file", L"Error: Unable to open file for writing." },
{ L"status_no_valid_items_in_csv", L"No valid items found in the CSV file." },
{ L"status_list_exported_to_bash", L"List exported to BASH script." },

// Dynamic Status message entries
{ L"status_rows_shifted", L"$REPLACE_STRING rows successfully shifted." },
{ L"status_lines_deleted", L"$REPLACE_STRING lines deleted." },
{ L"status_column_sorted", L"Column sorted in $REPLACE_STRING order." },
{ L"status_occurrences_replaced", L"$REPLACE_STRING occurrences were replaced." },
{ L"status_replace_next_found", L"Replace: $REPLACE_STRING replaced. Next occurrence found." },
{ L"status_replace_none_left", L"Replace: $REPLACE_STRING replaced. None left." },
{ L"status_no_matches_found_for", L"No matches found for '$REPLACE_STRING'." },
{ L"status_actual_position", L"Actual Position $REPLACE_STRING" },
{ L"status_items_loaded_from_csv", L"$REPLACE_STRING items loaded from CSV." },
{ L"status_wrapped_position", L"Wrapped at $REPLACE_STRING" },
{ L"status_deleted_fields", L"Deleted $REPLACE_STRING fields." },
{ L"status_occurrences_marked", L"$REPLACE_STRING occurrences were marked." },
{ L"status_items_copied_to_clipboard", L"$REPLACE_STRING items copied into Clipboard." },
{ L"status_no_matches_after_wrap_for", L"No matches found for '$REPLACE_STRING' after wrap." },
{ L"status_deleted_fields_count", L"Deleted $REPLACE_STRING fields." },
{ L"status_wrapped_find", L"Wrapped '$REPLACE_STRING1'. Position: $REPLACE_STRING2" },
{ L"status_wrapped_no_find", L"Wrapped. Position: $REPLACE_STRING" },
{ L"status_line_and_column_position", L" (Line: $REPLACE_STRING, Column: $REPLACE_STRING1)" },

// MessageBox Titles
{ L"msgbox_title_error", L"Error" },
{ L"msgbox_title_confirm", L"Confirm" },
{ L"msgbox_title_use_variables_syntax_error", L"Use Variables: Syntax Error" },
{ L"msgbox_title_use_variables_execution_error", L"Use Variables: Execution Error" },

// MessageBox Messages
{ L"msgbox_failed_create_control", L"Failed to create control with ID: $REPLACE_STRING1, GetLastError returned: $REPLACE_STRING2" },
{ L"msgbox_confirm_replace_all", L"Are you sure you want to replace all occurrences in all open documents?" },
{ L"msgbox_confirm_delete_columns", L"Are you sure you want to delete $REPLACE_STRING column(s)?" },
{ L"msgbox_error_saving_settings", L"An error occurred while saving the settings:<br/>$REPLACE_STRING" },
{ L"msgbox_use_variables_execution_error", L"Execution halted due to execution failure in:<br/>$REPLACE_STRING" },
};