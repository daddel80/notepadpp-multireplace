// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ListCodec.h
// Bridges the list row data model (ReplaceItemData) and the CSV text
// representation. Owns the single source of truth for column order,
// row<->item mapping, the legacy positional layout, and the
// pre-V5 Regex/Extended swap detection. Sits above CsvListFormat
// (tokenizing/header) and StringUtils (cell escaping); the panel and
// clipboard paths go through here instead of hand-building rows.

#pragma once

#include "CsvListFormat.h"
#include "ReplaceItemData.h"

#include <string>
#include <vector>

namespace ListCodec {

    // The canonical CSV header row (no trailing newline). Single source
    // for both file save and any future writer.
    const std::wstring& headerLine();

    // Same header with a chosen field delimiter (for the plain-CSV
    // dialect when a non-comma separator is configured).
    std::wstring headerLine(wchar_t delimiter);

    // ---- Read: CSV row -> item ----------------------------------------

    // Detect the pre-V5 layout where a "UseVariables" column placed Regex
    // before Extended, requiring the two flags to be swapped on read.
    // Returns true if the swap applies for this header.
    bool needsRegexExtendedSwap(const CsvListFormat::HeaderIndex& hdr,
        const std::vector<std::wstring>& headerCells);

    // True if the header carries the old "Use Variables" column, the sole
    // reliable fingerprint of a legacy MultiReplace list. Used to detect a
    // legacy list that was opened through the plain "CSV (Excel)" path.
    bool hasLegacyUseVariablesHeader(const std::vector<std::wstring>& headerCells);

    // Build an item from a name-based row using the resolved header.
    // Missing cells fall back to defaults; swap applies the legacy fix.
    // lastModified is read when the column is present.
    ReplaceItemData rowToItem(const CsvListFormat::HeaderIndex& hdr,
        const std::vector<std::wstring>& columns,
        bool swapRegexExtended);

    // Build an item from a legacy positional row (no header). Accepts
    // 8..10 columns: 8 = pre-comment, 9 = +Comments, 10 = +LastModified.
    // Throws std::exception on the wrong column count or bad integers,
    // mirroring the historical std::stoi-based behaviour; callers decide
    // whether to skip or abort.
    ReplaceItemData rowToItemPositional(const std::vector<std::wstring>& columns);

    // ---- Write: item -> CSV row ---------------------------------------

    // Serialize one item to a CSV row (no trailing newline). The dialect
    // selects cell encoding: Mr uses StringUtils::escapeQuoted (always
    // quoted, backslash-escaped); Rfc4180 uses StringUtils::rfcQuote
    // (quoted only when needed, backslashes literal). delimiter is the
    // field separator. withLastModified adds the trailing LastModified
    // cell (file save) or omits it (clipboard).
    std::wstring itemToRow(const ReplaceItemData& item, bool withLastModified,
        CsvListFormat::Dialect dialect = CsvListFormat::Dialect::Mr,
        wchar_t delimiter = L',');

}  // namespace ListCodec