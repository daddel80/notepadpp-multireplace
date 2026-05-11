// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// CsvListFormat.h
// MR list CSV format: column field enum, header parsing, row lookup,
// and CSV line tokenization. Shared by CSV file load/save and
// clipboard copy/paste. Generic quoted-string escape helpers (used
// for cell bodies and INI values alike) live in StringUtils.

#pragma once

#include <array>
#include <string>
#include <vector>

namespace CsvListFormat {

    // Split one CSV line into its cells. Handles quoted cells with
    // embedded commas and "" escapes; trailing CR/LF stripped; cell
    // bodies are passed through StringUtils::unescapeQuoted.
    std::vector<std::wstring> parseLine(const std::wstring& line);

    // ---- Header parsing -----------------------------------------------

    enum class Field {
        Unknown,
        Selected,
        Find,
        Replace,
        WholeWord,
        MatchCase,
        FormulaSupport,
        Extended,
        Regex,
        Comments,
        LastModified
    };

    // Header-to-column resolution. idx[Field] = -1 means the field
    // was not present in the parsed header row.
    struct HeaderIndex {
        std::array<int, 11> idx;
        bool hasFind = false;
        bool hasReplace = false;
        bool looksLikeNames = false;
    };

    // Lowercase ASCII, strip whitespace/underscore/hyphen.
    std::wstring normalizeName(const std::wstring& raw);

    // Map a normalized header cell to a Field. "usevariables" maps
    // to FormulaSupport as a legacy alias.
    Field fieldFromName(const std::wstring& normalized);

    // Build a HeaderIndex from a parsed header row. looksLikeNames
    // is set if "Find" is present, or at least two fields are
    // recognized. The first occurrence wins on duplicates.
    HeaderIndex buildIndex(const std::vector<std::wstring>& headerCells);

    // Return the raw cell for a field, or def if the field is not
    // in the header or the row is shorter than the column index.
    std::wstring cellAt(const HeaderIndex& m,
        const std::vector<std::wstring>& row,
        Field f,
        const std::wstring& def = L"");

    // Same as cellAt but parses the cell as a 0/1 boolean. Returns
    // def for missing, empty, or non-integer cells.
    bool cellAtBool(const HeaderIndex& m,
        const std::vector<std::wstring>& row,
        Field f,
        bool def = false);

}  // namespace CsvListFormat