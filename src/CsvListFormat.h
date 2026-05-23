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

    // The body dialect of a list file. Mr is the internal escaped format
    // (\\ , \n , \r escapes plus "" quoting); Rfc4180 is plain CSV as
    // produced by Excel (only "" quoting, backslashes literal, newlines
    // allowed verbatim inside quoted fields). The dialect is always
    // declared by the caller, never sniffed from content.
    enum class Dialect {
        Mr,
        Rfc4180
    };

    // Split a whole CSV text into records (each a vector of cells).
    // Quote-aware across line breaks: a newline inside a quoted cell is
    // part of that cell, so multi-line fields survive. CRLF and lone CR
    // outside quotes both end a record. Cell bodies are decoded by the
    // chosen dialect: Mr uses StringUtils::unescapeQuoted, Rfc4180 uses
    // StringUtils::rfcUnquote. delimiter selects the field separator
    // (comma for the internal format; comma or semicolon for plain CSV).
    std::vector<std::vector<std::wstring>> readRecords(const std::wstring& text,
        Dialect dialect = Dialect::Mr, wchar_t delimiter = L',');

    // Sniff the field delimiter of a plain-CSV text by parsing its first
    // line with each candidate and picking the one whose cells resolve to
    // a recognizable header (Find present, or >=2 known names). Falls back
    // to the supplied default when neither candidate looks like a header
    // (e.g. a headerless file). Only meaningful for the Rfc4180 dialect.
    wchar_t detectDelimiter(const std::wstring& text, wchar_t fallback = L',');

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
        std::array<int, 11> idx{};   // 0 here; buildIndex fills with -1
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