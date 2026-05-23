// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// ListCodec.cpp

#include "ListCodec.h"
#include "StringUtils.h"

#include <stdexcept>

namespace ListCodec {

    const std::wstring& headerLine() {
        static const std::wstring header =
            L"Selected,Find,Replace,WholeWord,MatchCase,FormulaSupport,Extended,Regex,Comments,LastModified";
        return header;
    }

    std::wstring headerLine(wchar_t delimiter) {
        if (delimiter == L',') return headerLine();
        std::wstring h = headerLine();
        for (wchar_t& c : h) {
            if (c == L',') c = delimiter;
        }
        return h;
    }

    bool hasLegacyUseVariablesHeader(const std::vector<std::wstring>& headerCells) {
        for (const auto& cell : headerCells) {
            if (CsvListFormat::normalizeName(cell) == L"usevariables") return true;
        }
        return false;
    }

    bool needsRegexExtendedSwap(const CsvListFormat::HeaderIndex& hdr,
        const std::vector<std::wstring>& headerCells) {
        const int rIdx = hdr.idx[static_cast<int>(CsvListFormat::Field::Regex)];
        const int eIdx = hdr.idx[static_cast<int>(CsvListFormat::Field::Extended)];
        if (rIdx < 0 || eIdx < 0 || rIdx >= eIdx) return false;
        return hasLegacyUseVariablesHeader(headerCells);
    }

    ReplaceItemData rowToItem(const CsvListFormat::HeaderIndex& hdr,
        const std::vector<std::wstring>& columns,
        bool swapRegexExtended) {
        ReplaceItemData item;
        item.isEnabled = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::Selected, false);
        item.findText = CsvListFormat::cellAt(hdr, columns, CsvListFormat::Field::Find, L"");
        item.replaceText = CsvListFormat::cellAt(hdr, columns, CsvListFormat::Field::Replace, L"");
        item.wholeWord = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::WholeWord, false);
        item.matchCase = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::MatchCase, false);
        item.formulaSupport = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::FormulaSupport, false);
        item.extended = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::Extended, false);
        item.regex = CsvListFormat::cellAtBool(hdr, columns, CsvListFormat::Field::Regex, false);
        if (swapRegexExtended) std::swap(item.extended, item.regex);
        item.comments = CsvListFormat::cellAt(hdr, columns, CsvListFormat::Field::Comments, L"");
        item.lastModified = CsvListFormat::cellAt(hdr, columns, CsvListFormat::Field::LastModified, L"");
        return item;
    }

    ReplaceItemData rowToItemPositional(const std::vector<std::wstring>& columns) {
        if (columns.size() < 8 || columns.size() > 10) {
            throw std::out_of_range("positional CSV row: expected 8..10 columns");
        }
        ReplaceItemData item;
        item.isEnabled = std::stoi(columns[0]) != 0;
        item.findText = columns[1];
        item.replaceText = columns[2];
        item.wholeWord = std::stoi(columns[3]) != 0;
        item.matchCase = std::stoi(columns[4]) != 0;
        item.formulaSupport = std::stoi(columns[5]) != 0;
        item.extended = std::stoi(columns[6]) != 0;
        item.regex = std::stoi(columns[7]) != 0;
        item.comments = (columns.size() >= 9) ? columns[8] : L"";
        item.lastModified = (columns.size() >= 10) ? columns[9] : L"";
        return item;
    }

    std::wstring itemToRow(const ReplaceItemData& item, bool withLastModified,
        CsvListFormat::Dialect dialect, wchar_t delimiter) {
        // Text cells switch encoder by dialect; the 0/1 flag cells are the
        // same in both and never need quoting. For the plain dialect the
        // internal CRLF is first lowered to the Excel convention (real \n
        // for CRLF, literal "\n" for a foreign lone LF), then quoted.
        auto enc = [dialect, delimiter](const std::wstring& v) {
            return dialect == CsvListFormat::Dialect::Rfc4180
                ? StringUtils::rfcQuote(StringUtils::newlinesForPlainCsv(v), delimiter)
                : StringUtils::escapeQuoted(v);
            };
        const std::wstring d(1, delimiter);
        std::wstring row =
            std::to_wstring(item.isEnabled) + d +
            enc(item.findText) + d +
            enc(item.replaceText) + d +
            std::to_wstring(item.wholeWord) + d +
            std::to_wstring(item.matchCase) + d +
            std::to_wstring(item.formulaSupport) + d +
            std::to_wstring(item.extended) + d +
            std::to_wstring(item.regex) + d +
            enc(item.comments);
        if (withLastModified) {
            row += d + enc(item.lastModified);
        }
        return row;
    }

}  // namespace ListCodec