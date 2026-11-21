// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
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

#pragma once

#include <cstddef>
#include <string_view>
#include <vector>

// One CSV line: where its delimiters are (outside quotes, every quote character toggles quoting) and what deleting
// columns takes out of it
namespace CsvLine {

    // Calls onDelimiter(offset) for each delimiter, left to right
    template <class OnDelimiter>
    void forEachDelimiter(std::string_view line, std::string_view delimiter, char quote, OnDelimiter&& onDelimiter)
    {
        if (delimiter.empty()) return;
        const bool quoted = quote != '\0';
        const char first = delimiter[0];
        bool inQuotes = false;
        std::size_t pos = 0;

        while (pos < line.size()) {
            if (quoted && line[pos] == quote) {
                inQuotes = !inQuotes;
                ++pos;
                continue;
            }
            if (!inQuotes) {
                if (delimiter.size() == 1) {
                    if (line[pos] == first) {
                        onDelimiter(pos);
                        ++pos;
                        continue;
                    }
                }
                else {
                    const std::size_t found = line.find(delimiter, pos);
                    if (found == std::string_view::npos) break;
                    if (quoted) {
                        const std::size_t nextQuote = line.find(quote, pos);
                        if (nextQuote != std::string_view::npos && nextQuote < found) {
                            pos = nextQuote;
                            continue;
                        }
                    }
                    onDelimiter(found);
                    pos = found + delimiter.size();
                    continue;
                }
            }
            ++pos;
        }
    }

    struct Range {
        std::size_t begin = 0;
        std::size_t end = 0;
    };

    struct ColumnDeletion {
        std::vector<Range> ranges;   // ascending, touching ranges merged
        std::size_t fields = 0;      // deleted fields; an empty line has none
    };

    // Text to delete for the 1-based columns (ascending) of a line; delimiterAt(k) is the offset of delimiter k.
    // A column goes with its leading delimiter, or with its trailing one while no kept column precedes it
    template <class Columns, class DelimiterAt>
    ColumnDeletion columnDeletion(const Columns& columns, std::size_t delimiterCount, DelimiterAt&& delimiterAt,
        std::size_t delimiterLength, std::size_t contentEnd)
    {
        ColumnDeletion deletion;
        std::size_t deletedBefore = 0;
        for (const auto value : columns) {
            if (value < 1) continue;
            const std::size_t column = static_cast<std::size_t>(value);
            if (column > delimiterCount + 1) break;

            const bool last = column == delimiterCount + 1;
            const std::size_t contentBegin = (column == 1) ? 0 : delimiterAt(column - 2) + delimiterLength;
            const Range range = (deletedBefore == column - 1)
                ? Range{ contentBegin, last ? contentEnd : delimiterAt(column - 1) + delimiterLength }
                : Range{ delimiterAt(column - 2), last ? contentEnd : delimiterAt(column - 1) };
            ++deletedBefore;
            if (delimiterCount > 0 || contentEnd > 0) ++deletion.fields;

            if (range.end <= range.begin) continue;
            if (!deletion.ranges.empty() && deletion.ranges.back().end == range.begin)
                deletion.ranges.back().end = range.end;
            else
                deletion.ranges.push_back(range);
        }
        return deletion;
    }

    // The typed columns a line of at most fieldCount fields can have, in typed order. The others are empty in every
    // line, so leaving them out changes no comparison and no copied text
    inline std::vector<int> columnsWithin(const std::vector<int>& columns, std::size_t fieldCount)
    {
        std::vector<int> within;
        for (const int column : columns) {
            if (column >= 1 && static_cast<std::size_t>(column) <= fieldCount) within.push_back(column);
        }
        return within;
    }
}
