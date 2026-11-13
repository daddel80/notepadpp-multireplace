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

// Where the delimiters of one CSV line are: outside quotes, and every quote character toggles quoting
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
}
