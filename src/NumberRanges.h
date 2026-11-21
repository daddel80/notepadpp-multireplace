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

#include <algorithm>
#include <cstddef>
#include <exception>
#include <limits>
#include <sstream>
#include <string>
#include <unordered_set>
#include <vector>

// Number lists as typed in the panel, such as "1,3,5-7,10-8": numbers from 1, ranges in either direction.
// A range stays a range, so a list like "1-999999999" costs nothing until its numbers are listed one by one
namespace NumberRanges {

    struct Range {
        int from = 0;
        int to = 0;     // below from for a range typed downwards
    };

    constexpr int noLimit = (std::numeric_limits<int>::max)();

    // Each part between commas is a number or a range, read like std::stoi; empty parts are skipped.
    // False if a part is neither, or holds a number below 1 or above maxValue
    inline bool parse(const std::wstring& input, int maxValue, std::vector<Range>& ranges)
    {
        ranges.clear();
        std::wistringstream stream(input);
        std::wstring part;
        while (std::getline(stream, part, L',')) {
            if (part.empty()) continue;

            Range range;
            try {
                const std::size_t dash = part.find(L'-');
                range.from = std::stoi(part.substr(0, dash));
                range.to = dash == std::wstring::npos ? range.from : std::stoi(part.substr(dash + 1));
            }
            catch (const std::exception&) {
                ranges.clear();
                return false;
            }
            if (range.from < 1 || range.to < 1 || range.from > maxValue || range.to > maxValue) {
                ranges.clear();
                return false;
            }
            ranges.push_back(range);
        }
        return true;
    }

    // Every number once, in the order typed
    inline std::vector<int> expand(const std::vector<Range>& ranges)
    {
        std::vector<int> numbers;
        std::unordered_set<int> seen;
        for (const Range& range : ranges) {
            const int step = range.from <= range.to ? 1 : -1;
            for (int n = range.from; ; n += step) {
                if (seen.insert(n).second) numbers.push_back(n);
                if (n == range.to) break;
            }
        }
        return numbers;
    }

    inline bool contains(const std::vector<Range>& ranges, int number)
    {
        return std::any_of(ranges.begin(), ranges.end(), [number](const Range& range) {
            return range.from <= range.to ? number >= range.from && number <= range.to
                                          : number >= range.to && number <= range.from;
        });
    }
}
