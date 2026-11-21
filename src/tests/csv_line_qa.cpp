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

// QA for CsvLine: delimiters outside quotes, and Delete Columns. Deleting columns leaves the other fields joined by
// one delimiter each, also when the first columns go (a,b,c,d without 1,2 is c,d); the line break stays.
// Build: g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. csv_line_qa.cpp -o csv_line_qa
#include <cstddef>
#include <cstdio>
#include <random>
#include <set>
#include <string>
#include <vector>

#include "../CsvLine.h"

namespace {

    int g_checks = 0;
    int g_fail = 0;

    void check(bool ok, const std::string& what)
    {
        ++g_checks;
        if (!ok) {
            ++g_fail;
            if (g_fail <= 20) std::printf("FAIL  %s\n", what.c_str());
        }
    }

    std::vector<std::size_t> delimiters(const std::string& line, const std::string& delimiter, char quote)
    {
        std::vector<std::size_t> out;
        CsvLine::forEachDelimiter(line, delimiter, quote, [&](std::size_t pos) { out.push_back(pos); });
        return out;
    }

    std::size_t eolLength(const std::string& line)
    {
        if (line.size() >= 2 && line.compare(line.size() - 2, 2, "\r\n") == 0) return 2;
        if (!line.empty() && (line.back() == '\r' || line.back() == '\n')) return 1;
        return 0;
    }

    struct Applied {
        std::string text;
        std::size_t fields = 0;
        bool ordered = true;   // ascending, apart, inside the content
    };

    // The line as Delete Columns leaves it, the delimiters scanned with the line break as the panel does
    Applied deleteColumns(const std::string& line, const std::string& delimiter, char quote, const std::set<int>& columns)
    {
        const std::vector<std::size_t> delims = delimiters(line, delimiter, quote);
        const std::size_t contentEnd = line.size() - eolLength(line);
        const CsvLine::ColumnDeletion deletion = CsvLine::columnDeletion(columns, delims.size(),
            [&](std::size_t k) { return delims[k]; }, delimiter.size(), contentEnd);

        Applied applied{ line, deletion.fields, true };
        std::size_t previousEnd = 0;
        for (std::size_t i = 0; i < deletion.ranges.size(); ++i) {
            const CsvLine::Range& r = deletion.ranges[i];
            if (r.begin >= r.end || r.end > contentEnd || (i > 0 && r.begin <= previousEnd)) applied.ordered = false;
            previousEnd = r.end;
        }
        for (auto it = deletion.ranges.rbegin(); it != deletion.ranges.rend(); ++it)
            applied.text.erase(it->begin, it->end - it->begin);
        return applied;
    }

    // Model: the fields that stay, joined by the delimiter, then the line break
    Applied model(const std::string& line, const std::string& delimiter, char quote, const std::set<int>& columns)
    {
        const std::vector<std::size_t> delims = delimiters(line, delimiter, quote);
        const std::size_t contentEnd = line.size() - eolLength(line);
        std::vector<std::string> fields;
        std::size_t begin = 0;
        for (const std::size_t d : delims) {
            fields.push_back(line.substr(begin, d - begin));
            begin = d + delimiter.size();
        }
        fields.push_back(line.substr(begin, contentEnd - begin));

        Applied m;
        bool first = true;
        for (std::size_t c = 1; c <= fields.size(); ++c) {
            if (columns.count(static_cast<int>(c))) {
                if (fields.size() > 1 || !fields[c - 1].empty()) ++m.fields;
                continue;
            }
            if (!first) m.text += delimiter;
            m.text += fields[c - 1];
            first = false;
        }
        m.text += line.substr(contentEnd);
        return m;
    }

    void expect(const std::string& line, const std::set<int>& columns, const std::string& expected, std::size_t fields,
        const std::string& delimiter = ",", char quote = '"')
    {
        const Applied got = deleteColumns(line, delimiter, quote, columns);
        std::string set;
        for (const int c : columns) set += std::to_string(c) + " ";
        check(got.text == expected && got.fields == fields && got.ordered,
            "\"" + line + "\" without { " + set + "}: \"" + got.text + "\" (" + std::to_string(got.fields) + " fields), expected \""
            + expected + "\" (" + std::to_string(fields) + ")");
    }
}

int main()
{
    // Delimiters
    check(delimiters("a,b,c", ",", '"') == std::vector<std::size_t>{ 1, 3 }, "plain delimiters");
    check(delimiters("\"x,y\",b", ",", '"') == std::vector<std::size_t>{ 5 }, "a delimiter in quotes does not count");
    check(delimiters("\"x,y\",b", ",", '\0') == std::vector<std::size_t>{ 2, 5 }, "without a quote character it counts");
    check(delimiters("a;;b;;;c", ";;", '"') == std::vector<std::size_t>{ 1, 4 }, "a delimiter of two characters");
    check(delimiters("", ",", '"').empty(), "empty line");

    // The first columns
    expect("a,b,c,d", { 1, 2 }, "c,d", 2);
    expect("a,b,c,d", { 1, 2, 4 }, "c", 3);
    expect("a,b,c,d", { 1, 2, 3 }, "d", 3);
    expect("a,b,c,d", { 1 }, "b,c,d", 1);
    expect("a,b,c,d", { 1, 3 }, "b,d", 2);

    // Further columns
    expect("a,b,c,d", { 2, 3 }, "a,d", 2);
    expect("a,b,c,d", { 2, 4 }, "a,c", 2);
    expect("a,b,c,d", { 4 }, "a,b,c", 1);
    expect("a,b,c,d", { 3, 4 }, "a,b", 2);
    expect("a,b,c,d", { 1, 2, 3, 4 }, "", 4);
    expect("a,b,c,d", { 5, 6 }, "a,b,c,d", 0);
    expect("a,b,c,d", { 0, -1, 2 }, "a,c,d", 1);

    // Line breaks, quotes, empty fields, longer delimiters
    expect("a,b,c\r\n", { 1, 2 }, "c\r\n", 2);
    expect("a,b,c\n", { 3 }, "a,b\n", 1);
    expect("a,b,c\r", { 1, 2, 3 }, "\r", 3);
    expect("\"x,y\",b,c", { 1 }, "b,c", 1);
    expect("a,\"x,y\",c", { 2 }, "a,c", 1);
    expect("a,,c", { 2 }, "a,c", 1);
    expect(",,", { 1, 2 }, "", 2);
    expect("a;;b;;c", { 1, 2 }, "c", 2, ";;");
    expect("a\tb\tc", { 2 }, "a\tc", 1, "\t");

    // Short lines
    expect("a", { 1 }, "", 1);
    expect("a", { 2 }, "a", 0);
    expect("a,b", { 2, 3 }, "a", 1);
    expect("", { 1 }, "", 0);
    expect("\r\n", { 1, 2 }, "\r\n", 0);

    // Random lines and column sets against the model
    const std::string alphabet = "ab,,;\" x";
    const std::vector<std::string> delims = { ",", ";", ",;" };
    const std::vector<std::string> eols = { "", "\n", "\r\n" };
    std::mt19937 rng(20261121u);
    int differences = 0;
    for (int round = 0; round < 200000; ++round) {
        std::string line(std::uniform_int_distribution<std::size_t>(0, 12)(rng), ' ');
        for (char& c : line) c = alphabet[std::uniform_int_distribution<std::size_t>(0, alphabet.size() - 1)(rng)];
        line += eols[std::uniform_int_distribution<std::size_t>(0, eols.size() - 1)(rng)];
        const std::string& delimiter = delims[std::uniform_int_distribution<std::size_t>(0, delims.size() - 1)(rng)];
        const char quote = (round % 3 == 0) ? '\0' : '"';
        std::set<int> columns;
        const int picks = std::uniform_int_distribution<int>(1, 4)(rng);
        for (int k = 0; k < picks; ++k) columns.insert(std::uniform_int_distribution<int>(0, 7)(rng));

        const Applied got = deleteColumns(line, delimiter, quote, columns);
        const Applied want = model(line, delimiter, quote, columns);
        ++g_checks;
        if (got.text != want.text || got.fields != want.fields || !got.ordered) {
            ++g_fail;
            if (++differences <= 10)
                std::printf("FAIL  random \"%s\" delimiter \"%s\": \"%s\" (%zu), model \"%s\" (%zu)\n", line.c_str(),
                    delimiter.c_str(), got.text.c_str(), got.fields, want.text.c_str(), want.fields);
        }
    }

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
