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


// QA for ListCodec::clipboardItems, the one reading of pasted text behind Ctrl+V and the Paste entry of the list's
// context menu. It reads like the paste loop it replaced, except that a copied entry with "find" as one of its texts
// is no longer taken for a header, and the menu entry is offered exactly when the paste inserts something.
// Build and run: list_codec_qa.sh
#include <cstddef>
#include <cstdio>
#include <exception>
#include <random>
#include <string>
#include <vector>

#include "../CsvListFormat.h"
#include "../ListCodec.h"

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

    // The paste loop and the menu check before the change, apart from names
    namespace Before {

        std::vector<ReplaceItemData> pasted(const std::wstring& content)
        {
            std::vector<ReplaceItemData> itemsToInsert;
            std::vector<std::vector<std::wstring>> records = CsvListFormat::readRecords(content);

            CsvListFormat::HeaderIndex hdr;
            bool nameBased = false;
            bool swapRegexExtended = false;
            size_t firstDataRow = 0;
            if (!records.empty()) {
                CsvListFormat::HeaderIndex maybeHdr = CsvListFormat::buildIndex(records.front());
                if (maybeHdr.looksLikeNames && maybeHdr.hasFind) {
                    hdr = maybeHdr;
                    nameBased = true;
                    firstDataRow = 1;
                    swapRegexExtended = ListCodec::needsRegexExtendedSwap(hdr, records.front());
                }
            }

            for (size_t r = firstDataRow; r < records.size(); ++r) {
                const std::vector<std::wstring>& columns = records[r];
                if (columns.size() == 1 && columns[0].empty()) continue;

                ReplaceItemData item;
                try {
                    if (nameBased) {
                        item = ListCodec::rowToItem(hdr, columns, swapRegexExtended);
                    }
                    else {
                        if (columns.size() < 8 || columns.size() > 10) continue;
                        item = ListCodec::rowToItemPositional(columns);
                    }
                }
                catch (const std::exception&) {
                    continue;
                }
                itemsToInsert.push_back(item);
            }
            return itemsToInsert;
        }

        bool canPaste(const std::wstring& content)
        {
            std::vector<std::vector<std::wstring>> records = CsvListFormat::readRecords(content);
            for (const auto& columns : records) {
                if (columns.size() == 8 || columns.size() == 9 || columns.size() == 10) {
                    return true;
                }
            }
            return false;
        }
    }

    bool isEntryRow(const std::vector<std::wstring>& cells)
    {
        try {
            ListCodec::rowToItemPositional(cells);
            return true;
        }
        catch (const std::exception&) {
            return false;
        }
    }

    // What the paste reads now: as before, but a first row that is a copied entry stays an entry
    std::vector<ReplaceItemData> expected(const std::wstring& content, bool& entryKept)
    {
        const std::vector<std::vector<std::wstring>> records = CsvListFormat::readRecords(content);
        entryKept = false;
        if (!records.empty()) {
            const CsvListFormat::HeaderIndex hdr = CsvListFormat::buildIndex(records.front());
            entryKept = hdr.looksLikeNames && hdr.hasFind && isEntryRow(records.front());
        }
        if (!entryKept) return Before::pasted(content);

        std::vector<ReplaceItemData> items;
        for (const auto& columns : records) {
            if (isEntryRow(columns)) items.push_back(ListCodec::rowToItemPositional(columns));
        }
        return items;
    }

    bool same(const ReplaceItemData& a, const ReplaceItemData& b)
    {
        return a.id == b.id && a.findCount == b.findCount && a.replaceCount == b.replaceCount
            && a.isEnabled == b.isEnabled && a.findText == b.findText && a.replaceText == b.replaceText
            && a.wholeWord == b.wholeWord && a.matchCase == b.matchCase && a.formulaSupport == b.formulaSupport
            && a.extended == b.extended && a.regex == b.regex && a.comments == b.comments
            && a.isDirty == b.isDirty && a.lastModified == b.lastModified;
    }

    bool same(const std::vector<ReplaceItemData>& a, const std::vector<ReplaceItemData>& b, std::size_t count)
    {
        if (a.size() != count || b.size() < count) return false;
        for (std::size_t i = 0; i < count; ++i) {
            if (!same(a[i], b[i])) return false;
        }
        return true;
    }

    bool same(const std::vector<ReplaceItemData>& a, const std::vector<ReplaceItemData>& b)
    {
        return a.size() == b.size() && same(a, b, b.size());
    }

    ReplaceItemData entry(bool enabled, const std::wstring& find, const std::wstring& replace, int flags,
        const std::wstring& comments)
    {
        ReplaceItemData item;
        item.isEnabled = enabled;
        item.findText = find;
        item.replaceText = replace;
        item.wholeWord = (flags & 1) != 0;
        item.matchCase = (flags & 2) != 0;
        item.formulaSupport = (flags & 4) != 0;
        item.extended = (flags & 8) != 0;
        item.regex = (flags & 16) != 0;
        item.comments = comments;
        return item;
    }

    // The clipboard text MR writes for these entries
    std::wstring copied(const std::vector<ReplaceItemData>& items)
    {
        std::wstring text;
        for (const auto& item : items) text += ListCodec::itemToRow(item, false) + L'\n';
        return text;
    }
}

int main()
{
    check(ListCodec::clipboardItems(L"").empty(), "F1 empty text: nothing");

    const std::vector<ReplaceItemData> entries = {
        entry(true, L"find", L"search", 0, L""),
        entry(false, L"a,b", L"say \"hi\"", 1 | 16, L"two\r\nlines"),
        entry(true, L"C:\\path\\", L"", 2 | 4 | 8, L"comment, with comma"),
    };
    check(same(ListCodec::clipboardItems(copied(entries)), entries), "F2 entries as MR copies them come back unchanged");
    check(Before::pasted(copied(entries)).size() == entries.size() - 1,
        "F2 before: the first entry, Find text \"find\", was taken for a header");

    const std::wstring named = L"Find,Replace\nfoo,bar\n";
    const std::vector<ReplaceItemData> namedItems = ListCodec::clipboardItems(named);
    check(namedItems.size() == 1 && namedItems[0].findText == L"foo" && namedItems[0].replaceText == L"bar",
        "F3 a header with two columns: one entry");
    check(!Before::canPaste(named) && !Before::pasted(named).empty(), "F3 before: pasted, but the menu entry was gray");
    check(!ListCodec::clipboardItems(named, 1).empty(), "F3 now the menu entry is offered");

    check(ListCodec::clipboardItems(L"Selected,Find,Replace\n").empty(), "F4 a header alone: nothing");

    const std::vector<ReplaceItemData> byCount = ListCodec::clipboardItems(
        L"1,a,b,0,0,0,0\n1,b,b,0,0,0,0,0\n1,c,b,0,0,0,0,0,x\n1,d,b,0,0,0,0,0,x,y\n1,e,b,0,0,0,0,0,x,y,z\n");
    check(byCount.size() == 3 && byCount[0].findText == L"b" && byCount[2].findText == L"d" && byCount[2].lastModified == L"y",
        "F5 by position: 8 to 10 cells, 7 and 11 skipped");

    check(ListCodec::clipboardItems(L"1,a,b,0,x,0,0,0\n1,b,b,0,0,0,0,0\n").size() == 1, "F6 a flag cell that is no number: row skipped");

    const std::vector<ReplaceItemData> blank = ListCodec::clipboardItems(L"\r\n1,a,b,0,0,0,0,0\r\n\r\n1,b,b,0,0,0,0,0\r\n");
    check(blank.size() == 2 && blank[1].findText == L"b", "F7 blank lines and CRLF");

    const std::vector<ReplaceItemData> legacy = ListCodec::clipboardItems(
        L"Selected,Find,Replace,WholeWord,MatchCase,UseVariables,Regex,Extended,Comments\n1,a,b,0,0,0,1,0,c\n");
    check(legacy.size() == 1 && legacy[0].extended && !legacy[0].regex, "F8 a list before V5: Regex and Extended swapped");

    const std::wstring three = L"Find\na\nb\nc\n";
    const std::vector<ReplaceItemData> all = ListCodec::clipboardItems(three);
    const std::vector<ReplaceItemData> two = ListCodec::clipboardItems(three, 2);
    const std::vector<ReplaceItemData> one = ListCodec::clipboardItems(three, 1);
    check(all.size() == 3 && same(two, all, 2) && same(one, all, 1) && ListCodec::clipboardItems(three, 0).empty(),
        "F9 limit: the first entries");

    // Random texts: read as before, apart from a copied first entry; the menu check agrees with the paste
    std::mt19937 rng(20261121u);
    auto pick = [&](std::size_t n) { return std::uniform_int_distribution<std::size_t>(0, n - 1)(rng); };
    const std::vector<std::wstring> names = { L"Selected", L"Find", L"Replace", L"WholeWord", L"MatchCase",
        L"FormulaSupport", L"UseVariables", L"Extended", L"Regex", L"Comments", L"LastModified", L"match_case",
        L"Whole Word", L"find", L"Other" };
    const std::vector<std::wstring> flags = { L"0", L"1", L"0", L"1", L"2", L"x", L"\" 1\"" };
    const std::vector<std::wstring> texts = { L"", L"a", L"find", L"Find", L"F-IND", L"Regex", L"Extended", L"0", L"1",
        L"\"a,b\"", L"\"say \"\"hi\"\"\"", L"\"two\nlines\"", L"\"\\\\\"", L"x y", L"\"" };
    const std::vector<std::wstring> eols = { L"\n", L"\r\n", L"\r" };

    auto row = [&](int kind) {
        std::wstring line;
        if (kind == 0) {  // a header
            const std::size_t n = 1 + pick(6);
            for (std::size_t c = 0; c < n; ++c) line += (c ? L"," : L"") + names[pick(names.size())];
            return line;
        }
        const std::size_t n = kind == 1 ? 8 + pick(3) : pick(13);
        for (std::size_t c = 0; c < n; ++c) {
            const bool flag = kind == 1 && (c == 0 || (c >= 3 && c <= 7));
            line += (c ? L"," : L"") + (flag ? flags[pick(flags.size())] : texts[pick(texts.size())]);
        }
        return line;
    };

    int kept = 0, grayBefore = 0, emptyOfferBefore = 0, differences = 0;
    for (int round = 0; round < 200000; ++round) {
        std::wstring text;
        const std::size_t rows = pick(6);
        for (std::size_t r = 0; r < rows; ++r) {
            const int kind = r == 0 ? static_cast<int>(pick(3)) : 1 + static_cast<int>(pick(2));
            text += (pick(10) == 0 ? L"" : row(kind)) + eols[pick(eols.size())];
        }

        bool entryKept = false;
        const std::vector<ReplaceItemData> got = ListCodec::clipboardItems(text);
        const std::vector<ReplaceItemData> want = expected(text, entryKept);
        const std::size_t limit = pick(4);
        const std::vector<ReplaceItemData> first = ListCodec::clipboardItems(text, limit);
        const bool ok = same(got, want) && same(first, got, limit < got.size() ? limit : got.size())
            && ListCodec::clipboardItems(text, 1).empty() == got.empty();
        ++g_checks;
        if (!ok) {
            ++g_fail;
            if (++differences <= 5) std::printf("FAIL  random round %d: %zu entries, expected %zu\n", round, got.size(), want.size());
        }

        kept += entryKept ? 1 : 0;
        const bool pastedBefore = !Before::pasted(text).empty();
        grayBefore += !Before::canPaste(text) && pastedBefore ? 1 : 0;
        emptyOfferBefore += Before::canPaste(text) && !pastedBefore ? 1 : 0;
    }
    std::printf("random texts: %d with a copied first entry now kept; before, %d were pasted with the menu entry gray "
        "and %d offered with nothing to paste\n", kept, grayBefore, emptyOfferBefore);

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
