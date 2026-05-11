// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.

#include "CsvListFormat.h"
#include "StringUtils.h"

namespace CsvListFormat {

    // ---- Row tokenizer ------------------------------------------------

    std::vector<std::wstring> parseLine(const std::wstring& line) {
        std::wstring cleanLine = line;
        while (!cleanLine.empty() && (cleanLine.back() == L'\r' || cleanLine.back() == L'\n')) {
            cleanLine.pop_back();
        }

        std::vector<std::wstring> columns;
        std::wstring current;
        bool insideQuotes = false;

        for (size_t i = 0; i < cleanLine.size(); ++i) {
            wchar_t ch = cleanLine[i];
            if (ch == L'"') {
                if (insideQuotes && i + 1 < cleanLine.size() && cleanLine[i + 1] == L'"') {
                    current += L'"';
                    ++i;
                }
                else {
                    insideQuotes = !insideQuotes;
                }
            }
            else if (ch == L',' && !insideQuotes) {
                columns.push_back(StringUtils::unescapeQuoted(current));
                current.clear();
            }
            else {
                current += ch;
            }
        }
        columns.push_back(StringUtils::unescapeQuoted(current));
        return columns;
    }

    // ---- Header parsing -----------------------------------------------

    std::wstring normalizeName(const std::wstring& raw) {
        std::wstring s;
        s.reserve(raw.size());
        for (wchar_t c : raw) {
            if (c == L' ' || c == L'\t' || c == L'_' || c == L'-') continue;
            if (c >= L'A' && c <= L'Z') s.push_back(c + (L'a' - L'A'));
            else s.push_back(c);
        }
        return s;
    }

    Field fieldFromName(const std::wstring& normalized) {
        if (normalized == L"selected")       return Field::Selected;
        if (normalized == L"find")           return Field::Find;
        if (normalized == L"replace")        return Field::Replace;
        if (normalized == L"wholeword")      return Field::WholeWord;
        if (normalized == L"matchcase")      return Field::MatchCase;
        if (normalized == L"formulasupport") return Field::FormulaSupport;
        if (normalized == L"usevariables")   return Field::FormulaSupport; // legacy
        if (normalized == L"extended")       return Field::Extended;
        if (normalized == L"regex")          return Field::Regex;
        if (normalized == L"comments")       return Field::Comments;
        if (normalized == L"lastmodified")   return Field::LastModified;
        return Field::Unknown;
    }

    HeaderIndex buildIndex(const std::vector<std::wstring>& headerCells) {
        HeaderIndex m;
        m.idx.fill(-1);

        int recognized = 0;
        for (size_t i = 0; i < headerCells.size(); ++i) {
            const Field f = fieldFromName(normalizeName(headerCells[i]));
            if (f == Field::Unknown) continue;
            ++recognized;
            const int slot = static_cast<int>(f);
            if (m.idx[slot] < 0) m.idx[slot] = static_cast<int>(i); // first wins
            if (f == Field::Find)    m.hasFind = true;
            if (f == Field::Replace) m.hasReplace = true;
        }
        m.looksLikeNames = m.hasFind || (recognized >= 2);
        return m;
    }

    std::wstring cellAt(const HeaderIndex& m,
        const std::vector<std::wstring>& row,
        Field f,
        const std::wstring& def) {
        const int col = m.idx[static_cast<int>(f)];
        if (col < 0) return def;
        if (static_cast<size_t>(col) >= row.size()) return def;
        return row[col];
    }

    bool cellAtBool(const HeaderIndex& m,
        const std::vector<std::wstring>& row,
        Field f,
        bool def) {
        const int col = m.idx[static_cast<int>(f)];
        if (col < 0) return def;
        if (static_cast<size_t>(col) >= row.size()) return def;
        const std::wstring& v = row[col];
        if (v.empty()) return def;
        try { return std::stoi(v) != 0; }
        catch (...) { return def; }
    }

}  // namespace CsvListFormat