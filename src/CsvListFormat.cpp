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

    std::vector<std::vector<std::wstring>> readRecords(const std::wstring& text,
        Dialect dialect, wchar_t delimiter) {
        std::vector<std::vector<std::wstring>> records;
        std::vector<std::wstring> cells;
        std::wstring current;
        bool insideQuotes = false;
        bool pending = false;  // a cell/record is being built and must be flushed

        auto decode = [&](const std::wstring& raw) {
            // RFC body: unquote, then bring real in-cell line breaks to the
            // internal CRLF convention. MR (escaped) body keeps its own escapes.
            return dialect == Dialect::Rfc4180
                ? StringUtils::newlinesToCrlf(StringUtils::rfcUnquote(raw))
                : StringUtils::unescapeQuoted(raw);
            };
        auto endCell = [&]() {
            cells.push_back(decode(current));
            current.clear();
            };
        auto endRecord = [&]() {
            endCell();
            records.push_back(cells);
            cells.clear();
            pending = false;
            };

        for (size_t i = 0; i < text.size(); ++i) {
            wchar_t ch = text[i];
            if (ch == L'"') {
                pending = true;
                current += ch;  // quotes stay in the raw cell, decoded at cell end
                if (insideQuotes && i + 1 < text.size() && text[i + 1] == L'"') {
                    current += text[++i];  // doubled quote: keep the second too
                }
                else {
                    insideQuotes = !insideQuotes;
                }
            }
            else if (!insideQuotes && (ch == L'\n' || ch == L'\r')) {
                if (ch == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n') ++i;
                endRecord();
            }
            else if (ch == delimiter && !insideQuotes) {
                pending = true;
                endCell();
            }
            else {
                pending = true;
                current += ch;
            }
        }
        // Trailing record with no closing newline; an empty tail is not a record.
        if (pending || !cells.empty() || !current.empty()) endRecord();
        return records;
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

    wchar_t detectDelimiter(const std::wstring& text, wchar_t fallback) {
        // Read only the first record with each candidate; a delimiter that
        // makes the header resolve to recognizable column names is the one
        // the file actually uses. Plain CSV, so Rfc4180 decoding.
        for (wchar_t cand : { L',', L';' }) {
            std::vector<std::vector<std::wstring>> recs =
                readRecords(text, Dialect::Rfc4180, cand);
            if (recs.empty()) continue;
            if (buildIndex(recs.front()).looksLikeNames) return cand;
        }
        return fallback;
    }

}  // namespace CsvListFormat