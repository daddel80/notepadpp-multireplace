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

// QA for ExtendedEscapes::expand: a number escape reads only its own digits and stays literal otherwise, so paths
// like C:\backup keep their text. Random input is compared with the rules of Searching::convertExtendedToString.
// Build: g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. extended_escapes_qa.cpp -o extended_escapes_qa
#include <cstddef>
#include <cstdio>
#include <random>
#include <string>

#include "../ExtendedEscapes.h"

namespace {

    int g_checks = 0;
    int g_fail = 0;

    std::string show(const std::wstring& s)
    {
        std::string out;
        for (const wchar_t c : s) {
            if (c >= 0x20 && c < 0x7F) {
                out.push_back(static_cast<char>(c));
            }
            else {
                char buf[16];
                std::snprintf(buf, sizeof buf, "<%X>", static_cast<unsigned>(c));
                out += buf;
            }
        }
        return out;
    }

    void expect(const std::wstring& input, const std::wstring& expected, const char* what)
    {
        ++g_checks;
        const std::wstring got = ExtendedEscapes::expand(input);
        if (got != expected) {
            ++g_fail;
            if (g_fail <= 20)
                std::printf("FAIL  %s: %s -> %s, expected %s\n", what, show(input).c_str(), show(got).c_str(), show(expected).c_str());
        }
    }

    // Searching::convertExtendedToString with readBase (Notepad++, FindReplaceDlg.cpp) for ASCII input without '`',
    // counting the characters left exactly
    std::wstring nppRules(const std::wstring& query)
    {
        const auto readBase = [](const wchar_t* str, int& value, int base, int size) {
            int temp = 0;
            const wchar_t max = static_cast<wchar_t>(L'0' + base - 1);
            for (int i = 0; i < size; ++i) {
                wchar_t current = str[i];
                if (current >= L'A') {
                    current &= 0xdf;
                    current -= (L'A' - L'0' - 10);
                }
                else if (current > L'9') {
                    return false;
                }
                if (current < L'0' || current > max) return false;
                temp = temp * base + (current - L'0');
            }
            value = temp;
            return true;
        };

        std::wstring result;
        const std::size_t length = query.size();
        for (std::size_t i = 0; i < length; ++i) {
            if (query[i] != L'\\' || i + 1 >= length) {
                result.push_back(query[i]);
                continue;
            }
            const wchar_t current = query[++i];
            int size = 0, base = 0;
            switch (current) {
            case L'r':  result.push_back(L'\r'); continue;
            case L'n':  result.push_back(L'\n'); continue;
            case L'0':  result.push_back(L'\0'); continue;
            case L't':  result.push_back(L'\t'); continue;
            case L'\\': result.push_back(L'\\'); continue;
            case L'b':  size = 8; base = 2;  break;
            case L'o':  size = 3; base = 8;  break;
            case L'd':  size = 3; base = 10; break;
            case L'x':  size = 2; base = 16; break;
            case L'u':  size = 4; base = 16; break;
            default:    break;
            }
            int res = 0;
            if (size > 0 && length - i - 1 >= static_cast<std::size_t>(size) && readBase(query.c_str() + i + 1, res, base, size)) {
                result.push_back(static_cast<wchar_t>(res));
                i += static_cast<std::size_t>(size);
            }
            else {
                result.push_back(L'\\');
                result.push_back(current);
            }
        }
        return result;
    }
}

int main()
{
    // Escapes that read
    expect(L"\\r\\n\\t\\\\\\0", std::wstring(L"\r\n\t\\\0", 5), "control characters");
    expect(L"\\x41\\u20AC\\d065\\o101\\b01000001", L"A\u20AC" L"AAA", "numbers of every base");
    expect(L"\\xaF\\xBd", std::wstring{ wchar_t(0xAF), wchar_t(0xBD) }, "hex digits in both cases");
    expect(L"\\d999\\o777\\uFFFF", std::wstring{ wchar_t(999), wchar_t(511), wchar_t(0xFFFF) }, "largest values");
    expect(L"\\x414", L"A4", "digits after the number stay text");
    expect(L"\\\\n", L"\\n", "an escaped backslash before a letter");

    // A number that does not read stays literal and does not try another base
    expect(L"\\backup", L"\\backup", "\\b is no hex escape");
    expect(L"\\office", L"\\office", "\\o is no hex escape");
    expect(L"\\debug", L"\\debug", "\\d is no hex escape");
    expect(L"\\default", L"\\default", "\\d is no hex escape (2)");
    expect(L"\\dead", L"\\dead", "\\d is no hex escape (3)");
    expect(L"\\b100", L"\\b100", "\\b with three digits is no decimal");
    expect(L"\\o189", L"\\o189", "\\o with 8 or 9 is no decimal");
    expect(L"C:\\backup\\data\\office", L"C:\\backup\\data\\office", "a Windows path");
    expect(L"\\b0101010", L"\\b0101010", "\\b with seven digits");
    expect(L"\\b01010102", L"\\b01010102", "\\b with a non-binary digit");
    expect(L"\\d0x1", L"\\d0x1", "\\d with a letter");
    expect(L"\\bx\\n", L"\\bx\n", "the next escape still reads");

    // Incomplete and unknown sequences
    expect(L"\\x4", L"\\x4", "\\x with one digit");
    expect(L"\\u12", L"\\u12", "\\u with two digits");
    expect(L"\\", L"\\", "a lone backslash");
    expect(L"a\\", L"a\\", "a trailing backslash");
    expect(L"\\x41\\", L"A\\", "a trailing backslash after a number");
    expect(L"\\q\\N", L"\\q\\N", "unknown letters");
    expect(L"", L"", "empty input");

    // Random input against the rules of Notepad++
    const std::wstring alphabet = L"\\\\\\rnt012789afAFgbdoxuq ";
    std::mt19937 rng(20261121u);
    std::uniform_int_distribution<std::size_t> pick(0, alphabet.size() - 1), len(0, 14);
    int differences = 0;
    for (int round = 0; round < 300000; ++round) {
        std::wstring s(len(rng), L' ');
        for (wchar_t& c : s) c = alphabet[pick(rng)];
        const std::wstring got = ExtendedEscapes::expand(s);
        const std::wstring want = nppRules(s);
        ++g_checks;
        if (got != want) {
            ++g_fail;
            if (++differences <= 10)
                std::printf("FAIL  random: %s -> %s, rules %s\n", show(s).c_str(), show(got).c_str(), show(want).c_str());
        }
    }

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
