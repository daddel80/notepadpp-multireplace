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

// QA for NumberRanges. Up to the upper limit it reads every input exactly like parseNumberRanges of the versions
// before (reference below, without its status message), also the quirks of std::stoi. Above the limit it refuses,
// and a range is tested without listing its numbers.
// Build: g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. number_ranges_qa.cpp -o number_ranges_qa
#include <cstdio>
#include <random>
#include <string>
#include <unordered_set>
#include <vector>

#include "../NumberRanges.h"

namespace {

    int g_checks = 0;
    int g_fail = 0;

    void check(bool ok, const std::string& what)
    {
        ++g_checks;
        if (!ok) {
            ++g_fail;
            if (g_fail <= 30) std::printf("FAIL  %s\n", what.c_str());
        }
    }

    std::string narrow(const std::wstring& text)
    {
        return std::string(text.begin(), text.end());
    }

    enum class Before { Ok, Error, TooLarge };

    // parseNumberRanges as it was: the numbers in typed order, each once; Error where it showed its error.
    // TooLarge stops it before it lists a number above guard, which could take it hours
    Before reference(const std::wstring& input, int guard, std::vector<int>& result)
    {
        result.clear();
        if (input.empty()) return Before::Ok;
        bool tooLarge = false;
        std::unordered_set<int> seen;
        std::wistringstream stream(input);
        std::wstring token;
        auto pushUnique = [&](int n) { if (seen.insert(n).second) result.push_back(n); };
        auto processToken = [&](const std::wstring& token) -> bool {
            if (token.empty()) return true;
            try {
                size_t dashPos = token.find(L'-');
                if (dashPos != std::wstring::npos) {
                    int startRange = std::stoi(token.substr(0, dashPos));
                    int endRange = std::stoi(token.substr(dashPos + 1));
                    if (startRange < 1 || endRange < 1) return false;
                    if (startRange > guard || endRange > guard) { tooLarge = true; return false; }
                    if (endRange >= startRange) { for (int i = startRange; i <= endRange; ++i) pushUnique(i); }
                    else { for (int i = startRange; i >= endRange; --i) pushUnique(i); }
                }
                else {
                    int number = std::stoi(token);
                    if (number < 1) return false;
                    if (number > guard) { tooLarge = true; return false; }
                    pushUnique(number);
                }
            }
            catch (const std::exception&) {
                return false;
            }
            return true;
        };
        while (std::getline(stream, token, L',')) {
            if (!processToken(token)) { result.clear(); return tooLarge ? Before::TooLarge : Before::Error; }
        }
        return Before::Ok;
    }

    bool parsed(const std::wstring& input, int maxValue, std::vector<int>& numbers)
    {
        std::vector<NumberRanges::Range> ranges;
        const bool ok = NumberRanges::parse(input, maxValue, ranges);
        numbers = NumberRanges::expand(ranges);
        return ok;
    }

    std::wstring randomInput(std::mt19937& rng)
    {
        static const std::wstring parts[] = { L"1", L"2", L"7", L"12", L"30", L"0", L"999", L"00", L"-", L"-", L",", L",",
            L" ", L"+", L"x", L"\t", L"--", L"3-1", L"5 - 9" };
        std::wstring text;
        const int count = std::uniform_int_distribution<int>(0, 8)(rng);
        for (int i = 0; i < count; ++i) text += parts[std::uniform_int_distribution<int>(0, 18)(rng)];
        return text;
    }
}

int main()
{
    // ---- Up to the limit: the same as before, value for value
    const std::wstring fixed[] = { L"", L"1", L"1-50", L"1,3,5-7", L"10-8", L"1-3,2-4", L" 1 , 3 - 5 ", L"1,,3", L",", L",1,",
        L"1-2-3", L"5abc", L"+5", L"0", L"-5", L"5-", L"-", L"1-0", L"x", L"1, x", L"99999999999", L"2-2", L"3,3,3", L"12-10,9" };
    for (const std::wstring& input : fixed) {
        std::vector<int> before, now;
        const Before was = reference(input, 1000, before);
        const bool okNow = parsed(input, 1000, now);
        check(was != Before::TooLarge && (was == Before::Ok) == okNow && before == now, "as before: \"" + narrow(input) + "\"");
    }

    std::mt19937 rng(20261121u);
    for (int round = 0; round < 200000; ++round) {
        const std::wstring input = randomInput(rng);
        std::vector<int> before, now;
        const Before was = reference(input, 2000, before);
        const bool okNow = parsed(input, 2000, now);
        const bool same = was == Before::TooLarge ? !okNow && now.empty() : (was == Before::Ok) == okNow && before == now;
        if (!same) check(false, "as before, random: \"" + narrow(input) + "\"");
        else ++g_checks;
    }

    // ---- Above the limit: refused, nothing listed
    std::vector<int> numbers;
    check(parsed(L"1-16384", 16384, numbers) && numbers.size() == 16384 && numbers.back() == 16384, "up to the limit");
    check(!parsed(L"1-16385", 16384, numbers) && numbers.empty(), "a range above the limit");
    check(!parsed(L"16385", 16384, numbers) && numbers.empty(), "a number above the limit");
    check(!parsed(L"16385-1", 16384, numbers) && numbers.empty(), "a range down from above the limit");
    check(!parsed(L"1,2,1-999999999", 16384, numbers) && numbers.empty(), "the whole list is refused");

    // ---- A range is tested as a range
    std::vector<NumberRanges::Range> ranges;
    check(NumberRanges::parse(L"1-999999999", NumberRanges::noLimit, ranges) && ranges.size() == 1, "1-999999999 is one range");
    check(NumberRanges::contains(ranges, 1) && NumberRanges::contains(ranges, 999999999)
        && !NumberRanges::contains(ranges, 1000000000), "its ends belong to it, the next number not");
    check(NumberRanges::parse(L"5-3, 10", NumberRanges::noLimit, ranges) && NumberRanges::contains(ranges, 4)
        && NumberRanges::contains(ranges, 10) && !NumberRanges::contains(ranges, 2) && !NumberRanges::contains(ranges, 6)
        && !NumberRanges::contains(ranges, 9), "a range typed downwards and a single number");
    check(NumberRanges::parse(L"2147483646-2147483647", NumberRanges::noLimit, ranges) && NumberRanges::expand(ranges)
        == std::vector<int>{ 2147483646, 2147483647 }, "listing ends at the largest int without overflow");
    check(NumberRanges::parse(L"", NumberRanges::noLimit, ranges) && ranges.empty() && !NumberRanges::contains(ranges, 1), "an empty list holds nothing");

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
