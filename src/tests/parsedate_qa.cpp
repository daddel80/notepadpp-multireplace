// Standalone tests for DateParse.
// Compile:
//   g++ -std=c++20 -Wall -Wextra parsedate_qa.cpp ../fs/DateParse.cpp -o parsedate_qa
//   ./parsedate_qa [-v]

#include "../exprtk/DateParse.h"

#include <cstdio>
#include <cstring>
#include <ctime>
#include <iostream>
#include <string>
#include <string_view>

namespace {

int passed = 0;
int failed = 0;
bool verbose = false;

void checkFields(std::string_view input, std::string_view fmt,
                 int year, int mon0based, int mday,
                 int hour, int min, int sec,
                 const char* label)
{
    std::tm tm{};
    const bool ok = MultiReplace::parseDateTime(input, fmt, tm);
    if (!ok) {
        std::printf("FAIL [%s] parse returned false\n", label);
        ++failed;
        return;
    }
    if (tm.tm_year + 1900 != year || tm.tm_mon != mon0based
        || tm.tm_mday != mday || tm.tm_hour != hour
        || tm.tm_min != min || tm.tm_sec != sec)
    {
        std::printf("FAIL [%s] expected %d-%02d-%02d %02d:%02d:%02d  "
                    "got %d-%02d-%02d %02d:%02d:%02d\n",
                    label, year, mon0based + 1, mday, hour, min, sec,
                    tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday,
                    tm.tm_hour, tm.tm_min, tm.tm_sec);
        ++failed;
        return;
    }
    if (verbose) {
        std::printf("PASS  \"%-26.*s\" fmt=\"%-18.*s\" -> %d-%02d-%02d %02d:%02d:%02d  (%s)\n",
                    (int)input.size(), input.data(),
                    (int)fmt.size(), fmt.data(),
                    year, mon0based + 1, mday, hour, min, sec, label);
    }
    ++passed;
}

void checkFail(std::string_view input, std::string_view fmt, const char* label)
{
    std::tm tm{};
    const bool ok = MultiReplace::parseDateTime(input, fmt, tm);
    if (ok) {
        std::printf("FAIL [%s] expected parse failure for \"%.*s\" but it succeeded\n",
                    label, (int)input.size(), input.data());
        ++failed;
        return;
    }
    if (verbose) {
        std::printf("PASS  \"%-26.*s\" fmt=\"%-18.*s\" -> <fail>                  (%s)\n",
                    (int)input.size(), input.data(),
                    (int)fmt.size(), fmt.data(), label);
    }
    ++passed;
}

}  // namespace

int main(int argc, char** argv) {
    for (int i = 1; i < argc; ++i) {
        if (std::strcmp(argv[i], "-v") == 0
            || std::strcmp(argv[i], "--verbose") == 0)
        {
            verbose = true;
        }
    }

    // ---- ISO date ----
    checkFields("2023-11-14", "%Y-%m-%d",       2023, 10, 14, 0, 0, 0, "ISO date");
    checkFields("1970-01-01", "%Y-%m-%d",       1970,  0,  1, 0, 0, 0, "Unix epoch date");
    checkFields("9999-12-31", "%Y-%m-%d",       9999, 11, 31, 0, 0, 0, "max year");

    // ---- ISO datetime ----
    checkFields("2023-11-14 22:13:20", "%Y-%m-%d %H:%M:%S",
                2023, 10, 14, 22, 13, 20, "ISO datetime");
    checkFields("2023-11-14T22:13:20", "%Y-%m-%dT%H:%M:%S",
                2023, 10, 14, 22, 13, 20, "ISO 8601 T separator");
    checkFields("2023-11-14T22:13:20Z", "%Y-%m-%dT%H:%M:%SZ",
                2023, 10, 14, 22, 13, 20, "ISO 8601 with Z");

    // ---- European format ----
    checkFields("14.11.2023", "%d.%m.%Y",       2023, 10, 14, 0, 0, 0, "DE date DD.MM.YYYY");
    checkFields("14/11/2023", "%d/%m/%Y",       2023, 10, 14, 0, 0, 0, "EU date DD/MM/YYYY");
    checkFields("14.11.2023 22:13:20", "%d.%m.%Y %H:%M:%S",
                2023, 10, 14, 22, 13, 20, "DE datetime");

    // ---- US format ----
    checkFields("11/14/2023", "%m/%d/%Y",       2023, 10, 14, 0, 0, 0, "US date MM/DD/YYYY");

    // ---- Shortcuts %F and %T ----
    checkFields("2023-11-14", "%F",             2023, 10, 14, 0, 0, 0, "%F shortcut");
    checkFields("22:13:20", "%T",               1900, 0, 0, 22, 13, 20, "%T shortcut");
    checkFields("2023-11-14 22:13:20", "%F %T",
                2023, 10, 14, 22, 13, 20, "%F %T combined");

    // ---- Two-digit year ----
    checkFields("23-11-14", "%y-%m-%d",         2023, 10, 14, 0, 0, 0, "year 23 -> 2023");
    checkFields("68-01-01", "%y-%m-%d",         2068, 0, 1, 0, 0, 0, "year 68 -> 2068");
    checkFields("69-01-01", "%y-%m-%d",         1969, 0, 1, 0, 0, 0, "year 69 -> 1969");
    checkFields("99-12-31", "%y-%m-%d",         1999, 11, 31, 0, 0, 0, "year 99 -> 1999");
    checkFields("00-01-01", "%y-%m-%d",         2000, 0, 1, 0, 0, 0, "year 00 -> 2000");

    // ---- AM/PM with %I ----
    checkFields("10:30:00 AM", "%I:%M:%S %p",   1900, 0, 0, 10, 30, 0, "10:30 AM");
    checkFields("10:30:00 PM", "%I:%M:%S %p",   1900, 0, 0, 22, 30, 0, "10:30 PM");
    checkFields("12:00:00 AM", "%I:%M:%S %p",   1900, 0, 0,  0,  0, 0, "12:00 AM midnight");
    checkFields("12:00:00 PM", "%I:%M:%S %p",   1900, 0, 0, 12,  0, 0, "12:00 PM noon");
    checkFields("01:00:00 AM", "%I:%M:%S %p",   1900, 0, 0,  1,  0, 0, "01:00 AM");
    checkFields("11:59:59 PM", "%I:%M:%S %p",   1900, 0, 0, 23, 59, 59, "11:59:59 PM");
    checkFields("10:30:00 am", "%I:%M:%S %p",   1900, 0, 0, 10, 30, 0, "lowercase am");
    checkFields("10:30:00 pM", "%I:%M:%S %p",   1900, 0, 0, 22, 30, 0, "mixed case pM");

    // ---- Whitespace flexibility ----
    checkFields("2023-11-14   22:13:20", "%Y-%m-%d %H:%M:%S",
                2023, 10, 14, 22, 13, 20, "extra whitespace in input");
    checkFields("2023-11-14\t22:13:20", "%Y-%m-%d %H:%M:%S",
                2023, 10, 14, 22, 13, 20, "tab as whitespace");

    // ---- Trailing input is OK ----
    checkFields("2023-11-14 trailing junk", "%Y-%m-%d",
                2023, 10, 14, 0, 0, 0, "trailing input ignored");

    // ---- Failures ----
    checkFail("not-a-date", "%Y-%m-%d",     "complete garbage");
    checkFail("2023-13-14", "%Y-%m-%d",     "month 13");
    checkFail("2023-00-14", "%Y-%m-%d",     "month 0");
    checkFail("2023-11-32", "%Y-%m-%d",     "day 32");
    checkFail("2023-11-00", "%Y-%m-%d",     "day 0");
    checkFail("25:00:00", "%H:%M:%S",       "hour 25");
    checkFail("10:60:00", "%H:%M:%S",       "min 60");
    checkFail("00:00:61", "%H:%M:%S",       "sec 61");
    checkFail("2023-11", "%Y-%m-%d",        "incomplete input");
    checkFail("2023.11.14", "%Y-%m-%d",     "wrong separator");
    checkFail("ABC", "%Y",                  "non-digits for %Y");
    checkFail("13:00:00 AM", "%I:%M:%S %p", "12-hour out of range");
    checkFail("10:00:00 XM", "%I:%M:%S %p", "bad AM/PM token");
    checkFail("2023%", "%Y%",               "trailing percent in format");
    checkFail("2023", "%Q",                 "unknown specifier");

    // ---- Result ----
    std::printf("\n========================================\n");
    std::printf("PASSED: %d\n", passed);
    std::printf("FAILED: %d\n", failed);
    std::printf("========================================\n");
    return failed == 0 ? 0 : 1;
}