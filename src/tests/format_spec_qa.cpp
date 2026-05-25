// Test harness for FormatSpec. Builds standalone with the FormatSpec
// library only - no MR, no ExprTk, no Win32. Run from the src/tests dir.
//
// MSVC (x64 Native Tools Command Prompt for VS 2022):
//   cl /std:c++17 /EHsc /I.. format_spec_qa.cpp ..\exprtk\FormatSpec.cpp /Fe:format_spec_qa.exe
//   format_spec_qa.exe
//
// MinGW / g++:
//   g++ -std=c++17 -Wall -Wextra -I.. format_spec_qa.cpp ../exprtk/FormatSpec.cpp -o format_spec_qa
//   ./format_spec_qa
//
// Pass -v for a verbose pass-by-pass listing.

#include "../exprtk/FormatSpec.h"

#include <cstdio>
#include <cstdlib>
#include <iostream>
#include <string>

namespace {

int passed = 0;
int failed = 0;
bool verbose = false;  // toggled via -v command line flag

void check(const std::wstring& specText, double value,
           const std::wstring& expected, const char* label)
{
    auto spec = FormatSpec::parse(specText);
    if (!spec.valid) {
        std::wcout << L"FAIL [" << label << L"] parse error: "
                   << std::wstring(spec.errorMessage.begin(), spec.errorMessage.end())
                   << L"\n";
        ++failed;
        return;
    }
    std::wstring got = FormatSpec::apply(spec, value);
    if (got != expected) {
        std::wcout << L"FAIL [" << label << L"] spec=\"" << specText
                   << L"\" value=" << value
                   << L" expected=\"" << expected
                   << L"\" got=\"" << got << L"\"\n";
        ++failed;
        return;
    }
    if (verbose) {
        std::wprintf(L"PASS  %-14ls  %-14g  -> %-20ls  (%hs)\n",
                     specText.c_str(), value, got.c_str(), label);
    }
    ++passed;
}

void checkText(const std::wstring& specText, const std::string& value,
               const std::wstring& expected, const char* label)
{
    auto spec = FormatSpec::parse(specText);
    if (!spec.valid) {
        std::wcout << L"FAIL [" << label << L"] parse error: "
                   << std::wstring(spec.errorMessage.begin(), spec.errorMessage.end())
                   << L"\n";
        ++failed;
        return;
    }
    std::wstring got = FormatSpec::apply(spec, value);
    if (got != expected) {
        std::wcout << L"FAIL [" << label << L"] spec=\"" << specText
                   << L"\" expected=\"" << expected
                   << L"\" got=\"" << got << L"\"\n";
        ++failed;
        return;
    }
    if (verbose) {
        std::wprintf(L"PASS  %-14ls  (text)          -> %-20ls  (%hs)\n",
                     specText.c_str(), got.c_str(), label);
    }
    ++passed;
}

void checkParseError(const std::wstring& specText, const char* label)
{
    auto spec = FormatSpec::parse(specText);
    if (spec.valid) {
        std::wcout << L"FAIL [" << label << L"] expected parse error for \""
                   << specText << L"\" but parsed OK\n";
        ++failed;
        return;
    }
    if (verbose) {
        std::wprintf(L"PASS  %-14ls                  -> <error>             (%hs: %hs)\n",
                     specText.c_str(), label, spec.errorMessage.c_str());
    }
    ++passed;
}

void checkSplit(const std::string& body,
                const std::string& expectedFormula,
                const std::string& expectedSpec,
                bool expectHasSpec,
                const char* label)
{
    auto s = FormatSpec::splitFormulaSpec(body);
    bool ok = (s.formula == expectedFormula)
           && (s.spec == expectedSpec)
           && (s.hasSpec == expectHasSpec);
    if (!ok) {
        std::wcout << L"FAIL [" << label << L"] body=\""
                   << std::wstring(body.begin(), body.end())
                   << L"\"\n"
                   << L"  expected formula=\""
                   << std::wstring(expectedFormula.begin(), expectedFormula.end())
                   << L"\" spec=\""
                   << std::wstring(expectedSpec.begin(), expectedSpec.end())
                   << L"\" hasSpec=" << expectHasSpec << L"\n"
                   << L"  got      formula=\""
                   << std::wstring(s.formula.begin(), s.formula.end())
                   << L"\" spec=\""
                   << std::wstring(s.spec.begin(), s.spec.end())
                   << L"\" hasSpec=" << s.hasSpec << L"\n";
        ++failed;
        return;
    }
    if (verbose) {
        std::wcout << L"PASS  split \""
                   << std::wstring(body.begin(), body.end())
                   << L"\" -> formula=\""
                   << std::wstring(s.formula.begin(), s.formula.end())
                   << L"\" spec=\""
                   << std::wstring(s.spec.begin(), s.spec.end())
                   << L"\" (" << label << L")\n";
    }
    ++passed;
}

}  // namespace

int main(int argc, char** argv) {
    for (int i = 1; i < argc; ++i) {
        std::string a(argv[i]);
        if (a == "-v" || a == "--verbose") verbose = true;
    }

    if (verbose) {
        std::wprintf(L"%-6ls %-14ls  %-14ls     %-20ls  %ls\n",
                     L"status", L"spec", L"value", L"output", L"note");
        std::wprintf(L"%-6ls %-14ls  %-14ls     %-20ls  %ls\n",
                     L"------", L"------------", L"------------", L"-------------------", L"----");
    }
    // ---- Numeric: type letter only ----
    check(L"f",     3.14,        L"3.140000",     "f default precision");
    check(L"e",     12345.0,     L"1.234500e+04", "e default precision");
    check(L"g",     0.0001,      L"0.0001",       "g small");
    check(L"x",     255.0,       L"ff",           "x");
    check(L"o",     8.0,         L"10",           "o");
    check(L"b",     5.0,         L"101",          "b");
    check(L"b",     0.0,         L"0",            "b zero");

    // ---- Numeric: width ----
    check(L"5f",    3.14,        L"3.140000",     "5f width below printf default");
    check(L"10f",   3.14,        L"  3.140000",   "10f width padding spaces");
    check(L"010f",  3.14,        L"003.140000",   "010f zero pad");
    check(L"8x",    255.0,       L"      ff",     "8x width");
    check(L"08x",   255.0,       L"000000ff",     "08x zero pad");
    check(L"8b",    5.0,         L"     101",     "8b width");
    check(L"08b",   5.0,         L"00000101",     "08b zero pad");

    // ---- Numeric: precision ----
    check(L".2f",   3.14159,     L"3.14",         ".2f");
    check(L".4f",   3.14159,     L"3.1416",       ".4f rounds");
    check(L".0f",   3.7,         L"4",            ".0f rounds");
    check(L"05.2f", 3.14159,     L"03.14",        "05.2f");
    check(L"8.2f",  3.14159,     L"    3.14",     "8.2f");
    check(L".3g",   123456.0,    L"1.23e+05",     ".3g");

    // ---- Numeric: force sign ----
    check(L"+f",    3.14,        L"+3.140000",    "+f positive");
    check(L"+f",    -3.14,       L"-3.140000",    "+f negative");
    check(L"+05.2f", 3.14,       L"+3.14",        "+05.2f positive");

    // ---- Numeric: min-max precision (trailing zero trim) ----
    check(L".2-5f", 3.1,         L"3.10",         ".2-5f keeps min");
    check(L".2-5f", 3.14,        L"3.14",         ".2-5f mid");
    check(L".2-5f", 3.14159,     L"3.14159",      ".2-5f max");
    check(L".2-5f", 3.141592,    L"3.14159",      ".2-5f rounds at max");
    check(L".0-3f", 3.0,         L"3",            ".0-3f strips dot when zero");
    check(L".0-3f", 3.140,       L"3.14",         ".0-3f keeps non-zero");

    // ---- Numeric: default (no type letter) ----
    check(L"+",     3.14,        L"+3.14",        "+ only, default type");
    check(L"+",     -3.14,       L"-3.14",        "+ only negative");
    check(L"08",    3.14,        L"00003.14",     "08 only");
    check(L"8",     3.14,        L"    3.14",     "8 only");

    // ---- Numeric: negative values ----
    check(L"05.2f", -3.14,       L"-3.14",        "neg eats width");
    check(L"06.2f", -3.14,       L"-03.14",       "neg with zeropad");
    check(L"x",     -1.0,        L"ffffffffffffffff", "x negative wraps");

    // ---- Numeric: v2 universal frame (align/fill on numbers) ----
    check(L"5",      123,        L"  123",        "num default align = right");
    check(L"<5",     123,        L"123  ",        "num align left");
    check(L">5",     123,        L"  123",        "num align right explicit");
    check(L"^5",     123,        L" 123 ",        "num align center");
    check(L">5.2f",  3.14159,    L" 3.14",        "float right + precision");
    check(L"<5.2f",  3.14159,    L"3.14 ",        "float left + precision");
    check(L"^5.2f",  3.14159,    L"3.14 ",        "float center + precision");
    check(L"*>5.2f", 3.14159,    L"*3.14",        "float fill '*' right");
    check(L"*<5.2f", 3.14159,    L"3.14*",        "float fill '*' left");
    check(L"<8d",    42,         L"42      ",     "int left");
    checkParseError(L">05d",                      "align + zeropad conflict");

    // ---- Numeric: default type with precision behaves like 'g' ----
    check(L".3",     3.14159,    L"3.14",         "default prec = significant (g)");
    check(L".3",     1234.5678,  L"1.23e+03",     "default prec switches to sci (g)");
    check(L".3",     3.0,        L"3",            "default prec trims zeros (g)");
    check(L".5",     3.14159,    L"3.1416",       "default prec 5 significant");

    // ---- Duration: ts ----
    check(L"ts:hms",  3725.0,    L"1:02:05",      "ts:hms 1h2m5s");
    check(L"ts:hms",  0.0,       L"0:00:00",      "ts:hms zero");
    check(L"ts:ms",   90.0,      L"1:30",         "ts:ms 90sec");
    check(L"ts:ms",   125.0,     L"2:05",         "ts:ms 125sec");
    check(L"ts:hms",  -3725.0,   L"-1:02:05",     "ts:hms negative");

    // ---- Duration: tm ----
    check(L"tm:hm",   90.0,      L"1:30",         "tm:hm 90min");
    check(L"tm:hms",  90.0,      L"1:30:00",      "tm:hms 90min");
    check(L"tm:hm",   480.0,     L"8:00",         "tm:hm 8h workday");

    // ---- Duration: th ----
    check(L"th:hms",  2.5,       L"2:30:00",      "th:hms 2.5h");
    check(L"th:dh",   25.0,      L"1 01",         "th:dh 25h = 1d 1h");

    // ---- Duration: td ----
    check(L"td:dhms", 1.5,       L"1 12:00:00",   "td:dhms 1.5d");
    check(L"td:dhm",  1.5,       L"1 12:00",      "td:dhm 1.5d");

    // ---- v2 universal frame on duration / date ----
    check(L"^12 ts:hms", 3725.0,  L"  1:02:05   ", "dur align center");
    check(L">12 ts:hms", 3725.0,  L"     1:02:05", "dur align right");
    check(L"<12 ts:hms", 3725.0,  L"1:02:05     ", "dur align left");
    check(L"*>12ts:hms", 3725.0,  L"*****1:02:05", "dur fill, no space before marker");
    check(L">15 d:!%Y-%m-%d", 1700000000.0, L"     2023-11-14", "date align right");
    check(L"<15 d:!%Y-%m-%d", 1700000000.0, L"2023-11-14     ", "date align left");

    // ---- v2 marker-free: bare frame on a STRING value (t: dropped) ----
    checkText(L">8",    "abc",   L"     abc",     "bare frame on text, right");
    checkText(L"<8",    "abc",   L"abc     ",     "bare frame on text, left");
    checkText(L"^8",    "abc",   L"  abc   ",     "bare frame on text, center");
    checkText(L"*>8",   "abc",   L"*****abc",     "bare frame on text, fill");
    checkText(L">8.3",  "hello", L"     hel",     "bare frame on text, .3 = truncation");
    // v2 marker in the MIDDLE: frame, then marker, then body (uniform with d:/ts:)
    checkText(L">8 t:",  "abc",   L"     abc",     "t: marker after frame");
    checkText(L">8 t:.3","hello", L"     hel",     "t: after frame + truncation");
    // n: optional == bare (numeric value), marker after the frame
    check(L">8 n:.2f", 3.14159,   L"    3.14",     "n: after frame == bare");
    check(L">8.2f",    3.14159,   L"    3.14",     "bare == n:");
    // Old marker-FIRST spelling is no longer valid (frame must lead)
    checkParseError(L"n:>8.2f", "n: before frame rejected");
    checkParseError(L"t:>8",    "t: before frame rejected");
    checkParseError(L"t:>8.3",  "t: before frame+trunc rejected");
    // n: in the middle, marker is cosmetic: same output as bare, all bodies
    check(L"n:5.2f",  3.14,      L" 3.14",        "n: no frame, body only");
    check(L"<8 n:.2f",3.14159,   L"3.14    ",     "n: after left-align frame");
    check(L"^10 n:.2f",3.14159,  L"   3.14   ",   "n: after center frame");
    check(L"*>8 n:x", 255,       L"******ff",     "n: after fill+align, hex body");
    // align + zero-pad still conflicts whether the align is in the body (>05d)
    // or inherited from a frame before n: (>8 n:05)
    checkParseError(L">05d",     "align+zeropad conflict (body)");
    checkParseError(L">8 n:05",  "align+zeropad conflict (frame+n:)");
    // n:/t: alone (no body, no frame): n: empty is an error, t: is a no-op frame
    checkParseError(L"n:",       "n: with empty body");

    // ---- Parse errors ----
    checkParseError(L"",         "empty");
    checkParseError(L"  ",       "all whitespace");
    checkParseError(L"q",        "unknown type letter");
    checkParseError(L"05qf",     "bad char before type");
    checkParseError(L"05.2fx",   "trailing junk");
    checkParseError(L".f",       "precision dot no digit");
    checkParseError(L".5-",      "range no max");
    checkParseError(L".5-2f",    "range max < min");
    checkParseError(L".2x",      "precision with int type");
    checkParseError(L"+x",       "force-sign with int type");
    checkParseError(L"t",        "duration too short");
    checkParseError(L"tq:hms",   "unknown duration unit");
    checkParseError(L"ts",       "duration no mode");
    checkParseError(L"ts:",      "duration empty mode");
    checkParseError(L"ts:foo",   "unknown duration mode");
    checkParseError(L"tshms",    "duration no colon");

    // ---- More edge cases ----
    check(L"f",     0.0,         L"0.000000",     "f zero");
    check(L"+f",    0.0,         L"+0.000000",    "+f zero stays positive");
    check(L".0f",   0.4,         L"0",            ".0f rounds down");
    check(L".0f",   0.5,         L"0",            ".0f banker's rounding (varies)");
    check(L".2-4f", 1.0,         L"1.00",         ".2-4f integer keeps min");
    check(L".2-4f", 1.23456,     L"1.2346",       ".2-4f at max");
    check(L"x",     16.0,        L"10",           "x small");
    check(L"b",     -1.0,        L"1111111111111111111111111111111111111111111111111111111111111111",  "b -1 all bits");

    // Very large values
    check(L".2f",   1e10,        L"10000000000.00", ".2f large");
    check(L"g",     1234.5,      L"1234.5",       "g default no scientific");

    // Duration edge cases
    check(L"ts:hms",  0.5,       L"0:00:00",      "ts:hms half-second truncates");
    check(L"ts:hms",  86400.0,   L"24:00:00",     "ts:hms one day in hours");
    check(L"td:dhms", 0.0,       L"0 00:00:00",   "td:dhms zero");
    check(L"ts:ms",   3600.0,    L"60:00",        "ts:ms hour in minutes");

    // ---- Splitter tests --------------------------------------------------

    // Basic: no separator
    checkSplit("num(1)",                      "num(1)",           "",       false, "no spec");
    checkSplit("  num(1)  ",                  "num(1)",           "",       false, "no spec with whitespace");
    checkSplit("num(1)*0.9",                  "num(1)*0.9",       "",       false, "expression no spec");

    // Basic: one separator
    checkSplit("num(1) ~ 05.2f",              "num(1)",           "05.2f",  true,  "simple split");
    checkSplit("num(1)*0.9~05.2f",            "num(1)*0.9",       "05.2f",  true,  "no whitespace around tilde");
    checkSplit("  num(1)  ~  05.2f  ",        "num(1)",           "05.2f",  true,  "whitespace around all");

    // Multiple separators - LAST one wins
    checkSplit("a ~ b ~ c",                   "a ~ b",            "c",      true,  "last tilde wins");
    checkSplit("num(1) ~ num(2) ~ 05.2f",     "num(1) ~ num(2)",  "05.2f",  true,  "two tildes formula has one");

    // Quoted strings - tildes inside are NOT separators
    checkSplit("'a~b' ~ 05.2f",               "'a~b'",            "05.2f",  true,  "tilde in single-quoted string");
    checkSplit("\"a~b\" ~ 05.2f",             "\"a~b\"",          "05.2f",  true,  "tilde in double-quoted string");
    checkSplit("'a~b'",                       "'a~b'",            "",       false, "only tilde inside string");
    checkSplit("'~' ~ '~~' ~ 05.2f",          "'~' ~ '~~'",       "05.2f",  true,  "many tildes in strings");

    // Escaped quotes inside strings
    checkSplit("'a\\'b' ~ 05.2f",             "'a\\'b'",          "05.2f",  true,  "escaped single quote");
    checkSplit("\"a\\\"b\" ~ 05.2f",          "\"a\\\"b\"",       "05.2f",  true,  "escaped double quote");

    // Edge: only tilde, no formula
    checkSplit("~ 05.2f",                     "",                 "05.2f",  true,  "empty formula");

    // Edge: trailing tilde, no spec
    checkSplit("num(1) ~",                    "num(1)",           "",       true,  "trailing tilde empty spec");

    // Empty input
    checkSplit("",                            "",                 "",       false, "empty input");

    // ---- Date tests (UTC for determinism) ------------------------------

    // Known anchor: 1700000000 = 2023-11-14 22:13:20 UTC
    check(L"d:!%Y-%m-%d",        1700000000.0, L"2023-11-14",         "date UTC y-m-d");
    check(L"d:!%H:%M:%S",        1700000000.0, L"22:13:20",           "date UTC h:m:s");
    check(L"d:!%Y-%m-%d %H:%M:%S", 1700000000.0, L"2023-11-14 22:13:20","date UTC datetime");

    // Unix epoch
    check(L"d:!%Y-%m-%d",        0.0,          L"1970-01-01",         "date UTC epoch");
    check(L"d:!%Y-%m-%dT%H:%M:%SZ", 0.0,       L"1970-01-01T00:00:00Z","date UTC ISO 8601");

    // Subsecond truncation: 1.9 still equals 1970-01-01 00:00:01 (truncation, not rounding)
    check(L"d:!%H:%M:%S",        1.9,          L"00:00:01",           "date subsecond truncation");

    // Literal text inside the strftime pattern
    check(L"d:!Year %Y",         1700000000.0, L"Year 2023",          "date literal text");
    check(L"d:!%d.%m.%Y",        1700000000.0, L"14.11.2023",         "date european format");

    // Locale-dependent (just check non-empty, exact value depends on test runner TZ)
    {
        auto sp = FormatSpec::parse(L"d:%Y-%m-%d");
        if (sp.valid && !FormatSpec::apply(sp, 1700000000.0).empty()) {
            ++passed;
            if (verbose) {
                std::wprintf(L"PASS  %-14ls  %-14g  -> (non-empty local)   (date local time path)\n",
                             L"d:%Y-%m-%d", 1700000000.0);
            }
        } else {
            std::wcout << L"FAIL [date local time path]\n";
            ++failed;
        }
    }

    // Date parse errors
    checkParseError(L"d:",                   "date empty pattern");
    checkParseError(L"d:!",                  "date only utc marker");

    // Date negative timestamp returns empty (defensive, would have been
    // routed through dialog upstream).
    {
        auto sp = FormatSpec::parse(L"d:!%Y-%m-%d");
        std::wstring r = FormatSpec::apply(sp, -1.0);
        if (r.empty()) {
            ++passed;
            if (verbose) {
                std::wprintf(L"PASS  d:!%%Y-%%m-%%d   -1              -> <empty>             (date negative returns empty)\n");
            }
        } else {
            std::wcout << L"FAIL [date negative] got: " << r << L"\n";
            ++failed;
        }
    }

    // More date edge cases
    check(L"d:!%A",              0.0,          L"Thursday",           "date weekday name (epoch=Thu)");
    check(L"d:!%B",              0.0,          L"January",            "date month name");
    check(L"d:!%j",              0.0,          L"001",                "date day of year");

    // Y2038 boundary: 2147483647 = 2038-01-19 03:14:07 UTC (last 32-bit signed)
    check(L"d:!%Y-%m-%d %H:%M:%S", 2147483647.0, L"2038-01-19 03:14:07","date Y2038 boundary");

    // Brackets in pattern - test that strftime literals work
    check(L"d:!(%Y)",            1700000000.0, L"(2023)",             "date parens in pattern");

    // ---- Result summary ----
    std::wcout << L"\n========================================\n";
    std::wcout << L"PASSED: " << passed << L"\n";
    std::wcout << L"FAILED: " << failed << L"\n";
    std::wcout << L"========================================\n";
    return failed == 0 ? 0 : 1;
}