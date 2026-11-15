// Tests for parseNumber (NumberParse.cpp), the text-to-number helper of
// the ExprTk engine, linked directly instead of mirrored. The default
// number output is covered by format_spec_qa (FormatSpec::formatDefault).
// Compile from the repository root with:
//   g++ -std=c++17 -Wall -Wextra -O2 -Isrc/exprtk src/exprtk/NumberParse.cpp
//       src/engine/tests/test_engine_helpers.cpp -o test_engine_helpers

#include "NumberParse.h"

#include <cmath>
#include <iostream>
#include <limits>
#include <string>

using MultiReplaceEngine::parseNumber;

static int g_pass = 0;
static int g_fail = 0;

static void check_parse(const std::string& input, double expected)
{
    const double got = parseNumber(input);
    const bool ok = (std::isnan(expected) && std::isnan(got)) || got == expected;
    if (ok) {
        ++g_pass;
        std::cout << "[PASS] parse(\"" << input << "\") = " << got << "\n";
    }
    else {
        ++g_fail;
        std::cout << "[FAIL] parse(\"" << input << "\") = " << got
                  << " (expected " << expected << ")\n";
    }
}

int main()
{
    const double nan = std::numeric_limits<double>::quiet_NaN();

    std::cout << "=== parseNumber ===\n";

    // Plain numbers, both decimal separators
    check_parse("0", 0.0);
    check_parse("42", 42.0);
    check_parse("-42", -42.0);
    check_parse("3.14", 3.14);
    check_parse("3,14", 3.14);
    check_parse("-7,5", -7.5);
    check_parse(".5", 0.5);
    check_parse("5.", 5.0);

    // Scientific notation
    check_parse("1e10", 1e10);
    check_parse("1.5e-3", 1.5e-3);
    check_parse("-2.5e+5", -2.5e5);
    check_parse("3e", 3.0);

    // Leading spaces and a sign
    check_parse(" 42", 42.0);
    check_parse("\t+5", 5.0);
    check_parse("+5", 5.0);
    check_parse("+-5", nan);
    check_parse("- 5", nan);

    // Trailing text is ignored, a separator in it does not count
    check_parse("3,14abc", 3.14);
    check_parse("1.5, rest", 1.5);
    check_parse("12,5 kg, stock", 12.5);

    // More than one separator is ambiguous
    check_parse("1,234,567", nan);
    check_parse("1.000.000", nan);
    check_parse("1.234,56", nan);
    check_parse("1,234.56", nan);

    // Not a number
    check_parse("", nan);
    check_parse("abc", nan);
    check_parse("$100", nan);
    check_parse("inf", nan);
    check_parse("Info", nan);
    check_parse("nan", nan);
    check_parse(".", nan);
    check_parse("-", nan);
    check_parse("1e999", nan);

    std::cout << "\n=== summary ===\n";
    std::cout << "passed: " << g_pass << "\n";
    std::cout << "failed: " << g_fail << "\n";

    return g_fail == 0 ? 0 : 1;
}
