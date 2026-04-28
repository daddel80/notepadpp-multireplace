// Standalone tests for the ExprTk-independent helpers in
// ExprTkEngine: parseCaptureToDouble and formatDouble. These do not
// link the full engine; they re-implement the helpers in this TU so
// they can be verified without the ExprTk header.
//
// The intent is to lock in the contract for these helpers before the
// full engine is built, so the engine code can rely on documented
// behaviour. When the full engine compiles successfully on Thomas's
// box, we'll also build proper end-to-end tests.

#include <array>
#include <cassert>
#include <charconv>
#include <cmath>
#include <cstring>
#include <iostream>
#include <iomanip>
#include <limits>
#include <sstream>
#include <string>
#include <system_error>
#include <vector>

// ---- mirror of ExprTkEngine::parseCaptureToDouble -------------------

static double parseCaptureToDouble(const std::string& s)
{
    if (s.empty()) {
        return 0.0;
    }

    double value = 0.0;
    const char* first = s.data();
    const char* last  = s.data() + s.size();

    auto res = std::from_chars(first, last, value);
    if (res.ec != std::errc{}) {
        return 0.0;
    }
    return value;
}

// ---- mirror of ExprTkEngine::formatDouble ---------------------------

static std::string formatDouble(double value)
{
    if (std::isnan(value)) {
        return "nan";
    }
    if (std::isinf(value)) {
        return value < 0.0 ? "-inf" : "inf";
    }

    std::array<char, 64> buf{};
    auto res = std::to_chars(buf.data(), buf.data() + buf.size(), value);
    if (res.ec != std::errc{}) {
        return "nan";
    }
    return std::string(buf.data(), res.ptr);
}

// ---- mirror of RegFunction's index logic ----------------------------

static double regLookup(const std::vector<double>& caps, double index)
{
    if (!std::isfinite(index)) {
        return 0.0;
    }
    const long long idx = static_cast<long long>(index);
    if (idx < 0) {
        return 0.0;
    }
    if (static_cast<std::size_t>(idx) >= caps.size()) {
        return 0.0;
    }
    return caps[static_cast<std::size_t>(idx)];
}

// ---- test infrastructure --------------------------------------------

static int g_pass = 0;
static int g_fail = 0;

static void check_parse(const std::string& input, double expected)
{
    double got = parseCaptureToDouble(input);
    bool ok = (std::isnan(expected) && std::isnan(got))
           || (got == expected);
    if (ok) {
        ++g_pass;
        std::cout << "[PASS] parse(\"" << input << "\") = " << got << "\n";
    } else {
        ++g_fail;
        std::cout << "[FAIL] parse(\"" << input << "\") = " << got
                  << " (expected " << expected << ")\n";
    }
}

static void check_format(double value, const std::string& expected)
{
    std::string got = formatDouble(value);
    if (got == expected) {
        ++g_pass;
        std::cout << "[PASS] format(" << value << ") = \"" << got << "\"\n";
    } else {
        ++g_fail;
        std::cout << "[FAIL] format(" << value << ") = \"" << got
                  << "\" (expected \"" << expected << "\")\n";
    }
}

static void check_reg(const std::vector<double>& caps, double idx, double expected)
{
    double got = regLookup(caps, idx);
    bool ok = (std::isnan(expected) && std::isnan(got))
           || (got == expected);
    if (ok) {
        ++g_pass;
        std::cout << "[PASS] reg(" << idx << ") = " << got << "\n";
    } else {
        ++g_fail;
        std::cout << "[FAIL] reg(" << idx << ") = " << got
                  << " (expected " << expected << ")\n";
    }
}

// =====================================================================

int main()
{
    std::cout << "=== parseCaptureToDouble tests ===\n";

    // Empty -> 0
    check_parse("", 0.0);

    // Plain integers
    check_parse("0", 0.0);
    check_parse("1", 1.0);
    check_parse("42", 42.0);
    check_parse("-1", -1.0);
    check_parse("-42", -42.0);

    // Decimals (locale-independent: '.' only)
    check_parse("1.5", 1.5);
    check_parse("0.5", 0.5);
    check_parse("-3.14", -3.14);
    check_parse("100.001", 100.001);

    // Scientific notation
    check_parse("1e10", 1e10);
    check_parse("1.5e-3", 1.5e-3);
    check_parse("-2.5e+5", -2.5e5);

    // Leading whitespace -> NOT consumed by from_chars by default,
    // so this fails -> 0.0. This is a design choice: we accept it as
    // a non-numeric capture rather than trimming.
    check_parse(" 42", 0.0);

    // Trailing junk -> consumes the numeric prefix, returns the
    // parsed value. (from_chars sets res.ec=success even with junk
    // at the tail, since it was able to parse SOMETHING.)
    check_parse("1.5abc", 1.5);
    check_parse("42xyz",  42.0);

    // Comma as decimal: fails -> 0 (we explicitly want this for
    // locale-robustness)
    check_parse("1,5", 1.0);   // parses "1", stops at ","

    // Pure non-numeric
    check_parse("abc", 0.0);
    check_parse("hello world", 0.0);

    // Edge: just a sign
    check_parse("-", 0.0);
    check_parse("+", 0.0);

    // Edge: just a dot
    check_parse(".", 0.0);

    // Edge: dot+digit (some implementations require leading 0)
    check_parse(".5", 0.5);

    std::cout << "\n=== formatDouble tests ===\n";

    // Whole numbers -> no decimal
    check_format(0.0, "0");
    check_format(1.0, "1");
    check_format(42.0, "42");
    check_format(-1.0, "-1");

    // Decimals -> shortest
    check_format(1.5, "1.5");
    check_format(0.5, "0.5");
    check_format(-3.14, "-3.14");

    // Special floats (we override the default representation)
    check_format(std::nan(""), "nan");
    check_format(std::numeric_limits<double>::infinity(), "inf");
    check_format(-std::numeric_limits<double>::infinity(), "-inf");

    // Scientific notation kicks in for extreme magnitudes -
    // exact spelling is impl-defined ("1e+20" vs "1e20"), so we
    // just confirm it's not absurdly long.
    {
        std::string s = formatDouble(1e20);
        bool ok = (s.size() < 25) && (s.find('e') != std::string::npos
                                   || s.size() < 22);
        if (ok) { ++g_pass; std::cout << "[PASS] format(1e20) = \"" << s << "\"\n"; }
        else    { ++g_fail; std::cout << "[FAIL] format(1e20) too long: \"" << s << "\"\n"; }
    }

    std::cout << "\n=== regLookup tests ===\n";

    std::vector<double> caps = { 100.0, 1.0, 2.0, 3.0 };
    // index 0 = "match", 1..3 = capture groups

    check_reg(caps, 0, 100.0);
    check_reg(caps, 1, 1.0);
    check_reg(caps, 2, 2.0);
    check_reg(caps, 3, 3.0);

    // Out-of-range -> 0
    check_reg(caps, 4, 0.0);
    check_reg(caps, 100, 0.0);

    // Negative -> 0 (no error)
    check_reg(caps, -1, 0.0);
    check_reg(caps, -0.5, 100.0);  // truncates to 0

    // Truncation of fractional indices
    check_reg(caps, 1.5, 1.0);
    check_reg(caps, 2.999, 2.0);

    // NaN/inf -> 0
    check_reg(caps, std::nan(""), 0.0);
    check_reg(caps, std::numeric_limits<double>::infinity(), 0.0);
    check_reg(caps, -std::numeric_limits<double>::infinity(), 0.0);

    // Empty capture vector
    std::vector<double> empty;
    check_reg(empty, 0, 0.0);
    check_reg(empty, 5, 0.0);

    // ==============================================================

    std::cout << "\n=== summary ===\n";
    std::cout << "passed: " << g_pass << "\n";
    std::cout << "failed: " << g_fail << "\n";

    return g_fail == 0 ? 0 : 1;
}
