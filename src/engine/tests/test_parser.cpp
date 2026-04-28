// Test harness for ExprTkPatternParser. Compile with:
//   g++ -std=c++17 -Wall -Wextra -Wpedantic -O2 \
//       -I/home/claude/engine \
//       /home/claude/engine/ExprTkPatternParser.cpp \
//       /home/claude/engine/test_parser.cpp \
//       -o /tmp/test_parser
//
// The tests cover all 18 edge cases from the design phase plus a few
// additional adversarial cases.

#include "ExprTkPatternParser.h"

#include <cassert>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

using MultiReplaceEngine::ExprTkPatternParser;
using ST = MultiReplaceEngine::ExprTkPatternParser::SegmentType;

static int g_pass = 0;
static int g_fail = 0;

// Helper: pretty-print a string with visible escapes for diagnostic
// output. \ becomes \\, control chars become \xNN.
static std::string visible(const std::string& s)
{
    std::ostringstream o;
    for (unsigned char c : s) {
        if (c == '\\') {
            o << "\\\\";
        }
        else if (c < 32 || c >= 127) {
            o << "\\x" << std::hex << std::setw(2) << std::setfill('0') << (int)c << std::dec;
        }
        else {
            o << (char)c;
        }
    }
    return o.str();
}

struct ExpectedSeg {
    ST          type;
    std::string text;
};

static void check(const std::string& name,
                  const std::string& input,
                  const std::vector<ExpectedSeg>& expected,
                  bool expectSuccess = true,
                  const std::string& expectedErrorContains = "")
{
    auto r = ExprTkPatternParser::parse(input);

    bool ok = true;
    std::ostringstream diag;

    if (r.success != expectSuccess) {
        ok = false;
        diag << "  expected success=" << expectSuccess
             << " got " << r.success
             << " (errorMessage=\"" << r.errorMessage << "\")\n";
    }

    if (expectSuccess) {
        if (r.segments.size() != expected.size()) {
            ok = false;
            diag << "  expected " << expected.size() << " segments, got "
                 << r.segments.size() << "\n";
        }
        else {
            for (std::size_t i = 0; i < expected.size(); ++i) {
                if (r.segments[i].type != expected[i].type) {
                    ok = false;
                    diag << "  segment[" << i << "] type mismatch: "
                         << "expected " << (expected[i].type == ST::Literal ? "Literal" : "Expression")
                         << " got "      << (r.segments[i].type == ST::Literal ? "Literal" : "Expression")
                         << "\n";
                }
                if (r.segments[i].text != expected[i].text) {
                    ok = false;
                    diag << "  segment[" << i << "] text mismatch:\n"
                         << "    expected: \"" << visible(expected[i].text) << "\"\n"
                         << "    got:      \"" << visible(r.segments[i].text) << "\"\n";
                }
            }
        }
    }
    else {
        if (!expectedErrorContains.empty() &&
            r.errorMessage.find(expectedErrorContains) == std::string::npos) {
            ok = false;
            diag << "  error message did not contain \"" << expectedErrorContains << "\"\n"
                 << "    actual: \"" << r.errorMessage << "\"\n";
        }
        if (!r.segments.empty()) {
            ok = false;
            diag << "  on failure segments must be empty, got " << r.segments.size() << "\n";
        }
    }

    if (ok) {
        ++g_pass;
        std::cout << "[PASS] " << name << "\n";
    }
    else {
        ++g_fail;
        std::cout << "[FAIL] " << name << "\n";
        std::cout << "  input: \"" << visible(input) << "\"\n";
        std::cout << diag.str();
    }
}

// Convenience constructors
static ExpectedSeg L(const std::string& t) { return { ST::Literal,    t }; }
static ExpectedSeg E(const std::string& t) { return { ST::Expression, t }; }


int main()
{
    std::cout << "=== ExprTkPatternParser test suite ===\n\n";

    // ---------- 18 edge cases from the design phase ----------

    // 1. Normal case
    check("01_normal_case",
          "a=(?=reg(1)+5)b",
          { L("a="), E("reg(1)+5"), L("b") });

    // 2. Multiple expressions
    check("02_multiple_expressions",
          "(?=reg(1))(?=reg(2))",
          { E("reg(1)"), E("reg(2)") });

    // 3. Template without expressions (common case)
    check("03_no_expressions",
          "Hello World",
          { L("Hello World") });

    // 4. Single expression, no literals
    check("04_only_expression",
          "(?=reg(1)*2)",
          { E("reg(1)*2") });

    // 5. Nested parens inside expression
    check("05_nested_parens",
          "(?=min(reg(1), reg(2)))",
          { E("min(reg(1), reg(2))") });

    // 6. Escaped marker (literal "(?=...)"
    check("06_escaped_marker",
          "\\(?=foo)",
          { L("(?=foo)") });

    // 7a. Backslash itself (\\\\ -> \\)
    check("07a_double_backslash_in_text",
          "a\\\\b",
          { L("a\\b") });

    // 7b. Backslash followed by escaped marker
    //     "\\\\(?=foo)" should be: literal "\" + expression "foo"
    //     because \\ consumes 2 chars -> "\", then (?= starts expr.
    check("07b_backslash_then_marker",
          "\\\\(?=foo)",
          { L("\\"), E("foo") });

    // 8. Unmatched (?= at end
    check("08_unmatched_marker",
          "a(?=reg(1)",
          {},
          false,
          "Unmatched");

    // 9. Empty expression
    check("09_empty_expression",
          "(?=)",
          {},
          false,
          "Empty");

    // 10. Expression with escape characters - they are NOT escapes
    //     inside an expression. The expression text is verbatim.
    check("10_no_escape_in_expression",
          "(?=foo\\bar)",
          { E("foo\\bar") });

    // 11. UTF-8 in literal text - bytes pass through unchanged
    //     "Über=" in UTF-8: 55 CC 88 62 65 72 3D ... actually it's
    //     C3 9C 62 65 72 3D - either way, the parser only cares that
    //     the marker bytes don't appear in the multi-byte sequence.
    check("11_utf8_literal",
          "Über=(?=reg(1))",
          { L("Über="), E("reg(1)") });

    // 12. UTF-8 in expression - flows through to engine, parser
    //     must not corrupt bytes
    check("12_utf8_in_expression",
          "(?=reg(1)+ä)",
          { E("reg(1)+ä") });

    // 13. Multiple backslashes before marker
    //     "\\\\\\\\(?=foo)" = 4 backslashes + (?=foo)
    //     parses as: "\\" + "\\" + "(?=foo)"
    //     -> "\\" + "\\" = literal "\\\\" (two backslashes)
    //     -> then (?= starts expression
    check("13_four_backslashes_marker",
          "\\\\\\\\(?=foo)",
          { L("\\\\"), E("foo") });

    // 14. Lone (?= without proper following char is unmatched
    check("14_lone_open_marker",
          "(?=",
          {},
          false,
          "Unmatched");

    // 15. Long literal text without markers (one segment)
    check("15_long_literal",
          "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
          { L("Lorem ipsum dolor sit amet, consectetur adipiscing elit.") });

    // 16. ( with whitespace before ?= is NOT a marker
    check("16_space_in_marker",
          "( ?=reg(1))",
          { L("( ?=reg(1))") });

    // 17. Nested (?= inside an expression - the outer depth-counter
    //     swallows the inner ( because we don't recursively detect
    //     markers. The inner ?= is just three chars in the expression.
    //     "(?=reg(1)+(?=reg(2)))" parses as one expression:
    //         text = "reg(1)+(?=reg(2))"
    //     The closing ')' that balances the outer '(?=' is the last
    //     one. Depth: opens at (?=(1), then (?=(2), opens 2x for
    //     reg(...), so depth tracking yields the rightmost ')'.
    //     Let's count manually:
    //         (?=    -> we entered, depth=1, i past the marker
    //         r e g  -> 3 chars
    //         (      -> depth=2
    //         1
    //         )      -> depth=1
    //         + ( ?  -> ( bumps depth to 2
    //         = r e g
    //         (      -> depth=3
    //         2
    //         )      -> depth=2
    //         )      -> depth=1
    //         )      -> depth=0, STOP
    //     Final expression text = "reg(1)+(?=reg(2))"
    check("17_nested_marker_in_expr",
          "(?=reg(1)+(?=reg(2)))",
          { E("reg(1)+(?=reg(2))") });

    // 18. Backslash at end of input
    check("18_trailing_backslash",
          "abc\\",
          { L("abc\\") });

    // ---------- additional adversarial cases ----------

    // A1. Empty input
    check("A1_empty_input",
          "",
          {});

    // A2. Just a backslash
    check("A2_just_backslash",
          "\\",
          { L("\\") });

    // A3. Just \\\\ (two backslashes -> one literal \)
    check("A3_just_two_backslashes",
          "\\\\",
          { L("\\") });

    // A4. \(?= with no closing ) is still escaped - so literal "(?="
    //     and we don't enter expression mode
    check("A4_escaped_marker_no_close",
          "\\(?=",
          { L("(?=") });

    // A5. Adjacent expressions
    check("A5_adjacent_expressions",
          "(?=a)(?=b)(?=c)",
          { E("a"), E("b"), E("c") });

    // A6. Expression containing string-like content
    check("A6_expression_with_string",
          "(?=if(reg(1) > 0, 1, 0))",
          { E("if(reg(1) > 0, 1, 0)") });

    // A7. Mix: literal -> expr -> literal -> expr -> literal
    check("A7_mixed_pattern",
          "before(?=a)middle(?=b)after",
          { L("before"), E("a"), L("middle"), E("b"), L("after") });

    // A8. Large depth in expression (brace stress)
    //     Input: "(?=((((reg(1))))))" -> the parser skips "(?=",
    //     then depth-counts: each '(' bumps depth, each ')' lowers,
    //     stops when depth hits 0. So the trailing ')' that closes
    //     "(?=" is consumed; the expression text is everything in
    //     between: "((((reg(1)))))".
    check("A8_deep_nesting",
          "(?=((((reg(1))))))",
          { E("((((reg(1)))))") });

    // A9. Just one (?=...) wrapping nothing useful
    check("A9_minimal_expression",
          "(?=0)",
          { E("0") });

    // A10. Very long literal followed by expression
    {
        std::string longLit(1000, 'x');
        check("A10_long_literal_then_expr",
              longLit + "(?=reg(1))",
              { L(longLit), E("reg(1)") });
    }

    // A11. Single-char marker case: "(?=)" is empty -> error
    //     Already tested in #9

    // A12. Special chars in literal pass through
    check("A12_special_chars_literal",
          "$@#%^&*+=",
          { L("$@#%^&*+=") });

    // A13. Newlines in literal
    check("A13_newlines_in_literal",
          "line1\nline2",
          { L("line1\nline2") });

    // A14. Tabs in literal
    check("A14_tabs_in_literal",
          "col1\tcol2",
          { L("col1\tcol2") });

    // A15. Marker-like sequence that's NOT (?= - "(?<" for example
    check("A15_marker_lookalike",
          "(?<reg(1)>)",
          { L("(?<reg(1)>)") });

    // A16. \\ inside expression has no special meaning -> verbatim
    check("A16_double_backslash_in_expr",
          "(?=\\\\)",
          { E("\\\\") });

    // A17. Mixed escapes in literal
    check("A17_mixed_escapes",
          "\\(?=a)\\\\(?=reg(1))",
          { L("(?=a)\\"), E("reg(1)") });

    // ---------- hasExpressions() pre-check ----------

    if (ExprTkPatternParser::hasExpressions("Hello World")) {
        ++g_fail;
        std::cout << "[FAIL] hasExpressions: false positive on plain literal\n";
    }
    else {
        ++g_pass;
        std::cout << "[PASS] hasExpressions: plain literal -> false\n";
    }

    if (!ExprTkPatternParser::hasExpressions("a(?=b)c")) {
        ++g_fail;
        std::cout << "[FAIL] hasExpressions: false negative on simple expression\n";
    }
    else {
        ++g_pass;
        std::cout << "[PASS] hasExpressions: with marker -> true\n";
    }

    if (ExprTkPatternParser::hasExpressions("\\(?=b)")) {
        ++g_fail;
        std::cout << "[FAIL] hasExpressions: escaped marker counted as expression\n";
    }
    else {
        ++g_pass;
        std::cout << "[PASS] hasExpressions: escaped marker -> false\n";
    }

    // ---------- summary ----------
    std::cout << "\n=== summary ===\n";
    std::cout << "passed: " << g_pass << "\n";
    std::cout << "failed: " << g_fail << "\n";

    return g_fail == 0 ? 0 : 1;
}
