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

// QA harness for the results of (?=...) formulas on the patched ExprTk, read as ExprTkEngine reads
// them: a string-typed formula through expression_helper::get_string(), a numeric one through
// value(), one ending in return [...] through its result list.
//   S. single expressions ending in text: literal, concatenation, ternary
//   L. several statements ending in text: text, not the number ExprTk gives the last statement
//   N. several statements ending in a number or a comparison stay numeric
//   R. return [...] keeps its path
//
// Build (from src/tests):
//   g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. -o exprtk_string_qa exprtk_string_qa.cpp
// Run:   ./exprtk_string_qa
#include "../exprtk/third_party/exprtk.hpp"

#include <cstdio>
#include <string>

using helper_t = exprtk::expression_helper<double>;

static int failures = 0;
static double g_x = 0;

// "S:text" from get_string(), "N:number" from value(), "R:text" from return [...]
static std::string run(const std::string& formula, double x)
{
    exprtk::symbol_table<double> symbols;
    symbols.add_variable("x", g_x);
    exprtk::expression<double> expr;
    expr.register_symbol_table(symbols);
    exprtk::parser<double> parser;
    if (!parser.compile(formula, expr)) return "compile error: " + parser.error();
    g_x = x;
    std::string text;
    if (helper_t::is_string(expr)) return helper_t::get_string(expr, text) ? "S:" + text : "S: failed";
    const double value = expr.value();
    if (expr.return_invoked()) return expr.results().get_string(0, text) ? "R:" + text : "R: no text";
    char number[64];
    std::snprintf(number, sizeof number, "%g", value);
    return std::string("N:") + number;
}

static void check(const char* label, const std::string& formula, double x, const std::string& expected)
{
    const std::string got = run(formula, x);
    const bool ok = got == expected;
    std::printf("%s %s: %s with x = %g -> %s", ok ? "PASS" : "FAIL", label, formula.c_str(), x, got.c_str());
    if (!ok) { std::printf(" (expected %s)", expected.c_str()); ++failures; }
    std::printf("\n");
}

int main()
{
    check("S1", "'abc'", 0, "S:abc");
    check("S2", "'a' + 'b'", 0, "S:ab");
    check("S3", "x > 1 ? 'A' : 'B'", 2, "S:A");
    check("S4", "x > 1 ? 'A' : 'B'", 1, "S:B");

    check("L1", "var y := x; y > 1 ? 'A' : 'B'", 2, "S:A");
    check("L2", "var y := x; y > 1 ? 'A' : 'B'", 1, "S:B");
    check("L3", "var y := x; 'X'", 0, "S:X");
    check("L4", "var s := 'a'; s + 'b'", 0, "S:ab");
    check("L5", "var t := ''; for (var i := 0; i < x; i += 1) { t := t + '*'; }; t", 3, "S:***");

    check("N1", "var y := x; y * 10", 2, "N:20");
    check("N2", "var y := x; y > 1 ? 100 : 200", 2, "N:100");
    check("N3", "var s := 'a'; s == 'a'", 0, "N:1");
    check("N4", "x * 3", 2, "N:6");

    check("R1", "var y := x; return [y > 1 ? 'A' : 'B']", 2, "R:A");
    check("R2", "return ['abc']", 0, "R:abc");

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
