// Standalone tests for EcmdParser.
// Compile:
//   g++ -std=c++20 -Wall -Wextra ecmd_parser_qa.cpp ../engine/EcmdParser.cpp -o ecmd_parser_qa
//   ./ecmd_parser_qa [-v]

#include "../exprtk/EcmdParser.h"

#include <cstdio>
#include <cstring>
#include <string>

using MultiReplaceEngine::EcmdParser;

namespace {

int  passed = 0;
int  failed = 0;
bool verbose = false;

void checkOk(const std::string& src, const char* label) {
    auto r = EcmdParser::parse(src);
    if (!r.success) {
        std::printf("FAIL [%s] expected success, got error: %s @ %zu\n",
                    label, r.errorMessage.c_str(), r.errorPos);
        ++failed;
        return;
    }
    if (verbose) std::printf("PASS  [%s]\n", label);
    ++passed;
}

void checkFail(const std::string& src, const char* expectedSubstring,
               const char* label)
{
    auto r = EcmdParser::parse(src);
    if (r.success) {
        std::printf("FAIL [%s] expected error containing \"%s\", got success\n",
                    label, expectedSubstring);
        ++failed;
        return;
    }
    if (r.errorMessage.find(expectedSubstring) == std::string::npos) {
        std::printf("FAIL [%s] expected error containing \"%s\", got: %s\n",
                    label, expectedSubstring, r.errorMessage.c_str());
        ++failed;
        return;
    }
    if (verbose) std::printf("PASS  [%s] -> %s\n", label, r.errorMessage.c_str());
    ++passed;
}

void checkFunctionShape(const std::string& src,
                        const char* expectedName,
                        std::size_t expectedParamCount,
                        EcmdParser::ValueType expectedReturn,
                        const char* label)
{
    auto r = EcmdParser::parse(src);
    if (!r.success) {
        std::printf("FAIL [%s] parse failed: %s\n", label, r.errorMessage.c_str());
        ++failed;
        return;
    }
    if (r.functions.size() != 1) {
        std::printf("FAIL [%s] expected 1 function, got %zu\n",
                    label, r.functions.size());
        ++failed;
        return;
    }
    const auto& fn = r.functions[0];
    if (fn.name != expectedName) {
        std::printf("FAIL [%s] name: expected '%s', got '%s'\n",
                    label, expectedName, fn.name.c_str());
        ++failed;
        return;
    }
    if (fn.params.size() != expectedParamCount) {
        std::printf("FAIL [%s] param count: expected %zu, got %zu\n",
                    label, expectedParamCount, fn.params.size());
        ++failed;
        return;
    }
    if (fn.returnType != expectedReturn) {
        std::printf("FAIL [%s] return type mismatch\n", label);
        ++failed;
        return;
    }
    if (verbose) std::printf("PASS  [%s]\n", label);
    ++passed;
}

} // namespace


int main(int argc, char** argv) {
    for (int i = 1; i < argc; ++i) {
        if (std::strcmp(argv[i], "-v") == 0 || std::strcmp(argv[i], "--verbose") == 0) {
            verbose = true;
        }
    }

    using VT = EcmdParser::ValueType;

    // ---- Empty / trivial ----
    {
        auto r = EcmdParser::parse("");
        if (!r.success || !r.functions.empty()) {
            std::printf("FAIL [empty] expected success+0 functions\n");
            ++failed;
        } else { ++passed; if (verbose) std::printf("PASS  [empty]\n"); }
    }
    checkOk("   \n\t  \n", "whitespace_only");
    checkOk("/* just a comment */", "comment_only");
    checkOk("// line comment only\n", "line_comment_only");

    // ---- Basic shapes ----
    checkFunctionShape(
        "function foo(n) return n; end",
        "foo", 1, VT::Scalar, "basic_T_T");

    checkFunctionShape(
        "function foo(n) : S return [n]; end",
        "foo", 1, VT::String, "T_to_S");

    checkFunctionShape(
        "function foo(s: S) return 1; end",
        "foo", 1, VT::Scalar, "S_to_T_implicit_return");

    checkFunctionShape(
        "function foo(s: S) : T return 1; end",
        "foo", 1, VT::Scalar, "S_to_T_explicit");

    checkFunctionShape(
        "function foo(s: S) : S return [s]; end",
        "foo", 1, VT::String, "S_to_S");

    checkFunctionShape(
        "function noargs() return 42; end",
        "noargs", 0, VT::Scalar, "zero_args");

    checkFunctionShape(
        "function noargs() : S return ['x']; end",
        "noargs", 0, VT::String, "zero_args_S");

    // ---- Multi-arg ----
    checkFunctionShape(
        "function padleft(s: S, n, fill: S) : S return [s]; end",
        "padleft", 3, VT::String, "three_args_mixed");

    checkFunctionShape(
        "function f(a, b, c, d, e) return a; end",
        "f", 5, VT::Scalar, "five_scalars");

    {
        auto r = EcmdParser::parse("function f(a: S, b: T, c: S) : S return [a]; end");
        if (r.success && r.functions.size() == 1
            && r.functions[0].params[0].type == VT::String
            && r.functions[0].params[1].type == VT::Scalar
            && r.functions[0].params[2].type == VT::String) {
            ++passed; if (verbose) std::printf("PASS  [param_types_individual]\n");
        } else {
            std::printf("FAIL [param_types_individual]\n");
            ++failed;
        }
    }

    // ---- Multiple functions ----
    {
        const std::string src =
            "function a(n) return n; end\n"
            "function b(s: S) : S return [s]; end\n"
            "function c() return 0; end\n";
        auto r = EcmdParser::parse(src);
        if (r.success && r.functions.size() == 3
            && r.functions[0].name == "a"
            && r.functions[1].name == "b"
            && r.functions[2].name == "c") {
            ++passed; if (verbose) std::printf("PASS  [three_functions]\n");
        } else {
            std::printf("FAIL [three_functions]\n"); ++failed;
        }
    }

    // ---- Body preservation ----
    {
        const std::string src =
            "function r2(n) : S\n"
            "    var r := '';\n"
            "    var v := n;\n"
            "    while (v >= 10) { r := r + 'X'; v := v - 10; };\n"
            "    while (v >= 1)  { r := r + 'I'; v := v - 1;  };\n"
            "    return [r];\n"
            "end\n";
        auto r = EcmdParser::parse(src);
        if (!r.success) {
            std::printf("FAIL [body_with_blocks] %s\n", r.errorMessage.c_str());
            ++failed;
        } else if (r.functions.size() != 1
                   || r.functions[0].body.find("while (v >= 10)") == std::string::npos
                   || r.functions[0].body.find("return [r];") == std::string::npos) {
            std::printf("FAIL [body_with_blocks] body content lost\n");
            ++failed;
        } else { ++passed; if (verbose) std::printf("PASS  [body_with_blocks]\n"); }
    }

    // ---- String literal in body, must not eat 'end' ----
    {
        // The string contains the substring "end" which the scanner
        // must NOT treat as the function terminator.
        const std::string src =
            "function f(n) : S\n"
            "    var s := 'pretend';\n"
            "    return [s];\n"
            "end\n";
        auto r = EcmdParser::parse(src);
        if (r.success && r.functions.size() == 1
            && r.functions[0].body.find("'pretend'") != std::string::npos) {
            ++passed; if (verbose) std::printf("PASS  [end_inside_string]\n");
        } else {
            std::printf("FAIL [end_inside_string] %s\n",
                        r.success ? "(body lost)" : r.errorMessage.c_str());
            ++failed;
        }
    }

    // ---- Identifier containing 'end' as substring ----
    {
        // 'extended' starts with the letters of 'end' partway through;
        // the scanner must use whole-token matching.
        const std::string src =
            "function f(n)\n"
            "    var extended := n;\n"
            "    return extended;\n"
            "end\n";
        auto r = EcmdParser::parse(src);
        if (r.success && r.functions.size() == 1
            && r.functions[0].body.find("extended") != std::string::npos) {
            ++passed; if (verbose) std::printf("PASS  [end_substring_in_ident]\n");
        } else {
            std::printf("FAIL [end_substring_in_ident]\n");
            ++failed;
        }
    }

    // ---- Block comment inside body ----
    {
        const std::string src =
            "function f(n) /* tricky: end */ return n; end";
        auto r = EcmdParser::parse(src);
        if (r.success && r.functions.size() == 1) {
            ++passed; if (verbose) std::printf("PASS  [block_comment_with_end_word]\n");
        } else {
            std::printf("FAIL [block_comment_with_end_word]\n"); ++failed;
        }
    }

    // ---- Comments between functions ----
    checkOk(
        "/* lib doc */\n"
        "function a(n) return n; end\n"
        "// separator\n"
        "function b(n) return n*2; end\n",
        "comments_between_functions");

    // ---- Doubled '' inside a string ----
    {
        // ExprTk's escape for a single quote inside a string is ''.
        // The scanner must treat 'it''s' as a single string token.
        const std::string src =
            "function f(n) : S\n"
            "    var s := 'it''s ok';\n"
            "    return [s];\n"
            "end\n";
        auto r = EcmdParser::parse(src);
        if (r.success && r.functions.size() == 1) {
            ++passed; if (verbose) std::printf("PASS  [doubled_quote_escape]\n");
        } else {
            std::printf("FAIL [doubled_quote_escape] %s\n", r.errorMessage.c_str());
            ++failed;
        }
    }

    // ---- Error cases: type annotation ----
    checkFail("function f(s: x) return 0; end",
              "Unknown type", "unknown_type_letter");
    checkFail("function f(s: s) return 0; end",
              "uppercase", "lowercase_s");
    checkFail("function f(s: t) return 0; end",
              "uppercase", "lowercase_t");
    checkFail("function f(s: STRING) return 0; end",
              "single uppercase letter", "multichar_type");
    checkFail("function f(s:) return 0; end",
              "Expected type", "type_missing_after_colon");
    checkFail("function f(s S) return 0; end",
              "Expected ',' or ')'", "missing_colon_before_type");

    // ---- Error cases: return type ----
    checkFail("function f(n) : x return n; end",
              "Unknown return type", "unknown_return_type");
    checkFail("function f(n) : s return n; end",
              "uppercase", "lowercase_return_type");
    checkFail("function f(n) : SS return n; end",
              "single uppercase letter", "multichar_return_type");

    // ---- Error cases: structural ----
    checkFail("function", "Expected function name", "lone_function_keyword");
    checkFail("function f", "Expected '('", "no_paren_after_name");
    checkFail("function f(", "Expected parameter name", "open_paren_no_close");
    checkFail("function f(n,) return n; end",
              "Expected parameter name", "trailing_comma");
    checkFail("function f(n", "Expected ',' or ')'", "param_eof");
    checkFail("function f(n) return n;",
              "no matching 'end'", "unterminated_body");
    checkFail("function f(n)\n  /* unterminated\n  return n; end",
              "Unterminated", "unterminated_comment_in_body");
    checkFail("function f(n) return 'open string ; end",
              "Unterminated string", "unterminated_string");
    checkFail("function f(n, n) return n; end",
              "Duplicate parameter", "dup_param");
    checkFail(
        "function a(n) return n; end\n"
        "function a(n) return n*2; end\n",
        "Duplicate function", "dup_function");
    checkFail(
        "function f(n)\n"
        "  function g(n) return n; end\n"
        "end",
        "Unexpected 'function' inside body", "missing_end");

    // ---- Real-world example: roman numerals ----
    {
        const std::string roman =
            "function num2rom(n) : S\n"
            "    if (n < 1 or n > 3999) { return ['']; };\n"
            "    var r := '';\n"
            "    var v := n;\n"
            "    while (v >= 1000) { r := r + 'M';  v := v - 1000; };\n"
            "    if  (v >= 900)    { r := r + 'CM'; v := v - 900;  };\n"
            "    if  (v >= 500)    { r := r + 'D';  v := v - 500;  };\n"
            "    if  (v >= 400)    { r := r + 'CD'; v := v - 400;  };\n"
            "    while (v >= 100)  { r := r + 'C';  v := v - 100;  };\n"
            "    if  (v >= 90)     { r := r + 'XC'; v := v - 90;   };\n"
            "    if  (v >= 50)     { r := r + 'L';  v := v - 50;   };\n"
            "    if  (v >= 40)     { r := r + 'XL'; v := v - 40;   };\n"
            "    while (v >= 10)   { r := r + 'X';  v := v - 10;   };\n"
            "    if  (v >= 9)      { r := r + 'IX'; v := v - 9;    };\n"
            "    if  (v >= 5)      { r := r + 'V';  v := v - 5;    };\n"
            "    if  (v >= 4)      { r := r + 'IV'; v := v - 4;    };\n"
            "    while (v >= 1)    { r := r + 'I';  v := v - 1;    };\n"
            "    return [r];\n"
            "end\n"
            "\n"
            "function rom2num(s: S)\n"
            "    var total := 0;\n"
            "    var prev := 0;\n"
            "    var i := s[] - 1;\n"
            "    while (i >= 0) {\n"
            "        var c := s[i:i+1];\n"
            "        var cur := 0;\n"
            "        if      (c == 'I') { cur := 1;    }\n"
            "        else if (c == 'V') { cur := 5;    }\n"
            "        else if (c == 'X') { cur := 10;   }\n"
            "        else if (c == 'L') { cur := 50;   }\n"
            "        else if (c == 'C') { cur := 100;  }\n"
            "        else if (c == 'D') { cur := 500;  }\n"
            "        else if (c == 'M') { cur := 1000; };\n"
            "        if (cur < prev) { total := total - cur; }\n"
            "        else            { total := total + cur; };\n"
            "        prev := cur;\n"
            "        i := i - 1;\n"
            "    };\n"
            "    return total;\n"
            "end\n";
        auto r = EcmdParser::parse(roman);
        if (!r.success) {
            std::printf("FAIL [roman_full] %s @ %zu\n",
                        r.errorMessage.c_str(), r.errorPos);
            ++failed;
        } else if (r.functions.size() != 2
                   || r.functions[0].name != "num2rom"
                   || r.functions[0].returnType != VT::String
                   || r.functions[1].name != "rom2num"
                   || r.functions[1].returnType != VT::Scalar
                   || r.functions[1].params.size() != 1
                   || r.functions[1].params[0].type != VT::String)
        {
            std::printf("FAIL [roman_full] structural mismatch\n");
            ++failed;
        } else {
            ++passed;
            if (verbose) std::printf("PASS  [roman_full]\n");
        }
    }

    std::printf("\n%d passed, %d failed\n", passed, failed);
    return failed == 0 ? 0 : 1;
}