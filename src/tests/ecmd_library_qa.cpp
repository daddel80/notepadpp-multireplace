// Standalone tests for the EcmdLibrary and EcmdFunctionInstance.
// These exercise the wrapper at the same depth ExprTkEngine does,
// but without dragging in the rest of the engine: we build a tiny
// outer expression that calls into the library and verify results.
//
// Compile:
//   g++ -std=c++20 -Wall -Wextra ecmd_library_qa.cpp ../engine/EcmdParser.cpp
//       -I.. -o ecmd_library_qa
//   ./ecmd_library_qa [-v]
//
// Note: EcmdLibrary / EcmdFunctionInstance are private inner classes of
// ExprTkEngine. The shape of this test file mirrors what the engine
// does internally, instantiating a tiny adapter that exposes them
// through a thin public surface.

#include "../exprtk/EcmdParser.h"
#include "../exprtk/exprtk.hpp"

#include <cstdio>
#include <cstring>
#include <memory>
#include <string>
#include <vector>

using MultiReplaceEngine::EcmdParser;

// ----- Mirror of EcmdFunctionInstance (kept self-contained here) -----
// The engine version lives inside ExprTkEngine; this is the same logic
// rebuilt in a single TU so the tests don't depend on Windows headers
// that the engine pulls in transitively.

using T = double;
using sym_t = exprtk::symbol_table<T>;
using expr_t = exprtk::expression<T>;
using parser_t = exprtk::parser<T>;

class TestInstance : public exprtk::igeneric_function<T> {
public:
    using igenfunct_t = exprtk::igeneric_function<T>;
    using generic_t = typename igenfunct_t::generic_type;

    static std::string makeSeq(const std::vector<EcmdParser::ParamDef>& ps) {
        std::string s;
        for (const auto& p : ps) s.push_back(static_cast<char>(p.type));
        return s;
    }
    static igenfunct_t::return_type makeRT(EcmdParser::ValueType rt) {
        return (rt == EcmdParser::ValueType::String)
            ? igenfunct_t::e_rtrn_string : igenfunct_t::e_rtrn_scalar;
    }

    TestInstance(const EcmdParser::FunctionDef& def)
        : igenfunct_t(makeSeq(def.params), makeRT(def.returnType))
        , _name(def.name), _returnType(def.returnType), _body(def.body)
    {
        for (const auto& p : def.params) {
            if (p.type == EcmdParser::ValueType::String) {
                _strings.push_back(std::make_unique<std::string>());
                _syms.add_stringvar(p.name, *_strings.back());
                _routes.push_back({ true, _strings.size() - 1 });
            } else {
                _scalars.push_back(std::make_unique<T>(0.0));
                _syms.add_variable(p.name, *_scalars.back());
                _routes.push_back({ false, _scalars.size() - 1 });
            }
        }
        _syms.add_constants();
        if (def.params.empty()) {
            this->allow_zero_parameters() = true;
        }
    }

    void prepareTables(sym_t& lib) {
        _expr.register_symbol_table(_syms);
        _expr.register_symbol_table(lib);
    }
    bool compile(parser_t& p, std::string& err) {
        if (p.compile(_body, _expr)) return true;
        err = "In function '" + _name + "': " + p.error();
        return false;
    }

    T operator()(parameter_list_t params) override {
        return runScalar(params);
    }
    T operator()(std::string& result, parameter_list_t params) override {
        runString(params, result);
        return 0;
    }
    const std::string& name() const { return _name; }

private:
    T runScalar(parameter_list_t params) {
        std::vector<T> sv;  std::vector<std::string> ss;
        save(params, sv, ss);
        marshal(params);
        const T v = _expr.value();
        T result = std::numeric_limits<T>::quiet_NaN();
        if (_expr.return_invoked()) {
            if (_expr.results().count() > 0) _expr.results().get_scalar(0, result);
        } else {
            result = v;
        }
        restore(params, sv, ss);
        return result;
    }
    void runString(parameter_list_t params, std::string& out) {
        std::vector<T> sv;  std::vector<std::string> ss;
        save(params, sv, ss);
        marshal(params);
        _expr.value();
        out.clear();
        if (_expr.return_invoked() && _expr.results().count() > 0) {
            _expr.results().get_string(0, out);
        }
        restore(params, sv, ss);
    }
    void save(parameter_list_t params, std::vector<T>& sv, std::vector<std::string>& ss) {
        for (std::size_t i = 0; i < params.size() && i < _routes.size(); ++i) {
            const auto& r = _routes[i];
            if (r.isString) ss.push_back(*_strings[r.slotIndex]);
            else            sv.push_back(*_scalars[r.slotIndex]);
        }
    }
    void marshal(parameter_list_t params) {
        for (std::size_t i = 0; i < params.size() && i < _routes.size(); ++i) {
            const auto& r = _routes[i];
            if (r.isString) {
                const typename generic_t::string_view svv(params[i]);
                _strings[r.slotIndex]->assign(svv.begin(), svv.size());
            } else {
                const typename generic_t::scalar_view sc(params[i]);
                *_scalars[r.slotIndex] = sc();
            }
        }
    }
    void restore(parameter_list_t params, std::vector<T>& sv, std::vector<std::string>& ss) {
        std::size_t si = 0, ti = 0;
        for (std::size_t i = 0; i < params.size() && i < _routes.size(); ++i) {
            const auto& r = _routes[i];
            if (r.isString) *_strings[r.slotIndex] = std::move(ss[ti++]);
            else            *_scalars[r.slotIndex] = sv[si++];
        }
    }

    struct Route { bool isString; std::size_t slotIndex; };
    std::string                                 _name;
    EcmdParser::ValueType                       _returnType;
    std::string                                 _body;
    sym_t                                       _syms;
    expr_t                                      _expr;
    std::vector<std::unique_ptr<T>>             _scalars;
    std::vector<std::unique_ptr<std::string>>   _strings;
    std::vector<Route>                          _routes;
};

class TestLibrary {
public:
    sym_t& symbolTable() { return _lib; }

    bool load(const std::string& src, std::string& err) {
        auto parsed = EcmdParser::parse(src);
        if (!parsed.success) {
            err = parsed.errorMessage;
            return false;
        }
        std::vector<std::unique_ptr<TestInstance>> pending;
        for (auto& def : parsed.functions) {
            auto inst = std::make_unique<TestInstance>(def);
            if (!_lib.add_function(inst->name(), *inst)) {
                err = "name clash: " + inst->name();
                for (auto& p : pending) _lib.remove_function(p->name());
                return false;
            }
            pending.push_back(std::move(inst));
        }
        for (auto& inst : pending) {
            inst->prepareTables(_lib);
            std::string e;
            if (!inst->compile(_parser, e)) {
                err = e;
                for (auto& p : pending) _lib.remove_function(p->name());
                return false;
            }
        }
        for (auto& inst : pending) _instances.push_back(std::move(inst));
        return true;
    }
private:
    sym_t                                          _lib;
    std::vector<std::unique_ptr<TestInstance>>     _instances;
    parser_t                                       _parser;
};

// ----- Test helpers -----

namespace {

int passed = 0, failed = 0;
bool verbose = false;

struct Eval {
    bool   ok;
    double scalar;
    std::string str;
    std::string err;
};

Eval evalAgainst(TestLibrary& lib, const std::string& outerSrc) {
    sym_t outerSyms;
    outerSyms.add_constants();
    expr_t e;
    e.register_symbol_table(outerSyms);
    e.register_symbol_table(lib.symbolTable());
    parser_t p;
    if (!p.compile(outerSrc, e)) return { false, 0, "", p.error() };
    const T v = e.value();
    Eval r{ true, v, "", "" };
    if (e.return_invoked() && e.results().count() > 0) {
        using ts = exprtk::type_store<T>;
        const auto& t = e.results()[0];
        if (t.type == ts::e_string) {
            typename ts::string_view sv(t);
            r.str.assign(sv.begin(), sv.size());
        } else if (t.type == ts::e_scalar) {
            typename ts::scalar_view sc(t);
            r.scalar = sc();
        }
    }
    return r;
}

void expectScalar(const char* label, const std::string& lib_src,
                  const std::string& expr_src, double expected)
{
    TestLibrary lib;
    std::string err;
    if (!lib.load(lib_src, err)) {
        std::printf("FAIL [%s] library load: %s\n", label, err.c_str());
        ++failed; return;
    }
    auto r = evalAgainst(lib, expr_src);
    if (!r.ok) {
        std::printf("FAIL [%s] outer compile: %s\n", label, r.err.c_str());
        ++failed; return;
    }
    if (r.scalar != expected) {
        std::printf("FAIL [%s] expected %.6f, got %.6f\n",
                    label, expected, r.scalar);
        ++failed; return;
    }
    if (verbose) std::printf("PASS  [%s] -> %.6f\n", label, r.scalar);
    ++passed;
}

void expectString(const char* label, const std::string& lib_src,
                  const std::string& expr_src, const std::string& expected)
{
    TestLibrary lib;
    std::string err;
    if (!lib.load(lib_src, err)) {
        std::printf("FAIL [%s] library load: %s\n", label, err.c_str());
        ++failed; return;
    }
    auto r = evalAgainst(lib, expr_src);
    if (!r.ok) {
        std::printf("FAIL [%s] outer compile: %s\n", label, r.err.c_str());
        ++failed; return;
    }
    if (r.str != expected) {
        std::printf("FAIL [%s] expected '%s', got '%s'\n",
                    label, expected.c_str(), r.str.c_str());
        ++failed; return;
    }
    if (verbose) std::printf("PASS  [%s] -> '%s'\n", label, r.str.c_str());
    ++passed;
}

void expectLoadFail(const char* label, const std::string& lib_src,
                    const char* expectedSubstring)
{
    TestLibrary lib;
    std::string err;
    if (lib.load(lib_src, err)) {
        std::printf("FAIL [%s] expected load failure containing '%s', got success\n",
                    label, expectedSubstring);
        ++failed; return;
    }
    if (err.find(expectedSubstring) == std::string::npos) {
        std::printf("FAIL [%s] error did not contain '%s', got: %s\n",
                    label, expectedSubstring, err.c_str());
        ++failed; return;
    }
    if (verbose) std::printf("PASS  [%s] -> %s\n", label, err.c_str());
    ++passed;
}

} // namespace


int main(int argc, char** argv) {
    for (int i = 1; i < argc; ++i) {
        if (std::strcmp(argv[i], "-v") == 0 || std::strcmp(argv[i], "--verbose") == 0)
            verbose = true;
    }

    // ---- α: T -> S (Roman numerals) ----
    {
        const std::string lib =
            "function num2rom(n) : S\n"
            "    if (n < 1 or n > 3999) { return ['']; };\n"
            "    var r := ''; var v := n;\n"
            "    while (v >= 1000) { r := r + 'M';  v := v - 1000; };\n"
            "    if    (v >= 900)  { r := r + 'CM'; v := v - 900;  };\n"
            "    if    (v >= 500)  { r := r + 'D';  v := v - 500;  };\n"
            "    if    (v >= 400)  { r := r + 'CD'; v := v - 400;  };\n"
            "    while (v >= 100)  { r := r + 'C';  v := v - 100;  };\n"
            "    if    (v >= 90)   { r := r + 'XC'; v := v - 90;   };\n"
            "    if    (v >= 50)   { r := r + 'L';  v := v - 50;   };\n"
            "    if    (v >= 40)   { r := r + 'XL'; v := v - 40;   };\n"
            "    while (v >= 10)   { r := r + 'X';  v := v - 10;   };\n"
            "    if    (v >= 9)    { r := r + 'IX'; v := v - 9;    };\n"
            "    if    (v >= 5)    { r := r + 'V';  v := v - 5;    };\n"
            "    if    (v >= 4)    { r := r + 'IV'; v := v - 4;    };\n"
            "    while (v >= 1)    { r := r + 'I';  v := v - 1;    };\n"
            "    return [r];\n"
            "end\n";
        expectString("num2rom_1",     lib, "return [num2rom(1)];",       "I");
        expectString("num2rom_4",     lib, "return [num2rom(4)];",       "IV");
        expectString("num2rom_1994",  lib, "return [num2rom(1994)];",    "MCMXCIV");
        expectString("num2rom_2026",  lib, "return [num2rom(2026)];",    "MMXXVI");
        expectString("num2rom_3999",  lib, "return [num2rom(3999)];",    "MMMCMXCIX");
        expectString("num2rom_0",     lib, "return [num2rom(0)];",       "");
        expectString("num2rom_4000",  lib, "return [num2rom(4000)];",    "");
    }

    // ---- β: S -> T (Roman parsing) ----
    {
        const std::string lib =
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
        expectScalar("rom2num_I",        lib, "rom2num('I');",         1.0);
        expectScalar("rom2num_IV",       lib, "rom2num('IV');",        4.0);
        expectScalar("rom2num_MCMXCIV",  lib, "rom2num('MCMXCIV');",   1994.0);
        expectScalar("rom2num_MMXXVI",   lib, "rom2num('MMXXVI');",    2026.0);
        expectScalar("rom2num_MMMCMXCIX",lib, "rom2num('MMMCMXCIX');", 3999.0);
    }

    // ---- Default T -> T ----
    {
        const std::string lib =
            "function double_it(n) return n * 2; end\n"
            "function square(n)    return n * n; end\n";
        expectScalar("double_it",      lib, "double_it(7);",         14.0);
        expectScalar("square",         lib, "square(9);",            81.0);
        // Chain: double_it inside square
        expectScalar("nested_in_outer", lib, "square(double_it(3));", 36.0);
    }

    // ---- Zero-arg function ----
    {
        const std::string lib =
            "function pi_sq() return pi * pi; end\n";
        TestLibrary L;
        std::string err;
        if (L.load(lib, err)) {
            auto rr = evalAgainst(L, "return [pi_sq()];");
            if (!rr.ok) {
                std::printf("FAIL [zero_arg] outer compile: %s\n", rr.err.c_str());
                ++failed;
            } else if (std::abs(rr.scalar - 9.869604401) < 1e-6) {
                ++passed; if (verbose) std::printf("PASS  [zero_arg]\n");
            } else {
                std::printf("FAIL [zero_arg] expected ~9.8696, got %.6f\n", rr.scalar);
                ++failed;
            }
        } else { std::printf("FAIL [zero_arg load] %s\n", err.c_str()); ++failed; }
    }

    // ---- Multi-arg: S, T, S -> S ----
    {
        const std::string lib =
            "function repeat_pad(s: S, n, sep: S) : S\n"
            "    var r := s;\n"
            "    var k := 1;\n"
            "    while (k < n) {\n"
            "        r := r + sep + s;\n"
            "        k := k + 1;\n"
            "    };\n"
            "    return [r];\n"
            "end\n";
        expectString("repeat_pad_3x", lib,
                     "return [repeat_pad('ab', 3, '-')];", "ab-ab-ab");
        expectString("repeat_pad_1x", lib,
                     "return [repeat_pad('xy', 1, '-')];", "xy");
    }

    // ---- Cross-call (function calls another function) ----
    {
        const std::string lib =
            "function num2rom(n) : S\n"
            "    var r := ''; var v := n;\n"
            "    while (v >= 10) { r := r + 'X'; v := v - 10; };\n"
            "    while (v >= 1)  { r := r + 'I'; v := v - 1;  };\n"
            "    return [r];\n"
            "end\n"
            "function double_rom(n) : S\n"
            "    return [num2rom(n * 2)];\n"
            "end\n";
        expectString("cross_call", lib, "return [double_rom(5)];", "X");
        expectString("cross_call2",lib, "return [double_rom(11)];", "XXII");
    }

    // ---- Self-recursion ----
    {
        const std::string lib =
            "function fact(n)\n"
            "    if (n <= 1) { return [1]; };\n"
            "    return [n * fact(n - 1)];\n"
            "end\n";
        expectScalar("fact_1",  lib, "fact(1);",   1.0);
        expectScalar("fact_5",  lib, "fact(5);", 120.0);
        expectScalar("fact_10", lib, "fact(10);", 3628800.0);
    }

    // ---- Mutual recursion (even/odd) ----
    {
        const std::string lib =
            "function isEven(n)\n"
            "    if (n == 0) { return [1]; };\n"
            "    return [isOdd(n - 1)];\n"
            "end\n"
            "function isOdd(n)\n"
            "    if (n == 0) { return [0]; };\n"
            "    return [isEven(n - 1)];\n"
            "end\n";
        expectScalar("isEven_4",  lib, "isEven(4);", 1.0);
        expectScalar("isEven_7",  lib, "isEven(7);", 0.0);
        expectScalar("isOdd_5",   lib, "isOdd(5);",  1.0);
        expectScalar("isOdd_10",  lib, "isOdd(10);", 0.0);
    }

    // ---- Re-entrancy safety: function calls another that calls back ----
    {
        // a(n) calls b(n-1); b(n) calls a(n-1) plus a(n)+1 - needs n to be
        // preserved across the inner call.
        const std::string lib =
            "function a(n)\n"
            "    if (n <= 0) { return [0]; };\n"
            "    var x := b(n - 1);\n"
            "    return [n + x];\n"
            "end\n"
            "function b(n)\n"
            "    if (n <= 0) { return [0]; };\n"
            "    var x := a(n - 1);\n"
            "    return [n + x];\n"
            "end\n";
        // a(5) = 5 + b(4) = 5 + (4 + a(3)) = 5 + 4 + 3 + b(2) = ...
        //     = 5+4+3+2+1+0 = 15
        expectScalar("mutual_recursion_sum", lib, "a(5);", 15.0);
    }

    // ---- Multiple bodies, mixed defaults ----
    {
        const std::string lib =
            "function add(a, b) return a + b; end\n"
            "function greet(name: S) : S return ['Hi, ' + name]; end\n";
        expectScalar("multi_add",   lib, "add(3, 4);",          7.0);
        expectString("multi_greet", lib, "return [greet('Tom')];", "Hi, Tom");
    }

    // ---- Compose: ecmd-function + outer ExprTk + ecmd-function ----
    {
        const std::string lib =
            "function num2rom(n) : S\n"
            "    var r := ''; var v := n;\n"
            "    while (v >= 10) { r := r + 'X'; v := v - 10; };\n"
            "    while (v >= 1)  { r := r + 'I'; v := v - 1;  };\n"
            "    return [r];\n"
            "end\n"
            "function rom2num(s: S)\n"
            "    var total := 0;\n"
            "    var i := s[] - 1;\n"
            "    while (i >= 0) {\n"
            "        if (s[i:i+1] == 'I') { total := total + 1; };\n"
            "        if (s[i:i+1] == 'X') { total := total + 10; };\n"
            "        i := i - 1;\n"
            "    };\n"
            "    return total;\n"
            "end\n";
        // num2rom(rom2num('XX') + 5)  =  num2rom(20 + 5) = "XXV"... wait
        // our toy num2rom doesn't handle V. Let's pick a clean example:
        // num2rom(rom2num('XXIII') * 2) = num2rom(23 * 2) = num2rom(46)
        // = X X X X V I -> but no V again. Stay with 'XII' * 2 = 24
        expectString("compose_pipeline", lib,
                     "return [num2rom(rom2num('XII') * 2)];", "XXIIII");
    }

    // ---- Load errors ----
    expectLoadFail("dup_function",
                   "function f(n) return n; end\n"
                   "function f(n) return n*2; end\n",
                   "Duplicate function");

    expectLoadFail("body_undefined_var",
                   "function f(n) return undefined_var; end\n",
                   "In function 'f'");

    // NOTE: A scalar-declared function whose body returns a string
    // ('function f(n) return [\\'x\\']; end') currently loads without a
    // diagnostic - ExprTk doesn't expose static return-type checking we
    // could hook into. At call time the scalar path will yield NaN,
    // which flows into the engine's recoverable-error handling.
    // Documented as a known limitation; not enforced here.

    // ---- BOM handling ----
    {
        // \xEF\xBB\xBF in three separate adjacent literals to avoid the
        // C++ wide-hex-escape merge that would consume the 'f' that follows.
        std::string src = "\xEF\xBB\xBF" "function f(n) return n + 1; end\n";
        // EcmdLibrary doesn't strip BOM by itself - the engine's
        // loadEcmdFile does. The parser only sees clean content. So we
        // skip those bytes here for the test.
        if (src.size() >= 3 && static_cast<unsigned char>(src[0]) == 0xEF) {
            src.erase(0, 3);
        }
        TestLibrary L;
        std::string err;
        if (L.load(src, err)) {
            auto r = evalAgainst(L, "f(41);");
            if (r.ok && r.scalar == 42.0) {
                ++passed; if (verbose) std::printf("PASS  [bom_stripped]\n");
            } else { std::printf("FAIL [bom_stripped]\n"); ++failed; }
        } else { std::printf("FAIL [bom_stripped load] %s\n", err.c_str()); ++failed; }
    }

    std::printf("\n%d passed, %d failed\n", passed, failed);
    return failed == 0 ? 0 : 1;
}