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

// QA harness for Encoding::isValidUtf8 with the eight-byte ASCII step: every verdict must
// equal the byte-by-byte validator it replaces (a43e162), which serves as the reference.
//   E. all inputs of one to three bytes, and four-byte inputs over every lead byte with
//      boundary continuation bytes
//   B. every short probe at every offset in ASCII runs of 0 to 24 bytes, so the eight-byte
//      blocks start, end and get cut at every position, the probe also at the very end
//   R. random mixes of ASCII runs, valid sequences, stray and truncated bytes
//   C. optional: every file below the given folders, whole and cut at the last bytes
//   V. the copy below is the code in ../Encoding.cpp (text compared, line ends ignored)
//
// Build (from src/tests):
//   g++ -std=c++20 -O1 -Wall -Wextra -fsanitize=address,undefined -o utf8_validation_qa utf8_validation_qa.cpp
// Run:   ./utf8_validation_qa [folder ...]
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <random>
#include <sstream>
#include <string>
#include <vector>

// ----------------------------- Reference: Encoding::isValidUtf8 as of a43e162
namespace reference {
    bool isValidUtf8(const char* data, size_t len) {
        const unsigned char* s = reinterpret_cast<const unsigned char*>(data);
        size_t i = 0;
        while (i < len) {
            unsigned char c = s[i];
            if (c < 0x80) { ++i; continue; }
            size_t need = 0;
            if ((c & 0xE0) == 0xC0) { need = 1; if ((c & 0xFE) == 0xC0) return false; }
            else if ((c & 0xF0) == 0xE0) { need = 2; }
            else if ((c & 0xF8) == 0xF0) { need = 3; if (c > 0xF4) return false; }
            else return false;
            if (i + need >= len) return false;
            for (size_t k = 1; k <= need; ++k)
                if ((s[i + k] & 0xC0) != 0x80) return false;
            if (need == 2) {
                unsigned char c1 = s[i + 1];
                if ((c == 0xE0 && (c1 & 0xE0) == 0x80) || (c == 0xED && (c1 & 0xE0) == 0xA0)) return false;
            }
            else if (need == 3) {
                unsigned char c1 = s[i + 1];
                if ((c == 0xF0 && (c1 & 0xF0) == 0x80) || (c == 0xF4 && (c1 & 0xF0) != 0x80)) return false;
            }
            i += (need + 1);
        }
        return true;
    }
}

// ----------------------------- VERBATIM: Encoding::isValidUtf8 (Encoding.cpp)
namespace Encoding {
    bool isValidUtf8(const char* data, size_t len) {
        const unsigned char* s = reinterpret_cast<const unsigned char*>(data);
        size_t i = 0;
        while (i < len) {
            if (len - i >= 8) {                                      // eight ASCII bytes at once
                uint64_t block;
                std::memcpy(&block, s + i, 8);
                if ((block & 0x8080808080808080ull) == 0) { i += 8; continue; }
            }
            unsigned char c = s[i];
            if (c < 0x80) { ++i; continue; }                         // ASCII
            size_t need = 0;
            if ((c & 0xE0) == 0xC0) { need = 1; if ((c & 0xFE) == 0xC0) return false; } // overlong 2-byte
            else if ((c & 0xF0) == 0xE0) { need = 2; }
            else if ((c & 0xF8) == 0xF0) { need = 3; if (c > 0xF4) return false; }      // >U+10FFFF
            else return false;

            if (i + need >= len) return false;
            for (size_t k = 1; k <= need; ++k)
                if ((s[i + k] & 0xC0) != 0x80) return false;

            if (need == 2) {
                unsigned char c1 = s[i + 1];
                if ((c == 0xE0 && (c1 & 0xE0) == 0x80) ||            // overlong 3-byte
                    (c == 0xED && (c1 & 0xE0) == 0xA0))              // surrogate halves
                    return false;
            }
            else if (need == 3) {
                unsigned char c1 = s[i + 1];
                if ((c == 0xF0 && (c1 & 0xF0) == 0x80) ||            // overlong 4-byte
                    (c == 0xF4 && (c1 & 0xF0) != 0x80))              // > U+10FFFF
                    return false;
            }
            i += (need + 1);
        }
        return true;
    }
}

static long long g_cases = 0, g_mismatches = 0;

// Exact copy in a buffer of its own size, so the sanitizer sees any read past the end
static void check(const std::string& s, const char* what) {
    ++g_cases;
    std::vector<char> exact(s.begin(), s.end());
    const char* p = exact.empty() ? "" : exact.data();
    const bool want = reference::isValidUtf8(p, exact.size());
    const bool got = Encoding::isValidUtf8(p, exact.size());
    if (want == got) return;
    if (++g_mismatches <= 10) {
        std::printf("  MISMATCH [%s] reference %d new %d, %zu bytes:", what, want, got, s.size());
        for (size_t i = 0; i < s.size() && i < 40; ++i) std::printf(" %02X", static_cast<unsigned char>(s[i]));
        std::printf("\n");
    }
}

static int section(const char* name, long long before, long long mismatchesBefore) {
    std::printf("%-58s %10lld cases, %lld mismatches\n", name, g_cases - before, g_mismatches - mismatchesBefore);
    return 0;
}

// Leads and continuation bytes at every boundary the rules depend on
static const unsigned char kInteresting[] = {
    0x00, 0x41, 0x7F, 0x80, 0x8F, 0x90, 0x9F, 0xA0, 0xBF, 0xC0, 0xC1, 0xC2, 0xDF, 0xE0, 0xE1, 0xEC,
    0xED, 0xEE, 0xEF, 0xF0, 0xF1, 0xF3, 0xF4, 0xF5, 0xF7, 0xF8, 0xFB, 0xFC, 0xFE, 0xFF };

static void exhaustiveShort() {
    const long long c0 = g_cases, m0 = g_mismatches;
    std::string s;
    for (int a = 0; a < 256; ++a) {
        check(std::string(1, static_cast<char>(a)), "E1");
        for (int b = 0; b < 256; ++b) {
            check(std::string{ static_cast<char>(a), static_cast<char>(b) }, "E2");
            for (int c = 0; c < 256; ++c)
                check(std::string{ static_cast<char>(a), static_cast<char>(b), static_cast<char>(c) }, "E3");
        }
    }
    for (int a = 0xC0; a < 256; ++a)
        for (unsigned char b : kInteresting)
            for (unsigned char c : kInteresting)
                for (unsigned char d : kInteresting)
                    check(std::string{ static_cast<char>(a), static_cast<char>(b), static_cast<char>(c), static_cast<char>(d) }, "E4");
    section("E. one to three bytes complete, four over lead bytes", c0, m0);
}

static std::vector<std::string> probes() {
    std::vector<std::string> p;
    for (unsigned char a : kInteresting) {
        p.push_back(std::string(1, static_cast<char>(a)));
        for (unsigned char b : kInteresting) {
            p.push_back(std::string{ static_cast<char>(a), static_cast<char>(b) });
            if (a >= 0xE0)
                for (unsigned char c : { 0x80, 0x9F, 0xA0, 0xBF, 0x41 })
                    p.push_back(std::string{ static_cast<char>(a), static_cast<char>(b), static_cast<char>(c) });
        }
    }
    for (const char* v : { "\xC3\xA4", "\xE2\x82\xAC", "\xF0\x9F\x98\x80", "\xF4\x8F\xBF\xBF", "\xEF\xBB\xBF", "\xED\x9F\xBF", "\xEE\x80\x80" })
        p.push_back(v);
    return p;
}

static void boundaries() {
    const long long c0 = g_cases, m0 = g_mismatches;
    for (const std::string& probe : probes())
        for (size_t before = 0; before <= 24; ++before)
            for (size_t after = 0; after <= 16; ++after) {
                const std::string s = std::string(before, 'a') + probe + std::string(after, 'b');
                check(s, "B");
                check(s + probe, "B2");
            }
    section("B. probes at every offset in ASCII runs", c0, m0);
}

static void randomMix(uint64_t seed, int count) {
    const long long c0 = g_cases, m0 = g_mismatches;
    std::mt19937_64 rng(seed);
    static const char* valid[] = { "\xC3\xA4", "\xC2\x80", "\xDF\xBF", "\xE0\xA0\x80", "\xE2\x82\xAC", "\xED\x9F\xBF",
        "\xEF\xBF\xBD", "\xF0\x90\x80\x80", "\xF0\x9F\x98\x80", "\xF4\x8F\xBF\xBF", "\xE4\xB8\xAD" };
    for (int n = 0; n < count; ++n) {
        std::string s;
        const int parts = static_cast<int>(rng() % 12);
        for (int k = 0; k < parts; ++k) {
            switch (rng() % 6) {
            case 0: case 1: s.append(static_cast<size_t>(rng() % 40), static_cast<char>('a' + rng() % 26)); break;
            case 2: case 3: s += valid[rng() % (sizeof(valid) / sizeof(valid[0]))]; break;
            case 4: s += static_cast<char>(kInteresting[rng() % sizeof(kInteresting)]); break;
            default: { std::string v = valid[rng() % (sizeof(valid) / sizeof(valid[0]))]; s += v.substr(0, 1 + rng() % (v.size() - 1)); break; }
            }
        }
        if (!s.empty() && rng() % 4 == 0) s[rng() % s.size()] = static_cast<char>(rng() & 0xFF);
        check(s, "R");
    }
    section("R. random mixes", c0, m0);
}

static void corpus(const std::vector<std::string>& dirs) {
    const long long c0 = g_cases, m0 = g_mismatches;
    size_t files = 0;
    for (const std::string& dir : dirs) {
        std::error_code ec;
        for (auto it = std::filesystem::recursive_directory_iterator(dir, ec); !ec && it != std::filesystem::recursive_directory_iterator(); it.increment(ec)) {
            if (!it->is_regular_file(ec)) continue;
            std::ifstream in(it->path(), std::ios::binary);
            const std::string data((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
            ++files;
            check(data, "C");
            for (size_t cut = 1; cut <= 4 && cut <= data.size(); ++cut)
                check(data.substr(0, data.size() - cut), "C-cut");
        }
    }
    char name[96];
    std::snprintf(name, sizeof name, "C. corpus, %zu files", files);
    section(name, c0, m0);
}

// The text of isValidUtf8 in a source file, from its signature to the closing brace
static std::string functionText(const std::string& path, const char* signature, int skip) {
    std::ifstream in(path, std::ios::binary);
    if (!in) return {};
    std::string line, out;
    bool inside = false;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (!inside && line.find(signature) != std::string::npos && skip-- == 0) inside = true;
        if (!inside) continue;
        out += line + "\n";
        if (line == "    }") break;
    }
    return out;
}

static int verbatimCheck() {
    const std::filesystem::path here = std::filesystem::path(__FILE__).parent_path();
    const std::string self = (here / "utf8_validation_qa.cpp").string();
    const std::string product = (here / ".." / "Encoding.cpp").string();
    const char* signature = "bool isValidUtf8(const char* data, size_t len) {";
    const std::string mine = functionText(self, signature, 1);   // the first one is the reference
    const std::string theirs = functionText(product, signature, 0);
    if (mine.empty() || theirs.empty()) { std::printf("V. copy vs ../Encoding.cpp: SKIP (source not found)\n"); return 0; }
    const bool same = (mine == theirs);
    std::printf("V. copy vs ../Encoding.cpp: %s\n", same ? "identical" : "DIFFERENT");
    return same ? 0 : 1;
}

int main(int argc, char** argv) {
    std::printf("Encoding::isValidUtf8 with the eight-byte ASCII step against a43e162\n\n");
    exhaustiveShort();
    boundaries();
    randomMix(20261010, 2000000);
    std::vector<std::string> dirs(argv + 1, argv + argc);
    if (!dirs.empty()) corpus(dirs);
    const int verbatim = verbatimCheck();
    std::printf("\n%lld cases, %lld mismatches: %s\n", g_cases, g_mismatches, (g_mismatches == 0 && verbatim == 0) ? "PASS" : "FAIL");
    return (g_mismatches == 0 && verbatim == 0) ? 0 : 1;
}
