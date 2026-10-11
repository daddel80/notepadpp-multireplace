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

// QA harness for LiteralPrefilter (Find/Replace in Files). Whenever it says a search cannot
// match, the real Scintilla search must find nothing: Document::FindText of Notepad++ 697f46b
// with the case folders Scintilla uses on Windows (UTF-8: CaseFolderUnicode; code page 0:
// a code page 1252 table as ScintillaWin builds it).
//   T. targeted: folding beyond ASCII (ß and ẞ to ss, Kelvin sign, long s, ligatures, dotted I,
//      final sigma, Ångström and micro sign), control bytes, NUL, CR/LF, invalid UTF-8, code page 0
//   F. random texts and patterns over an alphabet of exactly those characters, with and without
//      match case and whole word, over the whole text and over random ranges (column mode
//      searches ranges of a line)
//   C. optional folders: per file, patterns cut from the file itself, taken from other files,
//      with changed case and one changed byte; plus the per-file flow of Find in Files: the hits
//      with the filter equal the hits without it
// Reported: how often "cannot match" was wrong (must be 0) and how often the filter decided.
//
// Needs the Scintilla objects of find_in_files_bench.sh (<work> is its work folder).
// Build (from src/tests, w = <work>):
//   g++ -std=c++17 -O1 -Wall -Wextra -fsanitize=address,undefined -I.. -I$w/npp/scintilla/include
//       -I$w/npp/scintilla/src literal_prefilter_qa.cpp $w/obj/*.o -o literal_prefilter_qa
// Run:   ./literal_prefilter_qa [folder ...]
#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <forward_list>
#include <fstream>
#include <iterator>
#include <map>
#include <memory>
#include <optional>
#include <random>
#include <set>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>

#include "LiteralPrefilter.h"

#include "ScintillaTypes.h"
#include "ILoader.h"
#include "ILexer.h"
#include "Debugging.h"
#include "Geometry.h"
#include "Platform.h"
#include "CharacterType.h"
#include "CharacterCategoryMap.h"
#include "Position.h"
#include "SplitVector.h"
#include "Partitioning.h"
#include "RunStyles.h"
#include "CellBuffer.h"
#include "PerLine.h"
#include "CharClassify.h"
#include "Decoration.h"
#include "CaseFolder.h"
#include "Document.h"

using namespace Scintilla;
using namespace Scintilla::Internal;
using namespace std::literals;

namespace Scintilla::Internal {
void Platform::Assert(const char* c, const char* file, int line) noexcept { std::fprintf(stderr, "Assertion [%s] failed at %s %d\n", c, file, line); std::abort(); }
void Platform::DebugPrintf(const char*, ...) noexcept {}
void Platform::DebugDisplay(const char*) noexcept {}
bool Platform::ShowAssertionPopUps(bool) noexcept { return false; }
ColourRGBA Platform::Chrome() { return ColourRGBA(0xe0, 0xe0, 0xe0); }
ColourRGBA Platform::ChromeHighlight() { return ColourRGBA(0xff, 0xff, 0xff); }
const char* Platform::DefaultFont() { return "monospace"; }
int Platform::DefaultFontSize() { return 10; }
unsigned int Platform::DoubleClickTime() { return 500; }
}

// A Scintilla document as MR's hidden buffer holds it: no styles, no undo
class SciText {
public:
    SciText(const std::string& text, int codepage) {
        _d = new Document(static_cast<DocumentOption>(0x1 | 0x100));
        _d->AddRef();
        _d->SetUndoCollection(false);
        _d->SetDBCSCodePage(codepage);
        if (codepage == SC_CP_UTF8) _d->SetCaseFolder(std::make_unique<CaseFolderUnicode>());
        else _d->SetCaseFolder(cp1252Folder());
        if (!text.empty()) _d->InsertString(0, text.data(), static_cast<Sci::Position>(text.size()));
    }
    ~SciText() { _d->Release(); }
    SciText(const SciText&) = delete;
    SciText& operator=(const SciText&) = delete;

    Sci::Position length() const { return _d->Length(); }

    // SCI_SETTARGETRANGE + SCI_SEARCHINTARGET as MR's performSingleSearch takes it: with match case
    // Scintilla can report a match running past the end of the range, at the end of the document
    // into bytes that are not there (a pattern ending in NUL); MR drops a match past the document
    // end, a match past the range end is real text and counts
    Sci::Position find(Sci::Position from, Sci::Position to, const std::string& pattern, int flags, Sci::Position& len) const {
        len = static_cast<Sci::Position>(pattern.size());
        const Sci::Position p = _d->FindText(from, to, pattern.c_str(), static_cast<FindOption>(flags), &len);
        return (p >= 0 && p + len > length()) ? -1 : p;
    }

    // MR's collect loop: every hit from the start, empty patterns never searched
    std::vector<std::pair<Sci::Position, Sci::Position>> allHits(const std::string& pattern, int flags) const {
        std::vector<std::pair<Sci::Position, Sci::Position>> hits;
        if (pattern.empty()) return hits;
        Sci::Position pos = 0;
        const Sci::Position end = length();
        while (pos <= end) {
            Sci::Position len = 0;
            const Sci::Position p = find(pos, end, pattern, flags, len);
            if (p < 0) break;
            hits.push_back({ p, len });
            pos = len > 0 ? p + len : std::max(_d->NextPosition(p, 1), p + 1);
        }
        return hits;
    }

private:
    // ScintillaWin::CaseFolderForEncoding for code page 1252
    static std::unique_ptr<CaseFolder> cp1252Folder() {
        auto f = std::make_unique<CaseFolderTable>();
        for (int ch = 0xC0; ch <= 0xDE; ++ch)
            if (ch != 0xD7) f->SetTranslation(static_cast<char>(ch), static_cast<char>(ch + 0x20));
        f->SetTranslation('\x8A', '\x9A');
        f->SetTranslation('\x8C', '\x9C');
        f->SetTranslation('\x8E', '\x9E');
        f->SetTranslation('\x9F', '\xFF');
        return f;
    }

    Document* _d = nullptr;
};

struct Stats { long long checks = 0, wrong = 0, decided = 0, sciNone = 0; };
static Stats g_all;

static std::string hex(const std::string& s) {
    std::string out;
    char b[4];
    for (size_t i = 0; i < s.size() && i < 24; ++i) { std::snprintf(b, sizeof b, "%02X ", static_cast<unsigned char>(s[i])); out += b; }
    return out;
}

static const char* flagName(int flags) {
    switch (flags & (SCFIND_MATCHCASE | SCFIND_WHOLEWORD | SCFIND_WORDSTART)) {
    case 0: return "ignore case";
    case SCFIND_MATCHCASE: return "match case";
    case SCFIND_WHOLEWORD: return "ignore case, whole word";
    case SCFIND_MATCHCASE | SCFIND_WHOLEWORD: return "match case, whole word";
    case SCFIND_WORDSTART: return "ignore case, word start";
    default: return "other";
    }
}

// The claim: mayMatch false implies no Scintilla match in [from, to]
static void check(Stats& st, const SciText& sci, LiteralPrefilter& filter, const std::string& pattern, int flags,
    Sci::Position from, Sci::Position to, const char* where)
{
    ++st.checks;
    const bool may = filter.mayMatch(pattern, flags);
    Sci::Position len = 0;
    const Sci::Position hit = pattern.empty() ? -1 : sci.find(from, to, pattern, flags, len);
    if (hit < 0) ++st.sciNone;
    if (!may) ++st.decided;
    if (!may && hit >= 0) {
        if (++st.wrong <= 10)
            std::printf("  WRONG [%s] %s: pattern %s matched at %ld\n", where, flagName(flags), hex(pattern).c_str(), static_cast<long>(hit));
    }
}

static const int kFlagSets[] = { 0, SCFIND_MATCHCASE, SCFIND_WHOLEWORD, SCFIND_MATCHCASE | SCFIND_WHOLEWORD, SCFIND_WORDSTART };

static void report(const char* name, const Stats& st) {
    std::printf("%-44s %9lld checks, wrong %lld, decided %lld of %lld without a match\n", name, st.checks, st.wrong, st.decided, st.sciNone);
    g_all.checks += st.checks; g_all.wrong += st.wrong; g_all.decided += st.decided; g_all.sciNone += st.sciNone;
}

// ------------------------------------------------------------------------------------- T
static void targeted() {
    struct Case { std::string_view text, pattern; int codepage; };
    static const Case cases[] = {
        { "Straße und Strasse"sv, "STRASSE"sv, SC_CP_UTF8 }, { "Straße"sv, "strasse"sv, SC_CP_UTF8 }, { "Straße"sv, "ss"sv, SC_CP_UTF8 },
        { "GROẞE"sv, "grosse"sv, SC_CP_UTF8 }, { "\xE2\x84\xAA" "elvin"sv, "kelvin"sv, SC_CP_UTF8 }, { "\xE2\x84\xAA"sv, "k"sv, SC_CP_UTF8 },
        { "Me\xC5\xBF" "s"sv, "mess"sv, SC_CP_UTF8 }, { "\xC5\xBF"sv, "S"sv, SC_CP_UTF8 }, { "\xEF\xAC\x81" "le"sv, "file"sv, SC_CP_UTF8 },
        { "\xEF\xAC\x81"sv, "fi"sv, SC_CP_UTF8 }, { "\xC4\xB0stanbul"sv, "istanbul"sv, SC_CP_UTF8 }, { "\xC4\xB0"sv, "i"sv, SC_CP_UTF8 },
        { "\xCF\x82\xCF\x83\xCE\xA3"sv, "\xCF\x83"sv, SC_CP_UTF8 }, { "\xE2\x84\xAB" "ngstr\xC3\xB6m"sv, "\xC3\xA5"sv, SC_CP_UTF8 },
        { "\xC2\xB5m"sv, "\xCE\xBC"sv, SC_CP_UTF8 }, { "\xC3\x84rger"sv, "\xC3\xA4rger"sv, SC_CP_UTF8 }, { "\xC3\xA4rger"sv, "\xC3\x84RGER"sv, SC_CP_UTF8 },
        { "plain ascii text"sv, "ASCII"sv, SC_CP_UTF8 }, { "plain ascii text"sv, "xyz"sv, SC_CP_UTF8 }, { "plain ascii text"sv, "PLAIN TEXT"sv, SC_CP_UTF8 },
        { "a\r\nb"sv, "\r\n"sv, SC_CP_UTF8 }, { "a\r\nb"sv, "\n\r"sv, SC_CP_UTF8 }, { "tab\there"sv, "\t"sv, SC_CP_UTF8 },
        { "nul\x01" "byte"sv, "\x01"sv, SC_CP_UTF8 }, { "bad \xC3 lead"sv, "\xC3"sv, SC_CP_UTF8 }, { "bad \x80 cont"sv, "\x80"sv, SC_CP_UTF8 },
        { "\xC3\xA4"sv, "\xA4"sv, SC_CP_UTF8 }, { "\xC3\xA4"sv, "\xC3"sv, SC_CP_UTF8 }, { "word words sword"sv, "WORD"sv, SC_CP_UTF8 },
        { "Stra\xDF" "e"sv, "STRASSE"sv, 0 }, { "Stra\xDF" "e"sv, "stra\xDF" "e"sv, 0 }, { "\xC4rger"sv, "\xE4rger"sv, 0 }, { "\xC4rger"sv, "\xE4RGER"sv, 0 },
        { "\x8A" "koda"sv, "\x9A" "KODA"sv, 0 }, { "raw \0 bytes"sv, "RAW"sv, 0 }, { "raw \0 bytes"sv, "\0 b"sv, 0 }, { "raw \0 bytes"sv, "w \0"sv, 0 },
        { "raw \0 bytes"sv, "\0\0"sv, 0 }, { "raw \0 bytes"sv, "\0"sv, SC_CP_UTF8 }, { "raw ascii"sv, "ASCII"sv, 0 }, { "raw ascii"sv, "zz"sv, 0 },
        { ""sv, "a"sv, SC_CP_UTF8 }, { "a"sv, "ab"sv, SC_CP_UTF8 }, { "ab"sv, "ab"sv, 0 },
    };
    Stats st;
    for (const Case& c : cases) {
        const std::string text(c.text);
        SciText sci(text, c.codepage);
        LiteralPrefilter filter(text);
        for (int flags : kFlagSets)
            check(st, sci, filter, std::string(c.pattern), flags, 0, sci.length(), "T");
    }
    report("T. targeted cases", st);
}

// ------------------------------------------------------------------------------------- F
static void fuzz(uint64_t seed, int texts) {
    static const std::vector<std::string> alphabet = {
        "a", "A", "b", "B", "k", "K", "s", "S", "f", "F", "i", "I", "l", "L", "e", "E", "t", "T", "x", " ", "_", "1", "-",
        "\xC3\x9F", "\xE1\xBA\x9E", "\xE2\x84\xAA", "\xC5\xBF", "\xEF\xAC\x81", "\xEF\xAC\x82", "\xEF\xAC\x80", "\xEF\xAC\x86",
        "\xC4\xB0", "\xC4\xB1", "\xCF\x83", "\xCF\x82", "\xCE\xA3", "\xE2\x84\xAB", "\xC3\xA5", "\xC3\x85", "\xC2\xB5", "\xCE\xBC",
        "\xCE\x9C", "\xC3\xA4", "\xC3\x84", "\r", "\n", "\r\n", "\t", std::string(1, '\0'), "\x01", "\xC3", "\x80", "\xFF" };
    auto token = [&](std::mt19937_64& rng, bool asciiOnly) -> const std::string& {
        for (;;) {
            const std::string& t = alphabet[rng() % alphabet.size()];
            if (!asciiOnly || LiteralPrefilter::isAscii(t)) return t;
        }
    };
    std::mt19937_64 rng(seed);
    Stats st;
    for (int t = 0; t < texts; ++t) {
        const bool asciiText = (rng() % 3) == 0;
        std::string text;
        const int tokens = static_cast<int>(rng() % 60);
        for (int i = 0; i < tokens; ++i) text += token(rng, asciiText);
        for (int codepage : { SC_CP_UTF8, 0 }) {
            SciText sci(text, codepage);
            LiteralPrefilter filter(text);
            for (int p = 0; p < 12; ++p) {
                std::string pattern;
                if (p < 4 && !text.empty()) {
                    const size_t a = rng() % text.size();
                    pattern = text.substr(a, 1 + rng() % std::min<size_t>(6, text.size() - a));
                    if (p & 1) for (auto& ch : pattern) if (std::isalpha(static_cast<unsigned char>(ch))) ch ^= 0x20;
                }
                else {
                    const int len = 1 + static_cast<int>(rng() % 4);
                    for (int i = 0; i < len; ++i) pattern += token(rng, (p & 1) != 0);
                }
                for (int flags : kFlagSets) {
                    check(st, sci, filter, pattern, flags, 0, sci.length(), "F");
                    const Sci::Position len = sci.length();
                    if (len > 0) {
                        Sci::Position a = static_cast<Sci::Position>(rng() % (len + 1)), b = static_cast<Sci::Position>(rng() % (len + 1));
                        if (a > b) std::swap(a, b);
                        check(st, sci, filter, pattern, flags, a, b, "F-range");
                    }
                }
            }
        }
    }
    report("F. random texts and patterns", st);
}

// ------------------------------------------------------------------------------------- C
static bool isUtf8(const std::string& s) {
    const unsigned char* p = reinterpret_cast<const unsigned char*>(s.data());
    size_t i = 0;
    while (i < s.size()) {
        const unsigned char c = p[i];
        if (c < 0x80) { ++i; continue; }
        size_t need = (c >= 0xC2 && c <= 0xDF) ? 1 : (c & 0xF0) == 0xE0 ? 2 : (c >= 0xF0 && c <= 0xF4) ? 3 : 0;
        if (!need || i + need >= s.size()) return false;
        for (size_t k = 1; k <= need; ++k) if ((p[i + k] & 0xC0) != 0x80) return false;
        i += need + 1;
    }
    return true;
}

static void appendUtf8(std::string& out, uint32_t cp) {
    if (cp < 0x80) out += static_cast<char>(cp);
    else if (cp < 0x800) { out += static_cast<char>(0xC0 | (cp >> 6)); out += static_cast<char>(0x80 | (cp & 0x3F)); }
    else { out += static_cast<char>(0xE0 | (cp >> 12)); out += static_cast<char>(0x80 | ((cp >> 6) & 0x3F)); out += static_cast<char>(0x80 | (cp & 0x3F)); }
}

// The text MR puts into the hidden buffer: UTF-8 as is, UTF-16 and ANSI (here: Latin-1) decoded;
// binaries as raw bytes in code page 0 unless they are valid UTF-8
static bool loadLikeMr(const std::string& raw, std::string& text, int& codepage) {
    const bool bom8 = raw.size() >= 3 && raw.compare(0, 3, "\xEF\xBB\xBF") == 0;
    const bool bom16le = raw.size() >= 2 && raw.compare(0, 2, "\xFF\xFE") == 0;
    const bool binary = !bom8 && !bom16le && std::memchr(raw.data(), 0, std::min<size_t>(raw.size(), 8192)) != nullptr;
    codepage = SC_CP_UTF8;
    if (binary) { text = raw; if (!isUtf8(raw)) codepage = 0; return true; }
    if (bom8) { text = raw.substr(3); return true; }
    if (bom16le) {
        text.clear();
        for (size_t i = 2; i + 1 < raw.size(); i += 2) appendUtf8(text, static_cast<unsigned char>(raw[i]) | (static_cast<unsigned char>(raw[i + 1]) << 8));
        return true;
    }
    if (isUtf8(raw)) { text = raw; return true; }
    text.clear();
    for (unsigned char c : raw) appendUtf8(text, c);
    return true;
}

static void corpus(const std::vector<std::string>& dirs, uint64_t seed) {
    std::vector<std::filesystem::path> files;
    for (const std::string& d : dirs) {
        std::error_code ec;
        for (auto it = std::filesystem::recursive_directory_iterator(d, ec); !ec && it != std::filesystem::recursive_directory_iterator(); it.increment(ec))
            if (it->is_regular_file(ec)) files.push_back(it->path());
    }
    std::sort(files.begin(), files.end());
    std::mt19937_64 rng(seed);
    std::vector<std::string> words;   // collected from earlier files
    Stats st;
    long long flowFiles = 0, flowSkipped = 0, flowMismatch = 0;
    for (const auto& path : files) {
        std::ifstream in(path, std::ios::binary);
        const std::string raw((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
        std::string text;
        int codepage = SC_CP_UTF8;
        if (!loadLikeMr(raw, text, codepage)) continue;
        SciText sci(text, codepage);
        LiteralPrefilter filter(text);

        std::vector<std::string> patterns = { "zqxjvqk", "WM_SETREDRAW", "Fi", "in", "cache", "Cache", "stra\xC3\x9F" "e", "\r\n\r\n", "\t\t" };
        for (int k = 0; k < 8 && !text.empty(); ++k) {
            size_t a = rng() % text.size();
            if (codepage == SC_CP_UTF8) while (a > 0 && (static_cast<unsigned char>(text[a]) & 0xC0) == 0x80) --a;
            std::string cut = text.substr(a, 1 + rng() % std::min<size_t>(16, text.size() - a));
            patterns.push_back(cut);
            std::string flipped = cut;
            for (auto& ch : flipped) if (std::isalpha(static_cast<unsigned char>(ch))) ch ^= 0x20;
            patterns.push_back(flipped);
            std::string mutated = cut;
            mutated[rng() % mutated.size()] = static_cast<char>(rng() & 0x7F);
            patterns.push_back(mutated);
            if (words.size() < 4000 && cut.size() >= 3) words.push_back(cut);
        }
        for (int k = 0; k < 6 && !words.empty(); ++k) patterns.push_back(words[rng() % words.size()]);

        for (const std::string& p : patterns)
            for (int flags : kFlagSets) check(st, sci, filter, p, flags, 0, sci.length(), "C");

        // Find in Files per file: a list of patterns, the file is searched only if one may match
        ++flowFiles;
        std::vector<std::pair<std::string, int>> list;
        for (int k = 0; k < 3; ++k) list.push_back({ patterns[rng() % patterns.size()], kFlagSets[rng() % 5] });
        std::vector<std::pair<Sci::Position, Sci::Position>> without, with;
        for (const auto& [p, f] : list) { auto h = sci.allHits(p, f); without.insert(without.end(), h.begin(), h.end()); }
        bool any = false;
        for (const auto& [p, f] : list) any = any || filter.mayMatch(p, f);
        if (!any) ++flowSkipped;
        else for (const auto& [p, f] : list) if (filter.mayMatch(p, f)) { auto h = sci.allHits(p, f); with.insert(with.end(), h.begin(), h.end()); }
        if (with != without) { if (++flowMismatch <= 5) std::printf("  FLOW MISMATCH in %s\n", path.string().c_str()); }
    }
    char name[96];
    std::snprintf(name, sizeof name, "C. corpus, %zu files", files.size());
    report(name, st);
    std::printf("%-44s %9lld files, %lld skipped by the filter, hits differ in %lld\n", "   Find in Files flow with three patterns", flowFiles, flowSkipped, flowMismatch);
    g_all.wrong += flowMismatch;
}

int main(int argc, char** argv) {
    std::printf("LiteralPrefilter against Scintilla's FindText (Notepad++ 697f46b)\n\n");
    targeted();
    fuzz(20261010, 20000);
    std::vector<std::string> dirs(argv + 1, argv + argc);
    if (!dirs.empty()) corpus(dirs, 20261010);
    std::printf("\n%lld checks, \"cannot match\" wrong %lld times: %s\n", g_all.checks, g_all.wrong, g_all.wrong == 0 ? "PASS" : "FAIL");
    return g_all.wrong == 0 ? 0 : 1;
}
