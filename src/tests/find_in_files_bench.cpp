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

// Test bench for Find in Files. It runs the per-file and per-hit work of three implementations
// on the same corpus and compares them statistically:
//   MR      MultiReplace a43e162: HiddenSciGuard::loadTextFile, Encoding::detectEncoding,
//           the collect loop of handleFindInFiles, ResultDock::appendFileBlock/closeSearchBlock
//   NppNew  Notepad++ 697f46b, which contains 20d8f02 "Improve Find in Files performance":
//           findInFilelist, FileManager::loadFileContentForSearch, processRange(ProcessFindAll), Finder
//   NppOld  Notepad++ 572650c, the state before 20d8f02: a full Buffer per file (loadFile, closeBuffer)
// MR's later steps run as variants of MR, alone or together:
//   MR+utf8       Encoding::isValidUtf8 with the eight-byte ASCII step
//   MR+prefilter  LiteralPrefilter (../LiteralPrefilter.h) before a file is filled
//   MR+dock       RawHit/CritAgg per search, one RowMatch list per dock line, line numbers
//                 and matches as styles set in one call
//   MR+all        all three
// The Notepad++ parts are ports of the Notepad++ sources (GPL v3). All three run on the real
// Scintilla document, the Notepad++ Boost regex bridge and the uchardet bundled with Notepad++;
// Win32 text conversion is replaced by equivalent code for UTF-8 and code page 1252.
// Not measured: Windows file system and virus scanner, window messages and painting, and the
// per-call work of the Scintilla window (Editor) around the document.
//
// Build, corpora and baseline run: find_in_files_bench.sh
// Usage: find_in_files_bench [--source DIR] [--source-text DIR] [--mixed DIR] [--rounds N]
//                            [--seed N] [--tools MR,NppNew,NppOld,...] [--queries NAME,...] [--out DIR]
//        find_in_files_bench --generate-mixed DIR [--gen-seed N] [--gen-files N]

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <cstdarg>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <ctime>
#include <filesystem>
#include <forward_list>
#include <fstream>
#include <map>
#include <memory>
#include <numeric>
#include <optional>
#include <random>
#include <set>
#include <stdexcept>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>
#include <fcntl.h>
#include <strings.h>
#include <sys/stat.h>
#include <sys/utsname.h>
#include <unistd.h>

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
#include "uchardet.h"
#include "LiteralPrefilter.h"

using namespace Scintilla;
using namespace Scintilla::Internal;

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

namespace sc {
constexpr int WHOLEWORD = 0x2, MATCHCASE = 0x4, REGEXP = 0x00200000, POSIX = 0x00400000;
constexpr int SKIPCRLFASONE = 0x08000000, EMPTY_NOTAFTER = 0x20000000, EMPTY_ALL = 0x40000000;
constexpr int CP_UTF8 = 65001;
}

using Clock = std::chrono::steady_clock;
static double msBetween(Clock::time_point a, Clock::time_point b) { return std::chrono::duration<double, std::milli>(b - a).count(); }

// ---------------------------------------------------------------------------------------------
// Win32 text conversion stand-ins (UTF-16 code units as on Windows, ANSI code page 1252)
// ---------------------------------------------------------------------------------------------
namespace w32 {
using wchar = char16_t;
constexpr unsigned ACP = 1252;

static const char16_t k1252[32] = {
    0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
    0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178 };

inline char16_t sbToU(unsigned char c) { return (c >= 0x80 && c < 0xA0) ? k1252[c - 0x80] : c; }
inline int uToSb(char16_t u) {
    if (u < 0x80 || (u >= 0xA0 && u <= 0xFF)) return u;
    for (int i = 0; i < 32; ++i) if (k1252[i] == u) return 0x80 + i;
    return -1;
}

// MultiByteToWideChar without MB_ERR_INVALID_CHARS: invalid bytes become U+FFFD
int multiByteToWide(unsigned cp, const char* s, int n, wchar* out) {
    const unsigned char* p = reinterpret_cast<const unsigned char*>(s);
    if (cp != sc::CP_UTF8) {
        if (out) for (int i = 0; i < n; ++i) out[i] = sbToU(p[i]);
        return n;
    }
    int w = 0, i = 0;
    while (i < n) {
        const unsigned char c = p[i];
        if (c < 0x80) { if (out) out[w] = c; ++w; ++i; continue; }
        int need = 0; uint32_t v = 0;
        if (c >= 0xC2 && c <= 0xDF) { need = 1; v = c & 0x1F; }
        else if ((c & 0xF0) == 0xE0) { need = 2; v = c & 0x0F; }
        else if (c >= 0xF0 && c <= 0xF4) { need = 3; v = c & 0x07; }
        bool ok = need > 0 && i + need < n;
        for (int k = 1; ok && k <= need; ++k) {
            const unsigned char t = p[i + k];
            if ((t & 0xC0) != 0x80) ok = false; else v = (v << 6) | (t & 0x3F);
        }
        if (ok && ((need == 2 && (v < 0x800 || (v >= 0xD800 && v <= 0xDFFF))) || (need == 3 && (v < 0x10000 || v > 0x10FFFF)))) ok = false;
        if (!ok) { if (out) out[w] = 0xFFFD; ++w; ++i; continue; }
        if (v >= 0x10000) {
            if (out) { v -= 0x10000; out[w] = static_cast<wchar>(0xD800 + (v >> 10)); out[w + 1] = static_cast<wchar>(0xDC00 + (v & 0x3FF)); }
            w += 2;
        }
        else { if (out) out[w] = static_cast<wchar>(v); ++w; }
        i += need + 1;
    }
    return w;
}

int wideToMultiByte(unsigned cp, const wchar* s, int n, char* out) {
    int b = 0;
    if (cp != sc::CP_UTF8) {
        for (int i = 0; i < n; ++i) { const int c = uToSb(s[i]); if (out) out[b] = static_cast<char>(c < 0 ? '?' : c); ++b; }
        return b;
    }
    for (int i = 0; i < n; ++i) {
        uint32_t c = s[i];
        if (c >= 0xD800 && c <= 0xDBFF && i + 1 < n && s[i + 1] >= 0xDC00 && s[i + 1] <= 0xDFFF) { c = 0x10000 + ((c - 0xD800) << 10) + (s[i + 1] - 0xDC00); ++i; }
        else if (c >= 0xD800 && c <= 0xDFFF) c = 0xFFFD;
        if (c < 0x80) { if (out) out[b] = static_cast<char>(c); b += 1; }
        else if (c < 0x800) { if (out) { out[b] = static_cast<char>(0xC0 | (c >> 6)); out[b + 1] = static_cast<char>(0x80 | (c & 0x3F)); } b += 2; }
        else if (c < 0x10000) { if (out) { out[b] = static_cast<char>(0xE0 | (c >> 12)); out[b + 1] = static_cast<char>(0x80 | ((c >> 6) & 0x3F)); out[b + 2] = static_cast<char>(0x80 | (c & 0x3F)); } b += 3; }
        else { if (out) { out[b] = static_cast<char>(0xF0 | (c >> 18)); out[b + 1] = static_cast<char>(0x80 | ((c >> 12) & 0x3F)); out[b + 2] = static_cast<char>(0x80 | ((c >> 6) & 0x3F)); out[b + 3] = static_cast<char>(0x80 | (c & 0x3F)); } b += 4; }
    }
    return b;
}

inline bool isDbcsLeadByte(unsigned cp, unsigned char b) {
    if (cp == 932) return (b >= 0x81 && b <= 0x9F) || (b >= 0xE0 && b <= 0xFC);
    return b >= 0x81 && b <= 0xFE;
}

// MultiByteToWideChar(MB_ERR_INVALID_CHARS) on a CJK code page: structure check only
inline bool dbcsValid(unsigned cp, const char* s, int n) {
    const unsigned char* p = reinterpret_cast<const unsigned char*>(s);
    for (int i = 0; i < n; ++i) {
        const unsigned char b = p[i];
        if (b < 0x80) continue;
        if (cp == 932 && b >= 0xA1 && b <= 0xDF) continue;
        if (!isDbcsLeadByte(cp, b) || i + 1 >= n) return false;
        const unsigned char t = p[i + 1];
        const bool trail = (cp == 932) ? ((t >= 0x40 && t <= 0x7E) || (t >= 0x80 && t <= 0xFC)) : (t >= 0x40 && t <= 0xFE && t != 0x7F);
        if (!trail) return false;
        ++i;
    }
    return true;
}

inline size_t wlen(const wchar* s) { size_t n = 0; while (s[n]) ++n; return n; }

std::u16string widen(const std::string& u8) {
    std::u16string w(static_cast<size_t>(multiByteToWide(sc::CP_UTF8, u8.data(), static_cast<int>(u8.size()), nullptr)), u'\0');
    multiByteToWide(sc::CP_UTF8, u8.data(), static_cast<int>(u8.size()), w.data());
    return w;
}
std::string narrow(const std::u16string& w) {
    std::string s(static_cast<size_t>(wideToMultiByte(sc::CP_UTF8, w.data(), static_cast<int>(w.size()), nullptr)), '\0');
    wideToMultiByte(sc::CP_UTF8, w.data(), static_cast<int>(w.size()), s.data());
    return s;
}
}

// ---------------------------------------------------------------------------------------------
// Scintilla document driven like the SCI_* messages the tools send
// ---------------------------------------------------------------------------------------------
struct SciDoc {
    Document* d = nullptr;
    int cp = -1;
    bool folder = false;

    // SC_DOCUMENTOPTION_TEXT_LARGE (0x100), without styles also SC_DOCUMENTOPTION_STYLES_NONE (0x1)
    explicit SciDoc(bool styles) {
        d = new Document(static_cast<DocumentOption>(styles ? 0x100 : (0x1 | 0x100)));
        d->AddRef();
        d->SetUndoCollection(false);
        setCodePage(sc::CP_UTF8);
    }
    ~SciDoc() { d->Release(); }
    SciDoc(const SciDoc&) = delete;
    SciDoc& operator=(const SciDoc&) = delete;

    void setCodePage(int c) {                    // SCI_SETCODEPAGE: a change drops the case folder
        if (c == cp) return;
        cp = c;
        d->SetDBCSCodePage(c);
        folder = false;
    }
    void ensureFolder() {                        // Editor::SearchInTarget -> CaseFolderForEncoding
        if (folder) return;
        if (cp == sc::CP_UTF8) d->SetCaseFolder(std::make_unique<CaseFolderUnicode>());
        else {
            auto f = std::make_unique<CaseFolderTable>();
            f->StandardASCII();
            for (int ch = 0xC0; ch <= 0xDE; ++ch) if (ch != 0xD7) f->SetTranslation(static_cast<char>(ch), static_cast<char>(ch + 0x20));
            d->SetCaseFolder(std::move(f));
        }
        folder = true;
    }
    void clearAll() { if (d->Length()) d->DeleteChars(0, d->Length()); }
    void allocate(Sci::Position n) { d->Allocate(n); }
    void append(const char* s, Sci::Position n) { if (n > 0) d->InsertString(d->Length(), s, n); }
    void insert(Sci::Position at, const char* s, Sci::Position n) { if (n > 0) d->InsertString(at, s, n); }
    Sci::Position length() const { return d->Length(); }
    Sci::Line lineCount() const { return d->LinesTotal(); }
    Sci::Line lineFromPos(Sci::Position p) const { return d->SciLineFromPosition(p); }
    Sci::Position lineStart(Sci::Line l) const { return d->LineStart(l); }
    Sci::Position lineEnd(Sci::Line l) const { return d->LineEnd(l); }
    Sci::Position lineLength(Sci::Line l) const { return d->LineStart(l + 1) - d->LineStart(l); }
    void getText(char* buf, Sci::Position start, Sci::Position end) const { d->GetCharRange(buf, start, end - start); buf[end - start] = '\0'; }
    Sci::Position positionAfter(Sci::Position p) const { return d->NextPosition(p, 1); }
    // SCI_SETTARGETRANGE + SCI_SEARCHINTARGET (+ SCI_GETTARGETEND)
    Sci::Position searchInTarget(Sci::Position from, Sci::Position to, const char* s, size_t n, int flags, Sci::Position& end) {
        ensureFolder();
        Sci::Position len = static_cast<Sci::Position>(n);
        const Sci::Position p = d->FindText(from, to, s, static_cast<FindOption>(flags), &len);
        if (p >= 0) end = p + len;
        return p;
    }
};

// ---------------------------------------------------------------------------------------------
// Corpus and queries
// ---------------------------------------------------------------------------------------------
struct FileEntry { std::string path; std::u16string path16; uint64_t size = 0; bool binary = false; };

struct Corpus {
    std::string name, dir;
    std::vector<FileEntry> files;
    uint64_t bytes = 0, fingerprint = 0;
    size_t binaryFiles = 0;
};

struct Query { std::string name, text; bool regex = false, matchCase = true; };

static uint64_t fnv1a(uint64_t h, const void* data, size_t n) {
    const unsigned char* p = static_cast<const unsigned char*>(data);
    for (size_t i = 0; i < n; ++i) { h ^= p[i]; h *= 1099511628211ull; }
    return h;
}

static bool readWholeFile(const std::string& path, std::string& out) {
    FILE* f = std::fopen(path.c_str(), "rb");
    if (!f) return false;
    std::fseek(f, 0, SEEK_END); const long n = std::ftell(f); std::fseek(f, 0, SEEK_SET);
    out.resize(n > 0 ? static_cast<size_t>(n) : 0);
    const size_t got = n > 0 ? std::fread(out.data(), 1, out.size(), f) : 0;
    std::fclose(f);
    out.resize(got);
    return true;
}

// MR's binary rule (NUL in the first 8 KB without BOM), used to separate hits Notepad++ finds in binaries
static bool mrBinaryRule(const std::string& s) {
    const unsigned char* u = reinterpret_cast<const unsigned char*>(s.data());
    const size_t n = s.size();
    if (n >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) return false;
    if (n >= 2 && ((u[0] == 0xFF && u[1] == 0xFE) || (u[0] == 0xFE && u[1] == 0xFF))) return false;
    return std::memchr(s.data(), 0, std::min<size_t>(n, 8192)) != nullptr;
}

// textOnly leaves out the files MR's rule calls binary, as if the folder held no binaries
static Corpus scanCorpus(const std::string& name, const std::string& dir, bool textOnly) {
    Corpus c; c.name = name; c.dir = dir;
    std::vector<std::string> paths;
    for (auto& e : std::filesystem::recursive_directory_iterator(dir)) if (e.is_regular_file()) paths.push_back(e.path().string());
    std::sort(paths.begin(), paths.end());
    uint64_t h = 1469598103934665603ull;
    std::string buf;
    for (const auto& p : paths) {
        FileEntry fe; fe.path = p; fe.path16 = w32::widen(p);
        readWholeFile(p, buf);
        fe.size = buf.size(); fe.binary = mrBinaryRule(buf);
        if (textOnly && fe.binary) continue;
        const std::string rel = p.substr(dir.size());
        h = fnv1a(h, rel.data(), rel.size());
        h = fnv1a(h, buf.data(), buf.size());
        c.bytes += fe.size; c.binaryFiles += fe.binary;
        c.files.push_back(std::move(fe));
    }
    c.fingerprint = h;
    return c;
}

// Deterministic mixed-encoding text corpus
static void generateMixedCorpus(const std::string& dir, uint64_t seed, int nFiles) {
    static const char* words[] = {
        "und", "der", "die", "das", "mit", "für", "über", "Straße", "Grüße", "Mädchen", "Bäcker", "Schlüssel", "Größe", "Übung",
        "Prüfung", "Lösung", "Rechnung", "Zeitung", "Wohnung", "Bewegung", "Führung", "Öffnung", "Erfahrung", "Beziehung",
        "Haus", "Baum", "Fluss", "Brücke", "Fähre", "Hafen", "Stadt", "Dorf", "Feld", "Wald", "Wiese", "Garten", "Fenster",
        "the", "and", "with", "from", "File", "Find", "Filter", "first", "fixed", "final", "value", "result", "search",
        "replace", "column", "line", "text", "data", "cache", "Cache", "CACHE", "buffer", "window", "error", "warning",
        "naïve", "café", "résumé", "Ærø", "Øresund", "Åland", "señor", "niño", "Ελλάδα", "Москва", "привет", "日本語", "中文",
        "1024", "2048", "3.14", "0x1F", "(", ")", ",", ".", ";", ":", "-", "!", "?" };
    constexpr int nWords = sizeof(words) / sizeof(words[0]);
    std::mt19937_64 rng(seed);
    // Log-normal sizes via Box-Muller on raw mt19937_64 output, so the corpus does not depend on
    // a standard library's distribution algorithm
    auto uniform01 = [&rng]() { return (static_cast<double>(rng() >> 11) + 0.5) / 9007199254740992.0; };
    auto logNormal = [&](double mu, double sigma) {
        const double u1 = uniform01(), u2 = uniform01();
        return std::exp(mu + sigma * std::sqrt(-2.0 * std::log(u1)) * std::cos(6.283185307179586 * u2));
    };
    std::filesystem::create_directories(dir);
    for (int i = 0; i < nFiles; ++i) {
        const int r = static_cast<int>(rng() % 100);
        const char* kind = r < 58 ? "utf8" : r < 66 ? "bom" : r < 81 ? "ansi" : r < 86 ? "utf16" : r < 94 ? "ascii" : "bin";
        const size_t target = static_cast<size_t>(std::clamp(logNormal(std::log(12000.0), 1.2), 200.0, 4.0e6));
        const bool crlf = (rng() & 1) != 0;
        const bool rare = (rng() % 100) == 0;
        std::string text;
        if (std::strcmp(kind, "bin") == 0) {
            text.resize(target);
            for (auto& ch : text) { const unsigned v = static_cast<unsigned>(rng() % 160); ch = static_cast<char>(v < 32 ? v : v + 96); }
        }
        else {
            int lineWords = 0;
            while (text.size() < target) {
                const char* w = words[rng() % nWords];
                if (std::strcmp(kind, "ascii") == 0) { bool ascii = true; for (const char* q = w; *q; ++q) if (static_cast<unsigned char>(*q) >= 0x80) ascii = false; if (!ascii) w = "plain"; }
                if (lineWords) text += ' ';
                text += w;
                if (++lineWords >= 8 + static_cast<int>(rng() % 9)) { text += crlf ? "\r\n" : "\n"; lineWords = 0; }
            }
            if (rare) { text += crlf ? "\r\nZebrafinkenkäfig am Ende\r\n" : "\nZebrafinkenkäfig am Ende\n"; }
        }
        std::string bytes;
        if (std::strcmp(kind, "bom") == 0) bytes = "\xEF\xBB\xBF" + text;
        else if (std::strcmp(kind, "ansi") == 0) {
            const std::u16string w = w32::widen(text);
            bytes.resize(w.size());
            w32::wideToMultiByte(w32::ACP, w.data(), static_cast<int>(w.size()), bytes.data());
        }
        else if (std::strcmp(kind, "utf16") == 0) {
            const std::u16string w = w32::widen(text);
            bytes = "\xFF\xFE";
            for (char16_t u : w) { bytes += static_cast<char>(u & 0xFF); bytes += static_cast<char>(u >> 8); }
        }
        else bytes = text;
        char sub[32]; std::snprintf(sub, sizeof sub, "/d%02d", i % 20);
        std::filesystem::create_directories(dir + sub);
        char name[64]; std::snprintf(name, sizeof name, "/f%05d_%s.txt", i, kind);
        std::ofstream(dir + sub + name, std::ios::binary).write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    }
}

// ---------------------------------------------------------------------------------------------
// Results shared by all tools
// ---------------------------------------------------------------------------------------------
enum Phase { PH_OPEN, PH_DETECT, PH_FILL, PH_SEARCH, PH_RESULT, PH_FINAL, PH_BUFFER, PH_COUNT };
static const char* kPhaseNames[PH_COUNT] = { "read", "detect", "fill", "search", "results", "final", "buffer" };

struct RunResult {
    double total = 0;
    double phase[PH_COUNT] = {};
    size_t hits = 0, filesWithHits = 0, binaryHits = 0;
    uint64_t textBytesLoaded = 0;
};

struct PhaseTimer {
    double& slot; Clock::time_point t0;
    explicit PhaseTimer(double& s) : slot(s), t0(Clock::now()) {}
    ~PhaseTimer() { slot += msBetween(t0, Clock::now()); }
};

// ---------------------------------------------------------------------------------------------
// MultiReplace a43e162
// ---------------------------------------------------------------------------------------------
namespace mr {

enum class Kind { UTF8, UTF16LE, UTF16BE, ANSI };
struct EncodingInfo { Kind kind = Kind::ANSI; unsigned codepage = 0; bool withBOM = false; int bomBytes = 0; };
enum class Skip { None, Binary, TooLarge, Unreadable, Undecodable };

bool isValidUtf8(const char* data, size_t len) {
    const unsigned char* s = reinterpret_cast<const unsigned char*>(data);
    size_t i = 0;
    while (i < len) {
        const unsigned char c = s[i];
        if (c < 0x80) { ++i; continue; }
        size_t need = 0;
        if ((c & 0xE0) == 0xC0) { need = 1; if ((c & 0xFE) == 0xC0) return false; }
        else if ((c & 0xF0) == 0xE0) { need = 2; }
        else if ((c & 0xF8) == 0xF0) { need = 3; if (c > 0xF4) return false; }
        else return false;
        if (i + need >= len) return false;
        for (size_t k = 1; k <= need; ++k) if ((s[i + k] & 0xC0) != 0x80) return false;
        if (need == 2) { const unsigned char c1 = s[i + 1]; if ((c == 0xE0 && (c1 & 0xE0) == 0x80) || (c == 0xED && (c1 & 0xE0) == 0xA0)) return false; }
        else if (need == 3) { const unsigned char c1 = s[i + 1]; if ((c == 0xF0 && (c1 & 0xF0) == 0x80) || (c == 0xF4 && (c1 & 0xF0) != 0x80)) return false; }
        i += need + 1;
    }
    return true;
}

static bool isMostlyAscii(const char* data, int len, double threshold) {
    if (len <= 0) return true;
    int ascii = 0;
    const unsigned char* p = reinterpret_cast<const unsigned char*>(data);
    for (int i = 0; i < len; ++i) if (p[i] < 0x80) ++ascii;
    return static_cast<double>(ascii) / len >= threshold;
}

static int dbcsPairScore(const char* data, int len, unsigned cp) {
    int score = 0;
    for (int i = 0; i < len; ++i) if (w32::isDbcsLeadByte(cp, static_cast<unsigned char>(data[i]))) { if (i + 1 < len) { ++score; ++i; } }
    return score;
}

static unsigned autoDetectAnsiCodepage(const char* data, int len, unsigned acp) {
    if (isMostlyAscii(data, len, 0.98)) return acp;
    const int cap = 128 * 1024;
    const char* s1 = data; const int n1 = len > cap ? cap : len;
    const char* s2 = nullptr; int n2 = 0;
    if (len > cap * 3) { s2 = data + len / 2; n2 = std::min(cap, len - len / 2); }
    static const unsigned cjk[] = { 932u, 936u, 949u, 950u };
    for (unsigned cp : cjk) {
        const int score = dbcsPairScore(s1, n1, cp) + ((s2 && n2 > 0) ? dbcsPairScore(s2, n2, cp) : 0);
        if (score >= 3 && w32::dbcsValid(cp, s1, n1) && (!s2 || w32::dbcsValid(cp, s2, n2))) return cp;
    }
    return acp;
}

// Encoding::isValidUtf8 with the eight-byte ASCII step
bool isValidUtf8Fast(const char* data, size_t len) {
    const unsigned char* s = reinterpret_cast<const unsigned char*>(data);
    size_t i = 0;
    while (i < len) {
        if (len - i >= 8) {
            uint64_t block;
            std::memcpy(&block, s + i, 8);
            if ((block & 0x8080808080808080ull) == 0) { i += 8; continue; }
        }
        const unsigned char c = s[i];
        if (c < 0x80) { ++i; continue; }
        size_t need = 0;
        if ((c & 0xE0) == 0xC0) { need = 1; if ((c & 0xFE) == 0xC0) return false; }
        else if ((c & 0xF0) == 0xE0) { need = 2; }
        else if ((c & 0xF8) == 0xF0) { need = 3; if (c > 0xF4) return false; }
        else return false;
        if (i + need >= len) return false;
        for (size_t k = 1; k <= need; ++k) if ((s[i + k] & 0xC0) != 0x80) return false;
        if (need == 2) { const unsigned char c1 = s[i + 1]; if ((c == 0xE0 && (c1 & 0xE0) == 0x80) || (c == 0xED && (c1 & 0xE0) == 0xA0)) return false; }
        else if (need == 3) { const unsigned char c1 = s[i + 1]; if ((c == 0xF0 && (c1 & 0xF0) == 0x80) || (c == 0xF4 && (c1 & 0xF0) != 0x80)) return false; }
        i += need + 1;
    }
    return true;
}

static EncodingInfo detectEncoding(const char* data, size_t len, bool fastUtf8) {
    EncodingInfo ei;
    if (!data || len == 0) { ei.kind = Kind::ANSI; ei.codepage = w32::ACP; return ei; }
    const unsigned char* p = reinterpret_cast<const unsigned char*>(data);
    if (len >= 3 && p[0] == 0xEF && p[1] == 0xBB && p[2] == 0xBF) { ei.kind = Kind::UTF8; ei.withBOM = true; ei.bomBytes = 3; return ei; }
    if (len >= 2 && p[0] == 0xFF && p[1] == 0xFE) { ei.kind = Kind::UTF16LE; ei.withBOM = true; ei.bomBytes = 2; return ei; }
    if (len >= 2 && p[0] == 0xFE && p[1] == 0xFF) { ei.kind = Kind::UTF16BE; ei.withBOM = true; ei.bomBytes = 2; return ei; }
    if (fastUtf8 ? isValidUtf8Fast(data, len) : isValidUtf8(data, len)) { ei.kind = Kind::UTF8; return ei; }
    ei.kind = Kind::ANSI; ei.codepage = w32::ACP;
    ei.codepage = autoDetectAnsiCodepage(data, static_cast<int>(len), ei.codepage);
    return ei;
}

static bool convertBufferToUtf8(const char* data, size_t len, const EncodingInfo& src, std::string& out) {
    out.clear();
    if (src.bomBytes > 0 && static_cast<size_t>(src.bomBytes) <= len) { data += src.bomBytes; len -= src.bomBytes; }
    if (src.kind == Kind::UTF16LE || src.kind == Kind::UTF16BE) {
        const size_t wc = len / 2;
        std::u16string w(wc, u'\0');
        const unsigned char* b = reinterpret_cast<const unsigned char*>(data);
        if (src.kind == Kind::UTF16LE) std::memcpy(w.data(), data, wc * 2);
        else for (size_t i = 0; i < wc; ++i) w[i] = static_cast<char16_t>((b[i * 2] << 8) | b[i * 2 + 1]);
        const int n = w32::wideToMultiByte(sc::CP_UTF8, w.data(), static_cast<int>(wc), nullptr);
        out.resize(static_cast<size_t>(n));
        w32::wideToMultiByte(sc::CP_UTF8, w.data(), static_cast<int>(wc), out.data());
        return true;
    }
    const int wn = w32::multiByteToWide(src.codepage, data, static_cast<int>(len), nullptr);
    std::u16string w(static_cast<size_t>(wn), u'\0');
    w32::multiByteToWide(src.codepage, data, static_cast<int>(len), w.data());
    const int n = w32::wideToMultiByte(sc::CP_UTF8, w.data(), wn, nullptr);
    out.resize(static_cast<size_t>(n));
    w32::wideToMultiByte(sc::CP_UTF8, w.data(), wn, out.data());
    return true;
}

// a43e162: one Hit per match
struct Hit {
    std::string fullPathUtf8;
    Sci::Position pos{}, length{};
    int docLine{ -1 }, searchFlags{ 0 };
    std::u16string findTextW;
    std::vector<std::u16string> allFindTexts;
    std::vector<Sci::Position> allPositions, allLengths;
    std::vector<int> allSearchFlags;
    int displayLineStart{ -1 }, numberStart{ 0 }, numberLen{ 0 };
    std::vector<int> matchStarts, matchLens, matchColors;
    int colorIndex{ -1 };
};
struct CritAgg { std::u16string text; std::vector<Hit> hits; };
struct FileAgg { std::u16string wPath; int hitCount = 0; std::vector<CritAgg> crits; };
using FileMap = std::unordered_map<std::string, FileAgg>;

// MR+dock: the search's text, flags and color once per criterion, one RowMatch list per dock line
struct RawHit { Sci::Position pos{}, length{}; int docLine{ -1 }; };
struct RowMatch {
    std::shared_ptr<const std::u16string> findTextW;
    Sci::Position pos{}, length{};
    int searchFlags{ 0 }, colorIndex{ -1 }, dispStart{ -1 }, dispLen{ 0 };
};
struct Row {
    std::string fullPathUtf8;
    Sci::Position pos{}, length{};
    int docLine{ -1 }, searchFlags{ 0 };
    std::u16string findTextW;
    std::vector<RowMatch> matches;
    int displayLineStart{ -1 }, numberStart{ 0 }, numberLen{ 0 };
};
struct LeanCrit {
    std::u16string text;
    std::shared_ptr<const std::u16string> findTextW;
    int searchFlags = 0, colorIndex = -1;
    std::vector<RawHit> hits;
};
struct LeanFile { std::u16string wPath; int hitCount = 0; std::vector<LeanCrit> crits; };
using LeanMap = std::unordered_map<std::string, LeanFile>;

struct Variant { bool utf8 = false, prefilter = false, dock = false; };

class Tool {
public:
    explicit Tool(Variant v) : _v(v) {}

    RunResult run(const Corpus& corpus, const Query& q) {
        RunResult rr;
        const auto t0 = Clock::now();
        SciDoc scan(false);
        scan.setCodePage(sc::CP_UTF8);
        _pendingText.clear(); _pendingHits.clear(); _pendingRows.clear();
        std::unordered_set<std::string> uniqueFiles;
        const std::u16string patW = w32::widen(q.text);
        const std::u16string label = patW;
        const auto patShared = std::make_shared<const std::u16string>(patW);
        int flags = (q.matchCase ? sc::MATCHCASE : 0);
        if (q.regex) flags |= sc::REGEXP | sc::POSIX | sc::SKIPCRLFASONE | sc::EMPTY_ALL;
        std::string patBytes; int patCp = -1;
        startSearchBlock(u"Search \"" + label + u"\" (0 hits in 0 files)");
        std::string content;
        for (const FileEntry& fe : corpus.files) {
            EncodingInfo enc;
            const Skip s = loadTextFile(fe.path, content, enc, rr);
            if (s != Skip::None) continue;
            rr.textBytesLoaded += content.size();
            if (_v.prefilter) {
                PhaseTimer pt(rr.phase[PH_SEARCH]);
                if (patCp != sc::CP_UTF8) { patBytes = w32::narrow(patW); patCp = sc::CP_UTF8; }
                LiteralPrefilter prefilter(content);
                if (!prefilter.mayMatch(patBytes, flags)) continue;
            }
            {
                PhaseTimer pt(rr.phase[PH_FILL]);
                scan.clearAll();
                scan.setCodePage(sc::CP_UTF8);
                scan.append(content.data(), static_cast<Sci::Position>(content.size()));
            }
            const std::u16string wPath = fe.path16;
            const std::string u8Path = w32::narrow(wPath);
            const Sci::Position docLength = scan.length();
            FileMap fileMap;
            LeanMap leanMap;
            int hitsInFile = 0;
            {
                PhaseTimer pt(rr.phase[PH_SEARCH]);
                if (patCp != scan.cp) { patBytes = w32::narrow(patW); patCp = scan.cp; }
                std::vector<Hit> raw;
                std::vector<RawHit> lean;
                Sci::Position pos = 0;
                while (true) {
                    Sci::Position end = 0;
                    const Sci::Position p = scan.searchInTarget(pos, docLength, patBytes.c_str(), patBytes.size(), flags, end);
                    if (p < 0 || end < p || end > docLength) break;
                    const Sci::Position len = end - p;
                    if (len > 0) pos = p + len;
                    else { const Sci::Position after = scan.positionAfter(p); pos = std::min(after > p ? after : p + 1, scan.length()); }
                    if (_v.dock) {
                        RawHit h{ p, len, static_cast<int>(scan.lineFromPos(p)) };
                        if (flags & sc::REGEXP) trimToFirstLine(scan, h.pos, h.length, h.docLine);
                        lean.push_back(h);
                        continue;
                    }
                    Hit h{};
                    h.fullPathUtf8 = u8Path;
                    h.pos = p; h.length = len;
                    h.docLine = static_cast<int>(scan.lineFromPos(p));
                    h.searchFlags = flags;
                    if (flags & sc::REGEXP) trimToFirstLine(scan, h.pos, h.length, h.docLine);
                    h.findTextW = patW;
                    h.colorIndex = 0;
                    raw.push_back(std::move(h));
                }
                const int n = static_cast<int>(_v.dock ? lean.size() : raw.size());
                if (n > 0 && _v.dock) {
                    auto& agg = leanMap[u8Path];
                    agg.wPath = wPath;
                    agg.hitCount += n;
                    agg.crits.push_back({ label, patShared, flags, 0, std::move(lean) });
                }
                else if (n > 0) {
                    auto& agg = fileMap[u8Path];
                    agg.wPath = wPath;
                    agg.hitCount += n;
                    agg.crits.push_back({ label, std::move(raw) });
                }
                hitsInFile += n;
            }
            if (hitsInFile > 0) {
                PhaseTimer pt(rr.phase[PH_RESULT]);
                if (_v.dock) appendFileBlockLean(leanMap, scan);
                else appendFileBlock(fileMap, scan);
                uniqueFiles.insert(u8Path);
                rr.hits += static_cast<size_t>(hitsInFile); ++rr.filesWithHits;
            }
        }
        {
            PhaseTimer pt(rr.phase[PH_FINAL]);
            closeSearchBlock(static_cast<int>(rr.hits), static_cast<int>(uniqueFiles.size()));
        }
        rr.total = msBetween(t0, Clock::now());
        return rr;   // the dock (document and hits) lives on with the Tool, as in the plugin
    }

private:
    Variant _v;
    std::string _pendingText;
    std::vector<Hit> _pendingHits;
    std::vector<Row> _pendingRows;
    std::unique_ptr<SciDoc> _dock;
    std::vector<Hit> _dockHits;
    std::vector<Row> _dockRows;
    std::unordered_map<int, int> _index;

    Skip loadTextFile(const std::string& path, std::string& content, EncodingInfo& enc, RunResult& rr) {
        content.clear(); enc = EncodingInfo{};
        auto t = Clock::now();
        const int fd = ::open(path.c_str(), O_RDONLY);
        if (fd < 0) { rr.phase[PH_OPEN] += msBetween(t, Clock::now()); return Skip::Unreadable; }
        struct stat st {};
        ::fstat(fd, &st);
        const uint64_t fileSize = static_cast<uint64_t>(st.st_size);
        const size_t headerSize = static_cast<size_t>(std::min<uint64_t>(fileSize, 8192));
        content.resize(headerSize);
        size_t headerLen = readFully(fd, content.data(), headerSize);
        content.resize(headerLen);
        rr.phase[PH_OPEN] += msBetween(t, Clock::now());

        t = Clock::now();
        const bool binary = shouldSkipAsBinary(content.data(), content.size());
        rr.phase[PH_DETECT] += msBetween(t, Clock::now());
        if (binary) { ::close(fd); content.clear(); return Skip::Binary; }

        t = Clock::now();
        if (headerLen == headerSize && fileSize > headerSize) {
            content.resize(static_cast<size_t>(fileSize));
            const size_t rest = readFully(fd, content.data() + headerLen, content.size() - headerLen);
            content.resize(headerLen + rest);
        }
        ::close(fd);
        rr.phase[PH_OPEN] += msBetween(t, Clock::now());

        PhaseTimer pt(rr.phase[PH_DETECT]);
        enc = detectEncoding(content.data(), content.size(), _v.utf8);
        if (enc.kind == Kind::UTF8) { content.erase(0, static_cast<size_t>(enc.bomBytes)); return Skip::None; }
        std::string u8;
        if (!convertBufferToUtf8(content.data(), content.size(), enc, u8)) { content.clear(); return Skip::Undecodable; }
        content.swap(u8);
        return Skip::None;
    }

    static size_t readFully(int fd, char* dst, size_t count) {
        size_t got = 0;
        while (got < count) { const ssize_t n = ::read(fd, dst + got, count - got); if (n <= 0) break; got += static_cast<size_t>(n); }
        return got;
    }

    static bool shouldSkipAsBinary(const char* data, size_t len) {
        const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
        if (len >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) return false;
        if (len >= 2 && ((u[0] == 0xFF && u[1] == 0xFE) || (u[0] == 0xFE && u[1] == 0xFF))) return false;
        return std::memchr(data, '\0', std::min<size_t>(len, 8192)) != nullptr;
    }

    static void trimToFirstLine(SciDoc& sci, Sci::Position pos, Sci::Position& length, int docLine) {
        const int line0 = docLine >= 0 ? docLine : static_cast<int>(sci.lineFromPos(pos));
        const Sci::Position ls = sci.lineStart(line0), le = sci.lineEnd(line0);
        if (pos >= ls && pos + length <= le) return;
        if (pos >= le) length = std::min<Sci::Position>(length, 1);
        else length = le - pos;
    }

    static std::string indent(int level) { static const int sp[] = { 1, 2, 3, 4 }; return std::string(static_cast<size_t>(sp[level]), ' '); }

    void startSearchBlock(const std::u16string& header) {
        _pendingText = indent(0) + w32::narrow(header) + "\r\n";
        _pendingHits.clear();
        _pendingRows.clear();
    }

    void appendFileBlock(FileMap& fm, SciDoc& sci) {
        std::string partText;
        std::vector<Hit> partHits;
        buildListText(fm, sci, partText, partHits);
        for (auto& h : partHits) h.displayLineStart += static_cast<int>(_pendingText.size());
        _pendingText += partText;
        _pendingHits.insert(_pendingHits.end(), std::make_move_iterator(partHits.begin()), std::make_move_iterator(partHits.end()));
    }

    void buildListText(FileMap& files, SciDoc& sci, std::string& outText, std::vector<Hit>& outHits) {
        std::string body;
        for (auto& [path, f] : files) {
            body.append(indent(1));
            body.append(w32::narrow(f.wPath + u" (" + w32::widen(std::to_string(f.hitCount)) + u" hits)"));
            body.append("\r\n", 2);
            std::vector<Hit> merged = std::move(f.crits.front().hits);
            formatHitsLines(sci, merged, body);
            outHits.insert(outHits.end(), std::make_move_iterator(merged.begin()), std::make_move_iterator(merged.end()));
        }
        outText = std::move(body);
    }

    // Display text of a document line as ResultDock::formatHitsLines builds it (UTF-8, no FlowTabs padding)
    struct LineView {
        int prevDocLine = -1;
        std::string raw, display;
        Sci::Position absLineStart = 0;
        std::vector<size_t> mapOrigToDisp;
        static constexpr size_t kMaxHitTextUtf8 = 2048;

        void load(SciDoc& sci, int line0) {
            const int rawLen = static_cast<int>(sci.lineLength(line0));
            raw.resize(static_cast<size_t>(rawLen) + 1);
            sci.getText(raw.data(), sci.lineStart(line0), sci.lineStart(line0) + rawLen);
            raw.resize(static_cast<size_t>(rawLen));
            while (!raw.empty() && (raw.back() == '\r' || raw.back() == '\n')) raw.pop_back();
            absLineStart = sci.lineStart(line0);
            auto isCtlCp = [](uint32_t cp) { return cp == 0 || (cp <= 0x1F && cp != 0x09 && cp != 0x0A && cp != 0x0D) || cp == 0x7F || (cp >= 0x80 && cp <= 0x9F) || cp == 0x2028 || cp == 0x2029; };
            display.clear(); mapOrigToDisp.clear();
            size_t o = 0, d = 0;
            while (o < raw.size() && d <= kMaxHitTextUtf8) {
                const unsigned char c = static_cast<unsigned char>(raw[o]);
                size_t clen;
                if ((c & 0x80) == 0) clen = 1;
                else if ((c & 0xE0) == 0xC0 && o + 1 < raw.size()) clen = 2;
                else if ((c & 0xF0) == 0xE0 && o + 2 < raw.size()) clen = 3;
                else if ((c & 0xF8) == 0xF0 && o + 3 < raw.size()) clen = 4;
                else clen = 1;
                uint32_t cp = 0;
                if (clen == 1) cp = c;
                else if (clen == 2) cp = ((c & 0x1Fu) << 6) | (uint32_t(raw[o + 1]) & 0x3Fu);
                else if (clen == 3) cp = ((c & 0x0Fu) << 12) | ((uint32_t(raw[o + 1]) & 0x3Fu) << 6) | (uint32_t(raw[o + 2]) & 0x3Fu);
                else cp = ((c & 0x07u) << 18) | ((uint32_t(raw[o + 1]) & 0x3Fu) << 12) | ((uint32_t(raw[o + 2]) & 0x3Fu) << 6) | (uint32_t(raw[o + 3]) & 0x3Fu);
                if (cp == 0x00AD) mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d);
                else if (cp == 0x00A0 || isCtlCp(cp)) { mapOrigToDisp.insert(mapOrigToDisp.end(), clen, d); display.push_back(' '); ++d; }
                else { for (size_t k = 0; k < clen; ++k) mapOrigToDisp.push_back(d + k); display.append(raw, o, clen); d += clen; }
                o += clen;
            }
            mapOrigToDisp.push_back(d);
            if (display.size() > kMaxHitTextUtf8) {
                size_t cut = kMaxHitTextUtf8 - 3;
                while (cut > 0 && (static_cast<unsigned char>(display[cut]) & 0xC0) == 0x80) --cut;
                display.resize(cut); display.append("...");
            }
        }
        size_t toDisp(size_t off) const { return off >= mapOrigToDisp.size() ? (mapOrigToDisp.empty() ? 0 : mapOrigToDisp.back()) : mapOrigToDisp[off]; }
    };

    static size_t countDigits(int n) { size_t c = 0; if (n <= 0) return 1; while (n > 0) { ++c; n /= 10; } return c; }

    // ResultDock::formatHitsLines of a43e162
    void formatHitsLines(SciDoc& sci, std::vector<Hit>& hits, std::string& out) {
        const std::string indentHit = indent(3);
        const std::string kLine = "Line ";
        std::vector<int> lineNumbers; lineNumbers.reserve(hits.size());
        size_t maxDigits = 0;
        for (const Hit& h : hits) {
            const int line1 = static_cast<int>(sci.lineFromPos(h.pos)) + 1;
            lineNumbers.push_back(line1);
            maxDigits = std::max(maxDigits, countDigits(line1));
        }
        LineView lv;
        Hit* firstHitOnRow = nullptr; size_t hitIdx = 0;
        const size_t prefixLen = indentHit.size() + kLine.size() + maxDigits + 2;
        for (Hit& h : hits) {
            const int line1 = lineNumbers[hitIdx++];
            const int line0 = line1 - 1;
            if (line0 != lv.prevDocLine) lv.load(sci, line0);
            const size_t hitStart = static_cast<size_t>(h.pos - lv.absLineStart);
            const size_t dispStart = lv.toDisp(hitStart);
            const size_t dispEnd = lv.toDisp(hitStart + static_cast<size_t>(h.length));
            const size_t dispLen = dispEnd > dispStart ? dispEnd - dispStart : 0;
            if (line0 != lv.prevDocLine) {
                const size_t line1Digits = countDigits(line1);
                const size_t rowStartPos = out.size();
                out.append(indentHit); out.append(kLine); out.append(maxDigits - line1Digits, ' '); out.append(std::to_string(line1)); out.append(": ", 2);
                out.append(lv.display); out.append("\r\n", 2);
                size_t lineHitCount = 1;
                for (size_t j = hitIdx; j < lineNumbers.size() && lineNumbers[j] == line1; ++j) ++lineHitCount;
                h.displayLineStart = static_cast<int>(rowStartPos);
                h.numberStart = static_cast<int>(indentHit.size() + kLine.size() + maxDigits - line1Digits);
                h.numberLen = static_cast<int>(line1Digits);
                h.matchStarts.clear(); h.matchLens.clear(); h.matchColors.clear();
                h.matchStarts.reserve(lineHitCount); h.matchLens.reserve(lineHitCount); h.matchColors.reserve(lineHitCount);
                firstHitOnRow = &h;
                firstHitOnRow->allFindTexts.clear(); firstHitOnRow->allPositions.clear(); firstHitOnRow->allLengths.clear(); firstHitOnRow->allSearchFlags.clear();
                firstHitOnRow->allFindTexts.reserve(lineHitCount); firstHitOnRow->allPositions.reserve(lineHitCount);
                firstHitOnRow->allLengths.reserve(lineHitCount); firstHitOnRow->allSearchFlags.reserve(lineHitCount);
                firstHitOnRow->allFindTexts.push_back(h.findTextW); firstHitOnRow->allPositions.push_back(h.pos);
                firstHitOnRow->allLengths.push_back(h.length); firstHitOnRow->allSearchFlags.push_back(h.searchFlags);
                if (dispStart < lv.display.size()) {
                    const size_t safeLen = std::min(dispLen, lv.display.size() - dispStart);
                    if (safeLen > 0) { h.matchStarts.push_back(static_cast<int>(prefixLen + dispStart)); h.matchLens.push_back(static_cast<int>(safeLen)); h.matchColors.push_back(h.colorIndex); }
                }
                lv.prevDocLine = line0;
            }
            else if (firstHitOnRow) {
                if (dispStart < lv.display.size()) {
                    const size_t safeLen = std::min(dispLen, lv.display.size() - dispStart);
                    if (safeLen > 0) { firstHitOnRow->matchStarts.push_back(static_cast<int>(prefixLen + dispStart)); firstHitOnRow->matchLens.push_back(static_cast<int>(safeLen)); firstHitOnRow->matchColors.push_back(h.colorIndex); }
                }
                firstHitOnRow->allFindTexts.push_back(h.findTextW); firstHitOnRow->allPositions.push_back(h.pos);
                firstHitOnRow->allLengths.push_back(h.length); firstHitOnRow->allSearchFlags.push_back(h.searchFlags);
                h.displayLineStart = -1;
            }
        }
        hits.erase(std::remove_if(hits.begin(), hits.end(), [](const Hit& e) { return e.displayLineStart < 0; }), hits.end());
    }

    // MR+dock: ResultDock::buildListText and formatHitsLines with RawHit, CritAgg and RowMatch
    void appendFileBlockLean(LeanMap& fm, SciDoc& sci) {
        std::string body;
        std::vector<Row> rows;
        for (auto& [path, f] : fm) {
            body.append(indent(1));
            body.append(w32::narrow(f.wPath + u" (" + w32::widen(std::to_string(f.hitCount)) + u" hits)"));
            body.append("\r\n", 2);
            const LeanCrit& c = f.crits.front();
            formatRows(sci, path, c, body, rows);
        }
        for (auto& r : rows) r.displayLineStart += static_cast<int>(_pendingText.size());
        _pendingText += body;
        _pendingRows.insert(_pendingRows.end(), std::make_move_iterator(rows.begin()), std::make_move_iterator(rows.end()));
    }

    void formatRows(SciDoc& sci, const std::string& path, const LeanCrit& crit, std::string& out, std::vector<Row>& rows) {
        const std::string indentHit = indent(3);
        const std::string kLine = "Line ";
        std::vector<int> lineNumbers; lineNumbers.reserve(crit.hits.size());
        size_t maxDigits = 0;
        for (const RawHit& h : crit.hits) {
            const int line1 = (h.docLine >= 0 ? h.docLine : static_cast<int>(sci.lineFromPos(h.pos))) + 1;
            lineNumbers.push_back(line1);
            maxDigits = std::max(maxDigits, countDigits(line1));
        }
        LineView lv;
        size_t hitIdx = 0;
        const size_t prefixLen = indentHit.size() + kLine.size() + maxDigits + 2;
        for (const RawHit& h : crit.hits) {
            const int line1 = lineNumbers[hitIdx++];
            const int line0 = line1 - 1;
            if (line0 != lv.prevDocLine) lv.load(sci, line0);
            const size_t hitStart = static_cast<size_t>(h.pos - lv.absLineStart);
            const size_t dispStart = lv.toDisp(hitStart);
            const size_t dispEnd = lv.toDisp(hitStart + static_cast<size_t>(h.length));
            const size_t dispLen = dispEnd > dispStart ? dispEnd - dispStart : 0;
            if (line0 != lv.prevDocLine) {
                const size_t line1Digits = countDigits(line1);
                const size_t rowStartPos = out.size();
                out.append(indentHit); out.append(kLine); out.append(maxDigits - line1Digits, ' '); out.append(std::to_string(line1)); out.append(": ", 2);
                out.append(lv.display); out.append("\r\n", 2);
                size_t lineHitCount = 1;
                for (size_t j = hitIdx; j < lineNumbers.size() && lineNumbers[j] == line1; ++j) ++lineHitCount;
                Row& row = rows.emplace_back();
                row.fullPathUtf8 = path;
                row.pos = h.pos; row.length = h.length; row.docLine = h.docLine;
                row.searchFlags = crit.searchFlags;
                row.findTextW = *crit.findTextW;
                row.displayLineStart = static_cast<int>(rowStartPos);
                row.numberStart = static_cast<int>(indentHit.size() + kLine.size() + maxDigits - line1Digits);
                row.numberLen = static_cast<int>(line1Digits);
                row.matches.reserve(lineHitCount);
                lv.prevDocLine = line0;
            }
            RowMatch& m = rows.back().matches.emplace_back();
            m.findTextW = crit.findTextW;
            m.pos = h.pos; m.length = h.length;
            m.searchFlags = crit.searchFlags;
            m.colorIndex = crit.colorIndex;
            if (dispStart < lv.display.size()) {
                const size_t safeLen = std::min(dispLen, lv.display.size() - dispStart);
                if (safeLen > 0) { m.dispStart = static_cast<int>(prefixLen + dispStart); m.dispLen = static_cast<int>(safeLen); }
            }
        }
    }

    // ResultDock::closeSearchBlock + prependBlock into an empty dock
    void closeSearchBlock(int totalHits, int totalFiles) {
        const std::string hitsStr = std::to_string(totalHits), filesStr = std::to_string(totalFiles);
        const size_t headerEnd = _pendingText.find("\r\n");
        size_t pos = headerEnd == std::string::npos ? _pendingText.size() : headerEnd;
        ptrdiff_t delta = 0;
        auto replaceLastNumber = [&](const std::string& with) {
            while (pos > 0 && !std::isdigit(static_cast<unsigned char>(_pendingText[pos - 1]))) --pos;
            const size_t e = pos; while (pos > 0 && std::isdigit(static_cast<unsigned char>(_pendingText[pos - 1]))) --pos;
            _pendingText.replace(pos, e - pos, with);
            delta += static_cast<ptrdiff_t>(with.size()) - static_cast<ptrdiff_t>(e - pos);
        };
        replaceLastNumber(filesStr);
        replaceLastNumber(hitsStr);
        if (delta != 0) {
            for (auto& h : _pendingHits) h.displayLineStart += static_cast<int>(delta);
            for (auto& r : _pendingRows) r.displayLineStart += static_cast<int>(delta);
        }

        _dock = std::make_unique<SciDoc>(true);
        SciDoc& dock = *_dock;
        dock.allocate(static_cast<Sci::Position>(_pendingText.size()) + 65536);
        dock.insert(0, _pendingText.data(), static_cast<Sci::Position>(_pendingText.size()));
        Sci::Line line = 0;
        for (size_t p = 0; p < _pendingText.size();) {
            int ind = 0; while (p < _pendingText.size() && _pendingText[p] == ' ') { ++ind; ++p; }
            while (p < _pendingText.size() && _pendingText[p] != '\n') ++p;
            ++p;
            dock.d->SetLevel(line++, 0x400 + ind);
        }
        if (_v.dock) {
            // one call: line styles by indent, line numbers and matches as styles
            std::string styles(_pendingText.size(), static_cast<char>(32));
            for (size_t ls = 0; ls < _pendingText.size();) {
                size_t next = _pendingText.find('\n', ls);
                next = (next == std::string::npos) ? _pendingText.size() : next + 1;
                int ind = 0; for (size_t i = ls; i < _pendingText.size() && _pendingText[i] == ' '; ++i) ++ind;
                if (ind >= 1 && ind <= 3) std::fill(styles.begin() + static_cast<ptrdiff_t>(ls), styles.begin() + static_cast<ptrdiff_t>(next), static_cast<char>(32 + ind));
                ls = next;
            }
            for (const Row& r : _pendingRows) {
                std::fill_n(styles.begin() + r.displayLineStart + r.numberStart, r.numberLen, static_cast<char>(40));
                for (const RowMatch& m : r.matches)
                    if (m.dispStart >= 0) std::fill_n(styles.begin() + r.displayLineStart + m.dispStart, m.dispLen, static_cast<char>(41));
            }
            dock.d->StartStyling(0);
            dock.d->SetStyles(static_cast<Sci::Position>(styles.size()), styles.data());
            _dockRows.insert(_dockRows.begin(), std::make_move_iterator(_pendingRows.begin()), std::make_move_iterator(_pendingRows.end()));
            _index.reserve(_dockRows.size());
            for (int i = 0; i < static_cast<int>(_dockRows.size()); ++i) _index[_dockRows[i].displayLineStart] = i;
            _pendingRows.clear();
            return;
        }
        const Sci::Line lines = dock.lineCount();
        dock.d->StartStyling(0);
        for (Sci::Line l = 0; l < lines; ++l) {
            const Sci::Position ll = dock.lineLength(l);
            const int ind = dock.d->GetLineIndentation(l);
            if (ll > 0) dock.d->SetStyleFor(ll, static_cast<char>(ind >= 4 ? 32 : 32 + ind));
        }
        dock.d->DecorationSetCurrentIndicator(29);
        for (const auto& h : _pendingHits) if (h.displayLineStart >= 0) dock.d->DecorationFillRange(h.displayLineStart + h.numberStart, 1, h.numberLen);
        dock.d->DecorationSetCurrentIndicator(30);
        for (const auto& h : _pendingHits) {
            if (h.displayLineStart < 0) continue;
            for (size_t i = 0; i < h.matchStarts.size(); ++i)
                dock.d->DecorationFillRange(h.displayLineStart + h.matchStarts[i], 1, h.matchLens[i]);
        }
        _dockHits.insert(_dockHits.begin(), std::make_move_iterator(_pendingHits.begin()), std::make_move_iterator(_pendingHits.end()));
        _index.reserve(_dockHits.size());
        for (int i = 0; i < static_cast<int>(_dockHits.size()); ++i) if (_dockHits[i].displayLineStart >= 0) _index[_dockHits[i].displayLineStart] = i;
        _pendingHits.clear();
    }
};
}

// ---------------------------------------------------------------------------------------------
// Notepad++ (shared by the new and the old loading path)
// ---------------------------------------------------------------------------------------------
namespace npp {

constexpr size_t blockSize = 128 * 1024 + 4;
constexpr size_t kLineBufMax = 2048;            // SC_SEARCHRESULT_LINEBUFFERMAXLENGTH
constexpr size_t kFindTempMax = 1024 * 1024;    // FINDTEMPSTRING_MAXSIZE

enum UniMode { uni8Bit = 0, uniUTF8 = 1, uni16BE = 2, uni16LE = 3, uniCookie = 4, uni7Bit = 5, uni16BE_NoBOM = 6, uni16LE_NoBOM = 7, uniUTF8_NoBOM = 8 };
enum class Eol { windows, macos, unix_, unknown };
enum LangType { L_TEXT = 0, L_PHP, L_HTML, L_XML, L_BASH, L_PYTHON, L_PERL, L_RUBY, L_JAVASCRIPT };

// WcharMbcsConvertor: two calls per conversion as the Win32 API is used
class Wmc {
    std::unique_ptr<w32::wchar[]> _w; size_t _wcap = 0, _wlen = 0;
    std::unique_ptr<char[]> _m; size_t _mcap = 0;
    void wsize(size_t n) { if (n > _wcap) { _wcap = std::max<size_t>(n, 1024); _w.reset(new w32::wchar[_wcap]); } }
    void msize(size_t n) { if (n > _mcap) { _mcap = std::max<size_t>(n, 1024); _m.reset(new char[_mcap]); } }
public:
    size_t sizeW() const { return _wlen; }
    const w32::wchar* char2wchar(const char* s, unsigned cp, int len, int* lenWc = nullptr) {
        if (len < 0) len = static_cast<int>(std::strlen(s));
        int n = w32::multiByteToWide(cp, s, len, nullptr);
        wsize(static_cast<size_t>(n) + 1);
        n = w32::multiByteToWide(cp, s, len, _w.get());
        _w[n] = 0; _wlen = static_cast<size_t>(n);
        if (lenWc) *lenWc = n;
        return _w.get();
    }
    const w32::wchar* char2wcharMarks(const char* s, unsigned cp, intptr_t* ms, intptr_t* me, int mbcsLen) {
        if (mbcsLen == 0) { _wlen = 0; *ms = 0; *me = 0; wsize(1); _w[0] = 0; return _w.get(); }
        int len = w32::multiByteToWide(cp, s, mbcsLen, nullptr);
        if (len > 0) {
            wsize(static_cast<size_t>(len) + 1);
            len = w32::multiByteToWide(cp, s, mbcsLen, _w.get());
            _w[len] = 0; _wlen = static_cast<size_t>(len);
            if (*ms < mbcsLen && *me <= mbcsLen) {
                *ms = w32::multiByteToWide(cp, s, static_cast<int>(*ms), nullptr);
                *me = w32::multiByteToWide(cp, s, static_cast<int>(*me), nullptr);
                if (*ms >= len || *me > len) { *ms = 0; *me = 0; }
            }
        }
        else { _wlen = 0; *ms = 0; *me = 0; }
        return _w.get();
    }
    const char* wchar2char(const w32::wchar* s, unsigned cp, int lenWc = -1, int* lenOut = nullptr) {
        if (lenWc < 0) lenWc = static_cast<int>(w32::wlen(s));
        int n = w32::wideToMultiByte(cp, s, lenWc, nullptr);
        msize(static_cast<size_t>(n) + 1);
        n = w32::wideToMultiByte(cp, s, lenWc, _m.get());
        _m[n] = 0;
        if (lenOut) *lenOut = n;
        return _m.get();
    }
    const char* wchar2charMarks(const w32::wchar* s, unsigned cp, intptr_t* ms, intptr_t* me, int wcLen, int* lenOut) {
        int len = w32::wideToMultiByte(cp, s, wcLen, nullptr);
        if (len > 0) {
            msize(static_cast<size_t>(len) + 1);
            len = w32::wideToMultiByte(cp, s, wcLen, _m.get());
            _m[len] = 0;
            if (*ms < wcLen && *me < wcLen) {
                *ms = w32::wideToMultiByte(cp, s, static_cast<int>(*ms), nullptr);
                *me = w32::wideToMultiByte(cp, s, static_cast<int>(*me), nullptr);
                if (*ms >= len || *me > len) { *ms = 0; *me = 0; }
            }
        }
        else { msize(1); _m[0] = 0; len = 0; *ms = 0; *me = 0; }
        if (lenOut) *lenOut = len;
        return _m.get();
    }
    const char* encode(unsigned fromCp, unsigned toCp, const char* s, int len, int* outLen) {
        int lw = 0;
        const w32::wchar* w = char2wchar(s, fromCp, len, &lw);
        return wchar2char(w, toCp, lw, outLen);
    }
};

class Utf8_16_Read {
public:
    ~Utf8_16_Read() { if (is16()) delete[] _newBuf; }
    UniMode encoding() const { return _enc; }
    const char* newBuf() const { return reinterpret_cast<const char*>(_newBuf); }

    static UniMode fromBOM(const unsigned char* b, size_t n) {
        if (n > 1 && b[0] == 0xFE && b[1] == 0xFF) return uni16BE;
        if (n > 1 && b[0] == 0xFF && b[1] == 0xFE) return uni16LE;
        if (n > 2 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return uniUTF8;
        return uni8Bit;
    }

    size_t convert(char* buf, size_t len) {
        static size_t nSkip = 0;
        _buf = reinterpret_cast<unsigned char*>(buf); _len = len; _newSize = 0;
        if (_first) { determineEncoding(); nSkip = _skip; _first = false; }
        switch (_enc) {
        case uni7Bit: case uni8Bit: case uniUTF8_NoBOM: _allocated = 0; _newBuf = _buf; _newSize = len; break;
        case uniUTF8: _allocated = 0; _newBuf = _buf + nSkip; _newSize = len - nSkip; break;
        case uni16BE_NoBOM: case uni16LE_NoBOM: case uni16BE: case uni16LE: {
            const size_t newSize = (len + len % 2) + (len + len % 2) / 2 + 4;
            if (_allocated != newSize) { delete[] (is16() && _newBuf != _buf ? _newBuf : nullptr); _newBuf = new unsigned char[newSize]; _allocated = newSize; }
            unsigned char* out = _newBuf;
            const bool be = (_enc == uni16BE || _enc == uni16BE_NoBOM);
            const unsigned char* p = _buf + nSkip;
            const size_t n = (len - nSkip) / 2;
            for (size_t i = 0; i < n; ++i) {
                uint32_t c = be ? (p[i * 2] << 8 | p[i * 2 + 1]) : (p[i * 2] | p[i * 2 + 1] << 8);
                if (c >= 0xD800 && c <= 0xDBFF && i + 1 < n) {
                    const uint32_t lo = be ? (p[i * 2 + 2] << 8 | p[i * 2 + 3]) : (p[i * 2 + 2] | p[i * 2 + 3] << 8);
                    if (lo >= 0xDC00 && lo <= 0xDFFF) { c = 0x10000 + ((c - 0xD800) << 10) + (lo - 0xDC00); ++i; }
                }
                if (c < 0x80) *out++ = static_cast<unsigned char>(c);
                else if (c < 0x800) { *out++ = static_cast<unsigned char>(0xC0 | (c >> 6)); *out++ = static_cast<unsigned char>(0x80 | (c & 0x3F)); }
                else if (c < 0x10000) { *out++ = static_cast<unsigned char>(0xE0 | (c >> 12)); *out++ = static_cast<unsigned char>(0x80 | ((c >> 6) & 0x3F)); *out++ = static_cast<unsigned char>(0x80 | (c & 0x3F)); }
                else { *out++ = static_cast<unsigned char>(0xF0 | (c >> 18)); *out++ = static_cast<unsigned char>(0x80 | ((c >> 12) & 0x3F)); *out++ = static_cast<unsigned char>(0x80 | ((c >> 6) & 0x3F)); *out++ = static_cast<unsigned char>(0x80 | (c & 0x3F)); }
            }
            _newSize = static_cast<size_t>(out - _newBuf);
            break;
        }
        default: break;
        }
        nSkip = 0;
        return _newSize;
    }

private:
    UniMode _enc = uni8Bit;
    unsigned char* _buf = nullptr; size_t _len = 0;
    unsigned char* _newBuf = nullptr; size_t _newSize = 0, _allocated = 0, _skip = 0;
    bool _first = true;
    bool is16() const { return _enc == uni16BE || _enc == uni16LE || _enc == uni16BE_NoBOM || _enc == uni16LE_NoBOM; }

    void determineEncoding() {
        _enc = uni8Bit; _skip = 0;
        if (_len > 1 && _buf[0] == 0xFE && _buf[1] == 0xFF) { _enc = uni16BE; _skip = 2; }
        else if (_len > 1 && _buf[0] == 0xFF && _buf[1] == 0xFE) { _enc = uni16LE; _skip = 2; }
        else if (_len > 2 && _buf[0] == 0xEF && _buf[1] == 0xBB && _buf[2] == 0xBF) { _enc = uniUTF8; _skip = 3; }
        else {
            const int d = utf8_7bits_8bits();
            _enc = d == 0 ? uniUTF8_NoBOM : d == 1 ? uni7Bit : uni8Bit;
        }
    }
    int utf8_7bits_8bits() const {   // 0 utf8, 1 7 bits, 2 8 bits
        int rv = 1, ascii7only = 1;
        const unsigned char* sx = _buf; const unsigned char* endx = _buf + _len;
        while (sx < endx) {
            if (*sx == 0) { ascii7only = 0; rv = 0; break; }
            else if ((*sx & 0x80) == 0) ++sx;
            else if ((*sx & 0xC0) == 0x80) { ascii7only = 0; rv = 0; break; }
            else if ((*sx & 0xE0) == 0xC0) { ascii7only = 0; if (endx - sx < 2 || (sx[1] & 0xC0) != 0x80) { rv = 0; break; } sx += 2; }
            else if ((*sx & 0xF0) == 0xE0) { ascii7only = 0; if (endx - sx < 3 || (sx[1] & 0xC0) != 0x80 || (sx[2] & 0xC0) != 0x80) { rv = 0; break; } sx += 3; }
            else if ((*sx & 0xF8) == 0xF0) { ascii7only = 0; if (endx - sx < 4 || (sx[1] & 0xC0) != 0x80 || (sx[2] & 0xC0) != 0x80 || (sx[3] & 0xC0) != 0x80) { rv = 0; break; } sx += 4; }
            else { ascii7only = 0; rv = 0; break; }
        }
        if (ascii7only) return 1;
        if (rv) return 0;
        return 2;
    }
};

static bool inAliasList(const char* token, const char* list) {
    char word[64]; size_t j = 0;
    for (size_t i = 0, len = std::strlen(list); i <= len; ++i) {
        if (list[i] == ' ' || list[i] == '\0') { if (j) { word[j] = 0; j = 0; if (strcasecmp(token, word) == 0) return true; } }
        else if (j < sizeof(word) - 1) word[j++] = list[i];
    }
    return false;
}

static int encodingFromString(const char* alias) {
    static const std::pair<int, const char*> table[] = {
        { 1250, "windows-1250" }, { 1251, "windows-1251" }, { 1252, "windows-1252" }, { 1253, "windows-1253" }, { 1254, "windows-1254" },
        { 1255, "windows-1255" }, { 1256, "windows-1256" }, { 1257, "windows-1257" }, { 1258, "windows-1258" },
        { 28591, "latin1 ISO_8859-1 ISO-8859-1 CP819 IBM819 csISOLatin1 iso-ir-100 l1" }, { 28592, "latin2 ISO_8859-2 ISO-8859-2 csISOLatin2 iso-ir-101 l2" },
        { 28593, "latin3 ISO_8859-3 ISO-8859-3 csISOLatin3 iso-ir-109 l3" }, { 28594, "latin4 ISO_8859-4 ISO-8859-4 csISOLatin4 iso-ir-110 l4" },
        { 28595, "cyrillic ISO_8859-5 ISO-8859-5 csISOLatinCyrillic iso-ir-144" }, { 28596, "arabic ISO_8859-6 ISO-8859-6 csISOLatinArabic iso-ir-127 ASMO-708 ECMA-114" },
        { 28597, "greek ISO_8859-7 ISO-8859-7 csISOLatinGreek greek8 iso-ir-126 ELOT_928 ECMA-118" }, { 28598, "hebrew ISO_8859-8 ISO-8859-8 csISOLatinHebrew iso-ir-138" },
        { 28599, "latin5 ISO_8859-9 ISO-8859-9 csISOLatin5 iso-ir-148 l5" }, { 28603, "ISO_8859-13 ISO-8859-13" }, { 28604, "iso-celtic latin8 ISO_8859-14 ISO-8859-14 18 iso-ir-199" },
        { 28605, "Latin-9 ISO_8859-15 ISO-8859-15" }, { 437, "IBM437 cp437 437 csPC8CodePage437" }, { 720, "IBM720 cp720 oem720 720" }, { 737, "IBM737 cp737 oem737 737" },
        { 775, "IBM775 cp775 oem775 775" }, { 850, "IBM850 cp850 oem850 850" }, { 852, "IBM852 cp852 oem852 852" }, { 855, "IBM855 cp855 oem855 855 csIBM855" },
        { 857, "IBM857 cp857 oem857 857" }, { 858, "IBM858 cp858 oem858 858" }, { 860, "IBM860 cp860 oem860 860" }, { 861, "IBM861 cp861 oem861 861" },
        { 862, "IBM862 cp862 oem862 862" }, { 863, "IBM863 cp863 oem863 863" }, { 865, "IBM865 cp865 oem865 865" }, { 866, "IBM866 cp866 oem866 866" },
        { 869, "IBM869 cp869 oem869 869" }, { 950, "big5 csBig5" }, { 936, "gb2312 gbk csGB2312 gb18030" }, { 932, "Shift_JIS MS_Kanji csShiftJIS csWindows31J" },
        { 949, "windows-949 korean" }, { 51949, "euc-kr csEUCKR" }, { 874, "tis-620" }, { 10007, "x-mac-cyrillic xmaccyrillic" }, { 21866, "koi8_u" }, { 20866, "koi8_r csKOI8R" } };
    if (inAliasList(alias, "utf-8 utf8")) return sc::CP_UTF8;
    for (const auto& e : table) if (inAliasList(alias, e.second)) return e.first;
    return -1;
}

static int detectCodepage(const char* buf, size_t len) {
    int cp = -1;
    uchardet_t ud = uchardet_new();
    uchardet_handle_data(ud, buf, len);
    uchardet_data_end(ud);
    const char* cs = uchardet_get_charset(ud);
    if (strcasecmp(cs, "TIS-620") != 0) cp = encodingFromString(cs);
    uchardet_delete(ud);
    return cp;
}

static LangType detectLanguageFromTextBeginning(const unsigned char* data, size_t dataLen) {
    if (dataLen <= 3) return L_TEXT;
    size_t i = 0;
    if ((data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) || (data[0] == 0xFE && data[1] == 0xFF && data[2] == 0x00) || (data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00)) i += 3;
    for (; i < dataLen; ++i) if (data[i] != ' ' && data[i] != '\t' && data[i] != '\n' && data[i] != '\r') break;
    if (i == dataLen) return L_TEXT;
    const size_t longest = std::min<size_t>(40, dataLen - i);
    std::string buf2Test(reinterpret_cast<const char*>(data) + i, longest);
    const size_t crnl = std::min(buf2Test.find('\r'), buf2Test.find('\n'));
    if (crnl != std::string::npos && crnl < longest) buf2Test = std::string(reinterpret_cast<const char*>(data) + i, crnl);
    if (buf2Test.find("#!") == 0) {
        static const std::pair<const char*, LangType> sheb[] = { { "sh", L_BASH }, { "python", L_PYTHON }, { "perl", L_PERL }, { "php", L_PHP }, { "ruby", L_RUBY }, { "node", L_JAVASCRIPT } };
        for (const auto& s : sheb) if (buf2Test.find(s.first) != std::string::npos) return s.second;
        return L_TEXT;
    }
    static const std::pair<const char*, LangType> first[] = { { "<?xml", L_XML }, { "<?php", L_PHP }, { "<html", L_HTML }, { "<!DOCTYPE html", L_HTML }, { "<?", L_PHP } };
    for (const auto& f : first) if (buf2Test.find(f.first) == 0) return f.second;
    return L_TEXT;
}

static Eol eolFormat(const char* data, size_t n) {
    for (size_t i = 0; i < n; ++i) {
        if (data[i] == '\r') return (i + 1 < n && data[i + 1] == '\n') ? Eol::windows : Eol::macos;
        if (data[i] == '\n') return Eol::unix_;
    }
    return Eol::unknown;
}

struct FoundInfo {
    FoundInfo() = default;
    FoundInfo(intptr_t s, intptr_t e, size_t line, const char16_t* path) : _lineNumber(line), _fullPath(path) { _ranges.push_back({ s, e }); }
    std::vector<std::pair<intptr_t, intptr_t>> _ranges;
    size_t _lineNumber = 0;
    std::u16string _fullPath;
};
struct Marking { std::vector<std::pair<intptr_t, intptr_t>> _segmentPostions; };

class Finder {
public:
    SciDoc view{ true };
    Wmc& wmc;
    explicit Finder(Wmc& w) : wmc(w) {}

    void beginNewFilesSearch() { _pMainInfos = &_infos1; _pMainMarks = &_marks1; _nbFoundFiles = 0; }
    void addSearchLine(const std::u16string& name) {
        std::u16string str = u"Search \""; str += name; str += u"\" \r\n";
        addGenericText(str);
        _lastSearchHeaderPos = view.length() - 2;
        _pMainInfos->push_back(FoundInfo{}); _pMainMarks->push_back(Marking{}); _previousLine = -1;
    }
    void addFileNameTitle(const char16_t* fileName) {
        std::u16string str = u"  "; str += fileName; str += u"\r\n";
        addGenericText(str);
        _lastFileHeaderPos = view.length() - 2;
        _pMainInfos->push_back(FoundInfo{}); _pMainMarks->push_back(Marking{}); _previousLine = -1;
    }
    void addFileHitCount(int count) {
        std::u16string text = u" ("; text += w32::widen(std::to_string(count)); text += count == 1 ? u" hit)" : u" hits)";
        int n = 0; const char* a = wmc.wchar2char(text.c_str(), sc::CP_UTF8, static_cast<int>(text.size()), &n);
        view.insert(_lastFileHeaderPos, a, n);
        ++_nbFoundFiles;
    }
    std::string foundLine(FoundInfo fi, Marking mi, const w32::wchar* line, size_t lineLen, size_t totalLineNumber) {
        bool repeated = false;
        if (_previousLine == -1) _previousLine = static_cast<intptr_t>(fi._lineNumber);
        else if (_previousLine == static_cast<intptr_t>(fi._lineNumber)) repeated = true;
        else _previousLine = static_cast<intptr_t>(fi._lineNumber);
        std::u16string headerStr = u"\t"; headerStr += _prefixLine; headerStr += u" ";
        const size_t totalDigits = nbDigits(totalLineNumber) + 1, curDigits = nbDigits(fi._lineNumber) + 1;
        std::u16string lineNumberStr; lineNumberStr.append(totalDigits - curDigits, u' '); lineNumberStr.append(w32::widen(std::to_string(fi._lineNumber)));
        headerStr += lineNumberStr; headerStr += u": ";
        mi._segmentPostions[0].first += static_cast<intptr_t>(headerStr.length());
        mi._segmentPostions[0].second += static_cast<intptr_t>(headerStr.length());
        headerStr += std::u16string(line, lineLen);
        int len8 = 0;
        const char* text8 = wmc.wchar2charMarks(headerStr.c_str(), sc::CP_UTF8, &mi._segmentPostions[0].first, &mi._segmentPostions[0].second, static_cast<int>(headerStr.length()), &len8);
        if (repeated && (len8 < static_cast<int>(kLineBufMax) || mi._segmentPostions[0].second < static_cast<intptr_t>(kLineBufMax))) {
            _pMainMarks->back()._segmentPostions.push_back(mi._segmentPostions[0]);
            _pMainInfos->back()._ranges.push_back(fi._ranges.back());
            return "";
        }
        _pMainInfos->push_back(fi);
        if (len8 >= static_cast<int>(kLineBufMax)) {
            static const char endOfLongLine[] = " ...\r\n";
            size_t cut = kLineBufMax - (sizeof(endOfLongLine) - 1) - 1;
            while (cut > 0 && (static_cast<unsigned char>(text8[cut]) & 0xC0) == 0x80) --cut;
            std::string s(text8, cut); s += endOfLongLine;
            _pMainMarks->push_back(mi);
            return s;
        }
        _pMainMarks->push_back(mi);
        return std::string(text8, static_cast<size_t>(len8));
    }
    void finishFilesSearch(int count, int searched) {
        std::vector<FoundInfo>* old = _pMainInfos == &_infos1 ? &_infos2 : &_infos1;
        std::vector<Marking>* oldM = _pMainMarks == &_marks1 ? &_marks2 : &_marks1;
        old->insert(old->begin(), _pMainInfos->begin(), _pMainInfos->end());
        oldM->insert(oldM->begin(), _pMainMarks->begin(), _pMainMarks->end());
        _pMainInfos->clear(); _pMainMarks->clear();
        _pMainInfos = old; _pMainMarks = oldM;
        std::u16string info = u" (" + w32::widen(std::to_string(count)) + u" hits in " + w32::widen(std::to_string(_nbFoundFiles)) + u" files of " + w32::widen(std::to_string(searched)) + u" searched)";
        int n = 0; const char* a = wmc.wchar2char(info.c_str(), sc::CP_UTF8, static_cast<int>(info.size()), &n);
        view.insert(_lastSearchHeaderPos, a, n);
        _previousLine = -1;
    }

private:
    std::vector<FoundInfo> _infos1, _infos2; std::vector<Marking> _marks1, _marks2;
    std::vector<FoundInfo>* _pMainInfos = &_infos1; std::vector<Marking>* _pMainMarks = &_marks1;
    intptr_t _previousLine = -1;
    Sci::Position _lastFileHeaderPos = 0, _lastSearchHeaderPos = 0;
    int _nbFoundFiles = 0;
    const std::u16string _prefixLine = u"Line";

    static size_t nbDigits(size_t n) { size_t d = 1; while (n >= 10) { n /= 10; ++d; } return d; }
    void addGenericText(const std::u16string& s) {
        int n = 0; const char* a = wmc.wchar2char(s.c_str(), sc::CP_UTF8, static_cast<int>(s.size()), &n);
        view.append(a, n);
    }
};

// Buffer as the old path creates one per file (fields that cost allocations)
struct Buffer {
    std::u16string fullPath, fileName, tabCreatedTime, userReadOnlyReason;
    std::unique_ptr<SciDoc> doc;
    UniMode unicodeMode = uniUTF8_NoBOM;
    int encoding = -1;
    LangType lang = L_TEXT;
};

class Tool {
public:
    explicit Tool(bool newPath) : _newPath(newPath) {}

    RunResult run(const Corpus& corpus, const Query& q) {
        RunResult rr;
        const auto t0 = Clock::now();
        Wmc wmc;
        Finder finder(wmc);
        finder.beginNewFilesSearch();
        const std::u16string pat16 = w32::widen(q.text);
        finder.addSearchLine(pat16);
        std::unique_ptr<SciDoc> scratch;
        if (_newPath) scratch = std::make_unique<SciDoc>(false);
        int nbTotal = 0;
        for (const FileEntry& fe : corpus.files) {
            {   // getBufferFromName: no tab holds this file
                PhaseTimer pt(rr.phase[PH_BUFFER]);
                static const char16_t* openTabs[] = { u"C:\\notes\\todo.txt", u"C:\\work\\plan.md", u"new 1" };
                for (const char16_t* t : openTabs) if (fe.path16 == t) return rr;
            }
            SciDoc* doc = nullptr;
            UniMode mode = uni8Bit;
            std::unique_ptr<Buffer> buf;
            if (_newPath) {
                if (!loadFileContentForSearch(*scratch, fe, rr, wmc, mode)) continue;
                doc = scratch.get();
                doc->setCodePage(0);
                setCodePageForInvisibleView(*doc, mode);
            }
            else {
                buf = loadFile(fe, rr, wmc);
                if (!buf) continue;
                doc = buf->doc.get();
                setCodePageForInvisibleView(*doc, buf->unicodeMode);
            }
            rr.textBytesLoaded += static_cast<uint64_t>(doc->length());
            int nb = 0;
            {
                PhaseTimer pt(rr.phase[PH_SEARCH]);
                nb = processRangeFindAll(*doc, finder, wmc, fe.path16.c_str(), pat16, q);
            }
            if (nb > 0) { rr.hits += static_cast<size_t>(nb); ++rr.filesWithHits; if (fe.binary) rr.binaryHits += static_cast<size_t>(nb); }
            nbTotal += nb;
            if (buf) { PhaseTimer pt(rr.phase[PH_BUFFER]); buf.reset(); }   // closeBuffer: SCI_RELEASEDOCUMENT + delete
        }
        {
            PhaseTimer pt(rr.phase[PH_FINAL]);
            finder.finishFilesSearch(nbTotal, static_cast<int>(corpus.files.size()));
        }
        rr.total = msBetween(t0, Clock::now());
        return rr;
    }

private:
    bool _newPath;

    struct LoadedFileFormat { int encoding = -1; Eol eol = Eol::unknown; LangType lang = L_TEXT; };

    static void setCodePageForInvisibleView(SciDoc& doc, UniMode mode) {
        const int detected = doc.cp;
        int cp2set = sc::CP_UTF8;
        if (mode == uni8Bit) cp2set = (detected == sc::CP_UTF8 ? 0 : detected);
        doc.setCodePage(cp2set);
    }

    // copyFileContentToScintilla (forSearching: no language detection; forEditing: with it)
    bool copyFileContent(SciDoc& doc, FILE* fp, char* data, Utf8_16_Read& conv, LoadedFileFormat& ff, bool detectLanguage, RunResult& rr, Wmc& wmc) {
        size_t lenFile = 0;
        bool isFirstTime = true;
        int incomplete = 0;
        Eol format = Eol::unknown;
        do {
            auto t = Clock::now();
            lenFile = std::fread(data + incomplete, 1, blockSize - incomplete, fp) + incomplete;
            rr.phase[PH_OPEN] += msBetween(t, Clock::now());
            if (std::ferror(fp)) return false;
            if (lenFile == 0) break;
            if (isFirstTime) {
                PhaseTimer pt(rr.phase[PH_DETECT]);
                if (Utf8_16_Read::fromBOM(reinterpret_cast<unsigned char*>(data), lenFile) != uni8Bit) ff.encoding = -1;
                else if (ff.encoding == -1) ff.encoding = detectCodepage(data, lenFile);
                if (detectLanguage && ff.lang == L_TEXT) ff.lang = detectLanguageFromTextBeginning(reinterpret_cast<unsigned char*>(data), lenFile);
                isFirstTime = false;
            }
            if (ff.encoding != -1) {
                if (ff.encoding == sc::CP_UTF8) { PhaseTimer pt(rr.phase[PH_FILL]); doc.append(data, static_cast<Sci::Position>(lenFile)); }
                else {
                    int newLen = 0;
                    const char* newData;
                    { PhaseTimer pt(rr.phase[PH_DETECT]); newData = wmc.encode(static_cast<unsigned>(ff.encoding), sc::CP_UTF8, data, static_cast<int>(lenFile), &newLen); }
                    PhaseTimer pt(rr.phase[PH_FILL]); doc.append(newData, newLen);
                }
                if (format == Eol::unknown) format = eolFormat(data, lenFile);
            }
            else {
                size_t lenConvert;
                { PhaseTimer pt(rr.phase[PH_DETECT]); lenConvert = conv.convert(data, lenFile); }
                { PhaseTimer pt(rr.phase[PH_FILL]); doc.append(conv.newBuf(), static_cast<Sci::Position>(lenConvert)); }
                if (format == Eol::unknown) format = eolFormat(conv.newBuf(), lenConvert);
            }
        } while (lenFile > 0);
        ff.eol = format;
        return true;
    }

    static void resolveLoadedEncoding(const Utf8_16_Read& conv, int& enc, UniMode& mode) {
        mode = conv.encoding();
        if (enc == -1) { if (mode == uni7Bit) mode = uniUTF8_NoBOM; }   // openAnsiAsUtf8
        else { enc = (enc == sc::CP_UTF8) ? -1 : enc; mode = uniUTF8_NoBOM; }
    }

    bool loadFileContentForSearch(SciDoc& scratch, const FileEntry& fe, RunResult& rr, Wmc& wmc, UniMode& mode) {
        auto t = Clock::now();
        FILE* fp = std::fopen(fe.path.c_str(), "rb");
        rr.phase[PH_OPEN] += msBetween(t, Clock::now());
        if (!fp) return false;
        std::unique_ptr<char[]> data(new char[blockSize + 8]);
        Utf8_16_Read conv;
        LoadedFileFormat ff;
        { PhaseTimer pt(rr.phase[PH_FILL]); scratch.clearAll(); scratch.allocate(static_cast<Sci::Position>(fe.size)); }
        const bool ok = copyFileContent(scratch, fp, data.get(), conv, ff, false, rr, wmc);
        t = Clock::now(); std::fclose(fp); rr.phase[PH_OPEN] += msBetween(t, Clock::now());
        if (!ok) return false;
        int enc = ff.encoding;
        resolveLoadedEncoding(conv, enc, mode);
        return true;
    }

    // getFileAttributesExWithTimeout: a worker thread per call
    static bool attributesWithTimeout(const std::string& path, uint64_t& size) {
        bool ok = false;
        std::thread worker([&] { struct stat st {}; ok = ::stat(path.c_str(), &st) == 0; size = static_cast<uint64_t>(st.st_size); });
        worker.join();
        return ok;
    }

    std::unique_ptr<Buffer> loadFile(const FileEntry& fe, RunResult& rr, Wmc& wmc) {
        uint64_t size = 0;
        {
            PhaseTimer pt(rr.phase[PH_BUFFER]);
            uint64_t dummy = 0;
            if (!attributesWithTimeout(fe.path, dummy)) return nullptr;     // doesFileExist
            if (!attributesWithTimeout(fe.path, size)) return nullptr;      // file size
        }
        const bool isLarge = size >= 200ull * 1024 * 1024;
        auto doc = std::make_unique<SciDoc>(!isLarge);
        auto t = Clock::now();
        FILE* fp = std::fopen(fe.path.c_str(), "rb");
        rr.phase[PH_OPEN] += msBetween(t, Clock::now());
        if (!fp) return nullptr;
        std::unique_ptr<char[]> data(new char[blockSize + 8]);
        Utf8_16_Read conv;
        LoadedFileFormat ff;
        const int64_t requested = static_cast<int64_t>(size) + std::min<int64_t>(1LL << 20, static_cast<int64_t>(size) / 6);
        { PhaseTimer pt(rr.phase[PH_FILL]); doc->clearAll(); doc->allocate(requested); }
        const bool ok = copyFileContent(*doc, fp, data.get(), conv, ff, true, rr, wmc);
        t = Clock::now(); std::fclose(fp); rr.phase[PH_OPEN] += msBetween(t, Clock::now());
        if (!ok) return nullptr;
        PhaseTimer pt(rr.phase[PH_BUFFER]);
        auto buf = std::make_unique<Buffer>();
        buf->fullPath = fe.path16;
        const size_t slash = fe.path16.find_last_of(u'/');
        buf->fileName = fe.path16.substr(slash == std::u16string::npos ? 0 : slash + 1);
        static const char16_t* exts[] = { u"txt", u"php", u"c", u"cpp", u"h", u"hpp", u"cs", u"java", u"js", u"ts", u"py", u"rb", u"pl", u"sh", u"bat", u"ps1",
            u"xml", u"html", u"htm", u"css", u"json", u"yaml", u"yml", u"ini", u"md", u"sql", u"lua", u"go", u"rs", u"swift", u"kt", u"vb", u"pas", u"asm",
            u"tex", u"r", u"m", u"f90", u"cmake", u"mak", u"diff", u"props", u"nsi", u"iss", u"inno", u"au3", u"ada", u"lisp", u"scm", u"erl", u"hs" };
        const size_t dot = buf->fileName.find_last_of(u'.');
        const std::u16string ext = dot == std::u16string::npos ? u"" : buf->fileName.substr(dot + 1);
        for (const char16_t* e : exts) if (ext == e) { buf->lang = L_TEXT; break; }
        uint64_t s2 = 0;
        attributesWithTimeout(fe.path, s2);        // updateTimeStamp
        attributesWithTimeout(fe.path, s2);        // checkFileState
        int enc = ff.encoding;
        resolveLoadedEncoding(conv, enc, buf->unicodeMode);
        buf->encoding = enc;
        buf->doc = std::move(doc);
        return buf;
    }

    // FindReplaceDlg::processAll / processRange with ProcessFindAll
    int processRangeFindAll(SciDoc& doc, Finder& finder, Wmc& wmc, const char16_t* fileName, const std::u16string& pat16, const Query& q) {
        if (doc.cp != sc::CP_UTF8) {   // isSearchUnicodeCharOnAnsi
            for (char16_t c : pat16) if (w32::uToSb(c) < 0) return 0;
        }
        const size_t docLength = static_cast<size_t>(doc.length());
        if (docLength == 0) return 0;
        const size_t len16 = pat16.size();
        std::unique_ptr<char16_t[]> pTextFind(new char16_t[len16 + 1]);
        std::memcpy(pTextFind.get(), pat16.c_str(), (len16 + 1) * sizeof(char16_t));
        int flags = (q.matchCase ? sc::MATCHCASE : 0) | (q.regex ? (sc::REGEXP | sc::POSIX) : 0) | sc::SKIPCRLFASONE | sc::EMPTY_NOTAFTER;
        auto text2AddUtf8 = std::make_unique<std::string>();
        bool fileTitleAdded = false;
        intptr_t targetStart = 0;
        size_t startRange = 0, endRange = docLength;
        int nb = 0;
        while (targetStart >= 0) {
            const char* pat8 = wmc.wchar2char(pTextFind.get(), static_cast<unsigned>(doc.cp == sc::CP_UTF8 ? sc::CP_UTF8 : w32::ACP));
            const size_t pat8Len = std::strlen(pat8);
            Sci::Position targetEnd = 0;
            targetStart = doc.searchInTarget(static_cast<Sci::Position>(startRange), static_cast<Sci::Position>(endRange), pat8, std::max(len16, pat8Len), flags, targetEnd);
            if (targetStart == -1) break;
            if (targetStart < -1) return 0;
            if (static_cast<size_t>(targetEnd) > endRange) break;
            const intptr_t foundTextLen = targetEnd - targetStart;
            if (!fileTitleAdded) { finder.addFileNameTitle(fileName); fileTitleAdded = true; }
            const size_t totalLineNumber = static_cast<size_t>(doc.lineCount());
            const Sci::Line lineNumber = doc.lineFromPos(targetStart);
            intptr_t lend = doc.lineEnd(lineNumber);
            const intptr_t lstart = doc.lineStart(lineNumber);
            intptr_t nbChar = lend - lstart;
            char16_t lineBuf[kLineBufMax]{};
            if (nbChar > static_cast<intptr_t>(kLineBufMax) - 3) { lend = lstart + kLineBufMax - 4; nbChar = kLineBufMax - 4; }
            intptr_t startMark = targetStart - lstart, endMark = targetEnd - lstart;
            getGenericText(doc, wmc, lineBuf, kLineBufMax, static_cast<size_t>(lstart), static_cast<size_t>(lend), &startMark, &endMark, &nbChar);
            lineBuf[nbChar++] = u'\r';
            lineBuf[nbChar++] = u'\n';
            Marking srml; srml._segmentPostions.push_back({ startMark, endMark });
            text2AddUtf8->append(finder.foundLine(FoundInfo(targetStart, targetEnd, static_cast<size_t>(lineNumber) + 1, fileName), srml, lineBuf, static_cast<size_t>(nbChar), totalLineNumber));
            if (text2AddUtf8->length() > kFindTempMax) { finder.view.append(text2AddUtf8->data(), static_cast<Sci::Position>(text2AddUtf8->size())); text2AddUtf8->clear(); }
            ++nb;
            if (static_cast<size_t>(targetStart + foundTextLen) == endRange) break;
            startRange = static_cast<size_t>(targetStart + foundTextLen);
        }
        if (nb > 0) {
            finder.view.append(text2AddUtf8->data(), static_cast<Sci::Position>(text2AddUtf8->size()));
            finder.addFileHitCount(nb);
        }
        return nb;
    }

    static void getGenericText(SciDoc& doc, Wmc& wmc, char16_t* dest, size_t destlen, size_t start, size_t end, intptr_t* ms, intptr_t* me, intptr_t* outLen) {
        const size_t nbChar = end - start;
        if (nbChar == 0) { dest[0] = 0; return; }
        std::unique_ptr<char[]> destA(new char[nbChar + 1]);
        doc.getText(destA.get(), static_cast<Sci::Position>(start), static_cast<Sci::Position>(end));
        const unsigned cp = static_cast<unsigned>(doc.cp == sc::CP_UTF8 ? sc::CP_UTF8 : w32::ACP);
        const char16_t* destW = wmc.char2wcharMarks(destA.get(), cp, ms, me, static_cast<int>(nbChar));
        size_t lenW = wmc.sizeW();
        if (lenW >= destlen) lenW = destlen - 1;
        std::memcpy(dest, destW, lenW * sizeof(char16_t));
        dest[lenW] = 0;
        if (outLen) *outLen = static_cast<intptr_t>(lenW);
    }
};
}

// ---------------------------------------------------------------------------------------------
// Statistics
// ---------------------------------------------------------------------------------------------
static double tQuantile975(int df) {
    static const double t[] = { 0, 12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.160, 2.145, 2.131,
        2.120, 2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042 };
    if (df <= 0) return 0;
    if (df <= 30) return t[df];
    if (df <= 60) return 2.000;
    return 1.960;
}

struct Summary { int n = 0; double median = 0, mean = 0, sd = 0, cv = 0, ciLo = 0, ciHi = 0, min = 0, max = 0; };

static double medianOf(std::vector<double> v) {
    if (v.empty()) return 0;
    std::sort(v.begin(), v.end());
    const size_t n = v.size();
    return n % 2 ? v[n / 2] : 0.5 * (v[n / 2 - 1] + v[n / 2]);
}

static Summary summarize(const std::vector<double>& v) {
    Summary s; s.n = static_cast<int>(v.size());
    if (v.empty()) return s;
    s.median = medianOf(v);
    s.mean = std::accumulate(v.begin(), v.end(), 0.0) / v.size();
    double ss = 0; for (double x : v) ss += (x - s.mean) * (x - s.mean);
    s.sd = v.size() > 1 ? std::sqrt(ss / (v.size() - 1)) : 0;
    s.cv = s.mean > 0 ? 100.0 * s.sd / s.mean : 0;
    const double h = tQuantile975(s.n - 1) * s.sd / std::sqrt(static_cast<double>(s.n));
    s.ciLo = s.mean - h; s.ciHi = s.mean + h;
    s.min = *std::min_element(v.begin(), v.end()); s.max = *std::max_element(v.begin(), v.end());
    return s;
}

// Paired ratio a_i / b_i over rounds: median and 95 % percentile bootstrap interval
static void pairedRatio(const std::vector<double>& a, const std::vector<double>& b, uint64_t seed, double& med, double& lo, double& hi) {
    std::vector<double> r;
    for (size_t i = 0; i < a.size() && i < b.size(); ++i) if (b[i] > 0) r.push_back(a[i] / b[i]);
    med = medianOf(r); lo = hi = med;
    if (r.size() < 3) return;
    std::mt19937_64 rng(seed);
    std::vector<double> boots, sample(r.size());
    for (int it = 0; it < 4000; ++it) { for (auto& x : sample) x = r[rng() % r.size()]; boots.push_back(medianOf(sample)); }
    std::sort(boots.begin(), boots.end());
    lo = boots[static_cast<size_t>(0.025 * boots.size())];
    hi = boots[static_cast<size_t>(0.975 * boots.size()) - 1];
}

// ---------------------------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------------------------
struct Cell { const Corpus* corpus; const Query* query; std::string tool; };
struct Sample { std::string corpus, query, tool; int round; RunResult r; };

// MR and its variants for a step-by-step comparison
static bool mrVariant(const std::string& name, mr::Variant& v) {
    if (name == "MR") v = {};
    else if (name == "MR+utf8") v = { true, false, false };
    else if (name == "MR+prefilter") v = { false, true, false };
    else if (name == "MR+dock") v = { false, false, true };
    else if (name == "MR+all") v = { true, true, true };
    else return false;
    return true;
}

static RunResult runCell(const Cell& c) {
    mr::Variant v;
    if (mrVariant(c.tool, v)) { mr::Tool t(v); return t.run(*c.corpus, *c.query); }
    npp::Tool t(c.tool == "NppNew");
    return t.run(*c.corpus, *c.query);
}

static bool isKnownTool(const std::string& name) { mr::Variant v; return mrVariant(name, v) || name == "NppNew" || name == "NppOld"; }

static std::vector<Query> queriesFor(const std::string& kind) {
    if (kind == "source" || kind == "source-text") return {
        { "literal absent", "zqxjvqk", false, true },
        { "literal rare", "WM_SETREDRAW", false, true },
        { "literal common", "Fi", false, true },
        { "literal very common", "in", false, true },
        { "regex Fi (guy038)", "Fi", true, true },
        { "ignore case", "cache", false, false },
        { "regex \\bstd::\\w+", "\\bstd::\\w+", true, true } };
    return {
        { "literal absent", "zqxjvqk", false, true },
        { "literal rare", "Zebrafinkenkäfig", false, true },
        { "literal common", "und", false, true },
        { "regex Fi (guy038)", "Fi", true, true },
        { "ignore case", "straße", false, false },
        { "regex \\b\\w+ung\\b", "\\b\\w+ung\\b", true, true } };
}

// Writes to stdout and to the summary file
static FILE* g_summary = nullptr;
__attribute__((format(printf, 1, 2))) static void both(const char* fmt, ...) {
    va_list a, b;
    va_start(a, fmt);
    va_copy(b, a);
    std::vprintf(fmt, a);
    if (g_summary) std::vfprintf(g_summary, fmt, b);
    va_end(b);
    va_end(a);
}

static std::string cpuModel() {
    std::ifstream f("/proc/cpuinfo");
    std::string line;
    while (std::getline(f, line))
        if (line.rfind("model name", 0) == 0) { const size_t c = line.find(':'); return c == std::string::npos ? line : line.substr(c + 2); }
    return "unknown CPU";
}

static std::string utcNow() {
    const std::time_t now = std::time(nullptr);
    char buf[32] = "unknown";
    if (const std::tm* t = std::gmtime(&now)) std::strftime(buf, sizeof buf, "%Y-%m-%d %H:%M UTC", t);
    return buf;
}

static std::string kernelName() {
    struct utsname u {};
    if (uname(&u) != 0) return "unknown system";
    return std::string(u.sysname) + " " + u.release;
}

#ifdef NDEBUG
#define BENCH_BUILD_MODE "release (NDEBUG)"
#else
#define BENCH_BUILD_MODE "debug"
#endif

int main(int argc, char** argv) {
    std::vector<std::pair<std::string, std::string>> corpusArgs;   // kind, dir
    int rounds = 15; uint64_t seed = 20261009; std::string out = "fif_results";
    std::string genDir; uint64_t genSeed = 1; int genFiles = 1500;
    std::vector<std::string> tools = { "MR", "NppNew", "NppOld" };
    std::vector<std::string> onlyQueries;
    auto splitList = [](const std::string& s) {
        std::vector<std::string> out;
        for (size_t p = 0; p <= s.size();) { size_t q = s.find(',', p); if (q == std::string::npos) q = s.size(); out.push_back(s.substr(p, q - p)); p = q + 1; }
        return out;
    };
    for (int i = 1; i < argc; ++i) {
        const std::string a = argv[i];
        auto next = [&]() -> std::string { return i + 1 < argc ? argv[++i] : ""; };
        if (a == "--source") corpusArgs.push_back({ "source", next() });
        else if (a == "--source-text") corpusArgs.push_back({ "source-text", next() });
        else if (a == "--mixed") corpusArgs.push_back({ "mixed", next() });
        else if (a == "--rounds") rounds = std::atoi(next().c_str());
        else if (a == "--seed") seed = std::strtoull(next().c_str(), nullptr, 10);
        else if (a == "--out") out = next();
        else if (a == "--generate-mixed") genDir = next();
        else if (a == "--gen-seed") genSeed = std::strtoull(next().c_str(), nullptr, 10);
        else if (a == "--gen-files") genFiles = std::atoi(next().c_str());
        else if (a == "--tools") tools = splitList(next());
        else if (a == "--queries") onlyQueries = splitList(next());
    }
    for (const auto& t : tools)
        if (!isKnownTool(t)) { std::fprintf(stderr, "unknown tool '%s' (known: MR, MR+utf8, MR+prefilter, MR+dock, MR+all, NppNew, NppOld)\n", t.c_str()); return 2; }
    if (rounds < 1) { std::fprintf(stderr, "--rounds needs at least 1\n"); return 2; }
    if (!genDir.empty()) {
        generateMixedCorpus(genDir, genSeed, genFiles);
        std::printf("generated %d files in %s (seed %llu)\n", genFiles, genDir.c_str(), static_cast<unsigned long long>(genSeed));
        if (corpusArgs.empty()) return 0;
    }
    std::filesystem::create_directories(out);

    std::vector<Corpus> corpora;
    for (auto& [kind, dir] : corpusArgs) corpora.push_back(scanCorpus(kind, dir, kind == "source-text"));
    std::vector<std::vector<Query>> queries;
    for (auto& c : corpora) {
        std::vector<Query> qs;
        for (const Query& q : queriesFor(c.name))
            if (onlyQueries.empty() || std::find(onlyQueries.begin(), onlyQueries.end(), q.name) != onlyQueries.end()) qs.push_back(q);
        queries.push_back(qs);
    }

    std::vector<Cell> cells;
    for (size_t ci = 0; ci < corpora.size(); ++ci)
        for (const Query& q : queries[ci])
            for (const auto& t : tools) cells.push_back({ &corpora[ci], &q, t });

    std::fprintf(stderr, "warm-up: %zu cells\n", cells.size());
    for (const Cell& c : cells) (void)runCell(c);

    std::vector<Sample> samples;
    std::mt19937_64 order(seed);
    for (int r = 0; r < rounds; ++r) {
        std::vector<size_t> idx(cells.size()); std::iota(idx.begin(), idx.end(), 0);
        std::shuffle(idx.begin(), idx.end(), order);
        for (size_t k : idx) samples.push_back({ cells[k].corpus->name, cells[k].query->name, cells[k].tool, r, runCell(cells[k]) });
        std::fprintf(stderr, "round %d/%d done\n", r + 1, rounds);
    }

    {
        std::ofstream csv(out + "/raw.csv");
        csv << "corpus,query,tool,round,total_ms";
        for (int p = 0; p < PH_COUNT; ++p) csv << ',' << kPhaseNames[p] << "_ms";
        csv << ",hits,files_with_hits,binary_hits,text_bytes\n";
        for (const auto& s : samples) {
            csv << s.corpus << ",\"" << s.query << "\"," << s.tool << ',' << s.round << ',' << s.r.total;
            for (int p = 0; p < PH_COUNT; ++p) csv << ',' << s.r.phase[p];
            csv << ',' << s.r.hits << ',' << s.r.filesWithHits << ',' << s.r.binaryHits << ',' << s.r.textBytesLoaded << '\n';
        }
    }

    FILE* md = std::fopen((out + "/summary.md").c_str(), "w");
    g_summary = md;
    auto pick = [&](const std::string& corpus, const std::string& query, const std::string& tool) {
        std::vector<const Sample*> v(static_cast<size_t>(rounds), nullptr);
        for (const auto& s : samples) if (s.corpus == corpus && s.query == query && s.tool == tool) v[static_cast<size_t>(s.round)] = &s;
        return v;
    };

    both("# Find in Files test bench: results\n\n");
    both("- Machine: %s, %u hardware threads, %s\n", cpuModel().c_str(), std::thread::hardware_concurrency(), kernelName().c_str());
    both("- Build: g++ %s, %s\n", __VERSION__, BENCH_BUILD_MODE);
    both("- Date: %s\n", utcNow().c_str());
    both("- Rounds: %d, cells per round: %zu, cell order shuffled per round (seed %llu), one warm-up pass per cell\n", rounds, cells.size(), static_cast<unsigned long long>(seed));
    both("- Tools: ");
    for (size_t t = 0; t < tools.size(); ++t) both("%s%s", t ? ", " : "", tools[t].c_str());
    both("\n");
    for (const auto& c : corpora)
        both("- Corpus `%s`: %zu files, %.1f MB, %zu binary by MR's rule, fingerprint %016llx\n", c.name.c_str(), c.files.size(), c.bytes / 1048576.0, c.binaryFiles, static_cast<unsigned long long>(c.fingerprint));
    both("\nTimes are wall-clock ms per complete search over the corpus, warm file cache.\n");
    both("Ratio A / B below 1 means A is faster. A difference counts as reliable when the 95 %% interval excludes 1.\n\n");

    for (size_t ci = 0; ci < corpora.size(); ++ci) {
        const std::string& cn = corpora[ci].name;
        both("## Corpus `%s`\n\n", cn.c_str());
        both("| Query | Tool | median ms | mean ms | 95 %% CI of mean | CV %% | min | max | hits | hits in binaries |\n|---|---|---:|---:|---|---:|---:|---:|---:|---:|\n");
        for (const Query& q : queries[ci]) {
            std::set<size_t> textHits;
            for (const auto& t : tools) {
                std::vector<double> times;
                std::set<std::pair<size_t, size_t>> hitSets;
                for (const Sample* s : pick(cn, q.name, t)) if (s) { times.push_back(s->r.total); hitSets.insert({ s->r.hits - s->r.binaryHits, s->r.binaryHits }); }
                const Summary sm = summarize(times);
                const auto h = hitSets.empty() ? std::pair<size_t, size_t>{ 0, 0 } : *hitSets.begin();
                textHits.insert(h.first);
                both("| %s | %s | %.1f | %.1f | %.1f to %.1f | %.1f | %.1f | %.1f | %zu%s | %zu |\n", q.name.c_str(), t.c_str(), sm.median, sm.mean, sm.ciLo, sm.ciHi, sm.cv, sm.min, sm.max,
                    h.first, hitSets.size() > 1 ? " (varies between rounds)" : "", h.second);
            }
            if (textHits.size() > 1) both("| %s | **hit counts in text files differ between tools** | | | | | | | | |\n", q.name.c_str());
        }

        both("\n### Paired ratios per round: median and 95 %% bootstrap interval\n\n| Query | A / B | median | 95 %% interval | verdict |\n|---|---|---:|---|---|\n");
        uint64_t pairSeed = seed;
        for (const Query& q : queries[ci])
            for (size_t a = 0; a < tools.size(); ++a)
                for (size_t b = a + 1; b < tools.size(); ++b) {
                    std::vector<double> ta, tb;
                    const auto sa = pick(cn, q.name, tools[a]), sb = pick(cn, q.name, tools[b]);
                    for (int r = 0; r < rounds; ++r) if (sa[static_cast<size_t>(r)] && sb[static_cast<size_t>(r)]) { ta.push_back(sa[static_cast<size_t>(r)]->r.total); tb.push_back(sb[static_cast<size_t>(r)]->r.total); }
                    double m = 0, lo = 0, hi = 0;
                    pairedRatio(ta, tb, ++pairSeed, m, lo, hi);
                    const std::string verdict = hi < 1.0 ? tools[a] + " faster" : lo > 1.0 ? tools[b] + " faster" : "no reliable difference";
                    both("| %s | %s / %s | %.2f | %.2f to %.2f | %s |\n", q.name.c_str(), tools[a].c_str(), tools[b].c_str(), m, lo, hi, verdict.c_str());
                }

        both("\n### Phases, median ms per run\n\n");
        both("Notepad++ builds the result lines inside its search loop, so its `search` includes them and its `results` stays 0.\n\n| Query | Tool |");
        for (int p = 0; p < PH_COUNT; ++p) both(" %s |", kPhaseNames[p]);
        both(" search+results |\n|---|---|");
        for (int p = 0; p <= PH_COUNT; ++p) both("---:|");
        both("\n");
        for (const Query& q : queries[ci])
            for (const auto& t : tools) {
                both("| %s | %s |", q.name.c_str(), t.c_str());
                std::vector<double> sr;
                const auto runs = pick(cn, q.name, t);
                for (int p = 0; p < PH_COUNT; ++p) {
                    std::vector<double> v;
                    for (const Sample* s : runs) if (s) v.push_back(s->r.phase[p]);
                    both(" %.1f |", medianOf(v));
                }
                for (const Sample* s : runs) if (s) sr.push_back(s->r.phase[PH_SEARCH] + s->r.phase[PH_RESULT]);
                both(" %.1f |\n", medianOf(sr));
            }
        both("\n");
    }
    if (md) std::fclose(md);
    g_summary = nullptr;
    return 0;
}
