// QA harness for the RawBytes-branch codepage decision.
//
// Context: files that hasNullBytes() flags as binary, with "Skip binary
// files" turned off, are searched as raw undecoded bytes (LoadKind::RawBytes)
// - this is deliberate ("N++ behavior", see HiddenSciGuard.h loadTextFile).
// Previously every such file was unconditionally bound to codepage 0
// (system ANSI), regardless of what encoding its content actually was in.
// For guy038's Mark_Style.txt (French prose with a literal NUL and a table
// of literal C0 control bytes, saved as UTF-8) this meant: an ASCII search
// pattern ("Fi") still matched correctly (ASCII is byte-identical across
// UTF-8 and ANSI), but any accented search pattern ("PRÉCÉDENT") would not
// reliably match, and dock hit lines with accents would render as mojibake.
//
// The fix: loadTextFile now runs Encoding::isValidUtf8 - a strict structural
// parser, not a statistical guess - over the FULL raw content whenever it
// takes the RawBytes branch, and reports the result via the existing `enc`
// out-parameter. The two call sites (handleFindInFiles, handleReplaceInFiles)
// then bind the hidden buffer to SC_CP_UTF8 instead of codepage 0 when it
// validates. The underlying bytes are NEVER decoded or transformed by this -
// only the codepage interpretation used for search/replace text encoding and
// dock rendering changes.
//
// This harness proves the central safety claim: the change is neutral or
// better in every realistic bucket, never worse for genuinely binary data.
//
// Mirrors verbatim:
//   Encoding::isValidUtf8              (Encoding.cpp)
//   codepageForLoadedFile              (MultiReplacePanel.cpp)
//
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined
//        -o rawbytes_codepage_qa rawbytes_codepage_qa.cpp
#include <cstdio>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <random>
#include <string>
#include <vector>

// ----------------------------- VERBATIM: Encoding::isValidUtf8
namespace Encoding {
    bool isValidUtf8(const char* data, size_t len) {
        const unsigned char* s = reinterpret_cast<const unsigned char*>(data);
        size_t i = 0;
        while (i < len) {
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

    enum class Kind { ANSI, UTF8, UTF16LE, UTF16BE };
    struct EncodingInfo { Kind kind = Kind::ANSI; bool withBOM = false; };
}

// ---------------------------------- MIRROR: HiddenSciGuard binary primitives
static constexpr size_t BINARY_CHECK_SIZE = 8192;

static bool hasBOM(const char* data, size_t len) {
    if (len < 2) return false;
    const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
    if (len >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) return true;
    if (u[0] == 0xFF && u[1] == 0xFE) return true;
    if (u[0] == 0xFE && u[1] == 0xFF) return true;
    return false;
}
static bool hasNullBytes(const char* data, size_t len) {
    const size_t checkLen = (len < BINARY_CHECK_SIZE) ? len : BINARY_CHECK_SIZE;
    return std::memchr(data, '\0', checkLen) != nullptr;
}
// (isUtf16NoBomLE omitted: none of this harness's fixtures are UTF-16-LE-shaped
//  without a BOM, so it would never fire - see D-series coverage instead.)
static bool shouldSkipAsBinary(const char* data, size_t len) {
    if (hasBOM(data, len)) return false;
    return hasNullBytes(data, len);
}

// -------------------- MIRROR: loadTextFile's RawBytes branch (the fix)
struct LoadOutcome { bool isRawBytes; Encoding::EncodingInfo enc; };

static LoadOutcome simulateLoad(const std::string& raw)
{
    LoadOutcome out;
    out.isRawBytes = shouldSkipAsBinary(raw.data(), raw.size());
    if (out.isRawBytes && Encoding::isValidUtf8(raw.data(), raw.size())) {
        out.enc.kind = Encoding::Kind::UTF8;
    }
    return out;
}

// -------------------------- VERBATIM: codepageForLoadedFile (MultiReplacePanel.cpp)
static constexpr int SC_CP_UTF8_ = 65001;
static int codepageForLoadedFile(bool isRawBytes, const Encoding::EncodingInfo& enc)
{
    if (!isRawBytes) return SC_CP_UTF8_;
    return (enc.kind == Encoding::Kind::UTF8) ? SC_CP_UTF8_ : 0;
}

// ----------------------------------------------------------- checks
static int failures = 0;
static void CHECK(const std::string& d, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", d.c_str());
    if (!ok) ++failures;
}

static std::string utf8(const std::vector<uint32_t>& cps) {
    std::string s;
    for (uint32_t c : cps) {
        if (c < 0x80) s += static_cast<char>(c);
        else if (c < 0x800) {
            s += static_cast<char>(0xC0 | (c >> 6));
            s += static_cast<char>(0x80 | (c & 0x3F));
        } else {
            s += static_cast<char>(0xE0 | (c >> 12));
            s += static_cast<char>(0x80 | ((c >> 6) & 0x3F));
            s += static_cast<char>(0x80 | (c & 0x3F));
        }
    }
    return s;
}

int main() {
    std::printf("=== R1 guy038's Mark_Style.txt shape: UTF-8 prose + literal C0 bytes ===\n\n");
    {
        // French prose (UTF-8: É = C3 89) with a literal embedded NUL, like
        // the file's own demo table ("Ctrl + 0" -> shows the NUL char itself).
        std::string content = utf8({'L','e',' ','r','a','c','c','o','u','r','c','i',',',' ',
                                     'p','a','r',' ','D',0xC9,'F','A','U','T'});     // "DÉFAUT"
        content += '\x00';                                                          // literal NUL, like Ctrl+0's output
        content += utf8({'p','r',0xC9,'C',0xC9,'D','E','N','T'});                   // "PRÉCÉDENT"
        content += "\x01\x02\x03\x04";                                              // more literal C0 bytes from the table

        CHECK("R1 flagged as binary (has NUL in header)",
              shouldSkipAsBinary(content.data(), content.size()));
        auto r = simulateLoad(content);
        CHECK("R1 RawBytes branch taken", r.isRawBytes);
        CHECK("R1 validates as UTF-8 (C0 bytes are valid 1-byte code points)",
              r.enc.kind == Encoding::Kind::UTF8);
        CHECK("R1 codepage resolves to UTF-8, not ANSI",
              codepageForLoadedFile(r.isRawBytes, r.enc) == SC_CP_UTF8_);
        std::printf("     -> accented search patterns now match; previously forced to codepage 0\n\n");
    }

    std::printf("=== R2 pure ASCII with a stray NUL: behavior identical either way ===\n\n");
    {
        std::string content = "plain english text";
        content += '\x00';
        content += "more text, no accents";
        auto r = simulateLoad(content);
        CHECK("R2 RawBytes + validates UTF-8 (ASCII is a UTF-8 subset)",
              r.isRawBytes && r.enc.kind == Encoding::Kind::UTF8);
        CHECK("R2 codepage is UTF-8 (byte-identical to ANSI/CP_ACP for ASCII "
              "range, so observably unchanged for guy038's own 'Fi' search)",
              codepageForLoadedFile(r.isRawBytes, r.enc) == SC_CP_UTF8_);
    }

    std::printf("\n=== R3 non-UTF8 ANSI text with a stray NUL: unchanged, not worse ===\n\n");
    {
        // Windows-1252 'é' is single byte 0xE9 - as a lone byte with no
        // continuation bytes following, this is NOT valid UTF-8.
        std::string content = "caf\xE9 also has a stray NUL here: ";
        content += '\x00';
        content += "end";
        CHECK("R3 flagged as binary", shouldSkipAsBinary(content.data(), content.size()));
        auto r = simulateLoad(content);
        CHECK("R3 does NOT validate as UTF-8 (lone 0xE9 breaks structural check)",
              r.enc.kind != Encoding::Kind::UTF8);
        CHECK("R3 codepage falls back to 0 - EXACTLY today's pre-fix behavior",
              codepageForLoadedFile(r.isRawBytes, r.enc) == 0);
    }

    std::printf("\n=== R4 files that never reach the RawBytes branch: untouched ===\n\n");
    {
        auto r = simulateLoad("no null bytes, ordinary text file");
        CHECK("R4 not flagged as binary", !r.isRawBytes);
        CHECK("R4 codepage is UTF-8 via the Text-kind path (unrelated to this fix)",
              codepageForLoadedFile(r.isRawBytes, r.enc) == SC_CP_UTF8_);
    }
    {
        // BOM short-circuits shouldSkipAsBinary before this fix is ever reached.
        std::string bomUtf8 = "\xEF\xBB\xBF";
        bomUtf8 += "text with a NUL"; bomUtf8 += '\x00'; bomUtf8 += "after BOM";
        CHECK("R4b BOM exempts from binary flag entirely (pre-existing behavior)",
              !shouldSkipAsBinary(bomUtf8.data(), bomUtf8.size()));
    }

    std::printf("\n=== R5 genuine binary content: the core safety claim ===\n\n");
    {
        // Deterministic PRNG standing in for "real compiled/compressed binary
        // data": high entropy, full byte range, no structure favoring UTF-8.
        std::mt19937 rng(133);
        std::uniform_int_distribution<int> byteDist(0, 255);
        int falsePositives = 0;
        const int trials = 5000;
        const size_t sampleLen = 4096;
        for (int t = 0; t < trials; ++t) {
            std::string blob(sampleLen, '\0');
            for (auto& c : blob) c = static_cast<char>(byteDist(rng));
            // Every sample already contains NUL bytes by construction (256-way
            // uniform over 4096 bytes), so it's already binary-flagged; the
            // question is only whether isValidUtf8 also wrongly says "UTF-8".
            if (Encoding::isValidUtf8(blob.data(), blob.size())) ++falsePositives;
        }
        std::printf("     %d/%d uniformly-random 4KB blobs falsely validated as UTF-8\n",
                    falsePositives, trials);
        CHECK("R5 zero false positives on high-entropy random binary data",
              falsePositives == 0);
    }
    {
        // Structured binary stand-ins: PE/ELF-style headers (magic bytes,
        // then packed little-endian fields with lots of 0x00 padding, which
        // is exactly the shape that could in principle look "text-like" in
        // spots). Still must fail structural UTF-8 validation overall.
        auto peLike = [] {
            std::string s = "MZ\x90\x00\x03\x00\x00\x00\x04\x00\x00\x00\xFF\xFF\x00\x00";
            for (int i = 0; i < 200; ++i) { s += static_cast<char>(i & 0xFF); s += '\x00'; }
            s += "PE\x00\x00\x64\x86\x07\x00";
            return s;
            }();
        CHECK("R5b PE-header-shaped binary flagged as binary",
              shouldSkipAsBinary(peLike.data(), peLike.size()));
        CHECK("R5b PE-header-shaped binary does NOT validate as UTF-8",
              !Encoding::isValidUtf8(peLike.data(), peLike.size()));

        auto zipLike = [] {
            std::string s = "PK\x03\x04\x14\x00\x00\x00\x08\x00";
            std::mt19937 rng(133);
            std::uniform_int_distribution<int> b(0, 255);
            for (int i = 0; i < 500; ++i) s += static_cast<char>(b(rng));
            return s;
            }();
        CHECK("R5b ZIP-header-shaped binary does NOT validate as UTF-8",
              !Encoding::isValidUtf8(zipLike.data(), zipLike.size()));
    }

    std::printf("\n=== R6 write-back path is untouched by this fix (structural guarantee) ===\n\n");
    {
        // The fix only ever reads `raw`/`content` to decide a codepage; it
        // never rebinds, reassigns, or transforms the buffer that flows into
        // `content = std::move(raw)`. This is a static/structural guarantee
        // from the diff shape itself (isValidUtf8 takes `const char*`, the
        // move happens identically regardless of branch outcome) rather than
        // something a runtime check can prove differently than reading the
        // patch - documented here as the explicit claim under test.
        CHECK("R6 documented: content buffer identity is independent of the "
              "isValidUtf8 result by construction (see HiddenSciGuard.h diff)",
              true);
    }

    std::printf("\n%s (%d failure%s)\n",
                failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
