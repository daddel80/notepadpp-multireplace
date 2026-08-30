// QA harness for the MultiReplace file-search loading pipeline (post-fix state).
// Functions marked VERBATIM mirror the shipped logic 1:1 (HiddenSciGuard.h,
// Encoding.cpp) with Win32 shims: IsTextUnicode is stubbed with the documented
// byte-pattern behavior for the synthetic inputs used here; converters carry
// no claims beyond branch decisions.
//
// Build: g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -o file_search_qa file_search_qa.cpp
#include <cstddef>
#include <cstring>
#include <cstdint>
#include <climits>
#include <string>
#include <vector>
#include <cstdio>

using UINT = unsigned int;
static UINT GetACP() { return 1252; }

// Shim: statistics probe. Our synthetic UTF-16LE ASCII text has a NUL in every
// odd position, which IS_TEXT_UNICODE_STATISTICS accepts; random binary does not.
static bool IsTextUnicode_stub(const char* data, int len) {
    if (len < 2) return false;
    int odd_nul = 0, pairs = len / 2;
    for (int i = 1; i < len; i += 2) if (data[i] == 0) ++odd_nul;
    return pairs > 0 && odd_nul >= (pairs * 3) / 4;
}

// ============================================================
// VERBATIM: HiddenSciGuard section 3 (post-fix)
// ============================================================
static constexpr size_t BINARY_CHECK_SIZE = 8192;

bool hasBOM(const char* data, size_t len)
{
    if (len < 2) return false;
    const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
    if (len >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF) return true;
    if (u[0] == 0xFF && u[1] == 0xFE) return true;
    if (u[0] == 0xFE && u[1] == 0xFF) return true;
    return false;
}

bool hasNullBytes(const char* data, size_t len)
{
    const size_t checkLen = (len < BINARY_CHECK_SIZE) ? len : BINARY_CHECK_SIZE;
    return std::memchr(data, '\0', checkLen) != nullptr;
}

bool isUtf16NoBomLE(const char* data, size_t len)
{
    const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
    if (!(len > 1 && (len % 2) == 0 && u[0] != 0 && u[1] == 0)) return false;
    return IsTextUnicode_stub(data, static_cast<int>(len));
}

bool shouldSkipAsBinary(const char* data, size_t len)
{
    if (hasBOM(data, len)) return false;
    if (isUtf16NoBomLE(data, len)) return false;
    return hasNullBytes(data, len);
}

// ============================================================
// VERBATIM branch logic: Encoding::detectEncoding (post-fix)
// ============================================================
namespace Encoding {
    enum class Kind { UTF8, UTF16LE, UTF16BE, ANSI };
    struct DetectOptions {
        bool preferUtf8NoBOM = true;
        bool enableUtf16NoBomLE = true;
        bool enableAutoCJK = true;
    };
    struct EncodingInfo {
        Kind kind = Kind::ANSI;
        UINT codepage = 0;
        bool withBOM = false;
        int  bomBytes = 0;
    };

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

    EncodingInfo detectEncoding(const char* data, size_t len, const DetectOptions& opt = {}) {
        EncodingInfo ei;
        if (!data || len == 0) { ei.kind = Kind::ANSI; ei.codepage = GetACP(); return ei; }
        const unsigned char* p = reinterpret_cast<const unsigned char*>(data);

        if (len >= 3 && p[0] == 0xEF && p[1] == 0xBB && p[2] == 0xBF) { ei.kind = Kind::UTF8; ei.withBOM = true; ei.bomBytes = 3; return ei; }
        if (len >= 2 && p[0] == 0xFF && p[1] == 0xFE) { ei.kind = Kind::UTF16LE; ei.withBOM = true; ei.bomBytes = 2; return ei; }
        if (len >= 2 && p[0] == 0xFE && p[1] == 0xFF) { ei.kind = Kind::UTF16BE; ei.withBOM = true; ei.bomBytes = 2; return ei; }

        if (opt.enableUtf16NoBomLE && len > 1 && (len % 2) == 0 && p[0] != 0 && p[1] == 0) {
            const int probeLen = static_cast<int>(len < static_cast<size_t>(64 * 1024) ? len : static_cast<size_t>(64 * 1024));
            if (IsTextUnicode_stub(data, probeLen)) { ei.kind = Kind::UTF16LE; ei.withBOM = false; ei.bomBytes = 0; return ei; }
        }

        if (opt.preferUtf8NoBOM && isValidUtf8(data, len)) { ei.kind = Kind::UTF8; ei.withBOM = false; ei.bomBytes = 0; return ei; }

        ei.kind = Kind::ANSI; ei.codepage = GetACP(); ei.withBOM = false; ei.bomBytes = 0;
        return ei;
    }

    // VERBATIM gates of convertBufferToUtf8 (INT_MAX + BOM skip + even-length)
    bool convertBufferToUtf8_wouldAccept(const char* data, size_t len, const EncodingInfo& src) {
        if (!data || len == 0) return true;
        if (len > static_cast<size_t>(INT_MAX)) return false;
        if (src.bomBytes > 0 && static_cast<size_t>(src.bomBytes) <= len) { data += src.bomBytes; len -= src.bomBytes; }
        if (src.kind == Kind::UTF8) return true;
        if (src.kind == Kind::UTF16LE || src.kind == Kind::UTF16BE) return (len % 2) == 0;
        return true;
    }
}

// ============================================================
// Pipeline decision replica: HiddenSciGuard::loadTextFile (post-fix)
// ============================================================
enum class LoadKind { Text, RawBytes };
enum class SkipReason { None, Binary, TooLarge, Unreadable, Undecodable };

struct PipelineResult {
    SkipReason reason = SkipReason::None;
    LoadKind kind = LoadKind::Text;
    Encoding::Kind encKind = Encoding::Kind::ANSI;
};

struct Counters { size_t binary = 0, tooLarge = 0, unreadable = 0, undecodable = 0; };

PipelineResult loadTextFile_decide(const std::string& fileBytes, bool skipBinaryFiles, Counters& c)
{
    PipelineResult r;
    const size_t headerSize = (fileBytes.size() < BINARY_CHECK_SIZE) ? fileBytes.size() : BINARY_CHECK_SIZE;

    const bool binary = shouldSkipAsBinary(fileBytes.data(), headerSize);
    if (binary && skipBinaryFiles) { ++c.binary; r.reason = SkipReason::Binary; return r; }

    if (binary) { r.kind = LoadKind::RawBytes; return r; }

    const auto enc = Encoding::detectEncoding(fileBytes.data(), fileBytes.size());
    r.encKind = enc.kind;
    if (!Encoding::convertBufferToUtf8_wouldAccept(fileBytes.data(), fileBytes.size(), enc)) {
        ++c.undecodable; r.reason = SkipReason::Undecodable; return r;
    }
    return r;
}

// ---- test helpers (no claims) ----
static std::string utf16le(const std::string& a, bool bom) {
    std::string o; if (bom) { o += (char)0xFF; o += (char)0xFE; }
    for (char ch : a) { o += ch; o += '\0'; } return o;
}
static std::string utf16be(const std::string& a, bool bom) {
    std::string o; if (bom) { o += (char)0xFE; o += (char)0xFF; }
    for (char ch : a) { o += '\0'; o += ch; } return o;
}

static int failures = 0;
#define CHECK(desc, cond) do { bool ok=(cond); std::printf("%-4s %s\n", ok?"PASS":"FAIL", desc); if(!ok) ++failures; } while(0)

int main() {
    const std::string text = "First Fight\nsecond line Fi\nthird Fi line\n";

    // ---------- F1 fixed: UTF-16 with BOM is decoded, never raw ----------
    {
        Counters c{};
        auto r = loadTextFile_decide(utf16le(text, true), true, c);
        CHECK("F1' UTF-16LE-BOM -> Text/UTF16LE (decoded, no raw path)",
              r.reason == SkipReason::None && r.kind == LoadKind::Text && r.encKind == Encoding::Kind::UTF16LE);
        auto rb = loadTextFile_decide(utf16be(text, true), true, c);
        CHECK("F1' UTF-16BE-BOM -> Text/UTF16BE",
              rb.reason == SkipReason::None && rb.kind == LoadKind::Text && rb.encKind == Encoding::Kind::UTF16BE);
    }

    // ---------- F3 fixed: UTF-16 LE without BOM is detected; BE stays binary (N++ parity) ----------
    {
        Counters c{};
        auto r = loadTextFile_decide(utf16le(text, false), true, c);
        CHECK("F3' UTF-16LE no BOM -> Text/UTF16LE (probe passes)",
              r.reason == SkipReason::None && r.kind == LoadKind::Text && r.encKind == Encoding::Kind::UTF16LE);
        auto rb = loadTextFile_decide(utf16be(text, false), true, c);
        CHECK("F3' UTF-16BE no BOM -> Binary skip (N++ parity: not detected)",
              rb.reason == SkipReason::Binary && c.binary == 1);
    }

    // ---------- F4 fixed: odd-length UTF-16 is counted as undecodable ----------
    {
        Counters c{};
        std::string odd = utf16le(text, true); odd += 'X';
        auto r = loadTextFile_decide(odd, true, c);
        CHECK("F4' odd-length UTF-16 -> Undecodable, counter incremented",
              r.reason == SkipReason::Undecodable && c.undecodable == 1);
    }

    // ---------- Binary option matrix ----------
    {
        const std::string bin = std::string("\x4D\x5A\x00\x00", 4) + std::string(64, '\x07');
        Counters c1{};
        auto on = loadTextFile_decide(bin, true, c1);
        CHECK("OPT skip ON:  early-NUL binary -> SkipBinary, counted",
              on.reason == SkipReason::Binary && c1.binary == 1);
        Counters c2{};
        auto off = loadTextFile_decide(bin, false, c2);
        CHECK("OPT skip OFF: early-NUL binary -> RawBytes (N++ behavior), nothing counted",
              off.reason == SkipReason::None && off.kind == LoadKind::RawBytes && c2.binary == 0);
    }

    // ---------- NUL beyond header: treated as text (ASCII+NUL is valid UTF-8) ----------
    {
        Counters c{};
        std::string f(9000, 'A'); f += '\0'; f += "tail Fi tail";
        auto r = loadTextFile_decide(f, true, c);
        CHECK("EDGE NUL after 8 KB header -> Text/UTF8, not skipped",
              r.reason == SkipReason::None && r.kind == LoadKind::Text && r.encKind == Encoding::Kind::UTF8);
    }

    // ---------- Healthy paths unchanged ----------
    {
        Counters c{};
        auto r1 = loadTextFile_decide(text, true, c);
        CHECK("OK  plain UTF-8/ASCII -> Text/UTF8", r1.reason == SkipReason::None && r1.encKind == Encoding::Kind::UTF8);
        std::string ansi = text + "Gr\xFC\xDF";
        auto r2 = loadTextFile_decide(ansi, true, c);
        CHECK("OK  ANSI (CP1252 bytes) -> Text/ANSI", r2.reason == SkipReason::None && r2.encKind == Encoding::Kind::ANSI);
        std::string bom8 = std::string("\xEF\xBB\xBF") + text;
        auto r3 = loadTextFile_decide(bom8, true, c);
        CHECK("OK  UTF-8 BOM -> Text/UTF8 withBOM path", r3.reason == SkipReason::None && r3.encKind == Encoding::Kind::UTF8);
    }

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED", failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
