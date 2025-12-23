// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
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

#include "Encoding.h"
#include <algorithm>
#include <cstring>
#include <windows.h>

namespace Encoding {

    // ---------- internal helpers ----------

    static inline bool isMostlyAscii(const char* data, int len, double threshold) {
        if (len <= 0) return true;
        int ascii = 0;
        const unsigned char* p = reinterpret_cast<const unsigned char*>(data);
        for (int i = 0; i < len; ++i) if (p[i] < 0x80) ++ascii;
        return (static_cast<double>(ascii) / (len ? len : 1)) >= threshold;
    }

    static inline void pickSamples(const char* base, int len,
        const char*& s1, int& n1,
        const char*& s2, int& n2,
        int maxKB)
    {
        const int cap = maxKB * 1024;
        s1 = base;
        n1 = (len > cap) ? cap : len;
        s2 = nullptr;
        n2 = 0;

        // Optional middle slice for very large files
        if (len > cap * 3) {
            s2 = base + (len / 2);
            n2 = (std::min)(cap, len - (len / 2));
        }
    }

    // ---------- UTF-8 validation ----------

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

    // ---------- roundtrip check ----------

    bool roundtripLossless(const char* data, int len, UINT cp) {
        if (!data || len <= 0) return true;

        // MBCS -> UTF-16
        DWORD err = 0;
        int wlen = MultiByteToWideChar(cp, MB_ERR_INVALID_CHARS, data, len, nullptr, 0);
        if (wlen == 0) {
            err = GetLastError();
            if (err == ERROR_INVALID_FLAGS) {
                wlen = MultiByteToWideChar(cp, 0, data, len, nullptr, 0);
                if (wlen == 0) return false;
            }
            else {
                return false;
            }
        }

        std::vector<wchar_t> wbuf(static_cast<size_t>(wlen));
        if (MultiByteToWideChar(cp, 0, data, len, wbuf.data(), wlen) != wlen)
            return false;

        // UTF-16 -> MBCS (no best-fit; detect default-char usage)
        BOOL usedDefault = FALSE;
        int mlen = WideCharToMultiByte(cp, WC_NO_BEST_FIT_CHARS,
            wbuf.data(), wlen, nullptr, 0,
            nullptr, &usedDefault);
        if (mlen <= 0 || usedDefault) return false;

        std::vector<char> back(static_cast<size_t>(mlen));
        if (WideCharToMultiByte(cp, WC_NO_BEST_FIT_CHARS,
            wbuf.data(), wlen, back.data(), mlen,
            nullptr, &usedDefault) != mlen)
            return false;
        if (usedDefault) return false;
        if (mlen != len) return false;
        return std::memcmp(back.data(), data, len) == 0;
    }

    // Quick DBCS plausibility: count lead+trail pairs for a given code page.
    static int dbcsPairScore(const char* data, int len, UINT cp) {
        int score = 0;
        for (int i = 0; i < len; ++i) {
            unsigned char b = static_cast<unsigned char>(data[i]);
            if (IsDBCSLeadByteEx(cp, b)) {
                if (i + 1 < len) { ++score; ++i; }
            }
        }
        return score;
    }


    UINT autoDetectAnsiCodepage(const char* data, int len, UINT acp, const DetectOptions& opt) {
        if (!data || len <= 0) return acp;

        // 0) ASCII fast path
        if (isMostlyAscii(data, len, opt.asciiQuickPathThreshold))
            return acp;

        const char* s1; int n1; const char* s2; int n2;
        pickSamples(data, len, s1, n1, s2, n2, static_cast<int>(opt.sampleKB));

        // 1) Probe CJK first ONLY if we see plausible DBCS structure
        static const UINT cjk[] = { 932u, 936u, 949u, 950u };
        constexpr int MIN_DBCS_PAIRS = 3;

        for (UINT cp : cjk) {
            int score1 = dbcsPairScore(s1, n1, cp);
            int score2 = (!s2 || n2 <= 0) ? 0 : dbcsPairScore(s2, n2, cp);
            if (score1 + score2 >= MIN_DBCS_PAIRS) {
                bool ok1 = roundtripLossless(s1, n1, cp);
                bool ok2 = (!s2 || n2 <= 0) ? true : roundtripLossless(s2, n2, cp);
                if (ok1 && ok2) return cp;
            }
        }

        // 2) Optional extras (only if caller set them)
        for (UINT cp : opt.extraAnsiCandidates) {
            if (cp == acp) continue;
            bool ok1 = roundtripLossless(s1, n1, cp);
            bool ok2 = (!s2 || n2 <= 0) ? true : roundtripLossless(s2, n2, cp);
            if (ok1 && ok2) return cp;
        }

        // 3) Fallback: system ACP
        return acp;
    }

    EncodingInfo detectEncoding(const char* data, size_t len, const DetectOptions& opt) {
        EncodingInfo ei;

        if (!data || len == 0) {
            ei.kind = Kind::ANSI;
            ei.codepage = GetACP();
            return ei;
        }

        const unsigned char* p = reinterpret_cast<const unsigned char*>(data);

        // BOM checks
        if (len >= 3 && p[0] == 0xEF && p[1] == 0xBB && p[2] == 0xBF) {
            ei.kind = Kind::UTF8;
            ei.withBOM = true; ei.bomBytes = 3;
            return ei;
        }
        if (len >= 2 && p[0] == 0xFF && p[1] == 0xFE) {
            ei.kind = Kind::UTF16LE;
            ei.withBOM = true; ei.bomBytes = 2;
            return ei;
        }
        if (len >= 2 && p[0] == 0xFE && p[1] == 0xFF) {
            ei.kind = Kind::UTF16BE;
            ei.withBOM = true; ei.bomBytes = 2;
            return ei;
        }

        // UTF-8 (no BOM)
        if (opt.preferUtf8NoBOM && isValidUtf8(data, len)) {
            ei.kind = Kind::UTF8;
            ei.withBOM = false; ei.bomBytes = 0;
            return ei;
        }

        // ANSI fallback
        ei.kind = Kind::ANSI;
        ei.codepage = GetACP();
        ei.withBOM = false; ei.bomBytes = 0;

        if (opt.enableAutoCJK) {
            UINT cp = autoDetectAnsiCodepage(data, static_cast<int>(len), ei.codepage, opt);
            ei.codepage = cp;
        }
        return ei;
    }


    // ---------- string conversions ----------

    std::wstring bytesToWString(const std::string& bytes, UINT cp) {
        if (bytes.empty()) return std::wstring();
        int wlen = MultiByteToWideChar(cp, 0, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
        if (wlen <= 0) return std::wstring();
        std::wstring w(static_cast<size_t>(wlen), L'\0');
        MultiByteToWideChar(cp, 0, bytes.data(), static_cast<int>(bytes.size()), w.data(), wlen);
        return w;
    }

    // FIX: Restored to v4.4 behavior - always use permissive mode (flags = 0).
    // The strict WC_NO_BEST_FIT_CHARS + usedDefault check caused search failures
    // on certain codepages (CP866, CP1251, etc.) where even simple ASCII failed.
    std::string wstringToBytes(const std::wstring& w, UINT cp) {
        if (w.empty()) return std::string();
        int mlen = WideCharToMultiByte(cp, 0, w.data(), static_cast<int>(w.size()),
            nullptr, 0, nullptr, nullptr);
        if (mlen <= 0) return std::string();
        std::string out(static_cast<size_t>(mlen), '\0');
        WideCharToMultiByte(cp, 0, w.data(), static_cast<int>(w.size()),
            out.data(), mlen, nullptr, nullptr);
        return out;
    }

    std::string bytesToUtf8(const std::string& bytes, UINT cp) {
        return wstringToUtf8(bytesToWString(bytes, cp));
    }

    std::string wstringToUtf8(const std::wstring& w) {
        if (w.empty()) return std::string();
        int mlen = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()),
            nullptr, 0, nullptr, nullptr);
        if (mlen <= 0) return std::string();
        std::string out(static_cast<size_t>(mlen), '\0');
        WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()),
            out.data(), mlen, nullptr, nullptr);
        return out;
    }

    std::wstring utf8ToWString(const std::string& u8) {
        if (u8.empty()) return std::wstring();
        int wlen = MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()),
            nullptr, 0);
        if (wlen <= 0) return std::wstring();
        std::wstring w(static_cast<size_t>(wlen), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()),
            w.data(), wlen);
        return w;
    }

    std::string utf8ToBytes(const std::string& u8, UINT cp) {
        return wstringToBytes(utf8ToWString(u8), cp);
    }

    // ---------- buffer conversions + BOM ----------

    static inline void appendBOM(Kind kind, std::vector<char>& out) {
        if (kind == Kind::UTF8) {
            const unsigned char b[3] = { 0xEF, 0xBB, 0xBF };
            out.insert(out.end(), reinterpret_cast<const char*>(b), reinterpret_cast<const char*>(b) + 3);
        }
        else if (kind == Kind::UTF16LE) {
            const unsigned char b[2] = { 0xFF, 0xFE };
            out.insert(out.end(), reinterpret_cast<const char*>(b), reinterpret_cast<const char*>(b) + 2);
        }
        else if (kind == Kind::UTF16BE) {
            const unsigned char b[2] = { 0xFE, 0xFF };
            out.insert(out.end(), reinterpret_cast<const char*>(b), reinterpret_cast<const char*>(b) + 2);
        }
    }

    bool convertBufferToUtf8(const std::vector<char>& in, const EncodingInfo& src, std::string& outUtf8) {
        outUtf8.clear();
        if (in.empty()) return true;

        const char* data = in.data();
        size_t len = in.size();

        // Skip BOM bytes if present
        if (src.bomBytes > 0 && static_cast<size_t>(src.bomBytes) <= len) {
            data += src.bomBytes;
            len -= src.bomBytes;
        }

        if (src.kind == Kind::UTF8) {
            outUtf8.assign(data, data + len);
            return true;
        }

        if (src.kind == Kind::UTF16LE || src.kind == Kind::UTF16BE) {
            UINT cp = (src.kind == Kind::UTF16LE) ? 1200u : 1201u; // Windows Unicode codepages
            int wlen = MultiByteToWideChar(cp, 0, data, static_cast<int>(len), nullptr, 0);
            if (wlen <= 0) return false;
            std::vector<wchar_t> w(static_cast<size_t>(wlen));
            if (MultiByteToWideChar(cp, 0, data, static_cast<int>(len), w.data(), wlen) != wlen)
                return false;
            outUtf8 = wstringToUtf8(std::wstring(w.begin(), w.end()));
            return true;
        }

        // ANSI
        outUtf8 = bytesToUtf8(std::string(data, data + len), src.codepage);
        return !outUtf8.empty() || len == 0;
    }

    bool convertUtf8ToOriginal(const std::string& u8, const EncodingInfo& dst, std::vector<char>& outBytes) {
        outBytes.clear();

        if (dst.kind == Kind::UTF8) {
            if (dst.withBOM) appendBOM(Kind::UTF8, outBytes);
            outBytes.insert(outBytes.end(), u8.begin(), u8.end());
            return true;
        }

        std::wstring w = utf8ToWString(u8);

        if (dst.kind == Kind::UTF16LE || dst.kind == Kind::UTF16BE) {
            if (dst.withBOM) appendBOM(dst.kind, outBytes);
            UINT cp = (dst.kind == Kind::UTF16LE) ? 1200u : 1201u;
            int blen = WideCharToMultiByte(cp, 0, w.data(), static_cast<int>(w.size()), nullptr, 0, nullptr, nullptr);
            if (blen <= 0) return false;
            outBytes.resize(static_cast<size_t>(blen));
            if (WideCharToMultiByte(cp, 0, w.data(), static_cast<int>(w.size()), outBytes.data(), blen, nullptr, nullptr) != blen)
                return false;
            return true;
        }

        // ANSI
        {
            std::string mbs = wstringToBytes(w, dst.codepage);
            if (mbs.empty() && !w.empty()) return false;
            // ANSI doesn't use BOM; ignore dst.withBOM
            outBytes.insert(outBytes.end(), mbs.begin(), mbs.end());
            return true;
        }
    }

} // namespace Encoding
