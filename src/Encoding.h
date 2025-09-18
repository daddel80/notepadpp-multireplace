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

#pragma once
// Encoding.h
// Robust Windows encoding utilities for MultiReplace.
// - BOM & charset detection (UTF-8/UTF-16LE/UTF-16BE/ANSI)
// - Optional Auto-CJK detection on ANSI fallback (932/936/949/950) via lossless roundtrip
// - Optional extra ANSI candidates (e.g., 1250) per deployment
// - Safe conversions (no best-fit by default), BOM preservation

#include <string>
#include <vector>
#include <cstdint>
#include <windows.h>

namespace Encoding {

    // ---------- Kinds & options ----------

    enum class Kind {
        UTF8,
        UTF16LE,
        UTF16BE,
        ANSI
    };

    struct DetectOptions {
        bool  preferUtf8NoBOM = true;       // treat valid UTF-8 (no BOM) as UTF-8
        bool  enableAutoCJK = true;      // probe CJK codepages on ANSI fallback
        std::vector<UINT> extraAnsiCandidates{}; // user-provided extra candidates (e.g., 1250)
        double asciiQuickPathThreshold = 0.98;   // skip probing if mostly ASCII
        size_t sampleKB = 128;              // size per sample for probing
    };

    struct EncodingInfo {
        Kind  kind = Kind::ANSI;
        UINT  codepage = CP_ACP;  // used when kind==ANSI
        bool  withBOM = false;   // preserve BOM on write
        int   bomBytes = 0;       // 0, 2, or 3
    };

    struct ConvertOptions {
        bool allowBestFit = false; // default strict: no best-fit
    };

    // ---------- Validation ----------
    bool isValidUtf8(const char* data, size_t len);

    // ---------- Detection ----------
    EncodingInfo detectEncoding(const char* data, size_t len, const DetectOptions& opt = {});
    UINT autoDetectAnsiCodepage(const char* data, int len, UINT acp, const DetectOptions& opt);

    // ---------- Low-level helper ----------
    bool roundtripLossless(const char* data, int len, UINT cp);

    // ---------- String conversions ----------
    std::wstring bytesToWString(const std::string& bytes, UINT cp);
    std::string  wstringToBytes(const std::wstring& w, UINT cp, const ConvertOptions& copt = {});
    std::string  bytesToUtf8(const std::string& bytes, UINT cp);
    std::string  wstringToUtf8(const std::wstring& w);
    std::wstring utf8ToWString(const std::string& u8);
    std::string  utf8ToBytes(const std::string& u8, UINT cp, const ConvertOptions& copt = {});

    // ---------- Buffer conversions with BOM handling ----------
    bool convertBufferToUtf8(const std::vector<char>& in, const EncodingInfo& src, std::string& outUtf8);
    bool convertUtf8ToOriginal(const std::string& u8, const EncodingInfo& dst, std::vector<char>& outBytes, const ConvertOptions& copt = {});

    std::wstring trim(const std::wstring& str);

} // namespace Encoding
