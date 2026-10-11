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

#pragma once

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <functional>
#include <string>

#include "Encoding.h"
#include "Scintilla.h"

// Find/Replace in Files: loads one file and decides once how it is searched. Holds no state and
// counts nothing, so it runs on any thread; the caller counts the skips.
namespace FileLoad {

    // Bytes to check for binary detection (8 KB - sufficient and fast)
    constexpr size_t BINARY_CHECK_SIZE = 8192;

    // How a loaded file should be fed into the hidden buffer
    enum class LoadKind { Text, RawBytes };

    // Why a file was skipped (None = processed)
    enum class SkipReason { None, Binary, TooLarge, Unreadable, Undecodable, ReadOnly, OpenUnsaved, Unencodable };

    struct Options {
        uint64_t maxFileSize = 0;        // 0 = unlimited
        bool skipBinary = true;          // false: binary files are searched as raw bytes
        bool verifyRoundtrip = false;    // refuse decodes that would not write back losslessly
        bool skipReadOnly = false;       // Replace in Files: read-only files are left out before opening
        bool stamp = false;              // record the file's state for unchangedSince
    };

    // The file as it was opened
    struct Stamp {
        uint32_t volume = 0;
        uint64_t fileId = 0;
        uint64_t size = 0;
        uint64_t lastWrite = 0;
        bool valid = false;
    };

    struct Result {
        SkipReason skip = SkipReason::None;
        LoadKind kind = LoadKind::Text;
        Encoding::EncodingInfo enc{};
        Stamp stamp{};
    };

    // A worker's hold on a load: beforeRead gets the file size before the full read and may wait,
    // stop is checked between blocks of a large file. Either one ends the load as Unreadable.
    struct Pace {
        std::function<bool(uint64_t fileSize)> beforeRead;
        const std::atomic<bool>* stop = nullptr;
    };

    // header -> BOM/UTF-16 detection -> binary check -> text size -> full read -> decode.
    // Text: content holds UTF-8, enc describes the source encoding.
    // RawBytes: content holds the raw file bytes (binary skip disabled).
    // content keeps its capacity: a caller loading file after file passes the same string.
    Result load(const std::filesystem::path& fp, const Options& options, std::string& content, const Pace* pace = nullptr);

    // Codepage to bind the hidden buffer to. Text is always UTF-8. Raw bytes go in undecoded: UTF-8 if
    // they are valid UTF-8 (control bytes in UTF-8 text, as in guy038's Mark_Style.txt), else codepage 0.
    inline int codepageFor(LoadKind kind, const Encoding::EncodingInfo& enc)
    {
        if (kind != LoadKind::RawBytes) return SC_CP_UTF8;
        return (enc.kind == Encoding::Kind::UTF8) ? SC_CP_UTF8 : 0;
    }

    // True if the file at fp is still the one the stamp was taken of (links followed)
    bool unchangedSince(const std::filesystem::path& fp, const Stamp& stamp);

    // How a drive copes with several reads at once. Network: each file waits for the server.
    // Fast: SSD, USB stick, SD card. Slow: optical drive, hard disk, and a disk on USB or FireWire
    // that does not report whether it is one.
    enum class DriveKind { Network, Fast, Slow };

    DriveKind driveKind(const std::filesystem::path& folder);
}
