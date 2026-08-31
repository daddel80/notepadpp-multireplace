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

#include <windows.h>
#include <shlwapi.h>           // For PathMatchSpecW
#include <algorithm>
#include <string>
#include <vector>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <cstring>             // For std::memchr
#include "Encoding.h"
#include "StringUtils.h"       // For splitFilterPatterns
#include "Notepad_plus_msgs.h" // For NPPM_*
#include "Scintilla.h"         // For SCI_*
#pragma comment(lib, "shlwapi.lib")

extern NppData nppData;       // From your plugin definition

class HiddenSciGuard {
public:
    // ========================================================================
    // Configuration
    // ========================================================================

    // Bytes to check for binary detection (8 KB - sufficient and fast)
    static constexpr size_t BINARY_CHECK_SIZE = 8192;

    // Default max file size in MB (0 = unlimited, same as N++)
    static constexpr size_t DEFAULT_MAX_FILE_SIZE_MB = 0;

    // How a loaded file should be fed into the hidden buffer
    enum class LoadKind { Text, RawBytes };

    // Why a file was not loaded (None = loaded successfully)
    enum class SkipReason { None, Binary, TooLarge, Unreadable, Undecodable };

    // ========================================================================
    // Configuration setters/getters (for INI/Config Panel)
    // ========================================================================

    // Enable/disable file size limit (default: disabled = unlimited)
    void setFileSizeLimitEnabled(bool enabled) { _limitFileSize = enabled; }
    bool isFileSizeLimitEnabled() const { return _limitFileSize; }

    // Set max file size in MB (only applies if limit is enabled)
    void setMaxFileSizeMB(size_t sizeMB) { _maxFileSizeMB = sizeMB; }
    size_t getMaxFileSizeMB() const { return _maxFileSizeMB; }

    // Enable/disable skipping of binary files (default: enabled)
    void setSkipBinaryEnabled(bool enabled) { _skipBinaryFiles = enabled; }
    bool isSkipBinaryEnabled() const { return _skipBinaryFiles; }

    // Enable/disable lossless-roundtrip verification after decode (Replace in Files)
    void setVerifyRoundtrip(bool enabled) { _verifyRoundtrip = enabled; }
    bool isVerifyRoundtrip() const { return _verifyRoundtrip; }

    // Get effective max size in bytes (0 if unlimited)
    size_t getEffectiveMaxFileSize() const {
        if (!_limitFileSize || _maxFileSizeMB == 0)
            return 0;  // unlimited
        return _maxFileSizeMB * 1024 * 1024;
    }

    // ========================================================================
    // Constructor / Destructor
    // ========================================================================

    HiddenSciGuard() = default;
    ~HiddenSciGuard()
    {
        if (hSci) {
            ::DestroyWindow(hSci);
            hSci = nullptr;
        }
        fn = nullptr;
        pData = 0;
    }

    HiddenSciGuard(const HiddenSciGuard&) = delete;
    HiddenSciGuard& operator=(const HiddenSciGuard&) = delete;

    // ========================================================================
    // 0) Create the hidden Scintilla buffer
    // ========================================================================

    bool create()
    {
        // Destroy existing hidden Scintilla if any (safe when null)
        if (hSci) {
            ::DestroyWindow(hSci);
            hSci = nullptr;
            fn = nullptr;
            pData = 0;
        }

        // Create new hidden Scintilla via Notepad++
        hSci = reinterpret_cast<HWND>(
            ::SendMessage(nppData._nppHandle,
                NPPM_CREATESCINTILLAHANDLE,
                0, 0));
        if (!hSci)
            return false;

        fn = reinterpret_cast<SciFnDirect>(
            ::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
        pData = ::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);

        if (fn && pData)
        {
            // set safe default and avoid unnecessary memory usage
            fn(pData, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
            fn(pData, SCI_SETUNDOCOLLECTION, 0, 0);
            fn(pData, SCI_EMPTYUNDOBUFFER, 0, 0);
            fn(pData, SCI_CLEARALL, 0, 0);
        }

        resetSkipCounters();

        return fn && pData;
    }

    // ========================================================================
    // 1) Filter parsing
    // ========================================================================

    void parseFilter(const std::wstring& filterString) {
        include_patterns.clear();
        exclude_patterns.clear();
        exclude_folders.clear();
        exclude_folders_recursive.clear();

        // Semicolon is the only separator - splitting on whitespace would make
        // a pattern containing a space ("my report.txt") impossible to express.
        for (const std::wstring& tok : StringUtils::splitFilterPatterns(filterString)) {
            if (tok.rfind(L"!+", 0) == 0) {
                exclude_folders_recursive.push_back(tok.substr(2));
            }
            else if (tok.rfind(L"!", 0) == 0) {
                if (tok.size() > 1 && tok[1] == L'\\')
                    exclude_folders.push_back(tok.substr(2));
                else
                    exclude_patterns.push_back(tok.substr(1));
            }
            else {
                include_patterns.push_back(tok);
            }
        }
        // If the user only provides exclusion patterns, assume a base of *.* for inclusion
        if (include_patterns.empty() &&
            (!exclude_patterns.empty() ||
                !exclude_folders.empty() ||
                !exclude_folders_recursive.empty()))
        {
            include_patterns.push_back(L"*.*");
        }
    }

    // ========================================================================
    // 2) Test a path against the filter
    // ========================================================================

    // Pattern-only match; hidden-folder handling lives in the directory
    // enumeration (N++ semantics: prune hidden folders, keep hidden files).
    bool matchPath(const std::filesystem::path& path) const
    {
        const std::wstring fname = path.filename().wstring();
        const std::filesystem::path parentPath = path.parent_path();

        // 1) Non-recursive folder excludes (!) – only the *direct* parent folder
        if (!parentPath.empty()) {
            const std::wstring parentName = parentPath.filename().wstring();
            for (const auto& pat : exclude_folders)
                if (PathMatchSpecW(parentName.c_str(), pat.c_str()))
                    return false;
        }

        // 2) Recursive folder excludes (!+) – walk every ancestor folder
        for (auto dir = parentPath; !dir.empty() && dir != dir.root_path(); dir = dir.parent_path()) {
            const std::wstring dirName = dir.filename().wstring();

            for (const auto& rawPat : exclude_folders_recursive) {
                std::wstring_view pat = rawPat;
                if (!pat.empty() && (pat.front() == L'\\' || pat.front() == L'/'))
                    pat.remove_prefix(1);

                if (PathMatchSpecW(dirName.c_str(), std::wstring{ pat }.c_str()))
                    return false;
            }
        }

        // 3) File-level excludes (!*.log)
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return false;

        // 4) File-level includes (*.cpp…)
        if (include_patterns.empty())
            return true;

        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return true;

        return false;
    }

    // ========================================================================
    // 3) Binary Detection
    // ========================================================================

    // Check for BOM (Byte Order Mark) - files with BOM are definitely text
    bool hasBOM(const char* data, size_t len) const
    {
        if (len < 2) return false;

        const unsigned char* u = reinterpret_cast<const unsigned char*>(data);

        // UTF-8 BOM: EF BB BF
        if (len >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF)
            return true;

        // UTF-16 LE BOM: FF FE
        if (u[0] == 0xFF && u[1] == 0xFE)
            return true;

        // UTF-16 BE BOM: FE FF
        if (u[0] == 0xFE && u[1] == 0xFF)
            return true;

        return false;
    }

    // Check if content is binary by looking for NULL bytes
    // This is the industry standard approach (same as grep)
    bool hasNullBytes(const char* data, size_t len) const
    {
        const size_t checkLen = (len < BINARY_CHECK_SIZE) ? len : BINARY_CHECK_SIZE;
        // std::memchr is highly optimized (uses SIMD on modern CPUs)
        return std::memchr(data, '\0', checkLen) != nullptr;
    }

    // UTF-16 LE without BOM: same pre-checks and probe as N++ and Encoding::detectEncoding
    bool isUtf16NoBomLE(const char* data, size_t len) const
    {
        const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
        if (!(len > 1 && (len % 2) == 0 && u[0] != 0 && u[1] == 0))
            return false;
        INT uniTest = IS_TEXT_UNICODE_STATISTICS;
        return ::IsTextUnicode(data, static_cast<int>(len), &uniTest) != FALSE;
    }

    // Combined check: returns true if content is binary (no BOM, not UTF-16, NUL bytes)
    bool shouldSkipAsBinary(const char* data, size_t len) const
    {
        if (hasBOM(data, len))
            return false;

        if (isUtf16NoBomLE(data, len))
            return false;

        return hasNullBytes(data, len);
    }

    // ========================================================================
    // 4) File Loading Pipeline (shared by Find in Files and Replace in Files)
    // ========================================================================

    // Loads a file and decides ONCE how it is to be searched:
    // header -> BOM/UTF-16 detection -> binary check -> full read -> decode.
    // Text: content holds UTF-8, enc describes the source encoding.
    // RawBytes: content holds the raw file bytes (binary skip disabled).
    SkipReason loadTextFile(const std::filesystem::path& fp, std::string& content,
        Encoding::EncodingInfo& enc, LoadKind& kind)
    {
        content.clear();
        enc = Encoding::EncodingInfo{};
        kind = LoadKind::Text;

        try {
            std::error_code ec;
            const auto fileSize = std::filesystem::file_size(fp, ec);
            if (ec) return fail(SkipReason::Unreadable);

            const size_t maxSize = getEffectiveMaxFileSize();
            if (maxSize > 0 && fileSize > maxSize) return fail(SkipReason::TooLarge);

            std::ifstream in(fp, std::ios::binary);
            if (!in) return fail(SkipReason::Unreadable);

            // Read header for the binary/encoding decision
            const size_t headerSize = (fileSize < BINARY_CHECK_SIZE)
                ? static_cast<size_t>(fileSize)
                : BINARY_CHECK_SIZE;
            std::string raw(headerSize, '\0');
            in.read(raw.data(), static_cast<std::streamsize>(headerSize));
            const std::streamsize headerLen = in.gcount();
            if (headerLen <= 0 && fileSize > 0) return fail(SkipReason::Unreadable);
            raw.resize(static_cast<size_t>((std::max)(headerLen, std::streamsize(0))));

            const bool binary = shouldSkipAsBinary(raw.data(), raw.size());
            if (binary && _skipBinaryFiles) return fail(SkipReason::Binary);

            // Append remainder
            if (fileSize > headerSize) {
                const size_t offset = raw.size();
                raw.resize(offset + (static_cast<size_t>(fileSize) - headerSize));
                in.read(raw.data() + offset, static_cast<std::streamsize>(raw.size() - offset));
                raw.resize(offset + static_cast<size_t>((std::max)(in.gcount(), std::streamsize(0))));
            }

            if (binary) {
                // Skip disabled: search the bytes as-is (N++ behavior).
                // This only picks the codepage the hidden buffer is bound
                // to (pattern encoding + dock rendering); the content is
                // never decoded or transformed. isValidUtf8 is a strict
                // structural parser over the FULL content - genuine binary
                // essentially never validates, and a false "yes" could only
                // mis-encode the search pattern, never corrupt buffer or
                // disk (write-back stays verbatim). Deliberately NOT
                // detectEncoding: its CJK/UTF-16 heuristics are tuned for
                // text files and would guess confidently wrong here.
                if (Encoding::isValidUtf8(raw.data(), raw.size())) {
                    enc.kind = Encoding::Kind::UTF8;
                    enc.withBOM = false;
                    enc.bomBytes = 0;
                }
                content = std::move(raw);
                kind = LoadKind::RawBytes;
                return SkipReason::None;
            }

            enc = Encoding::detectEncoding(raw.data(), raw.size());
            std::string u8;
            if (!Encoding::convertBufferToUtf8(raw.data(), raw.size(), enc, u8))
                return fail(SkipReason::Undecodable);

            // Replace path: refuse files whose decode would not write back losslessly
            if (_verifyRoundtrip && !Encoding::verifyLosslessDecode(raw.data(), raw.size(), enc, u8))
                return fail(SkipReason::Undecodable);

            content = std::move(u8);
            return SkipReason::None;
        }
        catch (...) {
            content.clear();
            return fail(SkipReason::Unreadable);
        }
    }

    // Skip counters (surfaced in the search summary)
    size_t getSkippedBinaryCount() const      { return _skippedBinaryCount; }
    size_t getSkippedLargeCount() const       { return _skippedLargeCount; }
    size_t getSkippedUnreadableCount() const  { return _skippedUnreadableCount; }
    size_t getSkippedUndecodableCount() const { return _skippedUndecodableCount; }
    size_t getSkippedTotalCount() const {
        return _skippedBinaryCount + _skippedLargeCount
             + _skippedUnreadableCount + _skippedUndecodableCount;
    }

    void resetSkipCounters() {
        _skippedBinaryCount = 0;
        _skippedLargeCount = 0;
        _skippedUnreadableCount = 0;
        _skippedUndecodableCount = 0;
    }

    // For skips the caller decides (e.g. attached live document over the size limit)
    void noteSkip(SkipReason reason) { fail(reason); }

    // ========================================================================
    // 4b) Attach a live N++ document for searching (N++'s findInFilelist
    //     technique via SCI_SETDOCPOINTER). While attached, no document-
    //     mutating call may run - the document belongs to the user's tab.
    // ========================================================================

    struct AttachedDoc {
        HiddenSciGuard* g = nullptr;
        sptr_t ownDoc = 0;

        AttachedDoc() = default;
        AttachedDoc(const AttachedDoc&) = delete;
        AttachedDoc& operator=(const AttachedDoc&) = delete;

        void attach(HiddenSciGuard& guard, sptr_t foreignDoc) {
            g = &guard;
            ownDoc = g->fn(g->pData, SCI_GETDOCPOINTER, 0, 0);
            g->fn(g->pData, SCI_ADDREFDOCUMENT, 0, ownDoc);      // keep ours alive
            g->fn(g->pData, SCI_SETDOCPOINTER, 0, foreignDoc);   // releases ours, refs foreign
        }
        ~AttachedDoc() {
            if (!g) return;
            g->fn(g->pData, SCI_SETDOCPOINTER, 0, ownDoc);       // releases foreign, refs ours
            g->fn(g->pData, SCI_RELEASEDOCUMENT, 0, ownDoc);     // drop keep-alive ref
        }
    };

    // ========================================================================
    // 5) Write file to disk (atomic: temp file + replace)
    // ========================================================================

    bool writeFile(const std::filesystem::path& fp, const std::string& data) const {
        const std::filesystem::path tmp = fp.wstring() + L".mr_tmp";
        {
            std::ofstream o(tmp, std::ios::binary | std::ios::trunc);
            if (!o) return false;
            o.write(data.data(), data.size());
            if (!o.good()) { o.close(); std::error_code ec; std::filesystem::remove(tmp, ec); return false; }
        }

        // ReplaceFileW keeps attributes/ACLs of the target; MoveFileExW covers new files
        if (::ReplaceFileW(fp.c_str(), tmp.c_str(), nullptr, 0, nullptr, nullptr))
            return true;
        if (::MoveFileExW(tmp.c_str(), fp.c_str(), MOVEFILE_REPLACE_EXISTING))
            return true;

        std::error_code ec;
        std::filesystem::remove(tmp, ec);
        return false;
    }

    // ========================================================================
    // 6) Hidden-buffer helpers
    // ========================================================================

    void setText(const std::string& txt) {
        if (!fn || !pData) return;
        fn(pData, SCI_CLEARALL, 0, 0);
        fn(pData, SCI_ADDTEXT, txt.length(), reinterpret_cast<sptr_t>(txt.data()));
    }

    std::string getText() const
    {
        if (!fn || !pData) return {};
        Sci_Position len = fn(pData, SCI_GETLENGTH, 0, 0);
        if (len <= 0) return {};
        std::string buf(static_cast<size_t>(len), '\0');
        Sci_TextRangeFull tr;
        tr.chrg.cpMin = 0;
        tr.chrg.cpMax = len;
        tr.lpstrText = buf.data();
        fn(pData, SCI_GETTEXTRANGEFULL, 0, reinterpret_cast<sptr_t>(&tr));
        return buf;
    }

    void replaceAllInBuffer(const std::string& findUtf8,
        const std::string& replUtf8,
        int searchFlags)
    {
        if (!fn || !pData) return;

        fn(pData, SCI_SETSEARCHFLAGS, searchFlags, 0);
        fn(pData, SCI_BEGINUNDOACTION, 0, 0);

        sptr_t docLen = fn(pData, SCI_GETLENGTH, 0, 0);
        sptr_t start = 0;
        fn(pData, SCI_SETTARGETRANGE, start, docLen);

        while (fn(pData,
            SCI_SEARCHINTARGET,
            static_cast<sptr_t>(findUtf8.size()),
            reinterpret_cast<sptr_t>(findUtf8.c_str())
        ) != -1)
        {
            fn(pData,
                SCI_REPLACETARGET,
                static_cast<sptr_t>(replUtf8.size()),
                reinterpret_cast<sptr_t>(replUtf8.c_str())
            );
            start = fn(pData, SCI_GETTARGETEND, 0, 0);
            docLen = fn(pData, SCI_GETLENGTH, 0, 0);
            fn(pData, SCI_SETTARGETRANGE, start, docLen);
        }

        fn(pData, SCI_ENDUNDOACTION, 0, 0);
    }

    // ========================================================================
    // 7) Debug helpers
    // ========================================================================

    std::wstring getFilterDebugString() const {
        std::wstringstream dbg;
        dbg << L"--- Internal Filter State ---\n";

        dbg << L"Include Patterns (" << include_patterns.size() << L"):\n";
        if (include_patterns.empty()) dbg << L"  (none)\n";
        for (const auto& p : include_patterns) dbg << L"  '" << p << L"'\n";

        dbg << L"\nExclude Patterns (" << exclude_patterns.size() << L"):\n";
        if (exclude_patterns.empty()) dbg << L"  (none)\n";
        for (const auto& p : exclude_patterns) dbg << L"  '!" << p << L"'\n";

        dbg << L"\nExclude Folders (" << exclude_folders.size() << L"):\n";
        if (exclude_folders.empty()) dbg << L"  (none)\n";
        for (const auto& p : exclude_folders) dbg << L"  '!\\" << p << L"'\n";

        dbg << L"\nExclude Folders (recursive) (" << exclude_folders_recursive.size() << L"):\n";
        if (exclude_folders_recursive.empty()) dbg << L"  (none)\n";
        for (const auto& p : exclude_folders_recursive) dbg << L"  '!+" << p << L"'\n";

        dbg << L"\n--- File Size Limit ---\n";
        if (_limitFileSize) {
            dbg << L"  Enabled: " << _maxFileSizeMB << L" MB\n";
        }
        else {
            dbg << L"  Disabled (unlimited)\n";
        }

        dbg << L"\n--- Skip Statistics ---\n";
        dbg << L"  Binary Files:      " << _skippedBinaryCount << L"\n";
        dbg << L"  Large Files:       " << _skippedLargeCount << L"\n";
        dbg << L"  Unreadable Files:  " << _skippedUnreadableCount << L"\n";
        dbg << L"  Undecodable Files: " << _skippedUndecodableCount << L"\n";

        return dbg.str();
    }

    // ========================================================================
    // Public members
    // ========================================================================

    HWND        hSci = nullptr;
    SciFnDirect fn = nullptr;
    sptr_t      pData = 0;

private:
    // Counts the skip and hands the reason back to the caller
    SkipReason fail(SkipReason reason) {
        switch (reason) {
        case SkipReason::Binary:      ++_skippedBinaryCount;      break;
        case SkipReason::TooLarge:    ++_skippedLargeCount;       break;
        case SkipReason::Unreadable:  ++_skippedUnreadableCount;  break;
        case SkipReason::Undecodable: ++_skippedUndecodableCount; break;
        case SkipReason::None:        break;
        }
        return reason;
    }

    std::vector<std::wstring> include_patterns;
    std::vector<std::wstring> exclude_patterns;
    std::vector<std::wstring> exclude_folders;
    std::vector<std::wstring> exclude_folders_recursive;

    // Skip counters (per operation; reset in create())
    size_t _skippedBinaryCount = 0;
    size_t _skippedLargeCount = 0;
    size_t _skippedUnreadableCount = 0;
    size_t _skippedUndecodableCount = 0;

    // Configuration
    size_t _maxFileSizeMB = DEFAULT_MAX_FILE_SIZE_MB;
    bool _limitFileSize = false;    // false = unlimited (default)
    bool _skipBinaryFiles = true;   // true = grep-style binary skip (default)
    bool _verifyRoundtrip = false;  // true = refuse lossy decodes (Replace in Files)
};