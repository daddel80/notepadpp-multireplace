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
#include <string_view>
#include <vector>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <cstdint>
#include <cstring>             // For std::memchr
#include <new>                 // For std::bad_alloc
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

    // Why a file was skipped (None = processed)
    enum class SkipReason { None, Binary, TooLarge, Unreadable, Undecodable, ReadOnly, OpenUnsaved, Unencodable };

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

        resetSkipCounters();

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
        if (!fn || !pData)
            return false;

        // Own document: a scan never styles, so no style byte per character, and 64-bit
        // line positions for binaries over 2 GB. The view holds the only reference.
        const sptr_t doc = fn(pData, SCI_CREATEDOCUMENT, 0, SC_DOCUMENTOPTION_STYLES_NONE | SC_DOCUMENTOPTION_TEXT_LARGE);
        if (!doc)
            return false;
        fn(pData, SCI_SETDOCPOINTER, 0, doc);
        fn(pData, SCI_RELEASEDOCUMENT, 0, doc);

        fn(pData, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
        fn(pData, SCI_SETUNDOCOLLECTION, 0, 0);
        fn(pData, SCI_SETMODEVENTMASK, SC_MOD_NONE, 0);   // N++ would pass every file load on to all plugins
        return true;
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
    // 2) Apply the filter: folder rules once per folder, file rules on the name
    //    (hidden-folder handling lives in the directory enumeration)
    // ========================================================================

    // !+x: the folder is left out with everything below it
    bool excludesFolder(const std::wstring& folderName) const
    {
        for (const auto& rawPat : exclude_folders_recursive) {
            std::wstring_view pat = rawPat;
            if (!pat.empty() && (pat.front() == L'\\' || pat.front() == L'/'))
                pat.remove_prefix(1);

            if (PathMatchSpecW(folderName.c_str(), std::wstring{ pat }.c_str()))
                return true;
        }
        return false;
    }

    // !+x on the scan root: the root and every folder above it count, up to the drive
    bool excludesRoot(const std::filesystem::path& root) const
    {
        for (auto dir = root; !dir.empty() && dir != dir.root_path(); dir = dir.parent_path())
            if (excludesFolder(dir.filename().wstring()))
                return true;
        return false;
    }

    // !\x: the files directly inside the folder are left out, its subfolders are not
    bool excludesFilesIn(const std::wstring& folderName) const
    {
        for (const auto& pat : exclude_folders)
            if (PathMatchSpecW(folderName.c_str(), pat.c_str()))
                return true;
        return false;
    }

    // !*.log excludes, *.cpp includes: the file name alone decides
    bool matchFileName(const wchar_t* fileName) const
    {
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fileName, pat.c_str()))
                return false;

        if (include_patterns.empty())
            return true;

        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fileName, pat.c_str()))
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
    // header -> BOM/UTF-16 detection -> binary check -> text size -> full read -> decode.
    // Text: content holds UTF-8, enc describes the source encoding.
    // RawBytes: content holds the raw file bytes (binary skip disabled).
    // content keeps its capacity: a caller loading file after file passes the same string.
    SkipReason loadTextFile(const std::filesystem::path& fp, std::string& content,
        Encoding::EncodingInfo& enc, LoadKind& kind)
    {
        content.clear();
        enc = Encoding::EncodingInfo{};
        kind = LoadKind::Text;
        auto skip = [&](SkipReason reason) { content.clear(); return fail(reason); };

        try {
            InputFile in(fp);
            uint64_t fileSize = 0;
            if (!in.isOpen() || !in.size(fileSize)) return fail(SkipReason::Unreadable);

            const size_t maxSize = getEffectiveMaxFileSize();
            if (maxSize > 0 && fileSize > maxSize) return fail(SkipReason::TooLarge);
            if (fileSize > content.max_size()) return fail(SkipReason::TooLarge);   // 32-bit build

            // Read header for the binary/encoding decision
            const size_t headerSize = static_cast<size_t>((std::min)(fileSize, uint64_t{ BINARY_CHECK_SIZE }));
            content.resize(headerSize);
            size_t headerLen = 0;
            if (!in.read(content.data(), headerSize, headerLen)) return skip(SkipReason::Unreadable);
            content.resize(headerLen);

            const bool binary = shouldSkipAsBinary(content.data(), content.size());
            if (binary && _skipBinaryFiles) return skip(SkipReason::Binary);

            // Text beyond the converters' reach is refused before the full read
            if (!binary && fileSize > Encoding::MAX_CONVERT_LENGTH) return skip(SkipReason::TooLarge);

            // Append remainder (a file that shrank since the size query ends where it ends now)
            if (headerLen == headerSize && fileSize > headerSize) {
                content.resize(static_cast<size_t>(fileSize));
                size_t restLen = 0;
                if (!in.read(content.data() + headerLen, content.size() - headerLen, restLen))
                    return skip(SkipReason::Unreadable);
                content.resize(headerLen + restLen);
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
                if (Encoding::isValidUtf8(content.data(), content.size())) {
                    enc.kind = Encoding::Kind::UTF8;
                    enc.withBOM = false;
                    enc.bomBytes = 0;
                }
                kind = LoadKind::RawBytes;
                return SkipReason::None;
            }

            enc = Encoding::detectEncoding(content.data(), content.size());

            // UTF-8 is searched as it is: decoding would copy the same bytes, and they write back unchanged
            if (enc.kind == Encoding::Kind::UTF8) {
                content.erase(0, static_cast<size_t>(enc.bomBytes));
                return SkipReason::None;
            }

            std::string u8;
            if (!Encoding::convertBufferToUtf8(content.data(), content.size(), enc, u8))
                return skip(SkipReason::Undecodable);

            // Replace path: refuse files whose decode would not write back losslessly
            if (_verifyRoundtrip && !Encoding::verifyLosslessDecode(content.data(), content.size(), enc, u8))
                return skip(SkipReason::Undecodable);

            content.swap(u8);
            return SkipReason::None;
        }
        catch (const std::bad_alloc&) {
            return skip(SkipReason::TooLarge);
        }
        catch (...) {
            return skip(SkipReason::Unreadable);
        }
    }

    // Skip counters (surfaced in the search summary)
    size_t getSkippedBinaryCount() const      { return _skippedBinaryCount; }
    size_t getSkippedLargeCount() const       { return _skippedLargeCount; }
    size_t getSkippedUnreadableCount() const  { return _skippedUnreadableCount; }
    size_t getSkippedUndecodableCount() const { return _skippedUndecodableCount; }
    size_t getSkippedReadOnlyCount() const    { return _skippedReadOnlyCount; }
    size_t getSkippedOpenUnsavedCount() const { return _skippedOpenUnsavedCount; }
    size_t getSkippedUnencodableCount() const { return _skippedUnencodableCount; }
    size_t getSkippedTotalCount() const {
        return _skippedBinaryCount + _skippedLargeCount
             + _skippedUnreadableCount + _skippedUndecodableCount
             + _skippedReadOnlyCount + _skippedOpenUnsavedCount + _skippedUnencodableCount;
    }

    void resetSkipCounters() {
        _skippedBinaryCount = 0;
        _skippedLargeCount = 0;
        _skippedUnreadableCount = 0;
        _skippedUndecodableCount = 0;
        _skippedReadOnlyCount = 0;
        _skippedOpenUnsavedCount = 0;
        _skippedUnencodableCount = 0;
        _unreadableFolderCount = 0;
    }

    // For skips the caller decides: read-only, open with unsaved changes,
    // oversized live document, replacement not encodable in the file's codepage
    void noteSkip(SkipReason reason) { fail(reason); }

    // Folders the directory scan could not read; not part of the file total
    void noteUnreadableFolders(size_t count) { _unreadableFolderCount += count; }
    size_t getUnreadableFolderCount() const  { return _unreadableFolderCount; }

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

    // Caret ends at 0 (SCI_ADDTEXT leaves it at the end): a defined scan start.
    // False if Scintilla could not take the whole text (out of memory).
    bool setText(const std::string& txt, int codepage) {
        if (!fn || !pData) return false;
        fn(pData, SCI_SETSTATUS, SC_STATUS_OK, 0);
        fn(pData, SCI_CLEARALL, 0, 0);
        fn(pData, SCI_SETCODEPAGE, codepage, 0);
        fn(pData, SCI_ADDTEXT, txt.length(), reinterpret_cast<sptr_t>(txt.data()));
        fn(pData, SCI_GOTOPOS, 0, 0);
        return bufferIntact() && static_cast<size_t>(fn(pData, SCI_GETLENGTH, 0, 0)) == txt.length();
    }

    // False once Scintilla failed an operation on the buffer; regex warnings do not count
    bool bufferIntact() const {
        const sptr_t status = fn(pData, SCI_GETSTATUS, 0, 0);
        return !(status > SC_STATUS_OK && status < SC_STATUS_WARN_START);
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
        dbg << L"  Read-only Files:   " << _skippedReadOnlyCount << L"\n";
        dbg << L"  Open Unsaved:      " << _skippedOpenUnsavedCount << L"\n";
        dbg << L"  Unencodable Files: " << _skippedUnencodableCount << L"\n";
        dbg << L"  Unreadable Dirs:   " << _unreadableFolderCount << L"\n";

        return dbg.str();
    }

    // ========================================================================
    // Public members
    // ========================================================================

    HWND        hSci = nullptr;
    SciFnDirect fn = nullptr;
    sptr_t      pData = 0;

private:
    // Read access that leaves the file to others meanwhile (they may write, rename or delete it)
    class InputFile {
    public:
        explicit InputFile(const std::filesystem::path& fp)
            : _h(::CreateFileW(fp.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)) {}
        ~InputFile() { if (_h != INVALID_HANDLE_VALUE) ::CloseHandle(_h); }
        InputFile(const InputFile&) = delete;
        InputFile& operator=(const InputFile&) = delete;

        bool isOpen() const { return _h != INVALID_HANDLE_VALUE; }

        bool size(uint64_t& bytes) const {
            LARGE_INTEGER s{};
            if (!::GetFileSizeEx(_h, &s)) return false;
            bytes = static_cast<uint64_t>(s.QuadPart);
            return true;
        }

        // Up to count bytes, fewer only at the end of the file; false on a read error
        bool read(char* dst, size_t count, size_t& got) {
            got = 0;
            while (got < count) {
                const DWORD chunk = static_cast<DWORD>((std::min)(count - got, size_t{ 1 } << 30));
                DWORD n = 0;
                if (!::ReadFile(_h, dst + got, chunk, &n, nullptr)) return false;
                if (n == 0) break;
                got += n;
            }
            return true;
        }

    private:
        HANDLE _h;
    };

    // Counts the skip and hands the reason back to the caller
    SkipReason fail(SkipReason reason) {
        switch (reason) {
        case SkipReason::Binary:      ++_skippedBinaryCount;      break;
        case SkipReason::TooLarge:    ++_skippedLargeCount;       break;
        case SkipReason::Unreadable:  ++_skippedUnreadableCount;  break;
        case SkipReason::Undecodable: ++_skippedUndecodableCount; break;
        case SkipReason::ReadOnly:    ++_skippedReadOnlyCount;    break;
        case SkipReason::OpenUnsaved: ++_skippedOpenUnsavedCount; break;
        case SkipReason::Unencodable: ++_skippedUnencodableCount; break;
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
    size_t _skippedReadOnlyCount = 0;
    size_t _skippedOpenUnsavedCount = 0;
    size_t _skippedUnencodableCount = 0;
    size_t _unreadableFolderCount = 0;

    // Configuration
    size_t _maxFileSizeMB = DEFAULT_MAX_FILE_SIZE_MB;
    bool _limitFileSize = false;    // false = unlimited (default)
    bool _skipBinaryFiles = true;   // true = grep-style binary skip (default)
    bool _verifyRoundtrip = false;  // true = refuse lossy decodes (Replace in Files)
};