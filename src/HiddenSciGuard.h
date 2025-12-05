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
#include <string>
#include <vector>
#include <filesystem>
#include <fstream>
#include <sstream>
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

    // ========================================================================
    // Configuration setters/getters (for INI/Config Panel)
    // ========================================================================

    // Enable/disable file size limit (default: disabled = unlimited)
    void setFileSizeLimitEnabled(bool enabled) { _limitFileSize = enabled; }
    bool isFileSizeLimitEnabled() const { return _limitFileSize; }

    // Set max file size in MB (only applies if limit is enabled)
    void setMaxFileSizeMB(size_t sizeMB) { _maxFileSizeMB = sizeMB; }
    size_t getMaxFileSizeMB() const { return _maxFileSizeMB; }

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

        // Reset skip counters
        _skippedBinaryCount = 0;
        _skippedLargeCount = 0;

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

        std::wstringstream ss(filterString);
        std::wstring tok;
        while (ss >> tok) {
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

    bool matchPath(const std::filesystem::path& path, bool includeHidden) const
    {
        // 1) Hidden attribute
        if (!includeHidden) {
            const DWORD a = GetFileAttributesW(path.c_str());
            if (a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_HIDDEN))
                return false;
        }

        const std::wstring fname = path.filename().wstring();
        const std::filesystem::path parentPath = path.parent_path();

        // 2) Non-recursive folder excludes (!) – only the *direct* parent folder
        if (!parentPath.empty()) {
            const std::wstring parentName = parentPath.filename().wstring();
            for (const auto& pat : exclude_folders)
                if (PathMatchSpecW(parentName.c_str(), pat.c_str()))
                    return false;
        }

        // 3) Recursive folder excludes (!+) – walk every ancestor folder
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

        // 4) File-level excludes (!*.log)
        for (const auto& pat : exclude_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return false;

        // 5) File-level includes (*.cpp…)
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

        for (size_t i = 0; i < checkLen; ++i)
        {
            if (data[i] == 0x00)
                return true;
        }

        return false;
    }

    // Combined check: returns true if file should be skipped as binary
    bool shouldSkipAsBinary(const char* data, size_t len) const
    {
        // 1. Files with BOM are definitely text files
        if (hasBOM(data, len))
            return false;

        // 2. Check for NULL bytes (binary indicator)
        return hasNullBytes(data, len);
    }

    // ========================================================================
    // 4) File Loading with Binary Detection
    // ========================================================================

    // Load file with automatic binary detection
    // Returns true on success, false on any failure (including binary skip)
    bool loadFile(const std::filesystem::path& fp, std::string& out)
    {
        out.clear();

        try {
            // Get file size
            std::error_code ec;
            auto fileSize = std::filesystem::file_size(fp, ec);
            if (ec) return false;

            // Check file size limit (if enabled)
            size_t maxSize = getEffectiveMaxFileSize();
            if (maxSize > 0 && fileSize > maxSize)
            {
                ++_skippedLargeCount;
                return false;
            }

            std::ifstream in(fp, std::ios::binary);
            if (!in) return false;

            // Read header for binary detection
            char header[BINARY_CHECK_SIZE];
            in.read(header, sizeof(header));
            std::streamsize headerLen = in.gcount();

            // Check if binary - skip if so
            if (headerLen > 0 && shouldSkipAsBinary(header, static_cast<size_t>(headerLen)))
            {
                ++_skippedBinaryCount;
                return false;
            }

            // Not binary - read full file
            in.clear();
            in.seekg(0, std::ios::beg);

            out.reserve(static_cast<size_t>(fileSize));
            out.assign(std::istreambuf_iterator<char>(in),
                std::istreambuf_iterator<char>());

            return true;
        }
        catch (...) {
            return false;
        }
    }

    // Get count of skipped binary files (for status messages)
    size_t getSkippedBinaryCount() const { return _skippedBinaryCount; }

    // Get count of skipped large files (for status messages)
    size_t getSkippedLargeCount() const { return _skippedLargeCount; }

    // Reset the skip counters
    void resetSkipCounters() {
        _skippedBinaryCount = 0;
        _skippedLargeCount = 0;
    }

    // ========================================================================
    // 5) Write file to disk
    // ========================================================================

    bool writeFile(const std::filesystem::path& fp, const std::string& data) const {
        std::ofstream o(fp, std::ios::binary | std::ios::trunc);
        if (!o) return false;
        o.write(data.data(), data.size());
        return o.good();
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
        dbg << L"  Binary Files: " << _skippedBinaryCount << L"\n";
        dbg << L"  Large Files:  " << _skippedLargeCount << L"\n";

        return dbg.str();
    }

    // ========================================================================
    // Public members
    // ========================================================================

    HWND        hSci = nullptr;
    SciFnDirect fn = nullptr;
    sptr_t      pData = 0;

private:
    std::vector<std::wstring> include_patterns;
    std::vector<std::wstring> exclude_patterns;
    std::vector<std::wstring> exclude_folders;
    std::vector<std::wstring> exclude_folders_recursive;

    // Counter for skipped binary files
    size_t _skippedBinaryCount = 0;

    // Counter for skipped large files (when limit enabled)
    size_t _skippedLargeCount = 0;

    // Configurable max file size (0 = unlimited)
    size_t _maxFileSizeMB = DEFAULT_MAX_FILE_SIZE_MB;
    bool _limitFileSize = false;  // false = unlimited (default)
};