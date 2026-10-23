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
#include <string_view>
#include <vector>
#include <filesystem>
#include <fstream>
#include <sstream>
#include "Encoding.h"
#include "FileLoad.h"
#include "StringUtils.h"       // For splitFilterPatterns
#include "Notepad_plus_msgs.h" // For NPPM_*
#include "Scintilla.h"         // For SCI_*
#pragma comment(lib, "shlwapi.lib")

extern NppData nppData;       // defined in PluginDefinition.cpp

class HiddenSciGuard {
public:
    // ========================================================================
    // Configuration
    // ========================================================================

    // Default max file size in MB (0 = unlimited, same as N++)
    static constexpr size_t DEFAULT_MAX_FILE_SIZE_MB = 0;

    using LoadKind = FileLoad::LoadKind;
    using SkipReason = FileLoad::SkipReason;

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
    ~HiddenSciGuard() { release(); }

    HiddenSciGuard(const HiddenSciGuard&) = delete;
    HiddenSciGuard& operator=(const HiddenSciGuard&) = delete;

    // ========================================================================
    // 0) Bind the hidden Scintilla buffer
    // ========================================================================

    // Notepad++ frees the Scintillas it creates for plugins only on exit, so one hidden
    // view serves all runs, one run at a time; each run gets a document of its own
    bool create()
    {
        release();
        resetSkipCounters();

        View& view = sharedView();
        if (view.busy)
            return false;
        if (!view.hSci || !::IsWindow(view.hSci)) {
            view = View{};
            const HWND h = reinterpret_cast<HWND>(::SendMessage(nppData._nppHandle, NPPM_CREATESCINTILLAHANDLE, 0, 0));
            if (!h)
                return false;
            const auto f = reinterpret_cast<SciFnDirect>(::SendMessage(h, SCI_GETDIRECTFUNCTION, 0, 0));
            const sptr_t d = ::SendMessage(h, SCI_GETDIRECTPOINTER, 0, 0);
            if (!f || !d)
                return false;
            f(d, SCI_SETMODEVENTMASK, SC_MOD_NONE, 0);   // N++ would pass every file load on to all plugins
            view = View{ h, f, d, false };
        }

        view.busy = true;
        hSci = view.hSci;
        fn = view.fn;
        pData = view.pData;
        return newDocument();
    }

    // Ends the run: its document is freed, the view stays for the next run.
    // Closing Notepad++ during a run destroys the view first: then nothing is left to free.
    void release()
    {
        if (!fn)
            return;
        if (::IsWindow(hSci))
            fn(pData, SCI_SETDOCPOINTER, 0, 0);
        sharedView().busy = false;
        hSci = nullptr;
        fn = nullptr;
        pData = 0;
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
    // 3) File Loading Pipeline (shared by Find in Files and Replace in Files)
    // ========================================================================

    // Loads a file and decides ONCE how it is to be searched (FileLoad::load), counting a skip.
    SkipReason loadTextFile(const std::filesystem::path& fp, std::string& content,
        Encoding::EncodingInfo& enc, LoadKind& kind)
    {
        const FileLoad::Result loaded = FileLoad::load(fp, loadOptions(), content);
        enc = loaded.enc;
        kind = loaded.kind;
        return fail(loaded.skip);
    }

    // This run's load settings, for loading on worker threads
    FileLoad::Options loadOptions() const
    {
        FileLoad::Options options;
        options.maxFileSize = getEffectiveMaxFileSize();
        options.skipBinary = _skipBinaryFiles;
        options.verifyRoundtrip = _verifyRoundtrip;
        return options;
    }

    // Skip counters (surfaced in the search summary)
    size_t getSkippedBinaryCount() const      { return _skippedBinaryCount; }
    size_t getSkippedLargeCount() const       { return _skippedLargeCount; }
    size_t getSkippedUnreadableCount() const  { return _skippedUnreadableCount; }
    size_t getSkippedUndecodableCount() const { return _skippedUndecodableCount; }
    size_t getSkippedReadOnlyCount() const    { return _skippedReadOnlyCount; }
    size_t getSkippedOpenUnsavedCount() const { return _skippedOpenUnsavedCount; }
    size_t getSkippedUnencodableCount() const { return _skippedUnencodableCount; }
    size_t getSkippedUnwritableCount() const  { return _skippedUnwritableCount; }
    size_t getSkippedTotalCount() const {
        return _skippedBinaryCount + _skippedLargeCount
             + _skippedUnreadableCount + _skippedUndecodableCount
             + _skippedReadOnlyCount + _skippedOpenUnsavedCount
             + _skippedUnencodableCount + _skippedUnwritableCount;
    }

    void resetSkipCounters() {
        _skippedBinaryCount = 0;
        _skippedLargeCount = 0;
        _skippedUnreadableCount = 0;
        _skippedUndecodableCount = 0;
        _skippedReadOnlyCount = 0;
        _skippedOpenUnsavedCount = 0;
        _skippedUnencodableCount = 0;
        _skippedUnwritableCount = 0;
        _unreadableFolderCount = 0;
    }

    // For skips the caller decides: read-only, open with unsaved changes, oversized
    // live document, replacement not encodable in the file's codepage, failed write
    void noteSkip(SkipReason reason) { fail(reason); }

    // Folders the directory scan could not read; not part of the file total
    void noteUnreadableFolders(size_t count) { _unreadableFolderCount += count; }
    size_t getUnreadableFolderCount() const  { return _unreadableFolderCount; }

    // ========================================================================
    // 3b) Attach a live N++ document for searching (N++'s findInFilelist
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
            if (!g || !::IsWindow(g->hSci)) return;
            g->fn(g->pData, SCI_SETDOCPOINTER, 0, ownDoc);       // releases foreign, refs ours
            g->fn(g->pData, SCI_RELEASEDOCUMENT, 0, ownDoc);     // drop keep-alive ref
        }
    };

    // ========================================================================
    // 4) Write file to disk (atomic: temp file + replace)
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
    // 5) Hidden-buffer helpers
    // ========================================================================

    // Ends with a single empty selection at 0: a defined scan start. Without window work per
    // file: SCI_CLEARALL would discard the font data of the view, SCI_GOTOPOS lay out line 0.
    // False if Scintilla could not take the whole text (out of memory).
    bool setText(const std::string& txt, int codepage) {
        if (!fn || !pData) return false;
        // After a failed allocation Scintilla refuses every further change to that document
        if (!bufferIntact() && !newDocument()) return false;
        fn(pData, SCI_SETSTATUS, SC_STATUS_OK, 0);
        fn(pData, SCI_DELETERANGE, 0, fn(pData, SCI_GETLENGTH, 0, 0));
        fn(pData, SCI_SETCODEPAGE, codepage, 0);
        fn(pData, SCI_APPENDTEXT, txt.length(), reinterpret_cast<sptr_t>(txt.data()));
        fn(pData, SCI_SETEMPTYSELECTION, 0, 0);
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
    // 6) Debug helpers
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
        dbg << L"  Unwritable Files:  " << _skippedUnwritableCount << L"\n";
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
    // The hidden view all runs share
    struct View {
        HWND        hSci = nullptr;
        SciFnDirect fn = nullptr;
        sptr_t      pData = 0;
        bool        busy = false;
    };

    static View& sharedView()
    {
        static View view;
        return view;
    }

    // Own document: a scan never styles, so no style byte per character, and 64-bit
    // line positions for binaries over 2 GB. The view holds the only reference.
    bool newDocument()
    {
        const sptr_t doc = fn(pData, SCI_CREATEDOCUMENT, 0, SC_DOCUMENTOPTION_STYLES_NONE | SC_DOCUMENTOPTION_TEXT_LARGE);
        if (!doc)
            return false;
        fn(pData, SCI_SETDOCPOINTER, 0, doc);
        fn(pData, SCI_RELEASEDOCUMENT, 0, doc);
        fn(pData, SCI_SETCODEPAGE, SC_CP_UTF8, 0);
        fn(pData, SCI_SETUNDOCOLLECTION, 0, 0);
        return true;
    }

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
        case SkipReason::Unwritable:  ++_skippedUnwritableCount;  break;
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
    size_t _skippedUnwritableCount = 0;
    size_t _unreadableFolderCount = 0;

    // Configuration
    size_t _maxFileSizeMB = DEFAULT_MAX_FILE_SIZE_MB;
    bool _limitFileSize = false;    // false = unlimited (default)
    bool _skipBinaryFiles = true;   // true = grep-style binary skip (default)
    bool _verifyRoundtrip = false;  // true = refuse lossy decodes (Replace in Files)
};