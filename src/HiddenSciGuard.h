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
    HiddenSciGuard() = default;
    ~HiddenSciGuard() {
        if (hSci) {
            ::SendMessage(nppData._nppHandle,
                NPPM_DESTROYSCINTILLAHANDLE,
                0,
                reinterpret_cast<LPARAM>(hSci));
        }
    }
    HiddenSciGuard(const HiddenSciGuard&) = delete;
    HiddenSciGuard& operator=(const HiddenSciGuard&) = delete;

    // 0) Create the hidden Scintilla buffer
    bool create() {
        hSci = reinterpret_cast<HWND>(
            ::SendMessage(nppData._nppHandle,
                NPPM_CREATESCINTILLAHANDLE,
                0, 0));
        if (!hSci) return false;
        fn = reinterpret_cast<SciFnDirect>(
            ::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
        pData = ::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
        
        if (fn && pData)
            fn(pData, SCI_SETCODEPAGE, SC_CP_UTF8, 0);     // set safe default

        return fn && pData;
    }

    // 1) Filter parsing
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

    // 2) Test a path against the filter
    // Requires <ranges> (C++20) for clarity; switch to a loop if C++17.
    bool matchPath(const std::filesystem::path& path, bool includeHidden) const
    {
        // 1) Hidden attribute ----------------------------------------------------
        if (!includeHidden) {
            const DWORD a = GetFileAttributesW(path.c_str());
            if (a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_HIDDEN))
                return false;
        }

        const std::wstring fname = path.filename().wstring();
        const std::filesystem::path parentPath = path.parent_path();

        // 2) Non-recursive folder excludes (!)  – only the *direct* parent folder
        if (!parentPath.empty()) {
            const std::wstring parentName = parentPath.filename().wstring();
            for (const auto& pat : exclude_folders)
                if (PathMatchSpecW(parentName.c_str(), pat.c_str()))
                    return false;
        }

        // 3) Recursive folder excludes (!+)  – walk every ancestor folder
        for (auto dir = parentPath; !dir.empty() && dir != dir.root_path(); dir = dir.parent_path()) {
            const std::wstring dirName = dir.filename().wstring();

            for (const auto& rawPat : exclude_folders_recursive) {
                // Strip an optional leading backslash the user may have typed
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

        // 5) File-level includes (*.cpp…)  – if no include pattern, everything left is in
        if (include_patterns.empty())
            return true;

        for (const auto& pat : include_patterns)
            if (PathMatchSpecW(fname.c_str(), pat.c_str()))
                return true;

        return false;
    }

    // 3) Read file into string
    bool loadFile(const std::filesystem::path& fp, std::string& out) const {
        std::ifstream in(fp, std::ios::binary);
        if (!in) return false;
        out.assign(std::istreambuf_iterator<char>(in),
            std::istreambuf_iterator<char>());
        return true;
    }

    // 4) Write string back to disk
    bool writeFile(const std::filesystem::path& fp, const std::string& data) const {
        std::ofstream o(fp, std::ios::binary | std::ios::trunc);
        if (!o) return false;
        o.write(data.data(), data.size());
        return true;
    }

    // 5) Hidden-buffer helpers
    // FIX: Null-safe method for setting the text
    void setText(const std::string& txt) {
        fn(pData, SCI_CLEARALL, 0, 0);
        fn(pData, SCI_ADDTEXT, txt.length(), reinterpret_cast<sptr_t>(txt.data()));
    }

    std::string getText() const
    {
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

        return dbg.str();
    }

    HWND        hSci = nullptr;
    SciFnDirect fn = nullptr;
    sptr_t      pData = 0;

private:
    std::vector<std::wstring> include_patterns;
    std::vector<std::wstring> exclude_patterns;
    std::vector<std::wstring> exclude_folders;
    std::vector<std::wstring> exclude_folders_recursive;
};