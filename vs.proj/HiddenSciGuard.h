#pragma once

#include <windows.h>
#include <shlwapi.h>           // For PathMatchSpecW
#include <string>
#include <vector>
#include <filesystem>
#include "C:\Users\knoetho\source\repos\notepadpp-multireplace\src\Notepad_plus_msgs.h" // For NPPM_CREATESCINTILLAHANDLE etc.
#include "C:\Users\knoetho\source\repos\notepadpp-multireplace\src\Scintilla.h"         // For SCI_* constants
#include "C:\Users\knoetho\source\repos\notepadpp-multireplace\src\MultiReplacePanel.h"

#pragma comment(lib, "shlwapi.lib")

// Type alias for the Scintilla direct‐function pointer
typedef sptr_t(__stdcall* SciFnDirect)(sptr_t, unsigned int, uptr_t, sptr_t);

class HiddenSciGuard {
public:
    // --- KORRIGIERT: Konstruktor hinzugefügt, um die NppData-Referenz zu initialisieren ---
    // Der Konstruktor nimmt eine Referenz auf NppData, wodurch die Abhängigkeit klar wird.
    inline HiddenSciGuard(NppData& nppData) : m_nppData(nppData) {}

    inline ~HiddenSciGuard() {
        // Stellt sicher, dass die Scintilla-Instanz zerstört wird, wenn das Objekt den Geltungsbereich verlässt.
        if (m_hSci) {
            ::SendMessage(m_nppData._nppHandle, NPPM_DESTROYSCINTILLAHANDLE, 0, reinterpret_cast<LPARAM>(m_hSci));
        }
    }

    // Kopieren und Zuweisen deaktivieren, um den RAII-Prinzipien zu folgen
    HiddenSciGuard(const HiddenSciGuard&) = delete;
    HiddenSciGuard& operator=(const HiddenSciGuard&) = delete;

    // --- HINZUGEFÜGT: Die fehlenden Getter-Methoden ---
    inline HWND getHandle() const { return m_hSci; }
    inline SciFnDirect getDirectFn() const { return m_fn; }
    inline sptr_t getDirectData() const { return m_pData; }
    // --- ENDE DER HINZUGEFÜGTEN METHODEN ---


    // --- Scintilla Operations ---

    inline bool create() {
        m_hSci = reinterpret_cast<HWND>(::SendMessage(m_nppData._nppHandle, NPPM_CREATESCINTILLAHANDLE, 0, 0));
        if (!m_hSci) return false;

        m_fn = reinterpret_cast<SciFnDirect>(::SendMessage(m_hSci, SCI_GETDIRECTFUNCTION, 0, 0));
        m_pData = ::SendMessage(m_hSci, SCI_GETDIRECTPOINTER, 0, 0);
        return (m_fn != nullptr) && (m_pData != 0);
    }

    inline void setText(const std::string& txt) {
        if (m_fn && m_pData) {
            m_fn(m_pData, SCI_SETTEXT, 0, reinterpret_cast<sptr_t>(txt.c_str()));
        }
    }

    inline std::string getText() const {
        if (!m_fn || !m_pData) return {};
        int len = static_cast<int>(m_fn(m_pData, SCI_GETLENGTH, 0, 0));
        if (len <= 0) return {};
        std::string s(len, '\0');
        Sci_TextRange tr{ {0, len}, &s[0] };
        m_fn(m_pData, SCI_GETTEXTRANGE, 0, reinterpret_cast<sptr_t>(&tr));
        return s;
    }

    inline void replaceAllInBuffer(const std::string& findUtf8, const std::string& replUtf8, int searchFlags) {
        if (!m_fn || !m_pData) return;

        m_fn(m_pData, SCI_SETSEARCHFLAGS, searchFlags, 0);
        m_fn(m_pData, SCI_BEGINUNDOACTION, 0, 0);
        sptr_t docLen = m_fn(pData, SCI_GETLENGTH, 0, 0);
        sptr_t start = 0;
        m_fn(m_pData, SCI_SETTARGETRANGE, start, docLen);
        while (m_fn(m_pData, SCI_SEARCHINTARGET, static_cast<sptr_t>(findUtf8.size()), reinterpret_cast<sptr_t>(findUtf8.c_str())) != -1) {
            m_fn(m_pData, SCI_REPLACETARGET, static_cast<sptr_t>(replUtf8.size()), reinterpret_cast<sptr_t>(replUtf8.c_str()));
            start = m_fn(m_pData, SCI_GETTARGETEND, 0, 0);
            docLen = m_fn(m_pData, SCI_GETLENGTH, 0, 0);
            if (start >= docLen) break;
            m_fn(m_pData, SCI_SETTARGETRANGE, start, docLen);
        }
        m_fn(m_pData, SCI_ENDUNDOACTION, 0, 0);
    }

    // --- File Operations ---

    inline bool loadFile(const std::filesystem::path& fp, std::string& out) const {
        std::ifstream in(fp, std::ios::binary);
        if (!in) return false;
        out.assign((std::istreambuf_iterator<char>(in)), (std::istreambuf_iterator<char>()));
        return true;
    }

    inline bool writeFile(const std::filesystem::path& fp, const std::string& data) const {
        std::ofstream o(fp, std::ios::binary | std::ios::trunc);
        if (!o) return false;
        o.write(data.data(), data.size());
        return true;
    }

    // --- Filter Logic ---

    inline void parseFilter(const std::wstring& filterString) {
        m_includePatterns.clear();
        m_excludePatterns.clear();
        m_excludeFolders.clear();
        m_excludeFoldersRecursive.clear();

        std::wstringstream ss(filterString);
        std::wstring token;

        while (ss >> token) {
            if (token.empty()) continue;

            if (token.rfind(L"!+", 0) == 0 && token.length() > 2) {
                m_excludeFoldersRecursive.push_back(token.substr(2));
            }
            else if (token.rfind(L"!\\", 0) == 0 && token.length() > 2) {
                m_excludeFolders.push_back(token.substr(2));
            }
            else if (token.rfind(L'!', 0) == 0 && token.length() > 1) {
                if (!m_excludePatterns.empty()) m_excludePatterns += L';';
                m_excludePatterns += token.substr(1);
            }
            else {
                if (!m_includePatterns.empty()) m_includePatterns += L';';
                m_includePatterns += token;
            }
        }
    }

    inline bool matchPath(const std::filesystem::path& path, bool includeHidden) const {
        if (!includeHidden) {
            DWORD attrs = GetFileAttributesW(path.c_str());
            if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_HIDDEN)) {
                return false;
            }
        }

        for (auto p = path.parent_path(); p.has_parent_path(); p = p.parent_path()) {
            const std::wstring folder_name = p.filename().wstring();
            for (const auto& pattern : m_excludeFoldersRecursive) {
                if (::PathMatchSpecW(folder_name.c_str(), pattern.c_str())) return false;
            }
            if (p == path.parent_path()) {
                for (const auto& pattern : m_excludeFolders) {
                    if (::PathMatchSpecW(folder_name.c_str(), pattern.c_str())) return false;
                }
            }
        }

        const std::wstring filename = path.filename().wstring();

        if (!m_excludePatterns.empty() && ::PathMatchSpecW(filename.c_str(), m_excludePatterns.c_str())) {
            return false;
        }

        if (m_includePatterns.empty()) {
            return true;
        }

        if (::PathMatchSpecW(filename.c_str(), m_includePatterns.c_str())) {
            return true;
        }

        return false;
    }

private:
    NppData& m_nppData; // Referenz auf Notepad++ Daten, wird im Konstruktor übergeben.
    HWND        m_hSci = nullptr;
    SciFnDirect m_fn = nullptr;
    sptr_t      m_pData = 0;
    std::wstring m_includePatterns;
    std::wstring m_excludePatterns;
    std::vector<std::wstring> m_excludeFolders;
    std::vector<std::wstring> m_excludeFoldersRecursive;
};