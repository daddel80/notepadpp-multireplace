// ============================================================================
// DEBUG_SEARCH.h - Complete diagnostic logging for search bug
// REMOVE after bugfix!
// ============================================================================
#pragma once

#include <string>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <ctime>
#include <windows.h>

namespace DebugSearch {

    inline bool g_enabled = true;
    inline bool g_initialized = false;

    inline std::wstring getLogPath() {
        wchar_t temp[MAX_PATH];
        GetTempPathW(MAX_PATH, temp);
        return std::wstring(temp) + L"MultiReplace_Debug.log";
    }

    inline void log(const std::wstring& msg) {
        if (!g_enabled) return;
        std::wofstream f(getLogPath(), std::ios::app);
        if (f.is_open()) { f << msg << L"\n"; f.close(); }
    }

    inline std::wstring toHex(const std::string& bytes, size_t max = 100) {
        std::wstringstream ss;
        ss << L"[";
        size_t n = (std::min)(bytes.size(), max);
        for (size_t i = 0; i < n; ++i) {
            if (i > 0) ss << L" ";
            ss << std::hex << std::setfill(L'0') << std::setw(2)
                << static_cast<unsigned>(static_cast<unsigned char>(bytes[i]));
        }
        if (bytes.size() > max) ss << L" ...+" << std::dec << (bytes.size() - max);
        ss << L"]";
        return ss.str();
    }

    inline std::wstring toHexW(const std::wstring& str, size_t max = 50) {
        std::wstringstream ss;
        ss << L"[";
        size_t n = (std::min)(str.size(), max);
        for (size_t i = 0; i < n; ++i) {
            if (i > 0) ss << L" ";
            ss << std::hex << std::setfill(L'0') << std::setw(4)
                << static_cast<unsigned>(str[i]);
        }
        if (str.size() > max) ss << L" ...+" << std::dec << (str.size() - max);
        ss << L"]";
        return ss.str();
    }

    inline std::wstring flagsStr(int f) {
        std::wstringstream ss;
        ss << L"0x" << std::hex << f << std::dec << L"(";
        bool first = true;
        if (f & 0x200000) { ss << L"REGEX"; first = false; }
        if (f & 0x4) { if (!first) ss << L"|"; ss << L"CASE"; first = false; }
        if (f & 0x2) { if (!first) ss << L"|"; ss << L"WORD"; first = false; }
        if (first) ss << L"NONE";
        ss << L")";
        return ss.str();
    }

    inline void init(const std::wstring& pluginVer, const std::wstring& nppVer) {
        if (g_initialized) return;
        g_initialized = true;

        // Truncate/create log
        std::wofstream f(getLogPath(), std::ios::trunc);
        if (!f.is_open()) return;

        std::time_t t = std::time(nullptr);
        std::tm tm; localtime_s(&tm, &t);
        wchar_t ts[64]; wcsftime(ts, 64, L"%Y-%m-%d %H:%M:%S", &tm);

        // System info
        OSVERSIONINFOEXW os = { sizeof(os) };
        typedef NTSTATUS(WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
        auto RtlGetVersion = (RtlGetVersionPtr)GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlGetVersion");
        if (RtlGetVersion) RtlGetVersion((PRTL_OSVERSIONINFOW)&os);

        wchar_t locale[LOCALE_NAME_MAX_LENGTH];
        GetUserDefaultLocaleName(locale, LOCALE_NAME_MAX_LENGTH);

        f << L"================================================================\n";
        f << L"MultiReplace Debug Log - " << ts << L"\n";
        f << L"================================================================\n\n";
        f << L"[SYSTEM]\n";
        f << L"  Plugin: " << pluginVer << L"\n";
        f << L"  Notepad++: " << nppVer << L"\n";
        f << L"  Windows: " << os.dwMajorVersion << L"." << os.dwMinorVersion << L" Build " << os.dwBuildNumber << L"\n";
        f << L"  Locale: " << locale << L"\n";
        f << L"  ACP: " << GetACP() << L"  OEM: " << GetOEMCP() << L"\n";
        f << L"================================================================\n\n";
        f.close();

        // Show message once
        std::wstring msg = L"DEBUG MODE\n\nLog: " + getLogPath() + L"\n\nTest and send log file.";
        MessageBoxW(NULL, msg.c_str(), L"MultiReplace Debug", MB_OK | MB_ICONINFORMATION);
    }

}  // namespace DebugSearch
