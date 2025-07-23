#pragma once
#include <string>
#include <Windows.h>

/// Collection of static helper functions for common text‑encoding tasks.

class Encoding final
{
public:
    // raw bytes(document code page) → UTF‑8 std::string
    static std::string bytesToUtf8(const std::string_view src, UINT codePage);
    static std::wstring Encoding::bytesToWString(const std::string& raw, UINT cp);

    // UTF‑8 ⇄ std::wstring
    static std::wstring utf8ToWString(const std::string& utf8);
    static std::string  wstringToUtf8(const std::wstring& ws);

    // ANSI code page ⇄ std::wstring
    static std::wstring ansiToWString(const std::string& input, UINT codePage = CP_ACP);
    static std::string  wstringToString(const std::wstring& ws, UINT codePage = CP_ACP);

    // multibyte (generic) ⇄ std::wstring
    static std::wstring stringToWString(const std::string& input, UINT codePage = CP_ACP);

    // utilities
    static std::wstring trim(const std::wstring& str);
    static bool         isValidUtf8(const std::string& data);

private:
    Encoding() = delete;  // static‑only class
    ~Encoding() = delete;
};
