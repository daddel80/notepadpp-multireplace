#include <windows.h>
#include <string>
#include <vector>
#include <functional>
#include "Encoding.h"

//
// raw bytes (document code page) → UTF‑8 std::string
//
std::string Encoding::bytesToUtf8(const std::string_view src, UINT codePage)
{
    // 1) If the document is already UTF‑8, simply copy the bytes
    if (codePage == CP_UTF8)
        return std::string(src);

    // 2) Convert raw bytes → UTF‑16 using the document’s code page
    int wlen = ::MultiByteToWideChar(
        codePage ? codePage : CP_ACP, 0,
        src.data(), static_cast<int>(src.size()),
        nullptr, 0);

    std::wstring w(wlen, L'\0');
    ::MultiByteToWideChar(
        codePage ? codePage : CP_ACP, 0,
        src.data(), static_cast<int>(src.size()),
        w.data(), wlen);

    // 3) Convert UTF‑16 → UTF‑8
    int u8len = ::WideCharToMultiByte(
        CP_UTF8, 0,
        w.data(), wlen,
        nullptr, 0,
        nullptr, nullptr);

    std::string u8(u8len, '\0');
    ::WideCharToMultiByte(
        CP_UTF8, 0,
        w.data(), wlen,
        u8.data(), u8len,
        nullptr, nullptr);

    return u8;
}


//
// UTF‑8 → std::wstring
//
std::wstring Encoding::utf8ToWString(const std::string& utf8) {
    if (utf8.empty())
        return {};

    int required = ::MultiByteToWideChar(CP_UTF8, 0,
        utf8.data(), static_cast<int>(utf8.size()),
        nullptr, 0);
    if (required == 0)
        return {};

    std::wstring result(required, L'\0');
    int converted = ::MultiByteToWideChar(CP_UTF8, 0,
        utf8.data(), static_cast<int>(utf8.size()),
        &result[0], required);
    return (converted > 0) ? result : std::wstring();
}

//
// ANSI → std::wstring
//
std::wstring Encoding::ansiToWString(const std::string& input, UINT codePage)
{
    if (input.empty())
        return {};

    int required = ::MultiByteToWideChar(codePage, 0, input.data(), static_cast<int>(input.size()), nullptr, 0);
    if (required <= 0)
        return {};

    std::wstring result(required, L'\0');
    ::MultiByteToWideChar(codePage, 0, input.data(), static_cast<int>(input.size()), &result[0], required);
    return result;
}

//
// std::wstring → UTF‑8
//
std::string Encoding::wstringToUtf8(const std::wstring& ws)
{
    if (ws.empty())
        return {};

    int required = ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, ws.data(), static_cast<int>(ws.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
        return {};

    std::string result(required, '\0');
    ::WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, ws.data(), static_cast<int>(ws.size()), &result[0], required, nullptr, nullptr);
    return result;
}

//
// std::wstring → multibyte (chosen code page)
//
std::string Encoding::wstringToString(const std::wstring& ws, UINT codePage)
{
    if (ws.empty())
        return {};

    if (codePage == 0)          // 0 means “ANSI” for many Win32 APIs
        codePage = CP_ACP;

    int required = ::WideCharToMultiByte(codePage, 0, ws.data(), static_cast<int>(ws.size()), nullptr, 0, nullptr, nullptr);
    if (required == 0)
        return {};

    std::string result(required, '\0');
    ::WideCharToMultiByte(codePage, 0, ws.data(), static_cast<int>(ws.size()), &result[0], required, nullptr, nullptr);
    return result;
}

//
// multibyte (chosen code page) → std::wstring
//
std::wstring Encoding::stringToWString(const std::string& input, UINT codePage)
{
    if (input.empty())
        return {};

    if (codePage == 0)
        codePage = CP_ACP;

    return (codePage == CP_UTF8)
        ? utf8ToWString(input)
        : ansiToWString(input, codePage);
}

//
// trim leading/trailing whitespace & line breaks
//
std::wstring Encoding::trim(const std::wstring& str)
{
    const auto first = str.find_first_not_of(L" \t\r\n");
    if (first == std::wstring::npos)
        return L"";

    const auto last = str.find_last_not_of(L" \t\r\n");
    return str.substr(first, last - first + 1);
}

//
// quick UTF‑8 validity check
//
bool Encoding::isValidUtf8(const std::string& data)
{
    return ::MultiByteToWideChar(CP_UTF8,
        MB_ERR_INVALID_CHARS,
        data.data(),
        static_cast<int>(data.size()),
        nullptr,
        0) != 0;
}
