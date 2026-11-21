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

#include <windows.h>

#include <algorithm>
#include <cstddef>
#include <filesystem>
#include <string>
#include <system_error>

// Writes a file through a temp file beside it, so a failed write leaves the file as it was. The temp file is created
// new, under a name no file has yet, so no other file is ever overwritten or removed
namespace AtomicFile {

    // The file a write ends up in: a symbolic link stays, the file it points to gets the content
    inline std::filesystem::path writeTarget(const std::filesystem::path& path)
    {
        std::error_code ec;
        if (!std::filesystem::is_symlink(path, ec)) return path;
        const std::filesystem::path real = std::filesystem::canonical(path, ec);
        return ec ? path : real;
    }

    // False if the content did not arrive: then the file is as it was, or, should Windows have removed it while
    // replacing, the content is kept in the temp file
    inline bool write(const std::filesystem::path& path, const std::string& bytes)
    {
        const std::filesystem::path target = writeTarget(path);
        std::wstring temp;
        HANDLE file = INVALID_HANDLE_VALUE;
        for (int n = 0; n < 100 && file == INVALID_HANDLE_VALUE; ++n) {
            temp = target.wstring() + (n ? L"." + std::to_wstring(n) : std::wstring()) + L".mr_tmp";
            file = ::CreateFileW(temp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (file == INVALID_HANDLE_VALUE && ::GetLastError() != ERROR_FILE_EXISTS) return false;
        }
        if (file == INVALID_HANDLE_VALUE) return false;

        bool written = true;
        for (std::size_t done = 0; written && done < bytes.size(); ) {
            const DWORD chunk = static_cast<DWORD>((std::min)(bytes.size() - done, std::size_t{ 1 } << 30));
            DWORD count = 0;
            written = ::WriteFile(file, bytes.data() + done, chunk, &count, nullptr) != FALSE && count == chunk;
            done += count;
        }
        written = ::CloseHandle(file) != FALSE && written;

        // ReplaceFileW keeps attributes and ACLs of the target; MoveFileExW covers a new file
        if (written && (::ReplaceFileW(target.c_str(), temp.c_str(), nullptr, 0, nullptr, nullptr)
            || ::MoveFileExW(temp.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING)))
            return true;

        if (!written || ::GetFileAttributesW(target.c_str()) != INVALID_FILE_ATTRIBUTES)
            ::DeleteFileW(temp.c_str());
        return false;
    }
}
