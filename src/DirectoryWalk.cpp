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

#include "DirectoryWalk.h"

#include <windows.h>

namespace fs = std::filesystem;

namespace DirectoryWalk {

    namespace {

        // Folder listings from the file system: one FindFirstFileExW per folder, entries
        // judged by the attributes that come with the listing
        class FileSystemLister {
        public:
            class Listing {
            public:
                Listing() = default;
                Listing(Listing&& other) noexcept
                    : _find(std::exchange(other._find, INVALID_HANDLE_VALUE)), _data(other._data), _hasEntry(other._hasEntry) {}
                Listing(const Listing&) = delete;
                Listing& operator=(const Listing&) = delete;
                Listing& operator=(Listing&&) = delete;
                ~Listing() { if (_find != INVALID_HANDLE_VALUE) ::FindClose(_find); }

                bool hasEntry() const { return _hasEntry; }

                Entry entry() const
                {
                    const bool link = isLink(_data);
                    return { _data.cFileName, (_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && !link, link,
                        (_data.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0 };
                }

                // False when the rest of the folder could not be read
                bool next()
                {
                    do {
                        if (!::FindNextFileW(_find, &_data)) {
                            _hasEntry = false;
                            return ::GetLastError() == ERROR_NO_MORE_FILES;
                        }
                    } while (isDotEntry(_data.cFileName));
                    return true;
                }

            private:
                friend class FileSystemLister;

                // Symbolic links and junctions; other reparse points (cloud placeholders,
                // deduplicated files) are ordinary files and folders, as for std::filesystem
                static bool isLink(const WIN32_FIND_DATAW& data)
                {
                    return (data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT)
                        && (data.dwReserved0 == IO_REPARSE_TAG_SYMLINK || data.dwReserved0 == IO_REPARSE_TAG_MOUNT_POINT);
                }

                static bool isDotEntry(const wchar_t* name)
                {
                    return name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'));
                }

                HANDLE _find = INVALID_HANDLE_VALUE;
                WIN32_FIND_DATAW _data{};
                bool _hasEntry = false;
            };

            // An empty folder lists nothing; false with the error when it cannot be listed,
            // which includes failing before the first entry after . and ..
            bool open(const fs::path& folder, Listing& listing, std::error_code& error) const
            {
                listing._find = ::FindFirstFileExW((folder / L"*").c_str(), FindExInfoBasic, &listing._data,
                    FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
                listing._hasEntry = (listing._find != INVALID_HANDLE_VALUE);
                if (listing._hasEntry && (!Listing::isDotEntry(listing._data.cFileName) || listing.next()))
                    return true;

                const DWORD code = ::GetLastError();
                if (code == ERROR_FILE_NOT_FOUND) return true;   // a drive root has no . and .. entries
                error = std::error_code(static_cast<int>(code), std::system_category());
                return false;
            }

            // A link counts as what it points to (std::filesystem::status): a file, not a
            // folder and not a link that leads nowhere
            bool linksToFile(const fs::path& link) const
            {
                const HANDLE h = ::CreateFileW(link.c_str(), FILE_READ_ATTRIBUTES,
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, nullptr);
                if (h == INVALID_HANDLE_VALUE) return false;
                BY_HANDLE_FILE_INFORMATION info{};
                const bool ok = ::GetFileInformationByHandle(h, &info) != FALSE;
                ::CloseHandle(h);
                return ok && !(info.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
            }
        };

    } // namespace

    Result collect(const fs::path& root, const Options& options)
    {
        FileSystemLister lister;
        return walk(lister, root, options);
    }

} // namespace DirectoryWalk
