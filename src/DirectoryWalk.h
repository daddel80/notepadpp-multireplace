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

#include <cstddef>
#include <filesystem>
#include <functional>
#include <system_error>
#include <vector>

// Folder walk for Find/Replace in Files, in the order of recursive_directory_iterator
// (a folder's content right after the folder), links and junctions not entered.
// A folder that cannot be read is counted and skipped instead of ending the walk,
// which the iterator does on its first error. File system errors never throw.
namespace DirectoryWalk {

    struct Options {
        bool recurse = true;
        std::function<bool(const std::filesystem::path& folder)> enterFolder;   // subfolder may be entered (empty: every one)
        std::function<bool(const std::filesystem::path& file)> keepFile;        // regular file goes into the result (empty: every one)
        std::function<bool(size_t entriesSeen)> keepGoing;                      // called for every entry, false cancels (empty: never)
    };

    enum class Status { Done, Canceled, RootUnreadable };

    struct Result {
        Status status = Status::Done;
        std::error_code rootError;                   // why the root could not be listed (RootUnreadable)
        size_t unreadableFolders = 0;                // listed partly or not at all, the walk went on without them
        std::vector<std::filesystem::path> files;    // complete only with Done
    };

    Result collect(const std::filesystem::path& root, const Options& options);

} // namespace DirectoryWalk
