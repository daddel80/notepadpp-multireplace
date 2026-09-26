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
#include <utility>
#include <vector>

// Folder walk for Find/Replace in Files, in the order of recursive_directory_iterator
// (a folder's content right after the folder). Every folder is listed once: a subfolder
// is judged once from its listing entry, a file by its name, and only the files taken
// get a path. Links and junctions to folders are not entered, a link to a file counts
// as that file. A folder that cannot be read is counted and skipped instead of ending
// the walk. File system errors never throw.
namespace DirectoryWalk {

    // What a folder contributes: the files directly inside it, its subfolders
    struct Take {
        bool files = true;
        bool subfolders = true;
    };

    struct Options {
        Take root;                                                                          // what the root contributes
        std::function<Take(const std::filesystem::path& folder, bool hidden)> subfolder;   // a subfolder reached (empty: all of it)
        std::function<bool(const wchar_t* fileName)> keepFile;                             // a file of a folder whose files are taken (empty: every one)
        std::function<bool(size_t entriesSeen)> keepGoing;                                 // called for every entry, false cancels (empty: never)
    };

    enum class Status { Done, Canceled, RootUnreadable };

    struct Result {
        Status status = Status::Done;
        std::error_code rootError;                   // why the root could not be listed (RootUnreadable)
        size_t unreadableFolders = 0;                // listed partly or not at all, the walk went on without them
        std::vector<std::filesystem::path> files;    // complete only with Done
    };

    // The walk over the file system (FindFirstFileExW)
    Result collect(const std::filesystem::path& root, const Options& options);

    // One entry of a folder listing, . and .. left out
    struct Entry {
        const wchar_t* name = nullptr;   // valid until the listing moves on
        bool folder = false;             // a real folder, not a link or junction to one
        bool link = false;               // a symbolic link or junction
        bool hidden = false;
    };

    // The walk over any lister: lister.open(folder, listing, error) starts a folder's listing
    // (false: unreadable), listing.hasEntry/entry/next go through it (next false: the rest
    // is lost), lister.linksToFile(path) says where a link leads. collect() runs it on the
    // file system, the QA harness on a model with injected faults.
    template <class Lister>
    Result walk(Lister& lister, const std::filesystem::path& root, const Options& options)
    {
        struct Level {
            typename Lister::Listing listing;
            std::filesystem::path folder;
            Take take;
        };

        Result result;
        std::error_code error;

        // One open listing per folder level, the innermost at the back
        std::vector<Level> levels;
        levels.push_back({ {}, root, options.root });
        if (!lister.open(root, levels.back().listing, error)) {
            result.status = Status::RootUnreadable;
            result.rootError = error;
            return result;
        }
        if (!options.root.files && !options.root.subfolders)
            return result;

        size_t seen = 0;
        while (!levels.empty()) {
            Level& level = levels.back();
            if (!level.listing.hasEntry()) {
                levels.pop_back();
                continue;
            }

            if (options.keepGoing && !options.keepGoing(++seen)) {
                result.status = Status::Canceled;
                return result;
            }

            const Entry entry = level.listing.entry();
            if (entry.folder) {
                if (level.take.subfolders) {
                    std::filesystem::path path = level.folder / entry.name;
                    const Take take = options.subfolder ? options.subfolder(path, entry.hidden) : Take{};
                    if (take.files || take.subfolders) {
                        Level sub{ {}, std::move(path), take };
                        const bool opened = lister.open(sub.folder, sub.listing, error);
                        if (!level.listing.next()) ++result.unreadableFolders;   // advance before the push moves level
                        if (opened) levels.push_back(std::move(sub));
                        else ++result.unreadableFolders;
                        continue;
                    }
                }
            }
            else if (level.take.files && (!options.keepFile || options.keepFile(entry.name))) {
                std::filesystem::path path = level.folder / entry.name;
                if (!entry.link || lister.linksToFile(path))
                    result.files.push_back(std::move(path));
            }

            if (!level.listing.next()) ++result.unreadableFolders;   // the rest of this folder is lost
        }
        return result;
    }

} // namespace DirectoryWalk
