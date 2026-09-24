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

#include <utility>

namespace fs = std::filesystem;

namespace DirectoryWalk {

    Result collect(const fs::path& root, const Options& options)
    {
        Result result;
        std::error_code ec;
        fs::directory_iterator top(root, ec);
        if (ec) {
            result.status = Status::RootUnreadable;
            result.rootError = ec;
            return result;
        }

        // One open iterator per folder level, the innermost at the back
        std::vector<fs::directory_iterator> levels;
        levels.push_back(std::move(top));
        size_t seen = 0;

        while (!levels.empty()) {
            fs::directory_iterator& level = levels.back();
            if (level == fs::directory_iterator()) {
                levels.pop_back();
                continue;
            }

            // Copy and advance first: entering a subfolder invalidates level
            const fs::directory_entry entry = *level;
            level.increment(ec);
            if (ec) {                                   // the rest of this folder is lost
                ++result.unreadableFolders;
                level = fs::directory_iterator();       // MSVC leaves it on the failed entry, not at the end
            }

            if (options.keepGoing && !options.keepGoing(++seen)) {
                result.status = Status::Canceled;
                return result;
            }

            std::error_code ignored;   // an entry gone since the listing is neither folder nor file
            if (entry.symlink_status(ignored).type() == fs::file_type::directory) {
                if (!options.recurse || (options.enterFolder && !options.enterFolder(entry.path())))
                    continue;
                fs::directory_iterator sub(entry.path(), ec);
                if (ec) {
                    ++result.unreadableFolders;
                    continue;
                }
                levels.push_back(std::move(sub));
            }
            else if (entry.is_regular_file(ignored) && (!options.keepFile || options.keepFile(entry.path()))) {
                result.files.push_back(entry.path());
            }
        }
        return result;
    }

} // namespace DirectoryWalk
