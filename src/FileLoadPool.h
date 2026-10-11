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

#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include "FileLoad.h"

// Find/Replace in Files: worker threads load, decode and prefilter the files of a run ahead of the
// UI thread, which takes them strictly in list order. Neither Scintilla nor the panel is used here.
// Ahead are at most `ahead` files and about `budget` bytes; a larger file is read once it is the next.
class FileLoadPool {
public:
    static constexpr unsigned MAX_WORKERS = 8;

    struct Limits {
        size_t ahead = 64;
        uint64_t budget = uint64_t{ 64 } << 20;
    };

    // A find text for the prefilter: the bytes Scintilla searches for in codepage 0 and in UTF-8
    struct Pattern {
        std::string ansi;
        std::string utf8;
        int searchFlags = 0;
    };

    // One file as the workers hand it over
    struct File {
        FileLoad::Result result;
        std::vector<char> mayMatch;   // per pattern, as LiteralPrefilter decides
        bool handOver = false;        // loaded, and no patterns or one that may match: text holds the file
        std::string text;
    };

    enum class Next { Ready, Pending, Done };

    // Workers for a run below root: a setting of 0 to MAX_WORKERS as it is; below 0 by the drive,
    // 8 on a network drive, 4 on a fast one, 1 on a slow one (FileLoad::DriveKind)
    static unsigned workersFor(int setting, const std::filesystem::path& root);

    // Starts the workers; workers() tells how many are running (none: the caller loads by itself)
    FileLoadPool(std::vector<std::filesystem::path> files, const FileLoad::Options& options,
                 std::vector<Pattern> patterns, unsigned workerCount);
    FileLoadPool(std::vector<std::filesystem::path> files, const FileLoad::Options& options,
                 std::vector<Pattern> patterns, unsigned workerCount, Limits limits);
    ~FileLoadPool();   // cancels and waits for the workers

    FileLoadPool(const FileLoadPool&) = delete;
    FileLoadPool& operator=(const FileLoadPool&) = delete;

    unsigned workers() const { return static_cast<unsigned>(_threads.size()); }

    // The next file in list order. Waits about 50 ms at most, then Pending; Done after the last file or a cancel
    Next next(File& file);

    // A buffer the caller is done with, for the workers to load into
    void recycle(std::string buffer);

    void cancel();

private:
    struct Slot {
        bool ready = false;
        uint64_t bytes = 0;   // counted against the budget while the file waits here
        File file;
    };

    void work();
    bool waitTurn(size_t index, uint64_t bytes, uint64_t& reserved);
    void prefilter(File& file) const;

    const std::vector<std::filesystem::path> _files;
    const FileLoad::Options _options;
    const std::vector<Pattern> _patterns;
    const Limits _limits;

    std::mutex _m;
    std::condition_variable _ready;   // the file the consumer waits for is ready, or cancel
    std::condition_variable _room;    // a place in the ring came free, or cancel
    std::condition_variable _turn;    // the consumer moved on or bytes came free, or cancel
    std::vector<Slot> _slots;         // ring: file i in slot i % ahead
    std::vector<std::string> _free;   // buffers to load into
    size_t _claimed = 0;              // files handed to workers
    size_t _taken = 0;                // files taken by the consumer
    uint64_t _bytes = 0;              // reserved by files being read and held by ready ones
    unsigned _waitingForTurn = 0;
    bool _canceled = false;
    std::atomic<bool> _stop{ false }; // _canceled for the reads in progress
    std::vector<std::thread> _threads;
};
