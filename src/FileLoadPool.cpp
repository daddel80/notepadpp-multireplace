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

#include "FileLoadPool.h"

#include <algorithm>
#include <chrono>
#include <new>
#include <utility>

#include "LiteralPrefilter.h"

namespace {
    // Buffers worth keeping for the next file: not tiny, and not so large that keeping them costs memory
    constexpr size_t REUSE_MIN = 4096;
    constexpr size_t REUSE_MAX = size_t{ 1 } << 20;

    bool reusable(const std::string& buffer)
    {
        return buffer.capacity() >= REUSE_MIN && buffer.capacity() <= REUSE_MAX;
    }

    // Longest wait in next(); the caller's message loop runs in between
    constexpr std::chrono::milliseconds NEXT_WAIT{ 50 };

    // Workers for LoadThreads=auto. Over the network every file waits for the server, so more reads at
    // once still pay off there; on local drives the gain ends at four.
    constexpr unsigned NETWORK_WORKERS = FileLoadPool::MAX_WORKERS;
    constexpr unsigned FAST_WORKERS = 4;
}

unsigned FileLoadPool::workersFor(int setting, const std::filesystem::path& root)
{
    if (setting >= 0) return (std::min)(static_cast<unsigned>(setting), MAX_WORKERS);
    switch (FileLoad::driveKind(root)) {
    case FileLoad::DriveKind::Network: return NETWORK_WORKERS;
    case FileLoad::DriveKind::Slow: return 1;
    case FileLoad::DriveKind::Fast: break;
    }
    return FAST_WORKERS;
}

FileLoadPool::FileLoadPool(std::vector<std::filesystem::path> files, const FileLoad::Options& options,
                           std::vector<Pattern> patterns, unsigned workerCount)
    : FileLoadPool(std::move(files), options, std::move(patterns), workerCount, Limits{})
{
}

FileLoadPool::FileLoadPool(std::vector<std::filesystem::path> files, const FileLoad::Options& options,
                           std::vector<Pattern> patterns, unsigned workerCount, Limits limits)
    : _files(std::move(files)), _options(options), _patterns(std::move(patterns)),
      _limits{ (std::max)(limits.ahead, size_t{ 1 }), limits.budget }, _slots(_limits.ahead)
{
    try {
        _threads.reserve(workerCount);
        _free.reserve(workerCount + 1);
        for (unsigned w = 0; w < workerCount; ++w)
            _threads.emplace_back([this] { work(); });
    }
    catch (...) {}   // fewer threads than asked for; with none the caller loads by itself
}

FileLoadPool::~FileLoadPool()
{
    cancel();
    for (std::thread& t : _threads) t.join();
}

FileLoadPool::Next FileLoadPool::next(File& file)
{
    std::unique_lock<std::mutex> lock(_m);
    if (_canceled || _taken >= _files.size()) return Next::Done;
    Slot& slot = _slots[_taken % _limits.ahead];
    if (!_ready.wait_for(lock, NEXT_WAIT, [&] { return _canceled || slot.ready; })) return Next::Pending;
    if (_canceled) return Next::Done;

    file = std::move(slot.file);
    _bytes -= slot.bytes;
    slot.ready = false;
    slot.bytes = 0;
    ++_taken;
    const bool turnWaiting = _waitingForTurn > 0;
    lock.unlock();

    _room.notify_one();
    if (turnWaiting) _turn.notify_all();
    return Next::Ready;
}

void FileLoadPool::recycle(std::string buffer)
{
    if (!reusable(buffer)) return;
    buffer.clear();
    std::lock_guard<std::mutex> lock(_m);
    if (_free.size() <= _threads.size()) _free.push_back(std::move(buffer));
}

void FileLoadPool::cancel()
{
    {
        std::lock_guard<std::mutex> lock(_m);
        _canceled = true;
        _stop.store(true, std::memory_order_relaxed);
    }
    _ready.notify_all();
    _room.notify_all();
    _turn.notify_all();
}

// A file may be read once it is the next one, or when it fits into the budget beside the others
bool FileLoadPool::waitTurn(size_t index, uint64_t bytes, uint64_t& reserved)
{
    std::unique_lock<std::mutex> lock(_m);
    auto mayRead = [&] { return _canceled || index == _taken || _bytes + bytes <= _limits.budget; };
    if (!mayRead()) {
        ++_waitingForTurn;
        _turn.wait(lock, mayRead);
        --_waitingForTurn;
    }
    if (_canceled) return false;
    _bytes += bytes;
    reserved = bytes;
    return true;
}

void FileLoadPool::prefilter(File& file) const
{
    if (_patterns.empty()) {
        file.handOver = true;
        return;
    }
    const bool utf8 = FileLoad::codepageFor(file.result.kind, file.result.enc) == SC_CP_UTF8;
    LiteralPrefilter literal(file.text);
    file.mayMatch.assign(_patterns.size(), 0);
    for (size_t i = 0; i < _patterns.size(); ++i) {
        const Pattern& pattern = _patterns[i];
        if (literal.mayMatch(utf8 ? pattern.utf8 : pattern.ansi, pattern.searchFlags)) {
            file.mayMatch[i] = 1;
            file.handOver = true;
        }
    }
}

void FileLoadPool::work()
{
    std::string spare;
    for (;;) {
        size_t index = 0;
        {
            std::unique_lock<std::mutex> lock(_m);
            _room.wait(lock, [&] { return _canceled || _claimed >= _files.size() || _claimed < _taken + _limits.ahead; });
            if (_canceled || _claimed >= _files.size()) return;
            index = _claimed++;
            if (!reusable(spare) && !_free.empty()) {
                spare = std::move(_free.back());
                _free.pop_back();
            }
        }

        File file;
        file.text = std::move(spare);
        spare = std::string();
        uint64_t reserved = 0;
        try {
            FileLoad::Pace pace;
            pace.beforeRead = [&](uint64_t fileSize) { return waitTurn(index, fileSize, reserved); };
            pace.stop = &_stop;
            file.result = FileLoad::load(_files[index], _options, file.text, &pace);
            if (file.result.skip == FileLoad::SkipReason::None)
                prefilter(file);
        }
        catch (const std::bad_alloc&) {
            file.result.skip = FileLoad::SkipReason::TooLarge;
            file.handOver = false;
        }
        catch (...) {
            file.result.skip = FileLoad::SkipReason::Unreadable;
            file.handOver = false;
        }

        // Nothing to hand over: the buffer stays with the worker
        if (!file.handOver) {
            if (reusable(file.text)) spare = std::move(file.text);
            file.text = std::string();
        }

        bool wakeConsumer = false;
        bool wakeTurn = false;
        {
            std::lock_guard<std::mutex> lock(_m);
            const uint64_t held = file.handOver ? file.text.capacity() : 0;
            _bytes = _bytes - reserved + held;
            wakeTurn = held < reserved && _waitingForTurn > 0;
            Slot& slot = _slots[index % _limits.ahead];
            slot.file = std::move(file);
            slot.bytes = held;
            slot.ready = true;
            wakeConsumer = (index == _taken);
        }
        if (wakeConsumer) _ready.notify_one();
        if (wakeTurn) _turn.notify_all();
    }
}
