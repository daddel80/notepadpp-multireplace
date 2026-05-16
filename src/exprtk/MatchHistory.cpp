// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// MatchHistory.cpp
// Ring buffer mechanics. The class is tiny - three counters and a
// vector of entries - so the implementation is essentially just
// index arithmetic with two carefully-chosen guards (depth == 0
// short-circuits, modular indexing avoids unsigned underflow).

#include "MatchHistory.h"

#include <algorithm>
#include <utility>

namespace MultiReplaceEngine {

    MatchHistory::MatchHistory(std::size_t depth,
        std::size_t captureSlots,
        std::size_t blockCount)
        : _depth(depth)
    {
        if (_depth == 0) {
            return;
        }
        _ring.resize(_depth);
        for (auto& entry : _ring) {
            entry.captures.resize(captureSlots);
            entry.blockOutputs.resize(blockCount);
        }
    }

    void MatchHistory::pushSwap(std::vector<CaptureSlot>& currentCaptures,
        std::vector<BlockOutput>& currentBlockOutputs)
    {
        if (_depth == 0) {
            return;
        }

        // Write into the slot at _head, then advance. The swap moves the
        // engine's current per-match content into the ring and lifts the
        // ring's previous content (now stale) back into the engine's
        // working buffers. Both ends keep their allocated capacities.
        auto& slot = _ring[_head];
        std::swap(slot.captures, currentCaptures);
        std::swap(slot.blockOutputs, currentBlockOutputs);

        _head = (_head + 1) % _depth;
        if (_filled < _depth) {
            ++_filled;
        }
    }

    MatchHistoryEntry* MatchHistory::lookback(std::size_t p)
    {
        if (_depth == 0 || p == 0 || p > _filled) {
            return nullptr;
        }

        // _head points at the next write position, which is one past the
        // most recent entry. So the most recent entry is at (_head - 1)
        // mod _depth. We compute (_head + _depth - p) instead of
        // (_head - p) to avoid wrapping size_t around zero when p > _head.
        const std::size_t idx = (_head + _depth - p) % _depth;
        return &_ring[idx];
    }

    const MatchHistoryEntry* MatchHistory::lookback(std::size_t p) const
    {
        // Const path forwards through a const_cast - the non-const
        // overload performs no mutation, so this is safe and avoids
        // duplicating the index arithmetic.
        return const_cast<MatchHistory*>(this)->lookback(p);
    }

    void MatchHistory::resizeCaptureSlots(std::size_t newSize)
    {
        for (auto& entry : _ring) {
            entry.captures.resize(newSize);
        }
    }

    void MatchHistory::clear() noexcept
    {
        _filled = 0;
        _head = 0;
        // Entry contents (and their internal string capacities) stay put.
        // The next pushSwap overwrites the slot it lands on; entries it
        // never reaches simply hold their old strings invisibly. That
        // memory is reclaimed when the engine destroys or reassigns the
        // whole MatchHistory.
    }

} // namespace MultiReplaceEngine