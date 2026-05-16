// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// MatchHistory.h
// Bounded per-match-state buffer that lets the formula engine reach
// across regex matches. The ring stores, for each past match that
// produced output:
//
//   - the text of every regex capture (including the full match at
//     slot 0), wrapped in a CaptureSlot that lazily caches its numeric
//     parse so repeated num() reads of the same slot are O(1);
//
//   - the output of every (?=...) block in segment order, as a
//     BlockOutput tagged union that distinguishes a numeric result
//     from a string result without implicit conversion.
//
// The module is deliberately engine-agnostic. It depends only on the
// shared parseNumber helper - no ExprTk, no ExprTkEngine, no Notepad++
// headers. That keeps the unit tests self-contained and the module
// reusable by other future engines that need the same cross-match
// reach (the Lua engine, for instance, could plug it in unchanged).

#pragma once

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "NumberParse.h"

namespace MultiReplaceEngine {

    // ---------------------------------------------------------------------
    // CaptureSlot - capture text + lazy numeric cache
    // ---------------------------------------------------------------------
    //
    // A capture is stored as text because both num() and txt() reads
    // are valid: txt() needs the original characters, num() needs the
    // parsed value. To avoid re-parsing on every num() read of the
    // same match we cache the parse result; the cache is invalidated
    // on assign().
    //
    // numCache and numCached are mutable because asNumber() is a const
    // observer that fills the cache on first call. A reader that holds
    // a `const CaptureSlot&` (history lookups return const access)
    // must still be able to drive the parse; mutable + const is the
    // standard C++ idiom for this lazy-computed-but-logically-const
    // pattern.
    struct CaptureSlot {
        std::string    text;
        mutable double numCache = 0.0;
        mutable bool   numCached = false;

        // Replace the slot's text. Capacity in `text` is retained
        // across calls (std::string::assign keeps the buffer when the
        // new content fits), so subsequent matches with similar-length
        // captures don't allocate.
        void assign(std::string_view s) {
            text.assign(s.data(), s.size());
            numCached = false;
        }

        // Lazy parse. First call performs the work; subsequent calls
        // return the cached result, including the NaN sentinel for
        // captures whose text is not a number.
        double asNumber() const {
            if (!numCached) {
                numCache = parseNumber(text);
                numCached = true;
            }
            return numCache;
        }

        // Const observer for the text. Returned by reference so the
        // caller can avoid copying for txt() reads.
        const std::string& asString() const { return text; }
    };


    // ---------------------------------------------------------------------
    // BlockOutput - tagged union for an Expression segment's result
    // ---------------------------------------------------------------------
    //
    // A (?=...) block can produce either a numeric value (the result
    // of the ExprTk expression) or a string (via ExprTk's return
    // statement). The two cases are stored side-by-side rather than
    // in a std::variant so a slot reused across matches retains the
    // string capacity even when the current match is numeric - and
    // vice versa. That capacity reuse is the main reason the steady-
    // state allocation count is zero.
    //
    // Type-mismatched reads (txt* against a Number slot, num* against
    // a String slot) return the caller's `v` fallback rather than
    // doing an implicit conversion. The whole point of the tagged
    // representation is to make that mismatch detectable at runtime.
    struct BlockOutput {
        enum Type : std::uint8_t { Number, String };

        Type        type = Number;
        double      numValue = 0.0;
        std::string strValue;

        void setNumber(double v) {
            type = Number;
            numValue = v;
            // strValue capacity is intentionally retained.
        }

        void setString(std::string_view s) {
            type = String;
            strValue.assign(s.data(), s.size());
        }
    };


    // ---------------------------------------------------------------------
    // MatchHistoryEntry - one row of the ring
    // ---------------------------------------------------------------------
    //
    // Sized once at engine compile (in MatchHistory's constructor) to
    // hold exactly the configured number of capture slots and block
    // outputs. After construction the two vectors never grow except
    // through resizeCaptureSlots, which is called at most once per
    // run.
    struct MatchHistoryEntry {
        std::vector<CaptureSlot> captures;
        std::vector<BlockOutput> blockOutputs;
    };


    // ---------------------------------------------------------------------
    // MatchHistory - the ring buffer
    // ---------------------------------------------------------------------
    //
    // Stores up to `depth` past matches in insertion order. Pushes are
    // O(1) and reuse the engine's working vectors via swap so no
    // allocation happens after construction. Lookback is O(1) index
    // math.
    //
    // Class invariants (maintained by every public mutator):
    //   - _filled <= _depth
    //   - _depth > 0  =>  _head in [0, _depth)
    //   - _ring.size() == _depth after construction (and after
    //     copy/move - but the class is non-copyable to keep the
    //     ownership of the contained string buffers unambiguous).
    //   - Each _ring[i].captures.size() equals the value of
    //     captureSlots passed at construction, modulo any subsequent
    //     resizeCaptureSlots call.
    //   - Each _ring[i].blockOutputs.size() equals the value of
    //     blockCount passed at construction.
    class MatchHistory {
    public:
        // Empty buffer: depth 0, no allocation. Valid for scripts
        // that have no history calls at all - pushSwap is a no-op
        // and every lookback returns null. The default constructor
        // gives the engine a clean object before compile() decides
        // how to size things.
        MatchHistory() = default;

        // Pre-allocate. After this call the ring is empty (size() ==
        // 0) but the underlying entries are already sized so the
        // first depth+1 pushes are guaranteed allocation-free.
        MatchHistory(std::size_t depth,
            std::size_t captureSlots,
            std::size_t blockCount);

        // Non-copyable, non-movable. The engine holds one instance
        // and resizes via assignment from a freshly-constructed
        // temporary; that path goes through move assignment which is
        // explicitly defaulted below. Copying a ring buffer is rarely
        // what anyone wants and would silently double the memory.
        MatchHistory(const MatchHistory&) = delete;
        MatchHistory& operator=(const MatchHistory&) = delete;
        MatchHistory(MatchHistory&&) = default;
        MatchHistory& operator=(MatchHistory&&) = default;

        // Push the engine's current per-match buffers into the ring
        // by swap. After return, the engine's vectors hold the
        // previous occupant of the freshly-written slot - irrelevant
        // for the engine, which is about to overwrite them for the
        // next match anyway, but importantly: their capacities are
        // preserved through this rotation, so the engine doesn't
        // re-allocate on the next per-match overwrite.
        //
        // No-op when depth == 0.
        void pushSwap(std::vector<CaptureSlot>& currentCaptures,
            std::vector<BlockOutput>& currentBlockOutputs);

        // Look back `p` matches. p == 1 is the most recent push,
        // p == 2 the one before that, and so on. Returns nullptr if
        // depth == 0, p == 0, or p exceeds the number of pushes
        // performed so far. The non-const overload exists only to
        // keep callers that already have non-const ownership of the
        // ring happy; both overloads return read-only data from the
        // perspective of the engine.
        MatchHistoryEntry* lookback(std::size_t p);
        const MatchHistoryEntry* lookback(std::size_t p) const;

        // Grow every ring entry's capture vector to `newSize`. Used
        // when the runtime regex turns out to provide more captures
        // than the compile-time literal analysis predicted (because
        // the user wrote num(myvar, 1) with a non-literal index).
        //
        // Growing is safe whether the ring is empty or not: the
        // backing std::vector::resize preserves existing slots and
        // default-constructs new ones at the tail. A previously
        // pushed entry that only populated its first N slots will
        // simply hold N populated + (newSize - N) empty slots, which
        // is the correct answer because that earlier match did not
        // capture the new positions either.
        //
        // Shrinking is permitted but not used by the engine in the
        // wild: the engine only ever grows during runtime refresh.
        // Shrinking a non-empty ring would silently discard data
        // from existing entries, which is the caller's responsibility
        // to avoid.
        void resizeCaptureSlots(std::size_t newSize);

        // Discard all stored matches. Entry storage and string
        // capacities inside entries are retained - subsequent pushes
        // overwrite in place without re-allocating.
        void clear() noexcept;

        std::size_t depth() const noexcept { return _depth; }
        std::size_t size()  const noexcept { return _filled; }

    private:
        std::size_t _depth = 0;
        std::size_t _filled = 0;
        std::size_t _head = 0;
        std::vector<MatchHistoryEntry> _ring;
    };

} // namespace MultiReplaceEngine