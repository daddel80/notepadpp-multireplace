// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// MatchHistoryAnalysis.h
// Compile-time scan of a parsed replace template. Walks every
// Expression segment and finds calls to the six history-aware
// functions (num, txt, numout, txtout, numprev, txtprev). From those
// calls it determines:
//
//   - how deep the history ring needs to be (`maxLookback`),
//   - how many capture slots each ring entry must hold
//     (`maxCaptureIndex + 1` for full match at slot 0, plus all the
//     literal `n` values seen), with a flag set when a non-literal
//     `n` forces runtime resizing,
//   - how many (?=...) blocks the template contains (`blockCount`).
//
// The scan also enforces three compile-time errors:
//
//   - `p < 0` literal in any arity (including the no-args defaults).
//   - `p > HISTORY_HARD_CAP_DEPTH` literal: explicit lookback that
//     exceeds the hard cap, set so a typo like `numprev(1000000)`
//     doesn't quietly allocate a multi-megabyte ring.
//   - `n < 0` literal in arity 2/3 calls of num/txt (arity-1 keeps
//     the legacy tolerance for backward compatibility).
//
// The module has no engine dependency. It consumes only the
// ExprTkPatternParser's result struct.

#pragma once

#include <cstddef>
#include <string>

#include "ExprTkPatternParser.h"

namespace MultiReplaceEngine {

    // Depth used when the user wrote a non-literal expression in a
    // `p` position. The scanner can't know at compile time how far
    // back the runtime expression will reach, so the ring is sized
    // to this fallback. Reads beyond actual depth return v at runtime.
    constexpr std::size_t HISTORY_FALLBACK_DEPTH = 64;

    // Hard ceiling for any literal `p`. A user typing
    // `numprev(99999999)` is almost certainly a bug, not an intent
    // to allocate hundreds of megabytes. Caught at compile time with
    // a descriptive error.
    constexpr std::size_t HISTORY_HARD_CAP_DEPTH = 1024;

    struct HistoryAnalysis {
        // Maximum literal `p` seen across all history calls (with
        // HISTORY_FALLBACK_DEPTH counted whenever any call had a
        // non-literal `p`). Used to size the ring.
        std::size_t maxLookback = 0;

        // Maximum literal `n` seen across num/txt arity-2/3 calls
        // and across all arity-1 num/txt calls (where the legacy
        // tolerance lets us still pick up the slot-count hint, even
        // though negative values are silently ignored at runtime).
        // The ring's captures vector is sized to maxCaptureIndex + 1
        // for the slot-0 full-match entry.
        std::size_t maxCaptureIndex = 0;

        // Count of Expression segments. Each (?=...) becomes one
        // BlockOutput slot in the per-match buffer and in every
        // ring entry.
        std::size_t blockCount = 0;

        // True if at least one history function (num arity-2/3, txt
        // arity-2/3, or any of the four new functions) was found in
        // the template. False means the engine can skip the ring
        // entirely (depth 0 in the MatchHistory).
        bool hasHistory = false;

        // True if any num/txt call uses a non-literal expression in
        // the `n` position. When set, the engine cannot know at
        // compile time how many capture slots the runtime regex
        // will need, so it sizes for slot 0 only and grows on the
        // first execute() (see MatchHistory::resizeCaptureSlots).
        bool hasNonLiteralCaptureIdx = false;
    };

    // Analyse `parsed`. Returns the sizing summary. On compile error
    // writes a human-readable message into `errorOut` and stops
    // processing further calls; the partial result is still safe to
    // read but the caller is expected to check `errorOut.empty()`
    // before using it.
    HistoryAnalysis analyzeHistory(const ExprTkPatternParser::ParseResult& parsed,
        std::string& errorOut);

} // namespace MultiReplaceEngine