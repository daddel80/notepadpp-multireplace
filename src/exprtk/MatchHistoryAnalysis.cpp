// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// MatchHistoryAnalysis.cpp
// Byte-level scanner. Function names are ASCII, so multi-byte UTF-8
// characters in surrounding expression text flow through transparently
// (their bytes never collide with the ASCII letters we look for).

#include "MatchHistoryAnalysis.h"

#include <algorithm>
#include <cctype>
#include <cstring>
#include <string_view>
#include <vector>

namespace MultiReplaceEngine {

    namespace {

        // ---------------------------------------------------------------------
        // Function kinds and intermediate data
        // ---------------------------------------------------------------------

        enum class FuncKind {
            NumF,        // num(n) / num(n,p) / num(n,p,v)
            TxtF,        // txt(n) / txt(n,p) / txt(n,p,v)
            NumOutF,     // numout(n) / numout(n,p) / numout(n,p,v)
            TxtOutF,     // txtout(n) / txtout(n,p) / txtout(n,p,v)
            NumPrevF,    // numprev() / numprev(p) / numprev(p,v)
            TxtPrevF     // txtprev() / txtprev(p) / txtprev(p,v)
        };

        // A single argument as the scanner found it. `isLiteral` is true only
        // when the entire (trimmed) text is a decimal integer with an optional
        // leading minus sign - the analysis cares about literal `n` and `p`
        // because those are the only values it can use to size the buffer.
        struct ParsedArg {
            std::string text;
            bool        isLiteral = false;
            long long   value = 0;
        };

        struct CallInfo {
            FuncKind                kind = FuncKind::NumF;
            std::vector<ParsedArg>  args;
            std::size_t             sourcePos = 0;
        };


        // ---------------------------------------------------------------------
        // Name dispatch table
        // ---------------------------------------------------------------------
        //
        // Order matters: longer names must come first so the scanner matches
        // `numprev` before `num` at the same starting position. Returning the
        // length lets the caller advance past the name in one step.

        struct NameEntry {
            const char* name;
            std::size_t length;
            FuncKind    kind;
        };

        constexpr NameEntry NAME_TABLE[] = {
            { "numprev", 7, FuncKind::NumPrevF },
            { "txtprev", 7, FuncKind::TxtPrevF },
            { "numout",  6, FuncKind::NumOutF  },
            { "txtout",  6, FuncKind::TxtOutF  },
            { "num",     3, FuncKind::NumF     },
            { "txt",     3, FuncKind::TxtF     }
        };

        // Identifier characters as ExprTk understands them: ASCII letters,
        // digits, underscore. Anything else terminates an identifier.
        inline bool isIdentChar(char c)
        {
            const unsigned char uc = static_cast<unsigned char>(c);
            return (uc >= 'a' && uc <= 'z')
                || (uc >= 'A' && uc <= 'Z')
                || (uc >= '0' && uc <= '9')
                || uc == '_';
        }

        // Try to match one of the names at position `i` in `text`. Returns
        // the matching entry (or nullptr) and, when found, requires:
        //   - The byte before position i is either not present or is not an
        //     identifier character (so `mynum(...)` doesn't trip on `num`).
        //   - The byte immediately after the name is not an identifier
        //     character (so `numerical` doesn't trip on `num`).
        // ASCII tolower. We deliberately avoid std::tolower(int) here because
        // passing a signed char with a high bit set is undefined behaviour
        // under <cctype>. All our function names are pure ASCII, so a manual
        // fold is both safer and faster.
        inline char asciiToLower(char c)
        {
            const unsigned char uc = static_cast<unsigned char>(c);
            if (uc >= 'A' && uc <= 'Z') {
                return static_cast<char>(uc + ('a' - 'A'));
            }
            return c;
        }

        // Case-insensitive byte compare for ASCII names. NAME_TABLE entries
        // are stored in lowercase, so we only fold the input side.
        inline bool nameMatchesCI(const char* input, const char* name, std::size_t n)
        {
            for (std::size_t k = 0; k < n; ++k) {
                if (asciiToLower(input[k]) != name[k]) {
                    return false;
                }
            }
            return true;
        }

        const NameEntry* matchNameAt(std::string_view text, std::size_t i)
        {
            if (i > 0 && isIdentChar(text[i - 1])) {
                return nullptr;
            }
            for (const auto& entry : NAME_TABLE) {
                if (i + entry.length > text.size()) {
                    continue;
                }
                // Case-insensitive match: ExprTk's symbol_table uses
                // details::ilesscompare, so `Num(1, 1)` and `NUMPREV()` resolve
                // to our functions at runtime. The analyzer must follow the
                // same rule or it would silently skip those calls and the ring
                // would be sized at 0, making `Num(1, 1)` always return NaN.
                if (!nameMatchesCI(text.data() + i, entry.name, entry.length)) {
                    continue;
                }
                const std::size_t after = i + entry.length;
                if (after < text.size() && isIdentChar(text[after])) {
                    continue;
                }
                return &entry;
            }
            return nullptr;
        }


        // ---------------------------------------------------------------------
        // Argument parsing
        // ---------------------------------------------------------------------
        //
        // Starts at the byte immediately after the opening `(`. Walks forward
        // tracking paren depth and string-literal state, splitting on top-level
        // commas. Returns the position one past the matching `)`, or
        // std::string::npos if the parens were unbalanced (in which case the
        // caller treats the call as malformed and skips it - ExprTk will
        // report its own parse error downstream, so we don't double-report).

        constexpr std::size_t NPOS = std::string::npos;

        // Trim ASCII whitespace from both ends of a substring view.
        std::string_view trim(std::string_view s)
        {
            std::size_t lo = 0, hi = s.size();
            while (lo < hi && std::isspace(static_cast<unsigned char>(s[lo]))) {
                ++lo;
            }
            while (hi > lo && std::isspace(static_cast<unsigned char>(s[hi - 1]))) {
                --hi;
            }
            return s.substr(lo, hi - lo);
        }

        // Try parsing `s` as a decimal integer with optional leading minus.
        // Returns true on success and writes the value; false otherwise. We
        // do not use std::from_chars here because we want to reject inputs
        // that have *anything* trailing - a literal must be the entire arg.
        bool parseLiteralInt(std::string_view s, long long& out)
        {
            if (s.empty()) {
                return false;
            }
            std::size_t i = 0;
            bool negative = false;
            if (s[i] == '-') {
                negative = true;
                ++i;
                if (i == s.size()) {
                    return false;   // bare "-"
                }
            }
            long long value = 0;
            for (; i < s.size(); ++i) {
                const char c = s[i];
                if (c < '0' || c > '9') {
                    return false;
                }
                value = value * 10 + (c - '0');
                // Overflow guard. A literal that doesn't fit in long long is
                // almost certainly a typo; treat as non-literal so the engine
                // hits the fallback path with a sane size.
                if (value < 0) {
                    return false;
                }
            }
            out = negative ? -value : value;
            return true;
        }

        // Collect arguments from position `start` (just past the opening
        // paren) and return the position one past the matching `)`, or NPOS
        // on unbalanced input. `args` is appended to.
        std::size_t collectArgs(std::string_view text,
            std::size_t       start,
            std::vector<ParsedArg>& args)
        {
            std::size_t parenDepth = 1;
            char        stringQuote = 0;        // 0 = not in string
            std::size_t argStart = start;

            auto pushArg = [&](std::size_t end) {
                const std::string_view raw = text.substr(argStart, end - argStart);
                const std::string_view tr = trim(raw);
                ParsedArg a;
                a.text.assign(tr.data(), tr.size());
                if (parseLiteralInt(tr, a.value)) {
                    a.isLiteral = true;
                }
                args.push_back(std::move(a));
                };

            for (std::size_t i = start; i < text.size(); ++i) {
                const char c = text[i];

                if (stringQuote != 0) {
                    if (c == '\\' && i + 1 < text.size()) {
                        ++i;            // skip the escaped character
                        continue;
                    }
                    if (c == stringQuote) {
                        stringQuote = 0;
                    }
                    continue;
                }

                if (c == '"' || c == '\'') {
                    stringQuote = c;
                    continue;
                }
                if (c == '(') {
                    ++parenDepth;
                    continue;
                }
                if (c == ')') {
                    --parenDepth;
                    if (parenDepth == 0) {
                        // End of arg list. If we have any non-whitespace
                        // content since the last comma, that's the final arg;
                        // an empty (whitespace-only) tail means no-args call
                        // like `numprev()`.
                        const std::string_view tail = trim(text.substr(argStart, i - argStart));
                        if (!tail.empty() || !args.empty()) {
                            pushArg(i);
                        }
                        return i + 1;
                    }
                    continue;
                }
                if (c == ',' && parenDepth == 1) {
                    pushArg(i);
                    argStart = i + 1;
                }
            }
            return NPOS;
        }


        // ---------------------------------------------------------------------
        // Scanner
        // ---------------------------------------------------------------------
        //
        // Walks one expression segment's text and records every history-call
        // it can identify. After finding a call we resume scanning at the
        // position immediately *after* the opening paren (not after the
        // closing paren), which lets nested calls inside argument lists be
        // discovered as separate entries: in `num(1, num(2))` both `num`s
        // surface.

        void scanHistoryCalls(std::string_view text, std::vector<CallInfo>& out)
        {
            char stringQuote = 0;       // 0 outside a string, else the opening quote
            std::size_t i = 0;
            while (i < text.size()) {

                // String tracking has to happen at the top level too. Without
                // this, a function-name substring inside an outer string
                // literal - e.g., the `"num(99)"` argument of
                // `txt(0, 1, "num(99)")` - would be misidentified as a call.
                // collectArgs() also tracks string state when it descends into
                // a call's argument list, so the two layers stay consistent.
                if (stringQuote != 0) {
                    const char c = text[i];
                    if (c == '\\' && i + 1 < text.size()) {
                        i += 2;
                        continue;
                    }
                    if (c == stringQuote) {
                        stringQuote = 0;
                    }
                    ++i;
                    continue;
                }
                if (text[i] == '"' || text[i] == '\'') {
                    stringQuote = text[i];
                    ++i;
                    continue;
                }

                // Comment skipping. ExprTk's lexer recognises three styles
                // (//, #, /* ... */) and strips them before parsing. The
                // analyzer must do the same, otherwise commented-out history
                // calls would influence ring sizing - or worse, a literal
                // negative p inside a comment (`/* numprev(-1) */`) would
                // raise a compile-time error for code the engine itself would
                // silently skip.
                if (text[i] == '#') {
                    // Line comment until newline or end of segment.
                    ++i;
                    while (i < text.size() && text[i] != '\n') ++i;
                    continue;
                }
                if (text[i] == '/' && i + 1 < text.size()) {
                    if (text[i + 1] == '/') {
                        // Line comment.
                        i += 2;
                        while (i < text.size() && text[i] != '\n') ++i;
                        continue;
                    }
                    if (text[i + 1] == '*') {
                        // Block comment until matching */ or end of segment.
                        i += 2;
                        while (i + 1 < text.size()
                            && !(text[i] == '*' && text[i + 1] == '/')) {
                            ++i;
                        }
                        if (i + 1 < text.size()) {
                            i += 2;     // skip the closing */
                        }
                        else {
                            i = text.size();    // unterminated; bail out
                        }
                        continue;
                    }
                }

                const NameEntry* hit = matchNameAt(text, i);
                if (!hit) {
                    ++i;
                    continue;
                }

                // Skip whitespace between the name and the expected `(`.
                std::size_t afterName = i + hit->length;
                while (afterName < text.size() &&
                    std::isspace(static_cast<unsigned char>(text[afterName]))) {
                    ++afterName;
                }
                if (afterName >= text.size() || text[afterName] != '(') {
                    // A bare identifier (e.g., the user defined `num` as a
                    // variable, or it's part of `numerical`). Not a call.
                    ++i;
                    continue;
                }

                CallInfo call;
                call.kind = hit->kind;
                call.sourcePos = i;

                const std::size_t resume = collectArgs(text, afterName + 1, call.args);
                if (resume == NPOS) {
                    // Unbalanced parens. Stop scanning this segment; ExprTk's
                    // own parser will surface the error to the user.
                    return;
                }
                out.push_back(std::move(call));

                // Resume *after the opening paren*, not at `resume`. That way
                // a nested `num(2)` inside an outer `num(1, num(2))` is
                // discovered in its own right.
                i = afterName + 1;
            }
        }


        // ---------------------------------------------------------------------
        // Per-call contribution to the analysis
        // ---------------------------------------------------------------------
        //
        // Returns false if a compile error was raised (caller stops further
        // processing).

        const char* kindName(FuncKind k)
        {
            switch (k) {
            case FuncKind::NumF:     return "num";
            case FuncKind::TxtF:     return "txt";
            case FuncKind::NumOutF:  return "numout";
            case FuncKind::TxtOutF:  return "txtout";
            case FuncKind::NumPrevF: return "numprev";
            case FuncKind::TxtPrevF: return "txtprev";
            }
            return "?";
        }

        // Where does the `p` argument live (which index in args), and what's
        // the default if it's omitted? `n` is always at index 0 when present
        // (only for num/txt/numout/txtout); numprev/txtprev have no n.

        struct ArgLayout {
            bool        hasN;
            std::size_t pIndex;       // index of p in args (when present)
            std::size_t defaultP;     // value of p when args.size() <= pIndex
        };

        ArgLayout layoutOf(FuncKind k)
        {
            switch (k) {
            case FuncKind::NumF:
            case FuncKind::TxtF:
                return { /*hasN*/ true, /*pIndex*/ 1, /*defaultP*/ 0 };
            case FuncKind::NumOutF:
            case FuncKind::TxtOutF:
                return { /*hasN*/ true, /*pIndex*/ 1, /*defaultP*/ 1 };
            case FuncKind::NumPrevF:
            case FuncKind::TxtPrevF:
                return { /*hasN*/ false, /*pIndex*/ 0, /*defaultP*/ 1 };
            }
            return { false, 0, 0 };
        }

        // Format a "in `fname(...)` at position N" suffix for error messages.
        // Cheap concatenation; never on the hot path.
        std::string errorSuffix(FuncKind k, std::size_t pos)
        {
            std::string s = " in `";
            s += kindName(k);
            s += "(...)` at position ";
            s += std::to_string(pos);
            return s;
        }

        bool processCall(const CallInfo& call,
            HistoryAnalysis& out,
            std::string& errorOut)
        {
            const ArgLayout layout = layoutOf(call.kind);
            const bool      isArity1 = layout.hasN && call.args.size() == 1;

            // -------- p position --------
            std::size_t effectiveP = layout.defaultP;
            if (call.args.size() > layout.pIndex) {
                const ParsedArg& pArg = call.args[layout.pIndex];
                if (pArg.isLiteral) {
                    if (pArg.value < 0) {
                        errorOut = "p must be non-negative";
                        errorOut += errorSuffix(call.kind, call.sourcePos);
                        return false;
                    }
                    if (static_cast<std::size_t>(pArg.value) > HISTORY_HARD_CAP_DEPTH) {
                        errorOut = "p exceeds hard cap (";
                        errorOut += std::to_string(HISTORY_HARD_CAP_DEPTH);
                        errorOut += ")";
                        errorOut += errorSuffix(call.kind, call.sourcePos);
                        return false;
                    }
                    effectiveP = static_cast<std::size_t>(pArg.value);
                }
                else {
                    effectiveP = HISTORY_FALLBACK_DEPTH;
                }
            }

            if (effectiveP > out.maxLookback) {
                out.maxLookback = effectiveP;
            }

            // -------- n position (only for num/txt/numout/txtout) --------
            if (layout.hasN && !call.args.empty()) {
                const ParsedArg& nArg = call.args[0];
                if (nArg.isLiteral) {
                    if (nArg.value < 0) {
                        // Backward-compat carve-out: arity-1 num(-1) / txt(-1)
                        // is silently tolerated at runtime (returns 0 / "")
                        // because some long-standing rules in the wild rely
                        // on it. For arity-2/3 we're stricter - if you're
                        // opting into the new API, you have to use it
                        // correctly.
                        if (!isArity1) {
                            errorOut = "n must be non-negative";
                            errorOut += errorSuffix(call.kind, call.sourcePos);
                            return false;
                        }
                        // Negative arity-1 literal: ignored for sizing.
                    }
                    else {
                        // maxCaptureIndex only governs the per-entry capture-
                        // slot count. For num/txt, n IS the capture index. For
                        // numout/txtout, n is a BLOCK index (addresses a
                        // different array entirely), so it has no business
                        // sizing the capture slots.
                        if (call.kind == FuncKind::NumF || call.kind == FuncKind::TxtF) {
                            const std::size_t idx = static_cast<std::size_t>(nArg.value);
                            if (idx > out.maxCaptureIndex) {
                                out.maxCaptureIndex = idx;
                            }
                        }
                    }
                }
                else {
                    // Only num/txt care about hasNonLiteralCaptureIdx:
                    // numout/txtout treat `n` as a block index where
                    // out-of-range simply returns v at runtime.
                    if (call.kind == FuncKind::NumF || call.kind == FuncKind::TxtF) {
                        out.hasNonLiteralCaptureIdx = true;
                    }
                }
            }

            // -------- hasHistory --------
            // Arity-1 num/txt are the legacy single-match readers - no ring
            // needed for them alone. Everything else either reaches into the
            // ring directly or has the potential to (effectiveP could still
            // be 0 for, e.g., numout(0, 0); but the engine code path goes
            // through the history-aware function classes, and the ring being
            // size 0 is fine for those).
            if (call.kind == FuncKind::NumF || call.kind == FuncKind::TxtF) {
                if (call.args.size() >= 2) {
                    out.hasHistory = true;
                }
            }
            else {
                out.hasHistory = true;
            }

            return true;
        }

    } // anonymous namespace


    // ---------------------------------------------------------------------
    // Public entry point
    // ---------------------------------------------------------------------

    HistoryAnalysis analyzeHistory(const ExprTkPatternParser::ParseResult& parsed,
        std::string& errorOut)
    {
        errorOut.clear();
        HistoryAnalysis result;

        for (const auto& seg : parsed.segments) {
            if (seg.type != ExprTkPatternParser::SegmentType::Expression) {
                continue;
            }
            ++result.blockCount;

            std::vector<CallInfo> calls;
            scanHistoryCalls(seg.text, calls);

            for (const auto& call : calls) {
                if (!processCall(call, result, errorOut)) {
                    // Stop on first error - the message is already set.
                    // Partial result fields are still consistent (we only
                    // wrote to maxLookback / maxCaptureIndex on success).
                    return result;
                }
            }
        }

        return result;
    }

} // namespace MultiReplaceEngine