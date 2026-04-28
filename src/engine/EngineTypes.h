// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// EngineTypes.h
// Shared data structures used across the formula engine boundary.
// These are intentionally engine-agnostic: every IFormulaEngine
// implementation receives FormulaVars and returns a FormulaResult,
// without exposing any Lua, ExprTk, or future engine specifics.

#pragma once

#include <string>
#include <vector>

namespace MultiReplaceEngine {

    // Identifies a concrete engine implementation. Persisted to INI as a
    // string (see engineTypeToString / engineTypeFromString) so future
    // additions don't shift magic numbers in user config files.
    enum class EngineType {
        Lua,
        ExprTk
    };

    // Per-match input variables that every engine receives. Mirrors the
    // surface that user scripts see: numeric counters, file metadata, the
    // matched text, and any regex captures.
    //
    // Engines map these into their own variable space (Lua globals,
    // ExprTk symbol_table, ...) but the input format is shared.
    struct FormulaVars {
        // Counters (1-based, populated by the replace pipeline)
        int CNT = 0;   // Replacement count across the whole run
        int LCNT = 0;   // Replacement count within the current line
        int LINE = 0;   // 1-based line number of the current match
        int LPOS = 0;   // Column position within the line (UTF-8 bytes)
        int APOS = 0;   // Absolute byte position in the document
        int COL = 0;   // CSV column index (CSV mode), 0 otherwise

        // Match metadata
        std::string MATCH;   // The matched text (UTF-8)
        std::string FPATH;   // Full path of the document being processed
        std::string FNAME;   // Filename without path

        // Regex captures CAP1..CAPn, populated only when the rule uses
        // regex search. Index 0 corresponds to CAP1 (CAP0 is intentionally
        // omitted; users address captures starting at 1).
        std::vector<std::string> captures;
    };

    // Result handed back from any engine after evaluating a script.
    // Engines never throw across this boundary; failures are reported
    // via success=false plus errorMessage.
    struct FormulaResult {
        std::string output;        // Resolved replacement text (UTF-8)
        bool        skip = false;  // True if the engine asked to skip
        // this match (Lua: skip(); ExprTk:
        // pattern signals no replacement)
        bool        success = true;
        std::string errorMessage;  // Populated when success == false
    };

    // String round-trip for INI persistence. Keeping these inline makes
    // the mapping obvious in one place.
    inline const wchar_t* engineTypeToString(EngineType t) {
        switch (t) {
        case EngineType::Lua:    return L"Lua";
        case EngineType::ExprTk: return L"ExprTk";
        }
        return L"Lua";
    }

    inline EngineType engineTypeFromString(const std::wstring& s) {
        if (s == L"ExprTk") return EngineType::ExprTk;
        return EngineType::Lua;  // Default for unknown / legacy entries
    }

} // namespace MultiReplaceEngine