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
//
// EngineFactory.h
// Single entry point for creating concrete IFormulaEngine instances.
// Keeps the rest of the codebase (TabState, the replace pipeline)
// decoupled from concrete engine headers - callers include only this
// factory header and the abstract IFormulaEngine.

#pragma once

#include "IFormulaEngine.h"

namespace MultiReplaceEngine {

    // Forward declaration so the factory header doesn't pull in <lua.hpp>
    // or other engine-specific includes.
    class ILuaEngineHost;

    class EngineFactory {
    public:
        // Construct a fresh engine of the given type. The returned engine
        // is uninitialised; the caller is expected to invoke initialize()
        // before the first compile/execute. Returns nullptr on an
        // unrecognised type.
        //
        // The host pointer is forwarded to engines that need to call back
        // into MR (Lua needs UI hooks; ExprTk currently does not but the
        // parameter is accepted for symmetry and future use).
        //
        // The host must outlive every engine produced from this call.
        static FormulaEnginePtr create(EngineType type, ILuaEngineHost* host);

        EngineFactory() = delete;  // static-only utility class
    };

} // namespace MultiReplaceEngine