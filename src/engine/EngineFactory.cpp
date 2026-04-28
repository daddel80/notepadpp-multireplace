// This file is part of MultiReplace.
//
// MultiReplace is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 2 of the License, or
// (at your option) any later version.
//
// EngineFactory.cpp
// Concrete engine selection. The factory is the only place in the
// codebase that knows about both LuaEngine and ExprTkEngine; every
// other module uses IFormulaEngine through this single entry point.

#include "EngineFactory.h"
#include "LuaEngine.h"

// ExprTkEngine.h gets included here once it lands. Until then, the
// ExprTk branch returns nullptr so calling code can detect the
// unimplemented case without crashing.
// #include "ExprTkEngine.h"

namespace MultiReplaceEngine {

    FormulaEnginePtr EngineFactory::create(EngineType type, ILuaEngineHost* host)
    {
        switch (type) {
        case EngineType::Lua:
            return std::make_unique<LuaEngine>(host);

        case EngineType::ExprTk:
            // Will be: return std::make_unique<ExprTkEngine>(host);
            // Returns nullptr until the ExprTk implementation is wired in;
            // the caller is expected to fall back to Lua or surface an
            // error. We deliberately don't silently substitute Lua here -
            // that would mask integration gaps.
            return nullptr;
        }

        // Unknown enum value (future-proofing for additional engine types).
        return nullptr;
    }

} // namespace MultiReplaceEngine