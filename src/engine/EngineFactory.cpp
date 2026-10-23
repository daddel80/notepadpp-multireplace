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
// EngineFactory.cpp
// Concrete engine selection. The factory is the only place in the
// codebase that knows about both LuaEngine and ExprTkEngine; every
// other module uses IFormulaEngine through this single entry point.

#include "EngineFactory.h"
#include "LuaEngine.h"
#include "ExprTkEngine.h"

namespace MultiReplaceEngine {

    FormulaEnginePtr EngineFactory::create(EngineType type, ILuaEngineHost* host)
    {
        switch (type) {
        case EngineType::Lua:
            return std::make_unique<LuaEngine>(host);

        case EngineType::ExprTk:
            return std::make_unique<ExprTkEngine>(host);
        }

        // Unknown enum value (future-proofing for additional engine types).
        return nullptr;
    }

} // namespace MultiReplaceEngine