// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025  Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#pragma once
// --------------------------------------------------------------
//  UndoRedoManager  (singleton)
//  • Keeps two stacks of lambda pairs (undo / redo)
//  • Totally framework‑agnostic – STL only
//  • Optional label field enables later history UI
// --------------------------------------------------------------

#include <functional>
#include <string>
#include <vector>

class UndoRedoManager
{
public:
    /// Global accessor – same idiom as ConfigManager / LanguageManager
    static UndoRedoManager& instance();

    using Action = std::function<void()>;

    void push(Action undoAction,
        Action redoAction,
        const std::wstring& label = L"");

    bool undo();          // returns false if nothing to undo
    bool redo();          // returns false if nothing to redo

    void clear();

    [[nodiscard]] bool   canUndo()   const { return !_undo.empty(); }
    [[nodiscard]] bool   canRedo()   const { return !_redo.empty(); }
    [[nodiscard]] size_t undoCount() const { return _undo.size(); }
    [[nodiscard]] size_t redoCount() const { return _redo.size(); }

    [[nodiscard]] std::wstring peekUndoLabel() const;
    [[nodiscard]] std::wstring peekRedoLabel() const;

private:
    UndoRedoManager() = default;
    ~UndoRedoManager() = default;
    UndoRedoManager(const UndoRedoManager&) = delete;
    UndoRedoManager& operator=(const UndoRedoManager&) = delete;

    struct Item {
        Action        undo;
        Action        redo;
        std::wstring  label;
    };

    std::vector<Item> _undo;
    std::vector<Item> _redo;
};
