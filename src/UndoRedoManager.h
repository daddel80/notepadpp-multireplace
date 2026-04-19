// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025  Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#pragma once
// --------------------------------------------------------------
//  UndoRedoManager  (singleton)
//  • Keeps undo/redo stacks per tab id
//  • push/undo/redo/clear operate on the currently active tab
//  • Totally framework-agnostic - STL only
// --------------------------------------------------------------

#include <deque>
#include <functional>
#include <string>
#include <unordered_map>

class UndoRedoManager
{
public:
    static UndoRedoManager& instance();

    using Action = std::function<void()>;

    // Tab management. Default active tab id is 0, which matches the
    // default id used by the primary tab, so callers that never touch
    // these methods keep working as a single-tab setup.
    void setActiveTab(int tabId);
    void removeTab(int tabId);
    void clearAll();

    // Operate on the active tab.
    void push(Action undoAction,
        Action redoAction,
        std::wstring label = L"");

    bool undo();          // returns false if nothing to undo
    bool redo();          // returns false if nothing to redo
    void clear();         // clears only the active tab

    [[nodiscard]] bool   canUndo()   const;
    [[nodiscard]] bool   canRedo()   const;
    [[nodiscard]] size_t undoCount() const;
    [[nodiscard]] size_t redoCount() const;

    [[nodiscard]] std::wstring peekUndoLabel() const;
    [[nodiscard]] std::wstring peekRedoLabel() const;

    void   setCapacity(size_t cap);
    size_t capacity() const { return _capacity; }

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

    struct Stacks {
        std::deque<Item> undo;   // deque for O(1) pop_front during trim
        std::deque<Item> redo;
    };

    // Default capacity of 200 steps is suitable for typical usage.
    // Setting capacity to 0 means unlimited (use with caution).
    size_t _capacity = 200;

    int _activeTabId = 0;
    std::unordered_map<int, Stacks> _perTab;

    Stacks& active();
    const Stacks& active() const;
    void          trim(Stacks& s);
};