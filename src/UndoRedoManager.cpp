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

#include "UndoRedoManager.h"

UndoRedoManager& UndoRedoManager::instance()
{
    static UndoRedoManager mgr;
    return mgr;
}

void UndoRedoManager::setActiveTab(int tabId)
{
    _activeTabId = tabId;
    // Implicitly create an empty Stacks entry for new tabs so the
    // const accessors can find it without modifying state.
    (void)_perTab[tabId];
}

void UndoRedoManager::removeTab(int tabId)
{
    _perTab.erase(tabId);
}

void UndoRedoManager::clearAll()
{
    _perTab.clear();
    (void)_perTab[_activeTabId];  // re-create empty entry for current tab
}

UndoRedoManager::Stacks& UndoRedoManager::active()
{
    return _perTab[_activeTabId];
}

const UndoRedoManager::Stacks& UndoRedoManager::active() const
{
    auto it = _perTab.find(_activeTabId);
    if (it != _perTab.end()) return it->second;

    static const Stacks empty;
    return empty;
}

void UndoRedoManager::setCapacity(size_t cap)
{
    _capacity = cap;
    for (auto& kv : _perTab) trim(kv.second);
}

void UndoRedoManager::trim(Stacks& s)
{
    if (_capacity == 0) return;
    while (s.undo.size() > _capacity) s.undo.pop_front();
    while (s.redo.size() > _capacity) s.redo.pop_front();
}

void UndoRedoManager::push(Action undoAction,
    Action redoAction,
    std::wstring label)
{
    Stacks& s = active();
    s.undo.push_back({ std::move(undoAction),
                       std::move(redoAction),
                       std::move(label) });
    s.redo.clear();
    trim(s);
}

bool UndoRedoManager::undo()
{
    Stacks& s = active();
    if (s.undo.empty())
        return false;

    Item cmd = std::move(s.undo.back());
    s.undo.pop_back();

    if (cmd.undo)
        cmd.undo();

    s.redo.push_back(std::move(cmd));
    return true;
}

bool UndoRedoManager::redo()
{
    Stacks& s = active();
    if (s.redo.empty())
        return false;

    Item cmd = std::move(s.redo.back());
    s.redo.pop_back();

    if (cmd.redo)
        cmd.redo();

    s.undo.push_back(std::move(cmd));
    return true;
}

void UndoRedoManager::clear()
{
    Stacks& s = active();
    s.undo.clear();
    s.redo.clear();
}

bool UndoRedoManager::canUndo() const
{
    return !active().undo.empty();
}

bool UndoRedoManager::canRedo() const
{
    return !active().redo.empty();
}

size_t UndoRedoManager::undoCount() const
{
    return active().undo.size();
}

size_t UndoRedoManager::redoCount() const
{
    return active().redo.size();
}

std::wstring UndoRedoManager::peekUndoLabel() const
{
    const Stacks& s = active();
    return s.undo.empty() ? L"" : s.undo.back().label;
}

std::wstring UndoRedoManager::peekRedoLabel() const
{
    const Stacks& s = active();
    return s.redo.empty() ? L"" : s.redo.back().label;
}