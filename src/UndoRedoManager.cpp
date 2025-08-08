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

// -----------------------------------------------------------------------------
//  Centralised undo / redo stack – no plugin‑specific code in here
// -----------------------------------------------------------------------------

#include "UndoRedoManager.h"

// -----------------------------------------------------------------------------
//  Singleton
// -----------------------------------------------------------------------------
UndoRedoManager& UndoRedoManager::instance()
{
    static UndoRedoManager mgr;
    return mgr;
}

// -----------------------------------------------------------------------------
//  push – store new command and invalidate redo history
// -----------------------------------------------------------------------------
void UndoRedoManager::push(Action undoAction,
    Action redoAction,
    const std::wstring& label)
{
    _undo.push_back({ std::move(undoAction),
                      std::move(redoAction),
                      label });
    _redo.clear();
    trim();
}

// -----------------------------------------------------------------------------
//  undo – run last undo lambda and move it to redo stack
// -----------------------------------------------------------------------------
bool UndoRedoManager::undo()
{
    if (_undo.empty())
        return false;

    Item cmd = std::move(_undo.back());
    _undo.pop_back();

    if (cmd.undo)
        cmd.undo();

    _redo.push_back(std::move(cmd));
    return true;
}

// -----------------------------------------------------------------------------
//  redo – run last redo lambda and move it back to undo stack
// -----------------------------------------------------------------------------
bool UndoRedoManager::redo()
{
    if (_redo.empty())
        return false;

    Item cmd = std::move(_redo.back());
    _redo.pop_back();

    if (cmd.redo)
        cmd.redo();

    _undo.push_back(std::move(cmd));
    return true;
}

// -----------------------------------------------------------------------------
//  Helpers
// -----------------------------------------------------------------------------
void UndoRedoManager::clear()
{
    _undo.clear();
    _redo.clear();
}

std::wstring UndoRedoManager::peekUndoLabel() const
{
    return _undo.empty() ? L"" : _undo.back().label;
}

std::wstring UndoRedoManager::peekRedoLabel() const
{
    return _redo.empty() ? L"" : _redo.back().label;
}
