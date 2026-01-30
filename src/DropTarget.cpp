// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
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

#include "DropTarget.h"
#include "MultiReplacePanel.h"

DropTarget::DropTarget(MultiReplace* parent)
    : _refCount(1), _parent(parent) {
}

HRESULT DropTarget::DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
    UNREFERENCED_PARAMETER(pDataObj);
    UNREFERENCED_PARAMETER(grfKeyState);
    UNREFERENCED_PARAMETER(pt);

    *pdwEffect = DROPEFFECT_COPY;  // Indicate that only copy operations are allowed
    return S_OK;
}

HRESULT DropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
    UNREFERENCED_PARAMETER(grfKeyState);
    UNREFERENCED_PARAMETER(pt);

    *pdwEffect = DROPEFFECT_COPY;  // Continue to allow only copy operations
    return S_OK;
}

HRESULT DropTarget::DragLeave() {
    // Handle the drag leave event with minimal overhead
    return S_OK;
}

HRESULT DropTarget::Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
    UNREFERENCED_PARAMETER(grfKeyState);
    UNREFERENCED_PARAMETER(pt);

    try {
        STGMEDIUM stgMedium;
        FORMATETC formatEtc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

        // Check if the data is in file format
        if (pDataObj->GetData(&formatEtc, &stgMedium) == S_OK) {
            HDROP hDrop = static_cast<HDROP>(GlobalLock(stgMedium.hGlobal));
            if (hDrop) {
                UINT numFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
                if (numFiles > 0) {
                    wchar_t filePath[MAX_PATH];
                    DragQueryFile(hDrop, 0, filePath, MAX_PATH);
                    _parent->loadListFromCsv(filePath);  // Load CSV file
                }
                GlobalUnlock(stgMedium.hGlobal); // Unlock the global memory object
            }
            ReleaseStgMedium(&stgMedium); // Release the storage medium
        }

        *pdwEffect = DROPEFFECT_COPY; // Indicate that the operation was a copy
        return S_OK;
    }
    catch (...) {
        // Handle errors silently
        *pdwEffect = DROPEFFECT_NONE;
        return E_FAIL;
    }
}

HRESULT DropTarget::QueryInterface(REFIID riid, void** ppvObject) {
    if (riid == IID_IUnknown || riid == IID_IDropTarget) {
        *ppvObject = this;
        AddRef();
        return S_OK;
    }
    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG DropTarget::AddRef() {
    return InterlockedIncrement(&_refCount);
}

ULONG DropTarget::Release() {
    ULONG count = InterlockedDecrement(&_refCount);
    if (count == 0) {
        delete this;
    }
    return count;
}