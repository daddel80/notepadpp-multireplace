#include "DropTarget.h"
#include "MultiReplacePanel.h"

DropTarget::DropTarget(HWND hwnd, MultiReplace* parent)
    : _refCount(1), _hwnd(hwnd), _parent(parent) {}

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
                    _parent->showListFilePath(); // Update List statistics
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
