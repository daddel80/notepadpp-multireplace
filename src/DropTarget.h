#ifndef DROP_TARGET_H
#define DROP_TARGET_H

#include <Windows.h>
#include <shlobj.h>  // Include for IDropTarget and other shell functions

class MultiReplace;  // Forward declaration to avoid circular dependency

class DropTarget : public IDropTarget {
public:
    DropTarget(HWND hwnd, MultiReplace* parent);
    virtual ~DropTarget() {}

    // Implement all IDropTarget methods
    HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override;
    HRESULT STDMETHODCALLTYPE DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override;
    HRESULT STDMETHODCALLTYPE DragLeave() override;
    HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override;

    // IUnknown methods
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;

private:
    LONG _refCount;
    HWND _hwnd;
    MultiReplace* _parent;
};

#endif // DROP_TARGET_H
