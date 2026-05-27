// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2025 Thomas Knoefel
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

#include "FileDialogUtil.h"

#include <new>
#include <shobjidl.h>

namespace {

    // Extract the first concrete extension from a filter spec like
    // "*.mrl;*.csv" -> ".mrl". Returns "" if no dot is present or the
    // first token ends with ".*" (a wildcard filter we never want to
    // append to a filename).
    std::wstring firstExtensionFromSpec(const std::wstring& spec)
    {
        const size_t dot = spec.find(L'.');
        if (dot == std::wstring::npos) return {};

        const size_t semi = spec.find(L';', dot + 1);
        std::wstring ext = (semi == std::wstring::npos)
            ? spec.substr(dot)
            : spec.substr(dot, semi - dot);

        if (ext == L".*") return {};
        return ext;
    }

    // Swap the trailing extension of `name` for `ext`. If `name` has no
    // extension, `ext` is appended. Returns true when the name actually
    // changed (so callers can skip a redundant SetFileName).
    bool swapExtension(std::wstring& name, const std::wstring& ext)
    {
        if (ext.empty()) return false;

        std::wstring stem;
        const size_t dot = name.find_last_of(L'.');
        if (dot != std::wstring::npos) {
            // Don't treat a dot in a folder segment as the file extension.
            const size_t sep = name.find_last_of(L"\\/");
            if (sep == std::wstring::npos || dot > sep)
                stem = name.substr(0, dot);
            else
                stem = name;
        }
        else {
            stem = name;
        }

        std::wstring next = stem + ext;
        if (next == name) return false;
        name.swap(next);
        return true;
    }

    // Event sink that handles the OnTypeChange callback: when the user
    // switches the file-type dropdown the filename's extension is
    // rewritten to match the new filter. Mirrors what Notepad++ does in
    // PowerEditor/src/WinControls/OpenSaveFileDialog/CustomFileDialog.cpp.
    class TypeFollowingEvents : public IFileDialogEvents
    {
    public:
        explicit TypeFollowingEvents(const std::vector<FileDialogUtil::Filter>& filters)
            : _refs(1), _filters(filters) {
        }

        // IUnknown
        IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override
        {
            if (!ppv) return E_INVALIDARG;
            *ppv = nullptr;
            if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileDialogEvents)) {
                *ppv = static_cast<IFileDialogEvents*>(this);
                AddRef();
                return S_OK;
            }
            return E_NOINTERFACE;
        }
        IFACEMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&_refs); }
        IFACEMETHODIMP_(ULONG) Release() override
        {
            const LONG r = InterlockedDecrement(&_refs);
            if (r == 0) delete this;
            return r;
        }

        // IFileDialogEvents
        IFACEMETHODIMP OnFileOk(IFileDialog*) override { return S_OK; }
        IFACEMETHODIMP OnFolderChanging(IFileDialog*, IShellItem*) override { return S_OK; }
        IFACEMETHODIMP OnFolderChange(IFileDialog*) override { return S_OK; }
        IFACEMETHODIMP OnSelectionChange(IFileDialog*) override { return S_OK; }
        IFACEMETHODIMP OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE*) override { return S_OK; }
        IFACEMETHODIMP OnOverwrite(IFileDialog*, IShellItem*, FDE_OVERWRITE_RESPONSE*) override { return S_OK; }

        IFACEMETHODIMP OnTypeChange(IFileDialog* dlg) override
        {
            if (!dlg) return S_OK;

            UINT one = 0;
            if (FAILED(dlg->GetFileTypeIndex(&one)) || one == 0) return S_OK;
            const size_t idx = static_cast<size_t>(one) - 1;
            if (idx >= _filters.size()) return S_OK;

            const std::wstring ext = firstExtensionFromSpec(_filters[idx].ext);
            if (ext.empty()) return S_OK;  // wildcard filter: leave the name alone

            // Keep the dialog's default extension in sync with the filter
            // so a name typed without an extension picks up the right one
            // (SetDefaultExtension wants the suffix without the leading dot).
            dlg->SetDefaultExtension(ext.c_str() + 1);

            PWSTR raw = nullptr;
            if (FAILED(dlg->GetFileName(&raw)) || !raw) return S_OK;
            std::wstring name = raw;
            CoTaskMemFree(raw);
            if (name.empty()) return S_OK;

            if (swapExtension(name, ext))
                dlg->SetFileName(name.c_str());

            return S_OK;
        }

    private:
        ~TypeFollowingEvents() = default;
        TypeFollowingEvents(const TypeFollowingEvents&) = delete;
        TypeFollowingEvents& operator=(const TypeFollowingEvents&) = delete;

        LONG                                  _refs;
        std::vector<FileDialogUtil::Filter>   _filters;
    };

    // Split a path-like "defaultPath" into folder + filename for the
    // dialog. The folder is applied via SetFolder() (which needs an
    // IShellItem of an existing directory); the filename via
    // SetFileName().
    void splitPathHint(const std::wstring& src,
        std::wstring& folderOut,
        std::wstring& nameOut)
    {
        if (src.empty()) { folderOut.clear(); nameOut.clear(); return; }

        const size_t sep = src.find_last_of(L"\\/");
        if (sep == std::wstring::npos) {
            folderOut.clear();
            nameOut = src;
            return;
        }
        folderOut = src.substr(0, sep);
        nameOut = src.substr(sep + 1);
    }

    // Run a Save or Open dialog with the shared parameter set. The
    // CLSID picks the flavor.
    FileDialogUtil::Result runDialog(REFCLSID clsid, const FileDialogUtil::Params& p)
    {
        FileDialogUtil::Result out;

        const HRESULT hrInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (FAILED(hrInit)) return out;
        const bool ownsCom = (hrInit == S_OK);  // S_FALSE means COM was already up here

        IFileDialog* dlg = nullptr;
        HRESULT hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&dlg));
        if (FAILED(hr) || !dlg) {
            if (ownsCom) CoUninitialize();
            return out;
        }

        // Filter spec.
        std::vector<COMDLG_FILTERSPEC> spec;
        spec.reserve(p.filters.size());
        for (const auto& f : p.filters)
            spec.push_back({ f.name.c_str(), f.ext.c_str() });
        if (!spec.empty())
            dlg->SetFileTypes(static_cast<UINT>(spec.size()), spec.data());

        const int idxClamped = (p.initialFilterIndex >= 0 &&
            p.initialFilterIndex < static_cast<int>(p.filters.size()))
            ? p.initialFilterIndex : 0;
        if (!p.filters.empty())
            dlg->SetFileTypeIndex(static_cast<UINT>(idxClamped) + 1);

        // Title + default name + default extension.
        if (!p.title.empty()) dlg->SetTitle(p.title.c_str());
        if (!p.defaultExtension.empty()) dlg->SetDefaultExtension(p.defaultExtension.c_str());

        std::wstring folder, name;
        splitPathHint(p.defaultPath, folder, name);
        if (!folder.empty()) {
            IShellItem* psi = nullptr;
            if (SUCCEEDED(SHCreateItemFromParsingName(folder.c_str(), nullptr,
                IID_PPV_ARGS(&psi))) && psi) {
                dlg->SetFolder(psi);
                psi->Release();
            }
        }
        if (!name.empty()) dlg->SetFileName(name.c_str());

        // Option flags.
        DWORD opts = 0;
        dlg->GetOptions(&opts);
        if (p.pathMustExist) opts |= FOS_PATHMUSTEXIST;
        if (p.fileMustExist) opts |= FOS_FILEMUSTEXIST;
        opts &= ~FOS_OVERWRITEPROMPT;  // callers run their own overwrite check on the final path
        dlg->SetOptions(opts);

        // Hook the events sink only when there's something to do (a
        // type change with a real extension to apply). new(nothrow) keeps
        // a failed allocation from leaking dlg on the way out.
        DWORD cookie = 0;
        bool adviseOk = false;
        IFileDialogEvents* events = nullptr;
        if (!p.filters.empty()) {
            events = new (std::nothrow) TypeFollowingEvents(p.filters);
            if (events)
                adviseOk = SUCCEEDED(dlg->Advise(events, &cookie));
        }

        hr = dlg->Show(p.owner);
        if (SUCCEEDED(hr)) {
            IShellItem* result = nullptr;
            if (SUCCEEDED(dlg->GetResult(&result)) && result) {
                PWSTR raw = nullptr;
                if (SUCCEEDED(result->GetDisplayName(SIGDN_FILESYSPATH, &raw)) && raw) {
                    out.path = raw;
                    CoTaskMemFree(raw);
                    out.ok = true;
                }
                result->Release();
            }
            UINT one = 0;
            if (SUCCEEDED(dlg->GetFileTypeIndex(&one)) && one > 0)
                out.filterIndex = static_cast<int>(one) - 1;
        }

        if (events) {
            if (adviseOk) dlg->Unadvise(cookie);
            events->Release();
        }
        dlg->Release();
        if (ownsCom) CoUninitialize();
        return out;
    }

} // namespace

namespace FileDialogUtil {

    Result showSave(const Params& p) { return runDialog(CLSID_FileSaveDialog, p); }
    Result showOpen(const Params& p) { return runDialog(CLSID_FileOpenDialog, p); }

} // namespace FileDialogUtil