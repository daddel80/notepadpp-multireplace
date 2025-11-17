#pragma once
#include "StaticDialog/StaticDialog.h"
#include "DPIManager.h"
#include <CommCtrl.h>
#include <windowsx.h>

class MultiReplaceConfigDialog : public StaticDialog
{
public:
    MultiReplaceConfigDialog() = default;
    ~MultiReplaceConfigDialog();

    void init(HINSTANCE hInst, HWND hParent);

protected:
    intptr_t CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam) override;

private:
    void createUI();
    void initUI();
    void showCategory(int index);
    void resizeUI();

    HWND _hCategoryList = nullptr;
    HWND _hGeneralPanel = nullptr;
    HWND _hResultPanel = nullptr;
    HWND _hCloseButton = nullptr;

    int _currentCategory = -1;
    DPIManager* dpiMgr = nullptr;
};
