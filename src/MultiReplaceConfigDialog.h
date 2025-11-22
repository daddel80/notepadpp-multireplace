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
    void applyPanelFonts();
    void resizeUI();
    void showCategory(int index);
    HWND createPanel();

    // DPI helpers
    int scaleX(int value) const;
    int scaleY(int value) const;

    // Control helpers for config UI
    HWND createGroupBox(HWND parent, int left, int top, int width, int height,
        int id, const TCHAR* text);
    HWND createCheckBox(HWND parent, int left, int top, int width,
        int id, const TCHAR* text);
    HWND createStaticText(HWND parent, int left, int top, int width, int height,
        int id, const TCHAR* text, DWORD extraStyle = 0);
    HWND createNumberEdit(HWND parent, int left, int top, int width, int height,
        int id);
    HWND createComboDropDownList(HWND parent, int left, int top, int width, int height,
        int id);
    HWND createTrackbarHorizontal(HWND parent, int left, int top, int width, int height,
        int id);
    HWND createSlider(HWND parent, int left, int top, int width, int height,
        int id, int minValue, int maxValue);
    INT_PTR handleCtlColorStatic(WPARAM wParam, LPARAM lParam);

    // Build controls for each category panel
    void createSearchReplacePanelControls();
    void createListViewLayoutPanelControls();
    void createAppearancePanelControls();
    void createVariablesAutomationPanelControls();
    void createImportScopePanelControls();
    void createCsvFlowTabsPanelControls();

    // Settings binding
    void loadSettingsFromConfig();
    void applyConfigToSettings();

    HWND _hCategoryList = nullptr;

    HWND _hSearchReplacePanel = nullptr;
    HWND _hListViewLayoutPanel = nullptr;
    HWND _hInterfaceTooltipsPanel = nullptr;
    HWND _hAppearancePanel = nullptr;
    HWND _hVariablesAutomationPanel = nullptr;
    HWND _hImportScopePanel = nullptr;
    HWND _hCsvFlowTabsPanel = nullptr;

    HWND _hCloseButton = nullptr;

    int _currentCategory = -1;
    DPIManager* dpiMgr = nullptr;

    HFONT _hCategoryFont = nullptr;
};
