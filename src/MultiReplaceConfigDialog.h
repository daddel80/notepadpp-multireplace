#pragma once
#include "StaticDialog/StaticDialog.h"
#include "DPIManager.h"
#include <CommCtrl.h>
#include <windowsx.h>
#include <vector>
#include <cstddef>

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
    HWND _hAppearancePanel = nullptr;
    HWND _hVariablesAutomationPanel = nullptr;
    HWND _hImportScopePanel = nullptr;
    HWND _hCsvFlowTabsPanel = nullptr;

    HWND _hCloseButton = nullptr;

    int _currentCategory = -1;
    DPIManager* dpiMgr = nullptr;

    // Store user scale locally since DPIManager is read-only/system-only
    double _userScaleFactor = 1.0;

    HFONT _hCategoryFont = nullptr;

    // ---- Minimal layout helper ----
    struct LayoutBuilder {
        MultiReplaceConfigDialog* dlg;
        HWND parent;
        int x, y, width;
        int stepY;
        LayoutBuilder(MultiReplaceConfigDialog* d, HWND p, int px, int py, int w, int step)
            : dlg(d), parent(p), x(px), y(py), width(w), stepY(step) {
        }

        void AddCheckbox(int id, const TCHAR* text) {
            dlg->createCheckBox(parent, x, y, width, id, text);
            y += stepY;
        }
        void AddLabel(int id, const TCHAR* text, int w = 160, int h = 18) {
            dlg->createStaticText(parent, x, y, w, h, id, text);
        }
        void AddNumberEdit(int id, int ex, int ey, int ew, int eh) {
            dlg->createNumberEdit(parent, x + ex, y + ey, ew, eh, id);
        }
        void AddSpace(int pixels) { y += pixels; }

        LayoutBuilder BeginGroup(int gx, int gy, int gw, int gh, int padX, int padTop, int id, const TCHAR* title) {
            // group container
            dlg->createGroupBox(parent, gx, gy, gw, gh, id, title);
            // inner content builder
            return LayoutBuilder(dlg, parent, gx + padX, gy + padTop, gw - 2 * padX, stepY);
        }
        void AddGroupBox(int id, const TCHAR* title, int gx, int gy, int gw, int gh) {
            dlg->createGroupBox(parent, gx, gy, gw, gh, id, title);
        }


        // slider helpers
        HWND AddSlider(int id, int ex, int ey, int ew, int eh, int minV, int maxV) {
            return dlg->createSlider(parent, x + ex, y + ey, ew, eh, id, minV, maxV);
        }
        void AddLabeledSlider(int labelId, const TCHAR* text, int sliderId,
            int sliderX, int sliderW, int minV, int maxV,
            int rowAdvance = 40,
            int labelW = 150, int labelH = 18,
            int sliderH = 26, int sliderYOffset = -4) {
            AddLabel(labelId, text, labelW, labelH);
            AddSlider(sliderId, sliderX, sliderYOffset, sliderW, sliderH, minV, maxV);
            y += rowAdvance;
        }
    };

    // ---- Binding (type-erased; implemented in .cpp) ----
    enum class ControlType { Checkbox, IntEdit };
    enum class ValueType { Bool, Int };
    struct Binding {
        HWND* panelHandlePtr;
        int controlID;
        ControlType control;
        ValueType value;
        std::size_t offset;
        int minVal;
        int maxVal;
    };
    std::vector<Binding> _bindings;
    bool _bindingsRegistered = false;
    void registerBindingsOnce();
    void applyBindingsToUI_Generic(void* settingsPtr);
    void readBindingsFromUI_Generic(void* settingsPtr);

};