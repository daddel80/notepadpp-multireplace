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

#pragma once
#include "StaticDialog/StaticDialog.h"
#include "DPIManager.h"
#include <CommCtrl.h>
#include <windowsx.h>
#include <vector>
#include <cstddef>

#define IDC_BTN_RESET 3001

class MultiReplaceConfigDialog : public StaticDialog
{
public:
    MultiReplaceConfigDialog() = default;
    ~MultiReplaceConfigDialog();

    void init(HINSTANCE hInst, HWND hParent);

protected:
    intptr_t CALLBACK run_dlgProc(UINT message, WPARAM wParam, LPARAM lParam) override;

private:
    // -------------------------------------------------------------------------
    // UI Creation & Management
    // -------------------------------------------------------------------------
    void createUI();
    void initUI();
    void resizeUI();
    void showCategory(int index);
    HWND createPanel();
    void applyInternalTheme();

    // -------------------------------------------------------------------------
    // Font & DPI Handling
    // -------------------------------------------------------------------------
    void createFonts();
    void cleanupFonts();
    void applyFonts();
    int scaleX(int value) const;
    int scaleY(int value) const;

    // -------------------------------------------------------------------------
    // Business Logic
    // -------------------------------------------------------------------------
    void resetToDefaults();
    void loadSettingsToConfigUI(bool reloadFile = true);
    void applyConfigToSettings();

    // -------------------------------------------------------------------------
    // Control Creation Helpers
    // -------------------------------------------------------------------------
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
        int id, int minValue, int maxValue, int tickMark = -1);

    // Message Handlers
    INT_PTR handleCtlColorStatic(WPARAM wParam, LPARAM lParam);

    // -------------------------------------------------------------------------
    // Specific Panel Builders
    // -------------------------------------------------------------------------
    void createSearchReplacePanelControls();
    void createListViewLayoutPanelControls();
    void createAppearancePanelControls();
    void createVariablesAutomationPanelControls();
    void createCsvOptionsPanelControls();

    // -------------------------------------------------------------------------
    // Members
    // -------------------------------------------------------------------------
    // UI Handles
    HWND _hCategoryList = nullptr;
    HWND _hCloseButton = nullptr;
    HWND _hResetButton = nullptr;

    HWND _hSearchReplacePanel = nullptr;
    HWND _hListViewLayoutPanel = nullptr;
    HWND _hAppearancePanel = nullptr;
    HWND _hVariablesAutomationPanel = nullptr;
    HWND _hImportScopePanel = nullptr;
    HWND _hCsvFlowTabsPanel = nullptr;

    // Subclass procedure for checkbox to notify dialog on click
    static LRESULT CALLBACK CheckboxSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

    // State
    int _currentCategory = -1;
    DPIManager* dpiMgr = nullptr;
    HFONT _hCategoryFont = nullptr;

    // Store user scale locally (from INI) to combine with System DPI
    double _userScaleFactor = 1.0;

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    // Simplified layout manager for grid-like positioning
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
            dlg->createGroupBox(parent, gx, gy, gw, gh, id, title);
            return LayoutBuilder(dlg, parent, gx + padX, gy + padTop, gw - 2 * padX, stepY);
        }

        void AddLabeledSlider(int labelId, const TCHAR* text, int sliderId,
            int sliderX, int sliderW, int minV, int maxV,
            int rowAdvance = 40,
            int labelW = 150,
            int sliderH = 18,
            int sliderYOffset = -4,
            int tickMark = -1)
        {
            int labelH = 18;
            AddLabel(labelId, text, labelW, labelH);
            HWND hTrack = dlg->createSlider(parent, x + sliderX, y + sliderYOffset, sliderW, sliderH, sliderId, minV, maxV, tickMark);
            (void)hTrack;
            y += rowAdvance;
        }
    };

    // Data Binding System
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