// This file is part of Notepad++ project
// Copyright (C)2022 Don HO <don.h@free.fr>

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include <windows.h>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <algorithm>
#include <cstdint>
#include <string>

#include "Scintilla.h"

namespace NppStyleKit {

    // ---------- Theme ----------
    namespace ThemeUtils {
        // Query Notepad++ dark mode state (uses NPPM_ISDARKMODEENABLED).
        bool isDarkMode(HWND hNpp);
    }

    // ---------- Colors ----------
    namespace ColorTools {
        // Hash-based stable RGB color (0xRRGGBB) for a given string.
        uint32_t djb2Color(const std::string& s);
    }

    // ---------- IndicatorRegistry (color -> indicator id) ----------
    class IndicatorRegistry {
    public:
        // Initialize registry with usable indicator pool shared across editors.
        bool init(HWND hSciA,
            HWND hSciB,
            const std::vector<int>& poolUsable,
            int defaultAlpha = 100);

        // Get or assign indicator id for RGB color; returns -1 if none available.
        int  acquireForColor(uint32_t rgb);

        // Apply indicator range for given id.
        void fillRange(HWND hSci, int id, Sci_Position pos, Sci_Position len) const;

        // Clear all ranges for all pooled indicators in given editor.
        void clearAll(HWND hSci) const;

        // Clear internal color->id mapping; keeps pool/alpha.
        void resetColorMap();

    private:
        void ensureConfiguredOnBoth(int id, uint32_t rgb) const;
        void setStyleForId(HWND hSci, int id, uint32_t rgb) const;

        HWND hA_{ nullptr };
        HWND hB_{ nullptr };
        std::vector<int> pool_;
        int alpha_{ 100 };
        std::unordered_map<uint32_t, int> color2id_;
        size_t rotateIndex_ = 0;
    };

    // ---------- IndicatorCoordinator (usable/reserved ids) ----------
    class IndicatorCoordinator {
    public:
        // Build sanitized pool from preferred ids, excluding reservedInitial.
        bool init(HWND hSciA,
            HWND hSciB,
            const std::vector<int>& preferredIds,
            const std::vector<int>& reservedInitial);

        // Try preferredId first; otherwise first free usable id.
        int  reservePreferredOrFirstIndicator(const char* owner, int preferredId);

        // Remaining free indicator ids.
        std::vector<int> availableIndicatorPool() const;

        // Query if id is reserved.
        bool isIndicatorReserved(int id) const;

        // Re-init only if editors changed; returns true if re-init happened.
        bool ensureIndicatorsInitialized(HWND hSciA,
            HWND hSciB,
            const std::vector<int>& preferredIds,
            const std::vector<int>& reservedIds);

    private:
        bool usableOnBoth(int id) const;

        HWND hA_{ nullptr };
        HWND hB_{ nullptr };
        std::vector<int> pool_;
        std::unordered_map<std::string, int> owner2id_;
        std::unordered_set<int> used_;
    };

    // ---------- Globals ----------
    extern IndicatorCoordinator gIndicatorCoord;
    extern IndicatorRegistry    gIndicatorReg;
    extern int                  gColumnTabsIndicatorId;
    extern int                  gResultDockTrackingIndicatorId;  // INDIC_HIDDEN for position tracking

} // namespace NppStyleKit