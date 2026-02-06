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

#include "NppStyleKit.h"
#include "Notepad_plus_msgs.h"
#include "Scintilla.h"

namespace NppStyleKit {

    // ---------- Globals ----------
    IndicatorCoordinator gIndicatorCoord;
    IndicatorRegistry    gIndicatorReg;
    int                  gColumnTabsIndicatorId = -1;
    int                  gResultDockTrackingIndicatorId = -1;

    // --- Scintilla direct ----------------------------------------------------
    // Local helper: fast Scintilla calls with per-thread cached direct function.
    static inline sptr_t S(HWND hSci, UINT m, uptr_t w = 0, sptr_t l = 0)
    {
        static thread_local HWND        cachedHwnd = nullptr;
        static thread_local SciFnDirect cachedFn = nullptr;
        static thread_local sptr_t      cachedPtr = 0;

        if (!hSci)
            return 0;

        if (hSci != cachedHwnd || !cachedFn || !cachedPtr)
        {
            cachedFn = reinterpret_cast<SciFnDirect>(
                ::SendMessage(hSci, SCI_GETDIRECTFUNCTION, 0, 0));
            cachedPtr = (sptr_t)::SendMessage(hSci, SCI_GETDIRECTPOINTER, 0, 0);
            cachedHwnd = hSci;
        }

        return cachedFn
            ? cachedFn(cachedPtr, m, w, l)
            : ::SendMessage(hSci, m, (WPARAM)w, (LPARAM)l);
    }

    // ---------- internals ----------
    static bool indicatorUsable(HWND hSci, int id)
    {
        const sptr_t orig = S(hSci, SCI_INDICGETSTYLE, (uptr_t)id);
        S(hSci, SCI_INDICSETSTYLE, (uptr_t)id, INDIC_STRAIGHTBOX);
        const sptr_t now = S(hSci, SCI_INDICGETSTYLE, (uptr_t)id);
        S(hSci, SCI_INDICSETSTYLE, (uptr_t)id, orig);
        return (now == INDIC_STRAIGHTBOX);
    }

    static std::vector<int> pruneForEditor(HWND hSci,
        const std::vector<int>& in,
        const std::vector<int>& reserved)
    {
        std::vector<int> out;
        out.reserve(in.size());
        for (int id : in)
        {
            if (std::find(reserved.begin(), reserved.end(), id) != reserved.end())
                continue;
            if (indicatorUsable(hSci, id))
                out.push_back(id);
        }
        return out;
    }

    static std::vector<int> buildUsablePool(HWND hSciA,
        HWND hSciB,
        const std::vector<int>& preferred,
        const std::vector<int>& reservedIds)
    {
        auto poolA = pruneForEditor(hSciA, preferred, reservedIds);
        if (!hSciB) return poolA;

        auto poolB = pruneForEditor(hSciB, preferred, reservedIds);

        std::vector<int> out;
        out.reserve((std::min)(poolA.size(), poolB.size()));
        for (int id : poolA)
        {
            if (std::find(poolB.begin(), poolB.end(), id) != poolB.end())
                out.push_back(id);
        }
        return out;
    }

    // ---------- ThemeUtils ----------
    bool ThemeUtils::isDarkMode(HWND hNpp)
    {
        return (::SendMessage(hNpp, NPPM_ISDARKMODEENABLED, 0, 0) != 0);
    }

    // ---------- ColorTools ----------
    uint32_t ColorTools::djb2Color(const std::string& s)
    {
        unsigned long h = 5381ul;
        for (unsigned char c : s)
            h = ((h << 5) + h) + c;

        const int r = (int)((h >> 16) & 0xFF);
        const int g = (int)((h >> 8) & 0xFF);
        const int b = (int)(h & 0xFF);

        return (uint32_t)((r << 16) | (g << 8) | b);
    }

    // ---------- IndicatorRegistry ----------
    bool IndicatorRegistry::init(HWND hSciA,
        HWND hSciB,
        const std::vector<int>& poolUsable,
        int defaultAlpha)
    {
        hA_ = hSciA;
        hB_ = hSciB;
        pool_ = poolUsable;
        alpha_ = defaultAlpha;
        color2id_.clear();
        return (!pool_.empty() && hA_);
    }

    void IndicatorRegistry::setStyleForId(HWND hSci, int id, uint32_t rgb) const
    {
        if (!hSci || id < 0) return;

        S(hSci, SCI_INDICSETSTYLE, (uptr_t)id, INDIC_STRAIGHTBOX);
        S(hSci, SCI_INDICSETFORE, (uptr_t)id, (sptr_t)rgb);
        S(hSci, SCI_INDICSETALPHA, (uptr_t)id, (sptr_t)alpha_);
    }

    void IndicatorRegistry::ensureConfiguredOnBoth(int id, uint32_t rgb) const
    {
        if (hA_) setStyleForId(hA_, id, rgb);
        if (hB_) setStyleForId(hB_, id, rgb);
    }

    int IndicatorRegistry::acquireForColor(uint32_t rgb)
    {
        // 1) Existing mapping: stable color for same word
        auto it = color2id_.find(rgb);
        if (it != color2id_.end())
            return it->second;

        if (pool_.empty())
            return -1;

        // 2) First pass: use a free id from pool
        for (int id : pool_) {
            bool used = false;
            for (const auto& kv : color2id_) {
                if (kv.second == id) {
                    used = true;
                    break;
                }
            }
            if (!used) {
                ensureConfiguredOnBoth(id, rgb);
                color2id_.emplace(rgb, id);
                return id;
            }
        }

        // 3) Pool exhausted: cyclic reuse like old implementation
        if (rotateIndex_ >= pool_.size())
            rotateIndex_ = 0;

        const int id = pool_[rotateIndex_++];

        // Drop previous color bound to this id (we reassign)
        for (auto it2 = color2id_.begin(); it2 != color2id_.end(); ++it2) {
            if (it2->second == id) {
                color2id_.erase(it2);
                break;
            }
        }

        ensureConfiguredOnBoth(id, rgb);
        color2id_[rgb] = id;
        return id;
    }


    void IndicatorRegistry::fillRange(HWND hSci, int id, Sci_Position pos, Sci_Position len) const
    {
        if (!hSci || id < 0 || len <= 0)
            return;

        S(hSci, SCI_SETINDICATORCURRENT, (uptr_t)id);
        S(hSci, SCI_INDICATORFILLRANGE, (uptr_t)pos, (sptr_t)len);
    }

    void IndicatorRegistry::clearAll(HWND hSci) const
    {
        if (!hSci)
            return;

        const Sci_Position len =
            (Sci_Position)S(hSci, SCI_GETLENGTH);

        for (int id : pool_)
        {
            S(hSci, SCI_SETINDICATORCURRENT, (uptr_t)id);
            S(hSci, SCI_INDICATORCLEARRANGE, 0, (sptr_t)len);
        }
    }

    void IndicatorRegistry::resetColorMap()
    {
        color2id_.clear();
    }

    // ---------- IndicatorCoordinator ----------
    bool IndicatorCoordinator::init(HWND hSciA,
        HWND hSciB,
        const std::vector<int>& preferredIds,
        const std::vector<int>& reservedInitial)
    {
        hA_ = hSciA;
        hB_ = hSciB;

        pool_.clear();
        used_.clear();
        owner2id_.clear();

        auto usable = buildUsablePool(hA_, hB_, preferredIds, reservedInitial);
        pool_.assign(usable.begin(), usable.end());

        for (int id : reservedInitial)
            used_.insert(id);

        return !pool_.empty();
    }

    bool IndicatorCoordinator::usableOnBoth(int id) const
    {
        if (!hA_) return false;
        if (!indicatorUsable(hA_, id)) return false;
        if (hB_ && !indicatorUsable(hB_, id)) return false;
        return true;
    }

    int IndicatorCoordinator::reservePreferredOrFirstIndicator(const char* owner, int preferredId)
    {
        const std::string key = owner ? owner : "";

        if (preferredId >= 0 &&
            std::find(pool_.begin(), pool_.end(), preferredId) != pool_.end() &&
            used_.find(preferredId) == used_.end() &&
            usableOnBoth(preferredId))
        {
            used_.insert(preferredId);
            owner2id_[key] = preferredId;
            return preferredId;
        }

        for (int id : pool_)
        {
            if (used_.find(id) == used_.end() && usableOnBoth(id))
            {
                used_.insert(id);
                owner2id_[key] = id;
                return id;
            }
        }
        return -1;
    }

    std::vector<int> IndicatorCoordinator::availableIndicatorPool() const
    {
        std::vector<int> out;
        out.reserve(pool_.size());
        for (int id : pool_)
        {
            if (used_.find(id) == used_.end())
                out.push_back(id);
        }
        return out;
    }

    bool IndicatorCoordinator::isIndicatorReserved(int id) const
    {
        return (used_.find(id) != used_.end());
    }

    bool IndicatorCoordinator::ensureIndicatorsInitialized(
        HWND hSciA,
        HWND hSciB,
        const std::vector<int>& preferredIds,
        const std::vector<int>& reservedIds)
    {
        if (!hSciA)
            return false;

        const bool changed = (hSciA != hA_) || (hSciB != hB_);
        if (!changed)
            return false;

        return init(hSciA, hSciB, preferredIds, reservedIds);
    }

} // namespace NppStyleKit