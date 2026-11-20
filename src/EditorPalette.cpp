// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
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

#include "EditorPalette.h"

#include <algorithm>
#include <cmath>

namespace EditorPalette {

    namespace {
        constexpr Color WHITE = 0xFFFFFF;
        constexpr Color BLACK = 0x000000;
        constexpr int STEPS = 64;                // fineness of the way towards white or black

        int channel(Color color, int shift) { return static_cast<int>((color >> shift) & 0xFF); }

        double linear(int value)
        {
            const double c = value / 255.0;
            return c <= 0.03928 ? c / 12.92 : std::pow((c + 0.055) / 1.055, 2.4);
        }

        double luminance(Color color)
        {
            return 0.2126 * linear(channel(color, 0)) + 0.7152 * linear(channel(color, 8)) + 0.0722 * linear(channel(color, 16));
        }

        // The first color on the way from color to end that passes the test; end itself when none does
        template <typename Test>
        Color towards(Color color, Color end, Test test)
        {
            for (int step = 0; step <= STEPS; ++step) {
                const Color candidate = blend(end, color, step * 255 / STEPS);
                if (test(candidate)) return candidate;
            }
            return end;
        }
    }

    Theme makeTheme(Color background, Color text, Color caretLine)
    {
        Theme theme;
        theme.background = background & 0xFFFFFF;
        theme.text = text & 0xFFFFFF;
        theme.caretLine = caretLine == NO_COLOR ? NO_COLOR : caretLine & 0xFFFFFF;
        theme.dark = isDark(theme.background);
        return theme;
    }

    double contrast(Color a, Color b)
    {
        const double la = luminance(a);
        const double lb = luminance(b);
        return (std::max(la, lb) + 0.05) / (std::min(la, lb) + 0.05);
    }

    bool isDark(Color background)
    {
        return contrast(background, WHITE) > contrast(background, BLACK);
    }

    Color blend(Color over, Color under, int alpha)
    {
        alpha = std::clamp(alpha, 0, 255);
        Color result = 0;
        for (const int shift : { 0, 8, 16 }) {
            const int mixed = (channel(over, shift) * alpha + channel(under, shift) * (255 - alpha) + 127) / 255;
            result |= static_cast<Color>(mixed) << shift;
        }
        return result;
    }

    Color visible(Color color, const Theme& theme)
    {
        const Color away = isDark(theme.background) ? WHITE : BLACK;   // white or black always reaches 4.58:1
        const auto offBackground = [&](Color c) { return contrast(c, theme.background) >= MARK_CONTRAST; };
        const Color onBackground = towards(color & 0xFFFFFF, away, offBackground);
        if (theme.caretLine == NO_COLOR) return onBackground;
        const auto offBoth = [&](Color c) { return offBackground(c) && contrast(c, theme.caretLine) >= MARK_CONTRAST; };
        const Color onBoth = towards(color & 0xFFFFFF, away, offBoth);
        return offBoth(onBoth) ? onBoth : onBackground;
    }

    Color readable(Color text, Color background, double minimum)
    {
        const Color away = isDark(background) ? WHITE : BLACK;
        return towards(text & 0xFFFFFF, away, [&](Color c) { return contrast(c, background) >= minimum; });
    }

    int fillAlpha(Color fill, int alpha, const Theme& theme)
    {
        const double wanted = std::min(FILL_CONTRAST, contrast(theme.text, theme.background));
        for (int a = std::clamp(alpha, 0, 255); a > 0; --a) {
            if (contrast(theme.text, blend(fill, theme.background, a)) >= wanted) return a;
        }
        return 0;
    }

    Color fillColor(Color fill, int alpha, const Theme& theme)
    {
        alpha = std::clamp(alpha, 0, 255);
        if (alpha == 0) return fill & 0xFFFFFF;
        const double wanted = std::min(FILL_CONTRAST, contrast(theme.text, theme.background));
        // As far towards the background as fillAlpha lowers the alpha, a step further where rounding costs contrast
        for (int share = fillAlpha(fill, alpha, theme) * 255 / alpha; share > 0; --share) {
            const Color candidate = blend(fill, theme.background, share);
            if (contrast(theme.text, blend(candidate, theme.background, alpha)) >= wanted) return candidate;
        }
        return theme.background;
    }
}
