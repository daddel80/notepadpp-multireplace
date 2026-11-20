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

// QA harness for EditorPalette, the colors MR fits to the editor's theme. On the themes Notepad++ ships,
// a grey ramp and random colors:
//   A. contrast, darkness and blending: the WCAG formula, a background is dark when white stands out more
//   B. visible: a mark keeps 3:1 to the background, and to the current line whenever a color on the way to
//      white or black allows both; a mark that stands out already keeps its color
//   C. readable: text reaches the asked contrast; text that reaches it already keeps its color
//   D. fillAlpha: the theme's text keeps 3:1 over the fill, or the theme's own contrast where that is lower;
//      the alpha is the highest that does so and never above the given one. fillColor: drawn at the given alpha
//      the text keeps the same contrast and a fill that keeps its alpha keeps its color; on the ramp and the shipped
//      themes it looks as the fill at fillAlpha (random themes hold colors of equal contrast that look different)
//   E. the default themes of Notepad++ keep their look: marks, frames, fills, column colors, Result Dock
//   F. every shipped theme with the colors MR uses: marks, frames, column text and labels keep their
//      contrast, and a fill whose alpha is lowered still stands out (CIELAB difference of at least 5)
// Build (from src/tests):
//   g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. -o editor_palette_qa editor_palette_qa.cpp ../EditorPalette.cpp
//   cl /std:c++17 /EHsc /I.. editor_palette_qa.cpp ..\EditorPalette.cpp /Fe:editor_palette_qa.exe
// Usage: editor_palette_qa [seed] [cases]
#include "../EditorPalette.h"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <random>
#include <string>
#include <vector>

using namespace EditorPalette;

static int fails = 0, checks = 0;
static void check(const std::string& name, bool ok)
{
    ++checks;
    if (!ok && ++fails <= 40) std::printf("FAIL %s\n", name.c_str());
}

static std::string hex(Color c)
{
    char text[16];
    std::snprintf(text, sizeof text, "%06X", static_cast<unsigned>(c));
    return text;
}

// Global styles of the themes Notepad++ ships (installer/themes, stylers.model.xml for the default), as
// COLORREF: default background and text, current line, line number margin
struct ShippedTheme { const char* name; Color background, text, caretLine, margin; };
static const ShippedTheme THEMES[] = {
    { "Bespin",            0x1C212A, 0x9DAEBD, 0x343C4B, 0x414A4C },
    { "Black board",       0x21100C, 0xF8F8F8, 0x301812, 0x36342E },
    { "Choco",             0x0B0F1A, 0x98BEC3, 0x111728, 0x36342E },
    { "DansLeRuSH-Dark",   0x2E2E2E, 0xC7C7C7, 0x363636, 0x363636 },
    { "DarkModeDefault",   0x3F3F3F, 0xCCDCDC, 0x101010, 0x0C0C0C },
    { "Deep Black",        0x000000, 0xFFFFFF, 0x333333, 0x333333 },
    { "Hello Kitty",       0xFFB0FF, 0x000000, 0xC080FF, 0xFF80FF },
    { "HotFudgeSundae",    0x010F2B, 0x5D97B7, 0x132C43, 0x0B2543 },
    { "Mono Industrial",   0x282C22, 0xFFFFFF, 0x33382C, 0x36342E },
    { "Monokai",           0x222827, 0xF2F8F8, 0x323D3E, 0x36342E },
    { "MossyLawn",         0x3D6958, 0x76C4F2, 0x397A71, 0x288A7E },
    { "Navajo",            0x809CBA, 0x000000, 0x7496B3, 0x808080 },
    { "Obsidian",          0x343129, 0xE4E2E0, 0x3C392F, 0x4E4B3F },
    { "Plastic Code Wrap", 0x1D160B, 0xF8F8F8, 0x2D2211, 0x36342E },
    { "Ruby Blue",         0x352411, 0xFFFFFF, 0x4B3A27, 0x61461F },
    { "Solarized-light",   0xE3F6FD, 0x837B65, 0xD5E8EE, 0xD5E8EE },
    { "Solarized",         0x362B00, 0x969483, 0x423607, 0x423607 },
    { "Twilight",          0x141414, 0xF8F8F8, 0x292929, 0x36342E },
    { "Vibrant Ink",       0x000000, 0xFFFFFF, 0x333333, 0x333333 },
    { "Zenburn",           0x3F3F3F, 0xCCDCDC, 0x101010, 0x0C0C0C },
    { "khaki",             0xAFD7D7, 0x005F5F, 0x87AFAF, 0x87AFAF },
    { "vim Dark Blue",     0x400000, 0xBFFFFF, 0x400000, 0x400000 },
    { "stylers.model",     0xFFFFFF, 0x000000, 0xFFE8E8, 0xE4E4E4 },
};
static const ShippedTheme& shipped(const std::string& name)
{
    for (const ShippedTheme& t : THEMES) if (name == t.name) return t;
    std::printf("no theme %s\n", name.c_str());
    std::exit(2);
}
static Theme themeOf(const ShippedTheme& t) { return makeTheme(t.background, t.text, t.caretLine); }

// The colors MR draws with (ReplaceMarks, SelectionFrame, MultiReplacePanel, ResultDock), light and dark
static constexpr Color MARKS[2] = { 0x505050, 0xB0B0B0 };
static constexpr Color FRAME[2] = { 0xBD814F, 0xD19B6A };
static constexpr int FRAME_ALPHA = 36;
static constexpr Color SINGLE_MARKER[2] = { 0x00A5FF, 0x0050D2 };
static constexpr Color DUPLICATE_MARKER[2] = { 0xE6E0B0, 0x909060 };
static constexpr int MARK_ALPHA[2] = { 150, 130 };
static constexpr Color COLUMNS[2][10] = {
    { 0xFFE0E0, 0xC0E0FF, 0x80FF80, 0xFFE0FF, 0xB0E0E0, 0xFFFF80, 0xE0C0C0, 0x80FFFF, 0xFFB0FF, 0xC0FFC0 },
    { 0x553333, 0x335577, 0x225522, 0x553355, 0x335555, 0x555522, 0x774444, 0x225555, 0x553366, 0x336633 } };
struct DockColors { Color lineNr, matchFg, filePathFg, headerBg, headerFg, critHdrBg, critHdrFg, foldGlyph, foldHighlight, caretLine; int caretLineAlpha; };
static constexpr DockColors DOCK[2] = {
    { 0xBF8040, 0x343FFA, 0x5080A0, 0xD5FFD5, 0x000000, 0xC4EBC4, 0x000000, 0x505050, 0x0000FF, 0xCD7887, 45 },
    { 0xFFC080, 0x2EE2A6, 0x8BCBEB, 0x9FAF8F, 0x000000, 0x849478, 0x000000, 0x808080, 0x869479, 0xAAAAAA, 64 } };
static constexpr Color DOCK_MARGIN_DARK = 0x000000;          // the margin's background in dark themes, else the editor's
static constexpr Color DOCK_SELECTION[2] = { 0xE0E0E0, 0x606060 };

// The column text of dark themes, tinted with the column color as MultiReplace::adjustForegroundForDarkMode does
static Color tintForDark(Color text, Color background)
{
    Color result = 0;
    for (const int shift : { 0, 8, 16 }) {
        const float blendFactor = 0.8f;
        int c = static_cast<int>(((text >> shift) & 0xFF) * (1.0f - blendFactor) + ((background >> shift) & 0xFF) * blendFactor);
        c = std::min(255, static_cast<int>(c * 1.9f));
        result |= static_cast<Color>(c) << shift;
    }
    return result;
}

// Colors over the range of ResultDock::generateColorFromText: every hue with its least, middle and most
// saturation and lightness, corrected per hue as it does
static Color hsl(double hue, double s, double l)
{
    s = std::clamp(s, 0.45, 0.85);
    l = std::clamp(l, 0.35, 0.72);
    const auto part = [](double p, double q, double t) {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 0.5) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    };
    const double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const double p = 2 * l - q;
    const auto byte = [](double v) { return static_cast<Color>(v * 255 + 0.5); };
    return byte(part(p, q, hue + 1.0 / 3)) | byte(part(p, q, hue)) << 8 | byte(part(p, q, hue - 1.0 / 3)) << 16;
}
static std::vector<Color> listColors(bool dark)
{
    std::vector<Color> colors;
    for (int step = 0; step < 144; ++step) {
        const double hue = step / 144.0, degrees = hue * 360;
        for (const double satVar : { -0.15, 0.0, 0.15 }) {
            for (const double litVar : { -0.10, -0.05, 0.0, 0.05, 0.10 }) {
                double s = (dark ? 0.60 : 0.65) + satVar, l = (dark ? 0.52 : 0.58) + litVar;
                if (degrees >= 40 && degrees < 110) {
                    const double t = (degrees - 40) / 70;
                    l -= 0.07 * (1 - (2 * t - 1) * (2 * t - 1));
                    if (degrees >= 50 && degrees < 70) s -= 0.08 * (1 - std::fabs(degrees - 60) / 10);
                }
                else if (degrees >= 210 && degrees < 280) {
                    const double t = (degrees - 210) / 70;
                    l += 0.06 * (1 - (2 * t - 1) * (2 * t - 1));
                }
                else if (degrees >= 330 || degrees < 15) {
                    const double t = degrees >= 330 ? (degrees - 330) / 30 : (degrees + 30) / 45;
                    s -= 0.06 * (1 - (2 * t - 1) * (2 * t - 1));
                }
                colors.push_back(hsl(hue, s, l));
            }
        }
    }
    return colors;
}

// CIELAB distance, how far apart two colors look
static double labDistance(Color a, Color b)
{
    const auto lab = [](Color c, double out[3]) {
        const auto lin = [](Color v) { const double x = v / 255.0; return x <= 0.04045 ? x / 12.92 : std::pow((x + 0.055) / 1.055, 2.4); };
        const double r = lin(c & 0xFF), g = lin((c >> 8) & 0xFF), bl = lin((c >> 16) & 0xFF);
        const double xyz[3] = { (0.4124 * r + 0.3576 * g + 0.1805 * bl) / 0.95047, 0.2126 * r + 0.7152 * g + 0.0722 * bl,
                                (0.0193 * r + 0.1192 * g + 0.9505 * bl) / 1.08883 };
        double f[3];
        for (int i = 0; i < 3; ++i) f[i] = xyz[i] > 0.008856 ? std::cbrt(xyz[i]) : 7.787 * xyz[i] + 16.0 / 116;
        out[0] = 116 * f[1] - 16;
        out[1] = 500 * (f[0] - f[1]);
        out[2] = 200 * (f[1] - f[2]);
    };
    double la[3], lb[3];
    lab(a, la);
    lab(b, lb);
    return std::sqrt((la[0] - lb[0]) * (la[0] - lb[0]) + (la[1] - lb[1]) * (la[1] - lb[1]) + (la[2] - lb[2]) * (la[2] - lb[2]));
}

// The guarantees of B, C and D for one case
static void checkVisible(const std::string& where, Color color, const Theme& theme)
{
    const Color mark = visible(color, theme);
    check(where + " visible " + hex(color) + ": 3:1 to the background", contrast(mark, theme.background) >= MARK_CONTRAST);
    if (contrast(color, theme.background) >= MARK_CONTRAST && (theme.caretLine == NO_COLOR || contrast(color, theme.caretLine) >= MARK_CONTRAST))
        check(where + " visible " + hex(color) + ": kept when it stands out", mark == color);
    if (theme.caretLine != NO_COLOR) {
        bool possible = false;   // any color on the way allows both
        const Color away = theme.dark ? 0xFFFFFF : 0x000000;
        for (int step = 0; step <= 64 && !possible; ++step) {
            const Color c = blend(away, color, step * 255 / 64);
            possible = contrast(c, theme.background) >= MARK_CONTRAST && contrast(c, theme.caretLine) >= MARK_CONTRAST;
        }
        if (possible) check(where + " visible " + hex(color) + ": 3:1 to the current line", contrast(mark, theme.caretLine) >= MARK_CONTRAST);
        else check(where + " visible " + hex(color) + ": as without the current line where both are out of reach",
            mark == visible(color, makeTheme(theme.background, theme.text)));
    }
}
static void checkReadable(const std::string& where, Color text, Color background, double minimum)
{
    const Color r = readable(text, background, minimum);
    check(where + " readable " + hex(text) + " on " + hex(background), contrast(r, background) >= minimum);
    if (contrast(text, background) >= minimum) check(where + " readable " + hex(text) + ": kept when readable", r == text);
}
static double g_mostLookDifference = 0;   // fillColor against fillAlpha, CIELAB, on the ramp and the shipped themes
static Color checkFillColor(const std::string& where, Color fill, int alpha, const Theme& theme, bool sameLook = true)
{
    const Color c = fillColor(fill, alpha, theme);
    const Color drawn = blend(c, theme.background, alpha);
    const int a = fillAlpha(fill, alpha, theme);
    const double wanted = std::min(FILL_CONTRAST, contrast(theme.text, theme.background));
    check(where + " fill color " + hex(fill) + ": the text keeps its contrast", contrast(theme.text, drawn) >= wanted);
    if (sameLook) {
        const double difference = labDistance(drawn, blend(fill, theme.background, a));
        g_mostLookDifference = std::max(g_mostLookDifference, difference);
        // At most a just noticeable difference: rounding twice costs a step where the text barely reaches 3:1
        check(where + " fill color " + hex(fill) + ": looks as at fillAlpha", difference <= 3.0);
    }
    if (a == std::clamp(alpha, 0, 255)) check(where + " fill color " + hex(fill) + ": kept with its alpha", c == (fill & 0xFFFFFF));
    return c;
}
static int checkFill(const std::string& where, Color fill, int alpha, const Theme& theme)
{
    const int a = fillAlpha(fill, alpha, theme);
    const double wanted = std::min(FILL_CONTRAST, contrast(theme.text, theme.background));
    check(where + " fill " + hex(fill) + ": within 0 and the given alpha", a >= 0 && a <= std::max(alpha, 0));
    check(where + " fill " + hex(fill) + ": the text keeps its contrast", contrast(theme.text, blend(fill, theme.background, a)) >= wanted);
    if (a < alpha) check(where + " fill " + hex(fill) + ": the highest alpha", contrast(theme.text, blend(fill, theme.background, a + 1)) < wanted);
    return a;
}

int main(int argc, char** argv)
{
    const unsigned seed = argc > 1 ? static_cast<unsigned>(std::strtoul(argv[1], nullptr, 10)) : 20261119u;
    const int cases = argc > 2 ? std::atoi(argv[2]) : 200000;
    std::mt19937 rng(seed);
    const auto randomColor = [&]() { return static_cast<Color>(rng() & 0xFFFFFF); };

    // A. contrast, darkness and blending
    check("A: white on black 21:1", std::fabs(contrast(0xFFFFFF, 0x000000) - 21.0) < 1e-9);
    check("A: a color on itself 1:1", std::fabs(contrast(0x3F7FBF, 0x3F7FBF) - 1.0) < 1e-12);
    check("A: grey 777777 on white 4.48:1", std::fabs(contrast(0x777777, 0xFFFFFF) - 4.48) < 0.01);
    check("A: near black 0A0A0A on black 1.061:1", std::fabs(contrast(0x0A0A0A, 0x000000) - 1.0607) < 0.0005);
    check("A: the order does not count", contrast(0x123456, 0xABCDEF) == contrast(0xABCDEF, 0x123456));
    check("A: green weighs more than blue", contrast(0x00FF00, 0x000000) > contrast(0xFF0000, 0x000000));
    check("A: black and 3F3F3F are dark", isDark(0x000000) && isDark(0x3F3F3F));
    check("A: white and 808080 are light", !isDark(0xFFFFFF) && !isDark(0x808080));
    check("A: blending at 0 and 255", blend(0x102030, 0xA0B0C0, 0) == 0xA0B0C0 && blend(0x102030, 0xA0B0C0, 255) == 0x102030);
    check("A: blending halfway", blend(0xFFFFFF, 0x000000, 128) == 0x808080);
    check("A: blending rounds", blend(0x000001, 0x000000, 128) == 0x000001 && blend(0x000001, 0x000000, 127) == 0x000000);
    check("A: blending clamps the alpha", blend(0xFFFFFF, 0x000000, 400) == 0xFFFFFF && blend(0xFFFFFF, 0x000000, -5) == 0x000000);
    {
        const Theme t = makeTheme(0xFF3F3F3F, 0xFFDCDCCC, 0x80101010);
        check("A: a theme drops the alpha bytes", t.background == 0x3F3F3F && t.text == 0xDCDCCC && t.caretLine == 0x101010 && t.dark);
        check("A: a theme without a current line", makeTheme(0xFFFFFF, 0).caretLine == NO_COLOR && !makeTheme(0xFFFFFF, 0).dark);
    }
    int darkSwitch = -1;   // the grey ramp turns light once
    for (int v = 0; v < 256; ++v) {
        const Color grey = static_cast<Color>(v * 0x010101);
        check("A: grey " + hex(grey) + " dark by the rule", isDark(grey) == (contrast(grey, 0xFFFFFF) > contrast(grey, 0x000000)));
        if (!isDark(grey) && darkSwitch < 0) darkSwitch = v;
        if (darkSwitch >= 0) check("A: lighter greys stay light " + hex(grey), !isDark(grey));
    }

    // B, C, D over the grey ramp and random colors
    for (int v = 0; v < 256; ++v) {
        const Color grey = static_cast<Color>(v * 0x010101);
        const Theme theme = makeTheme(grey, isDark(grey) ? 0xFFFFFF : 0x000000);
        for (const Color c : { MARKS[0], MARKS[1], FRAME[0], FRAME[1], grey, 0x000000u, 0xFFFFFFu }) checkVisible("B: ramp", c, theme);
        for (const Color c : { 0x000000u, 0xFFFFFFu, 0x808080u, grey }) checkReadable("C: ramp", c, grey, TEXT_CONTRAST);
        for (const Color c : { SINGLE_MARKER[0], SINGLE_MARKER[1], DUPLICATE_MARKER[0], DUPLICATE_MARKER[1], 0xFFFFFFu, 0x000000u }) {
            checkFill("D: ramp", c, 255, theme);
            checkFillColor("D: ramp", c, MARK_ALPHA[theme.dark ? 1 : 0], theme);
        }
    }
    for (int i = 0; i < cases; ++i) {
        const Theme theme = makeTheme(randomColor(), randomColor(), (i & 1) ? randomColor() : NO_COLOR);
        checkVisible("B: random", randomColor(), theme);
        checkReadable("C: random", randomColor(), randomColor(), (i & 2) ? TEXT_CONTRAST : MARK_CONTRAST);
        checkFill("D: random", randomColor(), static_cast<int>(rng() % 256), theme);
        checkFillColor("D: random", randomColor(), static_cast<int>(rng() % 256), theme, false);
        if (fails > 40) break;
    }

    // E. the default themes keep their look
    const Theme light = themeOf(shipped("stylers.model"));
    const Theme dark = themeOf(shipped("DarkModeDefault"));
    check("E: the default theme is light, the dark mode theme dark", !light.dark && dark.dark);
    for (int d = 0; d < 2; ++d) {
        const Theme& t = d ? dark : light;
        const std::string name = d ? "E: dark default" : "E: light default";
        check(name + ": replacement marks keep their color", visible(MARKS[d], t) == MARKS[d]);
        check(name + ": the scope frame keeps its color and fill", visible(FRAME[d], t) == FRAME[d] && fillAlpha(FRAME[d], FRAME_ALPHA, t) == FRAME_ALPHA);
        check(name + ": single and duplicate marks keep their fill", fillAlpha(SINGLE_MARKER[d], MARK_ALPHA[d], t) == MARK_ALPHA[d]
            && fillAlpha(DUPLICATE_MARKER[d], MARK_ALPHA[d], t) == MARK_ALPHA[d]);
        for (const Color c : COLUMNS[d]) {
            const Color text = d ? tintForDark(t.text, c) : t.text;
            check(name + ": column " + hex(c) + " keeps its text color", readable(text, c) == text);
        }
        const DockColors& dock = DOCK[d];
        const Color margin = d ? DOCK_MARGIN_DARK : t.background;
        check(name + ": dock labels keep their colors", readable(dock.lineNr, t.background, MARK_CONTRAST) == dock.lineNr
            && readable(dock.matchFg, t.background, MARK_CONTRAST) == dock.matchFg && readable(dock.filePathFg, t.background, MARK_CONTRAST) == dock.filePathFg);
        check(name + ": dock headers keep their colors", readable(dock.headerFg, dock.headerBg) == dock.headerFg && readable(dock.critHdrFg, dock.critHdrBg) == dock.critHdrFg);
        check(name + ": dock fold glyphs keep their colors", readable(dock.foldGlyph, margin, MARK_CONTRAST) == dock.foldGlyph
            && readable(dock.foldHighlight, margin, MARK_CONTRAST) == dock.foldHighlight);
        check(name + ": dock current line keeps its fill", fillAlpha(dock.caretLine, dock.caretLineAlpha, t) == dock.caretLineAlpha);
        const Color selText = d ? 0xFFFFFF : t.text;
        check(name + ": dock selection keeps its text color", readable(selText, DOCK_SELECTION[d]) == selText);
    }
    for (const Color c : listColors(false)) check("E: light default: list color " + hex(c) + " keeps its fill", fillAlpha(c, MARK_ALPHA[0], light) == MARK_ALPHA[0]);
    {
        int kept = 0, least = 255;
        const std::vector<Color> colors = listColors(true);
        for (const Color c : colors) {
            const int a = fillAlpha(c, MARK_ALPHA[1], dark);
            kept += a == MARK_ALPHA[1];
            least = std::min(least, a);
        }
        std::printf("E: dark default: %d of %zu list colors keep alpha %d, the least is %d\n", kept, colors.size(), MARK_ALPHA[1], least);
        check("E: dark default: list colors lose at most a fifth of their alpha", least >= MARK_ALPHA[1] * 4 / 5);
    }

    // F. every shipped theme with the colors MR uses
    for (const ShippedTheme& s : THEMES) {
        const Theme t = themeOf(s);
        const int d = t.dark ? 1 : 0;
        const std::string name = std::string("F: ") + s.name;
        checkVisible(name + " marks", MARKS[d], t);
        checkVisible(name + " frame", FRAME[d], t);
        checkFill(name + " frame", visible(FRAME[d], t), FRAME_ALPHA, t);
        for (const Color c : COLUMNS[d]) checkReadable(name + " column", d ? tintForDark(t.text, c) : t.text, c, TEXT_CONTRAST);
        const DockColors& dock = DOCK[d];
        const Color margin = d ? DOCK_MARGIN_DARK : t.background;
        for (const Color c : { dock.lineNr, dock.matchFg, dock.filePathFg }) checkReadable(name + " dock label", c, t.background, MARK_CONTRAST);
        checkReadable(name + " dock header", dock.headerFg, dock.headerBg, TEXT_CONTRAST);
        checkReadable(name + " dock fold glyph", dock.foldGlyph, margin, MARK_CONTRAST);
        checkReadable(name + " dock selection", d ? 0xFFFFFF : t.text, DOCK_SELECTION[d], TEXT_CONTRAST);
        checkFill(name + " dock current line", dock.caretLine, dock.caretLineAlpha, t);
        double leastSeen = 1e9;
        int lowered = 0;
        std::vector<Color> fills = listColors(t.dark);
        fills.push_back(SINGLE_MARKER[d]);
        fills.push_back(DUPLICATE_MARKER[d]);
        for (const Color c : fills) {
            const int a = checkFill(name, c, MARK_ALPHA[d], t);
            const Color drawn = blend(checkFillColor(name, c, MARK_ALPHA[d], t), t.background, MARK_ALPHA[d]);
            if (a == MARK_ALPHA[d]) continue;
            ++lowered;
            const double seen = labDistance(blend(c, t.background, a), t.background);
            leastSeen = std::min(leastSeen, seen);
            check(name + " fill " + hex(c) + ": stands out at alpha " + std::to_string(a), seen >= 5.0);
            check(name + " fill color " + hex(c) + ": stands out", labDistance(drawn, t.background) >= 5.0);
        }
        std::printf("%-20s %s  text %5.2f:1  marks %s  frame %s  fills lowered %3d of %zu%s\n", s.name, t.dark ? "dark " : "light",
            contrast(t.text, t.background), hex(visible(MARKS[d], t)).c_str(), hex(visible(FRAME[d], t)).c_str(), lowered, fills.size(),
            lowered ? (", least difference " + std::to_string(leastSeen).substr(0, 4)).c_str() : "");
    }

    std::printf("fillColor against fillAlpha: at most %.2f apart (CIELAB)\n", g_mostLookDifference);
    std::printf("%d/%d checks passed\n", checks - fails, checks);
    return fails ? 1 : 0;
}
