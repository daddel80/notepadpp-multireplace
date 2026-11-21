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

// QA for ListHash. legacy() is the hash of the versions before 6.3, value for value: the 64-bit values here were
// written by such a build into its INI. of() tells every change of a list that counts as content. A baseline finds
// its way out of an INI of those versions and back into it, also when they have run in between. That holds for
// untitled tabs as well, whose baseline those versions set to the list at exit.
// Build: g++ -std=c++17 -Wall -Wextra -fsanitize=address,undefined -I.. list_hash_qa.cpp -o list_hash_qa
// With MSVC it also compares legacy() with std::hash as those versions called it (x64 and x86).
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <functional>
#include <random>
#include <set>
#include <string>
#include <vector>

#include "../ListHash.h"

namespace {

    using List = std::vector<ReplaceItemData>;
    using ListHash::Baseline;

    int g_checks = 0;
    int g_fail = 0;

    void check(bool ok, const std::string& what)
    {
        ++g_checks;
        if (!ok) {
            ++g_fail;
            if (g_fail <= 30) std::printf("FAIL  %s\n", what.c_str());
        }
    }

    const unsigned En = 1, WW = 2, MC = 4, FS = 8, EXT = 16, RX = 32;

    ReplaceItemData entry(const std::wstring& find, const std::wstring& replace, unsigned options = En, const std::wstring& comments = L"")
    {
        ReplaceItemData item;
        item.isEnabled = (options & En) != 0;
        item.findText = find;
        item.replaceText = replace;
        item.wholeWord = (options & WW) != 0;
        item.matchCase = (options & MC) != 0;
        item.formulaSupport = (options & FS) != 0;
        item.extended = (options & EXT) != 0;
        item.regex = (options & RX) != 0;
        item.comments = comments;
        return item;
    }

    List many()
    {
        List list;
        for (int i = 0; i < 300; ++i) {
            const unsigned options = (i % 3 != 0 ? En : 0) | (i % 2 ? WW : 0) | (i % 5 == 0 ? MC : 0) | (i % 7 == 0 ? RX : 0);
            list.push_back(entry(L"find " + std::to_wstring(i), L"repl " + std::to_wstring(i * 7 % 13), options,
                i % 4 ? std::wstring() : L"c" + std::to_wstring(i)));
        }
        return list;
    }

    // legacy64: as a 64-bit build of 6.x wrote it into TabN_BaselineHash; legacy32: the same for a 32-bit build
    struct Golden {
        const char* name;
        List list;
        std::uint64_t legacy64;
        std::uint32_t legacy32;
        std::uint64_t hash;
    };

    std::vector<Golden> goldens()
    {
        return {
            { "single", { entry(L"a", L"b") }, 0x663A131D208BC467ULL, 0xFE737E07U, 0xA9BDC372CD08914CULL },
            { "disabled", { entry(L"alpha", L"beta", 0) }, 0x5B9003994F8A7C90ULL, 0x7AF02590U, 0x707CCBD26C872D5AULL },
            { "flags", { entry(L"f1", L"r", En | WW), entry(L"f2", L"r", En | MC), entry(L"f3", L"r", En | FS),
                entry(L"f4", L"r", En | EXT), entry(L"f5", L"r", En | RX) }, 0xE8739C8B4AEA7E4AULL, 0x2D55A576U, 0xC7CACC7D031EFE26ULL },
            { "unicode", { entry(L"Gr\u00F6\u00DFe \u00DF \u20AC", L"\u65E5\u672C\u8A9E", En | MC, L"emoji \U0001F600 \U0001D11E") },
                0xDAB096262827F8D1ULL, 0x2C379231U, 0xFFE572044001A275ULL },
            { "specials", { entry(L"He said \"hi\"", L"C:\\temp\\x", En | EXT, L"a, b; c"), entry(L"line one\r\nline two", L"joined", En | EXT),
                entry(L"tab\there", L""), entry(L"", L"only replace", En, L"empty find") }, 0xF5BC4A78238526A3ULL, 0xFF3811FEU, 0xE47AF222729F8FA8ULL },
            { "orderAB", { entry(L"one", L"1"), entry(L"two", L"2") }, 0xE175FED9AF28AC20ULL, 0x95F58428U, 0x381B50CD6222EE72ULL },
            { "orderBA", { entry(L"two", L"2"), entry(L"one", L"1") }, 0xE96905F3388D7FABULL, 0xFFE43DC3U, 0xDB5A2E68A5284466ULL },
            { "long", { entry(std::wstring(3000, L'x') + L"end", std::wstring(2000, L'y')) }, 0x3F9CE482DC5D1D3FULL, 0xC1ADC51FU, 0xD949F87B87BDAE1CULL },
            { "extended", { entry(L"\\t", L"TAB", En | EXT) }, 0x9AB00327E40E9809ULL, 0xC3F78FA9U, 0x396B94DDE0F3C2B4ULL },
            { "regex", { entry(L"\\t", L"TAB", En | RX) }, 0x9AB00327E40E9809ULL, 0xC3F78FA9U, 0x9F1EE93C7F754BC4ULL },
            { "sameSame", { entry(L"same", L"same", En | MC) }, 0x6641FD1D207E15F7ULL, 0xAC6550D7U, 0x2C0839E723CE3961ULL },
            { "otherOther", { entry(L"other", L"other", En | MC) }, 0x6641FD1D207E15F7ULL, 0xAC6550D7U, 0x68F9378CB893B50DULL },
            { "leftRight", { entry(L"left", L"right") }, 0x4AE0258AA780DCB7ULL, 0x4E2AF3D7U, 0x2EB283C227963228ULL },
            { "rightLeft", { entry(L"right", L"left") }, 0x4AE0258AA780DCB7ULL, 0x4E2AF3D7U, 0xB9A871AFADC96884ULL },
            { "twoFlags", { entry(L"word", L"w", En | WW | MC) }, 0xA53F1CBD927A9E73ULL, 0x2FC5B393U, 0x4912EAD350A84BBCULL },
            { "noFlags", { entry(L"word", L"w") }, 0xA53F1CBD927A9E73ULL, 0x2FC5B393U, 0xA71FA446C675D8B2ULL },
            { "noteText", { entry(L"note", L"r", En, L"text") }, 0x2D7EABFA98F3E95FULL, 0x318EF0FFU, 0xEDB022D75493EB48ULL },
            { "textNote", { entry(L"text", L"r", En, L"note") }, 0x2D7EABFA98F3E95FULL, 0x318EF0FFU, 0x6BBAFE6E938CB2BAULL },
            { "allFlags", { entry(L"all", L"flags", En | WW | MC | FS | EXT | RX, L"c") }, 0x0A43CD1EF4533C4DULL, 0x7872F76DU, 0x84BC1007262AFBAAULL },
            { "emptyTexts", { entry(L"", L"") }, 0x6641FF1D207E1991ULL, 0xAA6553F1U, 0x8E081D9C1B473B7DULL },
            { "many", many(), 0x1A875AE9D48FDB90ULL, 0xA6CCE4AEU, 0x9FC2600774273DD2ULL },
            { "commentOnly", { entry(L"", L"", En, L"only a comment") }, 0x3F3B19F9B8ABEEBDULL, 0xD98E929DU, 0x427401F9F831553DULL },
            { "empty", {}, 0, 0, 0 },
        };
    }

    const List& golden(const std::vector<Golden>& all, const std::string& name)
    {
        for (const Golden& g : all)
            if (name == g.name) return g.list;
        static const List none;
        return none;
    }

    void appendCodePoint(std::wstring& text, std::uint32_t codePoint)
    {
        if (sizeof(wchar_t) == 2 && codePoint > 0xFFFF) {
            text.push_back(static_cast<wchar_t>(0xD800 + ((codePoint - 0x10000) >> 10)));
            text.push_back(static_cast<wchar_t>(0xDC00 + ((codePoint - 0x10000) & 0x3FF)));
        }
        else {
            text.push_back(static_cast<wchar_t>(codePoint));
        }
    }

    std::wstring randomText(std::mt19937& rng)
    {
        static const std::wstring common = L"ab \\\"\t\r\n,;:0";
        std::wstring text;
        const int length = std::uniform_int_distribution<int>(0, 9)(rng);
        for (int i = 0; i < length; ++i) {
            const int kind = std::uniform_int_distribution<int>(0, 9)(rng);
            if (kind < 7) text.push_back(common[std::uniform_int_distribution<std::size_t>(0, common.size() - 1)(rng)]);
            else if (kind < 9) appendCodePoint(text, std::uniform_int_distribution<std::uint32_t>(0, 0xD7FF)(rng));
            else appendCodePoint(text, std::uniform_int_distribution<std::uint32_t>(0x10000, 0x10FFFF)(rng));
        }
        return text;
    }

    List randomList(std::mt19937& rng)
    {
        List list;
        const int count = std::uniform_int_distribution<int>(0, 5)(rng);
        for (int i = 0; i < count; ++i)
            list.push_back(entry(randomText(rng), randomText(rng), std::uniform_int_distribution<unsigned>(0, 63)(rng), randomText(rng)));
        return list;
    }

#ifdef _MSC_VER
    // The list hash as the versions before 6.3 have it in ReplaceItemData.h and computeListHash
    struct ReplaceItemDataHasher {
        std::size_t operator()(const ReplaceItemData& item) const {
            std::size_t hash = std::hash<bool>{}(item.isEnabled);
            hash ^= std::hash<std::wstring>{}(item.findText) << 1;
            hash ^= std::hash<std::wstring>{}(item.replaceText) << 1;
            hash ^= std::hash<bool>{}(item.wholeWord) << 1;
            hash ^= std::hash<bool>{}(item.matchCase) << 1;
            hash ^= std::hash<bool>{}(item.formulaSupport) << 1;
            hash ^= std::hash<bool>{}(item.extended) << 1;
            hash ^= std::hash<bool>{}(item.regex) << 1;
            hash ^= std::hash<std::wstring>{}(item.comments) << 1;
            return hash;
        }
    };

    std::size_t computeListHash(const List& list)
    {
        const std::size_t golden_ratio_constant = 0x9e3779b9;
        std::size_t combinedHash = 0;
        ReplaceItemDataHasher hasher;

        for (const auto& item : list) {
            combinedHash ^= hasher(item) + golden_ratio_constant + (combinedHash << 6) + (combinedHash >> 2);
        }

        return combinedHash;
    }
#endif

    // ---- A tab across versions: what its INI keys and files hold, and what the user gets to see at the next start

    struct Ini {
        bool hasLegacy = false;
        std::uint64_t legacy = 0;      // TabN_BaselineHash
        std::wstring text;             // TabN_ListHash; empty: no such key
        bool dirty = false;            // TabN_IsDirty
    };

    struct Disk {
        List snapshot;                 // tabN.mrtab
        bool hasFile = false;          // the tab has a list file and it can be read
        List file;
    };

    struct Seen {
        bool dirty = false;            // dirty marker on the tab
        bool close = false;            // closing the tab asks to save it
        bool prompt = false;           // "file was modified" at startup
    };

    bool same(const Seen& a, const Seen& b) { return a.dirty == b.dirty && a.close == b.close && a.prompt == b.prompt; }

    // 6.0 to 6.2: loadTabsFromConfig, checkForUnsavedChanges and checkSingleTabForFileChange
    Seen oldStart(const Ini& ini, const Disk& disk, std::uint64_t& baseline)
    {
        baseline = ini.hasLegacy ? ini.legacy : disk.hasFile ? ListHash::legacy(disk.file) : ListHash::legacy(disk.snapshot);
        const bool edited = ListHash::legacy(disk.snapshot) != baseline;
        return { ini.dirty || edited, edited, disk.hasFile && ListHash::legacy(disk.file) != baseline };
    }

    // 6.0 to 6.2: writeTabsToConfig, which sets the baseline of an untitled tab to its list. A key they do not know stays as it is
    Ini oldExit(const Ini& before, const Disk& disk, std::uint64_t baseline, bool dirty, bool untitled)
    {
        Ini ini = before;
        ini.hasLegacy = true;
        ini.legacy = untitled ? ListHash::legacy(disk.snapshot) : baseline;
        ini.dirty = dirty;
        return ini;
    }

    Seen newStart(const Ini& ini, const Disk& disk, bool untitled, Baseline& baseline)
    {
        if (ini.hasLegacy) {
            baseline = ListHash::fromIni(ini.text, ini.legacy, disk.snapshot, [&disk](List& file) {
                if (!disk.hasFile) return false;
                file = disk.file;
                return true;
            });
        }
        else {
            baseline = ListHash::baselineOf(disk.hasFile ? disk.file : disk.snapshot);
        }
        if (untitled) baseline = ListHash::untitledBaseline(baseline, disk.snapshot, ini.dirty);
        const bool edited = !baseline.matches(disk.snapshot);
        return { edited, edited, disk.hasFile && !baseline.matches(disk.file) };
    }

    Ini newExit(const Baseline& baseline, bool dirty)
    {
        return { true, baseline.legacy, ListHash::toText(baseline), dirty };
    }

    // The tab as a version before 6.3 leaves it: saved with the content base, then closed with snapshot in the list
    Ini writtenByOld(const List& base, const Disk& disk, bool untitled)
    {
        const std::uint64_t baseline = ListHash::legacy(base);
        return oldExit({}, disk, baseline, ListHash::legacy(disk.snapshot) != baseline, untitled);
    }
}

int main()
{
    const std::vector<Golden> all = goldens();

    // ---- legacy(): the values of the versions before 6.3
    for (const Golden& g : all) {
        check(ListHash::detail::legacyAs<std::uint64_t>(g.list) == g.legacy64, std::string("legacy, 64-bit: ") + g.name);
        check(ListHash::detail::legacyAs<std::uint32_t>(g.list) == g.legacy32, std::string("legacy, 32-bit: ") + g.name);
        const std::uint64_t native = sizeof(std::size_t) == 8 ? g.legacy64 : g.legacy32;
        check(ListHash::legacy(g.list) == native, std::string("legacy, this platform: ") + g.name);
    }

    // What it misses, and of() does not
    const char* missed[][2] = { { "extended", "regex" }, { "twoFlags", "noFlags" }, { "leftRight", "rightLeft" },
        { "sameSame", "otherOther" }, { "noteText", "textNote" } };
    for (const auto& pair : missed) {
        const List& a = golden(all, pair[0]);
        const List& b = golden(all, pair[1]);
        check(!a.empty() && !b.empty() && ListHash::legacy(a) == ListHash::legacy(b), std::string("legacy misses: ") + pair[0] + " / " + pair[1]);
        check(ListHash::of(a) != ListHash::of(b), std::string("of tells apart: ") + pair[0] + " / " + pair[1]);
    }

    // ---- of(): its values stay as they are, they are in the INI
    for (const Golden& g : all)
        check(ListHash::of(g.list) == g.hash, std::string("of: ") + g.name);

    check(ListHash::of({}) == 0 && ListHash::legacy({}) == 0, "the empty list is 0 in both");
    check(Baseline{}.matches({}), "a baseline that was never set is the empty list");
    check(!Baseline{}.matches(golden(all, "single")), "and no other");

    // Every combination of the options is its own value
    {
        std::set<std::uint64_t> values;
        for (unsigned options = 0; options < 64; ++options)
            values.insert(ListHash::of({ entry(L"find", L"replace", options, L"c") }));
        check(values.size() == 64, "64 option combinations, 64 values");
    }

    // What counts and what does not
    {
        const List base = { entry(L"ab", L"c", En, L"d"), entry(L"e", L"f") };
        const std::uint64_t hash = ListHash::of(base);

        List list = base;
        list[0].id = 7; list[0].findCount = 3; list[0].replaceCount = 2; list[0].isDirty = true; list[0].lastModified = L"2026-11-21 12:00:00";
        check(ListHash::of(list) == hash, "ids, counts, the row marker and the timestamp do not count");

        check(ListHash::of({ entry(L"a", L"bc", En, L"d"), entry(L"e", L"f") }) != hash, "Find and Replace do not run together");
        check(ListHash::of({ entry(L"ab", L"cd", En, L""), entry(L"e", L"f") }) != hash, "Replace and Comments do not run together");
        check(ListHash::of({ entry(L"ab", L"c", En, L"de"), entry(L"", L"f") }) != hash, "entries do not run together");
        check(ListHash::of({ entry(L"e", L"f"), entry(L"ab", L"c", En, L"d") }) != hash, "the order counts");
        check(ListHash::of({ entry(L"ab", L"c", En, L"d") }) != hash, "an entry less");
        check(ListHash::of({ entry(L"ab", L"c", En, L"d"), entry(L"e", L"f"), entry(L"e", L"f") }) != hash, "an entry twice");
        check(ListHash::of({ entry(L"ab", L"c", En, L"d"), entry(L"E", L"f") }) != hash, "upper and lower case");
        check(ListHash::of({ entry(std::wstring(1, L'\0'), L"") }) != ListHash::of({ entry(L"", std::wstring(1, L'\0')) }), "a NUL character counts where it is");
    }

    // Random lists: equal content, equal value; one change, another value
    {
        std::mt19937 rng(20261121u);
        int misses = 0;
        for (int round = 0; round < 200000; ++round) {
            const List list = randomList(rng);
            List copy = list;
            if (ListHash::of(copy) != ListHash::of(list) || ListHash::legacy(copy) != ListHash::legacy(list)) ++misses;

            if (list.empty()) {
                copy.push_back(entry(L"", L""));
            }
            else {
                ReplaceItemData& item = copy[std::uniform_int_distribution<std::size_t>(0, copy.size() - 1)(rng)];
                switch (std::uniform_int_distribution<int>(0, 9)(rng)) {
                case 0: item.isEnabled = !item.isEnabled; break;
                case 1: item.wholeWord = !item.wholeWord; break;
                case 2: item.matchCase = !item.matchCase; break;
                case 3: item.formulaSupport = !item.formulaSupport; break;
                case 4: { const bool extended = item.extended; item.extended = item.regex; item.regex = !extended; break; }   // two at once
                case 5: item.findText += L'x'; break;
                case 6: item.replaceText += L'x'; break;
                case 7: item.comments += L'x'; break;
                case 8: { if (item.findText == item.replaceText) item.findText += L'y'; std::swap(item.findText, item.replaceText); break; }
                default: copy.push_back(item); break;
                }
            }
            if (ListHash::of(copy) == ListHash::of(list)) ++misses;
        }
        g_checks += 200000;
        g_fail += misses;
        if (misses) std::printf("FAIL  random lists: %d misses\n", misses);
    }

    // ---- The INI value and the way from the legacy hash
    {
        const Baseline baseline{ 0x0123456789ABCDEFULL, 0xFEDCBA9876543210ULL };
        const std::wstring text = ListHash::toText(baseline);
        check(text == L"0123456789ABCDEF:FEDCBA9876543210", "toText");

        Baseline read;
        check(ListHash::fromText(text, baseline.legacy, read) && read.hash == baseline.hash && read.legacy == baseline.legacy, "fromText reads toText");
        check(ListHash::fromText(L"0123456789abcdef:fedcba9876543210", baseline.legacy, read) && read.hash == baseline.hash, "lower case digits");
        check(ListHash::fromText(ListHash::toText({}), 0, read) && read.hash == 0 && read.legacy == 0, "the empty list");

        read = { 1, 2 };
        check(!ListHash::fromText(text, baseline.legacy + 1, read) && read.hash == 1 && read.legacy == 2, "another legacy hash beside it: out of date, nothing read");
        const wchar_t* broken[] = { L"", L"0123456789ABCDEF", L"0123456789ABCDEF:FEDCBA987654321", L"0123456789ABCDEF:FEDCBA98765432100",
            L"0123456789ABCDEF;FEDCBA9876543210", L"0123456789ABCDEG:FEDCBA9876543210", L" 123456789ABCDEF:FEDCBA9876543210",
            L"0x23456789ABCDEF:FEDCBA9876543210", L"0123456789ABCDEF:FEDCBA987654321 " };
        for (const wchar_t* value : broken)
            check(!ListHash::fromText(value, baseline.legacy, read), "no baseline text");
    }
    {
        const List& saved = golden(all, "flags");
        const List& edited = golden(all, "specials");
        const List& other = golden(all, "many");
        const std::uint64_t stored = ListHash::legacy(saved);

        Baseline b = ListHash::rebase(stored, saved, &saved);
        check(b.hash == ListHash::of(saved) && b.legacy == stored, "rebase: list and file are the baseline");
        b = ListHash::rebase(stored, edited, &saved);
        check(b.hash == ListHash::of(saved) && b.legacy == stored && !b.matches(edited), "rebase: the file is the baseline, the list is edited");
        b = ListHash::rebase(stored, saved, &other);
        check(b.hash == ListHash::of(saved) && b.legacy == stored && !b.matches(other), "rebase: the list is the baseline, the file has changed");
        b = ListHash::rebase(stored, saved, nullptr);
        check(b.hash == ListHash::of(saved) && b.legacy == stored, "rebase: no file, the list is the baseline");
        b = ListHash::rebase(stored, edited, &other);
        check(b.legacy == stored && !b.matches(edited) && !b.matches(other) && !b.matches(saved) && !b.matches({}), "rebase: unknown, no list matches");
        b = ListHash::rebase(0, edited, nullptr);
        check(b.hash == 0 && b.legacy == 0 && b.matches({}) && !b.matches(edited), "rebase: 0 is the empty list");

        // A change the legacy hash missed: the file is the baseline, so the list shows as changed
        const List& hidden = golden(all, "regex");
        const List& onDisk = golden(all, "extended");
        b = ListHash::rebase(ListHash::legacy(onDisk), hidden, &onDisk);
        check(b.matches(onDisk) && !b.matches(hidden), "rebase: a missed change comes to light");
    }

    // ---- The baseline of an untitled tab
    {
        const List& s = golden(all, "specials");
        const Baseline reset = ListHash::baselineOf(s);            // as 6.2 sets it at exit
        check(ListHash::untitledBaseline(reset, s, false).matches(s), "untitled, not flagged: its list is the baseline");
        check(ListHash::untitledBaseline(reset, s, true).matches({}), "untitled, flagged dirty: the baseline of a new tab");
        const Baseline own = ListHash::baselineOf(golden(all, "single"));
        const Baseline kept = ListHash::untitledBaseline(own, s, true);
        check(kept.hash == own.hash && kept.legacy == own.legacy, "untitled, flagged and edited: its baseline stays");
        check(ListHash::untitledBaseline(Baseline{}, {}, true).matches({}), "untitled, flagged and emptied: saved");
    }

    // ---- A tab across versions
    {
        const List& f0 = golden(all, "flags");       // the list file as last saved
        const List& f1 = golden(all, "many");        // the file changed by someone else
        const List& s1 = golden(all, "specials");    // the list edited in the tab
        const List& s2 = golden(all, "unicode");     // another list edited in the tab
        const List& ext = golden(all, "extended");
        const List& rx = golden(all, "regex");       // ext with a change the legacy hash misses

        // old: 6.2 with its own INI; now: 6.3 after the update and at every start after; down: 6.2 with the INI of 6.3
        struct Case { const char* name; List base; Disk disk; bool untitled; Seen old, now, down; };
        const Seen none{}, unsaved{ true, true, false }, changed{ false, false, true }, both{ true, true, true };
        const Seen dotOnly{ true, false, false };    // 6.2 with an untitled tab whose baseline it set to the list
        const Case cases[] = {
            { "file tab, clean",                     f0, { f0, true, f0 },  false, none,    none,    none },
            { "file tab, edited",                    f0, { s1, true, f0 },  false, unsaved, unsaved, unsaved },
            { "file tab, file changed outside",      f0, { f0, true, f1 },  false, changed, changed, changed },
            { "file tab, edited and file changed",   f0, { s1, true, f1 },  false, both,    both,    both },
            { "file tab, file gone",                 f0, { f0, false, {} }, false, none,    none,    none },
            { "untitled tab, edited",                {}, { s1, false, {} }, true,  dotOnly, unsaved, unsaved },
            { "untitled tab, a duplicate",           s1, { s1, false, {} }, true,  none,    none,    none },
            { "untitled tab, a duplicate edited",    s1, { s2, false, {} }, true,  dotOnly, unsaved, unsaved },
            { "untitled tab, emptied",               s1, { {}, false, {} }, true,  dotOnly, none,    none },
            { "untitled tab, empty",                 {}, { {}, false, {} }, true,  none,    none,    none },
        };
        for (const Case& c : cases) {
            const Ini fromOld = writtenByOld(c.base, c.disk, c.untitled);
            std::uint64_t oldBaseline = 0;
            check(same(oldStart(fromOld, c.disk, oldBaseline), c.old), std::string("6.2 itself: ") + c.name);

            // Update
            Baseline baseline;
            const Seen updated = newStart(fromOld, c.disk, c.untitled, baseline);
            check(same(updated, c.now), std::string("update: ") + c.name);

            // And back: 6.2 reads the legacy hash of the baseline; for a file tab it is the one 6.2 writes itself
            const Ini fromNew = newExit(baseline, updated.dirty);
            if (!c.untitled)
                check(fromNew.legacy == oldExit(fromOld, c.disk, oldBaseline, false, false).legacy, std::string("downgrade, the legacy key: ") + c.name);
            std::uint64_t downBaseline = 0;
            const Seen down = oldStart(fromNew, c.disk, downBaseline);
            check(same(down, c.down), std::string("downgrade: ") + c.name);

            // Next start of 6.3: its own key is up to date, the file is not read for it
            bool fileRead = false;
            Baseline again = ListHash::fromIni(fromNew.text, fromNew.legacy, c.disk.snapshot, [&fileRead](List&) { fileRead = true; return false; });
            check(!fileRead && again.hash == baseline.hash, std::string("next start: ") + c.name);
            check(same(newStart(fromNew, c.disk, c.untitled, again), c.now), std::string("next start, the picture: ") + c.name);

            // 6.2 ran in between and closed with the tab unchanged: 6.3 shows the same as before
            const Ini backAgain = oldExit(fromNew, c.disk, downBaseline, down.dirty, c.untitled);
            check(same(newStart(backAgain, c.disk, c.untitled, again), c.now), std::string("after a downgrade: ") + c.name);
        }

        // The INI has no baseline at all (key removed by hand): both fall back the same way
        {
            const Disk disk{ s1, true, f0 };
            Ini ini; ini.dirty = true;
            std::uint64_t oldBaseline = 0;
            Baseline baseline;
            check(same(oldStart(ini, disk, oldBaseline), newStart(ini, disk, false, baseline)) && baseline.matches(f0), "no baseline key: the file is the baseline");
            const Disk untitled{ s1, false, {} };
            check(same(oldStart({}, untitled, oldBaseline), newStart({}, untitled, true, baseline)) && baseline.matches(s1), "no baseline key, no file: the list is");
            check(same(newStart(ini, untitled, true, baseline), unsaved) && baseline.matches({}), "no baseline key, no file, flagged dirty: the baseline of a new tab");
        }

        // A change that 6.2 missed and closed with: after the update the tab shows it and the file is not blamed
        {
            const Disk disk{ rx, true, ext };
            const Ini fromOld = writtenByOld(ext, disk, false);
            std::uint64_t oldBaseline = 0;
            check(!fromOld.dirty && same(oldStart(fromOld, disk, oldBaseline), none), "6.2: a missed change, the tab looks saved");
            Baseline baseline;
            const Seen updated = newStart(fromOld, disk, false, baseline);
            check(same(updated, unsaved), "update: the missed change shows as unsaved");
            const Ini fromNew = newExit(baseline, updated.dirty);
            check(same(oldStart(fromNew, disk, oldBaseline), dotOnly), "downgrade: 6.2 marks it by the dirty flag, but misses it on closing");
        }

        // 6.3, then 6.2 saves other content into the tab and leaves the key of 6.3 behind, then 6.3 again
        {
            Disk disk{ f0, true, f0 };
            Baseline baseline = ListHash::baselineOf(f0);
            const Ini fromNew = newExit(baseline, false);

            disk = { s1, true, s1 };                                       // edited and saved with 6.2
            const Ini fromOld = oldExit(fromNew, disk, ListHash::legacy(s1), false, false);
            check(fromOld.text == fromNew.text && fromOld.text == ListHash::toText(ListHash::baselineOf(f0)), "6.2 leaves the key of 6.3 as it was");

            const Seen seen = newStart(fromOld, disk, false, baseline);
            check(same(seen, none) && baseline.matches(s1), "6.3 again: the key left behind is not taken");

            // The same key on a tab that moved to this index in 6.2 (another tab was closed)
            const Disk moved{ f1, true, f1 };
            const Ini shifted = oldExit(fromNew, moved, ListHash::legacy(f1), false, false);
            check(same(newStart(shifted, moved, false, baseline), none) && baseline.matches(f1), "6.3 again: the key of another tab is not taken");

            // 6.2 ran, but left this tab alone: the key still holds
            const Disk kept{ f0, true, f0 };
            const Ini untouched = oldExit(fromNew, kept, fromNew.legacy, false, false);
            bool fileRead = false;
            baseline = ListHash::fromIni(untouched.text, untouched.legacy, kept.snapshot, [&fileRead](List&) { fileRead = true; return false; });
            check(!fileRead && baseline.matches(f0), "6.3 again: an untouched tab keeps its key");
        }

        // An untitled tab of 6.3 stays unsaved across restarts until it is saved
        {
            const Disk filled{ s1, false, {} };
            Baseline baseline = ListHash::baselineOf({});              // a new tab, then filled
            Ini ini = newExit(baseline, true);
            for (int start = 1; start <= 3; ++start) {
                check(same(newStart(ini, filled, true, baseline), unsaved) && baseline.matches({}), "a new tab, filled: unsaved at start " + std::to_string(start));
                ini = newExit(baseline, true);
            }

            // Saved in 6.2, then whole word on and match case off: a change legacy() misses
            const List& same0 = golden(all, "sameSame");
            List flipped = same0;
            flipped[0].wholeWord = true;
            flipped[0].matchCase = false;
            check(ListHash::legacy(flipped) == ListHash::legacy(same0) && ListHash::of(flipped) != ListHash::of(same0), "the flip is one legacy() misses");

            const Disk before{ same0, false, {} };
            check(same(newStart(writtenByOld(same0, before, true), before, true, baseline), none), "untitled from 6.2, unchanged: saved");
            const Disk after{ flipped, false, {} };
            const Ini fromNew = newExit(baseline, !baseline.matches(flipped));
            check(fromNew.dirty && fromNew.text == ListHash::toText(ListHash::baselineOf(same0)), "the INI keeps the baseline from before the flip");
            check(same(newStart(fromNew, after, true, baseline), unsaved), "next start: still unsaved, closing asks");

            // 6.2 in between: the dot by the flag only; at exit it sets the same legacy hash, so the key of 6.3 holds
            std::uint64_t oldBaseline = 0;
            const Seen old = oldStart(fromNew, after, oldBaseline);
            check(same(old, dotOnly), "6.2: the dot by the flag, no question on closing");
            const Ini fromOld = oldExit(fromNew, after, oldBaseline, old.dirty, true);
            check(fromOld.legacy == fromNew.legacy && fromOld.text == fromNew.text, "6.2 writes the same legacy hash");
            check(same(newStart(fromOld, after, true, baseline), unsaved) && baseline.matches(same0), "6.3 again: unsaved, against the baseline from before the flip");
        }

        // First start after 5.x: [File] OriginalListHash and the working list, no list file
        {
            Baseline b = ListHash::rebase(ListHash::legacy(f0), f0, nullptr);
            check(b.matches(f0), "from 5.x: the working list as it was saved");
            b = ListHash::rebase(ListHash::legacy(f0), s1, nullptr);
            check(!b.matches(s1), "from 5.x: the working list edited since");
            b = ListHash::rebase(0, {}, nullptr);
            check(b.matches({}) && b.hash == 0 && b.legacy == 0, "first start ever: an empty list, saved");
        }
    }

#ifdef _MSC_VER
    // ---- legacy() against std::hash, the way the versions before 6.3 call it
    {
        int differences = 0;
        for (const Golden& g : all)
            if (ListHash::legacy(g.list) != computeListHash(g.list)) ++differences;
        std::mt19937 rng(20261121u);
        for (int round = 0; round < 200000; ++round) {
            const List list = randomList(rng);
            if (ListHash::legacy(list) != computeListHash(list)) ++differences;
        }
        g_checks += 200000 + static_cast<int>(all.size());
        g_fail += differences;
        std::printf("std::hash, size_t of %d bytes: %d differences\n", static_cast<int>(sizeof(std::size_t)), differences);
    }
#endif

    std::printf("%d/%d checks passed\n", g_checks - g_fail, g_checks);
    return g_fail ? 1 : 0;
}
