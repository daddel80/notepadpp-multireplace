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

#include <cstdint>
#include <functional>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

// Flow Tabs: CSV columns aligned on tab stops. Delimiters other than TAB, and numbers, need padding in the text: a TAB
// before each delimiter, spaces before numbers. The padding carries an indicator and stays out of the undo history;
// each switch leaves one marker there, and undo or redo of a marker puts the padding back exactly or takes it out.
namespace ColumnTabs {

    // A message to one Scintilla view
    using Send = std::function<std::intptr_t(unsigned int, std::uintptr_t, std::intptr_t)>;

    struct Settings {
        std::string delimiter;      // as it stands in the text
        char quote = '\0';          // '\0': fields are not quoted
        bool alignNumbers = true;
    };

    // [start, end) in the document
    struct Range {
        std::intptr_t start;
        std::intptr_t end;
    };

    enum class SwitchedOn { Padded, StopsOnly, NothingToAlign, ReadOnly };
    enum class Replayed { NotOurs, Padded, Unpadded, Reset };

    // How far a long run is: padding in (true) or out, percent done; 100 once it has finished. Short runs report nothing
    using Progress = std::function<void(bool padding, int percent)>;

    class Aligner {
    public:
        Aligner();
        ~Aligner();
        Aligner(const Aligner&) = delete;
        Aligner& operator=(const Aligner&) = delete;

        void setIndicator(int id) noexcept { indicator_ = id; }
        int indicator() const noexcept { return indicator_; }
        // Reported after delayMs at the earliest, then every intervalMs at most
        void setProgress(Progress progress, int delayMs = 300, int intervalMs = 150);

        // The document the view shows. inserted, removed: the padding in document order (removed: positions before)
        SwitchedOn switchOn(const Send& view, const Settings& settings, std::vector<Range>* inserted);
        bool switchOff(const Send& view, std::vector<Range>* removed);
        // After switchOff: the same padding again while the text is unchanged, else aligned anew
        SwitchedOn realign(const Send& view, const Settings& settings, std::vector<Range>* inserted);

        bool aligned(const Send& view) const;   // padding or tab stops
        bool padded(const Send& view) const;
        bool padded(std::intptr_t document) const;
        // The view switched to another document and dropped the tab stops: aligned by stops only means aligned no more
        void leave(std::intptr_t document) noexcept;

        // Undo or redo reached a marker: replay it once Scintilla has finished that step, a saved document stays saved.
        // Reset: the undo history no longer matched the text and was cleared, the document keeps its unsaved state
        static bool isMarker(std::intptr_t token) noexcept;
        Replayed replay(const Send& view, std::intptr_t token, bool undo, std::vector<Range>* changed);

        // Every SCN_MODIFIED of a view: inserted text never becomes padding, padding that comes back is marked again
        void onModified(const Send& view, int modificationType, std::intptr_t position, std::intptr_t length,
            const char* text, std::intptr_t linesAdded);

        void refreshStops(const Send& view);   // measured again, after zoom
        // The document's tab stops in this view: as aligned, none once no longer aligned
        void restoreStops(const Send& view);
        // A view that shows an aligned document without its stops (the other view aligned it) gets them
        void ensureStops(const Send& view);
        std::string withoutPadding(const Send& view, std::intptr_t start, std::intptr_t end) const;
        // Text the view just put on the clipboard, freed of its padding: the selection as Scintilla joins its ranges,
        // the line copied without one, or the text just cut. False if it is none of these or holds no padding
        bool copiedWithoutPadding(const Send& view, std::string& text) const;
        void forget(std::intptr_t document) noexcept;

        struct Doc;     // per document, in ColumnTabs.cpp

    private:
        Doc* find(const Send& view) const;
        Doc& at(const Send& view);

        int indicator_ = -1;
        bool busy_ = false;
        Progress progress_;
        int progressDelayMs_ = 300;
        int progressIntervalMs_ = 150;
        std::unordered_map<std::intptr_t, std::unique_ptr<Doc>> docs_;
    };
}
