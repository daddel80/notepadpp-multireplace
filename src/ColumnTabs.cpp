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

#include "ColumnTabs.h"

#include "CsvLine.h"
#include "NumericToken.h"
#include "Scintilla.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <deque>
#include <map>
#include <string_view>
#include <unordered_map>

namespace ColumnTabs {

    namespace {

        using Pos = std::intptr_t;

        // Marker tokens: 'MR' in the upper half, a number per document in the lower
        constexpr Pos MARKER_TAG = 0x4D52;
        constexpr Pos MARKER_IDS = 0xFFFF;

        constexpr int GAP_SPACES = 2;                       // between a column and the next delimiter
        constexpr std::size_t MAX_RECORDS = 1024;           // markers per document that can be replayed
        constexpr std::size_t MAX_JOURNAL_BYTES = 32u << 20;

        // Padding of a line without padding: count times ch at offset
        struct Run {
            std::int32_t offset;
            std::int32_t count;
            char ch;
            bool operator==(const Run& o) const noexcept { return offset == o.offset && count == o.count && ch == o.ch; }
        };
        using Runs = std::vector<Run>;
        using Exceptions = std::map<Pos, Runs>;     // line -> padding that differs from the recipe

        // Marked bytes, relative to the start of a piece of text
        struct Mark {
            std::int32_t offset;
            std::int32_t length;
        };

        struct Fingerprint {
            Pos length = -1;
            std::uint64_t hash = 0;
            bool operator==(const Fingerprint& o) const noexcept { return length == o.length && hash == o.hash; }
            bool operator!=(const Fingerprint& o) const noexcept { return !(*this == o); }
        };

        std::uint64_t fnv1a(const char* p, std::size_t n) noexcept
        {
            std::uint64_t h = 1469598103934665603ull;
            for (std::size_t i = 0; i < n; ++i) {
                h ^= static_cast<unsigned char>(p[i]);
                h *= 1099511628211ull;
            }
            return h;
        }

        Fingerprint fingerprint(const Send& v)
        {
            const Pos length = v(SCI_GETLENGTH, 0, 0);
            const char* text = reinterpret_cast<const char*>(v(SCI_GETCHARACTERPOINTER, 0, 0));
            return { length, text ? fnv1a(text, static_cast<std::size_t>(length)) : 0 };
        }

        // One line without its line end; valid until the next change
        std::string_view lineText(const Send& v, Pos line, Pos* start = nullptr)
        {
            const Pos from = v(SCI_POSITIONFROMLINE, static_cast<std::uintptr_t>(line), 0);
            const Pos to = v(SCI_GETLINEENDPOSITION, static_cast<std::uintptr_t>(line), 0);
            if (start) *start = from;
            if (to <= from) return {};
            const char* p = reinterpret_cast<const char*>(v(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(from), to - from));
            return p ? std::string_view(p, static_cast<std::size_t>(to - from)) : std::string_view();
        }

        bool isBlank(char c) noexcept { return c == ' ' || c == '\t'; }

        // Widths at one zoom: ASCII measured once, other text measured as it comes
        struct Metrics {
            std::array<double, 128> ascii{};
            double gap = 0;                 // between a column and the next delimiter
            double delimiter = 0;           // width of the delimiter, 0 for TAB
        };

        // How a document is aligned: settings and what was measured when switching on
        struct Recipe {
            Settings settings;
            bool tabDelimited = false;      // the delimiters are TABs already
            bool numbers = false;           // spaces before numbers
            std::vector<int> intWidth;      // per column: digits of the widest integer part, 0 = no numbers
            Metrics metrics;
            std::vector<int> stops;         // pixels, shared by all lines
            Pos zoom = 0;                   // zoom the stops were measured at
        };

        // A field's number as numeric alignment sees it: prefix and suffix allowed, as in CSV sorting
        struct Number {
            std::size_t start = 0;          // sign or first digit
            int intDigits = 0;
            bool sign = false;
            bool aligned = false;           // has digits or a decimal separator
        };

        bool numberIn(std::string_view line, std::size_t s, std::size_t e, Number& out)
        {
            while (s < e && isBlank(line[s])) ++s;
            while (e > s && isBlank(line[e - 1])) --e;
            if (e <= s) return false;
            mr::num::NumericToken tok;
            if (!mr::num::classify_numeric_field(line.substr(s, e - s), tok)) return false;

            std::size_t i = s + tok.start;
            out.start = i;
            out.sign = i < e && (line[i] == '+' || line[i] == '-');
            if (out.sign) ++i;
            int digits = 0;
            while (i < e && line[i] >= '0' && line[i] <= '9') { ++digits; ++i; }
            const bool separator = i < e && (line[i] == '.' || line[i] == ',');
            out.intDigits = digits;
            out.aligned = digits > 0 || separator || tok.hasDecimal;
            return true;
        }

        void delimitersOf(std::string_view line, const Settings& s, std::vector<std::size_t>& out)
        {
            out.clear();
            CsvLine::forEachDelimiter(line, s.delimiter, s.quote, [&](std::size_t d) { out.push_back(d); });
        }

        // Field c of a line with these delimiters: [first, second)
        std::pair<std::size_t, std::size_t> field(std::string_view line, const std::vector<std::size_t>& delims,
            std::size_t delimiterLength, std::size_t c)
        {
            const std::size_t s = c == 0 ? 0 : delims[c - 1] + delimiterLength;
            const std::size_t e = c < delims.size() ? delims[c] : line.size();
            return { (std::min)(s, e), e };
        }

        // Padding of one line without padding, left to right
        void planLine(std::string_view line, const Recipe& r, std::vector<std::size_t>& delims, Runs& runs)
        {
            runs.clear();
            delimitersOf(line, r.settings, delims);
            const std::size_t dl = r.settings.delimiter.size();
            for (std::size_t c = 0; c <= delims.size(); ++c) {
                const auto [s, e] = field(line, delims, dl, c);
                if (r.numbers && c < r.intWidth.size() && r.intWidth[c] > 0) {
                    Number n;
                    if (numberIn(line, s, e, n) && n.aligned) {
                        int touching = 0;
                        for (std::size_t p = n.start; p > s && line[p - 1] == ' '; --p) ++touching;
                        int need = 1 + r.intWidth[c] - n.intDigits;
                        if (need < 0) need = 0;
                        if (n.sign && need > 0) --need;
                        if (need > touching)
                            runs.push_back({ static_cast<std::int32_t>(n.start), need - touching, ' ' });
                    }
                }
                // A TAB already in front of the delimiter does the job
                if (c < delims.size() && !r.tabDelimited && !(e > 0 && line[e - 1] == '\t'))
                    runs.push_back({ static_cast<std::int32_t>(e), 1, '\t' });
            }
        }

        // Pixel widths of text in STYLE_DEFAULT, TABs not counted: ASCII from the table, the rest measured
        class Widths {
        public:
            Widths(const Send& v, const std::array<double, 128>& ascii) : v_(v), ascii_(ascii) {}

            static std::array<double, 128> measureAscii(const Send& v)
            {
                std::array<double, 128> ascii{};
                char sample[101] = {};
                for (int c = 32; c < 127; ++c) {
                    std::fill(sample, sample + 100, static_cast<char>(c));
                    ascii[c] = static_cast<double>(v(SCI_TEXTWIDTH, STYLE_DEFAULT, reinterpret_cast<std::intptr_t>(sample))) / 100.0;
                }
                return ascii;
            }

            double of(std::string_view text)
            {
                double w = 0;
                for (const char ch : text) {
                    const unsigned char c = static_cast<unsigned char>(ch);
                    if (c == '\t') continue;
                    if (c < 32 || c > 126) return measured(text);
                    w += ascii_[c];
                }
                return w;
            }

            double space() const noexcept { return ascii_[' ']; }

        private:
            double measured(std::string_view text)
            {
                buffer_.clear();
                for (const char ch : text)
                    if (ch != '\t') buffer_.push_back(ch);
                return static_cast<double>(v_(SCI_TEXTWIDTH, STYLE_DEFAULT, reinterpret_cast<std::intptr_t>(buffer_.c_str())));
            }

            const Send& v_;
            const std::array<double, 128>& ascii_;
            std::string buffer_;
        };

        // Progress over the passes of one run, checked every 1024 lines; reported after a delay, then at intervals
        class Meter {
        public:
            Meter(const Progress& report, bool padding, int delayMs, int intervalMs)
                : report_(report), padding_(padding), interval_(intervalMs),
                next_(Clock::now() + std::chrono::milliseconds(delayMs)) {}
            ~Meter()
            {
                if (!shown_ || !report_) return;
                try { report_(padding_, 100); } catch (...) {}
            }
            Meter(const Meter&) = delete;
            Meter& operator=(const Meter&) = delete;

            // The next pass is this share of the run, in percent
            void pass(int share) noexcept { base_ += share_; share_ = share; }

            void at(Pos done, Pos total)
            {
                if (!report_ || (done & 1023) != 0 || total <= 0) return;
                const Clock::time_point now = Clock::now();
                if (now < next_) return;
                next_ = now + std::chrono::milliseconds(interval_);
                const int percent = base_ + static_cast<int>(static_cast<long long>(share_) * done / total);
                shown_ = (std::min)(99, (std::max)(shown_, (std::max)(1, percent)));
                report_(padding_, shown_);
            }

        private:
            using Clock = std::chrono::steady_clock;
            const Progress& report_;
            bool padding_;
            int interval_;
            Clock::time_point next_;
            int base_ = 0;
            int share_ = 0;
            int shown_ = 0;
        };

        // Widths at the current zoom, and stops from the widest cell per column: the gap, half a space to spare, then
        // the delimiter
        void measure(const Send& v, Recipe& r, Meter* meter = nullptr)
        {
            r.metrics.ascii = Widths::measureAscii(v);
            Widths w(v, r.metrics.ascii);
            const double space = (std::max)(1.0, std::round(w.space()));
            r.metrics.gap = GAP_SPACES * space + std::ceil(space / 2);
            r.metrics.delimiter = r.tabDelimited ? 0.0 : std::ceil(w.of(r.settings.delimiter));
            r.zoom = v(SCI_GETZOOM, 0, 0);

            std::vector<double> widest;
            std::vector<std::size_t> delims;
            const std::size_t dl = r.settings.delimiter.size();
            const Pos lines = v(SCI_GETLINECOUNT, 0, 0);
            for (Pos line = 0; line < lines; ++line) {
                if (meter) meter->at(line, lines);
                const std::string_view text = lineText(v, line);
                delimitersOf(text, r.settings, delims);
                if (delims.empty()) continue;
                if (widest.size() < delims.size() + 1) widest.resize(delims.size() + 1, 0.0);
                for (std::size_t c = 0; c <= delims.size(); ++c) {
                    const auto [s, e] = field(text, delims, dl, c);
                    widest[c] = (std::max)(widest[c], w.of(text.substr(s, e - s)));
                }
            }

            r.stops.clear();
            double x = 0;
            for (std::size_t c = 0; c + 1 < widest.size(); ++c) {
                x += std::ceil(widest[c]) + r.metrics.gap;
                r.stops.push_back(static_cast<int>(x));
                x += r.metrics.delimiter;
            }
        }

        // Stops of one line: the shared ones, pushed right only where a cell of this line needs more room. The TAB
        // in front of a delimiter jumps to the stop, text typed behind it starts there; a cell without one uses none
        void lineStops(std::string_view text, const Recipe& r, Widths& w, std::vector<std::size_t>& delims,
            std::vector<int>& out)
        {
            out.clear();
            delimitersOf(text, r.settings, delims);
            const std::size_t dl = r.settings.delimiter.size();
            double x = 0;
            for (std::size_t c = 0; c < delims.size(); ++c) {
                const auto [s, e] = field(text, delims, dl, c);
                const std::string_view cell = text.substr(s, e - s);
                const std::size_t tab = r.tabDelimited ? cell.size() : cell.rfind('\t');
                if (tab == std::string_view::npos) {
                    x += w.of(cell) + r.metrics.delimiter;
                    continue;
                }
                const int fit = static_cast<int>(std::ceil(x + w.of(cell.substr(0, tab))) + r.metrics.gap);
                const int stop = (std::max)(c < r.stops.size() ? r.stops[c] : 0, fit);
                out.push_back(stop);
                x = stop + w.of(cell.substr((std::min)(tab + 1, cell.size()))) + r.metrics.delimiter;
            }
            // Stops of columns this line lacks, as far as they lie behind its own
            for (std::size_t c = delims.size(); c < r.stops.size(); ++c)
                if (out.empty() || r.stops[c] > out.back()) out.push_back(r.stops[c]);
        }

        void setStops(const Send& v, Pos first, Pos last, const std::vector<int>& stops, Meter* meter = nullptr)
        {
            for (Pos line = first; line <= last; ++line) {
                if (meter) meter->at(line - first, last - first + 1);
                v(SCI_CLEARTABSTOPS, static_cast<std::uintptr_t>(line), 0);
                for (const int x : stops) v(SCI_ADDTABSTOP, static_cast<std::uintptr_t>(line), x);
            }
        }

        // Each line its own stops, measured from its cells; cells that fit keep the shared ones
        void fitStops(const Send& v, Pos first, Pos last, const Recipe& r, Meter* meter = nullptr)
        {
            Widths w(v, r.metrics.ascii);
            std::vector<std::size_t> delims;
            std::vector<int> stops;
            for (Pos line = first; line <= last; ++line) {
                if (meter) meter->at(line - first, last - first + 1);
                lineStops(lineText(v, line), r, w, delims, stops);
                v(SCI_CLEARTABSTOPS, static_cast<std::uintptr_t>(line), 0);
                for (const int x : stops) v(SCI_ADDTABSTOP, static_cast<std::uintptr_t>(line), x);
            }
        }

        void clearStops(const Send& v, Meter* meter = nullptr)
        {
            const Pos lines = v(SCI_GETLINECOUNT, 0, 0);
            for (Pos line = 0; line < lines; ++line) {
                if (meter) meter->at(line, lines);
                v(SCI_CLEARTABSTOPS, static_cast<std::uintptr_t>(line), 0);
            }
        }

        // Notepad++ and other plugins fill with the document's current indicator and value: put both back
        class KeepCurrent {
        public:
            explicit KeepCurrent(const Send& v)
                : v_(v), indicator_(v(SCI_GETINDICATORCURRENT, 0, 0)), value_(v(SCI_GETINDICATORVALUE, 0, 0)) {}
            ~KeepCurrent()
            {
                v_(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(indicator_), 0);
                v_(SCI_SETINDICATORVALUE, static_cast<std::uintptr_t>(value_), 0);
            }
            KeepCurrent(const KeepCurrent&) = delete;
            KeepCurrent& operator=(const KeepCurrent&) = delete;

        private:
            const Send& v_;
            Pos indicator_;
            Pos value_;
        };

        // While Flow Tabs change the text: outside the undo history, unannounced, past a read-only flag. Scintilla's
        // change history replays the undo history and breaks on changes outside it (a crash on a later undo): the
        // document's change history goes first. The view keeps its setting, it starts a new history only without undo
        class Quiet {
        public:
            Quiet(const Send& v, int indicator)
                : v_(v), keep_(v), undo_(v(SCI_GETUNDOCOLLECTION, 0, 0) != 0), mask_(v(SCI_GETMODEVENTMASK, 0, 0)),
                readOnly_(v(SCI_GETREADONLY, 0, 0) != 0), changeHistory_(v(SCI_GETCHANGEHISTORY, 0, 0))
            {
                v(SCI_SETCHANGEHISTORY, SC_CHANGE_HISTORY_DISABLED, 0);
                if (undo_) v(SCI_SETUNDOCOLLECTION, 0, 0);
                v(SCI_SETMODEVENTMASK, 0, 0);
                if (readOnly_) v(SCI_SETREADONLY, 0, 0);
                v(SCI_SETINDICATORCURRENT, static_cast<std::uintptr_t>(indicator), 0);
                v(SCI_SETINDICATORVALUE, 1, 0);
            }
            ~Quiet()
            {
                if (readOnly_) v_(SCI_SETREADONLY, 1, 0);
                v_(SCI_SETMODEVENTMASK, static_cast<std::uintptr_t>(mask_), 0);
                if (undo_) v_(SCI_SETUNDOCOLLECTION, 1, 0);
                if (changeHistory_) v_(SCI_SETCHANGEHISTORY, static_cast<std::uintptr_t>(changeHistory_), 0);
            }
            Quiet(const Quiet&) = delete;
            Quiet& operator=(const Quiet&) = delete;

            // Notepad++ rewrites its session backup only after a change it was told about
            void announce(bool on) { v_(SCI_SETMODEVENTMASK, on ? static_cast<std::uintptr_t>(mask_) : 0, 0); }

        private:
            const Send& v_;
            KeepCurrent keep_;
            bool undo_;
            Pos mask_;
            bool readOnly_;
            Pos changeHistory_;
        };

        void insertRuns(const Send& v, Pos lineStart, const Runs& runs, std::vector<Range>* inserted, std::string& buffer)
        {
            Pos delta = 0;
            for (const Run& r : runs) {
                const Pos at = lineStart + r.offset + delta;
                buffer.assign(static_cast<std::size_t>(r.count), r.ch);
                v(SCI_SETTARGETRANGE, static_cast<std::uintptr_t>(at), at);
                v(SCI_REPLACETARGET, buffer.size(), reinterpret_cast<std::intptr_t>(buffer.data()));
                v(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(at), r.count);
                if (inserted) inserted->push_back({ at, at + r.count });
                delta += r.count;
            }
        }

        // Pads every line: the exception where there is one, else the recipe. Returns the number of runs
        std::size_t padLines(const Send& v, const Recipe* recipe, const Exceptions* exceptions, std::vector<Range>* inserted,
            Meter* meter)
        {
            std::size_t count = 0;
            std::string buffer;
            std::vector<std::size_t> delims;
            Runs runs;
            const Pos lines = v(SCI_GETLINECOUNT, 0, 0);
            for (Pos ln = 0; ln < lines; ++ln) {
                if (meter) meter->at(ln, lines);
                Pos start = 0;
                const Runs* use = nullptr;
                if (exceptions) {
                    const auto it = exceptions->find(ln);
                    if (it != exceptions->end()) use = &it->second;
                }
                if (!use) {
                    if (!recipe) continue;
                    planLine(lineText(v, ln, &start), *recipe, delims, runs);
                    use = &runs;
                }
                else {
                    start = v(SCI_POSITIONFROMLINE, static_cast<std::uintptr_t>(ln), 0);
                }
                insertRuns(v, start, *use, inserted, buffer);
                count += use->size();
            }
            return count;
        }

        // Takes the marked padding out, line by line; the first deletion is announced. Lines whose padding differs from
        // the recipe go into exceptions exactly, so padLines(recipe, exceptions) puts back what was taken out
        void stripLines(const Send& v, int indicator, const Recipe* recipe, Quiet& quiet, Exceptions& exceptions,
            std::vector<Range>* removed, Meter* meter)
        {
            std::vector<std::size_t> delims;
            std::vector<Range> marked;
            Runs actual, planned;
            Pos removedBefore = 0;
            bool announced = false;
            const Pos lines = v(SCI_GETLINECOUNT, 0, 0);
            for (Pos ln = 0; ln < lines; ++ln) {
                if (meter) meter->at(ln, lines);
                Pos start = 0;
                const std::string_view text = lineText(v, ln, &start);
                const Pos end = start + static_cast<Pos>(text.size());

                marked.clear();
                for (Pos p = start; p < end;) {
                    Pos e = v(SCI_INDICATOREND, static_cast<std::uintptr_t>(indicator), p);
                    if (e <= p) break;
                    if (e > end) e = end;
                    if (v(SCI_INDICATORVALUEAT, static_cast<std::uintptr_t>(indicator), p)) marked.push_back({ p, e });
                    p = e;
                }

                actual.clear();
                Pos inLine = 0;
                for (const Range& m : marked) {
                    for (Pos p = m.start; p < m.end;) {
                        const char ch = text[static_cast<std::size_t>(p - start)];
                        Pos q = p + 1;
                        while (q < m.end && text[static_cast<std::size_t>(q - start)] == ch) ++q;
                        actual.push_back({ static_cast<std::int32_t>(p - start - inLine), static_cast<std::int32_t>(q - p), ch });
                        inLine += q - p;
                        p = q;
                    }
                    if (removed) removed->push_back({ m.start + removedBefore, m.end + removedBefore });
                }
                for (std::size_t k = marked.size(); k-- > 0;) {
                    if (!announced) quiet.announce(true);
                    v(SCI_DELETERANGE, static_cast<std::uintptr_t>(marked[k].start), marked[k].end - marked[k].start);
                    if (!announced) { quiet.announce(false); announced = true; }
                }
                removedBefore += inLine;

                if (recipe) {
                    planLine(lineText(v, ln), *recipe, delims, planned);
                    if (planned != actual) exceptions.emplace(ln, actual);
                }
                else if (!actual.empty()) {
                    exceptions.emplace(ln, actual);
                }
            }
        }

        // Every line of a piece of text without its line end: visit(offset, content)
        template <class Visit>
        void forEachPiece(std::string_view text, Visit visit)
        {
            std::size_t from = 0;
            for (;;) {
                const std::size_t nl = text.find('\n', from);
                std::size_t to = nl == std::string_view::npos ? text.size() : nl;
                if (to > from && text[to - 1] == '\r') --to;
                visit(from, text.substr(from, to - from));
                if (nl == std::string_view::npos) break;
                from = nl + 1;
            }
        }

        // Deleted text that held padding: enough to mark the padding again when the text comes back
        struct Deleted {
            struct Piece {
                std::uint64_t hash = 0;      // of the line content
                std::int32_t offset = 0;     // in the deleted text
                std::int32_t length = 0;
                std::vector<Mark> marks;     // relative to the piece
            };
            Pos position = 0;
            Pos length = 0;
            Pos lengthAfter = 0;     // document length after the deletion
            std::uint64_t hash = 0;
            std::vector<Piece> pieces;
            std::size_t bytes = 0;
        };

        void fillMarks(const Send& v, Pos at, const std::vector<Mark>& marks)
        {
            for (const Mark& m : marks) v(SCI_INDICATORFILLRANGE, static_cast<std::uintptr_t>(at + m.offset), m.length);
        }

        // A change as Scintilla reports it, once for each view that shows the document
        struct Change {
            int type = 0;
            Pos position = -1;
            Pos length = 0;
            Pos documentLength = -1;
            bool operator==(const Change& o) const noexcept
            {
                return type == o.type && position == o.position && length == o.length && documentLength == o.documentLength;
            }
        };

        // Undo history entry of one switch: which padding, and the text with and without it
        struct Record {
            std::shared_ptr<const Recipe> recipe;
            std::shared_ptr<const Exceptions> exceptions;
            Fingerprint unpadded;
            Fingerprint padded;
            bool padsOnRedo = true;      // switched on; a switch off pads on undo
        };
    }

    struct Aligner::Doc {
        std::shared_ptr<const Recipe> recipe;    // while aligned
        bool padded = false;                     // padding in the text, else tab stops only
        std::unordered_map<Pos, Record> records;
        std::deque<Pos> order;                   // tokens, oldest first
        Pos nextId = 0;
        Pos lastOff = 0;                         // marker of the last switch off
        std::deque<Deleted> journal;
        std::size_t journalBytes = 0;
        const Deleted* pending = nullptr;        // deleted just now: the next insertion may replace it
        Change lastChange;                       // the other view reports it again
    };

    namespace {

        Pos addMarker(const Send& v, Aligner::Doc& d, Record record)
        {
            d.nextId = d.nextId % MARKER_IDS + 1;
            const Pos token = (MARKER_TAG << 16) | d.nextId;
            if (d.records.erase(token)) d.order.erase(std::find(d.order.begin(), d.order.end(), token));
            d.records.emplace(token, std::move(record));
            d.order.push_back(token);
            while (d.order.size() > MAX_RECORDS) {
                d.records.erase(d.order.front());
                d.order.pop_front();
            }
            // The marker leaves the save point unannounced: a saved document stays saved, Notepad++ keeps track
            const bool saved = v(SCI_GETMODIFY, 0, 0) == 0;
            v(SCI_ADDUNDOACTION, static_cast<std::uintptr_t>(token), 0);
            if (saved) v(SCI_SETSAVEPOINT, 0, 0);
            return token;
        }

        // The undo history no longer matches the text: clear it, an unsaved document stays unsaved
        void resetHistory(const Send& v, Aligner::Doc& d)
        {
            const bool modified = v(SCI_GETMODIFY, 0, 0) != 0;
            v(SCI_EMPTYUNDOBUFFER, 0, 0);
            if (modified) v(SCI_SETUNDOSAVEPOINT, static_cast<std::uintptr_t>(-1), 0);
            d.records.clear();
            d.order.clear();
            d.journal.clear();
            d.journalBytes = 0;
            d.pending = nullptr;
            d.lastOff = 0;
        }

        // Was the document at its save point before the undo or redo step that reached a marker? A marker never changes
        // that: the text without padding is the same on both sides of it
        bool savedBeforeStep(const Send& v, bool undo)
        {
            const Pos savePoint = v(SCI_GETUNDOSAVEPOINT, 0, 0);
            return savePoint >= 0 && savePoint == v(SCI_GETUNDOCURRENT, 0, 0) + (undo ? 1 : -1);
        }

        void stripAll(const Send& v, int indicator, const Recipe* recipe, Exceptions& exceptions, std::vector<Range>* removed,
            Meter* meter = nullptr)
        {
            Quiet quiet(v, indicator);
            stripLines(v, indicator, recipe, quiet, exceptions, removed, meter);
        }

        std::size_t padAll(const Send& v, int indicator, const Recipe* recipe, const Exceptions* exceptions,
            std::vector<Range>* inserted, Meter* meter = nullptr)
        {
            Quiet quiet(v, indicator);
            return padLines(v, recipe, exceptions, inserted, meter);
        }

        // Stops of a recipe for the whole text: measured anew when the zoom changed since, else line by line, as lines
        // edited meanwhile may need more room
        std::shared_ptr<const Recipe> applyStops(const Send& v, std::shared_ptr<const Recipe> recipe, Meter* meter = nullptr)
        {
            const Pos last = v(SCI_GETLINECOUNT, 0, 0) - 1;
            if (v(SCI_GETZOOM, 0, 0) != recipe->zoom) {
                auto fresh = std::make_shared<Recipe>(*recipe);
                measure(v, *fresh, meter);
                setStops(v, 0, last, fresh->stops, meter);
                return fresh;
            }
            fitStops(v, 0, last, *recipe, meter);
            return recipe;
        }

        struct Busy {
            explicit Busy(bool& flag) : flag_(flag) { flag_ = true; }
            ~Busy() { flag_ = false; }
            Busy(const Busy&) = delete;
            Busy& operator=(const Busy&) = delete;
            bool& flag_;
        };
    }

    Aligner::Aligner() = default;
    Aligner::~Aligner() = default;

    Aligner::Doc* Aligner::find(const Send& view) const
    {
        const auto it = docs_.find(view(SCI_GETDOCPOINTER, 0, 0));
        return it == docs_.end() ? nullptr : it->second.get();
    }

    Aligner::Doc& Aligner::at(const Send& view)
    {
        auto& slot = docs_[view(SCI_GETDOCPOINTER, 0, 0)];
        if (!slot) slot = std::make_unique<Doc>();
        return *slot;
    }

    bool Aligner::aligned(const Send& view) const
    {
        const Doc* d = find(view);
        return d && d->recipe;
    }

    bool Aligner::padded(const Send& view) const
    {
        const Doc* d = find(view);
        return d && d->recipe && d->padded;
    }

    bool Aligner::padded(std::intptr_t document) const
    {
        const auto it = docs_.find(document);
        return it != docs_.end() && it->second->recipe && it->second->padded;
    }

    void Aligner::leave(std::intptr_t document) noexcept
    {
        const auto it = docs_.find(document);
        if (it != docs_.end() && !it->second->padded) it->second->recipe.reset();
    }

    bool Aligner::isMarker(std::intptr_t token) noexcept
    {
        return (token >> 16) == MARKER_TAG && (token & MARKER_IDS) != 0;
    }

    void Aligner::forget(std::intptr_t document) noexcept
    {
        docs_.erase(document);
    }

    void Aligner::setProgress(Progress progress, int delayMs, int intervalMs)
    {
        progress_ = std::move(progress);
        progressDelayMs_ = delayMs;
        progressIntervalMs_ = intervalMs;
    }

    SwitchedOn Aligner::switchOn(const Send& view, const Settings& settings, std::vector<Range>* inserted)
    {
        Doc& d = at(view);
        if (d.recipe) return d.padded ? SwitchedOn::Padded : SwitchedOn::StopsOnly;
        if (indicator_ < 0 || settings.delimiter.empty()) return SwitchedOn::NothingToAlign;
        if (view(SCI_GETREADONLY, 0, 0)) return SwitchedOn::ReadOnly;
        Busy busy(busy_);
        Meter meter(progress_, true, progressDelayMs_, progressIntervalMs_);
        meter.pass(5);

        auto recipe = std::make_shared<Recipe>();
        recipe->settings = settings;
        recipe->tabDelimited = settings.delimiter == "\t";
        // Spaces before numbers would turn into delimiters
        recipe->numbers = settings.alignNumbers && settings.delimiter.find(' ') == std::string::npos;

        bool anyDelimiter = false;
        std::vector<std::size_t> delims;
        const std::size_t dl = settings.delimiter.size();
        const Pos lines = view(SCI_GETLINECOUNT, 0, 0);
        for (Pos ln = 0; ln < lines; ++ln) {
            meter.at(ln, lines);
            const std::string_view text = lineText(view, ln);
            delimitersOf(text, settings, delims);
            if (delims.empty()) continue;
            anyDelimiter = true;
            if (!recipe->numbers) continue;
            for (std::size_t c = 0; c <= delims.size(); ++c) {
                const auto [s, e] = field(text, delims, dl, c);
                Number n;
                if (!numberIn(text, s, e, n)) continue;
                if (recipe->intWidth.size() <= c) recipe->intWidth.resize(c + 1, 0);
                recipe->intWidth[c] = (std::max)(recipe->intWidth[c], n.intDigits);
            }
        }
        if (!anyDelimiter) return SwitchedOn::NothingToAlign;

        const Fingerprint unpadded = fingerprint(view);
        meter.pass(50);
        const std::size_t runs = padAll(view, indicator_, recipe.get(), nullptr, inserted, &meter);
        meter.pass(10);
        measure(view, *recipe, &meter);
        meter.pass(35);
        setStops(view, 0, lines - 1, recipe->stops, &meter);

        d.recipe = recipe;
        d.padded = runs > 0;
        if (!d.padded) return SwitchedOn::StopsOnly;
        addMarker(view, d, { recipe, nullptr, unpadded, fingerprint(view), true });
        return SwitchedOn::Padded;
    }

    bool Aligner::switchOff(const Send& view, std::vector<Range>* removed)
    {
        Doc* d = find(view);
        if (!d || !d->recipe) return false;
        Busy busy(busy_);
        Meter meter(progress_, false, progressDelayMs_, progressIntervalMs_);
        if (!d->padded) {
            meter.pass(100);
            clearStops(view, &meter);
            d->recipe.reset();
            return true;
        }

        const Fingerprint padded = fingerprint(view);
        auto exceptions = std::make_shared<Exceptions>();
        meter.pass(85);
        stripAll(view, indicator_, d->recipe.get(), *exceptions, removed, &meter);
        meter.pass(15);
        clearStops(view, &meter);
        // No padding left to take out (the file was reloaded, say): no step for undo
        const Fingerprint unpadded = fingerprint(view);
        d->lastOff = unpadded == padded ? 0 : addMarker(view, *d, { d->recipe, exceptions, unpadded, padded, false });
        d->recipe.reset();
        d->padded = false;
        return true;
    }

    SwitchedOn Aligner::realign(const Send& view, const Settings& settings, std::vector<Range>* inserted)
    {
        Doc& d = at(view);
        if (d.recipe) return d.padded ? SwitchedOn::Padded : SwitchedOn::StopsOnly;
        const auto it = d.records.find(d.lastOff);
        if (it != d.records.end() && !view(SCI_GETREADONLY, 0, 0)) {
            const Record previous = it->second;
            Busy busy(busy_);
            if (fingerprint(view) == previous.unpadded) {
                Meter meter(progress_, true, progressDelayMs_, progressIntervalMs_);
                meter.pass(60);
                std::vector<Range> added;
                padAll(view, indicator_, previous.recipe.get(), previous.exceptions.get(), &added, &meter);
                if (fingerprint(view) == previous.padded) {
                    meter.pass(40);
                    d.recipe = applyStops(view, previous.recipe, &meter);
                    d.padded = true;
                    addMarker(view, d, { previous.recipe, previous.exceptions, previous.unpadded, previous.padded, true });
                    if (inserted) inserted->insert(inserted->end(), added.begin(), added.end());
                    return SwitchedOn::Padded;
                }
                Exceptions none;
                stripAll(view, indicator_, nullptr, none, nullptr);
            }
        }
        return switchOn(view, settings, inserted);
    }

    Replayed Aligner::replay(const Send& view, std::intptr_t token, bool undo, std::vector<Range>* changed)
    {
        if (!isMarker(token)) return Replayed::NotOurs;
        Doc& d = at(view);
        Busy busy(busy_);
        d.pending = nullptr;
        const auto it = d.records.find(token);
        if (it == d.records.end()) {
            resetHistory(view, d);
            return Replayed::Reset;
        }
        const Record record = it->second;
        const bool saved = savedBeforeStep(view, undo);
        std::vector<Range> ranges;

        if (record.padsOnRedo != undo) {
            if (d.recipe || fingerprint(view) != record.unpadded) {
                resetHistory(view, d);
                return Replayed::Reset;
            }
            Meter meter(progress_, true, progressDelayMs_, progressIntervalMs_);
            meter.pass(60);
            padAll(view, indicator_, record.recipe.get(), record.exceptions.get(), &ranges, &meter);
            if (fingerprint(view) != record.padded) {
                Exceptions none;
                stripAll(view, indicator_, nullptr, none, nullptr);
                resetHistory(view, d);
                return Replayed::Reset;
            }
            meter.pass(40);
            d.recipe = applyStops(view, record.recipe, &meter);
            d.padded = true;
            if (saved) view(SCI_SETSAVEPOINT, 0, 0);
            if (changed) *changed = std::move(ranges);
            return Replayed::Padded;
        }

        if (!d.padded || fingerprint(view) != record.padded) {
            resetHistory(view, d);
            return Replayed::Reset;
        }
        Meter meter(progress_, false, progressDelayMs_, progressIntervalMs_);
        meter.pass(85);
        Exceptions taken;
        stripAll(view, indicator_, record.recipe.get(), taken, &ranges, &meter);
        if (fingerprint(view) != record.unpadded) {
            padAll(view, indicator_, record.recipe.get(), &taken, nullptr);
            resetHistory(view, d);
            return Replayed::Reset;
        }
        meter.pass(15);
        clearStops(view, &meter);
        d.recipe.reset();
        d.padded = false;
        if (saved) view(SCI_SETSAVEPOINT, 0, 0);
        if (changed) *changed = std::move(ranges);
        return Replayed::Unpadded;
    }

    void Aligner::onModified(const Send& view, int modificationType, std::intptr_t position, std::intptr_t length,
        const char* text, std::intptr_t linesAdded)
    {
        constexpr int handled = SC_MOD_BEFOREDELETE | SC_MOD_DELETETEXT | SC_MOD_INSERTTEXT;
        if (busy_ || indicator_ < 0 || length <= 0 || !(modificationType & handled)) return;
        Doc* d = find(view);
        if (!d || !d->recipe) return;
        const auto ind = static_cast<std::uintptr_t>(indicator_);
        // A document in both views reports each change twice: the padding follows it once, the stops in each view
        const Change change{ modificationType, position, length, view(SCI_GETLENGTH, 0, 0) };
        const bool repeated = change == d->lastChange;
        d->lastChange = change;

        if (modificationType & SC_MOD_BEFOREDELETE) {
            if (repeated) return;
            d->pending = nullptr;
            if (!d->padded) return;
            std::vector<Range> marked;
            for (Pos p = position; p < position + length;) {
                Pos e = view(SCI_INDICATOREND, ind, p);
                if (e <= p) break;
                if (e > position + length) e = position + length;
                if (view(SCI_INDICATORVALUEAT, ind, p)) marked.push_back({ p, e });
                p = e;
            }
            if (marked.empty()) return;

            const char* p = reinterpret_cast<const char*>(view(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(position), length));
            if (!p) return;
            const std::string_view gone(p, static_cast<std::size_t>(length));
            Deleted entry;
            entry.position = position;
            entry.length = length;
            entry.lengthAfter = view(SCI_GETLENGTH, 0, 0) - length;
            entry.hash = fnv1a(gone.data(), gone.size());
            forEachPiece(gone, [&](std::size_t from, std::string_view piece) {
                Deleted::Piece out;
                out.hash = fnv1a(piece.data(), piece.size());
                out.offset = static_cast<std::int32_t>(from);
                out.length = static_cast<std::int32_t>(piece.size());
                const Pos pieceStart = position + static_cast<Pos>(from);
                const Pos pieceEnd = pieceStart + static_cast<Pos>(piece.size());
                for (const Range& m : marked) {
                    const Pos s = (std::max)(m.start, pieceStart);
                    const Pos e = (std::min)(m.end, pieceEnd);
                    if (s < e) out.marks.push_back({ static_cast<std::int32_t>(s - pieceStart), static_cast<std::int32_t>(e - s) });
                }
                entry.bytes += sizeof(Deleted::Piece) + out.marks.size() * sizeof(Mark);
                entry.pieces.push_back(std::move(out));
            });

            for (auto it = d->journal.begin(); it != d->journal.end(); ++it) {
                if (it->position == entry.position && it->lengthAfter == entry.lengthAfter && it->hash == entry.hash) {
                    d->journalBytes -= it->bytes;
                    d->journal.erase(it);
                    break;
                }
            }
            d->journalBytes += entry.bytes;
            d->journal.push_back(std::move(entry));
            while (d->journalBytes > MAX_JOURNAL_BYTES && d->journal.size() > 1) {
                d->journalBytes -= d->journal.front().bytes;
                d->journal.pop_front();
            }
            d->pending = &d->journal.back();
            return;
        }

        // What is left of a line after a deletion gets its stops
        if (modificationType & SC_MOD_DELETETEXT) {
            const Pos line = view(SCI_LINEFROMPOSITION, static_cast<std::uintptr_t>(position), 0);
            fitStops(view, line, line, *d->recipe);
            return;
        }

        if (!(modificationType & SC_MOD_INSERTTEXT)) return;
        const Deleted* pending = repeated ? nullptr : d->pending;
        if (!repeated) d->pending = nullptr;

        if (d->padded && !repeated) {
            std::string_view added;
            if (text) added = std::string_view(text, static_cast<std::size_t>(length));
            else if (const char* p = reinterpret_cast<const char*>(view(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(position), length)))
                added = std::string_view(p, static_cast<std::size_t>(length));
            const Pos lengthBefore = view(SCI_GETLENGTH, 0, 0) - length;
            const std::uint64_t hash = fnv1a(added.data(), added.size());

            KeepCurrent keep(view);
            view(SCI_SETINDICATORCURRENT, ind, 0);
            view(SCI_SETINDICATORVALUE, 1, 0);
            // Inserted text never becomes padding
            view(SCI_INDICATORCLEARRANGE, static_cast<std::uintptr_t>(position), length);

            const Deleted* back = nullptr;
            if (modificationType & SC_PERFORMED_UNDO) {
                for (auto it = d->journal.rbegin(); it != d->journal.rend(); ++it) {
                    if (it->position == position && it->length == length && it->lengthAfter == lengthBefore && it->hash == hash) {
                        back = &*it;
                        break;
                    }
                }
            }
            if (back) {
                // Undo brings deleted padding back
                for (const Deleted::Piece& piece : back->pieces)
                    fillMarks(view, position + piece.offset, piece.marks);
            }
            else if (pending && pending->position == position && pending->lengthAfter == lengthBefore) {
                // Replaced lines (sorted, moved): lines that came back unchanged keep their padding
                std::unordered_multimap<std::uint64_t, const Deleted::Piece*> byHash;
                for (const Deleted::Piece& piece : pending->pieces)
                    if (!piece.marks.empty()) byHash.emplace(piece.hash, &piece);
                forEachPiece(added, [&](std::size_t from, std::string_view piece) {
                    const auto range = byHash.equal_range(fnv1a(piece.data(), piece.size()));
                    for (auto it = range.first; it != range.second; ++it) {
                        if (it->second->length != static_cast<std::int32_t>(piece.size())) continue;
                        fillMarks(view, position + static_cast<Pos>(from), it->second->marks);
                        byHash.erase(it);
                        break;
                    }
                });
            }
        }

        // Inserted lines get stops; a cell that grew past its column moves only its own line on
        const Pos first = view(SCI_LINEFROMPOSITION, static_cast<std::uintptr_t>(position), 0);
        fitStops(view, first, first + (std::max)(linesAdded, static_cast<std::intptr_t>(0)), *d->recipe);
    }

    void Aligner::refreshStops(const Send& view)
    {
        Doc* d = find(view);
        if (!d || !d->recipe) return;
        Busy busy(busy_);
        Meter meter(progress_, true, progressDelayMs_, progressIntervalMs_);
        meter.pass(25);
        auto fresh = std::make_shared<Recipe>(*d->recipe);
        measure(view, *fresh, &meter);
        meter.pass(75);
        setStops(view, 0, view(SCI_GETLINECOUNT, 0, 0) - 1, fresh->stops, &meter);
        d->recipe = fresh;
    }

    void Aligner::restoreStops(const Send& view)
    {
        Doc* d = find(view);
        if (!d) return;
        Busy busy(busy_);
        if (!d->recipe) {
            clearStops(view);
            return;
        }
        Meter meter(progress_, true, progressDelayMs_, progressIntervalMs_);
        meter.pass(100);
        d->recipe = applyStops(view, d->recipe, &meter);
    }

    void Aligner::ensureStops(const Send& view)
    {
        const Doc* d = find(view);
        if (!d || !d->recipe || d->recipe->stops.empty() || view(SCI_GETNEXTTABSTOP, 0, 0) > 0) return;
        restoreStops(view);
    }

    std::string Aligner::withoutPadding(const Send& view, std::intptr_t start, std::intptr_t end) const
    {
        std::string out;
        if (end <= start) return out;
        const auto copy = [&](Pos from, Pos to) {
            if (to <= from) return;
            if (const char* p = reinterpret_cast<const char*>(view(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(from), to - from)))
                out.append(p, static_cast<std::size_t>(to - from));
        };
        if (!padded(view)) {
            copy(start, end);
            return out;
        }
        const auto ind = static_cast<std::uintptr_t>(indicator_);
        for (Pos p = start; p < end;) {
            Pos e = view(SCI_INDICATOREND, ind, p);
            if (e <= p) { copy(p, end); break; }
            if (e > end) e = end;
            if (!view(SCI_INDICATORVALUEAT, ind, p)) copy(p, e);
            p = e;
        }
        return out;
    }

    bool Aligner::copiedWithoutPadding(const Send& view, std::string& text) const
    {
        const Doc* d = find(view);
        if (!d || !d->recipe || !d->padded || indicator_ < 0) return false;

        // The selection as Scintilla copies it: the ranges in order, sorted in a rectangle, each followed by the line
        // end when there are several; with nothing selected the caret line and the line end
        static constexpr std::string_view eols[] = { "\r\n", "\r", "\n" };
        const Pos mode = view(SCI_GETEOLMODE, 0, 0);
        const std::string_view eol = eols[(mode >= 0 && mode <= 2) ? mode : 0];
        std::vector<Range> ranges;
        const Pos count = view(SCI_GETSELECTIONS, 0, 0);
        for (Pos i = 0; i < count; ++i) {
            const auto n = static_cast<std::uintptr_t>(i);
            ranges.push_back({ view(SCI_GETSELECTIONNSTART, n, 0), view(SCI_GETSELECTIONNEND, n, 0) });
        }
        bool lineEnds = ranges.size() > 1;
        if (std::all_of(ranges.begin(), ranges.end(), [](const Range& r) { return r.end <= r.start; })) {
            const Pos line = view(SCI_LINEFROMPOSITION, static_cast<std::uintptr_t>(view(SCI_GETCURRENTPOS, 0, 0)), 0);
            ranges.assign(1, { view(SCI_POSITIONFROMLINE, static_cast<std::uintptr_t>(line), 0),
                view(SCI_GETLINEENDPOSITION, static_cast<std::uintptr_t>(line), 0) });
            lineEnds = true;
        }
        else if (view(SCI_SELECTIONISRECTANGLE, 0, 0)) {
            std::sort(ranges.begin(), ranges.end(), [](const Range& a, const Range& b) { return a.start < b.start; });
        }
        std::size_t at = 0;
        bool selection = true;
        for (const Range& r : ranges) {
            const auto n = static_cast<std::size_t>((std::max)(r.end - r.start, Pos(0)));
            const char* p = n ? reinterpret_cast<const char*>(view(SCI_GETRANGEPOINTER, static_cast<std::uintptr_t>(r.start), r.end - r.start)) : "";
            const std::size_t e = lineEnds ? eol.size() : 0;
            if (!p || at + n + e > text.size() || text.compare(at, n, p, n) != 0 || text.compare(at + n, e, eol.data(), e) != 0) {
                selection = false;
                break;
            }
            at += n + e;
        }
        if (selection && at == text.size()) {
            std::string out;
            out.reserve(text.size());
            for (const Range& r : ranges) {
                out += withoutPadding(view, r.start, r.end);
                if (lineEnds) out += eol;
            }
            if (out.size() == text.size()) return false;
            text = std::move(out);
            return true;
        }

        // The text just cut: the last change, a deletion the journal kept together with its padding
        if (d->journal.empty() || !(d->lastChange.type & SC_MOD_DELETETEXT)) return false;
        const Deleted& cut = d->journal.back();
        if (d->lastChange.position != cut.position || d->lastChange.length != cut.length || d->lastChange.documentLength != cut.lengthAfter
            || text.size() != static_cast<std::size_t>(cut.length) || fnv1a(text.data(), text.size()) != cut.hash)
            return false;
        std::string out;
        out.reserve(text.size());
        std::size_t from = 0;
        for (const Deleted::Piece& piece : cut.pieces) {
            for (const Mark& m : piece.marks) {
                const auto s = static_cast<std::size_t>(piece.offset) + static_cast<std::size_t>(m.offset);
                out.append(text, from, s - from);
                from = s + static_cast<std::size_t>(m.length);
            }
        }
        out.append(text, from, std::string::npos);
        if (out.size() == text.size()) return false;
        text = std::move(out);
        return true;
    }
}
