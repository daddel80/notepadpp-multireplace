// QA harness for the Find/Replace-in-Files result wording.
//
// Mirrors the shipped composition 1:1:
//   MultiReplace::buildSkipBreakdown / buildScanSuffix / buildSkipSentence
//   (MultiReplacePanel.cpp) plus the status-line and dock-header assembly.
// String values are the English defaults verbatim from language_mapping.cpp;
// LM_get() reproduces LanguageManager::get()'s $REPLACE_STRINGn substitution
// (highest index first, then the bare form), so placeholder handling is tested
// the same way the plugin does it.
//
// Build: g++ -std=c++20 -Wall -Wextra -o scan_summary_qa scan_summary_qa.cpp
#include <cstdio>
#include <string>
#include <vector>
#include <algorithm>

// ----------------------------------------------------------- language strings
// verbatim from src/language_mapping.cpp (English base = fallback for all)
static const std::string L_replace_summary =
    "Replace in files: $REPLACE_STRING1 of $REPLACE_STRING2 file(s) modified.";
static const std::string L_scan_suffix   = " [$REPLACE_STRING1 file(s) searched$REPLACE_STRING2]";
static const std::string L_scan_skipped  = ", $REPLACE_STRING1 skipped: $REPLACE_STRING2";
static const std::string L_status_skipped= " $REPLACE_STRING1 file(s) skipped: $REPLACE_STRING2.";
static const std::string L_open_unsaved  = " $REPLACE_STRING of them in open document(s), not saved.";
static const std::string L_canceled      = "Canceled";
static const std::string L_single_header =
    "Search \"$REPLACE_STRING1\" ($REPLACE_STRING2 hits in $REPLACE_STRING3 file(s))";
static const std::string L_skip_binary      = "binary";
static const std::string L_skip_large       = "too large";
static const std::string L_skip_unreadable  = "unreadable";
static const std::string L_skip_undecodable = "not decodable";
static const std::string L_skip_readonly    = "read-only";
static const std::string L_skip_open_unsaved= "open with unsaved changes";
static const std::string L_skip_unencodable = "not encodable";

// the pre-change wording, kept so the tests can show the difference
static const std::string OLD_scan_suffix = " [of $REPLACE_STRING1 searched$REPLACE_STRING2]";

// VERBATIM behavior of LanguageManager::get(): substitute $REPLACE_STRINGn for
// n = size..1 (descending, so $..1 never eats the '1' of $..12), then the bare form.
static std::string LM_get(const std::string& tpl, const std::vector<std::string>& repl = {}) {
    std::string r = tpl;
    const std::string base = "$REPLACE_STRING";
    for (size_t i = repl.size(); i > 0; --i) {
        const std::string ph = base + std::to_string(i);
        for (size_t p = r.find(ph); p != std::string::npos; p = r.find(ph, p)) {
            r.replace(p, ph.size(), repl[i - 1]);
            p += repl[i - 1].size();
        }
    }
    const std::string v = repl.empty() ? std::string() : repl[0];
    for (size_t p = r.find(base); p != std::string::npos; p = r.find(base, p)) {
        r.replace(p, base.size(), v);
        p += v.size();
    }
    return r;
}

// ------------------------------------------------------------- skip bookkeeping
struct Guard {
    size_t binary = 0, large = 0, unreadable = 0, undecodable = 0;
    size_t readOnly = 0, openUnsaved = 0, unencodable = 0;
    size_t total() const {
        return binary + large + unreadable + undecodable + readOnly + openUnsaved + unencodable;
    }
};

// VERBATIM: MultiReplace::buildSkipBreakdown
static std::string buildSkipBreakdown(const Guard& g) {
    std::string b;
    auto add = [&](size_t n, const std::string& word) {
        if (n == 0) return;
        if (!b.empty()) b += ", ";
        b += std::to_string(n) + " " + word;
    };
    add(g.binary, L_skip_binary);
    add(g.large, L_skip_large);
    add(g.unreadable, L_skip_unreadable);
    add(g.undecodable, L_skip_undecodable);
    add(g.readOnly, L_skip_readonly);
    add(g.openUnsaved, L_skip_open_unsaved);
    add(g.unencodable, L_skip_unencodable);
    return b;
}

// VERBATIM: MultiReplace::buildScanSuffix
static std::string buildScanSuffix(const Guard& g, size_t searched, bool useOldWording = false) {
    std::string clause;
    if (g.total() > 0)
        clause = LM_get(L_scan_skipped, { std::to_string(g.total()), buildSkipBreakdown(g) });
    return LM_get(useOldWording ? OLD_scan_suffix : L_scan_suffix,
                  { std::to_string(searched), clause });
}

// VERBATIM: MultiReplace::buildSkipSentence
static std::string buildSkipSentence(const Guard& g) {
    if (g.total() == 0) return std::string();
    return LM_get(L_status_skipped, { std::to_string(g.total()), buildSkipBreakdown(g) });
}

// ------------------------------------------------- call-site assembly (verbatim)
// handleReplaceInFiles(): searched = files the loop reached, minus everything
// the guard recorded as skipped; idx < candidates when the run was canceled.
// changedOpenUnsaved is the part of changed that was edited in an open
// document with unsaved changes and therefore left unsaved.
static std::string replaceStatusLine(size_t candidates, size_t reachedIdx, size_t changed,
                                     const Guard& g, bool canceled, size_t changedOpenUnsaved = 0) {
    (void)candidates;
    const size_t skipped = g.total();
    const size_t searched = (reachedIdx > skipped) ? (reachedIdx - skipped) : 0;
    std::string msg = LM_get(L_replace_summary, { std::to_string(changed), std::to_string(searched) });
    if (changedOpenUnsaved > 0)
        msg += LM_get(L_open_unsaved, { std::to_string(changedOpenUnsaved) });
    msg += buildSkipSentence(g);
    if (canceled) msg += " - " + L_canceled;
    return msg;
}

// handleFindInFiles(): dock header + scan suffix
static std::string findDockHeader(const std::string& pattern, size_t hits, size_t filesWithHits,
                                  size_t reachedIdx, const Guard& g) {
    const size_t searched = (reachedIdx > g.total()) ? (reachedIdx - g.total()) : 0;
    std::string h = LM_get(L_single_header,
        { pattern, std::to_string(hits), std::to_string(filesWithHits) });
    return h + buildScanSuffix(g, searched);
}

// --------------------------------------------------------------------- checks
static int failures = 0;
static void CHECK(const std::string& desc, bool ok) {
    std::printf("%-4s %s\n", ok ? "PASS" : "FAIL", desc.c_str());
    if (!ok) ++failures;
}
// Catches "modified.," / ".." / " ." style seams from concatenating parts.
static bool wellFormed(const std::string& s) {
    static const char* bad[] = { "..", ".,", " ,", " .", "::", ",,", "  ", "[]", "( )" };
    for (const char* b : bad) if (s.find(b) != std::string::npos) return false;
    return !s.empty();
}

int main() {
    std::printf("=== Replace in Files - status line ===\n\n");

    // A) the reported screenshot case: one file, no hits, nothing skipped
    {
        Guard g{};
        const std::string now = replaceStatusLine(1, 1, 0, g, false);
        const std::string before = LM_get(L_replace_summary, { "0", "1" })
                                 + buildScanSuffix(g, 1, /*useOldWording=*/true);
        std::printf("  before: %s\n  now:    %s\n\n", before.c_str(), now.c_str());
        CHECK("A no skips -> no bracket suffix at all", now.find('[') == std::string::npos);
        CHECK("A reads as one sentence", now == "Replace in files: 0 of 1 file(s) modified.");
        CHECK("A well formed", wellFormed(now));
    }

    // B) skips present: breakdown becomes its own sentence
    {
        Guard g{ 38, 2, 1, 0 };
        const std::string now = replaceStatusLine(1772, 1772, 3, g, false);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("B denominator excludes the 41 skipped (1772-41=1731)",
              now.find("3 of 1731 file(s) modified.") != std::string::npos);
        CHECK("B skip breakdown is a separate sentence",
              now.find(" 41 file(s) skipped: 38 binary, 2 too large, 1 unreadable.") != std::string::npos);
        CHECK("B well formed", wellFormed(now));
    }

    // C) read-only files were reached but never searched
    {
        Guard g{}; g.readOnly = 3;
        const std::string now = replaceStatusLine(10, 10, 2, g, false);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("C read-only files excluded from the searched count (10-3=7)",
              now.find("2 of 7 file(s) modified.") != std::string::npos);
        CHECK("C read-only is a breakdown item, not its own sentence",
              now.find(" 3 file(s) skipped: 3 read-only.") != std::string::npos);
        CHECK("C well formed", wellFormed(now));
    }

    // D) canceled midway: only what was reached may be claimed as searched
    {
        Guard g{ 5, 0, 0, 0 };
        const std::string now = replaceStatusLine(1772, 100, 7, g, true);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("D canceled run reports reached-minus-skipped (100-5=95), not 1772",
              now.find("7 of 95 file(s) modified.") != std::string::npos);
        CHECK("D cancel marker kept", now.find("- Canceled") != std::string::npos);
        CHECK("D well formed", wellFormed(now));
    }

    // E) everything skipped -> honest zero, no underflow
    {
        Guard g{ 4, 0, 0, 0 };
        const std::string now = replaceStatusLine(4, 4, 0, g, false);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("E all skipped -> 0 of 0, no underflow",
              now.find("0 of 0 file(s) modified.") != std::string::npos);
        CHECK("E well formed", wellFormed(now));
    }

    std::printf("=== Find in Files - dock header ===\n\n");

    // F) no skips
    {
        Guard g{};
        const std::string now = findDockHeader("Fi", 4, 2, 7, g);
        const std::string before = LM_get(L_single_header, { "Fi", "4", "2" })
                                 + buildScanSuffix(g, 7, /*useOldWording=*/true);
        std::printf("  before: %s\n  now:    %s\n\n", before.c_str(), now.c_str());
        CHECK("F dangling \"of\" is gone", now.find("[of ") == std::string::npos);
        CHECK("F states what the number counts",
              now.find("[7 file(s) searched]") != std::string::npos);
        CHECK("F well formed", wellFormed(now));
    }

    // G) with skips
    {
        Guard g{ 38, 2, 1, 0 };
        const std::string now = findDockHeader("Fi", 900, 120, 1772, g);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("G searched count excludes skips (1772-41=1731)",
              now.find("[1731 file(s) searched,") != std::string::npos);
        CHECK("G breakdown inside the bracket",
              now.find("41 skipped: 38 binary, 2 too large, 1 unreadable]") != std::string::npos);
        CHECK("G well formed", wellFormed(now));
    }

    // H) undecodable bucket reachable (Replace verifies a lossless roundtrip)
    {
        Guard g{ 0, 0, 0, 6 };
        const std::string now = replaceStatusLine(20, 20, 1, g, false);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("H undecodable counted and named",
              now.find("6 file(s) skipped: 6 not decodable.") != std::string::npos);
        CHECK("H well formed", wellFormed(now));
    }

    // J) replacement not representable in the file's codepage: refused, counted, named
    {
        Guard g{}; g.unencodable = 2;
        const std::string now = replaceStatusLine(10, 10, 5, g, false);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("J unencodable files leave the denominator (10-2=8)",
              now.find("5 of 8 file(s) modified.") != std::string::npos);
        CHECK("J unencodable counted and named",
              now.find(" 2 file(s) skipped: 2 not encodable.") != std::string::npos);
        CHECK("J well formed", wellFormed(now));
    }

    // K) every reason at once: one summary, one skip sentence, one breakdown
    {
        Guard g{ 1, 2, 0, 0 }; g.readOnly = 2; g.openUnsaved = 1; g.unencodable = 2;
        const std::string now = replaceStatusLine(1225, 1225, 12, g, true);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("K exact combined wording",
              now == "Replace in files: 12 of 1217 file(s) modified."
                     " 8 file(s) skipped: 1 binary, 2 too large, 2 read-only,"
                     " 1 open with unsaved changes, 2 not encodable. - Canceled");
        auto count = [&](const char* word) {
            size_t n = 0;
            for (size_t p = now.find(word); p != std::string::npos; p = now.find(word, p + 1)) ++n;
            return n;
        };
        CHECK("K 'skipped' appears exactly once", count("skipped") == 1);
        CHECK("K 'file(s)' appears exactly twice", count("file(s)") == 2);
        CHECK("K well formed", wellFormed(now));
    }

    // L) open documents with unsaved changes are changed, not skipped: the
    //    addendum names them, the skip sentence stays a single sentence
    {
        Guard g{ 1, 2, 0, 0 }; g.readOnly = 2; g.unencodable = 2;
        const std::string now = replaceStatusLine(1225, 1225, 12, g, true, 2);
        std::printf("  now:    %s\n\n", now.c_str());
        CHECK("L exact wording with unsaved open documents",
              now == "Replace in files: 12 of 1218 file(s) modified."
                     " 2 of them in open document(s), not saved."
                     " 7 file(s) skipped: 1 binary, 2 too large, 2 read-only,"
                     " 2 not encodable. - Canceled");
        auto count = [&](const char* word) {
            size_t n = 0;
            for (size_t p = now.find(word); p != std::string::npos; p = now.find(word, p + 1)) ++n;
            return n;
        };
        CHECK("L 'skipped' appears exactly once", count("skipped") == 1);
        CHECK("L 'file(s)' appears exactly twice", count("file(s)") == 2);
        CHECK("L well formed", wellFormed(now));
    }

    // M) addendum alone, nothing skipped, not canceled
    {
        Guard g{};
        const std::string now = replaceStatusLine(10, 10, 3, g, false, 3);
        CHECK("M addendum without skip sentence",
              now == "Replace in files: 3 of 10 file(s) modified."
                     " 3 of them in open document(s), not saved.");
        CHECK("M well formed", wellFormed(now));
    }

    // N) no unsaved open documents: no addendum at all
    {
        Guard g{};
        const std::string now = replaceStatusLine(10, 10, 3, g, false, 0);
        CHECK("N no addendum", now == "Replace in files: 3 of 10 file(s) modified.");
    }

    // I) placeholder safety: substitution must not corrupt multi-digit numbers
    {
        Guard g{ 12, 0, 0, 0 };
        const std::string s = buildScanSuffix(g, 1234);
        CHECK("I multi-digit counts survive $REPLACE_STRINGn substitution",
              s == " [1234 file(s) searched, 12 skipped: 12 binary]");
    }

    std::printf("\n%s (%d failure%s)\n", failures == 0 ? "ALL CHECKS PASSED" : "CHECKS FAILED",
                failures, failures == 1 ? "" : "s");
    return failures == 0 ? 0 : 1;
}
