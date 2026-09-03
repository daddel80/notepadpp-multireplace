# One-Pass Replace All: stabilization plan

Working document for the staged rollout of `src/OnePassHits.h/.cpp`. Each stage
ends in a decision point and can be shipped or reverted on its own. Work top to
bottom; do not start a stage before the previous one meets its exit criteria.

> **Status: ready for release.** Stage 7, the manual checklist
> (`ONEPASS-CHECKLIST.md`), has been run by the owner in real Notepad++: all 20
> numbered cases green, including one live catch mid-run (an unrebuilt DLL made
> C3 look like a regression - a fresh build cleared it, confirmed by both the
> on-pass and the reference run). J5 (undo: one Ctrl+Z restores the whole run in
> one step) was run separately as the highest-value item from HANDOVER section
> 8.3 and is also green. J3/J4 (formula skip/error) and multi-document Replace
> All were not run - see the open item below. It took eight findings to get
> here (the last found by reading, not by a test); section 11 of
> `ONEPASS-HANDOVER.md` analyses why, and that analysis is the useful part. The
> feature can be reduced later (`remember = false`, one word) or removed (one
> file plus ~40 panel lines). `ONEPASS-HANDOVER.md` is the entry point for
> review; `ONEPASS-CHECKLIST.md` is the list that was run.

## Where we are

| Item | State |
| --- | --- |
| `src/OnePassHits.h/.cpp` | written; the constructor's `remember` parameter is a test lever only, the plugin always passes the default; Stages 1 to 6 of this plan worked through (results below) |
| `src/MultiReplacePanel.cpp` | `onePassReplaceAll` binds the class through `OnePassDoc` lambdas, ~30 lines, and constructs it with remembering on; there is no setting |
| `README.md` | one-pass note |
| `src/tests/onepass_qa.cpp` | standalone, builds with g++/cl against the real class; sections A to L3 |
| `src/tests/onepass_engine_qa.cpp` | same checks on the real Scintilla document and the N++ Boost bridge; sections A to L3 plus K (classification completeness), `bench`, `trace` |
| Fixed along the way | empty matches (`^ $ \b` lookarounds) are replaced like Replace All; the EOF freeze (`^`, `x*`) is gone; skipped hits no longer move the walk past other entries |
| Found by the analytical passes (Stages 1, 6 and 6b) | eight defects the fuzz had not reached, all fixed and each pinned by a fixed test case except F8, which was removed rather than pinned: see "Findings" |
| Still open | Stage 8 (ship). Stage 7 ran green (20/20 cases plus J5/undo); J3/J4 (formula skip/error) and multi-document Replace All were not run - see `ONEPASS-HANDOVER.md` section 9. |
| `src/MultiReplacePanel.cpp` (Stage 6) | `performSingleSearch` marks a refused engine answer; `replaceOne` takes the position its confirming search starts from |

## Findings of the analytical pass

Each of these was reasoned out from the code and the engine's behaviour first, then
confirmed with a probe against the real Boost bridge, then pinned as a fixed test
case that fails on the old class and passes on the new one. None of them had been
hit by 15 seeds of fuzzing with 600 cases each, which is the practical argument
for reading the code instead of only running it.

| # | Finding | Fix | Pinned by |
| --- | --- | --- | --- |
| F1 | Boost treats a form feed (`\f`) as a line separator for `^`, `$` and `.`; the context classes only knew `\r` and `\n`, so a form feed inserted right before a remembered `^` entry was not re-checked and the new match was missed. | `\f` maps to the `Lf` class. | engine C30, standalone A15 |
| F2 | Notepad++ can set custom word characters (`SCI_SETWORDCHARS`, Preferences > Delimiter); Scintilla's whole-word check follows that table, the class used a fixed alnum + `_` classification, so freeing a whole-word hit by replacing `-` with a space went unnoticed. | whole-word literal entries compare the actual byte before the replacement end; regex entries keep the fixed classes, because the Boost bridge ignores the document's word table. | engine C32, standalone A16, engine section H |
| F3 | Boost's search restart heuristic and its per-position matcher disagree: before a final `\f`, plain `\Z` reports no match, `\G(?:\Z)` reports one. The anchored probe put a hit into the cache that the verification search then refused (an extra "found", no replacement). | structural: every remembered hit now comes from a plain search (or its shift); the anchored probe is only a filter that decides whether that plain search is run. The cache can no longer disagree with the verification by construction. | engine section D with `\f` in the texts |
| F4 | Column scope: a replacement that inserts the delimiter (`a` to `a,`) moves the rest of the line into another column, so text that did not change can enter or leave the search scope. Invalidating hits on that line was not enough: the line can gain a hit for an entry whose remembered hit was on a later line. | in column scope nothing remembered survives a replacement (as in selection scope). Remembered hits still survive skips. Column mode therefore keeps today's cost per step. | engine section E with delimiter-inserting entries |
| F5 | The regex bridge decodes the text from wherever the search starts. If a replacement ends inside a character, the untouched bytes after it belong to a sequence that now starts inside the replacement and decode to other characters than before, so a remembered hit and the classification of the byte before it are both stale. Only reachable in malformed UTF-8 (`x` to `\xC3` in front of a lone `\xA4`). | a trail byte at the replacement end, or a walk position on one, drops everything remembered (`insideChar`). Not reached in an ANSI document, where every byte is a character. | engine C34 to C37, section M4 |
| F6 | The bridge can answer a search with a match that ends **before** it starts (`\b` in malformed UTF-8), and with -2 or -3 for an invalid pattern or an internal exception. `performSingleSearch` refuses all three, which the walk could not tell apart from "there is nothing here" and used as proof that the entry has no hits left. | `performSingleSearch` sets `_searchRefused`; the walk's document view reports a refused search as length -1 and never concludes from it that an entry is exhausted. | engine section M4 (0 of 600 differences after the fix, 111 refused answers observed) |
| F8 | `reset()` (an edit from outside the run) cleared the empty-match bars but not `_netDelta`, the offset that maps a current position to its original one. With a stale offset the sentinel -1 of `lastEmptyOrig` can be hit by a real position, and one empty match would be barred that should not be. Found by reading, not by a test. | `reset()` restarts the mapping with the bars. The collision is then impossible rather than unlikely, so there is no test case for it; C55 covers the path (growth, then outside edits). | removed by construction, path covered by C55 |
| F7 | An entry whose match is bound to the search start (`\K`, `\G`, lookbehind) is not visible from its own match position, so `replaceOne`'s confirming search refused every such hit: one-pass counted `a\Kb` hits and replaced none, while Replace All with the same entry replaced them. Pre-existing, not caused by the bookkeeping. | `replaceOne` takes the position the confirming search starts from; the walk passes the position it searched from for those entries (`OnePassHits::verifyFrom`), the match itself for all others. The confirmation is still an independent search that must return the same position and length. | engine B9a to B9c (`a\Kb`, `\Ka`, `(?<=a)b` against Replace All), C41 |

The safety property held throughout: F3 and F4 could produce an extra counted
hit, F1, F2, F5, F6, F7 and F8 a missed replacement; none could produce a
replacement of text the entry does not match.

Why they kept coming: every round until Stage 6 widened the testing, and
widening has a long tail. F7 in particular sat behind a test exclusion - the
comparison against Notepad++'s Replace All skipped `\K` patterns, so the one
defect the comparison existed to find was the one it could not see. Stage 6b
therefore stops widening and closes the contract instead: section P compares
**every** pattern the harness knows, with no exclusion, and the table below
lists every assumption with the way it is discharged.

## The safety property this plan rests on

`replaceOne` re-searches at the proposed position with the entry's own flags and
refuses the replacement when position or length differ. A wrong proposal from
`OnePassHits` can therefore cost a **missed** or **reordered** replacement, never
a replacement of text the entry does not match. Corrupted output in the sense of
"wrote over something that never matched" is structurally excluded. Keep this
property: it is what makes the staged rollout defensible. Any future change that
removes the verification search removes the safety net with it.

---

## Stage 0: reference baseline (harness only)

The plain walk, `remember = false` in the constructor, is the baseline: the walk
with corrected empty-match semantics and no bookkeeping at all. Behaviour equals
the old one-pass minus the two bugs, with the old cost per step. It is reachable
only from the two harnesses, where every differential section compares the
remembering run against it; the plugin does not expose it. Build in VS (x64
Release and Debug) and run cases 1, 3, 5, 14 of Stage 7 before anything else.

## Stage 1: write the invariants down (done)

Goal: make the correctness argument explicit so future changes can be checked
against it instead of re-derived. Every row was read against `OnePassHits.cpp`;
"holds" means confirmed, the findings above came out of the rows marked with them.

| # | Invariant | Rests on |
| --- | --- | --- |
| I1 | The walk position never decreases; a step ends at the hit end, plus delta when replaced. | `onePassReplaceAll` loop |
| I2 | A proposed hit is only ever a proposal; `replaceOne` verifies it. | see safety property above |
| I3 | After a replacement, text at and after the replacement end is untouched, so a remembered hit strictly after the old end keeps its content and only moves by delta. Holds in full scope only: in column scope the rest of the line can change column (F4), in selection scope the ranges move. | `afterReplace`, shift branch; `broken` for the other scopes |
| I4 | A match that depends on the character before its start can only newly appear at the replacement end, which holds as long as the replacement ends at a character boundary (F5). Regex entries: the byte before is classified (Bof, Cr, Lf incl. form feed, Word, Other, NonAscii), a class change triggers the probe; F1 came out of this row. Whole-word literals: any change of the byte triggers it, because the word table is the host's; F2 came out of this row. The probe only gates a plain search, which alone decides; F3 came out of this row. | `ctxStart`, `classAt`, `anchoredAt`, `keep`, `search` |
| I5 | An entry that found nothing from p finds nothing later, provided the engine answered the search at all (F6): a refused answer is not an answer. Measured as K4. a match created by a replacement starts at or before the replacement end, the walk continues at that end, and a non-`ctxStart` match at or after it lies in untouched text and would have been found by the earlier search. Holds. | `none`, restricted to `Scope::Full` |
| I6 | A changed line count invalidates everything, including a CRLF pair split or joined by a replacement. Holds; note that a form feed is not a Scintilla line end, which is why it needs the class in I4. | `broken`, `lineCount()` |
| I7 | In column and selection scope the end of a range is the end of the text for the engine, so an empty hit at such an end is invisible from a search that starts there and is not remembered. Holds. | `keep()`, `r.len == 0 && _scope != Scope::Full` |
| I8 | An entry never takes an empty match at its own `consumedEnd`, and never twice at the same original position; this is what terminates nullable patterns. Holds; the original position of the walk never decreases, each entry gets at most one empty event per original position, non-empty hits consume text, so the number of steps is bounded. | `emptyBarred`, `lastEmptyOrig`, `_netDelta` |
| I9 | `(?<=`, `(?<!`, `\K`, `\G` reach outside the match or bind to the search start and are never remembered. Holds; named groups `(?<name>` are correctly not matched by the `(?<=` and `(?<!` tokens. | `uncached` |
| I10 | An edit from outside the walk drops everything remembered. Holds, and the length check is sufficient: modification events are masked during the run and the formula engine host interface has no document access (S6). | `reset()`, length check in `onePassReplaceAll` |
| I11 | The engine is stateless for the walk: `EMPTYMATCH_ALLOWATSTART` bypasses the bridge's continuation logic, and the `_lastDirection` quirk for empty ranges gives the same result either way. So the searches issued with remembering on and off cannot influence each other's results. | `OnePassEntry::flags` |
| I12 | Every remembered hit is the result of a plain search from some position, or its shift. Introduced with F3; it is what makes I2 airtight. | `keep()` |
| I13 | A hit found by a search from p is found again by a search from its own position, so `replaceOne` can confirm it there. Measured as K5: true for every construct except those bound to the search start, which are the ones `verifyFrom` handles (F7). | `verifyFrom()`, K5 |
| I14 | Every search start the walk uses is a character boundary for the regex bridge, so two searches see the same characters. Malformed text is the exception and is handled by dropping everything remembered (F5). | `insideChar()`, `forget()` |
| I15 | A position and its original counterpart differ by `_netDelta`, which is only meaningful while the walk owns every change. An edit from outside restarts the mapping together with the bars (F8). | `reset()` |

## Stage 2: panel seams (done by code reading; the manual cases remain in Stage 7)

These are independent of the bookkeeping and affect the walk either way.

| # | Seam | What to check |
| --- | --- | --- |
| S1 | `replaceOne` outcomes | Confirmed. Verification mismatch returns `false` with `newPos` untouched (-1) and a differing `verify`; engine `skip()` returns `false` with `newPos` set; engine error, missing engine and debug stop (`LuaEngine` sets `success = false` when the debug window stops) return `false` with `newPos` -1 and a matching `verify`. The loop maps these to skip, skip, abort. Non-formula entries always replace once verified. |
| S2 | Start position | Confirmed. `computeAllStartPos`: 0 with wrap, caret without wrap and "from cursor", first stored selection in selection mode. The walk starts there, hits before it are never searched. Fuzzed in sections J (engine) and H (standalone). |
| S3 | Column delimiter bookkeeping | Confirmed, and no longer load-bearing: since F4 nothing remembered survives a replacement in column scope, so the only requirement is that the next column search sees the updated delimiters, which `replaceOne` guarantees by calling `updateLineDelimiterAfterReplace` right after the edit. |
| S4 | Selection scope | Confirmed. `updateSelectionScope()` fills `m_selectionScope` before the run, `startCtx.useStoredSelections` is set, `adjustSelectionScope` moves the ranges after each replacement, the selection is restored from them after the run. |
| S5 | Codepage | Confirmed in code and fuzzed in engine section I (code page 0 with single high bytes). Boost's ANSI path classifies 0xE4 as non-word; the class treats every byte at or above 0x80 as "context changed", which is conservative. |
| S6 | Length preserving outside edit | Closed. `ILuaEngineHost` offers `escapeForRegex`, the debug window, list refresh, error dialogs and three setting queries; neither engine sends a Scintilla message. A script cannot edit the document, so no same-length edit can happen inside the run and the length check is sufficient. |
| S7 | Counting and UI | Confirmed: `updateCountColumns(w, findTotals[w], replTotals[w])` after every step, `selectListItem` returns early while `_bulkReplaceInProgress`, `ScopedUndoAction` and a masked `SCI_SETMODEVENTMASK` wrap the whole run in `handleReplaceAllButton`. |
| S8 | Construction | The panel constructs `OnePassHits(doc, entries, scope)` and leaves the `remember` parameter at its default. No setting, no INI key: the plain walk exists only as the reference in the two harnesses. Nothing else in the panel depends on the class beyond the five calls in the loop (`next`, `beforeReplace`, `afterReplace`, `afterSkip`, `reset`). |

## Stage 3: widen the automated tests (done)

Added to both QA files: start position inside the text, custom word characters,
form feed, tab, `-` and `x.y` in the texts, delimiter-inserting replacements
(`a` to `a,`), entries that delete or insert form feeds, disabled entries in
random lists, match lists combined with every scope, and the level pools of
Stage 4. Engine harness only: an ANSI code page section with single high bytes.
Not added: larger documents beyond 3000 characters, because the walk's cost per
step, not its correctness, depends on the size, and the `bench` command covers
that; lists made only of empty-capable entries appear often enough in the random
lists (the pool is about half empty-capable).

Result: engine harness 452 fixed checks plus 10 fuzz sections, 14 seeds with 600
cases each and one seed with texts up to 3000 characters, 0 mismatches.
Standalone harness 16 fixed checks plus 9 fuzz sections, 6 seeds with 2000 to
3000 cases each, 0 mismatches.

## Stage 4: remembering by entry class (done as test sections, not as a knob)

Decision: no level parameter in the production code. The isolation the levels
were meant to give comes from restricting the fuzz pool instead: sections L1,
L2 and L3 run lists drawn only from the entry class of that level, so every
mechanism is exercised alone before the full pool (section D) combines them. A
production knob would have added a parameter that only tests use, and the
`remember` lever of the harnesses already isolates the whole bookkeeping.

| Level | Remembered | Mechanism added | Bugs found there so far |
| --- | --- | --- | --- |
| 0 | nothing | plain walk (Stage 0) | none |
| 1 | literals without whole word | shift by delta, invalidate overlapping | none |
| 2 | plus regex without context tokens | same mechanism, other engine | none |
| 3 | plus whole word and `\b ^ $ \A \Z` entries | `classAt`, `anchoredAt`, `\G` | 1 of 6 |
| 4 | plus entries that can match empty | bars, sidestep, `_netDelta`, `none` at the end of text | 3 of 6 |

Result: L1, L2, L3 and D green over all seeds in both QA files.

## Stage 5: what ships (decided)

Two decisions.

**Which entries are remembered when remembering is on: all of them, including
empty-capable ones.** The earlier suggestion was to stop remembering an entry
once it returned an empty hit. Rejected for a robustness reason rather than a
speed one: an empty-capable entry with rare hits, `(?=</body>)` to insert text
before a closing tag being a realistic list entry, would then be searched from
the walk position at every step and scan to its next hit each time. In a large
file with thousands of other replacements that is the quadratic case this whole
unit exists to remove, and the user experiences it as the freeze that is being
fixed. The mechanisms empty hits need (bars, sidestep, `_netDelta`, `none` at
the end of text) are covered by I8, the level sections and the fixed cases.

**Whether remembering is on when this ships: yes, always, without a setting.**
(Superseded on the release question by the status note at the top: the branch is
parked and 6.1 does not carry it. The reasoning below is why there is no switch
if and when it does ship.) An
INI-only switch with default off was built first and removed again. The reason:
a switch is a statement that the bookkeeping is not trusted, and it would not
even have covered the risk it pretended to cover. Of the 138 lines of
`OnePassHits.cpp` only 58 are bookkeeping; the empty-match rules, the CRLF
sidestep and the termination guard (48 lines) run with remembering on and off
alike; that shared part is new code too, it is where the two original one-pass
bugs were fixed and where the CRLF handling lives, and the switch would not
have taken it out of the build.
A unit that is not robust enough to run without a switch is not robust enough
to ship, so the acceptance bar is the analysis (Stages 1, 2, 5), the harness
(Stages 3, 4, 6) and the manual list (Stage 7), not a fallback. The plain walk
stays as the reference the harness compares against.

Completeness of the classification (section K of the engine harness): 68 regex
constructs measured against the engine with 18 different bytes before the match
position; 13 depend on the byte before and all 13 are classified, 44 are proven
independent and may be remembered, the rest are end anchors or syntax this Boost
build rejects. No gaps. This is the one assumption of the design that was a
hand-written list; it is now a measured one, and the test reruns on every change.

## Stage 6: the exotic cases, in the harness (done)

The cases that cannot be tried by hand: constructs and text shapes the earlier
pools do not reach, each run differentially (remembering against the plain walk)
and, for a single entry, against Notepad++'s own `processRange`.

**Section K rewritten from a classification check into a measurement of the five
engine properties the bookkeeping rests on.** 221 regex constructs and 11
literals (conditionals, recursion in every spelling, branch reset, back
references, `(?x)` with comments, `\Q` without `\E`, `\p{...}`, `[[=a=]]`,
possessive and atomic groups, every anchor and lookaround combination, NEL,
U+2028/29, ZWSP, BOM, emoji, NUL, control and invalid bytes), measured over 53
probe texts, at every start position, case sensitive and insensitive:

| | Property | Why the class needs it | Result |
| --- | --- | --- | --- |
| K1 | Which constructs see the byte before a match | a remembered hit must not survive a change in front of it | 40 depend on it, all classified, 0 gaps |
| K2 | Where a plain search hits, the anchored `\G(?:...)` probe hits too | the probe must not hide a real hit | 0 violations |
| K3 | Two bytes of the same context class give the same hit | a class change is the only trigger of the probe | 0 gaps |
| K4 | Nothing from p means nothing from any later position | an entry without a hit is not searched again | 0 violations in well-formed text; F6 found what the violations in malformed text really were |
| K5 | A hit found from p is found again from its own position | `replaceOne` confirms the hit there | 0 violations except the constructs bound to the search start, which F7 handles |
| K6 | An empty search range gives the same answer after a backward search | a run after Find Previous must behave like any other | 0 violations |

New fixed cases (engine harness): C34 to C54 - replacements that end inside a
character, lone lead and trail bytes, U+2028 and NEL around remembered anchors,
never-matching and always-empty patterns, conditionals and recursion, `(?x)`
comments and open `\Q`, `\K` verified against Replace All, NUL in text, find and
replacement, regex combined with whole word, emoji around replacements, a 64 KB
replacement, CR-only and form-feed-only documents, every separator in one text,
500 growing and 500 shrinking replacements in a row, 201 entries on 3000 bytes,
ten identical entries with a match list, and a start at the end of the text
after a Find Previous. New fuzz sections: M1 exotic but valid UTF-8, M2 the same
in an ANSI document, M3 with external edits, M5 and M6 in column and selection
scope, N up to 48 entries on texts up to 3000 bytes, M4 malformed UTF-8.

Result: seven seeds of the full engine harness plus one run with texts up to
3000 characters, and four seeds of the standalone harness with sanitizers: 0
mismatches. M4 (malformed UTF-8, entries that deliberately cut characters apart)
went from 12 differences in 150 cases to 0 with F5 and F6 fixed. 2400
single-entry runs on malformed text: remembering, the plain walk and Notepad++'s
own Replace All agree in all of them, although the engine answered with a match
ending before its start 111 times.

## The closed contract (Stage 6b)

Every earlier round found something because it widened the testing. Widening has
a long tail, so this section does the opposite: it lists **every** assumption the
walk makes, and next to each one how it is discharged. Three ways count as
discharged - **proven** (the code cannot depend on it being false), **measured**
(checked against the engine over a space, not over examples) and **removed** (the
code was changed so the assumption is no longer needed). Anything that is only
argued is a gap, and the table says so.

### About the regex engine

| # | Assumption | Discharged by | Status |
| --- | --- | --- | --- |
| E1 | A hit found by a search from p is found again by a search from the hit's own position. Without it `replaceOne` could not confirm a proposal. | K5 over 232 constructs x 53 texts x every position x both case modes. The only violators are the constructs bound to the search start; for those the walk passes the position it searched from (`verifyFrom`). | measured + removed |
| E2 | "Nothing from p" means nothing from any later position. This is what lets an entry without a hit be left alone, and it is the whole speed argument. | K4 over the same space: 0 violations in well-formed text. The violations in malformed text turned out to be refused answers, which are no longer read as "nothing" (F6). | measured + removed |
| E3 | Only certain constructs look at the byte before a match; the rest may be remembered across a change in front of them. | K1 over the hand-written list (40 dependent, all classified) **and** K8 over 4000 patterns generated from a grammar (1232 dependent, 0 misclassified). The hand-written list is no longer the only evidence. | measured |
| E4 | Bytes of the same context class before a match give the same answer, so a class change is a sufficient trigger. | K3, same space, 0 gaps. | measured |
| E5 | The anchored `\G(?:...)` probe never hides a hit the plain search would find. | K2, 0 violations. And structurally since F3: the probe only decides whether the plain search runs, it never supplies a hit. | measured + proven |
| E6 | The probe window for a whole-word literal (4x the find text plus 4 bytes) is long enough. | K7 over every literal, every position, every flag combination: 0 disagreements with a search over the whole document. | measured |
| E7 | The engine gives the same answer for an empty search range whether or not a backward search ran before. | K6, 0 quirks; half of every fuzz section now runs after a backward search. | measured |
| E8 | An answer of "no match" is an answer. | Removed: `performSingleSearch` marks a refused answer (invalid pattern, exception, a match ending before its start) and the walk never concludes from it (F6). | removed |
| E9 | Two searches see the same characters, so a position means the same thing to both. | Removed for the case where it fails: a walk position or a replacement end inside a character drops everything remembered (F5). Only reachable in malformed UTF-8. | removed |
| E10 | The engine's empty-match rules: never directly after the entry's own match, CRLF stepped as one character, the end of a range is the end of the text. | Sections A, B and P: every pattern in the harness, as a single entry, on every probe text, must equal Notepad++'s own Replace All - 30820 runs, 0 differences. | measured |

### About the document and the panel

| # | Assumption | Discharged by | Status |
| --- | --- | --- | --- |
| P1 | A wrong proposal cannot become a wrong replacement. | Proven: `replaceOne` re-searches with the entry's own flags and refuses unless position and length match exactly. Independent of everything above. | proven |
| P2 | Text at and after the replacement end is untouched, so a remembered hit only moves by delta. | Proven from the edit itself, restricted to full scope; column and selection scope keep nothing across a replacement (F4). | proven |
| P3 | A new context-dependent match can only appear at the replacement end. | Proven given E3/E4/E9: matches further on consist of untouched bytes with an untouched byte in front of them, and the walk never returns to text before its position. | proven |
| P4 | The line count catches every change of line structure, including a CRLF pair split or joined. | Proven: any such change adds or removes a line. The form feed, which is a line separator for the engine but not for Scintilla, is covered by the context class instead (F1). | proven |
| P5 | Nothing edits the document behind the walk. | Proven for the formula engine: `ILuaEngineHost` has no Scintilla access. Any other change is caught by the length check and `reset()`. | proven |
| P6 | The bars that terminate nullable patterns stay meaningful. | Proven: the walk position never decreases, each entry takes at most one empty match per original position, non-empty hits consume text. `reset()` restarts the position mapping together with the bars (F8). | proven |
| P7 | `_searchRefused` is read only for the search that set it. | Proven by inspection: the walk's document view clears it immediately before every search whose answer it reads. | proven |
| P8 | The code page seen by the walk is the one the searches use. | Proven: taken once from the run's `SearchContext`, which cannot change during a run. | proven |

### What is left

Two things, and they are named rather than hidden.

1. **The engine is a third party.** Everything in the first table is measured
   over a large space, not proven. A construct outside both the hand-written
   list and the grammar, behaving unlike all 232 constructs and all 3697 random
   ones, would not be covered. P1 bounds what that could cost: a missed or a
   reordered replacement, visible in the count columns, never a wrong one.
2. **A pre-existing difference outside this work.** Section P found that the
   panel's own `replaceAll` loop counts one extra empty match between `\r` and
   `\n` for a nullable pattern with an empty replacement, where Notepad++'s
   Replace All does not: `ensureForwardProgress` steps by a character, the engine
   steps a CRLF pair as one. The resulting text is identical, only the count in
   the Find column differs, and the one-pass walk agrees with Notepad++. It is
   noted here because the measurement found it, not because one-pass caused it.

## Stage 7: manual checklist in Notepad++ - done, all green

One build. Run the list on a UTF-8 file and repeat on a CRLF file. Where the
expected result is "same as Replace All", run Notepad++'s own Replace All with
the single entry on a copy of the file and compare text and count.

| # | Case | Expected |
| --- | --- | --- |
| 1 | `$` to `;` | one per line plus one at the end of the text, no freeze |
| 2 | `^` to `> ` on a file without a trailing newline | every line including the last |
| 3 | `x*` to `y` | terminates; this was the freeze |
| 4 | `\b` to `\|` on mixed text | same result as Replace All with a single entry |
| 5 | `cat` to `dog` plus `dog` to `cat` | swap, not double replacement |
| 6 | column mode, 3 columns, entries with `\b` and one replacement that removes a delimiter | only selected columns change, delimiters stay consistent |
| 7 | selection mode, two selections, one growing and one shrinking replacement | scope stays on the intended text |
| 8 | match list `1,3,4` with an empty-capable entry | the right hits, counts consistent |
| 9 | formula entry with `skip()` on every second hit | counts correct, walk continues |
| 10 | formula entry with a syntax error | aborts before any edit |
| 11 | debug stop in the middle of a run | run ends, edits made so far stay, no freeze |
| 12 | from cursor without wrap, then with wrap | start position honoured |
| 13 | ANSI file (CP1252) with umlauts, `\b` and whole word | same result as sequential Replace All |
| 14 | undo after a run | one Ctrl+Z restores the whole run |
| 15 | Find and Replace count columns | match the actual number of hits and replacements |
| 16 | large file with a list containing one regex without hits | finishes in well under a second |
| 17 | Preferences > Delimiter with `-` added to the word characters, whole-word entry `foo`, text `x-foo foo` and an entry `-` to space listed above it | both `foo` replaced |
| 18 | a file containing form feeds, entry `^` to `> ` | Replace All and one-pass agree on the count |
| 19 | entry `a\Kb` to `K` on `abc abx ab` | same as Replace All with that entry (this did not replace anything before F7) |
| 20 | one-pass off (the tab's own option): the same list must still run one Replace All per entry from the start of the file | unchanged behaviour |

Exit: all 20 pass on both line endings. **Run, by the owner, in real Notepad++: all 20 numbered cases (A1-J2 in the superseding `ONEPASS-CHECKLIST.md`, which is what was actually executed - one .mrl/.txt pair per case) plus J5 (undo) came back green.** One live catch on the way: C3 (`a\Kb`) first looked like the F7 regression had returned - Find 3, Replace 0 - which turned out to be an unrebuilt DLL older than the source fix; a fresh build produced the correct `aKc aKx aK`, 3/3, on both the on-pass and the off-pass reference run. Not run: J3/J4 (formula skip/error - build-dependent) and multi-document "Replace All in Open Documents", which calls `onePassReplaceAll` once per open document and was never exercised by any harness or the checklist - see Open Items.

## Stage 8: ship

1. README is updated (one-pass note).
2. Changelog: empty matches (`^ $ \b` lookarounds) are replaced like Replace
   All; no freeze when a pattern matches empty at the end of the text; skipped
   hits no longer hide other entries' hits at the same position; Replace All
   in one pass no longer searches every entry at every step.
3. Commit `src/OnePassHits.*`, the two QA files, `ONEPASS-PLAN.md`, the panel
   diff, the project files, README.

The one-pass mode itself stays the option it is: the tab's "Replace All in one
pass" toggle, off by default. What was removed is only the switch for the
bookkeeping inside that mode.

## Verdict

The question was whether a one-pass run with remembering gives the same result
as one without, and whether that can be trusted. The evidence, in the order of
its weight:

1. By construction: `replaceOne` verifies every proposed hit with an
   independent plain search before replacing. A defect in the bookkeeping can
   cost a missed or a reordered replacement, never a replacement of text the
   entry does not match. This does not depend on any test.
2. By construction since F3: every remembered hit is the result of a plain
   search or its shift, so the bookkeeping cannot disagree with the
   verification about what the engine finds.
3. Measured: the properties the design assumes about the engine are no longer
   assumptions. Section K measures eight of them - over 232 hand-picked
   constructs and 53 texts at every position, and over 4000 patterns generated
   from a grammar. The two that carry the whole design, "nothing here stays
   nothing" (K4) and "a hit is found again at its own position" (K5), are
   measured, not argued; so is the classification (K1 and K8, 0 gaps in 3697
   measurable random patterns) and the probe window (K7).
3a. Complete rather than wide: section P runs every pattern the harness knows,
   as a single entry, on every probe text, in UTF-8 and ANSI, against
   Notepad++'s own Replace All - 30820 runs, no exclusions, 0 differences.
   The exclusion that used to sit here is what hid F7 for four rounds.
4. Differential: remembering against the plain walk over 16 dimensions (three
   scopes, outside edits, custom word characters, ANSI, start position, runs
   after a backward search, three entry classes in isolation, exotic UTF-8,
   malformed UTF-8, 48-entry lists, full mix), seven seeds with 600 cases each
   plus one with texts up to 3000 characters on the real engine, 0 mismatches;
   single entries against N++'s own processRange, 0 mismatches; the standalone
   harness the same with sanitizers.
5. Read: fifteen invariants and eight panel seams confirmed line by line, and
   a closed table of every assumption with the way each one is discharged
   ("The closed contract" above). Two entries in that table are marked as
   remaining: the engine is a third party, and one pre-existing difference in
   the panel's own Replace All that this work did not cause and did not change.

What this does not give is a proof. The bookkeeping is stateful invalidation
logic against a third-party engine that surprised the analysis five times (form
feed, the document's word table, `\G` versus plain search, decoding that
depends on the search start, answers that are refusals rather than "nothing
here"). Every defect sat in a narrow corner - form feed in the file, custom
word characters configured, a `\Z`-type pattern before a trailing form feed,
column mode with a delimiter-inserting replacement, malformed UTF-8 - and one
of them (F7, `\K` never replaced) was not the bookkeeping's at all but a
property of the one-pass walk that the analysis only found because the
bookkeeping was being measured against Replace All.

The pattern of the findings is worth stating plainly: they were all found by
asking what the engine actually does, not by running more random cases. Both
rounds of fuzzing were green before each of them.

Position: for a normal list on a normal file the two runs are the same and a
difference would surprise me. For exotic combinations the remaining risk is
bounded by the safety property: a corner the analysis missed costs a missed
replacement, which the Find and Replace count columns make visible, never a
wrong one. After Stage 6 the exotic corners that can be built at all have been
built, including the malformed-encoding case where the engine stops being a
function of the text and the pattern alone.

What changed in Stage 6b is the kind of confidence, not the amount of testing.
Before it, the answer to "are there more bugs" was "the last rounds found some,
so probably". Now every assumption is written down with the way it is
discharged, and the ones that cannot be proven are measured over spaces rather
than examples. A defect from here on has to be an assumption nobody wrote down
- and the safety property still bounds what that costs.

Verdict from the analysis: ship it, remembering on, no switch. The one-pass mode
stays the option it always was.

Decision by the owner: ship. The manual checklist has been run in real
Notepad++ (Stage 7, all 20 cases plus J5/undo green, one live catch along the
way - see the status note at the top), and the owner reviewed the findings in
`ONEPASS-HANDOVER.md` before deciding. It goes into the codebase, not a side
branch. Not run before this decision: J3/J4 (formula skip/error) and
multi-document Replace All in Open Documents; that gap is recorded rather than
closed, and either surface is where a next defect is most likely to be found -
see `ONEPASS-HANDOVER.md` section 9.

## Revert path

If a bug report arrives after shipping:

1. Ask the reporter for the list and the file. Put both into section A of
   `onepass_engine_qa.cpp` as a fixed case; the harness runs it with
   remembering on and off. Same result means the bookkeeping is not the cause.
2. If it is the bookkeeping, the fix goes through the harness first and the
   case stays as its pin. The one-line emergency revert is `false` as the
   fourth constructor argument in `onePassReplaceAll` (rebuild needed); the
   empty-match fixes stay in place either way.
