# One-Pass Replace All: handover to whoever reviews or continues this

Written for another model (or another engineer) picking this up cold. It is the
**reasoning**: what the thing is, why it is built the way it is, every defect
found so far and why the tests had not found it, what is actually proven versus
measured versus argued, and where I would attack it next.

Read alongside, not instead of:

| File | What it is |
| --- | --- |
| `src/OnePassHits.h` / `.cpp` | the unit, 119 + 234 lines |
| `src/tests/ONEPASS-PLAN.md` | the staged work record: invariants I1-I15, panel seams S1-S8, findings F1-F8, the closed contract, the manual checklist |
| `src/tests/onepass_engine_qa.cpp` | QA against the real Scintilla document and the Notepad++ Boost bridge, 1018 lines |
| `src/tests/onepass_qa.cpp` | the same logic against a fake document, builds anywhere, sanitizers, 380 lines |
| `src/MultiReplacePanel.cpp` | `onePassReplaceAll` (the loop), `replaceOne` (the confirmation), `performSingleSearch` (the refusal signal) |

**Status: shipped.** Owner's decision after reviewing the findings: it went
into the codebase (not a side branch). The release gate,
`ONEPASS-CHECKLIST.md` passing in real Notepad++, was met - all 20 numbered
cases plus J5 (undo) green, with one live catch on the way (C3 briefly looked
like the F7 regression had returned; a stale, unrebuilt DLL was the cause, not
the code - see the checklist's own run log via section 9 item 1 below). The
reason it took eight findings to get here is analysed below and is itself the
useful part - see section 11. The feature can still be reduced later
(`remember = false`, one word) or removed (one file plus ~40 panel lines);
section 10 covers that. "Should this ship" keeps both sides of the argument on
record even though the decision is made.

---

## 1. What the feature is

MultiReplace replaces a **list** of find/replace entries. Per tab there is an
option **"Replace All in one pass"** (default off, stored as `OnePass` per tab).

- **Off:** every enabled entry gets its own Replace All run over the whole
  document, one after another. Entry 2 therefore sees the text entry 1 produced.
- **On:** the document is walked **once**. At every position the nearest hit of
  any enabled entry wins (tie: the entry listed first). Replaced text is never
  looked at again, so `cat`->`dog` and `dog`->`cat` in one list swap instead of
  cancelling.

That option is a user feature and is **not** what this work is about. What this
work changed is what happens **inside** the on-case.

### The problem

The old implementation searched **every enabled entry from the walk position at
every step**. With one entry that has no hits in the file, every step scanned the
rest of the document for it. 300 KB of prose, 5 frequent literals plus 5 regex
of which 2 have no hits: **10.5 seconds**. Users experienced it as a freeze.

### The fix

`OnePassHits` remembers each entry's next hit and searches an entry again only
when the last replacement could have changed its answer. Same file, same list:
**11 ms**. Two entries that never hit are searched once, not 7000 times.

```
  10 frequent literals              on     9 ms   /  off    24 ms  /  old   22 ms
  5 literals + 3 rare + 2 no-hit    on    11 ms   /  off 10513 ms  /  old 10453 ms
  1 literal + 1 regex no-hit        on     3 ms   /  off   774 ms  /  old  780 ms
  ^ -> "> " + 2 literals            on     7 ms   /  off     9 ms
  \s+$ -> "" + 3 literals           on     8 ms   /  off 11553 ms  /  old 11533 ms
```

(`onepass_engine_qa bench 300`. "off" is the same class with the bookkeeping
switched off, "old" is the flow before this work. The two agree, as they should.)

Two behaviour bugs were fixed on the way, and they matter for the shipping
decision because they are **not** part of the risky half:

- empty matches (`^`, `$`, `\b`, lookarounds) are now replaced the way Replace
  All does for a single entry; before, one-pass replaced almost none of them
- `x*` at the end of the text no longer loops forever
- a skipped hit no longer hides another entry's hit at the same position

---

## 2. The interface, in five calls

```cpp
OnePassHits hits(doc, entries, scope);          // doc = lambdas onto Scintilla
while (...) {
    OnePassHit h = hits.next(pos, w);           // nearest hit at/after pos, w = its entry
    if (h.pos < 0) break;
    if (replaceThisHit) {
        hits.beforeReplace(h);                  // records the context before the edit
        replaceOne(..., hits.verifyFrom(w, h.pos));   // confirms, then edits
        hits.afterReplace(w, h, delta);         // or hits.afterSkip(w, h)
    }
    else hits.afterSkip(w, h);
}
hits.reset();                                   // the document changed from outside
```

`OnePassDoc` is a plain struct of `std::function`s (`search`, `searchRange`,
`charAt`, `positionAfter`, `lineCount`, `length`, plus a `utf8` flag). That is
the whole coupling: the class never touches Scintilla, which is why the QA can
drive the same class against a fake document.

The constructor takes `bool remember = true`. **This is a test lever, not a
setting.** With `false` every step searches every entry from the walk position -
the reference the differential sections compare against. An INI switch existed
briefly and was removed: a switch is a statement that the code is not trusted,
and it would not even have covered the risk it pretended to, because only 58 of
the 138 statements are bookkeeping and three of the historical defects sat in
the shared half.

---

## 3. The three load-bearing ideas

Understand these and the rest of the code reads itself.

### 3.1 The safety property (`replaceOne`)

Before anything is written, `replaceOne` runs **its own search with the entry's
own flags** and refuses unless the found position **and** length are exactly what
the walk proposed. So:

> A defect in `OnePassHits` can cost a **missed** or a **reordered**
> replacement. It cannot produce a replacement of text the entry does not match.

This holds without any test, and every risk assessment in this work is anchored
to it. **Anything that removes or weakens that search removes the safety net.**
If you refactor, keep it.

### 3.2 `keep()`: the bookkeeping cannot disagree with the confirmation

Finding F3 taught this. Originally the anchored probe `\G(?:pattern)` could put
a hit into the cache; the engine's plain search then refused it, so the walk
counted a hit it could not replace. The fix was structural, not a special case:

> Every remembered hit is the result of a **plain search** (or its shift by
> delta). The anchored probe is only a **gate** that decides whether that plain
> search runs. It never supplies a hit.

Since `replaceOne` also uses a plain search, the two cannot disagree about what
the engine finds. That is `keep()`, and it is the reason the class survived the
later rounds as well as it did.

### 3.3 What is remembered, what never is

Three classes, decided textually from the find text at construction:

| Class | Test | Behaviour |
| --- | --- | --- |
| plain | everything else | hit remembered; after a replacement the position is shifted by delta if the hit lies strictly after the replacement end, otherwise dropped |
| context-sensitive (`ctxStart`) | whole word, or the pattern contains `\b \B \< \> ^ $ \A \Z \` [[:<:]] [[:>:]]` | same, plus: after a replacement whose **context class** at the end changed, an anchored probe at the replacement end decides whether a plain search is re-run there |
| never remembered (`uncached`) | `(?<=`, `(?<!`, `\K`, `\G`, `(?R)`, `(?0)`, an `(?x)` flag group, an unterminated `\Q` | searched from the walk position at every step; nothing is kept |

The `uncached` list has two different reasons in it, which is easy to miss:

- `(?<=`, `(?<!`, `\K`, `\G` reach outside the match or bind to the search start,
  so a remembered position means nothing
- `(?R)`, `(?0)`, `(?x)` with a `#` comment, an open `\Q` would be **broken by
  the probe's own wrapper** `\G(?:...)`: whole-pattern recursion would recurse
  into the wrapper, a free-spacing comment would swallow the closing bracket

The **context class** of the byte before a position is `Bof | Cr | Lf | Word |
Other | NonAscii`, with two deliberate oddities:

- a **form feed** maps to `Lf`, because Boost treats `\f` as a line separator for
  `^` and `$` while Scintilla does not count it as a line (F1)
- **every byte >= 0x80 always counts as a change**, because the engine's word
  classification is ASCII-only and multi-byte characters must never be assumed
  equivalent

For a **whole-word literal** the class compares the raw byte instead, because
Scintilla's whole-word test follows the document's word-character table, which
the user can change in Preferences > Delimiter (F2).

### 3.4 Termination

An entry may not take an empty match at its own `consumedEnd`, nor twice at the
same **original** position (`lastEmptyOrig`, mapped through `_netDelta`). A
barred empty match is stepped over CRLF-aware (`nextChar`, mirroring the bridge's
`nextCharacter`). Since the walk position never decreases, each entry gets at
most one empty event per original position and non-empty hits consume text, the
number of steps is bounded. `reset()` restarts that mapping together with the
bars - that is F8.

---

## 4. Every defect found, and why the tests had not found it

This table is the most useful thing in this document. The **pattern** of past
defects is the best available predictor of where the next one is.

| # | What was wrong | Why the tests missed it | What closed it |
| --- | --- | --- | --- |
| F1 | Boost treats `\f` as a line separator for `^`/`$`; the context classes knew only `\r` and `\n`, so a form feed inserted in front of a remembered `^` entry did not trigger a re-check | no test text contained a form feed | `\f` maps to `Lf`; form feeds added to the fuzz alphabet |
| F2 | Whole-word matching follows the document's word table (`SCI_SETWORDCHARS`); the class used fixed alnum + `_`, so replacing `-` with a space did not free a whole-word hit | the harness never configured custom word characters | whole-word literals compare the raw byte; section H configures a word table |
| F3 | The engine's search heuristic and its per-position matcher disagree: before a final `\f`, plain `\Z` finds nothing but `\G(?:\Z)` matches. The probe put a hit in the cache that the confirmation refused | needed a `\Z`-type pattern **and** a trailing form feed in the same case | structural: the probe only gates, the plain search decides (3.2) |
| F4 | Column scope: a replacement that inserts the delimiter moves the rest of the line into another column, so untouched text can enter or leave the search scope | the delimiter-inserting entry (`a` -> `a,`) was not in the pool | nothing remembered survives a replacement in column or selection scope |
| F5 | A replacement ending **inside** a character makes untouched bytes after it decode differently, so remembered hits and the class of the byte before them are stale. Reachable only in malformed UTF-8 | no test produced malformed UTF-8 by replacing | a trail byte at the replacement end, or a walk position on one, drops everything remembered |
| F6 | The bridge can answer with a match **ending before it starts** (`\b` in malformed UTF-8), and with -2/-3 for an invalid pattern or an internal exception. `performSingleSearch` refuses all three, and the walk read the refusal as "there is nothing here" and marked the entry exhausted | the harness mapped refusals to "not found" exactly like the panel did, so it reproduced the bug instead of exposing it | `performSingleSearch` sets `_searchRefused`; the document view reports length -1; `keep()` never concludes from it |
| F7 | An entry bound to the search start (`\K`, `\G`, lookbehind) is invisible from its own match position, so `replaceOne`'s confirmation refused **every** such hit: one-pass counted `a\Kb` hits and replaced none, while Replace All replaced them. **Pre-existing, not caused by the bookkeeping** | the comparison against Notepad++'s Replace All contained `find("\\K") == npos` - the one defect it existed to find was the one it could not see | `replaceOne` takes the position the confirmation starts from; `OnePassHits::verifyFrom` supplies it. Section P now runs with **no** exclusions |
| F8 | `reset()` cleared the empty-match bars but not `_netDelta`. With a stale offset a real position can hit the -1 sentinel of `lastEmptyOrig`, barring an empty match that should be taken | not reachable by construction from the fuzz shapes; found by reading | the mapping restarts with the bars, so the collision is impossible rather than unlikely |

**Read the "why missed" column as a list of failure modes of the testing, not of
the code.** Four distinct ones appear:

1. the input alphabet did not contain the byte (F1, F5)
2. the environment was never varied (F2)
3. two rare conditions had to coincide (F3, F4)
4. **the test agreed with the bug** - either because it modelled the panel's own
   mistake (F6) or because an exclusion hid it (F7)

Category 4 is the dangerous one, and it is the reason section P exists and why
section 8 below lists every place the harness could still be lying.

---

## 5. What is proven, measured, argued

`ONEPASS-PLAN.md` has the full table ("The closed contract"). The short form:

**Proven** (the code cannot depend on it being false): the safety property; that
text at and after a replacement end is untouched; that a new context-dependent
match can only appear at the replacement end; that a line-structure change always
changes the line count; that the formula engine cannot edit the document; that
the termination bars stay meaningful; that `_searchRefused` is read only for the
search that set it.

**Measured** (against the real engine over a space, not over examples):

| | What | Result |
| --- | --- | --- |
| K1 | which constructs look at the byte before a match, over 221 regex constructs + 11 literals x 53 texts x every position x both case modes | 40 dependent, all classified, 0 gaps |
| K2 | the anchored probe never hides a hit the plain search finds | 0 violations |
| K3 | two bytes of the same context class give the same answer | 0 gaps |
| K4 | "nothing from p" implies nothing from any later position | 0 in well-formed text; the malformed-text violations were F6 |
| K5 | a hit is found again from its own position (what makes the confirmation sound) | 0, except the constructs bound to the search start - which is F7 |
| K6 | an empty search range answers the same after a backward search (`_lastDirection`) | 0 quirks |
| K7 | the 4x+4 probe window for whole-word literals is long enough | 12804 positions, 0 disagreements with a full search |
| K8 | the classification against **4000 patterns generated from a grammar** | 3697 measurable, 1232 dependent, **0 misclassified** |
| P | every pattern the harness knows, single entry, every probe text, UTF-8 and ANSI, against Notepad++'s own Replace All | **30820 runs, no exclusions, 0 differences** |

**Argued only:** nothing in the design, by intent. What remains is that the
engine is a third party - K1/K8 cover a large space of constructs but not all
possible ones. The safety property bounds what an unknown construct could cost.

---

## 6. Build and run

Both harnesses build the **real** `../OnePassHits.cpp`, not a copy.

```sh
# standalone: fake document, no dependencies, sanitizers. Run this first.
cd src/tests
g++ -std=c++20 -Wall -Wextra -fsanitize=address,undefined -I.. \
    -o onepass_qa onepass_qa.cpp ../OnePassHits.cpp
./onepass_qa            # optional: ./onepass_qa <seed>

# engine harness: needs the Notepad++ sources and Boost.Regex
NPP=/path/to/notepad-plus-plus; S=$NPP/scintilla
for f in CellBuffer Document PerLine RunStyles Decoration CaseFolder CaseConvert \
         CharClassify CharacterCategoryMap CharacterType UniConversion DBCS \
         ChangeHistory UndoHistory RESearch; do
  g++ -std=c++17 -O2 -DNDEBUG -DSCI_OWNREGEX -I$S/include -I$S/src -c $S/src/$f.cxx
done
g++ -std=c++17 -O2 -DNDEBUG -DSCI_OWNREGEX -I$S/include -I$S/src -I$NPP/boostregex \
    -c $NPP/boostregex/BoostRegExSearch.cxx $NPP/boostregex/UTF8DocumentIterator.cxx
g++ -std=c++17 -O2 -DNDEBUG -I$S/include -I$S/src -I.. \
    onepass_engine_qa.cpp ../OnePassHits.cpp *.o -lboost_regex -o onepass_engine_qa
```

The platform stubs (`Platform::Assert` and friends) are **inside**
`onepass_engine_qa.cpp`, so do not also link a `PlatStub.o` - duplicate symbols.

```sh
./onepass_engine_qa                     # full, seed 12345, texts up to 120 bytes
./onepass_engine_qa quick               # 150 cases per section instead of 600
./onepass_engine_qa full 4242 3000      # seed, max text length
./onepass_engine_qa bench 300           # the timings above, 300 KB of prose
./onepass_engine_qa trace 'a|b|c' '^' '>' r          # step trace, | = LF, ~ = CR
./onepass_engine_qa find 'a|b' '\bb' 0               # raw engine answer, plain and anchored
TRACE=1 ./onepass_engine_qa full 777    # on a fuzz mismatch, replay it with traces
```

`find` and `trace` are the two tools that made the analytical findings possible.
When something looks wrong, ask the engine directly with `find` before theorising
- F3, F5 and F6 were all identified that way in a few minutes each.

Sections: A (vs the old flow), B (single entry vs Replace All, fixed cases),
C1-C55 (fixed edge cases), **P** (exhaustive single entry vs Replace All),
**K** (the engine-property measurements), then 17 fuzz sections: D full text,
E column, F selection, G external edits, H custom word chars, I ANSI, J start
position, L1-L3 entry classes in isolation, M1 exotic valid UTF-8, M2 the same
in ANSI, M3 the same with external edits, M5/M6 exotic in column/selection,
N scale (48 entries, 3000 bytes), M4 malformed UTF-8 (informational, see 8.4).

Green means: 521 checks pass, and the run ends with `ALL CHECKS PASSED`.
Last full state: seeds 12345, 999, 20260903, 777, 4242, 31337, 5150 plus one run
with 3000-byte texts, and 4 sanitizer runs of the standalone harness.

---

## 7. How the harness models the panel

`HitsPass::go()` in `onepass_engine_qa.cpp` is a **copy of the panel loop**. That
is the point and also the danger: if it drifts from `onePassReplaceAll`, every
green run proves something about a program that does not exist. F6 is exactly
that failure - the harness reproduced the panel's inability to distinguish a
refused search, so both were wrong together and agreed.

What the harness deliberately models:

| Panel | Harness |
| --- | --- |
| `performSearchForward` with the entry's flags + `EMPTYMATCH_ALLOWATSTART` | `Scoping::search` with `newFlags(entry)` |
| `performSingleSearch` refusing `matchEnd < pos`, -2, -3 | `Doc::find` returning -1 and setting `refused` |
| `replaceOne`'s confirming search from `verifyFrom` | the explicit verification search before `doc.replace` |
| column scope = delimiter fields per line | `Scoping::scope == Column` |
| selection scope = stored ranges moved by `adjustSelectionScope` | `Scoping::afterReplace` |
| the length check + `reset()` on an outside edit | `extEdits` |
| formula entries whose replacement varies per hit | `dynamicRepl` |
| a match list (`Replace at matches`) | `matchList` |
| a run started after Find Previous (`_lastDirection`) | `backwardSearch()`, on half the fuzz cases |

**If you change the panel loop, change `HitsPass::go` in the same commit**, or
the harness is measuring the wrong program. This is the single most important
maintenance rule in this work.

---

## 8. Where I would attack it next

Ranked by where I would actually expect to find something, with a concrete method
for each. Items 1 and 2 are the ones I would not skip.

### 8.1 Re-derive the harness/panel correspondence line by line

Not a test - a reading. Put `onePassReplaceAll` and `HitsPass::go` side by side
and check every branch maps. Then do the same for `Doc::find` against
`performSingleSearch` and `Scoping::search` against `performSearchForward` /
`performSearchColumn` / `performSearchSelection`. F6 lived in that gap for four
rounds. Any difference you find is worth more than another 10 fuzz seeds.

### 8.2 Look for more exclusions and more silent agreement

Grep both harnesses for `continue`, `npos`, `if (...) return`, `unsupported`,
`capped`, and for every one ask: *what would this hide?* The `\K` exclusion was
one line and it cost four rounds. Also check every place the harness and the
class share an assumption - a shared assumption is not a test.

### 8.3 The panel seams that the harness cannot reach at all

The harness has no Notepad++, so these are only covered by reading (S1-S8 in the
plan) and by the manual checklist in Stage 7:

- the formula engine returning `skip()`, an error, or a debug stop mid-run
- undo: one Ctrl+Z must restore the whole run (`ScopedUndoAction`)
- the count columns after a run with a match list
- multi-document Replace All in Files
- a document whose code page changes between runs
- rectangular / multiple selections with zero-length sub-selections

### 8.4 Malformed UTF-8, once more

Section M4 is informational by design: in malformed text the engine is not a
function of (text, pattern) - it decodes from wherever the search starts. After
F5 and F6 the differences are **0 of 600**, and 2400 single-entry runs on
malformed text agree with Notepad++ exactly (even though the engine answered with
a match ending before its start 111 times). If a difference reappears there,
first ask whether the engine is being consistent at all - use `find` at the two
positions - before suspecting the bookkeeping.

### 8.5 Things I considered and deliberately did not do

- **Stop remembering an entry once it matched empty.** Rejected: a rare-hit
  empty-capable entry such as `(?=</body>)` would then be scanned from the walk
  position at every step, which is the quadratic case this unit exists to remove.
- **Scan the document for malformed UTF-8 once per run and disable the
  "exhausted" optimisation if it is malformed.** Rejected: that trades a possible
  missed replacement for a possible freeze, which is the worse failure.
- **A `remember` INI switch.** Built, then removed - see section 2.
- **Fixing the panel's own `replaceAll` CRLF counting difference** (below).
  Out of scope; it is not a one-pass question and changing it touches the
  non-one-pass path.

---

## 9. Open items

1. **Stage 7, the manual checklist** (`ONEPASS-CHECKLIST.md`, superseding the
   20-case sketch in `ONEPASS-PLAN.md`) has been run by the owner in real
   Notepad++: all 20 numbered cases plus J5 (undo, run separately as the
   highest-value item below) came back green. C3 (`a\Kb`) briefly looked like
   the F7 regression had returned - Find 3, Replace 0 - which turned out to be
   an unrebuilt DLL older than the source fix; a fresh build gave the correct
   `aKc aKx aK`, 3/3, on both the on-pass run and the off-pass reference run.
   **Not run:** J3/J4 (formula skip/error - depends on the build having that
   feature) and multi-document Replace All in Open Documents
   (`isReplaceAllInDocs`, which calls `onePassReplaceAll` once per open
   document) - no harness and no checklist case ever exercised that surface.
   If a report comes in after shipping, that is the first place to look.
2. **A pre-existing difference, found by section P and left alone.** The panel's
   own `replaceAll` loop counts **one extra empty match between `\r` and `\n`**
   for a nullable pattern with an empty replacement (`x*` -> ``), where Notepad++'s
   Replace All does not: `ensureForwardProgress` steps by a character while the
   engine steps a CRLF pair as one. The resulting **text is identical**, only the
   number in the Find column differs, and the one-pass walk agrees with
   Notepad++. Fixing it means making `ensureForwardProgress` CRLF-aware when
   `SCFIND_REGEXP_SKIPCRLFASONE` is set, and it would change the non-one-pass
   path, so it needs its own decision and its own tests.
3. **The Visual Studio project** already lists `OnePassHits.h/.cpp`; nothing else
   in the build needs changing.
4. The branch also carries stray `.git/*.lock` and `.git/objects/*/tmp_obj_*`
   files from an interrupted git run on the owner's machine - unrelated to the
   code, but they block git until deleted.

---

## 10. Should this ship

Both sides, honestly, because the decision was not mine - it is recorded here
for whoever reviews it next, not as a live question.

**Against (the case that was weighed, and it stayed defensible even after the
decision):** eight defects, the last found by reading after the work was called
done. 234 lines of stateful invalidation against a third-party regex engine
that surprised the analysis five separate times. The benefit is speed, not
correctness, and a plugin with many users does not need to take a correctness
risk for speed.

**For:** the risk is bounded in a way that is unusual. The confirmation search
means a defect costs a missed replacement, never a wrong one, and a missed
replacement is visible in the count columns. Every defect so far sat in a corner
(form feed in the file, custom word characters configured, `\Z` before a trailing
form feed, column mode with a delimiter-inserting replacement, malformed UTF-8),
none in mainstream use. 30820 single-entry runs match Notepad++'s own Replace
All exactly, with no exclusions. And the one surface no harness could reach -
the UI, undo, real Notepad++ - has since been run by hand (Stage 7) and come
back green, including one case (C3) that looked like a regression and turned
out to be a stale build, which is itself evidence the checklist does what it is
for.

**What was actually decided:** ship, remembering on, no switch, into the
codebase rather than a side branch or `remember = false`. The manual checklist
passing was the condition set for that decision, and it was met. What was not
run before shipping - J3/J4 (formula skip/error) and multi-document Replace All
in Open Documents, which calls `onePassReplaceAll` once per open document and
was never exercised by any harness or the checklist - is not a reason to
reverse the decision, but it is the first place to look if a report comes in
from either of those surfaces. See the open items below.

---

## 11. Process notes, for whoever tests this next

The four testing rounds before the last one all **widened** the tests and all
found something. Widening has a long tail: it tells you a bug exists, never that
none remains. What ended the sequence was the opposite move - closing the
contract:

- write down **every** assumption, then discharge each one as proven, measured
  or removed. An assumption that is only argued is a gap, and it will be the
  next defect.
- when an assumption is about a third-party component, **measure it over a
  space**, not over examples. K8 (4000 generated patterns) is worth more than
  any number of hand-picked cases, because it covers what nobody thought of.
- when an assumption cannot be measured, **change the code so it is no longer
  needed**. F5, F6 and F8 were closed that way, not by handling a case.
- prefer a check with **no exclusions** over a broader check with one. Section P
  is 30820 crude runs and it is the strongest evidence in the whole file.
- when a fuzz mismatch appears, do not shrink it by hand: ask the engine with
  `find` what it actually does at the two positions. Every analytical finding
  here came from that, and each took minutes rather than hours.
