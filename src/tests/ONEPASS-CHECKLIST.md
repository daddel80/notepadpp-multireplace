# One-Pass Replace All: manual acceptance checklist

The cases to run by hand in Notepad++ before this goes into the code. Each result
below was verified against the Notepad++ Boost engine in the harness, so the
"Expected" column is what the engine actually does, not a guess.

## How to run each case

1. New file. Set the encoding named in the case (menu **Encoding**): most cases
   are **UTF-8**, a few say **ANSI**.
2. Paste the **Text** block exactly. Where a case needs a special character
   (form feed, no trailing newline) it says so in **Setup** — follow that instead
   of pasting blindly.
3. Open MultiReplace. Clear the list. Enter the **List** rows: Find, Replace,
   and the flags (□ = off, ☑ = on) for **Match case**, **Whole word**, **Regex**.
   Enter the rows in the given order — order decides ties.
4. Turn the tab option **"Replace All in One Pass"** ☑ on (tab context menu, or
   wherever your build exposes it).
5. Click **Replace All**. Compare the document to **Expected text** and the
   **Find** / **Replace** counts in the list columns to **Expected counts**.
6. **The reference run** applies only to cases tagged *(vs Replace All)*. For
   those, repeat on a fresh paste with one-pass **off**; the resulting **text**
   must be identical (counts may differ between on and off — only text is the
   promise). Cases without the tag either differ on/off on purpose (A2, C4, D2)
   or cannot be run off at all (scopes, formula) — do not run a reference for
   those, just check the given Expected result.

A case passes only if its **Expected text** matches, and where a case gives
**Expected counts**, those match too. A wrong count with correct text is still a
fail — it is the visible symptom the safety net leaves behind (a defect in the
bookkeeping shows up as a missed replacement, i.e. a Find/Replace number that
does not fit the text).

Legend in the text blocks: `·` marks a trailing space you must actually type;
everything else is literal. "no trailing newline" means the cursor sits right
after the last character, no empty line below.

---

## A. The everyday cases (these must simply be boring)

### A1 — plain literal list *(vs Replace All)*
Text:
```
the quick brown fox
the lazy dog
```
List: `the`→`THE` □☑□ · then · `dog`→`cat` □☑□
Expected text:
```
THE quick brown fox
THE lazy cat
```
Expected counts: `the` Find 2 / Replace 2, `dog` Find 1 / Replace 1.

### A2 — the swap (the reason one-pass exists) — on/off differ by design
Text:
```
cat dog cat
```
List: `cat`→`dog` ☑□□ · then · `dog`→`cat` ☑□□
Expected text: `dog cat dog`
Why it matters: with one-pass **off** you would get `cat cat cat` (entry 1 makes
everything `dog`, entry 2 turns it all back). On, replaced text is never touched
again, so it is a true swap. Counts: `cat` Find 2 / Replace 2, `dog` Find 1 /
Replace 1.

### A3 — overlap and ties
Text:
```
aaaa
```
List: `aa`→`b` ☑□□ · then · `a`→`x` ☑□□
Expected text: `bb`
Why: at each position the nearest hit wins, tie goes to the entry listed first,
so `aa` wins at 0 and at 2. `a`→`x` never fires. On odd input `aaa` you get `bx`.

### A4 — many entries, few hits (the speed case, checked for correctness)
Text:
```
hello world
```
List: `world`→`W` □☑□ · then · `zzz`→`q` □□□ · then · `qqq`→`r` □□□
Expected text: `hello W`. Counts: `world` 1/1, the other two 0/0.
On a large file this is the case that used to freeze; here just confirm the
result is right and the run is instant.

---

## B. Empty matches (this is where one-pass used to be wrong)

### B1 — `$` at every line end *(vs Replace All)*
Setup: paste the three lines, **no trailing newline** after `gamma`.
Text:
```
alpha
beta
gamma
```
List: `$`→`;` □□☑
Expected text:
```
alpha;
beta;
gamma;
```
Expected counts: Find 3 / Replace 3 (one per line; the last line counts even
without a trailing newline). No freeze.

### B2 — `^` at every line start
Text (no trailing newline):
```
alpha
beta
gamma
```
List: `^`→`> ` □□☑
Expected text:
```
> alpha
> beta
> gamma
```
Counts: 3 / 3, including the very first line.

### B3 — `x*` terminates *(vs Replace All)*
Text:
```
axbxc
```
List: `x*`→`y` □□☑
Expected text: `yaybycy`
This is the old freeze. It must finish instantly and produce exactly that.

### B4 — `\b` word boundaries *(vs Replace All)*
Text:
```
foo bar
```
List: `\b`→`|` □□☑
Expected text: `|foo| |bar|`
Counts: Find 4 / Replace 4.

### B5 — empty line, `^$`
Text (three lines, the middle one empty):
```
alpha

gamma
```
List: `^$`→`#` □□☑
Expected: the empty middle line becomes `#`; `alpha` and `gamma` untouched.

---

## C. Regex that looks at its surroundings (the classified half)

### C1 — lookahead *(vs Replace All)*
Text:
```
bc ba bc
```
List: `b(?=c)`→`X` ☑□☑
Expected text: `Xc ba Xc`

### C2 — lookbehind *(vs Replace All)*
Text:
```
ab cb ab
```
List: `(?<=a)b`→`B` ☑□☑
Expected text: `aB cb aB`

### C3 — `\K` (this was counted-but-never-replaced before the fix) *(vs Replace All)*
Text:
```
abc abx ab
```
List: `a\Kb`→`K` ☑□☑
Expected text: `aKc aKx aK`
Expected counts: Find 3 / Replace 3. If Replace shows 0 while Find shows 3, this
is the old F7 defect — fail.

### C4 — replaced text is behind the walk and is not revisited
Text:
```
ab ab ab
```
List: `a`→`c` ☑□□ · then · `cb`→`!` ☑□□
Expected text: `cb cb cb`
This is the defining behaviour of one-pass, and it is worth seeing on purpose:
`a`→`c` turns each `ab` into `cb`, but the walk has already moved past that spot,
so `cb`→`!` never fires on text the first entry just produced. Counts: `a` 3/3,
`cb` 0/0. With one-pass **off** you would instead get `! ! !`, because there the
second entry runs over the whole document after the first. Both are correct — the
difference is exactly what the option means, and this case makes it visible.

---

## D. Text size changes under the walk

### D1 — big growth then later hits *(vs Replace All)*
Text:
```
a bc bc bc
```
List: `a`→(type 200 x the letter `q`) ☑□□ · then · `bc`→`X` ☑□□
Expected: the `a` becomes 200 q's, and all three `bc` still become `X`. Counts:
`a` 1/1, `bc` 3/3. (The point is that the remembered `bc` hits survive a large
insertion in front of them.)

### D2 — deletions with anchored entries (a deliberate on/off difference)
Text:
```
aXbXc
```
List: `a`→`` (empty replace) ☑□□ · then · `c`→`` ☑□□ · then · `^X`→`1` ☑□☑ · then · `X$`→`2` ☑□☑
Expected text (one-pass **on**): `1bX`
What happens: `a` and `c` are deleted, leaving `XbX`; `^X` fires on the new
leading `X` → `1bX`. The trailing `X` is behind the walk by the time `X$` is
considered, so it is not replaced.
**Do NOT use the reference run here.** One-pass **off** gives `1b2` instead,
because there `X$`→`2` runs over the whole finished text and catches the trailing
`X`. That difference is correct and expected — it is the same "off revisits, on
does not" behaviour as C4. This case is here to confirm the **on** result is
exactly `1bX` and that the run is stable; it is not a vs-Replace-All case.

### D3 — replacement recreates the find text
Text:
```
abbb abbb
```
List: `ab`→`a` ☑□□
Expected text: `abb abb` (each match consumes `ab`→`a`; the walk does not go
back over the new `a`, so you get one replacement per original `ab`, not a
cascade). Counts: 2/2.

---

## E. CRLF (Windows line endings)

Run B1, B2, B3 again on a file saved with **Windows (CR LF)** line endings
(Edit ▸ EOL Conversion ▸ Windows). Text and the visible result must be identical
to the UTF-8/LF runs.

### E1 — the one known count quirk, so it does not alarm you
Text (Windows CR LF), **no trailing newline**:
```
a
b
```
List: `x*`→`` (empty replace) □□☑
Expected text: **unchanged** (`x*` matches empty everywhere, replacing empty with
empty changes nothing).
Expected count: this is the **one place** where MultiReplace and Notepad++'s own
Replace All disagree by design — the built-in Replace All counts one extra empty
match between the CR and the LF. **The text is identical either way.** One-pass
agrees with Notepad++'s Replace All. So: confirm the text is unchanged and do not
worry if the Find number here looks one-off versus a hand count around the line
break. (This is documented in ONEPASS-PLAN.md, open item 2 — it is pre-existing
and not part of this feature.)

---

## F. Form feed (the F1 case)

### F1 — a form feed in the file
Setup: a form feed (0x0C) is hard to type. Easiest way to get one: with a
throwaway entry in **Extended** mode, Find `ab` Replace `a\fb` (the `\f` inserts a
form feed), run it once on the text `ab`, then delete that entry. You now have
`a`FF`b`. If you cannot make a form feed, skip and note it — the harness covers
it (engine C30/C31, section K); this case is only a real-UI confirmation.
Text: `a`FF`b` (one form feed between the two letters).
List: `^`→`> ` □□☑
Expected text: `>a>b`. The engine treats the form feed as a line break, so `^`
matches both at the start of the file and right after the form feed — two `>`.

---

## G. Scopes

### G1 — Column mode
Setup: a 3-column CSV, comma delimiter. Select column mode, mark columns 1 and 3.
Text:
```
a,b,a
a,b,a
```
List: `a`→`X` ☑□□ · then · `b`→`Y` ☑□□
Expected: only columns 1 and 3 change; the `b` in column 2 is out of scope and
stays. Result:
```
X,b,X
X,b,X
```
(`b`→`Y` finds nothing in the selected columns → 0 replacements.)

### G2 — Column mode, a replacement that changes a field's length
Text (3 columns):
```
a,b,c
a,b,c
```
Select columns 1 and 3. List: `a`→`AA` ☑□□
Expected:
```
AA,b,c
AA,b,c
```
Only column 1 changes; the commas stay where they are and column 3 is untouched.
The point is that the run keeps the columns straight after a field grows. Confirm
every row still has exactly two commas in the same places.

### G3 — Selection mode
Setup: select two separate regions (Ctrl-drag a second selection), e.g. the word
`aa` in two places. Text:
```
aa bb aa bb
```
Select the two `aa`. List: `aa`→`a` ☑□□ · then · `bb`→`Z` ☑□□
Expected: only the selected `aa` regions change to `a`; the `bb` outside the
selection stays. `bb`→`Z` replaces nothing (out of scope).

---

## H. Custom word characters (the F2 case)

### H1 — a hyphen made into a word character *(vs Replace All)*
Setup: Settings ▸ Preferences ▸ Delimiter, add `-` to the word-character set (so
`foo-bar` is one word). Encoding UTF-8.
Text:
```
x-foo foo
```
List: `-`→`·`(a single space) □□□ · then · `foo`→`F` □☑□ (Whole word ON)
Expected text: `x foo foo` → then both `foo` become `F`: `x F F`.
Why: replacing `-` with a space breaks `x-foo` into two words, so the first `foo`
becomes a whole word and gets replaced. Before the F2 fix the second entry would
have missed it. Counts: `foo` Find 2 / Replace 2.
Afterwards: **restore your delimiter setting.**

---

## I. ANSI file

### I1 — umlauts, whole word, `\b` on a CP1252 file *(vs Replace All)*
Setup: Encoding ▸ Character sets ▸ Western European ▸ ANSI (or ISO-8859-1). Type
the umlauts directly so they are single CP1252 bytes, not UTF-8.
Text:
```
für die tür
```
List: `\bdie\b`→`DIE` ☑□☑ · then · `ü`→`ue` □□□
Expected text: `fuer DIE tuer`
Why: the whole-word `die` is replaced (the `\b` boundaries hold around it), and
every `ü` becomes `ue`. Counts: `\bdie\b` 1/1, `ü` 2/2. Run it once with one-pass
off and the text must be identical — that is the ANSI equivalent of section I in
the harness.

---

## J. The engine and the UI around the run

### J1 — start position: from cursor, no wrap
Setup: turn **wrap-around off**, put the cursor in the middle of the text.
Text:
```
foo foo foo
```
Cursor right before the second `foo`. List: `foo`→`X` □□□
Expected: only the second and third `foo` change → `foo X X`. The first, before
the cursor, is untouched.

### J2 — match list (Replace at matches)
Text:
```
a a a a a a
```
List: `a`→`X` ☑□□, and in **Replace at matches** enter `1,3,4`.
Expected: the 1st, 3rd and 4th `a` become `X`, the rest stay → `X a X X a a`.
Counts: Find 6 / Replace 3.

### J3 — a formula entry that skips *(if your build has the formula/Lua feature)*
A formula entry whose script calls `skip()` on every second hit. Expected: the
walk continues past skipped hits, the counts stay consistent, no freeze. (If you
do not use the formula feature, skip.)

### J4 — formula syntax error aborts cleanly *(formula builds only)*
A formula entry with a deliberate syntax error. Expected: the run aborts **before
any edit** — the document is unchanged, an error is shown.

### J5 — undo
After any successful multi-replacement run, press **Ctrl+Z once**. Expected: the
entire run is undone in one step, the document is back to the original.

### J6 — counts sanity
After any run, the **Find** and **Replace** columns per entry must match what you
can see happened. This is the check that would surface a bookkeeping defect,
since a wrong proposal costs a missed replacement, which shows here as a Find/
Replace number that does not match the visible text.

---

## What "all green" means here

Every case's text matches, every count matches, and every *(vs Replace All)* case
gives the identical text to the slow one-at-a-time run. If any case fails, note
its letter+number and the exact Text/List/Expected/Actual — that maps straight
onto a fixed case in `onepass_engine_qa.cpp` so it can be reproduced without the
UI.
