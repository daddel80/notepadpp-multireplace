# Use Variables — Engine Reference

The **Use Variables** option enables dynamic replacements that go beyond plain text substitution. Instead of a fixed replacement string, you write a small expression that is evaluated for each match — accessing the matched text, capture groups, counters, line and file information, and producing a computed result.

Two engines are available, selected via the `(L)` / `(E)` indicator next to the option in the panel. See the [Engine Overview](README.md#engine-overview) in the main README for a side-by-side comparison and decision guidance. This document is the syntax reference for both engines.

## Table of Contents
- [Lua Reference](#lua-reference)
  - [Quick Start: Use Variables](#quick-start-use-variables)
  - [Variables Overview](#variables-overview)
  - [Available Functions](#available-functions)
  - [Command Reference](#command-reference)
  - [String Formatting Helpers](#string-formatting-helpers)
  - [Preload Variables and Helpers](#preload-variables-and-helpers)
  - [Operators](#operators)
  - [If-Then Logic](#if-then-logic)
  - [Debug Mode](#debug-mode)
  - [More Examples](#more-examples)
- [ExprTk Reference](#exprtk-reference)
  - [Quick Start: ExprTk](#quick-start-exprtk)
  - [Pattern Syntax](#pattern-syntax---)
  - [Variables and Captures](#variables-and-captures)
  - [Math Functions](#math-functions)
  - [Control Flow](#control-flow)
  - [String Output via `return [...]`](#string-output-via-return-)
  - [Operators](#operators-1)
  - [Skipping Matches](#skipping-matches)
  - [Sequence Generator](#sequence-generator)
  - [CSV Column Access](#csv-column-access)
  - [More Examples](#more-examples-1)
  - [Limitations](#limitations)

<br>

## Lua Reference

---

### Quick Start: Use Variables

1. **Enable "Use Variables"** — Check the "**Use Variables**" option in the Replace interface.

2. **Use a Command:**

   **Option 1:** [`set(...)`](#setstrorcalc) → Directly replaces with a value or calculation.
   - **Find**: `(\d+)`
   - **Replace**: `set(tonum(CAP1) * 2)`
   - **Result**: Doubles numbers (e.g., `50` → `100`).

   *(Enable 'Regular Expression' in 'Search Mode' to use `(\d+)` as a capture group.)*

   **Option 2:** [`cond(...)`](#condcondition-trueval-falseval) → Conditional replacement.
   - **Find**: `word`
   - **Replace**: `cond(CNT==1, "FirstWord")`
   - **Result**: Changes only the **first** occurrence of "word" to "FirstWord".

3. **Use Basic Variables:**
   - **`CNT`**: Inserts the current match number (e.g., "1" for the first match, "2" for the second).
   - **`CAP1`**, **`CAP2`**, etc.: Holds captured groups when Regex is enabled.
     > **Capture Groups:** With a regex in parentheses `(...)`, matched text is stored in `CAP` variables (e.g., `(\d+)` in `Item 123` stores `123` in `CAP1`). For more details, refer to regex documentation.

   See the [Variables Overview](#variables-overview) for a complete list.

---

### Variables Overview

All variables below are registered under both upper- and lowercase. Pick whichever style reads better — `CNT` and `cnt` refer to the same value, as do `match` and `MATCH`, `cap1` and `CAP1`, and so on.

| Variable        | Description |
|-----------------|-------------|
| `CNT` / `cnt`   | Count of the detected string. |
| `LCNT` / `lcnt` | Count of the detected string within the line. |
| `LINE` / `line` | Line number where the string is found. |
| `LPOS` / `lpos` | Relative character position within the line. |
| `APOS` / `apos` | Absolute character position in the document. |
| `COL` / `col`   | Column number where the string was found (CSV-Scope option selected). |
| `MATCH` / `match` | Contains the text of the detected string, in contrast to `CAP` variables which correspond to capture groups in regex patterns. |
| `FNAME` / `fname` | Filename or window title for new, unsaved files. |
| `FPATH` / `fpath` | Full path including the filename, or empty for new, unsaved files. |
| `CAP1` / `cap1`, `CAP2` / `cap2`, ... | Equivalents to regex capture groups, designed for use in the 'Use Variables' environment. Always strings; use `tonum(CAP1)` for calculations. Note: Standard counterparts (`$1`, `$2`, ...) cannot be used in this environment. |

**Notes:**
- `FNAME` and `FPATH` are updated for each file processed by `Replace All in Open Documents` and `Replace All in Files`. This ensures that variables always refer to the file currently being modified.
- **String Variables:** `MATCH` and `CAP` variables are always strings. For calculations, use `tonum(CAP1)`. Both dot (.) and comma (,) are recognized as decimal separators. Thousands separators are not supported.

<br>

### Available Functions

| Function | Description | Example |
|----------|-------------|---------|
| `set(v)` | Returns `v` as the replacement string. If `v` is `nil`, the replacement is skipped (original text remains). | `set("Value: " .. CNT)` |
| `cond(c, t, f)` | If `c` is true, returns `t`. If `c` is false, returns `f`. If `f` is omitted, skips replacement on false. | `cond(CNT > 5, "Over 5")` |
| `fmtN(n, d, f)` | Formats number `n` with `d` decimal places and optional fixed flag `f`. | `fmtN(math.pi, 2)` → `"3.14"` |
| `trim(s)` | Removes leading and trailing whitespace from string `s`. | `trim("  abc  ")` → `"abc"` |
| `tonum(s)` | Converts string `s` to number. Accepts both dot (.) and comma (,) as decimal separator. | `tonum(CAP1) * 2` |
| `padL(s, w, c)` | Pads string `s` on the **left** to width `w` with char `c`. Ideal for zero-padding IDs. | `padL(CNT, 3, "0")` → `"001"` |
| `padR(s, w, c)` | Pads string `s` on the **right** to width `w` with char `c`. Ideal for text alignment. | `padR(MATCH, 10, ".")` → `"Val......."` |
| `vars(t)` | Declares custom variables for use across replacements. Essential for carrying state between matches. | `vars({sum=0})` |
| `lkp(k, p, c)` | Looks up key `k` in file `p` and returns the value (or specific column `c`). | `set(lkp(MATCH, "C:\\list.lkp"))` |
| `lvars(p)` | Loads all key-value pairs from file `p` as global variables. | `lvars("C:\\config.vars")` |
| `lcmd(p)` | Loads external Lua functions from file `p` to extend functionality. | `lcmd("C:\\helpers.lcmd")` |

<br>

### Command Reference

#### String Composition
`..` is employed for concatenation.
E.g., `"Detected "..CNT.." times."`

<br>

#### set(strOrCalc)
Directly outputs strings or numbers, replacing the matched text with a specified or calculated value.

| Find      | Replace with                                 | Regex | Description/Expected Output                                                 |
|-----------|----------------------------------------------|-------|-----------------------------------------------------------------------------|
| `apple`   | `set("banana")`                              | No    | Replaces every occurrence of `apple` with the string `"banana"`.           |
| `(\d+)`   | `set(tonum(CAP1) * 2)`                    | Yes   | Doubles any found number; e.g., `10` becomes `20`.                         |
| `found`   | `set("Found #"..CNT.." at position "..APOS)` | No    | Shows how many times `found` was detected and its absolute position.       |

<br>

#### cond(condition, trueVal, [falseVal])
Evaluates the condition and outputs `trueVal` if the condition is true, otherwise `falseVal`. If `falseVal` is omitted, the original text remains unchanged when the condition is false.

| Find        | Replace with                                                                     | Regex | Description/Expected Output                                                                                             |
|-------------|----------------------------------------------------------------------------------|-------|-------------------------------------------------------------------------------------------------------------------------|
| `word`      | `cond(CNT==1, "First 'word'", "Another 'word'")`                                 | No    | For the first occurrence of `word` → `"First 'word'"`; for subsequent matches → `"Another 'word'"`.                    |
| `(\d+)`     | `cond(tonum(CAP1)>100, "Large", cond(tonum(CAP1)>50, "Medium", "Small"))`  | Yes   | If > 100 → `"Large"`, if > 50 → `"Medium"`, otherwise → `"Small"`.                                                      |
| `anymatch`  | `cond(APOS<50, "Early in document", "Later in document")`                        | No    | If the absolute position `APOS` is under 50 → `"Early in document"`, otherwise → `"Later in document"`.                |

<br>

#### vars({Variable1=Value1, Variable2=Value2, ...})
**Note:** This command was previously named `init()` and has been renamed to `vars()`. For compatibility, `init()` still works.

Initializes custom variables for use in various commands, extending beyond standard variables like `CNT`, `MATCH`, `CAP1`. These variables can carry the status of previous find-and-replace operations to subsequent ones.

Custom variables persist from match to match within a single **'Replace All'** operation and can transfer values between list entries. Each new operation (**button click**) starts with a fresh state. In multi-document or multi-file replacements, variables **persist across documents**. Use `FPATH` or `FNAME` to detect document changes and reset variables conditionally if needed.

**Init usage:** Can be used as an init entry (empty Find) to preload before replacements; not mandatory. See [Preload Variables and Helpers](#preload-variables-and-helpers) for workflow and examples.

| **Find**         | **Replace**                                                                                                                            | **Before**                                           | **After**                                                    | **Regex** | **Scope CSV** | **Description**                                                                                                              |
|------------------|----------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------|--------------------------------------------------------------|-----------|---------------|------------------------------------------------------------------------------------------------------------------------------|
| `(\d+)`          | `vars({COL2=0,COL4=0}); cond(LCNT==4, COL2+COL4);`<br>`if COL==2 then COL2=tonum(CAP1) end;`<br>`if COL==4 then COL4=tonum(CAP1) end;`               | `1,20,text,2,0`<br>`2,30,text,3,0`<br>`3,40,text,4,0` | `1,20,text,2,22.0`<br>`2,30,text,3,33.0`<br>`3,40,text,4,44.0` | Yes       | Yes           | Tracks values from columns 2 and 4, sums them, and updates the result for the 4th match in the current line.                |
| `\d{2}-[A-Z]{3}` | `vars({MATCH_PREV=''}); cond(LCNT==1,'Moved', MATCH_PREV); MATCH_PREV=MATCH;`                                                          | `12-POV,00-PLC`<br>`65-SUB,00-PLC`<br>`43-VOL,00-PLC` | `Moved,12-POV`<br>`Moved,65-SUB`<br>`Moved,43-VOL`            | Yes       | No            | Uses `MATCH_PREV` to track the first match in the line and shift it to the 2nd (`LCNT`) match during replacements.           |

<br>

#### lvars(filePath)

Loads custom variables from an external file. The file specifies variable names and their corresponding values. The loaded variables can then be used throughout the Replace process, similar to how variables defined with [`vars`](#varsvariable1value1-variable2value2-) work.

The parameter **filePath** must specify a valid path to a file. Supported path formats include:
- Escaped Backslashes: `"C:\\path\\to\\file.vars"`
- Forward Slashes: `"C:/path/to/file.vars"`
- Long Bracket String: `[[C:\path\to\file.vars]]`

**File:**
```lua
-- Local variables remain private
local PATH = [[C:\Data\Projects\]]

-- Only the returned variables are accessible in Replace operations
return {
    userName = "Alice",
    threshold = 10,
    enableFeature = true,
    fullPath = PATH .. "dataFile.lkp"
}
```

**Init usage:** Can be used as an init entry (empty Find) to preload before replacements; not mandatory. See [Preload Variables and Helpers](#preload-variables-and-helpers) for workflow and examples.

| Find          | Replace                                          | Regex | Scope CSV | Description                                                                     |
|---------------|--------------------------------------------------|-------|-----------|---------------------------------------------------------------------------------|
| *(empty)*     | `lvars([[C:\tmp\myVars.vars]])`                  | No    | No        | Loads variables such as `userName = "Alice"` and `threshold = 10`.             |
| `Hello`       | `set(userName)`                                  | No    | No        | Replaces `Hello` with the value of the variable `userName`, e.g., `"Alice"`.   |
| `(\d+)`       | `cond(threshold > 5, "Above", "Below")`          | Yes   | No        | Replaces the match based on the condition evaluated using the variable `threshold`. |

**Key Points:**
- **Conditional Loading**: Variables can be loaded conditionally by placing `lvars()` alongside a specific Find pattern. In this case, the variables are only initialized when the pattern matches.
- **Local vs. Returned Variables**: Only variables explicitly included in the return table of the .vars file are available for use. Any local variables remain private to the file.

<br>

#### lkp(key, hpath, inner)
Performs an external lookup of **key** against an indexed data file located at **hpath** and returns the corresponding value. By default, if the **key** is not found, `lkp()` simply reverts to the key itself. Setting **inner** to `true` instead yields a `nil` result when the key is missing, allowing for conditional checks or deeper nested logic.

**Key and File Path:**
- **Key** — The **key** can be either a string or a number. Numbers are automatically converted to strings to ensure compatibility in the lookup process.
- **File Path (hpath)** — The **hpath** must point to a valid `.lkp` file that returns a table of data. Supported path formats: Escaped Backslashes, Forward Slashes, or Long Bracket String.

**Data File Format:**

Each lkp file must be defined as a table of entries in the form `{ [keys], value }`, where `[keys]` can be:
- A single key (e.g., `"001"`).
- An array of keys (e.g., `{ "001", "1", 1 }`) mapping to the same value.

**Example:**
```lua
return {
    { {"001", "1", 1}, "One" },
    { 2, "Two" },
    { "003", "Three" }
}
```

In this example:
- `'001'`, `'1'`, and `1` all correspond to `"One"`.
- `2` corresponds to `"Two"`.
- `'003'` directly maps to `"Three"`.

**Caching Mechanism:**
Once `lkp()` loads the data file for **hpath**, the parsed table is cached in memory for the duration of the Replace-All operation.

**inner Flag:**
- **`false` (default, can be omitted)**: If the key is not found, `lkp()` returns the **search term itself** (e.g., `MATCH`, `CAP1`), instead of a mapped value.
- **`true`**: If the key is not found, `lkp()` returns `nil`, allowing conditional handling.

**Examples:**

| **Find**   | **Replace**                                                    | **Regex** | **Scope CSV** | **Description**                                                                                      |
|------------|----------------------------------------------------------------|-----------|---------------|------------------------------------------------------------------------------------------------------|
| `\b\w+\b`  | `lkp(MATCH, [[C:\tmp\hash.lkp]], true)`                        | Yes       | No            | Uses **inner = true**: If found, replaces with mapped value. If not found, original word is removed. |
| `(\d+)`    | `lkp(CAP1, "C:/path/to/myLookupFile.lkp")`                     | Yes       | No            | Uses **inner = false** (default): If found, replaces with mapped value. If not, returns `CAP1`.      |
| `\b\w+\b`  | `output = lkp(MATCH, [[C:\tmp\hash.lkp]], true); set(output or "NoKey")` | Yes | No         | Uses **inner = true**: If lookup result is non-`nil`, replaces with mapped value; otherwise `"NoKey"`. |
| `\b\w+\b`  | `cond(COL==3, lkp(MATCH, [[C:/tmp/col3_hash.lkp]]))`           | No        | Yes           | Looks up values in the third column (`COL==3`) using a separate lookup file.                         |

<br>

#### fmtN(num, maxDecimals, fixedDecimals)
Formats numbers based on precision (maxDecimals) and whether the number of decimals is fixed (fixedDecimals being true or false).

**Note**: The `fmtN()` command can exclusively be used within the `set()` and `cond()` commands.

| Example                             | Result  |
|-------------------------------------|---------|
| `set(fmtN(5.73652, 2, true))`       | 5.74    |
| `set(fmtN(5.0, 2, true))`           | 5.00    |
| `set(fmtN(5.73652, 4, false))`      | 5.7365  |
| `set(fmtN(5.0, 4, false))`          | 5       |

<br>

#### lcmd(path)

Load user-defined helper functions from a Lua file. The file must `return` a table of functions. `lcmd` registers those functions as globals for the current run.

**Purpose:** Add reusable helper functions (formatters, slugifiers, padding, small logic). Helpers **must return a string or number** and are intended to be called from **action** commands (e.g. `set(...)`, `cond(...)`).

**Init usage:** Can be used as an init entry (empty Find) to preload before replacements; not mandatory. See [Preload Variables and Helpers](#preload-variables-and-helpers) for workflow and examples.

| Find      | Replace                                                        | Regex | Description |
|-----------|----------------------------------------------------------------|-------|-------------|
| *(empty)* | `lcmd([[C:\tmp\mycmds.lcmd]])`                                 | No    | Load helpers from file (init row — no replacement). |
| `(\d+)`   | `set(padLeft(CAP1, 6, '0'))`                                   | Yes   | Zero-pad captured number to width 6 using `padLeft`. |
| `(.+)`    | `set(slug(CAP1))`                                              | Yes   | Create a URL-safe slug from the whole line using `slug`. |
| `\{\{(.*?)\}\}` | `set(file_log_simple(MATCH, [[C:\tmp\out.txt]]))`        | Yes   | Logs the entire search hit to a custom file, leaving the original text unchanged. |

**File format:**
```lua
-- C:\tmp\helpers.lcmd
return {
  -- slug: create a URL-friendly slug
  -- Usage: set(slug("Hello World!"))  → "hello-world"
  slug = function(s)
    s = tostring(s or ""):lower()
    s = s:gsub("%s+", "-"):gsub("[^%w%-]", "")
    return s
  end,

  -- titleCase: convert snake_case or space-separated to Title Case
  -- Usage: set(titleCase("hello_world"))  → "Hello World"
  titleCase = function(s)
    s = tostring(s or "")
    s = s:gsub("_", " ")
    s = s:gsub("(%a)([%w]*)", function(first, rest)
      return first:upper() .. rest:lower()
    end)
    return s
  end,

  -- wrap: wrap text at specified width
  -- Usage: set(wrap("long text here", 40))  → wrapped text
  wrap = function(s, width)
    s = tostring(s or "")
    width = tonum(width) or 80
    local result = {}
    for line in s:gmatch("[^\n]+") do
      while #line > width do
        local pos = line:sub(1, width):match(".*()%s") or width
        table.insert(result, line:sub(1, pos - 1))
        line = line:sub(pos + 1)
      end
      table.insert(result, line)
    end
    return table.concat(result, "\n")
  end,

  -- file_log: append match to file, return original match
  -- Usage: set(file_log(MATCH, [[C:\tmp\out.txt]]))
  file_log = function(match, path)
    if match == nil then return "" end
    path = path or [[C:\tmp\matches.txt]]
    local f = io.open(path, "a")
    if not f then return match end
    f:write(tostring(match) .. "\n")
    f:close()
    return match
  end,
}
```

<br>

### String Formatting Helpers

The following built-in helper functions are available for string manipulation:

| Function | Description | Example |
|----------|-------------|---------|
| `trim(s)` | Removes leading and trailing whitespace from string `s`. Useful for cleaning up database dumps or messy CSVs. | `trim("  abc  ")` → `"abc"` |
| `tonum(s)` | Converts string `s` to number. Accepts both dot (.) and comma (,) as decimal separator. Returns `nil` if conversion fails. | `tonum("3,14")` → `3.14` |
| `padL(s, w, c)` | Pads string `s` on the **left** to width `w` with character `c` (default: space). Use `padL(CNT, 3, "0")` to generate sorted file names like `file_001.txt`. | `padL("42", 5, "0")` → `"00042"` |
| `padR(s, w, c)` | Pads string `s` on the **right** to width `w` with character `c` (default: space). Useful to align text into fixed-width columns. | `padR("Val", 10, ".")` → `"Val......."` |

#### `tonum` parsing behaviour

The `tonum()` helper normalises commas to periods and delegates to Lua's built-in `tonumber()`. The strict return semantics (e.g. `nil` on trailing junk) follow Lua convention — combine with `or` for a default: `tonum(CAP1) or 0`.

| Input          | Result    | Note                                            |
|----------------|-----------|-------------------------------------------------|
| `"42"`         | `42`      |                                                 |
| `"3.14"`       | `3.14`    |                                                 |
| `"3,14"`       | `3.14`    | comma accepted as decimal                       |
| `"-7,5"`       | `-7.5`    |                                                 |
| `""`, `nil`, `"abc"`     | `nil` | not a number                                    |
| `"$100"`       | `nil`     | leading non-digit fails                         |
| `"3,14abc"`    | `nil`     | trailing characters cause failure               |
| `"  42  "`     | `42`      | surrounding whitespace tolerated                |
| `"0x1F"`, `"1e3"` | `31`, `1000` | hex and scientific notation supported   |

<br>

### Preload Variables and Helpers
Use init entries (empty Find) to preload variables or helper functions before any replacements run. Init entries run once per Replace-All (or per list pass) and do not change text directly.

**How it works:**
- **Place `vars()`, `lvars()` or `lcmd()` next to an empty Find field.**
- This entry does **not** search for matches but runs before replacements begin.
- It ensures that **variables and helpers are loaded once**, regardless of their position in the list.
- Use **Use Variables = ON** for init rows so loaded variables/helpers are available to later rows.

**Examples:**

| **Find**    | **Replace**                                 | **Description** |
|-------------|---------------------------------------------|-----------------|
| *(empty)*   | `vars({prefix = "ID_"})`                    | Set `prefix` before replacements. |
| *(empty)*   | `lvars([[C:\path\to\myVars.vars]])`         | Load variables from file (file must `return { ... }`). |
| *(empty)*   | `lcmd([[C:\tmp\mycmds.lcmd]])`              | Load helpers from file (e.g. `padLeft`, `slug`). |
| `(\d+)`     | `set(prefix .. CAP1)`                       | Uses `prefix` from init (`123` → `ID_123`). |
| `(\d+)`     | `set(padLeft(CAP1, 6, '0'))`                | Use helper loaded by `lcmd` to zero-pad (`123` → `000123`). |
| `(.+)`      | `set(slug(CAP1))`                           | Use helper loaded by `lcmd` to create a slug (`Hello World!` → `hello-world`). |

<br>

### Operators

| Type           | Operators                                | Example                                                |
|----------------|------------------------------------------|--------------------------------------------------------|
| Concatenation  | `..`                                     | `set("Found "..CNT)`                                   |
| Arithmetic     | `+`, `-`, `*`, `/`, `^`, `%`             | `set(CNT * 2)`                                         |
| Relational     | `==`, `~=`, `<`, `>`, `<=`, `>=`         | `cond(LINE == 1, "First", "Not first")`               |
| Logical        | `and`, `or`, `not`                       | `cond(LINE > 5 and CNT < 10, "Midrange", "Other")`    |

<br>

### If-Then Logic
If-then logic is integral for dynamic replacements, allowing users to set custom variables based on specific conditions. This enhances the versatility of find-and-replace operations.

**Note**: Do not embed `cond()`, `set()`, or `vars()` within if statements; `if statements` are exclusively for adjusting custom variables.

**Syntax Combinations:**
- `if condition then ... end`
- `if condition then ... else ... end`
- `if condition then ... elseif another_condition then ... end`
- `if condition then ... elseif another_condition then ... else ... end`

**Example:**

This example shows how to use `if` statements with `cond()` to manage variables based on conditions:

`vars({MVAR=""}); if CAP2~=nil then MVAR=MVAR..CAP2 end; cond(string.sub(CAP1,1,1)~="#", MVAR); if CAP2~=nil then MVAR=string.sub(CAP1,4,-1) end`

<br>

### Debug Mode

The **Debug Mode** provides a safe environment to test complex logic involving conditions or math. Instead of modifying the text immediately, it allows you to step through matches one by one and inspect the real-time values of all variables.

**How to enable:**
* **Via Menu:** Click the arrow on the **Replace All** button and select **Debug Mode**. This is the standard way to debug.
* **Via Script:** Initialize the `DEBUG` variable in your replacement string. This overrides the menu setting and allows for conditional debugging.

| Find | Replace | Description |
| :--- | :--- | :--- |
| *(empty)* | `vars({DEBUG=true})` | **Globally** enables Debug Mode for the run (acts as an [Init entry](#preload-variables-and-helpers)). |
| `(\d+)` | `if CNT==50 then DEBUG=true end; set(CAP1)` | Activates Debug Mode **starting from** the 50th match. |

<br>

### More Examples

| Find              | Replace                                                                                                     | Regex | Scope CSV | Description                                                                                     |
|-------------------|-------------------------------------------------------------------------------------------------------------|-------|-----------|-------------------------------------------------------------------------------------------------|
| `;`               | `cond(LCNT==5,";Column5;")`                                                                                 | No    | No        | Adds a 5th Column for each line into a `;` delimited file.                                      |
| `key`             | `set("key"..CNT)`                                                                                           | No    | No        | Enumerates key values by appending the count of detected strings. E.g., key1, key2, key3, etc.  |
| `(\d+)`           | `local n=tonum(CAP1); set(CAP1.."€ The VAT is: "..(n*0.15).."€ Total: "..(n*1.15).."€")`  | Yes   | No        | Finds a number and calculates VAT at 15%, displaying original, VAT, and total.                  |
| `---`             | `cond(COL==1 and LINE<3, "0-2", cond(COL==2 and LINE>2 and LINE<5, "3-4", cond(COL==3 and LINE>=5, "5-9")))` | No    | Yes       | Replaces `---` with a specific range based on `COL` and `LINE` values.                          |
| `(\d+)\.(\d+)\.(\d+)` | `local c1,c2,c3=tonum(CAP1),tonum(CAP2),tonum(CAP3); cond(c1>0 and c2==0 and c3==0, MATCH, cond(c2>0 and c3==0, " "..MATCH, "  "..MATCH))` | Yes   | No        | Alters spacing based on version number hierarchy, aligning lower hierarchies with spaces.        |
| `(\d+)`           | `set(tonum(CAP1) * 2)`                                                                                   | Yes   | No        | Doubles the matched number. E.g., `100` becomes `200`.                                          |
| `;`               | `cond(LCNT == 1, string.rep(" ", 20-(LPOS))..";")`                                                          | No    | No        | Inserts spaces before the semicolon to align it to the 20th character position.                 |
| `-`               | `cond(LINE == math.floor(10.5 + 6.25*math.sin((2*math.pi*LPOS)/50)), "*", " ")`                              | No    | No        | Draws a sine wave across a canvas of '-' characters.                                            |
| `^(.*)$`          | `vars({MATCH_PREV=1}); cond(MATCH == MATCH_PREV, ''); MATCH_PREV=MATCH;`                                    | Yes   | No        | Removes duplicate lines, keeping the first occurrence of each line.                             |

<br>

## ExprTk Reference

ExprTk is a numeric expression engine. Its design goal is concise inline math: where Lua wraps every formula in `set(...)`, ExprTk drops the formula directly into the replace string via `(?=...)` markers. This makes it the better choice when the work is mostly arithmetic, statistics, or value transformation on captured numbers.

---

### Quick Start: ExprTk

1. **Switch to ExprTk** — Click the `(L)` indicator next to **Use Variables** so it shows `(E)`.

2. **Enable Use Variables** — Check the **Use Variables** option in the Replace interface.

3. **Try a simple expression:**

   - **Find:** `(\d+)`
   - **Replace:** `(?=num(1) * 2)`
   - **Input:** `42`
   - **Output:** `84`

   Every captured number is doubled. The `(?=...)` block is the expression; everything outside it is literal text in the output.

4. **Mix expressions and literal text:**

   - **Find:** `(\d+)`
   - **Replace:** `value=(?=num(1)) doubled=(?=num(1)*2)`
   - **Input:** `42`
   - **Output:** `value=42 doubled=84`

   A single replace string can contain **any number** of `(?=...)` blocks interleaved with literal text. This is the core idiom — small, focused expressions instead of one big script.

<br>

### Pattern Syntax: `(?= ... )`

Every ExprTk expression in a replace string is wrapped in `(?=` and `)`. The marker is intentionally three characters so it cannot be confused with regex syntax.

| Syntax           | Meaning                                                               |
|------------------|-----------------------------------------------------------------------|
| `(?=expr)`       | Evaluate the expression and insert its result at this position.       |
| `\(?=`           | Literal `(?=` — backslash escapes the marker.                         |
| `\\`             | Literal backslash.                                                    |
| any other text   | Literal output, including UTF-8 byte sequences.                       |

The closing `)` is matched by depth-counting, so balanced parentheses inside the expression — like `min(num(1), num(2))` — work naturally without extra escaping.

<br>

### Variables and Captures

ExprTk receives the same per-match context as Lua, but exposed differently because ExprTk is numeric-first. All variables below are registered under both upper- and lowercase — pick the style you prefer.

#### Numeric variables (read-only)

| Name          | Description                                                       |
|---------------|-------------------------------------------------------------------|
| `CNT` / `cnt`   | Replacement count across the whole run (1-based).               |
| `LCNT` / `lcnt` | Replacement count within the current line (1-based).            |
| `LINE` / `line` | 1-based line number of the current match.                       |
| `LPOS` / `lpos` | Column position within the line (UTF-8 bytes).                  |
| `APOS` / `apos` | Absolute byte position in the document.                         |
| `COL` / `col`   | CSV column index (CSV mode), 0 otherwise.                       |
| `HIT` / `hit`   | Numeric value of the full match (same as `num(0)`).             |

#### String variables (read-only, usable only in `return [...]` lists)

| Name              | Description                                  |
|-------------------|----------------------------------------------|
| `FPATH` / `fpath` | Full path of the document being processed.   |
| `FNAME` / `fname` | Filename without path.                       |

To access the matched text as a string, use `txt(0)` inside a `return [...]` list. To reference the match in the replace template directly, use `$0` (whole match) or an explicit capture group `(...)` referenced with `\1`.

#### Capture access

| Function  | Returns | Use case                                                              |
|-----------|---------|-----------------------------------------------------------------------|
| `num(N)`  | number  | Capture group N as `double`. `num(0)` is the full match. Non-numeric captures yield `NaN`. |
| `txt(N)` | string  | Capture group N as text. `txt(0)` is the full match. Only valid inside a `return [...]` list. |

#### Match-aware functions

| Function       | Returns | Use case                                                                                            |
|----------------|---------|-----------------------------------------------------------------------------------------------------|
| `seq([s,[i]])` | number  | Sequence value `s + (CNT-1)*i`. Both arguments default to `1` so `seq()` yields 1, 2, 3, ... See [Sequence Generator](#sequence-generator). |
| `skip()`       | -       | Mark the current match as untouched. See [Skipping Matches](#skipping-matches).                   |
| `numcol(N)` / `numcol('name')` | number | CSV column on the current row, parsed as `double`. NaN if missing or non-numeric. See [CSV Column Access](#csv-column-access). |
| `txtcol(N)` / `txtcol('name')` | string | CSV column as text. Only valid inside `return [...]`. See [CSV Column Access](#csv-column-access). |

**Number parsing notes:**

`num(N)` parses captures with both `.` and `,` accepted as decimal separator, tolerates trailing non-numeric characters, and yields `NaN` when the input has no leading digits. When an expression evaluates to `NaN` (or to `±Infinity` from things like `0/0`, `log(-1)`, `sqrt(-1)`), MultiReplace shows a dialog with three choices: skip just this match, skip every NaN for the rest of the run, or stop the run. Skipped matches are left untouched, so no original data is lost.

| Capture text                         | `num(N)`            | Note                                            |
|--------------------------------------|---------------------|-------------------------------------------------|
| `"42"`                               | `42`                |                                                 |
| `"3.14"`                             | `3.14`              |                                                 |
| `"3,14"`                             | `3.14`              | comma accepted as decimal                       |
| `"-7,5"`                             | `-7.5`              |                                                 |
| `""`, `"abc"`                        | `NaN`               | unparseable                                     |
| `"$100"`                             | `NaN`               | leading non-digit                               |
| `"3,14abc"`                          | `3.14`              | trailing characters tolerated                   |
| `"1,234,567"` (US thousands)         | `1.234`             | thousands separators not recognised — strip them in the regex first |
| `"1.234,56"` or `"1,234.56"`         | `1.234` or `1`      | mixed separators are unreliable — avoid relying on them |

<br>

### Math Functions

ExprTk ships with a substantially richer numeric library than Lua's `math.*`. The most useful built-ins for replacement work:

#### Basic

| Function                | Description                                                |
|-------------------------|------------------------------------------------------------|
| `abs(x)`                | Absolute value.                                            |
| `sgn(x)`                | Sign: -1, 0, or 1.                                         |
| `floor(x)`, `ceil(x)`   | Round down / up.                                           |
| `round(x)`              | Round to nearest integer.                                  |
| `roundn(x, n)`          | Round to `n` decimal places.                               |
| `trunc(x)`              | Truncate toward zero.                                      |
| `frac(x)`               | Fractional part.                                           |
| `sqrt(x)`, `root(x, n)` | Square root, n-th root.                                    |
| `exp(x)`, `expm1(x)`    | `e^x`, `e^x - 1`.                                          |
| `log(x)`, `log2(x)`, `log10(x)`, `log1p(x)`, `logn(x, n)` | Natural log and variants.    |
| `pow(x, n)` or `x^n`    | Exponentiation.                                            |
| `mod(a, b)` or `a % b`  | Modulo.                                                    |
| `clamp(lo, x, hi)`      | Constrain `x` to the range `[lo, hi]`.                     |
| `inrange(lo, x, hi)`    | 1 if `lo <= x <= hi`, else 0.                              |

#### Aggregates (variadic)

| Function                | Description                                                |
|-------------------------|------------------------------------------------------------|
| `min(a, b, ...)`        | Smallest argument. Any number of operands.                 |
| `max(a, b, ...)`        | Largest argument.                                          |
| `avg(a, b, ...)`        | Arithmetic mean.                                           |
| `sum(a, b, ...)`        | Sum.                                                       |
| `mul(a, b, ...)`        | Product.                                                   |

#### Trigonometry

`sin`, `cos`, `tan`, `asin`, `acos`, `atan`, `atan2(y, x)`, `sinh`, `cosh`, `tanh`, `asinh`, `acosh`, `atanh`, `cot`, `csc`, `sec`, `hypot(x, y)`, `deg2rad(x)`, `rad2deg(x)`.

#### Statistics

| Function   | Description                                                      |
|------------|------------------------------------------------------------------|
| `erf(x)`   | Error function.                                                  |
| `erfc(x)`  | Complementary error function.                                    |
| `ncdf(x)`  | Standard normal cumulative distribution (Z to percentile).       |

#### Constants

`pi`, `epsilon`, `inf`.

<br>

### Control Flow

ExprTk has expression-level control flow. Everything returns a value — there are no statements in the Lua sense.

| Form                                    | Meaning                                          |
|-----------------------------------------|--------------------------------------------------|
| `if (cond) a else b`                    | If/else expression. Last expression is the value.|
| `cond ? a : b`                          | Ternary operator. Same as the if-else above.     |
| `if (cond) { a; b; c }`                 | Block form, value is the last statement.         |
| `switch { case c1: v1; case c2: v2; default: v3 }` | Multi-way branch.                     |

Example — clamp a captured value but show "OVER" if it exceeded:

```
(?=num(1) > 100 ? 100 : num(1))
```

<br>

### String Output via `return [...]`

ExprTk expressions normally produce a number. To emit a **string** — e.g. to combine literal text with `txt(N)` capture text — wrap the output in a `return` list:

```
(?=return ['prefix-', txt(1), '-suffix'])
```

The list elements can be string literals, the string variables (`FPATH`, `FNAME`), `txt(N)` capture text (use `txt(0)` for the full match), and numeric expressions (which are converted to text). This is the only path for any non-numeric output.

> **String literals are ASCII-only.** UTF-8 bytes in `'...'` will fail to compile. Non-ASCII text must come from the document via captures or string variables, or be placed in the literal portion of the replace string outside `(?=...)`. See [Limitations](#limitations) below.

<br>

### Operators

| Operator                                          | Meaning                                           |
|---------------------------------------------------|---------------------------------------------------|
| `+`, `-`, `*`, `/`, `%`, `^`                       | Arithmetic and exponentiation.                    |
| `==`, `!=`, `<`, `<=`, `>`, `>=`                   | Comparison.                                       |
| `and`, `or`, `not`, `xor`, `nand`, `nor`           | Logical operators (also `&&`, `\|\|`, `!`).        |
| `:=`                                              | Assignment to a local variable.                   |
| `+=`, `-=`, `*=`, `/=`, `%=`                      | Compound assignment.                              |
| `?:`                                              | Ternary.                                          |

<br>

### Skipping Matches

Calling `skip()` from within an expression tells the engine to leave this match unchanged. The expression's own value is ignored when skip is signalled. Useful as a numeric filter:

```
(?=num(1) < 100 ? skip() : num(1) * 2)
```

This doubles captured numbers ≥ 100 and leaves everything else untouched.

<br>

### Sequence Generator

`seq([start[, increment]])` returns the next value in an arithmetic sequence keyed off the match counter `CNT`. Both arguments are optional and default to `1`, so:

```
(?=seq())            -> 1, 2, 3, 4, ...        (default start=1, inc=1)
(?=seq(10))          -> 10, 11, 12, 13, ...    (start=10, inc defaults to 1)
(?=seq(0, 10))       -> 0, 10, 20, 30, ...     (custom start and step)
(?=seq(100, -10))    -> 100, 90, 80, 70, ...   (negative step counts down)
```

It is exactly equivalent to `start + (CNT - 1) * inc`, just shorter and clearer at the use site. Useful for numbering matches, generating IDs or any other arithmetic series across replacements.

```
Input:    apple
          banana
          cherry
Pattern:  ^(.+)$
Replace:  (?=seq(100, 10)). \1
Output:   100. apple
          110. banana
          120. cherry
```

<br>

### CSV Column Access

When CSV mode is active, `numcol(N)` and `txtcol(N)` return the value of column N from the **physical line** containing the current match. Index is 1-based; `numcol(1)` is the first column. Both also accept a header name as a string (`numcol('price')`); on first use the document's first line is parsed as a header row and cached for the rest of the run.

| Function                       | Returns | Notes                                                                            |
|--------------------------------|---------|----------------------------------------------------------------------------------|
| `numcol(N)` / `numcol('name')` | number  | Same parsing as `num()` - `.`/`,` decimal separator, NaN on non-numeric content. |
| `txtcol(N)` / `txtcol('name')` | string  | Raw cell text. Only meaningful inside `return [...]`.                            |

**Requirements:**
- CSV mode must be active (the **CSV** scope radio button) and a delimiter configured.
- Indexing reads the raw row regardless of any column-selection filter the user set.
- The cell is returned verbatim - no quote stripping, matching the existing CSV column convention used by sort and column highlighting.

**Behaviour at the edges:**
- Missing column (e.g. `numcol(99)` on a row with 5 columns): `NaN` for `numcol`, empty string for `txtcol`.
- Unknown header name: same.
- CSV mode off, no delimiter set, or no current match in progress: same.

**Example - lookup by name with arithmetic:**

```
Header:  product;price;qty
Row:     apple;1.50;10
Row:     banana;0.30;25
```

```
Find:    ^(?!product).+$
Replace: (?=return [txt(0), ' total=', numcol('price') * numcol('qty')])
```

```
product;price;qty
apple;1.50;10 total=15
banana;0.30;25 total=7.5
```

The negative lookahead in the Find pattern excludes the header row so the formula only runs on data rows. Inside the formula, `numcol('price')` reads the `price` column of whichever line the current match landed on.

<br>

### More Examples

The examples below are drawn from typical text-processing tasks. Each row in the table is a **single rule** — copy the Find and Replace columns straight into MultiReplace, set the engine to ExprTk, enable Use Variables and Regex.

#### Numeric transformation

| Find          | Replace                                  | Description                                        |
|---------------|------------------------------------------|----------------------------------------------------|
| `(\d+\.?\d*)` | `(?=roundn(num(1), 2))`                  | Round all numbers to 2 decimal places.             |
| `(-?\d+\.?\d*)` | `(?=clamp(0, num(1), 100))`            | Clamp every value to the range `[0..100]`.         |
| `(-?\d+\.?\d*)` | `(?=abs(num(1)))`                      | Take the absolute value of every number.           |
| `(\d+\.?\d*)` | `(?=num(1) * 2.54)`                      | Convert inches to centimetres (×2.54).             |
| `(\d+\.?\d*)` | `(?=(num(1) - 32) * 5/9)`                | Convert °F to °C.                                  |

#### Conditional / filtering

| Find          | Replace                                                       | Description                                                                |
|---------------|---------------------------------------------------------------|----------------------------------------------------------------------------|
| `(\d+)`       | `(?=num(1) < 100 ? skip() : num(1))`                          | Keep only numbers ≥ 100; smaller ones stay untouched (skipped, not zeroed).|
| `(\d+\.?\d*)` | `(?=inrange(0, num(1), 1) ? num(1) : skip())`                 | Keep only values in `[0..1]`, skip the rest.                               |
| `(\d+\.?\d*)` | `(?=num(1) > avg(0, 100, 200) ? num(1) : skip())`             | Skip below-average values (here against a fixed reference average).        |

#### Mixed expressions in one rule

| Find                  | Replace                                                         | Description                                                          |
|-----------------------|-----------------------------------------------------------------|----------------------------------------------------------------------|
| `(\d+),(\d+)`         | `total=(?=num(1)+num(2)) avg=(?=avg(num(1),num(2)))`            | Two independent expressions in one replace, each in its own `(?=)`.  |
| `(\d+\.?\d*)`         | `(?=num(1)) cm = (?=num(1)/2.54) inches`                        | Output original and converted value side-by-side.                    |
| `(\d+\.?\d*),(\d+\.?\d*)` | `distance=(?=hypot(num(1), num(2)))`                        | Euclidean distance from two coordinate values.                       |

#### Statistics

| Find          | Replace                                       | Description                                                              |
|---------------|-----------------------------------------------|--------------------------------------------------------------------------|
| `(-?\d+\.?\d*)` | `(?=ncdf(num(1)))`                          | Z-score to percentile via the standard normal CDF (e.g. `1.96` → `0.975`).|
| `(-?\d+\.?\d*)` | `(?=erf(num(1)))`                           | Error function — handy for diffusion / probability calculations.         |

#### String composition

| Find          | Replace                                                 | Description                                                              |
|---------------|---------------------------------------------------------|--------------------------------------------------------------------------|
| `(\w+)`       | `(?=return ['<', txt(1), '>'])`                        | Wrap every match in angle brackets.                                      |
| `(\d+)`       | `(?=return ['#', CNT, ': ', txt(1)])`                  | Prefix each match with a running global counter.                         |
| `(\w+)`       | `(?=return [FNAME, ': ', txt(1)])`                     | Tag each match with the source filename.                                 |

#### Counters

| Find          | Replace                       | Description                                                              |
|---------------|-------------------------------|--------------------------------------------------------------------------|
| `^`           | `(?=CNT). `                   | Number every line (insert at the start of each line).                    |
| `^`           | `(?=seq(100, 10)). `          | Number every line, starting at 100, step 10 (100, 110, 120, ...).        |
| `^`           | `(?=seq(0, -1)). `            | Number every line, counting **down** from 0 (0, -1, -2, ...).            |
| `;`           | `(?=LCNT == 1 ? skip() : 0);` | Replace only the second and later semicolons on each line.               |

<br>

### Limitations

ExprTk is deliberately scoped. Things it does **not** do:

- **No string manipulation** — there is no `substring`, `replace`, `find`, `format`, etc. Strings can only be passed through (via captures, `txt(N)`, `FNAME`, `FPATH`) and assembled in `return [...]` lists.
- **UTF-8 inside string literals fails to compile.** A literal like `'Größe'` between `'...'` will produce an `Invalid string token` error. Use captures or place non-ASCII text outside the `(?=...)` block.
- **No state across matches.** Each `(?=...)` evaluation starts fresh. The numeric counters (`CNT`, `LCNT`) are provided by the host and read-only; for accumulating user-defined state across matches, switch to the Lua engine and use `vars({...})`.
- **No file I/O, no external scripts.** ExprTk has no equivalent to Lua's `lvars`, `lkp`, or `lcmd`.
- **Numeric-only by design.** Operations on text that go beyond passing it through belong in Lua.