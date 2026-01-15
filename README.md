# MultiReplace for Notepad++
[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-brightgreen)](https://github.com/daddel80/notepadpp-multireplace/blob/main/license.txt)
[![Latest Stable Version](https://img.shields.io/badge/version-5.0.0.35-blue)](https://github.com/daddel80/notepadpp-multireplace/releases/tag/5.0.0.35)
[![Total Downloads](https://img.shields.io/github/downloads/daddel80/notepadpp-multireplace/total?logo=github)](https://github.com/daddel80/notepadpp-multireplace/releases)

MultiReplace is a Notepad++ plugin designed for complex and repeatable text transformations. Its foundation is the management of search/replace operations in reusable lists, suitable for use across sessions and projects.

For data-heavy tasks, the plugin provides dedicated tools. These include specific operations on CSV columns (e.g., sort, delete) and the use of external files as lookup tables.

At its core, a rule engine allows any replacement to be enhanced with conditional (if/then) and mathematical logic, providing the flexibility needed for specific and complex requirements.

![MultiReplace Screenshot](./MultiReplace.gif)

## Table of Contents
- [Key Features](#key-features)
- [Search Configuration](#search-configuration)
  - [Search Mode](#search-mode)
  - [Match Options](#match-options)
- [Search Scopes and Targets](#search-scopes-and-targets)
  - [Search Scopes](#search-scopes)
  - [CSV Scope and Column Operations](#csv-scope-and-column-operations)
  - [Execution Targets](#execution-targets)
- [Search Results Window](#search-results-window)
- [Option 'Use Variables'](#option-use-variables)
  - [Quick Start: Use Variables](#quick-start-use-variables)
  - [Variables Overview](#variables-overview)
  - [Available Functions](#available-functions)
  - [Command Reference](#command-reference)
    - [set](#setstrorcalc)
    - [cond](#condcondition-trueval-falseval)
    - [vars](#varsvariable1value1-variable2value2-)
    - [lvars](#lvarsfilepath)
    - [lkp](#lkpkey-hpath-inner)
    - [fmtN](#fmtnnum-maxdecimals-fixeddecimals)
    - [lcmd](#lcmdpath)
  - [String Formatting Helpers](#string-formatting-helpers)
  - [Preload Variables and Helpers](#preload-variables-and-helpers)
  - [Operators](#operators)
  - [If-Then Logic](#if-then-logic)
  - [DEBUG Mode](#debug-mode)
  - [More Examples](#more-examples)
  - [Engine Overview](#engine-overview)
- [User Interaction and List Management](#user-interaction-and-list-management)
  - [Entry Interaction and Limits](#entry-interaction-and-limits)
  - [Context Menu and Keyboard Shortcuts](#context-menu-and-keyboard-shortcuts)
  - [List Columns](#list-columns)
  - [List Toggling](#list-toggling)
- [Column Locking](#column-locking)
- [Data Handling](#data-handling)
  - [List Saving and Loading](#list-saving-and-loading)
- [Settings and Customization](#settings-and-customization)
  - [1. Search and Replace](#1-search-and-replace)
  - [2. List View and Layout](#2-list-view-and-layout)
  - [3. CSV Options](#3-csv-options)
  - [4. Export](#4-export)
  - [5. Appearance](#5-appearance)
- [Multilingual UI Support](#multilingual-ui-support)

## Key Features

- **Batch Replacement Lists** — Run any number of search-and-replace pairs in a single pass, either in the current document or across all open documents.
- **CSV Column Toolkit** — Search, replace, sort, or delete specific columns; numeric-aware sorting and header exclusion included.
- **Reusable Replacement Lists** — Save, load, and drag-and-drop lists with all options intact—perfect for recurring workflows.
- **Rule-Driven & Variable Replacements** — Lua-powered variables, conditions, and calculations enable dynamic, context-aware substitutions.
- **External Lookup Tables** — Swap matches with values from external hash/lookup files—ideal for large or frequently updated mapping tables.
- **Open Scripting API** — Add your own Lua functions to handle advanced formatting, data logging, and fully custom replacement logic.
- **Precision Scopes & Selections** — Rectangle and multi-selection support, column scopes, and "replace at specific match" for pinpoint operations.
- **Multi-Color Highlighting** — Highlight search hits in up to 28 distinct colors for rapid visual confirmation.
- **Search Results Window** — A dedicated dockable panel displays all search hits with folding, color-coding, and one-click navigation.

<br>

## Search Configuration
This section describes the different modes and options to control how search patterns are interpreted and matched.

### Search Mode
The search mode determines how the text in the **Find what** field is interpreted.

- **Normal** — Treats the search string literally. Special characters like `.` or `*` have no special meaning.

- **Extended** — Allows the use of backslash escape sequences to find or insert special and non-printable characters.

  | Sequence | Description            | Example                                                             |
  |----------|------------------------|---------------------------------------------------------------------|
  | `\n`     | Line Feed (LF)         | Replaces all `\n` with a space to join lines.                       |
  | `\r`     | Carriage Return (CR)   | Useful for cleaning up Windows-style line endings (`\r\n`).         |
  | `\t`     | Tab character          | Replaces four spaces with `\t`.                                     |
  | `\0`     | NULL character         | Finds or inserts the NULL character.                                |
  | `\\`     | Literal Backslash      | To search for a single `\`, you must enter `\\`.                    |
  | `\xHH`   | Hexadecimal value      | `\x41` finds or inserts the character **A**.                        |
  | `\oNNN`  | Octal value            | `\o101` finds or inserts **A**.                                     |
  | `\dNNN`  | Decimal value          | `\d065` finds or inserts **A**.                                     |
  | `\uXXXX` | Unicode character      | `\u20AC` finds or inserts **€**. *(Extension not available in standard Notepad++.)* |

- **Regular expression** — Enables powerful pattern matching using the regex engine integrated into the Notepad++ editor component. It supports common syntax for character classes (`[a-z]`), quantifiers (`*`, `+`, `?`), and capture groups (`(...)`). For a detailed reference, see the official Notepad++ Regular Expressions documentation.

### Match Options
These options refine the search behavior across all modes.

- **Match Whole Word Only** — The search term is matched only if it is a whole word, surrounded by non-word characters.
- **Match Case** — Makes the search case-sensitive, treating `Hello` and `hello` as distinct terms.
- **Use Variables** — Allows the use of variables within the replacement string for dynamic and conditional replacements. See the [chapter 'Use Variables'](#option-use-variables) for details.
- **Wrap Around** — If active, the search continues from the beginning of the document after reaching the end.
- **Replace matches** — Applies to all **Replace All** actions (current document, all open docs, and in files). Allows you to specify exactly which occurrences of a match to replace. Accepts single numbers, commas, or ranges (e.g., `1,3,5-7`).

## Search Scopes and Targets
This section describes **where** to search (Scopes) and **in which files** (Targets).

### Search Scopes
Search scopes define the area within a document for search and replace operations.

- **All Text** — The entire document is searched.
- **Selection** — Operations are restricted to the selected text. Supports standard, rectangular (columnar), and multi-selections.

### CSV Scope and Column Operations
Selecting the **CSV** scope enables powerful tools for working with delimited data.

**Scope Definition:**
- **Cols** — Specify the target columns for the operation (e.g., `1,3,5-7`). For sorting, the sequence is crucial as it defines the priority (e.g., `3,1` sorts by column 3, then 1). Descending ranges like `5-3` are also supported.
- **Delim** — Define the delimiter character.
- **Quote** — Specify a quote character (`"` or `'`) to ignore delimiters within quoted text.

**Available Column Operations:**
- **Sorting Lines by Columns** — Sort lines based on one or more columns in ascending or descending order. The sorting algorithm correctly handles mixed numeric and text values.
  - **Smart Undo (Toggle Sort)** — A second click on the same sort button reverts the lines to their original order. This powerful undo works even if rows have been modified, added, or deleted after the initial sort.
  - **Exclude Header Lines** — You can protect header rows from being sorted. Configure the number of header rows in [Settings > CSV Options](#3-csv-options).
- **Deleting Multiple Columns** — Remove specified columns at once, automatically cleaning up obsolete delimiters.
- **Clipboard Column Copying** — Copy the content of specified columns, including their delimiters, to the clipboard.
- **Flow Tabs Alignment** — Visually aligns columns in tab-delimited and CSV files for easier reading and editing.
  - For CSV files, temporary tabs are inserted to simulate uniform column spacing. For tab-delimited files, existing tabs are realigned.
  - The **Align Columns** button toggles alignment on or off; pressing it again restores the original spacing.
  - Numeric values are right-aligned by default; this behavior can be turned off in [Settings > CSV Options](#3-csv-options).

### Execution Targets
Execution targets define **which files** an operation is applied to. They are accessible via the **Replace All** split-button menu.

- **Replace All** — Executes the replacement in the **current document only**.
- **Replace All in All Open Docs** — Executes the replacement across **all open files** in Notepad++.
- **Replace in Files** — Extends the replacement scope to entire directory structures. When selected, the main window expands to show a dedicated panel for configuration.
  - **Directory:** The starting folder for the file search.
  - **Filters:** Space-separated list of patterns to include or exclude files and folders.
  - **In Subfolders:** Recursively include all subdirectories.
  - **In Hidden Files:** Include hidden files and folders.
- **Debug Mode** — Runs a simulation of the replacement to inspect variables without modifying the document. See [Debug Mode](#debug-mode) for details.

**Filter Syntax**

| Prefix   | Example      | Description                                                            |
|----------|--------------|------------------------------------------------------------------------|
| *(none)* | `*.cpp *.h`  | Includes files matching the pattern.                                   |
| `!`      | `!*.bak`     | Excludes files matching the pattern.                                   |
| `!\`     | `!\obj\`     | Excludes the specified folder *non-recursive*.                         |
| `!+\`    | `!+\logs\`   | Excludes the specified folder **and** all its subfolders *recursive*.  |

**Operation Control**

- **Progress Feedback** — A message line shows real-time progress (percentage and current file).
- **Cancel Button** — A **Cancel** button appears during long operations, allowing safe abort.
- **Encoding Handling** — Encoding (ANSI, UTF-8, UTF-16, etc.) is auto-detected and preserved when writing changes back to disk.

<br>

## Search Results Window

The **Search Results Window** is a dedicated dockable panel that displays all matches from **Find All** operations. It provides a structured, navigable view of search results across one or multiple files.

### Features

- **Hierarchical Display** — Results are organized in a collapsible tree structure: Search Header → File → Criterion → Individual Hits. Each level can be expanded or collapsed for easy navigation.
- **Color-Coded Matches** — When using the replacement list, each search term is highlighted in its own distinct color (up to 28 colors). This visual distinction makes it easy to identify which list entry produced each match.
- **One-Click Navigation** — Double-clicking any result line immediately opens the corresponding file (if not already open) and jumps to the exact match position in the editor.
- **Context Menu** — Right-click to access options like Copy Lines, Copy Paths, Open Selected Files, Fold/Unfold All, and Clear Results.
- **Persistent Results** — By default, new searches append to existing results, allowing you to accumulate findings. Enable "Purge on new search" in the context menu to clear previous results automatically.
- **Word Wrap** — Toggle word wrap via the context menu for long lines.

### Options

The following options in [Settings > Appearance](#4-appearance) control the Search Results Window behavior:

- **Use list colors in search results** — When checked, matches are color-coded to match their corresponding list entry. When unchecked, all results use a uniform standard color.

<br>

## Option 'Use Variables'

Enable the '**Use Variables**' option to enhance replacements with calculations and logic based on the matched text. This feature lets you create dynamic replacement patterns, handle conditions, and produce flexible outputs—all configured directly in the Replace field. This functionality relies on the [Lua engine](https://www.lua.org/).

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

| Variable | Description |
|----------|-------------|
| **CNT**  | Count of the detected string. |
| **LCNT** | Count of the detected string within the line. |
| **LINE** | Line number where the string is found. |
| **LPOS** | Relative character position within the line. |
| **APOS** | Absolute character position in the document. |
| **COL**  | Column number where the string was found (CSV-Scope option selected). |
| **MATCH**| Contains the text of the detected string, in contrast to `CAP` variables which correspond to capture groups in regex patterns. |
| **FNAME**| Filename or window title for new, unsaved files. |
| **FPATH**| Full path including the filename, or empty for new, unsaved files. |
| **CAP1**, **CAP2**, ... | Equivalents to regex capture groups, designed for use in the 'Use Variables' environment. Always strings; use `tonum(CAP1)` for calculations. Note: Standard counterparts (`$1`, `$2`, ...) cannot be used in this environment. |

**Notes:**
- `FNAME` and `FPATH` are updated for each file processed by `Replace All in All Open Docs` and `Replace All in Files`. This ensures that variables always refer to the file currently being modified.
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

Custom variables persist from match to match within a single **'Replace All'** operation and can transfer values between list entries. Each new operation (**button click**) starts with a fresh state. In multi-file replacements, variables also reset at each **document switch**.

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

### Engine Overview
MultiReplace uses the [Lua engine](https://www.lua.org/), allowing for Lua math operations and string methods. Refer to [Lua String Manipulation](https://www.lua.org/manual/5.4/manual.html#6.4) and [Lua Mathematical Functions](https://www.lua.org/manual/5.4/manual.html#6.6) for more information.

<br>

## User Interaction and List Management
Manage search and replace strings within the list using the context menu, which provides comprehensive functionalities accessible by right-clicking on an entry, using direct keyboard shortcuts, or mouse interactions.

### Entry Interaction and Limits
- **Manage Entries** — Manage search and replace strings in a list, and enable or disable entries for replacement, highlighting, or searching within the list.
- **Highlighting** — Highlight multiple find words in unique colors for better visual distinction, with up to 28 distinct colors available.
- **Character Limit** — Find and replace texts have no fixed length limit. Very long texts may slow down processing.

### Context Menu and Keyboard Shortcuts
Right-click on any entry in the list or use the corresponding keyboard shortcuts to access these options:

| Menu Item                | Shortcut      | Description                                                 |
|--------------------------|---------------|-------------------------------------------------------------|
| Undo                     | Ctrl+Z        | Reverts the last change made to the list, including sorting and moving rows. |
| Redo                     | Ctrl+Y        | Reapplies the last action that was undone, restoring previous changes. |
| Transfer to Input Fields | Alt+Up        | Transfers the selected entry to the input fields for editing. |
| Search in List           | Ctrl+F        | Initiates a search within the list entries. Inputs are entered in the "Find what" and "Replace with" fields. |
| Cut                      | Ctrl+X        | Cuts the selected entry to the clipboard.                   |
| Copy                     | Ctrl+C        | Copies the selected entry to the clipboard.                 |
| Paste                    | Ctrl+V        | Pastes content from the clipboard into the list.            |
| Edit Field               |               | Opens the selected list entry for direct editing.           |
| Delete                   | Del           | Removes the selected entry from the list.                   |
| Select All               | Ctrl+A        | Selects all entries in the list.                            |
| Enable                   | Alt+E         | Enables the selected entries, making them active for operations. |
| Disable                  | Alt+D         | Disables the selected entries to prevent them from being included in operations. |

**Note on 'Edit Field':** The edit field supports multiple lines, simplifying the management of complex 'Use Variables' statements.

**Additional Interactions:**
- **Space Key** — Toggles the activation state of selected entries.
- **Double-Click** — Allows direct in-place editing (configurable in Settings).

### List Columns
- **Find** — The text or pattern to search for.
- **Replace** — The text or pattern to replace with.
- **Options Columns:**
  - **W** — Match whole word only.
  - **C** — Match case.
  - **V** — Use Variables.
  - **E** — Extended search mode.
  - **R** — Regular expression mode.
- **Additional Columns:**
  - **Matches** — Displays the number of times each 'Find what' string is detected.
  - **Replaced** — Shows the number of replacements made for each 'Replace with' string.
  - **Comments** — Add custom comments to entries for annotations or additional context.
  - **Delete** — Contains a delete button for each entry, allowing quick removal from the list.

You can manage the visibility of the additional columns via the **Header Column Menu** by right-clicking on the header row.

### List Toggling
The **"Use List"** button toggles between processing the entire list or just the single "Find what" / "Replace with" fields.

## Column Locking

You can lock specific column widths to prevent them from resizing automatically when the window layout changes. This is particularly useful for keeping key columns like **Find**, **Replace**, or **Comments** visible at a fixed size.

- **How to Lock** — Double-click the column divider line in the list header.
- **Visual Feedback** — A lock icon (🔒) appears in the header of the locked column.
- **Effect** — Locked columns maintain their exact pixel width, while unlocked columns adjust dynamically to fill the remaining space.

## Data Handling

### List Saving and Loading
- **Save List / Load List** — Store and reload your search/replace entries as CSV files.
- **Drag & Drop** — You can load lists by dragging CSV files onto the plugin window.
- **Format** — Files follow RFC 4180 CSV standards.

## Settings and Customization

You can customize the behavior and appearance of MultiReplace via the dedicated **Settings Panel**. To open it, click the **Settings** entry in the plugin menu or the context menu.

### 1. Search and Replace
Control the behavior of search operations and cursor movement.

- **Replace: Don't move to the following occurrence** — If checked, the cursor remains on the current match after clicking "Replace". If unchecked, it automatically jumps to the next match (standard behavior).
- **Find: Search from cursor position** — If checked, "Find All" and "Replace All" start from the current cursor position. If unchecked, operations always process the entire defined scope.
- **Mute all sounds** — Disables the notification sound (beep) when a search yields no results. The window will still flash visually to indicate "not found".
- **File Search: Skip files larger than** — Defines a maximum size (in MB) for **Find/Replace in Files**. Skips larger files to ensure responsiveness and prevent high memory usage.

### 2. List View and Layout
Manage the visual elements and behavior of the replacement list to save screen space or increase information density.

- **Visible Columns:**
  - **Matches** — Toggles the column displaying the number of hits for each entry.
  - **Replaced** — Toggles the column displaying the number of replacements made.
  - **Comments** — Toggles the user comment column.
  - **Delete Button** — Toggles the column containing the 'X' button for deleting rows.
- **List Results:**
  - **Show list statistics** — Displays a summary line below the list (Active entries, Total items, Selected items).
  - **Find All: Group hits by list entry** — When On, search results in the docking window are grouped hierarchically by the list entry that found them. When Off, results are displayed as a flat list sorted by position.
- **List Interaction & View:**
  - **Highlight current match in list** — Automatically highlights the list row corresponding to the current find/replace operation.
  - **Edit in-place on double-click** — When On, double-clicking a cell allows editing the text directly. When Off, double-clicking transfers the entry content to the top input fields.
  - **Show full text on hover** — Displays a tooltip with the complete text for long entries that are truncated in the view.
  - **Expanded edit height (lines)** — Defines how many lines the in-place edit box expands to when modifying multiline text (Range: 2–20).

### 3. CSV Options
Settings specific to the CSV column manipulation and alignment features.

- **Flow Tabs: Right-align numeric columns** — When using the **Flow Tabs** feature (Column Alignment), numeric values will be right-aligned within their columns for better readability. Text remains left-aligned.
- **Flow Tabs: Don't show intro message** — Suppresses the informational dialog that appears when activating Flow Tabs for the first time.
- **CSV Sort: Header lines to exclude** — Specifies the number of lines at the top of the file to protect from sorting operations (e.g., set to `1` to keep the header row fixed at the top).

### 4. Export
Configure how list data is exported to the clipboard via **Export Data** from the context menu.

- **Template** — Defines the output format using placeholders. Available placeholders:
  - `%FIND%` — Find pattern
  - `%REPLACE%` — Replace text
  - `%COMMENT%` — Comment
  - `%FCOUNT%` — Find count
  - `%RCOUNT%` — Replace count
  - `%ROW%` — Row number
  - `%SEL%` — Selected state (1/0)
  - `%REGEX%`, `%CASE%`, `%WORD%`, `%EXT%`, `%VAR%` — Option flags (1/0)
  - Use `\t` for tab delimiter (e.g., `%FIND%\t%REPLACE%\t%COMMENT%` for TSV format).
- **Escape special characters** — When checked, converts newlines, tabs, and backslashes in field values to their escape sequences (`\n`, `\t`, `\\`). Useful for single-line formats.
- **Include header row** — When checked, adds a header line with column names before the data rows.

### 5. Appearance
Customize the look and feel of the plugin window.

- **Interface:**
  - **Foreground transparency** — Sets the window opacity when the plugin is active/focused.
  - **Background transparency** — Sets the window opacity when the plugin loses focus.
  - **Scale factor** — Scales the entire plugin UI (buttons, text, list size). Useful for high-DPI monitors or accessibility preferences (Range: 50% to 200%).
- **Display Options:**
  - **Enable tooltips** — Toggles the helper popups that appear when hovering over buttons and options.
  - **Use list colors in search results** — When checked, search hits in the result docking window are color-coded matching the specific color assigned to their list entry. When unchecked, all search results use a uniform standard color.
  - **Use list colors for text marking** — When checked, highlights matches in the editor using the specific color defined in the list entry. When unchecked, all marked text uses a single standard highlight color.

<br>

> **Note:** All settings configured in this panel are automatically saved to `MultiReplace.ini` in your Notepad++ plugins configuration directory.

### INI-Only Settings

Some advanced options are not exposed in the UI and can only be configured by editing `MultiReplace.ini` directly:

| Section | Key | Default | Description |
|---------|-----|---------|-------------|
| `[Lua]` | `SafeMode` | `0` | When set to `1`, disables Lua libraries that access the file system (`os`, `io`, `package`, `dofile`, etc.) for security. When `0`, full Lua functionality is enabled (required for `lvars`, `lkp`, and `lcmd`). |

## Multilingual UI Support

MultiReplace supports multiple languages. The plugin automatically detects the language selected in Notepad++ and applies the corresponding translation if available.

You can customize the translations by editing the `languages.ini` file located in:
`C:\Program Files\Notepad++\plugins\MultiReplace\` (or your specific installation path).

Contributions to the `languages.ini` file on GitHub are welcome!
