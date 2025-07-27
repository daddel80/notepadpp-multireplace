# MultiReplace for Notepad++
[![License: GPL-3.0](https://img.shields.io/badge/license-GPL--3.0-brightgreen)](https://github.com/daddel80/notepadpp-multireplace/blob/main/license.txt)
[![Latest Stable Version](https://img.shields.io/badge/version-4.5.0.30-blue)](https://github.com/daddel80/notepadpp-multireplace/releases/tag/4.5.0.30)
[![Total Downloads](https://img.shields.io/github/downloads/daddel80/notepadpp-multireplace/total?logo=github)](https://github.com/daddel80/notepadpp-multireplace/releases)

MultiReplace is a Notepad++ plugin that allows users to create, store, and manage search and replace strings within a list, perfect for use across different sessions or projects. It increases efficiency by enabling multiple replacements at once, supports sorting and applying operations to specific columns in CSV files, and offers flexible options for replacing text, including conditional and mathematical operations, as well as the use of external hash tables for dynamic data lookups.

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
- [Option 'Use Variables'](#option-use-variables)
  - [Quick Start: Use Variables](#quick-start-use-variables)  
  - [Variables Overview](#variables-overview)
  - [Command Overview](#command-overview)
    - [set](#setstrorcalc)
    - [cond](#condcondition-trueval-falseval)
    - [vars](#varsvariable1value1-variable2value2-)
    - [lvars](#lvarsfilepath)
    - [lkp](#lkpkey-hpath-inner)
    - [fmtN](#fmtnnum-maxdecimals-fixeddecimals)
  - [Preloading Variables](#preloading-variables)
  - [Operators](#operators)
  - [If-Then Logic](#if-then-logic)
  - [DEBUG option](#debug-option)
  - [Examples](#more-examples)
- [User Interaction and List Management](#user-interaction-and-list-management)
  - [Context Menu and Keyboard Shortcuts](#context-menu-and-keyboard-shortcuts)
  - [List Columns](#list-columns)
  - [List Toggling](#list-toggling)
- [Data Handling](#data-handling)
  - [List Saving and Loading](#list-saving-and-loading)
  - [Bash Script Export (optional)](#bash-script-export-optional)
- [UI and Behavior Settings](#ui-and-behavior-settings)
  - [Column Locking](#column-locking)
  - [Configuration Settings](#configuration-settings)
  - [Multilingual UI Support](#multilingual-ui-support)
  
## Key Features

- **Batch Replacement Lists** – Run any number of search-and-replace pairs in a single pass, either in the current document or across all open documents.
- **Rule-Driven & Variable Replacements** – Lua-powered variables, conditions, and calculations enable dynamic, context-aware substitutions.
- **CSV Column Toolkit** – Search, replace, sort, or delete specific columns; numeric-aware sorting and header exclusion included.
- **External Lookup Tables** – Swap matches with values from external hash/lookup files—ideal for large or frequently updated mapping tables.
- **Precision Scopes & Selections** – Rectangle and multi-selection support, column scopes, and “replace at specific match” for pinpoint operations.
- **Reusable Replacement Lists** – Save, load, and drag-and-drop lists with all options intact—perfect for recurring workflows.
- **Multi-Colour Highlighting** – Highlight search hits in up to 20 distinct colours for rapid visual confirmation.
- **Script Export** – Generate standalone bash scripts that reproduce the replacement logic outside of Notepad++.

<br>

## Search Configuration
This section describes the different modes and options to control how search patterns are interpreted and matched.

### Search Mode
The search mode determines how the text in the **Find what** field is interpreted.

- **Normal**  
  Treats the search string literally. Special characters like `.` or `*` have no special meaning.

- **Extended**  
  Allows the use of backslash escape sequences to find or insert special and non-printable characters.

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
  | `\uXXXX` | Unicode character      | `\u20AC` finds or inserts **€**. *(This is an extension not available in standard Notepad++.)* |

- **Regular expression**  
  Enables powerful pattern matching using the regex engine integrated into the Notepad++ editor component. It supports common syntax for character classes (`[a-z]`), quantifiers (`*`, `+`, `?`), and capture groups (`(...)`). For a detailed reference, see the official Notepad++ Regular Expressions documentation.

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

- **Scope Definition:**
  - **Cols**: Specify the columns for focused operations (e.g., `1,3-5`).
  - **Delim**: Define the delimiter character.
  - **Quote**: Specify a quote character (`"` or `'`) to ignore delimiters within quoted text.

- **Available Column Operations:**
  - **Sorting Lines by Columns**: Sort lines based on one or more columns in ascending or descending order. The sorting algorithm correctly handles mixed numeric and text values.  
    - **Smart Undo&nbsp;(Toggle Sort)**: A second click on the same sort button reverts the lines to their original order. This powerful undo works even if rows have been modified, added, or deleted after the initial sort.  
    - **Exclude Header Lines**: You can protect header rows from being sorted. Configure the number of header rows via the `HeaderLines` parameter in the INI file. For details, see the [`INI File Settings`](#configuration-settings).
  - **Deleting Multiple Columns**: Remove specified columns at once, automatically cleaning up obsolete delimiters.
  - **Clipboard Column Copying**: Copy the content of specified columns, including their delimiters, to the clipboard.

### Execution Targets  
Execution targets define **which files** an operation is applied to. They are accessible via the **Replace All** split-button menu.

- **Replace All**  
    Executes the replacement in the **current document only**.
- **Replace All in All Open Docs**  
    Executes the replacement across **all open files** in Notepad++.
- **Replace in Files**  
    Extends the replacement scope to entire directory structures. When selected, the main window expands to show a dedicated panel for configuration.
    - **Directory:** The starting folder for the file search.
    - **Filters:** Space-separated list of patterns to include or exclude files and folders.
    - **In Subfolders:** Recursively include all subdirectories.
    - **In Hidden Files;** Include hidden files and folders.

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

## Option 'Use Variables'
Enable the '**Use Variables**' option to enhance replacements with calculations and logic based on the matched text. This feature lets you create dynamic replacement patterns, handle conditions, and produce flexible outputs—all configured directly in the Replace field. This functionality relies on the [Lua engine](https://www.lua.org/).

---

### Quick Start: Use Variables

1. **Enable "Use Variables"**  
   Check the "**Use Variables**" option in the Replace interface.  

2. **Use a Command**  

   **Option 1:** [`set(...)`](#setstrorcalc) → Directly replaces with a value or calculation.  
   - **Find**: `(\d+)`  
   - **Replace**: `set(CAP1 * 2)`  
   - **Result**: Doubles numbers (e.g., `50` → `100`).  
   *(Enable "Regular Expression" in 'Search Mode' to use `(\d+)` as a capture group.)*  

   **Option 2:** [`cond(...)`](#condcondition-trueval-falseval) → Conditional replacement.  
   - **Find**: `word`  
   - **Replace**: `cond(CNT==1, "FirstWord")`  
   - **Result**: Changes only the **first** occurrence of "word" to "FirstWord".
    
3. **Use Basic Variables:**  
   - **`CNT`**: Inserts the current match number (e.g., "1" for the first match, "2" for the second).
   - **`CAP1`**, **`CAP2`**, etc.: Holds captured groups when Regex is enabled.  
     > **Capture Groups:**  
     > With a regex in parentheses `(...)`, matched text is stored in `CAP` variables (e.g., `(\d+)` in `Item 123` stores `123` in `CAP1`).  For more details, refer to regex documentation.

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
| **COL**  | Column number where the string was found (CSV-Scope option selected).|
| **MATCH**| Contains the text of the detected string, in contrast to `CAP` variables which correspond to capture groups in regex patterns. |
| **FNAME**| Filename or window title for new, unsaved files. |
| **FPATH**| Full path including the filename, or empty for new, unsaved files. |
| **CAP1**, **CAP2**, ...  | These variables are equivalents to regex capture groups, designed for use in the 'Use Variables' environment. They are specifically suited for calculations and conditional operations within this environment. Although their counterparts ($1, $2, ...) cannot be used here.|

**Note:**
- FNAME and FPATH are updated for each file processed by `Replace All in All Open Docs` and `Replace All in Files`. This ensures that variables always refer to the file currently being modified.
- **Decimal Separator:** When `MATCH` and `CAP` variables are used to read numerical values for further calculations, both dot (.) and comma (,) can serve as decimal separators. However, these variables do not support the use of thousands separators.

<br>

### Command Overview
#### String Composition
`..` is employed for concatenation.  
E.g., `"Detected "..CNT.." times."`

<br>

#### set(strOrCalc)
Directly outputs strings or numbers, replacing the matched text with a specified or calculated value.

| Find      | Replace with                            | Regex | Description/Expected Output                                                 |
|-----------|-----------------------------------------|-------|-----------------------------------------------------------------------------|
| `apple`   | `set("banana")`                         | No    | Replaces every occurrence of `apple` with the string `"banana"`.           |
| `(\d+)`   | `set(CAP1 * 2)`                         | Yes   | Doubles any found number; e.g., `10` becomes `20`.                         |
| `found`   | `set("Found #"..CNT.." at position "..APOS)` | No    | Shows how many times `found` was detected and its absolute position.       |

<br>

#### cond(condition, trueVal, [falseVal])
Evaluates the condition and outputs `trueVal` if the condition is true, otherwise `falseVal`. If `falseVal` is omitted, the original text remains unchanged when the condition is false.

| Find        | Replace with                                                                          | Regex | Description/Expected Output                                                                                             |
|-------------|---------------------------------------------------------------------------------------|-------|-------------------------------------------------------------------------------------------------------------------------|
| `word`      | `cond(CNT==1, "First 'word'", "Another 'word'")`                                      | No    | For the first occurrence of `word` → `"First 'word'"`; for subsequent matches → `"Another 'word'"`.                    |
| `(\d+)`     | `cond(CAP1>100, "Large number", cond(CAP1>50, "Medium number", "Small number"))`      | Yes   | For a numeric match: if > 100 → `"Large number"`, if > 50 → `"Medium number"`, otherwise → `"Small number"`.           |
| `anymatch`  | `cond(APOS<50, "Early in document", "Later in document")`                             | No    | If the absolute position `APOS` is under 50 → `"Early in document"`, otherwise → `"Later in document"`.                |

<br>

#### **vars({Variable1=Value1, Variable2=Value2, ...})**
**Note:** This command was previously named `init(...)` and has been renamed to `vars(...)`. For compatibility, `init(...)` still works.

Initializes custom variables for use in various commands, extending beyond standard variables like `CNT`, `MATCH`, `CAP1`. These variables can carry the status of previous find-and-replace operations to subsequent ones.

Custom variables maintain their values throughout a single Replace-All or within a list of multiple Replace operations. Thus, they can transfer values from one list entry to subsequent ones. They reset at the start of each new document in **'Replace All in All Open Documents'**.

> **Tip**: To learn how to preload variables using an empty Find field before the main replacement process starts, see [Preloading Variables](#preloading-variables).

| **Find**        | **Replace**                                                                                                                            | **Before**                                     | **After**                                              | **Regex** | **Scope CSV** | **Description**                                                                                                              |
|-----------------|----------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------|--------------------------------------------------------|----------|--------------|------------------------------------------------------------------------------------------------------------------------------|
| `(\d+)`         | `vars({COL2=0,COL4=0}); cond(LCNT==4, COL2+COL4);`<br>`if COL==2 then COL2=CAP1 end;`<br>`if COL==4 then COL4=CAP1 end;`               | `1,20,text,2,0`<br>`2,30,text,3,0`<br>`3,40,text,4,0` | `1,20,text,2,22.0`<br>`2,30,text,3,33.0`<br>`3,40,text,4,44.0` | Yes      | Yes          | Tracks values from columns 2 and 4, sums them, and updates the result for the 4th match in the current line.                |
| `\d{2}-[A-Z]{3}`| `vars({MATCH_PREV=''}); cond(LCNT==1,'Moved', MATCH_PREV); MATCH_PREV=MATCH;`                                                           | `12-POV,00-PLC`<br>`65-SUB,00-PLC`<br>`43-VOL,00-PLC`  | `Moved,12-POV`<br>`Moved,65-SUB`<br>`Moved,43-VOL`      | Yes      | No           | Uses `MATCH_PREV` to track the first match in the line and shift it to the 2nd (`LCNT`) match during replacements.           |

<br>

#### **lvars(filePath)**

Loads custom variables from an external file. The file specifies variable names and their corresponding values. The loaded variables can then be used throughout the Replace process, similar to how variables defined with [`vars`](#varsvariable1value1-variable2value2-) work.

The parameter **filePath** must specify a valid path to a file. Supported path formats include:
- Escaped Backslashes: `"C:\\path\\to\\file.vars"`
- Forward Slashes: `"C:/path/to/file.vars"`
- Long Bracket String: `[[C:\path\to\file.vars]]`

**File:**
```lua
-- Local variables remain private
local PATH = [[C:\Data\Projects\\]]

-- Only the returned variables are accessible in Replace operations
return {
    userName = "Alice",
    threshold = 10,
    enableFeature = true,
    fullPath = PATH .. "dataFile.lkp" -- Combine a local variable with a string
}
```

> **Tip**: To learn how to preload variables using an empty Find field before the main replacement process starts, see [Preloading Variables](#preloading-variables).

| Find          | Replace                                                                       | Regex | Scope CSV | Description                                                                                          |
|---------------|-------------------------------------------------------------------------------|-------|-----------|------------------------------------------------------------------------------------------------------|
| *(empty)*     | `lvars([[C:\tmp\m\Vars.vars]])`                                              | No    | No        | Loads variables such as `userName = "Alice"` and `threshold = 10` from `myVars.vars`.               |
| `Hello`       | `set(userName)`                                                              | No    | No        | Replaces `Hello` with the value of the variable `userName`, e.g., `"Alice"`.                        |
| `(\d+)`       | `cond(threshold > 5, "Above", "Below")`                                      | Yes   | No        | Replaces the match based on the condition evaluated using the variable `threshold`.                 |

**Key Points**
- **Conditional Loading**: Variables can be loaded conditionally by placing lvars(filePath) alongside a specific Find pattern. In this case, the variables are only initialized when the pattern matches.
- **Local vs. Returned Variables**: Only variables explicitly included in the return table of the .vars file are available for use. Any local variables remain private to the file.

<br>

#### **lkp(key, hpath, inner)**
Performs an external lookup of **key** against an indexed data file located at **hpath** and returns the corresponding value. By default, if the **key** is not found, `lkp()` simply reverts to the key itself. Setting **inner** to `true` instead yields a `nil` result when the key is missing, allowing for conditional checks or deeper nested logic.

##### Key and File Path
- **Key**:  
  The **key** can be either a string or a number. Numbers are automatically converted to strings to ensure compatibility in the lookup process.
  
- **File Path (hpath)**:  
  The **hpath** must point to a valid `.lkp` file that returns a table of data.

  **Supported Path Formats**:  
  - Escaped Backslashes: `"C:\\path\\to\\file.lkp"`  
  - Forward Slashes: `"C:/path/to/file.lkp"`  
  - Long Bracket String: `[[C:\path\to\file.lkp]]`  

##### Data File Format
Each lkp file must be defined as a table of entries in the form `{ [keys], value }`, where `[keys]` can be:
- A single key (e.g., `"001"`).  
- An array of keys (e.g., `{ "001", "1", 1 }`) mapping to the same value.

**Example**:
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

##### Caching Mechanism
Once `lkp()` loads the data file for **hpath**, the parsed table is cached in memory for the duration of the Replace-All operation.

##### inner Flag
- **`false` (default, can be omitted)**: If the key is not found, `lkp()` returns the **search term itself** (e.g., `MATCH`, `CAP1`), instead of a mapped value.
- **`true`**: If the key is not found, `lkp()` returns `nil`, allowing conditional handling.

##### Examples

| **Find**   | **Replace**                                                                                               | **Regex** | **Scope CSV** | **Description**                                                                                      |
|------------|-----------------------------------------------------------------------------------------------------------|-----------|---------------|------------------------------------------------------------------------------------------------------|
| `\b\w+\b`  | `lkp(MATCH, [[C:\tmp\hash.lkp]], true)`                                                                    | Yes       | No            | Uses **inner = true**: If a match is found in the lookup file, replaces it with the mapped value. If no match is found, the original word is removed. |
| `(\d+)`    | `lkp(CAP1, "C:/path/to/myLookupFile.lkp")`                                                                | Yes       | No            | Uses **inner = false** (default): If a match is found in the lookup file, replaces it with the mapped value. If no match is found, the original text (`CAP1`) is returned. |
| `\b\w+\b`  | `output = lkp(MATCH, [[C:\tmp\hash.lkp]], true).result; set(output or "NoKey")`                           | Yes       | No            | Uses **inner = true**: If the lookup result is non-`nil`, replaces with the mapped value. Otherwise, replaces with `"NoKey"`. |
| `\b\w+\b`  | `cond(COL==3, lkp(MATCH, [[C:/tmp/col3_hash.lkp]]))`                                               | No        | Yes           | Looks up values in the third column (`COL==3`) using a separate lookup file (`col3_hash.lkp`). If a match is found, replaces it; otherwise, leaves it unchanged. |

<br>

#### **fmtN(num, maxDecimals, fixedDecimals)**
Formats numbers based on precision (maxDecimals) and whether the number of decimals is fixed (fixedDecimals being true or false).

**Note**: The `fmtN` command can exclusively be used within the `set` and `cond` commands.
| Example                             | Result  |
|-------------------------------------|---------|
| `set(fmtN(5.73652, 2, true))`       | 5.74  |
| `set(fmtN(5.0, 2, true))`           | 5.00  |
| `set(fmtN(5.73652, 4, false))`      | 5.7365|
| `set(fmtN(5.0, 4, false))`          | 5     |

<br>

### **Preloading Variables**
MultiReplace supports **predefining or loading variables** before any replacements occur. By separating initialization from the actual replacements, operations stay clean and maintainable.

#### 🔹 **How it works:**
- **Place `vars()` or `lvars()` next to an empty Find field.**
- This entry does **not** search for matches but runs before replacements begin.
- It ensures that **variables are loaded once**, regardless of their position in the list.

**Examples**  

| **Find**    | **Replace**                                                 | **Description** |
|--------------|------------------------------------------------------------|----------------|
| *(empty)*   | `vars({prefix = "ID_"})`                                    | Sets `prefix = "ID_"` before replacements. |
| *(empty)*   | `lvars([[C:\path\to\myVars.vars]])`                         | Loads external variables from a file. |
| `(\d+)`     | `set(prefix .. CAP1)`                                       | Uses `prefix` from initialization (e.g., `123` → `ID_123`). |

<br>

### Operators 
| Type        | Operators                     |
|-------------|-------------------------------|
| Arithmetic  | `+`, `-`, `*`, `/`, `^`, `%`  |
| Relational  | `==`, `~=`, `<`, `>`, `<=`, `>=`|
| Logical     | `and`, `or`, `not`            |

<br>

### If-Then Logic
If-then logic is integral for dynamic replacements, allowing users to set custom variables based on specific conditions. This enhances the versatility of find-and-replace operations.

**Note**: Do not embed `cond()`, `set()`, or `vars()` within if statements; if statements are exclusively for adjusting custom variables.

##### Syntax Combinations
- `if condition then ... end`
- `if condition then ... else ... end`
- `if condition then ... elseif another_condition then ... end`
- `if condition then ... elseif another_condition then ... else ... end`

##### Example
This example shows how to use `if` statements with `cond()` to manage variables based on conditions:

`vars({MVAR=""}); if CAP2~=nil then MVAR=MVAR..CAP2 end; cond(string.sub(CAP1,1,1)~="#", MVAR); if CAP2~=nil then MVAR=string.sub(CAP1,4,-1) end`

<br>

### DEBUG option

The `DEBUG` option lets you inspect global variables during replacements. When enabled, it opens a message box displaying the current values of all global variables for each replacement hit, requiring confirmation to proceed to the next match. Initialize the `DEBUG` option in your replacement string to enable it.

| Find      | Replace               |
|------------|----------------------|
| *(empty)*  | `vars({DEBUG=true})`|

<br>

### More Examples

| Find             | Replace                                                                                                     | Regex | Scope CSV | Description                                                                                     |
|-------------------|-----------------------------------------------------------------------------------------------------------|-------|-----------|-------------------------------------------------------------------------------------------------|
| `;`              | `cond(LCNT==5,";Column5;")`                                                                               | No    | No        | Adds a 5th Column for each line into a `;` delimited file.                                      |
| `key`            | `set("key"..CNT)`                                                                                         | No    | No        | Enumerates key values by appending the count of detected strings. E.g., key1, key2, key3, etc.  |
| `(\d+)`          | `set(CAP1.."€ The VAT is: ".. (CAP1 * 0.15).."€ Total with VAT: ".. (CAP1 + (CAP1 * 0.15)).."€")`          | Yes   | No        | Finds a number and calculates the VAT at 15%, then displays the original amount, the VAT, and the total amount. E.g., `50` becomes `50€ The VAT is: 7.5€ Total with VAT: 57.5€`. |
| `---`            | `cond(COL==1 and LINE<3, "0-2", cond(COL==2 and LINE>2 and LINE<5, "3-4", cond(COL==3 and LINE>=5 and LINE<10, "5-9", cond(COL==4 and LINE>=10, "10+"))))` | No    | Yes       | Replaces `---` with a specific range based on the `COL` and `LINE` values. E.g., `3-4` in column 2 of lines 3-4, and `5-9` in column 3 of lines 5-9 assuming `---` is found in all lines and columns. |
| `(\d+)\.(\d+)\.(\d+)` | `cond(CAP1 > 0 and CAP2 == 0 and CAP3 == 0, MATCH, cond(CAP2 > 0 and CAP3 == 0, " " .. MATCH, " " .. MATCH))` | Yes   | No        | Alters the spacing based on the hierarchy of the version numbers, aligning lower hierarchies with spaces as needed. E.g., `1.0.0` remains `1.0.0`, `1.2.0` becomes ` 1.2.0`, indicating a second-level version change. |
| `(\d+)`          | `set(CAP1 * 2)`                                                                                           | Yes   | No        | Doubles the matched number. E.g., `100` becomes `200`.                                          |
| `;`              | `cond(LCNT == 1, string.rep(" ", 20- (LPOS))..";")`                                                       | No    | No        | Inserts spaces before the semicolon to align it to the 20th character position if it's the first occurrence. |
| `-`              | `cond(LINE == math.floor(10.5 + 6.25 * math.sin((2 * math.pi * LPOS) / 50)), "*", " ")`                    | No    | No        | Draws a sine wave across a canvas of '-' characters spanning at least 20 lines and 80 characters per line. |
| `^(.*)$`         | `vars({MATCH_PREV=1}); cond(MATCH == MATCH_PREV, ''); MATCH_PREV=MATCH;`                                   | Yes   | No        | Removes duplicate lines, keeping the first occurrence of each line. Matches an entire line and uses `MATCH_PREV` to identify and remove consecutive duplicates. |

#### Engine Overview
MultiReplace uses the [Lua engine](https://www.lua.org/), allowing for Lua math operations and string methods. Refer to [Lua String Manipulation](https://www.lua.org/manual/5.4/manual.html#6.4) and [Lua Mathematical Functions](https://www.lua.org/manual/5.4/manual.html#6.6) for more information.

<br>

### User Interaction and List Management
Manage search and replace strings within the list using the context menu, which provides comprehensive functionalities accessible by right-clicking on an entry, using direct keyboard shortcuts, or mouse interactions. Here are the detailed actions available:

#### Context Menu and Keyboard Shortcuts
Right-click on any entry in the list or use the corresponding keyboard shortcuts to access these options:

| Menu Item                | Shortcut      | Description                                     |
|--------------------------|---------------|-------------------------------------------------|
| Undo                     | Ctrl+Z        | Reverts the last change made to the list, including sorting and moving rows. |
| Redo                     | Ctrl+Y        | Reapplies the last action that was undone, restoring previous changes. |
| Transfer to Input Fields | Alt+Up        | Transfers the selected entry to the input fields for editing.|
| Search in List           | Ctrl+F        | Initiates a search within the list entries. Inputs are entered in the "Find what" and "Replace with" fields.|
| Cut                      | Ctrl+X        | Cuts the selected entry to the clipboard.       |
| Copy                     | Ctrl+C        | Copies the selected entry to the clipboard.     |
| Paste                    | Ctrl+V        | Pastes content from the clipboard into the list.|
| Edit Field               |               | Opens the selected list entry for direct editing.|
| Delete                   | Del           | Removes the selected entry from the list.       |
| Select All               | Ctrl+A        | Selects all entries in the list.                |
| Enable                   | Alt+E         | Enables the selected entries, making them active for operations. |
| Disable                  | Alt+D         | Disables the selected entries to prevent them from being included in operations. |

**Note on the 'Edit Field' option:**
The edit field supports multiple lines, preserving text with line breaks. This simplifies inserting and managing complex, structured 'Use Variables' statements.

Additional Interactions:
- **Space Key**: Toggles the activation state of selected entries, similar to using Alt+A to enable or Alt+D to disable.
- **Double-Click**: Double-clicking on a list entry allows direct in-place editing. This behavior can be adjusted via the [`DoubleClickEdits`](#configuration-settings) parameter.

### List Columns
- **Find**: The text or pattern to search for.
- **Replace**: The text or pattern to replace with.
- **Options Columns**:
  - **W**: Match whole word only.
  - **C**: Match case.
  - **V**: Use Variables.
  - **E**: Extended search mode.
  - **R**: Regular expression mode.
- **Additional Columns**:
  - **Find Count**: Displays the number of times each 'Find what' string is detected.
  - **Replace Count**: Shows the number of replacements made for each 'Replace with' string.
  - **Comments**: Add custom comments to entries for annotations or additional context.
  - **Delete**: Contains a delete button for each entry, allowing quick removal from the list.

You can manage the visibility of the additional columns via the **Header Column Menu** by right-clicking on the header row.

### List Toggling
- "Use List" checkbox toggles operation application between all list entries or the "Find what:" and "Replace with:" fields.

### Entry Interaction and Limits
- **Manage Entries**: Manage search and replace strings in a list, and enable or disable entries for replacement, highlighting or searching within the list.
- **Highlighting**: Highlight multiple find words in unique colors for better visual distinction, with over 20 distinct colors available.
- **Character Limit**: Field limits of 4096 characters for "Find what:" and "Replace with:" fields.

<br>

## Data Handling

### List Saving and Loading
-   **Save List** and **Load List** buttons allow you to store and reload your current list of search and replace entries (including all options and states) as a CSV file.
-   List files can also be loaded by dragging and dropping them onto the list area.
-   The CSV files follow RFC 4180 standards for compatibility with other tools.
-   Saved lists make it easy to reuse and share search/replace workflows across sessions and projects.
 
### Bash Script Export (optional)
-   You can export your search and replace list as a standalone bash script for use outside of Notepad++.
-   **Note:** This feature is disabled by default for most users. To enable it, set `BashExport=1` in the [`INI file`](#configuration-settings).
-   List entries using the "Use Variables" option are skipped, as variables are not supported in bash scripts.
-   The bash export does not support the value `\0` in Extended mode.
-   The script aims to reproduce the replacement logic as closely as possible, but some limitations exist due to differences in scripting environments.

<br>

## UI and Behavior Settings

### Column Locking

Lock column widths to prevent resizing during window adjustments. This is useful for key columns like **Find**, **Replace**, and **Comments**.
- **How to Lock**: Double-click the column divider in the header to toggle locking. A lock icon appears in the header for locked columns.
- **Effect**: Locked columns keep their width fixed, while unlocked ones adjust dynamically.

### Configuration Settings

The MultiReplace plugin provides several configuration options, including transparency, scaling, and behavior settings, that can be adjusted via the INI file located at:
`C:\Users\<Username>\AppData\Roaming\Notepad++\plugins\Config\MultiReplace.ini`

#### INI File Settings:

- **ScaleFactor**: Controls the scaling of the plugin window and UI elements.
  - **Default**: `ScaleFactor=1.0` (normal size).
  - **Range**: 0.5 to 2.0.
  - **Description**: Adjust this value to resize the plugin window and UI elements. A lower value shrinks the interface, while a higher value enlarges it.

- **Transparency Settings**: Controls the transparency of the plugin window depending on focus.
  - `ForegroundTransparency`: Transparency level when in focus (0-255, default 255).
  - `BackgroundTransparency`: Transparency level when not in focus (0-255, default 190).

- **EditFieldSize**: Configures the size adjustment of the edit field in the list during toggling.
  - **Default**: `EditFieldSize=5` (normal size).
  - **Range**: 2 to 20.
  - **Description**: Sets the factor by which the edit field in the list expands or collapses when toggling its size.

- **HeaderLines**: Specifies the number of top lines to exclude from sorting as headers.
  - **Default**: `HeaderLines=1` (first line is excluded from sorting).
  - **Description**: Set this value to exclude a specific number of lines at the top of the file from being sorted during CSV operations. Useful for preserving header rows in CSV files.
  - **Note**: If set to `0`, no lines are excluded from sorting.

- **HoverText**: Enables or disables the display of full text for truncated list entries when hovering over them.
  - **Default**: `HoverText=1` (enabled).
  - **Description**: When enabled (`1`), hovering over a truncated entry shows its full content in a pop-up. Set to `0` to disable this functionality.

- **Tooltips**: Controls the display of tooltips in the UI.
  - **Default**: `Tooltips=1` (enabled).
  - **Description**: To disable tooltips, set `Tooltips=0` in the INI file.

- **DoubleClickEdits**: Controls the behavior of double-clicking on list entries.
  - **Default**: `DoubleClickEdits=1` (enabled).
  - **Description**: When enabled (`1`), double-clicking on a list entry allows direct in-place editing. When disabled (`0`), double-clicking transfers the entry to the input fields for editing.

- **AlertNotFound**: Controls notifications for unsuccessful searches.
  - **Default**: `AlertNotFound=1` (enabled).
  - **Description**: To disable the bell sound for unsuccessful searches, set `AlertNotFound=0` in the INI file.
 
- **ListStatistics**: Controls whether list statistics are displayed below the list.
  - **Default**: `ListStatistics=0` (disabled).
  - **Description**: When enabled (`1`), a compact statistics field appears below the list, showing:
    - **A**: Number of activated entries
    - **L**: Total number of list items
    - **R**: Index of the currently focused row
    - **S**: Number of selected items
   
- **StayAfterReplace**: Controls whether it jumps to the next match after pressing Replace.
  - **Default**: `StayAfterReplace=0` (disabled).
  - **Description**: When enabled (`1`), pressing the **Replace** button replaces the current match without jumping to the next one. When disabled (`0`), it automatically jumps to the next match after replacing.
 
- **GroupResults**: Controls how 'Find All' search results are displayed.
  - **Default**: `GroupResults=0` (disabled).
  - **Description**: This option changes how 'Find All' results are presented. When enabled (`1`), results are grouped by their source list entry, creating a categorized view. When disabled (`0`), all results are displayed as a single, flat list, sorted by their position in the document, without any categorization.


### Multilingual UI Support

The UI language settings for the MultiReplace plugin can be customized by adjusting the `languages.ini` file located in `C:\Program Files\Notepad++\plugins\MultiReplace\`. These adjustments will ensure that the selected language in Notepad++ is applied within the plugin. 

Contributions to the `languages.ini` file on GitHub are welcome for future versions. Find the file [here](https://github.com/daddel80/notepadpp-multireplace/blob/main/languages.ini).

