# MultiReplace for Notepad++
Multi-string replacement plugin for Notepad++

<img src="./NppMultiReplace.gif" alt="MultiReplace Screenshot" width="540" height="440">

## Technical Features

### String Handling and Visualization
- Create and store search and replace strings in a list.
- Field limits of 4096 characters for "Find what:" and "Replace with:".
- Uses the djb2 hash algorithm to ensure consistent color marking of unique words of a Findstring.
- Supports over 20 distinct colors for marking.

### List Interaction
- **Add into List Button**: Adds field contents along with their options into the list.
- **Alt + Up Arrow / Double-Click**: Instantly transfer a row's contents with their options to fields.
- **Delete Button (X) / Delete Key**: Select and delete rows.

### Data Import/Export
- Supports import/export of search and replace strings with their options in CSV format.
- Adherence to RFC 4180 standards for CSV, enabling compatibility and easy interaction with other CSV handling tools.
- Enables reuse of search and replace operations across sessions and projects.

### Bash Script Export
- Exports Find and Replace strings into a runnable script.
- Strives to encapsulate the full functionality of the plugin in the script. However, due to differences in tooling, complete compatibility cannot be guaranteed.
- Intentionally does not support the value `\0` in the Extended Option to avoid escalating environment tooling requirements.

### Function Toggling
- "Use List" checkbox toggles operation application between all list entries or the "Find what:" and "Replace with:" fields.

### Compatibility
- Full compatibility with standard Notepad++ search parameters, including "Match whole word only", "Match case", and "Search Mode".

## Usage
MultiReplace enhances Notepad++ functionalities, enabling:
- Multiple replacements in a single operation.
- CSV-based string list storage.
- Scripted text replacements through bash script export.
