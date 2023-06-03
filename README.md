
# MultiReplace for Notepad++
multi-string replacement plugin for Notepad++

<img src="./NppMultiReplace.gif" alt="MultiReplace Screenshot" width="540" height="440">

## Overview

MultiReplace is a Notepad++ plugin that enhances the existing search and replace features. It enables users to conduct multiple search and replace operations all at once, saving time and increasing productivity.

## Features

### Multiple Search and Replace

MultiReplace introduces a functionality to store search and replace strings in a list. This enhances efficiency when multiple replacements need to be made concurrently. 

### Interaction with the List

- **Upward Arrow**: Clicking on the upward arrow on the right side of a row transfers the contents of the row back into the "Find what:" and "Replace with:" fields for editing and re-insertion.

- **Double-Click/ Space key**: Double-clicking a row will also copy its contents to the "Find what:" and "Replace with:" fields.

- **Delete Button (X)**: The X button deletes the corresponding row. You can also delete a row by selecting it and pressing the "Delete" key. Multiple rows can be selected and deleted simultaneously.

### CSV Import/Export

For more convenient management and reusability, MultiReplace supports storing and loading search and replace strings in CSV format. This enables the reuse of previously defined search and replace operations across different sessions and projects.

### Bash Script Export

MultiReplace offers the option to unload search and replace strings with all associated options into an executable bash script based on sed command.

### Use List Checkbox

The "Use List" checkbox determines the behavior of the Replace and Mark operations. When selected, these operations are applied to all entries in the list. When deselected, the "Find what:" and "Replace with:" fields act as the input for replace or mark operations, similar to the core functionality of Notepad++.

### Compatibility with Notepad++ Search Modes

MultiReplace adheres to the standard Notepad++ search parameters such as "Match whole word only", "Match case", and the "Search Mode" options.

## Usage

The MultiReplace plugin adds an advanced layer to the Notepad++ search and replace functionalities. By enabling multiple replacements in a single operation and allowing for the storage and loading of string lists in CSV format, it provides a more efficient workflow for complex text manipulation tasks. The bash script export feature opens up new possibilities for scripted or automated text replacements.
