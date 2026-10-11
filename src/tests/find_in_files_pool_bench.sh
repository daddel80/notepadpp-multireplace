#!/usr/bin/env bash
# This file is part of the MultiReplace plugin for Notepad++.
# Copyright (C) 2026 Thomas Knoefel
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

# Builds the Find in Files thread test (find_in_files_pool_bench.cpp) with MinGW-w64 against the pinned
# Notepad++ sources (Scintilla document and Boost regex bridge) and the plugin's loading code.
# The test puts a simulated USB stick in front of FileLoad::load: FileLoad.cpp is compiled once more under
# the name DiskLoad, and the test's own FileLoad::load calls it.
#
# Needs bash, git, sed and x86_64-w64-mingw32-g++ (posix threads) with C++17, access to github.com.
# Usage: find_in_files_pool_bench.sh [work dir]
#   work dir  sources, objects and the program; default ${TMPDIR:-/tmp}/find_in_files_pool_bench
# Result: <work dir>/find_in_files_pool_bench.exe, run next to the folders source-text and source
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
work="${1:-${TMPDIR:-/tmp}/find_in_files_pool_bench}"

# Notepad++ master containing 20d8f02 "Improve Find in Files performance", as in find_in_files_bench.sh
npp_commit=697f46b6b6d7245408185370b85306a1f03da179

mkdir -p "$work"
cd "$work"

# Sparse, blob-less checkout of the folders needed at a fixed commit
if [ ! -d npp/.git ]; then
    git clone --quiet --filter=blob:none --no-checkout https://github.com/notepad-plus-plus/notepad-plus-plus.git npp
fi
git -C npp sparse-checkout set --cone scintilla/include scintilla/src boostregex
git -C npp checkout --quiet "$npp_commit"

# The plugin sources include <Windows.h> and similar; MinGW's headers are lower case
mkdir -p winstub
for h in Windows.h CommCtrl.h Commctrl.h ShlObj.h Shlwapi.h; do
    [ -e "winstub/$h" ] || ln -s "$(echo '#include <windows.h>' | x86_64-w64-mingw32-g++ -E -x c++ - | grep -m1 -o '"/[^"]*/windows.h"' | tr -d '"' | xargs dirname)/$(echo "$h" | tr 'A-Z' 'a-z')" "winstub/$h"
done

cxx=x86_64-w64-mingw32-g++
flags=(-std=c++17 -O2 -DNDEBUG -DUNICODE -D_UNICODE -DNOMINMAX)
mkdir -p objw gen

# Scintilla document layer with the Notepad++ Boost regex bridge (SCI_OWNREGEX)
for f in CellBuffer Document PerLine RunStyles Decoration CaseFolder CaseConvert CharClassify \
         CharacterCategoryMap CharacterType UniConversion DBCS ChangeHistory UndoHistory RESearch; do
    [ -f "objw/$f.o" ] || $cxx "${flags[@]}" -DSCI_OWNREGEX -Inpp/scintilla/include -Inpp/scintilla/src \
        -c "npp/scintilla/src/$f.cxx" -o "objw/$f.o"
done
for f in BoostRegExSearch UTF8DocumentIterator; do
    [ -f "objw/$f.o" ] || $cxx "${flags[@]}" -DSCI_OWNREGEX -DBOOST_REGEX_STANDALONE -Inpp/scintilla/include \
        -Inpp/scintilla/src -Inpp/boostregex -c "npp/boostregex/$f.cxx" -o "objw/$f.o"
done

# The plugin's loading code as it is in the working tree, FileLoad as DiskLoad
rm -f objw/mr_*.o
sed 's/FileLoad/DiskLoad/g' "$here/../FileLoad.h" > gen/DiskLoad.h
sed 's/FileLoad/DiskLoad/g' "$here/../FileLoad.cpp" > gen/DiskLoad.cpp
for f in Encoding DirectoryWalk FileLoadPool; do
    $cxx "${flags[@]}" -Iwinstub -I"$here/.." -c "$here/../$f.cpp" -o "objw/mr_$f.o"
done
$cxx "${flags[@]}" -Iwinstub -I"$here/.." -c gen/DiskLoad.cpp -o objw/mr_DiskLoad.o

$cxx "${flags[@]}" -municode -Wall -Wextra -Wno-unknown-pragmas -Iwinstub -I"$here/.." -Igen -Inpp/scintilla/include -Inpp/scintilla/src \
    "$here/find_in_files_pool_bench.cpp" objw/*.o -static -lshlwapi -ladvapi32 -o find_in_files_pool_bench.exe
echo "built $work/find_in_files_pool_bench.exe"
