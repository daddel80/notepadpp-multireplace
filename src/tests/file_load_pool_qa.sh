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

# Builds and runs the QA for loading ahead on worker threads (file_load_pool_qa.cpp): on Linux the pool
# against the simulated loader under ThreadSanitizer and under AddressSanitizer with
# UndefinedBehaviorSanitizer, then the Windows program with MinGW-w64 (run it on Windows or with Wine).
#
# Needs bash, g++ and x86_64-w64-mingw32-g++ (posix threads) with C++17.
# Usage: file_load_pool_qa.sh [work dir]
#   work dir  programs and stubs; default ${TMPDIR:-/tmp}/file_load_pool_qa
# Result: <work dir>/file_load_pool_qa.exe [--dir PATH]...
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
src="$here/.."
work="${1:-${TMPDIR:-/tmp}/file_load_pool_qa}"
mkdir -p "$work/linux" "$work/winstub"

# Encoding.h includes <windows.h>; on Linux the pool needs only these two from it
printf '#pragma once\ntypedef unsigned int UINT;\n#define CP_ACP 0\n' > "$work/linux/windows.h"

for san in thread address,undefined; do
    exe="$work/selftest_${san%%,*}"
    g++ -std=c++17 -O1 -g -Wall -Wextra -fsanitize="$san" -pthread -I"$work/linux" -I"$src" \
        "$here/file_load_pool_qa.cpp" "$src/FileLoadPool.cpp" -o "$exe"
    "$exe"
done

# The plugin sources include <Windows.h> and similar; MinGW's headers are lower case
for h in Windows.h CommCtrl.h Commctrl.h ShlObj.h Shlwapi.h; do
    [ -e "$work/winstub/$h" ] || ln -s "$(echo '#include <windows.h>' | x86_64-w64-mingw32-g++ -E -x c++ - | grep -m1 -o '"/[^"]*/windows.h"' | tr -d '"' | xargs dirname)/$(echo "$h" | tr 'A-Z' 'a-z')" "$work/winstub/$h"
done

x86_64-w64-mingw32-g++ -std=c++17 -O2 -municode -DUNICODE -D_UNICODE -DNOMINMAX -Wall -Wextra -I"$work/winstub" -I"$src" \
    "$here/file_load_pool_qa.cpp" "$src/FileLoadPool.cpp" "$src/FileLoad.cpp" "$src/Encoding.cpp" "$src/DirectoryWalk.cpp" \
    "$src/StringUtils.cpp" -static -lshlwapi -ladvapi32 -o "$work/file_load_pool_qa.exe"
echo "built $work/file_load_pool_qa.exe"
