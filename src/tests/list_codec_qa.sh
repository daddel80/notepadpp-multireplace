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

# Builds and runs the QA for reading pasted list entries (list_codec_qa.cpp) under AddressSanitizer and
# UndefinedBehaviorSanitizer.
#
# Needs bash and g++ with C++20.
# Usage: list_codec_qa.sh [work dir]
#   work dir  program and stub; default ${TMPDIR:-/tmp}/list_codec_qa
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
src="$here/.."
work="${1:-${TMPDIR:-/tmp}/list_codec_qa}"
mkdir -p "$work/linux"

# StringUtils.cpp includes <windows.h>; reading pasted text calls none of it, so every stub aborts
cat > "$work/linux/windows.h" <<'STUB'
#pragma once
#include <cstddef>
#include <cstdlib>
typedef unsigned int UINT;
typedef unsigned long DWORD;
typedef unsigned long LCID;
typedef unsigned long LCTYPE;
#define CP_UTF8 65001
#define LOCALE_USER_DEFAULT 0x0400
#define LOCALE_STHOUSAND 0x0F
inline int MultiByteToWideChar(UINT, DWORD, const char*, int, wchar_t*, int) { std::abort(); }
inline int WideCharToMultiByte(UINT, DWORD, const wchar_t*, int, char*, int, const char*, int*) { std::abort(); }
inline DWORD CharLowerBuffW(wchar_t*, DWORD) { std::abort(); }
inline int GetLocaleInfoW(LCID, LCTYPE, wchar_t*, int) { std::abort(); }
template <std::size_t N> inline int wcscpy_s(wchar_t (&)[N], const wchar_t*) { std::abort(); }
STUB

g++ -std=c++20 -O1 -g -Wall -Wextra -fsanitize=address,undefined -I"$work/linux" -I"$src" \
    "$here/list_codec_qa.cpp" "$src/ListCodec.cpp" "$src/CsvListFormat.cpp" "$src/StringUtils.cpp" -o "$work/list_codec_qa"
"$work/list_codec_qa"
