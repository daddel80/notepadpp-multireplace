// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2026 Thomas Knoefel
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program. If not, see <https://www.gnu.org/licenses/>.
//
// DateParse.h
// Parser for date/time strings against a strftime-style format. Used
// by the ExprTk todate() function as the inverse of D[fmt] output.
//
// Supported specifiers:
//   %Y  4-digit year         %F  shortcut for %Y-%m-%d
//   %y  2-digit year         %T  shortcut for %H:%M:%S
//   %m  month (01..12)       %I  hour (01..12)
//   %d  day (01..31)         %p  AM/PM (case-insensitive)
//   %H  hour (00..23)        %%  literal '%'
//   %M  minute (00..59)
//   %S  second (00..60)
//
// Whitespace in the format matches zero or more whitespace chars in
// the input (matches POSIX strptime semantics). Literal characters
// must match exactly.
//
// Two-digit years follow the POSIX rule: 00..68 -> 2000..2068,
// 69..99 -> 1969..1999.

#pragma once

#include <ctime>
#include <string_view>

namespace MultiReplace {

    // Parses `input` against `format`. Writes the consumed fields into
    // `tm`; fields not mentioned in `format` are left untouched, so the
    // caller should zero-init the struct first if a fully-populated
    // value is wanted.
    //
    // Returns true on full success: the entire format was consumed and
    // every field landed inside its valid range. Returns false on any
    // mismatch or out-of-range value; tm contents are undefined after
    // a failed parse.
    bool parseDateTime(std::string_view input,
        std::string_view format,
        std::tm& tm);

}  // namespace MultiReplace