// This file is part of Notepad++ project
// Copyright (C)2023 Thomas Knoefel
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// at your option any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

#ifndef IMAGE_DATA_H
#define IMAGE_DATA_H

static const struct {
    unsigned int width;
    unsigned int height;
    unsigned int bytes_per_pixel; // 2:RGB16, 3:RGB, 4:RGBA
    unsigned char pixel_data[32 * 32 * 4 + 1];;
} gimp_image = {
  32, 32, 4,
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\361\000\000n\362\000\000\252\362\000\000\252"
  "\362\000\000\252\362\000\000\252\362\000\000\252\362\000\000\252\362\000\000\252\362\000\000\252"
  "\362\000\000\252\362\000\000\215\362\000\000O\377\000\000\001\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\366\000\000\035\362\000\000\377\362\000\000"
  "\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000"
  "\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\265\355\000\000"
  "\016\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377\000"
  "\000\005\362\000\000\305\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000"
  "\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000"
  "\377\362\000\000\377\362\000\000\232\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377\000\000\002\361\000\000\065\363\000\000\341\362\000\000\377"
  "\362\000\000\376\363\000\000\026\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\361\000\000n\362\000\000\377\362\000\000\377\361\000\000H"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\003\276\000\255\003\274\000\234\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\003\277"
  "\000\307\003\276\000\377\003\276\000\377\005\277\000\243\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\004\275\000\210\003\276\000"
  "\377\003\276\000\377\003\276\000\377\003\276\000\376\002\275\000m\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\004\277\000\203\003\276\000\377\003\276"
  "\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\376\002\276\000j\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\277\000\177\003\276\000\377\003\276\000"
  "\377\003\276\000\375\003\276\000\377\003\276\000\377\003\276\000\376\003\276\000\377\003\276\000"
  "\376\002\275\000h\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275\000x\003\276\000\377\003\276"
  "\000\377\003\276\000\373\003\276\000\234\003\276\000\377\003\276\000\377\003\275\000\232\003\276"
  "\000\377\003\276\000\377\003\276\000\375\003\277\000_\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\361\000\000Y\362\000\000\377\362\000"
  "\000\377\361\000\000]\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\277\000\024\003\276"
  "\000\375\003\276\000\377\003\276\000\374\003\276\000Z\002\275\000i\003\276\000\377\003\276\000\377"
  "\003\275\000M\002\277\000s\003\276\000\377\003\276\000\377\003\276\000\364\000\377\000\001\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\363\000\000)\361\000\000\\\377\000\000\007\000\000\000\000\361"
  "\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\000\000\000\000\377\000\000\003\362\000\000L\360\000\000"
  "\"\000\000\000\000\000\277\000\004\003\276\000\314\003\276\000\371\003\275\000]\000\000\000\000\002\275\000i\003"
  "\276\000\377\003\276\000\377\003\275\000M\000\000\000\000\002\275\000t\003\276\000\374\003\276\000\265"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377\000\000\005\362\000\000\356\362\000\000\377"
  "\363\000\000\271\377\000\000\010\361\000\000Y\362\000\000\377\362\000\000\377\361\000\000]\377\000\000"
  "\004\361\000\000\251\362\000\000\377\362\000\000\351\377\000\000\003\000\000\000\000\000\177\000\002\000\306"
  "\000\011\000\000\000\000\000\000\000\000\002\275\000i\003\276\000\377\003\276\000\377\003\275\000M\000\000\000\000\000"
  "\000\000\000\000\271\000\013\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377\000\000"
  "\003\363\000\000\341\362\000\000\377\362\000\000\377\362\000\000\265\361\000\000]\362\000\000\377\362"
  "\000\000\377\362\000\000_\363\000\000\245\362\000\000\377\362\000\000\377\362\000\000\334\377\000\000"
  "\001\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275\000i\003\276\000\377\003\276\000\377"
  "\003\275\000M\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\362\000\000'\362\000\000\347\362\000\000\377\362\000\000\377\363\000\000\317"
  "\362\000\000\377\362\000\000\377\362\000\000\277\362\000\000\377\362\000\000\377\362\000\000\350"
  "\363\000\000)\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275\000i\003\276"
  "\000\377\003\276\000\377\003\275\000M\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\364\000\000-\362\000\000\353\362\000\000\377"
  "\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\351"
  "\363\000\000*\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275\000"
  "i\003\276\000\377\003\276\000\377\003\275\000M\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\357\000\000\061\362"
  "\000\000\356\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\377\362\000\000\351\363"
  "\000\000+\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275"
  "\000i\003\276\000\377\003\276\000\377\003\275\000M\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\360"
  "\000\000\064\363\000\000\367\362\000\000\377\362\000\000\377\362\000\000\366\364\000\000.\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\275\000"
  "i\003\276\000\377\003\276\000\377\003\275\000M\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\357\000\000b\362\000\000\377\362\000\000\377\362\000\000e\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\002\277\000g\003\276\000\377"
  "\003\276\000\377\003\276\000N\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377"
  "\000\000\023\377\000\000\023\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\004\277\000D\003\276\000\377\003\276\000\377\004\275"
  "\000\200\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\277\000\020\003\276\000\364\003\276\000\377\003\276\000\370\004\277\000\207"
  "\003\276\000V\003\275\000U\003\275\000U\003\275\000U\003\275\000U\003\275\000U\003\275\000U\003\275\000U"
  "\003\275\000U\000\277\000$\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\002\275\000m\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377"
  "\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377"
  "\003\276\000\377\003\276\000\377\003\276\000\357\000\271\000\013\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\377\000\001\002\275\000h\003\276\000\363\003\276"
  "\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276"
  "\000\377\003\276\000\377\003\276\000\377\003\276\000\377\003\276\000\356\000\271\000\013\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\277\000\010\004\274\000\071\003\275\000U\003\275\000U\003\275\000U\003\275\000U\003\275\000U\003\275"
  "\000U\003\275\000U\003\275\000U\003\275\000U\000\301\000!\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000"
  "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000",
};

#endif // IMAGE_DATA_H