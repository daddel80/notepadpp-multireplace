// This file is part of the MultiReplace plugin for Notepad++.
// Copyright (C) 2023 Thomas Knoefel
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

#ifndef IMAGE_DATA_H
#define IMAGE_DATA_H

#include <windows.h>
#include <cstdint>

struct ImageData {
    unsigned int width;
    unsigned int height;
    unsigned int bytes_per_pixel;
    const unsigned char* pixel_data;
};

// Toolbar icons (colored: gray/white lines + blue arrow)
extern const ImageData gimp_image;
extern const ImageData gimp_image_light;
extern const ImageData gimp_image_dark;

// Tab icons (monochrome)
extern const ImageData gimp_image_tab_light;
extern const ImageData gimp_image_tab_dark;

// Icon creation functions
void CalculateIconSize(UINT dpi, int& iconWidth, int& iconHeight);
HBITMAP CreateBitmapFromImageData(const ImageData& img, UINT dpi);
HICON CreateIconFromImageData(const ImageData& img, UINT dpi);

#endif // IMAGE_DATA_H