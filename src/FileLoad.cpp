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

#include "FileLoad.h"

#include <windows.h>
#include <winioctl.h>
#include <algorithm>
#include <cstddef>
#include <cstring>             // For std::memchr
#include <new>                 // For std::bad_alloc

namespace FileLoad {

    namespace {

        // A large file is read in blocks of this size, so a stop does not wait for all of it
        constexpr size_t READ_BLOCK = size_t{ 16 } << 20;

        Stamp stampOf(HANDLE h)
        {
            BY_HANDLE_FILE_INFORMATION info{};
            if (!::GetFileInformationByHandle(h, &info)) return {};
            Stamp s;
            s.volume = info.dwVolumeSerialNumber;
            s.fileId = (uint64_t{ info.nFileIndexHigh } << 32) | info.nFileIndexLow;
            s.size = (uint64_t{ info.nFileSizeHigh } << 32) | info.nFileSizeLow;
            s.lastWrite = (uint64_t{ info.ftLastWriteTime.dwHighDateTime } << 32) | info.ftLastWriteTime.dwLowDateTime;
            s.valid = true;
            return s;
        }

        // Read access that leaves the file to others meanwhile (they may write, rename or delete it)
        class InputFile {
        public:
            explicit InputFile(const std::filesystem::path& fp)
                : _h(::CreateFileW(fp.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                    nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr)) {}
            ~InputFile() { if (_h != INVALID_HANDLE_VALUE) ::CloseHandle(_h); }
            InputFile(const InputFile&) = delete;
            InputFile& operator=(const InputFile&) = delete;

            bool isOpen() const { return _h != INVALID_HANDLE_VALUE; }

            bool size(uint64_t& bytes) const {
                LARGE_INTEGER s{};
                if (!::GetFileSizeEx(_h, &s)) return false;
                bytes = static_cast<uint64_t>(s.QuadPart);
                return true;
            }

            Stamp stamp() const { return stampOf(_h); }

            // Up to count bytes, fewer only at the end of the file; false on a read error or a stop
            bool read(char* dst, size_t count, size_t& got, const std::atomic<bool>* stop) {
                got = 0;
                while (got < count) {
                    if (stop && stop->load(std::memory_order_relaxed)) return false;
                    const DWORD chunk = static_cast<DWORD>((std::min)(count - got, READ_BLOCK));
                    DWORD n = 0;
                    if (!::ReadFile(_h, dst + got, chunk, &n, nullptr)) return false;
                    if (n == 0) break;
                    got += n;
                }
                return true;
            }

        private:
            HANDLE _h;
        };

        // Check for BOM (Byte Order Mark) - files with BOM are definitely text
        bool hasBOM(const char* data, size_t len)
        {
            if (len < 2) return false;

            const unsigned char* u = reinterpret_cast<const unsigned char*>(data);

            // UTF-8 BOM: EF BB BF
            if (len >= 3 && u[0] == 0xEF && u[1] == 0xBB && u[2] == 0xBF)
                return true;

            // UTF-16 LE BOM: FF FE
            if (u[0] == 0xFF && u[1] == 0xFE)
                return true;

            // UTF-16 BE BOM: FE FF
            if (u[0] == 0xFE && u[1] == 0xFF)
                return true;

            return false;
        }

        // Check if content is binary by looking for NULL bytes
        // This is the industry standard approach (same as grep)
        bool hasNullBytes(const char* data, size_t len)
        {
            const size_t checkLen = (len < BINARY_CHECK_SIZE) ? len : BINARY_CHECK_SIZE;
            // std::memchr is highly optimized (uses SIMD on modern CPUs)
            return std::memchr(data, '\0', checkLen) != nullptr;
        }

        // UTF-16 LE without BOM: same pre-checks and probe as N++ and Encoding::detectEncoding
        bool isUtf16NoBomLE(const char* data, size_t len)
        {
            const unsigned char* u = reinterpret_cast<const unsigned char*>(data);
            if (!(len > 1 && (len % 2) == 0 && u[0] != 0 && u[1] == 0))
                return false;
            INT uniTest = IS_TEXT_UNICODE_STATISTICS;
            return ::IsTextUnicode(data, static_cast<int>(len), &uniTest) != FALSE;
        }

        // Combined check: returns true if content is binary (no BOM, not UTF-16, NUL bytes)
        bool shouldSkipAsBinary(const char* data, size_t len)
        {
            if (hasBOM(data, len))
                return false;

            if (isUtf16NoBomLE(data, len))
                return false;

            return hasNullBytes(data, len);
        }
    }

    Result load(const std::filesystem::path& fp, const Options& options, std::string& content, const Pace* pace)
    {
        content.clear();
        Result r;
        auto skip = [&](SkipReason reason) { content.clear(); r.skip = reason; return r; };

        try {
            if (options.skipReadOnly) {
                const DWORD attrs = ::GetFileAttributesW(fp.c_str());
                if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) return skip(SkipReason::ReadOnly);
            }

            InputFile in(fp);
            if (!in.isOpen()) return skip(SkipReason::Unreadable);
            if (options.stamp) r.stamp = in.stamp();   // before the read: a later change shows as one
            uint64_t fileSize = r.stamp.size;
            if (!r.stamp.valid && !in.size(fileSize)) return skip(SkipReason::Unreadable);

            if (options.maxFileSize > 0 && fileSize > options.maxFileSize) return skip(SkipReason::TooLarge);
            if (fileSize > content.max_size()) return skip(SkipReason::TooLarge);   // 32-bit build

            // Read header for the binary/encoding decision
            const size_t headerSize = static_cast<size_t>((std::min)(fileSize, uint64_t{ BINARY_CHECK_SIZE }));
            content.resize(headerSize);
            size_t headerLen = 0;
            if (!in.read(content.data(), headerSize, headerLen, nullptr)) return skip(SkipReason::Unreadable);
            content.resize(headerLen);

            const bool binary = shouldSkipAsBinary(content.data(), content.size());
            if (binary && options.skipBinary) return skip(SkipReason::Binary);

            // Text beyond the converters' reach is refused before the full read
            if (!binary && fileSize > Encoding::MAX_CONVERT_LENGTH) return skip(SkipReason::TooLarge);

            if (pace && pace->beforeRead && !pace->beforeRead(fileSize)) return skip(SkipReason::Unreadable);

            // Append remainder (a file that shrank since the size query ends where it ends now)
            if (headerLen == headerSize && fileSize > headerSize) {
                content.resize(static_cast<size_t>(fileSize));
                size_t restLen = 0;
                if (!in.read(content.data() + headerLen, content.size() - headerLen, restLen, pace ? pace->stop : nullptr))
                    return skip(SkipReason::Unreadable);
                content.resize(headerLen + restLen);
            }

            if (binary) {
                // Skip disabled: search the bytes as-is (N++ behavior).
                // This only picks the codepage the hidden buffer is bound
                // to (pattern encoding + dock rendering); the content is
                // never decoded or transformed. isValidUtf8 is a strict
                // structural parser over the FULL content - genuine binary
                // essentially never validates, and a false "yes" could only
                // mis-encode the search pattern, never corrupt buffer or
                // disk (write-back stays verbatim). Deliberately NOT
                // detectEncoding: its CJK/UTF-16 heuristics are tuned for
                // text files and would guess confidently wrong here.
                if (Encoding::isValidUtf8(content.data(), content.size())) {
                    r.enc.kind = Encoding::Kind::UTF8;
                    r.enc.withBOM = false;
                    r.enc.bomBytes = 0;
                }
                r.kind = LoadKind::RawBytes;
                return r;
            }

            r.enc = Encoding::detectEncoding(content.data(), content.size());

            // UTF-8 is searched as it is: decoding would copy the same bytes, and they write back unchanged
            if (r.enc.kind == Encoding::Kind::UTF8) {
                content.erase(0, static_cast<size_t>(r.enc.bomBytes));
                return r;
            }

            std::string u8;
            if (!Encoding::convertBufferToUtf8(content.data(), content.size(), r.enc, u8))
                return skip(SkipReason::Undecodable);

            // Replace path: refuse files whose decode would not write back losslessly
            if (options.verifyRoundtrip && !Encoding::verifyLosslessDecode(content.data(), content.size(), r.enc, u8))
                return skip(SkipReason::Undecodable);

            content.swap(u8);
            return r;
        }
        catch (const std::bad_alloc&) {
            return skip(SkipReason::TooLarge);
        }
        catch (...) {
            return skip(SkipReason::Unreadable);
        }
    }

    bool unchangedSince(const std::filesystem::path& fp, const Stamp& stamp)
    {
        if (!stamp.valid) return false;
        const HANDLE h = ::CreateFileW(fp.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return false;
        const Stamp now = stampOf(h);
        ::CloseHandle(h);
        return now.valid && now.volume == stamp.volume && now.fileId == stamp.fileId
            && now.size == stamp.size && now.lastWrite == stamp.lastWrite;
    }

    DriveKind driveKind(const std::filesystem::path& folder)
    {
        // The root of the folder's volume, or of the path itself when the volume cannot be found
        wchar_t volumePath[MAX_PATH + 1] = {};
        const std::wstring root = ::GetVolumePathNameW(folder.c_str(), volumePath, MAX_PATH + 1)
            ? std::wstring(volumePath) : folder.root_path().wstring();

        switch (::GetDriveTypeW(root.c_str())) {
        case DRIVE_REMOTE:
            return DriveKind::Network;
        case DRIVE_CDROM:
            return DriveKind::Slow;
        case DRIVE_FIXED:
            break;
        default:   // removable flash media (USB sticks, SD cards), RAM disks, drives that cannot be told
            return DriveKind::Fast;
        }

        // A fixed disk, asked through its volume: "\\?\Volume{...}" without the last backslash
        wchar_t volumeName[MAX_PATH + 1] = {};
        if (!::GetVolumeNameForVolumeMountPointW(root.c_str(), volumeName, MAX_PATH + 1)) return DriveKind::Fast;
        std::wstring volume = volumeName;
        if (!volume.empty() && volume.back() == L'\\') volume.pop_back();
        const HANDLE h = ::CreateFileW(volume.c_str(), 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr);
        if (h == INVALID_HANDLE_VALUE) return DriveKind::Fast;

        STORAGE_PROPERTY_QUERY query{};
        query.QueryType = PropertyStandardQuery;
        query.PropertyId = StorageDeviceSeekPenaltyProperty;
        DEVICE_SEEK_PENALTY_DESCRIPTOR penalty{};
        DWORD bytes = 0;
        if (::DeviceIoControl(h, IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), &penalty, sizeof(penalty), &bytes, nullptr)
            && bytes >= sizeof(penalty)) {
            ::CloseHandle(h);
            return penalty.IncursSeekPenalty ? DriveKind::Slow : DriveKind::Fast;
        }

        // No answer: behind a USB or FireWire bridge this may be a hard disk
        query.PropertyId = StorageDeviceProperty;
        alignas(STORAGE_DEVICE_DESCRIPTOR) unsigned char buffer[512] = {};
        const bool described = ::DeviceIoControl(h, IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query), buffer, sizeof(buffer), &bytes, nullptr)
            && bytes >= offsetof(STORAGE_DEVICE_DESCRIPTOR, BusType) + sizeof(STORAGE_BUS_TYPE);
        ::CloseHandle(h);
        if (!described) return DriveKind::Fast;
        switch (reinterpret_cast<const STORAGE_DEVICE_DESCRIPTOR*>(buffer)->BusType) {
        case BusTypeUsb:
        case BusType1394:
            return DriveKind::Slow;
        default:
            return DriveKind::Fast;
        }
    }
}
