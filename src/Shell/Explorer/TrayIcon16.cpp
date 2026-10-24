#include "TrayIcon16.h"
#include <string.h>
#include <wchar.h>

namespace
{
    // Local ABI declaration keeps GetIconInfoExW optional on legacy systems.
    struct IconInfoEx
    {
        DWORD cbSize;
        BOOL fIcon;
        DWORD xHotspot;
        DWORD yHotspot;
        HBITMAP hbmMask;
        HBITMAP hbmColor;
        WORD wResID;
        WCHAR szModName[MAX_PATH];
        WCHAR szResName[MAX_PATH];
    };
    typedef BOOL (WINAPI* GetIconInfoExProc)(HICON, IconInfoEx*);

    GetIconInfoExProc GetMetadataProc()
    {
        static GetIconInfoExProc proc = reinterpret_cast<GetIconInfoExProc>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetIconInfoExW"));
        return proc;
    }

    bool Terminated(const WCHAR* text)
    {
        return wmemchr(text, L'\0', MAX_PATH) != nullptr;
    }

    bool AbsoluteModule(const WCHAR* path)
    {
        if (!Terminated(path))
            return false;
        const size_t length = wcslen(path);
        size_t start = 0;
        if (length >= 4 && ((path[0] >= L'A' && path[0] <= L'Z') ||
            (path[0] >= L'a' && path[0] <= L'z')) && path[1] == L':' && path[2] == L'\\')
            start = 3;
        else if (length >= 5 && path[0] == L'\\' && path[1] == L'\\' &&
            path[2] != L'?' && path[2] != L'.')
        {
            // UNC must contain a server, share, and file component.
            const WCHAR* serverEnd = wcschr(path + 2, L'\\');
            if (!serverEnd || serverEnd == path + 2)
                return false;
            const WCHAR* shareEnd = wcschr(serverEnd + 1, L'\\');
            if (!shareEnd || shareEnd == serverEnd + 1 || !shareEnd[1])
                return false;
            start = 2;
        }
        else
            return false; // Includes hexadecimal datafile tokens and device paths.

        size_t component = start;
        for (size_t i = start; i <= length; ++i)
        {
            const WCHAR ch = path[i];
            if (ch == L'\\' || ch == L'\0')
            {
                if (i == component || path[i - 1] == L'.' || path[i - 1] == L' ')
                    return false;
                component = i + 1;
            }
            else if (ch < 32 || ch == L'/' || ch == L':' || ch == L'*' ||
                ch == L'?' || ch == L'"' || ch == L'<' || ch == L'>' || ch == L'|')
                return false;
        }
        return true;
    }

    struct CacheEntry
    {
        TrayIconResource resource;
        int cx;
        int cy;
        HICON icon;
        bool occupied;
    };
    CacheEntry cache[32] = {};
    unsigned nextEntry = 0;

    WORD ReadWord(const BYTE* bytes)
    {
        WORD value;
        memcpy(&value, bytes, sizeof(value));
        return value;
    }

    DWORD ReadDword(const BYTE* bytes)
    {
        DWORD value;
        memcpy(&value, bytes, sizeof(value));
        return value;
    }

    const BYTE* ResourceBytes(HMODULE module, LPCWSTR type, LPCWSTR name, DWORD& size)
    {
        HRSRC resource = FindResourceW(module, name, type);
        if (!resource)
            return nullptr;
        size = SizeofResource(module, resource);
        HGLOBAL loaded = LoadResource(module, resource);
        return loaded ? static_cast<const BYTE*>(LockResource(loaded)) : nullptr;
    }

    bool IsFourBitDib(const BYTE* bytes, DWORD size, int width, int height)
    {
        if (size < sizeof(BITMAPINFOHEADER))
            return false;
        BITMAPINFOHEADER header;
        memcpy(&header, bytes, sizeof(header));
        if ((header.biSize != sizeof(BITMAPINFOHEADER) &&
             header.biSize != sizeof(BITMAPV4HEADER) && header.biSize != sizeof(BITMAPV5HEADER)) ||
            header.biSize > size || header.biWidth != width || header.biHeight != height * 2 ||
            header.biPlanes != 1 || header.biBitCount != 4 || header.biCompression != BI_RGB ||
            header.biClrUsed > 16)
            return false;

        // Dimensions come from group bytes (1..256), so all arithmetic is bounded.
        const DWORD palette = (header.biClrUsed ? header.biClrUsed : 16) * sizeof(RGBQUAD);
        const DWORD xorSize = ((width * 4 + 31) / 32) * 4 * height;
        const DWORD andSize = ((width + 31) / 32) * 4 * height;
        const DWORD needed = header.biSize + palette + xorSize + andSize;
        if (needed > size || (header.biSizeImage != 0 &&
            header.biSizeImage != xorSize && header.biSizeImage != xorSize + andSize))
            return false;
        // Embedded V5 profiles are not part of the supported icon DIB layout.
        if (header.biSize == sizeof(BITMAPV5HEADER) &&
            (ReadDword(bytes + 112) != 0 || ReadDword(bytes + 116) != 0))
            return false;
        return true;
    }

    HICON LoadFourBit(const TrayIconResource& identity, int cx, int cy)
    {
        HMODULE module = LoadLibraryExW(identity.module, nullptr,
            LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE);
        if (!module)
            return nullptr;

        DWORD groupSize = 0;
        const BYTE* group = ResourceBytes(module, MAKEINTRESOURCEW(14),
            identity.id ? MAKEINTRESOURCEW(identity.id) : identity.name, groupSize);
        HICON result = nullptr;
        if (group && groupSize >= 6 && ReadWord(group) == 0 && ReadWord(group + 2) == 1)
        {
            const WORD count = ReadWord(group + 4);
            if (count != 0 && count <= (groupSize - 6) / 14)
            {
                const BYTE* best = nullptr;
                DWORD bestSize = 0;
                unsigned bestDistance = ~0u;
                for (unsigned i = 0; i < count; ++i)
                {
                    const BYTE* entry = group + 6 + i * 14;
                    if (entry[3] != 0 || ReadWord(entry + 4) != 1 ||
                        ReadWord(entry + 6) != 4 || ReadWord(entry + 12) == 0)
                        continue;
                    const int width = entry[0] ? entry[0] : 256;
                    const int height = entry[1] ? entry[1] : 256;
                    const unsigned distance =
                        static_cast<unsigned>(width > cx ? width - cx : cx - width) +
                        static_cast<unsigned>(height > cy ? height - cy : cy - height);
                    if (distance >= bestDistance)
                        continue;
                    DWORD iconSize = 0;
                    const BYTE* icon = ResourceBytes(module, MAKEINTRESOURCEW(3),
                        MAKEINTRESOURCEW(ReadWord(entry + 12)), iconSize);
                    if (!icon || iconSize != ReadDword(entry + 8) ||
                        !IsFourBitDib(icon, iconSize, width, height))
                        continue;
                    best = icon;
                    bestSize = iconSize;
                    bestDistance = distance;
                }
                if (best)
                    result = CreateIconFromResourceEx(const_cast<PBYTE>(best), bestSize,
                        TRUE, 0x00030000, cx, cy, LR_DEFAULTCOLOR);
            }
        }
        FreeLibrary(module);
        return result;
    }
}

HICON CopyTrayIconSource(HICON source, TrayIconResource* resource)
{
    if (!resource)
        return source ? CopyIcon(source) : nullptr;
    ZeroMemory(resource, sizeof(*resource));
    if (!source)
        return nullptr;
    IconInfoEx info = {};
    info.cbSize = sizeof(info);
    GetIconInfoExProc getInfo = GetMetadataProc();
    const BOOL obtained = getInfo && getInfo(source, &info);
    HICON snapshot = nullptr;
    if (obtained && info.fIcon)
    {
        if (AbsoluteModule(info.szModName) && Terminated(info.szResName) &&
            (info.wResID != 0 || info.szResName[0] != L'\0'))
        {
            memcpy(resource->module, info.szModName, sizeof(resource->module));
            resource->id = info.wResID;
            if (!resource->id)
                memcpy(resource->name, info.szResName, sizeof(resource->name));
        }
        ICONINFO pixels = { TRUE, info.xHotspot, info.yHotspot, info.hbmMask, info.hbmColor };
        snapshot = CreateIconIndirect(&pixels);
    }
    if (info.hbmMask)
        DeleteObject(info.hbmMask);
    if (info.hbmColor)
        DeleteObject(info.hbmColor);
    if (snapshot)
        return snapshot;
    ZeroMemory(resource, sizeof(*resource));
    return CopyIcon(source);
}

HICON FindTrayIcon16(const TrayIconResource& resource, int cx, int cy)
{
    if (cx <= 0 || cy <= 0 ||
        !AbsoluteModule(resource.module) || !Terminated(resource.name) ||
        (resource.id == 0 && resource.name[0] == L'\0'))
        return nullptr;
    for (unsigned i = 0; i < ARRAYSIZE(cache); ++i)
    {
        const CacheEntry& entry = cache[i];
        if (entry.occupied && entry.cx == cx && entry.cy == cy &&
            entry.resource.id == resource.id &&
            _wcsicmp(entry.resource.module, resource.module) == 0 &&
            (resource.id != 0 || wcscmp(entry.resource.name, resource.name) == 0))
            return entry.icon; // Includes negative hits, without reopening the module.
    }
    CacheEntry& entry = cache[nextEntry];
    nextEntry = (nextEntry + 1) % ARRAYSIZE(cache);
    if (entry.icon)
        DestroyIcon(entry.icon);
    entry.resource = resource;
    entry.cx = cx;
    entry.cy = cy;
    entry.icon = LoadFourBit(resource, cx, cy);
    entry.occupied = true;
    return entry.icon;
}

void ClearTrayIcon16Cache()
{
    for (unsigned i = 0; i < ARRAYSIZE(cache); ++i)
    {
        if (cache[i].icon)
            DestroyIcon(cache[i].icon);
        cache[i] = {};
    }
    nextEntry = 0;
}
