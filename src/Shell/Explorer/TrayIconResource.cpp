#include "TrayIconLegacy.h"

namespace
{
    const int IconSize = 16;
    const unsigned FixedOne = 65536;

    struct ResourceDIB
    {
        BITMAPINFOHEADER header;
        RGBQUAD palette[256];
        const BYTE* color;
        const BYTE* mask;
        SIZE_T colorStride;
        unsigned paletteSize;
    };

    bool ReadDIB(const BYTE* bytes, DWORD size, ResourceDIB& dib)
    {
        if (!bytes || size < sizeof(BITMAPINFOHEADER))
            return false;
        CopyMemory(&dib.header, bytes, sizeof(dib.header));
        const BITMAPINFOHEADER& header = dib.header;
        if (header.biSize != sizeof(BITMAPINFOHEADER) ||
            header.biPlanes != 1 || header.biCompression != BI_RGB ||
            (header.biBitCount != 4 && header.biBitCount != 8 && header.biBitCount != 24) ||
            header.biWidth <= 0 || header.biHeight <= 0 || (header.biHeight & 1))
            return false;

        const unsigned maxColors = header.biBitCount == 24 ? 256 : 1u << header.biBitCount;
        dib.paletteSize = header.biClrUsed ? header.biClrUsed :
            (header.biBitCount == 24 ? 0 : maxColors);
        if (dib.paletteSize > maxColors)
            return false;

        // Use wide arithmetic before validating offsets, including on x86.
        const ULONGLONG stride = ((ULONGLONG(header.biWidth) * header.biBitCount + 31) / 32) * 4;
        const ULONGLONG maskStride = ((ULONGLONG(header.biWidth) + 31) / 32) * 4;
        const ULONGLONG height = header.biHeight / 2;
        const ULONGLONG colorOffset = sizeof(BITMAPINFOHEADER) + dib.paletteSize * sizeof(RGBQUAD);
        const ULONGLONG maskOffset = colorOffset + stride * height;
        if (maskOffset + maskStride * height > size)
            return false;

        if (dib.paletteSize)
            CopyMemory(dib.palette, bytes + sizeof(BITMAPINFOHEADER), dib.paletteSize * sizeof(RGBQUAD));
        dib.color = bytes + SIZE_T(colorOffset);
        dib.mask = bytes + SIZE_T(maskOffset);
        dib.colorStride = SIZE_T(stride);
        dib.header.biHeight /= 2;
        dib.header.biSizeImage = DWORD(stride * height);
        // Full palettes make every index valid. Reduced tables need a scan
        // before either our sampler or GDI can dereference an absent entry.
        if (header.biBitCount != 24 && dib.paletteSize < maxColors)
            for (LONG y = 0; y < header.biHeight; ++y)
            {
                const BYTE* row = dib.color + SIZE_T(y) * dib.colorStride;
                for (LONG x = 0; x < header.biWidth; ++x)
                {
                    const unsigned index = header.biBitCount == 8 ? row[x] :
                        ((x & 1) ? row[x / 2] & 15 : row[x / 2] >> 4);
                    if (index >= dib.paletteSize)
                        return false;
                }
            }
        return true;
    }

    bool ReadPixel(const ResourceDIB& dib, unsigned x, unsigned y, DWORD& rgb, unsigned& index)
    {
        const BYTE* row = dib.color + y * dib.colorStride;
        if (dib.header.biBitCount == 24)
        {
            const BYTE* pixel = row + x * 3;
            rgb = DWORD(pixel[0]) | (DWORD(pixel[1]) << 8) | (DWORD(pixel[2]) << 16);
            index = 0;
            return true;
        }
        index = dib.header.biBitCount == 8 ? row[x] :
            ((x & 1) ? row[x / 2] & 15 : row[x / 2] >> 4);
        if (index >= dib.paletteSize)
            return false;
        const RGBQUAD& entry = dib.palette[index];
        rgb = DWORD(entry.rgbBlue) | (DWORD(entry.rgbGreen) << 8) | (DWORD(entry.rgbRed) << 16);
        return true;
    }

    unsigned MixChannel(unsigned c0, unsigned c1, unsigned c2, unsigned c3,
        unsigned x, unsigned y)
    {
        // cldib.c::MixRGB intentionally truncates each horizontal product
        // before applying the vertical weight. A conventional bilinear blend
        // is not pixel-compatible, nor is a pair of rounded interpolations.
        if (!x && !y)
            return c0;
        if (!x)
            return ((FixedOne - y) * c0 + y * c2) / FixedOne;
        if (!y)
            return ((FixedOne - x) * c0 + x * c1) / FixedOne;
        if (x == FixedOne / 2 && y == FixedOne / 2)
            return (c0 + c1 + c2 + c3) / 4;
        return (c0 * (FixedOne - x) / FixedOne * (FixedOne - y) +
                c1 * x / FixedOne * (FixedOne - y) +
                c2 * (FixedOne - x) / FixedOne * y +
                c3 * x / FixedOne * y) / FixedOne;
    }

    bool ScaleDIB(const ResourceDIB& dib, DWORD (&pixels)[IconSize * IconSize])
    {
        const unsigned width = unsigned(dib.header.biWidth);
        const unsigned height = unsigned(dib.header.biHeight);
        const bool halfSize = width == IconSize * 2 && height == IconSize * 2;
        const unsigned start = halfSize ? FixedOne / 2 : 0;
        const unsigned stepX = halfSize ? FixedOne * 2 : (width - 1) * FixedOne / (IconSize - 1);
        const unsigned stepY = halfSize ? FixedOne * 2 : (height - 1) * FixedOne / (IconSize - 1);
        unsigned sy = start;
        // Interpolate in bottom-up DIB storage order, just like Scale424/824.
        // Reversing rows first changes staged truncation for non-2:1 sizes.
        for (unsigned y = 0; y < IconSize; ++y, sy += stepY)
        {
            unsigned sx = start;
            const unsigned y0 = sy / FixedOne;
            const unsigned y1 = y0 + 1 < height ? y0 + 1 : y0;
            for (unsigned x = 0; x < IconSize; ++x, sx += stepX)
            {
                const unsigned x0 = sx / FixedOne;
                const unsigned x1 = x0 + 1 < width ? x0 + 1 : x0;
                DWORD rgb[4];
                unsigned indices[4];
                if (!ReadPixel(dib, x0, y0, rgb[0], indices[0]) ||
                    !ReadPixel(dib, x1, y0, rgb[1], indices[1]) ||
                    !ReadPixel(dib, x0, y1, rgb[2], indices[2]) ||
                    !ReadPixel(dib, x1, y1, rgb[3], indices[3]))
                    return false;
                DWORD result = 0;
                // MixRGBI preserves identical indices even where MixRGB's
                // staged truncation would slightly darken a flat color.
                if (dib.header.biBitCount != 24 && indices[0] == indices[1] &&
                    indices[1] == indices[2] && indices[2] == indices[3])
                    result = rgb[0];
                else
                    for (unsigned shift = 0; shift < 24; shift += 8)
                        result |= MixChannel((rgb[0] >> shift) & 255, (rgb[1] >> shift) & 255,
                            (rgb[2] >> shift) & 255, (rgb[3] >> shift) & 255,
                            sx % FixedOne, sy % FixedOne) << shift;
                pixels[y * IconSize + x] = result;
            }
        }
        return true;
    }

    bool StretchBitmap(HBITMAP target, const BYTE* bits, const BITMAPINFO* info, int mode)
    {
        HDC dc = CreateCompatibleDC(NULL);
        if (!dc)
            return false;
        HGDIOBJ old = SelectObject(dc, target);
        bool success = false;
        if (old && old != HGDI_ERROR)
        {
            SetStretchBltMode(dc, mode);
            SetTextColor(dc, RGB(0, 0, 0));
            SetBkColor(dc, RGB(255, 255, 255));
            success = StretchDIBits(dc, 0, 0, IconSize, IconSize,
                0, 0, info->bmiHeader.biWidth, info->bmiHeader.biHeight,
                bits, info, DIB_RGB_COLORS, SRCCOPY) > 0;
            SelectObject(dc, old);
        }
        DeleteDC(dc);
        return success;
    }

    HICON DecodeModuleResource(HMODULE module, LPCWSTR name)
    {
        HRSRC group = FindResourceW(module, name, MAKEINTRESOURCEW(14));
        if (!group)
            return NULL;
        const DWORD size = SizeofResource(module, group);
        const BYTE* bytes = static_cast<const BYTE*>(LockResource(LoadResource(module, group)));
        if (!bytes || size < 6)
            return NULL;
        WORD directory[3];
        CopyMemory(directory, bytes, sizeof(directory));
        // LookupIconIdFromDirectoryEx has no size argument: validate its
        // complete group directory before allowing it to walk the entries.
        if (directory[0] != 0 || directory[1] != 1 || !directory[2] ||
            directory[2] > (size - 6) / 14)
            return NULL;
        const int id = LookupIconIdFromDirectoryEx(const_cast<BYTE*>(bytes), TRUE,
            IconSize, IconSize, LR_DEFAULTCOLOR);
        if (!id)
            return NULL;
        HRSRC image = FindResourceW(module, MAKEINTRESOURCEW(id), MAKEINTRESOURCEW(3));
        if (!image)
            return NULL;
        const BYTE* imageBytes = static_cast<const BYTE*>(LockResource(LoadResource(module, image)));
        return TrayIconLegacy::DecodeResource(imageBytes, SizeofResource(module, image));
    }

    bool IsAbsoluteModulePath(const WCHAR* path)
    {
        const bool drive = ((path[0] >= L'A' && path[0] <= L'Z') ||
            (path[0] >= L'a' && path[0] <= L'z')) && path[1] == L':' &&
            (path[2] == L'\\' || path[2] == L'/');
        const bool unc = path[0] == L'\\' && path[1] == L'\\' && path[2];
        return drive || unc;
    }
}

HICON TrayIconLegacy::DecodeResource(const BYTE* bytes, DWORD size)
{
    ResourceDIB dib;
    if (!ReadDIB(bytes, size, dib))
        return NULL;

    const bool smooth = (dib.header.biWidth != IconSize || dib.header.biHeight != IconSize) &&
        IconSize >= dib.header.biWidth / 2 && IconSize >= dib.header.biHeight / 2;
    DWORD pixels[IconSize * IconSize];
    if (smooth && !ScaleDIB(dib, pixels))
        return NULL;

    HDC screen = GetDC(NULL);
    if (!screen)
        return NULL;
    HBITMAP color = NULL;
    if (smooth)
    {
        BITMAPINFO output = {};
        output.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        output.bmiHeader.biWidth = IconSize;
        output.bmiHeader.biHeight = IconSize;
        output.bmiHeader.biPlanes = 1;
        output.bmiHeader.biBitCount = 32;
        color = CreateDIBitmap(screen, &output.bmiHeader, CBM_INIT, pixels, &output, DIB_RGB_COLORS);
    }
    else
    {
        color = CreateCompatibleBitmap(screen, IconSize, IconSize);
        // The >2:1 path stays COLORONCOLOR; never delegate resource scaling
        // to modern USER32, whose smart-loading policy differs from Win2k.
        if (color && !StretchBitmap(color, dib.color,
            reinterpret_cast<const BITMAPINFO*>(&dib), COLORONCOLOR))
        {
            DeleteObject(color);
            color = NULL;
        }
    }
    ReleaseDC(NULL, screen);
    if (!color)
        return NULL;

    struct MaskInfo { BITMAPINFOHEADER header; RGBQUAD colors[2]; } maskInfo = {};
    maskInfo.header.biSize = sizeof(BITMAPINFOHEADER);
    maskInfo.header.biWidth = dib.header.biWidth;
    maskInfo.header.biHeight = dib.header.biHeight;
    maskInfo.header.biPlanes = 1;
    maskInfo.header.biBitCount = 1;
    maskInfo.colors[1].rgbRed = maskInfo.colors[1].rgbGreen = maskInfo.colors[1].rgbBlue = 255;
    HBITMAP mask = CreateBitmap(IconSize, IconSize, 1, 1, NULL);
    HICON result = NULL;
    if (mask && StretchBitmap(mask, dib.mask,
        reinterpret_cast<const BITMAPINFO*>(&maskInfo), BLACKONWHITE))
    {
        ICONINFO icon = {};
        icon.fIcon = TRUE;
        icon.hbmColor = color;
        icon.hbmMask = mask;
        result = CreateIconIndirect(&icon);
    }
    if (mask) DeleteObject(mask);
    DeleteObject(color);
    return result;
}

HICON TrayIconLegacy::CopyResource(HICON source)
{
    typedef BOOL (WINAPI* GetIconInfoExProc)(HICON, PICONINFOEXW);
    if (!source)
        return NULL;
    static const GetIconInfoExProc getInfo = reinterpret_cast<GetIconInfoExProc>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetIconInfoExW"));
    if (!getInfo)
        return NULL;
    ICONINFOEXW info = {};
    info.cbSize = sizeof(info);
    const BOOL success = getInfo(source, &info);
    // GetIconInfoExW returns owned bitmap copies, not the icon's bitmaps.
    if (info.hbmColor) DeleteObject(info.hbmColor);
    if (info.hbmMask) DeleteObject(info.hbmMask);
    if (!success || !info.fIcon ||
        info.szModName[ARRAYSIZE(info.szModName) - 2] ||
        info.szResName[ARRAYSIZE(info.szResName) - 2] ||
        !IsAbsoluteModulePath(info.szModName))
        return NULL;
    LPCWSTR name = info.wResID ? MAKEINTRESOURCEW(info.wResID) : info.szResName;
    if (!info.wResID && !info.szResName[0])
        return NULL;

    // A full path avoids DLL search-order guesses. Data-only loading never
    // executes the owner module, and the mapping outlives all resource reads.
    HMODULE module = LoadLibraryExW(info.szModName, NULL,
        LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE);
    if (!module)
        return NULL;
    HICON result = DecodeModuleResource(module, name);
    FreeLibrary(module);
    return result;
}
