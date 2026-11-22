#include "TrayIconLegacy.h"

namespace
{
    struct MonoBitmapInfo
    {
        BITMAPINFOHEADER header;
        RGBQUAD colors[2];
    };

    // CopyIcoCur/CopyBmp in Win2k clres.c sends monochrome bitmaps through
    // BitmapFromDIB with BLACKONWHITE. Preserve both planes of a mono icon:
    // the top half is AND and the bottom half is XOR (including inversion).
    HBITMAP ResizeMask(HDC screen, HBITMAP source, int height)
    {
        BITMAP bitmap = {};
        if (!GetObject(source, sizeof(bitmap), &bitmap) ||
            bitmap.bmWidth <= 0 || bitmap.bmHeight <= 0)
            return NULL;

        const SIZE_T stride = ((SIZE_T(bitmap.bmWidth) + 31) / 32) * 4;
        if (stride > SIZE_T(-1) / SIZE_T(bitmap.bmHeight))
            return NULL;
        const SIZE_T size = stride * SIZE_T(bitmap.bmHeight);
        BYTE localBits[4 * 32];
        BYTE* bits = size <= sizeof(localBits) ? localBits :
            static_cast<BYTE*>(HeapAlloc(GetProcessHeap(), 0, size));
        if (!bits)
            return NULL;

        MonoBitmapInfo info = {};
        info.header.biSize = sizeof(info.header);
        info.header.biWidth = bitmap.bmWidth;
        info.header.biHeight = -bitmap.bmHeight;
        info.header.biPlanes = 1;
        info.header.biBitCount = 1;
        info.colors[1].rgbRed = 255;
        info.colors[1].rgbGreen = 255;
        info.colors[1].rgbBlue = 255;

        HBITMAP result = NULL;
        if (GetDIBits(screen, source, 0, bitmap.bmHeight, bits,
            reinterpret_cast<BITMAPINFO*>(&info), DIB_RGB_COLORS) == bitmap.bmHeight)
        {
            result = CreateBitmap(16, height, 1, 1, NULL);
            if (result)
            {
                bool copied = false;
                if (bitmap.bmWidth == 16 && bitmap.bmHeight == height)
                {
                    copied = SetDIBits(screen, result, 0, height, bits,
                        reinterpret_cast<BITMAPINFO*>(&info), DIB_RGB_COLORS) == height;
                }
                else
                {
                    HDC dc = CreateCompatibleDC(screen);
                    if (dc)
                    {
                        HGDIOBJ previous = SelectObject(dc, result);
                        if (previous && previous != HGDI_ERROR)
                        {
                            SetTextColor(dc, RGB(0, 0, 0));
                            SetBkColor(dc, RGB(255, 255, 255));
                            SetStretchBltMode(dc, BLACKONWHITE);
                            const int lines = StretchDIBits(dc, 0, 0, 16, height,
                                0, 0, bitmap.bmWidth, bitmap.bmHeight, bits,
                                reinterpret_cast<BITMAPINFO*>(&info), DIB_RGB_COLORS, SRCCOPY);
                            copied = lines != 0 && lines != GDI_ERROR;
                            SelectObject(dc, previous);
                        }
                        DeleteDC(dc);
                    }
                }
                if (!copied)
                {
                    DeleteObject(result);
                    result = NULL;
                }
            }
        }
        if (bits != localBits)
            HeapFree(GetProcessHeap(), 0, bits);
        return result;
    }

    HICON ResizeColor(HDC dc, const ICONINFO& source)
    {
        BITMAP bitmap = {};
        if (!GetObject(source.hbmColor, sizeof(bitmap), &bitmap) ||
            bitmap.bmWidth <= 0 || bitmap.bmHeight <= 0)
            return NULL;
        const SIZE_T width = SIZE_T(bitmap.bmWidth);
        const SIZE_T height = SIZE_T(bitmap.bmHeight);
        if (width > SIZE_T(-1) / sizeof(DWORD) / height)
            return NULL;
        const SIZE_T count = width * height;
        DWORD localPixels[16 * 16];
        DWORD* pixels = count <= ARRAYSIZE(localPixels) ? localPixels :
            static_cast<DWORD*>(HeapAlloc(GetProcessHeap(), 0, count * sizeof(DWORD)));
        if (!pixels)
            return NULL;

        BITMAPINFO info = {};
        info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        info.bmiHeader.biWidth = bitmap.bmWidth;
        info.bmiHeader.biHeight = -bitmap.bmHeight;
        info.bmiHeader.biPlanes = 1;
        info.bmiHeader.biBitCount = 32;
        HICON result = NULL;
        if (GetDIBits(dc, source.hbmColor, 0, bitmap.bmHeight, pixels,
            &info, DIB_RGB_COLORS) == bitmap.bmHeight)
        {
            bool hasAlpha = false;
            if (bitmap.bmBitsPixel == 32 && bitmap.bmPlanes == 1)
                for (SIZE_T i = 0; i < count; ++i)
                    if (pixels[i] >> 24)
                    {
                        hasAlpha = true;
                        break;
                    }

            DWORD resized[16 * 16];
            // CreateBitmap consumes WORD-aligned, top-to-bottom mono rows.
            BYTE alphaMask[16 * 2] = {};
            // Win2k CopyBmp converts a native screen bitmap to a 32-bit DIB,
            // bypassing cldib.c's 4/8/24-bit smoothing. COLORONCOLOR samples
            // destination pixel centers in the source. In particular 32->16
            // takes (2*x+1, 2*y+1), not a modern CopyImage filtered average.
            // Use 64-bit coordinates so arbitrary native dimensions cannot
            // overflow the center calculation; no source-size special cases.
            SIZE_T sourceX[16];
            for (unsigned x = 0; x < 16; ++x)
                sourceX[x] = SIZE_T((ULONGLONG(2 * x + 1) * width) / 32);
            for (unsigned y = 0; y < 16; ++y)
            {
                const SIZE_T sourceY = SIZE_T((ULONGLONG(2 * y + 1) * height) / 32);
                for (unsigned x = 0; x < 16; ++x)
                {
                    DWORD pixel = pixels[sourceY * width + sourceX[x]];
                    if (hasAlpha)
                    {
                        const unsigned alpha = pixel >> 24;
                        if (alpha < 128)
                        {
                            alphaMask[y * 2 + x / 8] |= BYTE(0x80 >> (x % 8));
                            pixel = 0;
                        }
                        else if (alpha < 255)
                        {
                            // USER32 stores modern icon RGB premultiplied.
                            DWORD rgb = 0;
                            for (unsigned shift = 0; shift < 24; shift += 8)
                            {
                                const unsigned channel = ((pixel >> shift) & 255) * 255 / alpha;
                                rgb |= (channel > 255 ? 255 : channel) << shift;
                            }
                            pixel = rgb;
                        }
                    }
                    // Alpha-all-zero legacy icons retain their XOR even
                    // beneath an AND bit; only meaningful alpha replaces it.
                    resized[y * 16 + x] = pixel & 0x00ffffff;
                }
            }

            info.bmiHeader.biWidth = 16;
            info.bmiHeader.biHeight = -16;
            info.bmiHeader.biSizeImage = 0;
            HBITMAP color = CreateDIBitmap(dc, &info.bmiHeader, CBM_INIT,
                resized, &info, DIB_RGB_COLORS);
            HBITMAP mask = hasAlpha ? CreateBitmap(16, 16, 1, 1, alphaMask) :
                ResizeMask(dc, source.hbmMask, 16);
            if (color && mask)
            {
                ICONINFO legacy = {};
                legacy.fIcon = TRUE;
                legacy.hbmColor = color;
                legacy.hbmMask = mask;
                result = CreateIconIndirect(&legacy);
            }
            if (color) DeleteObject(color);
            if (mask) DeleteObject(mask);
        }
        if (pixels != localPixels)
            HeapFree(GetProcessHeap(), 0, pixels);
        return result;
    }
}

HICON TrayIconLegacy::ResizeSnapshot(HICON source)
{
    if (!source)
        return NULL;
    ICONINFO info = {};
    if (!GetIconInfo(source, &info))
        return NULL;

    HICON result = NULL;
    HDC dc = GetDC(NULL);
    if (dc)
    {
        if (info.hbmColor)
            result = ResizeColor(dc, info);
        else if (info.hbmMask)
        {
            HBITMAP mask = ResizeMask(dc, info.hbmMask, 32);
            if (mask)
            {
                ICONINFO legacy = {};
                legacy.fIcon = TRUE;
                legacy.hbmMask = mask;
                result = CreateIconIndirect(&legacy);
                DeleteObject(mask);
            }
        }
        ReleaseDC(NULL, dc);
    }
    if (info.hbmColor) DeleteObject(info.hbmColor);
    if (info.hbmMask) DeleteObject(info.hbmMask);
    return result;
}
