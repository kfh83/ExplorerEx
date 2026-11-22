#pragma once

#include <windows.h>

// Tray-only USER32 resizing compatibility. Returned icons are caller-owned.
// No image-list or application-specific policy belongs in these helpers.
namespace TrayIconLegacy
{
    HICON ResizeSnapshot(HICON source);

    // Decode a classic ICO DIB using Win2k's resource-loading stretch policy.
    // NULL means invalid/unsupported DIB (PNG and modern alpha stay in USER32).
    HICON DecodeResource(const BYTE* bytes, DWORD size);

    // NULL means no supported, identifiable module resource is available.
    HICON CopyResource(HICON source);
}
