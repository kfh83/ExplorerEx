#pragma once

#include <windows.h>

struct TrayIconResource
{
    WCHAR module[MAX_PATH];
    WCHAR name[MAX_PATH];
    WORD id;
};

// UI thread only. Captures metadata before copying, since CopyIcon loses it.
// Returns an OWNED pixel snapshot (DestroyIcon); clears resource on unknown identity.
HICON CopyTrayIconSource(HICON source, TrayIconResource* resource);

// UI thread only. Returns a BORROWED cached resource icon, or nullptr.
// Only genuine 4-bit entries in the exact identified RT_GROUP_ICON are used.
// A 32-entry cache retains hits and misses; no owner-process enumeration or
// palette quantization here. Unknown/dynamic/no-variant icons use their original
// pixels; the tray's ILC_COLOR4 destination reduces those when enabled.
// Never destroy the result; it is valid only until the next helper/cache call.
HICON FindTrayIcon16(const TrayIconResource& resource, int cx, int cy);
void ClearTrayIcon16Cache();
