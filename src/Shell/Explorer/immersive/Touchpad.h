#pragma once

#include <windows.h>

// Three-/four-finger tap dispatch modeled on native CTray::_HandlePTPTap.
// Reads PrecisionTouchPad settings on each tap, including the 0xffff custom
// selector and per-finger KeyParams. Supports disabled, Search, Action Center,
// play/pause, middle/back/forward mouse buttons, and recorded keyboard shortcuts.
// Native Inject*Input is preferred; SendInput is used only when those APIs are
// absent, never after a failed injection. Failures are returned to the tray.
// Swipe recognition remains the immersive shell's responsibility.
HRESULT HandleTouchpadTap(bool fourFinger);
