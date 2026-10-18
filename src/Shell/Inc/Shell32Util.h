#pragma once

#include <shlobj.h>

// COM factories and direct cache exports must initialize this DLL's private
// imports before using them. Call outside DllMain/the loader lock.
EXTERN_C BOOL WINAPI ExplorerExShell32_EnsureRuntimeInit(void);

EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** ppsf, PIDLIST_ABSOLUTE* ppidl);
