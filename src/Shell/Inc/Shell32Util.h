#pragma once

#include <shlobj.h>

EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** ppsf, PIDLIST_ABSOLUTE* ppidl);
