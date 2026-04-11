#include "pch.h"

#include "Shell32Util.h"

// Thanks to ep_taskbar by @amrsatrio
EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** psfOut, ITEMIDLIST_ABSOLUTE** pidlOut)
{
    if (psfOut)
        *psfOut = nullptr;

    *pidlOut = nullptr;

    WCHAR szPath[47] = L"shell:::";
    StringFromGUID2(rclsid, &szPath[8], 39);

    ITEMIDLIST_ABSOLUTE* pidl;
    HRESULT hr = SHILCreateFromPath(szPath, &pidl, nullptr);
    if (SUCCEEDED(hr))
    {
        if (psfOut)
        {
            hr = SHBindToObject(nullptr, pidl, nullptr, IID_PPV_ARGS(psfOut));
        }

        if (SUCCEEDED(hr))
        {
            *pidlOut = pidl;
            pidl = nullptr;
        }

        ILFree(pidl);
    }

    return hr;
}
