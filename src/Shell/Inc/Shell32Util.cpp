#include "pch.h"

#include "Shell32Util.h"

// Thanks to ep_taskbar by @amrsatrio
EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** ppsf, ITEMIDLIST_ABSOLUTE** ppidl)
{
    if (ppsf)
        *ppsf = nullptr;

    *ppidl = nullptr;

    WCHAR szBindPath[47] = L"shell:::";
    StringFromGUID2(rclsid, &szBindPath[8], 39);

    ITEMIDLIST_ABSOLUTE* pidl;
    HRESULT hr = SHILCreateFromPath(szBindPath, &pidl, nullptr);
    if (SUCCEEDED(hr))
    {
        if (ppsf)
        {
            hr = SHBindToObject(nullptr, pidl, nullptr, IID_PPV_ARGS(ppsf));
        }

        if (SUCCEEDED(hr))
        {
            *ppidl = pidl;
            pidl = nullptr;
        }

        ILFree(pidl);
    }

    return hr;
}
