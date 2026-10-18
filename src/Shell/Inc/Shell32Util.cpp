#include "pch.h"
#include "Shell32Util.h"

EXTERN_C HRESULT BindToGetFolderAndPidl(REFCLSID rclsid, IShellFolder** ppsf, PIDLIST_ABSOLUTE* ppidl)
{
    if (ppsf)
        *ppsf = nullptr;
    if (!ppidl)
        return E_POINTER;
    *ppidl = nullptr;

    WCHAR szBindPath[47] = L"shell:::";
    if (!StringFromGUID2(rclsid, &szBindPath[8], 39))
        return E_FAIL;

    PIDLIST_ABSOLUTE pidl = nullptr;
    HRESULT hr = SHILCreateFromPath(szBindPath, &pidl, nullptr);
    if (SUCCEEDED(hr) && ppsf)
        hr = SHBindToObject(nullptr, pidl, nullptr, IID_PPV_ARGS(ppsf));

    if (SUCCEEDED(hr))
    {
        *ppidl = pidl;
        pidl = nullptr;
    }
    ILFree(pidl);
    return hr;
}
