#include "pch.h"

#include "ResourceStringHelpers.h"

#include "Win32ErrorHelpers.h"

HRESULT ResourceStringFindAndSizeEx(HMODULE hInstance, UINT uId, WORD wLanguage, const WCHAR** ppch, WORD* plen)
{
    HRESULT hr;

    HRSRC hRsrc = FindResourceEx(hInstance, RT_STRING, MAKEINTRESOURCE((uId >> 4) + 1), wLanguage);
    if (hRsrc)
    {
        HGLOBAL hgRsrc = LoadResource(hInstance, hRsrc);
        if (hgRsrc)
        {
            WORD* pRsrc = (WORD*)LockResource(hgRsrc);
            if (pRsrc)
            {
                for (UINT i = uId & 0xF; i; --i)
                    pRsrc += *pRsrc + 1;

                if (ppch)
                {
                    WORD len = *pRsrc;
                    *ppch = len ? (const WCHAR*)pRsrc + 1 : nullptr;
                    *plen = len;
                }
                else if (plen)
                {
                    *plen = *pRsrc;
                }

                hr = S_OK;
            }
            else
            {
                hr = E_FAIL;
            }
        }
        else
        {
            hr = HRESULTFromLastErrorError();
        }
    }
    else
    {
        hr = HRESULTFromLastErrorError();
    }

    return hr;
}

template <typename T>
HRESULT TResourceStringAllocCopyEx(
    HMODULE hInstance,
    UINT uId,
    WORD wLanguage,
    HRESULT (CALLBACK *pfnAlloc)(HANDLE, SIZE_T, T*),
    HANDLE hHeap,
    T* pt)
{
    *pt = nullptr;

    const WCHAR* rgch;
    WORD len;
    HRESULT hr = ResourceStringFindAndSizeEx(hInstance, uId, wLanguage, &rgch, &len);
    if (SUCCEEDED(hr))
    {
        SIZE_T elemSize = sizeof(*pt);

        SIZE_T cb = elemSize * len;
        T t;
        hr = pfnAlloc(hHeap, cb + elemSize, &t);
        if (SUCCEEDED(hr))
        {
            memcpy(t, rgch, cb);
            t[cb / elemSize] = 0;
            *pt = t;
        }
    }

    return hr;
}

template <typename T>
HRESULT TCoTaskMemAllocCb(SIZE_T cb, T **out)
{
    T *p = (T *)CoTaskMemAlloc(cb);
    HRESULT hr = p ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        *out = p;
    }
    return S_OK;
}

HRESULT CALLBACK _ResourceStringAllocCopyExCoAlloc(HANDLE hHeap, SIZE_T cb, WCHAR **ppsz)
{
    return TCoTaskMemAllocCb(cb, ppsz);
}

HRESULT ResourceStringCoAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR **ppsz)
{
    return TResourceStringAllocCopyEx(hModule, uId, wLanguage, _ResourceStringAllocCopyExCoAlloc, nullptr, ppsz);
}

HRESULT ResourceStringCoAllocCopy(HINSTANCE hModule, UINT uId, WCHAR** ppsz)
{
    return ResourceStringCoAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

int SysAllocCb(SIZE_T cb, WCHAR** ppsz)
{
    if (cb < sizeof(WCHAR))
        return E_INVALIDARG;
    WCHAR* psz = SysAllocStringByteLen(nullptr, (UINT)(cb - sizeof(WCHAR)));
    HRESULT hr = psz ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        *ppsz = psz;
    }
    return hr;
}

HRESULT CALLBACK ResourceStringAllocCopyExSysAlloc(HANDLE hHeap, SIZE_T cb, WCHAR** ppsz)
{
    return SysAllocCb(cb, ppsz);
}

HRESULT ResourceStringSysAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR** ppsz)
{
    return TResourceStringAllocCopyEx(hModule, uId, wLanguage, ResourceStringAllocCopyExSysAlloc, nullptr, ppsz);
}

HRESULT ResourceStringSysAllocCopy(HINSTANCE hModule, UINT uId, WCHAR** ppsz)
{
    return ResourceStringSysAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

template <typename T>
HRESULT TLocalAllocArrayEx(UINT uFlags, SIZE_T uBytes, T** out)
{
    T* p = (T*)LocalAlloc(uFlags, uBytes);
    HRESULT hr = p ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        *out = p;
    }
    return hr;
}

HRESULT CALLBACK _ResourceStringAllocCopyExLocalAlloc(HANDLE hHeap, SIZE_T cb, WCHAR** ppsz)
{
    return TLocalAllocArrayEx(0, cb, (BYTE**)ppsz);
}

HRESULT ResourceStringLocalAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR** ppsz)
{
    return TResourceStringAllocCopyEx(hModule, uId, wLanguage, _ResourceStringAllocCopyExLocalAlloc, nullptr, ppsz);
}

HRESULT ResourceStringLocalAllocCopy(HINSTANCE hModule, UINT uId, WCHAR** ppsz)
{
    return ResourceStringLocalAllocCopyEx(hModule, uId, LANG_NEUTRAL, ppsz);
}

HRESULT ResourceStringCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR* ppsz, UINT cch)
{
    const WCHAR* pch;
    WORD len;
    HRESULT hr = ResourceStringFindAndSizeEx(hModule, uId, wLanguage, &pch, &len);
    if (SUCCEEDED(hr))
    {
        if (len < cch)
        {
            memcpy(ppsz, pch, sizeof(WCHAR) * len);
            ppsz[len] = 0;
        }
        else
        {
            hr = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
        }
    }
    return hr;
}

HRESULT ResourceStringCchCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR* ppsz, UINT cch)
{
    return ResourceStringCopyEx(hModule, uId, wLanguage, ppsz, cch);
}
