#pragma once

#include <Windows.h>

inline HRESULT ResultFromWin32(__in DWORD dwErr)
{
    return HRESULT_FROM_WIN32(dwErr);
}

inline HRESULT ResultFromLastError()
{
    return HRESULT_FROM_WIN32(GetLastError());
}

inline HRESULT ResultFromKnownLastError()
{
    const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
    return (SUCCEEDED(hr) ? E_FAIL : hr);
}

inline HRESULT ResultFromWin32Bool(BOOL b)
{
    return b ? S_OK : ResultFromKnownLastError();
}

inline HRESULT ResultFromWin32Count(UINT cchResult, UINT cchBuffer)
{
    return cchResult && cchResult <= cchBuffer ? S_OK : ResultFromWin32(ERROR_INSUFFICIENT_BUFFER);
}

inline DWORD GetLastErrorError()
{
    DWORD dwError = GetLastError();
    return dwError == ERROR_SUCCESS ? 1 : dwError;
}

inline HRESULT HRESULTFromLastErrorError()
{
    DWORD error = GetLastError();
    if (error != ERROR_SUCCESS && (int)error <= 0)
        return (HRESULT)GetLastErrorError();
    else
        return (HRESULT)((GetLastErrorError() & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000);
}
