#pragma once

inline HRESULT ResultFromLastError()
{
    return HRESULT_FROM_WIN32(GetLastError());
}

inline DWORD GetLastErrorError()
{
    DWORD result = GetLastError();
    return result == ERROR_SUCCESS ? 1 : result;
}

inline HRESULT HRESULTFromLastErrorError()
{
    DWORD error = GetLastError();
    if (error != ERROR_SUCCESS && (int)error <= 0)
        return (HRESULT)GetLastErrorError();
    else
        return (HRESULT)((GetLastErrorError() & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000);
}
