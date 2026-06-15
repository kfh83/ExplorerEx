#pragma once

HRESULT ResourceStringFindAndSizeEx(HMODULE hInstance, UINT uId, WORD wLanguage, const WCHAR** ppch, WORD* plen);

HRESULT ResourceStringCoAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR** ppsz);
HRESULT ResourceStringCoAllocCopy(HMODULE hModule, UINT uId, WCHAR** ppsz);
HRESULT ResourceStringLocalAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR** ppsz);
HRESULT ResourceStringLocalAllocCopy(HMODULE hModule, UINT uId, WCHAR** ppsz);
HRESULT ResourceStringSysAllocCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR** ppsz);
HRESULT ResourceStringSysAllocCopy(HMODULE hModule, UINT uId, WCHAR** ppsz);
HRESULT ResourceStringCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR* ppsz, UINT cch);
HRESULT ResourceStringCchCopyEx(HMODULE hModule, UINT uId, WORD wLanguage, WCHAR* ppsz, UINT cch);
