#include "pch.h"

STDAPI_(BOOL) DllMain(HANDLE hDll, DWORD dwReason, void* lpReserved)
{
    switch (dwReason)
    {
        case DLL_PROCESS_ATTACH:
        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
        case DLL_PROCESS_DETACH:
            break;
    }
    return TRUE;
}