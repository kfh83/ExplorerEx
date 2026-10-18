#include "pch.h"
#include "ShUndoc.h"

EXTERN_C BOOL WINAPI ExplorerExShell32_EnsureRuntimeInit(void)
{
    // Initialize private shell imports outside the loader lock, once per module.
    static const BOOL initialized = SHUndocInit();
    return initialized;
}

STDAPI_(BOOL) DllMain(HINSTANCE hDll, DWORD dwReason, void* lpReserved)
{
    return TRUE;
}