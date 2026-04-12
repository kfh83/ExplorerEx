#include "pch.h"

/*
* This will eventually be the host for the following CLSIDs + Implementations.
* CLSID_StartMenu                   - CStartMenu_CreateInstance
* CLSID_PersonalStartMenu           - CPersonalStartMenu_CreateInstance
* CLSID_StartMenuFolder             - CStartMenuFolder_CreateInstance
* CLSID_ProgramsFolder              - CProgramsFolder_CreateInstance
* CLSID_ProgramsFolderAndFastItems  - CProgramsFolderAndFastItems_CreateInstance
*/

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