#pragma once
#include <windows.h>

// This cache is produced before shell startup. Runtime loading never downloads
// symbols or accepts an RVA from a different executable/PDB identity.
struct ExplorerImageIdentity
{
    GUID pdbGuid;
    DWORD pdbAge;
    DWORD timeDateStamp;
    DWORD sizeOfImage;
    BYTE sha256[32];
};

struct ExplorerHookCache
{
    DWORD magic;
    DWORD version;
    ExplorerImageIdentity image;
    DWORD functionRva;
    BYTE prologue[16];
};

constexpr DWORD ExplorerHookCacheMagic = 0x58455844; // DXEX
constexpr DWORD ExplorerHookCacheVersion = 1;
constexpr WCHAR ExplorerHookCacheName[] = L"explorer-hook.bin";
static_assert(sizeof(ExplorerHookCache) == 88, "Unexpected cache layout");

HRESULT GetExplorerExDirectory(WCHAR (&path)[MAX_PATH]);
HRESULT ReadExplorerImageIdentity(PCWSTR path, ExplorerImageIdentity* identity,
                                  WCHAR (&pdbName)[MAX_PATH]);
HRESULT ReadExplorerHookCache(PCWSTR directory, ExplorerHookCache* cache);
