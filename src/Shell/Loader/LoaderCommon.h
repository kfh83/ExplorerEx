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

// Exact-image cache for the native shell32 multicolumn menu sizing fix.
struct Shell32HookCache
{
    DWORD magic;
    DWORD version;
    ExplorerImageIdentity image;
    DWORD targetRva;      // CMenuSFToolbar::GetSize (hooked; prologue verified)
    DWORD replacementRva; // CMenuToolbarBase::GetSize (called instead)
    BYTE prologue[16];    // original bytes at targetRva
};

constexpr DWORD Shell32HookCacheVersion = 3;
constexpr WCHAR Shell32HookCacheName[] = L"shell32-hook.bin";
static_assert(sizeof(Shell32HookCache) == 92, "Unexpected cache layout");

HRESULT GetExplorerExDirectory(WCHAR (&path)[MAX_PATH]);
// dll=false requires an executable image, dll=true a DLL.
HRESULT ReadExplorerImageIdentity(PCWSTR path, ExplorerImageIdentity* identity,
                                  WCHAR (&pdbName)[MAX_PATH], bool dll = false);
HRESULT ReadExplorerHookCache(PCWSTR directory, ExplorerHookCache* cache);
HRESULT ReadShell32HookCache(PCWSTR directory, Shell32HookCache* cache);
