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

// In-process fixes for Windows' shell32.dll, one cache per exact shell32 image:
//  - multicolumn menu fix: CMenuSFToolbar::GetSize is redirected to
//    CMenuToolbarBase::GetSize (symbols exist in every supported build);
//  - pin gate: wil FeatureImpl<Feature_W32PTP>::__private_IsEnabled is forced
//    to false so CPinnedList keeps its classic add-pin path. Windows 10 builds
//    from 19045.7184 gate adding a pin behind that feature; earlier builds do
//    not have the symbol, so pinGateRva == 0 means "not present, nothing to do".
struct Shell32HookCache
{
    DWORD magic;
    DWORD version;
    ExplorerImageIdentity image;
    DWORD targetRva;      // CMenuSFToolbar::GetSize (hooked; prologue verified)
    DWORD replacementRva; // CMenuToolbarBase::GetSize (called instead)
    BYTE prologue[16];    // original bytes at targetRva
    DWORD pinGateRva;     // __private_IsEnabled for Feature_W32PTP, or 0
    BYTE pinGatePrologue[16];
};

constexpr DWORD Shell32HookCacheVersion = 2;
constexpr WCHAR Shell32HookCacheName[] = L"shell32-hook.bin";
static_assert(sizeof(Shell32HookCache) == 112, "Unexpected cache layout");

HRESULT GetExplorerExDirectory(WCHAR (&path)[MAX_PATH]);
// dll=false requires an executable image, dll=true a DLL.
HRESULT ReadExplorerImageIdentity(PCWSTR path, ExplorerImageIdentity* identity,
                                  WCHAR (&pdbName)[MAX_PATH], bool dll = false);
HRESULT ReadExplorerHookCache(PCWSTR directory, ExplorerHookCache* cache);
HRESULT ReadShell32HookCache(PCWSTR directory, Shell32HookCache* cache);
