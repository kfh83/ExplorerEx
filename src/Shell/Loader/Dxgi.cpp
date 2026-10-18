#include <windows.h>
#include <cstring>
#include <strsafe.h>
#include <KNSoft/SlimDetours/SlimDetours.h>
#include "LoaderCommon.h"

#pragma comment(lib, "KNSoft.SlimDetours.lib")

using CreateDesktopAndTrayFn = HANDLE (__cdecl*)();
using InitializeExplorerExFn = HRESULT (WINAPI*)();

static INIT_ONCE g_realDxgiOnce = INIT_ONCE_STATIC_INIT;
static INIT_ONCE g_bootstrapOnce = INIT_ONCE_STATIC_INIT;
static HMODULE g_realDxgi;
static volatile LONG g_bootstrapThread;
static CreateDesktopAndTrayFn g_originalCreate;
static CreateDesktopAndTrayFn g_replacementCreate;
static InitializeExplorerExFn g_initialize;

static void Report(PCWSTR operation, HRESULT hr)
{
    WCHAR message[256];
    if (SUCCEEDED(StringCchPrintfW(message, ARRAYSIZE(message),
        L"ExplorerEx DXGI loader: %s (0x%08lX)\n", operation, static_cast<ULONG>(hr))))
        OutputDebugStringW(message);
}

static BOOL CALLBACK LoadRealDxgi(PINIT_ONCE, PVOID, PVOID*)
{
    WCHAR path[MAX_PATH];
    UINT length = GetSystemDirectoryW(path, ARRAYSIZE(path));
    if (length && length < ARRAYSIZE(path) &&
        SUCCEEDED(StringCchCatW(path, ARRAYSIZE(path), L"\\dxgi.dll")))
        g_realDxgi = LoadLibraryExW(path, nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    return TRUE;
}

// The assembly forwarding stubs preserve all x64 argument registers and the
// caller's stack, then tail-call the genuine export. No guessed private ABI.
extern "C" void* ResolveDxgiExport(unsigned index)
{
    static const char* const names[] = {
        "ApplyCompatResolutionQuirking", "CompatString", "CompatValue",
        "CreateDXGIFactory", "CreateDXGIFactory1", "CreateDXGIFactory2",
        "DXGID3D10CreateDevice", "DXGID3D10CreateLayeredDevice",
        "DXGID3D10GetLayeredDeviceSize", "DXGID3D10RegisterLayers",
        "DXGIDeclareAdapterRemovalSupport", "DXGIDumpJournal",
        "DXGIGetDebugInterface1", "DXGIReportAdapterConfiguration",
        "PIXBeginCapture", "PIXEndCapture", "PIXGetCaptureState",
        "SetAppCompatStringPointer", "UpdateHMDEmulationStatus"
    };
    static PVOID volatile functions[ARRAYSIZE(names)] = {};
    if (index < ARRAYSIZE(names))
    {
        PVOID proc = InterlockedCompareExchangePointer(&functions[index], nullptr, nullptr);
        if (proc)
            return proc;
        InitOnceExecuteOnce(&g_realDxgiOnce, LoadRealDxgi, nullptr, nullptr);
        if (g_realDxgi)
        {
            proc = reinterpret_cast<PVOID>(GetProcAddress(g_realDxgi, names[index]));
            if (proc)
            {
                InterlockedCompareExchangePointer(&functions[index], proc, nullptr);
                return proc;
            }
        }
    }
    // A missing genuine system export is a loader error, not a fabricated DXGI
    // success. Normal forwarding never depends on ExplorerEx being available.
    RaiseException(0xC0000139, EXCEPTION_NONCONTINUABLE, 0, nullptr);
    TerminateProcess(GetCurrentProcess(), ERROR_PROC_NOT_FOUND);
    __assume(0);
}

static HANDLE __cdecl CreateDesktopAndTrayHook()
{
    // Preserve native Explorer's initialization order: the DXGI bootstrap only
    // installs the redirect. Initialize our shell at the original call site.
    HRESULT hr = g_initialize();
    if (FAILED(hr))
    {
        Report(L"payload initialization failed; retaining native shell", hr);
        return g_originalCreate();
    }
    return g_replacementCreate();
}

static bool IsExecutableRva(HMODULE module, DWORD rva, SIZE_T length)
{
    const BYTE* base = reinterpret_cast<const BYTE*>(module);
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
    if (rva >= nt->OptionalHeader.SizeOfImage || length > nt->OptionalHeader.SizeOfImage - rva)
        return false;
    const IMAGE_SECTION_HEADER* section = IMAGE_FIRST_SECTION(nt);
    for (unsigned i = 0; i < nt->FileHeader.NumberOfSections; ++i, ++section)
    {
        if ((section->Characteristics & IMAGE_SCN_MEM_EXECUTE) && rva >= section->VirtualAddress &&
            rva - section->VirtualAddress < section->Misc.VirtualSize &&
            length <= section->Misc.VirtualSize - (rva - section->VirtualAddress))
            return true;
    }
    return false;
}

using MenuGetSizeFn = void (__cdecl*)(void* self, SIZE* size);
static MenuGetSizeFn g_originalMenuGetSize;
static MenuGetSizeFn g_menuBaseGetSize;

// Windows' CMenuSFToolbar::GetSize expands multicolumn menus forever when the
// theme background has margins; the base-class implementation is correct.
static void __cdecl MenuSFToolbarGetSizeHook(void* self, SIZE* size)
{
    g_menuBaseGetSize(self, size);
}

static HRESULT InstallMenuFix(PCWSTR directory)
{
    Shell32HookCache cache{};
    HRESULT hr = ReadShell32HookCache(directory, &cache);
    if (FAILED(hr))
        return hr;
    // Explorer imports shell32, so it is already mapped here; never load it.
    HMODULE shell32 = GetModuleHandleW(L"shell32.dll");
    if (!shell32)
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    WCHAR path[MAX_PATH];
    DWORD length = GetModuleFileNameW(shell32, path, ARRAYSIZE(path));
    if (!length || length >= ARRAYSIZE(path))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    ExplorerImageIdentity identity{};
    WCHAR pdbName[MAX_PATH];
    hr = ReadExplorerImageIdentity(path, &identity, pdbName, true);
    if (FAILED(hr))
        return hr;
    if (memcmp(&identity, &cache.image, sizeof(identity)) != 0)
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    if (!IsExecutableRva(shell32, cache.targetRva, sizeof(cache.prologue)) ||
        !IsExecutableRva(shell32, cache.replacementRva, sizeof(cache.prologue)))
        return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
    BYTE* target = reinterpret_cast<BYTE*>(shell32) + cache.targetRva;
    if (memcmp(target, cache.prologue, sizeof(cache.prologue)) != 0)
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA); // Already patched or stale cache.
    g_menuBaseGetSize = reinterpret_cast<MenuGetSizeFn>(reinterpret_cast<BYTE*>(shell32) + cache.replacementRva);
    g_originalMenuGetSize = reinterpret_cast<MenuGetSizeFn>(target);
    hr = SlimDetoursInlineHook(TRUE, reinterpret_cast<PVOID*>(&g_originalMenuGetSize),
        reinterpret_cast<PVOID>(&MenuSFToolbarGetSizeHook));
    if (SUCCEEDED(hr))
        Report(L"shell32 multicolumn menu fix installed", hr);
    return hr;
}

static HRESULT InstallExplorerHook()
{
    if (GetSystemMetrics(SM_CLEANBOOT) || GetShellWindow() || FindWindowW(L"Shell_TrayWnd", nullptr))
        return S_FALSE;

    WCHAR directory[MAX_PATH];
    HRESULT hr = GetExplorerExDirectory(directory);
    if (FAILED(hr))
        return hr;
    ExplorerHookCache cache{};
    hr = ReadExplorerHookCache(directory, &cache);
    if (FAILED(hr))
        return hr;

    WCHAR imagePath[MAX_PATH];
    DWORD length = GetModuleFileNameW(nullptr, imagePath, ARRAYSIZE(imagePath));
    if (!length || length >= ARRAYSIZE(imagePath))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    ExplorerImageIdentity identity{};
    WCHAR pdbName[MAX_PATH];
    hr = ReadExplorerImageIdentity(imagePath, &identity, pdbName);
    if (FAILED(hr))
        return hr;
    if (memcmp(&identity, &cache.image, sizeof(identity)) != 0)
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);

    // Exact binary identity, rather than a basename-only check, also permits an
    // unmodified private copy of native Explorer for isolated-desktop testing.
    HMODULE explorer = GetModuleHandleW(nullptr);
    if (!IsExecutableRva(explorer, cache.functionRva, sizeof(cache.prologue)))
        return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
    BYTE* target = reinterpret_cast<BYTE*>(explorer) + cache.functionRva;
    if (memcmp(target, cache.prologue, sizeof(cache.prologue)) != 0)
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA); // Already patched or stale cache.

    WCHAR payloadPath[MAX_PATH];
    hr = StringCchPrintfW(payloadPath, ARRAYSIZE(payloadPath), L"%s\\ExplorerEx.dll", directory);
    if (FAILED(hr))
        return hr;
    HMODULE payload = LoadLibraryExW(payloadPath, nullptr,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    if (!payload)
        return HRESULT_FROM_WIN32(GetLastError());
    g_initialize = reinterpret_cast<InitializeExplorerExFn>(GetProcAddress(payload, "ExplorerExInitialize"));
    g_replacementCreate = reinterpret_cast<CreateDesktopAndTrayFn>(GetProcAddress(payload, "CreateDesktopAndTray"));
    if (!g_initialize || !g_replacementCreate)
    {
        FreeLibrary(payload);
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    // Pin both modules before exposing a branch into them. Shell callbacks and
    // the detour remain valid until process exit; there is no unsafe hot unload.
    HMODULE pinned;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
        reinterpret_cast<PCWSTR>(&CreateDesktopAndTrayHook), &pinned) ||
        !GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
            reinterpret_cast<PCWSTR>(g_replacementCreate), &pinned))
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        FreeLibrary(payload);
        return hr;
    }
    g_originalCreate = reinterpret_cast<CreateDesktopAndTrayFn>(target);
    hr = SlimDetoursInlineHook(TRUE, reinterpret_cast<PVOID*>(&g_originalCreate),
        reinterpret_cast<PVOID>(&CreateDesktopAndTrayHook));
    if (SUCCEEDED(hr))
    {
        Report(L"CreateDesktopAndTray hook installed", hr);
        // Optional and independent: a failure leaves ExplorerEx fully working.
        HRESULT menuFix = InstallMenuFix(directory);
        if (FAILED(menuFix))
            Report(L"shell32 menu fix not installed", menuFix);
    }
    return hr;
}

static BOOL CALLBACK Bootstrap(PINIT_ONCE, PVOID, PVOID*)
{
    InterlockedExchange(&g_bootstrapThread, static_cast<LONG>(GetCurrentThreadId()));
    HRESULT hr = InstallExplorerHook();
    if (FAILED(hr))
        Report(L"hook not installed; retaining native shell", hr);
    InterlockedExchange(&g_bootstrapThread, 0);
    return TRUE;
}

extern "C" HRESULT WINAPI DXGIDeclareAdapterRemovalSupport()
{
    // A dependency may recursively invoke DXGI on this same thread. Other
    // threads wait for bootstrap completion; the owner must not wait on itself.
    if (static_cast<DWORD>(InterlockedCompareExchange(&g_bootstrapThread, 0, 0)) != GetCurrentThreadId())
        InitOnceExecuteOnce(&g_bootstrapOnce, Bootstrap, nullptr, nullptr);
    using Function = HRESULT (WINAPI*)();
    return reinterpret_cast<Function>(ResolveDxgiExport(10))();
}

BOOL WINAPI DllMain(HINSTANCE, DWORD, LPVOID)
{
    return TRUE;
}
