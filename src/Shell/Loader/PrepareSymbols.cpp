#include "LoaderCommon.h"
#include <dbghelp.h>
#include <winhttp.h>
#include <shlobj.h>
#include <strsafe.h>
#include <cstdio>
#include <cstring>
#include <cwchar>

#pragma comment(lib, "dbghelp.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "shell32.lib")

namespace
{
    enum ExitCode
    {
        Success = 0,
        UsageError = 2,
        ImageError = 3,
        DirectoryError = 4,
        DownloadError = 5,
        SymbolError = 6,
        CacheError = 7,
        MenuFixError = 8
    };

    struct Handle
    {
        HANDLE value = INVALID_HANDLE_VALUE;
        void Close()
        {
            if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value);
            value = INVALID_HANDLE_VALUE;
        }
        ~Handle() { Close(); }
    };

    struct InternetHandle
    {
        HINTERNET value = nullptr;
        ~InternetHandle() { if (value) WinHttpCloseHandle(value); }
    };

    struct TemporaryFile
    {
        WCHAR path[MAX_PATH]{};
        ~TemporaryFile() { if (path[0]) DeleteFileW(path); }
    };

    struct SymbolSession
    {
        HANDLE process = GetCurrentProcess();
        bool initialized = false;
        ~SymbolSession() { if (initialized) SymCleanup(process); }
    };

    HRESULT LastErrorResult()
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error ? error : ERROR_GEN_FAILURE);
    }

    int Fail(ExitCode code, PCWSTR operation, HRESULT hr)
    {
        WCHAR message[512]{};
        DWORD error = HRESULT_FACILITY(hr) == FACILITY_WIN32 ? HRESULT_CODE(hr) : static_cast<DWORD>(hr);
        FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                       nullptr, error, 0, message, ARRAYSIZE(message), nullptr);
        fwprintf(stderr, L"%s failed (0x%08lX)%s%s\n", operation, static_cast<DWORD>(hr),
                 message[0] ? L": " : L"", message);
        return code;
    }

    HRESULT EnsureDirectory(PCWSTR directory)
    {
        const int error = SHCreateDirectoryExW(nullptr, directory, nullptr);
        if (error != ERROR_SUCCESS && error != ERROR_ALREADY_EXISTS && error != ERROR_FILE_EXISTS)
            return HRESULT_FROM_WIN32(error);
        const DWORD attributes = GetFileAttributesW(directory);
        if (attributes == INVALID_FILE_ATTRIBUTES) return LastErrorResult();
        return (attributes & FILE_ATTRIBUTE_DIRECTORY) ? S_OK : HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    bool SameIdentity(const ExplorerImageIdentity& a, const ExplorerImageIdentity& b)
    {
        return IsEqualGUID(a.pdbGuid, b.pdbGuid) && a.pdbAge == b.pdbAge &&
            a.timeDateStamp == b.timeDateStamp && a.sizeOfImage == b.sizeOfImage &&
            memcmp(a.sha256, b.sha256, sizeof(a.sha256)) == 0;
    }

    HRESULT VerifyPdb(PCWSTR path, const ExplorerImageIdentity& identity)
    {
        SYMSRV_INDEX_INFOW info{};
        info.sizeofstruct = sizeof(info);
        if (!SymSrvGetFileIndexInfoW(path, &info, 0)) return LastErrorResult();
        if (!IsEqualGUID(info.guid, identity.pdbGuid) || info.age != identity.pdbAge)
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        return S_OK;
    }

    HRESULT WriteAll(HANDLE file, const BYTE* data, DWORD length)
    {
        while (length)
        {
            DWORD written = 0;
            if (!WriteFile(file, data, length, &written, nullptr)) return LastErrorResult();
            if (!written) return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
            data += written;
            length -= written;
        }
        return S_OK;
    }

    HRESULT DownloadPdb(PCWSTR directory, PCWSTR path, PCWSTR requestPath,
                        const ExplorerImageIdentity& identity)
    {
        TemporaryFile temporary;
        if (!GetTempFileNameW(directory, L"pdb", 0, temporary.path)) return LastErrorResult();
        // DbgHelp selects its PDB reader by filename extension. Valid PDB bytes
        // under GetTempFileName's .tmp suffix are treated as an invalid PE.
        WCHAR pdbTemporary[MAX_PATH]{};
        HRESULT nameResult = StringCchPrintfW(pdbTemporary, ARRAYSIZE(pdbTemporary),
                                             L"%s.pdb", temporary.path);
        if (FAILED(nameResult)) return nameResult;
        if (!MoveFileExW(temporary.path, pdbTemporary, 0)) return LastErrorResult();
        StringCchCopyW(temporary.path, ARRAYSIZE(temporary.path), pdbTemporary);
        Handle file;
        file.value = CreateFileW(temporary.path, GENERIC_WRITE, 0, nullptr, TRUNCATE_EXISTING,
                                 FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
        InternetHandle session;
        session.value = WinHttpOpen(L"ExplorerEx.Loader/1", WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
                                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!session.value) return LastErrorResult();
        if (!WinHttpSetTimeouts(session.value, 30000, 30000, 30000, 60000)) return LastErrorResult();
        InternetHandle connection;
        connection.value = WinHttpConnect(session.value, L"msdl.microsoft.com", INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (!connection.value) return LastErrorResult();
        InternetHandle request;
        request.value = WinHttpOpenRequest(connection.value, L"GET", requestPath, nullptr,
            WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
        if (!request.value) return LastErrorResult();
        DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_DISALLOW_HTTPS_TO_HTTP;
        DWORD redirectLimit = 5;
        if (!WinHttpSetOption(request.value, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy)) ||
            !WinHttpSetOption(request.value, WINHTTP_OPTION_MAX_HTTP_AUTOMATIC_REDIRECTS, &redirectLimit, sizeof(redirectLimit)))
            return LastErrorResult();
        if (!WinHttpSendRequest(request.value, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
            !WinHttpReceiveResponse(request.value, nullptr))
            return LastErrorResult();
        DWORD status = 0;
        DWORD statusBytes = sizeof(status);
        if (!WinHttpQueryHeaders(request.value, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusBytes, WINHTTP_NO_HEADER_INDEX))
            return LastErrorResult();
        if (status != HTTP_STATUS_OK)
        {
            fwprintf(stderr, L"Microsoft symbol server returned HTTP %lu for %s\n", status, requestPath);
            return HRESULT_FROM_WIN32(ERROR_WINHTTP_INVALID_SERVER_RESPONSE);
        }
        ULONGLONG expected = 0;
        bool hasExpectedLength = false;
        WCHAR contentLength[32]{};
        DWORD contentLengthBytes = sizeof(contentLength);
        if (WinHttpQueryHeaders(request.value, WINHTTP_QUERY_CONTENT_LENGTH,
                               WINHTTP_HEADER_NAME_BY_INDEX, contentLength, &contentLengthBytes,
                               WINHTTP_NO_HEADER_INDEX))
        {
            WCHAR* end = nullptr;
            expected = _wcstoui64(contentLength, &end, 10);
            if (!contentLength[0] || !end || *end || contentLength[0] == L'-')
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            hasExpectedLength = true;
        }
        else if (GetLastError() != ERROR_WINHTTP_HEADER_NOT_FOUND)
            return LastErrorResult();
        BYTE buffer[64 * 1024];
        ULONGLONG total = 0;
        for (;;)
        {
            DWORD received = 0;
            if (!WinHttpReadData(request.value, buffer, sizeof(buffer), &received)) return LastErrorResult();
            if (!received) break;
            HRESULT hr = WriteAll(file.value, buffer, received);
            if (FAILED(hr)) return hr;
            total += received;
        }
        if (!total || (hasExpectedLength && total != expected))
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        if (!FlushFileBuffers(file.value)) return LastErrorResult();
        file.Close();
        HRESULT hr = VerifyPdb(temporary.path, identity);
        if (FAILED(hr)) return hr;
        if (!MoveFileExW(temporary.path, path, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            return LastErrorResult();
        temporary.path[0] = L'\0';
        wprintf(L"Downloaded and verified %llu bytes of matching symbols.\n", total);
        return S_OK;
    }

    struct SymbolRequest
    {
        PCWSTR mask; // SymEnumSymbols filter
        PCWSTR name; // exact name; a leading '?' means the exact decorated name
        DWORD rva = 0;
    };

    struct SymbolMatch
    {
        PCWSTR name = nullptr;
        DWORD64 address = 0;
        ULONG size = 0;
        unsigned count = 0;
    };

    BOOL CALLBACK FindSymbol(PSYMBOL_INFOW symbol, ULONG, PVOID context)
    {
        auto& match = *static_cast<SymbolMatch*>(context);
        WCHAR name[MAX_SYM_NAME]{};
        PCWSTR candidate = symbol->Name;
        if (match.name[0] != L'?' && candidate[0] == L'?' &&
            UnDecorateSymbolNameW(candidate, name, ARRAYSIZE(name), UNDNAME_NAME_ONLY))
            candidate = name;
        if (wcscmp(candidate, match.name) != 0) return TRUE;
        ++match.count;
        match.address = symbol->Address;
        match.size = symbol->Size;
        return TRUE;
    }

    HRESULT ResolveSymbols(PCWSTR imagePath, PCWSTR symbolsDirectory,
                           const ExplorerImageIdentity& identity, SymbolRequest* requests, unsigned count)
    {
        SymbolSession session;
        // Raw names: exact decorated matching must not see DbgHelp's undecoration.
        SymSetOptions(SYMOPT_EXACT_SYMBOLS | SYMOPT_FAIL_CRITICAL_ERRORS |
                      SYMOPT_NO_PROMPTS | SYMOPT_IGNORE_NT_SYMPATH);
        if (!SymInitializeW(session.process, symbolsDirectory, FALSE)) return LastErrorResult();
        session.initialized = true;
        const DWORD64 base = SymLoadModuleExW(session.process, nullptr, imagePath, nullptr, 0, 0, nullptr, 0);
        if (!base) return LastErrorResult();
        IMAGEHLP_MODULEW64 module{};
        module.SizeOfStruct = sizeof(module);
        if (!SymGetModuleInfoW64(session.process, base, &module)) return LastErrorResult();
        if (module.SymType != SymPdb || module.PdbUnmatched || module.DbgUnmatched ||
            !IsEqualGUID(module.PdbSig70, identity.pdbGuid) || module.PdbAge != identity.pdbAge ||
            module.ImageSize != identity.sizeOfImage || module.TimeDateStamp != identity.timeDateStamp)
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        for (unsigned i = 0; i < count; ++i)
        {
            SymbolMatch match;
            match.name = requests[i].name;
            if (!SymEnumSymbolsW(session.process, base, requests[i].mask, FindSymbol, &match))
                return LastErrorResult();
            if (match.count != 1)
            {
                fwprintf(stderr, L"Expected one exact %s symbol; found %u.\n", requests[i].name, match.count);
                return HRESULT_FROM_WIN32(match.count ? ERROR_DUP_NAME : ERROR_NOT_FOUND);
            }
            if (match.address < base || match.address - base >= identity.sizeOfImage ||
                identity.sizeOfImage - (match.address - base) < 16 || (match.size && match.size < 16))
                return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
            requests[i].rva = static_cast<DWORD>(match.address - base);
        }
        return S_OK;
    }

    // Read from the on-disk image, not a relocated or potentially hooked process.
    HRESULT ReadPrologue(PCWSTR path, DWORD rva, BYTE (&prologue)[16])
    {
        Handle file;
        file.value = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE,
                                 nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
        LARGE_INTEGER fileSize{};
        if (!GetFileSizeEx(file.value, &fileSize)) return LastErrorResult();
        IMAGE_DOS_HEADER dos{};
        DWORD read = 0;
        if (!ReadFile(file.value, &dos, sizeof(dos), &read, nullptr)) return LastErrorResult();
        if (read != sizeof(dos) || dos.e_magic != IMAGE_DOS_SIGNATURE || dos.e_lfanew < 0 ||
            static_cast<ULONGLONG>(dos.e_lfanew) + sizeof(IMAGE_NT_HEADERS64) > static_cast<ULONGLONG>(fileSize.QuadPart))
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        LARGE_INTEGER position{};
        position.QuadPart = dos.e_lfanew;
        if (!SetFilePointerEx(file.value, position, nullptr, FILE_BEGIN)) return LastErrorResult();
        IMAGE_NT_HEADERS64 nt{};
        if (!ReadFile(file.value, &nt, sizeof(nt), &read, nullptr)) return LastErrorResult();
        if (read != sizeof(nt) || nt.Signature != IMAGE_NT_SIGNATURE ||
            nt.FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 ||
            nt.OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC ||
            nt.FileHeader.SizeOfOptionalHeader < sizeof(IMAGE_OPTIONAL_HEADER64) ||
            !nt.FileHeader.NumberOfSections || nt.FileHeader.NumberOfSections > 96 ||
            rva >= nt.OptionalHeader.SizeOfImage || nt.OptionalHeader.SizeOfImage - rva < sizeof(prologue))
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        position.QuadPart = static_cast<LONGLONG>(dos.e_lfanew) +
            FIELD_OFFSET(IMAGE_NT_HEADERS64, OptionalHeader) + nt.FileHeader.SizeOfOptionalHeader;
        if (position.QuadPart + nt.FileHeader.NumberOfSections * sizeof(IMAGE_SECTION_HEADER) >
            static_cast<ULONGLONG>(fileSize.QuadPart))
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        if (!SetFilePointerEx(file.value, position, nullptr, FILE_BEGIN)) return LastErrorResult();
        for (WORD i = 0; i < nt.FileHeader.NumberOfSections; ++i)
        {
            IMAGE_SECTION_HEADER section{};
            if (!ReadFile(file.value, &section, sizeof(section), &read, nullptr)) return LastErrorResult();
            if (read != sizeof(section)) return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
            if (rva < section.VirtualAddress) continue;
            const DWORD offset = rva - section.VirtualAddress;
            if (offset >= section.Misc.VirtualSize) continue;
            if (!(section.Characteristics & IMAGE_SCN_MEM_EXECUTE) ||
                section.Misc.VirtualSize - offset < sizeof(prologue) ||
                offset > section.SizeOfRawData || section.SizeOfRawData - offset < sizeof(prologue))
                return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
            position.QuadPart = static_cast<ULONGLONG>(section.PointerToRawData) + offset;
            if (position.QuadPart + sizeof(prologue) > static_cast<ULONGLONG>(fileSize.QuadPart))
                return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
            if (!SetFilePointerEx(file.value, position, nullptr, FILE_BEGIN)) return LastErrorResult();
            if (!ReadFile(file.value, prologue, sizeof(prologue), &read, nullptr)) return LastErrorResult();
            return read == sizeof(prologue) ? S_OK : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }
        return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
    }

    HRESULT SaveCache(PCWSTR directory, PCWSTR name, const void* data, DWORD length)
    {
        WCHAR path[MAX_PATH]{};
        HRESULT hr = StringCchPrintfW(path, ARRAYSIZE(path), L"%s\\%s", directory, name);
        if (FAILED(hr)) return hr;
        TemporaryFile temporary;
        if (!GetTempFileNameW(directory, L"exc", 0, temporary.path)) return LastErrorResult();
        Handle file;
        file.value = CreateFileW(temporary.path, GENERIC_WRITE, 0, nullptr, TRUNCATE_EXISTING,
                                 FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
        hr = WriteAll(file.value, static_cast<const BYTE*>(data), length);
        if (FAILED(hr)) return hr;
        if (!FlushFileBuffers(file.value)) return LastErrorResult();
        file.Close();
        if (!MoveFileExW(temporary.path, path, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            return LastErrorResult();
        temporary.path[0] = L'\0';
        return S_OK;
    }

    // Identity, exact PDB, symbol resolution and original bytes for one native
    // image. prologues[i] receives the original bytes of requests[i].
    int PrepareImage(PCWSTR label, PCWSTR imagePath, bool dll, PCWSTR directory,
                     ExplorerImageIdentity* image, SymbolRequest* requests, unsigned count,
                     BYTE (*prologues)[16])
    {
        WCHAR pdbName[MAX_PATH]{};
        HRESULT hr = ReadExplorerImageIdentity(imagePath, image, pdbName, dll);
        if (FAILED(hr)) return Fail(ImageError, L"Validating native x64 image", hr);
        wprintf(L"%s image: %s\nPDB: %s, age %lu\n", label, imagePath, pdbName, image->pdbAge);
        hr = EnsureDirectory(directory);
        if (FAILED(hr)) return Fail(DirectoryError, L"Creating payload directory", hr);
        WCHAR key[48]{};
        const GUID& guid = image->pdbGuid;
        hr = StringCchPrintfW(key, ARRAYSIZE(key), L"%08lX%04X%04X%02X%02X%02X%02X%02X%02X%02X%02X%X",
            guid.Data1, guid.Data2, guid.Data3, guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
            guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7], image->pdbAge);
        if (FAILED(hr)) return Fail(ImageError, L"Formatting PDB identity", hr);
        WCHAR symbolsDirectory[MAX_PATH]{};
        hr = StringCchPrintfW(symbolsDirectory, ARRAYSIZE(symbolsDirectory), L"%s\\symbols\\%s", directory, key);
        if (FAILED(hr)) return Fail(DirectoryError, L"Resolving symbol directory", hr);
        hr = EnsureDirectory(symbolsDirectory);
        if (FAILED(hr)) return Fail(DirectoryError, L"Creating symbol directory", hr);
        WCHAR pdbPath[MAX_PATH]{};
        hr = StringCchPrintfW(pdbPath, ARRAYSIZE(pdbPath), L"%s\\%s", symbolsDirectory, pdbName);
        if (FAILED(hr)) return Fail(DirectoryError, L"Resolving PDB path", hr);
        if (SUCCEEDED(VerifyPdb(pdbPath, *image)))
            wprintf(L"Reusing verified PDB: %s\n", pdbPath);
        else
        {
            WCHAR requestPath[MAX_PATH * 3]{};
            hr = StringCchPrintfW(requestPath, ARRAYSIZE(requestPath), L"/download/symbols/%s/%s/%s", pdbName, key, pdbName);
            if (FAILED(hr)) return Fail(DownloadError, L"Formatting symbol URL", hr);
            wprintf(L"Downloading https://msdl.microsoft.com%s\n", requestPath);
            hr = DownloadPdb(symbolsDirectory, pdbPath, requestPath, *image);
            if (FAILED(hr)) return Fail(DownloadError, L"Downloading exact PDB", hr);
        }
        hr = ResolveSymbols(imagePath, symbolsDirectory, *image, requests, count);
        if (FAILED(hr)) return Fail(SymbolError, L"Resolving symbols from exact PDB", hr);
        for (unsigned i = 0; i < count; ++i)
        {
            hr = ReadPrologue(imagePath, requests[i].rva, prologues[i]);
            if (FAILED(hr)) return Fail(SymbolError, L"Validating executable symbol and original bytes", hr);
        }
        ExplorerImageIdentity currentIdentity{};
        WCHAR currentPdbName[MAX_PATH]{};
        hr = ReadExplorerImageIdentity(imagePath, &currentIdentity, currentPdbName, dll);
        if (FAILED(hr)) return Fail(ImageError, L"Rechecking native image", hr);
        if (!SameIdentity(*image, currentIdentity))
            return Fail(ImageError, L"Native image changed during preparation; run prepare again",
                        HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
        return Success;
    }
}

int wmain(int argc, wchar_t** argv)
{
    if (argc < 2 || argc > 3 || _wcsicmp(argv[1], L"prepare") != 0)
    {
        fwprintf(stderr, L"Usage: ExplorerEx.Loader.exe prepare [payload-directory]\n"
                         L"Default: native Program Files\\ExplorerEx. Run elevated when preparing that directory.\n"
                         L"Exit codes: 0 success, 2 usage, 3 native image, 4 directory, 5 PDB download/identity, "
                         L"6 symbol resolution, 7 cache publication, 8 optional shell32 fixes unavailable "
                         L"(ExplorerEx itself is still prepared).\n");
        return UsageError;
    }
    WCHAR directory[MAX_PATH]{};
    HRESULT hr = S_OK;
    if (argc == 3)
    {
        const DWORD length = GetFullPathNameW(argv[2], ARRAYSIZE(directory), directory, nullptr);
        if (!length) hr = LastErrorResult();
        else if (length >= ARRAYSIZE(directory)) hr = HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE);
    }
    else
        hr = GetExplorerExDirectory(directory);
    if (FAILED(hr)) return Fail(DirectoryError, L"Resolving payload directory", hr);
    WCHAR windows[MAX_PATH]{};
    const UINT windowsLength = GetWindowsDirectoryW(windows, ARRAYSIZE(windows));
    if (!windowsLength) return Fail(ImageError, L"Finding Windows directory", LastErrorResult());
    if (windowsLength >= ARRAYSIZE(windows))
        return Fail(ImageError, L"Finding Windows directory", HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE));
    WCHAR imagePath[MAX_PATH]{};
    hr = StringCchPrintfW(imagePath, ARRAYSIZE(imagePath), L"%s\\explorer.exe", windows);
    if (FAILED(hr)) return Fail(ImageError, L"Resolving native Explorer path", hr);

    ExplorerHookCache cache{};
    cache.magic = ExplorerHookCacheMagic;
    cache.version = ExplorerHookCacheVersion;
    SymbolRequest explorerRequests[] = { { L"*CreateDesktopAndTray*", L"CreateDesktopAndTray" } };
    BYTE explorerPrologues[1][16]{};
    int result = PrepareImage(L"Native", imagePath, false, directory, &cache.image,
                              explorerRequests, ARRAYSIZE(explorerRequests), explorerPrologues);
    if (result != Success) return result;
    memcpy(cache.prologue, explorerPrologues[0], sizeof(cache.prologue));
    cache.functionRva = explorerRequests[0].rva;
    hr = SaveCache(directory, ExplorerHookCacheName, &cache, sizeof(cache));
    if (FAILED(hr)) return Fail(CacheError, L"Atomically publishing hook cache", hr);
    wprintf(L"Prepared CreateDesktopAndTray RVA 0x%08lX for image timestamp 0x%08lX, size 0x%08lX.\n"
             L"Cache: %s\\%s\n", cache.functionRva, cache.image.timeDateStamp,
             cache.image.sizeOfImage, directory, ExplorerHookCacheName);

    // Optional: shell32 menu fix. Never leave a cache for a different shell32.
    WCHAR shell32CachePath[MAX_PATH]{};
    StringCchPrintfW(shell32CachePath, ARRAYSIZE(shell32CachePath), L"%s\\%s", directory, Shell32HookCacheName);
    DeleteFileW(shell32CachePath);
    WCHAR system[MAX_PATH]{};
    WCHAR shell32Path[MAX_PATH]{};
    const UINT systemLength = GetSystemDirectoryW(system, ARRAYSIZE(system));
    if (!systemLength || systemLength >= ARRAYSIZE(system) ||
        FAILED(StringCchPrintfW(shell32Path, ARRAYSIZE(shell32Path), L"%s\\shell32.dll", system)))
    {
        fwprintf(stderr, L"Cannot locate shell32.dll; menu fix not prepared.\n");
        return MenuFixError;
    }
    Shell32HookCache menuCache{};
    menuCache.magic = ExplorerHookCacheMagic;
    menuCache.version = Shell32HookCacheVersion;
    SymbolRequest menuRequests[] = {
        { L"*GetSize*", L"?GetSize@CMenuSFToolbar@@UEAAXPEAUtagSIZE@@@Z" },
        { L"*GetSize*", L"?GetSize@CMenuToolbarBase@@UEAAXPEAUtagSIZE@@@Z" },
    };
    BYTE menuPrologues[ARRAYSIZE(menuRequests)][16]{};
    result = PrepareImage(L"shell32", shell32Path, true, directory, &menuCache.image,
                          menuRequests, ARRAYSIZE(menuRequests), menuPrologues);
    if (result != Success)
    {
        fwprintf(stderr, L"Warning: shell32 menu fix not prepared (exit %d); ExplorerEx is otherwise ready.\n", result);
        return MenuFixError;
    }
    memcpy(menuCache.prologue, menuPrologues[0], sizeof(menuCache.prologue));
    menuCache.targetRva = menuRequests[0].rva;
    menuCache.replacementRva = menuRequests[1].rva;
    hr = SaveCache(directory, Shell32HookCacheName, &menuCache, sizeof(menuCache));
    if (FAILED(hr))
    {
        Fail(CacheError, L"Publishing shell32 menu fix cache", hr);
        return MenuFixError;
    }
    wprintf(L"Prepared shell32 menu fix: CMenuSFToolbar::GetSize 0x%08lX -> CMenuToolbarBase::GetSize 0x%08lX.\n",
            menuCache.targetRva, menuCache.replacementRva);
    return Success;
}
