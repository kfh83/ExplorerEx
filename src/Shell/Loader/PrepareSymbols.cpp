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
        CacheError = 7
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

    struct SymbolMatch
    {
        DWORD64 address = 0;
        ULONG size = 0;
        unsigned count = 0;
    };

    BOOL CALLBACK FindSymbol(PSYMBOL_INFOW symbol, ULONG, PVOID context)
    {
        WCHAR name[MAX_SYM_NAME]{};
        PCWSTR candidate = symbol->Name;
        if (candidate[0] == L'?' && UnDecorateSymbolNameW(candidate, name, ARRAYSIZE(name), UNDNAME_NAME_ONLY))
            candidate = name;
        if (wcscmp(candidate, L"CreateDesktopAndTray") != 0) return TRUE;
        auto& match = *static_cast<SymbolMatch*>(context);
        ++match.count;
        match.address = symbol->Address;
        match.size = symbol->Size;
        return TRUE;
    }

    HRESULT ResolveSymbol(PCWSTR imagePath, PCWSTR symbolsDirectory,
                          const ExplorerImageIdentity& identity, DWORD* rva)
    {
        SymbolSession session;
        SymSetOptions(SYMOPT_UNDNAME | SYMOPT_EXACT_SYMBOLS | SYMOPT_FAIL_CRITICAL_ERRORS |
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
        SymbolMatch match;
        if (!SymEnumSymbolsW(session.process, base, L"*CreateDesktopAndTray*", FindSymbol, &match))
            return LastErrorResult();
        if (match.count != 1)
        {
            fwprintf(stderr, L"Expected one exact CreateDesktopAndTray symbol; found %u.\n", match.count);
            return HRESULT_FROM_WIN32(match.count ? ERROR_DUP_NAME : ERROR_NOT_FOUND);
        }
        if (match.address < base || match.address - base >= identity.sizeOfImage ||
            identity.sizeOfImage - (match.address - base) < 16 || (match.size && match.size < 16))
            return HRESULT_FROM_WIN32(ERROR_INVALID_ADDRESS);
        *rva = static_cast<DWORD>(match.address - base);
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

    HRESULT SaveCache(PCWSTR directory, const ExplorerHookCache& cache)
    {
        WCHAR path[MAX_PATH]{};
        HRESULT hr = StringCchPrintfW(path, ARRAYSIZE(path), L"%s\\%s", directory, ExplorerHookCacheName);
        if (FAILED(hr)) return hr;
        TemporaryFile temporary;
        if (!GetTempFileNameW(directory, L"exc", 0, temporary.path)) return LastErrorResult();
        Handle file;
        file.value = CreateFileW(temporary.path, GENERIC_WRITE, 0, nullptr, TRUNCATE_EXISTING,
                                 FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
        hr = WriteAll(file.value, reinterpret_cast<const BYTE*>(&cache), sizeof(cache));
        if (FAILED(hr)) return hr;
        if (!FlushFileBuffers(file.value)) return LastErrorResult();
        file.Close();
        if (!MoveFileExW(temporary.path, path, MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
            return LastErrorResult();
        temporary.path[0] = L'\0';
        return S_OK;
    }
}

int wmain(int argc, wchar_t** argv)
{
    if (argc < 2 || argc > 3 || _wcsicmp(argv[1], L"prepare") != 0)
    {
        fwprintf(stderr, L"Usage: ExplorerEx.Loader.exe prepare [payload-directory]\n"
                         L"Default: native Program Files\\ExplorerEx. Run elevated when preparing that directory.\n"
                         L"Exit codes: 0 success, 2 usage, 3 native image, 4 directory, 5 PDB download/identity, "
                         L"6 symbol resolution, 7 cache publication.\n");
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
    WCHAR pdbName[MAX_PATH]{};
    hr = ReadExplorerImageIdentity(imagePath, &cache.image, pdbName);
    if (FAILED(hr)) return Fail(ImageError, L"Validating native x64 Explorer", hr);
    wprintf(L"Native image: %s\nPDB: %s, age %lu\n", imagePath, pdbName, cache.image.pdbAge);
    hr = EnsureDirectory(directory);
    if (FAILED(hr)) return Fail(DirectoryError, L"Creating payload directory", hr);
    WCHAR key[48]{};
    const GUID& guid = cache.image.pdbGuid;
    hr = StringCchPrintfW(key, ARRAYSIZE(key), L"%08lX%04X%04X%02X%02X%02X%02X%02X%02X%02X%02X%X",
        guid.Data1, guid.Data2, guid.Data3, guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
        guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7], cache.image.pdbAge);
    if (FAILED(hr)) return Fail(ImageError, L"Formatting PDB identity", hr);
    WCHAR symbolsDirectory[MAX_PATH]{};
    hr = StringCchPrintfW(symbolsDirectory, ARRAYSIZE(symbolsDirectory), L"%s\\symbols\\%s", directory, key);
    if (FAILED(hr)) return Fail(DirectoryError, L"Resolving symbol directory", hr);
    hr = EnsureDirectory(symbolsDirectory);
    if (FAILED(hr)) return Fail(DirectoryError, L"Creating symbol directory", hr);
    WCHAR pdbPath[MAX_PATH]{};
    hr = StringCchPrintfW(pdbPath, ARRAYSIZE(pdbPath), L"%s\\%s", symbolsDirectory, pdbName);
    if (FAILED(hr)) return Fail(DirectoryError, L"Resolving PDB path", hr);
    if (SUCCEEDED(VerifyPdb(pdbPath, cache.image)))
        wprintf(L"Reusing verified PDB: %s\n", pdbPath);
    else
    {
        WCHAR requestPath[MAX_PATH * 3]{};
        hr = StringCchPrintfW(requestPath, ARRAYSIZE(requestPath), L"/download/symbols/%s/%s/%s", pdbName, key, pdbName);
        if (FAILED(hr)) return Fail(DownloadError, L"Formatting symbol URL", hr);
        wprintf(L"Downloading https://msdl.microsoft.com%s\n", requestPath);
        hr = DownloadPdb(symbolsDirectory, pdbPath, requestPath, cache.image);
        if (FAILED(hr)) return Fail(DownloadError, L"Downloading exact Explorer PDB", hr);
    }
    hr = ResolveSymbol(imagePath, symbolsDirectory, cache.image, &cache.functionRva);
    if (FAILED(hr)) return Fail(SymbolError, L"Resolving CreateDesktopAndTray from exact PDB", hr);
    hr = ReadPrologue(imagePath, cache.functionRva, cache.prologue);
    if (FAILED(hr)) return Fail(SymbolError, L"Validating executable symbol and original bytes", hr);
    ExplorerImageIdentity currentIdentity{};
    WCHAR currentPdbName[MAX_PATH]{};
    hr = ReadExplorerImageIdentity(imagePath, &currentIdentity, currentPdbName);
    if (FAILED(hr)) return Fail(ImageError, L"Rechecking native Explorer", hr);
    if (!SameIdentity(cache.image, currentIdentity))
        return Fail(ImageError, L"Native Explorer changed during preparation; run prepare again",
                    HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH));
    hr = SaveCache(directory, cache);
    if (FAILED(hr)) return Fail(CacheError, L"Atomically publishing hook cache", hr);
    wprintf(L"Prepared CreateDesktopAndTray RVA 0x%08lX for image timestamp 0x%08lX, size 0x%08lX.\n"
             L"Cache: %s\\%s\n", cache.functionRva, cache.image.timeDateStamp,
             cache.image.sizeOfImage, directory, ExplorerHookCacheName);
    return Success;
}
