#include "LoaderCommon.h"
#include <bcrypt.h>
#include <shlobj.h>
#include <strsafe.h>
#include <cstring>
#include <cwchar>
#include <limits>

#pragma comment(lib, "bcrypt.lib")
#pragma comment(lib, "shell32.lib")

namespace
{
    struct Handle
    {
        HANDLE value = nullptr;
        ~Handle() { if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value); }
    };

    struct View
    {
        const BYTE* data = nullptr;
        ~View() { if (data) UnmapViewOfFile(data); }
    };

    HRESULT LastErrorResult()
    {
        const DWORD error = GetLastError();
        return HRESULT_FROM_WIN32(error ? error : ERROR_GEN_FAILURE);
    }

    bool Fits(size_t offset, size_t length, size_t size)
    {
        return offset <= size && length <= size - offset;
    }

    const BYTE* RvaData(const BYTE* data, size_t size, const IMAGE_NT_HEADERS64& nt,
                        const IMAGE_SECTION_HEADER* sections, DWORD rva, DWORD length)
    {
        if (rva < nt.OptionalHeader.SizeOfHeaders &&
            Fits(rva, length, nt.OptionalHeader.SizeOfHeaders) && Fits(rva, length, size))
            return data + rva;
        for (WORD i = 0; i < nt.FileHeader.NumberOfSections; ++i)
        {
            const auto& section = sections[i];
            if (rva >= section.VirtualAddress)
            {
                const DWORD offset = rva - section.VirtualAddress;
                if (Fits(offset, length, section.SizeOfRawData) &&
                    Fits(section.PointerToRawData, section.SizeOfRawData, size))
                    return data + section.PointerToRawData + offset;
            }
        }
        return nullptr;
    }

    HRESULT HashImage(const BYTE* data, size_t size, BYTE (&digest)[32])
    {
        BCRYPT_ALG_HANDLE algorithm = nullptr;
        BCRYPT_HASH_HANDLE hash = nullptr;
        NTSTATUS status = BCryptOpenAlgorithmProvider(&algorithm, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
        if (status >= 0)
            status = BCryptCreateHash(algorithm, &hash, nullptr, 0, nullptr, 0, 0);
        while (status >= 0 && size)
        {
            const ULONG chunk = static_cast<ULONG>((size > MAXDWORD) ? MAXDWORD : size);
            status = BCryptHashData(hash, const_cast<PUCHAR>(data), chunk, 0);
            data += chunk;
            size -= chunk;
        }
        if (status >= 0)
            status = BCryptFinishHash(hash, digest, sizeof(digest), 0);
        if (hash) BCryptDestroyHash(hash);
        if (algorithm) BCryptCloseAlgorithmProvider(algorithm, 0);
        return status >= 0 ? S_OK : HRESULT_FROM_NT(status);
    }
}

HRESULT GetExplorerExDirectory(WCHAR (&path)[MAX_PATH])
{
    path[0] = L'\0';
    PWSTR programFiles = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_ProgramFilesX64, KF_FLAG_DEFAULT, nullptr, &programFiles);
    if (SUCCEEDED(hr))
        hr = StringCchPrintfW(path, MAX_PATH, L"%s\\ExplorerEx", programFiles);
    CoTaskMemFree(programFiles);
    return hr;
}

HRESULT ReadExplorerImageIdentity(PCWSTR path, ExplorerImageIdentity* identity,
                                  WCHAR (&pdbName)[MAX_PATH])
{
    if (!identity || !path) return E_INVALIDARG;
    *identity = {};
    pdbName[0] = L'\0';
    Handle file;
    file.value = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE,
                             nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
    LARGE_INTEGER length{};
    if (!GetFileSizeEx(file.value, &length)) return LastErrorResult();
    if (length.QuadPart < sizeof(IMAGE_DOS_HEADER) ||
        static_cast<ULONGLONG>(length.QuadPart) > (std::numeric_limits<size_t>::max)())
        return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    const size_t size = static_cast<size_t>(length.QuadPart);
    Handle mapping;
    mapping.value = CreateFileMappingW(file.value, nullptr, PAGE_READONLY, 0, 0, nullptr);
    if (!mapping.value) return LastErrorResult();
    View view;
    view.data = static_cast<const BYTE*>(MapViewOfFile(mapping.value, FILE_MAP_READ, 0, 0, 0));
    if (!view.data) return LastErrorResult();
    const auto badImage = HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(view.data);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew < 0 ||
        !Fits(static_cast<size_t>(dos->e_lfanew), sizeof(IMAGE_NT_HEADERS64), size))
        return badImage;
    const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS64*>(view.data + dos->e_lfanew);
    if (nt->Signature != IMAGE_NT_SIGNATURE || nt->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 ||
        !(nt->FileHeader.Characteristics & IMAGE_FILE_EXECUTABLE_IMAGE) ||
        (nt->FileHeader.Characteristics & IMAGE_FILE_DLL) ||
        nt->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC ||
        nt->FileHeader.SizeOfOptionalHeader < sizeof(IMAGE_OPTIONAL_HEADER64) ||
        nt->OptionalHeader.NumberOfRvaAndSizes <= IMAGE_DIRECTORY_ENTRY_DEBUG ||
        nt->OptionalHeader.NumberOfRvaAndSizes > IMAGE_NUMBEROF_DIRECTORY_ENTRIES ||
        !nt->FileHeader.NumberOfSections || nt->FileHeader.NumberOfSections > 96 ||
        !nt->OptionalHeader.SizeOfImage || !nt->OptionalHeader.SizeOfHeaders ||
        nt->OptionalHeader.SizeOfHeaders > size ||
        nt->OptionalHeader.SizeOfHeaders > nt->OptionalHeader.SizeOfImage)
        return badImage;
    const size_t sectionOffset = static_cast<size_t>(dos->e_lfanew) +
        FIELD_OFFSET(IMAGE_NT_HEADERS64, OptionalHeader) + nt->FileHeader.SizeOfOptionalHeader;
    const size_t sectionBytes = nt->FileHeader.NumberOfSections * sizeof(IMAGE_SECTION_HEADER);
    if (!Fits(sectionOffset, sectionBytes, size) ||
        !Fits(sectionOffset, sectionBytes, nt->OptionalHeader.SizeOfHeaders))
        return badImage;
    const auto* sections = reinterpret_cast<const IMAGE_SECTION_HEADER*>(view.data + sectionOffset);
    DWORD previousEnd = nt->OptionalHeader.SizeOfHeaders;
    for (WORD i = 0; i < nt->FileHeader.NumberOfSections; ++i)
    {
        const auto& section = sections[i];
        const DWORD span = (section.Misc.VirtualSize > section.SizeOfRawData) ?
            section.Misc.VirtualSize : section.SizeOfRawData;
        if (section.VirtualAddress < previousEnd ||
            !Fits(section.VirtualAddress, span, nt->OptionalHeader.SizeOfImage) ||
            (section.SizeOfRawData &&
             (section.PointerToRawData < nt->OptionalHeader.SizeOfHeaders ||
              !Fits(section.PointerToRawData, section.SizeOfRawData, size))))
            return badImage;
        previousEnd = section.VirtualAddress + span;
    }
    const auto& debug = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DEBUG];
    if (!debug.VirtualAddress || !debug.Size || debug.Size % sizeof(IMAGE_DEBUG_DIRECTORY) ||
        !Fits(debug.VirtualAddress, debug.Size, nt->OptionalHeader.SizeOfImage))
        return badImage;
    const BYTE* debugBytes = RvaData(view.data, size, *nt, sections, debug.VirtualAddress, debug.Size);
    if (!debugBytes) return badImage;
    ExplorerImageIdentity result{};
    WCHAR resultName[MAX_PATH]{};
    bool found = false;
    for (DWORD offset = 0; offset < debug.Size; offset += sizeof(IMAGE_DEBUG_DIRECTORY))
    {
        IMAGE_DEBUG_DIRECTORY entry{};
        memcpy(&entry, debugBytes + offset, sizeof(entry));
        if (entry.Type != IMAGE_DEBUG_TYPE_CODEVIEW) continue;
        if (entry.SizeOfData < 25 || !Fits(entry.PointerToRawData, entry.SizeOfData, size))
            return badImage;
        const BYTE* cv = view.data + entry.PointerToRawData;
        if (memcmp(cv, "RSDS", 4) != 0) continue;
        if (found || (entry.AddressOfRawData &&
            RvaData(view.data, size, *nt, sections, entry.AddressOfRawData, entry.SizeOfData) != cv))
            return badImage;
        memcpy(&result.pdbGuid, cv + 4, sizeof(GUID));
        memcpy(&result.pdbAge, cv + 20, sizeof(DWORD));
        const char* name = reinterpret_cast<const char*>(cv + 24);
        const char* end = static_cast<const char*>(memchr(name, '\0', entry.SizeOfData - 24));
        if (!end) return badImage;
        const char* basename = name;
        for (const char* p = name; p < end; ++p)
            if (*p == '\\' || *p == '/') basename = p + 1;
        const size_t nameLength = static_cast<size_t>(end - basename);
        if (nameLength < 5 || nameLength >= MAX_PATH) return badImage;
        for (size_t i = 0; i < nameLength; ++i)
        {
            const unsigned char c = static_cast<unsigned char>(basename[i]);
            if (!((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') ||
                  (c >= '0' && c <= '9') || c == '.' || c == '_' || c == '-'))
                return badImage;
            resultName[i] = c;
        }
        const GUID empty{};
        if (_wcsicmp(resultName + nameLength - 4, L".pdb") ||
            !result.pdbAge || IsEqualGUID(result.pdbGuid, empty))
            return badImage;
        found = true;
    }
    if (!found) return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    result.timeDateStamp = nt->FileHeader.TimeDateStamp;
    result.sizeOfImage = nt->OptionalHeader.SizeOfImage;
    HRESULT hr = HashImage(view.data, size, result.sha256);
    if (FAILED(hr)) return hr;
    *identity = result;
    memcpy(pdbName, resultName, sizeof(resultName));
    return S_OK;
}

HRESULT ReadExplorerHookCache(PCWSTR directory, ExplorerHookCache* cache)
{
    if (!directory || !cache) return E_INVALIDARG;
    *cache = {};
    WCHAR path[MAX_PATH]{};
    HRESULT hr = StringCchPrintfW(path, MAX_PATH, L"%s\\%s", directory, ExplorerHookCacheName);
    if (FAILED(hr)) return hr;
    Handle file;
    file.value = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_DELETE,
                             nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file.value, &size)) return LastErrorResult();
    if (size.QuadPart != sizeof(ExplorerHookCache)) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    ExplorerHookCache result{};
    DWORD bytes = 0;
    if (!ReadFile(file.value, &result, sizeof(result), &bytes, nullptr)) return LastErrorResult();
    if (bytes != sizeof(result) || result.magic != ExplorerHookCacheMagic ||
        result.version != ExplorerHookCacheVersion || !result.functionRva ||
        !Fits(result.functionRva, sizeof(result.prologue), result.image.sizeOfImage))
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    *cache = result;
    return S_OK;
}
