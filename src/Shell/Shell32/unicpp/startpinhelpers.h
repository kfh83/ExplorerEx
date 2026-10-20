#pragma once

#include "pch.h"
#include "shundoc.h"
#include "path.h"
#include "cowsite.h"
#include "dpa.h"
#include "../startpinres.h"

// The XP wire format is retained, but it lives exclusively in ExplorerEx's key.
#define REGSTR_PATH_STARTFAVS TEXT("Software\\ExplorerEx\\StartPage")
#define REGSTR_VAL_STARTFAVS TEXT("Favorites")
#define REGSTR_VAL_STARTFAVCHANGES TEXT("FavoritesChanges")
#define REGSTR_VAL_STARTFAVLINKS TEXT("FavoritesResolve")

extern "C" IMAGE_DOS_HEADER __ImageBase;
#define PIN_RESOURCES reinterpret_cast<HINSTANCE>(&__ImageBase)

// Recursive because Load -> EnumObjects -> initialization shares this lock.
class CPinStoreLock
{
    static CRITICAL_SECTION& Section()
    {
        static struct Lock {
            CRITICAL_SECTION cs;
            Lock() { InitializeCriticalSection(&cs); }
            ~Lock() { DeleteCriticalSection(&cs); }
        } lock;
        return lock.cs;
    }
public:
    CPinStoreLock() { EnterCriticalSection(&Section()); }
    ~CPinStoreLock() { LeaveCriticalSection(&Section()); }
};

static bool PinValidPidl(const BYTE* data, SIZE_T size, bool exact)
{
    SIZE_T offset = 0;
    while (size - offset >= sizeof(USHORT))
    {
        USHORT cb;
        CopyMemory(&cb, data + offset, sizeof(cb));
        if (!cb)
            return !exact || offset + sizeof(cb) == size;
        if (cb < sizeof(cb) || cb > size - offset)
            return false;
        offset += cb;
    }
    return false;
}

static HRESULT PinReadPidl(IStream* stream, LPITEMIDLIST* result)
{
    *result = NULL;
    DWORD size = 0;
    HRESULT hr = IStream_Read(stream, &size, sizeof(size));
    if (FAILED(hr)) return hr;
    STATSTG stat = {};
    ULARGE_INTEGER pos = {};
    LARGE_INTEGER zero = {};
    if (FAILED(hr = stream->Stat(&stat, STATFLAG_NONAME)) ||
        FAILED(hr = stream->Seek(zero, STREAM_SEEK_CUR, &pos))) return hr;
    if (size < sizeof(USHORT) || size > 0x100000 ||
        pos.QuadPart > stat.cbSize.QuadPart || size > stat.cbSize.QuadPart - pos.QuadPart)
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    LPITEMIDLIST pidl = static_cast<LPITEMIDLIST>(CoTaskMemAlloc(size));
    if (!pidl) return E_OUTOFMEMORY;
    hr = IStream_Read(stream, pidl, size);
    if (SUCCEEDED(hr) && !PinValidPidl(reinterpret_cast<const BYTE*>(pidl), size, true))
        hr = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    if (FAILED(hr)) CoTaskMemFree(pidl);
    else *result = pidl;
    return hr;
}

static HRESULT PinWritePidl(IStream* stream, LPCITEMIDLIST pidl)
{
    DWORD size = ILGetSize(pidl);
    HRESULT hr = IStream_Write(stream, &size, sizeof(size));
    return FAILED(hr) ? hr : IStream_Write(stream, pidl, size);
}

static BOOL PinSetPidl(LPITEMIDLIST* target, LPCITEMIDLIST source)
{
    LPITEMIDLIST clone = ILClone(source);
    if (!clone) return FALSE;
    ILFree(*target);
    *target = clone;
    return TRUE;
}

static LPITEMIDLIST PinDataObjectPidl(IDataObject* object)
{
    FORMATETC format = { static_cast<CLIPFORMAT>(RegisterClipboardFormat(CFSTR_SHELLIDLIST)),
        NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM medium = {};
    if (FAILED(object->GetData(&format, &medium))) return NULL;
    LPITEMIDLIST result = NULL;
    if (medium.tymed == TYMED_HGLOBAL && medium.hGlobal)
    {
        SIZE_T size = GlobalSize(medium.hGlobal);
        const BYTE* data = static_cast<const BYTE*>(GlobalLock(medium.hGlobal));
        if (data && size >= sizeof(UINT) * 3)
        {
            UINT fields[3];
            CopyMemory(fields, data, sizeof(fields));
            if (fields[0] == 1 && fields[1] >= sizeof(fields) && fields[2] >= sizeof(fields) &&
                fields[1] < size && fields[2] < size &&
                PinValidPidl(data + fields[1], size - fields[1], false) &&
                PinValidPidl(data + fields[2], size - fields[2], false))
                result = ILCombine(reinterpret_cast<LPCITEMIDLIST>(data + fields[1]),
                                   reinterpret_cast<LPCITEMIDLIST>(data + fields[2]));
        }
        if (data) GlobalUnlock(medium.hGlobal);
    }
    ReleaseStgMedium(&medium);
    return result;
}

static HRESULT PinStoreStream(LPCTSTR value, IStream* stream)
{
    HGLOBAL memory = NULL;
    STATSTG stat = {};
    HRESULT hr = stream->Stat(&stat, STATFLAG_NONAME);
    if (FAILED(hr)) return hr;
    if (stat.cbSize.HighPart) return STG_E_MEDIUMFULL;
    hr = GetHGlobalFromStream(stream, &memory);
    if (FAILED(hr)) return hr;
    void* data = GlobalLock(memory);
    if (!data && stat.cbSize.LowPart) return E_OUTOFMEMORY;
    LSTATUS error = SHSetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS, value,
        REG_BINARY, data, stat.cbSize.LowPart);
    if (data) GlobalUnlock(memory);
    return HRESULT_FROM_WIN32(error);
}

static HRESULT PinReadStore(LPCTSTR value, IStream** result)
{
    *result = NULL;
    DWORD size = 0, type = 0;
    LSTATUS error = SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
        value, &type, NULL, &size);
    if (error != ERROR_SUCCESS) return HRESULT_FROM_WIN32(error);
    if (type != REG_BINARY) return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, size ? size : 1);
    if (!memory) return E_OUTOFMEMORY;
    void* data = GlobalLock(memory);
    if (!data) { GlobalFree(memory); return E_OUTOFMEMORY; }
    DWORD capacity = size;
    error = SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_STARTFAVS,
        value, &type, data, &capacity);
    GlobalUnlock(memory);
    if (error != ERROR_SUCCESS || type != REG_BINARY || capacity != size)
    {
        GlobalFree(memory);
        return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_INVALID_DATA);
    }
    HRESULT hr = CreateStreamOnHGlobal(memory, TRUE, result);
    if (FAILED(hr)) { GlobalFree(memory); return hr; }
    ULARGE_INTEGER length = {};
    length.QuadPart = size;
    hr = (*result)->SetSize(length);
    if (FAILED(hr)) { (*result)->Release(); *result = NULL; }
    return hr;
}
