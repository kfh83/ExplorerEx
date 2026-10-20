#include "pch.h"
#include "cocreateinstancehook.h"
#include "dbg.h"
#include "shundoc.h"

DWORD WINAPI BeepThread(LPVOID)
{
	Beep(1700, 200);
	return 0;
}

typedef HRESULT (STDAPICALLTYPE *PFNGETCLASSOBJECT)(REFCLSID, REFIID, void**);

static HRESULT GetLocalShellClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    // Keep the adjacent server loaded for the lifetime of its COM objects.
    // Function-local static initialization also serializes the first activation.
    static const HMODULE module = []() -> HMODULE
    {
        HMODULE owner = NULL;
        if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
            GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            reinterpret_cast<LPCWSTR>(&GetLocalShellClassObject), &owner))
            return NULL;
        WCHAR path[MAX_PATH];
        DWORD length = GetModuleFileNameW(owner, path, ARRAYSIZE(path));
        if (!length || length >= ARRAYSIZE(path) ||
            !PathRemoveFileSpecW(path) || !PathAppendW(path, L"ExplorerEx.Shell32.dll"))
            return NULL;
        return LoadLibraryExW(path, NULL, LOAD_WITH_ALTERED_SEARCH_PATH);
    }();
    if (!module)
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    static const PFNGETCLASSOBJECT getClassObject =
        reinterpret_cast<PFNGETCLASSOBJECT>(GetProcAddress(module, "DllGetClassObject"));
    return getClassObject ? getClassObject(rclsid, riid, ppv) :
        HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
}


HRESULT CoCreateInstanceHook(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID* ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;

    HRESULT hr;
    if (rclsid == CLSID_StartMenuPin)
    {
        if (!(dwClsContext & CLSCTX_INPROC_SERVER))
            return REGDB_E_CLASSNOTREG;
        IClassFactory* factory = NULL;
        hr = GetLocalShellClassObject(rclsid, IID_PPV_ARGS(&factory));
        if (SUCCEEDED(hr))
        {
            hr = factory->CreateInstance(pUnkOuter, riid, ppv);
            factory->Release();
        }
    }
    else
    {
        hr = CoCreateInstance(rclsid, pUnkOuter, dwClsContext, riid, ppv);
    }
#ifndef RELEASE
	if (FAILED(hr))
	{
		CreateThread(0, 0, BeepThread, 0, 0, 0);

		wchar_t* clsidstring = 0;
		if (FAILED(StringFromCLSID(rclsid, &clsidstring))) return hr;

		wchar_t* iidstring = 0;
		if (FAILED(StringFromCLSID(riid, &iidstring)))
		{
			if (clsidstring) CoTaskMemFree(clsidstring);
			return hr;
		}

		wprintf(L"COCREATEINSTANCE FAILED! clsid %s, riid %s\n", clsidstring, iidstring);
		if (clsidstring) CoTaskMemFree(clsidstring);
		if (iidstring) CoTaskMemFree(iidstring);
		dbg::printstacktrace();
	}
#endif
	return hr;
}

HRESULT SHCoCreateInstanceHook(_In_opt_ PCWSTR pszCLSID, _In_opt_ const CLSID* pclsid, _In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _Outptr_ void** ppv)
{
    CLSID clsid;
    const CLSID* target = pclsid;
    if (!target && pszCLSID && SUCCEEDED(CLSIDFromString(pszCLSID, &clsid)))
        target = &clsid;
    if (target && *target == CLSID_StartMenuPin)
        return CoCreateInstanceHook(*target, pUnkOuter, CLSCTX_INPROC_SERVER, riid, ppv);
	HRESULT hr = SHCoCreateInstance(pszCLSID,pclsid,pUnkOuter,riid,ppv);
#ifndef RELEASE
	if (FAILED(hr))
	{
		CreateThread(0, 0, BeepThread, 0, 0, 0);

		wchar_t* iidstring = 0;
		if (FAILED(StringFromCLSID(riid, &iidstring)))
		{
			return hr;
		}

		wchar_t* clsidStr = const_cast<wchar_t*>(pszCLSID);

		if (!clsidStr)
		{
			if (FAILED(StringFromCLSID(*pclsid, &clsidStr)))
			{
				return hr;
			}
		}

		wprintf(L"COCREATEINSTANCE FAILED! clsid %s, riid %s\n", pszCLSID, iidstring);
		if (iidstring) CoTaskMemFree(iidstring);
		dbg::printstacktrace();
	}
#endif
	return hr;
}

static bool IsStartPinVerb(PCWSTR verb)
{
    return !StrCmpIW(verb, L"pin") || !StrCmpIW(verb, L"unpin") ||
        !StrCmpIW(verb, L"startpin") || !StrCmpIW(verb, L"startunpin") ||
        !StrCmpIW(verb, L"pintostart") || !StrCmpIW(verb, L"unpinfromstart") ||
        !StrCmpIW(verb, L"pintostartscreen") || !StrCmpIW(verb, L"unpinfromstartscreen");
}

static void RemoveNativeStartPinCommands(HMENU menu, IContextMenu* pcm, UINT idFirst)
{
    for (int index = GetMenuItemCount(menu) - 1; index >= 0; --index)
    {
        HMENU submenu = GetSubMenu(menu, index);
        if (submenu)
            RemoveNativeStartPinCommands(submenu, pcm, idFirst);
        UINT id = GetMenuItemID(menu, index);
        if (id == UINT(-1) || id < idFirst)
            continue;
        WCHAR verb[MAX_PATH] = {};
        ContextMenu_GetCommandStringVerb(pcm, id - idFirst, verb, ARRAYSIZE(verb));
        if (IsStartPinVerb(verb))
            DeleteMenu(menu, index, MF_BYPOSITION);
    }
}

class CStartPinContextMenu : public IContextMenu3, public IObjectWithSite
{
public:
    CStartPinContextMenu(IContextMenu* nativeMenu, IContextMenu* pinMenu) :
        _refs(1), _native(nativeMenu), _pin(pinMenu), _native2(NULL),
        _native3(NULL), _site(NULL), _nativeFirst(0)
    {
        _native->QueryInterface(IID_PPV_ARGS(&_native2));
        _native->QueryInterface(IID_PPV_ARGS(&_native3));
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        if (!ppv)
            return E_POINTER;
        *ppv = NULL;
        if (riid == IID_IUnknown || riid == IID_IContextMenu ||
            riid == IID_IContextMenu2 || riid == IID_IContextMenu3)
            *ppv = static_cast<IContextMenu3*>(this);
        else if (riid == IID_IObjectWithSite)
            *ppv = static_cast<IObjectWithSite*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    STDMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_refs); }
    STDMETHODIMP_(ULONG) Release()
    {
        ULONG refs = InterlockedDecrement(&_refs);
        if (!refs)
            delete this;
        return refs;
    }
    STDMETHODIMP QueryContextMenu(HMENU menu, UINT index, UINT first, UINT last, UINT flags)
    {
        if (last < first || last - first < 2)
            return E_INVALIDARG;
        // The XP handler owns offsets 0 and 1, even when it adds no command.
        _nativeFirst = first + 2;
        HRESULT nativeResult = _native->QueryContextMenu(menu, index, _nativeFirst, last, flags);
        RemoveNativeStartPinCommands(menu, _native, _nativeFirst);
        if (_pin)
            _pin->QueryContextMenu(menu, index, first, first + 1, flags);
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0,
            2 + (SUCCEEDED(nativeResult) ? HRESULT_CODE(nativeResult) : 0));
    }
    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO info)
    {
        if (!info || info->cbSize < sizeof(CMINVOKECOMMANDINFO))
            return E_INVALIDARG;
        CMINVOKECOMMANDINFOEX command = {};
        const DWORD commandSize = min(info->cbSize, static_cast<DWORD>(sizeof(command)));
        CopyMemory(&command, info, commandSize);
        command.cbSize = commandSize;
        bool unicodeVerb = command.cbSize >= sizeof(command) &&
            (command.fMask & CMIC_MASK_UNICODE) && command.lpVerbW;
        bool pinCommand = false;
        if (unicodeVerb ? IS_INTRESOURCE(command.lpVerbW) : IS_INTRESOURCE(command.lpVerb))
        {
            UINT_PTR offset = unicodeVerb ? reinterpret_cast<UINT_PTR>(command.lpVerbW) :
                reinterpret_cast<UINT_PTR>(command.lpVerb);
            pinCommand = offset < 2;
            if (!pinCommand)
                offset -= 2;
            command.lpVerb = MAKEINTRESOURCEA(offset);
            command.lpVerbW = MAKEINTRESOURCEW(offset);
        }
        else
        {
            WCHAR verb[MAX_PATH] = {};
            if (unicodeVerb)
                StringCchCopyW(verb, ARRAYSIZE(verb), command.lpVerbW);
            else
                MultiByteToWideChar(CP_ACP, 0, info->lpVerb, -1, verb, ARRAYSIZE(verb));
            pinCommand = IsStartPinVerb(verb);
            if (pinCommand)
            {
                bool unpin = StrStrIW(verb, L"unpin") != NULL;
                command.lpVerb = MAKEINTRESOURCEA(unpin ? 1 : 0);
                command.lpVerbW = MAKEINTRESOURCEW(unpin ? 1 : 0);
            }
        }
        if (pinCommand)
            return _pin ? _pin->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&command)) : E_FAIL;
        return _native->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(&command));
    }
    STDMETHODIMP GetCommandString(UINT_PTR id, UINT type, UINT* reserved, LPSTR name, UINT count)
    {
        if (id < 2)
            return _pin ? _pin->GetCommandString(id, type, reserved, name, count) : E_INVALIDARG;
        return _native->GetCommandString(id - 2, type, reserved, name, count);
    }
    STDMETHODIMP HandleMenuMsg(UINT message, WPARAM wparam, LPARAM lparam)
    {
        LRESULT result = 0;
        return HandleMenuMsg2(message, wparam, lparam, &result);
    }
    STDMETHODIMP HandleMenuMsg2(UINT message, WPARAM wparam, LPARAM lparam, LRESULT* result)
    {
        if (result)
            *result = 0;
        HRESULT hr = _native3 ? _native3->HandleMenuMsg2(message, wparam, lparam, result) :
            (_native2 ? _native2->HandleMenuMsg(message, wparam, lparam) : E_NOTIMPL);
        if (message == WM_INITMENUPOPUP)
            RemoveNativeStartPinCommands(reinterpret_cast<HMENU>(wparam), _native, _nativeFirst);
        return hr;
    }
    STDMETHODIMP SetSite(IUnknown* site)
    {
        if (site)
            site->AddRef();
        if (_site)
            _site->Release();
        _site = site;
        IUnknown_SetSite(_native, site);
        if (_pin)
            IUnknown_SetSite(_pin, site);
        return S_OK;
    }
    STDMETHODIMP GetSite(REFIID riid, void** ppv)
    {
        if (!ppv)
            return E_POINTER;
        *ppv = NULL;
        return _site ? _site->QueryInterface(riid, ppv) : E_FAIL;
    }
private:
    ~CStartPinContextMenu()
    {
        if (_native3) _native3->Release();
        if (_native2) _native2->Release();
        _native->Release();
        if (_pin) _pin->Release();
        if (_site) _site->Release();
    }
    LONG _refs;
    IContextMenu* _native;
    IContextMenu* _pin;
    IContextMenu2* _native2;
    IContextMenu3* _native3;
    IUnknown* _site;
    UINT _nativeFirst;
};

HRESULT WrapStartMenuPinContextMenu(IContextMenu** ppcm, IDataObject* data)
{
    if (!ppcm || !*ppcm)
        return E_INVALIDARG;
    IContextMenu* pin = NULL;
    IShellExtInit* init = NULL;
    if (data && SUCCEEDED(CoCreateInstanceHook(CLSID_StartMenuPin, NULL,
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&init))))
    {
        if (SUCCEEDED(init->Initialize(NULL, data, NULL)))
            init->QueryInterface(IID_PPV_ARGS(&pin));
        init->Release();
    }
    CStartPinContextMenu* wrapper = new CStartPinContextMenu(*ppcm, pin);
    if (!wrapper)
    {
        if (pin)
            pin->Release();
        return E_OUTOFMEMORY;
    }
    *ppcm = wrapper;
    return S_OK;
}
