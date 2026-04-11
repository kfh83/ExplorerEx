#include "pch.h"
#include "shguidp.h"
#include "cabinet.h"
#include "shundoc.h"

///////////////////////////////////////////////////////////////////////////////////////
//
// class factory for explorer.exe
//
// These objects do not exist in the registry but rather are registered dynamically at
// runtime.  Since ClassFactory_Start is called on the the tray's thread, all objects
// will be registered on that thread.
//
///////////////////////////////////////////////////////////////////////////////////////

typedef HRESULT (*LPFNCREATEOBJINSTANCE)(IUnknown* pUnkOuter, IUnknown** ppunk);

class CDynamicClassFactory : public IClassFactory
{                                                                      
public:                                                                
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void **ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CDynamicClassFactory, IClassFactory),
            { 0 },
        };

        return QISearch(this, qit, riid, ppv);
    }

    STDMETHODIMP_(ULONG) AddRef() { return ++_cRef; }

    STDMETHODIMP_(ULONG) Release()
    {
        if (--_cRef > 0)
        {
            return _cRef;
        }
        delete this;
        return 0;
    }

    // *** IClassFactory ***
    STDMETHODIMP CreateInstance(IUnknown *punkOuter, REFIID riid, void **ppv)
    {
        *ppv = NULL;

        IUnknown *punk;
        HRESULT hr = _pfnCreate(punkOuter, &punk);
        if (SUCCEEDED(hr))
        {
            hr = punk->QueryInterface(riid, ppv);
            punk->Release();
        }

        return hr;
    }

    STDMETHODIMP LockServer(BOOL) { return S_OK; }

    // *** misc public methods ***
    HRESULT Register()
    {
        return CoRegisterClassObject(
            *_pclsid, this, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &_dwClassObject);
    }

    HRESULT Revoke()
    {
        HRESULT hr = CoRevokeClassObject(_dwClassObject);
        _dwClassObject = 0;
        return hr;
    }

    CDynamicClassFactory(CLSID const* pclsid, LPFNCREATEOBJINSTANCE pfnCreate) : _pclsid(pclsid),
                                                                    _pfnCreate(pfnCreate), _cRef(1) {}


private:

    CLSID const* _pclsid;
    LPFNCREATEOBJINSTANCE _pfnCreate;
    DWORD _dwClassObject;
    ULONG _cRef;
};

HRESULT CTaskBand_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk);
HRESULT CTrayBandSiteService_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk);
HRESULT CTrayNotifyStub_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk);

//these are defined in shell32startmnu.cpp
EXTERN_C HRESULT CStartMenu_CreateInstance(LPUNKNOWN punkOuter, REFIID riid, void** ppvOut);
EXTERN_C HRESULT CPersonalStartMenu_CreateInstance(LPUNKNOWN punkOuter, REFIID riid, void** ppvOut);
EXTERN_C HRESULT CProgramsFolderAndFastItems_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);

HRESULT CStartMenu_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk)
{
    return CStartMenu_CreateInstance(punkOuter,IID_PPV_ARGS(ppunk));
} 

HRESULT CPersonalStartMenu_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk)
{
    return CPersonalStartMenu_CreateInstance(punkOuter, IID_PPV_ARGS(ppunk));
}
HRESULT CStartMenuFolder_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk)
{
    return CStartMenuFolder_CreateInstance(punkOuter, IID_PPV_ARGS(ppunk));
}
HRESULT CProgramsFolder_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk)
{
    return CProgramsFolder_CreateInstance(punkOuter, IID_PPV_ARGS(ppunk));
}

HRESULT CProgramsFolderAndFastItems_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk)
{
    return CProgramsFolderAndFastItems_CreateInstance(punkOuter, IID_PPV_ARGS(ppunk));
}

static const struct
{
    CLSID const* pclsid;
    LPFNCREATEOBJINSTANCE pfnCreate;
}
c_ClassParams[] =
{
    { &CLSID_TaskBand,            CTaskBand_CreateInstance },
    { &CLSID_TrayBandSiteService, CTrayBandSiteService_CreateInstance },
    { &CLSID_TrayNotify,          CTrayNotifyStub_CreateInstance },

    // @Note: Custom to ExplorerEx, these would normally be instantiated inside shell32
    { &CLSID_StartMenu,                     CStartMenu_CreateInstance },
    { &CLSID_PersonalStartMenu,             CPersonalStartMenu_CreateInstance },
    { &CLSID_StartMenuFolder,               CStartMenuFolder_CreateInstance },
    { &CLSID_ProgramsFolder,                CProgramsFolder_CreateInstance },
    { &CLSID_ProgramsFolderAndFastItems,    CProgramsFolderAndFastItems_CreateInstance },
};

CDynamicClassFactory* g_rgpcf[ARRAYSIZE(c_ClassParams)] = {};


void ClassFactory_Start()
{
    for (int i = 0; i < ARRAYSIZE(c_ClassParams); i++)
    {
        g_rgpcf[i] = new CDynamicClassFactory(c_ClassParams[i].pclsid, c_ClassParams[i].pfnCreate);
        if (g_rgpcf[i])
        {
            g_rgpcf[i]->Register();
        }
    }
}

void ClassFactory_Stop()
{
    for (int i = 0; i < ARRAYSIZE(c_ClassParams); i++)
    {
        if (g_rgpcf[i])
        {
            g_rgpcf[i]->Revoke();

            g_rgpcf[i]->Release();
            g_rgpcf[i] = NULL;
        }
    }
}

const wchar_t pszCLSID[] = L"SOFTWARE\\Classes\\CLSID\\";

//https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/OneCodeTeam/UAC%20self-elevation%20(CppUACSelfElevation)/%5BC++%5D-UAC%20self-elevation%20(CppUACSelfElevation)/C++/CppUACSelfElevation/CppUACSelfElevation.cpp#L291
BOOL IsProcessElevated()
{
	BOOL fIsElevated = FALSE;
	DWORD dwError = ERROR_SUCCESS;
	HANDLE hToken = NULL;

	// Open the primary access token of the process with TOKEN_QUERY.
	if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
	{
		dwError = GetLastError();
		goto Cleanup;
	}

	// Retrieve token elevation information.
	TOKEN_ELEVATION elevation;
	DWORD dwSize;
	if (!GetTokenInformation(hToken, TokenElevation, &elevation,
		sizeof(elevation), &dwSize))
	{
		// When the process is run on operating systems prior to Windows 
		// Vista, GetTokenInformation returns FALSE with the 
		// ERROR_INVALID_PARAMETER error code because TokenElevation is 
		// not supported on those operating systems.
		dwError = GetLastError();
		goto Cleanup;
	}

	fIsElevated = elevation.TokenIsElevated;

Cleanup:
	// Centralized cleanup for all allocated resources.
	if (hToken)
	{
		CloseHandle(hToken);
		hToken = NULL;
	}

	// Throw the error if something failed in the function.
	if (ERROR_SUCCESS != dwError)
	{
		throw dwError;
	}

	return fIsElevated;
}

void SetupMergedFolderKeys(LPCTSTR clsid)
{
    WCHAR subKey[255];
    WCHAR subKeyClsid[255];
    WCHAR subKeyShellFolder[255];

    wcscpy_s(subKeyClsid, pszCLSID);
    wcsncat_s(subKeyClsid, clsid, 255);

    wcscpy_s(subKey, subKeyClsid);
    wcscpy_s(subKeyShellFolder, subKeyClsid);

    wcsncat_s(subKey, L"\\InProcServer32", 255);
    wcsncat_s(subKeyShellFolder, L"\\ShellFolder", 255);

    wprintf(L"subkey %s\n", subKey);
    wprintf(L"subKeyShellFolder %s\n", subKeyShellFolder);

    HKEY mainKeyToUse = HKEY_CURRENT_USER;
    if (IsProcessElevated())
        mainKeyToUse = HKEY_LOCAL_MACHINE;

    const WCHAR* pszDisplayName = nullptr;
    if (_wcsicmp(clsid, L"{865e5e76-ad83-4dca-a109-50dc2113ce9a}") == 0)
        pszDisplayName = L"Programs Folder and Fast Items";
    else if (_wcsicmp(clsid, L"{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}") == 0)
        pszDisplayName = L"Programs Folder";
    else if (_wcsicmp(clsid, L"{48e7caab-b918-4e58-a94d-505519c795dc}") == 0)
        pszDisplayName = L"Start Menu Folder";

    HKEY res;
    bool bHasKey = RegOpenKeyW(mainKeyToUse, subKeyClsid, &res) == S_OK;
    if (!bHasKey)
    {
        bHasKey = RegCreateKeyW(mainKeyToUse, subKeyClsid, &res) == S_OK;
    }

    if (!bHasKey)
    {
        wprintf(L"FAILED TO CREATE THE KEY!!!!\n");
        return;
    }

    if (pszDisplayName)
    {
        bool bWrote = RegSetValueExW(res, nullptr, 0, REG_SZ,
            (const BYTE*)pszDisplayName, (DWORD)((wcslen(pszDisplayName) + 1) * sizeof(WCHAR))) == S_OK;
        if (!bWrote)
            wprintf(L"FAILED TO WRITE CLSID DISPLAY NAME!!");
    }

    RegCloseKey(res);

    bHasKey = RegOpenKeyW(mainKeyToUse, subKey, &res) == S_OK;
    if (!bHasKey)
    {
        bHasKey = RegCreateKeyW(mainKeyToUse, subKey, &res) == S_OK;
    }

    if (!bHasKey)
    {
        wprintf(L"FAILED TO CREATE THE KEY!!!!\n");
        return;
    }

	WCHAR localExePath[MAX_PATH];
	GetModuleFileNameW(0, localExePath, MAX_PATH);

    wprintf(L"localExePath %s\n", localExePath);

    WCHAR exePath[MAX_PATH];
    DWORD size = sizeof(exePath);
    bool bRead = RegQueryValueExW(res, 0, 0, 0, (LPBYTE)exePath, &size) == S_OK;
    if (bRead)
    {
        //there is something different there!
        if (wcscmp(exePath, localExePath) != 0)
        {
            wprintf(L"WARNING!!! ALREADY A KEY OF A DIFFERENT EXE THERE!! OVERWRITING!\n");
            //return;
        }
    }

    //otherwise we overwrite
    bool bWrote = RegSetValueExW(res, 0, 0, REG_SZ, (LPBYTE)localExePath, sizeof(localExePath)) == S_OK;
    if (!bWrote)
        wprintf(L"FAILED TO WRITE!!");

    RegCloseKey(res);

	bHasKey = RegOpenKeyW(mainKeyToUse, subKeyShellFolder, &res) == S_OK;
	if (!bHasKey)
	{
		bHasKey = RegCreateKeyW(mainKeyToUse, subKeyShellFolder, &res) == S_OK;
	}

    // CLSID_ProgramsFolder = SFGAO_NONENUMERATED | SFGAO_BROWSABLE | SFGAO_FOLDER;
    // CLSID_ProgramsFolderAndFastItems = SFGAO_NONENUMERATED | SFGAO_BROWSABLE | SFGAO_FOLDER;
    DWORD attributes = SFGAO_NONENUMERATED | SFGAO_BROWSABLE | SFGAO_FOLDER;
    bWrote = RegSetValueExW(res, L"Attributes", 0, REG_DWORD, (LPBYTE)&attributes, sizeof(DWORD)) == S_OK;
    if (!bWrote)
        wprintf(L"FAILED TO WRITE!!");

    RegCloseKey(res);
}

void ComServer_Stop(LPCTSTR clsid)
{

}

