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

typedef struct
{
    HKEY hRootKey;
    const WCHAR* pszSubKey;
    const WCHAR* pszClassID;
    const WCHAR* pszValueName;
    BYTE* pszData;
    DWORD dwType;
} REGSTRUCT;

HRESULT InitializeRegistryKeys()
{
    WCHAR szPersonalStartMenuCLSID[64];
    StringFromGUID2(CLSID_PersonalStartMenu, szPersonalStartMenuCLSID, ARRAYSIZE(szPersonalStartMenuCLSID));

    WCHAR szStartMenuCLSID[64];
    StringFromGUID2(CLSID_StartMenu, szStartMenuCLSID, ARRAYSIZE(szStartMenuCLSID));

    WCHAR szStartMenuFolderCLSID[64];
    StringFromGUID2(CLSID_StartMenuFolder , szStartMenuFolderCLSID, ARRAYSIZE(szStartMenuFolderCLSID));

    WCHAR szProgramsFolderCLSID[64];
    StringFromGUID2(CLSID_ProgramsFolder, szProgramsFolderCLSID, ARRAYSIZE(szProgramsFolderCLSID));

    WCHAR szProgramsFolderAndFastItemsCLSID[64];
    StringFromGUID2(
        CLSID_ProgramsFolderAndFastItems, szProgramsFolderAndFastItemsCLSID,
        ARRAYSIZE(szProgramsFolderAndFastItemsCLSID));

    WCHAR szModulePathAndName[260];
    GetModuleFileNameW(GetModuleHandleW(L"ExplorerEx.Shell32.dll"), szModulePathAndName, ARRAYSIZE(szModulePathAndName));

    const DWORD dwSuppressionPolicy = 0x80;
    const SFGAOF dwBaseAttributes = SFGAO_BROWSABLE | SFGAO_FOLDER;
    const SFGAOF dwProgramsFolderAttributes = SFGAO_NONENUMERATED | dwBaseAttributes;

    const REGSTRUCT c_rgRegistryEntries[]
    {
        // CLSID_PersonalStartMenu
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szPersonalStartMenuCLSID, nullptr, (BYTE*)L"Personal Start Menu", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szPersonalStartMenuCLSID, nullptr, (BYTE*)L"%s", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szPersonalStartMenuCLSID, L"ThreadingModel", (BYTE*)L"Apartment", REG_SZ },

        // CLSID_StartMenu
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szStartMenuCLSID, nullptr, (BYTE*)L"Start Menu", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InProcServer32", szStartMenuCLSID, nullptr, (BYTE*)L"%s", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InProcServer32", szStartMenuCLSID, L"ThreadingModel", (BYTE*)L"Apartment", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\MergedFolder", szStartMenuCLSID, L"Location", (BYTE*)L"@shell32.dll,-4177", REG_SZ },

        // CLSID_StartMenuFolder
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szStartMenuFolderCLSID, nullptr, (BYTE*)L"Start Menu Folder", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szStartMenuFolderCLSID, L"LocalizedString", (BYTE*)L"@%SystemRoot%\\system32\\shell32.dll,-21786", REG_EXPAND_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szStartMenuFolderCLSID, nullptr, (BYTE*)L"%s", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szStartMenuFolderCLSID, L"ThreadingModel", (BYTE*)L"Apartment", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find", szStartMenuFolderCLSID, L"LegacyDisable", (BYTE*)L"", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find", szStartMenuFolderCLSID, L"SuppressionPolicy", (BYTE*)&dwSuppressionPolicy, REG_DWORD },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find\\command", szStartMenuFolderCLSID, nullptr, (BYTE*)L"%SystemRoot%\\Explorer.exe", REG_EXPAND_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find\\ddeexec", szStartMenuFolderCLSID, nullptr, (BYTE*)L"[FindFolder(\"%%l\", %%I)]", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find\\ddeexec\\application", szStartMenuFolderCLSID, nullptr, (BYTE*)L"Folders", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\shell\\find\\ddeexec\\topic", szStartMenuFolderCLSID, nullptr, (BYTE*)L"AppProperties", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\ShellFolder", szStartMenuFolderCLSID, L"Attributes", (BYTE*)&dwBaseAttributes, REG_DWORD },

        // CLSID_ProgramsFolder
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szProgramsFolderCLSID, nullptr, (BYTE*)L"Programs Folder", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szProgramsFolderCLSID, nullptr, (BYTE*)L"%s", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szProgramsFolderCLSID, L"ThreadingModel", (BYTE*)L"Apartment", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\MergedFolder", szProgramsFolderCLSID, L"Location", (BYTE*)L"@shell32.dll,-4177", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\ShellFolder", szProgramsFolderCLSID, L"Attributes", (BYTE*)&dwProgramsFolderAttributes, REG_DWORD },

        // CLSID_ProgramsFolderAndFastItems
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s", szProgramsFolderAndFastItemsCLSID, nullptr, (BYTE*)L"Programs Folder and Fast Items", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szProgramsFolderAndFastItemsCLSID, nullptr, (BYTE*)L"%s", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\InprocServer32", szProgramsFolderAndFastItemsCLSID, L"ThreadingModel", (BYTE*)L"Apartment", REG_SZ },
        { HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\%s\\ShellFolder", szProgramsFolderAndFastItemsCLSID, L"Attributes", (BYTE*)&dwProgramsFolderAttributes, REG_DWORD },
    };

    WCHAR szSubKey[260];
    WCHAR szData[260];
    HKEY hkey = nullptr;
    HRESULT hr = S_OK;

    for (UINT i = 0; SUCCEEDED(hr) && i < ARRAYSIZE(c_rgRegistryEntries); ++i)
    {
        hr = StringCchPrintfW(
            szSubKey, ARRAYSIZE(szSubKey), c_rgRegistryEntries[i].pszSubKey, c_rgRegistryEntries[i].pszClassID);
        if (SUCCEEDED(hr))
        {
            LRESULT lres = RegCreateKeyExW(
                c_rgRegistryEntries[i].hRootKey, szSubKey, 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr,
                &hkey, nullptr);
            if (lres == ERROR_SUCCESS)
            {
                switch (c_rgRegistryEntries[i].dwType)
                {
                    case REG_SZ:
                    {
                        hr = StringCchPrintfW(
                            szData, ARRAYSIZE(szData), (WCHAR*)c_rgRegistryEntries[i].pszData, szModulePathAndName);
                        if (SUCCEEDED(hr))
                        {
                            RegSetValueExW(
                                hkey, c_rgRegistryEntries[i].pszValueName, 0, c_rgRegistryEntries[i].dwType,
                                (BYTE*)szData, (lstrlenW(szData) + 1) * sizeof(WCHAR));
                        }
                        break;
                    }
                    case REG_EXPAND_SZ:
                    {
                        RegSetValueExW(
                            hkey, c_rgRegistryEntries[i].pszValueName, 0, c_rgRegistryEntries[i].dwType,
                            c_rgRegistryEntries[i].pszData,
                            (lstrlenW((WCHAR*)c_rgRegistryEntries[i].pszData) + 1) * sizeof(WCHAR));
                        break;
                    }
                    case REG_DWORD:
                    {
                        RegSetValueExW(
                            hkey, c_rgRegistryEntries[i].pszValueName, 0, c_rgRegistryEntries[i].dwType,
                            c_rgRegistryEntries[i].pszData, sizeof(DWORD));
                        break;
                    }
                }

                RegCloseKey(hkey);
            }
            else
            {
                hr = SELFREG_E_CLASS;
            }
        }
    }

    return hr;
}

void ComServer_Stop(LPCTSTR clsid)
{

}

