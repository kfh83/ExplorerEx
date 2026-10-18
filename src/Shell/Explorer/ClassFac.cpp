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
        return CoRegisterClassObject(*_pclsid, this, CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &_dwClassObject);
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

CDynamicClassFactory* g_rgpcf[ARRAYSIZE(c_ClassParams)] = {0};


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

static HRESULT SetClassValue(HKEY root, REFCLSID clsid, PCWSTR pszSubKey, PCWSTR pszName,
                             DWORD dwType, const void* pvData, DWORD cbData)
{
    WCHAR szClsid[39];
    WCHAR szKey[MAX_PATH];
    StringFromGUID2(clsid, szClsid, ARRAYSIZE(szClsid));
    HRESULT hr = StringCchPrintfW(szKey, ARRAYSIZE(szKey),
        L"Software\\Classes\\CLSID\\%s%s", szClsid, pszSubKey);
    if (FAILED(hr))
        return hr;

    HKEY hkey;
    LSTATUS status = RegCreateKeyExW(root, szKey, 0, nullptr,
        REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, nullptr, &hkey, nullptr);
    if (status == ERROR_SUCCESS)
    {
        status = RegSetValueExW(hkey, pszName, 0, dwType,
            static_cast<const BYTE*>(pvData), cbData);
        RegCloseKey(hkey);
    }
    return HRESULT_FROM_WIN32(status);
}

static HRESULT SetClassString(HKEY root, REFCLSID clsid, PCWSTR pszSubKey, PCWSTR pszName,
                              PCWSTR pszValue, DWORD dwType = REG_SZ)
{
    return SetClassValue(root, clsid, pszSubKey, pszName, dwType, pszValue,
        (lstrlenW(pszValue) + 1) * sizeof(WCHAR));
}

HRESULT InitializeRegistryKeys()
{
    // hinstCabinet is the Explorer module, including when hosted as a DLL.
    WCHAR szServer[MAX_PATH];
    DWORD cch = GetModuleFileNameW(hinstCabinet, szServer, ARRAYSIZE(szServer));
    if (!cch)
        return HRESULT_FROM_WIN32(GetLastError());
    if (cch >= ARRAYSIZE(szServer))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    if (!PathRemoveFileSpecW(szServer) || !PathAppendW(szServer, L"ExplorerEx.Shell32.dll"))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    DWORD attributes = GetFileAttributesW(szServer);
    if (attributes == INVALID_FILE_ATTRIBUTES)
        return HRESULT_FROM_WIN32(GetLastError());
    if (attributes & FILE_ATTRIBUTE_DIRECTORY)
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);

    HANDLE token;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token))
        return HRESULT_FROM_WIN32(GetLastError());
    TOKEN_ELEVATION elevation = {};
    DWORD cbElevation;
    BOOL gotElevation = GetTokenInformation(token, TokenElevation, &elevation,
        sizeof(elevation), &cbElevation);
    DWORD tokenError = gotElevation ? ERROR_SUCCESS : GetLastError();
    CloseHandle(token);
    if (tokenError != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(tokenError);

    // Shell namespace binding in an elevated process ignores per-user servers.
    // Only our private folder CLSIDs go into HKLM; never replace Windows-owned
    // StartMenu/PersonalStartMenu machine registrations.
    HKEY folderRoot = elevation.TokenIsElevated ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;

    const DWORD dwBaseAttributes = SFGAO_BROWSABLE | SFGAO_FOLDER;
    const DWORD dwProgramsAttributes = SFGAO_NONENUMERATED | dwBaseAttributes;
    const struct
    {
        const CLSID* pclsid;
        PCWSTR pszDescription;
        DWORD dwAttributes;
        bool fMergedLocation;
    } classes[] = {
        { &CLSID_StartMenu, L"Start Menu", 0, true },
        { &CLSID_PersonalStartMenu, L"Personal Start Menu", 0, false },
        { &CLSID_StartMenuFolder, L"Start Menu Folder", dwBaseAttributes, false },
        { &CLSID_ProgramsFolder, L"Programs Folder", dwProgramsAttributes, true },
        { &CLSID_ProgramsFolderAndFastItems, L"Programs Folder and Fast Items", dwProgramsAttributes, false },
    };
    HRESULT hr = S_OK;
    for (const auto& entry : classes)
    {
        HKEY root = entry.dwAttributes ? folderRoot : HKEY_CURRENT_USER;
        hr = SetClassString(root, *entry.pclsid, L"", nullptr, entry.pszDescription);
        if (SUCCEEDED(hr))
            hr = SetClassString(root, *entry.pclsid, L"\\InprocServer32", nullptr, szServer);
        if (SUCCEEDED(hr))
            hr = SetClassString(root, *entry.pclsid, L"\\InprocServer32", L"ThreadingModel", L"Apartment");
        if (SUCCEEDED(hr) && entry.dwAttributes)
            hr = SetClassValue(root, *entry.pclsid, L"\\ShellFolder", L"Attributes",
                REG_DWORD, &entry.dwAttributes, sizeof(entry.dwAttributes));
        if (SUCCEEDED(hr) && entry.fMergedLocation)
            hr = SetClassString(root, *entry.pclsid, L"\\MergedFolder", L"Location", L"@shell32.dll,-4177");
        if (FAILED(hr))
            return hr;
    }

    hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"", L"LocalizedString",
        L"@%SystemRoot%\\system32\\shell32.dll,-21786", REG_EXPAND_SZ);
    if (SUCCEEDED(hr))
        hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find", L"LegacyDisable", L"");
    const DWORD dwSuppressionPolicy = 0x80;
    if (SUCCEEDED(hr))
        hr = SetClassValue(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find", L"SuppressionPolicy",
            REG_DWORD, &dwSuppressionPolicy, sizeof(dwSuppressionPolicy));
    if (SUCCEEDED(hr))
        hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find\\command", nullptr,
            L"%SystemRoot%\\Explorer.exe", REG_EXPAND_SZ);
    if (SUCCEEDED(hr))
        hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find\\ddeexec", nullptr,
            L"[FindFolder(\"%l\", %I)]");
    if (SUCCEEDED(hr))
        hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find\\ddeexec\\application", nullptr, L"Folders");
    if (SUCCEEDED(hr))
        hr = SetClassString(folderRoot, CLSID_StartMenuFolder, L"\\shell\\find\\ddeexec\\topic", nullptr, L"AppProperties");
    return hr;
}
