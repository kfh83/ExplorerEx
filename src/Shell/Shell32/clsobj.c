#include <objbase.h>
#include <unknwn.h>
#include "Shell32Util.h"

EXTERN_C const CLSID CLSID_StartMenu;
EXTERN_C const CLSID CLSID_PersonalStartMenu;
EXTERN_C const CLSID CLSID_StartMenuFolder;
EXTERN_C const CLSID CLSID_ProgramsFolder;
EXTERN_C const CLSID CLSID_ProgramsFolderAndFastItems;

EXTERN_C HRESULT CStartMenu_CreateInstance(IUnknown*, REFIID, void**);
EXTERN_C HRESULT CPersonalStartMenu_CreateInstance(IUnknown*, REFIID, void**);
EXTERN_C HRESULT CStartMenuFolder_CreateInstance(IUnknown*, REFIID, void**);
EXTERN_C HRESULT CProgramsFolder_CreateInstance(IUnknown*, REFIID, void**);
EXTERN_C HRESULT CProgramsFolderAndFastItems_CreateInstance(IUnknown*, REFIID, void**);

typedef struct
{
    const IClassFactoryVtbl* lpVtbl;
    const CLSID* pclsid;
    HRESULT (*pfnCreate)(IUnknown*, REFIID, void**);
} CLASS_ENTRY;

static HRESULT STDMETHODCALLTYPE FactoryQueryInterface(IClassFactory* factory, REFIID riid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;
    if (!IsEqualIID(riid, &IID_IUnknown) && !IsEqualIID(riid, &IID_IClassFactory))
        return E_NOINTERFACE;
    *ppv = factory;
    return S_OK;
}

// Like the system shell server, this module remains loaded for the process lifetime.
// Its static factories and menu objects therefore require no unload refcount.
static ULONG STDMETHODCALLTYPE FactoryAddRef(IClassFactory* factory)
{
    return 2;
}

static ULONG STDMETHODCALLTYPE FactoryRelease(IClassFactory* factory)
{
    return 1;
}

static HRESULT STDMETHODCALLTYPE FactoryCreateInstance(IClassFactory* factory, IUnknown* outer,
                                                       REFIID riid, void** ppv)
{
    const CLASS_ENTRY* entry = (const CLASS_ENTRY*)factory;
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;
    if (outer)
        return CLASS_E_NOAGGREGATION;
    return entry->pfnCreate(outer, riid, ppv);
}

static HRESULT STDMETHODCALLTYPE FactoryLockServer(IClassFactory* factory, BOOL lock)
{
    return S_OK;
}

static const IClassFactoryVtbl factoryVtbl = {
    FactoryQueryInterface, FactoryAddRef, FactoryRelease, FactoryCreateInstance, FactoryLockServer
};

static const CLASS_ENTRY classes[] = {
    { &factoryVtbl, &CLSID_StartMenu, CStartMenu_CreateInstance },
    { &factoryVtbl, &CLSID_PersonalStartMenu, CPersonalStartMenu_CreateInstance },
    { &factoryVtbl, &CLSID_StartMenuFolder, CStartMenuFolder_CreateInstance },
    { &factoryVtbl, &CLSID_ProgramsFolder, CProgramsFolder_CreateInstance },
    { &factoryVtbl, &CLSID_ProgramsFolderAndFastItems, CProgramsFolderAndFastItems_CreateInstance }
};

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    unsigned int i;
    if (!ppv)
        return E_POINTER;
    *ppv = NULL;
    for (i = 0; i < sizeof(classes) / sizeof(classes[0]); ++i)
    {
        if (IsEqualCLSID(rclsid, classes[i].pclsid))
        {
            if (!ExplorerExShell32_EnsureRuntimeInit())
                return CO_E_ERRORINDLL;
            return FactoryQueryInterface((IClassFactory*)&classes[i], riid, ppv);
        }
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}
