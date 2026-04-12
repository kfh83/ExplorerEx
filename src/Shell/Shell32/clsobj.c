#define _IOffset(class, itf)         ((UINT)(UINT_PTR)&(((class *)0)->itf))
#define IToClass(class, itf, pitf)   ((class  *)(((LPSTR)pitf)-_IOffset(class, itf)))

#include <objbase.h>
#include <Unknwn.h>

#define DllAddRef()
#define DllRelease()    

extern const CLSID CLSID_StartMenu;
extern const CLSID CLSID_PersonalStartMenu;
extern const CLSID CLSID_StartMenuFolder;
extern const CLSID CLSID_ProgramsFolder;
extern const CLSID CLSID_ProgramsFolderAndFastItems;

extern HRESULT CStartMenu_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
extern HRESULT CPersonalStartMenu_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
extern HRESULT CStartMenuFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
extern HRESULT CProgramsFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
extern HRESULT CProgramsFolderAndFastItems_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);

typedef struct
{
    const IClassFactoryVtbl* cf;
    REFCLSID rclsid;
    HRESULT(*pfnCreate)(IUnknown*, REFIID, void**);
    ULONG flags;
} OBJ_ENTRY;

#define OBJ_AGGREGATABLE 1

extern const IClassFactoryVtbl c_CFVtbl;

const OBJ_ENTRY c_clsmap[] =
{
    { &c_CFVtbl, &CLSID_StartMenu, CStartMenu_CreateInstance, 0 },
    { &c_CFVtbl, &CLSID_PersonalStartMenu, CPersonalStartMenu_CreateInstance, 0 },
    { &c_CFVtbl, &CLSID_StartMenuFolder, CStartMenuFolder_CreateInstance, 0 },
    { &c_CFVtbl, &CLSID_ProgramsFolder, CProgramsFolder_CreateInstance, 0 },
    { &c_CFVtbl, &CLSID_ProgramsFolderAndFastItems, CProgramsFolderAndFastItems_CreateInstance, 0 },
    { NULL, NULL, NULL, 0 }
};

STDMETHODIMP CCF_QueryInterface(IClassFactory* pcf, REFIID riid, void** ppvObj)
{
    // OBJ_ENTRY *this = IToClass(OBJ_ENTRY, cf, pcf);
    if (IsEqualIID(riid, &IID_IClassFactory) || IsEqualIID(riid, &IID_IUnknown))
    {
        *ppvObj = (void*)pcf;
        DllAddRef();
        return NOERROR;
    }

    *ppvObj = NULL;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CCF_AddRef(IClassFactory* pcf)
{
    DllAddRef();
    return 2;
}

STDMETHODIMP_(ULONG) CCF_Release(IClassFactory* pcf)
{
    DllRelease();
    return 1;
}

STDMETHODIMP CCF_CreateInstance(IClassFactory* pcf, IUnknown* punkOuter, REFIID riid, void** ppvObject)
{
    OBJ_ENTRY* this = IToClass(OBJ_ENTRY, cf, pcf);

    *ppvObject = NULL; // to avoid nulling it out in every create function...

    if (punkOuter && !(this->flags & OBJ_AGGREGATABLE))
        return CLASS_E_NOAGGREGATION;

    return this->pfnCreate(punkOuter, riid, ppvObject);
}

STDMETHODIMP CCF_LockServer(IClassFactory* pcf, BOOL fLock)
{
    /*  SHELL32.DLL does not implement DllCanUnloadNow, thus does not have a DLL refcount
        This means we can never unload!

        if (fLock)
            DllAddRef();
        else
            DllRelease();
    */
    return S_OK;
}

const IClassFactoryVtbl c_CFVtbl = {
    CCF_QueryInterface, CCF_AddRef, CCF_Release,
    CCF_CreateInstance,
    CCF_LockServer
};

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (IsEqualIID(riid, &IID_IClassFactory) || IsEqualIID(riid, &IID_IUnknown))
    {
        const OBJ_ENTRY* pcls;
        for (pcls = c_clsmap; pcls->rclsid; pcls++)
        {
            if (IsEqualIID(rclsid, pcls->rclsid))
            {
                *ppv = (void*)&(pcls->cf);
                DllAddRef();
                return NOERROR;
            }
        }
    }

    *ppv = NULL;
    return CLASS_E_CLASSNOTAVAILABLE;
}