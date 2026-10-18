#pragma once

#include "ShUndoc.h"

class CPinnedListWrapper : public IStartMenuPin
{
public:
    CPinnedListWrapper(IUnknown* punk, const GUID& riid);
    ~CPinnedListWrapper();

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

    STDMETHODIMP EnumObjects(IEnumFullIDList**);
    STDMETHODIMP Modify(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE);
    STDMETHODIMP GetChangeCount(ULONG*);
    STDMETHODIMP GetPinnableInfo(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*);
    STDMETHODIMP IsPinnable(IDataObject*, int);
    STDMETHODIMP Resolve(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*);
    STDMETHODIMP IsPinned(PCIDLIST_ABSOLUTE);
    STDMETHODIMP GetPinnedItem(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*);
    STDMETHODIMP GetAppIDForPinnedItem(PCIDLIST_ABSOLUTE, PWSTR*);
    STDMETHODIMP ItemChangeNotify(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE);
    STDMETHODIMP UpdateForRemovedItemsAsNecessary(VOID);

private:
    IFlexibleTaskbarPinnedList* m_flexList = nullptr;
    IPinnedList3* m_pinnedList3 = nullptr;
    IPinnedList25* m_pinnedList25 = nullptr;
};
