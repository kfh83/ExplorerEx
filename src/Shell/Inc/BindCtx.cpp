#include "pch.h"

#include "BindCtx.h"

#include "shundoc.h"

class CDummyUnknown
    : public IOleWindow
    , public IPersist
    , public IObjectWithFolderEnumMode
{
public:
    CDummyUnknown();
    CDummyUnknown(HWND hwnd);
    CDummyUnknown(FOLDER_ENUM_MODE feMode);
    CDummyUnknown(GUID clsid);

    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

    //~ Begin IOleWindow Interface
    STDMETHODIMP GetWindow(HWND* phwnd) override;
    STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode) override;
    //~ End IOleWindow Interface

    //~ Begin IPersist Interface
    STDMETHODIMP GetClassID(CLSID* pClassID) override;
    //~ End IPersist Interface

    //~ Begin IObjectWithFolderEnumMode Interface
    STDMETHODIMP SetMode(FOLDER_ENUM_MODE feMode) override;
    STDMETHODIMP GetMode(FOLDER_ENUM_MODE* pfeMode) override;
    //~ End IObjectWithFolderEnumMode Interface

protected:
    LONG _cRef;
    HWND _hwnd;
    GUID _clsid;
    FOLDER_ENUM_MODE _feMode;
};

CDummyUnknown::CDummyUnknown()
    : _cRef(1)
    , _hwnd(nullptr)
    , _clsid(GUID_NULL)
    , _feMode(FEM_VIEWRESULT)
{
}

CDummyUnknown::CDummyUnknown(HWND hwnd)
    : _cRef(1)
    , _hwnd(hwnd)
    , _clsid(GUID_NULL)
    , _feMode(FEM_VIEWRESULT)
{
}

CDummyUnknown::CDummyUnknown(FOLDER_ENUM_MODE feMode)
    : _cRef(1)
    , _hwnd(nullptr)
    , _clsid(GUID_NULL)
    , _feMode(feMode)
{
}

CDummyUnknown::CDummyUnknown(GUID clsid)
    : _cRef(1)
    , _hwnd(nullptr)
    , _clsid(clsid)
    , _feMode(FEM_VIEWRESULT)
{
}

HRESULT CDummyUnknown::QueryInterface(REFIID riid, void** ppvObject)
{
    static const QITAB qit[] = {
        QITABENT(CDummyUnknown, IOleWindow),
        QITABENT(CDummyUnknown, IPersist),
        QITABENT(CDummyUnknown, IObjectWithFolderEnumMode),
        {}
    };

    return QISearch(this, qit, riid, ppvObject);
}

ULONG CDummyUnknown::AddRef()
{
    return ++_cRef;
}

ULONG CDummyUnknown::Release()
{
    if (--_cRef <= 0)
    {
        delete this;
        return 0;
    }
    return _cRef;
}

HRESULT CDummyUnknown::GetWindow(HWND* phwnd)
{
    *phwnd = _hwnd;
    return _hwnd ? S_OK : E_NOTIMPL;
}

HRESULT CDummyUnknown::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

HRESULT CDummyUnknown::GetClassID(CLSID* pClassID)
{
    *pClassID = _clsid;
    return S_OK;
}

HRESULT CDummyUnknown::SetMode(FOLDER_ENUM_MODE feMode)
{
    return S_OK;
}

HRESULT CDummyUnknown::GetMode(FOLDER_ENUM_MODE* pfeMode)
{
    *pfeMode = _feMode;
    return S_OK;
}

STDAPI BindCtx_AddObjectParam(IBindCtx* pbc, const WCHAR* pszKey, IUnknown* punkValue)
{
    HRESULT hr;

    CDummyUnknown* dummy = nullptr;
    if (!punkValue)
    {
        dummy = new CDummyUnknown();
        hr = dummy ? S_OK : E_OUTOFMEMORY;
        if (FAILED(hr))
            return hr;
        punkValue = (IOleWindow*)dummy;
    }

    hr = pbc->RegisterObjectParam((LPOLESTR)pszKey, punkValue);

    IUnknown_SafeReleaseAndNullPtr(dummy);
    return hr;
}

STDAPI BindCtx_RegisterObjectParam(IBindCtx* pbc, const WCHAR* pszKey, IUnknown* punk, IBindCtx** ppbc)
{
    HRESULT hr;

    *ppbc = pbc;
    if (pbc)
    {
        pbc->AddRef();
    }
    else
    {
        hr = CreateBindCtx(0, ppbc);
        if (FAILED(hr))
            return hr;
    }

    hr = BindCtx_AddObjectParam(*ppbc, pszKey, punk);

    if (FAILED(hr))
        IUnknown_SafeReleaseAndNullPtr(*ppbc);

    return hr;
}

STDAPI BindCtx_RegisterUIWindow(IBindCtx* pbc, HWND hwnd, IBindCtx** ppbc)
{
    CDummyUnknown* dummy = new CDummyUnknown(hwnd);
    HRESULT hr = dummy ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = BindCtx_RegisterObjectParam(pbc, L"UI During Binding", (IOleWindow*)dummy, ppbc);
        dummy->Release();
    }
    else
    {
        *ppbc = nullptr;
    }
    return hr;
}

STDAPI BindCtx_SetMode(IBindCtx* pbcIn, DWORD grfMode, IBindCtx** ppbcOut)
{
    HRESULT hr = S_OK;

    *ppbcOut = pbcIn;
    if (pbcIn)
    {
        pbcIn->AddRef();
    }
    else
    {
        hr = CreateBindCtx(0, ppbcOut);
    }

    if (SUCCEEDED(hr))
    {
        BIND_OPTS bo;
        bo.cbStruct = sizeof(bo);
        bo.grfFlags = 0;
        bo.grfMode = grfMode;
        bo.dwTickCountDeadline = 0;
        if (pbcIn)
        {
            hr = pbcIn->GetBindOptions(&bo);
            bo.grfMode = grfMode;
        }

        if (SUCCEEDED(hr)) // @Note: This check is new in Windows 10
            hr = (*ppbcOut)->SetBindOptions(&bo);

        if (FAILED(hr))
            IUnknown_SafeReleaseAndNullPtr(*ppbcOut);
    }

    return hr;
}

STDAPI BindCtx_CreateWithMode(DWORD grfMode, IBindCtx** ppbcOut)
{
    return BindCtx_SetMode(nullptr, grfMode, ppbcOut);
}
