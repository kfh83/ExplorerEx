#include "pch.h"

#include "OpenBox.h"

#include "cabinet.h"
#include <propvarutil.h>
#include "SFTHost.h"

HRESULT COpenBoxHost::QueryInterface(REFIID riid, void** ppvObj)
{
	static const QITAB qit[] =
	{
		QITABENT(COpenBoxHost, IServiceProvider),
		QITABENT(COpenBoxHost, IOleCommandTarget),
		QITABENT(COpenBoxHost, IObjectWithSite),
		QITABENT(COpenBoxHost, IShellSearchTarget),
		QITABENT(COpenBoxHost, IInputObjectSite),
		QITABENT(COpenBoxHost, IAccessible),
		QITABENT(COpenBoxHost, IDispatch),
		QITABENT(COpenBoxHost, IEnumVARIANT),
		{ 0 },
	};

	return QISearch(this, qit, riid, ppvObj);
}

ULONG COpenBoxHost::AddRef()
{
	return InterlockedIncrement(&_lRef);
}

ULONG COpenBoxHost::Release()
{
	LONG lRef = InterlockedDecrement(&_lRef);
	if (lRef == 0 && this)
	{
		delete this;
	}
	return lRef;
}

HRESULT COpenBoxHost::get_accRole(VARIANT varChild, VARIANT* pvarRole)
{
	return CAccessible::get_accRole(varChild, pvarRole);
}

HRESULT COpenBoxHost::get_accState(VARIANT varChild, VARIANT* pvarState)
{
	return CAccessible::get_accState(varChild, pvarState);
}

HRESULT COpenBoxHost::get_accDefaultAction(VARIANT varChild, BSTR* pszDefAction)
{
	return CAccessible::GetRoleString(0xFFFFFCB6, pszDefAction);
}

HRESULT COpenBoxHost::accDoDefaultAction(VARIANT varChild)
{
	return CAccessible::accDoDefaultAction(varChild);
}

HRESULT COpenBoxHost::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
{
	if (IsEqualGUID(guidService, SID_SM_OpenBox))
	{
		HRESULT hr = QueryInterface(riid, ppvObject);
		if (SUCCEEDED(hr))
			return hr;
	}
	return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

HRESULT COpenBoxHost::SetSite(IUnknown* punkSite)
{
	CObjectWithSite::SetSite(punkSite);

	if (!punkSite && _pssc)
	{
		IUnknown_SetSite(_pssc, NULL);
		IUnknown_SafeReleaseAndNullPtr(&_pssc);
	}
	return S_OK;
}

HRESULT COpenBoxHost::QueryStatus(const GUID* pguidCmdGroup,
	ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pCmdText)
{
	return E_NOTIMPL;
}

HRESULT COpenBoxHost::Exec(const GUID *pguidCmdGroup,
    DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
    HRESULT hr = E_INVALIDARG;

    if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
    {
        switch (nCmdID)
        {
        case 300:
            return _UpdateSearch();
        case 306:
        {
            ASSERT(pvarargIn->vt == VT_BYREF); // 627
            HWND hwnd = (HWND)pvarargIn->byref;
            ASSERT(pvarargOut->vt == VT_BYREF); // 629
            if (SHIsChildOrSelf(_hwnd, hwnd) == S_OK)
                pvarargOut->byref = _hwnd;
            goto LABEL_32;
        }
        case 307:
            WCHAR szSearchQuery[MAX_PATH];
            _pssc->GetText(szSearchQuery, ARRAYSIZE(szSearchQuery));
            if (szSearchQuery[0])
            {
                _pssc->SetText(L"", SSCTEXT_TAKEFOCUS);
                pvarargOut->iVal = 0;
            }
            return 0;
        case 308:
            ASSERT(pvarargOut); // 618
            WCHAR szSearchQuery2[MAX_PATH];
            _pssc->GetText(szSearchQuery2, ARRAYSIZE(szSearchQuery2));
            return InitVariantFromString(szSearchQuery2, pvarargOut);
        case 315:
            ASSERT(pvarargIn->vt == VT_BYREF); // 608
            field_48 = 1;
            field_4C = 0;
            _pssc->SetText(pvarargIn->bstrVal, SSCTEXT_DEFAULT);
            field_48 = 0;
            return hr;
        case 319:
            OnMenuCommand(2);
            goto LABEL_32;
        case 320:
            OnMenuCommand(1);
            goto LABEL_32;
        case 321:
            OnMenuCommand(3);
            goto LABEL_32;
        case 322:
            OnMenuCommand(4);
        LABEL_32:
            hr = 0;
            break;
        default:
            return hr;
        }
    }
    return hr;
}

HRESULT COpenBoxHost::Search(LPCWSTR a2, DWORD a3)
{
    if (!a3)
    {
        // SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_WordWheel_Activated, 36);
        return E_NOTIMPL;
    }
}

LONG g_nOpenBoxCharThreadId;
FILETIME g_ftOpenBoxChar;

HRESULT COpenBoxHost::OnSearchTextNotify(LPCWSTR a2, LPCWSTR a3, SHELLSEARCHNOTIFY a4)
{
#if 1
    HRESULT hr = 0;

    if ((a4 & 9) != 0 && !field_48)
    {
        int v4 = field_4C;
        field_4C = 1;
        if (v4)
        {
            if (a2 && a2[0])
            {
                DWORD dwThreadId = GetCurrentThreadId();
                InterlockedExchange(&g_nOpenBoxCharThreadId, dwThreadId);
                GetSystemTimeAsFileTime(&g_ftOpenBoxChar);
                // SHTracePerf(&ShellTraceId_Explorer_StartPane_OpenBox_Char_Info);
            }
            hr = COpenBoxHost::_UpdateSearch();
        }
    }
    return hr;
#else
    HRESULT hr = S_OK;

    if ((a4 & (SSC_KEYPRESS | SSC_FORCE)) != 0 && !field_48)
    {
        if (a2 && a2[0])
        {
            // InterlockedExchange(&g_nOpenBoxCharThreadId, GetCurrentThreadId());
            // GetSystemTimeAsFileTime(&g_ftOpenBoxChar);
            // Skipped telemetry Explorer_StartPane_OpenBox_Char_Info
        }
        hr = _UpdateSearch();
    }

    return hr;
#endif
}

HRESULT COpenBoxHost::GetSearchText(LPWSTR pszText, UINT cchText)
{
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::GetPromptText(LPWSTR pszText, UINT cchText)
{
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::GetMenu(HMENU* phMenu)
{
	*phMenu = NULL;
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::InitMenuPopup(HMENU hhenu)
{
    // SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Dropdown_Count, 146);
    return S_OK;
}

HRESULT WWMenuCommand(LPCWSTR pszText, DWORD dwCmd, DWORD a3)
{
    return S_OK;
}

HRESULT COpenBoxHost::OnMenuCommand(DWORD dwCmd)
{
    WCHAR szText[MAX_PATH];
    HRESULT hr = _pssc->GetText(szText, ARRAYSIZE(szText));
    if (SUCCEEDED(hr))
    {
        hr = WWMenuCommand(szText, dwCmd, 0);
        if (SUCCEEDED(hr))
        {
            RECT rc;
            GetWindowRect(_hwnd, &rc);
            MapWindowRect(_hwnd, NULL, &rc);

            NMHDR nm;
            _SendNotify(GetParent(_hwnd), SMN_COMMANDINVOKED, &nm);
        }
    }
    return hr;
}

HRESULT COpenBoxHost::Enter(IShellSearchScope* pScope)
{
	return E_NOTIMPL;
}

HRESULT COpenBoxHost::Exit()
{
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::OnFocusChangeIS(IUnknown* punk, BOOL fSetFocus)
{
    VARIANT vt;
    vt.vt = VT_INT;
    vt.lVal = fSetFocus ? 0x20000 : 0;
    return IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 304, 0, &vt, NULL);
}

COpenBoxHost::COpenBoxHost(HWND hwnd)
	: _hwnd(hwnd)
    , _lRef(1)
    , field_4C(1)
{
}

COpenBoxHost::~COpenBoxHost()
{
    IUnknown_SafeReleaseAndNullPtr(&_pssc);
	CoTaskMemFree(_pszSearchQuery);
}

LRESULT COpenBoxHost::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    COpenBoxHost* self = reinterpret_cast<COpenBoxHost*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

    switch (uMsg)
    {
    case WM_CREATE:
        return self->_OnCreate(hwnd, uMsg, wParam, lParam);

    case WM_ERASEBKGND:
        return self->_OnEraseBkgnd(hwnd, uMsg, wParam, lParam);

    case WM_NOTIFY:
    {
        HRESULT hr = self->_OnNotify(hwnd, uMsg, wParam, lParam);
        if (SUCCEEDED(hr))
        {
            return hr;
        }
        break;
    }

    case WM_NCCREATE:
        return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);

    case WM_NCDESTROY:
        return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT COpenBoxHost::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	COpenBoxHost* pobh = new COpenBoxHost(hwnd);
    if (pobh)
    {
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pobh);
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

DEFINE_GUID(CLSID_SearchBox, 0x64BC32B5, 0x4EEC, 0x4DE7, 0x97, 0x2D, 0xBD, 0x8B, 0xD0, 0x32, 0x45, 0x37);

LRESULT COpenBoxHost::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPCREATESTRUCT lpcs = (LPCREATESTRUCT)lParam;
    SMPANEDATA* psmpd = (SMPANEDATA*)lpcs->lpCreateParams;

    IUnknown_Set(&psmpd->punk, SAFECAST(this, IServiceProvider*));

    MARGINS margins = {0};
    _hTheme = psmpd->hTheme;
    if (_hTheme)
    {
        GetThemeMargins(_hTheme, NULL, SPP_OPENBOX, 0, TMT_CONTENTMARGINS, 0, &margins);
    }
    else
    {
        _clrBk = GetSysColor(COLOR_MENU);

        margins.cxLeftWidth = 2 * GetSystemMetrics(SM_CXEDGE);
        margins.cxRightWidth = 2 * GetSystemMetrics(SM_CXEDGE);
        margins.cyTopHeight = 2 * GetSystemMetrics(SM_CYEDGE);
        margins.cyBottomHeight = lpcs->cy / 3;
    }

    SetWindowFont(_hwnd, GetWindowFont(GetParent(_hwnd)), FALSE);

    HRESULT hr = CoCreateInstance(/*CLSID_SearchControl*/CLSID_SearchBox, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pssc));
    if (SUCCEEDED(hr))
    {
        RECT rc;
        GetClientRect(_hwnd, &rc);

        rc.left += margins.cxLeftWidth;
        rc.top += margins.cyTopHeight;
        rc.right -= margins.cxRightWidth;
        rc.bottom -= margins.cyBottomHeight;

        hr = _pssc->Initialize(_hwnd, &rc);
        if (SUCCEEDED(hr))
        {
            // Build 6002 = SSCSTATE_DRAWCUETEXTFOCUS | SSCSTATE_APPENDTEXT | SSCSTATE_NODROPDOWN (0x2C)
            // Build 6519+ = SSCSTATE_DRAWCUETEXTFOCUS | SSCSTATE_APPENDTEXT | SSCSTATE_NODROPDOWN | SSCSTATE_NOSUGGESTIONS (0xAC)
            hr = _pssc->SetFlags(SSCSTATE_DRAWCUETEXTFOCUS | SSCSTATE_APPENDTEXT | SSCSTATE_NODROPDOWN);
            if (SUCCEEDED(hr))
            {
                WCHAR szText[MAX_PATH];
                LoadString(g_hinstCabinet, 8246, szText, ARRAYSIZE(szText));
                hr = _pssc->SetCueAndTooltipText(szText, 0);
                if (SUCCEEDED(hr))
                {
                    IUnknown_SetSite(_pssc, (IShellSearchControl*)this);
                    hr = IUnknown_GetWindow(_pssc, &_hwndEdit);
                }
            }
        }
    }
    return SUCCEEDED(hr) ? 0 : -1;
}

LRESULT COpenBoxHost::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
    LRESULT lRes = DefWindowProc(hwnd, uMsg, wParam, lParam);
    if (this)
        this->Release();
    return lRes;
}

LRESULT COpenBoxHost::_OnEraseBkgnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    RECT rc;
    GetClientRect(hwnd, &rc);
    if (_hTheme)
    {
        if (IsCompositionActive())
        {
            SHFillRectClr((HDC)wParam, &rc, 0);
        }
        DrawThemeBackground(_hTheme, (HDC)wParam, SPP_OPENBOX, 0, &rc, NULL);
    }
    else
    {
        SHFillRectClr((HDC)wParam, &rc, _clrBk);
        DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_TOP);
    }
    return 1;
}

LRESULT COpenBoxHost::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    HRESULT hr = E_FAIL;

    LPNMHDR pnmh = reinterpret_cast<LPNMHDR>(lParam);
    if (pnmh)
    {
        switch (pnmh->code)
        {
        case 215:
            return _OnSMNFindItem((SMNDIALOGMESSAGE *)pnmh);
        case 222:
            if (_pssc)
            {
                _pssc->SetText(L"", SSCTEXT_TAKEFOCUS);
            }
            break;
        case 223:
            return SetSite(((SMNSETSITE *)pnmh)->punkSite);
        }
    }
    return hr;
}

LRESULT COpenBoxHost::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
    pdm->flags |= 0x7800u;
    switch (pdm->flags & SMNDM_FINDMASK)
    {
        case SMNDM_FINDFIRSTMATCH:
        case SMNDM_FINDNEXTMATCH:
        case 8u:
        case 9u:
        case 0xAu:
        case 0xBu:
            return 0;
        case 2u:
            pdm->itemID = 0;
            return 1;
        case 3u:
        case 4u:
        case 7u:
            return _Mark(pdm, 1);
        case 5u:
        {
            WCHAR szQuery[MAX_PATH];
            _pssc->GetText(szQuery, ARRAYSIZE(szQuery));
            if (pdm->pmsg->wParam == VK_LEFT || pdm->pmsg->wParam == VK_RIGHT)
            {
                if (!wcslen(szQuery))
                {
                    RECT rc;
                    GetClientRect(_hwndEdit, &rc);
                    pdm->pt.x = (rc.left + rc.right) / 2;
                    pdm->pt.y = (rc.top + rc.bottom) / 2;
                    pdm->flags &= ~0x1000u;

                    if (pdm->pmsg->wParam == VK_LEFT)
                    {
                        pdm->flags |= 0x10000u;
                    }
                    return 0;
                }
            }

            if (pdm->pmsg->wParam != VK_UP && pdm->pmsg->wParam != VK_DOWN)
                return 0;

            VARIANT vt;
            vt.vt = VT_BYREF;
            vt.byref = pdm->pmsg;
            pdm->flags &= ~0x1000u;
            if (IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 324, 0, &vt, 0) < 0)
                return 0;
            pdm->flags &= ~0x100u;
            return 1;
        }
        case 6u:
        {
            if ((pdm->flags & 0x400) != 0)
            {
                pdm->flags = pdm->flags & 0xFFFFEFFF;
                IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 301, 0, 0, 0);
                //SHTracePerf(&ShellTraceId_Explorer_StartPane_OpenBox_Launch_Info);
            }
            return 0;
        }
        default:
        {
            ASSERT(!"Unknown SMNDM command"); // 384
            return 0;
        }
    }
}

BOOL COpenBoxHost::_Mark(PSMNDIALOGMESSAGE pdm, UINT a2)
{
    pdm->flags |= 0x8000;
    if (a2 == 1)
    {
        pdm->itemID = 1;
        pdm->flags |= 0x20000;
        if (_pssc)
        {
            IInputObject *pio = NULL;
            if (SUCCEEDED(_pssc->QueryInterface(IID_PPV_ARGS(&pio))))
            {
                pio->UIActivateIO(1, pdm->pmsg);
                pdm->flags &= ~0x100;
                return 1;
            }

            if (pio)
            {
                pio->Release();
            }
        }
    }
    return 0;
}

DWORD SHExpandEnvironmentStringsW(const WCHAR *pszIn, WCHAR *pszOut, DWORD cchOut);

HRESULT COpenBoxHost::_UpdateSearch()
{
    WCHAR pszStr1[260];
    HRESULT hr = this->_pssc->GetText(pszStr1, ARRAYSIZE(pszStr1));
    if (hr >= 0 && (!this->_pszSearchQuery || StrCmpC(pszStr1, this->_pszSearchQuery)))
    {
        CoTaskMemFree(this->_pszSearchQuery);
        SHStrDupW(pszStr1, &this->_pszSearchQuery);

        WCHAR pszPath[260];
        SHExpandEnvironmentStringsW(pszStr1, pszPath, 260);
        PathRemoveBlanksW(pszPath);

        VARIANTARG varg;
        hr = InitVariantFromString(pszPath, &varg);
        if (hr >= 0)
        {
            hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 300, 0, &varg, 0);
            IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 300, 0, 0, &varg);
            VariantClear(&varg);
        }
    }
    return hr;
}

BOOL OpenBoxHost_RegisterClass()
{
    WNDCLASSEX wcex;
	ZeroMemory(&wcex, sizeof(wcex));

    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_GLOBALCLASS;
    wcex.lpfnWndProc = (WNDPROC)COpenBoxHost::s_WndProc;
    wcex.hInstance = g_hinstCabinet;
    wcex.hbrBackground = 0;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.lpszClassName = WC_OPENBOXHOST;
    return RegisterClassEx(&wcex);
}