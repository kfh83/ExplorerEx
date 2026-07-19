#include "pch.h"

#include "OpenBox.h"

#include "cabinet.h"
#include "SFTHost.h"
#include "shstr.h"

#include <propvarutil.h>

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
        {},
    };

    return QISearch(this, qit, riid, ppvObj);
}

ULONG COpenBoxHost::AddRef()
{
    return InterlockedIncrement(&_lRef);
}

ULONG COpenBoxHost::Release()
{
    _ASSERTE(0 != _lRef);

    LONG lRef = InterlockedDecrement(&_lRef);
    if (lRef == 0)
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
    return GetRoleString(0xFFFFFCB6, pszDefAction);
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
        IUnknown_SetSite(_pssc, nullptr);
        IUnknown_SafeReleaseAndNullPtr(_pssc);
    }
    return S_OK;
}

HRESULT COpenBoxHost::QueryStatus(const GUID* pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pCmdText)
{
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::Exec(
    const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvarargIn, VARIANT* pvarargOut)
{
    HRESULT hr = E_INVALIDARG;

    if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
    {
        switch (nCmdID)
        {
            case 300:
            {
                return _UpdateSearch();
            }
            case 306:
            {
                _ASSERTE(pvarargIn->vt == VT_BYREF); // 627
                HWND hwnd = static_cast<HWND>(pvarargIn->byref);
                _ASSERTE(pvarargOut->vt == VT_BYREF); // 629
                if (SHIsChildOrSelf(_hwnd, hwnd) == S_OK)
                {
                    pvarargOut->byref = _hwnd;
                }
                hr = S_OK;
                break;
            }
            case 307:
            {
                WCHAR szSearchQuery[260];
                _pssc->GetText(szSearchQuery, ARRAYSIZE(szSearchQuery));
                if (szSearchQuery[0])
                {
                    // @MOD: SSCTEXT_FORCE taken from build 6519
                    _pssc->SetText(L"", SSCTEXT_TAKEFOCUS | SSCTEXT_FORCE);
                    pvarargOut->boolVal = VARIANT_FALSE;
                }
                return S_OK;
            }
            case 308:
            {
                ASSERT(pvarargOut); // 618
                WCHAR szSearchQuery[260];
                _pssc->GetText(szSearchQuery, ARRAYSIZE(szSearchQuery));
                return InitVariantFromString(szSearchQuery, pvarargOut);
            }
            case 315:
            {
                ASSERT(pvarargIn->vt == VT_BYREF); // 608
                _fInSetText = TRUE;
                field_4C = 0;
                // @MOD: SSCTEXT_FORCE taken from build 6519
                _pssc->SetText((const WCHAR*)pvarargIn->byref, SSCTEXT_DEFAULT | SSCTEXT_FORCE);
                _fInSetText = FALSE;
                return hr;
            }
            case 319:
                OnMenuCommand(2);
                hr = S_OK;
                break;
            case 320:
                OnMenuCommand(1);
                hr = S_OK;
                break;
            case 321:
                OnMenuCommand(3);
                hr = S_OK;
                break;
            case 322:
                OnMenuCommand(4);
                hr = S_OK;
                break;
            default:
                break;
        }
    }

    return hr;
}

HRESULT COpenBoxHost::Search(const WCHAR* pszSearchText, DWORD dwCommand)
{
    if (dwCommand == 0)
    {
        //SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_WordWheel_Activated, 548);
    }
    return E_NOTIMPL;
}

LONG g_nOpenBoxCharThreadId;
FILETIME g_ftOpenBoxChar;

HRESULT COpenBoxHost::OnSearchTextNotify(const WCHAR* pszSearchText, const WCHAR* pszStrippedText, SHELLSEARCHNOTIFY sn)
{
    HRESULT hr = S_OK;

    if ((sn & (SSC_KEYPRESS | SSC_FORCE)) != 0 && !_fInSetText)
    {
        int v4 = field_4C;
        field_4C = 1;
        if (v4)
        {
            if (pszSearchText && pszSearchText[0])
            {
                // InterlockedExchange(&g_nOpenBoxCharThreadId, GetCurrentThreadId());
                // GetSystemTimeAsFileTime(&g_ftOpenBoxChar);
                // Skipped telemetry ShellTraceId_Explorer_StartPane_OpenBox_Char_Info
            }
            hr = _UpdateSearch();
        }
    }

    return hr;
}

HRESULT COpenBoxHost::GetSearchText(LPWSTR pszSearchText, UINT cch)
{
    return E_NOTIMPL;
}

HRESULT COpenBoxHost::GetPromptText(LPWSTR pszPromptText, UINT cch)
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

const struct
{
    WCHAR ch;
    const WCHAR* pszEscaped;
} c_rgEscapeCharacters[] =
{
    { L' ', L"%20" },
    { L'%', L"%25" },
    { L'&', L"%26" },
    { L'/', L"%2F" },
    { L':', L"%3A" },
    { L'=', L"%3D" },
    { L'\\', L"%5C" }
};

HRESULT AppendSearchParsingNameEscaped(const WCHAR* pszSearchQuery, SHSTRW* pstrParsingName)
{
    HRESULT hr = S_OK;
    for (const WCHAR* pchCurrent = pszSearchQuery; SUCCEEDED(hr) && *pchCurrent; ++pchCurrent)
    {
        BOOL bEscaped = FALSE;
        for (UINT i = 0; i < ARRAYSIZE(c_rgEscapeCharacters); ++i)
        {
            if (*pchCurrent == c_rgEscapeCharacters[i].ch)
            {
                bEscaped = TRUE;
                hr = pstrParsingName->Append(c_rgEscapeCharacters[i].pszEscaped);
                break;
            }
        }

        if (!bEscaped)
        {
            hr = pstrParsingName->Append(*pchCurrent);
        }
    }

    return hr;
}

HRESULT _InvokeDefaultBrowserSearch(const WCHAR* pszSearchQuery)
{
    WCHAR szName[260];
    DWORD cchName = ARRAYSIZE(szName);
    HRESULT hr = AssocQueryStringW(ASSOCF_NONE, ASSOCSTR_EXECUTABLE, L"http", L"open", szName, &cchName);
    if (SUCCEEDED(hr))
    {
        WCHAR szBuf[260];
        StringCchPrintfW(szBuf, ARRAYSIZE(szBuf), L"\"? %s\"", pszSearchQuery);

        SHELLEXECUTEINFOW shei = {};
        shei.cbSize = sizeof(shei);
        shei.lpFile = szName;
        shei.lpParameters = szBuf;
        shei.nShow = SW_SHOWNORMAL;
        if (!ShellExecuteExW(&shei))
        {
            hr = E_FAIL;
        }
    }
    return hr;
}

HRESULT _InvokeInternetSearchExtension(const WCHAR* pszSearchQuery, HWND hwnd)
{
    HRESULT hr = E_FAIL;

    WCHAR szAction[2084];
    DWORD cbData = sizeof(szAction);
    if (_SHRegGetValueFromHKCUHKLM(
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer\\SearchExtensions",
        L"InternetExtensionAction", SRRF_RT_ANY, nullptr, szAction, &cbData) == ERROR_SUCCESS)
    {
        if (StrCmpNICW(szAction, L"http://", lstrlenW(L"http://")) == 0 || StrCmpNICW(szAction, L"https://", lstrlenW(L"https://")) == 0)
        {
            WCHAR* pszSubstitute = StrStrIW(szAction, L"%w");
            if (pszSubstitute)
            {
                WCHAR szFile[2048];
                StringCchCopyNW(szFile, ARRAYSIZE(szFile), szAction, pszSubstitute - szAction);
                StringCchCatW(szFile, ARRAYSIZE(szFile), pszSearchQuery);
                StringCchCatW(szFile, ARRAYSIZE(szFile), &pszSubstitute[lstrlenW(L"%w")]);

                SHELLEXECUTEINFOW shei = {};
                shei.cbSize = sizeof(shei);
                shei.hwnd = hwnd;
                shei.lpFile = szFile;
                shei.nShow = SW_SHOWNORMAL;
                hr = ShellExecuteExW(&shei) ? S_OK : E_FAIL;
            }
        }
    }

    return hr;
}

HRESULT WWMenuCommand(const WCHAR* pszSearchQuery, DWORD dwCmd, HWND hwnd)
{
    HRESULT hr = E_FAIL;

    switch (dwCmd)
    {
        case 1:
        case 4:
        {
            SHSTRW strSearchQuery;
            hr = strSearchQuery.SetStr(L"search:query=");
            if (SUCCEEDED(hr))
            {
                hr = AppendSearchParsingNameEscaped(pszSearchQuery, &strSearchQuery);
            }
            if (SUCCEEDED(hr))
            {
                SHELLEXECUTEINFOW shei = {};
                shei.cbSize = sizeof(shei);
                shei.hwnd = hwnd;
                shei.lpFile = strSearchQuery.GetStr();
                shei.nShow = SW_SHOWNORMAL;
                hr = ShellExecuteExW(&shei) ? S_OK : E_FAIL;
            }
            break;
        }
        case 2:
        {
            hr = _InvokeDefaultBrowserSearch(pszSearchQuery);
            break;
        }
        case 3:
        {
            hr = _InvokeInternetSearchExtension(pszSearchQuery, hwnd);
            break;
        }
    }

    return hr;
}

HRESULT COpenBoxHost::OnMenuCommand(DWORD dwCmd)
{
    WCHAR szText[MAX_PATH];
    HRESULT hr = _pssc->GetText(szText, ARRAYSIZE(szText));
    if (SUCCEEDED(hr))
    {
        hr = WWMenuCommand(szText, dwCmd, nullptr);
        if (SUCCEEDED(hr))
        {
            RECT rc;
            GetWindowRect(_hwnd, &rc);
            MapWindowRect(_hwnd, nullptr, &rc);

            SMNMCOMMANDINVOKED ci;
            _SendNotify(GetParent(_hwnd), SMN_COMMANDINVOKED, &ci.hdr);
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
    return IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 304, 0, &vt, nullptr);
}

COpenBoxHost::COpenBoxHost(HWND hwnd)
	: _hwnd(hwnd)
    , _lRef(1)
    , field_4C(1)
{
}

COpenBoxHost::~COpenBoxHost()
{
    IUnknown_SafeReleaseAndNullPtr(_pssc);
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
        GetThemeMargins(_hTheme, nullptr, SPP_OPENBOX, 0, TMT_CONTENTMARGINS, nullptr, &margins);
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

    HRESULT hr = CoCreateInstance(/*CLSID_SearchControl*/CLSID_SearchBox, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pssc));
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
                WCHAR szText[260];
                LoadStringW(g_hinstCabinet, 8246, szText, ARRAYSIZE(szText));
                hr = _pssc->SetCueAndTooltipText(szText, nullptr);
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

    NMHDR* pnmh = reinterpret_cast<NMHDR*>(lParam);
    if (pnmh)
    {
        switch (pnmh->code)
        {
            case 215:
                return _OnSMNFindItem((SMNDIALOGMESSAGE*)pnmh);
            case 222:
                if (_pssc)
                {
                    // @MOD: SSCTEXT_FORCE taken from build 6519
                    _pssc->SetText(L"", SSCTEXT_TAKEFOCUS | SSCTEXT_FORCE);
                }
                break;
            case 223:
                return SetSite(((SMNSETSITE*)pnmh)->punkSite);
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
            {
                return 0;
            }
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
            IInputObject* pio = nullptr;
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

DWORD SHExpandEnvironmentStringsW(const WCHAR* pszIn, WCHAR* pszOut, DWORD cchOut);

HRESULT COpenBoxHost::_UpdateSearch()
{
    WCHAR szSearchQuery[260];
    HRESULT hr = _pssc->GetText(szSearchQuery, ARRAYSIZE(szSearchQuery));
    if (SUCCEEDED(hr) && (!_pszSearchQuery || StrCmpCW(szSearchQuery, _pszSearchQuery)))
    {
        CoTaskMemFree(_pszSearchQuery);
        SHStrDupW(szSearchQuery, &_pszSearchQuery);

        WCHAR szPath[260];
        SHExpandEnvironmentStringsW(szSearchQuery, szPath, ARRAYSIZE(szPath));
        PathRemoveBlanksW(szPath);

        VARIANTARG varg;
        hr = InitVariantFromString(szPath, &varg);
        if (SUCCEEDED(hr))
        {
            hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 300, 0, &varg, nullptr);
            IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 300, 0, nullptr, &varg);
            VariantClear(&varg);
        }
    }

    return hr;
}

BOOL OpenBoxHost_RegisterClass()
{
    WNDCLASSEX wcex = {};

    wcex.cbSize = sizeof(wcex);
    wcex.style = CS_GLOBALCLASS;
    wcex.lpfnWndProc = COpenBoxHost::s_WndProc;
    wcex.hInstance = g_hinstCabinet;
    wcex.hbrBackground = nullptr;
    wcex.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wcex.lpszClassName = L"Desktop OpenBox Host";
    return RegisterClassExW(&wcex);
}
