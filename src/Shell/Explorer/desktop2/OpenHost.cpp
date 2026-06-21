#include "pch.h"

#include "OpenHost.h"

#include "cabinet.h"
#include "HostUtil.h"

#define OPENVIEW_MFU			0
#define OPENVIEW_NSCHOST		1
#define OPENVIEW_SEARCHPANE		2
#define OPENVIEW_VIEWCONTROL	3
#define OPENVIEW_TOPMATCH		4

const SMPANEDATA g_aopaDefault[] =
{
	{ L"DesktopSFTBarHost",				0,	SPP_PROGLIST,		{ 0,  0}, nullptr, nullptr, FALSE, nullptr },
	{ L"Desktop NSCHost",				0,	SPP_NSCHOST,		{ 0,  0}, nullptr, nullptr, FALSE, nullptr },
	{ L"Desktop Search Open View",		0,	SPP_SEARCHVIEW,		{ 0,  0}, nullptr, nullptr, FALSE, nullptr },
	{ L"Desktop More Programs Pane",	0,	SPP_MOREPROGRAMS,	{ 0, 30}, nullptr, nullptr, FALSE, nullptr },
	{ L"Desktop top match",				0,	SPP_TOPMATCH,		{ 0, 20}, nullptr, nullptr, FALSE, nullptr }
};

HRESULT COpenViewHost::QueryInterface(REFIID riid, void** ppvObj)
{
	static const QITAB qit[] =
	{
		QITABENT(COpenViewHost, IServiceProvider),
		QITABENT(COpenViewHost, IObjectWithSite),
		QITABENT(COpenViewHost, IOleCommandTarget),
		{},
	};
	return QISearch(this, qit, riid, ppvObj);
}

ULONG COpenViewHost::AddRef()
{
	return InterlockedIncrement(&_lRef);
}

ULONG COpenViewHost::Release()
{
	LONG lRef = InterlockedDecrement(&_lRef);
	if (lRef == 0)
	{
		delete this;
	}
	return lRef;
}

HRESULT COpenViewHost::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
	HRESULT hr = E_FAIL;

	if (IsEqualGUID(guidService, SID_SM_OpenView))
	{
		ASSERT(_aopa[OPENVIEW_SEARCHPANE].punk != NULL); // 430
		hr = IUnknown_QueryService(_aopa[OPENVIEW_SEARCHPANE].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_ViewControl))
	{
		ASSERT(_aopa[OPENVIEW_VIEWCONTROL].punk != NULL); // 436
		hr = IUnknown_QueryService(_aopa[OPENVIEW_VIEWCONTROL].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_TopMatch))
	{
		ASSERT(_aopa[OPENVIEW_TOPMATCH].punk != NULL); // 442
		hr = IUnknown_QueryService(_aopa[OPENVIEW_TOPMATCH].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_OpenHost))
	{
		hr = QueryInterface(riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_MFU))
	{
		hr = IUnknown_QueryService(_aopa[OPENVIEW_MFU].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_NSCHOST))
	{
		hr = IUnknown_QueryService(_aopa[OPENVIEW_NSCHOST].punk, guidService, riid, ppvObject);
	}

	if (FAILED(hr))
	{
		return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
	}
	return hr;
}

HRESULT COpenViewHost::QueryStatus(const GUID* pguidCmdGroup,
	ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pCmdText)
{
	return E_NOTIMPL;
}

HRESULT COpenViewHost::Exec(const GUID *pguidCmdGroup,
	DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
	HRESULT hr = E_INVALIDARG;

	if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
	{
		switch (nCmdID)
		{
			case 324:
				ASSERT(pvarargIn->vt == VT_BYREF); // 611
				return _HandleOpenBoxArrowKey(static_cast<MSG*>(pvarargIn->byref));
			case 330:
				ASSERT(pvarargIn->vt == VT_BYREF); // 618
				return _HandleOpenBoxContextMenu(static_cast<MSG*>(pvarargIn->byref));
			case 307:
				IUnknown_QueryServiceExec(static_cast<IServiceProvider*>(this), SID_SM_OpenBox, &SID_SM_DV2ControlHost, nCmdID, 0, nullptr, pvarargOut);
				if (pvarargOut->iVal && _iCurView)
				{
					_SetCurrentView(0, pvarargIn);
					IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 310, 0, nullptr, nullptr);
					pvarargOut->iVal = 0;
				}
				return S_OK;
			case 303:
				ASSERT(pvarargOut->vt == VT_I4); // 638
				pvarargOut->lVal = _iCurView;
				return S_OK;
			case 316:
				ASSERT(pvarargIn->vt == VT_I4); // 646
				_aopa[OPENVIEW_TOPMATCH].size.cy = pvarargIn->lVal;

				RECT rc;
				GetClientRect(_hwnd, &rc);
				_Layout(rc.right, rc.bottom);
				return S_OK;
			case 302:
				return _SetCurrentView(nCmdexecopt == -1 ? field_1C : nCmdexecopt, pvarargIn);
			case 306:
				if (_iCurView == 2 || _iCurView == 1)
				{
					ASSERT(pvarargIn->vt == VT_BYREF); // 659
					HWND hwndParent = static_cast<HWND>(pvarargIn->byref);
					ASSERT(pvarargOut->vt == VT_BYREF); // 661
					if (hwndParent && GetParent(hwndParent) == _hwnd)
					{
						pvarargOut->byref = hwndParent;
					}
				}
				return S_OK;
		}
	}

	return hr;
}

HRESULT COpenViewHost::SetSite(IUnknown* punkSite)
{
	CObjectWithSite::SetSite(punkSite);

	if (!_punkSite)
	{
		for (int i = 0; i < ARRAYSIZE(_aopa); i++)
		{
			if (_aopa[i].punk)
			{
				IUnknown_SetSite(_aopa[i].punk, nullptr);
				IUnknown_SafeReleaseAndNullPtr(&_aopa[i].punk);
			}
		}
	}
	return S_OK;
}

COpenViewHost::COpenViewHost(HWND hwnd)
	: _hwnd(hwnd)
	, _lRef(1)
{
	memcpy(_aopa, g_aopaDefault, sizeof(_aopa));
}

COpenViewHost::~COpenViewHost()
{
}

LRESULT COpenViewHost::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	COpenViewHost* self = reinterpret_cast<COpenViewHost*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	switch (uMsg)
	{
	case WM_CREATE:
		return self->_OnCreate(hwnd, uMsg, wParam, lParam);

	case WM_SIZE:
		return self->_OnSize(hwnd, uMsg, wParam, lParam);

	case WM_SETTINGCHANGE:
		SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
		break;

	case WM_NOTIFY:
		return self->_OnNotify(hwnd, uMsg, wParam, lParam);

	case WM_NCCREATE:
		return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);

	case WM_NCDESTROY:
		return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT COpenViewHost::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	COpenViewHost* povh = new COpenViewHost(hwnd);
	SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)povh);
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

extern void RemapSizeForHighDPI(SIZE *psiz);

LRESULT COpenViewHost::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CREATESTRUCTW* lpcs = reinterpret_cast<CREATESTRUCTW*>(lParam);
	SMPANEDATA* psmpd = static_cast<SMPANEDATA*>(lpcs->lpCreateParams);

	HTHEME hTheme = psmpd->hTheme;

	IUnknown_Set(&psmpd->punk, static_cast<IServiceProvider*>(this));

	for (int i = 0; i < ARRAYSIZE(_aopa); i++)
	{
		DWORD dwStyle = WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_CHILD;

		if (_iCurView == i)
			dwStyle |= WS_VISIBLE;

		if (i == 3 && _iCurView < 2)
			dwStyle |= WS_VISIBLE;

		if (hTheme)
		{
			_aopa[i].bPartDefined = IsThemePartDefined(hTheme, _aopa[i].iPartId, 0);

			if (_aopa[i].bPartDefined)
				_aopa[i].hTheme = hTheme;

			if (i == 3)
			{
				RECT rc;
				if (SUCCEEDED(GetThemeRect(hTheme, _aopa[i].iPartId, 0, TMT_DEFAULTPANESIZE, &rc)))
				{
					_aopa[i].size.cx = RECTWIDTH(rc);
					_aopa[i].size.cy = RECTHEIGHT(rc);
				}
			}
		}

		RemapSizeForHighDPI(&_aopa[i].size);

		HWND hwndPane = CreateWindowExW(
			0, _aopa[i].pszClassName, nullptr, dwStyle, 0, 0, 0, _aopa[i].size.cy, hwnd, IntToPtr_(HMENU, i),
			nullptr, &_aopa[i]);
		if (!hwndPane || !GetWindowLongPtrW(hwnd, GWL_STYLE))
			return -1;

		_aopa[i].hwnd = hwndPane;

		SMNSETSITE nmss;
		nmss.hdr.hwndFrom = hwndPane;
		nmss.hdr.idFrom = GetDlgCtrlID(hwndPane);
		nmss.hdr.code = 223;
		nmss.punkSite = static_cast<IServiceProvider*>(this);
		SendMessageW(hwndPane, WM_NOTIFY, nmss.hdr.idFrom, (LPARAM)&nmss);
	}
	return 0;
}

LRESULT COpenViewHost::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
	LRESULT lRes = DefWindowProcW(hwnd, uMsg, wParam, lParam);
	if (this)
		Release();
	return lRes;
}

LRESULT COpenViewHost::_OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int cx = GET_X_LPARAM(lParam);
	int cy = GET_Y_LPARAM(lParam);
	_Layout(cx, cy);
	return 0;
}

LRESULT COpenViewHost::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HWND Parent;
	int cy;
	HWND hwndInner;
	int DlgCtrlID;

	NMHDR* pnm = reinterpret_cast<NMHDR*>(lParam);
	HWND hwndFrom = pnm->hwndFrom;
	if (hwndFrom == GetParent(_hwnd))
	{
		switch (pnm->code)
		{
			case 206:
			{
				SMNGETMINSIZE* pnmgs = reinterpret_cast<SMNGETMINSIZE*>(lParam);

				for (__int64 i = 0; i < 5; ++i)
				{
					SMNGETMINSIZE nmgms = {};
					nmgms.hdr.hwndFrom = pnm->hwndFrom;
					nmgms.hdr.code = 206;
					nmgms.siz = pnmgs->siz;
					SendMessageW(_aopa[i].hwnd, WM_NOTIFY, nmgms.hdr.idFrom, reinterpret_cast<LPARAM>(&nmgms));

					cy = nmgms.siz.cy;
					if (i < 2)
					{
						cy = _aopa[3].size.cy + nmgms.siz.cy;
					}
					if (pnmgs->siz.cy < cy)
					{
						pnmgs->siz.cy = cy;
					}
					if (pnmgs->field_14.cy < nmgms.field_14.cy)
					{
						pnmgs->field_14.cy = nmgms.field_14.cy;
					}
				}
				return 0;
			}
			case 210:
				goto LABEL_14;
			case 215:
				return 0;
			default:
				break;
		}

		if (pnm->code != 217)
		{
			if (pnm->code != 221)
			{
				if (pnm->code != 222)
				{
					if (pnm->code == 223)
					{
						return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
					}
					Parent = _aopa[_iCurView].hwnd;
					return SendMessageW(Parent, uMsg, wParam, lParam);
				}
			LABEL_14:
				_SetCurrentView(0, nullptr);
				SHPropagateMessage(_hwnd, uMsg, wParam, lParam, 3);
				return 0;
			}
			_SetCurrentView(0, nullptr);
			return 0;
		}
		if (_iCurView >= 2)
		{
			return 0;
		}
		Parent = _aopa[3].hwnd;
	}
	else
	{
		if (hwndFrom == _hwnd)
		{
			hwndInner = GetWindow(_aopa[_iCurView].hwnd, 5u);
			pnm->hwndFrom = hwndInner;
			DlgCtrlID = GetDlgCtrlID(hwndInner);
			uMsg = 0x4E;
			wParam = DlgCtrlID;
			pnm->idFrom = DlgCtrlID;
			Parent = _aopa[_iCurView].hwnd;
		}
		else
		{
			if (pnm->code != 202
				&& pnm->code != 204
				&& pnm->code != 207
				&& pnm->code != 209
				&& pnm->code != 212
				&& pnm->code != 218)
			{
				return DefWindowProcW(hwnd, uMsg, wParam, lParam);
			}
			Parent = GetParent(_hwnd);
		}
	}

	return SendMessageW(Parent, uMsg, wParam, lParam);
}

void COpenViewHost::_Layout(int cx, int cy)
{
	HWND hwnd;
	int y;
	int cyPane;

	for (int i = 0; (unsigned int)i < ARRAYSIZE(_aopa); ++i)
	{
		hwnd = _aopa[i].hwnd;
		if (!hwnd)
			continue;

		y = 0;
		if (i < 2)
		{
			cyPane = cy - _aopa[3].size.cy;
		LABEL_11:
			SetWindowPos(hwnd, nullptr, 0, y, cx, cyPane, SWP_NOZORDER | SWP_NOOWNERZORDER);
			continue;
		}
		if (i == 3)
		{
			cyPane = _aopa[3].size.cy;
		}
		else
		{
			if (i == 2)
			{
				cyPane = cy - _aopa[OPENVIEW_TOPMATCH].size.cy;
				goto LABEL_11;
			}
			cyPane = _aopa[OPENVIEW_TOPMATCH].size.cy;
		}

		y = cy - cyPane;
		goto LABEL_11;
	}
}

void COpenViewHost::_ShowEnableWindow(HWND hwnd, BOOL bEnable)
{
	ShowWindow(hwnd, bEnable ? SW_SHOW : SW_HIDE);
	EnableWindow(hwnd, bEnable);
}

HRESULT COpenViewHost::_SetCurrentView(int iView, VARIANT* pvararg)
{
	HRESULT hr = S_FALSE;

	if (_iCurView != iView)
	{
		_ShowEnableWindow(_aopa[_iCurView].hwnd, FALSE);

		if (_iCurView < OPENVIEW_SEARCHPANE)
		{
			_ShowEnableWindow(_aopa[OPENVIEW_VIEWCONTROL].hwnd, FALSE);
			IUnknown_QueryServiceExec(
				_aopa[OPENVIEW_VIEWCONTROL].punk, SID_SM_ViewControl, &SID_SM_DV2ControlHost, 302, iView, nullptr,
				nullptr);
		}
		if (_iCurView == OPENVIEW_SEARCHPANE)
		{
			_ShowEnableWindow(_aopa[OPENVIEW_TOPMATCH].hwnd, FALSE);
		}

		if (iView == OPENVIEW_SEARCHPANE)
		{
			if (_iCurView == OPENVIEW_NSCHOST)
			{
				// SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_AllPrograms_Search_Usage, 583);
			}
			else
			{
				// SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Usage, 584);
			}
		}

		field_1C = _iCurView;
		if (field_1C == 2)
		{
			field_1C = 0;
		}
		_iCurView = iView;
		_ShowEnableWindow(_aopa[iView].hwnd, TRUE);

		if (_iCurView < OPENVIEW_SEARCHPANE)
		{
			_ShowEnableWindow(_aopa[OPENVIEW_VIEWCONTROL].hwnd, TRUE);
		}
		else if (_iCurView == OPENVIEW_SEARCHPANE)
		{
			_ShowEnableWindow(_aopa[OPENVIEW_TOPMATCH].hwnd, TRUE);
		}
		if (_iCurView == OPENVIEW_NSCHOST)
		{
			IUnknown_QueryServiceExec(
				_aopa[OPENVIEW_NSCHOST].punk, SID_SM_NSCHOST, &SID_SM_DV2ControlHost, 302, iView, pvararg, nullptr);
		}
		hr = S_OK;
	}
	return hr;
}

HRESULT COpenViewHost::_HandleOpenBoxArrowKey(MSG* pmsg)
{
	return E_NOTIMPL; // EXEX-VISTA(allison): TODO.
}

HRESULT COpenViewHost::_HandleOpenBoxContextMenu(MSG* pmsg)
{
	HRESULT hr = E_FAIL;

	if (_iCurView == 2)
	{
		VARIANT vt;
		vt.lVal = -1;
		vt.vt = VT_I4;
		if (SUCCEEDED(IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 325, 0, nullptr, &vt)) && vt.lVal != -1)
		{
			pmsg->hwnd = _aopa[OPENVIEW_SEARCHPANE].hwnd;
		}
		else
		{
			hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 325, 0, nullptr, &vt);
			pmsg->hwnd = _aopa[OPENVIEW_TOPMATCH].hwnd;
			if (vt.lVal == -1)
			{
				return hr;
			}
		}
		return 0;
	}

	return hr;
}

BOOL WINAPI OpenViewHost_RegisterClass()
{
	WNDCLASSEXW wc		= {};
	wc.cbSize			= sizeof(wc);
	wc.style			= CS_GLOBALCLASS;
	wc.lpfnWndProc		= COpenViewHost::s_WndProc;
	wc.hInstance		= g_hinstCabinet;
	wc.hbrBackground	= (HBRUSH)nullptr;
	wc.hCursor =		LoadCursorW(nullptr, IDC_ARROW);
	wc.lpszClassName =	L"Desktop Open Pane Host";

	return RegisterClassEx(&wc);
}
