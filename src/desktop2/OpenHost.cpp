#include "pch.h"

#include "OpenHost.h"

#include "cabinet.h"
#include "HostUtil.h"

#define OPENVIEW_MFU			0
#define OPENVIEW_NSCHOST		1
#define OPENVIEW_SEARCHPANE		2
#define OPENVIEW_VIEWCONTROL	3
#define OPENVIEW_TOPMATCH		4

#define COMPILE_OPENVIEW

#ifdef COMPILE_OPENVIEW

const SMPANEDATA g_aopaDefault[] =
{
	{ L"DesktopSFTBarHost",				0,	SPP_PROGLIST,		{ 0,  0},	NULL, NULL, FALSE, NULL },
	{ L"Desktop NSCHost",				0,	SPP_NSCHOST,		{ 0,  0},	NULL, NULL, FALSE, NULL },
	{ L"Desktop Search Open View",		0,	SPP_SEARCHVIEW,		{ 0,  0},	NULL, NULL, FALSE, NULL },
	{ L"Desktop More Programs Pane",	0,	SPP_MOREPROGRAMS,	{ 0, 30},	NULL, NULL, FALSE, NULL },
	{ L"Desktop top match",				0,	SPP_TOPMATCH,		{ 0, 20},	NULL, NULL, FALSE, NULL }
};

HRESULT COpenViewHost::QueryInterface(REFIID riid, void** ppvObj)
{
	static const QITAB qit[] = {
		QITABENT(COpenViewHost, IServiceProvider),
		QITABENT(COpenViewHost, IObjectWithSite),
		QITABENT(COpenViewHost, IOleCommandTarget),
		{ 0 },
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
	if (lRef == 0 && this)
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
		hr = IUnknown_QueryService(_aopa[2].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_ViewControl))
	{
		ASSERT(_aopa[OPENVIEW_VIEWCONTROL].punk != NULL); // 436
		hr = IUnknown_QueryService(_aopa[3].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_TopMatch))
	{
		ASSERT(_aopa[OPENVIEW_TOPMATCH].punk != NULL); // 442
		hr = IUnknown_QueryService(_aopa[4].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_OpenHost))
	{
		hr = QueryInterface(riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_MFU))
	{
		hr = IUnknown_QueryService(_aopa[0].punk, guidService, riid, ppvObject);
	}
	else if (IsEqualGUID(guidService, SID_SM_NSCHOST))
	{
		hr = IUnknown_QueryService(_aopa[1].punk, guidService, riid, ppvObject);
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
			return _HandleOpenBoxArrowKey((MSG *)pvarargIn->lVal);

		case 330:
			ASSERT(pvarargIn->vt == VT_BYREF); // 618
			return _HandleOpenBoxContextMenu((MSG *)pvarargIn->lVal);

		case 307:
			IUnknown_QueryServiceExec((IServiceProvider *)this, SID_SM_OpenBox, &SID_SM_DV2ControlHost, nCmdID, 0, 0, pvarargOut);
			if (pvarargOut->iVal && field_18)
			{
				_SetCurrentView(0, pvarargIn);
				IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 310, 0, 0, 0);
				pvarargOut->iVal = 0;
			}
			return S_OK;

		case 303:
			ASSERT(pvarargOut->vt == VT_I4); // 638
			pvarargOut->decVal.Lo32 = field_18;
			return S_OK;

		case 316:
			ASSERT(pvarargIn->vt == VT_I4); // 646
			_aopa[4].size.cy = pvarargIn->decVal.Lo32;

			RECT rc;
			GetClientRect(_hwnd, &rc);
			_Layout(rc.right, rc.bottom);
			return S_OK;

		case 302:
			return _SetCurrentView(nCmdexecopt == -1 ? field_1C : nCmdexecopt, pvarargIn);

		case 306:
			if (field_18 == 2 || field_18 == 1)
			{
				ASSERT(pvarargIn->vt == VT_BYREF); // 659
				HWND hwndParent = (HWND)pvarargIn->lVal;
				ASSERT(pvarargOut->vt == VT_BYREF); // 661

				if (hwndParent && GetParent(hwndParent) == _hwnd)
					pvarargOut->lVal = (LONG)hwndParent;
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
				IUnknown_SetSite(_aopa[i].punk, NULL);
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
	LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);
	SMPANEDATA *psmpd = reinterpret_cast<SMPANEDATA *>(lpcs->lpCreateParams);

	HTHEME hTheme = psmpd->hTheme;

	IUnknown_Set(&psmpd->punk, static_cast<IServiceProvider *>(this));

	for (int i = 0; i < ARRAYSIZE(_aopa); i++)
	{
		DWORD dwStyle = WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_CHILD;

		if (field_18 == i)
			dwStyle |= WS_VISIBLE;

		if (i == 3 && field_18 < 2)
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

		HWND hwndPane = CreateWindowExW(0, _aopa[i].pszClassName, NULL, dwStyle, 0, 0, 0, _aopa[i].size.cy, hwnd, IntToPtr_(HMENU, i), NULL, &_aopa[i]);
		if (!hwndPane || !GetWindowLongPtr(hwnd, GWL_STYLE))
			return -1;

		_aopa[i].hwnd = hwndPane;

		SMNSETSITE nmss;
		nmss.hdr.hwndFrom = hwndPane;
		nmss.hdr.idFrom = GetDlgCtrlID(hwndPane);
		nmss.hdr.code = 223;
		nmss.punkSite = static_cast<IServiceProvider *>(this);
		SendMessage(hwndPane, WM_NOTIFY, nmss.hdr.idFrom, (LPARAM)&nmss);
	}
	return 0;
}

LRESULT COpenViewHost::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
	LRESULT lRes = DefWindowProc(hwnd, uMsg, wParam, lParam);
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
	LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);

	if (pnm->hwndFrom == GetParent(_hwnd))
	{
		if (pnm->code == SMN_GETMINSIZE)
		{
			//LPARAM lParama[7]; // [esp+Ch] [ebp-1Ch] BYREF
			//signed int v19 = 0;
			//HWND* Msg = &_aopa[0].hwnd;
			//do
			//{
			//	memset(&lParama[1], 0, 0x18u);
			//	lParama[0] = *(_DWORD*)lParam;
			//	lParama[3] = *(_DWORD*)(lParam + 12);
			//	lParama[4] = *(_DWORD*)(lParam + 16);
			//	HWND v12 = *Msg;
			//	lParama[2] = 206;
			//	SendMessageW(v12, WM_NOTIFY, lParama[1], (LPARAM)lParama);
			//	int v8 = lParama[4];
			//	if (v19 < 2)
			//		v8 = _aopa[3].size.cy + lParama[4];
			//	if (*(_DWORD*)(lParam + 16) < v8)
			//		*(_DWORD*)(lParam + 16) = v8;
			//	if (*(_DWORD*)(lParam + 24) < lParama[6])
			//		*(_DWORD*)(lParam + 24) = lParama[6];
			//	++v19;
			//	Msg += 9;
			//} while ((unsigned int)v19 < 5);
			//return 0;

			SMNGETMINSIZE *pnmgs = reinterpret_cast<SMNGETMINSIZE *>(lParam);

			for (int i = 0; i < ARRAYSIZE(_aopa); i++)
			{
				SMNGETMINSIZE nmgms = { 0 };

				nmgms.hdr.hwndFrom = pnmgs->hdr.hwndFrom;
				nmgms.siz = pnmgs->siz;
				nmgms.hdr.code = pnmgs->hdr.code;

				SendMessage(_aopa[i].hwnd, WM_NOTIFY, nmgms.hdr.idFrom, (LPARAM)&nmgms);

				// int cy = nmgms.siz.cy;
				if (i < 2)
				{
					nmgms.siz.cy += _aopa[3].size.cy;
					//cy = _aopa[3].size.cy + nmgms.siz.cy;
				}

				if (pnmgs->siz.cy < /*cy*/ nmgms.siz.cy)
				{
					pnmgs->siz.cy = /*cy*/ nmgms.siz.cy;
				}

				if (pnmgs->field_14.cx < nmgms.field_14.cx)
				{
					pnmgs->field_14.cx = nmgms.field_14.cx;
				}
			}
			return 0;
		}


		if (pnm->code == 210)
			goto LABEL_15;

		if (pnm->code != 215)
		{
			if (pnm->code == 217)
			{
				if (field_18 < 2)
				{
					return SendMessage(_aopa[3].hwnd, uMsg, wParam, lParam);
				}
				return 0;
			}
			if (pnm->code != 221)
			{
				if (pnm->code != 222)
				{
					if (pnm->code == 223)
					{
						return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
					}
					return SendMessage(_aopa[field_18].hwnd, uMsg, wParam, lParam);
				}
			LABEL_15:
				_SetCurrentView(0, NULL);
				SHPropagateMessage(_hwnd, uMsg, wParam, lParam, SPM_POST | SPM_ONELEVEL);
				return 0;
			}
			_SetCurrentView(0, NULL);
		}
		return 0;
	}

	if (pnm->hwndFrom == _hwnd)
	{
		HWND hwndInner = ::GetWindow(_aopa[field_18].hwnd, GW_CHILD);
		pnm->hwndFrom = hwndInner;
		pnm->idFrom = GetDlgCtrlID(hwndInner);
		return SendMessage(_aopa[field_18].hwnd, WM_NOTIFY, pnm->idFrom, lParam);
	}

	if (pnm->code == 202 || pnm->code == 204 || pnm->code == 207 || pnm->code == 209 || pnm->code == 212 || pnm->code == 218)
	{
		return SendMessage(GetParent(_hwnd), uMsg, wParam, lParam);
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void COpenViewHost::_Layout(int cx, int cy)
{
	// ebx
	HWND hwnd; // edx
	int v6; // eax
	LONG cy1; // ecx

	for (int i = 0; (unsigned int)i < ARRAYSIZE(_aopa); ++i)
	{
		hwnd = _aopa[i].hwnd;
		if (!hwnd)
			continue;
		v6 = 0;
		if (i >= 2)
		{
			if (i == 3)
			{
				cy1 = _aopa[3].size.cy;
			}
			else
			{
				if (i == 2)
				{
					cy1 = cy - _aopa[4].size.cy;
					goto LABEL_11;
				}
				cy1 = _aopa[4].size.cy;
			}
			v6 = cy - cy1;
			goto LABEL_11;
		}
		cy1 = cy - _aopa[3].size.cy;
	LABEL_11:
		SetWindowPos(hwnd, NULL, 0, v6, cx, cy1, 0x204);
	}
}

void COpenViewHost::_ShowEnableWindow(HWND hwnd, BOOL bEnable)
{
	ShowWindow(hwnd, bEnable ? SW_SHOW : SW_HIDE);
	EnableWindow(hwnd, bEnable);
}

HRESULT COpenViewHost::_SetCurrentView(int a2, VARIANT* pvararg)
{
	HRESULT hr = S_FALSE;
	
	if (field_18 != a2)
	{
		_ShowEnableWindow(_aopa[field_18].hwnd, FALSE);
		
		bool v6 = field_18 == 2;
		if (field_18 < 2)
		{
			_ShowEnableWindow(_aopa[3].hwnd, FALSE);
			IUnknown_QueryServiceExec(_aopa[3].punk, SID_SM_ViewControl, &SID_SM_DV2ControlHost, 302, a2, NULL, NULL);
			v6 = field_18 == 2;
		}
		
		if (v6)
		{
			_ShowEnableWindow(_aopa[4].hwnd, FALSE);
		}

		//if (a2 == 2)
		//{
		//	if (field_18 == 1)
		//		SHTracePerfSQMCountImpl((PCEVENT_DESCRIPTOR)ShellTraceId_StartMenu_AllPrograms_Search_Usage, 71);
		//	else
		//		SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Usage, 72);
		//}

		int v7 = field_18;
		field_1C = v7;
		if (v7 == 2)
			field_1C = 0;
		field_18 = a2;
		
		_ShowEnableWindow(_aopa[a2].hwnd, TRUE);
		
		if (field_18 < 2)
		{
			_ShowEnableWindow(_aopa[3].hwnd, TRUE);
		}
		else if (field_18 == 2)
		{
			_ShowEnableWindow(_aopa[4].hwnd, TRUE);
		}

		if (field_18 == 1)
		{
			IUnknown_QueryServiceExec(_aopa[1].punk, SID_SM_NSCHOST, &SID_SM_DV2ControlHost, 302, a2, pvararg, NULL);
		}
		return S_OK;
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
	
	if (field_18 == 2)
	{
		VARIANT vt;
		vt.lVal = -1;
		vt.vt = VT_I4;
		if (IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenView, &SID_SM_DV2ControlHost, 325, 0, NULL, &vt) >= 0 && vt.lVal != -1)
		{
			pmsg->hwnd  = _aopa[2].hwnd;
		}
		else
		{
			hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_TopMatch, &SID_SM_DV2ControlHost, 325, 0, NULL, &vt);
			pmsg->hwnd  = _aopa[4].hwnd;
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
	WNDCLASSEX wc;
	ZeroMemory(&wc, sizeof(wc));

	wc.cbSize			= sizeof(wc);
	wc.style			= CS_GLOBALCLASS;
	wc.lpfnWndProc		= COpenViewHost::s_WndProc;
	wc.hInstance		= g_hinstCabinet;
	wc.hbrBackground	= (HBRUSH)NULL;
	wc.hCursor			= LoadCursor(NULL, IDC_ARROW);
	wc.lpszClassName	= WC_OPENPANEHOST;

	return RegisterClassEx(&wc);
}

#endif
