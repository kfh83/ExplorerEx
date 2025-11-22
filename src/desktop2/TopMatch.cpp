#include "pch.h"

#include "TopMatch.h"

#include <propvarutil.h>
#include "SFTHost.h"

#ifdef COMPILE_TOPMATCH

CTopMatch::CTopMatch(HWND hwnd)
	: _hwnd(hwnd)
	, _lRef(1)
{
}

CTopMatch::~CTopMatch()
{
}

HRESULT CTopMatch::QueryInterface(REFIID riid, void **ppvObj)
{
	static const QITAB qit[] =
	{
		QITABENT(CTopMatch, IServiceProvider),
		QITABENT(CTopMatch, IOleCommandTarget),
		QITABENT(CTopMatch, IObjectWithSite),
		QITABENT(CTopMatch, IAccessible),
		QITABENT(CTopMatch, IDispatch),
		QITABENT(CTopMatch, IEnumVARIANT),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObj);
}

ULONG CTopMatch::AddRef()
{
	return InterlockedIncrement(&_lRef);
}

ULONG CTopMatch::Release()
{
	LONG lRef = InterlockedDecrement(&_lRef);
	if (lRef == 0 && this)
	{
		delete this;
	}
	return lRef;
}

HRESULT CTopMatch::get_accState(VARIANT varChild, VARIANT *pvarState)
{
	HRESULT hr = _paccInner->get_accState(varChild, pvarState);
	if (SUCCEEDED(hr) && pvarState->vt == VT_I4)
	{
		pvarState->lVal |= 0x100004u;
	}
	return hr;
}

HRESULT CTopMatch::QueryService(REFGUID guidService, REFIID riid, void** ppvObject)
{
	if (IsEqualGUID(guidService, SID_SM_TopMatch))
	{
		HRESULT hr = QueryInterface(riid, ppvObject);
		if (hr >= 0)
		{
			return hr;
		}
	}
	return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

GUID POLID_NoSearchInternetLinkInStartMenu =
{
  2349109124u,
  16169u,
  18090u,
  { 149u, 144u, 236u, 147u, 63u, 146u, 255u, 51u }
};

const GUID POLID_NoSearchComputerLinkInStartMenu =
{
  2183995476u,
  37952u,
  18650u,
  { 146u, 246u, 250u, 144u, 222u, 14u, 156u, 150u }
};

HRESULT CTopMatch::SetSite(IUnknown *punkSite)
{
	HRESULT hr = CObjectWithSite::SetSite(punkSite);
	if (punkSite)
	{
		_SetTileWidth(-1);
		if (!SHWindowsPolicy(POLID_NoSearchComputerLinkInStartMenu))
		{
			SHWindowsPolicy(POLID_NoSearchInternetLinkInStartMenu);

			LVITEM lvi = {0};
			LVFINDINFO lvfi = {0};

			lvi.lParam = 0;
			lvfi.lParam = 0;
			lvi.mask = 5;
			lvi.pszText = _sz;
			lvfi.flags = 1;
			if (ListView_FindItem(_hwndList, -1, &lvfi) < 0)
			{
				lvi.mask |= 2u;
				lvi.iItem = 0;
				lvi.iImage = 22;
				ListView_InsertItem(_hwndList, &lvi);
			}
		}
		_AddSearchItem(1, _sz2);
		_AddSearchExtension();
		_UpdateTopMatchSizeInOpenView();
	}
	return hr;
}

HRESULT CTopMatch::QueryStatus(const GUID *pguidCmdGroup,
	ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

HRESULT CTopMatch::Exec(const GUID *pguidCmdGroup,
	DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
	HRESULT hr = E_INVALIDARG;

	if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
	{
		switch (nCmdID)
		{
			case 301:
			{
				int iItem = _GetLVCurSel();
				if (iItem >= 0)
				{
					_ActivateItem(iItem, 1);
					return S_OK;
				}
				return E_FAIL;
			}
			case 317:
			{
				int iIndex = VariantToInt32WithDefault(*pvarargIn, 0);
				if (iIndex >= 0)
				{
					if (!field_464)
					{
						field_464 = 1;
						ListView_SetItemState(_hwndList, iIndex, LVIS_SELECTED, LVIS_SELECTED);
					}
				}
				else
				{
					field_464 = 0;
					ListView_SetItemState(_hwndList, -1, 0, LVIS_SELECTED);
				}
				return 0;
			}
			case 325:
			{
				pvarargOut->vt = VT_I4;
				pvarargOut->lVal = _GetLVCurSel();
				return 0;
			}
			case 328:
			{
				ListView_SetItemText(_hwndList, 0, 0, _sz);
				break;
			}
		}
		return hr;
	}
	return hr;
}

LRESULT CTopMatch::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CTopMatch *self = reinterpret_cast<CTopMatch *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	switch (uMsg)
	{
		case WM_NOTIFY:
			return self->_OnNotify(hwnd, uMsg, wParam, lParam);
		case WM_CREATE:
			return self->_OnCreate(hwnd, uMsg, wParam, lParam);
		case WM_SIZE:
			return self->_OnSize(hwnd, uMsg, wParam, lParam);
		case WM_ERASEBKGND:
			return self->_OnEraseBackground(hwnd, uMsg, wParam, lParam);
		case WM_SYSCOLORCHANGE:
			return self->_OnSysColorChange(hwnd, uMsg, wParam, lParam);
		case WM_SETTINGCHANGE:
			return self->_OnSettingChange(hwnd, uMsg, wParam, lParam);
		case WM_CONTEXTMENU:
			return self->_OnContextMenu(hwnd, uMsg, wParam, lParam);
		case WM_DISPLAYCHANGE:
			return self->_OnSettingChange(hwnd, uMsg, wParam, lParam);
		case WM_NCCREATE:
			return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);
		case WM_NCDESTROY:
			return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CTopMatch::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);

	if (pnm->hwndFrom == _hwndList)
	{
		switch (pnm->code)
		{
			case NM_RETURN:
			{
				return _ActivateItem(_GetLVCurSel(), 1);
			}
			case NM_CLICK:
			{
				return _ActivateItem(lParam, 1);
			}
		}
	}

	switch (pnm->code)
	{
		case SMN_GETMINSIZE:
		{
			return _OnSMNGetMinSize((PSMNGETMINSIZE)pnm);
		}
		case 215:
		{
			return _OnSMNFindItem((PSMNDIALOGMESSAGE)pnm);
		}
		case 223:
		{
			return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
		}
		case NM_KILLFOCUS:
		{
			ListView_SetItemState(_hwndList, -1, 0, LVIS_SELECTED);
			field_464 = 0;
			break;
		}
	}

	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}


// EXEX-Vista(allison): TODO, move me to a common header
#define NTDDI_VERSION NTDDI_VISTA
#define _WIN32_WINNT _WIN32_WINNT_VISTA
#include <commoncontrols.h>

LRESULT CTopMatch::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SMPANEDATA *psmpd = PaneDataFromCreateStruct(lParam);
	IUnknown_Set(&psmpd->punk, SAFECAST(this, IServiceProvider*));

	RECT rc;
	GetClientRect(_hwnd, &rc);
	_hTheme = psmpd->hTheme;

	field_10 = 1;

	if (_hTheme)
	{
		GetThemeMargins(_hTheme, NULL, SPP_TOPMATCH, 0, TMT_CONTENTMARGINS, &rc, &_margins);
	}
	else
	{
		_margins.cyTopHeight = 0;
		_margins.cyBottomHeight = 0;
		_margins.cxLeftWidth = 2 * SHGetSystemMetricsScaled(SM_CXEDGE);
		_margins.cxRightWidth = 2 * SHGetSystemMetricsScaled(SM_CXEDGE);
	}

	_hwndList = SHFusionCreateWindowEx(0, L"SysListView32", NULL, 0x5600204F,
		_margins.cxLeftWidth, _margins.cyTopHeight,
		RECTWIDTH(rc), RECTHEIGHT(rc),
		_hwnd, 0, g_hinstCabinet, 0);
	
	if (!_hwndList)
		return -1;

	DWORD dwLvExStyle = LVS_EX_FULLROWSELECT | LVS_EX_INFOTIP;
	if (!GetSystemMetrics(SM_REMOTESESSION) && !GetSystemMetrics(SM_REMOTECONTROL))
	{
		dwLvExStyle |= LVS_EX_DOUBLEBUFFER;
	}
	
	SendMessage(_hwndList, LVM_SETEXTENDEDLISTVIEWSTYLE, dwLvExStyle, dwLvExStyle);
	if (_hTheme)
	{
		GetThemeColor(_hTheme, SPP_TOPMATCH, 0, TMT_HOTTRACKING, &field_44);
		COLORREF clrText;
		GetThemeColor(_hTheme, SPP_TOPMATCH, 0, TMT_TEXTCOLOR, &clrText);
		_clrBk = CLR_INVALID;
		ListView_SetTextColor(_hwndList, clrText);
		ListView_SetBkColor(_hwndList, _clrBk);
		ListView_SetTextBkColor(_hwndList, _clrBk);
		ListView_SetOutlineColor(_hwndList, field_44);
	}
	else
	{
		ListView_SetTextColor(_hwndList, GetSysColor(COLOR_MENUTEXT));
		field_44 = GetSysColor(COLOR_MENUTEXT);
		_clrBk = GetSysColor(COLOR_MENU);
	}
	
	LPCWSTR pszTheme = IsCompositionActive() ? L"StartMenuComposited" : L"StartMenu";
	SetWindowTheme(_hwndList, pszTheme, NULL);
	_cyIcon = GetSystemMetrics(SM_CYSMICON);


	IImageList2 *piml;
	if (SUCCEEDED(SHGetImageList(-1, IID_PPV_ARGS(&piml))))
	{
		if (SUCCEEDED(piml->Resize(_cyIcon, _cyIcon)))
			_himl = (HIMAGELIST)piml;
		else
			piml->Release();
	}

	ListView_SetView(_hwndList, LV_VIEW_TILE);

	if (!_himl)
		return -1;

	_cyIcon += 2 * SHGetSystemMetricsScaled(SM_CYEDGE) + 2;
	ListView_SetImageList(_hwndList, _himl, LVSIL_NORMAL);

	LoadString(g_hinstCabinet, 8243, _sz, ARRAYSIZE(_sz));
	LoadString(g_hinstCabinet, 8244, _sz2, ARRAYSIZE(_sz2));
	SetAccessibleSubclassWindow(_hwndList);
	return 0;
}

LRESULT CTopMatch::_OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	int cxLeftWidth = _margins.cxLeftWidth;
	int cx = GET_X_LPARAM(lParam) - (cxLeftWidth + _margins.cxRightWidth);
	int cyTopHeight = _margins.cyTopHeight;
	int cy = GET_Y_LPARAM(lParam) - (cyTopHeight + _margins.cyBottomHeight);
	if (cx < 0)
		cx = 0;
	if (cy < 0)
		cy = 0;
	SetWindowPos(_hwndList, NULL, cxLeftWidth, cyTopHeight, cx, cy, 0x204u);
	_SetTileWidth(cx);
	return 0;
}

LRESULT CTopMatch::_OnEraseBackground(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	RECT rc;
	GetClientRect(hwnd, &rc);

	if (_hTheme)
	{
		if (IsCompositionActive())
		{
			SHFillRectClr((HDC)wParam, &rc, 0);
		}
		DrawThemeBackground(_hTheme, (HDC)wParam, SPP_TOPMATCH, 0, &rc, NULL);
	}
	else
	{
		SHFillRectClr((HDC)wParam, &rc, _clrBk);
	}
	return 1;
}

LRESULT CTopMatch::_OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (_hTheme)
	{
		ListView_SetTextColor(_hwndList, GetSysColor(COLOR_MENUTEXT));
		field_44 = GetSysColor(COLOR_MENUTEXT);
		_clrBk = GetSysColor(COLOR_MENU);
		ListView_SetBkColor(_hwndList, _clrBk);
		ListView_SetTextBkColor(_hwndList, _clrBk);
	}
	SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
	return 0;
}

LRESULT CTopMatch::_OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
	return 0;
}

LRESULT CTopMatch::_OnContextMenu(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (!SHRestricted(REST_NOSMMOREPROGRAMS))
	{
		if (IS_WM_CONTEXTMENU_KEYBOARD(lParam))
		{
			RECT rc;
			GetWindowRect(_hwnd, &rc);
			lParam = MAKELPARAM(rc.left, rc.top);
		}

		SMNGETISTARTBUTTON nm;
		nm.pstb = NULL;
		_SendNotify(_hwnd, 218, &nm.hdr);
		if (nm.pstb)
		{
			nm.pstb->OnContextMenu(_hwnd, lParam);
			nm.pstb->Release();
		}
	}
	return 0;
}

LRESULT CTopMatch::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CTopMatch *self = new CTopMatch(hwnd);
	if (!self)
		return 0;
	SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)self);
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CTopMatch::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
	LRESULT lRes = DefWindowProc(hwnd, uMsg, wParam, lParam);
	if (this)
		this->Release();
	return lRes;
}

LRESULT CTopMatch::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
	LRESULT lres = _OnSMNFindItemWorker(pdm);
	if (lres)
	{
		if ((pdm->flags & 0x100 | 0x800) != 0)
		{
			field_464 = 1;
			ListView_SetItemState(_hwndList, pdm->itemID, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
			if ((pdm->flags & SMNDM_FINDMASK) != SMNDM_HITTEST)
			{
				ListView_KeyboardSelected(_hwndList, pdm->itemID);
			}
		}
	}
	else
	{
		pdm->flags |= 0x4000u;

		int iItem = _GetLVCurSel();
		RECT rc;
		if (iItem >= 0 && ListView_GetItemRect(_hwndList, iItem, &rc, LVIR_BOUNDS))
		{
			pdm->pt.x = (RECTWIDTH(rc)) / 2;
			pdm->pt.y = (RECTHEIGHT(rc)) / 2;
		}
		else
		{
			pdm->pt.x = 0;
			pdm->pt.y = 0;
		}
	}
	return lres;
}

LRESULT CTopMatch::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
	LRESULT iItem1;
	WPARAM wParam;
	LRESULT iCurSel1;
	LRESULT iItem;
	LRESULT iCurSel;
	LVFINDINFO lvfi;

	switch (pdm->flags & 0xF)
	{
		case 0u:
		case 1u:
		case 8u:
		case 0xBu:
			return 0;
		case 2u:
			lvfi.pt = pdm->pt;
			lvfi.vkDirection = VK_UP;
			goto LABEL_13;
		case 3u:
			goto LABEL_12;
		case 4u:
			lvfi.vkDirection = VK_END;
			goto LABEL_13;
		case 5u:
			wParam = pdm->pmsg->wParam;
			if (wParam == VK_UP)
			{
				iCurSel1 = _GetLVCurSel();
				iItem = SendMessageW(_hwndList, LVM_GETNEXTITEM, iCurSel1, 0x100);
			LABEL_7:
				pdm->itemID = iItem;
				return iItem != iCurSel1 && iItem >= 0;
			}
			if (wParam != VK_DOWN)
				return 0;
			iCurSel1 = _GetLVCurSel();
			if (iCurSel1 != -1)
			{
				iItem = SendMessage(this->_hwndList, LVM_GETNEXTITEM, iCurSel1, 0x200);
				goto LABEL_7;
			}
		LABEL_12:
			lvfi.vkDirection = VK_HOME;
		LABEL_13:
			lvfi.flags = 0x40;
			iItem1 = SendMessage(this->_hwndList, LVM_FINDITEMW, -1, (LPARAM)&lvfi);
		LABEL_14:
			pdm->itemID = iItem1;
			return iItem1 >= 0;
		case 6u:
			iCurSel = _GetLVCurSel();
			if (iCurSel < 0)
				return 0;
			_ActivateItem(iCurSel, pdm->flags & 0x400);
			return 1;
		case 7u:
			iItem1 = SendMessage(this->_hwndList, LVM_HITTEST, 0, (LPARAM)&pdm->pt);
			goto LABEL_14;
		case 9u:
		case 0xAu:
			return 1;
		default:
			ASSERT(!"Unknown SMNDM command"); // 513
			return 0;
	}
}

LRESULT CTopMatch::_OnSMNGetMinSize(PSMNGETMINSIZE pgms)
{
	pgms->field_14.cy = -1;
	pgms->siz.cy = _margins.cyTopHeight + _margins.cyBottomHeight + _cyIcon;
	return 0;
}

void CTopMatch::_SetTileWidth(int cxTile)
{
	LVTILEVIEWINFO tvi = {0};
	tvi.cbSize = sizeof(tvi);
	tvi.dwMask = LVTVIM_TILESIZE | LVTVIM_COLUMNS;
	tvi.dwFlags = LVTVIF_FIXEDWIDTH | LVTVIF_FIXEDHEIGHT;
	tvi.sizeTile.cx = cxTile;
	tvi.sizeTile.cy = _cyIcon;
	tvi.cLines = 0;
	ListView_SetTileViewInfo(_hwndList, &tvi);
}

void CTopMatch::_UpdateTopMatchSizeInOpenView()
{
	int cItems = ListView_GetItemCount(_hwndList);
	VARIANT varg;
	varg.vt = VT_I4;
	varg.iVal = _margins.cyTopHeight + _margins.cyBottomHeight + cItems * _cyIcon;
	IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &SID_SM_DV2ControlHost, 316, 0, &varg, NULL);
	VariantClear(&varg);
}

void CTopMatch::_AddSearchExtension()
{
	WCHAR szExtensionName[MAX_PATH];
	DWORD cbData = sizeof(szExtensionName);
	if (!_SHRegGetValueFromHKCUHKLM(
		TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer\\SearchExtensions"),
		TEXT("InternetExtensionName"),
		SRRF_RT_ANY,
		0,
		szExtensionName,
		&cbData))
	{
		WCHAR szExtensionNameOut[MAX_PATH];
		if (SUCCEEDED(SHLoadIndirectString(szExtensionName, szExtensionNameOut, ARRAYSIZE(szExtensionNameOut), NULL)))
		{
			_AddSearchItem(2, szExtensionNameOut);
		}
	}
}

void CTopMatch::_AddSearchItem(LPARAM lParam, LPWSTR pszText)
{
	HWND hwndList; // [esp-10h] [ebp-70h]
	HWND v5; // [esp-10h] [ebp-70h]
	LVITEMW lvi; // [esp+Ch] [ebp-54h] BYREF
	LVFINDINFOW lvfi; // [esp+48h] [ebp-18h] BYREF

	if ((!SHWindowsPolicy(POLID_NoSearchComputerLinkInStartMenu) || lParam)
		&& (!SHWindowsPolicy(POLID_NoSearchInternetLinkInStartMenu) || lParam != 1))
	{
		memset(&lvi.iItem, 0, 0x38u);
		lvi.pszText = pszText;
		hwndList = this->_hwndList;
		lvfi.flags = 1;
		lvi.mask = 5;
		lvi.lParam = lParam;
		lvfi.lParam = lParam;
		if (SendMessageW(hwndList, LVM_FINDITEMW, 0xFFFFFFFF, (LPARAM)&lvfi) < 0)
		{
			if (pszText)
			{
				lvi.mask |= 2u;
				v5 = this->_hwndList;
				lvi.iImage = 22;
				lvi.iItem = lParam;
				SendMessageW(v5, LVM_INSERTITEMW, 0, (LPARAM)&lvi);
			}
		}
	}
}

LRESULT CTopMatch::_ActivateItem(int iItem, int b)
{
	HRESULT hr; // ebx MAPDST
	HWND Parent; // eax
	HWND hwndList; // [esp-10h] [ebp-74h]
	LVITEM lvi; // [esp+Ch] [ebp-58h] BYREF
	NMHDR nm; // [esp+48h] [ebp-1Ch] BYREF

	hr = 0x80004005;
	//SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_TopResult_Launch, 73);
	memset(&lvi.iItem, 0, 0x38u);
	lvi.iItem = iItem;
	hwndList = this->_hwndList;
	lvi.mask = 4;
	if (SendMessageW(hwndList, LVM_GETITEMW, 0, (LPARAM)&lvi))
	{
		if (lvi.lParam)
		{
			if (lvi.lParam == 1)
			{
				if (SHWindowsPolicy(POLID_NoSearchInternetLinkInStartMenu))
					return hr >= 0;
				//SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Internet_Count, 118);
				hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenBox, &SID_SM_DV2ControlHost, 319, 0, 0, 0);
			}
			else
			{
				if (lvi.lParam != 2)
					return hr >= 0;
				hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenBox, &SID_SM_DV2ControlHost, 321, 0, 0, 0);
			}
		}
		else
		{
			if (SHWindowsPolicy(POLID_NoSearchComputerLinkInStartMenu))
				return hr >= 0;
			//SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Search_Computer_Count, 117);
			hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenBox, &SID_SM_DV2ControlHost, 322, 0, 0, 0);
		}
		if (hr >= 0)
		{
			Parent = GetParent(this->_hwnd);
			_SendNotify(Parent, 204u, &nm);
		}
	}
	return hr >= 0;
}

int CTopMatch::_GetLVCurSel()
{
	return ListView_GetNextItem(_hwndList, -1, LVIS_SELECTED);
}

BOOL TopMatch_RegisterClass()
{
	WNDCLASSEX wc;
	ZeroMemory(&wc, sizeof(wc));

	wc.cbSize = sizeof(wc);
	wc.style = CS_GLOBALCLASS;
	wc.lpfnWndProc = CTopMatch::s_WndProc;
	wc.hInstance = g_hinstCabinet;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.lpszClassName = WC_TOPMATCH;
	return RegisterClassEx(&wc);
}

#endif