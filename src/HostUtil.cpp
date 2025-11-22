#include "pch.h"
#include "stdafx.h"
#include "hostutil.h"

#define DEFAULT_BALLOON_TIMEOUT     (10*1000)       // 10 seconds

LRESULT CALLBACK BalloonTipSubclassProc(
                         HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
                         UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (uMsg)
    {
    case WM_TIMER:
        // Our autodismiss timer
        if (uIdSubclass == wParam)
        {
            KillTimer(hwnd, wParam);
            DestroyWindow(hwnd);
            return 0;
        }
        break;


    // On a settings change, recompute our size and margins
    case WM_SETTINGCHANGE:
        MakeMultilineTT(hwnd);
        break;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, BalloonTipSubclassProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

//
//  A "fire and forget" balloon tip.  Tell it where to go, what font
//  to use, and what to say, and it pops up and times out.
//
HWND CreateBalloonTip(HWND hwndOwner, int x, int y, HFONT hf,
                      UINT idsTitle, UINT idsText)
{
    DWORD dwStyle = TTS_ALWAYSTIP | TTS_BALLOON | TTS_NOPREFIX;

    HWND hwnd = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, dwStyle,
                               0, 0, 0, 0,
                               hwndOwner, NULL,
                               _AtlBaseModule.GetModuleInstance(), NULL);
    if (hwnd)
    {
        MakeMultilineTT(hwnd);

        TCHAR szBuf[MAX_PATH];
        TOOLINFO ti;
        ti.cbSize = sizeof(ti);
        ti.hwnd = hwndOwner;
        ti.uId = reinterpret_cast<UINT_PTR>(hwndOwner);
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS | TTF_TRACK;
        ti.hinst = _AtlBaseModule.GetResourceInstance();

        // We can't use MAKEINTRESOURCE because that allows only up to 80
        // characters for text, and our text can be longer than that.
        ti.lpszText = szBuf;
        if (LoadString(_AtlBaseModule.GetResourceInstance(), idsText, szBuf, ARRAYSIZE(szBuf)))
        {
            SendMessage(hwnd, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));

            if (idsTitle &&
                LoadString(_AtlBaseModule.GetResourceInstance(), idsTitle, szBuf, ARRAYSIZE(szBuf)))
            {
                SendMessage(hwnd, TTM_SETTITLE, TTI_INFO, reinterpret_cast<LPARAM>(szBuf));
            }

            SendMessage(hwnd, TTM_TRACKPOSITION, 0, MAKELONG(x, y));

            if (hf)
            {
                SetWindowFont(hwnd, hf, FALSE);
            }

            SendMessage(hwnd, TTM_TRACKACTIVATE, TRUE, reinterpret_cast<LPARAM>(&ti));

            // Set the autodismiss timer
            if (SetWindowSubclass(hwnd, BalloonTipSubclassProc, (UINT_PTR)hwndOwner, 0))
            {
                SetTimer(hwnd, (UINT_PTR)hwndOwner, DEFAULT_BALLOON_TIMEOUT, NULL);
            }
        }
    }

    return hwnd;
}

// Make the tooltip control multiline (infotip or balloon tip).
// The size computations are the same ones that comctl32 uses
// for listview and treeview infotips.

void MakeMultilineTT(HWND hwndTT)
{
    HWND hwndOwner = GetWindow(hwndTT, GW_OWNER);
    HDC hdc = GetDC(hwndOwner);
    if (hdc)
    {
        int iWidth = MulDiv(GetDeviceCaps(hdc, LOGPIXELSX), 300, 72);
        int iMaxWidth = GetDeviceCaps(hdc, HORZRES) * 3 / 4;
        SendMessage(hwndTT, TTM_SETMAXTIPWIDTH, 0, min(iWidth, iMaxWidth));

        static const RECT rcMargin = {4, 4, 4, 4};
        SendMessage(hwndTT, TTM_SETMARGIN, 0, (LPARAM)&rcMargin);

        ReleaseDC(hwndOwner, hdc);
    }
}



CPropBagFromReg::CPropBagFromReg(HKEY hk)
{
    _cref = 1;
    _hk = hk;
};
CPropBagFromReg::~CPropBagFromReg()
{
    RegCloseKey(_hk);
}

STDMETHODIMP CPropBagFromReg::QueryInterface(REFIID riid, PVOID *ppvObject)
{
    if (IsEqualIID(riid, IID_IPropertyBag))
        *ppvObject = (IPropertyBag *)this;
    else if (IsEqualIID(riid, IID_IUnknown))
        *ppvObject = this;
    else
    {
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG CPropBagFromReg::AddRef(void)
{
    return ++_cref; // on the stack
}
ULONG CPropBagFromReg::Release(void)
{
    if (--_cref)
        return _cref;

    delete this;
    return 0;
}

STDMETHODIMP CPropBagFromReg::Read(LPCOLESTR pszPropName, VARIANT *pVar, IErrorLog *pErrorLog)
{
    VARTYPE vtDesired = pVar->vt;

    WCHAR szTmp[100];
    DWORD cb = sizeof(szTmp);
    DWORD dwType;
    if (ERROR_SUCCESS == RegQueryValueExW(_hk, pszPropName, NULL, &dwType, (LPBYTE)szTmp, &cb) && (REG_SZ==dwType))
    {
        // TODO - use dwType to set the vt properly
        pVar->bstrVal = SysAllocString(szTmp);
        if (pVar->bstrVal)
        {
            pVar->vt = VT_BSTR;
            return VariantChangeTypeForRead(pVar, vtDesired);
        }
        else
            return E_OUTOFMEMORY;
    }
    else
        return E_INVALIDARG;

}

HRESULT CreatePropBagFromReg(LPCTSTR pszKey, IPropertyBag**pppb)
{
    HRESULT hr = E_OUTOFMEMORY;

    *pppb = NULL;

    // Try current user 1st, if that fails, fall back to localmachine
    HKEY hk;
    if (ERROR_SUCCESS == RegOpenKeyEx(HKEY_CURRENT_USER, pszKey, NULL, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE, &hk)
     || ERROR_SUCCESS == RegOpenKeyEx(HKEY_LOCAL_MACHINE, pszKey, NULL, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE, &hk))
    {
        CPropBagFromReg* pcpbfi = new CPropBagFromReg(hk);
        if (pcpbfi)
        {
            hr = pcpbfi->QueryInterface(IID_IPropertyBag, (void**) pppb);
            pcpbfi->Release();
        }
        else
        {
            RegCloseKey(hk);
        }
    }

    return hr;
};

BOOL RectFromStrW(LPCWSTR pwsz, RECT *pr)
{
    pr->left = StrToIntW(pwsz);
    pwsz = StrChrW(pwsz, L',');
    if (!pwsz)
        return FALSE;
    pr->top = StrToIntW(++pwsz);
    pwsz = StrChrW(pwsz, L',');
    if (!pwsz)
        return FALSE;
    pr->right = StrToIntW(++pwsz);
    pwsz = StrChrW(pwsz, L',');
    if (!pwsz)
        return FALSE;
    pr->bottom = StrToIntW(++pwsz);
    return TRUE;
}

LRESULT HandleApplyRegion(HWND hwnd, HTHEME hTheme,
                          PSMNMAPPLYREGION par, int iPartId, int iStateId)
{
    if (hTheme)
    {
        RECT rc;
        GetWindowRect(hwnd, &rc);

        // Map to caller's coordinates
        MapWindowRect(NULL, par->hdr.hwndFrom, &rc);

        HRGN hrgn;
        if (SUCCEEDED(GetThemeBackgroundRegion(hTheme, NULL, iPartId, iStateId, &rc, &hrgn)) && hrgn)
        {
            // Replace our window rectangle with the region
            HRGN hrgnRect = CreateRectRgnIndirect(&rc);
            if (hrgnRect)
            {
                // We want to take par->hrgn, subtract hrgnRect and add hrgn.
                // But we want to do this with a single operation to par->hrgn
                // so we don't end up with a corrupted region on low memory failure.
                // So we do
                //
                //  par->hrgn ^= hrgnRect ^ hrgn.
                //
                // If hrgnRect ^ hrgn == NULLREGION then the background
                // does not want to customize the rectangle so we can just
                // leave par->hrgn alone.

                int iResult = CombineRgn(hrgn, hrgn, hrgnRect, RGN_XOR);
                if (iResult != ERROR && iResult != NULLREGION)
                {
                    CombineRgn(par->hrgn, par->hrgn, hrgn, RGN_XOR);
                }
                DeleteObject(hrgnRect);
            }
            DeleteObject(hrgn);
        }
    }
    return 0;
}

void HandleApplyRegionFromRect(const RECT& rc, HTHEME hTheme, SMNMAPPLYREGION* par, int iPartId, int iStateId)
{
    HRGN hrgn;
    if (SUCCEEDED(GetThemeBackgroundRegion(hTheme, nullptr, iPartId, iStateId, &rc, &hrgn)) && hrgn)
    {
        HRGN hrgnRect = CreateRectRgnIndirect(&rc);
        if (hrgnRect)
        {
            int iResult = CombineRgn(hrgn, hrgn, hrgnRect, RGN_XOR);
            if (iResult != ERROR && iResult != NULLREGION)
            {
                CombineRgn(par->hrgn, par->hrgn, hrgn, RGN_XOR);
            }
            DeleteObject(hrgnRect);
        }
        DeleteObject(hrgn);
    }
}

//****************************************************************************
//
//  CAccessible - Most of this class is just forwarders

#define ACCESSIBILITY_FORWARD(fn, typedargs, args)  \
HRESULT CAccessible::fn typedargs                   \
{                                                   \
    if (_paccInner)                                 \
    {                                               \
        return _paccInner->fn args;                 \
    }                                               \
    else                                            \
    {                                               \
        return E_FAIL;                              \
    }                                               \
}                                                   \

#define ENUMVARIANT_FORWARD(fn, typedargs, args)    \
HRESULT CAccessible::fn typedargs                   \
{                                                   \
    if (_pevarInner)                                \
    {                                               \
        return _pevarInner->fn args;                \
    }                                               \
    else                                            \
    {                                               \
        return E_FAIL;                              \
    }                                               \
}                                                   \

ACCESSIBILITY_FORWARD(get_accParent,
                      (IDispatch **ppdispParent),
                      (ppdispParent))
ACCESSIBILITY_FORWARD(GetTypeInfoCount,
                      (UINT *pctinfo),
                      (pctinfo))
ACCESSIBILITY_FORWARD(GetTypeInfo,
                      (UINT itinfo, LCID lcid, ITypeInfo **pptinfo),
                      (itinfo, lcid, pptinfo))
ACCESSIBILITY_FORWARD(GetIDsOfNames,
                      (REFIID riid, OLECHAR **rgszNames, UINT cNames,
                       LCID lcid, DISPID *rgdispid),
                      (riid, rgszNames, cNames, lcid, rgdispid))
ACCESSIBILITY_FORWARD(Invoke,
                      (DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags,
                       DISPPARAMS *pdispparams, VARIANT *pvarResult,
                       EXCEPINFO *pexcepinfo, UINT *puArgErr),
                      (dispidMember, riid, lcid, wFlags,
                       pdispparams, pvarResult,
                       pexcepinfo, puArgErr))
ACCESSIBILITY_FORWARD(get_accChildCount,
                      (long *pChildCount),
                      (pChildCount))
ACCESSIBILITY_FORWARD(get_accChild,
                      (VARIANT varChildIndex, IDispatch **ppdispChild),
                      (varChildIndex, ppdispChild))
ACCESSIBILITY_FORWARD(get_accName,
                      (VARIANT varChild, BSTR *pszName),
                      (varChild, pszName))
ACCESSIBILITY_FORWARD(get_accValue,
                      (VARIANT varChild, BSTR *pszValue),
                      (varChild, pszValue))
ACCESSIBILITY_FORWARD(get_accDescription,
                      (VARIANT varChild, BSTR *pszDescription),
                      (varChild, pszDescription))
ACCESSIBILITY_FORWARD(get_accRole,
                      (VARIANT varChild, VARIANT *pvarRole),
                      (varChild, pvarRole))
ACCESSIBILITY_FORWARD(get_accState,
                      (VARIANT varChild, VARIANT *pvarState),
                      (varChild, pvarState))
ACCESSIBILITY_FORWARD(get_accHelp,
                      (VARIANT varChild, BSTR *pszHelp),
                      (varChild, pszHelp))
ACCESSIBILITY_FORWARD(get_accHelpTopic,
                      (BSTR *pszHelpFile, VARIANT varChild, long *pidTopic),
                      (pszHelpFile, varChild, pidTopic))
ACCESSIBILITY_FORWARD(get_accKeyboardShortcut,
                      (VARIANT varChild, BSTR *pszKeyboardShortcut),
                      (varChild, pszKeyboardShortcut))
ACCESSIBILITY_FORWARD(get_accFocus,
                      (VARIANT *pvarFocusChild),
                      (pvarFocusChild))
ACCESSIBILITY_FORWARD(get_accSelection,
                      (VARIANT *pvarSelectedChildren),
                      (pvarSelectedChildren))
ACCESSIBILITY_FORWARD(get_accDefaultAction,
                      (VARIANT varChild, BSTR *pszDefaultAction),
                      (varChild, pszDefaultAction))
ACCESSIBILITY_FORWARD(accSelect,
                      (long flagsSelect, VARIANT varChild),
                      (flagsSelect, varChild))
ACCESSIBILITY_FORWARD(accLocation,
                      (long *pxLeft, long *pyTop, long *pcxWidth, long *pcyHeight, VARIANT varChild),
                      (pxLeft, pyTop, pcxWidth, pcyHeight, varChild))
ACCESSIBILITY_FORWARD(accNavigate,
                      (long navDir, VARIANT varStart, VARIANT *pvarEndUpAt),
                      (navDir, varStart, pvarEndUpAt))
ACCESSIBILITY_FORWARD(accHitTest,
                      (long xLeft, long yTop, VARIANT *pvarChildAtPoint),
                      (xLeft, yTop, pvarChildAtPoint))
ACCESSIBILITY_FORWARD(accDoDefaultAction,
                      (VARIANT varChild),
                      (varChild));
ACCESSIBILITY_FORWARD(put_accName,
                      (VARIANT varChild, BSTR szName),
                      (varChild, szName))
ACCESSIBILITY_FORWARD(put_accValue,
                      (VARIANT varChild, BSTR pszValue),
                      (varChild, pszValue));

ENUMVARIANT_FORWARD(Next,
                   (ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched),
                   (celt, rgVar, pCeltFetched));
ENUMVARIANT_FORWARD(Skip,
                   (ULONG celt),
                   (celt));
ENUMVARIANT_FORWARD(Reset,
                   (),
	               ());
ENUMVARIANT_FORWARD(Clone,
                   (IEnumVARIANT **ppEnum),
	               (ppEnum));

HRESULT CAccessible::GetInnerObject(HWND hwnd, LONG idObject)
{
    if (_pevarInner)
        return S_OK;

    HRESULT hr = CreateStdAccessibleObject(hwnd, idObject, IID_PPV_ARGS(&this->_paccInner));
    if (SUCCEEDED(hr))
    {
        hr = _paccInner->QueryInterface(IID_PPV_ARGS(&_pevarInner));
        if (FAILED(hr))
        {
            IUnknown_SafeReleaseAndNullPtr(&_paccInner);
        }
        return hr;
    }
    return hr;
}

LRESULT CALLBACK CAccessible::s_SubclassProc(
                         HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
                         UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    CAccessible *self = reinterpret_cast<CAccessible *>(dwRefData);

    switch (uMsg)
    {

    case WM_GETOBJECT:
        if ((DWORD)lParam == OBJID_CLIENT)
        {
            HRESULT hr = self->GetInnerObject(hwnd, (LONG)lParam);
            ASSERT((self->_paccInner == NULL) == (self->_pevarInner == NULL)); // 424
            if (SUCCEEDED(hr))
            {
                return LresultFromObject(IID_IAccessible, wParam, SAFECAST(self, IAccessible *));
            }
            else
            {
                return hr;
            }
        };
        break;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, s_SubclassProc, 0);
        break;

    case 0x10C1:
        if (self && self->field_10)
            return S_OK;
        break;
    }
    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

HRESULT CAccessible::GetRoleString(DWORD dwRole, BSTR *pbsOut)
{
    *pbsOut = NULL;

    WCHAR szBuf[MAX_PATH];
    if (GetRoleTextW(dwRole, szBuf, ARRAYSIZE(szBuf)))
    {
        *pbsOut = SysAllocString(szBuf);
    }

    return *pbsOut ? S_OK : E_OUTOFMEMORY;
}

HRESULT CAccessible::CreateAcceleratorBSTR(TCHAR tch, BSTR *pbsOut)
{
    TCHAR sz[2] = { tch, 0 };
    *pbsOut = SysAllocString(sz);
    return *pbsOut ? S_OK : E_OUTOFMEMORY;
}

// EXEX-TODO: Move? Currently in TaskBand.cpp
extern int g_iLPX;
extern int g_iLPY;
extern BOOL g_fHighDPI;

extern void InitDPI();

BOOL IsHighDPI()
{
    InitDPI();
    return g_fHighDPI;
}

int SHGetSystemMetricsScaled(int nIndex)
{
    int SystemMetrics; // eax MAPDST
    int result; // eax

    InitDPI();
    SystemMetrics = GetSystemMetrics(nIndex);
    switch (nIndex)
    {
    case SM_CXSCREEN:
    case SM_CXVSCROLL:
    case SM_CXBORDER:
    case SM_CXDLGFRAME:
    case SM_CXHTHUMB:
    case SM_CXICON:
    case SM_CXCURSOR:
    case SM_CXFULLSCREEN:
    case SM_CXHSCROLL:
    case SM_CXMIN:
    case SM_CXSIZE:
    case SM_CXFRAME:
    case SM_CXMINTRACK:
    case SM_CXDOUBLECLK:
    case SM_CXICONSPACING:
    case SM_CXEDGE:
    case SM_CXMINSPACING:
    case SM_CXSMICON:
    case SM_CXSMSIZE:
    case SM_CXMENUSIZE:
    case SM_CXMINIMIZED:
    case SM_CXMAXTRACK:
    case SM_CXMAXIMIZED:
    case SM_CXDRAG:
    case SM_CXMENUCHECK:
    case SM_XVIRTUALSCREEN:
    case SM_CXVIRTUALSCREEN:
    case SM_CXFOCUSBORDER:
        result = MulDiv(SystemMetrics, g_iLPX, 96);
        break;
    case SM_CYSCREEN:
    case SM_CYHSCROLL:
    case SM_CYCAPTION:
    case SM_CYBORDER:
    case SM_CYDLGFRAME:
    case SM_CYVTHUMB:
    case SM_CYICON:
    case SM_CYCURSOR:
    case SM_CYMENU:
    case SM_CYFULLSCREEN:
    case SM_CYKANJIWINDOW:
    case SM_CYVSCROLL:
    case SM_CYMIN:
    case SM_CYSIZE:
    case SM_CYFRAME:
    case SM_CYMINTRACK:
    case SM_CYDOUBLECLK:
    case SM_CYICONSPACING:
    case SM_CYEDGE:
    case SM_CYMINSPACING:
    case SM_CYSMICON:
    case SM_CYSMCAPTION:
    case SM_CYSMSIZE:
    case SM_CYMENUSIZE:
    case SM_CYMINIMIZED:
    case SM_CYMAXTRACK:
    case SM_CYMAXIMIZED:
    case SM_CYDRAG:
    case SM_CYMENUCHECK:
    case SM_YVIRTUALSCREEN:
    case SM_CYVIRTUALSCREEN:
    case SM_CYFOCUSBORDER:
        result = MulDiv(SystemMetrics, g_iLPY, 96);
        break;
    default:
        ASSERTMSG(FALSE, "SHGetSystemMetricsScaled called with non-scaling metric!");
        result = SystemMetrics;
        break;
    }
    return result;
}

HBITMAP CreateBitmap(HDC hdc, int cx, int cy)
{
    if (!IsCompositionActive())
        return CreateCompatibleBitmap(hdc, cx, cy);

    BITMAPINFO bmi = { 0 };
    bmi.bmiHeader.biWidth = cx;
    bmi.bmiHeader.biHeight = -cy;
    bmi.bmiHeader.biCompression = 0;
    bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    return CreateDIBSection(hdc, &bmi, 0, 0, 0, 0);
}