#include "pch.h"
#include "cocreateinstancehook.h"
#include "shguidp.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "hostutil.h"
#include "moreprog.h"
#include "tray.h"           // To get access to c_tray

//
//  Unfortunately, WTL #undef's SelectFont, so we have to define it again.
//

inline HFONT SelectFont(HDC hdc, HFONT hf)
{
    return (HFONT)SelectObject(hdc, hf);
}

// EXEX-VISTA(allison): Validated.
CMorePrograms::CMorePrograms(HWND hwnd) :
    _lRef(1),
    _hwnd(hwnd),
    _clrText(CLR_INVALID),
    _clrBk(CLR_INVALID),
    dwordB8(1)
{
}

// EXEX-VISTA(allison): Validated.
CMorePrograms::~CMorePrograms()
{
    if (_hf)
      DeleteObject(_hf);

    if (_hfTTBold)
      DeleteObject(_hfTTBold);

    if (_hfMarlett)
      DeleteObject(_hfMarlett);

    IUnknown_SafeReleaseAndNullPtr(&_pdth);

    // Note that we do not need to clean up our HWNDs.
    // USER does that for us automatically.
}

//
//  Metrics changed -- update.
//
// EXEX-VISTA(allison): Validated.
void CMorePrograms::_InitMetrics()
{
    if (_hwndTT)
    {
        MakeMultilineTT(_hwndTT);

        // Disable/enable infotips based on user preference
        SendMessage(_hwndTT, TTM_ACTIVATE, ShowInfoTip(), 0);
    }
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
        
    CMorePrograms *self = new CMorePrograms(hwnd);

    if (self)
    {
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LPARAM)self);
        return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return FALSE;
}

//
//  Create an inner button that is exactly the right size.
//
//  Height of inner button = height of text.
//  Width of inner button = full width.
//
//  This allows us to let USER do most of the work of hit-testing and
//  focus rectangling.
//

// EXEX-VISTA(allison): Validated. Still needs slight cleanup.
LRESULT CMorePrograms::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPCREATESTRUCT lpcs = (LPCREATESTRUCT)lParam;
    SMPANEDATA *psmpd = (SMPANEDATA *)lpcs->lpCreateParams;

    IUnknown_Set(&psmpd->punk, SAFECAST(this, IServiceProvider *));

    _hTheme = psmpd->hTheme;
    if (_hTheme)
    {
        GetThemeColor(_hTheme, SPP_MOREPROGRAMS, 0, TMT_TEXTCOLOR, &_clrText);
        _hbrBk = GetStockBrush(HOLLOW_BRUSH);
        _colorHighlight = COLOR_MENUHILIGHT;
        _colorHighlightText = COLOR_HIGHLIGHTTEXT;

        GetThemeMargins(_hTheme, NULL, SPP_MOREPROGRAMS, 0, TMT_CONTENTMARGINS, NULL, &_margins);

        SIZE siz = {0};
        GetThemePartSize(_hTheme, NULL, SPP_MOREPROGRAMSARROW, 0, NULL, TS_TRUE, &siz);
        _cxArrow = siz.cx;
    }
    else
    {
        _clrText = GetSysColor(COLOR_MENUTEXT);
        _clrBk = GetSysColor(COLOR_MENU);
        _hbrBk = GetSysColorBrush(COLOR_MENU);
        _colorHighlight = COLOR_HIGHLIGHT;
        _colorHighlightText = COLOR_HIGHLIGHTTEXT;
        _margins.cxLeftWidth = 2 * GetSystemMetrics(SM_CXEDGE);
        _margins.cxRightWidth = 2 * GetSystemMetrics(SM_CXEDGE);
    }

    if (!SHRestricted(REST_NOSMMOREPROGRAMS))
    {
        if (!LoadString(g_hinstCabinet,     8226,    _szMessage,        ARRAYSIZE(_szMessage))
            || !LoadString(g_hinstCabinet,  8241,    _szMessageBack,    ARRAYSIZE(_szMessageBack))
            || !LoadString(g_hinstCabinet,  8227,    _szTool,           ARRAYSIZE(_szTool))
            || !LoadString(g_hinstCabinet,  8245,    _szToolBack,       ARRAYSIZE(_szToolBack)))
        {
            return -1;
        }

        _chMnem = (WCHAR)CharUpper((LPWSTR)SHFindMnemonic(_szMessage));
        _chMnemBack = (WCHAR)CharUpper((LPWSTR)SHFindMnemonic(_szMessageBack));

        _hf = LoadControlFont(_hTheme, SPP_MOREPROGRAMS, FALSE, 0);

        HDC hdc = GetDC(hwnd);
        if (hdc)
        {
            HFONT hfPrev = (HFONT)SelectObject(hdc, _hf);
            if (hfPrev)
            {
                SIZE sizText;
                GetTextExtentPoint32(hdc, _szMessage, lstrlen(_szMessage), &sizText);
                _cxText = sizText.cx + SHGetSystemMetricsScaled(SM_CXEDGE);

                GetTextExtentPoint32(hdc, _szMessageBack, lstrlen(_szMessageBack), &sizText);
                _cxText2 = sizText.cx + SHGetSystemMetricsScaled(SM_CXEDGE);

                TEXTMETRIC tm;
                if (GetTextMetrics(hdc, &tm))
                {
                    _tmAscent = tm.tmAscent;

                    LOGFONT lf;
                    ZeroMemory(&lf, sizeof(lf));

                    lf.lfHeight = _tmAscent;
                    lf.lfWeight = FW_NORMAL;
                    lf.lfCharSet = SYMBOL_CHARSET;
                    StringCchCopy(lf.lfFaceName, ARRAYSIZE(lf.lfFaceName), TEXT("Marlett"));
                    _hfMarlett = CreateFontIndirect(&lf);

                    if (_hfMarlett)
                    {
                        SelectFont(hdc, _hfMarlett);
                        if (GetTextMetrics(hdc, &tm))
                        {
                            _tmAscentMarlett = tm.tmAscent;
                        }
                    }
                }
                SelectFont(hdc, hfPrev);
            }
            ReleaseDC(hwnd, hdc);
        }

        if (!_tmAscentMarlett)
            return -1;

        BOOL bLargeIcons = _SHRegGetBoolValueFromHKCUHKLM(REGSTR_EXPLORER_ADVANCED, REGSTR_VAL_DV2_LARGEICONS, TRUE /* default to large*/);

        RECT rc;
        GetClientRect(_hwnd, &rc);
        rc.left += _margins.cxLeftWidth;
        rc.right -= _margins.cxRightWidth;
        rc.top += _margins.cyTopHeight;
        rc.bottom -= _margins.cyBottomHeight;

        _cxTextIndent = 3 * GetSystemMetrics(SM_CXEDGE) +
            GetSystemMetrics(bLargeIcons ? SM_CXICON : SM_CXSMICON);

        ASSERT(RECTHEIGHT(rc) > _tmAscent); // 208

        _iTextCenterVal = (RECTHEIGHT(rc) - _tmAscent) / 2;

        _hwndButton = SHFusionCreateWindowEx(
            0,
            WC_BUTTON,
            _szMessage,
            0x5600000B,
            rc.left,
            rc.top,
            RECTWIDTH(rc),
            RECTHEIGHT(rc),
            _hwnd,
            (HMENU)IDC_ALL,
            g_hinstCabinet,
            NULL);

        if (!_hwndButton)
            return -1;

        CAccessible::SetAccessibleSubclassWindow(_hwndButton);

        if (_hf)
            SetWindowFont(_hwndButton, _hf, FALSE);

        CoCreateInstanceHook(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pdth));

        RegisterDragDrop(_hwndButton, this);

        _hwndTT = _CreateTooltip();

        _TooltipAddTool();

        _InitMetrics();
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
void CMorePrograms::_TooltipAddTool()
{
    if (_hwndTT)
    {
        TOOLINFO ti;
        ti.hwnd = _hwnd;
        ti.uId = (UINT_PTR)_hwndButton;
        ti.hinst = g_hinstCabinet;
        ti.cbSize = sizeof(ti);
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;

        OPENHOSTVIEW view;
        if (SUCCEEDED(_GetCurView(&view)))
        {
            SendMessage(_hwndTT, TTM_DELTOOL, 0, (LPARAM)&ti);
            ti.lpszText = view == OHVIEW_0 ? _szTool : _szToolBack;
            SendMessage(_hwndTT, TTM_ADDTOOL, 0, (LPARAM)&ti);
        }
    }
}

HWND CMorePrograms::_CreateTooltip()
{
    DWORD dwStyle = WS_BORDER | TTS_NOPREFIX;

    HWND hwnd = CreateWindowEx(0, TOOLTIPS_CLASS, NULL, dwStyle,
                               0, 0, 0, 0,
                               _hwndButton, NULL,
                               _AtlBaseModule.GetModuleInstance(), NULL);
    if (hwnd)
    {
        TCHAR szBuf[MAX_PATH];
        TOOLINFO ti;
        ti.cbSize = sizeof(ti);
        ti.hwnd = _hwnd;
        ti.uId = reinterpret_cast<UINT_PTR>(_hwndButton);
        ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        ti.hinst = _AtlBaseModule.GetResourceInstance();

        // We can't use MAKEINTRESOURCE because that allows only up to 80
        // characters for text, and our text can be longer than that.
        UINT ids = IDS_STARTPANE_MOREPROGRAMS_TIP;

        ti.lpszText = szBuf;
        if (LoadString(_AtlBaseModule.GetResourceInstance(), ids, szBuf, ARRAYSIZE(szBuf)))
        {
            SendMessage(hwnd, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&ti));

        }
    }

    return hwnd;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    RevokeDragDrop(_hwndButton);
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}


LRESULT CMorePrograms::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // WARNING!  "this" might be invalid (if WM_NCCREATE failed), so
    // do not use any member variables!
    LRESULT lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
    SetWindowLongPtr(hwnd, 0, 0);
    if (this)
    {
        Release();
    }
    return lres;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnCtlColorBtn(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    HDC hdc = reinterpret_cast<HDC>(wParam);

    if (_clrText != CLR_INVALID)
    {
        SetTextColor(hdc, _clrText);
    }

    if (_clrBk != CLR_INVALID)
    {
        SetBkColor(hdc, _clrBk);
    }

    return reinterpret_cast<LRESULT>(_hbrBk);
}

BOOL
WINAPI
SHExtTextOutW(
    HDC hdc,
    int x,
    int y,
    UINT options,
    CONST RECT *lprect,
    LPCWSTR lpString,
    UINT c,
    CONST INT *lpDx);

// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::_GetCurView(OPENHOSTVIEW *pView)
{   
    VARIANT var;
	var.vt = VT_I4;
    HRESULT hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &SID_SM_DV2ControlHost, 303, 0, 0, &var);
    if (SUCCEEDED(hr))
    {
        ASSERT(var.vt == VT_I4); // 487
        *pView = (OPENHOSTVIEW)var.lVal;
    }
    return hr;
}

// EXEX-VISTA(allison): Validated.
int CMorePrograms::_OnSetCurView(OPENHOSTVIEW view)
{
    InvalidateRect(_hwnd, NULL, TRUE);
    field_A0 = 1;
    return 0;
}

// EXEX-VISTA(allison): Validated. Still needs major cleanup.
LRESULT CMorePrograms::_OnDrawItem(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    LPDRAWITEMSTRUCT pdis = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);
    ASSERT(pdis->CtlType == ODT_BUTTON);
    ASSERT(pdis->CtlID == IDC_BUTTON);

    if (pdis->itemAction & (ODA_DRAWENTIRE | ODA_FOCUS))
    {
        BOOL fRTLReading = GetLayout(pdis->hDC) & LAYOUT_RTL;
        UINT fuOptions = 0;
        if (fRTLReading)
        {
            fuOptions |= ETO_RTLREADING;
        }

        HFONT hfPrev = SelectFont(pdis->hDC, _hf);
        if (hfPrev)
        {
            BOOL fHot = (pdis->itemState & ODS_FOCUS) || _tmHoverStart;
            if (fHot)
            {
                // hot background
                FillRect(pdis->hDC, &pdis->rcItem, GetSysColorBrush(_colorHighlight));
                SetTextColor(pdis->hDC, GetSysColor(_colorHighlightText));
            }
            else if (_hTheme)
            {
                // Themed non-hot background = custom
                RECT rc;
                GetClientRect(hwnd, &rc);
                MapWindowRect(hwnd, pdis->hwndItem, &rc);
                DrawThemeBackground(_hTheme, pdis->hDC, SPP_MOREPROGRAMS, 0, &rc, 0);
            }
            else
            {
                // non-themed non-hot background
                FillRect(pdis->hDC, &pdis->rcItem, _hbrBk);
            }

            int iOldMode = SetBkMode(pdis->hDC, TRANSPARENT);

            // _cxTextIndent will move it in the current width of an icon (small or large), plus the space we add between an icon and the text
            pdis->rcItem.left += _cxTextIndent;

            UINT dtFlags = DT_VCENTER | DT_SINGLELINE | DT_EDITCONTROL;
            if (fRTLReading)
            {
                dtFlags |= DT_RTLREADING;
            }
            if (pdis->itemState & ODS_NOACCEL)
            {
                dtFlags |= DT_HIDEPREFIX;
            }

            DrawText(pdis->hDC, _szMessage, -1, &pdis->rcItem, dtFlags);

            RECT rc = pdis->rcItem;
            rc.left += _cxText;

            if (_hTheme)
            {
                if (_iTextCenterVal < 0) // text is taller than the bitmap
                    rc.top += (-_iTextCenterVal);

                rc.right = rc.left + _cxArrow;       // clip rectangle down to the minumum size...
                DrawThemeBackground(_hTheme, pdis->hDC, SPP_MOREPROGRAMSARROW,
                    fHot ? SPS_HOT : 0, &rc, 0);
            }
            else
            {
                if (SelectFont(pdis->hDC, _hfMarlett))
                {
                    rc.top = rc.top + _tmAscent - _tmAscentMarlett + (_iTextCenterVal > 0 ? _iTextCenterVal : 0);
                    TCHAR chOut = fRTLReading ? TEXT('w') : TEXT('8');
                    if (EVAL(!IsRectEmpty(&rc)))
                    {
                        ExtTextOut(pdis->hDC, rc.left, rc.top, fuOptions, &rc, &chOut, 1, NULL);
                        rc.right = rc.left + _cxArrow;
                    }
                }
            }

            _rcExclude = rc;
            _rcExclude.left -= _cxText;  // includes the text in the exclusion rectangle.

            MapWindowRect(pdis->hwndItem, NULL, &_rcExclude);
            SetBkMode(pdis->hDC, iOldMode);
            SelectFont(pdis->hDC, hfPrev);
        }
    }

    //
    //  Since we are emulating a menu item, we don't need to draw a
    //  focus rectangle.
    //
    return TRUE;
#else
    HBITMAP Bitmap; // eax
    bool v8; // zf
    HTHEME hTheme; // eax
    HBRUSH SysColorBrush; // eax
    COLORREF SysColor; // eax
    HBRUSH v12; // eax
    DWORD cxText2; // edx
    int cxTextIndent; // eax
    WCHAR *szMessage; // edi
    HTHEME v16; // ecx
    int iTextCenterVal; // eax
    int v18; // eax
    DTTOPTS pOptions; // [esp+10h] [ebp-A8h] BYREF
    int mode; // [esp+50h] [ebp-68h]
    HGDIOBJ h; // [esp+54h] [ebp-64h]
    HGDIOBJ v23; // [esp+58h] [ebp-60h]
    DWORD v24; // [esp+5Ch] [ebp-5Ch]
    HGDIOBJ ho; // [esp+60h] [ebp-58h]
    int v26; // [esp+64h] [ebp-54h]
    RECT rc; // [esp+68h] [ebp-50h] BYREF
    RECT pRect; // [esp+78h] [ebp-40h] BYREF
    int iStateId; // [esp+88h] [ebp-30h]
    OPENHOSTVIEW view; // [esp+94h] [ebp-24h] BYREF
    int v33; // [esp+98h] [ebp-20h] BYREF
    DWORD dwTextFlags; // [esp+9Ch] [ebp-1Ch]
    HDC hdc; // [esp+CCh] [ebp+14h]

    LPDRAWITEMSTRUCT pdis = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ASSERT(pdis->CtlType == ODT_BUTTON); // 323
	ASSERT(pdis->CtlID == IDC_ALL);      // 324

    if (this->_pdth)
        this->_pdth->Show(0);

    if ((pdis->itemAction & 0x45) != 0)
    {
        hdc = CreateCompatibleDC(pdis->hDC);
        if (hdc)
        {
            pRect.left = 0;
            pRect.top = 0;
            pRect.right = pdis->rcItem.right - pdis->rcItem.left;
            pRect.bottom = pdis->rcItem.bottom - pdis->rcItem.top;
            
            Bitmap = CreateBitmap(pdis->hDC, pRect.right, pRect.bottom);
            ho = Bitmap;
            if (!Bitmap)
            {
            LABEL_70:
                DeleteDC(hdc);
                goto LABEL_71;
            }
            
            v23 = SelectObject(hdc, Bitmap);
            v24 = GetLayout(pdis->hDC) & 1;
            v26 = 0;
            if (v24)
                v26 = 128;
            h = SelectObject(hdc, this->_hf);
            if (!h)
            {
            LABEL_69:
                SelectObject(hdc, v23);
                DeleteObject(ho);
                goto LABEL_70;
            }
            if ((pdis->itemState & 0x50) != 0 || this->_tmHoverStart || (v8 = this->field_BC == 0, v33 = 0, !v8))
                v33 = 1;

            view = OHVIEW_0;
            _GetCurView(&view);

            iStateId = 1;
            hTheme = this->_hTheme;
            if (hTheme)
            {
                if (v33)
                {
                    iStateId = 2;
                }
                else if (this->field_B4)
                {
                    iStateId = 3;
                }
                DrawThemeBackground(hTheme, hdc, 12, iStateId, &pRect, 0);
                goto LABEL_33;
            }
            if (v33)
            {
                SysColorBrush = GetSysColorBrush(this->_colorHighlight);
                FillRect(hdc, &pRect, SysColorBrush);
                SysColor = GetSysColor(this->_colorHighlightText);
            }
            else
            {
                if (!this->field_B4)
                {
                    FillRect(hdc, &pRect, this->_hbrBk);
                    SetTextColor(hdc, this->_clrText);
                LABEL_33:
                    mode = SetBkMode(hdc, 1);
                    ASSERT(pdis->CtlID == IDC_ALL); // 391
                    if (view)
                        cxText2 = this->_cxText2;
                    else
                        cxText2 = this->_cxText;
                    dwTextFlags = cxText2;
                    cxTextIndent = this->_cxTextIndent;
                    if (cxTextIndent > (int)(pRect.right - cxText2 - pRect.left))
                    {
                        //CcshellDebugMsgW(
                        //    1,
                        //    "StartMenu: 'maximum of (%s, %s) ' is %dpx, only room for %d- notify localizers!",
                        //    (const char *)this->_szMessage,
                        //    (const char *)this->_szMessageBack,
                        //    cxText2,
                        //    pRect.right - cxText2 - pRect.left);
                        cxTextIndent = pRect.right - dwTextFlags - pRect.left;
                        if (cxTextIndent < 0)
                            cxTextIndent = 0;
                    }
                    pRect.left += cxTextIndent;
                    dwTextFlags = 0x2024;
					ASSERT(pdis->CtlID == IDC_ALL); // 403
                    szMessage = this->_szMessage;
                    if (view)
                        szMessage = this->_szMessageBack;

                    if (v24)
                        dwTextFlags |= 0x20000u;
                    if ((pdis->itemState & 0x100) != 0)
                        dwTextFlags |= 0x100000u;

                    if (this->_hTheme)
                    {
                        pOptions.dwSize = 64;
                        memset(&pOptions.dwFlags, 0, 0x3Cu);
                        pOptions.dwFlags = IsCompositionActive() ? 0x2000 : 0;
                        DrawThemeTextEx(this->_hTheme, hdc, 12, iStateId, szMessage, -1, dwTextFlags, &pRect, &pOptions);
                    }
                    else
                    {
                        DrawTextW(hdc, szMessage, -1, &pRect, dwTextFlags);
                    }
                    rc = pRect;
                    rc.left = GetSystemMetrics(45);
                    v16 = this->_hTheme;
                    if (v16)
                    {
                        iTextCenterVal = this->_iTextCenterVal;
                        if (iTextCenterVal < 0)
                            rc.top -= iTextCenterVal;
                        rc.right = rc.left + this->_cxArrow;
                        DrawThemeBackground(v16, hdc, view != OHVIEW_0 ? 17 : 3, v33 != 0 ? 2 : 0, &rc, 0);
                    }
                    else if (SelectObject(hdc, this->_hfMarlett))
                    {
                        v18 = this->_iTextCenterVal;
                        if (v18 <= 0)
                            v18 = 0;
                        rc.top += v18 + this->_tmAscent - this->_tmAscentMarlett;
                        v33 = (unsigned __int16)(view != OHVIEW_0 ? 'w' : '8');
                        if (v24)
                            v33 = (unsigned __int16)(view != OHVIEW_0 ? '8' : 'w');

                        if (EVAL(!IsRectEmpty(&rc))) // 450
                        {
                            SHExtTextOutW(hdc, rc.left, rc.top, v26, &rc, (const WCHAR *)&v33, 1u, 0);
                            rc.right = rc.left + this->_cxArrow;
                        }
                    }

                    BitBlt(pdis->hDC, pdis->rcItem.left, pdis->rcItem.top,RECTWIDTH(pdis->rcItem),RECTHEIGHT(pdis->rcItem), hdc, 0, 0,SRCCOPY);
                    SelectObject(hdc, h);
                    SetBkMode(hdc, mode);
                    goto LABEL_69;
                }
                v12 = GetSysColorBrush(24);
                FillRect(hdc, &pRect, v12);
                SysColor = GetSysColor(23);
            }
            SetTextColor(hdc, SysColor);
            goto LABEL_33;
        }
    }
LABEL_71:
    if (this->_pdth)
        this->_pdth->Show(1);
    return 1;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT CMorePrograms::_OnCommand(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (GET_WM_COMMAND_ID(wParam, lParam))
    {
        case IDC_ALL:
            switch (GET_WM_COMMAND_CMD(wParam, lParam))
            {
                case BN_CLICKED:
                {
                    KillTimer(_hwnd, 1);

                    OPENHOSTVIEW view;
                    if (SUCCEEDED(_GetCurView(&view)))
                    {
#if 0
                        if (wParam)
                            SHTracePerfSQMCountImpl(&ShellTraceId_Explorer_StartPane_AllPrograms_BackButton, 16);
                        else
                            SHTracePerfSQMCountImpl(&ShellTraceId_Explorer_StartPane_AllPrograms_Show_Start, 15);
#endif
                        int v12 = view == OHVIEW_0 ? 1 : 0;
                        if (SUCCEEDED(IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &SID_SM_DV2ControlHost, 302, v12, NULL, NULL)))
                        {
                            IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 326, 0, NULL, NULL);
                        }


                        LPWSTR pszTitle = v12 == 0 ? _szMessageBack : _szMessage;
                        SetWindowText(_hwndButton, pszTitle);

                        _TooltipAddTool();
                        SendMessage(_hwndTT, TTM_ACTIVATE, ShowInfoTip(), 0);
                        if (v12 == 1)
                        {
                            // SHTracePerf(&ShellTraceId_Explorer_StartPane_AllPrograms_Show_Stop
                            field_B4 = 0;
                            _SendNotify(_hwnd, SMN_SEENNEWITEMS);
                        }
                    }
                    break;
                }
            }
            break;
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnEraseBkgnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    RECT rc;
    GetClientRect(hwnd, &rc);
    if (_hTheme)
    {
        if (IsCompositionActive())
        {
            SHFillRectClr((HDC)wParam, &rc, 0);
        }
        DrawThemeBackground(_hTheme, (HDC)wParam, SPP_MOREPROGRAMS, 0, &rc, NULL);
    }
    else
    {
        SHFillRectClr((HDC)wParam, &rc, _clrBk);
    }
    return 0;
}

LRESULT CMorePrograms::_OnMouseLeave()
{
    KillTimer(this->_hwnd, 1u);
    this->field_A0 = 0;
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);

    switch (pnm->code)
    {
        case SMN_APPLYREGION:
            return HandleApplyRegion(_hwnd, _hTheme, (SMNMAPPLYREGION *)lParam, SPP_MOREPROGRAMS, 0);
        case SMN_DISMISS:
            return _OnSMNDismiss();
        case 215:
            return _OnSMNFindItem(CONTAINING_RECORD(pnm, SMNDIALOGMESSAGE, hdr));
        case SMN_SHOWNEWAPPSTIP:
            return _OnSMNShowNewAppsTip(CONTAINING_RECORD(pnm, SMNMBOOL, hdr));
        case 223:
            return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
        case 225:
            _OnMouseLeave();
            break;
        case NM_KILLFOCUS:
            field_BC = 0;
            InvalidateRect(_hwndButton, NULL, TRUE);
            break;
    }
    return 0;
}

int CMorePrograms::_Mark(SMNDIALOGMESSAGE *pdm, UINT a3)
{
    int result; // eax
    UINT flags; // esi

    pdm->flags |= 0x8000u;
    result = 1;
    flags = pdm->flags;
    if (a3 != 1)
        return 0;
    pdm->itemID = 1;
    pdm->flags = flags | 0x80000;
    pdm->field_24 = this->_hwndButton;
    return result;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
LRESULT CMorePrograms::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
#ifdef DEAD_CODE
    if (SHRestricted(REST_NOSMMOREPROGRAMS))
        return 0;

    switch (pdm->flags & SMNDM_FINDMASK)
    {

        // Life is simple if you have only one item -- all searches succeed!
        case SMNDM_FINDFIRST:
        case SMNDM_FINDLAST:
        case SMNDM_FINDNEAREST:
        case SMNDM_HITTEST:
            pdm->itemID = 0;
            return TRUE;

        case SMNDM_FINDFIRSTMATCH:
        {
            TCHAR tch = CharUpperCharW((TCHAR)pdm->pmsg->wParam);
            if (tch == _chMnem)
            {
                pdm->itemID = 0;
                return TRUE;
            }
        }
        break;      // not found

        case SMNDM_FINDNEXTMATCH:
            break;      // there is only one item so there can't be a "next"


        case SMNDM_FINDNEXTARROW:
            if (pdm->flags & SMNDM_TRYCASCADE)
            {
                FORWARD_WM_COMMAND(_hwnd, IDC_KEYPRESS, NULL, 0, PostMessage);
                return TRUE;
            }
            break;      // not found

        case SMNDM_INVOKECURRENTITEM:
        case SMNDM_OPENCASCADE:
#ifdef DEAD_CODE
            if (pdm->flags & SMNDM_KEYBOARD)
            {
                FORWARD_WM_COMMAND(_hwnd, IDC_KEYPRESS, NULL, 0, PostMessage);
            }
            else
            {
                FORWARD_WM_COMMAND(_hwnd, IDC_ALL, NULL, 0, PostMessage);
            }
            return TRUE;
#else
            if (!this->field_A0)
            {
                UINT pvParam;
                if (!SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, &pvParam, 0))
                    pvParam = 0;
                SetTimer(this->_hwnd, 1u, 2 * pvParam, 0);
            }
            return 1;
#endif

        case SMNDM_FINDITEMID:
            return TRUE;

        default:
            ASSERT(!"Unknown SMNDM command");
            break;
    }

    //
    //  If not found, then tell caller what our orientation is (vertical)
    //  and where the currently-selected item is.
    //
    pdm->flags |= SMNDM_VERTICAL;
    pdm->pt.x = 0;
    pdm->pt.y = 0;
    return FALSE;
#else
    MSG *pmsg; // eax
    WCHAR v6; // ax
    MSG *v7; // edi
    MSG *v8; // eax
    WCHAR v9; // [esp+10h] [ebp-1Ch]
    OPENHOSTVIEW view; // [esp+34h] [ebp+8h] SPLIT BYREF
    OPENHOSTVIEW view1; // [esp+34h] [ebp+8h] SPLIT BYREF
    UINT v12; // [esp+34h] [ebp+8h] SPLIT BYREF

    if (SHRestricted(REST_NOSMMOREPROGRAMS))
        return 0;

    if (!this->field_BC && (pdm->flags & 0x800) != 0)
    {
        this->field_BC = 1;
        InvalidateRect(this->_hwndButton, 0, 1);
    }

    switch (pdm->flags & 0xF)
    {
        case 0u:
            pmsg = pdm->pmsg;
            if (!pmsg)
                goto LABEL_32;
            v9 = (unsigned __int16)CharUpperW((LPWSTR)LOWORD(pmsg->wParam));
            view = OHVIEW_0;
            _GetCurView(&view);
            v6 = view ? this->_chMnemBack : this->_chMnem;
            if (v9 != v6)
            {
                goto LABEL_32;
            }
            return _Mark(pdm, 1u);
        case 1u:
            goto LABEL_32;
        case 2u:
        case 3u:
        case 4u:
        case 7u:
            pdm->itemID = 0;
            return 1;
        case 5u:
        {
            view1 = OHVIEW_0;
            _GetCurView(&view1);
            v8 = pdm->pmsg;
            if (v8 && LODWORD(v8->wParam) == 39 && view1 == OHVIEW_0 || LODWORD(v8->wParam) == 37 && view1 == OHVIEW_1)
            {
                goto LABEL_18;
            }
            goto LABEL_32;
        }
        case 6u:
        case 10u:
        {
            v7 = pdm->pmsg;
            if (v7 && v7->message == 514 && v7->hwnd == this->_hwndButton)
                ReleaseCapture();
        LABEL_18:
            PostMessageW(this->_hwnd, WM_COMMAND, 1u, (LPARAM)this->_hwndButton);
            return 1;
        }
        case 8u:
        {
            if (!this->field_A0)
            {
                if (!SystemParametersInfoW(SPI_GETMOUSEHOVERTIME, 0, &v12, 0))
                    v12 = 0;
                SetTimer(this->_hwnd, 1u, 2 * v12, 0);
            }
            return 1;
        }
        case 9u:
            return 1;
        case 11u:
            return 0;
        default:
        {
            ASSERT(!"Unknown SMNDM command"); // 748
        LABEL_32:
            pdm->flags |= 0x4000u;
            pdm->pt.x = 0;
            pdm->pt.y = 0;
            return 0;
        }
    }
#endif
}

//
//  The boolean parameter in the SMNMBOOL tells us whether to display or
//  hide the balloon tip.
//
LRESULT CMorePrograms::_OnSMNShowNewAppsTip(PSMNMBOOL psmb)
{
    if(SHRestricted(REST_NOSMMOREPROGRAMS))
        return 0;

    if (psmb->f)
    {
        if (_hwndTT)
        {
            SendMessage(_hwndTT, TTM_ACTIVATE, FALSE, 0);
        }

        if (!_hwndBalloon)
        {
            RECT rc;
            GetWindowRect(_hwndButton, &rc);

            if (!_hfTTBold)
            {
                NONCLIENTMETRICS ncm;
                ncm.cbSize = sizeof(ncm);
                if (SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
                {
                    ncm.lfStatusFont.lfWeight = FW_BOLD;
                    SHAdjustLOGFONT(&ncm.lfStatusFont);
                    _hfTTBold = CreateFontIndirect(&ncm.lfStatusFont);
                }
            }

            _hwndBalloon = CreateBalloonTip(_hwnd,
                            rc.left + _cxTextIndent + _cxText,
                            (rc.top + rc.bottom)/2,
                            _hfTTBold, 0,
                            IDS_STARTPANE_MOREPROGRAMS_BALLOONTITLE);
            if (_hwndBalloon)
            {
                SetProp(_hwndBalloon, PROP_DV2_BALLOONTIP, DV2_BALLOONTIP_MOREPROG);
            }

        }
    }
    else
    {
        _PopBalloon();
    }

    return 0;
}

void CMorePrograms::_PopBalloon()
{
    if (_hwndBalloon)
    {
        DestroyWindow(_hwndBalloon);
        _hwndBalloon = NULL;
    }
    if (_hwndTT)
    {
        SendMessage(_hwndTT, TTM_ACTIVATE, TRUE, 0);
    }

}

void CMorePrograms::_BuildHoverRect(const LPPOINT ppt)
{
    UINT pvParam;
    if (!SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0, &pvParam, 0))
        pvParam = 4;

    UINT v4;
    if (!SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0, &v4, 0))
        v4 = 4;

    this->field_64.left = ppt->x - pvParam;
    this->field_64.right = pvParam + ppt->x;
    this->field_64.top = ppt->y - v4;
    this->field_64.bottom = v4 + ppt->y;
    this->_tmHoverStart = NonzeroGetTickCount();
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnSMNDismiss()
{
    SetWindowText(_hwndButton, _szMessage);
    _TooltipAddTool();
    KillTimer(_hwnd, 1);
    field_A0 = 0;
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // update colors in classic mode
    if (!_hTheme)
    {
        _clrText = GetSysColor(COLOR_MENUTEXT);
        _clrBk = GetSysColor(COLOR_MENU);
        _hbrBk = GetSysColorBrush(COLOR_MENU);
    }

    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    return 0;
}

LRESULT CMorePrograms::_OnDisplayChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    _InitMetrics();
    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // _InitMetrics() is so cheap it's not worth getting too upset about
    // calling it too many times.
    _InitMetrics();
    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnContextMenu(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    if(SHRestricted(REST_NOSMMOREPROGRAMS))
        return 0;

    if (IS_WM_CONTEXTMENU_KEYBOARD(lParam))
    {
        RECT rc;
        GetWindowRect(_hwnd, &rc);
        lParam = MAKELPARAM(rc.left, rc.top);
    }

    c_tray._stb.OnContextMenu(_hwnd, (DWORD)lParam);
    return 0;
#else
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
#endif
}

LRESULT CMorePrograms::_OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    SetWindowPos(
        _hwndButton,
        NULL,
        _margins.cxLeftWidth,
        _margins.cyTopHeight,
        GET_X_LPARAM(lParam) - (_margins.cxLeftWidth + _margins.cxRightWidth),
        GET_Y_LPARAM(lParam) - (_margins.cyTopHeight + _margins.cyBottomHeight),
        SWP_NOZORDER | SWP_NOOWNERZORDER);
    return 0;
}

LRESULT CMorePrograms::_OnTimer(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (wParam != 1 || field_A0)
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    SendMessage(_hwnd, 0x111u, 1u, (LPARAM)_hwndButton);
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT CALLBACK CMorePrograms::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CMorePrograms *self = reinterpret_cast<CMorePrograms *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

    switch (uMsg)
    {
        case WM_NCCREATE:
            return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);
        case WM_CREATE:
            return self->_OnCreate(hwnd, uMsg, wParam, lParam);
        case WM_DESTROY:
            return self->_OnDestroy(hwnd, uMsg, wParam, lParam);
        case WM_NCDESTROY:
            return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);
        case WM_CTLCOLORBTN:
            return self->_OnCtlColorBtn(hwnd, uMsg, wParam, lParam);
        case WM_DRAWITEM:
            return self->_OnDrawItem(hwnd, uMsg, wParam, lParam);
        case WM_ERASEBKGND:
            return self->_OnEraseBkgnd(hwnd, uMsg, wParam, lParam);
        case WM_COMMAND:
            return self->_OnCommand(hwnd, uMsg, wParam, lParam);
        case WM_SYSCOLORCHANGE:
            return self->_OnSysColorChange(hwnd, uMsg, wParam, lParam);
        case WM_DISPLAYCHANGE:
            return self->_OnDisplayChange(hwnd, uMsg, wParam, lParam);
        case WM_SETTINGCHANGE:
            return self->_OnSettingChange(hwnd, uMsg, wParam, lParam);
        case WM_NOTIFY:
            return self->_OnNotify(hwnd, uMsg, wParam, lParam);
        case WM_CONTEXTMENU:
            return self->_OnContextMenu(hwnd, uMsg, wParam, lParam);
        case WM_SIZE:
            return self->_OnSize(hwnd, uMsg, wParam, lParam);
		case WM_TIMER:
			return self->_OnTimer(hwnd, uMsg, wParam, lParam);
    }

    return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
}

BOOL WINAPI MorePrograms_RegisterClass()
{
    WNDCLASSEX wc;
    ZeroMemory(&wc, sizeof(wc));
    
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_GLOBALCLASS;
    wc.lpfnWndProc   = CMorePrograms::s_WndProc;
    wc.hInstance     = _AtlBaseModule.GetModuleInstance();
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszClassName = WC_MOREPROGRAMS;

    return RegisterClassEx(&wc);
}

// We implement a minimal drop target so we can auto-open the More Programs
// list when the user hovers over the More Programs button.

// *** IUnknown ***

HRESULT CMorePrograms::QueryInterface(REFIID riid, void * *ppvOut)
{
    static const QITAB qit[] = {
        QITABENT(CMorePrograms, IDropTarget),
        QITABENT(CMorePrograms, IServiceProvider),
        QITABENT(CMorePrograms, IOleCommandTarget),
        QITABENT(CMorePrograms, IObjectWithSite),
        QITABENT(CMorePrograms, IAccessible),
        QITABENT(CMorePrograms, IDispatch), // IAccessible derives from IDispatch
        QITABENT(CMorePrograms, IEnumVARIANT),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvOut);
}

ULONG CMorePrograms::AddRef()
{
    return InterlockedIncrement(&_lRef);
}

ULONG CMorePrograms::Release()
{
    ASSERT( 0 != _lRef );
    ULONG cRef = InterlockedDecrement(&_lRef);
    if ( 0 == cRef) 
    {
        delete this;
    }
    return cRef;
}


// *** IDropTarget::DragEnter ***
// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::DragEnter(IDataObject *pdto, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
    POINT pt = { ptl.x, ptl.y };
    if (_pdth) {
        _pdth->DragEnter(_hwnd, pdto, &pt, *pdwEffect);
    }

	_BuildHoverRect(&pt);
    InvalidateRect(_hwndButton, NULL, TRUE); // draw with drop highlight

    return DragOver(grfKeyState, ptl, pdwEffect);
}

// *** IDropTarget::DragOver ***
// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::DragOver(DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
#ifdef DEAD_CODE
    POINT pt = { ptl.x, ptl.y };
    if (_pdth) {
        _pdth->DragOver(&pt, *pdwEffect);
    }

    //  Hover time is 1 second, the same as the hard-coded value for the
    //  Start Button.
    if (_tmHoverStart && GetTickCount() - _tmHoverStart > 1000)
    {
        _tmHoverStart = 0;
        FORWARD_WM_COMMAND(_hwnd, IDC_ALL, _hwndButton, BN_CLICKED, PostMessage);
    }

    *pdwEffect = DROPEFFECT_NONE;
    return S_OK;
#else
    POINT pt = { ptl.x, ptl.y };
    if (_pdth)
        _pdth->DragOver(&pt, *pdwEffect);

    if (_tmHoverStart)
    {
        if (PtInRect(&field_64, pt))
        {
            if (GetTickCount() - _tmHoverStart > 1000)
            {
                _tmHoverStart = 0;
                PostMessage(_hwnd, WM_COMMAND, 1u, (LPARAM)_hwndButton);
            }
        }
        else
        {
            _BuildHoverRect(&pt);
        }
    }

    *pdwEffect = DROPEFFECT_NONE;
    return S_OK;
#endif
}

// *** IDropTarget::DragLeave ***
// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::DragLeave()
{
    if (_pdth)
    {
        _pdth->DragLeave();
    }

    _tmHoverStart = 0;
    InvalidateRect(_hwndButton, NULL, TRUE); // draw without drop highlight

    return S_OK;
}

// *** IDropTarget::Drop ***
// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::Drop(IDataObject *pdto, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
    POINT pt = { ptl.x, ptl.y };
    if (_pdth)
    {
        _pdth->Drop(pdto, &pt, *pdwEffect);
    }

    _tmHoverStart = 0;
    InvalidateRect(_hwndButton, NULL, TRUE); // draw without drop highlight

    return S_OK;
}

//****************************************************************************
//
//  Accessibility
//

//
//  The default accessibility object reports buttons as
//  ROLE_SYSTEM_PUSHBUTTON, but we know that we are really a menu.
//
// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
    HRESULT hr = CAccessible::get_accRole(varChild, pvarRole);
    if (SUCCEEDED(hr) && V_VT(pvarRole) == VT_I4)
    {
        switch (V_I4(pvarRole))
        {
        case ROLE_SYSTEM_PUSHBUTTON:
            V_I4(pvarRole) = ROLE_SYSTEM_MENUITEM;
            break;
        }
    }
    return hr;
}

// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::get_accState(VARIANT varChild, VARIANT *pvarState)
{
    HRESULT hr = CAccessible::get_accState(varChild, pvarState);
    if (SUCCEEDED(hr) && V_VT(pvarState) == VT_I4)
    {
        V_I4(pvarState) |= STATE_SYSTEM_HASPOPUP;
    }
    return hr;
}

HRESULT CMorePrograms::get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut)
{
    return CreateAcceleratorBSTR(_chMnem, pszKeyboardShortcut);
}

HRESULT CMorePrograms::get_accDefaultAction(VARIANT varChild, BSTR *pszDefAction)
{
    DWORD dwRole = _fMenuOpen ? ACCSTR_CLOSE : ACCSTR_OPEN;
    return GetRoleString(dwRole, pszDefAction);
}

HRESULT CMorePrograms::accDoDefaultAction(VARIANT varChild)
{
    if (_fMenuOpen)
    {
        _SendNotify(_hwnd, SMN_CANCELSHELLMENU);
        return S_OK;
    }
    else
    {
        return CAccessible::accDoDefaultAction(varChild);
    }
}

// ****************************************************************************
// IServiceProvider
//

HRESULT CMorePrograms::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    if (IsEqualGUID(guidService, SID_SM_ViewControl))
    {
        HRESULT hr = QueryInterface(riid, ppvObject);
        if (SUCCEEDED(hr))
        {
            return hr;
        }
    }
    return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

// ****************************************************************************
// IOleCommandTarget
//

HRESULT CMorePrograms::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
	return E_NOTIMPL;
}

HRESULT CMorePrograms::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
    HRESULT hr = E_INVALIDARG;
    if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup) && nCmdID == 302 && nCmdexecopt != -1)
    {
        _OnSetCurView((OPENHOSTVIEW)nCmdexecopt);
        return S_OK;
    }
    return hr;
}