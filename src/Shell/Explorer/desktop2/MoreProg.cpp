#include "pch.h"
#include "cocreateinstancehook.h"
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
CMorePrograms::CMorePrograms(HWND hwnd)
    : _hwnd(hwnd),
    _clrText(CLR_INVALID),
    _clrBk(CLR_INVALID),
    _lRef(1)
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

    IUnknown_SafeReleaseAndNullPtr(_pdth);

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
    SMPANEDATA* psmpd = PaneDataFromCreateStruct(lParam);

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
        if (!LoadString(_AtlBaseModule.GetResourceInstance(), 8226, _szMessage, ARRAYSIZE(_szMessage))
            || !LoadString(_AtlBaseModule.GetResourceInstance(), 8241, _szMessageBack, ARRAYSIZE(_szMessageBack))
            || !LoadString(_AtlBaseModule.GetResourceInstance(), 8227, _szTool, ARRAYSIZE(_szTool))
            || !LoadString(_AtlBaseModule.GetResourceInstance(), 8245, _szToolBack, ARRAYSIZE(_szToolBack)))
        {
            return -1;
        }

        _chMnem     = CharUpperCharW(SHFindMnemonic(_szMessage));
        _chMnemBack = CharUpperCharW(SHFindMnemonic(_szMessageBack));

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
                _cxTextBack = sizText.cx + SHGetSystemMetricsScaled(SM_CXEDGE);

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
        rc.left     += _margins.cxLeftWidth;
        rc.right    -= _margins.cxRightWidth;
        rc.top      += _margins.cyTopHeight;
        rc.bottom   -= _margins.cyBottomHeight;

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
            _AtlBaseModule.GetModuleInstance(),
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
        ti.hinst = _AtlBaseModule.GetResourceInstance();
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
    return SHFusionCreateWindowEx(
        0, TOOLTIPS_CLASS, nullptr, TTS_ALWAYSTIP | TTS_NOPREFIX | WS_BORDER, 0, 0, 0, 0, _hwndButton, nullptr,
        _AtlBaseModule.GetModuleInstance(), nullptr);
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

// EXEX-VISTA(allison): Validated.
HRESULT CMorePrograms::_GetCurView(OPENHOSTVIEW *pView)
{   
    VARIANT var;
	var.vt = VT_I4;
    HRESULT hr = IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &CGID_DV2ControlHost, 303, 0, 0, &var);
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

    if (_pdth)
        _pdth->Show(0);

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
            h = SelectObject(hdc, _hf);
            if (!h)
            {
            LABEL_69:
                SelectObject(hdc, v23);
                DeleteObject(ho);
                goto LABEL_70;
            }
            if ((pdis->itemState & 0x50) != 0 || _tmHoverStart || (v8 = field_BC == 0, v33 = 0, !v8))
                v33 = 1;

            view = OHVIEW_0;
            _GetCurView(&view);

            iStateId = 1;
            hTheme = _hTheme;
            if (hTheme)
            {
                if (v33)
                {
                    iStateId = 2;
                }
                else if (field_B4)
                {
                    iStateId = 3;
                }
                DrawThemeBackground(hTheme, hdc, 12, iStateId, &pRect, 0);
                goto LABEL_33;
            }
            if (v33)
            {
                SysColorBrush = GetSysColorBrush(_colorHighlight);
                FillRect(hdc, &pRect, SysColorBrush);
                SysColor = GetSysColor(_colorHighlightText);
            }
            else
            {
                if (!field_B4)
                {
                    FillRect(hdc, &pRect, _hbrBk);
                    SetTextColor(hdc, _clrText);
                LABEL_33:
                    mode = SetBkMode(hdc, 1);
                    ASSERT(pdis->CtlID == IDC_ALL); // 391
                    if (view)
                        cxText2 = _cxTextBack;
                    else
                        cxText2 = _cxText;
                    dwTextFlags = cxText2;
                    cxTextIndent = _cxTextIndent;
                    if (cxTextIndent > (int)(pRect.right - cxText2 - pRect.left))
                    {
                        //CcshellDebugMsgW(
                        //    1,
                        //    "StartMenu: 'maximum of (%s, %s) ' is %dpx, only room for %d- notify localizers!",
                        //    (const char *)_szMessage,
                        //    (const char *)_szMessageBack,
                        //    cxText2,
                        //    pRect.right - cxText2 - pRect.left);
                        cxTextIndent = pRect.right - dwTextFlags - pRect.left;
                        if (cxTextIndent < 0)
                            cxTextIndent = 0;
                    }
                    pRect.left += cxTextIndent;
                    dwTextFlags = 0x2024;
					ASSERT(pdis->CtlID == IDC_ALL); // 403
                    szMessage = _szMessage;
                    if (view)
                        szMessage = _szMessageBack;

                    if (v24)
                        dwTextFlags |= 0x20000u;
                    if ((pdis->itemState & 0x100) != 0)
                        dwTextFlags |= 0x100000u;

                    if (_hTheme)
                    {
                        pOptions.dwSize = 64;
                        memset(&pOptions.dwFlags, 0, 0x3Cu);
                        pOptions.dwFlags = IsCompositionActive() ? 0x2000 : 0;
                        DrawThemeTextEx(_hTheme, hdc, 12, iStateId, szMessage, -1, dwTextFlags, &pRect, &pOptions);
                    }
                    else
                    {
                        DrawTextW(hdc, szMessage, -1, &pRect, dwTextFlags);
                    }
                    rc = pRect;
                    rc.left = GetSystemMetrics(45);
                    v16 = _hTheme;
                    if (v16)
                    {
                        iTextCenterVal = _iTextCenterVal;
                        if (iTextCenterVal < 0)
                            rc.top -= iTextCenterVal;
                        rc.right = rc.left + _cxArrow;
                        DrawThemeBackground(v16, hdc, view != OHVIEW_0 ? 17 : 3, v33 != 0 ? 2 : 0, &rc, 0);
                    }
                    else if (SelectObject(hdc, _hfMarlett))
                    {
                        v18 = _iTextCenterVal;
                        if (v18 <= 0)
                            v18 = 0;
                        rc.top += v18 + _tmAscent - _tmAscentMarlett;
                        v33 = (unsigned __int16)(view != OHVIEW_0 ? 'w' : '8');
                        if (v24)
                            v33 = (unsigned __int16)(view != OHVIEW_0 ? '8' : 'w');

                        if (EVAL(!IsRectEmpty(&rc))) // 450
                        {
                            SHExtTextOutW(hdc, rc.left, rc.top, v26, &rc, (const WCHAR *)&v33, 1u, 0);
                            rc.right = rc.left + _cxArrow;
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
    if (_pdth)
        _pdth->Show(1);
    return 1;
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
                        BOOL v12 = view == OHVIEW_0 ? 1 : 0;
                        if (SUCCEEDED(IUnknown_QueryServiceExec(_punkSite, SID_SM_OpenHost, &CGID_DV2ControlHost, 302, v12, NULL, NULL)))
                        {
                            IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &CGID_DV2ControlHost, 326, 0, NULL, NULL);
                        }


                        LPWSTR pszTitle = v12 != 0 ? _szMessage : _szMessageBack;
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
    KillTimer(_hwnd, 1);
    field_A0 = 0;
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
        case 217:
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
        default:
            break;
    }
    return 0;
}

int CMorePrograms::_Mark(SMNDIALOGMESSAGE* pdm, UINT a3)
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
    pdm->hwnd2 = _hwndButton;
    return result;
}

// EXEX-VISTA(allison): Validated.
LRESULT CMorePrograms::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
    if (SHRestricted(REST_NOSMMOREPROGRAMS) != 0)
        return 0;

    if (!field_BC && (pdm->flags & 0x800) != 0)
    {
        field_BC = 1;
        InvalidateRect(_hwndButton, nullptr, true);
    }

    switch (pdm->flags & SMNDM_FINDMASK)
    {
        case SMNDM_FINDFIRSTMATCH:
        {
            if (pdm->pmsg != nullptr)
            {
                WCHAR tch = CharUpperCharW((WCHAR)pdm->pmsg->wParam);

                OPENHOSTVIEW view = OHVIEW_0;
                _GetCurView(&view);
                WCHAR chMnem = view == OHVIEW_0 ? _chMnem : _chMnemBack;
                if (tch == chMnem)
                {
                    return _Mark(pdm, 1u);
                }
            }
            break;
        }

        case SMNDM_FINDNEXTMATCH:
            break;

        case SMNDM_FINDNEAREST:
        case SMNDM_FINDFIRST:
        case SMNDM_FINDLAST:
        case SMNDM_HITTEST:
            pdm->itemID = 0;
            return 1;

        case SMNDM_FINDNEXTARROW:
        {
            OPENHOSTVIEW view = OHVIEW_0;
            _GetCurView(&view);

            if (pdm->pmsg != nullptr && pdm->pmsg->wParam == 39 && view == OHVIEW_0 || pdm->pmsg->wParam == 37 && view == OHVIEW_1)
            {
                PostMessage(_hwnd, WM_COMMAND, 1, (LPARAM)_hwndButton);
                return 1;
            }
            break;
        }

        case SMNDM_INVOKECURRENTITEM:
        case SMNDM_MOUSEDOWN:
            if (pdm->pmsg != nullptr && pdm->pmsg->message == WM_LBUTTONUP && pdm->pmsg->hwnd == _hwndButton)
            {
                ReleaseCapture();
            }
            PostMessage(_hwnd, WM_COMMAND, 1, (LPARAM)_hwndButton);
            return 1;

        case 8:
        {
            if (!field_A0)
            {
                UINT uHoverTime;
                if (!SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, &uHoverTime, 0))
                    uHoverTime = 0;
                SetTimer(_hwnd, 1, uHoverTime * 2, nullptr);
            }
            return 1;
        }

        case SMNDM_FINDITEMID:
            return 1;

        case 11:
            return 0;

        default:
            ASSERT(!"Unknown SMNDM command"); // 748
            break;
    }

    pdm->flags |= SMNDM_VERTICAL;
    pdm->pt.x = 0;
    pdm->pt.y = 0;
    return FALSE;
}

//
//  The boolean parameter in the SMNMBOOL tells us whether to display or
//  hide the balloon tip.
//
LRESULT CMorePrograms::_OnSMNShowNewAppsTip(PSMNMBOOL psmb)
{
    if (!SHRestricted(REST_NOSMMOREPROGRAMS))
    {
        field_B4 = psmb->f;
        InvalidateRect(_hwndButton, NULL, TRUE);
    }
    return 0;
}

void CMorePrograms::_BuildHoverRect(const LPPOINT ppt)
{
    UINT cxHover;
    if (!SystemParametersInfo(SPI_GETMOUSEHOVERWIDTH, 0, &cxHover, 0))
        cxHover = 4;

    UINT cyHover;
    if (!SystemParametersInfo(SPI_GETMOUSEHOVERHEIGHT, 0, &cyHover, 0))
        cyHover = 4;

    _rcHover.left       = ppt->x - cxHover;
    _rcHover.right      = ppt->x + cxHover;

    _rcHover.top        = ppt->y - cyHover;
    _rcHover.bottom     = ppt->y + cyHover;
    _tmHoverStart = NonzeroGetTickCount();
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
    if (!SHRestricted(REST_NOSMMOREPROGRAMS))
    {
        if (IS_WM_CONTEXTMENU_KEYBOARD(lParam))
        {
            RECT rc;
            GetWindowRect(_hwnd, &rc);
			lParam = MAKELPARAM(rc.left, rc.top);
        }

        SMNMISTARTBUTTON isb;
        isb.psb = NULL;
        _SendNotify(_hwnd, 218, &isb.hdr);
        if (isb.psb)
        {
            isb.psb->OnContextMenu(_hwnd, lParam);
			isb.psb->Release();
        }
    }
    return 0;
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
    if (wParam == 1 && !field_A0)
    {
        SendMessage(_hwnd, WM_COMMAND, 1u, (LPARAM)_hwndButton);
        return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
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
    POINT pt = { ptl.x, ptl.y };
    if (_pdth)
        _pdth->DragOver(&pt, *pdwEffect);

    if (_tmHoverStart)
    {
        if (PtInRect(&_rcHover, pt))
        {
            if (GetTickCount() - _tmHoverStart > 1000)
            {
                _tmHoverStart = 0;
                PostMessage(_hwnd, WM_COMMAND, IDC_ALL, (LPARAM)_hwndButton);
            }
        }
        else
        {
            _BuildHoverRect(&pt);
        }
    }

    *pdwEffect = DROPEFFECT_NONE;
    return S_OK;
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
    return GetRoleString(ACCSTR_OPEN, pszDefAction);
}

HRESULT CMorePrograms::accDoDefaultAction(VARIANT varChild)
{
    return CAccessible::accDoDefaultAction(varChild);
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

HRESULT CMorePrograms::Exec(const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvarargIn, VARIANT* pvarargOut)
{
    HRESULT hr = E_INVALIDARG;
    if (IsEqualGUID(CGID_DV2ControlHost, *pguidCmdGroup) && nCmdID == 302 && nCmdexecopt != -1)
    {
        _OnSetCurView((OPENHOSTVIEW)nCmdexecopt);
        hr = S_OK;
    }
    return hr;
}
