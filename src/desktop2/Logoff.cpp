#include "pch.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "shundoc.h"

// WARNING!  Must be in sync with c_rgidmLegacy

#define NUM_TBBUTTON_IMAGES 6
static const TBBUTTON tbButtonsCreate [] = 
{
    {0, SMNLC_EJECT,    TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_EJECT, 0},
    {1, SMNLC_LOGOFF,   TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_LOGOFF, 1},
    {2, SMNLC_TURNOFF,  TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_SHUTDOWN, 2},
    {2,SMNLC_DISCONNECT,TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_DISCONNECT, 3},
};

// WARNING!  Must be in sync with tbButtonsCreate
static const UINT c_rgidmLegacy[] =
{
    IDM_EJECTPC,
    IDM_LOGOFF,
    IDM_EXITWIN,
    IDM_MU_DISCONNECT,
};

// EXEX-TODO: Move? (This one is still in use by Windows 10.)
// {14CE31DC-ABC2-484C-B061-CF3416AED8FF}
DEFINE_GUID(CLSID_AuthUIShutdownChoices, 0x14CE31DC, 0xABC2, 0x484C, 0xB0, 0x61, 0xCF, 0x34, 0x16, 0xAE, 0xD8, 0xFF);

// EXEX-TODO: Move? Copied from https://github.com/AllieTheFox/AuthUX-Styles/blob/e05fda9e431ce5715e2e05093febc45cf8146d5c/sdk/inc/logoninterfaces.h#L1044-L1051
MIDL_INTERFACE("a0b16477-52f1-4cfe-b1cc-388cd0e3e23a")
IEnumShutdownChoices : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Next(DWORD, DWORD*, DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE Skip(DWORD) = 0;
    virtual HRESULT STDMETHODCALLTYPE Reset() = 0;
    virtual HRESULT STDMETHODCALLTYPE Clone(IEnumShutdownChoices**) = 0;
};

// EXEX-TODO: Move? Vista implementation.
MIDL_INTERFACE("79D6926F-5F85-4F37-B3AE-427A9B015CB2")
IShutdownChoiceListener : IUnknown
{
    STDMETHOD(SetNotifyWnd)(HWND hwnd, UINT a3) PURE;
    STDMETHOD(GetMessageWnd)(HWND *phwnd) PURE;
    STDMETHOD(ScanForPassiveChange)() PURE;
    STDMETHOD(StartListening)() PURE;
    STDMETHOD(StopListening)() PURE;
}

// EXEX-TODO: Move? Also, this interface changed between Windows versions. This is the Vista version.
typedef SHUTDOWN_CHOICE; // Temporary type to represent while I figure out the real type.
MIDL_INTERFACE("A5DBD3DC-EE32-497A-AB84-F2C6AA5913F5")
IShutdownChoices : IUnknown
{
    STDMETHOD(Refresh)() PURE;
    STDMETHOD(CreateListener)(IShutdownChoiceListener **ppOut) PURE;
    STDMETHOD(SetShowBadChoices)(BOOL fShow) PURE;
    STDMETHOD(GetChoiceEnumerator)(IEnumShutdownChoices **ppOut) PURE;
    STDMETHOD(GetDefaultChoice)(int **piOut) PURE; // This returns an ID representing the choice.
    STDMETHOD(GetMenuChoices)(IEnumShutdownChoices **ppOut) PURE;
    STDMETHOD_(BOOL, UserHasShutdownRights)() PURE;
    STDMETHOD(GetChoiceName)(SHUTDOWN_CHOICE choice, int a3, LPWSTR a4, UINT a5) PURE;
    STDMETHOD(GetChoiceDesc)(SHUTDOWN_CHOICE, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceVerb)(SHUTDOWN_CHOICE, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceIcon)(SHUTDOWN_CHOICE choice, enum SHUTDOWN_CHOICE_ICON *pIcon) PURE;
};

class CLogoffPane
    : public CUnknown
    , public CAccessible
{
public:

    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj);
    STDMETHODIMP_(ULONG) AddRef(void) { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release(void) { return CUnknown::Release(); }

    // *** IAccessible overridden methods ***
    STDMETHODIMP get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut);
    STDMETHODIMP get_accDefaultAction(VARIANT varChild, BSTR *pszDefAction);

    CLogoffPane();
    ~CLogoffPane();

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCreate(LPARAM lParam);
    void _OnDestroy();
    LRESULT _OnNCCreate(HWND hwnd,  UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNotify(NMHDR *pnm);
    LRESULT _OnCommand(int id);
    LRESULT _OnCustomDraw(NMTBCUSTOMDRAW *pnmcd);
    LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);
    LRESULT _OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm);
    LRESULT _OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnDisplayChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    void    _InitMetrics();
    LRESULT _OnSize(int x, int y);

private:
    HWND _hwnd;
    HWND _hwndTB;
    HWND _hwndSplit;
    HWND _hwndTT;   //Tooltip window.
    COLORREF _clr;
    int      _colorHighlight;
    int      _colorHighlightText;
    HTHEME _hTheme;
    BOOL   _fSettingHotItem;
    int field_40; // EXEX-TODO(isabella): Rename.
    MARGINS _margins;
    HFONT _hfMarlett;
    int field_58;
    int field_5c;
    HIMAGELIST _himl;
    int field_64;
    int field_68;
    IShutdownChoices *_psdc;
    IShutdownChoiceListener *_psdListen;
    HWND _hwndSdListenMsg;
    int field_78;


    // helper functions
    int _GetCurButton();
    LRESULT _NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction);
    BOOL _IsButtonHidden(int i);
    TCHAR _GetButtonAccelerator(int i);
    void _RightAlign();
    void _ApplyOptions();
    HRESULT _InitShutdownObjects();

    BOOL _SetTBButtons(int id, UINT iMsg);

    friend BOOL CLogoffPane_RegisterClass();
};

CLogoffPane::CLogoffPane()
{
    //ASSERT(_hwndTB == NULL);
    //ASSERT(_hwndTT == NULL);
    _clr = CLR_INVALID;
}

CLogoffPane::~CLogoffPane()
{
}

HRESULT CLogoffPane::QueryInterface(REFIID riid, void * *ppvOut)
{
    static const QITAB qit[] = {
        QITABENT(CLogoffPane, IAccessible),
        QITABENT(CLogoffPane, IDispatch), // IAccessible derives from IDispatch
        QITABENT(CLogoffPane, IEnumVARIANT),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvOut);
}

LRESULT CALLBACK CLogoffPane::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CLogoffPane *self = reinterpret_cast<CLogoffPane *>(GetWindowLongPtr(hwnd, 0));

    switch (uMsg)
    {
        case WM_NCCREATE:
            return self->_OnNCCreate(hwnd, uMsg, wParam, lParam);

        case WM_CREATE:
            return self->_OnCreate(lParam);

        case WM_DESTROY:
            self->_OnDestroy();
            break;

        case WM_NCDESTROY:
            return self->_OnNCDestroy(hwnd, uMsg, wParam, lParam);

        case WM_ERASEBKGND:
        {
            RECT rc;
            GetClientRect(hwnd, &rc);
            if (self->_hTheme)
            {
                DrawThemeBackground(self->_hTheme, (HDC)wParam, SPP_LOGOFF, 0, &rc, 0);
            }
            else
            {
                SHFillRectClr((HDC)wParam, &rc, GetSysColor(COLOR_MENU));
                DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_TOP);
            }
    
            return 1;
        }

        case WM_COMMAND:
            return self->_OnCommand(GET_WM_COMMAND_ID(wParam, lParam));

        case WM_NOTIFY:
            return self->_OnNotify((NMHDR*)(lParam));

        case WM_SIZE:
            return self->_OnSize(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

        case WM_SYSCOLORCHANGE:
            return self->_OnSysColorChange(hwnd, uMsg, wParam, lParam);
            
        case WM_DISPLAYCHANGE:
            return self->_OnDisplayChange(hwnd, uMsg, wParam, lParam);
            
        case WM_SETTINGCHANGE:
            return self->_OnSettingChange(hwnd, uMsg, wParam, lParam);
    }
    return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CLogoffPane::_OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);

    CLogoffPane *self = new CLogoffPane;

    if (self)
    {
        SetWindowLongPtr(hwnd, 0, (LONG_PTR)self);

        self->_hwnd = hwnd;

        return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return FALSE;
}

// EXEX-VISTA: Validated.
void AddBitmapToToolbar(HWND hwndTB, HBITMAP hBitmap, int cxTotal, int cy, int iNumImages, UINT iMsg)
{
    HIMAGELIST himl = ImageList_Create(cxTotal / iNumImages, cy, ILC_COLOR32, 0, iNumImages);

    if (himl)
    {
        ImageList_Add(himl, hBitmap, NULL);

        HIMAGELIST himlPrevious = (HIMAGELIST) SendMessage(hwndTB, iMsg, 0, (LPARAM)himl);
        if (himlPrevious)
        {
            ImageList_Destroy(himlPrevious);
        }
    }
}

// EXEX-VISTA: Validated.
BOOL CLogoffPane::_SetTBButtons(int id, UINT iMsg)
{
    HBITMAP hBitmap = LoadBitmap(_AtlBaseModule.GetModuleInstance(), MAKEINTRESOURCE(id));
    if (hBitmap)
    {
        BITMAP bm;
        if (GetObject(hBitmap, sizeof(BITMAP), &bm))
        {
            AddBitmapToToolbar(_hwndTB, hBitmap, bm.bmWidth, bm.bmHeight, NUM_TBBUTTON_IMAGES, iMsg);
        }
        DeleteObject(hBitmap);
    }
    return BOOLIFY(hBitmap);
}

// EXEX-VISTA: Validated.
void CLogoffPane::_OnDestroy()
{
    if (IsWindow(_hwndTB))
    {
        HIMAGELIST himl = (HIMAGELIST) SendMessage(_hwndTB, TB_GETIMAGELIST, 0, 0);

        if (himl)
        {
            ImageList_Destroy(himl);
        }

        himl = (HIMAGELIST) SendMessage(_hwndTB, TB_GETHOTIMAGELIST, 0, 0);
        if (himl)
        {
            ImageList_Destroy(himl);
        }

        himl = (HIMAGELIST)SendMessage(_hwndTB, TB_GETPRESSEDIMAGELIST, 0, 0);
        if (himl)
        {
            ImageList_Destroy(himl);
        }
    }
}

// EXEX-VISTA: Partially reversed.
LRESULT CLogoffPane::_OnCreate(LPARAM lParam)
{
    _InitShutdownObjects();

    // Do not set WS_TABSTOP here; that's CLogoffPane's job

    DWORD dwStyle = WS_CHILD|WS_CLIPSIBLINGS|WS_VISIBLE | CCS_NORESIZE|CCS_NODIVIDER | TBSTYLE_FLAT|TBSTYLE_LIST|TBSTYLE_TOOLTIPS;
    RECT rc;

    _hTheme = (PaneDataFromCreateStruct(lParam))->hTheme;

    if (_hTheme)
    {
        GetThemeColor(_hTheme, SPP_LOGOFF, 0, TMT_TEXTCOLOR, &_clr);
        _colorHighlight = COLOR_MENUHILIGHT;
        _colorHighlightText = COLOR_HIGHLIGHTTEXT;

        GetThemeMargins(_hTheme, NULL, SPP_LOGOFF, 0, TMT_CONTENTMARGINS, NULL, &_margins);
    }
    else
    {
        _clr = GetSysColor(COLOR_MENUTEXT);
        _colorHighlight = COLOR_HIGHLIGHT;
        _colorHighlightText = COLOR_HIGHLIGHTTEXT;

        _margins.cyTopHeight = _margins.cyBottomHeight = 2 * GetSystemMetrics(SM_CYEDGE);
        ASSERT(_margins.cxLeftWidth == 0);
        ASSERT(_margins.cxRightWidth == 0);
    }

    GetClientRect(_hwnd, &rc);
    rc.left += _margins.cxLeftWidth;
    rc.right -= _margins.cxRightWidth;
    rc.top += _margins.cyTopHeight;
    rc.bottom -= _margins.cyBottomHeight;

    _hwndTB = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, dwStyle,
                               rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc), _hwnd, 
                               NULL, NULL, NULL );
    if (_hwndTB)
    {
        //
        //  Don't freak out if this fails.  It just means that the accessibility
        //  stuff won't be perfect.
        //
        SetAccessibleSubclassWindow(_hwndTB);

        // we do our own themed drawing...
        SetWindowTheme(_hwndTB, L"", L"");

        // Scale up on HIDPI
        SendMessage(_hwndTB, CCM_DPISCALE, TRUE, 0);

        SendMessage(_hwndTB, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

        if (!_hTheme ||
            !_ThemedSetTBButtons(SPLS_NORMAL, TB_SETIMAGELIST) ||
            !_ThemedSetTBButtons(SPLS_HOT, TB_SETHOTIMAGELIST) )
        {
            // if we don't have a theme, or failed at setting the images from the theme
            // set buttons images from the rc file
            _SetTBButtons(IDB_LOGOFF_NORMAL, TB_SETIMAGELIST);
            _SetTBButtons(IDB_LOGOFF_PRESSED, TB_SETHOTIMAGELIST);
        }
        
        SendMessage(_hwndTB, TB_ADDBUTTONS, ARRAYSIZE(tbButtonsCreate), (LPARAM) tbButtonsCreate);
        BOOL fGina = SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("GinaUI"), FALSE, FALSE);
        int idText = !fGina ? IDS_LOGOFF_TEXT_FRIENDLY : IDS_LOGOFF_TEXT_DOMAIN;
        SendMessage(_hwndTB, TB_ADDSTRING, (WPARAM) _AtlBaseModule.GetModuleInstance(), (LPARAM) idText);

        _ApplyOptions();

        _hwndTT = (HWND)SendMessage(_hwndTB, TB_GETTOOLTIPS, 0, 0); //Get the tooltip window.

        _InitMetrics();
        
        return 0;
    }

    return -1; // no point in sticking around if we couldn't create the toolbar
}

BOOL CLogoffPane::_IsButtonHidden(int i)
{
    TBBUTTON but;
    SendMessage(_hwndTB, TB_GETBUTTON, i, (LPARAM) &but);
    return but.fsState & TBSTATE_HIDDEN;
}

void CLogoffPane::_RightAlign()
{
    int iWidthOfButtons=0;

    // add up the width of all the non-hidden buttons
    for(int i=0;i<ARRAYSIZE(tbButtonsCreate);i++)
    {
        if (!_IsButtonHidden(i))
        {
            RECT rc;
            SendMessage(_hwndTB, TB_GETITEMRECT, i, (LPARAM) &rc);
            iWidthOfButtons += RECTWIDTH(rc);
        }
    }

    if (iWidthOfButtons)
    {
        RECT rc;
        GetClientRect(_hwndTB, &rc);

        int iIndent = RECTWIDTH(rc) - iWidthOfButtons - GetSystemMetrics(SM_CXEDGE);

        if (iIndent < 0)
            iIndent = 0;
        SendMessage(_hwndTB, TB_SETINDENT, iIndent, 0);
    }
}

LRESULT CLogoffPane::_OnSize(int x, int y)
{
    if (_hwndTB)
    {
        SetWindowPos(_hwndTB, NULL, _margins.cxLeftWidth, _margins.cyTopHeight, 
            x-(_margins.cxRightWidth+_margins.cxLeftWidth), y-(_margins.cyBottomHeight+_margins.cyTopHeight), 
            SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        _RightAlign();
    }
    return 0;
}

LRESULT CLogoffPane::_OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // WARNING!  "this" might be invalid (if WM_NCCREATE failed), so
    // do not use any member variables!
    LRESULT lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
    SetWindowLongPtr(hwnd, 0, 0);
    if (this)
    {
        this->Release();
    }
    return lres;
}

void CLogoffPane::_ApplyOptions()
{
    SMNFILTEROPTIONS nmopt;
    nmopt.smnop = SMNOP_LOGOFF | SMNOP_TURNOFF | SMNOP_DISCONNECT | SMNOP_EJECT;
    _SendNotify(_hwnd, SMN_FILTEROPTIONS, &nmopt.hdr);

    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_EJECT, !(nmopt.smnop & SMNOP_EJECT));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_LOGOFF, !(nmopt.smnop & SMNOP_LOGOFF));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, !(nmopt.smnop & SMNOP_TURNOFF));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, !(nmopt.smnop & SMNOP_DISCONNECT));

    _RightAlign();
}


LRESULT CLogoffPane::_OnNotify(NMHDR *pnm)
{
    if (pnm->hwndFrom == _hwndTB)
    {
        switch (pnm->code)
        {
            case NM_CUSTOMDRAW:
                return _OnCustomDraw((NMTBCUSTOMDRAW*)pnm);
            case TBN_WRAPACCELERATOR:
                return TRUE;            // Disable wraparound; we want DeskHost to do navigation
            case TBN_GETINFOTIP:
                {
                    NMTBGETINFOTIP *ptbgit = (NMTBGETINFOTIP *)pnm;
                    ASSERT(ptbgit->lParam >= IDS_LOGOFF_TIP_EJECT && ptbgit->lParam <= IDS_LOGOFF_TIP_DISCONNECT);
                    LoadString(_AtlBaseModule.GetModuleInstance(), ptbgit->lParam, ptbgit->pszText, ptbgit->cchTextMax);
                    return TRUE;
                }
            case TBN_HOTITEMCHANGE:
                {
                    // Disallow setting a hot item if we are not focus
                    // (unless it was specifically our idea in the first place)
                    // Otherwise we interfere with keyboard navigation

                    NMTBHOTITEM *phot = (NMTBHOTITEM*)pnm;
                    if (!(phot->dwFlags & HICF_LEAVING) &&
                        GetFocus() != pnm->hwndFrom &&
                        !_fSettingHotItem)
                    {
                        return TRUE; // deny hot item change
                    }
                }
                break;

        }
    }
    else // from host
    {
        switch (pnm->code)
        {
            case SMN_REFRESHLOGOFF:
                _ApplyOptions();
                return TRUE;
            case SMN_FINDITEM:
                return _OnSMNFindItem(CONTAINING_RECORD(pnm, SMNDIALOGMESSAGE, hdr));
            case SMN_APPLYREGION:
                return HandleApplyRegion(_hwnd, _hTheme, (SMNMAPPLYREGION *)pnm, SPP_LOGOFF, 0);
        }
    }

    return FALSE;
}

LRESULT CLogoffPane::_OnCommand(int id)
{
    int i;
    for (i = 0; i < ARRAYSIZE(tbButtonsCreate); i++)
    {
        if (tbButtonsCreate[i].idCommand == id)
        {
            if (!_IsButtonHidden(i))
            {
                PostMessage(v_hwndTray, WM_COMMAND, c_rgidmLegacy[i], 0);
                SMNMCOMMANDINVOKED ci;
                SendMessage(_hwndTB, TB_GETITEMRECT, i,  (LPARAM)&ci.rcItem);
                MapWindowRect(_hwndTB, NULL, &ci.rcItem);
                _SendNotify(_hwnd, SMN_COMMANDINVOKED, &ci.hdr);
            }
            break;
        }
    }
    return 0;
}

LRESULT CLogoffPane::_OnCustomDraw(NMTBCUSTOMDRAW *pnmtbcd)
{
    LRESULT lres;

    switch (pnmtbcd->nmcd.dwDrawStage)
    {
        case CDDS_PREPAINT:
            return CDRF_NOTIFYITEMDRAW;
        
        case CDDS_ITEMPREPAINT:
#if 0       //Do we still need this?
            pnmtbcd->nHLStringBkMode = TRANSPARENT; // needed to reduce flicker- bug in toolbar?
#endif

            pnmtbcd->clrText = _clr;
            pnmtbcd->clrTextHighlight = GetSysColor(_colorHighlightText);
            pnmtbcd->clrHighlightHotTrack = GetSysColor(_colorHighlight);

            lres = TBCDRF_NOEDGES | TBCDRF_HILITEHOTTRACK;

            // todo - FIX TOOLBAR to respect clrTextHighlight when item is hot.
            if (pnmtbcd->nmcd.uItemState == CDIS_HOT)
            {
                pnmtbcd->clrText = pnmtbcd->clrTextHighlight;
            }

            return lres;
    }
    return CDRF_DODEFAULT;
}

LRESULT CLogoffPane::_NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction)
{
    ASSERT(direction == +1 || direction == -1);

    i += direction;
    while (i >= 0 && i < ARRAYSIZE(tbButtonsCreate))
    {
        if (!_IsButtonHidden(i))
        {
            pdm->itemID = i;
            return TRUE;
        }
        i += direction;
    }
    return FALSE;
}

int CLogoffPane::_GetCurButton()
{
    return (int)SendMessage(_hwndTB, TB_GETHOTITEM, 0, 0);
}

LRESULT CLogoffPane::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
    LRESULT lres = _OnSMNFindItemWorker(pdm);

    if (lres)
    {
        //
        //  If caller requested that the item also be selected, then do so.
        //
        if (pdm->flags & SMNDM_SELECT)
        {
            if (_GetCurButton() != pdm->itemID)
            {
                // Explicitly pop the tooltip so we don't have the problem
                // of a "virtual" tooltip causing the infotip to appear
                // the instant the mouse moves into the window.
                if (_hwndTT)
                {
                    SendMessage(_hwndTT, TTM_POP, 0, 0);
                }

                // _fSettingHotItem tells our WM_NOTIFY handler to
                // allow this hot item change to go through
                _fSettingHotItem = TRUE;
                SendMessage(_hwndTB, TB_SETHOTITEM, pdm->itemID, 0);
                _fSettingHotItem = FALSE;

                // Do the SetFocus after setting the hot item to prevent
                // toolbar from autoselecting the first button (which
                // is what it does if you SetFocus when there is no hot
                // item).  SetFocus returns the previous focus window.
                if (SetFocus(_hwndTB) != _hwndTB)
                {
                    // Send the notify since we tricked toolbar into not sending it
                    // (Toolbar doesn't send a subobject focus notification
                    // on WM_SETFOCUS if the item was already hot when it gained focus)
                    NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndTB, OBJID_CLIENT, pdm->itemID + 1);
                }
            }
        }
    }
    else
    {
        //
        //  If not found, then tell caller what our orientation is (horizontal)
        //  and where the currently-selected item is.
        //

        pdm->flags |= SMNDM_HORIZONTAL;
        int i = _GetCurButton();
        RECT rc;
        if (i >= 0 && SendMessage(_hwndTB, TB_GETITEMRECT, i, (LPARAM)&rc))
        {
            pdm->pt.x = (rc.left + rc.right)/2;
            pdm->pt.y = (rc.top + rc.bottom)/2;
        }
        else
        {
            pdm->pt.x = 0;
            pdm->pt.y = 0;
        }

    }
    return lres;
}

TCHAR CLogoffPane::_GetButtonAccelerator(int i)
{
    TCHAR szText[MAX_PATH];

    // First get the length of the text
    LRESULT lRes = SendMessage(_hwndTB, TB_GETBUTTONTEXT, tbButtonsCreate[i].idCommand, (LPARAM)NULL);

    // Check if the text will fit in our buffer.
    if (lRes > 0 && lRes < MAX_PATH)
    {
        if (SendMessage(_hwndTB, TB_GETBUTTONTEXT, tbButtonsCreate[i].idCommand, (LPARAM)szText) > 0)
        {
            return CharUpperCharW(SHFindMnemonic(szText));
        }
    }
    return 0;
}

//
//  Metrics changed -- update.
//
void CLogoffPane::_InitMetrics()
{
    if (_hwndTT)
    {
        // Disable/enable infotips based on user preference
        SendMessage(_hwndTT, TTM_ACTIVATE, ShowInfoTip(), 0);

        // Toolbar control doesn't set the tooltip font so we have to do it ourselves
        SetWindowFont(_hwndTT, GetWindowFont(_hwndTB), FALSE);
    }
}

LRESULT CLogoffPane::_OnDisplayChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // Propagate first because _InitMetrics needs to talk to the updated toolbar
    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    _InitMetrics();
    return 0;
}

LRESULT CLogoffPane::_OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // Propagate first because _InitMetrics needs to talk to the updated toolbar
    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    // _InitMetrics() is so cheap it's not worth getting too upset about
    // calling it too many times.
    _InitMetrics();
    _RightAlign();
    return 0;
}

LRESULT CLogoffPane::_OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // update colors in classic mode
    if (!_hTheme)
    {
        _clr = GetSysColor(COLOR_MENUTEXT);
    }

    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    return 0;
}



LRESULT CLogoffPane::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
    int i;
    switch (pdm->flags & SMNDM_FINDMASK)
    {
    case SMNDM_FINDFIRST:
        return _NextVisibleButton(pdm, -1, +1);

    case SMNDM_FINDLAST:
        return _NextVisibleButton(pdm, ARRAYSIZE(tbButtonsCreate), -1);

    case SMNDM_FINDNEAREST:
        // HACK! but we know that we are the only control in our group
        // so this doesn't need to be implemented
        return FALSE;

    case SMNDM_HITTEST:
        pdm->itemID = SendMessage(_hwndTB, TB_HITTEST, 0, (LPARAM)&pdm->pt);
        return pdm->itemID >= 0;

    case SMNDM_FINDFIRSTMATCH:
    case SMNDM_FINDNEXTMATCH:
        {
            if ((pdm->flags & SMNDM_FINDMASK) == SMNDM_FINDFIRSTMATCH)
            {
                i = 0;
            }
            else
            {
                i = _GetCurButton() + 1;
            }

            TCHAR tch = CharUpperCharW((TCHAR)pdm->pmsg->wParam);
            for ( ; i < ARRAYSIZE(tbButtonsCreate); i++)
            {
                if (_IsButtonHidden(i))
                    continue;               // skip hidden buttons
                if (_GetButtonAccelerator(i) == tch)
                {
                    pdm->itemID = i;
                    return TRUE;
                }
            }
        }
        break;      // not found

    case SMNDM_FINDNEXTARROW:
        switch (pdm->pmsg->wParam)
        {
        case VK_LEFT:
        case VK_UP:
            return _NextVisibleButton(pdm, _GetCurButton(), -1);

        case VK_RIGHT:
        case VK_DOWN:
            return _NextVisibleButton(pdm, _GetCurButton(), +1);
        }

        return FALSE;           // not found

    case SMNDM_INVOKECURRENTITEM:
        i = _GetCurButton();
        if (i >= 0 && i < ARRAYSIZE(tbButtonsCreate))
        {
            FORWARD_WM_COMMAND(_hwnd, tbButtonsCreate[i].idCommand, _hwndTB, BN_CLICKED, PostMessage);
            return TRUE;
        }
        return FALSE;

    case SMNDM_OPENCASCADE:
        return FALSE;           // none of our items cascade

    default:
        ASSERT(!"Unknown SMNDM command");
        break;
    }

    return FALSE;
}

HRESULT CLogoffPane::get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut)
{
    if (varChild.lVal)
    {
        return CreateAcceleratorBSTR(_GetButtonAccelerator(varChild.lVal - 1), pszKeyboardShortcut);
    }
    return CAccessible::get_accKeyboardShortcut(varChild, pszKeyboardShortcut);
}

HRESULT CLogoffPane::get_accDefaultAction(VARIANT varChild, BSTR *pszDefAction)
{
    if (varChild.lVal)
    {
        return GetRoleString(ACCSTR_EXECUTE, pszDefAction);
    }
    return CAccessible::get_accDefaultAction(varChild, pszDefAction);
}

// EXEX-VISTA: Partially reversed.
HRESULT CLogoffPane::_InitShutdownObjects()
{
    ASSERT(NULL == _psdc); // 805

    HRESULT hr = CoCreateInstance(CLSID_AuthUIShutdownChoices, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_psdc));

    if (SUCCEEDED(hr))
    {
        _psdc->SetShowBadChoices(TRUE);
        ASSERT(!_psdListen); // 811

        hr = _psdc->CreateListener(&_psdListen);

        if (SUCCEEDED(hr))
        {
            _psdListen->SetNotifyWnd(_hwnd, 0);
            hr = _psdListen->StartListening();
            if (SUCCEEDED(hr))
            {
                _psdListen->GetMessageWnd(&_hwndSdListenMsg);
            }
            else
            {
                // TODO: Trace
            }
        }
        else
        {
            // TODO: Trace.
        }
    }
    else
    {
        // TODO: Trace.
    }

    return hr;
}


BOOL WINAPI LogoffPane_RegisterClass()
{
    WNDCLASSEX wc;
    ZeroMemory( &wc, sizeof(wc) );
    
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_GLOBALCLASS;
    wc.cbWndExtra    = sizeof(LPVOID);
    wc.lpfnWndProc   = CLogoffPane::WndProc;
    wc.hInstance     = _AtlBaseModule.GetModuleInstance();
    wc.hCursor       = LoadCursor( NULL, IDC_ARROW );
    wc.hbrBackground = (HBRUSH)(NULL);
    wc.lpszClassName = TEXT("DesktopLogoffPane");

    return RegisterClassEx( &wc );
   
}
