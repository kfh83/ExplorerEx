#include "pch.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "shundoc.h"

class CLogoffPane;

// WARNING!  Must be in sync with c_rgidmLegacy

#define NUM_TBBUTTON_IMAGES 6

#ifdef DEAD_CODE
static const TBBUTTON tbButtonsCreate [] = 
{
    {0, SMNLC_EJECT,    TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_EJECT, 0},
    {1, SMNLC_LOGOFF,   TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_LOGOFF, 1},
    {2, SMNLC_TURNOFF,  TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_SHUTDOWN, 2},
    {2,SMNLC_DISCONNECT,TBSTATE_ENABLED, BTNS_SHOWTEXT|BTNS_AUTOSIZE, {0,0}, IDS_LOGOFF_TIP_DISCONNECT, 3},
};
#else
static const TBBUTTON tbButtonsCreate[] =
{
  { 1, SMNLC_TURNOFF,       TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7032, 1 },
  { 0, SMNLC_DISCONNECT,    TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7033, 2 },
  { 5, SMNLC_LOGOFF,        TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7034, 0 }
};
#endif

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
};

// EXEX-TODO: Move? Also, this interface changed between Windows versions. This is the Vista version.

enum SHUTDOWN_CHOICE
{
    SHUTDOWN_CHOICE_0 = 0,
	SHUTDOWN_CHOICE_1 = 1,
    SHUTDOWN_CHOICE_2 = 2,
    SHUTDOWN_CHOICE_3 = 3,
};

MIDL_INTERFACE("bfdc5e2f-3402-49b3-8740-91d6dc5dbb15")
IShutdownChoices : IUnknown
{
    STDMETHOD(Refresh)() PURE;
    STDMETHOD(CreateListener)(IShutdownChoiceListener **ppOut) PURE;
    STDMETHOD(SetShowBadChoices)(BOOL fShow) PURE;
    STDMETHOD(GetChoiceEnumerator)(IEnumShutdownChoices **ppOut) PURE;
    STDMETHOD(GetDefaultChoice)(DWORD *piOut) PURE; // This returns an ID representing the choice.
    STDMETHOD(GetMenuChoices)(IEnumShutdownChoices **ppOut) PURE;
    STDMETHOD_(BOOL, UserHasShutdownRights)() PURE;
    STDMETHOD(GetChoiceName)(DWORD choice, int a3, LPWSTR a4, UINT a5) PURE;
    STDMETHOD(GetChoiceDesc)(DWORD, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceVerb)(DWORD, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceIcon)(DWORD choice, int *pIcon) PURE;
};

MIDL_INTERFACE("bfdc5e2f-3402-49b3-8740-91d6dc5dbb15")
IShutdownChoices10 : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Refresh() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetChoiceMask(DWORD) = 0;
    virtual void STDMETHODCALLTYPE GetChoiceMask(DWORD *) = 0;
    virtual void STDMETHODCALLTYPE GetDefaultUIChoiceMask(DWORD *) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetShowBadChoices(int) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceEnumerator(IEnumShutdownChoices **) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDefaultChoice(SHUTDOWN_CHOICE *) = 0;
    virtual int STDMETHODCALLTYPE UserHasShutdownRights() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceName(DWORD, int, WCHAR *, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceDesc(DWORD, WCHAR *, UINT) = 0;
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
    STDMETHODIMP get_accName(VARIANT varChild, BSTR *pszName);
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
    LRESULT _OnCommand(int id, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCustomDraw(NMTBCUSTOMDRAW *pnmcd);
	LRESULT _OnCustomDrawSplitButton(DRAWITEMSTRUCT* pdis);
    LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);
    LRESULT _OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm);
    LRESULT _OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnDisplayChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    void    _InitMetrics();
    LRESULT _OnSize(int x, int y);

	void TBPressButton(WPARAM wParam, LPARAM lParam);
	void AddShutdownOptions(HMENU hMenu);
	void ApplyLogoffMenuOption(HMENU hMenu);
    int GetTipIDFromIDM(int id);
	HRESULT GetShutdownItemDescription(ULONG a2, WCHAR *pszDesc, UINT a4);

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
    IShutdownChoices10 *_psdc;
    IShutdownChoiceListener *_psdListen;
    HWND _hwndSdListenMsg;
    IAccessible* _pAcc;


    // helper functions
    int _GetCurButton();
    LRESULT _NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction);
    BOOL _IsButtonHiddenOrDisabled(int i, DWORD dwFlags);
    TCHAR _GetButtonAccelerator(int i);
    void _RightAlign();
    void _ApplyOptions();
    HRESULT _InitShutdownObjects();
    HRESULT _AddShutdownChoiceToMenu(HMENU hMenu, ULONG_PTR a3, UINT idMenuItem, UINT item);

    BOOL _SetTBButtons(int id, UINT iMsg);
    int _GetThemeBitmapSize(int iPartId, int iStateId, int id);

    int _GetIdmFromIdstip(int id);
	int _GetIdstipFromCommand(int id);
    int _GetIdmFromCommand(int id);
    int _GetCurPressedButton();
	int _HitTest(HWND hwndTest, POINT ptTest);
    BOOL _IsPtInDropDownSplit(HWND hwnd, POINT pt);
    void _DoSplitButtonContextMenu(int a2);
    int _IsTBButtonEnabled(int a2);
    void _SetShutdownButtonProperties(int a2);
    int _GetLocalImageForShutdownChoice(DWORD a2);
    

    friend BOOL CLogoffPane_RegisterClass();

    friend class CLogOffMenuCallback;
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
    CLogoffPane* self = reinterpret_cast<CLogoffPane*>(GetWindowLongPtr(hwnd, 0));

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
            DrawPlacesListBackground(self->_hTheme, hwnd, (HDC)wParam);
        }
        else
        {
            SHFillRectClr((HDC)wParam, &rc, GetSysColor(COLOR_MENU));
            DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_TOP);
        }

        return 1;
    }

    case WM_COMMAND:
        return self->_OnCommand(GET_WM_COMMAND_ID(wParam, lParam), wParam, lParam);

    case WM_NOTIFY:
        return self->_OnNotify((NMHDR*)(lParam));

    case WM_SIZE:
        return self->_OnSize(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

    case WM_DRAWITEM:
        return self->_OnCustomDrawSplitButton((DRAWITEMSTRUCT*)lParam);

    case WM_SYSCOLORCHANGE:
        return self->_OnSysColorChange(hwnd, uMsg, wParam, lParam);

    case WM_DISPLAYCHANGE:
        return self->_OnDisplayChange(hwnd, uMsg, wParam, lParam);

    case WM_SETTINGCHANGE:
        return self->_OnDisplayChange(hwnd, uMsg, wParam, lParam);
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
void AddBitmapToToolbar(HWND hwndTB, HBITMAP hBitmap, int cxTotal, int cy, UINT iNumImages, UINT iMsg)
{
    HIMAGELIST himl = ImageList_Create(cxTotal / iNumImages, cy, (ILC_MASK | ILC_COLOR32 | ILC_HIGHQUALITYSCALE), 0, iNumImages);
    if (himl)
    {
        ImageList_AddMasked(himl, hBitmap, 0xFF00FF);

        HIMAGELIST himlPrevious = (HIMAGELIST)SendMessage(hwndTB, iMsg, 0, (LPARAM)himl);
        if (himlPrevious)
        {
            ImageList_Destroy(himlPrevious);
        }
    }
}

// EXEX-VISTA: Validated.
BOOL CLogoffPane::_SetTBButtons(int id, UINT iMsg)
{
    HBITMAP hBitmap = (HBITMAP)LoadImage(_AtlBaseModule.GetModuleInstance(), MAKEINTRESOURCE(id), IMAGE_BITMAP,
        0, 0, LR_CREATEDIBSECTION);
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

int CLogoffPane::_GetThemeBitmapSize(int iPartId, int iStateId, int id)
{
    int cx = 0;
    if (_hTheme && iPartId)
    {
        SIZE siz;
        if (GetThemePartSize(_hTheme, 0, iPartId, iStateId, 0, TS_TRUE, &siz) >= 0)
        {
            cx = siz.cx;
        }
    }
    else
    {
        HBITMAP hBitmap = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(id));
        if (hBitmap)
        {
            BITMAP bm; // [esp+4h] [ebp-24h] BYREF
            if (GetObject(hBitmap, sizeof(BITMAP), &bm))
            {
                cx = bm.bmWidth;
            }
            DeleteObject(hBitmap);
        }
    }
    //SHLogicalToPhysicalDPI(&cx, 0);
    return cx;
}

class CSplitButtonAccessible : public CAccessible
{
public:
    CSplitButtonAccessible(HWND hwnd)
        : _hwnd(hwnd)
        , _cRef(1)
    {
    }

    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj) override
    {
        static const QITAB qit[] =
        {
            QITABENT(CSplitButtonAccessible, IDispatch),
            QITABENT(CSplitButtonAccessible, IAccessible),
            QITABENT(CSplitButtonAccessible, IEnumVARIANT),
            { 0 },
        };
        return QISearch(this, qit, riid, ppvObj);
    }

    STDMETHODIMP_(ULONG) AddRef() override
    {
        return InterlockedIncrement(&_cRef);
    }

    STDMETHODIMP_(ULONG) Release() override
    {
        LONG cRef = InterlockedDecrement(&_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

    // *** IAccessible ***
    STDMETHODIMP get_accRole(VARIANT varChild, VARIANT *pvarRole) override
    {
        pvarRole->vt = VT_I4;
        pvarRole->lVal = ROLE_SYSTEM_MENUITEM;
        return S_OK;
    }

    STDMETHODIMP accDoDefaultAction(VARIANT varChild) override
    {
        if (IsWindow(this->_hwnd) && IsWindowVisible(this->_hwnd))
        {
            PostMessage(GetParent(_hwnd), WM_COMMAND, 0x20063u, (LPARAM)_hwnd);
        }
        return S_OK;
    }

private:
    LONG _cRef;
    HWND _hwnd;
};

// EXEX-VISTA: Partially reversed.
LRESULT CLogoffPane::_OnCreate(LPARAM lParam)
{
    _InitShutdownObjects();

    // Do not set WS_TABSTOP here; that's CLogoffPane's job

    DWORD dwStyle = WS_CHILD|WS_CLIPSIBLINGS|WS_VISIBLE | CCS_NORESIZE|CCS_NODIVIDER | TBSTYLE_FLAT|TBSTYLE_LIST|TBSTYLE_TOOLTIPS;
    RECT rc;

    _hTheme = (PaneDataFromCreateStruct(lParam))->hTheme;

    int nResId = _hTheme ? IsHighDPI() ? 7054 : 7051 : IsHighDPI() ? 7015 : 7010;
    field_64 = _GetThemeBitmapSize(0, 0, nResId);
    if (!field_64)
        return -1;

    field_68 = _GetThemeBitmapSize(SPP_LOGOFFSPLITBUTTONDROPDOWN, 0, IsHighDPI() ? 7017 : 7014);
    if (!field_68)
        return -1;

    if (_hTheme)
    {
        GetThemeColor(_hTheme, SPP_LOGOFF, 0, TMT_TEXTCOLOR, &_clr);
        _colorHighlight = _clr;
        _colorHighlightText = _clr;

        GetThemeMargins(_hTheme, NULL, SPP_LOGOFF, 0, TMT_CONTENTMARGINS, NULL, &_margins);
    }
    else
    {
        _clr = GetSysColor(COLOR_MENUTEXT);
        _colorHighlight = GetSysColor(COLOR_HIGHLIGHT);
        _colorHighlightText = GetSysColor(COLOR_HIGHLIGHTTEXT);

        _margins.cyTopHeight = _margins.cyBottomHeight = 2 * GetSystemMetrics(SM_CYEDGE);
        field_68 = (UINT)field_68 >> 1;
        _margins.cxLeftWidth = GetSystemMetrics(SM_CXEDGE);
        ASSERT(_margins.cxLeftWidth == 0);
        ASSERT(_margins.cxRightWidth == 0);

        HBITMAP hbm = (HBITMAP)LoadImage(g_hinstCabinet,
            MAKEINTRESOURCE((IsHighDPI() ? IDB_LOGOFF_LARGE_EXPANDER : IDB_LOGOFF_EXPANDER)),
            IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
        if (hbm)
        {
            BITMAP bm;
            if (GetObject(hbm, sizeof(BITMAP), &bm))
            {
                _himl = ImageList_Create(bm.bmWidth / 2, bm.bmHeight, (ILC_MASK | ILC_COLOR32 | ILC_HIGHQUALITYSCALE), 0, 2);
                if (_himl)
                {
                    ImageList_AddMasked(_himl, hbm, 0xFF00FF);
                }
            }
            DeleteObject(hbm);
        }
    }

#ifdef DEAD_CODE
    GetClientRect(_hwnd, &rc);
    rc.left += _margins.cxLeftWidth;
    rc.right -= _margins.cxRightWidth;
    rc.top += _margins.cyTopHeight;
    rc.bottom -= _margins.cyBottomHeight;

    _hwndTB = CreateWindowEx(0, TOOLBARCLASSNAME, NULL, dwStyle,
                               rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc), _hwnd, 
                               NULL, NULL, NULL );
#else
    GetClientRect(_hwnd, &rc);
    rc.left += _margins.cxLeftWidth;
    rc.top += _margins.cyTopHeight;
    rc.right = rc.left + field_64;
    rc.bottom -= _margins.cyBottomHeight;

    _hwndTB = SHFusionCreateWindowEx(0, TOOLBARCLASSNAME, NULL, 0x54000944,
        rc.left, rc.top, field_64, RECTHEIGHT(rc), _hwnd,
        NULL, NULL, NULL);
#endif
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
        // SendMessage(_hwndTB, CCM_DPISCALE, TRUE, 0);

        SendMessage(_hwndTB, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
		SendMessage(_hwndTB, TB_SETLISTGAP, 0, 0);
        SendMessage(_hwndTB, TB_SETPADDING, 0, 0);

        if (!_hTheme
            || (!_SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_AERO_NORMAL : IDB_LOGOFF_AERO_NORMAL, TB_SETIMAGELIST))
            || (!_SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_AERO_HOT : IDB_LOGOFF_AERO_HOT, TB_SETHOTIMAGELIST))
            || (!_SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_AERO_PRESSED : IDB_LOGOFF_AERO_PRESSED, TB_SETPRESSEDIMAGELIST)))
        {
            // if we don't have a theme, or failed at setting the images from the theme
            // set buttons images from the rc file
            _SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_NORMAL : IDB_LOGOFF_NORMAL, TB_SETIMAGELIST);
            _SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_NORMAL : IDB_LOGOFF_NORMAL, TB_SETHOTIMAGELIST);
            _SetTBButtons(IsHighDPI() ? IDB_LOGOFF_LARGE_PRESSED : IDB_LOGOFF_PRESSED, TB_SETPRESSEDIMAGELIST);
        }
        
        SendMessage(_hwndTB, TB_ADDBUTTONS, ARRAYSIZE(tbButtonsCreate), (LPARAM) tbButtonsCreate);


		TBBUTTONINFO tbbi = { 0 };
		tbbi.cbSize = sizeof(tbbi);
		tbbi.dwMask = TBIF_STYLE;
		tbbi.fsStyle = BTNS_BUTTON;
		SendMessage(_hwndTB, TB_SETBUTTONINFO, 2, (LPARAM)&tbbi);

        _ApplyOptions();

        _hwndTT = (HWND)SendMessage(_hwndTB, TB_GETTOOLTIPS, 0, 0); //Get the tooltip window.

        _InitMetrics();

		HDC hdc = GetWindowDC(_hwnd);
        if (hdc)
        {
			TEXTMETRIC tm;
            if (GetTextMetrics(hdc, &tm))
            {
                LOGFONT lf;
				ZeroMemory(&lf, sizeof(lf));
                lf.lfHeight = tm.tmAscent;
				lf.lfWeight = FW_NORMAL;
                lf.lfCharSet = SYMBOL_CHARSET;
				StringCchCopy(lf.lfFaceName, ARRAYSIZE(lf.lfFaceName), L"Marlett");

				_hfMarlett = CreateFontIndirect(&lf);
                if (_hfMarlett)
                {
					HFONT hfMarlett = (HFONT)SelectObject(hdc, _hfMarlett);
                    if (GetTextMetrics(hdc, &tm))
                    {
                        field_58 = tm.tmAscent;
                    }
					SelectObject(hdc, hfMarlett);
				}
            }
			ReleaseDC(_hwnd, hdc);
        }

        rc.left = rc.right;
        rc.right += field_68;

		TCHAR szTitle[200];
        memset(szTitle, 0, sizeof(szTitle));
        LoadString(g_hinstCabinet, IDS_STARTPANE_TITLE_SPLITTER, szTitle, ARRAYSIZE(szTitle));

        _hwndSplit = SHFusionCreateWindowEx(0, WC_BUTTON, szTitle, 0x5600000Bu, rc.left, rc.top,
            RECTWIDTH(rc), RECTHEIGHT(rc), _hwnd, (HMENU)0x63, g_hinstCabinet, NULL);

        if (_hwndSplit)
        {
			CSplitButtonAccessible *pSplitAcc = new CSplitButtonAccessible(_hwndSplit);
            if (pSplitAcc)
            {
                SetAccessibleSubclassWindow(_hwndSplit);
				QueryInterface(IID_PPV_ARGS(&_pAcc));
                pSplitAcc->Release();
			}
        }
        
        return 0;
    }

    return -1; // no point in sticking around if we couldn't create the toolbar
}

BOOL CLogoffPane::_IsButtonHiddenOrDisabled(int i, DWORD dwFlags)
{
#ifdef DEAD_CODE
    TBBUTTON but;
    SendMessage(_hwndTB, TB_GETBUTTON, i, (LPARAM) &but);
    return but.fsState & TBSTATE_HIDDEN;
#else
    TBBUTTONINFO tbbi;
    tbbi.dwMask = dwFlags | 4;
    tbbi.cbSize = 0x20;
    SendMessage(this->_hwndTB, TB_GETBUTTONINFO, i, (LPARAM)&tbbi);
    return (tbbi.fsState & 0x18) != 0;
#endif
}

void CLogoffPane::_RightAlign()
{
#if 0
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
#endif
}

// #define DEAD_CODE

LRESULT CLogoffPane::_OnSize(int x, int y)
{
#ifdef DEAD_CODE
    if (_hwndTB)
    {
        SetWindowPos(_hwndTB, NULL, _margins.cxLeftWidth, _margins.cyTopHeight,
            x - (_margins.cxRightWidth + _margins.cxLeftWidth), y - (_margins.cyBottomHeight + _margins.cyTopHeight),
            SWP_NOACTIVATE | SWP_NOOWNERZORDER);
        _RightAlign();
    }
    return 0;
#else
    if (_hwndSplit)
    {
        SetWindowPos(_hwndSplit, NULL, _margins.cxLeftWidth + field_64, _margins.cyTopHeight, field_68,
            y - (_margins.cyBottomHeight - _margins.cyTopHeight),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }

    if (_hwndTB)
    {
        SetWindowPos(_hwndTB, NULL, _margins.cxLeftWidth, _margins.cyTopHeight, field_64,
            y - (_margins.cyBottomHeight - _margins.cyTopHeight),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    }
    return 0;
#endif
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

int __thiscall CLogoffPane::_GetLocalImageForShutdownChoice(DWORD a2)
{
    bool v2; // zf
    int v5; // [esp+0h] [ebp-4h] BYREF

    v2 = this->_psdc == 0;
    v5 = 1;
    if (!v2)
    {
        //this->_psdc->GetChoiceIcon(a2, &v5);
        if (v5 >= 2)
        {
            if (v5 <= 3)
                return 3;
            if (v5 == 4)
                return 4;
            if (v5 == 5)
                return 2;
        }
    }
    return 1;
}

void CLogoffPane::_SetShutdownButtonProperties(int a2)
{
    SHUTDOWN_CHOICE sdChoice; // [esp+34h] [ebp-1Ch] BYREF
    //CPPEH_RECORD ms_exc; // [esp+38h] [ebp-18h]

    if (this->_psdc && this->_psdc->GetDefaultChoice(&sdChoice) >= 0)
    {
        //if (sdChoice == SHUTDOWN_CHOICE_0
        //    && CcshellAssertFailedW(L"d:\\longhorn\\shell\\explorer\\desktop2\\logoff.cpp", 780, L"sdChoice != SHTDN_NONE", 0))
        //{
        //    AttachUserModeDebugger();
        //    do
        //    {
        //        __debugbreak();
        //        ms_exc.registration.TryLevel = -2;
        //    } while (dword_108BA88);
        //}

        if ((sdChoice & 0x40000) != 0)
        {
            if (a2)
                SendMessageW(this->_hwndTB, TB_SETSTATE, 1u, 16);
        }

		TBBUTTONINFO tbbi = {0};
        tbbi.cbSize = sizeof(tbbi);
        tbbi.dwMask = 1;
        tbbi.iImage = /*_GetLocalImageForShutdownChoice(sdChoice)*/ 3;
        //SHTracePerfSQMSetValueImpl(&ShellTraceId_StartMenu_Right_Control_Button_Label, 55, sdChoice);
        SendMessageW(this->_hwndTB, TB_SETBUTTONINFOW, 1u, (LPARAM)&tbbi);
    }
}

extern BOOL _ShowStartMenuShutdown();
extern BOOL _ShowStartMenuDisconnect();
extern BOOL _AllowLockWorkStation();

void CLogoffPane::_ApplyOptions()
{
#ifdef DEAD_CODE
    SMNFILTEROPTIONS nmopt;
    nmopt.smnop = SMNOP_LOGOFF | SMNOP_TURNOFF | SMNOP_DISCONNECT | SMNOP_EJECT;
    _SendNotify(_hwnd, SMN_FILTEROPTIONS, &nmopt.hdr);

    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_EJECT, !(nmopt.smnop & SMNOP_EJECT));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_LOGOFF, !(nmopt.smnop & SMNOP_LOGOFF));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, !(nmopt.smnop & SMNOP_TURNOFF));
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, !(nmopt.smnop & SMNOP_DISCONNECT));

    _RightAlign();
#else
    if (_psdc)
        _psdc->Refresh();

    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, _ShowStartMenuShutdown() == 0);
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, _ShowStartMenuDisconnect() == 0);
    SendMessage(_hwndTB, TB_HIDEBUTTON, 0, _AllowLockWorkStation() == 0);

    SIZE siz; 
    if (SendMessage(_hwndTB, TB_GETMAXSIZE, 0, (LPARAM)&siz))
        field_64 = siz.cx;

    // SHTracePerfSQMSetValueImpl(&ShellTraceId_StartMenu_Left_Control_Button_Label, 54, 5);
    _SetShutdownButtonProperties(_ShowStartMenuShutdown());
#endif
}


LRESULT CLogoffPane::_OnNotify(NMHDR* pnm)
{
#ifdef DEAD_CODE
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
            NMTBGETINFOTIP* ptbgit = (NMTBGETINFOTIP*)pnm;
            ASSERT(ptbgit->lParam >= IDS_LOGOFF_TIP_EJECT && ptbgit->lParam <= IDS_LOGOFF_TIP_DISCONNECT);
            LoadString(_AtlBaseModule.GetModuleInstance(), ptbgit->lParam, ptbgit->pszText, ptbgit->cchTextMax);
            return TRUE;
        }
        case TBN_HOTITEMCHANGE:
        {
            // Disallow setting a hot item if we are not focus
            // (unless it was specifically our idea in the first place)
            // Otherwise we interfere with keyboard navigation

            NMTBHOTITEM* phot = (NMTBHOTITEM*)pnm;
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
            return HandleApplyRegion(_hwnd, _hTheme, (SMNMAPPLYREGION*)pnm, SPP_LOGOFF, 0);
        }
    }

    return FALSE;
#else
    // eax
    // eax
    // eax
    HWND v8; // eax
    UINT v9; // eax
    NMHDR* v10; // eax
    HWND v11; // edi
    LRESULT v13; // eax
    //CPPEH_RECORD ms_exc; // [esp+30h] [ebp-18h]

    if (pnm->hwndFrom == this->_hwndTB)
    {
        switch (pnm->code)
        {
        case 0xFFFFFD2A:
            pnm[1].idFrom = -1;
            break;
        case 0xFFFFFD31:
            if (_GetCurPressedButton() == -1)
            {
                NMTBGETINFOTIP* ptbgit = (NMTBGETINFOTIP*)pnm;
                if (ptbgit->lParam == 7032)
                {
                    if (this->_psdc)
                    {
                        SHUTDOWN_CHOICE dwChoice;
                        if (this->_psdc->GetDefaultChoice(&dwChoice) >= 0)
                        {
                            this->_psdc->GetChoiceDesc(dwChoice, ptbgit->pszText, ptbgit->cchTextMax);
                        }
                    }
                }
                else
                {
                    ASSERT(ptbgit->lParam >= IDS_LOGOFF_TIP_EJECT/*7030*/ && ptbgit->lParam <= IDS_LOGOFF_TIP_LAST/*7036*/); // 879
                    if (ptbgit->lParam)
                    {
						LoadString(g_hinstCabinet, ptbgit->lParam, ptbgit->pszText, ptbgit->cchTextMax);
                    }
                }
            }
            break;
        case 0xFFFFFFF4:
            return _OnCustomDraw((NMTBCUSTOMDRAW*)pnm);
        default:
            return 0;
        }
        return 1;
    }

    if (pnm->hwndFrom == this->_hwndSdListenMsg)
    {
        _ApplyOptions();
        _SendNotify(this->_hwnd, 209, 0);
        return 0;
    }

    v9 = pnm->code;
    switch (v9)
    {
    case 0xD5:
        _ApplyOptions();
        break;
    case 0xD6:
        if (GetFocus() == this->_hwndTB)
        {
            v13 = SendMessageW(this->_hwndTB, 0x447, 0, 0);
            NotifyWinEvent(0x8005u, this->_hwndTB, -4, v13 + 1);
        }
        goto LABEL_29;
    case 0xD7u:
        return _OnSMNFindItem((SMNDIALOGMESSAGE*)pnm);
    case 0xDDu:
        if (this->_psdListen)
        {
            // SHTracePerf(&ShellTraceId_Explorer_ShutdownUX_StartMenuCriticalPath_Start);
            this->_psdListen->ScanForPassiveChange();
            // SHTracePerf(&ShellTraceId_Explorer_ShutdownUX_StartMenuCriticalPath_Stop);
        }
        return 0;
    case 224:
        if (_GetCurPressedButton() == 99)
        {
            TBBUTTONINFO tbbi;
            tbbi.cbSize = sizeof(tbbi);
            tbbi.dwMask = 32;
            tbbi.idCommand = 0;
            v10 = (NMHDR*)SendMessage(this->_hwndTB, TB_GETBUTTONINFO, 0, (LPARAM)&tbbi);
            pnm = v10;
            this->_fSettingHotItem = 1;
            SendMessage(this->_hwndTB, TB_SETHOTITEM, (WPARAM)v10, 0);
            this->_fSettingHotItem = 0;
            SendMessage(this->_hwndSplit, 0xF3u, 0, 0);

            if (SetFocus(this->_hwndTB) != this->_hwndTB)
            {
                NotifyWinEvent(0x8005u, this->_hwndTB, -4, (LONG)&pnm->hwndFrom + 1);
            }
        }
        return 0;
    case 225:
        break;
    case 0xFFFFFFF8:
    LABEL_29:
        if (this->field_5c || CLogoffPane::_GetCurPressedButton() == 99)
        {
            this->field_5c = 0;
            SendMessageW(this->_hwndSplit, BM_SETSTATE, 0, 0);
            InvalidateRect(this->_hwndSplit, 0, 0);
        }
        return 0;
    default:
        return 0;
    }
    return 1;
#endif
}

static const struct
{
    int smnCmd;
    int idCmd;
    int idsTip;
} c_rgButtons[] =
{
  { 0, 5000, 7035 },
  { 0, 410, 7030 },
  { 0, 402, 7031 },
  { 0, 517, 7034 },
  { 1, 506, 7032 },
  { 2, 5000, 7033 }
};

int CLogoffPane::_GetIdmFromIdstip(int id)
{
    UINT i = 0;
    UINT i_p = 0;
    while (c_rgButtons[i_p].idsTip != id)
    {
        ++i_p;
        ++i;
        if (i_p >= 6)
        {
            return -1;
        }
    }
    return c_rgButtons[i].idCmd;
}

int CLogoffPane::_GetIdstipFromCommand(int id)
{
    TBBUTTONINFO tbbi = { 0 };
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = 0x10;
    if (SendMessage(_hwndTB, TB_GETBUTTONINFO, id, (LPARAM)&tbbi) == -1)
        return -1;
    else
        return tbbi.lParam;
}

int CLogoffPane::_GetIdmFromCommand(int id)
{
    int IdstipFromCommand = _GetIdstipFromCommand(id);
    return _GetIdmFromIdstip(IdstipFromCommand);
}

int CLogoffPane::_GetCurPressedButton()
{
    if (SendMessage(this->_hwndTB, TB_ISBUTTONPRESSED, 0, 0))
        return 0;
    if (SendMessage(this->_hwndTB, TB_ISBUTTONPRESSED, 1, 0))
        return 1;

    if ((SendMessage(this->_hwndSplit, BM_GETSTATE, 0, 0) & 4) != 0)
        return 99;
    return -1;
}

HRESULT CLogOffMenuCallback_CreateInstance(IShellMenuCallback **ppsmc, CLogoffPane *pLogOffPane);

void CLogoffPane::_DoSplitButtonContextMenu(int a2)
{
    SendMessage(_hwndTB, TB_SETHOTITEM, -1, 0);

    IShellMenu* psm;
    if (SUCCEEDED(CoCreateInstance(CLSID_MenuBand, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psm))))
    {
        IShellMenuCallback* psmc = NULL;
        HRESULT hr = CLogOffMenuCallback_CreateInstance(&psmc, this);

        RECT rc;
        GetClientRect(_hwndSplit, &rc);
        MapWindowRect(_hwndSplit, NULL, &rc);

        IMAGEINFO imageInfo;
        if (!_hTheme && ImageList_GetImageInfo(_himl, 0, (IMAGEINFO*)&imageInfo))
        {
            rc.right = imageInfo.rcImage.right + rc.left;
            rc.bottom = imageInfo.rcImage.bottom + rc.top;
        }

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(psm->Initialize(psmc, 0, -1, 0x10000047)))
            {
                SendMessage(this->_hwndSplit, 0xF3, 1u, 0);
                //SHTracePerfSQMCountImpl(&ShellTraceId_StartMenu_Left_Control_Button_Split_Open, 57);

                DWORD dwFlags = 0;
                if (a2)
                {
                    dwFlags |= MPPF_KEYBOARD | MPPF_FINALSELECT;
                }

                SMNTRACKSHELLMENU nmtsm;
                nmtsm.psm = psm;
                nmtsm.rcExclude = rc;
                nmtsm.itemID = 99;
                nmtsm.dwFlags = dwFlags;

                _SendNotify(_hwnd, 216, &nmtsm.hdr);
            }
            psmc->Release();
        }
        psm->Release();
    }
}

LRESULT CLogoffPane::_OnCommand(int id, WPARAM wParam, LPARAM lParam)
{
    LPARAM v5; // edi
    WPARAM IdmFromCommand; // ebx
    int v7; // eax
    IShutdownChoices* psdc; // eax
    // eax
    HWND hwndSplit; // [esp-10h] [ebp-3Ch]
    // [esp-4h] [ebp-30h]
    // [esp+28h] [ebp-4h]

    int v14 = 0;
    if (id == 4)
    {
        v5 = lParam;
        if (!lParam)
        {
            SendMessageW(this->_hwndTB, TB_PRESSBUTTON, 1u, 0);
            return 0;
        }
        IdmFromCommand = _GetIdmFromCommand(1);
        v7 = v5 | 0x10000000;
    }
    else
    {
        if (id == 99)
        {
            if (HIWORD(wParam) == 2)
            {
                _DoSplitButtonContextMenu(0);
            }
            else if (HIWORD(wParam) == 7 && (this->field_5c || _GetCurPressedButton() == 99))
            {
                hwndSplit = this->_hwndSplit;
                this->field_5c = 0;
                SendMessageW(hwndSplit, BM_SETSTATE, 0, 0);
                InvalidateRect(this->_hwndSplit, NULL, 0);
            }
            return 0;
        }
        if (_IsButtonHiddenOrDisabled(id, 0))
            return 0;

        IdmFromCommand = _GetIdmFromCommand(id);
        v7 = IdmFromCommand | 0x20000000;
        v14 = 1;
        if (id == 1)
        {
            // SHTracePerf(&ShellTraceId_Explorer_ShutdownUX_DefaultButtonPress_Start);
            if (_psdc && _psdc->GetDefaultChoice((SHUTDOWN_CHOICE*)&lParam) >= 0)
                v5 = lParam;
            else
                v5 = 2;
            v7 = v5 | 0x10000000;
        }
        else
        {
            v5 = lParam;
        }
    }
    
    // DWORD dwTick = GetTickCount();
    // SHTracePerfSQMStreamTwoImpl(&ShellTraceId_StartMenu_Logoff_Usage_Stream, 59, dwTick, v7);
    if (IdmFromCommand)
    {
        PostMessageW(v_hwndTray, WM_COMMAND, IdmFromCommand, v5);
        if (v14)
        {
            SMNMCOMMANDINVOKED ci;
            SendMessageW(_hwndTB, TB_GETITEMRECT, IdmFromCommand, (LPARAM)&ci.rcItem);
            MapWindowPoints(_hwndTB, NULL, (LPPOINT)&ci.rcItem, 2u);
            _SendNotify(_hwnd, 204u, &ci.hdr);
        }
    }
    return 0;
}

LRESULT CLogoffPane::_OnCustomDraw(NMTBCUSTOMDRAW* pnmtbcd)
{
    LRESULT lres;

    switch (pnmtbcd->nmcd.dwDrawStage)
    {
    case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;

    case CDDS_ITEMPREPAINT:
        pnmtbcd->clrText = _clr;
        pnmtbcd->clrTextHighlight = _colorHighlightText;
        pnmtbcd->clrHighlightHotTrack = _colorHighlight;

        lres = TBCDRF_NOEDGES | TBCDRF_NOOFFSET | TBCDRF_NOBACKGROUND;

        if (pnmtbcd->nmcd.dwItemSpec == 0)
        {
            if ((pnmtbcd->nmcd.uItemState & CDIS_HOT) != 0)
            {
                if (field_5c)
                {
                    pnmtbcd->nmcd.uItemState = pnmtbcd->nmcd.uItemState & 0xFFFFFFBF;
                }
            }
        }

        if (!_hTheme && (pnmtbcd->nmcd.uItemState & CDIS_HOT) != 0)
        {
            SHFillRectClr(pnmtbcd->nmcd.hdc, &pnmtbcd->nmcd.rc, GetSysColor(COLOR_HIGHLIGHT));
        }

        return lres;
    }
    return CDRF_DODEFAULT;
}

LRESULT CLogoffPane::_OnCustomDrawSplitButton(DRAWITEMSTRUCT* pdis)
{
    IMAGEINFO pImageInfo;
    RECT pRect;
    DWORD chText;

    if (pdis->hwndItem == this->_hwndSplit && (pdis->itemAction & 0x45) != 0)
    {
        unsigned int v4 = SendMessageW(this->_hwndTB, TB_GETBUTTONSIZE, 0, 0);
        LONG v5 = pdis->rcItem.right - pdis->rcItem.left;
        bool v6 = this->_hTheme == NULL;
        pRect.left = 0;
        pRect.top = 0;
        int v7 = this->field_5c;
        pRect.right = v5;
        pRect.bottom = HIWORD(v4);
        if (v6)
        {
            int ia = pdis->itemState & 1;
            DWORD SysColor = GetSysColor(4);
            SHFillRectClr(pdis->hDC, &pdis->rcItem, SysColor);
            int mode = SetBkMode(pdis->hDC, 1);
            if (ImageList_GetImageInfo(this->_himl, ia, &pImageInfo))
            {
                if (pRect.right > pImageInfo.rcImage.right)
                {
                    pRect.right = pImageInfo.rcImage.right;
                }
                
                if (v7)
                {
                    SHFillRectClr(pdis->hDC, &pRect, GetSysColor(13));
                }
                ImageList_DrawEx(
                    this->_himl,
                    ia,
                    pdis->hDC,
                    pRect.left,
                    pRect.top,
                    0,
                    pRect.bottom - pRect.top,
                    0xFFFFFFFF,
                    0xFFFFFFFF,
                    0);
            }
            
            HFONT h = (HFONT)SelectObject(pdis->hDC, this->_hfMarlett);
            if (h)
            {
                char Layout = GetLayout(pdis->hDC);
                UINT ib = 1;
                pRect.top += (pRect.bottom - field_58 - pRect.top) / 2;
                if ((Layout & 1) != 0)
                {
                    ib = 0x20001;
                }
                chText = (unsigned __int16)((Layout & 1) != 0 ? 119 : 56);
                DrawText(pdis->hDC, (LPCWSTR)&chText, 1, &pRect, ib);
                SetTextColor(pdis->hDC, SetTextColor(pdis->hDC, GetSysColor(v7 != 0 ? 14 : 7)));
                SelectObject(pdis->hDC, h);
            }
            
            if (mode)
            {
                SetBkMode(pdis->hDC, mode);
            }
        }
        else
        {
            int v8 = (v7 != 0) + 1;
            if ((pdis->itemState & 1) != 0)
            {
                v8 = 3;
            }
            DrawThemeParentBackground(this->_hwndSplit, pdis->hDC, &pdis->rcItem);
            DrawThemeBackground(this->_hTheme, pdis->hDC, 19, v8, &pRect, NULL);
        }
    }
    return 1;
}

LRESULT CLogoffPane::_NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction)
{
#ifdef DEAD_CODE
    ASSERT(direction == +1 || direction == -1);

    i += direction;
    while (i >= 0 && i < ARRAYSIZE(tbButtonsCreate))
    {
        if (!_IsButtonHiddenOrDisabled(i, 0x80000000))
        {
            pdm->itemID = i;
            return TRUE;
        }
        i += direction;
    }
    return FALSE;
#else
    int v4; // esi
    bool v5; // zf

    ASSERT(direction == +1 || direction == -1) // 1400

    v4 = i;
    if (i == 99)
        v4 = 3;
    while (1)
    {
        v4 += direction;
        if (v4 < 0)
            break;
        v5 = v4 == 3;
        if ((unsigned int)v4 >= 3)
            goto LABEL_12;
        if (!CLogoffPane::_IsButtonHiddenOrDisabled(v4, 0x80000000))
        {
            pdm->itemID = v4;
            return 1;
        }
    }
    v5 = v4 == 3;
LABEL_12:
    if (v5)
    {
        pdm->itemID = 99;
        return 1;
    }
    return 0;
#endif
}

int CLogoffPane::_GetCurButton()
{
    if (field_5c)
    {
        return 99;
    }
    return (int)SendMessage(_hwndTB, TB_GETHOTITEM, 0, 0);
}

BOOL CLogoffPane::_IsPtInDropDownSplit(HWND hwnd, POINT pt)
{
    return CLogoffPane::_HitTest(hwnd, pt) == 99;
}

LRESULT CLogoffPane::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
#ifdef DEAD_CODE
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
#else
    int CurPressedButton; // eax
    LRESULT CurButton; // eax
    LPARAM itemID; // ecx
    HWND hwndSplit; // edx
    HWND hwndTT; // eax
    int IsPtInDropDownSplit; // eax
    LRESULT v10; // eax
    LONG bottom; // ecx
    HWND hwndTB; // [esp-10h] [ebp-30h]
    HWND v14; // [esp-10h] [ebp-30h]
    RECT lParam; // [esp+Ch] [ebp-14h] BYREF
    LRESULT lres; // [esp+1Ch] [ebp-4h]
    LPARAM hwnda; // [esp+28h] [ebp+8h]
    HWND hwndb; // [esp+28h] [ebp+8h]

    lres = CLogoffPane::_OnSMNFindItemWorker(pdm);
    if (lres)
    {
        if ((pdm->flags & 0x900) != 0)
        {
            CurPressedButton = CLogoffPane::_GetCurPressedButton();
            if (CurPressedButton == -1 || CurPressedButton == 99)
            {
                CurButton = CLogoffPane::_GetCurButton();
                itemID = pdm->itemID;
                pdm->flags |= 0x80000u;
                if (itemID == 99)
                    hwndSplit = this->_hwndSplit;
                else
                    hwndSplit = this->_hwndTB;
                pdm[1].hdr.hwndFrom = hwndSplit;
                if (CurButton != itemID)
                {
                    hwndTT = this->_hwndTT;
                    if (hwndTT)
                        SendMessageW(hwndTT, 0x41Cu, 0, 0);
                    hwnda = pdm->itemID;
                    if (hwnda == 99)
                    {
                        hwnda = -1;
                    }
                    else
                    {
                        hwndTB = this->_hwndTB;
                        this->_fSettingHotItem = 1;
                        SendMessageW(hwndTB, 0x448u, hwnda, 0);
                        this->_fSettingHotItem = 0;
                    }
                    if ((pdm->flags & 0x100) != 0 && hwnda != -1)
                    {
                        SendMessageW(this->_hwndSplit, 0xF3u, 0, 0);
                        hwndb = this->_hwndTB;
                        if (SetFocus(hwndb) != hwndb)
                            NotifyWinEvent(0x8005u, hwndb, -4, pdm->itemID + 1);
                    }
                }
                IsPtInDropDownSplit = _IsPtInDropDownSplit(pdm->hwnd, pdm->pt);
                if (IsPtInDropDownSplit != this->field_5c)
                {
                    this->field_5c = IsPtInDropDownSplit;
                    if (!IsPtInDropDownSplit)
                        SendMessageW(this->_hwndSplit, 0xF3u, 0, 0);
                    goto LABEL_7;
                }
            }
            else
            {
                pdm->flags &= ~0x100u;
                if (this->field_5c)
                {
                    this->field_5c = 0;
                LABEL_7:
                    InvalidateRect(this->_hwndSplit, 0, 0);
                }
            }
        }
    }
    else
    {
        pdm->flags |= 0x4000u;
        v10 = CLogoffPane::_GetCurButton();
        if (v10 >= 0 && SendMessageW(this->_hwndTB, TB_GETITEMRECT, v10, (LPARAM)&lParam))
        {
            bottom = lParam.bottom;
            pdm->pt.x = (lParam.right + lParam.left) / 2;
            pdm->pt.y = (bottom + lParam.top) / 2;
        }
        else
        {
            pdm->pt.x = 0;
            pdm->pt.y = 0;
        }
        if (this->field_5c)
        {
            v14 = this->_hwndSplit;
            this->field_5c = 0;
            SendMessageW(v14, 0xF3u, 0, 0);
            InvalidateRect(this->_hwndSplit, 0, 0);
        }
    }
    return lres;
#endif
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

int CLogoffPane::_HitTest(HWND hwndTest, POINT ptTest)
{
    HWND v4; // eax
    HWND hwndTB; // ecx

    if (hwndTest)
        MapWindowPoints(hwndTest, this->_hwnd, &ptTest, 1u);
    v4 = ChildWindowFromPointEx(this->_hwnd, ptTest, 1u);
    hwndTB = this->_hwndTB;
    if (v4 != hwndTB)
        return v4 != this->_hwndSplit ? -1 : 99;
    MapWindowPoints(this->_hwnd, hwndTB, &ptTest, 1u);
    return SendMessageW(this->_hwndTB, 0x445u, 0, (LPARAM)&ptTest);
}

int __thiscall CLogoffPane::_IsTBButtonEnabled(int wParam)
{
    HWND hwndTB; // [esp-10h] [ebp-30h]
    TBBUTTONINFOW tbbi; // [esp+0h] [ebp-20h] BYREF

    tbbi.cbSize = 0x20;
    hwndTB = this->_hwndTB;
    tbbi.dwMask = 0x80000004;
    SendMessageW(hwndTB, 0x43Fu, wParam, (LPARAM)&tbbi);
    return (tbbi.fsState >> 2) & 1;
}

LRESULT CLogoffPane::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
#ifdef DEAD_CODE
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
                if (_IsButtonHiddenOrDisabled(i, 0))
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
#else
    UINT flags; // eax
    int VisibleButton; // ebx
    LRESULT v7; // eax
    MSG *pmsg; // eax
    LRESULT v9; // eax
    LRESULT CurButton; // eax
    unsigned int v11; // ebx
    LRESULT v12; // eax
    struct SMNDIALOGMESSAGE *pnmdma; // [esp+30h] [ebp+8h]

    flags = pdm->flags;
    pnmdma = (struct SMNDIALOGMESSAGE *)(flags & 0xF);

    switch (flags & 0xF)
    {
    case 0u:
    case 1u:
    case 0xBu:
        return 0;
    case 2u:
        if ((flags & 0x10000) == 0)
            return CLogoffPane::_NextVisibleButton(pdm, -1, 1);
        return CLogoffPane::_NextVisibleButton(pdm, 3, -1);
    case 3u:
        return CLogoffPane::_NextVisibleButton(pdm, -1, 1);
    case 4u:
        VisibleButton = CLogoffPane::_NextVisibleButton(pdm, 3, -1);
        if (VisibleButton)
        {
            if (pdm->itemID == 99)
                CLogoffPane::_DoSplitButtonContextMenu(1);
        }
        return VisibleButton;
    case 5u:
        if (pdm->pmsg->wParam == 37)
        {
            CurButton = CLogoffPane::_GetCurButton();
            if (!CLogoffPane::_NextVisibleButton(pdm, CurButton, -1))
                return 0;
            goto LABEL_21;
        }
        if (pdm->pmsg->wParam == 39)
        {
            v9 = CLogoffPane::_GetCurButton();
            if (CLogoffPane::_NextVisibleButton(pdm, v9, 1))
            {
                if (pdm->itemID == 99)
                {
                    CLogoffPane::_DoSplitButtonContextMenu(1);
                    pdm->flags &= ~0x100u;
                }
            LABEL_21:
                pdm->flags |= 0x1000u;
                return 1;
            }
        }
        return 0;
    case 6u:
    case 8u:
        v11 = _GetCurButton();
        v12 = _HitTest(pdm->hwnd, pdm->pt);
        pdm->itemID = v12;
        if (v12 >= 0)
            v11 = v12;

        if (v11 == 99)
        {
            if (CLogoffPane::_GetCurPressedButton() != 99)
                PostMessageW(this->_hwnd, WM_COMMAND, GET_WM_COMMAND_MPS(99, _hwndSplit, BN_HILITE));
            return 1;
        }
        if (v11 <= 2 && _IsTBButtonEnabled(v11))
        {
            if (pnmdma == (struct SMNDIALOGMESSAGE *)6)
                PostMessage(this->_hwnd, WM_COMMAND, LOWORD(tbButtonsCreate[v11].idCommand), (LPARAM)this->_hwndTB);
            return 1;
        }
        return 0;

    case 7u:
        v7 = _HitTest(pdm->hwnd, pdm->pt);
        pdm->itemID = v7;
        if (v7 == 99)
        {
            pmsg = pdm->pmsg;
            if (pmsg)
            {
                if (pmsg->message == 0x201 && CLogoffPane::_GetCurPressedButton() != 99)
                    PostMessageW(this->_hwnd, WM_COMMAND, GET_WM_COMMAND_MPS(99, _hwndSplit, BN_HILITE));
            }
        }
        return pdm->itemID >= 0;
    case 9u:
        pdm->flags = flags & 0xFFFFFEFF;
        return 1;
    case 0xAu:
        if (CLogoffPane::_GetCurButton() != 99)
            return 0;
        CLogoffPane::_DoSplitButtonContextMenu(1);
        return 1;
    default:
        ASSERT(!"Unknown SMNDM command"); // 1767
        return 0;
    }
#endif
}

HRESULT CLogoffPane::get_accName(VARIANT varChild, BSTR *pszName)
{
    OLECHAR *v3; // eax
    SHUTDOWN_CHOICE v5; // [esp+Ch] [ebp-D0h] BYREF
    OLECHAR psz[100]; // [esp+10h] [ebp-CCh] BYREF

    if (!varChild.lVal
        || SendMessageW(_hwndTB, TB_COMMANDTOINDEX, 1u, 0) != varChild.lVal - 1
        || !_psdc
        || _psdc->GetDefaultChoice(&v5) < 0
        || _psdc->GetChoiceName(v5, 0, psz, 100) < 0)
    {
        return CAccessible::get_accName(varChild, pszName);
    }
    v3 = SysAllocString(psz);
    *pszName = v3;
    return v3 != 0 ? 0 : 0x8007000E;
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

        //DWORD dwChoiceMask = 0x400781;
        //if (_ShowStartMenuShutdownVista())
        //{
        //    dwChoiceMask |= 0x20006;
        //    if (!IsOS(OS_ANYSERVER))
        //        dwChoiceMask |= 0x200050;
        //}
        //dwChoiceMask &= ~0x200000; // @MOD To add Switch user, Sign out, and Lock options. Taken from @ex7
        //_psdc->SetChoiceMask(dwChoiceMask);

        //ASSERT(!_psdListen); // 811

        //hr = _psdc->CreateListener(&_psdListen);

        //if (SUCCEEDED(hr))
        //{
        //    _psdListen->SetNotifyWnd(_hwnd, 0);
        //    hr = _psdListen->StartListening();
        //    if (SUCCEEDED(hr))
        //    {
        //        _psdListen->GetMessageWnd(&_hwndSdListenMsg);
        //    }
        //    else
        //    {
        //        // TODO: Trace
        //    }
        //}
        //else
        //{
        //    // TODO: Trace.
        //}
    }
    else
    {
        // TODO: Trace.
    }

    return hr;
}

void CLogoffPane::TBPressButton(WPARAM wParam, LPARAM lParam)
{
    SendMessage(_hwndTB, TB_PRESSBUTTON, wParam, lParam);
}

void CLogoffPane::AddShutdownOptions(HMENU hMenu)
{
#if 1
    if (_ShowStartMenuShutdown())
    {
        if (_psdc)
        {
			MENUITEMINFO mi = {0};
			mi.cbSize = sizeof(mi);
            mi.fType = 2048;
            InsertMenuItem(hMenu, 0xFFFFu, 1024, &mi);

            //IEnumShutdownChoices *pesc;
            //HRESULT v4 = _psdc->GetMenuChoices(&pesc);
            //if (v4 >= 0)
            //{
            //    DWORD v7;
            //    for (UINT i = 100; !pesc->Next(1, &v7, 0); ++i)
            //    {
            //        if ((v7 & 0x400000) != 0)
            //            InsertMenuItem(hMenu, 0xFFFF, 1024, &mi);
            //        else
            //            v4 = _AddShutdownChoiceToMenu(hMenu, v7, i, 0xFFFFu);
            //        if (v4 < 0)
            //            break;
            //    }
            //    pesc->Release();
            //}
        }
    }
#else
    if (!_psdc)
        return;

    _psdc->SetShowBadChoices(TRUE);

    IEnumShutdownChoices *iterator;
    if (FAILED(_psdc->GetChoiceEnumerator(&iterator)))
        return;

    DWORD sc;
    while (iterator->Next(1, &sc, nullptr) == S_OK)
    {
        if ((sc & 0x400000) != 0)
        {
            MENUITEMINFOW mi = { sizeof(mi),MIIM_TYPE, MFT_SEPARATOR };
            InsertMenuItemW(hMenu, 0xFFFF, TRUE, &mi);
        }
        else
        {
            BOOL bC0000 = (sc & 0xC0000) != 0;
            sc &= ~0xC0000;
            //if (sc != m_scDefault && sc != (m_scDefault & ~0x20000))
            {
                WCHAR szChoiceName[200];
                if (SUCCEEDED(_psdc->GetChoiceName(sc, TRUE, szChoiceName, ARRAYSIZE(szChoiceName))))
                {
                    MENUITEMINFOW mi = { sizeof(mi) };
                    mi.fMask = MIIM_STATE | MIIM_ID | MIIM_TYPE;
                    mi.fType = MFT_STRING;
                    mi.fState = bC0000 ? MFS_DISABLED : MFS_ENABLED;
                    mi.wID = sc;
                    mi.dwTypeData = szChoiceName;
                    InsertMenuItemW(hMenu, 0xFFFF, 1, &mi);
                }
            }
        }
    }

    _SHPrettyMenu(hMenu);
    iterator->Release();
#endif
}

HRESULT CLogoffPane::_AddShutdownChoiceToMenu(
    HMENU hMenu,
    ULONG_PTR a3,
    UINT idMenuItem,
    UINT item)
{
	ASSERT(idMenuItem <= 150 /*IDM_SHUTDOWN_LAST*/); // 1074
	ASSERT(NULL != this->_psdc); // 1075

    WCHAR v11[200]; // [esp+48h] [ebp-1ACh] BYREF
    HRESULT hr = this->_psdc->GetChoiceName(a3, 1, v11, 200);
    if (hr >= 0)
    {
		MENUITEMINFO mi = { 0 };
		mi.cbSize = sizeof(mi);

        mi.fMask = 0x33;
        mi.fType = 0;
        mi.dwTypeData = v11;
        mi.wID = idMenuItem;
        mi.dwItemData = a3;
        mi.fState = (a3 & 0xC0000) != 0 ? 3 : 0;
        if (!InsertMenuItemW(hMenu, item, 0x400, &mi))
        {
            DWORD dwLastError = GetLastError();
            if ((int)dwLastError > 0)
                return (unsigned __int16)dwLastError | 0x80070000;
            return dwLastError;
        }
    }
    return hr;
}

extern BOOL _ShowStartMenuEject();
extern BOOL _ShowStartMenuLogoff();

const GUID POLID_HideFastUserSwitching =
{
  1462751767u,
  63487u,
  18332u,
  { 141u, 150u, 188u, 147u, 141u, 104u, 103u, 245u }
};


void CLogoffPane::ApplyLogoffMenuOption(HMENU hMenu)
{
    if (!_AllowLockWorkStation())
    {
        DeleteMenu(hMenu, 517, 0);
        DeleteMenu(hMenu, 5000, 0);
    }
    if (!IsOS(OS_FASTUSERSWITCHING)
        || GetSystemMetrics(SM_REMOTESESSION)
        || GetSystemMetrics(SM_REMOTECONTROL)
        || SHWindowsPolicy(POLID_HideFastUserSwitching))
    {
        DeleteMenu(hMenu, 5000, 0);
    }
    if (!_ShowStartMenuLogoff())
        DeleteMenu(hMenu, 402, 0);
    if (!_ShowStartMenuEject())
        DeleteMenu(hMenu, 410, 0);
}

HRESULT CLogoffPane::GetShutdownItemDescription(ULONG a2, WCHAR *pszDesc, UINT a4)
{
	ASSERT(NULL != this->_psdc); // 1106
    return this->_psdc->GetChoiceDesc(a2, pszDesc, a4);
}

int CLogoffPane::GetTipIDFromIDM(int id)
{
    int v2; // ecx
    unsigned int v3; // eax

    v2 = 0;
    v3 = 0;
    while (c_rgButtons[v3].idCmd != id)
    {
        ++v3;
        ++v2;
        if (v3 >= 6)
            return -1;
    }
    return c_rgButtons[v2].idsTip;
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


class CLogOffMenuCallback
    : public CUnknown
    , public IShellMenuCallback
{
public:
    CLogOffMenuCallback(CLogoffPane *pLogoffPane);

private:
    ~CLogOffMenuCallback();

public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // *** IShellMenuCallback ***
    STDMETHODIMP CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam) override;

private:
    DWORD _ShutdownChoiceFromMenuChoice(LPSMDATA psmd);

private:
    CLogoffPane *_pLogoffPane;
};

CLogOffMenuCallback::CLogOffMenuCallback(CLogoffPane *pLogoffPane)
    : _pLogoffPane(pLogoffPane)
{
    if (_pLogoffPane)
        _pLogoffPane->AddRef();
}

CLogOffMenuCallback::~CLogOffMenuCallback()
{
    IUnknown_SafeReleaseAndNullPtr(&_pLogoffPane);
}

HRESULT CLogOffMenuCallback::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CLogOffMenuCallback, IShellMenuCallback),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvObj);
}

ULONG CLogOffMenuCallback::AddRef()
{
    return CUnknown::AddRef();
}

ULONG CLogOffMenuCallback::Release()
{
    return CUnknown::Release();
}

HRESULT CLogOffMenuCallback::CallbackSM(LPSMDATA psmd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    UINT v5; // eax
    ULONG v6; // eax
    UINT TipIDFromIDM; // eax
    UINT uId; // eax
    unsigned int v10; // esi
    unsigned int v13; // [esp-4h] [ebp-Ch]

    switch (uMsg)
    {
    case 1u:
        IShellMenu * psm;
        if (psmd->punk && SUCCEEDED(psmd->punk->QueryInterface(IID_PPV_ARGS(&psm))))
        {
            HMENU hMenu = (HMENU)SHLoadMenuPopup(g_hinstCabinet, 6001);
            _pLogoffPane->ApplyLogoffMenuOption(hMenu);
            _pLogoffPane->AddShutdownOptions(hMenu);

            psm->SetMenu(hMenu, _pLogoffPane->_hwnd, 0);
            psm->Release();
            return S_OK;
        }
        break;
    case 3u:
        _pLogoffPane->TBPressButton(0, 0);
        return S_OK;
    case 4u:
        uId = psmd->uId;
        if (uId < 0x64 || uId > 0x96)
        {
            DWORD v13 = uId | 0x20000000;
            DWORD TickCount = GetTickCount();
            // SHTracePerfSQMStreamTwoImpl(&ShellTraceId_StartMenu_Logoff_Usage_Stream, 59, TickCount, v13);
            PostMessage(v_hwndTray, 0x111u, psmd->uId, 0);
        }
        else
        {
            v10 = _ShutdownChoiceFromMenuChoice(psmd);
            // SHTracePerf(&ShellTraceId_Explorer_ShutdownUX_SelectMenuItem_Start);
            PostMessage(_pLogoffPane->_hwnd, 0x111u, 4u, v10 | 0x10000000);
        }
        return S_OK;
    case 5u:
    {
        SMINFO *psmi = (SMINFO *)lParam;
        if ((psmi->dwMask & SMDM_TOOLBAR) != 0)
            psmi->iIcon = -1;
        return S_OK;
    }
    case 0xDu:
        if (psmd->uId >= 0x64 && psmd->uId <= 0x96)
        {
            v6 = _ShutdownChoiceFromMenuChoice(psmd);
            return _pLogoffPane->GetShutdownItemDescription(v6, (WCHAR *)wParam, (UINT)lParam);
        }

        TipIDFromIDM = _pLogoffPane->GetTipIDFromIDM(psmd->uId);
        if (TipIDFromIDM != -1)
        {
            LoadString(g_hinstCabinet, TipIDFromIDM, (LPWSTR)wParam, (int)lParam);
            return 0;
        }
        break;
    }
    return 1;
}

DWORD CLogOffMenuCallback::_ShutdownChoiceFromMenuChoice(LPSMDATA psmd)
{
    MENUITEMINFO mii = { 0 };
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_DATA;
    if (GetMenuItemInfo(psmd->hmenu, psmd->uId, FALSE, &mii))
        return mii.dwItemData;
    return 0;
}

HRESULT CALLBACK CLogOffMenuCallback_CreateInstance(IShellMenuCallback **ppsmc, CLogoffPane *pLogoffPane)
{
    CLogOffMenuCallback *plmc = new CLogOffMenuCallback(pLogoffPane);
    *ppsmc = SAFECAST(plmc, IShellMenuCallback *);
    return *ppsmc != NULL ? S_OK : E_OUTOFMEMORY;
}