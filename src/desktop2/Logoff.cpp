#include "pch.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "shundoc.h"

class CLogoffPane;

// WARNING!  Must be in sync with c_rgidmLegacy

#define NUM_TBBUTTON_IMAGES 6

static const TBBUTTON tbButtonsCreate[] =
{
    { 1, 1, TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7032, 1 },
    { 0, 2, TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7033, 2 },
    { 5, 0, TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, 7034, 0 }
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

enum SHTDN
{
    SHTDN_NONE = 0x0,
    SHTDN_LOGOFF = 0x1,
    SHTDN_SHUTDOWN = 0x2,
    SHTDN_RESTART = 0x4,
    SHTDN_RESTART_DOS = 0x8,
    SHTDN_SLEEP = 0x10,
    SHTDN_SLEEP2 = 0x20,
    SHTDN_HIBERNATE = 0x40,
    SHTDN_DISCONNECT = 0x80,
    SHTDN_SWITCHUSER = 0x100,
    SHTDN_LOCK = 0x200,
    SHTDN_EJECT = 0x400,
    SHTDN_MODIFIER_FORCE = 0x10000,
    SHTDN_MODIFIER_INSTALL_UPDATES = 0x20000,
    SHTDN_MODIFIER_ACCESS_DENIED = 0x40000,
    SHTDN_MODIFIER_BROKEN = 0x80000,
    SHTDN_MODIFIER_NO_WAKE = 0x100000,
    SHTDN_MODIFIER_SEPARATOR = 0x400000,
    SHTDN_MODIFIER_HYBRID = 0x800000,
    SHTDN_MODIFIER_FORCE_OTHERS = 0x1000000,
    SHTDN_MODIFIER_BOOTOPTIONS = 0x2000000,
};

DEFINE_ENUM_FLAG_OPERATORS(SHTDN);

typedef DWORD SHUTDOWN_CHOICE;

enum SHUTDOWN_CHOICE_ICON
{
    SHUTDOWN_CHOICE_ICON_1 = 0x1,
    SHUTDOWN_CHOICE_ICON_2 = 0x2,
    SHUTDOWN_CHOICE_ICON_3 = 0x3,
    SHUTDOWN_CHOICE_ICON_4 = 0x4,
    SHUTDOWN_CHOICE_ICON_5 = 0x5,
};

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

// Windows Vista version of the interface. Here for reference purposes.
MIDL_INTERFACE("a5dbd3dc-ee32-497a-ab84-f2c6aa5913f5")
IShutdownChoices_Vista : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Refresh() = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateListener(IShutdownChoiceListener**) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetShowBadChoices(BOOL) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceEnumerator(IEnumShutdownChoices**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDefaultChoice(SHUTDOWN_CHOICE*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetMenuChoices(IEnumShutdownChoices**) = 0;
    virtual BOOL STDMETHODCALLTYPE UserHasShutdownRights() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceName(SHUTDOWN_CHOICE, int, WCHAR*, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceDesc(SHUTDOWN_CHOICE, WCHAR*, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceVerb(SHUTDOWN_CHOICE, WCHAR*, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceIcon(SHUTDOWN_CHOICE, SHUTDOWN_CHOICE_ICON*) = 0;
};

// Windows 10 version of the interface. Missing many of the important methods from the Vista version.
MIDL_INTERFACE("bfdc5e2f-3402-49b3-8740-91d6dc5dbb15")
IShutdownChoices10 : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Refresh() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetChoiceMask(SHUTDOWN_CHOICE) = 0;
    virtual void STDMETHODCALLTYPE GetChoiceMask(SHUTDOWN_CHOICE*) = 0;
    virtual void STDMETHODCALLTYPE GetDefaultUIChoiceMask(SHUTDOWN_CHOICE*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetShowBadChoices(BOOL) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceEnumerator(IEnumShutdownChoices**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDefaultChoice(SHUTDOWN_CHOICE*) = 0;
    virtual BOOL STDMETHODCALLTYPE UserHasShutdownRights() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceName(SHUTDOWN_CHOICE, int, WCHAR*, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChoiceDesc(SHUTDOWN_CHOICE, WCHAR*, UINT) = 0;
};

class CLogoffPane
    : public CUnknown
    , public CAccessible
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override { return CUnknown::AddRef(); }
    STDMETHODIMP_(ULONG) Release() override { return CUnknown::Release(); }

    // *** IAccessible overridden methods ***
    STDMETHODIMP get_accName(VARIANT varChild, BSTR* pszName);
    STDMETHODIMP get_accKeyboardShortcut(VARIANT varChild, BSTR* pszKeyboardShortcut);
    STDMETHODIMP get_accDefaultAction(VARIANT varChild, BSTR* pszDefAction);

    CLogoffPane();
    ~CLogoffPane() override;

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
    void    _InitMetrics();
    LRESULT _OnSize(int x, int y);

	void TBPressButton(WPARAM wParam, LPARAM lParam);
	void AddShutdownOptions(HMENU hMenu);
	void ApplyLogoffMenuOption(HMENU hMenu);
    int GetTipIDFromIDM(int id);
    HRESULT GetShutdownItemDescription(SHUTDOWN_CHOICE sdChoice, WCHAR* pszDesc, UINT cchDesc);

private:
    HWND _hwnd;
    HWND _hwndTB;
    HWND _hwndSplit;
    HWND _hwndTT;
    COLORREF _clr;
    int _colorHighlight;
    int _colorHighlightText;
    HTHEME _hTheme;
    BOOL _fSettingHotItem;
    int field_40; // EXEX-TODO(isabella): Rename.
    MARGINS _margins;
    HFONT _hfMarlett;
    int field_58;
    int _fSplitButtonHot;
    HIMAGELIST _himl;
    int field_64;
    int field_68;
    IShutdownChoices10* _psdc;
    IShutdownChoiceListener* _psdListen;
    HWND _hwndSdListenMsg;
    IAccessible* _pAcc;

    // helper functions
    int _GetCurButton();
    LRESULT _NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction);
    BOOL _IsButtonHiddenOrDisabled(int i, DWORD dwFlags);
    TCHAR _GetButtonAccelerator(int i);
    void _ApplyOptions();
    HRESULT _InitShutdownObjects();
    HRESULT _AddShutdownChoiceToMenu(HMENU hMenu, SHUTDOWN_CHOICE sdChoice, UINT idMenuItem, UINT item);

    BOOL _SetTBButtons(int id, UINT iMsg);
    int _GetThemeBitmapSize(int iPartId, int iStateId, int id);

    int _GetIdmFromIdstip(int id);
	int _GetIdstipFromCommand(int id);
    int _GetIdmFromCommand(int id);
    int _GetCurPressedButton();
    int _HitTest(HWND hwndTest, POINT ptTest);
    BOOL _IsPtInDropDownSplit(HWND hwnd, POINT pt);
    void _DoSplitButtonContextMenu(int a2);
    BOOL _IsTBButtonEnabled(int a2);
    void _SetShutdownButtonProperties(int a2);
    int _GetLocalImageForShutdownChoice(SHUTDOWN_CHOICE sdChoice);

    friend BOOL CLogoffPane_RegisterClass();

    friend class CLogOffMenuCallback;
};

CLogoffPane::CLogoffPane()
{
    ASSERT(_hwndTB == NULL);
    ASSERT(_hwndTT == NULL);
    _clr = CLR_INVALID;
}

CLogoffPane::~CLogoffPane()
{
}

HRESULT CLogoffPane::QueryInterface(REFIID riid, void** ppvOut)
{
    static const QITAB qit[] =
    {
        QITABENT(CLogoffPane, IAccessible),
        QITABENT(CLogoffPane, IDispatch), // IAccessible derives from IDispatch
        QITABENT(CLogoffPane, IEnumVARIANT),
        {},
    };
    return QISearch(this, qit, riid, ppvOut);
}

LRESULT CALLBACK CLogoffPane::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CLogoffPane* self = reinterpret_cast<CLogoffPane*>(GetWindowLongPtrW(hwnd, 0));

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
    HBITMAP hBitmap = (HBITMAP)LoadImageW(
        _AtlBaseModule.GetModuleInstance(), MAKEINTRESOURCEW(id), IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
    if (hBitmap)
    {
        BITMAP bm;
        if (GetObjectW(hBitmap, sizeof(BITMAP), &bm))
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
        if (GetThemePartSize(_hTheme, nullptr, iPartId, iStateId, nullptr, TS_TRUE, &siz) >= 0)
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
    TBBUTTONINFOW tbbi;
    tbbi.dwMask = dwFlags | TBIF_STATE;
    tbbi.cbSize = sizeof(tbbi);
    SendMessageW(_hwndTB, TB_GETBUTTONINFOW, i, reinterpret_cast<LPARAM>(&tbbi));
    return (tbbi.fsState & TBSTATE_HIDDEN | TBSTATE_INDETERMINATE) != 0;
}

LRESULT CLogoffPane::_OnSize(int x, int y)
{
    if (_hwndSplit)
    {
        SetWindowPos(
            _hwndSplit, nullptr, _margins.cxLeftWidth + field_64, _margins.cyTopHeight, field_68,
            y - (_margins.cyBottomHeight - _margins.cyTopHeight),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER
        );
    }

    if (_hwndTB)
    {
        SetWindowPos(
            _hwndTB, nullptr, _margins.cxLeftWidth, _margins.cyTopHeight, field_64,
            y - (_margins.cyBottomHeight - _margins.cyTopHeight),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER
        );
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

const struct
{
    SHUTDOWN_CHOICE sdc;
    UINT uIdName;
    UINT uIdAccName;
    UINT uIdVerb;
    UINT uIdDesc;
    SHUTDOWN_CHOICE_ICON sdcIcon;
} g_rgShutdownResources[] =
{
    { SHTDN_SHUTDOWN,                                   3008, 3009, 3010, 3011, SHUTDOWN_CHOICE_ICON_1 },
    { SHTDN_SLEEP,                                      3016, 3017, 3018, 3019, SHUTDOWN_CHOICE_ICON_2 },
    { SHTDN_RESTART,                                    3012, 3013, 3014, 3015, SHUTDOWN_CHOICE_ICON_4 },
    { SHTDN_SLEEP | 0x200000,                           3026, 3027, 3028, 3029, SHUTDOWN_CHOICE_ICON_3 }, // Hybrid sleep (s4)
    { SHTDN_HIBERNATE,                                  3022, 3023, 3024, 3025, SHUTDOWN_CHOICE_ICON_2 },
    { SHTDN_SHUTDOWN | SHTDN_MODIFIER_INSTALL_UPDATES,  3030, 3031, 3032, 3033, SHUTDOWN_CHOICE_ICON_5 }
};

int CLogoffPane::_GetLocalImageForShutdownChoice(SHUTDOWN_CHOICE sdChoice)
{
    if (_psdc != nullptr)
    {
        SHUTDOWN_CHOICE_ICON sdChoiceIcon = SHUTDOWN_CHOICE_ICON_1;
        // _psdc->GetChoiceIcon(sdChoice, &sdChoiceIcon); // @NOTE: Nuked sometime after vista
        switch (sdChoiceIcon)
        {
            case SHUTDOWN_CHOICE_ICON_2:
            case SHUTDOWN_CHOICE_ICON_3:
                return 3;
            case SHUTDOWN_CHOICE_ICON_4:
                return 4;
            case SHUTDOWN_CHOICE_ICON_5:
                return 2;
        }
    }
    return 1;
}

void CLogoffPane::_SetShutdownButtonProperties(int a2)
{
    SHUTDOWN_CHOICE sdChoice;
    if (_psdc && SUCCEEDED(_psdc->GetDefaultChoice(&sdChoice)))
    {
        _ASSERT(sdChoice != SHTDN_NONE); // 780
        if ((sdChoice & SHTDN_MODIFIER_ACCESS_DENIED) != 0 && a2)
        {
            SendMessageW(_hwndTB, TB_SETSTATE, 1, TBSTATE_INDETERMINATE);
        }

        TBBUTTONINFOW tbbi = {};
        tbbi.cbSize = sizeof(tbbi);
        tbbi.dwMask = TBIF_IMAGE;
        tbbi.iImage = _GetLocalImageForShutdownChoice(sdChoice);

        (void)sdChoice; // Skipped telemetry StartMenu_Right_Control_Button_Label
        SendMessageW(_hwndTB, TB_SETBUTTONINFOW, 1, (LPARAM)&tbbi);
    }
}

extern BOOL _ShowStartMenuShutdown();
extern BOOL _ShowStartMenuDisconnect();
extern BOOL _AllowLockWorkStation();

void CLogoffPane::_ApplyOptions()
{
    if (_psdc)
        _psdc->Refresh();

    BOOL fShowShutdown = _ShowStartMenuShutdown();
    SendMessageW(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, fShowShutdown == 0);
    SendMessageW(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, _ShowStartMenuDisconnect() == 0);
    SendMessageW(_hwndTB, TB_HIDEBUTTON, SMNLC_LOGOFF, _AllowLockWorkStation() == 0);

    SIZE siz;
    if (SendMessageW(_hwndTB, TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&siz)))
    {
        field_64 = siz.cx;
    }
    // SHTracePerfSQMSetValueImpl(&ShellTraceId_StartMenu_Left_Control_Button_Label, 54, 5);
    _SetShutdownButtonProperties(fShowShutdown);
}

// @NOTE This is not publicly in CommCtrl.h for some reason.
// ref: https://learn.microsoft.com/en-us/windows/win32/controls/tbn-wrapaccelerator
typedef struct tagNMTBWRAPACCELERATOR
{
    NMHDR hdr;
    UINT ch;
    int iButton;
} NMTBWRAPACCELERATOR, *LPNMTBWRAPACCELERATOR;

LRESULT CLogoffPane::_OnNotify(NMHDR* pnm)
{
    NMHDR* v10;
    LRESULT v13;

    if (pnm->hwndFrom == _hwndTB)
    {
        switch (pnm->code)
        {
            case TBN_WRAPACCELERATOR:
                reinterpret_cast<NMTBWRAPACCELERATOR*>(pnm)->iButton = -1;
                break;
            case TBN_GETINFOTIPW:
                if (_GetCurPressedButton() == -1)
                {
                    NMTBGETINFOTIPW* ptbgit = reinterpret_cast<NMTBGETINFOTIPW*>(pnm);
                    if (ptbgit->lParam != 7032)
                    {
                        _ASSERT(ptbgit->lParam >= IDS_LOGOFF_TIP_EJECT/*7030*/ && ptbgit->lParam <= IDS_LOGOFF_TIP_LAST/*7036*/); // 879
                        if (ptbgit->lParam)
                        {
                            LoadStringW(g_hinstCabinet, ptbgit->lParam, ptbgit->pszText, ptbgit->cchTextMax);
                        }
                    }
                    else if (_psdc)
                    {
                        DWORD sdChoice;
                        if (SUCCEEDED(_psdc->GetDefaultChoice(&sdChoice)))
                        {
                            _psdc->GetChoiceDesc(sdChoice, ptbgit->pszText, ptbgit->cchTextMax);
                        }
                    }
                }
                break;
            case NM_CUSTOMDRAW:
                return _OnCustomDraw(reinterpret_cast<NMTBCUSTOMDRAW*>(pnm));
            default:
                return 0;
        }
        return 1;
    }

    if (pnm->hwndFrom == _hwndSdListenMsg)
    {
        _ApplyOptions();
        _SendNotify(_hwnd, 209, nullptr);
        return 0;
    }

    switch (pnm->code)
    {
        case SMN_REFRESHLOGOFF:
            _ApplyOptions();
            break;
        case 214:
            if (GetFocus() == _hwndTB)
            {
                v13 = SendMessageW(_hwndTB, TB_GETHOTITEM, 0, 0);
                NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndTB, OBJID_CLIENT, v13 + 1);
            }
            goto LABEL_29;
        case 215:
            return _OnSMNFindItem(reinterpret_cast<SMNDIALOGMESSAGE*>(pnm));
        case 221:
            if (_psdListen)
            {
                // Skipped telemetry ShellTraceId_Explorer_ShutdownUX_StartMenuCriticalPath_Start
                _psdListen->ScanForPassiveChange();
                // Skipped telemetry ShellTraceId_Explorer_ShutdownUX_StartMenuCriticalPath_Stop
            }
            return 0;
        case 224:
            if (_GetCurPressedButton() == 99)
            {
                TBBUTTONINFOW tbbi;
                tbbi.cbSize = sizeof(tbbi);
                tbbi.dwMask = TBIF_COMMAND;
                tbbi.idCommand = 0;

                v10 = (NMHDR*)SendMessageW(_hwndTB, TB_GETBUTTONINFO, 0, (LPARAM)&tbbi);
                pnm = v10;

                _fSettingHotItem = 1;
                SendMessageW(_hwndTB, TB_SETHOTITEM, (WPARAM)v10, 0);
                _fSettingHotItem = 0;

                SendMessageW(_hwndSplit, BM_SETSTATE, 0, 0);
                if (SetFocus(_hwndTB) != _hwndTB)
                {
                    NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndTB, OBJID_CLIENT, (LONG)&pnm->hwndFrom + 1);
                }
            }
            return 0;
        case 225:
            break;
        case NM_KILLFOCUS:
        LABEL_29:
            if (_fSplitButtonHot || _GetCurPressedButton() == 99)
            {
                _fSplitButtonHot = 0;
                SendMessageW(_hwndSplit, BM_SETSTATE, 0, 0);
                InvalidateRect(_hwndSplit, nullptr, 0);
            }
            return 0;
        default:
            return 0;
    }
    return 1;
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
    TBBUTTONINFOW tbbi = {};
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = TBIF_LPARAM;
    if (SendMessageW(_hwndTB, TB_GETBUTTONINFOW, id, reinterpret_cast<LPARAM>(&tbbi)) != -1)
    {
        return tbbi.lParam;
    }
    return -1;
}

int CLogoffPane::_GetIdmFromCommand(int id)
{
    return _GetIdmFromIdstip(_GetIdstipFromCommand(id));
}

int CLogoffPane::_GetCurPressedButton()
{
    if (SendMessageW(_hwndTB, TB_ISBUTTONPRESSED, 0, 0))
        return 0;
    if (SendMessageW(_hwndTB, TB_ISBUTTONPRESSED, 1, 0))
        return 1;
    if ((SendMessageW(_hwndSplit, BM_GETSTATE, 0, 0) & BST_PUSHED) != 0)
        return 99;
    return -1;
}

HRESULT CLogOffMenuCallback_CreateInstance(IShellMenuCallback** ppsmc, CLogoffPane* pLogOffPane);

void CLogoffPane::_DoSplitButtonContextMenu(int a2)
{
    SendMessageW(_hwndTB, TB_SETHOTITEM, -1, 0);

    IShellMenu* psm;
    if (SUCCEEDED(CoCreateInstance(CLSID_MenuBand, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&psm))))
    {
        IShellMenuCallback* psmc = nullptr;
        HRESULT hr = CLogOffMenuCallback_CreateInstance(&psmc, this);

        RECT rc;
        GetClientRect(_hwndSplit, &rc);
        MapWindowRect(_hwndSplit, nullptr, &rc);

        IMAGEINFO imageInfo;
        if (!_hTheme && ImageList_GetImageInfo(_himl, 0, &imageInfo))
        {
            rc.right = imageInfo.rcImage.right + rc.left;
            rc.bottom = imageInfo.rcImage.bottom + rc.top;
        }

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(psm->Initialize(psmc, 0, ANCESTORDEFAULT, 0x1 | SMINIT_RESTRICT_DRAGDROP | SMINIT_TOPLEVEL | 0x40 | SMINIT_VERTICAL)))
            {
                SendMessageW(_hwndSplit, BM_SETSTATE, TRUE, 0);
                (void)57; // Skipped telemetry ShellTraceId_StartMenu_Left_Control_Button_Split_Open

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
    // eax
    // [esp-10h] [ebp-3Ch]
    // [esp-4h] [ebp-30h]
    // [esp+28h] [ebp-4h]

    int v14 = 0;
    if (id == 4)
    {
        v5 = lParam;
        if (!lParam)
        {
            SendMessageW(_hwndTB, TB_PRESSBUTTON, 1u, 0);
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
            else if (HIWORD(wParam) == 7 && (_fSplitButtonHot || _GetCurPressedButton() == 99))
            {
                _fSplitButtonHot = 0;
                SendMessageW(_hwndSplit, BM_SETSTATE, 0, 0);
                InvalidateRect(_hwndSplit, NULL, 0);
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
            if (_psdc && _psdc->GetDefaultChoice((DWORD*)&lParam) >= 0)
            {
                v5 = lParam;
            }
            else
            {
                v5 = 2;
            }
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
                if (_fSplitButtonHot)
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
    if (pdis->hwndItem == _hwndSplit && (pdis->itemAction & 0x45) != 0)
    {
        RECT rc;
        rc.left = 0;
        rc.top = 0;
        rc.right = RECTWIDTH(pdis->rcItem);
        rc.bottom = HIWORD(SendMessageW(_hwndTB, TB_GETBUTTONSIZE, 0, 0));

        if (_hTheme == nullptr)
        {
            int iImage = (pdis->itemState & ODS_SELECTED) != 0 ? 1 : 0;

            SHFillRectClr(pdis->hDC, &pdis->rcItem, GetSysColor(COLOR_MENU));
            int iOldMode = SetBkMode(pdis->hDC, TRANSPARENT);

            IMAGEINFO imageInfo;
            if (ImageList_GetImageInfo(_himl, iImage, &imageInfo))
            {
                if (rc.right > imageInfo.rcImage.right)
                {
                    rc.right = imageInfo.rcImage.right;
                }
                if (_fSplitButtonHot)
                {
                    SHFillRectClr(pdis->hDC, &rc, GetSysColor(COLOR_HIGHLIGHT));
                }
                ImageList_DrawEx(
                    _himl, iImage, pdis->hDC, rc.left, rc.top, 0, RECTHEIGHT(rc), CLR_NONE, CLR_NONE, 0);
            }

            HFONT hfMarlett = (HFONT)SelectObject(pdis->hDC, _hfMarlett);
            if (hfMarlett)
            {
                BOOL fRTL = (GetLayout(pdis->hDC) & LAYOUT_RTL) != 0;

                UINT dtFlags = DT_CENTER;
                rc.top += (rc.bottom - field_58 - rc.top) / 2;
                if (fRTL)
                {
                    dtFlags |= DT_RTLREADING;
                }

                WCHAR chOut = fRTL ? 'w' : '8';
                DrawTextW(pdis->hDC, &chOut, 1, &rc, dtFlags);
                SetTextColor(
                    pdis->hDC, SetTextColor(pdis->hDC, GetSysColor(_fSplitButtonHot != 0 ? COLOR_HIGHLIGHTTEXT : COLOR_MENUTEXT)));
                SelectObject(pdis->hDC, hfMarlett);
            }
            if (iOldMode)
            {
                SetBkMode(pdis->hDC, iOldMode);
            }
        }
        else
        {
            int iState = _fSplitButtonHot != 0 ? SPLS_HOT : SPLS_NORMAL;
            if ((pdis->itemState & ODS_SELECTED) != 0)
            {
                iState = SPLS_PRESSED;
            }
            DrawThemeParentBackground(_hwndSplit, pdis->hDC, &pdis->rcItem);
            DrawThemeBackground(_hTheme, pdis->hDC, SPP_LOGOFFSPLITBUTTONDROPDOWN, iState, &rc, nullptr);
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
    if (_fSplitButtonHot)
    {
        return 99;
    }
    return static_cast<int>(SendMessageW(_hwndTB, TB_GETHOTITEM, 0, 0));
}

BOOL CLogoffPane::_IsPtInDropDownSplit(HWND hwnd, POINT pt)
{
    return _HitTest(hwnd, pt) == 99;
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
                if (IsPtInDropDownSplit != this->_fSplitButtonHot)
                {
                    this->_fSplitButtonHot = IsPtInDropDownSplit;
                    if (!IsPtInDropDownSplit)
                        SendMessageW(this->_hwndSplit, 0xF3u, 0, 0);
                    goto LABEL_7;
                }
            }
            else
            {
                pdm->flags &= ~0x100u;
                if (this->_fSplitButtonHot)
                {
                    this->_fSplitButtonHot = 0;
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
        if (this->_fSplitButtonHot)
        {
            v14 = this->_hwndSplit;
            this->_fSplitButtonHot = 0;
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
    // eax
    DWORD v5; // [esp+Ch] [ebp-D0h] BYREF
    OLECHAR psz[100]; // [esp+10h] [ebp-CCh] BYREF

    if (!varChild.lVal
        || SendMessageW(_hwndTB, TB_COMMANDTOINDEX, 1u, 0) != varChild.lVal - 1
        || !_psdc
        || _psdc->GetDefaultChoice(&v5) < 0
        || _psdc->GetChoiceName(v5, 0, psz, 100) < 0)
    {
        return CAccessible::get_accName(varChild, pszName);
    }
    OLECHAR* v3 = SysAllocString(psz);
    *pszName = v3;
    return v3 != nullptr ? 0 : 0x8007000E;
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

    HRESULT hr = CoCreateInstance(CLSID_AuthUIShutdownChoices, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_psdc));
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

#if 0
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
#endif
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
    if (_ShowStartMenuShutdown() && _psdc)
    {
        MENUITEMINFOW mi = {};
        mi.cbSize = sizeof(mi);
        mi.fType = MFT_SEPARATOR;
        InsertMenuItemW(hMenu, 0xFFFF, 0x400, &mi);

        IEnumShutdownChoices* pesc;
        // HRESULT hr = _psdc->GetMenuChoices(&pesc);
        HRESULT hr = _psdc->GetChoiceEnumerator(&pesc);
        if (SUCCEEDED(hr))
        {
            SHUTDOWN_CHOICE sdChoice;
            for (UINT i = 100; pesc->Next(1, &sdChoice, nullptr) == S_OK; ++i)
            {
                if ((sdChoice & 0x400000) != 0)
                {
                    InsertMenuItemW(hMenu, 0xFFFF, 0x400, &mi);
                }
                else
                {
                    hr = _AddShutdownChoiceToMenu(hMenu, sdChoice, i, 0xFFFF);
                }
                if (FAILED(hr))
                    break;
            }
            pesc->Release();
        }
    }
}

HRESULT CLogoffPane::_AddShutdownChoiceToMenu(HMENU hMenu, DWORD sdChoice, UINT idMenuItem, UINT item)
{
    _ASSERT(idMenuItem <= IDM_SHUTDOWN_LAST /*150*/); // 1074
    _ASSERT(nullptr != _psdc); // 1075

    WCHAR szChoiceName[200];
    HRESULT hr = _psdc->GetChoiceName(sdChoice, 1, szChoiceName, ARRAYSIZE(szChoiceName));
    if (SUCCEEDED(hr))
    {
        MENUITEMINFOW mii = {};
        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_STATE | MIIM_ID | MIIM_TYPE | MIIM_DATA;
        mii.fType = MFT_STRING;
        mii.dwTypeData = szChoiceName;
        mii.wID = idMenuItem;
        mii.dwItemData = sdChoice;
        mii.fState = (sdChoice & (SHTDN_MODIFIER_ACCESS_DENIED | SHTDN_MODIFIER_BROKEN)) != 0
                         ? MFS_GRAYED
                         : MFS_ENABLED;
        if (!InsertMenuItemW(hMenu, item, 0x400, &mii))
        {
            hr = HRESULT_FROM_WIN32(GetLastError());
        }
    }

    return hr;
}

extern BOOL _ShowStartMenuEject();
extern BOOL _ShowStartMenuLogoff();

DEFINE_GUID(POLID_HideFastUserSwitching, 0x572FD217, 0xF7FF, 0x479C, 0x8D, 0x96, 0xBC, 0x93, 0x8D, 0x68, 0x67, 0xF5);

void CLogoffPane::ApplyLogoffMenuOption(HMENU hMenu)
{
    if (!_AllowLockWorkStation())
    {
        DeleteMenu(hMenu, IDM_LOCKWORKSTATION, 0);
        DeleteMenu(hMenu, 5000, 0);
    }
    if (!IsOS(OS_FASTUSERSWITCHING) || GetSystemMetrics(SM_REMOTESESSION) || GetSystemMetrics(SM_REMOTECONTROL)
        || SHWindowsPolicy(POLID_HideFastUserSwitching))
    {
        DeleteMenu(hMenu, 5000, 0);
    }
    if (!_ShowStartMenuLogoff())
    {
        DeleteMenu(hMenu, IDM_LOGOFF, 0);
    }
    if (!_ShowStartMenuEject())
    {
        DeleteMenu(hMenu, IDM_EJECTPC, 0);
    }
}

HRESULT CLogoffPane::GetShutdownItemDescription(SHUTDOWN_CHOICE sdChoice, WCHAR* pszDesc, UINT cchDesc)
{
    ASSERT(nullptr != _psdc); // 1106
    return _psdc->GetChoiceDesc(sdChoice, pszDesc, cchDesc);
}

int CLogoffPane::GetTipIDFromIDM(int id)
{
    int v2 = 0;
    unsigned int v3 = 0;
    while (c_rgButtons[v3].idCmd != id)
    {
        ++v3;
        ++v2;
        if (v3 >= 6)
        {
            return -1;
        }
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

HRESULT CALLBACK CLogOffMenuCallback_CreateInstance(IShellMenuCallback** ppsmc, CLogoffPane* pLogoffPane)
{
    CLogOffMenuCallback* plmc = new(std::nothrow) CLogOffMenuCallback(pLogoffPane);
    *ppsmc = static_cast<IShellMenuCallback*>(plmc);
    return *ppsmc != nullptr ? S_OK : E_OUTOFMEMORY;
}
