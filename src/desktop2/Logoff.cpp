#include "pch.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "shundoc.h"

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
    STDMETHOD(GetDefaultChoice)(DWORD *piOut) PURE; // This returns an ID representing the choice.
    STDMETHOD(GetMenuChoices)(IEnumShutdownChoices **ppOut) PURE;
    STDMETHOD_(BOOL, UserHasShutdownRights)() PURE;
    STDMETHOD(GetChoiceName)(DWORD choice, int a3, LPWSTR a4, UINT a5) PURE;
    STDMETHOD(GetChoiceDesc)(DWORD, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceVerb)(DWORD, LPWSTR, UINT) PURE;
    STDMETHOD(GetChoiceIcon)(DWORD choice, enum SHUTDOWN_CHOICE_ICON *pIcon) PURE;
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
    BOOL _IsButtonHiddenOrDisabled(int i, DWORD dwFlags);
    TCHAR _GetButtonAccelerator(int i);
    void _RightAlign();
    void _ApplyOptions();
    HRESULT _InitShutdownObjects();

    BOOL _SetTBButtons(int id, UINT iMsg);
    int _GetThemeBitmapSize(int iPartId, int iStateId, int id);

    int _GetIdmFromIdstip(int id);
	int _GetIdstipFromCommand(int id);
    int _GetIdmFromCommand(int id);

    int _GetCurPressedButton();

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

int __stdcall DrawPlacesListBackground(HTHEME hTheme, HWND hWnd, HDC hdc);

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
        return self->_OnCommand(GET_WM_COMMAND_ID(wParam), wParam, lParam);

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

// EXEX-TODO: Move? Currently in TaskBand.cpp
extern BOOL g_fHighDPI;
extern void InitDPI();

BOOL WINAPI IsHighDPI()
{
    InitDPI();
    return g_fHighDPI;
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
    TBBUTTONINFOW tbbi;
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

DEFINE_GUID(POLID_NoClose, 0x29B0CC43, 0x2F2B, 0x4D0C, 0xA0, 0x81, 0xA5, 0x28, 0xDD, 0x34, 0x96, 0x31);

BOOL WINAPI _ShowStartMenuShutdownVista()
{
    return !SHWindowsPolicy(POLID_NoClose) && (IsOS(OS_ANYSERVER) || !GetSystemMetrics(SM_REMOTESESSION));
}

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

    // SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, _ShowStartMenuShutdownVista() == 0);
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, _ShowStartMenuDisconnect() == 0);
    // SendMessage(_hwndTB, TB_HIDEBUTTON, 0, _AllowLockWorkStation() == 0);

    SIZE siz; 
    if (SendMessage(_hwndTB, TB_GETMAXSIZE, 0, (LPARAM)&siz))
        field_64 = siz.cx;

    // SHTracePerfSQMSetValueImpl(&ShellTraceId_StartMenu_Left_Control_Button_Label, 54, 5);
    // _SetShutdownButtonProperties(_ShowStartMenuShutdownVista());
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
                        DWORD dwChoice;
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
        _SendNotify(this->_hwnd, 0xD1u, 0);
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
    case 0xE0u:
        if (_GetCurPressedButton() == 99)
        {
            TBBUTTONINFOW tbbi;
            tbbi.cbSize = 32;
            tbbi.dwMask = 32;
            tbbi.idCommand = 0;
            v10 = (NMHDR*)SendMessageW(this->_hwndTB, TB_GETBUTTONINFOW, 0, (LPARAM)&tbbi);
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
    case 0xE1u:
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

LRESULT CLogoffPane::_OnCommand(int id, WPARAM wParam, LPARAM lParam)
{
    LPARAM v5; // edi
    WPARAM IdmFromCommand; // ebx
    int v7; // eax
    IShutdownChoices* psdc; // eax
    // eax
    HWND hwndSplit; // [esp-10h] [ebp-3Ch]
    // [esp-4h] [ebp-30h]
    SMNMCOMMANDINVOKED ci; // [esp+Ch] [ebp-20h] BYREF
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
                // _DoSplitButtonContextMenu(0);
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
            if (_psdc && _psdc->GetDefaultChoice((DWORD*)&lParam) >= 0)
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
            SendMessageW(this->_hwndTB, TB_GETITEMRECT, IdmFromCommand, (LPARAM)&ci.rcItem);
            MapWindowPoints(this->_hwndTB, NULL, (LPPOINT)&ci.rcItem, 2u);
            _SendNotify(this->_hwnd, 204u, &ci.hdr);
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
#if 0
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
#endif
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
