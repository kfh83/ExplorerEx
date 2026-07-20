#include "pch.h"

#include "AuthUI.h"
#include "DPIHelpers.h"
#include "SHUndoc.h"
#include "stdafx.h"
#include "SFTHost.h"

class CLogoffPane;

#define NUM_TBBUTTON_IMAGES 6

static const TBBUTTON tbButtonsCreate[] =
{
    { 1,    SMNLC_TURNOFF,     TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, IDS_LOGOFF_TIP_SHUTDOWN,     1 },
    { 0,    SMNLC_DISCONNECT,  TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, IDS_LOGOFF_TIP_DISCONNECT,   2 },
    { 5,    SMNLC_LOGOFF,      TBSTATE_ENABLED, BTNS_BUTTON, { 0, 0 }, IDS_LOGOFF_TIP_LOCK,         0 }
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
    STDMETHODIMP get_accName(VARIANT varChild, BSTR* pszName) override;
    STDMETHODIMP get_accKeyboardShortcut(VARIANT varChild, BSTR* pszKeyboardShortcut) override;
    STDMETHODIMP get_accDefaultAction(VARIANT varChild, BSTR* pszDefAction) override;

    CLogoffPane();
    ~CLogoffPane() override;

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCreate(LPARAM lParam);
    void _OnDestroy();
    LRESULT _OnNCCreate(HWND hwnd,  UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNotify(NMHDR *pnm);
    LRESULT _OnCommand(int id, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCustomDraw(NMTBCUSTOMDRAW *pnmtbcd);
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
    COLORREF _colorHighlight;
    COLORREF _colorHighlightText;
    HTHEME _hTheme;
    BOOL _fSettingHotItem;
    int field_40;
    MARGINS _margins;
    HFONT _hfMarlett;
    int _tmAscentMarlett;
    BOOL _fSplitButtonHot;
    HIMAGELIST _himl;
    int _cxToolbar;
    UINT _cxSplitButton;
    IShutdownChoices* _psdc;
    IShutdownChoiceListener* _psdListen;
    HWND _hwndSdListenMsg;
    IAccessible* _pAcc;

    // helper functions
    int _GetCurButton();
    int _GetCurPressedButton();
    int _GetThemeBitmapSize(int iPartId, int iStateId, int id);
    LRESULT _NextVisibleButton(PSMNDIALOGMESSAGE pdm, int i, int direction);
    BOOL _IsButtonHiddenOrDisabled(int i, DWORD dwFlags);
    TCHAR _GetButtonAccelerator(int i);

    void _ApplyOptions();
    HRESULT _InitShutdownObjects();
    HRESULT _AddShutdownChoiceToMenu(HMENU hMenu, SHUTDOWN_CHOICE sdChoice, UINT idMenuItem, UINT item);

    BOOL _SetTBButtons(int id, UINT iMsg);

    int _GetIdmFromIdstip(int id);
	int _GetIdstipFromCommand(int id);
    int _GetIdmFromCommand(int id);

    int _HitTest(HWND hwndTest, POINT ptTest);
    BOOL _IsPtInDropDownSplit(HWND hwnd, POINT pt);
    void _DoSplitButtonContextMenu(int a2);
    BOOL _IsTBButtonEnabled(int i);
    void _SetShutdownButtonProperties(int a2);
    int _GetLocalImageForShutdownChoice(SHUTDOWN_CHOICE sdChoice);

    friend BOOL CLogoffPane_RegisterClass();
    friend class CLogOffMenuCallback;
};

CLogoffPane::CLogoffPane()
{
    ASSERT(_hwndTB == nullptr); // 313
    ASSERT(_hwndTT == nullptr); // 314
    _clr = CLR_INVALID;
    ASSERT(_psdc == nullptr);   // 316
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
    CLogoffPane* self = static_cast<CLogoffPane*>(GetWindowPtr0(hwnd));

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
    CLogoffPane* self = new CLogoffPane;
    if (self)
    {
        SetWindowPtr0(hwnd, self);
        self->_hwnd = hwnd;
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
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

class CSplitButtonAccessible : public CAccessible
{
public:
    CSplitButtonAccessible(HWND hwnd)
        : _cRef(1)
        , _hwnd(hwnd)
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

    int nResId;
    if (_hTheme)
        nResId = IsHighDPI() ? IDB_LOGOFF_LARGE_AERO_NORMAL : IDB_LOGOFF_AERO_NORMAL;
    else
        nResId = IsHighDPI() ? IDB_LOGOFF_LARGE_NORMAL : IDB_LOGOFF_NORMAL;
    _cxToolbar = _GetThemeBitmapSize(0, 0, nResId);
    if (!_cxToolbar)
        return -1;

    _cxSplitButton = _GetThemeBitmapSize(
        SPP_LOGOFFSPLITBUTTONDROPDOWN, 0, IsHighDPI() ? IDB_LOGOFF_LARGE_EXPANDER : IDB_LOGOFF_EXPANDER);
    if (!_cxSplitButton)
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
        _cxSplitButton /= 2;
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

    GetClientRect(_hwnd, &rc);
    rc.left += _margins.cxLeftWidth;
    rc.top += _margins.cyTopHeight;
    rc.right = rc.left + _cxToolbar;
    rc.bottom -= _margins.cyBottomHeight;

    _hwndTB = SHFusionCreateWindowEx(0, TOOLBARCLASSNAME, NULL,
        0x4 | 0x40 | 0x100 | 0x800 | 0x4000000 | 0x10000000 | 0x40000000,
        rc.left, rc.top, _cxToolbar, RECTHEIGHT(rc), _hwnd,
        NULL, NULL, NULL);
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
                        _tmAscentMarlett = tm.tmAscent;
                    }
					SelectObject(hdc, hfMarlett);
				}
            }
			ReleaseDC(_hwnd, hdc);
        }

        rc.left = rc.right;
        rc.right += _cxSplitButton;

		TCHAR szTitle[200] = {};
        LoadString(g_hinstCabinet, IDS_STARTPANE_TITLE_SPLITTER, szTitle, ARRAYSIZE(szTitle));

        _hwndSplit = SHFusionCreateWindowEx(
            0, WC_BUTTON, szTitle, 0x1 | 0x2 | 0x8 | 0x2000000 | 0x4000000 | 0x10000000 | 0x40000000, rc.left, rc.top,
            RECTWIDTH(rc), RECTHEIGHT(rc), _hwnd, (HMENU)99, g_hinstCabinet, nullptr);
        if (_hwndSplit)
        {
            CSplitButtonAccessible* pSplitAcc = new CSplitButtonAccessible(_hwndSplit);
            if (pSplitAcc)
            {
                SetAccessibleSubclassWindow(_hwndSplit);
                QueryInterface(IID_PPV_ARGS(&_pAcc));
                pSplitAcc->Release();
            }
        }

        return 0;
    }

    return -1;
}

BOOL CLogoffPane::_IsButtonHiddenOrDisabled(int i, DWORD dwFlags)
{
    TBBUTTONINFO tbbi;
    tbbi.dwMask = dwFlags | TBIF_STATE;
    tbbi.cbSize = sizeof(tbbi);
    SendMessage(_hwndTB, TB_GETBUTTONINFO, i, reinterpret_cast<LPARAM>(&tbbi));
    return (tbbi.fsState & TBSTATE_HIDDEN | TBSTATE_INDETERMINATE) != 0;
}

LRESULT CLogoffPane::_OnSize(int x, int y)
{
    if (_hwndSplit)
    {
        SetWindowPos(
            _hwndSplit, nullptr, _margins.cxLeftWidth + _cxToolbar, _margins.cyTopHeight, _cxSplitButton,
            y - (_margins.cyBottomHeight - _margins.cyTopHeight),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER
        );
    }
    if (_hwndTB)
    {
        SetWindowPos(
            _hwndTB, nullptr, _margins.cxLeftWidth, _margins.cyTopHeight, _cxToolbar,
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

// Vista AuthUI.dll shutdown option struct for future reference.
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
    if (_psdc)
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
        ASSERT(sdChoice != SHTDN_NONE); // 780
        if (sdChoice & SHTDN_MODIFIER_ACCESS_DENIED && a2)
        {
            SendMessage(_hwndTB, TB_SETSTATE, 1, TBSTATE_INDETERMINATE);
        }

        TBBUTTONINFO tbbi;
        ZeroMemory(&tbbi, sizeof(tbbi));

        tbbi.cbSize = sizeof(tbbi);
        tbbi.dwMask = TBIF_IMAGE;
        tbbi.iImage = _GetLocalImageForShutdownChoice(sdChoice);

        (void)sdChoice; // Skipped telemetry StartMenu_Right_Control_Button_Label
        SendMessage(_hwndTB, TB_SETBUTTONINFO, 1, (LPARAM)&tbbi);
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
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_TURNOFF, fShowShutdown == 0);
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_DISCONNECT, _ShowStartMenuDisconnect() == 0);
    SendMessage(_hwndTB, TB_HIDEBUTTON, SMNLC_LOGOFF, _AllowLockWorkStation() == 0);

    SIZE siz;
    if (SendMessage(_hwndTB, TB_GETMAXSIZE, 0, (LPARAM)&siz))
    {
        _cxToolbar = siz.cx;
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
                    NMTBGETINFOTIP* ptbgit = reinterpret_cast<NMTBGETINFOTIP*>(pnm);
                    if (ptbgit->lParam != IDS_LOGOFF_TIP_SHUTDOWN)
                    {
                        ASSERT(ptbgit->lParam >= IDS_LOGOFF_TIP_EJECT && ptbgit->lParam <= IDS_LOGOFF_TIP_LAST); // 879
                        if (ptbgit->lParam)
                        {
                            LoadString(g_hinstCabinet, ptbgit->lParam, ptbgit->pszText, ptbgit->cchTextMax);
                        }
                    }
                    else if (_psdc)
                    {
                        SHUTDOWN_CHOICE sdChoice;
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
                v13 = SendMessage(_hwndTB, TB_GETHOTITEM, 0, 0);
                NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndTB, OBJID_CLIENT, v13 + 1);
            }
            goto L_SET_FOCUS;
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
        L_SET_FOCUS:
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
    { 0, 5000, IDS_LOGOFF_TIP_SWITCH },
    { 0, 410, IDS_LOGOFF_TIP_EJECT },
    { 0, 402, IDS_LOGOFF_TIP_LOGOFF },
    { 0, 517, IDS_LOGOFF_TIP_LOCK },
    { 1, 506, IDS_LOGOFF_TIP_SHUTDOWN },
    { 2, 5000, IDS_LOGOFF_TIP_DISCONNECT },
};

int CLogoffPane::_GetIdmFromIdstip(int id)
{
    UINT i = 0;
    UINT i_p = 0;

    while (true)
    {
        if (c_rgButtons[i_p].idsTip != id)
        {
            ++i_p;
            ++i;
            if (i_p >= 6)
            {
                return -1;
            }
            continue;
        }
        break;
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

HRESULT CLogOffMenuCallback_CreateInstance(IShellMenuCallback** ppsmc, CLogoffPane* pLogOffPane);

void CLogoffPane::_DoSplitButtonContextMenu(int a2)
{
    SendMessage(_hwndTB, TB_SETHOTITEM, -1, 0);

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
                SendMessage(_hwndSplit, BM_SETSTATE, TRUE, 0);
                (void)57; // Skipped telemetry ShellTraceId_StartMenu_Left_Control_Button_Split_Open

                DWORD dwFlags = 0;
                if (a2)
                {
                    dwFlags |= (MPPF_KEYBOARD | MPPF_FINALSELECT);
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

            if (pnmtbcd->nmcd.dwItemSpec == 0 && (pnmtbcd->nmcd.uItemState & CDIS_HOT) != 0 && _fSplitButtonHot)
            {
                pnmtbcd->nmcd.uItemState &= ~CDIS_HOT;
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
        rc.bottom = HIWORD(SendMessage(_hwndTB, TB_GETBUTTONSIZE, 0, 0));

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
                ImageList_DrawEx(_himl, iImage, pdis->hDC, rc.left, rc.top, 0, RECTHEIGHT(rc), CLR_NONE, CLR_NONE, 0);
            }

            HFONT hfMarlett = (HFONT)SelectObject(pdis->hDC, _hfMarlett);
            if (hfMarlett)
            {
                BOOL fRTL = (GetLayout(pdis->hDC) & LAYOUT_RTL) != 0;

                UINT dtFlags = DT_CENTER;
                rc.top += (rc.bottom - _tmAscentMarlett - rc.top) / 2;
                if (fRTL)
                {
                    dtFlags |= DT_RTLREADING;
                }

                WCHAR chOut = fRTL ? 'w' : '8';
                DrawText(pdis->hDC, &chOut, 1, &rc, dtFlags);

                COLORREF crText = GetSysColor(_fSplitButtonHot ? COLOR_HIGHLIGHTTEXT : COLOR_MENUTEXT);
                SetTextColor(pdis->hDC, SetTextColor(pdis->hDC, crText));
                SelectObject(pdis->hDC, hfMarlett);
            }
            if (iOldMode)
            {
                SetBkMode(pdis->hDC, iOldMode);
            }
        }
        else
        {
            int iState = _fSplitButtonHot ? SPLS_HOT : SPLS_NORMAL;
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
    ASSERT(direction == +1 || direction == -1); // 1400

    if (i == 99) i = 3;
    i += direction;
    while (i >= 0 && i < ARRAYSIZE(tbButtonsCreate))
    {
        if (!_IsButtonHiddenOrDisabled(i, TBIF_BYINDEX))
        {
            pdm->itemID = i;
            return TRUE;
        }
        i += direction;
    }
    if (i == 3)
    {
        pdm->itemID = 99;
        return TRUE;
    }
    return FALSE;
}

int CLogoffPane::_GetCurButton()
{
    if (_fSplitButtonHot)
    {
        return 99;
    }
    return static_cast<int>(SendMessage(_hwndTB, TB_GETHOTITEM, 0, 0));
}

int CLogoffPane::_GetCurPressedButton()
{
    if (SendMessage(_hwndTB, TB_ISBUTTONPRESSED, 0, 0))
        return 0;
    if (SendMessage(_hwndTB, TB_ISBUTTONPRESSED, 1, 0))
        return 1;
    if ((SendMessage(_hwndSplit, BM_GETSTATE, 0, 0) & BST_PUSHED) != 0)
        return 99;
    return -1;
}

int CLogoffPane::_GetThemeBitmapSize(int iPartId, int iStateId, int id)
{
    int cxBitmap = 0;

    if (_hTheme && iPartId)
    {
        SIZE siz;
        if (SUCCEEDED(GetThemePartSize(_hTheme, NULL, iPartId, iStateId, NULL, TS_TRUE, &siz)))
        {
            cxBitmap = siz.cx;
        }
    }
    else
    {
        HBITMAP hBitmap = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(id));
        if (hBitmap)
        {
            BITMAP bm;
            if (GetObject(hBitmap, sizeof(BITMAP), &bm))
            {
                cxBitmap = bm.bmWidth;
            }
            DeleteObject(hBitmap);
        }
    }

    SHLogicalToPhysicalDPI(&cxBitmap, NULL);
    return cxBitmap;
}

BOOL CLogoffPane::_IsPtInDropDownSplit(HWND hwnd, POINT pt)
{
    return _HitTest(hwnd, pt) == 99;
}

LRESULT CLogoffPane::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
{
    LRESULT lres = _OnSMNFindItemWorker(pdm);
    if (lres)
    {
        if ((pdm->flags & (SMNDM_SELECT | 0x800)) != 0)
        {
            int iButton = _GetCurPressedButton();
            if (iButton == -1 || iButton == 99)
            {
                pdm->flags |= SMNDM_HORIZONTAL;

                HWND hwnd;
                if (pdm->itemID == 99)
                    hwnd = _hwndSplit;
                else
                    hwnd = _hwndTB;
                pdm->hwnd2 = hwnd;

                if (_GetCurButton() != pdm->itemID)
                {
                    if (_hwndTT)
                    {
                        SendMessage(_hwndTT, TTM_POP, 0, 0);
                    }

                    int iHotItem = pdm->itemID;
                    if (iHotItem == 99)
                    {
                        iHotItem = -1;
                    }
                    else
                    {
                        _fSettingHotItem = TRUE;
                        SendMessage(_hwndTB, TB_SETHOTITEM, iHotItem, 0);
                        _fSettingHotItem = FALSE;
                    }

                    if ((pdm->flags & SMNDM_SELECT) != 0 && iHotItem != -1)
                    {
                        SendMessage(_hwndSplit, BM_SETSTATE, FALSE, 0);
                        if (SetFocus(_hwndTB) != _hwndTB)
                        {
                            NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndTB, OBJID_CLIENT, pdm->itemID + 1);
                        }
                    }
                }

                BOOL fIsInSplitButton = _IsPtInDropDownSplit(pdm->hwnd, pdm->pt);
                if (fIsInSplitButton != _fSplitButtonHot)
                {
                    _fSplitButtonHot = fIsInSplitButton;
                    if (!fIsInSplitButton)
                    {
                        SendMessage(_hwndSplit, BM_SETSTATE, FALSE, 0);
                    }
                    InvalidateRect(_hwndSplit, NULL, FALSE);
                    return lres;
                }
            }
            else
            {
                pdm->flags &= ~SMNDM_SELECT;

                if (_fSplitButtonHot)
                {
                    _fSplitButtonHot = 0;
                    InvalidateRect(_hwndSplit, NULL, FALSE);
                }
            }
        }
    }
    else
    {
        pdm->flags |= SMNDM_VERTICAL;

        int i = _GetCurButton();
        RECT rc;
        if (i >= 0 && SendMessage(_hwndTB, TB_GETITEMRECT, i, (LPARAM)&rc))
        {
            pdm->pt.x = (rc.left + rc.right) / 2;
            pdm->pt.y = (rc.top + rc.bottom) / 2;
        }
        else
        {
            pdm->pt.x = 0;
            pdm->pt.y = 0;
        }

        if (_fSplitButtonHot)
        {
            _fSplitButtonHot = FALSE;
            SendMessage(_hwndSplit, BM_SETSTATE, FALSE, 0);
            InvalidateRect(_hwndSplit, NULL, FALSE);
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
    if (hwndTest)
        MapWindowPoints(hwndTest, _hwnd, &ptTest, 1);

    HWND v4 = ChildWindowFromPointEx(_hwnd, ptTest, CWP_SKIPINVISIBLE);
    if (v4 != _hwndTB)
        return v4 != _hwndSplit ? -1 : 99;

    MapWindowPoints(_hwnd, _hwndTB, &ptTest, 1);
    return SendMessageW(_hwndTB, TB_HITTEST, 0, (LPARAM)&ptTest);
}

BOOL CLogoffPane::_IsTBButtonEnabled(int i)
{
    TBBUTTONINFO tbbi;
    tbbi.cbSize = sizeof(tbbi);
    tbbi.dwMask = 0x80000004;
    SendMessage(_hwndTB, TB_GETBUTTONINFO, i, (LPARAM)&tbbi);
    return tbbi.fsState & TBSTATE_ENABLED;
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
    SMNDIALOGMESSAGE *pnmdma; // [esp+30h] [ebp+8h]

    flags = pdm->flags;
    pnmdma = (SMNDIALOGMESSAGE *)(flags & 0xF);

    switch (flags & 0xF)
    {
    case 0u:
    case 1u:
    case 0xBu:
        return 0;
    case 2u:
        if ((flags & 0x10000) == 0)
            return _NextVisibleButton(pdm, -1, 1);
        return _NextVisibleButton(pdm, 3, -1);
    case 3u:
        return _NextVisibleButton(pdm, -1, 1);
    case 4u:
        VisibleButton = _NextVisibleButton(pdm, 3, -1);
        if (VisibleButton)
        {
            if (pdm->itemID == 99)
                _DoSplitButtonContextMenu(1);
        }
        return VisibleButton;
    case 5u:
        if (pdm->pmsg->wParam == 37)
        {
            CurButton = _GetCurButton();
            if (!_NextVisibleButton(pdm, CurButton, -1))
                return 0;
            goto LABEL_21;
        }
        if (pdm->pmsg->wParam == 39)
        {
            v9 = _GetCurButton();
            if (_NextVisibleButton(pdm, v9, 1))
            {
                if (pdm->itemID == 99)
                {
                    _DoSplitButtonContextMenu(1);
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
            if (_GetCurPressedButton() != 99)
                PostMessageW(this->_hwnd, WM_COMMAND, GET_WM_COMMAND_MPS(99, _hwndSplit, BN_HILITE));
            return 1;
        }
        if (v11 <= 2 && _IsTBButtonEnabled(v11))
        {
            if (pnmdma == (SMNDIALOGMESSAGE *)6)
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
                if (pmsg->message == 0x201 && _GetCurPressedButton() != 99)
                    PostMessageW(this->_hwnd, WM_COMMAND, GET_WM_COMMAND_MPS(99, _hwndSplit, BN_HILITE));
            }
        }
        return pdm->itemID >= 0;
    case 9u:
        pdm->flags = flags & 0xFFFFFEFF;
        return 1;
    case 0xAu:
        if (_GetCurButton() != 99)
            return 0;
        _DoSplitButtonContextMenu(1);
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
    CLogOffMenuCallback(CLogoffPane *pLogOffPane);

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

    CLogoffPane* _pLogOffPane;
};

CLogOffMenuCallback::CLogOffMenuCallback(CLogoffPane* pLogOffPane)
    : _pLogOffPane(pLogOffPane)
{
    if (_pLogOffPane)
    {
        _pLogOffPane->AddRef();
    }
}

CLogOffMenuCallback::~CLogOffMenuCallback()
{
    IUnknown_SafeReleaseAndNullPtr(_pLogOffPane);
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
    switch (uMsg)
    {
        case SMC_INITMENU:
        {
            IShellMenu* psm;
            if (psmd->punk && SUCCEEDED(psmd->punk->QueryInterface(IID_PPV_ARGS(&psm))))
            {
                HMENU hMenu = SHLoadMenuPopup(g_hinstCabinet, 6001);
                _pLogOffPane->ApplyLogoffMenuOption(hMenu);
                _pLogOffPane->AddShutdownOptions(hMenu);

                psm->SetMenu(hMenu, _pLogOffPane->_hwnd, 0);
                psm->Release();
                return S_OK;
            }
            break;
        }
        case SMC_EXITMENU:
        {
            _pLogOffPane->TBPressButton(0, 0);
            return S_OK;
        }
        case 4u:
        {
            if (psmd->uId < 100 || psmd->uId > 150)
            {
                UINT uId = psmd->uId | 0x20000000;
                DWORD dwTickCount = GetTickCount();
                // SHTracePerfSQMStreamTwoImpl(&ShellTraceId_StartMenu_Logoff_Usage_Stream, 571, dwTickCount, uId);
                PostMessage(v_hwndTray, WM_COMMAND, psmd->uId, 0);
            }
            else
            {
                DWORD dwChoice = _ShutdownChoiceFromMenuChoice(psmd);
                // SHTracePerf(&ShellTraceId_Explorer_ShutdownUX_SelectMenuItem_Start);
                PostMessage(_pLogOffPane->_hwnd, WM_COMMAND, 4, dwChoice | 0x10000000);
            }
            return S_OK;
        }
        case SMC_GETINFO:
        {
            PSMINFO psmi = (PSMINFO)lParam;
            if (psmi->dwMask & SMIM_ICON)
            {
                psmi->iIcon = -1;
            }
            return S_OK;
        }
        case 0xD:
        {
            if (psmd->uId >= 100 && psmd->uId <= 150)
            {
                DWORD dwChoice = _ShutdownChoiceFromMenuChoice(psmd);
                return _pLogOffPane->GetShutdownItemDescription(dwChoice, (LPWSTR)wParam, (UINT)lParam);
            }

            UINT uResID = _pLogOffPane->GetTipIDFromIDM(psmd->uId);
            if (uResID != -1)
            {
                LoadString(g_hinstCabinet, uResID, (LPWSTR)wParam, (int)lParam);
                return S_OK;
            }
            break;
        }
    }

    return S_FALSE;
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

HRESULT CALLBACK CLogOffMenuCallback_CreateInstance(IShellMenuCallback** ppsmc, CLogoffPane* pLogOffPane)
{
    CLogOffMenuCallback* plmc = new CLogOffMenuCallback(pLogOffPane);
    *ppsmc = static_cast<IShellMenuCallback*>(plmc);
    return *ppsmc ? S_OK : E_OUTOFMEMORY;
}
