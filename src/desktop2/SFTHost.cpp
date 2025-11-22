#include "pch.h"
#include "cocreateinstancehook.h"
#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "startmnu.h"

#define NTDDI_VERSION NTDDI_VISTA
#define _WIN32_WINNT _WIN32_WINNT_VISTA
#include <commoncontrols.h>


#define TF_HOST     0x00000010
#define TF_HOSTDD   0x00000040 // drag/drop
#define TF_HOSTPIN  0x00000080 // pin

#define ANIWND_WIDTH  80
#define ANIWND_HEIGHT 50

//---------BEGIN HACKS OF DEATH -------------

// HACKHACK - desktopp.h and browseui.h both define SHCreateFromDesktop
// What's worse, browseui.h includes desktopp.h! So you have to sneak it
// out in this totally wacky way.
//#include <desktopp.h>
#define SHCreateFromDesktop _SHCreateFromDesktop
//#include <browseui.h>

//---------END HACKS OF DEATH -------------


//
//  Unfortunately, WTL #undef's SelectFont, so we have to define it again.
//

inline HFONT SelectFont(HDC hdc, HFONT hf)
{
    return (HFONT)SelectObject(hdc, hf);
}


//****************************************************************************
//
//  Dummy IContextMenu
//
//  We use this when we can't get the real IContextMenu for an item.
//  If the user pins an object and then deletes the underlying
//  file, attempting to get the IContextMenu from the shell will fail,
//  but we need something there so we can add the "Remove from this list"
//  menu item.
//
//  Since this dummy context menu has no state, we can make it a static
//  singleton object.

class CEmptyContextMenu
    : public IContextMenu
{
public:
    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj)
    {
        static const QITAB qit[] = {
            QITABENT(CEmptyContextMenu, IContextMenu),
            { 0 },
        };
        return QISearch(this, qit, riid, ppvObj);
    }

    STDMETHODIMP_(ULONG) AddRef(void) { return 3; }
    STDMETHODIMP_(ULONG) Release(void) { return 2; }

    // *** IContextMenu ***
    STDMETHODIMP  QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
    {
        return ResultFromShort(0);  // No items added
    }

    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici)
    {
        ASSERT(FALSE);
        return E_FAIL;
    }

    STDMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pwRes, LPSTR pszName, UINT cchMax)
    {
        return E_INVALIDARG; // no commands; therefore, no command strings!
    }

public:
    IContextMenu *GetContextMenu()
    {
        // Don't need to AddRef since we are a static object
        return this;
    }
};

static CEmptyContextMenu s_EmptyContextMenu;

//****************************************************************************

#define WC_SFTBARHOST       TEXT("DesktopSFTBarHost")

// EXEX-VISTA(allison): Validated.
BOOL GetFileCreationTime(LPCTSTR pszFile, FILETIME *pftCreate)
{
    WIN32_FILE_ATTRIBUTE_DATA wfad;
    BOOL fRc = GetFileAttributesEx(pszFile, GetFileExInfoStandard, &wfad);
    if (fRc)
    {
        *pftCreate = wfad.ftCreationTime;
    }

    return fRc;
}

// {2A1339D7-523C-4E21-80D3-30C97B0698D2}
const CLSID TOID_SFTBarHostBackgroundEnum = {
    0x2A1339D7, 0x523C, 0x4E21,
    { 0x80, 0xD3, 0x30, 0xC9, 0x7B, 0x06, 0x98, 0xD2} };

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::Register()
{
    WNDCLASS wc;
    wc.style = 0;
    wc.lpfnWndProc = _WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = sizeof(void *);
    wc.hInstance = _AtlBaseModule.GetModuleInstance();
    wc.hIcon = 0;
    // We specify a cursor so the OOBE window gets something
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszMenuName = 0;
    wc.lpszClassName = WC_SFTBARHOST;
    return ::RegisterClass(&wc);
}

BOOL SFTBarHost::Unregister()
{
    return ::UnregisterClass(WC_SFTBARHOST, _AtlBaseModule.GetModuleInstance());
}

// EXEX-VISTA(allison): Partially validated. Recheck flow.
LRESULT CALLBACK SFTBarHost::_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    SFTBarHost *self = reinterpret_cast<SFTBarHost *>(GetWindowLongPtr(hwnd, 0));

    if (uMsg == WM_NCCREATE)
    {
        return _OnNcCreate(hwnd, uMsg, wParam, lParam);
    }
    else if (self)
    {

#define HANDLE_SFT_MESSAGE(wm, fn) case wm: return self->fn(hwnd, uMsg, wParam, lParam)

        switch (uMsg)
        {
            HANDLE_SFT_MESSAGE(WM_CREATE, _OnCreate);
            HANDLE_SFT_MESSAGE(WM_DESTROY, _OnDestroy);
            HANDLE_SFT_MESSAGE(WM_NCDESTROY, _OnNcDestroy);
            HANDLE_SFT_MESSAGE(WM_NOTIFY, _OnNotify);
            HANDLE_SFT_MESSAGE(WM_SIZE, _OnSize);
            HANDLE_SFT_MESSAGE(WM_ERASEBKGND, _OnEraseBackground);
            HANDLE_SFT_MESSAGE(WM_CONTEXTMENU, _OnContextMenu);
            HANDLE_SFT_MESSAGE(WM_CTLCOLORSTATIC, _OnCtlColorStatic);
            HANDLE_SFT_MESSAGE(WM_TIMER, _OnTimer);
            HANDLE_SFT_MESSAGE(WM_SETFOCUS, _OnSetFocus);

            HANDLE_SFT_MESSAGE(WM_INITMENUPOPUP, _OnMenuMessage);
            HANDLE_SFT_MESSAGE(WM_DRAWITEM, _OnMenuMessage);
            HANDLE_SFT_MESSAGE(WM_MENUCHAR, _OnMenuMessage);
            HANDLE_SFT_MESSAGE(WM_MEASUREITEM, _OnMenuMessage);

            HANDLE_SFT_MESSAGE(WM_SYSCOLORCHANGE, _OnSysColorChange);
            HANDLE_SFT_MESSAGE(WM_DISPLAYCHANGE, _OnForwardMessage);
            HANDLE_SFT_MESSAGE(WM_SETTINGCHANGE, _OnForwardMessage);

            HANDLE_SFT_MESSAGE(WM_UPDATEUISTATE, _OnUpdateUIState);

            HANDLE_SFT_MESSAGE(SFTBM_REPOPULATE, _OnRepopulate);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 0, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 1, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 2, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 3, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 4, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 5, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 6, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY + 7, _OnChangeNotify);
            HANDLE_SFT_MESSAGE(SFTBM_REFRESH, _OnRefresh);
            HANDLE_SFT_MESSAGE(SFTBM_CASCADE, _OnCascade);
            HANDLE_SFT_MESSAGE(SFTBM_ICONUPDATE, _OnIconUpdate);
        }

        // If this assert fires, you need to add more
        // HANDLE_SFT_MESSAGE(SFTBM_CHANGENOTIFY+... entries.
        COMPILETIME_ASSERT(SFTHOST_MAXNOTIFY == 9);

#undef HANDLE_SFT_MESSAGE

        return self->OnWndProc(hwnd, uMsg, wParam, lParam);
    }

    return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
#else
    LRESULT v6; // esi
    LRESULT updated; // eax
    unsigned int v8; // eax

    SFTBarHost *self = reinterpret_cast<SFTBarHost *>(GetWindowLongPtr(hwnd, 0));
    if (uMsg == 0x81)
    {
        return _OnNcCreate(hwnd, uMsg, wParam, lParam);
    }
    if (self)
    {
        self->AddRef();
        if (uMsg <= 0x121)
        {
            if (uMsg == 0x121)
            {
				//updated = self->_OnSetIdle(hwnd, uMsg, wParam, lParam); // EXEX-VISTA(allison): TODO: Uncomment when implemented.
                goto LABEL_74;
            }
            if (uMsg > 0x2C)
            {
                switch (uMsg)
                {
                    case 0x4Eu:
                        updated = self->_OnNotify(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x7Bu:
                        updated = self->_OnContextMenu(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x7Eu:
                        updated = self->_OnForwardMessage(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x82u:
                        updated = self->_OnNcDestroy(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x113u:
                        updated = self->_OnTimer(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x117u:
                        updated = self->_OnMenuMessage(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x120u:
                        updated = self->_OnMenuMessage(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                }
            }
            else
            {
                if (uMsg == 44)
                {
                    updated = self->_OnMenuMessage(hwnd, 0x2Cu, wParam, lParam);
                    goto LABEL_74;
                }
                if (uMsg <= 0x14)
                {
                    switch (uMsg)
                    {
                        case 0x14u:
                            updated = self->_OnEraseBackground(hwnd, uMsg, wParam, lParam);
                            break;
                        case 1u:
                            updated = self->_OnCreate(hwnd, uMsg, wParam, lParam);
                            break;
                        case 2u:
                            updated = self->_OnDestroy( hwnd, uMsg, wParam, lParam);
                            break;
                        case 5u:
                            updated = self->_OnSize(hwnd, uMsg, wParam, lParam);
                            break;
                        case 7u:
                            if (self->_hwndList)
                                SetFocus(self->_hwndList);
                            goto LABEL_16;
                        default:
                            goto LABEL_66;
                    }
                LABEL_74:
                    v6 = updated;
                    goto LABEL_75;
                }
                switch (uMsg)
                {
                    case 0x15u:
                        updated = self->_OnSysColorChange(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x1Au:
                        updated = self->_OnForwardMessage(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x2Bu:
                        updated = self->_OnMenuMessage(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                }
            }
        LABEL_66:
            updated = self->OnWndProc(hwnd, uMsg, wParam, lParam);
            goto LABEL_74;
        }
        v8 = 1030;
        if (uMsg > 0x406)
        {
            v8 = 1031;
            if (uMsg != 1031)
            {
                switch (uMsg)
                {
                    case 0x408u:
                        updated = self->_OnChangeNotify(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x409u:
                        updated = self->_OnChangeNotify(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x40Au:
                        self->_EnumerateContents(wParam);
                    LABEL_16:
                        v6 = 0;
                    LABEL_75:
                        self->Release();
                        return v6;
                    case 0x40Bu:
                        updated = self->_OnCascade(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x40Cu:
                        updated = self->_OnIconUpdate(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x40Du:
                        updated = self->_OnItemUpdate(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                }
                goto LABEL_66;
            }
        }
        else if (uMsg != 1030)
        {
            v8 = 1026;
            if (uMsg > 0x402)
            {
                v8 = 1027;
                if (uMsg != 1027)
                {
                    if (uMsg == 1028)
                        updated = self->_OnChangeNotify(hwnd, uMsg, wParam, lParam);
                    else
                        updated = self->_OnChangeNotify(hwnd, uMsg, wParam, lParam);
                    goto LABEL_74;
                }
            }
            else if (uMsg != 1026)
            {
                switch (uMsg)
                {
                    case 0x128u:
                        updated = self->_OnUpdateUIState(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x138u:
                        updated = self->_OnCtlColorStatic(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x400u:
                        updated = self->_OnRepopulate(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                    case 0x401u:
                        updated = self->_OnChangeNotify(hwnd, uMsg, wParam, lParam);
                        goto LABEL_74;
                }
                goto LABEL_66;
            }
        }
        updated = self->_OnChangeNotify(hwnd, v8, wParam, lParam);
        goto LABEL_74;
    }
    return DefWindowProcW(hwnd, uMsg, wParam, lParam);
#endif
}


// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnNcCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    SFTBarHost *self = NULL; // eax MAPDST

    SMPANEDATA *pspld = PaneDataFromCreateStruct(lParam);
    if (pspld->iPartId == SPP_PROGLIST)
    {
        self = ByUsage_CreateInstance();
    }
    else if (pspld->iPartId == SPP_PLACESLIST)
    {
        self = SpecList_CreateInstance();
    }
    else
    {
        //CcshellDebugMsgW(2, "Unknown panetype %d", pspld->iPartId);
        return 0;
    }

    if (self)
    {
        (void*)SetWindowLongPtr(hwnd, 0, (LONG_PTR)self);

        IUnknown_Set(&pspld->punk, static_cast<IServiceProvider *>(self));

        self->_hwnd = hwnd;
        self->_hTheme = pspld->hTheme;
        if (FAILED(self->Initialize()))
        {
            //CcshellDebugMsgW(2, "SFTBarHost::NcCreate Initialize call failed");
            return 0;
        }
        return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

//
//  The tile height is max(imagelist height, text height) + some margin
//  The margin is "scientifically computed" to be the value that looks
//  reasonably close to the bitmaps the designers gave us.
//

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_ComputeTileMetrics()
{
    int cyTile = _cyIcon;

    if (_iconsize == ICONSIZE_MEDIUM)
        cyTile /= 2;

    HDC hdc = GetDC(_hwndList);
    if (hdc)
    {
        HFONT hf = GetWindowFont(_hwndList);
        HFONT hfPrev = SelectFont(hdc, hf);
        SIZE siz;
        if (GetTextExtentPoint(hdc, TEXT("0"), 1, &siz))
        {
            if (_CanHaveSubtitles())
            {
                // Reserve space for the subtitle too
                siz.cy *= 2;
            }

            if (cyTile < siz.cy)
                cyTile = siz.cy;
        }

        SelectFont(hdc, hfPrev);
        ReleaseDC(_hwndList, hdc);
    }

    _cxIndent = _cxMargin + 2 * SHGetSystemMetricsScaled(SM_CXEDGE);
    if (_iconsize != ICONSIZE_MEDIUM)
        _cxIndent += _cxIcon;
    
    _cyTile = cyTile + _cyMargin * (_iconsize == ICONSIZE_MEDIUM ? 1 : 3);
    if (_iconsize == ICONSIZE_MEDIUM && _hTheme)
    {
        _cyTile += 1;
    }
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_SetTileWidth(int cxTile)
{
    LVTILEVIEWINFO tvi;
    tvi.cbSize = sizeof(tvi);
    tvi.dwMask = LVTVIM_TILESIZE | LVTVIM_COLUMNS;
    tvi.dwFlags = LVTVIF_FIXEDSIZE;

    // If we support cascading, then reserve space for the cascade arrows
    if (_dwFlags & HOSTF_CASCADEMENU)
    {
        // WARNING!  _OnLVItemPostPaint uses these margins
        tvi.dwMask |= LVTVIM_LABELMARGIN;
        tvi.rcLabelMargin.left   = 2;
        tvi.rcLabelMargin.top    = 0;
        tvi.rcLabelMargin.right  = _cxMarlett;
        tvi.rcLabelMargin.bottom = 0;
    }

    // Reserve space for subtitles if necessary
    tvi.cLines = _CanHaveSubtitles() ? 1 : 0;

    // _cyTile has the padding into account, but we want each item to be the height without padding
    tvi.sizeTile.cy = _cyTile;
    tvi.sizeTile.cx = cxTile;
    ListView_SetTileViewInfo(_hwndList, &tvi);
    _cxTile = cxTile;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup
void SFTBarHost::_CalculateSize(int a2)
{
    BOOL v3 = this->_cSep;
    if (!this->_cSep)
        v3 = this->_cPinnedDesired > 0;
    int cyTile = this->_cyTile;
    if (cyTile > 0)
        this->field_6C = (a2 - v3 * this->_cySepTile) / cyTile - this->_cPinned - this->_cSep;
    if (this->field_6C < 0)
        this->field_6C = 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (_hwndList)
    {
        SIZE sizeClient = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        sizeClient.cx -= (_margins.cxLeftWidth + _margins.cxRightWidth);
        sizeClient.cy -= (_margins.cyTopHeight + _margins.cyBottomHeight);

        SetWindowPos(_hwndList, NULL, _margins.cxLeftWidth, _margins.cyTopHeight,
                     sizeClient.cx, sizeClient.cy,
                     SWP_NOZORDER | SWP_NOOWNERZORDER);

        _SetTileWidth(sizeClient.cx);
        if (HasDynamicContent() || field_170)
        {
            _CalculateSize(sizeClient.cy);
            _InternalRepopulateList(field_170);
        }
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // if we're in unthemed mode, then we need to update our colors
    if (!_hTheme)
    {
        ListView_SetTextColor(_hwndList, GetSysColor(COLOR_MENUTEXT));
        _clrHot = GetSysColor(COLOR_MENUTEXT);
        _clrBG = GetSysColor(COLOR_MENU);
        _clrSubtitle = CLR_NONE;

        ListView_SetBkColor(_hwndList, _clrBG);
        ListView_SetTextBkColor(_hwndList, _clrBG);
    }

    return _OnForwardMessage(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
LRESULT SFTBarHost::_OnCtlColorStatic(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    // Use the same colors as the listview itself.
    HDC hdc = GET_WM_CTLCOLOR_HDC(wParam, lParam, uMsg);
    SetTextColor(hdc, ListView_GetTextColor(_hwndList));
    COLORREF clrBk = ListView_GetTextBkColor(_hwndList);
    if (clrBk == CLR_NONE)
    {
        // The animate control really wants to get a text background color.
        // It doesn't support transparency.
        if (GET_WM_CTLCOLOR_HWND(wParam, lParam, uMsg) == _hwndAni)
        {
            if (_hTheme)
            {
                if (!_hBrushAni)
                {
                    // We need to paint the theme background in a bitmap and use that
                    // to create a brush for the background of the flashlight animation
                    RECT rcClient;
                    GetClientRect(hwnd, &rcClient);
                    int x = (RECTWIDTH(rcClient) - ANIWND_WIDTH) / 2;     // IDA_SEARCH is ANIWND_WIDTH pix wide
                    int y = (RECTHEIGHT(rcClient) - ANIWND_HEIGHT) / 2;    // IDA_SEARCH is ANIWND_HEIGHT pix tall
                    RECT rc;
                    rc.top = y;
                    rc.bottom = y + ANIWND_HEIGHT;
                    rc.left = x;
                    rc.right = x + ANIWND_WIDTH;
                    HDC hdcBMP = CreateCompatibleDC(hdc);
                    HBITMAP hbmp = CreateCompatibleBitmap(hdc, ANIWND_WIDTH, ANIWND_HEIGHT);
                    POINT pt = { 0, 0 };

                    // Offset the viewport so that DrawThemeBackground draws the part that we care about
                    // at the right place
                    OffsetViewportOrgEx(hdcBMP, -x, -y, &pt);
                    SelectObject(hdcBMP, hbmp);
                    DrawThemeBackground(_hTheme, hdcBMP, _iThemePart, 0, &rcClient, 0);

                    // Our bitmap is now ready!
                    _hBrushAni = CreatePatternBrush(hbmp);

                    // Cleanup
                    SelectObject(hdcBMP, NULL);
                    DeleteObject(hbmp);
                    DeleteObject(hdcBMP);
                }
                return (LRESULT)_hBrushAni;
            }
            else
            {
                return (LRESULT)GetSysColorBrush(COLOR_MENU);
            }
        }

        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)GetStockBrush(HOLLOW_BRUSH);
    }
    else
    {
        return (LRESULT)GetSysColorBrush(COLOR_MENU);
    }
#else
    HDC hdc = GET_WM_CTLCOLOR_HDC(wParam, lParam, uMsg);


    SetTextColor(hdc, SendMessageW(_hwndList, LVM_GETTEXTCOLOR, 0, 0));

	COLORREF clrBk = SendMessageW(_hwndList, LVM_GETTEXTBKCOLOR, 0, 0);
    if (clrBk == -1)
    {
        if (GET_WM_CTLCOLOR_HWND(wParam, lParam, uMsg) != _hwndAni)
        {
            SetBkMode(hdc, 1);
            return (LRESULT)GetStockBrush(5);
        }

        if (_hTheme)
        {
            if (!_hBrushAni)
            {
                RECT rcClient;
                GetClientRect(hwnd, &rcClient);
                int x = (RECTWIDTH(rcClient) - ANIWND_WIDTH) / 2;
                int y = (RECTHEIGHT(rcClient) - ANIWND_HEIGHT) / 2;
                HDC hdcBMP = CreateCompatibleDC(hdc);
                if (hdcBMP)
                {
                    HBITMAP hbmp = CreateCompatibleBitmap(hdc, ANIWND_WIDTH, ANIWND_HEIGHT);
                    if (hbmp)
                    {
                        POINT pt = { 0, 0 };
                        OffsetViewportOrgEx(hdcBMP, -x, -y, &pt);
                        HGDIOBJ v9 = SelectObject(hdcBMP, hbmp);
                        DrawThemeBackground(_hTheme, hdcBMP, _iThemePart, 0, &rcClient, NULL);
                        _hBrushAni = CreatePatternBrush(hbmp);
                        
                        SelectObject(hdcBMP, v9);
                        DeleteObject(hbmp);
                    }
                    DeleteDC(hdcBMP);
                }
            }
            return (LRESULT)_hBrushAni;
        }
        else
        {
            return (LRESULT)GetSysColorBrush(COLOR_MENU);   
        }
    }
    else
    {
        return (LRESULT)GetSysColorBrush(COLOR_MENU);   
    }
#endif
}

//
//  Appends the PaneItem to _dpaEnum, or deletes it (and nulls it out)
//  if unable to append.
// 
// EXEX-VISTA(allison): Validated.
int SFTBarHost::_AppendEnumPaneItem(PaneItem *pitem)
{
    int iItem = _dpaEnumNew.AppendPtr(pitem);
    if (iItem < 0)
    {
		pitem->Release();
        iItem = -1;
    }
    return iItem;
}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::AddItem(PaneItem *pitem)
{
    BOOL fSuccess = FALSE;

    ASSERT(_fEnumerating);

    pitem->AddRef();
    if (_AppendEnumPaneItem(pitem) >= 0)
    {
        fSuccess = TRUE;
    }
    return fSuccess;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_RepositionItems()
{

    int iItem;
    for (iItem = ListView_GetItemCount(_hwndList) - 1; iItem >= 0; iItem--)
    {
        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem)
        {
            POINT pt;
            _ComputeListViewItemPosition(pitem->_iPos, &pt);
            ListView_SetItemPosition(_hwndList, iItem, pt.x, pt.y);
            pitem->Release();
        }
    }
}

//
//  pvData = the window to receive the icon
//  pvHint = pitem whose icon we just extracted
//  iIconIndex = the icon we got
//
// EXEX-VISTA(allison): Validated.
void SFTBarHost::SetIconAsync(LPVOID pvData, LPVOID pvHint, INT iIconIndex, INT iOpenIconIndex)
{
    HWND hwnd = (HWND)pvData;
    if (IsWindow(hwnd))
    {
        SFTBarHost *self = reinterpret_cast<SFTBarHost *>(GetWindowLongPtr(hwnd, 0));
        if (self)
        {
            IImageList2 *piml = NULL;
            if (SUCCEEDED(HIMAGELIST_QueryInterface(self->_himl, IID_PPV_ARGS(&piml))))
            {
                piml->ForceImagePresent(iIconIndex, ILFIP_ALWAYS);
            }

            PostMessage(hwnd, SFTBM_ICONUPDATE, iIconIndex, (LPARAM)pvHint);

            if (piml)
            {
                piml->Release();
            }
        }
    }
}

//
//  wParam = icon index
//  lParam = pitem to update
//
// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnIconUpdate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    //
    //  Do not dereference lParam (pitem) until we are sure it is valid.
    //

    LVFINDINFO fi;
    LVITEM lvi;

    fi.flags = LVFI_PARAM;
    fi.lParam = lParam;
    lvi.iItem = ListView_FindItem(_hwndList, -1, &fi);
    if (lvi.iItem >= 0)
    {
        lvi.mask = LVIF_IMAGE;
        lvi.iSubItem = 0;
        lvi.iImage = (int)wParam;
        ListView_SetItem(_hwndList, &lvi);
        // Now, we need to go update our cached bitmap version of the start menu.
        _SendNotify(_hwnd, SMN_NEEDREPAINT, NULL);

        RECT rc;
        if (ListView_GetItemRect(_hwndList, lvi.iItem, &rc, LVIR_ICON))
        {
            InvalidateRect(_hwndList, &rc, TRUE);
		}
    }
    return 0;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
BOOL SFTBarHost::_OnTextUpdate(int iItem)
{
    BOOL bRet = TRUE;
    WCHAR szText[260];

    LVITEM lvi;
    lvi.iSubItem = 0;
    lvi.pszText = szText;
    lvi.iItem = iItem;
    lvi.mask = 5;
    lvi.cchTextMax = ARRAYSIZE(szText);
    if (ListView_GetItem(_hwndList, &lvi))
    {
        PaneItem* pitem = _GetItemFromLVLParam(lvi.lParam);
        if (pitem)
        {
            lvi.iSubItem = 0;
            lvi.iItem = iItem;
            lvi.mask = 1;
            lvi.pszText = _DisplayNameOfItem(pitem, 0);
            if (lvi.pszText)
            {
                if (StrCmpN(szText, lvi.pszText, ARRAYSIZE(szText)))
                {
                    ListView_SetItem(_hwndList, &lvi);
                    _SendNotify(_hwnd, SMN_SHOWNEWAPPSTIP, 0);
                }
                CoTaskMemFree(lvi.pszText);
            }
            pitem->Release();
        }
        else
        {
            bRet = 0;
        }
    }
    return bRet;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnItemUpdate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    for (int i = 0; i < ListView_GetItemCount(_hwndList); ++i)
    {
        PaneItem *pitem = _GetItemFromLV(i);
        if (pitem)
        {
            if (OnItemUpdate(pitem, wParam, lParam) == S_OK)
            {
                _OnTextUpdate(i);
                if (_IsPrivateImageList())
                {
                    if (pitem->GetPrivateIcon() != -1)
                    {
                        _OnIconUpdate(_hwnd, 0x40Cu, pitem->GetPrivateIcon(), (LPARAM)pitem);
                    }
                }
                else
                {
                    IShellFolder *psf;
                    LPCITEMIDLIST pidl;
                    if (SUCCEEDED(GetFolderAndPidl(pitem, &psf, &pidl)))
                    {
                        int iImage = AddImageForItem(pitem, psf, pidl, 0);
                        if (iImage != -1)
                        {
                            _OnIconUpdate(_hwnd, 0x40C, iImage, (LPARAM)pitem);
                        }
                        psf->Release();
                    }
                }
            }
            pitem->Release();
        }
    }
    OnItemUpdateComplete(wParam, lParam);
    return 0;
}

// EXEX-VISTA(allison): Validated.
int SFTBarHost::AddImageForItem(PaneItem *pitem, IShellFolder *psf, LPCITEMIDLIST pidl, int iPos)
{
    int iIndex = pitem->GetPrivateIcon();
    printf("AddImageForItem: Private icon index = %d\n", iIndex);
    if (iIndex == -1)
    {
        if (NeedBackgroundEnum())
        {
            if (_psched
                && SUCCEEDED(SHMapIDListToSystemImageListIndexAsync(_psched, psf, pidl, SetIconAsync, _hwnd, pitem, &iIndex, NULL)))
            {
                IImageList2 *piml = NULL;
                if (SUCCEEDED(HIMAGELIST_QueryInterface(_himl, IID_PPV_ARGS(&piml)))
                    && FAILED(piml->ForceImagePresent(iIndex, ILFIP_FROMSTANDBY)))
                {
                    CLoadIconTask *pTask = new CLoadIconTask(_hwnd, pitem, iIndex);
                    _psched->AddTask(pTask, TOID_SFTBarHostBackgroundEnum, (DWORD_PTR)this, 0x10001000);
                    pTask->Release();
                }

                if (piml)
                {
                    piml->Release();
                }
            }
        }
        else
        {
            SHMapIDListToSystemImageListIndex(psf, pidl, &iIndex, NULL);
        }
    }
    return iIndex;
}

HICON _IconOf(IShellFolder *psf, PCUITEMID_CHILD pidl, int cxIcon, LPCWSTR pszPath, int a5)
{
    HICON hicoLarge = NULL;
    HICON hicoSmall = NULL;

    IExtractIcon *pxi;
    HRESULT hr = psf->GetUIObjectOf(NULL, 1, &pidl, IID_IExtractIcon, NULL, (void **)&pxi);
    if (SUCCEEDED(hr))
    {
        WCHAR szPath[MAX_PATH];
        int iIndex;
        UINT uiFlags;
        hr = pxi->GetIconLocation(0, szPath, ARRAYSIZE(szPath), &iIndex, &uiFlags);
        if (hr == S_FALSE)
        {
            StringCchCopy(szPath, ARRAYSIZE(szPath), TEXT("shell32.dll"));
            iIndex = II_DOCNOASSOC;
            hr = S_OK;
        }

        if (a5 == iIndex && a5 != -1 && pszPath && !StrCmpI(pszPath, szPath))
        {
            hr = S_OK;
        }
        else if (SUCCEEDED(hr))
        {
            hr = pxi->Extract(szPath, iIndex, &hicoLarge, &hicoSmall, cxIcon);
            if (hr == S_FALSE)
            {
                hr = SHDefExtractIcon(szPath, iIndex, uiFlags, &hicoLarge, &hicoSmall, cxIcon);
            }
        }

        pxi->Release();
    }

    if (FAILED(hr))
    {
        SFGAOF attr = SFGAO_FOLDER;
        int iIndex;
        if (SUCCEEDED(psf->GetAttributesOf(1, &pidl, &attr))
            && (attr & SFGAO_FOLDER))
        {
            iIndex = II_FOLDER;
        }
        else
        {
            iIndex = II_DOCNOASSOC;
        }
        SHDefExtractIcon(TEXT("shell32.dll"), iIndex, 0, &hicoLarge, &hicoSmall, cxIcon);
    }

    if (hicoSmall)
        DestroyIcon(hicoSmall);

    return hicoLarge;
}

//
//  There are two sets of numbers that keep track of items.  Sorry.
//  (I tried to reduce it to one, but things got hairy.)
//
//  1. Position numbers.  Separators occupy a position number.
//  2. Item numbers (listview).  Separators do not consume an item number.
//
//  Example:
//
//              iPos        iItem
//
//  A           0           0
//  B           1           1
//  ----        2           N/A
//  C           3           2
//  ----        4           N/A
//  D           5           3
//
//  _rgiSep[] = { 2, 4 };
//
//  _PosToItemNo and _ItemNoToPos do the conversion.

int SFTBarHost::_PosToItemNo(int iPos)
{
    // Subtract out the slots occupied by separators.
    int iItem = iPos;
    for (int i = 0; i < _cSep && _rgiSep[i] < iPos; i++)
    {
        iItem--;
    }
    return iItem;
}

int SFTBarHost::_ItemNoToPos(int iItem)
{
    // Add in the slots occupied by separators.
    int iPos = iItem;
    for (int i = 0; i < _cSep && _rgiSep[i] <= iPos; i++)
    {
        iPos++;
    }
    return iPos;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void SFTBarHost::_ComputeListViewItemPosition(int iItem, POINT* pptOut)
{
#ifdef DEAD_CODE
    // WARNING!  _InternalRepopulateList uses an incremental version of this
    // algorithm.  Keep the two in sync!

    ASSERT(_cyTilePadding >= 0);

    int y = iItem * _cyTile;

    // Adjust for all the separators in the list
    for (int i = 0; i < _cSep; i++)
    {
        if (_rgiSep[i] < iItem)
        {
            y = y - _cyTile + _cySepTile;
        }
    }

    pptOut->x = _cxMargin;
    pptOut->y = y;
#else

    int y = iItem * _cyTile;

    if (_cSep > 0)
    {
        int* rgiSep = this->_rgiSep;
        int iItema = this->_cSep;
        while (iItema)
        {
            if (*rgiSep < iItem)
            {
                y += this->_cySepTile - this->_cyTile;
            }
            ++rgiSep;
            --iItema;
        }
    }

    pptOut->x = _cxMargin;
    pptOut->y = y;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
int SFTBarHost::_InsertListViewItem(int iPos, PaneItem *pitem)
{
#ifdef DEAD_CODE
    ASSERT(pitem);

    int iItem = -1;
    IShellFolder *psf = NULL;
    LPCITEMIDLIST pidl = NULL;
    LVITEM lvi;
    lvi.pszText = NULL;

    lvi.mask = 0;

    // If necessary, tell listview that we want to use column 1
    // as the subtitle.
    if (_iconsize == ICONSIZE_LARGE && pitem->HasSubtitle())
    {
        const static UINT One = 1;
        lvi.mask = LVIF_COLUMNS;
        lvi.cColumns = 1;
        lvi.puColumns = const_cast<UINT*>(&One);
    }

    ASSERT(!pitem->IsSeparator());

    lvi.mask |= LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;
    if (FAILED(GetFolderAndPidl(pitem, &psf, &pidl)))
    {
        goto exit;
    }

    if (lvi.mask & LVIF_IMAGE)
    {
        lvi.iImage = AddImageForItem(pitem, psf, pidl, iPos);
    }

    if (lvi.mask & LVIF_TEXT)
    {
        if (_iconsize == ICONSIZE_SMALL && pitem->HasSubtitle())
        {
            lvi.pszText = SubtitleOfItem(pitem, psf, pidl);
        }
        else
        {
            lvi.pszText = DisplayNameOfItem(pitem, psf, pidl, SHGDN_NORMAL);
        }
        if (!lvi.pszText)
        {
            goto exit;
        }
    }

    lvi.iItem = iPos;
    lvi.iSubItem = 0;
    lvi.lParam = reinterpret_cast<LPARAM>(pitem);
    iItem = ListView_InsertItem(_hwndList, &lvi);

    // If the item has a subtitle, add it.
    // If this fails, don't worry.  The subtitle is just a fluffy bonus thing.
    if (iItem >= 0 && (lvi.mask & LVIF_COLUMNS))
    {
        lvi.iItem = iItem;
        lvi.iSubItem = 1;
        lvi.mask = LVIF_TEXT;
        SHFree(lvi.pszText);
        lvi.pszText = SubtitleOfItem(pitem, psf, pidl);
        if (lvi.pszText)
        {
            ListView_SetItem(_hwndList, &lvi);
        }
    }

exit:
    ATOMICRELEASE(psf);
    SHFree(lvi.pszText);
    return iItem;
#else
	ASSERT(pitem); // 687

    int iItem = -1;
    IShellFolder* psf = NULL;
    PITEMID_CHILD pidl = NULL;
    LVITEM lvi;
    lvi.pszText = NULL;

    lvi.mask = 0;

    if (_iconsize == ICONSIZE_LARGE && pitem->HasSubtitle())
    {
		const static UINT One = 1;
        lvi.mask = LVIF_COLUMNS;
        lvi.cColumns = 1;
        lvi.puColumns = const_cast<UINT*>(&One);
    }

    ASSERT(!pitem->IsSeparator()); // 707

    lvi.mask |= 7u;
    if (GetFolderAndPidl(pitem, &psf, (LPCITEMIDLIST*)&pidl) >= 0)
    {
        if ((lvi.mask & 2) != 0)
            lvi.iImage = AddImageForItem(pitem, psf, pidl, 0);

		WCHAR* v4;
        if ((lvi.mask & 1) == 0
            || (this->_iconsize || (pitem->_dwFlags & 2) == 0
                ? (v4 = this->DisplayNameOfItem(pitem, psf, pidl, 0))
                : (v4 = this->SubtitleOfItem(pitem, psf, pidl)),
                (lvi.pszText = v4) != NULL))
        {
            lvi.iItem = iPos;
            lvi.iSubItem = 0;
            lvi.lParam = reinterpret_cast<LPARAM>(pitem);;
			iItem = ListView_InsertItem(_hwndList, &lvi);

            if (iItem >= 0 && (lvi.mask & 0x200) != 0)
            {
                lvi.iItem = iItem;
                lvi.iSubItem = 1;
                lvi.mask = LVIF_TEXT;
                CoTaskMemFree(lvi.pszText);
                lvi.pszText = SubtitleOfItem(pitem, psf, (LPCITEMIDLIST)pidl);
                if (lvi.pszText)
                {
                    ListView_SetItem(_hwndList, &lvi);
                }
            }
        }
    }

    IUnknown_SafeReleaseAndNullPtr(&psf);
    CoTaskMemFree(lvi.pszText);
    return iItem;
#endif
}


// Add items to our view, or at least as many as will fit

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void SFTBarHost::_RepopulateList()
{
#ifdef DEAD_CODE
    //
    //  Kill the async enum animation now that we're ready
    //
    if (_idtAni)
    {
        KillTimer(_hwnd, _idtAni);
        _idtAni = 0;
    }
    if (_hwndAni)
    {
        if (_hBrushAni)
        {
            DeleteObject(_hBrushAni);
            _hBrushAni = NULL;
        }
        DestroyWindow(_hwndAni);
        _hwndAni = NULL;
    }

    // Let's see if anything changed
    BOOL fChanged = FALSE;
    if (_fForceChange)
    {
        _fForceChange = FALSE;
        fChanged = TRUE;
    }
    else if (_dpaEnum.GetPtrCount() == _dpaEnumNew.GetPtrCount())
    {
        int iMax = _dpaEnum.GetPtrCount();
        int i;
        for (i = 0; i < iMax; i++)
        {
            if (!_dpaEnum.FastGetPtr(i)->IsEqual(_dpaEnumNew.FastGetPtr(i)))
            {
                fChanged = TRUE;
                break;
            }
        }
    }
    else
    {
        fChanged = TRUE;
    }


    // No need to do any real work if nothing changed.
    if (fChanged)
    {
        // Now move the _dpaEnumNew to _dpaEnum
        // Clear out the old DPA, we don't need it anymore
        _dpaEnum.EnumCallbackEx(PaneItem::DPAEnumCallback, (void *)NULL);
        if (_dpaEnum)
        {
            _dpaEnum.DeleteAllPtrs();
        }

        // switch DPAs now
        CDPA<PaneItem, CTContainer_PolicyUnOwned<PaneItem>> dpaTemp = _dpaEnum;
        _dpaEnum = _dpaEnumNew;
        _dpaEnumNew = dpaTemp;

        _InternalRepopulateList(0);
    }
    else
    {
        // Clear out the new DPA, we don't need it anymore
        _dpaEnumNew.EnumCallbackEx(PaneItem::DPAEnumCallback, (void *)NULL);
        if (_dpaEnumNew)
        {
            _dpaEnumNew.DeleteAllPtrs();
        }
    }

    _fNeedsRepopulate = FALSE;
#else
    int cp; // ecx
    int v8; // eax
    int i; // edi
    CDPA<PaneItem, CTContainer_PolicyUnOwned<PaneItem>> dpaTemp; // [esp+Ch] [ebp-4h] SPLIT BYREF
    int iMax; // [esp+Ch] [ebp-4h]

    if (_idtAni)
    {
        KillTimer(_hwnd, _idtAni);
        _idtAni = 0;
    }

    if (_hwndAni)
    {
        if (_hBrushAni)
        {
            DeleteObject(_hBrushAni);
            _hBrushAni = 0;
        }
        NotifyWinEvent(0x8001u, _hwndAni, 0, 0);
        DestroyWindow(_hwndAni);
        _hwndAni = 0;
    }

    if (_fForceChange)
    {
        _fForceChange = 0;
        goto LABEL_9;
    }

    cp = _dpaEnum.GetPtrCount();
    v8 = _dpaEnumNew.GetPtrCount();

    if (cp != v8)
    {
    LABEL_9:
        _dpaEnum.EnumCallback(PaneItem::DPAEnumCallback);
        if (_dpaEnum)
        {
            _dpaEnum.DeleteAllPtrs();
        }

        dpaTemp.Attach(_dpaEnum.Detach());
        _dpaEnum.Attach(_dpaEnumNew.Detach());
        _dpaEnumNew.Attach(dpaTemp.Detach());

        _InternalRepopulateList(0);
    }
    else
    {
        iMax = _dpaEnum.GetPtrCount();
        i = 0;
        if (iMax > 0)
        {
            do
            {
                if (!_dpaEnum.FastGetPtr(i)->IsEqual(_dpaEnumNew.FastGetPtr(i)))
                {
                    goto LABEL_9;
                }
            } while (++i < iMax);
        }

        _dpaEnumNew.EnumCallback(PaneItem::DPAEnumCallback);
        if (_dpaEnumNew)
        {
            _dpaEnumNew.DeleteAllPtrs();
        }
    }
    _fNeedsRepopulate = 0;
#endif
}

// The internal version is when we decide to repopulate on our own,
// not at the prompting of the background thread.  (Therefore, we
// don't nuke the animation.)

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void SFTBarHost::_InternalRepopulateList(BOOL a2)
{
#ifdef DEAD_CODE

    //
    //  Start with a clean slate.
    //

    ListView_DeleteAllItems(_hwndList);
    if (_IsPrivateImageList())
    {
        ImageList_RemoveAll(_himl);
    }

    int cPinned = 0;
    int cNormal = 0;

    _DebugConsistencyCheck();

    SetWindowRedraw(_hwndList, FALSE);


    //
    //  To populate the list, we toss the pinned items at the top,
    //  then let the enumerated items flow beneath them.
    //
    //  Separator "items" don't get added to the listview.  They
    //  are added to the special "separators list".
    //
    //  WARNING!  We are computing incrementally the same values as
    //  _ComputeListViewItemPosition.  Keep the two in sync.
    //

    int iPos;                   // the slot we are trying to fill
    int iEnum;                  // the item index we will fill it from
    int y = 0;                  // where the next item should be placed
    BOOL fSepSeen = FALSE;      // have we seen a separator yet?
    PaneItem *pitem;            // the item that will fill it

    _cSep = 0;                  // no separators (yet)

    RECT rc;
    GetClientRect(_hwndList, &rc);
    //
    //  Subtract out the bonus separator used by SPP_PROGLIST
    //
    if (_iThemePart == SPP_PROGLIST)
    {
        rc.bottom -= _cySep;
    }

    // Note that the loop control must be a _dpaEnum.GetPtr(), not a
    // _dpaEnum.FastGetPtr(), because iEnum can go past the end of the
    // array if we do't have enough items to fill the view.
    //
    //
    // The "while" condition is "there is room for another non-separator
    // item and there are items remaining in the enumeration".

    BOOL fCheckMaxLength = HasDynamicContent();

    for (iPos = iEnum = 0;
        (pitem = _dpaEnum.GetPtr(iEnum)) != NULL;
        iEnum++)
    {
        if (fCheckMaxLength)
        {
            if (y + _cyTile > rc.bottom)
            {
                break;
            }

            // Once we hit a separator, check if we satisfied the number
            // of normal items.  We have to wait until a separator is
            // hit, because _cNormalDesired can be zero; otherwise we
            // would end up stopping before adding even the pinned items!
            if (fSepSeen && cNormal >= _cNormalDesired)
            {
                break;
            }
        }

#ifdef DEBUG
        // Make sure that we are in sync with _ComputeListViewItemPosition
        POINT pt;
        _ComputeListViewItemPosition(iPos, &pt);
        ASSERT(pt.x == _cxMargin);
        ASSERT(pt.y == y);
#endif
        if (pitem->IsSeparator())
        {
            fSepSeen = TRUE;

            // Add the separator, but only if it actually separate something.
            // If this EVAL fires, it means somebody added a separator
            // and MAX_SEPARATORS needs to be increased.
            if (iPos > 0 && EVAL(_cSep < ARRAYSIZE(_rgiSep)))
            {
                _rgiSep[_cSep++] = iPos++;
                y += _cySepTile;
            }
        }
        else
        {
            if (_InsertListViewItem(iPos, pitem) >= 0)
            {
                pitem->_iPos = iPos++;
                y += _cyTile;
                if (pitem->IsPinned())
                {
                    cPinned++;
                }
                else
                {
                    cNormal++;
                }
            }
        }
    }

    //
    //  If the last item was a separator, then delete it
    //  since it's not actually separating anything.
    //
    if (_cSep && _rgiSep[_cSep - 1] == iPos - 1)
    {
        _cSep--;
    }


    _cPinned = cPinned;

    //
    //  Now put the items where they belong.
    //
    _RepositionItems();


    SetWindowRedraw(_hwndList, TRUE);

    // Now, we need to go update our cached bitmap version of the start menu.
    _SendNotify(_hwnd, SMN_NEEDREPAINT, NULL);

    _DebugConsistencyCheck();
#else
    SetWindowRedraw(_hwndList, FALSE);
	ListView_DeleteAllItems(this->_hwndList);

    int cPinned = 0;
    int cNormal = 0;
    //++this->_fPopulating;
    int y = 0;
    int fSepSeen = 0;
    _cSep = 0;

    RECT rc;
    GetClientRect(this->_hwndList, &rc);
    if (this->_iThemePart == SPP_PROGLIST)
        rc.bottom -= this->_cySep;

    int fCheckMaxLength = this->HasDynamicContent();

    if (a2)
        a2 = this->_cNormalDesired > this->_cSep + this->field_6C;

    int iEnum = 0;
    int iPos = 0;
    while (1)
    {
        PaneItem *pitem = _dpaEnum.GetPtr(iEnum);
        if (!pitem || fCheckMaxLength && y + this->_cyTile > rc.bottom)
            break;
        if (fSepSeen && cNormal >= this->_cNormalDesired)
            break;

        POINT pt;
        _ComputeListViewItemPosition(iPos, &pt);

        ASSERT(pt.x == _cxMargin); // 941
        ASSERT(pt.y == y); // 942

        if (pitem->_iPinPos == -2)
        {
            fSepSeen = 1;
            if (iPos > 0 && this->_cSep < 3 && EVAL(_cSep < ARRAYSIZE(_rgiSep))) // 951
            {
                _rgiSep[_cSep++] = iPos++;
                y += _cySepTile;
            }
        }
        else if ((!a2 || !pitem->IsHiddenInSafeMode()) && _InsertListViewItem(iPos, pitem) >= 0)
        {
            pitem->_iPos = iPos++;
            y += this->_cyTile;
            if (pitem->_iPinPos < 0)
                ++cNormal;
            else
                ++cPinned;
        }
        ++iEnum;
    }

    if (_cSep && _rgiSep[_cSep - 1] == iPos - 1)
        _cSep--;

    this->_cPinned = cPinned;

    SFTBarHost::_RepositionItems();
    //--this->_fPopulating;
    SetWindowRedraw(_hwndList, TRUE);
    _SendNotify(_hwnd, SMN_NEEDREPAINT, 0);
#endif
}

void SHLogicalToPhysicalDPI(int *a1, int *a2);

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    RECT rc;
    GetClientRect(_hwnd, &rc);

    if (_hTheme)
    {
        GetThemeMargins(_hTheme, NULL, _iThemePart, 0, TMT_CONTENTMARGINS, &rc, &_margins);
    }
    else
    {
        _margins.cyTopHeight = _iThemePart == SPP_PLACESLIST ? 0 : 2 * GetSystemMetrics(SM_CXEDGE);
        _margins.cxLeftWidth = 2 * GetSystemMetrics(SM_CXEDGE);
        _margins.cxRightWidth = 2 * GetSystemMetrics(SM_CXEDGE);
    }


    //
    //  Now to create the listview.
    //

    DWORD dwStyle = WS_CHILD | WS_VISIBLE |
        WS_CLIPCHILDREN | WS_CLIPSIBLINGS |
        // Do not set WS_TABSTOP; SFTBarHost handles tabbing
        LVS_LIST |
        LVS_SINGLESEL |
        LVS_NOSCROLL |
        LVS_SHAREIMAGELISTS;

    if (_dwFlags & HOSTF_CANRENAME)
    {
        dwStyle |= LVS_EDITLABELS;
    }

    DWORD dwExStyle = 0;

    _hwndList = SHFusionCreateWindowEx(dwExStyle, WC_LISTVIEW, NULL, dwStyle,
        _margins.cxLeftWidth, _margins.cyTopHeight, rc.right, rc.bottom, // no point in being too exact, we'll be resized later
        _hwnd, NULL,
        _AtlBaseModule.GetModuleInstance(), NULL);
    if (!_hwndList)
        return -1;

    LPCWSTR pszTheme;
    if (IsCompositionActive())
    {
        pszTheme = _iThemePart == SPP_PLACESLIST ? L"StartMenuPlaceListComposited" : L"StartMenuComposited";
    }
    else
    {
        pszTheme = L"StartMenu";
    }

    SetWindowTheme(_hwndList, pszTheme, NULL);

    //
    //  Don't freak out if this fails.  It just means that the accessibility
    //  stuff won't be perfect.
    //
    SetAccessibleSubclassWindow(_hwndList);

    //
    //  Create two dummy columns.  We will never display them, but they
    //  are necessary so that we have someplace to put our subtitle.
    //
    LVCOLUMN lvc;
    lvc.mask = LVCF_WIDTH;
    lvc.cx = 1;
    if (ListView_InsertColumn(_hwndList, 0, &lvc) < 0 || ListView_InsertColumn(_hwndList, 1, &lvc) < 0)
        return -1;


    HWND hwndTT = ListView_GetToolTips(_hwndList);
    if (hwndTT)
    {
        SetWindowPos(hwndTT, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOOWNERZORDER | SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
    }


    // Must do Marlett after doing the listview font, because we base the
    // Marlett font metrics on the listview font metrics (so they match)
    if (_dwFlags & HOSTF_CASCADEMENU)
    {
        if (!_CreateMarlett())
            return -1;
    }

    // We can survive if these objects fail to be created
    CoCreateInstanceHook(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&_pdth));
    CoCreateInstanceHook(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&_pdsh));

    //
    // If this fails, no big whoop - you just don't get
    // drag/drop, boo hoo.
    //
    RegisterDragDrop(_hwndList, this);

    //  If this fails, then disable "fancy droptarget" since we won't be
    //  able to manage it properly.
    if (!SetWindowSubclass(_hwndList, s_DropTargetSubclassProc, 0,
        reinterpret_cast<DWORD_PTR>(this)))
    {
        IUnknown_SafeReleaseAndNullPtr(&_pdth);
    }

    if (!_dpaEnum.Create(4))
        return -1;

    if (!_dpaEnumNew.Create(4))
        return -1;

    //-------------------------
    // Imagelist goo
    int iImageList = -1;
    int iIconSize = ReadIconSize();

    _iconsize = (ICONSIZE)iIconSize;
    if (iIconSize == ICONSIZE_MEDIUM)
    {
        _cyIcon = 64;
        _cxIcon = 64;
        SHLogicalToPhysicalDPI(&_cxIcon, &_cyIcon);
    }
    else
    {
        _cyIcon = GetSystemMetrics(iIconSize ? SM_CXICON : SM_CXSMICON);
        _cxIcon = GetSystemMetrics(iIconSize ? SM_CXICON : SM_CXSMICON);
        iImageList = _iconsize == ICONSIZE_SMALL ? SHIL_SYSSMALL : SHIL_LARGE;
    }

    IImageList2 *piml;
    if (SUCCEEDED(SHGetImageList(iImageList, IID_PPV_ARGS(&piml))))
    {
        if (SUCCEEDED(piml->Resize(_cxIcon, _cyIcon)))
            _himl = (HIMAGELIST)piml;
        else
            piml->Release();
    }

    if (_IsPrivateImageList())
    {
        UINT flags = ImageList_GetFlags(_himl);
        ImageList_Destroy(_himl);
        _himl = ImageList_Create(_cxIcon, _cyIcon, flags, 8, 2);
    }

    if (!_himl)
        return -1;

    if (_iconsize != ICONSIZE_MEDIUM)
        ListView_SetImageList(_hwndList, _himl, LVSIL_NORMAL);

    // Register for SHCNE_UPDATEIMAGE so we know when to reload our icons
    _RegisterNotify(SFTHOST_HOSTNOTIFY_UPDATEIMAGE, SHCNE_UPDATEIMAGE, NULL, FALSE);

    //-------------------------

    _cxMargin = GetSystemMetrics(SM_CXEDGE);
    _cyMargin = GetSystemMetrics(SM_CYEDGE);

    _ComputeTileMetrics();

    //
    //  In the themed case, the designers want a narrow separator.
    //  In the nonthemed case, we need a fat separator because we need
    //  to draw an etch (which requires two pixels).
    //
    if (_hTheme)
    {
        SIZE siz = { 0, 0 };
        HDC hdc = GetDC(_hwndList);
        if (hdc)
        {
            GetThemePartSize(_hTheme, hdc, _iThemePartSep, 0, NULL, TS_DRAW, &siz);
            ReleaseDC(_hwndList, hdc);
        }
        _cySep = siz.cy;
    }
    else
    {
        _cySep = GetSystemMetrics(SM_CYEDGE);
    }

    _cySepTile = _iThemePart == SPP_PLACESLIST ? _cySep + 1 : 4 * _cySep;

    ASSERT(rc.left == 0 && rc.top == 0); // Should still be a client rectangle
    _SetTileWidth(rc.right);             // so rc.right = RCWIDTH and rc.bottom = RCHEIGHT

    DWORD dwLvExStyle = LVS_EX_COLUMNSNAPPOINTS |
        LVS_EX_INFOTIP |
        LVS_EX_FULLROWSELECT;

    if (!GetSystemMetrics(SM_REMOTESESSION) && !GetSystemMetrics(SM_REMOTECONTROL))
    {
        dwLvExStyle |= LVS_EX_DOUBLEBUFFER;
    }

    ListView_SetExtendedListViewStyleEx(_hwndList, dwLvExStyle,
        dwLvExStyle);
    if (!_hTheme)
    {
        ListView_SetTextColor(_hwndList, GetSysColor(COLOR_MENUTEXT));
        _clrHot = GetSysColor(COLOR_HIGHLIGHTTEXT);
        _clrBG = GetSysColor(COLOR_MENU);       // default color for no theme case
        _clrSubtitle = CLR_NONE;

    }
    else
    {
        COLORREF clrText;

        GetThemeColor(_hTheme, _iThemePart, 0, TMT_HOTTRACKING, &_clrHot);  // todo - use state
        GetThemeColor(_hTheme, _iThemePart, 0, TMT_CAPTIONTEXT, &_clrSubtitle);
        _clrBG = CLR_NONE;

        GetThemeColor(_hTheme, _iThemePart, 0, TMT_TEXTCOLOR, &clrText);
        ListView_SetTextColor(_hwndList, clrText);
        ListView_SetOutlineColor(_hwndList, _clrHot);
    }

    ListView_SetBkColor(_hwndList, _clrBG);
    ListView_SetTextBkColor(_hwndList, _clrBG);


    ListView_SetView(_hwndList, LV_VIEW_TILE);

    field_170 = GetSystemMetrics(SM_CYSCREEN) <= 480;

    // USER will send us a WM_SIZE after the WM_CREATE, which will cause
    // the listview to repopulate, if we chose to repopulate in the
    // foreground.

    return 0;

}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::_CreateMarlett()
{
    HDC hdc = GetDC(_hwndList);
    if (hdc)
    {
        HFONT hfPrev = SelectFont(hdc, GetWindowFont(_hwndList));
        if (hfPrev)
        {
            TEXTMETRIC tm;
            if (GetTextMetrics(hdc, &tm))
            {
                LOGFONT lf;
                ZeroMemory(&lf, sizeof(lf));
                lf.lfHeight = tm.tmAscent;
                lf.lfWeight = FW_NORMAL;
                lf.lfCharSet = SYMBOL_CHARSET;
                StrCpyN(lf.lfFaceName, TEXT("Marlett"), ARRAYSIZE(lf.lfFaceName));
                _hfMarlett = CreateFontIndirect(&lf);

                if (_hfMarlett)
                {
                    SelectFont(hdc, _hfMarlett);
                    if (GetTextMetrics(hdc, &tm))
                    {
                        _tmAscentMarlett = tm.tmAscent;
                        SIZE siz;
                        if (GetTextExtentPoint(hdc, TEXT("8"), 1, &siz))
                        {
                            _cxMarlett = siz.cx + 4;
                        }
                    }
                }
            }

            SelectFont(hdc, hfPrev);
        }
        ReleaseDC(_hwndList, hdc);
    }

    return _cxMarlett;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_CreateBoldFont()
{
    if (!_hfBold)
    {
        HFONT hf = GetWindowFont(_hwndList);
        if (hf)
        {
            LOGFONT lf;
            if (GetObject(hf, sizeof(lf), &lf))
            {
                lf.lfWeight = FW_BOLD;
                SHAdjustLOGFONT(&lf); // locale-specific adjustments
                _hfBold = CreateFontIndirect(&lf);
            }
        }
    }
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_ReloadText()
{
    int iItem;
    for (iItem = ListView_GetItemCount(_hwndList) - 1; iItem >= 0; iItem--)
    {
        if (!_OnTextUpdate(iItem))
        {
            break;
        }
    }
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
void SFTBarHost::_RevalidateItems()
{
    if (!(_dwFlags & HOSTF_REVALIDATE))
    {
        return;
    }

    int iItem;
    for (iItem = ListView_GetItemCount(_hwndList) - 1; iItem >= 0 && _fEnumValid; iItem--)
    {
        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem)
        {
            pitem->Release();
        }
        else
        {
            _fEnumValid = 0;
        }
    }
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_RevalidatePostPopup()
{
    _RevalidateItems();

    if (_dwFlags & HOSTF_RELOADTEXT)
    {
        SetTimer(_hwnd, IDT_RELOADTEXT, 250, NULL);
    }
    // If the list is still good, then don't bother redoing it
    if (!_fEnumValid)
    {
        _EnumerateContents(FALSE);
    }
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
void SFTBarHost::_EnumerateContents(BOOL fUrgent)
{
#ifdef DEAD_CODE
    // If we have deferred refreshes until the window closes, then
    // leave it alone.
    if (!fUrgent && _fNeedsRepopulate)
    {
        return;
    }

    // If we're already enumerating, then just remember to do it again
    if (_fBGTask)
    {
        // accumulate urgency so a low-priority request + an urgent request
        // is treated as urgent
        _fRestartUrgent |= fUrgent;
        _fRestartEnum = TRUE;
        return;
    }

    _fRestartEnum = FALSE;
    _fRestartUrgent = FALSE;

    // If the list is still good, then don't bother redoing it
    if (_fEnumValid && !fUrgent)
    {
        return;
    }

    // This re-enumeration will make everything valid.
    _fEnumValid = TRUE;

    // Clear out all the leftover stuff from the previous enumeration

    _dpaEnumNew.EnumCallbackEx(PaneItem::DPAEnumCallback, (void *)NULL);
    _dpaEnumNew.DeleteAllPtrs();

    // Let client do some work on the foreground thread
    PrePopulate();

    // Finish the enumeration either on the background thread (if requested)
    // or on the foreground thread (if can't enumerate in the background).

    HRESULT hr;
    if (NeedBackgroundEnum())
    {
        if (_psched)
        {
            hr = S_OK;
        }
        else
        {
            // We need a separate task scheduler for each instance
            hr = CoCreateInstanceHook(CLSID_ShellTaskScheduler, NULL, CLSCTX_INPROC_SERVER,
                                  IID_PPV_ARGS(&_psched));
        }

        if (SUCCEEDED(hr))
        {
            CBGEnum *penum = new CBGEnum(this, fUrgent);
            if (penum)
            {

            // We want to run at a priority slightly above normal
            // because the user is sitting there waiting for the
            // enumeration to complete.
#define ITSAT_BGENUM_PRIORITY (ITSAT_DEFAULT_PRIORITY + 0x1000)

                hr = _psched->AddTask(penum, TOID_SFTBarHostBackgroundEnum, (DWORD_PTR)this, ITSAT_BGENUM_PRIORITY);
                if (SUCCEEDED(hr))
                {
                    _fBGTask = TRUE;

                    if (ListView_GetItemCount(_hwndList) == 0)
                    {
                        //
                        //  Set a timer that will create the "please wait"
                        //  animation if the enumeration takes too long.
                        //
                        _idtAni = IDT_ASYNCENUM;
                        SetTimer(_hwnd, _idtAni, 1000, NULL);
                    }
                }
                penum->Release();
            }
        }
    }

    if (!_fBGTask)
    {
        // Fallback: Do it on the foreground thread
        _EnumerateContentsBackground();
        _RepopulateList();
    }
#else
    if (!fUrgent && _fNeedsRepopulate)
        return;
    
    if (_fBGTask)
    {
        _fRestartUrgent |= fUrgent;
        _fRestartEnum = 1;
    }
    else
    {
        _fRestartEnum = 0;
        _fRestartUrgent = 0;
        if (!_fEnumValid || fUrgent)
        {
            _fEnumValid = 1;

            _dpaEnumNew.EnumCallback(PaneItem::DPAEnumCallback, 0);
            if (_dpaEnumNew)
            {
                _dpaEnumNew.DeleteAllPtrs();
            }

            PrePopulate();
            if (NeedBackgroundEnum() && (_psched || CoCreateInstance(CLSID_ShellTaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_psched)) >= 0))
            {
                CBGEnum *penum = new CBGEnum(this, fUrgent);
                if (penum)
                {
                    if (_psched->AddTask(penum, TOID_SFTBarHostBackgroundEnum, (DWORD_PTR)this, 0x10001000u) >= 0)
                    {
                        _fBGTask = 1;

                        if (ListView_GetItemCount(_hwndList) == 0)
                        {
                            _idtAni = IDT_ASYNCENUM;
                            SetTimer(_hwnd, _idtAni, 1000, NULL);
                        }
                    }
                    penum->Release();
                }
            }

            if (!this->_fBGTask)
            {
                _EnumerateContentsBackground();
                _RepopulateList();
            }
        }
    }
#endif
}


// EXEX-VISTA(allison): Partially validated. Figure out the CDPA sort call.
void SFTBarHost::_EnumerateContentsBackground()
{
    // Start over

    EnumItems();

#ifdef _ALPHA_
    // Alpha compiler is lame
    _dpaEnumNew.Sort((CDPA<PaneItem>::_PFNDPACOMPARE)_SortItemsAfterEnum, (LPARAM)this);
#else
    _dpaEnumNew.SortEx(_SortItemsAfterEnum, this);
#endif

	// PostEnum(); // EXEX-VISTA(allison): TODO: Uncomment when implemented.
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
int CALLBACK SFTBarHost::_SortItemsAfterEnum(PaneItem *p1, PaneItem *p2, SFTBarHost *self)
{

#ifdef DEAD_CODE
    //
    //  Put all pinned items (sorted by pin position) ahead of unpinned items.
    //
    if (p1->IsPinned())
    {
        if (p2->IsPinned())
        {
            return p1->GetPinPos() - p2->GetPinPos();
        }
        return -1;
    }
    else if (p2->IsPinned())
    {
        return +1;
    }

    //
    //  Both unpinned - let the client decide.
    //
    return self->CompareItems(p1, p2);
#else
    int iPinPos; // eax
    int v4; // ecx

    iPinPos = p1->_iPinPos;
    if (iPinPos < 0)
    {
        if (p2->_iPinPos < 0)
            return self->CompareItems(p1, p2);
        else
            return 1;
    }
    else
    {
        v4 = p2->_iPinPos;
        if (v4 < 0)
            return -1;
        else
            return iPinPos - v4;
    }
#endif
}

// EXEX-VISTA(allison): Validated.
SFTBarHost::~SFTBarHost()
{
    // We shouldn't be destroyed while in these temporary states.
    // If this fires, it's possible that somebody incremented
    // _fListUnstable/_fPopulating and forgot to decrement it.
    ASSERT(!_fListUnstable);
    ASSERT(!_fPopulating);

    ATOMICRELEASE(_pdth);
    ATOMICRELEASE(_pdsh);
    ATOMICRELEASE(_psched);
    ASSERT(_pdtoDragOut == NULL);

    _dpaEnum.DestroyCallbackEx(PaneItem::DPAEnumCallback, (void *)NULL);

    _dpaEnumNew.DestroyCallbackEx(PaneItem::DPAEnumCallback, (void *)NULL);

    if (_himl)
    {
        if (_IsPrivateImageList())
        {
            VARIANT vt = {0};
            vt.vt = VT_BYREF;
            IUnknown_QueryServiceExec(_punkSite, SID_SM_UserPane, &SID_SM_DV2ControlHost, 314, 0, &vt, 0);
        }
        ImageList_Destroy(_himl);
		_himl = NULL;
    }

    if (_hfList)
    {
        DeleteObject(_hfList);
    }

    if (_hfBold)
    {
        DeleteObject(_hfBold);
    }

    if (_hfMarlett)
    {
        DeleteObject(_hfMarlett);
    }

    if (_hBrushAni)
    {
        DeleteObject(_hBrushAni);
    }
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    UINT id;
    for (id = 0; id < SFTHOST_MAXNOTIFY; id++)
    {
        UnregisterNotify(id);
    }

    if (_hwndList)
    {
        RevokeDragDrop(_hwndList);
    }
    return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnNcDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // WARNING!  "this" might be NULL (if WM_NCCREATE failed).
    LRESULT lres = DefWindowProc(hwnd, uMsg, wParam, lParam);
    SetWindowLongPtr(hwnd, 0, 0);
    if (this)
    {
        _hwndList = NULL;
        _hwnd = NULL;
        if (_psched)
        {
            // Remove all tasks now, and wait for them to finish
            _psched->RemoveTasks(TOID_NULL, ITSAT_DEFAULT_LPARAM, TRUE);
            ATOMICRELEASE(_psched);
        }
        Release();
    }
    return lres;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_UpdateHotTrackRect()
{
    if (!_hTheme)
    {
        int iHotItem = ListView_GetHotItem(_hwndList);
        if (iHotItem >= 0)
        {
            RECT rc;
            if (ListView_GetItemRect(_hwndList, iHotItem, &rc, LVIR_SELECTBOUNDS))
            {
                InvalidateRect(_hwndList, &rc, TRUE);
            }
        }
    }
}

// EXEX-VISTA(allison): Validated. Still needs major cleanup.
LRESULT SFTBarHost::_OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
#ifdef DEAD_CODE
    LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);
    if (pnm->hwndFrom == _hwndList)
    {
        switch (pnm->code)
        {
        case NM_CUSTOMDRAW:
            return _OnLVCustomDraw(CONTAINING_RECORD(
                                   CONTAINING_RECORD(pnm, NMCUSTOMDRAW, hdr),
                                                          NMLVCUSTOMDRAW, nmcd));
        case NM_CLICK:
            return _OnLVNItemActivate(CONTAINING_RECORD(pnm, NMITEMACTIVATE, hdr));

        case NM_RETURN:
            return _ActivateItem(_GetLVCurSel(), AIF_KEYBOARD);

        case NM_KILLFOCUS:
            // On loss of focus, deselect all items so they all draw
            // in the plain state.
            ListView_SetItemState(_hwndList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            break;

        case LVN_GETINFOTIP:
            return _OnLVNGetInfoTip(CONTAINING_RECORD(pnm, NMLVGETINFOTIP, hdr));

        case LVN_BEGINDRAG:
        case LVN_BEGINRDRAG:
            return _OnLVNBeginDrag(CONTAINING_RECORD(pnm, NMLISTVIEW, hdr));

        case LVN_BEGINLABELEDIT:
            return _OnLVNBeginLabelEdit(CONTAINING_RECORD(pnm, NMLVDISPINFO, hdr));

        case LVN_ENDLABELEDIT:
            return _OnLVNEndLabelEdit(CONTAINING_RECORD(pnm, NMLVDISPINFO, hdr));

        case LVN_KEYDOWN:
            return _OnLVNKeyDown(CONTAINING_RECORD(pnm, NMLVKEYDOWN, hdr));
        }
    }
    else
    {
        switch (pnm->code)
        {
        case SMN_INITIALUPDATE:
            _EnumerateContents(FALSE);
            break;

        case SMN_POSTPOPUP:
            _RevalidatePostPopup();
            break;

        case SMN_GETMINSIZE:
            return _OnSMNGetMinSize(CONTAINING_RECORD(pnm, SMNGETMINSIZE, hdr));
            break;

        case SMN_FINDITEM:
            return _OnSMNFindItem(CONTAINING_RECORD(pnm, SMNDIALOGMESSAGE, hdr));
        case SMN_DISMISS:
            return _OnSMNDismiss();

        case SMN_APPLYREGION:
            return HandleApplyRegion(_hwnd, _hTheme, (SMNMAPPLYREGION *)lParam, _iThemePart, 0);

        case SMN_SHELLMENUDISMISSED:
            _iCascading = -1;
            return 0;
        }
    }

    // Give derived class a chance to respond
    return OnWndProc(hwnd, uMsg, wParam, lParam);
#else
    LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);
    if (pnm->hwndFrom == this->_hwndList)
    {
        if (pnm->code <= 0xFFFFFF91)
        {
            if (pnm->code != -111)
            {
                switch (pnm->code)
                {
                case 0xFFFFFF46:
                    return _OnLVNAsyncDrawn(CONTAINING_RECORD(pnm, NMLVASYNCDRAWN, hdr));
                case 0xFFFFFF50:
                    return _OnLVNEndLabelEdit(CONTAINING_RECORD(pnm, NMLVDISPINFO, hdr));
                case 0xFFFFFF51:
                    return _OnLVNBeginLabelEdit(CONTAINING_RECORD(pnm, NMLVDISPINFO, hdr));
                case 0xFFFFFF62:
                    return _OnLVNGetInfoTip(CONTAINING_RECORD(pnm, NMLVGETINFOTIP, hdr));
                case 0xFFFFFF65:
                    return _OnLVNKeyDown(CONTAINING_RECORD(pnm, NMLVKEYDOWN, hdr));
                case 0xFFFFFF87:
                    _UpdateHotTrackRect();
                    return 0;
                }
                return this->OnWndProc(hwnd, uMsg, wParam, lParam);
            }
            return _OnLVNBeginDrag(CONTAINING_RECORD(pnm, NMLISTVIEW, hdr));
        }
        if (pnm->code == -109)
        {
            return _OnLVNBeginDrag(CONTAINING_RECORD(pnm, NMLISTVIEW, hdr));
        }
        if (pnm->code == NM_CUSTOMDRAW)
        {
            return _OnLVCustomDraw(CONTAINING_RECORD(CONTAINING_RECORD(pnm, NMCUSTOMDRAW, hdr), NMLVCUSTOMDRAW, nmcd));
        }
        if (pnm->code == NM_KILLFOCUS)
        {
            this->_NotifyHoverImage(-1);
            goto L_DESELECT_ALL;
        }
        if (pnm->code == -4)
        {
            return _ActivateItem(_GetLVCurSel(), AIF_KEYBOARD);
        }
        if (pnm->code == -2)
        {
            return _OnLVNItemActivate(CONTAINING_RECORD(pnm, NMITEMACTIVATE, hdr));
        }
        return this->OnWndProc(hwnd, uMsg, wParam, lParam);
    }
    if (pnm->code <= 214)
    {
        if (pnm->code == 214)
        {
            if (this->_hwndList == GetFocus())
            {
                NotifyWinEvent(0x8005u, this->_hwndList, -4, this->_iCascading + 1);
            }
            this->_iCascading = -1;
            return 0;
        }
        if (pnm->code == 200) // SMN_INITIALUPDATE
        {
            _EnumerateContents(0);
        }
        else
        {
            if (pnm->code == 201) // SMN_APPLYREGION
            {
                if (this->_iThemePart != 6)
                {
                    return HandleApplyRegion(this->_hwnd, this->_hTheme, (SMNMAPPLYREGION *)lParam, this->_iThemePart, 0);
                }
            }
            else
            {
                if (pnm->code == 206) // SMN_GETMINSIZE
                {
                    return _OnSMNGetMinSize(CONTAINING_RECORD(pnm, SMNGETMINSIZE, hdr));
                }
                if (pnm->code == 208) // SMN_POSTPOPUP
                {
                    _RevalidatePostPopup();
                }
                else if (pnm->code == 210) // SMN_DISMISS
                {
                    return _OnSMNDismiss();
                }
            }
        }
        return this->OnWndProc(hwnd, uMsg, wParam, lParam);
    }
    if (pnm->code == 215)
    {
        return _OnSMNFindItem((SMNDIALOGMESSAGE *)lParam);
    }
    if (pnm->code == 223)
    {
        return SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite));
    }
    if (pnm->code == 224)
    {
        if (_iCascading != -1)
        {
            NotifyWinEvent(EVENT_OBJECT_FOCUS, _hwndList, -4, _iCascading + 1);
        }
        return OnWndProc(hwnd, uMsg, wParam, lParam);
    }
    if (pnm->code == 225)
    {
        _NotifyHoverImage(-1);
        return OnWndProc(hwnd, uMsg, wParam, lParam);
    }
    if (pnm->code == -8)
    {
    L_DESELECT_ALL:
        _UpdateHotTrackRect();
        ListView_SetItemState(_hwndList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
        return OnWndProc(hwnd, uMsg, wParam, lParam);
    }
    return OnWndProc(hwnd, uMsg, wParam, lParam);
#endif
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnTimer(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (wParam)
    {
    case IDT_ASYNCENUM:
        KillTimer(hwnd, wParam);

        // For some reason, we sometimes get spurious WM_TIMER messages,
        // so ignore them if we aren't expecting them.
        if (_idtAni)
        {
            _idtAni = 0;
            if (_hwndList && !_hwndAni)
            {
                DWORD dwStyle = WS_CHILD | WS_VISIBLE |
                                WS_CLIPCHILDREN | WS_CLIPSIBLINGS |
                                ACS_AUTOPLAY | ACS_TIMER | ACS_TRANSPARENT;

                RECT rcClient;
                GetClientRect(_hwnd, &rcClient);
                int x = (RECTWIDTH(rcClient) - ANIWND_WIDTH)/2;     // IDA_SEARCH is ANIWND_WIDTH pix wide
                int y = (RECTHEIGHT(rcClient) - ANIWND_HEIGHT)/2;    // IDA_SEARCH is ANIWND_HEIGHT pix tall

                _hwndAni = SHFusionCreateWindow(ANIMATE_CLASS, NULL, dwStyle,
                                                x, y, 0, 0,
                                                _hwnd, NULL,
                                                _AtlBaseModule.GetModuleInstance(), NULL);
                if (_hwndAni)
                {
                    NotifyWinEvent(EVENT_OBJECT_SHOW, _hwndAni, 0, 0);

                    SetWindowPos(_hwndAni, HWND_TOP, 0, 0, 0, 0,
                                 SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
                    #define IDA_SEARCH 150 // from shell32
                    Animate_OpenEx(_hwndAni, GetModuleHandle(TEXT("SHELL32")), MAKEINTRESOURCE(IDA_SEARCH));
                }
            }
        }
        return 0;
    case IDT_RELOADTEXT:
        KillTimer(hwnd, wParam);
        _ReloadText();
        break;

    case IDT_REFRESH:
        KillTimer(hwnd, wParam);
        PostMessage(hwnd, SFTBM_REFRESH, FALSE, 0);
        break;
    }

    // Give derived class a chance to respond
    return OnWndProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnSetFocus(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (_hwndList)
    {
        SetFocus(_hwndList);
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
HRESULT DrawPlacesListBackground(HTHEME hTheme, HWND hwnd, HDC hdc)
{
    HWND hwndParent = GetParent(hwnd);
    RECT rcParent;
    GetClientRect(hwndParent, &rcParent);

    HRESULT hr = S_OK;

    if (IsCompositionActive())
    {
        RECT rc;
        GetClipBox(hdc, &rc);
        SHFillRectClr(hdc, &rc, 0);
        hr = S_FALSE;
    }

    if (hwndParent)
    {
        POINT pt = { 0, 0 };
        SetBkMode(hdc, TRANSPARENT);

        MapWindowPoints(hwnd, hwndParent, &pt, 1);
        rcParent.right -= pt.x;
        OffsetWindowOrgEx(hdc, 0, pt.y, &pt);

        DrawThemeBackground(hTheme, hdc, SPP_PLACESLIST, 0, &rcParent, NULL);

        SetWindowOrgEx(hdc, pt.x, pt.y, NULL);
    }
    return hr;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnEraseBackground(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    RECT rc;
    GetClientRect(hwnd, &rc);
    if (_hTheme)
    {
        if (_iThemePart == SPP_PLACESLIST)
        {
            DrawPlacesListBackground(_hTheme, hwnd, (HDC)wParam);
        }
        else
        {
            if (IsCompositionActive())
            {
                SHFillRectClr((HDC)wParam, &rc, 0);
            }
            DrawThemeBackground(_hTheme, (HDC)wParam, _iThemePart, 0, &rc, NULL);
        }
    }
    else
    {
        SHFillRectClr((HDC)wParam, &rc, _clrBG);
        if (_iThemePart == SPP_PLACESLIST)                  // we set this even in non-theme case, its how we tell them apart
            DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_LEFT);
    }

    return TRUE;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnLVCustomDraw(LPNMLVCUSTOMDRAW plvcd)
{
    _DebugConsistencyCheck();

    switch (plvcd->nmcd.dwDrawStage)
    {
    case CDDS_PREPAINT:
        return _OnLVPrePaint(plvcd);

    case CDDS_ITEMPREPAINT:
        return _OnLVItemPrePaint(plvcd);

    case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
        return _OnLVSubItemPrePaint(plvcd);

    case CDDS_ITEMPOSTPAINT:
        return _OnLVItemPostPaint(plvcd);

    case CDDS_POSTPAINT:
        return _OnLVPostPaint(plvcd);
    }

    return CDRF_DODEFAULT;
}

//
//  Catch WM_PAINT messages headed to ListView and hide any drop effect
//  so it doesn't interfere with painting.  WM_PAINT messages might nest
//  under extreme conditions, so do this only at outer level.
//
// EXEX-VISTA(allison): Validated.
LRESULT CALLBACK SFTBarHost::s_DropTargetSubclassProc(
                             HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
                             UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    SFTBarHost *self = reinterpret_cast<SFTBarHost *>(dwRefData);
    LRESULT lres;

    switch (uMsg)
    {
    case WM_PAINT:

        // If entering outermost paint cycle, hide the drop feedback
        ++self->_cPaint;
        if (self->_cPaint == 1 && self->_pdth)
        {
            self->_pdth->Show(FALSE);
        }
        lres = DefSubclassProc(hwnd, uMsg, wParam, lParam);

        // If exiting outermost paint cycle, restore the drop feedback
        // Don't decrement _cPaint until really finished because
        // Show() will call UpdateWindow and trigger a nested paint cycle.
        if (self->_cPaint == 1 && self->_pdth)
        {
            self->_pdth->Show(TRUE);
        }
        --self->_cPaint;

        return lres;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, s_DropTargetSubclassProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

//
//  Listview makes it hard to detect whether you are in a real customdraw
//  or a fake customdraw, since it frequently "gets confused" and gives
//  you a 0x0 rectangle even though it really wants you to draw something.
//
//  Even worse, within a single paint cycle, Listview uses multiple
//  NMLVCUSTOMDRAW structures so you can't stash state inside the customdraw
//  structure.  You have to save it externally.
//
//  The only trustworthy guy is CDDS_PREPAINT.  Use his rectangle to
//  determine whether this is a real or fake customdraw...
//
//  What's even weirder is that inside a regular paint cycle, you
//  can get re-entered with a sub-paint cycle, so we have to maintain
//  a stack of "is the current customdraw cycle real or fake?" bits.

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_CustomDrawPush(BOOL fReal)
{
    _dwCustomDrawState = (_dwCustomDrawState << 1) | fReal;
}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::_IsRealCustomDraw()
{
    return _dwCustomDrawState & 1;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_CustomDrawPop()
{
    _dwCustomDrawState >>= 1;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnLVPrePaint(LPNMLVCUSTOMDRAW plvcd)
{
    LRESULT lResult;

    // Always ask for postpaint so we can maintain our customdraw stack
    lResult = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
    BOOL fReal = !IsRectEmpty(&plvcd->nmcd.rc);
    _CustomDrawPush(fReal);

    return lResult;
}

//
//  Hack!  We want to know in _OnLvSubItemPrePaint whether the item
//  is selected or not,  We borrow the CDIS_CHECKED bit, which is
//  otherwise used only by toolbar controls.
//
#define CDIS_WASSELECTED        CDIS_CHECKED

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnLVItemPrePaint(LPNMLVCUSTOMDRAW plvcd)
{
#ifdef DEAD_CODE
    LRESULT lResult = CDRF_DODEFAULT;

    plvcd->nmcd.uItemState &= ~CDIS_WASSELECTED;

    if (GetFocus() == _hwndList &&
        (plvcd->nmcd.uItemState & CDIS_SELECTED))
    {
        plvcd->nmcd.uItemState |= CDIS_WASSELECTED;

        // menu-highlighted tiles are always opaque
        if (_hTheme)
        {
            plvcd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
            plvcd->clrFace = plvcd->clrTextBk = GetSysColor(COLOR_MENUHILIGHT);
        }
        else
        {
            plvcd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
            plvcd->clrFace = plvcd->clrTextBk = GetSysColor(COLOR_HIGHLIGHT);
        }
    }

    // Turn off CDIS_SELECTED because it causes the icon to get alphablended
    // and we don't want that.  Turn off CDIS_FOCUS because that draws a
    // focus rectangle and we don't want that either.

    plvcd->nmcd.uItemState &= ~(CDIS_SELECTED | CDIS_FOCUS);

    //
    if (plvcd->nmcd.uItemState & CDIS_HOT && _clrHot != CLR_NONE)
        plvcd->clrText = _clrHot;

    // Turn off selection highlighting for everyone except
    // the drop target highlight
    if ((int)plvcd->nmcd.dwItemSpec != _iDragOver || !_pdtDragOver)
    {
        lResult |= LVCDRF_NOSELECT;
    }

    PaneItem *pitem = _GetItemFromLVLParam(plvcd->nmcd.lItemlParam);
    if (!pitem)
    {
        // Sometimes ListView doesn't give us an lParam so we have to
        // get it ourselves
        pitem = _GetItemFromLV((int)plvcd->nmcd.dwItemSpec);
    }

    if (pitem)
    {
        if (IsBold(pitem))
        {
            _CreateBoldFont();
            SelectFont(plvcd->nmcd.hdc, _hfBold);
            lResult |= CDRF_NEWFONT;
        }
        if (pitem->IsCascade())
        {
            // Need subitem notification because that's what sets the colors
            lResult |= CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYSUBITEMDRAW;
        }
        if (pitem->HasAccelerator())
        {
            // Need subitem notification because that's what sets the colors
            lResult |= CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYSUBITEMDRAW;
        }
        if (pitem->HasSubtitle())
        {
            lResult |= CDRF_NOTIFYSUBITEMDRAW;
        }
    }
    return lResult;
#else
    COLORREF SysColor; // eax
    COLORREF v4; // eax

    LRESULT lResult = CDRF_DODEFAULT;

    plvcd->nmcd.uItemState &= ~8u;

    if ((plvcd->nmcd.uItemState & 0x41) != 0 || plvcd->nmcd.dwItemSpec == ListView_GetHotItem(_hwndList))
    {
        plvcd->nmcd.uItemState |= 8u;
        if (_hTheme)
        {
            SysColor = GetSysColor(COLOR_MENUHILIGHT);
        }
        else
        {
            plvcd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
            SysColor = GetSysColor(COLOR_HIGHLIGHT);
        }
        plvcd->clrFace = SysColor;
        plvcd->clrTextBk = SysColor;
    }

    plvcd->nmcd.uItemState &= 0xFFFFFFEE;
    if ((plvcd->nmcd.uItemState & 0x40) != 0 && this->_clrHot != -1
        || plvcd->nmcd.dwItemSpec == SendMessageW(this->_hwndList, 0x103Du, 0, 0))
    {
        v4 = GetSysColor(COLOR_HIGHLIGHT);
        plvcd->clrTextBk = v4;
        plvcd->clrFace = v4;
        plvcd->clrText = this->_clrHot;
    }

    if (plvcd->nmcd.dwItemSpec != this->_iDragOver || !this->_pdtDragOver)
        lResult = 0x10000;

    PaneItem* pitem = _GetItemFromLVLParam(plvcd->nmcd.lItemlParam);
    if (pitem || (pitem = _GetItemFromLV(plvcd->nmcd.dwItemSpec)) != 0)
    {
        if (IsBold(pitem))
        {
            _CreateBoldFont();
            SelectObject(plvcd->nmcd.hdc, _hfBold);
            lResult |= 2;
        }

        if ((pitem->_dwFlags & 1) != 0)
            lResult |= 0x30u;
        if (pitem->_pszAccelerator)
            lResult |= 0x30u;
        if ((pitem->_dwFlags & 2) != 0)
            lResult |= 0x20u;
        pitem->Release();
    }
    return lResult;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnLVSubItemPrePaint(LPNMLVCUSTOMDRAW plvcd)
{
#ifdef DEAD_CODE
    LRESULT lResult = CDRF_DODEFAULT;
    if (plvcd->iSubItem == 1)
    {
        // Second line uses the regular font (first line was bold)
        SelectFont(plvcd->nmcd.hdc, GetWindowFont(_hwndList));
        lResult |= CDRF_NEWFONT;

        if (GetFocus() == _hwndList &&
            (plvcd->nmcd.uItemState & CDIS_WASSELECTED))
        {
            plvcd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
        }
        else
        // Maybe there's a custom subtitle color
        if (_clrSubtitle != CLR_NONE)
        {
            plvcd->clrText = _clrSubtitle;
        }
        else
        {
            plvcd->clrText = GetSysColor(COLOR_MENUTEXT);
        }
    }
    return lResult;
#else
    LRESULT lResult = CDRF_DODEFAULT;
    if (plvcd->iSubItem == 1)
    {
        SelectFont(plvcd->nmcd.hdc, GetWindowFont(_hwndList));
        lResult |= CDRF_NEWFONT;
        
        if (!_hTheme && (plvcd->nmcd.uItemState & CDIS_WASSELECTED))
        {
            plvcd->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
        }
        else if (_clrSubtitle != CLR_NONE)
        {
            plvcd->clrText = _clrSubtitle;
        }
        else
        {
            plvcd->clrText = GetSysColor(COLOR_MENUTEXT);
        }
    }
    return lResult;
#endif
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

// QUIRK!  Listview often sends item postpaint messages even though we
// didn't ask for one.  It does this because we set NOTIFYPOSTPAINT on
// the CDDS_PREPAINT notification ("please notify me when the entire
// listview is finished painting") and it thinks that that flag also
// turns on postpaint notifications for each item...

// EXEX-VISTA(allison): Validated. Still needs major cleanup.
LRESULT SFTBarHost::_OnLVItemPostPaint(LPNMLVCUSTOMDRAW plvcd)
{
#ifdef DEAD_CODE
    PaneItem *pitem = _GetItemFromLVLParam(plvcd->nmcd.lItemlParam);
    if (_IsRealCustomDraw() && pitem)
    {
        RECT rc;
        if (ListView_GetItemRect(_hwndList, plvcd->nmcd.dwItemSpec, &rc, LVIR_LABEL))
        {
            COLORREF clrBkPrev = SetBkColor(plvcd->nmcd.hdc, plvcd->clrFace);
            COLORREF clrTextPrev = SetTextColor(plvcd->nmcd.hdc, plvcd->clrText);
            int iModePrev = SetBkMode(plvcd->nmcd.hdc, TRANSPARENT);
            BOOL fRTL = GetLayout(plvcd->nmcd.hdc) & LAYOUT_RTL;

            if (pitem->IsCascade())
            {
                {
                    HFONT hfPrev = SelectFont(plvcd->nmcd.hdc, _hfMarlett);
                    if (hfPrev)
                    {
                        TCHAR chOut = fRTL ? TEXT('w') : TEXT('8');
                        UINT fuOptions = 0;
                        if (fRTL)
                        {
                            fuOptions |= ETO_RTLREADING;
                        }

                        ExtTextOut(plvcd->nmcd.hdc, rc.right - _cxMarlett,
                                   rc.top + (rc.bottom - rc.top - _tmAscentMarlett)/2,
                                   fuOptions, &rc, &chOut, 1, NULL);
                        SelectFont(plvcd->nmcd.hdc, hfPrev);
                    }
                }
            }

            if (pitem->HasAccelerator() &&
                (plvcd->nmcd.uItemState & CDIS_SHOWKEYBOARDCUES))
            {
                // Subtitles mess up our computations...
                ASSERT(!pitem->HasSubtitle());

                rc.right -= _cxMarlett; // Subtract out our margin

                UINT uFormat = DT_VCENTER | DT_SINGLELINE | DT_PREFIXONLY |
                               DT_WORDBREAK | DT_EDITCONTROL | DT_WORD_ELLIPSIS;
                if (fRTL)
                {
                    uFormat |= DT_RTLREADING;
                }

                DrawText(plvcd->nmcd.hdc, pitem->_pszAccelerator, -1, &rc, uFormat);
                rc.right += _cxMarlett; // restore it
            }

            SetBkMode(plvcd->nmcd.hdc, iModePrev);
            SetTextColor(plvcd->nmcd.hdc, clrTextPrev);
            SetBkColor(plvcd->nmcd.hdc, clrBkPrev);
        }
    }

    return CDRF_DODEFAULT;
#else
    int fRTL; // edi MAPDST
    UINT fuOptions; // eax
    DTTOPTS opts; // [esp+10h] [ebp-9Ch] BYREF
    COLORREF clrBkPrev; // [esp+50h] [ebp-5Ch]
    COLORREF clrTextPrev; // [esp+54h] [ebp-58h]
    int iModePrev; // [esp+58h] [ebp-54h]
    HFONT hfPrev; // [esp+5Ch] [ebp-50h]
    RECT pRect; // [esp+60h] [ebp-4Ch] BYREF
    DWORD dwTextFlags; // [esp+74h] [ebp-38h]
    PaneItem *pitem; // [esp+80h] [ebp-2Ch] MAPDST
    RECT rc; // [esp+84h] [ebp-28h] BYREF
    //CPPEH_RECORD ms_exc; // [esp+94h] [ebp-18h]
    UINT uFormat; // [esp+B4h] [ebp+8h]

    pitem = SFTBarHost::_GetItemFromLVLParam(plvcd->nmcd.lItemlParam);
    if (pitem)
    {
        if ((this->_dwCustomDrawState & 1) != 0)
        {
            if (ListView_GetItemRect(_hwndList, plvcd->nmcd.dwItemSpec, &rc, LVIR_LABEL))
            {
                clrBkPrev = SetBkColor(plvcd->nmcd.hdc, plvcd->clrFace);
                clrTextPrev = SetTextColor(plvcd->nmcd.hdc, plvcd->clrText);
                iModePrev = SetBkMode(plvcd->nmcd.hdc, 1);
                fRTL = GetLayout(plvcd->nmcd.hdc) & 1;
                if ((pitem->_dwFlags & 1) != 0)
                {
                    hfPrev = (HFONT)SelectObject(plvcd->nmcd.hdc, this->_hfMarlett);
                    if (hfPrev)
                    {
                        WCHAR chOut[2] = { fRTL ? TEXT('w') : TEXT('8') };
                        dwTextFlags = 0;
                        fuOptions = 0;
                        if (fRTL)
                        {
                            fuOptions = 128;
                            dwTextFlags = 0x20000;
                        }

                        if (this->_hTheme)
                        {
                            opts.dwSize = 64;
                            memset(&opts.dwFlags, 0, 0x3Cu);
                            opts.dwFlags = IsCompositionActive() ? 0x2000 : 0;
                            pRect.left = rc.right - this->_cxMarlett;
                            pRect.top = rc.top + (rc.bottom - this->_tmAscentMarlett - rc.top) / 2;
                            pRect.right = rc.right;
                            pRect.bottom = rc.bottom;
                            DrawThemeTextEx(
                                this->_hTheme,
                                plvcd->nmcd.hdc,
                                this->_iThemePart,
                                plvcd->iStateId,
                                chOut,
                                1,
                                dwTextFlags,
                                &pRect,
                                &opts);
                        }
                        else
                        {
                            SHExtTextOutW(
                                plvcd->nmcd.hdc,
                                rc.right - this->_cxMarlett,
                                rc.top + (rc.bottom - this->_tmAscentMarlett - rc.top) / 2,
                                fuOptions,
                                &rc,
                                chOut,
                                1u,
                                0);
                        }
                        SelectObject(plvcd->nmcd.hdc, hfPrev);
                    }
                }
                if (pitem->_pszAccelerator && (plvcd->nmcd.uItemState & 0x200) != 0)
                {
                    ASSERT(!pitem->HasSubtitle()); // 2159
                    rc.right -= this->_cxMarlett;
                    rc.left += 2;

                    uFormat = 0x42010;
                    if (fRTL)
                        uFormat = 0x62010;
                    pRect = rc;
                    DrawTextW(plvcd->nmcd.hdc, pitem->_pszAccelerator, -1, &pRect, uFormat | 0x400);
                    rc.top += (rc.bottom + pRect.top - pRect.bottom - rc.top) / 2;
                    if (this->_hTheme)
                    {
                        opts.dwSize = 0x40;
                        memset(&opts.dwFlags, 0, 0x3Cu);
                        opts.dwFlags = IsCompositionActive() ? DTT_COMPOSITED : 0;
                        DrawThemeTextEx(
                            this->_hTheme,
                            plvcd->nmcd.hdc,
                            this->_iThemePart,
                            plvcd->iStateId,
                            pitem->_pszAccelerator,
                            -1,
                            uFormat | 0x200000,
                            &rc,
                            &opts);
                    }
                    else
                    {
                        DrawTextW(plvcd->nmcd.hdc, pitem->_pszAccelerator, -1, &rc, uFormat | 0x200000);
                    }
                    rc.right += this->_cxMarlett;
                }
                SetBkMode(plvcd->nmcd.hdc, iModePrev);
                SetTextColor(plvcd->nmcd.hdc, clrTextPrev);
                SetBkColor(plvcd->nmcd.hdc, clrBkPrev);
            }
        }
        pitem->Release();
    }
    return 0;
#endif
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnLVPostPaint(LPNMLVCUSTOMDRAW plvcd)
{
    if (_IsRealCustomDraw())
    {
        _DrawInsertionMark(plvcd);
        _DrawSeparators(plvcd);
    }
    _CustomDrawPop();
    return CDRF_DODEFAULT;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnUpdateUIState(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // Only need to do this when the Start Menu is visible; if not visible, then
    // don't waste your time invalidating useless rectangles (and paging them in!)
    if (IsWindowVisible(GetAncestor(_hwnd, GA_ROOT)))
    {
        // All UIS_SETs should happen when the Start Menu is hidden;
        // we assume that the only thing we will be asked to do is to
        // start showing the underlines

        ASSERT(LOWORD(wParam) != UIS_SET);

        DWORD dwLvExStyle = 0;

        if (!GetSystemMetrics(SM_REMOTESESSION) && !GetSystemMetrics(SM_REMOTECONTROL))
        {
            dwLvExStyle |= LVS_EX_DOUBLEBUFFER;
        }

        if ((ListView_GetExtendedListViewStyle(_hwndList) & LVS_EX_DOUBLEBUFFER) != dwLvExStyle)
        {
            ListView_SetExtendedListViewStyleEx(_hwndList, LVS_EX_DOUBLEBUFFER, dwLvExStyle);
        }

        int iItem;
        for (iItem = ListView_GetItemCount(_hwndList) - 1; iItem >= 0; iItem--)
        {
            PaneItem *pitem = _GetItemFromLV(iItem);
            if (pitem && pitem->HasAccelerator())
            {
                RECT rc;
                if (ListView_GetItemRect(_hwndList, iItem, &rc, LVIR_LABEL))
                {
                    // We need to repaint background because of cleartype double print issues
                    InvalidateRect(_hwndList, &rc, TRUE);
                }
            }
			pitem->Release();
        }
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
PaneItem *SFTBarHost::_GetItemFromLV(int iItem)
{
    LVITEM lvi;
    lvi.iItem = iItem;
    lvi.iSubItem = 0;
    lvi.mask = LVIF_PARAM;
    if (iItem >= 0 && ListView_GetItem(_hwndList, &lvi))
    {
        PaneItem *pitem = _GetItemFromLVLParam(lvi.lParam);
        return pitem;
    }
    return NULL;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnMenuMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LRESULT lres;
    if (_pcm3Pop && SUCCEEDED(_pcm3Pop->HandleMenuMsg2(uMsg, wParam, lParam, &lres)))
    {
        return lres;
    }

    if (_pcm2Pop && SUCCEEDED(_pcm2Pop->HandleMenuMsg(uMsg, wParam, lParam)))
    {
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnForwardMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    SHPropagateMessage(hwnd, uMsg, wParam, lParam, SPM_SEND | SPM_ONELEVEL);
    // Give derived class a chance to get the message, too
    return OnWndProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::UnregisterNotify(UINT id)
{
    ASSERT(id < SFTHOST_MAXNOTIFY);

    if (id < SFTHOST_MAXNOTIFY && _rguChangeNotify[id])
    {
        UINT uChangeNotify = _rguChangeNotify[id];
        _rguChangeNotify[id] = 0;
        return SHChangeNotifyDeregister(uChangeNotify);
    }
    return FALSE;
}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::_RegisterNotify(UINT id, LONG lEvents, PCIDLIST_ABSOLUTE pidl, BOOL fRecursive)
{
    ASSERT(id < SFTHOST_MAXNOTIFY);

    if (id < SFTHOST_MAXNOTIFY)
    {
        UnregisterNotify(id);

        SHChangeNotifyEntry fsne;
        fsne.fRecursive = fRecursive;
        fsne.pidl = pidl;

        int fSources = SHCNRF_NewDelivery | SHCNRF_ShellLevel | SHCNRF_InterruptLevel;
        if (fRecursive)
        {
            // SHCNRF_RecursiveInterrupt means "Please use a recursive FindFirstChangeNotify"
            fSources |= SHCNRF_RecursiveInterrupt;
        }
        _rguChangeNotify[id] = SHChangeNotifyRegister(_hwnd, fSources, lEvents,
                                                      SFTBM_CHANGENOTIFY + id, 1, &fsne);
        return _rguChangeNotify[id];
    }
    return FALSE;
}

//
//  wParam = 0 if this is not an urgent refresh (can be postponed)
//  wParam = 1 if this is urgent (must refresh even if menu is open)
//

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnRepopulate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    // Don't update the list now if we are visible, except if the list was empty
    _fBGTask = FALSE;

    if (wParam || !IsWindowVisible(_hwnd) || ListView_GetItemCount(_hwndList) == 0)
    {
        _RepopulateList();
    }
    else
    {
        _fNeedsRepopulate = TRUE;
    }

    if (_fRestartEnum)
    {
        _EnumerateContents(_fRestartUrgent);
    }

    return 0;
}

// EXEX-VISTA(allison): Partially validated. Recheck flow.
LRESULT SFTBarHost::_OnChangeNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPITEMIDLIST *ppidl;
    LONG lEvent;
    LPSHChangeNotificationLock pshcnl;
    pshcnl = SHChangeNotification_Lock((HANDLE)wParam, (DWORD)lParam, &ppidl, &lEvent);

    if (pshcnl)
    {
        UINT id = uMsg - SFTBM_CHANGENOTIFY;
        if (id < SFTHOST_MAXCLIENTNOTIFY)
        {
            OnChangeNotify(id, lEvent, ppidl[0], ppidl[1]);
        }
        else if (id == SFTHOST_HOSTNOTIFY_UPDATEIMAGE)
        {
            _OnUpdateImage(ppidl[0], ppidl[1]);
        }
        else
        {
            // Our wndproc shouldn't have dispatched to us
            ASSERT(0);
        }

        SHChangeNotification_Unlock(pshcnl);
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_OnUpdateImage(LPCITEMIDLIST pidl, LPCITEMIDLIST pidlExtra)
{
    // Must use pidl and not pidlExtra because pidlExtra is sometimes NULL
    SHChangeDWORDAsIDList *pdwidl = (SHChangeDWORDAsIDList *)pidl;
    if (pdwidl->dwItem1 == 0xFFFFFFFF)
    {
        _SendNotify(_hwnd, 225, 0);
    }
    else
    {
        int iImage = SHHandleUpdateImage(pidlExtra);
        if (iImage >= 0)
        {
            UpdateImage(iImage);
        }
    }
}

//
//  See if anybody is using this image; if so, invalidate the cached bitmap.
//

// EXEX-VISTA(allison): Validated.
void SFTBarHost::UpdateImage(int iImage)
{
	// Only cache bitmaps for small or large icons
    if (_iconsize != ICONSIZE_MEDIUM)
    {
        int iItem;
        for (iItem = ListView_GetItemCount(_hwndList) - 1; iItem >= 0; iItem--)
        {
            LVITEM lvi;
            lvi.iItem = iItem;
            lvi.iSubItem = 0;
            lvi.mask = LVIF_IMAGE;
            if (ListView_GetItem(_hwndList, &lvi) && lvi.iImage == iImage)
            {
                // The cached bitmap is no good; an icon changed
                _SendNotify(_hwnd, SMN_NEEDREPAINT, NULL);
                break;
            }
        }
    }
}

//
//  wParam = 0 if this is not an urgent refresh (can be postponed)
//  wParam = 1 if this is urgen (must refresh even if menu is open)
//
LRESULT SFTBarHost::_OnRefresh(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    _EnumerateContents((BOOL)wParam);
    return 0;
}

// EXEX-VISTA(allison): Validated.
LPTSTR _DisplayNameOf(IShellFolder *psf, LPCITEMIDLIST pidl, UINT shgno)
{
    LPWSTR pszOut = 0;

    LPITEMIDLIST pidlOut;
    IShellFolder *psfOut;
    if (SHBindToFolderIDListParent(psf, pidl, IID_PPV_ARGS(&psfOut), (LPCITEMIDLIST *)&pidlOut) >= 0)
    {
        DisplayNameOfAsString(psfOut, pidlOut, shgno, &pszOut);
        psfOut->Release();
    }
    return pszOut;
}

// EXEX-VISTA(allison): Validated.
LPTSTR SFTBarHost::_DisplayNameOfItem(PaneItem *pitem, UINT shgno)
{
    IShellFolder *psf;
    LPCITEMIDLIST pidl;
    LPTSTR pszOut = NULL;

    if (SUCCEEDED(_GetFolderAndPidl(pitem, &psf, &pidl)))
    {
        pszOut = DisplayNameOfItem(pitem, psf, pidl, (SHGDNF)shgno);
        psf->Release();
    }
    return pszOut;
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::_GetUIObjectOfItem(PaneItem *pitem, REFIID riid, void * *ppv)
{
    *ppv = NULL;

    IShellFolder *psf;
    LPCITEMIDLIST pidlItem;
    HRESULT hr = _GetFolderAndPidl(pitem, &psf, &pidlItem);
    if (SUCCEEDED(hr))
    {
        hr = psf->GetUIObjectOf(_hwnd, 1, &pidlItem, riid, NULL, ppv);
        psf->Release();
    }

    return hr;
}

HRESULT SFTBarHost::_GetUIObjectOfItem(int iItem, REFIID riid, void * *ppv)
{
    PaneItem *pitem = _GetItemFromLV(iItem);
    if (pitem)
    {
        HRESULT hr = _GetUIObjectOfItem(pitem, riid, ppv);
        return hr;
    }
    return E_FAIL;
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::_GetFolderAndPidl(PaneItem *pitem, IShellFolder **ppsfOut, LPCITEMIDLIST *ppidlOut)
{
    *ppsfOut = NULL;
    *ppidlOut = NULL;
    return pitem->IsSeparator() ? E_FAIL : GetFolderAndPidl(pitem, ppsfOut, ppidlOut);
}

HRESULT SFTBarHost::_GetFolderAndPidlForActivate(PaneItem *pitem, IShellFolder **ppsfOut, LPCITEMIDLIST *ppidlOut)
{
    *ppsfOut = NULL;
    *ppidlOut = NULL;
    return pitem->IsSeparator() ? E_FAIL : GetFolderAndPidlForActivate(pitem, ppsfOut, ppidlOut);
}

//
//  Given the coordinates of a context menu (lParam from WM_CONTEXTMENU),
//  determine which item's context menu should be activated, or -1 if the
//  context menu is not for us.
//
//  Also, returns on success in *ppt the coordinates at which the
//  context menu should be displayed.
//

// EXEX-VISTA(allison): Validated.
int SFTBarHost::_ContextMenuCoordsToItem(LPARAM lParam, POINT *ppt)
{
    int iItem;
    ppt->x = GET_X_LPARAM(lParam);
    ppt->y = GET_Y_LPARAM(lParam);

    // If initiated from keyboard, act like they clicked on the center
    // of the focused icon.
    if (IS_WM_CONTEXTMENU_KEYBOARD(lParam))
    {
        iItem = _GetLVCurSel();
        if (iItem >= 0)
        {
            RECT rc;
            if (ListView_GetItemRect(_hwndList, iItem, &rc, LVIR_ICON))
            {
                MapWindowRect(_hwndList, NULL, &rc);
                ppt->x = (rc.left+rc.right)/2;
                ppt->y = (rc.top+rc.bottom)/2;
            }
            else
            {
                iItem = -1;
            }
        }
    }
    else
    {
        // Initiated from mouse; find the item they clicked on
        LVHITTESTINFO hti;
        hti.pt = *ppt;
        MapWindowPoints(NULL, _hwndList, &hti.pt, 1);
        iItem = ListView_HitTest(_hwndList, &hti);
    }

    return iItem;
}

LRESULT SFTBarHost::_OnContextMenu(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if(_AreChangesRestricted())
    {
        return 0;
    }

    TCHAR szBuf[MAX_PATH];
    _DebugConsistencyCheck();

    BOOL fSuccess = FALSE;

    POINT pt;
    int iItem = _ContextMenuCoordsToItem(lParam, &pt);

    if (iItem >= 0)
    {
        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem)
        {
            // If we can't get the official shell context menu,
            // then use a dummy one.
            IContextMenu *pcm;
            if (FAILED(_GetUIObjectOfItem(pitem, IID_PPV_ARGS(&pcm))))
            {
                pcm = s_EmptyContextMenu.GetContextMenu();
            }

            HMENU hmenu = ::CreatePopupMenu();

            if (hmenu)
            {
                UINT uFlags = CMF_NORMAL;
                if (GetKeyState(VK_SHIFT) < 0)
                {
                    uFlags |= CMF_EXTENDEDVERBS;
                }

                if (_dwFlags & HOSTF_CANRENAME)
                {
                    uFlags |= CMF_CANRENAME;
                }

                pcm->QueryContextMenu(hmenu, 0, IDM_QCM_MIN, IDM_QCM_MAX, uFlags);

                // Remove "Create shortcut" from context menu because it creates
                // the shortcut on the desktop, which the user can't see...
                ContextMenu_DeleteCommandByName(pcm, hmenu, IDM_QCM_MIN, TEXT("link"));

                // Remove "Cut" from context menu because we don't want objects
                // to be deleted.
                ContextMenu_DeleteCommandByName(pcm, hmenu, IDM_QCM_MIN, TEXT("cut"));

                // Let clients override the "delete" option.

                // Change "Delete" to "Remove from this list".
                // If client doesn't support "delete" then nuke it outright.
                // If client supports "delete" but the IContextMenu didn't create one,
                // then create a fake one so we cn add the "Remove from list" option.
                UINT uPosDelete = GetMenuIndexForCanonicalVerb(hmenu, pcm, IDM_QCM_MIN, TEXT("delete"));
                UINT uiFlags = 0;
                UINT idsDelete = AdjustDeleteMenuItem(pitem, &uiFlags);
                if (idsDelete)
                {
                    if (LoadString(_AtlBaseModule.GetResourceInstance(), idsDelete, szBuf, ARRAYSIZE(szBuf)))
                    {
                        if (uPosDelete != -1)
                        {
                            ModifyMenu(hmenu, uPosDelete, uiFlags | MF_BYPOSITION | MF_STRING, IDM_REMOVEFROMLIST, szBuf);
                        }
                        else
                        {
                            AppendMenu(hmenu, MF_SEPARATOR, -1, NULL);
                            AppendMenu(hmenu, uiFlags | MF_STRING, IDM_REMOVEFROMLIST, szBuf);
                        }
                    }
                }
                else
                {
                    DeleteMenu(hmenu, uPosDelete, MF_BYPOSITION);
                }

                _SHPrettyMenu(hmenu);

                ASSERT(_pcm2Pop == NULL);   // Shouldn't be recursing
                pcm->QueryInterface(IID_PPV_ARGS(&_pcm2Pop));

                ASSERT(_pcm3Pop == NULL);   // Shouldn't be recursing
                pcm->QueryInterface(IID_PPV_ARGS(&_pcm3Pop));

                int idCmd = TrackPopupMenuEx(hmenu,
                    TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_LEFTALIGN,
                    pt.x, pt.y, hwnd, NULL);

                ATOMICRELEASE(_pcm2Pop);
                ATOMICRELEASE(_pcm3Pop);

                if (idCmd)
                {
                    switch (idCmd)
                    {
                    case IDM_REMOVEFROMLIST:
                        StrCpyN(szBuf, TEXT("delete"), ARRAYSIZE(szBuf));
                        break;

                    default:
                        ContextMenu_GetCommandStringVerb(pcm, idCmd - IDM_QCM_MIN, szBuf, ARRAYSIZE(szBuf));
                        break;
                    }

                    idCmd -= IDM_QCM_MIN;

                    CMINVOKECOMMANDINFOEX ici = {
                        sizeof(ici),            // cbSize
                        CMIC_MASK_FLAG_LOG_USAGE | // this was an explicit user action
                        CMIC_MASK_ASYNCOK,      // fMask
                        hwnd,                   // hwnd
                        (LPCSTR)IntToPtr(idCmd),// lpVerb
                        NULL,                   // lpParameters
                        NULL,                   // lpDirectory
                        SW_SHOWDEFAULT,         // nShow
                        0,                      // dwHotKey
                        0,                      // hIcon
                        NULL,                   // lpTitle
                        (LPCWSTR)IntToPtr(idCmd),// lpVerbW
                        NULL,                   // lpParametersW
                        NULL,                   // lpDirectoryW
                        NULL,                   // lpTitleW
                        { pt.x, pt.y },         // ptInvoke
                    };

                    if ((_dwFlags & HOSTF_CANRENAME) &&
                        StrCmpI(szBuf, TEXT("rename")) == 0)
                    {
                        _EditLabel(iItem);
                    }
                    else
                    {
                        ContextMenuInvokeItem(pitem, pcm, &ici, szBuf);
                    }
                }

                DestroyMenu(hmenu);

                fSuccess = TRUE;
            }
            pcm->Release();
        }

    }

    _DebugConsistencyCheck();

    return fSuccess ? 0 : DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_EditLabel(int iItem)
{
    _fAllowEditLabel = TRUE;
    ListView_EditLabel(_hwndList, iItem);
    _fAllowEditLabel = FALSE;
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::ContextMenuInvokeItem(PaneItem *pitem, IContextMenu *pcm, CMINVOKECOMMANDINFOEX *pici, LPCTSTR pszVerb)
{
    // Make sure none of our private menu items leaked through
    ASSERT(PtrToLong(pici->lpVerb) >= 0);

    // FUSION: When we call out to 3rd party code we want it to use 
    // the process default context. This means that the 3rd party code will get
    // v5 in the explorer process. However, if shell32 is hosted in a v6 process,
    // then the 3rd party code will still get v6. 
    ULONG_PTR cookie = 0;
    ActivateActCtx(NULL, &cookie); 

	// SetICIKeyModifiers(&pici->fMask); // EXEX-VISTA(allison): TODO: Uncomment when implemented.

    IUnknown_SetSite(pcm, SAFECAST(this, IServiceProvider *));
    HRESULT hr = pcm->InvokeCommand(reinterpret_cast<LPCMINVOKECOMMANDINFO>(pici));
	IUnknown_SetSite(pcm, NULL);
    
    if (cookie != 0)
    {
        DeactivateActCtx(0, cookie);
    }

    return hr;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnLVNItemActivate(LPNMITEMACTIVATE pnmia)
{
    return _ActivateItem(pnmia->iItem, 0);
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
LRESULT SFTBarHost::_ActivateItem(int iItem, DWORD dwFlags)
{
#ifdef DEAD_CODE
    PaneItem *pitem;
    IShellFolder *psf;
    LPCITEMIDLIST pidl;

    DWORD dwCascadeFlags = 0;
    if (dwFlags & AIF_KEYBOARD)
    {
        dwCascadeFlags = MPPF_KEYBOARD | MPPF_INITIALSELECT;
    }

    if (_OnCascade(iItem, dwCascadeFlags))
    {
        // We did the cascade thing; all finished!
    }
    else
        if ((pitem = _GetItemFromLV(iItem)) &&
            SUCCEEDED(_GetFolderAndPidl(pitem, &psf, &pidl)))
        {
            // See if the item is still valid.
            // Do this only for SFGAO_FILESYSTEM objects because
            // we can't be sure that other folders support SFGAO_VALIDATE,
            // and besides, you can't resolve any other types of objects
            // anyway...

            DWORD dwAttr = SFGAO_FILESYSTEM | SFGAO_VALIDATE;
            if (FAILED(psf->GetAttributesOf(1, &pidl, &dwAttr)) ||
                (dwAttr & SFGAO_FILESYSTEM | SFGAO_VALIDATE) == SFGAO_FILESYSTEM ||
                FAILED(_InvokeDefaultCommand(iItem, psf, pidl)))
            {
                // Object is bogus - offer to delete it
                if ((_dwFlags & HOSTF_CANDELETE) && pitem->IsPinned())
                {
                    _OfferDeleteBrokenItem(pitem, psf, pidl);
                }
            }

            psf->Release();
        }
    return 0;
#else
    // eax
    PaneItem *pitem; // edi
    HRESULT hr; // eax
    SMNGETISTARTBUTTON nm; // [esp+4h] [ebp-18h] BYREF
    const ITEMID_CHILD *pidl; // [esp+14h] [ebp-8h] BYREF
    IShellFolder *psf; // [esp+18h] [ebp-4h] BYREF
    DWORD dwAttr; // [esp+28h] [ebp+Ch] SPLIT BYREF

    nm.pstb = 0;
    _SendNotify(this->_hwnd, 218u, &nm.hdr);
    if (nm.pstb)
    {
		nm.pstb->LockStartPane();
        nm.pstb->Release();
    }

    DWORD dwCascadeFlags = 0;
    if ((dwFlags & 1) != 0)
        dwCascadeFlags = 0x12;
    
    if (!_OnCascade(iItem, dwCascadeFlags))
    {
        pitem = _GetItemFromLV(iItem);
        if (pitem)
        {
            this->_NotifyInvoke(pitem);
            if (_GetFolderAndPidlForActivate(pitem, &psf, &pidl) >= 0)
            {
                dwAttr = 0x1000000;
                if (psf->GetAttributesOf(1u, &pidl, &dwAttr) < 0)
                {
                    goto LABEL_14;
                }
                hr = _InvokeDefaultCommand(iItem, psf, pidl);
                if (hr >= 0 || hr == 0x800704C7)
                {
                    hr = 0;
                }
                
                if (hr < 0) 
                {
                LABEL_14:
                    if ((this->_dwFlags & 2) != 0 && pitem->_iPinPos >= 0)
                    {
                        _OfferDeleteBrokenItem(pitem, psf, pidl);
                    }
                }
                psf->Release();
            }
            pitem->Release();
        }
    }
    return 0;
#endif
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::_InvokeDefaultCommand(int iItem, IShellFolder *psf, LPCITEMIDLIST pidl)
{
    HRESULT hr = SHInvokeDefaultCommand(GetShellWindow(), psf, pidl);
    if (SUCCEEDED(hr))
    {
        if (_dwFlags & HOSTF_FIREUEMEVENTS)
        {
			// UAFireEventByFolderAndIDList(&UAIID_AUTOMATIC, 0, psf, pidl, 0); // EXEX-VISTA(allison): TODO: Uncomment when implemented.
            _FireUEMPidlEvent(psf, pidl);
        }
        SMNMCOMMANDINVOKED ci;
        ListView_GetItemRect(_hwndList, iItem, &ci.rcItem, LVIR_BOUNDS);
        MapWindowRect(_hwndList, NULL, &ci.rcItem);
        _SendNotify(_hwnd, SMN_COMMANDINVOKED, &ci.hdr);
    }
    return hr;
}

class OfferDelete
{
public:

    LPTSTR          _pszName;
    LPITEMIDLIST    _pidlFolder;
    LPITEMIDLIST    _pidlFull;
    IStartMenuPin * _psmpin;
    HWND            _hwnd;

    ~OfferDelete()
    {
        SHFree(_pszName);
        ILFree(_pidlFolder);
        ILFree(_pidlFull);
    }

    BOOL _RepairBrokenItem();
    void _ThreadProc();

    static DWORD s_ThreadProc(LPVOID lpParameter)
    {
        OfferDelete *poffer = (OfferDelete *)lpParameter;
        poffer->_ThreadProc();
        delete poffer;
        return 0;
    }
};


BOOL OfferDelete::_RepairBrokenItem()
{
    BOOL fSuccess = FALSE;
    LPITEMIDLIST pidlNew;
    HRESULT hr = _psmpin->Resolve(_hwnd, 0, _pidlFull, &pidlNew);
    if (pidlNew)
    {
        ASSERT(hr == S_OK); // only the S_OK case should alloc a new pidl

        // Update to reflect the new pidl
        ILFree(_pidlFull);
        _pidlFull = pidlNew;

        // Re-invoke the default command; if it fails the second time,
        // then I guess the Resolve didn't work after all.
        IShellFolder *psf;
        LPCITEMIDLIST pidlChild;
        if (SUCCEEDED(SHBindToIDListParent(_pidlFull, IID_PPV_ARGS(&psf), &pidlChild)))
        {
            if (SUCCEEDED(SHInvokeDefaultCommand(_hwnd, psf, pidlChild)))
            {
                fSuccess = TRUE;
            }
            psf->Release();
        }

    }
    return fSuccess;
}

void OfferDelete::_ThreadProc()
{
    _hwnd = SHCreateWorkerWindowW(NULL, NULL, 0, 0, NULL, NULL);
    if (_hwnd)
    {
        if (SUCCEEDED(CoCreateInstanceHook(CLSID_StartMenuPin, NULL, CLSCTX_INPROC_SERVER,
                                       IID_PPV_ARGS(&_psmpin))))
        {
            //
            //  First try to repair it by invoking the shortcut tracking code.
            //  If that fails, then offer to delete.
            if (!_RepairBrokenItem() &&
                ShellMessageBox(_AtlBaseModule.GetResourceInstance(), NULL,
                                MAKEINTRESOURCE(IDS_SFTHOST_OFFERREMOVEITEM),
                                _pszName, MB_YESNO) == IDYES)
            {
                _psmpin->Modify(_pidlFull, NULL);
            }
            ATOMICRELEASE(_psmpin);
        }
        DestroyWindow(_hwnd);
    }
}

void SFTBarHost::_OfferDeleteBrokenItem(PaneItem *pitem, IShellFolder *psf, LPCITEMIDLIST pidlChild)
{
    //
    //  The offer is done on a separate thread because putting up modal
    //  UI while the Start Menu is open creates all sorts of weirdness.
    //  (The user might decide to switch to Classic Start Menu
    //  while the dialog is still up, and we get our infrastructure
    //  ripped out from underneath us and then USER faults inside
    //  MessageBox...  Not good.)
    //
    OfferDelete *poffer = new OfferDelete;
    if (poffer)
    {
        if ((poffer->_pszName = DisplayNameOfItem(pitem, psf, pidlChild, SHGDN_NORMAL)) != NULL &&
            SUCCEEDED(SHGetIDListFromUnk(psf, &poffer->_pidlFolder)) &&
            (poffer->_pidlFull = ILCombine(poffer->_pidlFolder, pidlChild)) != NULL &&
            SHCreateThread(OfferDelete::s_ThreadProc, poffer, CTF_COINIT, NULL))
        {
            poffer = NULL;       // thread took ownership
        }
        delete poffer;
    }
}

BOOL ShowInfoTip()
{
    // find out if infotips are on or off, from the registry settings
    SHELLSTATE ss;
    // force a refresh
    SHGetSetSettings(&ss, 0, TRUE);
    SHGetSetSettings(&ss, SSF_SHOWINFOTIP, FALSE);
    return ss.fShowInfoTip;
}

// over-ridable method for getting the infotip on an item
// EXEX-VISTA(allison): Validated.
void SFTBarHost::GetItemInfoTip(PaneItem *pitem, LPTSTR pszText, DWORD cch)
{
    IShellFolder *psf;
    LPCITEMIDLIST pidl;

    if (pszText && cch)
    {
        *pszText = 0;

        if (SUCCEEDED(_GetFolderAndPidl(pitem, &psf, &pidl)))
        {
            GetInfoTip(psf, pidl, pszText, cch);
            psf->Release();
        }
    }
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnLVNGetInfoTip(LPNMLVGETINFOTIP plvn)
{
    _DebugConsistencyCheck();

	*plvn->pszText = 0;

    PaneItem *pitem;

    if (ShowInfoTip() && 
        (pitem = _GetItemFromLV(plvn->iItem)) &&
        !pitem->IsCascade())
    {
        int cchName = (plvn->dwFlags & LVGIT_UNFOLDED) ? 0 : lstrlen(plvn->pszText);

        if (cchName)
        {
            StringCchCat(plvn->pszText, plvn->cchTextMax, TEXT("\r\n"));
            cchName = lstrlen(plvn->pszText);
        }

        // If there is room in the buffer after we added CRLF, append the
        // infotip text.  We succeeded if there was nontrivial infotip text.

        if (cchName < plvn->cchTextMax)
        {
            GetItemInfoTip(pitem, &plvn->pszText[cchName], plvn->cchTextMax - cchName);
        }

		pitem->Release();
    }

    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT _SendNotify(HWND hwndFrom, UINT code, OPTIONAL NMHDR *pnm)
{
    NMHDR nm;
    if (pnm == NULL)
    {
        pnm = &nm;
    }
    pnm->hwndFrom = hwndFrom;
    pnm->idFrom = GetDlgCtrlID(hwndFrom);
    pnm->code = code;
    return SendMessage(GetParent(hwndFrom), WM_NOTIFY, pnm->idFrom, (LPARAM)pnm);
}

//****************************************************************************
//
//  Drag sourcing
//

// *** IDropSource::GiveFeedback ***

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::GiveFeedback(DWORD dwEffect)
{
    if (_fForceArrowCursor)
    {
        SetCursor(LoadCursor(NULL, IDC_ARROW));
        return S_OK;
    }

    return DRAGDROP_S_USEDEFAULTCURSORS;
}

// *** IDropSource::QueryContinueDrag ***

HRESULT SFTBarHost::QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
{
    if (fEscapePressed ||
        (grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == (MK_LBUTTON | MK_RBUTTON))
    {
        return DRAGDROP_S_CANCEL;
    }
    if ((grfKeyState & (MK_LBUTTON | MK_RBUTTON)) == 0)
    {
        return DRAGDROP_S_DROP;
    }
    return S_OK;
}

// *** IServiceProvider::QueryService ***
HRESULT SFTBarHost::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    HRESULT hr = E_FAIL;
    if (IsEqualGUID(guidService, IID_IFolderView) || IsEqualGUID(guidService, SID_SM_MFU))
    {
		hr = QueryInterface(riid, ppvObject);
    }
    return hr;
}

// *** IOleCommandTarget::QueryStatus ***
HRESULT SFTBarHost::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds,
                                OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
    return E_NOTIMPL;
}

// *** IOleCommandTarget::Exec ***
HRESULT SFTBarHost::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut)
{
    return E_NOTIMPL;
}

// EXEX-VISTA(allison): Validated. Might still need cleanup.
LRESULT SFTBarHost::_OnLVNBeginDrag(LPNMLISTVIEW plv)
{
#ifdef DEAD_CODE
    //If changes are restricted, don't allow drag and drop!
    if(_AreChangesRestricted())
        return 0;

    _DebugConsistencyCheck();

    ASSERT(_pdtoDragOut == NULL);
    _pdtoDragOut = NULL;

    PaneItem *pitem = _GetItemFromLV(plv->iItem);
    ASSERT(pitem);

    IDataObject *pdto;
    if (pitem && SUCCEEDED(_GetUIObjectOfItem(pitem, IID_PPV_ARGS(&pdto))))
    {
        POINT pt;

        pt = plv->ptAction;
        ClientToScreen(_hwndList, &pt);

        if (_pdsh)
        {
            _pdsh->InitializeFromWindow(_hwndList, &pt, pdto);
        }

        CLIPFORMAT cfOFFSETS = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLISTOFFSET);

        POINT *apts = (POINT*)GlobalAlloc(GPTR, sizeof(POINT)*2);
        if (NULL != apts)
        {
            POINT ptOrigin = {0};
            POINT ptItem = {0};

            ListView_GetOrigin(_hwndList, &ptOrigin);
            apts[0].x = plv->ptAction.x + ptOrigin.x;
            apts[0].y = plv->ptAction.y + ptOrigin.y;

            ListView_GetItemPosition(_hwndList,plv->iItem,&ptItem);
            apts[1].x = ptItem.x - apts[0].x;
            apts[1].y = ptItem.y - apts[0].y;

            HRESULT hr = DataObj_SetGlobal(pdto, cfOFFSETS, apts);
            if (FAILED(hr))
            {
                GlobalFree((HGLOBAL)apts);
            }
        }

        // We don't need to refcount _pdtoDragOut since its lifetime
        // is the same as pdto.
        _pdtoDragOut = pdto;
        _iDragOut = plv->iItem;
        _iPosDragOut = pitem->_iPos;

        // Notice that DROPEFFECT_MOVE is explicitly forbidden.
        // You cannot move things out of the control.
        DWORD dwEffect = DROPEFFECT_LINK | DROPEFFECT_COPY;
        DoDragDrop(pdto, this, dwEffect, &dwEffect);

        _pdtoDragOut = NULL;
        pdto->Release();
    }
    return 0;
#else
    if (_AreChangesRestricted())
    {
        return 0;   
    }
    
    ASSERT(_pdtoDragOut == NULL); // 3115
    _pdtoDragOut = 0;

    PaneItem* pitem = _GetItemFromLV(plv->iItem);
	ASSERT(pitem); // 3119

    if (pitem)
    {
        IDataObject *pdto;
        if (SUCCEEDED(_GetUIObjectOfItem(pitem, IID_PPV_ARGS(&pdto))))
        {
            POINT pt;

            pt = plv->ptAction;
            ClientToScreen(_hwndList, &pt);

            if (_pdsh)
            {
                _pdsh->InitializeFromWindow(NULL, &pt, pdto);
            }

            CLIPFORMAT cfOFFSETS = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLISTOFFSET);

            POINT *apts = (POINT *)GlobalAlloc(GPTR, sizeof(POINT) * 2);
            if (NULL != apts)
            {
                POINT ptOrigin = {0};
                POINT ptItem = {0};

                ListView_GetOrigin(_hwndList, &ptOrigin);
                apts[0].x = ptOrigin.x + plv->ptAction.x;
                apts[0].y = ptOrigin.y + plv->ptAction.y;

                ListView_GetItemPosition(_hwndList, plv->iItem, &ptItem);
                apts[1].x = ptItem.x - apts[0].x;
                apts[1].y = ptItem.y - apts[0].y;

                HRESULT hr = DataObj_SetGlobal(pdto, cfOFFSETS, apts);
                if (FAILED(hr))
                {
                    GlobalFree(apts);
                }
            }

            _pdtoDragOut = pdto;
            _iDragOut = plv->iItem;
            _iPosDragOut = pitem->_iPos;

            DWORD dwEffect = DROPEFFECT_LINK;
            DoDragDrop(pdto, this, dwEffect, &dwEffect);
            
            _pdtoDragOut = NULL;
            pdto->Release();
        }
        pitem->Release();
    }
    return 0;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
HRESULT SHBeginLabelEdit(IShellFolder *psf, LPCITEMIDLIST pidl, HWND hwndEdit, int cchLimit)
{
    HRESULT hr = 1;

    SFGAOF dwAttribs = SHGetAttributes(psf, pidl, 0x20400010u);
    if ((dwAttribs & 0x10) != 0)
    {
        hr = 0;

        WCHAR szName[260];
        WCHAR szNameNew[260];
        if (DisplayNameOf(psf, pidl, 0x1001, szName, 260u) >= 0 && DisplayNameOf(psf, pidl, 0x8001, szNameNew, 260u) >= 0)
        {
            SetWindowText(hwndEdit, szName);
            if ((dwAttribs & 0x20400000) != 0x20000000 && !StrCmpW(szNameNew, szName))
            {
                SendMessage(hwndEdit, EM_SETSEL, 0, PathFindExtension(szNameNew) - szNameNew);
            }

            IItemNameLimits *pinl;
            if (psf->QueryInterface(IID_PPV_ARGS(&pinl)) >= 0)
            {
                pinl->GetMaxLength(szNameNew, &cchLimit);
                pinl->Release();
            }
            SHLimitInputEdit(hwndEdit, psf);
        }

        if (!cchLimit)
            cchLimit = 128;

        SendMessage(hwndEdit, EM_LIMITTEXT, cchLimit, 0);
    }
    return hr;
}

//
//  Must perform validation of SFGAO_CANRENAME when the label edit begins
//  because John Gray somehow can trick the listview into going into edit
//  mode by clicking in the right magic place, so this is the only chance
//  we get to reject things that aren't renamable...
//

// EXEX-VISTA(allison): Validated. Might still need cleanup.
LRESULT SFTBarHost::_OnLVNBeginLabelEdit(NMLVDISPINFO *plvdi)
{
#ifdef DEAD_CODE
    LRESULT lres = 1;

    PaneItem *pitem = _GetItemFromLVLParam(plvdi->item.lParam);

    IShellFolder *psf;
    LPCITEMIDLIST pidl;

    if (_fAllowEditLabel &&
        pitem && SUCCEEDED(_GetFolderAndPidl(pitem, &psf, &pidl)))
    {
        DWORD dwAttr = SFGAO_CANRENAME;
        if (SUCCEEDED(psf->GetAttributesOf(1, &pidl, &dwAttr)) &&
            (dwAttr & SFGAO_CANRENAME))
        {
            LPTSTR ptszName = _DisplayNameOf(psf, pidl,
                                    SHGDN_INFOLDER | SHGDN_FOREDITING);
            if (ptszName)
            {
                HWND hwndEdit = ListView_GetEditControl(_hwndList);
                if (hwndEdit)
                {
                    SetWindowText(hwndEdit, ptszName);

                    int cchLimit = MAX_PATH;
                    IItemNameLimits *pinl;
                    if (SUCCEEDED(psf->QueryInterface(IID_PPV_ARGS(&pinl))))
                    {
                        pinl->GetMaxLength(ptszName, &cchLimit);
                        pinl->Release();
                    }
                    Edit_LimitText(hwndEdit, cchLimit);

                    // use way-cool helper which pops up baloon tips if they enter an invalid folder....
                    SHLimitInputEdit(hwndEdit, psf);

                    // Block menu mode during editing so the user won't
                    // accidentally cancel out of rename mode just because
                    // they moved the mouse.
                    SMNMBOOL nmb;
                    nmb.f = TRUE;
                    _SendNotify(_hwnd, SMN_BLOCKMENUMODE, &nmb.hdr);

                    lres = 0;
                }
                SHFree(ptszName);
            }
        }
        psf->Release();
    }

    return lres;
#else
    LRESULT lres = 1;
    
    PaneItem* pitem = _GetItemFromLVLParam(plvdi->item.lParam);
    if (_fAllowEditLabel && pitem)
    {
        IShellFolder *psf;
        LPCITEMIDLIST pidl;
        if (_GetFolderAndPidl(pitem, &psf, &pidl) >= 0)
        {
            HWND hwndEdit = ListView_GetEditControl(this->_hwndList);
            if (hwndEdit)
            {
                lres = SHBeginLabelEdit(psf, pidl, hwndEdit, 0);
                if (!lres)
                {
                    SMNMBOOL nmb;
                    nmb.f = TRUE;
                    _SendNotify(_hwnd, SMN_BLOCKMENUMODE, &nmb.hdr);
                }
            }
            psf->Release();
        }
        pitem->Release();
    }
    return lres;
#endif
}

LRESULT SFTBarHost::_OnLVNEndLabelEdit(NMLVDISPINFO *plvdi)
{
    // Unblock menu mode now that editing is over.
    SMNMBOOL nmb;
    nmb.f = FALSE;
    _SendNotify(_hwnd, SMN_BLOCKMENUMODE, &nmb.hdr);

    // If changing to NULL pointer, then user is cancelling
    if (!plvdi->item.pszText)
        return FALSE;

    // Note: We allow the user to type blanks. Regfolder treats a blank
    // name as "restore default name".
    PathRemoveBlanks(plvdi->item.pszText);
    PaneItem *pitem = _GetItemFromLVLParam(plvdi->item.lParam);

    HRESULT hr = ContextMenuRenameItem(pitem, plvdi->item.pszText);

    if (SUCCEEDED(hr))
    {
        LPTSTR ptszName = _DisplayNameOfItem(pitem, SHGDN_NORMAL);
        if (ptszName)
        {
            ListView_SetItemText(_hwndList, plvdi->item.iItem, 0, ptszName);
            _SendNotify(_hwnd, SMN_NEEDREPAINT, NULL);
        }
    }
    else if (hr != HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        _EditLabel(plvdi->item.iItem);
    }

    // Always return FALSE to prevent listview from changing the
    // item text to what the user typed.  If the rename succeeded,
    // we manually set the name to the new name (which might not be
    // the same as what the user typed).
    return FALSE;
}

LRESULT SFTBarHost::_OnLVNKeyDown(LPNMLVKEYDOWN pkd)
{
    // Plain F2 (no shift, ctrl or alt) = rename
    if (pkd->wVKey == VK_F2 && GetKeyState(VK_SHIFT) >= 0 &&
        GetKeyState(VK_CONTROL) >= 0 && GetKeyState(VK_MENU) >= 0 &&
        (_dwFlags & HOSTF_CANRENAME))
    {
        int iItem = _GetLVCurSel();
        if (iItem >= 0)
        {
            _EditLabel(iItem);
            // cannot return TRUE because listview mistakenly thinks
            // that all WM_KEYDOWNs lead to WM_CHARs (but this one doesn't)
        }
    }

    return 0;
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnLVNAsyncDrawn(NMLVASYNCDRAWN *plvad)
{
    plvad->dwRetFlags = 0x1;

    PaneItem *pitem = _GetItemFromLV(plvad->iItem);
    if (pitem)
    {
        IShellFolder *psf = NULL;
        LPCITEMIDLIST pidl = NULL;
        if (SUCCEEDED(GetFolderAndPidl(pitem, &psf, &pidl)))
        {
            AddImageForItem(pitem, psf, pidl, 0);
            psf->Release();
        }
        pitem->Release();
    }
    return 1;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
LRESULT SFTBarHost::_OnSMNGetMinSize(PSMNGETMINSIZE pgms)
{
#ifdef DEAD_CODE
    // We need to synchronize here to get the proper size
    if (_fBGTask && !HasDynamicContent())
    {
        // Wait for the enumeration to be done
        while (TRUE)
        {
            MSG msg;
            // Need to peek messages for all queues here or else WaitMessage will say
            // that some messages are ready to be processed and we'll end up with an
            // active loop
            if (PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
            {
                if (PeekMessage(&msg, _hwnd, SFTBM_REPOPULATE, SFTBM_REPOPULATE, PM_REMOVE))
                {
                    DispatchMessage(&msg);
                    break;
                }
            }
            WaitMessage();
        }
    }

    int cItems = _cPinnedDesired + _cNormalDesired;
    int cSep = _cSep;

    // if the repopulate hasn't happened yet, but we've got pinned items, we're going to have a separator
    if (_cSep == 0 && _cPinnedDesired > 0)
        cSep = 1;
    int cy = (_cyTile * cItems) + (_cySepTile * cSep);

    // add in theme margins
    cy += _margins.cyTopHeight + _margins.cyBottomHeight;

    // SPP_PROGLIST gets a bonus separator at the bottom
    if (_iThemePart == SPP_PROGLIST)
    {
        cy += _cySep;
    }

    pgms->siz.cy = cy;

    return 0;
#else
    if (this->_fBGTask && !this->HasDynamicContent())
    {
        MSG msg;
        while (!PeekMessage(&msg, 0, 0, 0, 0) || !PeekMessage(&msg, this->_hwnd, 0x400u, 0x400u, 1u))
        {
            WaitMessage();
        }
        DispatchMessage(&msg);
    }

    int cPinnedDesired = this->_cPinnedDesired;
    int cSep = this->_cSep;
    if (!cSep)
        cSep = cPinnedDesired > 0;
    LONG v5 = this->_margins.cyTopHeight
        + this->_margins.cyBottomHeight
        + (cPinnedDesired + this->_cNormalDesired) * this->_cyTile
        + cSep * this->_cySepTile;
    if (this->_iThemePart == 4)
        v5 += this->_cySep;
    pgms->field_14.cy = -1;
    pgms->siz.cy = v5;
    int v6 = this->field_6C;
    int cNormalDesired = this->_cNormalDesired;
    if (v6 < cNormalDesired)
        pgms->field_14.cy = v5 - this->_cyTile * (cNormalDesired - v6);
    pgms->field_14.cx = this->_cxIndent + this->GetMinTextWidth();
    return 0;
#endif
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
LRESULT SFTBarHost::_OnSMNFindItem(PSMNDIALOGMESSAGE pdm)
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
            ListView_SetItemState(_hwndList, pdm->itemID,
                                  LVIS_SELECTED | LVIS_FOCUSED,
                                  LVIS_SELECTED | LVIS_FOCUSED);
            if ((pdm->flags & SMNDM_FINDMASK) != SMNDM_HITTEST)
            {
                ListView_KeyboardSelected(_hwndList, pdm->itemID);
            }
        }
    }
    else
    {
        //
        //  If not found, then tell caller what our orientation is (vertical)
        //  and where the currently-selected item is.
        //

        pdm->flags |= SMNDM_VERTICAL;
        int iItem = _GetLVCurSel();
        RECT rc;
        if (iItem >= 0 &&
            ListView_GetItemRect(_hwndList, iItem, &rc, LVIR_BOUNDS))
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
    LRESULT lres = _OnSMNFindItemWorker(pdm);

    if (lres)
    {
        if ((pdm->flags & 0x100 | 0x800))
        {
            UINT state = LVIS_SELECTED;
            if ((pdm->flags & SMNDM_SELECT))
                state |= LVIS_FOCUSED;
            ListView_SetItemState(_hwndList, pdm->itemID, state, state);

            if ((pdm->flags & SMNDM_FINDMASK) != SMNDM_HITTEST)
            {
                ListView_KeyboardSelected(_hwndList, pdm->itemID);
            }

			LVITEM lvi = {0};
            lvi.iItem = (int)pdm->itemID;
            lvi.mask = LVIF_IMAGE;
            ListView_GetItem(_hwndList, &lvi);
            _NotifyHoverImage(lvi.iImage);
        }
    }
    else
    {
        pdm->flags |= 0x4000u;
        int iItem = _GetLVCurSel();
        RECT rc;
        if (iItem >= 0 && ListView_GetItemRect(_hwndList, iItem, &rc, LVIR_BOUNDS))
        {
            pdm->pt.x = (rc.left + rc.right) / 2;
            pdm->pt.y = (rc.top + rc.bottom) / 2;
        }
        else
        {
            pdm->pt.x = 0;
            pdm->pt.y = 0;
        }
    }
    return lres;
#endif
}

// EXEX-VISTA(allison): Validated.
TCHAR SFTBarHost::GetItemAccelerator(PaneItem *pitem, int iItemStart)
{
    TCHAR sz[2];
    ListView_GetItemText(_hwndList, iItemStart, 0, sz, ARRAYSIZE(sz));
    return CharUpperCharA(sz[0]);
}

// EXEX-VISTA(allison): Validated. Still needs major cleanup.
LRESULT SFTBarHost::_OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm)
{
#ifdef DEAD_CODE
    LVFINDINFO lvfi;
    LVHITTESTINFO lvhti;

    switch (pdm->flags & SMNDM_FINDMASK)
    {
        case SMNDM_FINDFIRST:
        L_SMNDM_FINDFIRST:
            // Note: We can't just return item 0 because drag/drop pinning
            // may have gotten the physical locations out of sync with the
            // item numbers.
            lvfi.vkDirection = VK_HOME;
            lvfi.flags = LVFI_NEARESTXY;
            pdm->itemID = ListView_FindItem(_hwndList, -1, &lvfi);
            return pdm->itemID >= 0;

        case SMNDM_FINDLAST:
            // Note: We can't just return cItems-1 because drag/drop pinning
            // may have gotten the physical locations out of sync with the
            // item numbers.
            lvfi.vkDirection = VK_END;
            lvfi.flags = LVFI_NEARESTXY;
            pdm->itemID = ListView_FindItem(_hwndList, -1, &lvfi);
            return pdm->itemID >= 0;

        case SMNDM_FINDNEAREST:
            lvfi.pt = pdm->pt;
            lvfi.vkDirection = VK_UP;
            lvfi.flags = LVFI_NEARESTXY;
            pdm->itemID = ListView_FindItem(_hwndList, -1, &lvfi);
            return pdm->itemID >= 0;

        case SMNDM_HITTEST:
            lvhti.pt = pdm->pt;
            pdm->itemID = ListView_HitTest(_hwndList, &lvhti);
            return pdm->itemID >= 0;

        case SMNDM_FINDFIRSTMATCH:
        case SMNDM_FINDNEXTMATCH:
        {
            int iItemStart;
            if ((pdm->flags & SMNDM_FINDMASK) == SMNDM_FINDFIRSTMATCH)
            {
                iItemStart = 0;
            }
            else
            {
                iItemStart = _GetLVCurSel() + 1;
            }
            TCHAR tch = CharUpperCharA((TCHAR)pdm->pmsg->wParam);
            int iItems = ListView_GetItemCount(_hwndList);
            for (iItemStart; iItemStart < iItems; iItemStart++)
            {
                PaneItem *pitem = _GetItemFromLV(iItemStart);
                if (GetItemAccelerator(pitem, iItemStart) == tch)
                {
                    pdm->itemID = iItemStart;
                    return TRUE;
                }
            }
            return FALSE;
        }
        break;

        case SMNDM_FINDNEXTARROW:
            if (pdm->pmsg->wParam == VK_UP)
            {
                pdm->itemID = ListView_GetNextItem(_hwndList, _GetLVCurSel(), LVNI_ABOVE);
                return pdm->itemID >= 0;
            }

            if (pdm->pmsg->wParam == VK_DOWN)
            {
                // HACK! ListView_GetNextItem explicitly fails to find a "next item"
                // if you tell it to start at -1 (no current item), so if there is no
                // focus item, we have to change it to a SMNDM_FINDFIRST.
                int iItem = _GetLVCurSel();
                if (iItem == -1)
                {
                    goto L_SMNDM_FINDFIRST;
                }
                pdm->itemID = ListView_GetNextItem(_hwndList, iItem, LVNI_BELOW);
                return pdm->itemID >= 0;
            }

            if (pdm->flags & SMNDM_TRYCASCADE)
            {
                pdm->itemID = _GetLVCurSel();
                return _OnCascade((int)pdm->itemID, MPPF_KEYBOARD | MPPF_INITIALSELECT);
            }

            return FALSE;

        case SMNDM_INVOKECURRENTITEM:
        {
            int iItem = _GetLVCurSel();
            if (iItem >= 0)
            {
                DWORD aif = 0;
                if (pdm->flags & SMNDM_KEYBOARD)
                {
                    aif |= AIF_KEYBOARD;
                }
                _ActivateItem(iItem, aif);
                return TRUE;
            }
        }
        return FALSE;

        case SMNDM_OPENCASCADE:
        {
            DWORD mppf = 0;
            if (pdm->flags & SMNDM_KEYBOARD)
            {
                mppf |= MPPF_KEYBOARD | MPPF_INITIALSELECT;
            }
            pdm->itemID = _GetLVCurSel();
            return _OnCascade((int)pdm->itemID, mppf);
        }

        case SMNDM_FINDITEMID:
            return TRUE;

        default:
            ASSERT(!"Unknown SMNDM command");
            break;
    }

    return FALSE;
#else
    UINT flags; // ecx
    LRESULT v5; // eax
    int iItemStart; // ebx
    int wParam; // eax
    LRESULT LVCurSel; // ebx
    LRESULT v11; // eax
    LRESULT v12; // eax
    LRESULT iItem; // eax
    DWORD mppf; // ebx
    DWORD v15; // [esp-4h] [ebp-64h]
    LVHITTESTINFO lvhti; // [esp+10h] [ebp-50h] BYREF
    LVFINDINFOW lvfi; // [esp+28h] [ebp-38h] BYREF
    int iItems; // [esp+40h] [ebp-20h]
    WCHAR tch; // [esp+44h] [ebp-1Ch]
    //CPPEH_RECORD ms_exc; // [esp+48h] [ebp-18h]
    PaneItem *pitem; // [esp+68h] [ebp+8h] MAPDST

    flags = pdm->flags;
    switch (flags & 0xF)
    {
        case 0u:                                    // SMNDM_FINDFIRSTMATCH
        case 1u:                                    // SMNDM_FINDNEXTMATCH
            if ((flags & 0xF) != 0)
                iItemStart = SFTBarHost::_GetLVCurSel() + 1;
            else
                iItemStart = 0;
            tch = (unsigned __int16)CharUpperW((LPWSTR)LOWORD(pdm->pmsg->wParam));
            iItems = SendMessageW(this->_hwndList, LVM_GETITEMCOUNT, 0, 0);
            if (iItemStart >= iItems)
                return 0;
            while (1)
            {
                pitem = SFTBarHost::_GetItemFromLV(iItemStart);
                if (pitem)
                {
                    if (this->GetItemAccelerator(pitem, iItemStart) == tch)
                    {
                        pdm->itemID = iItemStart;
                        pitem->Release();
                        return 1;
                    }
                    pitem->Release();
                }
                if (++iItemStart >= iItems)
                    return 0;
            }
        case 2u:                                    // SMNDM_FINDNEAREST
            lvfi.pt = pdm->pt;
            lvfi.vkDirection = 38;
            goto LABEL_24;
        case 3u:                                    // SMNDM_FINDFIRST
            goto LABEL_23;
        case 4u:                                    // SMNDM_FINDLAST
            lvfi.vkDirection = 35;
            goto LABEL_24;
        case 5u:                                    // SMNDM_FINDNEXTARROW
            wParam = pdm->pmsg->wParam;
            if (wParam == 38)
            {
                LVCurSel = SFTBarHost::_GetLVCurSel();
                v11 = SendMessageW(this->_hwndList, LVM_GETNEXTITEM, LVCurSel, 256);
            LABEL_18:
                pdm->itemID = v11;
                return v11 != LVCurSel && v11 >= 0;
            }
            if (wParam == VK_DOWN)
            {
                LVCurSel = SFTBarHost::_GetLVCurSel();
                if (LVCurSel != -1)
                {
                    v11 = SendMessageW(this->_hwndList, LVM_GETNEXTITEM, LVCurSel, 512);
                    goto LABEL_18;
                }
            LABEL_23:
                lvfi.vkDirection = VK_HOME;
            LABEL_24:
                lvfi.flags = LVFI_NEARESTXY;
                v5 = SendMessageW(this->_hwndList, LVM_FINDITEMW, -1, (LPARAM)&lvfi);
            LABEL_25:
                pdm->itemID = v5;
                return v5 >= 0;
            }
            else
            {
                if ((flags & 0x200) == 0)
                    return 0;
                v12 = SFTBarHost::_GetLVCurSel();
                v15 = 0x12;
            LABEL_29:
                pdm->itemID = v12;
                return SFTBarHost::_OnCascade(v12, v15);
            }
        case 6u:                                    // SMNDM_INVOKECURRENTITEM
            iItem = SFTBarHost::_GetLVCurSel();
            if (iItem < 0)
                return 0;
            SFTBarHost::_ActivateItem(iItem, (pdm->flags & 0x400) != 0);
            return 1;
        case 7u:                                    // SMNDM_HITTEST
            lvhti.pt = pdm->pt;
            v5 = SendMessageW(this->_hwndList, LVM_HITTEST, 0, (LPARAM)&lvhti);
            goto LABEL_25;
        case 8u:                                    // SMNDM_OPENCASCADE
            mppf = 0;
            if ((flags & 0x400) != 0)
                mppf = 0x12;
            v12 = SFTBarHost::_GetLVCurSel();
            v15 = mppf;
            goto LABEL_29;
        case 9u:                                    // SMNDM_FINDITEMID
        case 0xAu:
            return 1;
        case 0xBu:
            return 0;
        default:
            //if (CcshellAssertFailedW(
            //    L"d:\\longhorn\\shell\\explorer\\desktop2\\sfthost.cpp",
            //    3570,
            //    L"!\"Unknown SMNDM command\"",
            //    0))
            //{
            //    AttachUserModeDebugger();
            //    do
            //    {
            //        __debugbreak();
            //        ms_exc.registration.TryLevel = -2;
            //    } while (dword_108B97C);
            //}
            return 0;
    }
#endif
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnSMNDismiss()
{
    if (_fNeedsRepopulate)
    {
        _RepopulateList();
    }
    return 0;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::_OnCascade(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    return _OnCascade((int)wParam, (DWORD)lParam);
}

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
BOOL SFTBarHost::_OnCascade(int iItem, DWORD dwFlags)
{
#ifdef DEAD_CODE
    BOOL fSuccess = FALSE;
    SMNTRACKSHELLMENU tsm;
    tsm.dwFlags = dwFlags;
    tsm.itemID = iItem;

    if (iItem >= 0 &&
        ListView_GetItemRect(_hwndList, iItem, &tsm.rcExclude, LVIR_BOUNDS))
    {
        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem && pitem->IsCascade())
        {
            if (SUCCEEDED(GetCascadeMenu(pitem, &tsm.psm)))
            {
                MapWindowRect(_hwndList, NULL, &tsm.rcExclude);
                HWND hwnd = _hwnd;
                _iCascading = iItem;
                _SendNotify(_hwnd, SMN_TRACKSHELLMENU, &tsm.hdr);
                tsm.psm->Release();
                fSuccess = TRUE;
            }
        }
    }
    return fSuccess;
#else
    BOOL fSuccess = FALSE;
    SMNTRACKSHELLMENU tsm;
    tsm.dwFlags = dwFlags;
    tsm.itemID = iItem;
    if (iItem >= 0 &&
        ListView_GetItemRect(_hwndList, iItem, &tsm.rcExclude, LVIR_BOUNDS))
    {
        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem && pitem->IsCascade())
        {
            if (SUCCEEDED(GetCascadeMenu(pitem, &tsm.psm)))
            {
                MapWindowRect(this->_hwndList, 0, &tsm.rcExclude);
                HWND hwnd = _hwnd;
                _iCascading = iItem;
                _SendNotify(_hwnd, 216, &tsm.hdr);
                tsm.psm->Release();
                fSuccess = TRUE;
                //SHTracePerf(&ShellTraceId_Explorer_StartPane_Cascade_Show_Start);
                _NotifyCascade(pitem);
            }
            pitem->Release();
        }

    }
    return fSuccess;
#endif
}

HRESULT SFTBarHost::QueryInterface(REFIID riid, void * *ppvOut)
{
    static const QITAB qit[] = {
        QITABENT(SFTBarHost, IDropTarget),
        QITABENT(SFTBarHost, IDropSource),
        QITABENT(SFTBarHost, IServiceProvider),
        QITABENT(SFTBarHost, IOleCommandTarget),
        QITABENT(SFTBarHost, IObjectWithSite),
        QITABENT(SFTBarHost, IAccessible),
        QITABENT(SFTBarHost, IDispatch), // IAccessible derives from IDispatch
        { 0 },
    };
    return QISearch(this, qit, riid, ppvOut);
}

ULONG SFTBarHost::AddRef()
{
    return InterlockedIncrement(&_lRef);
}

ULONG SFTBarHost::Release()
{
    ASSERT( 0 != _lRef );
    ULONG cRef = InterlockedDecrement(&_lRef);
    if ( 0 == cRef ) 
    {
        delete this;
    }
    return cRef;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_SetDragOver(int iItem)
{
    if (_iDragOver >= 0)
    {
        ListView_SetItemState(_hwndList, _iDragOver, 0, LVIS_DROPHILITED);
    }

    _iDragOver = iItem;

    if (_iDragOver >= 0)
    {
        ListView_SetItemState(_hwndList, _iDragOver, LVIS_DROPHILITED, LVIS_DROPHILITED);

        _tmDragOver = NonzeroGetTickCount();
    }
    else
    {
        _tmDragOver = 0;
    }
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_ClearInnerDropTarget()
{
    if (_pdtDragOver)
    {
        ASSERT(_iDragState == DRAGSTATE_ENTERED);
        _pdtDragOver->DragLeave();
        _pdtDragOver->Release();
        _pdtDragOver = NULL;
#ifdef DEBUG
        _iDragState = DRAGSTATE_UNINITIALIZED;
#endif
    }

    ASSERT(_iDragState == DRAGSTATE_UNINITIALIZED); // 3668

    _SetDragOver(-1);
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
HRESULT SFTBarHost::_TryInnerDropTarget(int iItem, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
#ifdef DEAD_CODE
    HRESULT hr;

    if (_iDragOver != iItem)
    {
        _ClearInnerDropTarget();

        // Even if it fails, remember that we have this item so we don't
        // query for the drop target again (and have it fail again).
        _SetDragOver(iItem);

        ASSERT(_pdtDragOver == NULL);
        ASSERT(_iDragState == DRAGSTATE_UNINITIALIZED);

        PaneItem *pitem = _GetItemFromLV(iItem);
        if (pitem && pitem->IsDropTarget())
        {
            hr = _GetUIObjectOfItem(pitem, IID_PPV_ARGS(&_pdtDragOver));
            if (SUCCEEDED(hr))
            {
                hr = _pdtDragOver->DragEnter(_pdtoDragIn, grfKeyState, ptl, pdwEffect);
                if (SUCCEEDED(hr) && *pdwEffect)
                {
                }
                else
                {
                    ATOMICRELEASE(_pdtDragOver);
                }
            }
        }
    }

    ASSERT(_iDragOver == iItem);

    if (_pdtDragOver)
    {
        ASSERT(_iDragState == DRAGSTATE_ENTERED);
        hr = _pdtDragOver->DragOver(grfKeyState, ptl, pdwEffect);
    }
    else
    {
        hr = E_FAIL;            // No drop target
    }

    return hr;
#else
    IDropTarget **p_pdtDragOver; // edi
    PaneItem *pitem; // [esp+10h] [ebp-20h] MAPDST

    if (this->_iDragOver != iItem)
    {
        _ClearInnerDropTarget();
        _SetDragOver(iItem);

        //if (this->_pdtDragOver
        //    && CcshellAssertFailedW(L"d:\\longhorn\\shell\\explorer\\desktop2\\sfthost.cpp", 3704, L"_pdtDragOver == NULL", 0))
        //{
        //    AttachUserModeDebugger();
        //    do
        //        __debugbreak();
        //    while (dword_108B944);
        //}

        //if (this->_iDragState
        //    && CcshellAssertFailedW(
        //        L"d:\\longhorn\\shell\\explorer\\desktop2\\sfthost.cpp",
        //        3705,
        //        L"_iDragState == DRAGSTATE_UNINITIALIZED",
        //        0))
        //{
        //    AttachUserModeDebugger();
        //    do
        //        __debugbreak();
        //    while (dword_108B940);
        //}

        pitem = _GetItemFromLV(iItem);
        if (pitem)
        {
            if ((pitem->_dwFlags & 4) == 0)
                goto LABEL_18;
            p_pdtDragOver = &this->_pdtDragOver;
            if (SFTBarHost::_GetUIObjectOfItem(
                pitem,
                IID_IDropTarget,
                (void **)&this->_pdtDragOver) < 0)
                goto LABEL_18;
            if ((*p_pdtDragOver)->DragEnter(this->_pdtoDragIn, grfKeyState, ptl, pdwEffect) >= 0)
            {
                if (*pdwEffect)
                {
                    //this->_iDragState = 1;
                LABEL_18:
                    pitem->Release();
                    goto LABEL_19;
                }
                (*p_pdtDragOver)->DragLeave();
            }
            //this->_iDragState = 0;
            IUnknown_SafeReleaseAndNullPtr(&this->_pdtDragOver);
            goto LABEL_18;
        }
    }
LABEL_19:
    //if (this->_iDragOver != iItem
    //    && CcshellAssertFailedW(L"d:\\longhorn\\shell\\explorer\\desktop2\\sfthost.cpp", 3739, L"_iDragOver == iItem", 0))
    //{
    //    AttachUserModeDebugger();
    //    do
    //        __debugbreak();
    //    while (dword_108B93C);
    //}
    if (!this->_pdtDragOver)
        return 0x80004005;

    //if (this->_iDragState != 1
    //    && CcshellAssertFailedW(
    //        L"d:\\longhorn\\shell\\explorer\\desktop2\\sfthost.cpp",
    //        3743,
    //        L"_iDragState == DRAGSTATE_ENTERED",
    //        0))
    //{
    //    AttachUserModeDebugger();
    //    do
    //        __debugbreak();
    //    while (dword_108B938);
    //}
    return this->_pdtDragOver->DragOver(grfKeyState, ptl, pdwEffect);
#endif
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_PurgeDragDropData()
{
    _SetInsertMarkPosition(-1);
    _fForceArrowCursor = FALSE;
    _ClearInnerDropTarget();
    ATOMICRELEASE(_pdtoDragIn);
}

// *** IDropTarget::DragEnter ***

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::DragEnter(IDataObject *pdto, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
    if(_AreChangesRestricted())
    {
        *pdwEffect = DROPEFFECT_NONE;
        return S_OK;
    }
        
    POINT pt = { ptl.x, ptl.y };
    if (_pdth) {
        _pdth->DragEnter(_hwnd, pdto, &pt, *pdwEffect);
    }

    return _DragEnter(pdto, grfKeyState, ptl, pdwEffect);
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::_DragEnter(IDataObject *pdto, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
    _PurgeDragDropData();

    _fDragToSelf = SHIsSameObject(pdto, _pdtoDragOut);
    _fInsertable = IsInsertable(pdto);

    ASSERT(_pdtoDragIn == NULL);
    _pdtoDragIn = pdto;
    _pdtoDragIn->AddRef();

    return DragOver(grfKeyState, ptl, pdwEffect);
}

// *** IDropTarget::DragOver ***

HRESULT SFTBarHost::DragOver(DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
    if(_AreChangesRestricted())
    {
        *pdwEffect = DROPEFFECT_NONE;
        return S_OK;
    }
    
    _DebugConsistencyCheck();
    ASSERT(_pdtoDragIn);

    POINT pt = { ptl.x, ptl.y };
    if (_pdth) {
        _pdth->DragOver(&pt, *pdwEffect);
    }

    _fForceArrowCursor = FALSE;

    // Need to remember this because at the point of the drop, OLE gives
    // us the keystate after the user releases the button, so we can't
    // tell what kind of a drag operation the user performed!
    _grfKeyStateLast = grfKeyState;

#ifdef DEBUG
    if (_fDragToSelf)
    {
        ASSERT(_pdtoDragOut);
        ASSERT(_iDragOut >= 0);
        PaneItem *pitem = _GetItemFromLV(_iDragOut);
        ASSERT(pitem && (pitem->_iPos == _iPosDragOut));
    }
#endif

    //  Find the last item above the cursor position.  This allows us
    //  to treat the entire blank space at the bottom as belonging to the
    //  last item, and separators end up belonging to the item immediately
    //  above them.  Note that we don't bother testing item zero since
    //  he is always above everything (since he's the first item).

    ScreenToClient(_hwndList, &pt);

    POINT ptItem;
    int cItems = ListView_GetItemCount(_hwndList);
    int iItem;

    for (iItem = cItems - 1; iItem >= 1; iItem--)
    {
        ListView_GetItemPosition(_hwndList, iItem, &ptItem);
        if (ptItem.y <= pt.y)
        {
            break;
        }
    }

    //
    //  We didn't bother checking item 0 because we knew his position
    //  (by treating him special, this also causes all negative coordinates
    //  to be treated as belonging to item zero also).
    //
    if (iItem <= 0)
    {
        ptItem.y = 0;
        iItem = 0;
    }

    //
    //  Decide whether this is a drag-between or a drag-over...
    //
    //  For computational purposes, we treat each tile as four
    //  equal-sized "units" tall.  For each unit, we consider the
    //  possible actions in the order listed.
    //
    //  +-----
    //  |  0   insert above, drop on, reject
    //  | ----
    //  |  1                 drop on, reject
    //  | ----
    //  |  2                 drop on, reject
    //  | ----
    //  |  3   insert below, drop on, reject
    //  +-----
    //
    //  If the listview is empty, then treat as an
    //  insert before (imaginary) item zero; i.e., pin
    //  to top of the list.
    //

    UINT uUnit = 0;
    if (_cyTile && cItems)
    {
        int dy = pt.y - ptItem.y;

        // Peg out-of-bounds values to the nearest edge.
        if (dy < 0) dy = 0;
        if (dy >= _cyTile) dy = _cyTile - 1;

        // Decide which unit we are in.
        uUnit = 4 * dy / _cyTile;

        ASSERT(uUnit < 4);
    }

    //
    //  Now determine the appropriate action depending on which unit
    //  we are in.
    //

    int iInsert = -1;                   // Assume not inserting

    if (_fInsertable)
    {
        // Note!  Spec says that if you are in the non-pinned part of
        // the list, we draw the insert bar at the very bottom of
        // the pinned area.

        switch (uUnit)
        {
        case 0:
            iInsert = min(iItem, _cPinned);
            break;

        case 3:
            iInsert = min(iItem+1, _cPinned);
            break;
        }
    }

    //
    //  If inserting above or below isn't allowed, try dropping on.
    //
    if (iInsert < 0)
    {
        _SetInsertMarkPosition(-1);         // Not inserting

        // Up above, we let separators be hit-tested as if they
        // belongs to the item above them.  But that doesn't work for
        // drops, so reject them now.
        //
        // Also reject attempts to drop on the nonexistent item zero,
        // and don't let the user drop an item on itself.

        if (InRange(pt.y, ptItem.y, ptItem.y + _cyTile - 1) &&
            cItems &&
            !(_fDragToSelf && _iDragOut == iItem) &&
            SUCCEEDED(_TryInnerDropTarget(iItem, grfKeyState, ptl, pdwEffect)))
        {
            // Woo-hoo, happy joy!
        }
        else
        {
            // Note that we need to convert a failed drop into a DROPEFFECT_NONE
            // rather than returning a flat-out error code, because if we return
            // an error code, OLE will stop sending us drag/drop notifications!
            *pdwEffect = DROPEFFECT_NONE;
        }

        // If the user is hovering over a cascadable item, then open it.
        // First see if the user has hovered long enough...

        if (_tmDragOver && (GetTickCount() - _tmDragOver) >= _GetCascadeHoverTime())
        {
            _tmDragOver = 0;

            // Now see if it's cascadable
            PaneItem *pitem = _GetItemFromLV(_iDragOver);
            if (pitem && pitem->IsCascade())
            {
                // Must post this message because the cascading is modal
                // and we have to return a result to OLE
                PostMessage(_hwnd, SFTBM_CASCADE, _iDragOver, 0);
            }
        }
    }
    else
    {
        _ClearInnerDropTarget();    // Not dropping

        if (_fDragToSelf)
        {
            // Even though we're going to return DROPEFFECT_LINK,
            // tell the drag source (namely, ourselves) that we would
            // much prefer a regular arrow cursor because this is
            // a Move operation from the user's point of view.
            _fForceArrowCursor = TRUE;
        }

        //
        //  If user is dropping to a place where nothing would change,
        //  then don't draw an insert mark.
        //
        if (IsInsertMarkPointless(iInsert))
        {
            _SetInsertMarkPosition(-1);
        }
        else
        {
            _SetInsertMarkPosition(iInsert);
        }

        //  Sigh.  MergedFolder (used by the merged Start Menu)
        //  won't let you create shortcuts, so we pretend that
        //  we're copying if the data object doesn't permit
        //  linking.

        if (*pdwEffect & DROPEFFECT_LINK)
        {
            *pdwEffect = DROPEFFECT_LINK;
        }
        else
        {
            *pdwEffect = DROPEFFECT_COPY;
        }
    }

    return S_OK;
}

// *** IDropTarget::DragLeave ***

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::DragLeave()
{
    if(_AreChangesRestricted())
    {
        return S_OK;
    }
    
    if (_pdth) {
        _pdth->DragLeave();
    }

    _PurgeDragDropData();
    return S_OK;
}

// *** IDropTarget::Drop ***

// EXEX-VISTA(allison): Validated. Still needs minor cleanup.
HRESULT SFTBarHost::Drop(IDataObject *pdto, DWORD grfKeyState, POINTL ptl, DWORD *pdwEffect)
{
#ifdef DEAD_CODE
    if (_AreChangesRestricted())
    {
        *pdwEffect = DROPEFFECT_NONE;
        return S_OK;
    }

    _DebugConsistencyCheck();

    // Use the key state from the last DragOver call
    grfKeyState = _grfKeyStateLast;

    // Need to go through the whole _DragEnter thing again because who knows
    // maybe the data object and coordinates of the drop are different from
    // the ones we got in DragEnter/DragOver...  We use _DragEnter, which
    // bypasses the IDropTargetHelper::DragEnter.
    //
    _DragEnter(pdto, grfKeyState, ptl, pdwEffect);

    POINT pt = { ptl.x, ptl.y };
    if (_pdth) {
        _pdth->Drop(pdto, &pt, *pdwEffect);
    }

    int iInsert = _iInsert;
    _SetInsertMarkPosition(-1);

    if (*pdwEffect)
    {
        ASSERT(_pdtoDragIn);
        if (iInsert >= 0)                           // "add to pin" or "move"
        {
            BOOL fTriedMove = FALSE;

            // First see if it was just a move of an existing pinned item
            if (_fDragToSelf)
            {
                PaneItem *pitem = _GetItemFromLV(_iDragOut);
                if (pitem)
                {
                    if (pitem->IsPinned())
                    {
                        // Yup, it was a move - so move it.
                        if (SUCCEEDED(MovePinnedItem(pitem, iInsert)))
                        {
                            // We used to try to update all the item positions
                            // incrementally.  This was a major pain in the neck.
                            //
                            // So now we just do a full refresh.  Turns out that a
                            // full refresh is fast enough anyway.
                            //
                            PostMessage(_hwnd, SFTBM_REFRESH, TRUE, 0);
                        }

                        // We tried to move a pinned item (return TRUE even if
                        // we actually failed).
                        fTriedMove = TRUE;
                    }
                }
            }

            if (!fTriedMove)
            {
                if (SUCCEEDED(InsertPinnedItem(_pdtoDragIn, iInsert)))
                {
                    PostMessage(_hwnd, SFTBM_REFRESH, TRUE, 0);
                }
            }
        }
        else if (_pdtDragOver) // Not an insert, maybe it was a plain drop
        {
            ASSERT(_iDragState == DRAGSTATE_ENTERED);
            _pdtDragOver->Drop(_pdtoDragIn, grfKeyState, ptl, pdwEffect);
        }
    }

    _PurgeDragDropData();
    _DebugConsistencyCheck();

    return S_OK;
#else
    if (_AreChangesRestricted())
    {
        *pdwEffect = 0;
        return S_OK;
    }

    grfKeyState = _grfKeyStateLast;
    _DragEnter(pdto, grfKeyState, ptl, pdwEffect);

    POINT pt = { ptl.x, ptl.y };
    if (_pdth)
    {
        _pdth->Drop(pdto, &pt, *pdwEffect);
    }

    int iInsert = _iInsert;
    _SetInsertMarkPosition(-1);

    if (*pdwEffect)
    {
        ASSERT(_pdtoDragIn); // 4087

        if (iInsert >= 0)
        {
            BOOL fTriedMove = 0;

            if (_fDragToSelf)
            {
                PaneItem *pitem = _GetItemFromLV(_iDragOut);
                if (pitem)
                {
                    if (pitem->_iPinPos >= 0)
                    {
                        if (MovePinnedItem(pitem, iInsert) >= 0)
                        {
                            PostMessage(_hwnd, 0x40Au, 1, 0);
                        }

                        fTriedMove = 1;
                    }
                    pitem->Release();
                }
            }

            if (!fTriedMove && InsertPinnedItem(_pdtoDragIn, iInsert) >= 0)
            {
                PostMessage(_hwnd, 0x40Au, 1, 0);
            }
        }
        else if (_pdtDragOver)
        {
            ASSERT(_iDragState == DRAGSTATE_ENTERED); // 4131
            _pdtDragOver->Drop(_pdtoDragIn, grfKeyState, ptl, pdwEffect);
            IUnknown_SafeReleaseAndNullPtr(&_pdtDragOver);
#ifdef DEBUG
            _iDragState = DRAGSTATE_UNINITIALIZED;
#endif
        }
    }

    _PurgeDragDropData();
    return S_OK;
#endif
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_SetInsertMarkPosition(int iInsert)
{
    if (_iInsert != iInsert)
    {
        _InvalidateInsertMark();
        _iInsert = iInsert;
        _InvalidateInsertMark();
    }
}

// EXEX-VISTA(allison): Validated.
BOOL SFTBarHost::_GetInsertMarkRect(LPRECT prc)
{
    if (_iInsert >= 0)
    {
        GetClientRect(_hwndList, prc);
        POINT pt;
        _ComputeListViewItemPosition(_iInsert, &pt);
        int cyEdge = SHGetSystemMetricsScaled(SM_CYEDGE);
        prc->bottom = pt.y + cyEdge;
        prc->top = pt.y - cyEdge;
        return TRUE;
    }
    
    return FALSE;
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_InvalidateInsertMark()
{
    RECT rc;
    if (_GetInsertMarkRect(&rc))
    {
        InvalidateRect(_hwndList, &rc, TRUE);
    }
}

// EXEX-VISTA(allison): Validated.
void SFTBarHost::_DrawInsertionMark(LPNMLVCUSTOMDRAW plvcd)
{
    RECT rc;
    if (_GetInsertMarkRect(&rc))
    {
        FillRect(plvcd->nmcd.hdc, &rc, GetSysColorBrush(COLOR_WINDOWTEXT));
    }
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
int SFTBarHost::_CalcMaxTextWith()
{
    WCHAR chText[64]; // [esp+6Ch] [ebp-84h] BYREF

    HWND hwndList = this->_hwndList;
    int v20 = 0;
    HDC WindowDC = GetWindowDC(hwndList);
    HDC hdc = WindowDC;
    void* v3 = (void*)SendMessageW(this->_hwndList, WM_GETFONT, 0, 0);
    HGDIOBJ v4 = SelectObject(hdc, v3);
    HGDIOBJ h = v4;
    LRESULT v5 = SendMessageW(this->_hwndList, LVM_GETITEMCOUNT, 0, 0);
    while (--v5 >= 0)
    {
        if (hdc)
        {
            RECT rc;
            if (ListView_GetItemRect(this->_hwndList, v5, &rc, LVIR_LABEL))
            {
                PaneItem* pitem = _GetItemFromLV(v5);
                if (!pitem->CanItemWrap())
                {
                    LVITEMW lvi; // [esp+Ch] [ebp-E4h] BYREF
                    lvi.iSubItem = 0;
                    int v18 = this->_cxMarlett + rc.left + 4;
                    lvi.pszText = chText;
                    lvi.cchTextMax = 64;
                    SendMessageW(this->_hwndList, LVM_GETITEMTEXTW, v5, (LPARAM)&lvi);
                    DrawTextW(hdc, chText, -1, &rc, 0x400u);
                    if (v20 <= v18 + rc.right)
                    {
                        v20 = v18 + rc.right;
                    }
                }
                pitem->Release();
            }
        }
    }

    SelectObject(hdc, h);
    ReleaseDC(this->_hwndList, hdc);
    return v20;
}

// EXEX-VISTA(allison): Validated.
LRESULT SFTBarHost::GetLVText(const PaneItem *pitem, LPWSTR pszText, DWORD cch)
{
    HRESULT hr = E_FAIL;

    LVFINDINFO lvfi;
    lvfi.lParam = reinterpret_cast<LPARAM>(const_cast<PaneItem *>(pitem));
    lvfi.flags = LVIF_TEXT;

    int iItem = ListView_FindItem(_hwndList, -1, &lvfi);
    if (iItem >= 0)
    {
        ListView_GetItemText(_hwndList, iItem, 0, pszText, cch);
        return iItem;
    }
    return hr;
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void SFTBarHost::_DrawSeparator(HDC hdc, int x, int y)
{
#ifdef DEAD_CODE
    RECT rc;
    rc.left = x;
    rc.top = y;
    rc.right = rc.left + _cxTile;
    rc.bottom = rc.top + _cySep;

    if (!_hTheme)
    {
        DrawEdge(hdc, &rc, EDGE_ETCHED,BF_TOPLEFT);
    }
    else
    {
        DrawThemeBackground(_hTheme, hdc, _iThemePartSep, 0, &rc, 0);
    }
#else
    RECT rc; // [esp+4h] [ebp-10h] BYREF
    LONG v4 = x + this->_cxTile;
    rc.left = x;
    rc.bottom = y + this->_cySep;
    rc.right = v4;
    rc.top = y;
    if (_hTheme)
        DrawThemeBackground(_hTheme, hdc, this->_iThemePartSep, 0, &rc, 0);
    else
        DrawEdge(hdc, &rc, EDGE_ETCHED, BF_TOPLEFT);
#endif
}

// EXEX-VISTA(allison): Validated. Still needs cleanup.
void SFTBarHost::_DrawSeparators(LPNMLVCUSTOMDRAW plvcd)
{
#ifdef DEAD_CODE
    POINT pt;
    RECT rc;

    for (int iSep = 0; iSep < _cSep; iSep++)
    {
        _ComputeListViewItemPosition(_rgiSep[iSep], &pt);
        pt.y = pt.y - _cyTilePadding + (_cySepTile - _cySep + _cyTilePadding)/2;
        _DrawSeparator(plvcd->nmcd.hdc, pt.x, pt.y);
    }

    // Also draw a bonus separator at the bottom of the list to separate
    // the MFU list from the More Programs button.

    if (_iThemePart == SPP_PROGLIST)
    {
        _ComputeListViewItemPosition(0, &pt);
        GetClientRect(_hwndList, &rc);
        rc.bottom -= _cySep;
        _DrawSeparator(plvcd->nmcd.hdc, pt.x, rc.bottom);

    }
#else
    int v3; // ebx
    int* rgiSep; // edi
    struct tagRECT rc; // [esp+8h] [ebp-18h] BYREF
    struct tagPOINT pt; // [esp+18h] [ebp-8h] BYREF

    v3 = 0;
    if (this->_cSep > 0)
    {
        rgiSep = this->_rgiSep;
        do
        {
            SFTBarHost::_ComputeListViewItemPosition(*rgiSep, &pt);
            pt.y += (this->_cySepTile - this->_cySep) / 2;
            SFTBarHost::_DrawSeparator(plvcd->nmcd.hdc, pt.x, pt.y);
            ++v3;
            ++rgiSep;
        } while (v3 < this->_cSep);
    }

    if (this->_iThemePart == 4)
    {
        SFTBarHost::_ComputeListViewItemPosition(0, &pt);
        GetClientRect(this->_hwndList, &rc);
        SFTBarHost::_DrawSeparator(plvcd->nmcd.hdc, pt.x, rc.bottom - this->_cySep);
    }
#endif
}

//****************************************************************************
//
//  Accessibility
//

// EXEX-VISTA(allison): Validated.
PaneItem *SFTBarHost::_GetItemFromAccessibility(const VARIANT& varChild)
{
    if (varChild.lVal)
    {
        return _GetItemFromLV(varChild.lVal - 1);
    }
    return NULL;
}

//
//  The default accessibility object reports listview items as
//  ROLE_SYSTEM_LISTITEM, but we know that we are really a menu.
//
//  Our items are either ROLE_SYSTEM_MENUITEM or ROLE_SYSTEM_MENUPOPUP.
//

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::get_accRole(VARIANT varChild, VARIANT *pvarRole)
{
    HRESULT hr = _paccInner->get_accRole(varChild, pvarRole);
    if (SUCCEEDED(hr) && V_VT(pvarRole) == VT_I4)
    {
        switch (V_I4(pvarRole))
        {
        case ROLE_SYSTEM_LIST:
            V_I4(pvarRole) = ROLE_SYSTEM_MENUPOPUP;
            break;

        case ROLE_SYSTEM_LISTITEM:
            V_I4(pvarRole) = ROLE_SYSTEM_MENUITEM;
            break;
        }
    }
    return hr;
}

HRESULT SFTBarHost::get_accState(VARIANT varChild, VARIANT *pvarState)
{
    HRESULT hr = _paccInner->get_accState(varChild, pvarState);
    if (SUCCEEDED(hr) && V_VT(pvarState) == VT_I4)
    {
        PaneItem *pitem = _GetItemFromAccessibility(varChild);
        if (pitem && pitem->IsCascade())
        {
            V_I4(pvarState) |= STATE_SYSTEM_HASPOPUP;
        }

    }
    return hr;
}

HRESULT SFTBarHost::get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut)
{
    if (varChild.lVal)
    {
        PaneItem *pitem = _GetItemFromAccessibility(varChild);
        if (pitem)
        {
            return CreateAcceleratorBSTR(GetItemAccelerator(pitem, varChild.lVal - 1), pszKeyboardShortcut);
        }
    }
    *pszKeyboardShortcut = NULL;
    return E_NOT_APPLICABLE;
}


//
//  Default action for cascading menus is Open/Close (depending on
//  whether the item is already open); for regular items
//  is Execute.
//
HRESULT SFTBarHost::get_accDefaultAction(VARIANT varChild, BSTR *pszDefAction)
{
    *pszDefAction = NULL;

    ASSERT(pszDefAction); // 4365

    HRESULT hr = E_NOT_APPLICABLE;

    if (varChild.lVal)
    {
        PaneItem *pitem = _GetItemFromAccessibility(varChild);
        if (pitem)
        {
            if (pitem->IsCascade())
            {
                DWORD dwRole = varChild.lVal - 1 == _iCascading ? ACCSTR_CLOSE : ACCSTR_OPEN;
                hr = GetRoleString(dwRole, pszDefAction);
            }
            pitem->Release();
        }

        if (FAILED(hr))
        {
            hr = GetRoleString(ACCSTR_EXECUTE, pszDefAction);
        }
    }
    return hr;
}

// EXEX-VISTA(allison): Validated.
HRESULT SFTBarHost::accDoDefaultAction(VARIANT varChild)
{
    HRESULT hr = E_FAIL;
    
    if (varChild.lVal)
    {
        PaneItem* pitem = _GetItemFromAccessibility(varChild);
        if (pitem)
        {
            if (pitem->IsCascade() && varChild.lVal - 1 == _iCascading)
            {
                _SendNotify(_hwnd, SMN_CANCELSHELLMENU);
                hr = S_OK;
            }

            pitem->Release();

            if (SUCCEEDED(hr))
            {
                return hr;
            }
        }
    }
    return CAccessible::accDoDefaultAction(varChild);
}



//****************************************************************************
//
//  Debugging helpers
//

#ifdef FULL_DEBUG

void SFTBarHost::_DebugConsistencyCheck()
{
    int i;
    int citems;

    if (_hwndList && !_fListUnstable)
    {
        //
        //  Check that the items in the listview are in their correct positions.
        //

        citems = ListView_GetItemCount(_hwndList);
        for (i = 0; i < citems; i++)
        {
            PaneItem *pitem = _GetItemFromLV(i);
            if (pitem)
            {
                // Make sure the item number and the iPos are in agreement
                ASSERT(pitem->_iPos == _ItemNoToPos(i));
                ASSERT(_PosToItemNo(pitem->_iPos) == i);

                // Make sure the item is where it should be
                POINT pt, ptShould;
                _ComputeListViewItemPosition(pitem->_iPos, &ptShould);
                ListView_GetItemPosition(_hwndList, i, &pt);
                ASSERT(pt.x == ptShould.x);
                ASSERT(pt.y == ptShould.y);
            }
        }
    }

}
#endif

//  iFile is the zero-based index of the file being requested
//        or 0xFFFFFFFF if you don't care about any particular file
//
//  puFiles receives the number of files in the HDROP
//        or NULL if you don't care about the number of files
//

STDAPI_(HRESULT)
IDataObject_DragQueryFile(IDataObject *pdto, UINT iFile, LPTSTR pszBuf, UINT cch, UINT *puFiles)
{
    static FORMATETC const feHdrop =
        { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stgm;
    HRESULT hr;

    // Sigh.  IDataObject::GetData has a bad prototype and says that
    // the first parameter is a modifiable FORMATETC, even though it
    // isn't.
    hr = pdto->GetData(const_cast<FORMATETC*>(&feHdrop), &stgm);
    if (SUCCEEDED(hr))
    {
        HDROP hdrop = reinterpret_cast<HDROP>(stgm.hGlobal);
        if (puFiles)
        {
            *puFiles = DragQueryFile(hdrop, 0xFFFFFFFF, NULL, 0);
        }

        if (iFile != 0xFFFFFFFF)
        {
            hr = DragQueryFile(hdrop, iFile, pszBuf, cch) ? S_OK : E_FAIL;
        }
        ReleaseStgMedium(&stgm);
    }
    return hr;
}

/*
 * If pidl has an alias, free the original pidl and return the alias.
 * Otherwise, just return pidl unchanged.
 *
 * Expected usage is
 *
 *          pidlTarget = ConvertToLogIL(pidlTarget);
 *
 */
STDAPI_(LPITEMIDLIST) ConvertToLogIL(LPITEMIDLIST pidl)
{
    LPITEMIDLIST pidlAlias = SHLogILFromFSIL(pidl);
    if (pidlAlias)
    {
        ILFree(pidl);
        return pidlAlias;
    }
    return pidl;
}

//****************************************************************************
//

STDAPI_(HFONT) LoadControlFont(HTHEME hTheme, int iPart, BOOL fUnderline, DWORD dwSizePercentage)
{
    LOGFONT lf;
    BOOL bSuccess;

    if (hTheme)
    {
        bSuccess = SUCCEEDED(GetThemeFont(hTheme, NULL, iPart, 0, TMT_FONT, &lf));
    }
    else
    {
        bSuccess = SystemParametersInfo(SPI_GETICONTITLELOGFONT, sizeof(lf), &lf, FALSE);
    }

    if (bSuccess)
    {
        // only apply size scaling factor in non-theme case, for themes it makes sense to specify the exact font in the theme
        if (!hTheme && dwSizePercentage && dwSizePercentage != 100)
        {
            lf.lfHeight = (lf.lfHeight * (int)dwSizePercentage) / 100;
            lf.lfWidth = 0; // get the closest based on aspect ratio
        }

        if (fUnderline)
        {
            lf.lfUnderline = TRUE;
        }

       return CreateFontIndirect(&lf);
    }
    return NULL;
}
