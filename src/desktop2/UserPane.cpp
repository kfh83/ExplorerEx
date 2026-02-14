#include "pch.h"

#include "shundoc.h"
#include "stdafx.h"
#include "sfthost.h"
#include "userpane.h"

#define REGSTR_VAL_DV2_STARTPANEL_FADEIN TEXT("StartPanel_FadeIn")
#define REGSTR_VAL_DV2_STARTPANEL_FADEOUT TEXT("StartPanel_FadeOut")
#define REGSTR_VAL_DV2_STARTPANEL_FADEDELAY TEXT("StartPanel_FadeDelay")

CUserPane::CUserPane()
    : _lRef(1)
    , _fadeA(-1)
    , _fadeB(-1)
{
    ASSERT(_hwnd == NULL); // 19
    ASSERT(_hbmUserPicture == NULL); // 20

    DWORD cbData = sizeof(DWORD);
    _dwFadeIn = 350;
    _SHRegGetValueFromHKCUHKLM(
        DV2_REGPATH, REGSTR_VAL_DV2_STARTPANEL_FADEIN, SRRF_RT_ANY, NULL, &_dwFadeIn, &cbData);

    cbData = sizeof(DWORD);
    _dwFadeOut = 250;
    _SHRegGetValueFromHKCUHKLM(
        DV2_REGPATH, REGSTR_VAL_DV2_STARTPANEL_FADEOUT, SRRF_RT_ANY, NULL, &_dwFadeOut, &cbData);

    cbData = sizeof(DWORD);
    _dwFadeDelay = 400;
    _SHRegGetValueFromHKCUHKLM(
        DV2_REGPATH, REGSTR_VAL_DV2_STARTPANEL_FADEDELAY, SRRF_RT_ANY, NULL, &_dwFadeDelay, &cbData);
}

CUserPane::~CUserPane()
{
    if (_hbmUserPicture)
        DeleteObject(_hbmUserPicture);

    if (_pgdipImage)
        delete _pgdipImage;
}

HRESULT CUserPane::QueryInterface(REFIID riid, void **ppvObj)
{
    static const QITAB qit[] = {
        QITABENT(CUserPane, IServiceProvider),
        QITABENT(CUserPane, IOleCommandTarget),
        QITABENT(CUserPane, IObjectWithSite),
        { 0 },
    };
    return QISearch(this, qit, riid, ppvObj);
}

ULONG CUserPane::AddRef()
{
    return InterlockedIncrement(&_lRef);
}

ULONG CUserPane::Release()
{
    ASSERT(0 != _lRef); // 261
    LONG lRef = InterlockedDecrement(&_lRef);
    if (lRef == 0 && this)
    {
        delete this;
    }
    return lRef;
}

HRESULT CUserPane::SetSite(IUnknown *punkSite)
{
    return CObjectWithSite::SetSite(punkSite);
}

HRESULT CUserPane::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    if (IsEqualGUID(guidService, SID_SM_UserPane))
    {
        HRESULT hr = QueryInterface(riid, ppvObject);
        if (SUCCEEDED(hr))
        {
            return hr;
        }
    }
    return IUnknown_QueryService(_punkSite, guidService, riid, ppvObject);
}

HRESULT CUserPane::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText)
{
    return E_NOTIMPL;
}

HRESULT CUserPane::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut)
{
    HRESULT hr = E_INVALIDARG;

    if (IsEqualGUID(SID_SM_DV2ControlHost, *pguidCmdGroup))
    {
        switch (nCmdID)
        {
            case 313:
            {
                ASSERT(V_VT(pvarargIn) == VT_I4); // 298

                LONG lVal = pvarargIn->lVal;
                if (_fadeC != lVal)
                {
                    _fadeC = lVal;
                    KillTimer(_hwndStatic, 1);
                    SetTimer(_hwndStatic, 1, _dwFadeDelay, 0);
                }
                break;
            }
            case 314:
                ASSERT(V_VT(pvarargIn) == VT_BYREF); // 310
                _himl = (HIMAGELIST)pvarargIn->byref;
                break;
            case 323u:
                _HidePictureWindow();
                break;
            default:
                return hr;
        }
        return 0;
    }
    return hr;
}

LRESULT CALLBACK CUserPane::s_WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CUserPane *pThis = reinterpret_cast<CUserPane *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    if (pThis)
        return pThis->WndProcPane(hwnd, uMsg, wParam, lParam);

    if (uMsg != WM_NCDESTROY)
    {
        pThis = new CUserPane();
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pThis);
        if (pThis)
        {
            return pThis->WndProcPane(hwnd, uMsg, wParam, lParam);
        }
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void CUserPane::_UpdatePictureWindow(BYTE a2, BYTE a3)
{
    if (IsWindowVisible(_hwndStatic) || _fadeB == -1)
    {
        HDC hdc = GetDC(_hwndStatic);
        _PaintPictureWindow(hdc, a2, a3);
        ReleaseDC(_hwndStatic, hdc);
    }
}

void CUserPane::_UpdateDC(HDC hdc, int iIndex, BYTE a4)
{
    HDC hDC = GetDC(_hwndStatic);
    HDC hMemDC = CreateCompatibleDC(hDC);

    HGDIOBJ hOldObj = 0;
    HGDIOBJ hbmTemp = 0;

    if (hMemDC)
    {
        if (_himl && iIndex != -1)
        {
            RECT rc;
            GetClientRect(_hwndStatic, &rc);

            BITMAPINFO bmi = { 0 };
            bmi.bmiHeader.biWidth = rc.right;
            bmi.bmiHeader.biHeight = rc.bottom;
            bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = 0;

            void *pvBits;
            hbmTemp = CreateDIBSection(hMemDC, &bmi, 0, &pvBits, 0, 0);
            if (hbmTemp)
            {
                hOldObj = SelectObject(hMemDC, hbmTemp);
            }

            int imageCount = ImageList_GetImageCount(_himl);
            printf("ImageList count: %d, drawing index: %d\n", imageCount, iIndex);
            BOOL result = ImageList_DrawEx(_himl, iIndex, hMemDC, 0, 0, RECTWIDTH(rc), RECTHEIGHT(rc), 0xFFFFFFFF, 0xFFFFFFFF, 0x12001u);
            printf("ImageList_DrawEx result: %d\n", result);
        }
        else if (_hbmUserPicture)
        {
            RECT rc;
            GetClientRect(_hwndStatic, &rc);

            BITMAPINFO bmi = { 0 };
            bmi.bmiHeader.biWidth = rc.right;
            bmi.bmiHeader.biHeight = rc.bottom;
            bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
            bmi.bmiHeader.biPlanes = 1;
            bmi.bmiHeader.biBitCount = 32;
            bmi.bmiHeader.biCompression = BI_RGB;

            void *pvBits;
            hbmTemp = CreateDIBSection(hMemDC, &bmi, 0, &pvBits, 0, 0);
            if (hbmTemp)
            {
                hOldObj = SelectObject(hMemDC, hbmTemp);
            }

            int cxLeftWidth = field_1C;
            int cyTopHeight = field_24;
            int v9 = rc.right - _iFramedPicWidth - rc.left;
            int v22 = (rc.bottom - rc.top - _iFramedPicHeight) / 2;
            int v21 = v9 / 2;
            int v10 = v9 / 2 + cxLeftWidth;
            int v11 = v22 + cyTopHeight;

            HDC hdcTemp = CreateCompatibleDC(hMemDC);
            if (hdcTemp)
            {
                HGDIOBJ v24 = SelectObject(hdcTemp, _hbmUserPicture);
                BitBlt(hMemDC, v10, v11, _iUnframedPicWidth, _iUnframedPicHeight, hdcTemp, 0, 0, SRCCOPY);
                SelectObject(hdcTemp, v24);
                DeleteDC(hdcTemp);
            }

            if ((_iFramedPicWidth != 48 || _iFramedPicHeight != 48) && _pgdipImage)
            {
                int v13 = rc.right - v21;
                int v14 = rc.bottom - v22;

                Gdiplus::Graphics graphics(hMemDC);
                graphics.SetCompositingMode(Gdiplus::CompositingModeSourceOver);
                graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBilinear);
                graphics.DrawImage(_pgdipImage, v21, v22, v13 - v21, v14 - v22);
            }
        }

        BLENDFUNCTION bf;
        bf.BlendOp = AC_SRC_OVER;
        bf.BlendFlags = 0;
        bf.SourceConstantAlpha = a4;
        bf.AlphaFormat = 0;
        GdiAlphaBlend(hdc, 0, 0, _iFramedPicWidth, _iFramedPicHeight, hMemDC, 0, 0, _iFramedPicWidth, _iFramedPicHeight, bf);

        if (hOldObj)
            SelectObject(hMemDC, hOldObj);

        if (hbmTemp)
            DeleteObject(hbmTemp);

        DeleteDC(hMemDC);
    }
    ReleaseDC(_hwndStatic, hDC);
}

void CUserPane::_PaintPictureWindow(HDC hdc, BYTE a3, BYTE a4)
{
    HDC hMemDC = CreateCompatibleDC(hdc);
    HGDIOBJ hOldObj = NULL;

    if (hMemDC)
    {
        RECT rc;
        GetClientRect(_hwndStatic, &rc);

        BITMAPINFO bmi = { 0 };
        bmi.bmiHeader.biWidth = rc.right;
        bmi.bmiHeader.biHeight = rc.bottom;
        bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;

        void *pvBits;
        HBITMAP hbm = CreateDIBSection(hMemDC, &bmi, 0, &pvBits, NULL, 0);
        if (hbm)
            hOldObj = SelectObject(hMemDC, hbm);

        if (a3)
        {
            _UpdateDC(hMemDC, _fadeB, a3);
        }

        if (a4)
        {
            _UpdateDC(hMemDC, _fadeA, a4);
        }

        POINT pt;
        pt.x = 0;
        pt.y = 0;

        SIZE sizFramedPic;
        sizFramedPic.cx = _iFramedPicWidth;
        sizFramedPic.cy = _iFramedPicHeight;

        BLENDFUNCTION bf;
        bf.BlendOp = 0;
        bf.BlendFlags = 0;
        bf.SourceConstantAlpha = -1;
        bf.AlphaFormat = AC_SRC_ALPHA;
        UpdateLayeredWindow(_hwndStatic, hdc, NULL, &sizFramedPic, hMemDC, &pt, 0, &bf, LWA_ALPHA);

        if (hOldObj)
            SelectObject(hMemDC, hOldObj);

        if (hbm)
            DeleteObject(hbm);

        DeleteDC(hMemDC);
    }
}

void CUserPane::_HidePictureWindow()
{
    if (_fadeA != -1 || _fadeB != -1)
    {
        _fadeA = -1;
        _fadeB = -1;
        _UpdatePictureWindow(0xFF, 0);
    }

    EnableWindow(_hwndStatic, FALSE);
    SetWindowPos(
        _hwndStatic, nullptr, 0, 0, 0, 0,
        SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW | SWP_NOSENDCHANGING);
}

LRESULT CUserPane::_OnNcCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    SMPANEDATA *psmpd = PaneDataFromCreateStruct(lParam);

    _hwnd = hwnd;
    _hTheme = psmpd->hTheme;
    IUnknown_Set(&psmpd->punk, static_cast<IServiceProvider*>(this));
    if (SUCCEEDED(_UpdateUserInfo(1)))
    {
        _UpdatePictureWindow(0xFF, 0);
        return 1;
    }
    return -1;
}

HWND CUserPane::_GetPictureWindowPrevHnwd()
{
    HWND hwnd = NULL;

    SMNGETISTARTBUTTON nm;
    nm.pstb = NULL;
    _SendNotify(_hwnd, 218, &nm.hdr);
    if (nm.pstb)
    {
        nm.pstb->GetWindow(&hwnd);
		nm.pstb->Release();
    }

    if (hwnd)
        return 0;
    else
        return GetWindow(0, GW_HWNDPREV);
}

LRESULT CALLBACK CUserPane::WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_DESTROY:
        {
            if (_uidChangeRegister)
            {
                SHChangeNotifyDeregister(_uidChangeRegister);
                _uidChangeRegister = 0;
            }
            if (_hwndStatic)
            {
                DestroyWindow(_hwndStatic);
                _hwndStatic = nullptr;
            }
            break;
        }
        case WM_ERASEBKGND:
        {
            RECT rc;
            GetClientRect(_hwnd, &rc);
            if (_hTheme)
            {
                DrawPlacesListBackground(_hTheme, hwnd, (HDC)wParam);
            }
            else
            {
                SHFillRectClr((HDC)wParam, &rc, GetSysColor(COLOR_MENU));
                DrawEdge((HDC)wParam, &rc, EDGE_ETCHED, BF_LEFT);
            }
            return 1;
        }
        case WM_WINDOWPOSCHANGED:
        {
            LPWINDOWPOS pwp = reinterpret_cast<LPWINDOWPOS>(lParam);
            if (pwp)
            {
                if (pwp->flags & SWP_SHOWWINDOW)
                {
                    VARIANT vt;
                    vt.vt = VT_BOOL;
                    IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 327, 0, nullptr, &vt);
                    if (vt.boolVal == VARIANT_FALSE)
                    {
                        EnableWindow(_hwndStatic, TRUE);
                        SetWindowPos(
                            _hwndStatic, nullptr, 0, 0, 0, 0,
                            SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOSENDCHANGING);
                    }
                }
                else if (pwp->flags & SWP_HIDEWINDOW)
                {
                    EnableWindow(_hwndStatic, 0);
                    SetWindowPos(
                        _hwndStatic, nullptr, 0, 0, 0, 0,
                        SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_HIDEWINDOW | SWP_NOSENDCHANGING);
                }

                if ((pwp->flags & (SWP_NOSIZE | SWP_NOMOVE)) != (SWP_NOSIZE | SWP_NOMOVE))
                {
                    OnSize();
                }
            }
            break;
        }
        case WM_NOTIFY:
        {
            LPNMHDR pnm = reinterpret_cast<LPNMHDR>(lParam);
            switch (pnm->code)
            {
                case 208:
                {
                    VARIANT vt;
                    vt.vt = VT_BOOL;
                    IUnknown_QueryServiceExec(_punkSite, SID_SMenuPopup, &SID_SM_DV2ControlHost, 327, 0, nullptr, &vt);
                    if (vt.boolVal == VARIANT_FALSE)
                    {
                        EnableWindow(_hwndStatic, TRUE);
                        SetWindowPos(
                            _hwndStatic, _GetPictureWindowPrevHnwd(), 0, 0, 0, 0,
                            SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING);
                    }
                    break;
                }
                case 221:
                {
                    OnSize();
                    break;
                }
                case 223:
                {
                    if (SUCCEEDED(SetSite(((SMNSETSITE *)pnm)->punkSite)))
                    {
                        return 1;
                    }
                    break;
                }
            }
            break;
        }
        case WM_NCCREATE:
        {
            return _OnNcCreate(hwnd, uMsg, wParam, lParam);
        }

        case WM_NCDESTROY:
        {
            SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
            Release();
            break;
        }
        case UPM_CHANGENOTIFY:
        {
            LPITEMIDLIST* pppidl = 0;
            LONG plEvent = 0;
            HANDLE v5 = SHChangeNotification_Lock((HANDLE)wParam, (DWORD)lParam, &pppidl, &plEvent);
            if (v5)
            {
                if (plEvent == 0x4000000 && *pppidl && *(DWORD*)((*pppidl)->mkid.abID) == 11)
                {
                    _UpdateUserInfo(0);
                    _UpdatePictureWindow(0xFF, 0);
                }
                SHChangeNotification_Unlock(v5);
            }
            break;
        }
    }
    return ::DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CUserPane::s_WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CUserPane *pThis = reinterpret_cast<CUserPane *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    if (uMsg == 0x81)
    {
        LPCREATESTRUCT lpcs = (LPCREATESTRUCT)lParam;
        pThis = (CUserPane *)lpcs->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pThis);
    }
    
    if (pThis)
        return pThis->WndProcPicture(hwnd, uMsg, wParam, lParam);
    else
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void CUserPane::_DoFade()
{
    LARGE_INTEGER Frequency; // [esp+8h] [ebp-28h] BYREF
    LARGE_INTEGER PerformanceCount; // [esp+10h] [ebp-20h] BYREF
    LARGE_INTEGER v10; // [esp+18h] [ebp-18h] BYREF
    BYTE v13; // [esp+28h] [ebp-8h]
    BYTE a2; // [esp+2Ch] [ebp-4h]

    QueryPerformanceFrequency(&Frequency);
    QueryPerformanceCounter(&PerformanceCount);
    
    LONG* p_fadeA = &this->_fadeA;
    LONG fadeA = this->_fadeA;
    LONG fadeB = this->_fadeB;
    LONG v4 = this->_fadeA;
    LONG v12 = fadeA;
    if (fadeA == v4)
    {
        while (fadeB == this->_fadeB)
        {
            QueryPerformanceCounter(&v10);
            v10.QuadPart -= PerformanceCount.QuadPart;
            LONGLONG v5 = 1000 * v10.QuadPart / Frequency.QuadPart;
            if ((unsigned int)v5 >= this->_dwFadeOut && (unsigned int)v5 >= this->_dwFadeIn)
            {
                CUserPane::_UpdatePictureWindow(0xFF, 0);
                break;
            }
            
            DWORD dwFadeIn = this->_dwFadeIn;
            if ((unsigned int)v5 >= dwFadeIn)
                v13 = 0;
            else
                v13 = 255 * (int)v5 / dwFadeIn;
            if ((unsigned int)v5 >= this->_dwFadeOut)
                a2 = 0;
            else
                a2 = (unsigned int)(255 * (this->_dwFadeOut - v5)) / this->_dwFadeOut;
            CUserPane::_UpdatePictureWindow(v13, a2);
            Sleep(0xAu);
            if (v12 != *p_fadeA)
                break;
        }
    }
    if (v12 == *p_fadeA)
    {
        LONG v7 = this->_fadeB;
        if (fadeB == v7)
        {
            InterlockedExchange(p_fadeA, v7);
        }
    }
}

DWORD CUserPane::s_FadeThreadProc(LPVOID lpParameter)
{
    CUserPane* pThis = static_cast<CUserPane*>(lpParameter);
    pThis->_DoFade();
    pThis->Release();
    return 0;
}

void CUserPane::_FadePictureWindow()
{
    if (_fadeA != _fadeB || _fadeA != _fadeC)
    {
        _fadeB = _fadeC;
        if (GetSystemMetrics(SM_REMOTESESSION) || GetSystemMetrics(SM_REMOTECONTROL))
        {
            _UpdatePictureWindow(0xFF, 0);
            InterlockedExchange(&_fadeA, _fadeB);
        }
        else
        {
            AddRef();
            if (!SHCreateThread(s_FadeThreadProc, this, 0, nullptr))
            {
                Release();
            }
        }
    }
}

LRESULT CUserPane::WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_SETCURSOR:
        {
            if (_fadeA == -1)
            {
                SetCursor(LoadCursor(nullptr, IDC_HAND));
                return 1;
            }
            break;
        }
        case WM_TIMER:
        {
            if (wParam == 1)
            {
                KillTimer(hwnd, 1);
                _FadePictureWindow();
            }
            break;
        }
        case WM_LBUTTONUP:
        {
            IOpenControlPanel* pocp = nullptr;
            if (SUCCEEDED(CoCreateInstance(CLSID_OpenControlPanel, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pocp))))
            {
                pocp->Open(L"Microsoft.UserAccounts", nullptr, nullptr);
            }
            // Skipped telemetry StartMenu_UserTile_Clicked
            if (pocp)
            {
                pocp->Release();
            }
            return 0;
        }
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CUserPane::OnSize()
{
    RECT rc;
    GetClientRect(_hwnd, &rc);

    if (_hwndStatic)
    {
        int iPicOffset = rc.bottom - field_28;

        RECT rcFrame;
        rcFrame.left = (rc.right - _iFramedPicWidth - rc.left) / 2;
        rcFrame.top = iPicOffset - _iFramedPicHeight - rc.top;
        rcFrame.right = rcFrame.left + _iFramedPicWidth;
        rcFrame.bottom = iPicOffset - rc.top;
        MapWindowRect(_hwnd, NULL, &rcFrame);
        MoveWindow(_hwndStatic, rcFrame.left, rcFrame.top, RECTWIDTH(rcFrame), RECTHEIGHT(rcFrame), FALSE);
    }
    return 0;
}

HRESULT CUserPane::_CreateUserPicture()
{
    _hwndStatic = SHFusionCreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_LAYERED, L"Desktop User Picture", L"user picture", CW_USEDEFAULT,
        0, 0, _iFramedPicWidth, _iFramedPicHeight, _hwnd, nullptr, g_hinstCabinet, this);
    if (!_hwndStatic)
    {
        return E_FAIL;
    }

    SetWindowPos(
        _hwndStatic, GetWindow(GetAncestor(_hwnd, GA_ROOT), GW_HWNDPREV), 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE | SWP_NOSENDCHANGING);

    IStream* pstm = nullptr;
    HRESULT hr = SHCreateStreamOnModuleResourceW(g_hinstCabinet, MAKEINTRESOURCE(7013), L"PNGFILE", &pstm);
    if (SUCCEEDED(hr))
    {
        _pgdipImage = Gdiplus::Bitmap::FromStream(pstm);
        if (!_pgdipImage)
        {
            hr = E_OUTOFMEMORY;
        }
    }
    if (pstm)
    {
        pstm->Release();
    }
    return hr;
}

void CUserPane::_UpdateUserImage(Gdiplus::Image *pgdiImageUserPicture)
{
    ASSERT(pgdiImageUserPicture != NULL); // 672

    HDC hdc = GetDC(_hwndStatic);
    HDC hMemDC = CreateCompatibleDC(hdc);
    if (hMemDC)
    {
        RECT rc;
        GetClientRect(_hwndStatic, &rc);
        if (_hbmUserPicture)
            DeleteObject(_hbmUserPicture);

        BITMAPINFO bmi = {0};
        bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
        bmi.bmiHeader.biWidth = rc.right;
        bmi.bmiHeader.biHeight = rc.bottom;
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;

        LPVOID pvBits;
        _hbmUserPicture = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &pvBits, 0, 0);
        if (_hbmUserPicture)
        {
            HBITMAP hbmUserPicture = (HBITMAP)SelectObject(hMemDC, _hbmUserPicture);

            Gdiplus::Graphics graphics(hMemDC);
            graphics.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
            graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBilinear);
            graphics.DrawImage(pgdiImageUserPicture, 0, 0, _iUnframedPicWidth, _iUnframedPicHeight);

            SelectObject(hMemDC, hbmUserPicture);
        }
        DeleteDC(hMemDC);
    }
    ReleaseDC(_hwndStatic, hdc);
}

void SHLogicalToPhysicalDPI(int* a1, int* a2);
void RemapSizeForHighDPI(SIZE* psiz);

HRESULT CUserPane::_UpdateUserInfo(int a2)
{
    HRESULT hr = S_OK;

    BOOL bShowPicture = TRUE;
    if (_hTheme)
        GetThemeBool(_hTheme, SPP_USERPANE, 0, TMT_USERPICTURE, &bShowPicture);

    if (bShowPicture)
    {
        WCHAR szUserPicturePath[260];
        szUserPicturePath[0] = '0';
        SHGetUserPicturePath(nullptr, SHGUPP_FLAG_CREATE, szUserPicturePath, ARRAYSIZE(szUserPicturePath));
        if (szUserPicturePath[0])
        {
            Gdiplus::Image* pgdiImageUserPicture = Gdiplus::Image::FromFile(szUserPicturePath);
            if (pgdiImageUserPicture)
            {
                _iUnframedPicHeight = USERPICHEIGHT;
                _iUnframedPicWidth = USERPICWIDTH;

                if (pgdiImageUserPicture->GetWidth() > pgdiImageUserPicture->GetHeight())
                {
                    UINT iPicWidth = pgdiImageUserPicture->GetWidth();
                    UINT iPicHeight = pgdiImageUserPicture->GetHeight();
                    _iUnframedPicHeight = MulDiv(_iUnframedPicWidth, iPicHeight, iPicWidth);
                }
                else if (pgdiImageUserPicture->GetHeight() > pgdiImageUserPicture->GetWidth())
                {
                    UINT iPicHeight = pgdiImageUserPicture->GetHeight();
                    UINT iPicWidth = pgdiImageUserPicture->GetWidth();
                    _iUnframedPicWidth = MulDiv(_iUnframedPicHeight, iPicWidth, iPicHeight);
                }

                field_1C = 8;
                field_20 = 8;
                field_28 = 8;
                field_24 = 8;

                SIZE sizUserPic = { 64, 64 };
                RemapSizeForHighDPI(&sizUserPic);

                SHLogicalToPhysicalDPI(&_iUnframedPicWidth, &_iUnframedPicHeight);
                SHLogicalToPhysicalDPI(&field_1C, &field_24);
                SHLogicalToPhysicalDPI(&field_20, &field_28);

                _iFramedPicHeight = sizUserPic.cy;
                _iFramedPicWidth = sizUserPic.cx;

                if (a2)
                {
                    hr = _CreateUserPicture();
                }
                _UpdateUserImage(pgdiImageUserPicture);
                delete pgdiImageUserPicture;
            }
        }

        if (!_uidChangeRegister)
        {
            SHChangeNotifyEntry fsne;
            fsne.fRecursive = FALSE;
            fsne.pidl = nullptr;
            _uidChangeRegister = SHChangeNotifyRegister(
                _hwnd, SHCNRF_NewDelivery | SHCNRF_ShellLevel, SHCNE_EXTENDED_EVENT, UPM_CHANGENOTIFY, 1, &fsne);
        }

        ULONG cch = ARRAYSIZE(_szUserName);
        SHGetUserDisplayName(_szUserName, &cch);
        SetWindowText(_hwndStatic, _szUserName);

        if (!a2)
        {
            IUnknown_QueryServiceExec(_punkSite, IID_IFolderView, &SID_SM_DV2ControlHost, 329, 0, nullptr, nullptr);
        }
    }

    OnSize();
    NMHDR nm;
    nm.hwndFrom = _hwnd;
    nm.idFrom = 0;
    nm.code = SMN_NEEDREPAINT;
    SendMessage(GetParent(_hwnd), WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
    return hr;
}

BOOL WINAPI UserPicture_RegisterClass()
{
    WNDCLASSEX wc;
    ZeroMemory(&wc, sizeof(wc));

    wc.cbSize = sizeof(wc);
    wc.style = CS_GLOBALCLASS;
    wc.lpfnWndProc = CUserPane::s_WndProcPicture;
    wc.hInstance = g_hinstCabinet;
    wc.hbrBackground = nullptr;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = L"Desktop User Picture";
    return RegisterClassExW(&wc);
}

BOOL WINAPI UserPane_RegisterClass()
{
    WNDCLASSEX wc;
    ZeroMemory(&wc, sizeof(wc));
    
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_GLOBALCLASS;
    wc.lpfnWndProc   = CUserPane::s_WndProcPane;
    wc.hInstance     = GetModuleHandle(NULL);
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(NULL);
    wc.lpszClassName = WC_USERPANE;

    return RegisterClassEx(&wc);
}