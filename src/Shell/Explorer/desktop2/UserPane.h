#pragma once

#include "COWSite.h"

// window class name of user pane control
#define WC_USERPANE TEXT("Desktop User Pane")

// hardcoded width and height of user picture
#define USERPICWIDTH 48
#define USERPICHEIGHT 48

#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")

class CGraphicsInit
{
public:
    CGraphicsInit()
    {
        Gdiplus::GdiplusStartupInput gsi = { nullptr };
        gsi.GdiplusVersion = 1;
        Gdiplus::GdiplusStartup(&_token, &gsi, NULL);
    }

    ~CGraphicsInit()
    {
        if (_token)
        {
            Gdiplus::GdiplusShutdown(_token);
            _token = 0;
        }
    }

private:
    ULONG_PTR _token;
};

class CUserPane
    : public CObjectWithSite
    , public IServiceProvider
	, public IOleCommandTarget
{
public:
    CUserPane();
    ~CUserPane() override;

    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // *** IObjectWithSite ***
    STDMETHODIMP SetSite(IUnknown* punkSite) override;

    // *** IServiceProvider ***
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;

    // *** IOleCommandTarget ***
    STDMETHODIMP QueryStatus(const GUID* pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT* pCmdText) override;
    STDMETHODIMP Exec(
        const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvarargIn, VARIANT* pvarargOut) override;

    static LRESULT CALLBACK s_WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT CALLBACK WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    static LRESULT CALLBACK s_WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT CALLBACK WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    LRESULT _OnNcCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    LRESULT OnSize();
    void _UpdatePictureWindow(BYTE a2, BYTE a3);
    void _PaintPictureWindow(HDC hdc, BYTE a2, BYTE a3);
    HWND _GetPictureWindowPrevHnwd();
    void _UpdateDC(HDC hdcDest, int iIndex, BYTE bAlpha);
    void _HidePictureWindow();
    void _FadePictureWindow();
    void _DoFade();

    static DWORD s_FadeThreadProc(LPVOID lpParameter);

private:
    HRESULT _UpdateUserInfo(BOOL fInitial);
    HRESULT _CreateUserPicture();
    void _UpdateUserImage(Gdiplus::Image* pgdiImageUserPicture);
	
    HWND _hwnd;
    HWND _hwndStatic;
    HTHEME _hTheme;
    MARGINS _mrgnPictureFrame;
    int _cxPicInset;
    int _cxPicMargin; // Seems to  be unused anyway, the frame is always centered horizontally.
    int _cyPicInset;
    int _cyPicMargin;

    int _iFramedPicHeight;
    int _iFramedPicWidth;
    int _iUnframedPicHeight;
    int _iUnframedPicWidth;

    UINT _uidChangeRegister;

    HBITMAP _hbmUserPicture;
    LONG _lRef;
    LONG _lFadeA;
    LONG _lFadeB;
    LONG _lFadeC;
    HIMAGELIST _himl;
    CGraphicsInit _graphics;
    DWORD _dwFadeIn;
    DWORD _dwFadeOut;
    DWORD _dwFadeDelay;
    Gdiplus::Image* _pgdipImage;
    TCHAR _szUserName[UNLEN + 1];

    enum { UPM_CHANGENOTIFY = WM_USER };

    friend BOOL UserPane_RegisterClass();
};