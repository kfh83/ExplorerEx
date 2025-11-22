#include "pch.h"

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
        Gdiplus::GdiplusStartupInput gsi = {0};
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
    ~CUserPane();

	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// *** IObjectWithSite ***
	STDMETHODIMP SetSite(IUnknown *punkSite);

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
	STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut);

    static LRESULT CALLBACK s_WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT CALLBACK WndProcPane(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	static LRESULT CALLBACK s_WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT CALLBACK WndProcPicture(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    LRESULT _OnNcCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	
    LRESULT OnSize();
    void _UpdatePictureWindow(BYTE a2, BYTE a3);
	void _PaintPictureWindow(HDC hdc, BYTE a2, BYTE a3);
    HWND _GetPictureWindowPrevHnwd();
    void _UpdateDC(HDC hdc, int iIndex, BYTE a3);
	void _HidePictureWindow();
    void _FadePictureWindow();
    void _DoFade();

    static DWORD s_FadeThreadProc(LPVOID lpParameter);

private:
    HRESULT _UpdateUserInfo(int a2);
	HRESULT _CreateUserPicture();
    void _UpdateUserImage(Gdiplus::Image *pgdiImageUserPicture);
	
    HWND _hwnd;
    HWND _hwndStatic;
    HTHEME _hTheme;
    int field_1C;
    int field_20;
    int field_24;
    int field_28;
    int _iFramedPicHeight;
    int _iFramedPicWidth;
    int _iUnframedPicHeight;
    int _iUnframedPicWidth;

    UINT _uidChangeRegister;

    HBITMAP _hbmUserPicture;
	LONG _lRef;
    LONG _fadeA;
    LONG _fadeB;
    LONG _fadeC;
    HIMAGELIST _himl;
    CGraphicsInit _graphics;
    DWORD _dwFadeIn;
    DWORD _dwFadeOut;
    DWORD _dwFadeDelay;
    Gdiplus::Image *_pgdipImage;
    TCHAR _szUserName[UNLEN + 1];

    enum { UPM_CHANGENOTIFY = WM_USER };

    friend BOOL UserPane_RegisterClass();
};
