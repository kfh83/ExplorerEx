#pragma once

#include <Windows.h>
#include <Uxtheme.h>
#include <ShObjIdl_core.h>

struct DECLSPEC_NOVTABLE IStartButtonSite
{
     virtual void STDMETHODCALLTYPE EnableTooltips() = 0;
     virtual void STDMETHODCALLTYPE PurgeRebuildRequests() = 0;
     virtual bool STDMETHODCALLTYPE ShouldUseSmallIcons() = 0;
     virtual void STDMETHODCALLTYPE HandleFullScreenApp(HWND) = 0;
     virtual void STDMETHODCALLTYPE StartButtonClicked() = 0;
     virtual void STDMETHODCALLTYPE OnStartMenuDismissed() = 0;
     virtual int STDMETHODCALLTYPE GetStartButtonMinHeight() = 0;
     virtual void STDMETHODCALLTYPE GetStartMenuSuckPlace(POINT*) = 0;
     virtual void STDMETHODCALLTYPE SetUnhideTimer(LONG, LONG) = 0;
     virtual void STDMETHODCALLTYPE OnStartButtonClosing() = 0;
};

MIDL_INTERFACE("8B62940C-7ED5-4DE6-9BDC-4CA4346AAE3B")
IStartButton : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetFocusToStartButton() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnContextMenu(HWND, LONG) = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateStartButtonBalloon() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetStartPaneActive(BOOL bActive) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartMenuDismissed() = 0;
    virtual HRESULT STDMETHODCALLTYPE UnlockStartPane() = 0;
    virtual HRESULT STDMETHODCALLTYPE LockStartPane() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPopupPosition(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetWindow(HWND*) = 0;
};

class CStartButton : public IStartButton, public IServiceProvider
{
public:
    CStartButton();

    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

    //~ Begin IStartButton Interface
    STDMETHODIMP SetFocusToStartButton() override;
    STDMETHODIMP OnContextMenu(HWND, LONG) override;
    STDMETHODIMP CreateStartButtonBalloon() override;
    STDMETHODIMP SetStartPaneActive(BOOL bActive) override;
    STDMETHODIMP OnStartMenuDismissed() override;
    STDMETHODIMP UnlockStartPane() override;
    STDMETHODIMP LockStartPane() override;
    STDMETHODIMP GetPopupPosition(DWORD*) override;
    STDMETHODIMP GetWindow(HWND*) override;
    //~ End IStartButton Interface

    //~ Begin IServiceProvider Interface
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;
    //~ End IServiceProvider Interface

    WCHAR *pszCurrentThemeName;
    INT _nSomeSize;
    int _nUnkBool2;
    int _nUnkBool3;
    HWND _hwndStartBtn;
    HWND _hwndStartBalloon;
    LONG _lWidth;
    LONG _lHeight;
    HTHEME _hTheme;
    int _nShellBrdBitmap1001;
    HFONT _hStartFont;
    char padding[4];
    int _nUnkBool1;
    int _nStartBtnNotPressed;
    int _nIsOnContextMenu;
    UINT _uLockCode;
    int _nBackgroundBitmapInitialized;
    int _nSettingsChangeType;
    int _nStartPaneActiveState;
    DWORD _dwTickCount;
    HIMAGELIST _hIml;
    IStartButtonSite *_pStartButtonSite;
    IMenuBand *_pMenuBand;
    IMenuPopup *_pMenuPopup;
    IDeskBand *_pDv2CreateArg2;
    IMenuPopup *_pDv2CreateArg1;
    IDeskBand *_pUnk1;
    char padding5[4];
    WCHAR *_pszWindowName;
};