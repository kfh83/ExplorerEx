#pragma once

#include <Windows.h>
#include <Uxtheme.h>
#include <ShObjIdl_core.h>

MIDL_INTERFACE("F89C89D6-76E1-42D7-9D11-995976CF7B4C")
IStartButtonSite : IUnknown
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

    IStartButton* iStartBtn;
    IServiceProvider* iservProvider;
    WCHAR* currentThemeName;
    INT someSize;
    BOOL unkBool2;
    BOOL unkBool3;
    HWND m_hWndStartBtn;
    HWND m_hWndStartBalloon;
    LONG width;
    LONG height;
    HTHEME hTheme;
    int shellBrdBitmap1001;
    HFONT startFont;
    char padding[4];
    BOOL unkBool1;
    BOOL startBtnNotPressed;
    BOOL isOnContextMenu;
    UINT uLockCode;
    BOOL backgroundBitmapInitialized;
    int settingsChangeType;
    BOOL startPaneActiveState;
    DWORD tickCount;
    HIMAGELIST hIml;
    IStartButtonSite* m_btnSite;
    IMenuBand* mnuBand;
    IMenuPopup* mnuPopup;
    IDeskBand* dv2_create_arg2;
    IMenuBand* dv2_create_arg1;
    IDeskBand* unk1;
    char padding5[4];
    WCHAR* windowName;
};