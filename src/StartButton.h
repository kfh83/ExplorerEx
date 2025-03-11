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
     virtual UINT STDMETHODCALLTYPE GetStartMenuStuckPlace() = 0;
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
    STDMETHODIMP OnContextMenu(HWND, LONG) override;    // TODO: do
    STDMETHODIMP CreateStartButtonBalloon() override;   // TODO: do
    STDMETHODIMP SetStartPaneActive(BOOL bActive) override;
    STDMETHODIMP OnStartMenuDismissed() override;
    STDMETHODIMP UnlockStartPane() override;
    STDMETHODIMP LockStartPane() override;
    STDMETHODIMP GetPopupPosition(DWORD* out) override;
    STDMETHODIMP GetWindow(HWND* out) override;
    //~ End IStartButton Interface

    //~ Begin IServiceProvider Interface
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;
    //~ End IServiceProvider Interface

    // TODO: revise

    void BuildStartMenu();
    void CloseStartMenu();
    HWND CreateStartButton(HWND hwnd);
    void DestroyStartMenu();
    void DisplayStartMenu();
    void DrawStartButton(int iStateId, bool hdcSrc /*allegedly*/);
    void ExecRefresh();
    void ForceButtonUp();
    void GetRect(RECT* lpRect);
    void GetSizeAndFont(HTHEME hTheme);
    BOOL InitBackgroundBitmap();
    void InitTheme();
    BOOL IsButtonPushed();
    HRESULT IsMenuMessage(MSG* pmsg);
    BOOL IsPopupMenuVisible();
    void RecalcSize();
    void RepositionBalloon();
    void StartButtonReset();
    BOOL TrackMenu(HMENU hMenu);
    BOOL TranslateMenuMessage(MSG* pmsg, LRESULT* plRet);
    void UpdateStartButton(bool a2 /*allegedly*/);
    void _DestroyStartButtonBalloon();
    void _DontShowTheStartButtonBalloonAnyMore();

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
    BOOL _uLockCode;
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

private:

    // TODO: revise

    HRESULT OnMouseClick(HWND hWndTo, int a3);
    void _CalcExcludeRect(RECT* lprcDst);
    void _CalcStartButtonPos(POINT* a2, HRGN* a3);
    HFONT _CreateStartFont();
    void _ExploreCommonStartMenu();

    const WCHAR _GetCurrentThemeName();

    void _HandleDestroy();
    void _OnSettingChanged(UINT a2);
    char _OnThemeChanged(char a2);
    BOOL _ShouldDelayClip(RECT* a2, const RECT* lprcSrc2);
    LRESULT _StartButtonSubclassProc(HWND hWndTo, UINT uMsg, WPARAM wParam, LPARAM lParam);

    static LRESULT s_StartButtonSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, CStartButton *dwRefData);
    static LRESULT s_StartMenuSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

};