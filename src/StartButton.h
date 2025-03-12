#pragma once

#include <Windows.h>
#include <Uxtheme.h>
#include <ShObjIdl_core.h>

struct DECLSPEC_NOVTABLE IStartButtonSite
{
     STDMETHOD_(VOID,EnableTooltips(BOOL bEnable)) PURE;
     STDMETHOD_(VOID, PurgeRebuildRequests()) PURE;
     STDMETHOD_(BOOL, ShouldUseSmallIcons()) PURE;
     STDMETHOD_(VOID, HandleFullScreenApp(HWND)) PURE;
     STDMETHOD_(VOID, StartButtonClicked()) PURE;
     STDMETHOD_(VOID, OnStartMenuDismissed()) PURE;
     STDMETHOD_(INT, GetStartButtonMinHeight()) PURE;
     STDMETHOD_(UINT, GetStartMenuStuckPlace()) PURE;
     STDMETHOD_(VOID, SetUnhideTimer(LONG, LONG)) PURE;
     STDMETHOD_(VOID, OnStartButtonClosing()) PURE;
};

MIDL_INTERFACE("8B62940C-7ED5-4DE6-9BDC-4CA4346AAE3B")
IStartButton : IUnknown
{
    STDMETHOD(SetFocusToStartButton()) PURE;
    STDMETHOD(OnContextMenu(HWND hWnd, LPARAM lParam)) PURE;
    STDMETHOD(CreateStartButtonBalloon(UINT a2, UINT uID)) PURE;
    STDMETHOD(SetStartPaneActive(BOOL bActive)) PURE;
    STDMETHOD(OnStartMenuDismissed()) PURE;
    STDMETHOD(UnlockStartPane()) PURE;
    STDMETHOD(LockStartPane()) PURE;
    STDMETHOD(GetPopupPosition(DWORD* out)) PURE;
    STDMETHOD(GetWindow(HWND* hWndOut)) PURE;
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
    STDMETHODIMP OnContextMenu(HWND, LPARAM) override;
    STDMETHODIMP CreateStartButtonBalloon(UINT a2, UINT uID) override;
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
    HWND CreateStartButton(HWND);
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

    LPCWSTR pszCurrentThemeName;
    INT _nSomeSize;
    int _nUnkBool2;
    int _nUnkBool3;
    HWND _hwndStartBtn;
    HWND _hwndStartBalloon;
    LONG _lWidth;
    LONG _lHeight;
    HTHEME _hTheme;
    HBITMAP _hbmpStartBkg;
    HFONT _hStartFont;
    char padding[4];
    int _nUnkBool1;
    int _nStartBtnNotPressed;
    int _nIsOnContextMenu;
    BOOL _uLockCode;
    BOOL _fBackgroundBitmapInitialized;
    bool _nSettingsChangeType;
    int _nStartPaneActiveState;
    DWORD _dwTickCount;
    HIMAGELIST _hIml;
    IStartButtonSite *_pStartButtonSite;
    IMenuBand *_pOldStartMenuBand;
    IMenuPopup *_pOldStartMenu;
    IMenuBand *_pNewStartMenuBand;
    IMenuPopup *_pNewStartMenu;
    IDeskBand *_pUnk1;
    char padding5[4];
    WCHAR *_pszWindowName;
    UINT g_nReplaceMe1;
    WCHAR WindowName;

private:

    // TODO: revise

    LRESULT OnMouseClick(HWND hWndTo, LPARAM lParam);
    void _CalcExcludeRect(RECT* lprcDst);           // TODO
    BOOL _CalcStartButtonPos(POINT* a2, HRGN* a3);  // TODO
    HFONT _CreateStartFont();
    void _ExploreCommonStartMenu(BOOL bExplore);

    LPCWSTR _GetCurrentThemeName();

    void _HandleDestroy();
    void _OnSettingChanged(UINT a2);
    bool _OnThemeChanged(bool bForceUpdate);
    BOOL _ShouldDelayClip(const RECT* a2, const RECT* lprcSrc2);
    LRESULT _StartButtonSubclassProc(HWND hWndTo, UINT uMsg, WPARAM wParam, LPARAM lParam);

    static LRESULT s_StartButtonSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, CStartButton *dwRefData);
    static LRESULT s_StartMenuSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
};
