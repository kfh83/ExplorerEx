#pragma once

// from explorer\desktop2
STDAPI DesktopV2_Create(
    IMenuPopup** ppmp, IMenuBand** ppmb, void** ppvStartPane, IUnknown** ppunk, HWND hwnd);
STDAPI DesktopV2_Build(void* pvStartPane);

// from tray
EXTERN_C BOOL WINAPI Tray_StartPanelEnabled();


MIDL_INTERFACE("8B62940C-7ED5-4DE6-9BDC-4CA4346AAE3B")
IStartButton : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetFocusToStartButton() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnContextMenu(HWND, LPARAM) = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateStartButtonBalloon(UINT, UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetStartPaneActive(BOOL) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartMenuDismissed() = 0;
    virtual HRESULT STDMETHODCALLTYPE UnlockStartPane() = 0;
    virtual HRESULT STDMETHODCALLTYPE LockStartPane() = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPopupPosition(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetWindow(HWND*) = 0;
};

interface DECLSPEC_NOVTABLE IStartButtonSite
{
    virtual void STDMETHODCALLTYPE EnableTooltips(BOOL) = 0;
    virtual void STDMETHODCALLTYPE PurgeRebuildRequests() = 0;
    virtual BOOL STDMETHODCALLTYPE ShouldUseSmallIcons() = 0;
    virtual void STDMETHODCALLTYPE HandleFullScreenApp(HWND) = 0;
    virtual void STDMETHODCALLTYPE StartButtonClicked() = 0;
    virtual void STDMETHODCALLTYPE OnStartMenuDismissed() = 0;
    virtual int STDMETHODCALLTYPE GetStartButtonMinHeight() = 0;
    virtual UINT STDMETHODCALLTYPE GetStartMenuStuckPlace() = 0;
    virtual void STDMETHODCALLTYPE SetUnhideTimer(LONG, LONG) = 0;
    virtual void STDMETHODCALLTYPE OnStartButtonClosing() = 0;
};

class CStartButton
    : public IStartButton
    , public IServiceProvider
{
public:
    CStartButton(IStartButtonSite* pStartButtonSite);

    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

    //~ Begin IStartButton Interface
    STDMETHODIMP SetFocusToStartButton() override;
    STDMETHODIMP OnContextMenu(HWND hwnd, LPARAM lParam) override;
    STDMETHODIMP CreateStartButtonBalloon(UINT a2, UINT uID) override;
    STDMETHODIMP SetStartPaneActive(BOOL bActive) override;
    STDMETHODIMP OnStartMenuDismissed() override;
    STDMETHODIMP UnlockStartPane() override;
    STDMETHODIMP LockStartPane() override;
    STDMETHODIMP GetPopupPosition(DWORD* pdwPos) override;
    STDMETHODIMP GetWindow(HWND* phwndStart) override;
    //~ End IStartButton Interface

    //~ Begin IServiceProvider Interface
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;
    //~ End IServiceProvider Interface

    // TODO: revise

    void BuildStartMenu();
    void CloseStartMenu();
    HWND CreateStartButton(HWND hwndParent);
    void DestroyStartMenu();
    void DisplayStartMenu();
    void DrawStartButton(int iStateId, bool bRepaint /*allegedly*/);
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
    int TrackMenu(HMENU hMenu);
    HRESULT TranslateMenuMessage(MSG* pmsg, LRESULT* plRet);
    void UpdateStartButton(bool a2);
    void _DestroyStartButtonBalloon();
    void _DontShowTheStartButtonBalloonAnyMore();

    const WCHAR* _pszThemeName;
    int field_C;
    int field_10;
    BOOL field_14;
    HWND _hwndStart;
    HWND _hwndStartBalloon;
    SIZE _sizeStart;
    HTHEME _hTheme;
    HBITMAP _hbmpStartBkg;
    HFONT _hStartFont;
    UINT _uDown;
    BOOL _fAllowUp;
    int field_40;
    BOOL _fForegroundLocked;
    BOOL _fBackgroundBitmapInitialized;
    bool field_4C;
    UINT _uStartButtonState;
    DWORD _tmOpen;
    HIMAGELIST _himlStartFlag;
    IStartButtonSite* _pStartButtonSite;
    IMenuBand* _pmbStartMenu;
    IMenuPopup* _pmpStartMenu;
    IMenuBand* _pmbStartPane;
    IMenuPopup* _pmpStartPane;
    IUnknown* _punkSite;
    char padding5[4];
    WCHAR _szWindowName[50];

private:

    LRESULT OnMouseClick(HWND hWndTo, LPARAM lParam);
    void _CalcExcludeRect(RECTL* lprcDst);
    BOOL _CalcStartButtonPos(POINT* a2, HRGN* a3);
    HFONT _CreateStartFont();
    void _ExploreCommonStartMenu(BOOL bExplore);

    const WCHAR* _GetCurrentThemeName();

    void _HandleDestroy();
    void _OnSettingChanged(UINT a2);
    bool _OnThemeChanged(bool bForceUpdate);
    BOOL _ShouldDelayClip(const RECT* a2, const RECT* lprcSrc2);
    LRESULT _StartButtonSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    static LRESULT s_StartButtonSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    static LRESULT s_StartMenuSubclassProc(
        HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
};