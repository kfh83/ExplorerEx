#pragma once

// from explorer\desktop2
STDAPI DesktopV2_Create(
    IMenuPopup** ppmp, IMenuBand** ppmb, void** ppvStartPane, IUnknown** ppunkHost, HWND hwndOwner);
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

class IStartButtonSite
{
public:
    virtual void STDMETHODCALLTYPE EnableTooltips(BOOL fEnable) {}
    virtual void STDMETHODCALLTYPE PurgeRebuildRequests() {}
    virtual BOOL STDMETHODCALLTYPE ShouldUseSmallIcons() { return FALSE; }
    virtual void STDMETHODCALLTYPE HandleFullScreenApp(HWND hwnd) {}
    virtual void STDMETHODCALLTYPE StartButtonClicked() {}
    virtual void STDMETHODCALLTYPE OnStartMenuDismissed() = 0;
    virtual int STDMETHODCALLTYPE GetStartButtonMinHeight() { return 0; }
    virtual UINT STDMETHODCALLTYPE GetStartMenuStuckPlace() = 0;
    virtual void STDMETHODCALLTYPE SetUnhideTimer(LONG x, LONG y) = 0;
    virtual void STDMETHODCALLTYPE OnStartButtonClosing() = 0;
};

class CStartButton
    : public IStartButton
    , public IServiceProvider
{
public:
    CStartButton(IStartButtonSite* psbs);

    //~ Begin IUnknown Interface
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;
    //~ End IUnknown Interface

    //~ Begin IStartButton Interface
    STDMETHODIMP SetFocusToStartButton() override;
    STDMETHODIMP OnContextMenu(HWND hwnd, LPARAM lParam) override;
    STDMETHODIMP CreateStartButtonBalloon(UINT idsTitle, UINT idsMessage) override;
    STDMETHODIMP SetStartPaneActive(BOOL bActive) override;
    STDMETHODIMP OnStartMenuDismissed() override;
    STDMETHODIMP UnlockStartPane() override;
    STDMETHODIMP LockStartPane() override;
    STDMETHODIMP GetPopupPosition(DWORD* pdwPos) override;
    STDMETHODIMP GetWindow(HWND* phwndStart) override;
    //~ End IStartButton Interface

    //~ Begin IServiceProvider Interface
    STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObj) override;
    //~ End IServiceProvider Interface

    // TODO: revise

    void BuildStartMenu();
    void CloseStartMenu();
    HWND CreateStartButton(HWND hwndParent);
    void DestroyStartMenu();
    void DisplayStartMenu();
    void DrawStartButton(int iStateId, bool bRepaint);
    void ExecRefresh();
    void ForceButtonUp();
    void GetRect(RECT* prc);
    void GetSizeAndFont(HTHEME hTheme);
    BOOL InitBackgroundBitmap();
    void InitTheme();
    BOOL IsButtonPushed();
    HRESULT IsMenuMessage(MSG* pmsg);
    BOOL IsPopupMenuVisible();
    void RecalcSize();
    void RepositionBalloon();
    void StartButtonReset();
    int TrackMenu(HMENU hmenu);
    HRESULT TranslateMenuMessage(MSG* pmsg, LRESULT* plres);
    void UpdateStartButton(bool a2);

    void _DestroyStartButtonBalloon();
    void _DontShowTheStartButtonBalloonAnyMore();

    enum
    {
        STB_RECALCSIZE      = WM_APP,
        STB_GETIDEALSIZE    = WM_APP + 1,
    };

    enum { IDT_STARTBUTTONBALLOON = 1 };

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
    BOOL _fInContextMenu;
    BOOL _fForegroundLocked;
    BOOL _fBackgroundBitmapInitialized;
    bool field_4C;
    UINT _uStartButtonState;
    DWORD _tmOpen;
    HIMAGELIST _himlStartFlag;
    IStartButtonSite* _psbs;
    IMenuBand* _pmbStartMenu;
    IMenuPopup* _pmpStartMenu;
    IMenuBand* _pmbStartPane;
    IMenuPopup* _pmpStartPane;
    IUnknown* _punkSMHost;
    char padding5[4];
    WCHAR _szStart[50];

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