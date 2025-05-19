#ifndef TASKBAND_H_
#define TASKBAND_H_

#ifdef __cplusplus

#include "pch.h"
#include "cwndproc.h"
#include "dpa.h"
#include "commoncontrols.h"

#ifdef EXEX_DLL
#include "immersive/shpriv.h"
DEFINE_GUID(_GUID_23650f94_13b8_4f39_b2c3_817e6564a756,
    0x23650f94,
    0x13b8, 0x4f39, 0xb2, 0xc3,
    0x81, 0x7e, 0x65, 0x64, 0xa7, 0x56
);

//MIDL_INTERFACE("8b14e88b-5663-4caf-b196-c31479262831")
//IImmersiveApplication : IUnknown
//{
//    virtual HRESULT STDMETHODCALLTYPE GetWindows(IMMERSIVE_APPLICATION_GET_WINDOWS_FILTER, REFGUID, void**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetApplicationId(WCHAR**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetUniqueId(WCHAR**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE OpenPropertyStore(IMMAPPPROPERTYSTOREFLAGS, REFGUID, void**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsRunning(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsVisible(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsForeground(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetTimestamps(tagIMMAPPTIMESTAMPS*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsEqualByAppId(const WCHAR*, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsEqualByHwnd(HWND*, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsEqualByApp(IImmersiveApplication*, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsViewForSameApp(IImmersiveApplication*, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetPackageId(int, WCHAR**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE BelongsToPackage(const WCHAR*, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE QueryService(IMMERSIVE_APPLICATION_QUERY_SERVICE_OPTION, DWORD, REFGUID, REFGUID, void**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsServiceAvailable(IMMERSIVE_APPLICATION_QUERY_SERVICE_OPTION, REFGUID, int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsApplicationWindowStronglyNamed(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE ContainsStronglyNamedWindow(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsInteractive(int*) = 0;
//    virtual SPLASHSCREEN_ORIENTATION_PREFERENCE STDMETHODCALLTYPE GetManifestedOrientationPreference() = 0;
//    virtual HRESULT STDMETHODCALLTYPE NotifyApplicationWindows(UINT, WPARAM, LPARAM, NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_OPTION, NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_DELIVERY_TYPE) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetDestinationInformation(IImmersiveApplicationWindow**, tagRECT*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetRect(tagRECT*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetThumbnailPreviewState(IMMAPP_SETTHUMBNAIL_PREVIEW_STATE) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetViewData(IMMAPP_APPLICATION_VIEW_DATA*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetMonitor(IImmersiveMonitor*, MONITOR_CHANGE_FLAGS) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetMonitor(IImmersiveMonitor**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetGhostVisualStyle(GHOST_VISUAL_STYLE) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetTitle(WCHAR**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE AddBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS) = 0;
//    virtual HRESULT STDMETHODCALLTYPE RemoveBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IncrementGhostAnimationWaitCount(UINT) = 0;
//    virtual HRESULT STDMETHODCALLTYPE AddGhostStatusFlag(GHOST_STATUS_FLAG) = 0;
//    virtual HRESULT STDMETHODCALLTYPE RemoveGhostStatusFlag(GHOST_STATUS_FLAG) = 0;
//    virtual HRESULT STDMETHODCALLTYPE InvokeCharms() = 0;
//    virtual HRESULT STDMETHODCALLTYPE OnMinSizePreferencesUpdated(HWND*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsSplashScreenPresented(int*) = 0;
//    virtual int STDMETHODCALLTYPE IsQuirkEnabled(IMMERSIVE_APPLICATION_QUIRKS) = 0;
//    virtual HRESULT STDMETHODCALLTYPE TryInvokeBack(IAsyncCallback*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE RequestCloseAsync(REFGUID, void**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetCanHandleCloseRequest(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetPositionerMonitor(IImmersiveMonitor*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE IsTitleBarDrawnByApp(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetDisplayName(WCHAR**) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetIsOccluded(int*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetIsOccluded(int) = 0;
//    virtual HRESULT STDMETHODCALLTYPE SetWindowingEnvironmentConfig(IUnknown*) = 0;
//    virtual HRESULT STDMETHODCALLTYPE GetPersistingStateName(WCHAR**) = 0;
//};


DEFINE_GUID(_GUID_c71c41f1_ddad_42dc_a8fc_f5bfc61df957, 0xC71C41F1, 0xDDAD, 0x42DC, 0xA8, 0xFC, 0xF5, 0xBF, 0xC6, 0x1D, 0xF9, 0x57); // c71c41f1-ddad-42dc-a8fc-f5bfc61df957

interface IImmersiveShellCreationBehavior;

MIDL_INTERFACE("ffffffff-ffff-ffff-ffff-ffffffffffff")
IImmersiveShellController : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Start() = 0;
    virtual HRESULT STDMETHODCALLTYPE Stop() = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCreationBehavior(IImmersiveShellCreationBehavior*) = 0;
};

MIDL_INTERFACE("1c56b3e4-e6ea-4ced-8a74-73b72c6bd435")
IImmersiveShellBuilder : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE CreateImmersiveShellController(IImmersiveShellController**) = 0;
};

DEFINE_GUID(SID_ImmersiveShellHookService, 0x4624BD39, 0x5FC3, 0x44A8, 0xA8, 0x09, 0x16, 0x3A, 0x83, 0x6E, 0x90, 0x31); // 4624bd39-5fc3-44a8-a809-163a836e9031

interface IImmersiveShellHookNotification;

MIDL_INTERFACE("914d9b3a-5e53-4e14-bbba-46062acb35a4")
IImmersiveShellHookService : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Register(const UINT_PTR* const, UINT, IImmersiveShellHookNotification*, DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE Unregister(DWORD) = 0;
    virtual HRESULT STDMETHODCALLTYPE PostShellHookMessage(WPARAM, LPARAM) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetTargetWindowForSerialization(HWND) = 0;
    virtual HRESULT STDMETHODCALLTYPE PostShellHookMessageWithSerialization(WPARAM, LPARAM) = 0;
    virtual HRESULT STDMETHODCALLTYPE UpdateWindowApplicationId(HWND, const WCHAR*) = 0;
    virtual HRESULT STDMETHODCALLTYPE HandleWindowReplacement(HWND, HWND) = 0;
    virtual int STDMETHODCALLTYPE IsExecutionOnSerializedThread() = 0;
};

DEFINE_GUID(SID_AppCrusher, 0x3CF1532D, 0x66FB, 0x4F8C, 0x95, 0x92, 0xDC, 0x45, 0xC6, 0x3E, 0x65, 0x22); // 3cf1532d-66fb-4f8c-9592-dc45c63e6522

enum APPCRUSHER_INPUT
{
    ACI_TOUCH = 0,
    ACI_MOUSE = 1,
    ACI_KEYBOARD = 2,
    ACI_CONTEXTMENU = 3,
    ACI_TASKBAR = 4,
    ACI_TITLEBAR = 5,
    ACI_CAPTIONCONTROLS = 6,
    ACI_MULTITASKINGVIEW = 7,
    ACI_API = 8,
    ACI_HANGDETECTION = 9,
    ACI_INVALID = 10,
};

enum APPCRUSHER_SOURCE
{
    ACS_BACKSTACK = 0,
    ACS_VISIBLE = 1,
    ACS_LIST = 2,
    ACS_APPTIP = 3,
    ACS_PLACEMODE = 4,
    ACS_TASKBAR = 5,
    ACS_TITLEBAR = 6,
    ACS_CAPTIONCONTROLS = 7,
    ACS_MULTITASKINGVIEW = 8,
    ACS_API = 9,
    ACS_HANGDETECTION = 10,
    ACS_INVALID = 11,
};

enum APPCRUSHER_CLOSE_OPTIONS
{
    ACCO_NONE = 0x0,
    ACCO_NO_CLOSE_ANIMATION = 0x1,
    ACCO_NO_FOREGROUND_RIGHT_TRANSFER = 0x2,
    ACCO_FORCE = 0x4,
    ACCO_NO_CONFIRMAPPCLOSE_NOTIFY = 0x8,
    ACCO_NO_FALLBACK_FOREGROUND = 0x10,
    ACCO_SYSTEM_ACTION = 0x20,
};

DEFINE_ENUM_FLAG_OPERATORS(APPCRUSHER_CLOSE_OPTIONS);

enum APPCRUSHER_MINIMIZE_FLAGS
{
    ACMF_NONE = 0x0,
    ACMF_USE_DESKTOP_ANIMATIONS = 0x1,
};

DEFINE_ENUM_FLAG_OPERATORS(APPCRUSHER_MINIMIZE_FLAGS);

MIDL_INTERFACE("103231ae-04cb-4e5e-b63d-e3ce47cd0f0a")
IImmersiveAppCrusher : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE CloseAllApps() = 0;
    virtual HRESULT STDMETHODCALLTYPE CloseApp(IImmersiveApplication*, APPCRUSHER_INPUT, APPCRUSHER_SOURCE, APPCRUSHER_CLOSE_OPTIONS) = 0;
    virtual HRESULT STDMETHODCALLTYPE CloseAppWithTelemetryCookie(IImmersiveApplication*, APPCRUSHER_INPUT, APPCRUSHER_SOURCE, APPCRUSHER_CLOSE_OPTIONS, DWORD) = 0;
    virtual HRESULT STDMETHODCALLTYPE CloseApps(IObjectArray*, APPCRUSHER_INPUT, APPCRUSHER_SOURCE, APPCRUSHER_CLOSE_OPTIONS) = 0;
    virtual HRESULT STDMETHODCALLTYPE MinimizeApp(IImmersiveApplication*, APPCRUSHER_MINIMIZE_FLAGS) = 0;
    virtual HRESULT STDMETHODCALLTYPE Register(/*IImmersiveAppCrusherNotification*/void*, DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE Unregister(DWORD) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetTaskCompletionFormatString(size_t, /*HSTRING*/void*) = 0;
};
#endif

class CTray;
class CGroupItemContextMenu;

class TASKITEM
{
public:
    TASKITEM() { ZeroMemory(this, sizeof(TASKITEM)); };
    TASKITEM(TASKITEM* pti);
    ~TASKITEM();
    HWND hwnd;  // NULL if this item is a group of application entries
    DWORD dwFlags;
    class TaskShortcut *ptsh;
    DWORD dwTimeLastClicked;
    DWORD dwTimeFirstOpened;
    WCHAR* pszExeName;
    int iIconPref;
    BOOL fMarkedFullscreen;
    BOOL fHungApp;
};

typedef TASKITEM *PTASKITEM;

typedef struct
{
    PTASKITEM   pti; 
    UINT        fState;
    int         iIndex; // used to cache toolbar index
}
ANIMATIONITEMINFO, *PANIMATIONITEMINFO;

class CTaskBandSMC;

class CTaskBand : public IDeskBand
                , public IObjectWithSite
                , public IDropTarget
                , public IInputObject
                , public IPersistStream
                , public IWinEventHandler
                , public IOleCommandTarget
                , public CImpWndProc
{
public:
    // *** IUnknown methods ***
    STDMETHODIMP QueryInterface(REFIID riid, LPVOID * ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // *** IOleWindow methods ***
    STDMETHODIMP GetWindow(HWND * lphwnd);
    STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode)  { return E_NOTIMPL; }

    // *** IDockingWindow methods ***
    STDMETHODIMP ShowDW(BOOL fShow)         { return S_OK; }
    STDMETHODIMP CloseDW(DWORD dwReserved)  { return S_OK; }
    STDMETHODIMP ResizeBorderDW(LPCRECT prcBorder, IUnknown* punkToolbarSite, BOOL fReserved) { return E_NOTIMPL; }

    // *** IObjectWithSite methods ***
    STDMETHODIMP SetSite(IUnknown* punkSite);
    STDMETHODIMP GetSite(REFIID riid, void** ppvSite) { return E_NOTIMPL; };

    // *** IDeskBand methods ***
    STDMETHODIMP GetBandInfo(DWORD dwBandID, DWORD fViewMode, DESKBANDINFO* pdbi);

    // *** IDropTarget methods ***
    STDMETHODIMP DragEnter(IDataObject *pdtobj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
    STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
    STDMETHODIMP DragLeave(void);
    STDMETHODIMP Drop(IDataObject *pdtobj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);

    // *** IInputObject methods ***
    STDMETHODIMP TranslateAcceleratorIO(LPMSG lpMsg) { return E_NOTIMPL; }
    STDMETHODIMP HasFocusIO();
    STDMETHODIMP UIActivateIO(BOOL fActivate, LPMSG lpMsg);

    // *** IWinEventHandler methods ***
    STDMETHODIMP OnWinEvent(HWND hwnd, UINT dwMsg, WPARAM wParam, LPARAM lParam, LRESULT* plres);
    STDMETHODIMP IsWindowOwner(HWND hwnd);

    // *** IPersistStream methods ***
    STDMETHODIMP GetClassID(LPCLSID pClassID);
    STDMETHODIMP IsDirty(void)                  { return S_FALSE; }
    STDMETHODIMP Load(IStream *ps);
    STDMETHODIMP Save(LPSTREAM, BOOL)           { return S_OK; }
    STDMETHODIMP GetSizeMax(ULARGE_INTEGER*)    { return E_NOTIMPL; }

    // *** IOleCommandTarget methods ***
    STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pcmdtext);
    STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut);

protected:
    static void IconAsyncProc(HWND hwnd, UINT uMsg, ULONG_PTR dwData, LRESULT lResult);

    typedef struct
    {
        HWND hwnd;
        LPTSTR pszExeName;
        int iImage;
    }
    ICONCBPARAM, *PICONCBPARAM;

    typedef int (*PICONCALLBACK)(CTaskBand* ptb, PICONCBPARAM pip, LPARAM lParam, int iPref);

    static int GetIconCB(CTaskBand* ptb, PICONCBPARAM pip, LPARAM lParam, int iPref);
    static int GetSHILIconCB(CTaskBand* ptb, PICONCBPARAM pip, LPARAM lParam, int);
    static int GetDefaultIconCB(CTaskBand* ptb, PICONCBPARAM pip, LPARAM, int);
    static int GetClassIconCB(CTaskBand* ptb, PICONCBPARAM pip, LPARAM lParam, int);
    
    void _MoveGroup(HWND hwnd, WCHAR* szNewExeName);
    void _SetWindowIcon(HWND hwnd, HICON hicon, int iPref);
    static BOOL _ShouldMinimize(HWND hwnd);
    BOOL _CanMinimizeAll();
    BOOL _MinimizeAll(HWND hwndTray, BOOL fPostRaiseDesktop);
    int _HitTest(POINTL ptl);
    void _FreePopupMenu();

    void _RealityCheck();
    int  _FindIndexByHwnd(HWND hwnd);
    void _CheckNeedScrollbars(int cyRow, int cItems, int iCols, int iRows,
                                     int iItemWidth, LPRECT lprcView);
    void _NukeScrollbar(int fnBar);
    void _SetItemWidth(int iItem, int iWidth);
    int  _GetItemWidth(int iItem);
    int  _GetLastVisibleItem();
    int  _GetVisibleItemCount();
    int  _GetGroupWidth(int iIndexGroup);
    int  _GetIdealWidth(int *iRemainder);
    void _GetNumberOfRowsCols(int* piRows, int* piCols, BOOL fCurrentSize);
    int  _GetTextSpace();
    void _GetToolbarMetrics(LPTBMETRICS ptbm);
    void _CheckSize(void);
    void _SizeItems(int iButtonWidth, int iRemainder = 0);
    BOOL _AddWindow(HWND hwnd);
    BOOL _CheckButton(int iIndex, BOOL fCheck);
    BOOL _IsButtonChecked(int iIndex);
    int  _GetCurSel();
    void _SetCurSel(int iIndex, BOOL fIgnoreCtrlKey);
    int  _SelectWindow(HWND hwnd);
    void _SwitchToWindow(HWND hwnd);

    int  _GetSelectedItems(CDSA<PTASKITEM>* pdsa);
    int  _GetGroupItems(int iIndexGroup, CDSA<PTASKITEM>* pdsa);
    void _OnGroupCommand(int iRet, CDSA<PTASKITEM>* pdsa);
    void _SysMenuForItem(int i, int x, int y);
    static void CALLBACK FakeSystemMenuCB(HWND hwnd, UINT uMsg, ULONG_PTR dwData, LRESULT lres);
    HWND _CreateFakeWindow(HWND hwndOwner);
    void _HandleSysMenuTimeout();
    void _HandleSysMenu(HWND hwnd);
    void _FakeSystemMenu(HWND hwndTask, DWORD dwPos);
    BOOL _ContextMenu(DWORD dwPos);

    void _HandleCommand(WORD wCmd, WORD wID, HWND hwnd);
    void _DrawNumber(HDC hdc, int iValue, BOOL fCalcRect, LPRECT prc);
    LRESULT _HandleCustomDraw(LPNMTBCUSTOMDRAW ptbcd, PTASKITEM pti = NULL);
    void _RemoveImage(int iImage);
    void _OnButtonPressed(int iIndex, PTASKITEM pti, BOOL fForceRestore);
    LRESULT _HandleNotify(LPNMHDR lpnm);
    void _SwitchToItem(int iItem, HWND hwnd, BOOL fIgnoreCtrlKey);
    LRESULT _HandleCreate();
    LRESULT _HandleDestroy();
    LRESULT _HandleScroll(BOOL fHoriz, UINT code, int nPos);
    void _ScrollIntoView(int iItem);
    LRESULT _HandleSize(WPARAM fwSizeType);
    LRESULT _HandleActivate(HWND hwndActive);
    void _UpdateItemUsage(PTASKITEM pti);
    void _HandleOtherWindowDestroyed(HWND hwndDestroyed);
    void _HandleGetMinRect(HWND hwndShell, LPPOINTS lprc);
    void _HandleChangeNotify(WPARAM wParam, LPARAM lParam);
    LRESULT _HandleHardError(HARDERRORDATA *phed, DWORD cbData);
    BOOL _IsItemActive(HWND hwndActive);
    void _CreateTBImageLists();
    int _AddIconToNormalImageList(HICON hicon, int iImage);

    void _UpdateItemText(int iItem);
    void _UpdateItemIcon(int iItem);
    void _GetDispInfo(LPNMTBDISPINFO lptbdi);

    void _DoRedrawWhereNeeded();
    void _RedrawItem(HWND hwndShell, WPARAM code, int i = -1);
    void _SetActiveAlt(HWND hwndAlt);
    HWND _EnumForRudeWindow(HWND hwndSelected);
    HWND _FindRudeApp(HWND hwndPossible);
    LRESULT _OnAppCommand(int cmd);
    PTASKITEM _FindItemByHwnd(HWND hwnd);
    void _OnWindowActivated(HWND hwnd, BOOL fSuspectFullscreen);
    LRESULT _HandleShellHook(int iCode, LPARAM lParam);
    void _VerifyButtonHeight();
    void _InitFonts();
    void _SetItemImage(int iItem, int iImage, int iPref);
    void _UpdateAllIcons();
    LRESULT _HandleWinIniChange(WPARAM wParam, LPARAM lParam, BOOL fOnCreate);
    void _OnSetFocus();
    BOOL _RegisterWindowClass();
    void _UpdateFlashingFlag();
    void _ExecuteMenuOption(HWND hwnd, int iCmd);
    TASKITEM* _GetItem(int i, TBBUTTONINFO* ptbb = NULL, BOOL fByIndex = TRUE);

    static BOOL WINAPI BuildEnumProc(HWND hwnd, LPARAM lParam);
    static BOOL WINAPI IsRudeEnumProc(HWND hwnd, LPARAM lParam);
    static DWORD WINAPI MinimizeAllThreadProc(LPVOID lpv);
    void _OpenTheme();

    virtual LRESULT v_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

    int _GetGroupSize(int iIndexGroup);
    int _GetGroupIndex(int iIndexApp);
    DWORD _GetGroupAge(int iIndexGroup);
    int _GetGroupIndexFromExeName(WCHAR* szExeName);

    BOOL _IsHidden(int i);
    void _GetItemTitle(int iIndex, WCHAR* pszTitle, int cbTitle, BOOL fCustom);
    void _RefreshSettings();
    void _LoadSettings();
    void _Glom(int iIndexGroup, BOOL fGlom);
    void _HideGroup(int iIndexGroup, BOOL fHide);
    BOOL _AutoGlomGroup(BOOL fGlom, int iOpenSlots);
    void _DeleteTaskItem(int index, BOOL fDeletePTI);
    void _RealityCheckGroup(PTASKITEM pti);
    HRESULT _CreatePopupMenu(POINTL* ppt, RECTL* prcl);
    void _AddItemToDropDown(int iIndex);
    void _RemoveItemFromDropDown(int iIndex);
    void _RefreshItemFromDropDown(int iIndex, int iNewIndex, BOOL fRefresh);
    void _ClosePopupMenus();
    void _HandleDropDown(int index);
    void _UpdateProgramCount();
    BOOL _AddToTaskbar(PTASKITEM pti, int indexTaskbar, BOOL fVisible, BOOL fForceGetIcon);
    BOOL _InsertItem(HWND hwndTask, PTASKITEM ptiOveride = NULL, BOOL fForceGetIcon = FALSE);
    void _DeleteItem(HWND hWnd, int index = -1);
    void _AttachTaskShortcut(PTASKITEM pti, LPCTSTR pszExeName);
    void _ReattachTaskShortcut();
    void _BuildTaskList(CDPA<TASKITEM>* pDPA);

    // *** Async-Animation 
    BOOL _fAnimate;
    CDSA<ANIMATIONITEMINFO> _dsaAII;

    // animation methods
    BOOL _AnimateItems(int iIndex, BOOL fExpand, BOOL fGlomAnimation);
    void _AsyncAnimateItems();
    void _ResizeAnimationItems();
    int  _CheckAnimationSize();
    void _SizeNonAnimatingItems();

    // animation helpers
    void  _UpdateAnimationIndices();
    void  _UpdateAnimationIndicesSlow();
    int   _FindItem(PTASKITEM pti);
    void  _RemoveItemFromAnimationList(PTASKITEM ptiRemove);
    void  _SetAnimationState(PANIMATIONITEMINFO paii, BOOL fExpand, BOOL fGlomAnimation);
    int   _GetAnimationInsertPos(int iIndex);
    void  _SetAnimationItemWidth(PANIMATIONITEMINFO paii, int cxStep);
    int   _GetAnimationDistLeft(PANIMATIONITEMINFO paii, int iNormalWidth);
    void  _FinishAnimation(PANIMATIONITEMINFO paii);
    int   _GetAnimationWidth();
    int   _GetAnimationStep();
    DWORD _GetStepTime(int iStep);
    int _GetCurButtonHeight();

    void _SetThreadPriority(int iPriority, DWORD dwWakeupTime);
    void _RestoreThreadPriority();

    BOOL _IsHorizontal() { return !(_dwViewMode & DBIF_VIEWMODE_VERTICAL); }

    BOOL _fGlom;
    int _iGroupSize;
    CToolBarCtrl _tb;
    UINT WM_ShellHook;
    int _iSysMenuCount;
    int _iIndexActiveAtLDown;
    HWND    _hwndSysMenu;
    HWND    _hwndLastRude;
    HWND    _hwndPrevFocus;
    HWND    _hwndReplacing;
    BOOL _fIgnoreTaskbarActivate;
    BOOL _fFlashing;
    BOOL _fDenyHotItemChange;
    CTray* _ptray;
    HFONT _hfontSave;
    int _iTextSpace;
    DWORD _dwPos;
    DWORD _dwViewMode;
    HFONT _hfontCapNormal;
    HFONT _hfontCapBold;
    HTHEME _hTheme;
    int _iOldPriority;
    int _iNewPriority;

    ULONG _cRef;

    // Drag & drop stuff
    int _iDropItem;
    DWORD _dwTriggerStart;
    DWORD _dwTriggerDelay;

    // Variables for the ASYNC popup menu
    IShellMenu2* _psmPopup;
    IMenuPopup* _pmpPopup;
    IMenuBand*  _pmbPopup;
    int         _iIndexPopup;
    int         _iIndexLastPopup;
    CMenu       _menuPopup;

    IImageList* _pimlSHIL;

    // Rarely-used stuff
    ULONG       _uShortcutInvokeNotify;
    UINT        _uCDHardError;

#ifdef EXEX_DLL
    // Immersive stuff:
    CComPtr<IImmersiveShellController> m_immersiveShellController;
    CComPtr<IImmersiveShellHookService> m_immersiveShellHookService;
    //CComPtr<IImmersiveMonitorManager> m_immersiveMonitorManager;
    //CComPtr<IClassicWindowManager> m_classicWindowManager;
    CComPtr<IImmersiveAppCrusher> m_immersiveAppCrusher;
#endif

        
    CTaskBand();
    ~CTaskBand();
    HRESULT Init(CTray* ptray);

    HRESULT _BandInfoChanged();

    DWORD _dwBandID;

    IUnknown *  _punkSite;

    friend HRESULT CTaskBand_CreateInstance(IUnknown* punkOuter, IUnknown** ppunk);
    friend class CTaskBandSMC;
};

#endif  // __cplusplus

#endif //TASKBAND_H_
