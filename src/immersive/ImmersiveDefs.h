//+-------------------------------------------------------------------------
//
//	ExplorerEx - Windows XP Explorer
//	Copyright (C) Microsoft
// 
//	File:       ImmersiveDefs.h
// 
//	Description:	Definitions useful for initialising the Immersive shell.
// 
//	History:    May-19-25   kawapure  Created
//
//+-------------------------------------------------------------------------
#pragma once

#include <windows.h>
#include <initguid.h>

// {C2F03A33-21F5-47FA-B4BB-156362A2F239}
DEFINE_GUID(CLSID_ImmersiveShell, 0xC2F03A33, 0x21F5, 0x47FA, 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39);

// {4624BD39-5FC3-44A8-A809-163A836E9031}
DEFINE_GUID(SID_ImmersiveShellHookService, 0x4624BD39, 0x5FC3, 0x44A8, 0xA8, 0x09, 0x16, 0x3A, 0x83, 0x6E, 0x90, 0x31);

// {914D9B3A-5E53-4E14-BBBA-46062ACB35A4}
DEFINE_GUID(IID_IImmersiveShellHookService, 0x914D9B3A, 0x5E53, 0x4E14, 0xBB, 0xBA, 0x46, 0x06, 0x2A, 0xCB, 0x35, 0xA4);

// {C71C41F1-DDAD-42DC-A8FC-F5BFC61DF957}
DEFINE_GUID(CLSID_ImmersiveShellBuilder, 0xc71c41f1, 0xddad, 0x42dc, 0xa8, 0xfc, 0xf5, 0xbf, 0xc6, 0x1d, 0xf9, 0x57);

// {1C56B3E4-E6EA-4CED-8A74-73B72C6BD435}
DEFINE_GUID(IID_ImmersiveShellBuilder, 0x1c56b3e4, 0xe6ea, 0x4ced, 0x8a, 0x74, 0x73, 0xb7, 0x2c, 0x6b, 0xd4, 0x35);

// {139275E0-D644-4214-B45E-D9278C4A8501}
DEFINE_GUID(IID_ImmersiveBehavior, 0x139275e0, 0xd644, 0x4214, 0xb4, 0x5e, 0xd9, 0x27, 0x8c, 0x4a, 0x85, 0x01);

// {BCBB9860-C012-4AD7-A938-6E337AE6AB05}
DEFINE_GUID(CLSID_NowPlayingSessionManager, 0xbcbb9860, 0xc012, 0x4ad7, 0xa9, 0x38, 0x6e, 0x33, 0x7a, 0xe6, 0xab, 0xa5);

MIDL_INTERFACE("139275E0-D644-4214-B45E-D9278C4A8501")
IImmersiveBehavior : public IUnknown
{
public:
	STDMETHOD(OnImmersiveThreadStart)(void) PURE;
	STDMETHOD(OnImmersiveThreadStop)(void) PURE;
	STDMETHOD(GetMaximumComponentCount)(unsigned int *count) PURE;
	STDMETHOD(CreateComponent)(unsigned int number, IUnknown **component) PURE;
	STDMETHOD(ShouldCreateComponent)(unsigned int number, int *allowed) PURE;
};

/**
 * Interface of the controller object that the OS uses to manage the
 * Immersive shell.
 */
MIDL_INTERFACE("00000000-0000-0000-0000-000000000000")
IImmersiveShellController : public IUnknown
{
public:
	STDMETHOD(Start)(void) PURE;
	STDMETHOD(Stop)(void) PURE;
	STDMETHOD(SetCreationBehavior)(IImmersiveBehavior *) PURE;
};

/**
 * Builder interface given by the OS to create an Immersive shell controller.
 */
MIDL_INTERFACE("1C56B3E4-E6EA-4CED-8A74-73B72C6BD435")
IImmersiveShellBuilder : public IUnknown
{
public:
	STDMETHOD(CreateImmersiveShellController)(IImmersiveShellController **ppControllerOut) PURE;
};

MIDL_INTERFACE("914D9B3A-5E53-4E14-BBBA-46062ACB35A4")
IImmersiveShellHookService : IUnknown
{
    STDMETHOD(Register)(void** a1,
        IImmersiveShellHookService * thiss,
        const unsigned int* prgMessages,
        unsigned int cMessages,
        IUnknown * pNotification, //IImmersiveShellHookNotification
        unsigned int* pdwCookie) PURE;//todo:args
    STDMETHOD(Unregister)(UINT cookie) PURE;
    STDMETHOD(PostShellHookMessage)(WPARAM wParam, LPARAM lParam) PURE;
    STDMETHOD(SetTargetWindowForSerialization)(HWND hwnd) PURE;
    STDMETHOD(PostShellHookMessageWithSerialization)(bool a1,
        int a2,
        IImmersiveShellHookService* thiss,
        unsigned int msg,
        int msgParam) PURE; //todo:args
    STDMETHOD(UpdateWindowApplicationId)(HWND hwnd, LPCWSTR pszAppID) PURE;
    STDMETHOD(HandleWindowReplacement)(HWND hwndOld, HWND hwndNew) PURE;
    STDMETHOD_(BOOL, IsExecutionOnSerializedThread)() PURE;
};

interface IImmersiveWindowMessageService : IUnknown
{
    STDMETHOD(Register)(UINT msg, void* pNotification, UINT * pdwCookie);
    STDMETHOD(Unregister)(UINT dwCookie);
    STDMETHOD(SendMessageW)(UINT nsg, WPARAM wParam, LPARAM lParam);
    STDMETHOD(PostMessageW)(UINT nsg, WPARAM wParam, LPARAM lParam);
    STDMETHOD(RequestHotkeys)(); //todo: args
    STDMETHOD(UnrequestHotkeys)(UINT dwCookie);
    STDMETHOD(RequestWTSSessionNotification)(void* pNotification, unsigned int* pdwCookie);
    STDMETHOD(UnrequestWTSSessionNotification)(UINT dwCookie);
    STDMETHOD(RequestPowerSettingNotification)(const GUID* pPowerSettingGuid, void* pNotification, UINT* pdwCookie);
    STDMETHOD(UnrequestPowerSettingNotification)(UINT pdwCookie);
    STDMETHOD(RequestPointerDeviceNotification)(void* pNotification, int notificationType, UINT* pdwCookie);
    STDMETHOD(UnrequestPointerDeviceNotification)(UINT dwCookie);
    STDMETHOD(RegisterDwmIconicThumbnailWindow)();
};

enum IMMERSIVE_MONITOR_FILTER_FLAGS
{
    IMMERSIVE_MONITOR_FILTER_FLAGS_NONE = 0x0,
    IMMERSIVE_MONITOR_FILTER_FLAGS_DISABLE_TRAY = 0x1,
};

MIDL_INTERFACE("c6636ec2-eba1-4e6d-a995-8fa14b8b2891")
IImmersiveApplicationWindow : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetBandId(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetNativeWindow(HWND*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetApplicationId(WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetProcessId(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetThreadId(DWORD*) = 0;
};

typedef enum __MIDL___MIDL_itf_shpriv_0000_0034_0001
{
    NIE_NULL = 0x0,
    NIE_MouseMove = 0x1,
    NIE_MouseDown = 0x2,
    NIE_MouseUp = 0x3,
    NIE_MouseClick = 0x4,
    NIE_MouseDoubleClick = 0x5,
    NIE_ContextMenu = 0x6,
    NIE_KeyDown = 0x7,
    NIE_KeyUp = 0x8,
    NIE_Tooltip = 0x9,
    NIE_Reorder = 0xA,
} NotifyIconEvent;

enum IMMERSIVE_APPLICATION_GET_WINDOWS_FILTER
{
    IAGWF_ANY = 0,
    IAGWF_STRONGLY_NAMED = 1,
    IAGWF_PREFER_PENDING_PRESENTED = 2,
    IAGWF_ONLY_PENDING_PRESENTED = 3,
    IAGWF_PRESENTATION = 4,
    IAGWF_FRAME = 5
};

enum IMMAPPPROPERTYSTOREFLAGS
{
    IAGPS_DEFAULT = 0x0
};

DEFINE_ENUM_FLAG_OPERATORS(IMMAPPPROPERTYSTOREFLAGS);

typedef struct tagIMMAPPTIMESTAMPS
{
    FILETIME ftCreation;
    FILETIME ftClosed;
    FILETIME ftActivation;
    FILETIME ftInactive;
    FILETIME ftVisible;
    FILETIME ftHidden;
} IMMAPPTIMESTAMPS;

enum IMMERSIVE_APPLICATION_QUERY_SERVICE_OPTION
{
    IAQSO_DEFAULT = 0,
    IAQSO_FIRST_PACKAGED = 1,
    IAQSO_ANY_PACKAGED = 2
};

enum SPLASHSCREEN_ORIENTATION_PREFERENCE
{
    SSOP_NONE = 0,
    SSOP_LANDSCAPE = 1,
    SSOP_PORTRAIT = 2,
    SSOP_LANDSCAPE_FLIPPED = 4,
    SSOP_PORTRAIT_FLIPPED = 8
};

enum NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_OPTION
{
    NIAWO_ALL = 0,
    NIAWO_SKIP_SYSTEM_WINDOWS = 1,
    NIAWO_CURRENT_WINDOW_ONLY = 2,
    NIAWO_CURRENT_WINDOW_ONLY_IFF_APP = 3
};

enum NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_DELIVERY_TYPE
{
    NIAWDT_POST = 0,
    NIAWDT_SENDNOTIFY = 1
};

enum IMMAPP_SETTHUMBNAIL_PREVIEW_STATE
{
    IMMSPS_VISIBLE = 0,
    IMMSPS_HIDDEN = 1
};

enum USER_INTERACTION_MODE
{
    UIM_MOUSE = 0,
    UIM_TOUCH = 1
};

enum VIEW_PRESENTATION_MODE
{
    VPM_DESKTOP = 0,
    VPM_HOLOGRAPHIC = 1
};

enum APPLICATION_VIEW_MODE
{
    AVM_DEFAULT = 0,
    AVM_COMPACT_OVERLAY = 1,
    AVM_SPANNING = 2
};

enum APPLICATION_VIEW_MODE_FLAGS
{
    AVMF_DEFAULT = 0x1,
    AVMF_COMPACT_OVERLAY = 0x2,
    AVMF_SPANNING = 0x4
};

DEFINE_ENUM_FLAG_OPERATORS(APPLICATION_VIEW_MODE_FLAGS);

enum WindowTransparencyMode
{
    WTM_TransparentWhenActive = 0,
    WTM_AlwaysOpaque = 1,
    WTM_AlwaysTransparent = 2
};

struct APPLICATION_VIEW_DATA
{
    /*APPLICATION_VIEW_STATE*/ int viewState;
    /*APPLICATION_VIEW_ORIENTATION*/ int viewOrientation;
    /*ADJACENT_DISPLAY_EDGES*/ int displayEdges;
    BOOL fIsOnLockScreen;
    BOOL fIsFullScreenMode;
    USER_INTERACTION_MODE userInteractionMode;
    VIEW_PRESENTATION_MODE presentationMode;
    APPLICATION_VIEW_MODE viewMode;
    APPLICATION_VIEW_MODE_FLAGS allowedViewModes;
    WindowTransparencyMode windowTransparencyMode;
    BOOL canOpenInNewTab;
};

struct IMMAPP_APPLICATION_VIEW_DATA
{
    APPLICATION_VIEW_DATA current;
    APPLICATION_VIEW_DATA deferred;
};

typedef enum __MIDL___MIDL_itf_shpriv_core_0000_0325_0002
{
    MCF_FORCE = 0,
    MCF_IF_NOT_VISIBLE = 1
} MONITOR_CHANGE_FLAGS;

typedef enum __MIDL___MIDL_itf_shpriv_core_0000_0325_0001
{
    GVS_NORMAL = 0,
    GVS_USE_SPLASHSCREEN_VISUAL = 1,
    GVS_USE_SPLASHSCREEN_VISUAL_ONCE = 2
} GHOST_VISUAL_STYLE;

typedef enum __MIDL___MIDL_itf_shpriv_core_0000_0325_0003
{
    IABF_NONE = 0x0,
    IABF_INVALID_AUTOGLOM_DESTINATION = 0x1,
    IABF_AVOID_VIEW_FOR_SWITCH = 0x2,
    IABF_FORCE_TERMINATE_ON_CLOSE = 0x80000000
} IMM_APP_BEHAVIOR_FLAGS;

DEFINE_ENUM_FLAG_OPERATORS(IMM_APP_BEHAVIOR_FLAGS);

typedef enum __MIDL___MIDL_itf_shpriv_core_0000_0325_0004
{
    GSF_LOCKSCREENACTIVATION = 0x1,
    GSF_ACTIVATION = 0x2
} GHOST_STATUS_FLAG;

DEFINE_ENUM_FLAG_OPERATORS(GHOST_STATUS_FLAG);

typedef enum __MIDL___MIDL_itf_shpriv_core_0000_0325_0005
{
    IAQ_WIN8_WINDOWING_BEHAVIOR = 0,
    IAQ_REQUIRES_1366_PORTRAIT_MIN_HEIGHT = 1,
    IAQ_SHOW_ACTIONS_MENU = 2,
    IAQ_USE_WIN8X_COMPATIBILITY_SCALING = 3,
    IAQ_FULLSCREEN_8X_LEGACY_APP = 4,
    IAQ_WIN81_WINDOWING_BEHAVIOR = 5,
    IAQ_DONT_CACHE_TITLE_BAR_SETTINGS = 6,
    IAQ_USE_PREFERRED_STANDALONE_SIZE = 7
} IMMERSIVE_APPLICATION_QUIRKS;

interface IAsyncCallback;

MIDL_INTERFACE("880b26f8-9197-43d0-8045-8702d0d72000")
IImmersiveMonitor : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetIdentity(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE ConnectObject(IUnknown*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetHandle(HMONITOR*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsConnected(BOOL*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsPrimary(BOOL*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsImmersiveDisplayDevice(BOOL*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDisplayRect(RECT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetOrientation(DWORD*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetWorkArea(RECT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsEqual(IImmersiveMonitor*, BOOL*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsImmersiveCapable(BOOL*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetEffectiveDpi(UINT*, UINT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetFilterFlags(IMMERSIVE_MONITOR_FILTER_FLAGS*) = 0;
};

MIDL_INTERFACE("8b14e88b-5663-4caf-b196-c31479262831")
IImmersiveApplication : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetWindows(IMMERSIVE_APPLICATION_GET_WINDOWS_FILTER, REFGUID, void**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetApplicationId(WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetUniqueId(WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OpenPropertyStore(IMMAPPPROPERTYSTOREFLAGS, REFGUID, void**) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsRunning(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsVisible(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsForeground(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetTimestamps(tagIMMAPPTIMESTAMPS*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsEqualByAppId(const WCHAR*, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsEqualByHwnd(HWND*, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsEqualByApp(IImmersiveApplication*, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsViewForSameApp(IImmersiveApplication*, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPackageId(int, WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE BelongsToPackage(const WCHAR*, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE QueryService(IMMERSIVE_APPLICATION_QUERY_SERVICE_OPTION, DWORD, REFGUID, REFGUID, void**) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsServiceAvailable(IMMERSIVE_APPLICATION_QUERY_SERVICE_OPTION, REFGUID, int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsApplicationWindowStronglyNamed(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE ContainsStronglyNamedWindow(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsInteractive(int*) = 0;
    virtual SPLASHSCREEN_ORIENTATION_PREFERENCE STDMETHODCALLTYPE GetManifestedOrientationPreference() = 0;
    virtual HRESULT STDMETHODCALLTYPE NotifyApplicationWindows(UINT, WPARAM, LPARAM, NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_OPTION, NOTIFY_IMMERSIVE_APPLICATION_WINDOWS_DELIVERY_TYPE) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDestinationInformation(IImmersiveApplicationWindow**, tagRECT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetRect(tagRECT*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetThumbnailPreviewState(IMMAPP_SETTHUMBNAIL_PREVIEW_STATE) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewData(IMMAPP_APPLICATION_VIEW_DATA*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetMonitor(IImmersiveMonitor*, MONITOR_CHANGE_FLAGS) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetMonitor(IImmersiveMonitor**) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetGhostVisualStyle(GHOST_VISUAL_STYLE) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetTitle(WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS*) = 0;
    virtual HRESULT STDMETHODCALLTYPE AddBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS) = 0;
    virtual HRESULT STDMETHODCALLTYPE RemoveBehaviorFlags(IMM_APP_BEHAVIOR_FLAGS) = 0;
    virtual HRESULT STDMETHODCALLTYPE IncrementGhostAnimationWaitCount(UINT) = 0;
    virtual HRESULT STDMETHODCALLTYPE AddGhostStatusFlag(GHOST_STATUS_FLAG) = 0;
    virtual HRESULT STDMETHODCALLTYPE RemoveGhostStatusFlag(GHOST_STATUS_FLAG) = 0;
    virtual HRESULT STDMETHODCALLTYPE InvokeCharms() = 0;
    virtual HRESULT STDMETHODCALLTYPE OnMinSizePreferencesUpdated(HWND*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsSplashScreenPresented(int*) = 0;
    virtual int STDMETHODCALLTYPE IsQuirkEnabled(IMMERSIVE_APPLICATION_QUIRKS) = 0;
    virtual HRESULT STDMETHODCALLTYPE TryInvokeBack(IAsyncCallback*) = 0;
    virtual HRESULT STDMETHODCALLTYPE RequestCloseAsync(REFGUID, void**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCanHandleCloseRequest(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetPositionerMonitor(IImmersiveMonitor*) = 0;
    virtual HRESULT STDMETHODCALLTYPE IsTitleBarDrawnByApp(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDisplayName(WCHAR**) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetIsOccluded(int*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetIsOccluded(int) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetWindowingEnvironmentConfig(IUnknown*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetPersistingStateName(WCHAR**) = 0;
};