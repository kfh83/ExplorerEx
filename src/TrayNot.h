#ifndef _TRAYNOT_H
#define _TRAYNOT_H

#include "cwndproc.h"
#include "pch.h"
#include "traycmn.h"
#include "trayitem.h"
#include "trayreg.h"

#define TNM_GETCLOCK (WM_USER + 1)
#define TNM_HIDECLOCK (WM_USER + 2)
#define TNM_TRAYHIDE (WM_USER + 3)
#define TNM_TRAYPOSCHANGED (WM_USER + 4)
#define TNM_ASYNCINFOTIP   (WM_USER + 5)
#define TNM_ASYNCINFOTIPPOS (WM_USER + 6)
#define TNM_RUDEAPP         (WM_USER + 7)
#define TNM_SAVESTATE               (WM_USER + 8)
#define TNM_NOTIFY                  (WM_USER + 9)
#define TNM_STARTUPAPPSLAUNCHED     (WM_USER + 10)
#define TNM_ENABLEUSERTRACKINGINFOTIPS      (WM_USER + 11)

#define TNM_BANGICONMESSAGE         (WM_USER + 50)
#define TNM_ICONDEMOTETIMER         (WM_USER + 61)
#define TNM_INFOTIPTIMER            (WM_USER + 62)
#define TNM_UPDATEVERTICAL          (WM_USER + 63)
#define TNM_WORKSTATIONLOCKED       (WM_USER + 64)

#define TNM_SHOWTRAYBALLOON         (WM_USER + 90)

#define UID_CHEVRONBUTTON           (-1)


#define XXX_RENIM   0   // cache/reload ShellNotifyIcon's (see #if)


// Tray Notify Icon area implementation notes / details:

// - The icons are held in a toolbar with PTNPRIVICON on each button's lParam


#define INFO_TIMER       48
#define KEYBOARD_VERSION 3

#define NISP_SHAREDICONSOURCE 0x10000000 // says this is the source of a shared icon

#define INFO_INFO       0x00000001
#define INFO_WARNING    0x00000002
#define INFO_ERROR      0x00000003
#define INFO_ICON       0x00000003
#define ICON_HEIGHT      16
#define ICON_WIDTH       16
#define MAX_TIP_WIDTH    300

#define MIN_INFO_TIME   10000  // 10 secs is minimum time a balloon can be up
#define MAX_INFO_TIME   60000  // 1 min is the max time it can be up

//  For Win64 compat, the icon and hwnd are handed around as DWORDs
//  (so they won't change size as they travel between 32-bit and
//  64-bit processes).

#define GetHIcon(pnid)  ((HICON)ULongToPtr(pnid->dwIcon))
#define GetHWnd(pnid)   ((HWND)ULongToPtr(pnid->dwWnd))

#define ROWSCOLS(_nTot, _nROrC) ((_nTot+_nROrC-1)/_nROrC)

#define PADDING 1

typedef struct _TNINFOITEM
{
    INT_PTR nIcon;
    TCHAR szTitle[64];
    TCHAR szInfo[256];
    UINT uTimeout;
    DWORD dwFlags;
    BOOL bMinShown; // was this balloon shown for a min time?
} TNINFOITEM;

typedef struct _TNPRIVICON
{
    HWND hWnd;
    UINT uID;
    UINT uCallbackMessage;
    DWORD dwState;
    UINT uVersion;
    HICON hIcon;
} TNPRIVICON, *PTNPRIVICON;

#if XXX_RENIM
typedef enum {
    COP_ADD, COP_DEL,
} CACHEOP;
void CacheNID(CACHEOP op, INT_PTR nIcon, PNOTIFYICONDATA32 pNidMod);
#endif

//  Everybody has a copy of this function, so we will too!
STDAPI_(void) ExplorerPlaySound(LPCTSTR pszSound);

// defined in tray.cpp
extern BOOL IsPosInHwnd(LPARAM lParam, HWND hwnd);
// defined in taskband.cpp
extern BOOL ToolBar_IsVisible(HWND hwndToolBar, int iIndex);

#define _IsOverClock(lParam) IsPosInHwnd(lParam, _hwndClock)

// The anchor point is a new thing introduced in Vista. It is used for NOTIFYICON version 4.
// This is the internal form of the anchor point WPARAM documented on MSDN for version 4.
// It is the exact same type, but with a couple sentinel values used for internal Explorer
// logic.
#define TRAYITEM_ANCHORPOINT_INPUTTYPE_MOUSE    ((DWORD)(-1))
#define TRAYITEM_ANCHORPOINT_INPUTTYPE_KEYBOARD ((DWORD)(-2))


class CTrayNotify : public CImpWndProc
{
public:
    CTrayNotify() {};
    virtual ~CTrayNotify() {};

    // *** IUnknown methods ***
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // *** ITrayNotify methods, which are called from the CTrayNotifyStub ***
    STDMETHODIMP SetPreference(NOTIFYITEM pNotifyItem);
    STDMETHODIMP RegisterCallback(INotificationCB* pNotifyCB);
    STDMETHODIMP EnableAutoTray(BOOL bTraySetting);
    
    HWND TrayNotifyCreate(HWND hwndParent, UINT uID, HINSTANCE hInst);
    LRESULT TrayNotify(HWND hwndNotify, HWND hwndFrom, PCOPYDATASTRUCT pcds, BOOL *pbRefresh);
protected:
    // Used for notifications:
    WPARAM _CalculateAnchorPointWPARAMIfNecessary(DWORD inputType, HWND const hwnd, int itemIndex);

    LRESULT _SendNotify(PTNPRIVICON ptnpi, UINT uMsg, DWORD dwAnchorPoint = 0, HWND const hwnd = NULL, int itemIndex = 0);
    void _SetImage(INT_PTR iIndex, int iImage);
    void _SetText(INT_PTR iIndex, LPTSTR pszText);
    int _GetImage(INT_PTR iIndex);
    PTNPRIVICON _GetData(INT_PTR i, BOOL byIndex);
    INT_PTR _GetCount();
    INT_PTR _GetVisibleCount();
    int _FindImageIndex(HICON hIcon, BOOL fSetAsSharedSource);
    void _RemoveImage(UINT uIMLIndex);
    INT_PTR _FindNotify(PNOTIFYICONDATA32 pnid);
    BOOL _CheckAndResizeImages();
    void _ActivateTips(BOOL bActivate);
    void _InfoTipMouseClick(int x, int y);
    void _PositionInfoTip();
    void _ShowInfoTip(INT_PTR nIcon, BOOL bShow, BOOL bAsync);
    void _SetInfoTip(INT_PTR nIcon, PNOTIFYICONDATA32 pnid, BOOL bAsync);
    BOOL _ModifyNotify(PNOTIFYICONDATA32 pnid, INT_PTR nIcon, BOOL* pbRefresh);
    BOOL _SetVersionNotify(PNOTIFYICONDATA32 pnid, INT_PTR nIcon);
    void _FreeNotify(PTNPRIVICON ptnpi, int iImage);
    BOOL_PTR _DeleteNotify(INT_PTR nIcon);
    BOOL _InsertNotify(PNOTIFYICONDATA32 pnid);
    void _SetCursorPos(INT_PTR i);
    LRESULT CALLBACK _ToolbarWndProc(
        HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    
    // Initialization/Destroy functions
    LRESULT _Create(HWND hWnd);
    LRESULT _Destroy();

    LRESULT _Paint();
    int _MatchIconsHorz(int nMatchHorz, INT_PTR nIcons, POINT *ppt);
    int _MatchIconsVert(int nMatchVert, INT_PTR nIcons, POINT *ppt);
    UINT _CalcRects(int nMaxHorz, int nMaxVert, LPRECT prClock, LPRECT prNotifies);
    LRESULT _CalcMinSize(int nMaxHorz, int nMaxVert);
    LRESULT _Size();
    LRESULT _Timer(UINT uTimerID);
    LRESULT _MouseEvent(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL fClickDown);
    LRESULT _OnCDNotify(LPNMTBCUSTOMDRAW pnm);
    LRESULT _Notify(LPNMHDR pNmhdr);
    void _SysChange(UINT uMsg, WPARAM wParam, LPARAM lParam);
    void _Command(UINT id, UINT uCmd);
    LRESULT v_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) override;
    static LRESULT CALLBACK s_ToolbarWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
    BOOL TrayNotifyIcon(PTRAYNOTIFYDATA pnid, BOOL *pbRefresh);

    static int CALLBACK DeleteDPAPtrCB(void *pItem, void *pData);

    static const WCHAR c_szTrayNotify[];
    
private:
    LONG        m_cRef;
    
    BOOL _IsScreenSaverRunning();
    
    HWND _hwndNotify;
    HWND _hwndToolbar;
    HWND _hwndClock;
    HIMAGELIST _himlIcons;
    int _nCols;
    INT_PTR _iVisCount;
    BOOL _fKey : 1;
    BOOL _fReturn : 1;
    HWND _hwndInfoTip;
    UINT_PTR _uInfoTipTimer;
    TNINFOITEM* _pinfo; //current balloon being shown
    HDPA _hdpaInfo; // array of balloons waiting in queque
};
#pragma optimize( "", off )
//
// Stub for CTrayNotify, so as to not break the COM rules of refcounting a static object
//
class CTrayNotifyStub :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CTrayNotifyStub, &CLSID_TrayNotify>,
    public ITrayNotify
{
public:
    CTrayNotifyStub() {};
    virtual ~CTrayNotifyStub() {};

    //DECLARE_NOT_AGGREGATABLE(CTrayNotifyStub)

    BEGIN_COM_MAP(CTrayNotifyStub)
        COM_INTERFACE_ENTRY(ITrayNotify)
    END_COM_MAP()

    // *** ITrayNotify method ***
    virtual STDMETHODIMP RegisterCallback(INotificationCB* pNotifyCB, ULONG*) override;
    virtual STDMETHODIMP UnregisterCallback(ULONG*) override;
    virtual STDMETHODIMP SetPreference(NOTIFYITEM pNotifyItem) override;
    virtual STDMETHODIMP EnableAutoTray(BOOL bTraySetting) override;
    virtual STDMETHODIMP DoAction(BOOL bTraySetting) override;
    virtual STDMETHODIMP SetWindowingEnvironmentConfig(IUnknown* unk) override;
};
#pragma optimize( "", on )
#endif