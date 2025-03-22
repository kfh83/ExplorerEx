#pragma once
//+-------------------------------------------------------------------------
//
//  ExplorerEx - Windows XP Explorer
//  Copyright (C) Microsoft
//
//  File:       shundoc.h
//
//  History:    Oct-11-24   aubymori  Created
//  History:    Jan-20-25   kfh83     Modify for Explorer
//  History:    Jan-20-25   olive6841 Add COM Interaces
//
//--------------------------------------------------------------------------

#include "pch.h"
#include "criticalsection.h"
#include "winbase.h"
#include "winter.h"
#include "ieguidp.h"

enum ZBID
{
    ZBID_DEFAULT = 0,
    ZBID_DESKTOP = 1,
    ZBID_UIACCESS = 2,
    ZBID_IMMERSIVE_IHM = 3,
    ZBID_IMMERSIVE_NOTIFICATION = 4,
    ZBID_IMMERSIVE_APPCHROME = 5,
    ZBID_IMMERSIVE_MOGO = 6,
    ZBID_IMMERSIVE_EDGY = 7,
    ZBID_IMMERSIVE_INACTIVEMOBODY = 8,
    ZBID_IMMERSIVE_INACTIVEDOCK = 9,
    ZBID_IMMERSIVE_ACTIVEMOBODY = 10,
    ZBID_IMMERSIVE_ACTIVEDOCK = 11,
    ZBID_IMMERSIVE_BACKGROUND = 12,
    ZBID_IMMERSIVE_SEARCH = 13,
    ZBID_GENUINE_WINDOWS = 14,
    ZBID_IMMERSIVE_RESTRICTED = 15,
    ZBID_SYSTEM_TOOLS = 16,
    ZBID_LOCK = 17,
    ZBID_ABOVELOCK_UX = 18,
};

// path.cpp (private stuff) ---------------------

#define PQD_NOSTRIPDOTS 0x00000001

//void PathQualifyDef(LPTSTR psz, LPCTSTR szDefDir, DWORD dwFlags);

BOOL PathIsRemovable(LPCTSTR pszPath);
BOOL PathIsRemote(LPCTSTR pszPath);
BOOL PathIsTemporary(LPCTSTR pszPath);
BOOL PathIsWild(LPCTSTR pszPath);
BOOL PathIsLnk(LPCTSTR pszFile);
//STDAPI_(BOOL) PathIsSlow(LPCTSTR pszFile, DWORD dwFileAttr);
//BOOL PathIsInvalid(LPCTSTR pPath);
BOOL PathIsBinaryExe(LPCTSTR szFile);
//BOOL PathMergePathName(LPTSTR pPath, LPCTSTR pName);
BOOL PathGetMountPointFromPath(LPCTSTR pcszPath, LPTSTR pszMountPoint, int cchMountPoint);
BOOL PathIsShortcutToProgram(LPCTSTR pszFile);

//
// Constants
//

// from browseui/globals.h, possibly others
#define c_szNULL        TEXT("")

// Old shlobj.h things:
//
// Path processing function
//
#define PPCF_ADDQUOTES               0x00000001        // return a quoted name if required
#define PPCF_ADDARGUMENTS            0x00000003        // appends arguments (and wraps in quotes if required)
#define PPCF_NODIRECTORIES           0x00000010        // don't match to directories
#define PPCF_FORCEQUALIFY            0x00000040        // qualify even non-relative names
#define PPCF_LONGESTPOSSIBLE         0x00000080        // always find the longest possible name

// shellprv.h (including its own copied attributed headers):
//  Copy.c
#define SPEED_SLOW  400
DWORD GetPathSpeed(LPCTSTR pszPath);
// end shellprv.h

#define RRA_DEFAULT 0x0000
#define RRA_DELETE  0x0001
#define RRA_WAIT    0x0002

#define NIS_SHOWALWAYS          0x20000000      

// shlapip.h
#define NI_SIGNATURE    0x34753423
#define PFOPEX_NONE        0x00000000
#define PFOPEX_PIF         0x00000001
#define PFOPEX_COM         0x00000002
#define PFOPEX_EXE         0x00000004
#define PFOPEX_BAT         0x00000008
#define PFOPEX_LNK         0x00000010
#define PFOPEX_CMD         0x00000020
#define PFOPEX_OPTIONAL    0x00000040   // Search only if Extension not present
#define PFOPEX_DEFAULT     (PFOPEX_CMD | PFOPEX_COM | PFOPEX_BAT | PFOPEX_PIF | PFOPEX_EXE | PFOPEX_LNK)

// shlwapip.h

//
// flags for PathIsValidChar()
//
#define PIVC_ALLOW_QUESTIONMARK     0x00000001  // treat '?' as valid
#define PIVC_ALLOW_STAR             0x00000002  // treat '*' as valid
#define PIVC_ALLOW_DOT              0x00000004  // treat '.' as valid
#define PIVC_ALLOW_SLASH            0x00000008  // treat '\\' as valid
#define PIVC_ALLOW_COLON            0x00000010  // treat ':' as valid
#define PIVC_ALLOW_SEMICOLON        0x00000020  // treat ';' as valid
#define PIVC_ALLOW_COMMA            0x00000040  // treat ',' as valid
#define PIVC_ALLOW_SPACE            0x00000080  // treat ' ' as valid
#define PIVC_ALLOW_NONALPAHABETIC   0x00000100  // treat non-alphabetic exteneded chars as valid
#define PIVC_ALLOW_QUOTE            0x00000200  // treat '"' as valid

//
// standard masks for PathIsValidChar()
//
#define PIVC_SFN_NAME               (PIVC_ALLOW_DOT | PIVC_ALLOW_NONALPAHABETIC)
#define PIVC_SFN_FULLPATH           (PIVC_SFN_NAME | PIVC_ALLOW_COLON | PIVC_ALLOW_SLASH)
#define PIVC_LFN_NAME               (PIVC_ALLOW_DOT | PIVC_ALLOW_NONALPAHABETIC | PIVC_ALLOW_SEMICOLON | PIVC_ALLOW_COMMA | PIVC_ALLOW_SPACE)
#define PIVC_LFN_FULLPATH           (PIVC_LFN_NAME | PIVC_ALLOW_COLON | PIVC_ALLOW_SLASH)
#define PIVC_SFN_FILESPEC           (PIVC_SFN_FULLPATH | PIVC_ALLOW_STAR | PIVC_ALLOW_QUESTIONMARK)
#define PIVC_LFN_FILESPEC           (PIVC_LFN_FULLPATH | PIVC_ALLOW_STAR | PIVC_ALLOW_QUESTIONMARK)

const DWORD dwExStyleRTLMirrorWnd = WS_EX_LAYOUTRTL;

//
// Structs
// 

typedef struct {
    LPITEMIDLIST pidl;                          // IDlist for this item
    int          nOrder;                        // Ordinal indicating user preference
    DWORD        lParam;                        // store custom order info.
} ORDERITEM, * PORDERITEM;

// From comctrlp.h
typedef struct tagNMTBWRAPHOTITEM
{
    NMHDR hdr;
    int iStart;
    int iDir;
    UINT nReason;       // HICF_* flags
} NMTBWRAPHOTITEM, *LPNMTBWRAPHOTITEM;

typedef struct _SHDDEERR {      // sde (Software Design Engineer, Not!)
    UINT idMsg;
    TCHAR szParam[MAX_PATH];
} SHDDEERR, * PSHDDEERR;

typedef struct tagDDECONV {
    void*          head;           // HM header
    struct tagDDECONV* snext;
    struct tagDDECONV* spartnerConv;  // siamese twin
    void*                spwnd;          // associated pwnd
    void*                spwndPartner;   // associated partner pwnd
    struct tagXSTATE* spxsOut;       // transaction info queue - out point
    struct tagXSTATE* spxsIn;        // transaction info queue - in point
    struct tagFREELIST* pfl;           // free list
    DWORD               flags;          // CXF_ flags
    struct tagDDEIMP* pddei;         // impersonation information
} DDECONV, * PDDECONV;

typedef BOOL(*DDECOMMAND)(LPTSTR lpszBuf, UINT* lpwCmd, PDDECONV pddec);
typedef struct _DDECOMMANDINFO
{
    LPCTSTR     pszCommand;
    DDECOMMAND lpfnCommand;
} DDECOMMANDINFO;

typedef struct {
    NMHDR  hdr;
    CHAR   szCmd[MAX_PATH * 2];
    DWORD  dwHotKey;
} NMVIEWFOLDERA, FAR* LPNMVIEWFOLDERA;

typedef struct {
    NMHDR  hdr;
    WCHAR  szCmd[MAX_PATH * 2];
    DWORD  dwHotKey;
} NMVIEWFOLDERW, FAR* LPNMVIEWFOLDERW;

typedef struct tagMUIINSTALLLANG {
    LANGID LangID;
    BOOL   bInstalled;
} MUIINSTALLLANG, * LPMUIINSTALLLANG;

typedef struct
{
    CLSID clsid;
    DWORD dwFlags;
} LOADINPROCDATA, * PLOADINPROCDATA;
/*
typedef struct
{
    DWORD   cbSize;     // SIZEOF
    DWORD   dwMask;     // INOUT requested/given (UEIM_*)
    int     cHit;       // profile count
    DWORD   dwAttrs;    // attributes (UEIA_*)
    FILETIME ftExecute; // Last execute filetime
} UEMINFO, * LPUEMINFO;
*/
typedef struct tagUEMINFO {
    DWORD cbSize; /*  +0x0000  */
    DWORD dwMask; /*  +0x0004  */
    DWORD R; /*  +0x0008 40 00 00 00  */
    DWORD cHit; /*  +0x000c 75 00 00 00  */
    DWORD cSwitches; /*  +0x0010 75 00 00 00  */
    DWORD dwTime; /*  +0x0014  */
    FILETIME ftExecute; /*  +0x0018 ce 1a 00 00  */
    BOOL fExcludeFromMFU; /*  +0x0020 74 00 00 00  */
} UEMINFO, * LPUEMINFO;

typedef struct _SHShortcutInvokeAsIDList {
    USHORT  cb;
    DWORD   dwItem1;                    // SHCNEE_EXTENDED_EVENT requires this
    DWORD   dwPid;                      // PID of target application
    WCHAR   szShortcutName[MAX_PATH];   // Path to shortcut
    WCHAR   szTargetName[MAX_PATH];     // Path to target application
    USHORT  cbZero;
} SHShortcutInvokeAsIDList, * LPSHShortcutInvokeAsIDList;

typedef struct _AppBarData3264
{
    DWORD cbSize;
    DWORD dwWnd;
    UINT uCallbackMessage;
    UINT uEdge;
    RECT rc;
    DWORDLONG lParam; // message specific
} APPBARDATA3264, * PAPPBARDATA3264;

typedef struct _TRAYAPPBARDATA
{
    APPBARDATA3264 abd;
    DWORD dwMessage;
    DWORD hSharedABD;
    DWORD dwProcId;
} TRAYAPPBARDATA, * PTRAYAPPBARDATA;

#pragma pack(push,0x1)
typedef struct tagNOTIFYITEM
{
    LPWSTR      pszExeName;
    LPWSTR      pszIconText;
    HICON       hIcon;
    HWND        hWnd;
    DWORD       dwUserPref;
    UINT        uID;
    GUID        guidItem;
} NOTIFYITEM, * LPNOTIFYITEM;
#pragma pack(pop)

typedef struct _NOTIFYICONDATA32A {
    DWORD cbSize;
    DWORD dwWnd;                        // NB!
    UINT uID;
    UINT uFlags;
    UINT uCallbackMessage;
    DWORD dwIcon;                       // NB!
#if (_WIN32_IE < 0x0500)
    CHAR   szTip[64];
#else
    CHAR   szTip[128];
#endif
#if (_WIN32_IE >= 0x0500)
    DWORD dwState;
    DWORD dwStateMask;
    CHAR   szInfo[256];
    union {
        UINT  uTimeout;
        UINT  uVersion;
    } DUMMYUNIONNAME;
    CHAR   szInfoTitle[64];
    DWORD dwInfoFlags;
#endif
#if (_WIN32_IE >= 0x600)
    GUID guidItem;
#endif
} NOTIFYICONDATA32A, * PNOTIFYICONDATA32A;
typedef struct _NOTIFYICONDATA32W {
    DWORD cbSize;
    DWORD dwWnd;                        // NB!
    UINT uID;
    UINT uFlags;
    UINT uCallbackMessage;
    DWORD dwIcon;                       // NB!
#if (_WIN32_IE < 0x0500)
    WCHAR  szTip[64];
#else
    WCHAR  szTip[128];
#endif
#if (_WIN32_IE >= 0x0500)
    DWORD dwState;
    DWORD dwStateMask;
    WCHAR  szInfo[256];
    union {
        UINT  uTimeout;
        UINT  uVersion;
    } DUMMYUNIONNAME;
    WCHAR  szInfoTitle[64];
    DWORD dwInfoFlags;
#endif
#if (_WIN32_IE >= 0x600)
    GUID guidItem;
#endif
#if NTDDI_VERSION >= 0x06000000
    DWORD dwBalloonIcon;
#endif
} NOTIFYICONDATA32W, * PNOTIFYICONDATA32W;

typedef NOTIFYICONDATA32W NOTIFYICONDATA32;
typedef PNOTIFYICONDATA32W PNOTIFYICONDATA32;


typedef struct _TRAYNOTIFYDATAA {
    DWORD dwSignature;
    DWORD dwMessage;
    NOTIFYICONDATA32 nid;
} TRAYNOTIFYDATAA, * PTRAYNOTIFYDATAA;
typedef struct _TRAYNOTIFYDATAW {
    DWORD dwSignature;
    DWORD dwMessage;
    NOTIFYICONDATA32 nid;
} TRAYNOTIFYDATAW, * PTRAYNOTIFYDATAW;

typedef TRAYNOTIFYDATAW TRAYNOTIFYDATA;
typedef PTRAYNOTIFYDATAW PTRAYNOTIFYDATA;

typedef struct tagSHCNF_INSTRUMENT {
    USHORT uOffset;
    USHORT uAlign;
    DWORD dwEventType;
    DWORD dwEventStructure;
    SYSTEMTIME st;
    union tagEvents {
        struct tagSTRING {
            TCHAR sz[32];
        } string;
        struct tagHOTKEY {
            WPARAM wParam;
        } hotkey;
        struct tagWNDPROC {
            HWND hwnd;
            UINT uMsg;
            WPARAM wParam;
            LPARAM lParam;
        } wndproc;
        struct tagCOMMAND {
            HWND hwnd;
            UINT idCmd;
        } command;
        struct tagDROP {
            HWND hwnd;
            UINT idCmd;
            //          TCHAR sz[32]; // convert pDataObject into something we can log
        } drop;
    } e;
    USHORT uTerm;
} SHCNF_INSTRUMENT_INFO, * LPSHCNF_INSTRUMENT_INFO;

typedef struct _tagSHELLREMINDER
{
    DWORD  cbSize;
    LPCWSTR pszName;
    LPCWSTR pszTitle;
    LPCWSTR pszText;
    LPCWSTR pszTooltip;
    LPCWSTR pszIconResource;
    LPCWSTR pszShellExecute;
    GUID* pclsid;
    DWORD  dwShowTime;
    DWORD  dwRetryInterval;
    DWORD  dwRetryCount;
    DWORD  dwTypeFlags;
} SHELLREMINDER;

//
// Macros
//


// iepriv macros

#define SMC_INITMENU            0x00000001  // The callback is called to init a menuband
#define SMC_CREATE              0x00000002
#define SMC_EXITMENU            0x00000003  // The callback is called when menu is collapsing
#define SMC_EXEC                0x00000004  // The callback is called to execute an item
#define SMC_GETINFO             0x00000005  // The callback is called to return DWORD values
#define SMC_GETSFINFO           0x00000006  // The callback is called to return DWORD values
#define SMC_GETOBJECT           0x00000007  // The callback is called to get some object
#define SMC_GETSFOBJECT         0x00000008  // The callback is called to get some object
#define SMC_SFEXEC              0x00000009  // The callback is called to execute an shell folder item
#define SMC_SFSELECTITEM        0x0000000A  // The callback is called when an item is selected
#define SMC_SELECTITEM          0x0000000B  // The callback is called when an item is selected
#define SMC_GETSFINFOTIP        0x0000000C  // The callback is called to get some object
#define SMC_GETINFOTIP          0x0000000D  // The callback is called to get some object
#define SMC_INSERTINDEX         0x0000000E  // New item insert index
#define SMC_POPUP               0x0000000F  // InitMenu/InitMenuPopup (sort of)
#define SMC_REFRESH             0x00000010  // Menus have completely refreshed. Reset your state.
#define SMC_DEMOTE              0x00000011  // Demote an item
#define SMC_PROMOTE             0x00000012  // Promote an item, wParam = SMINV_* flag
#define SMC_BEGINENUM           0x00000013  // tell callback that we are beginning to ENUM the indicated parent
#define SMC_ENDENUM             0x00000014  // tell callback that we are ending the ENUM of the indicated paren
#define SMC_MAPACCELERATOR      0x00000015  // Called when processing an accelerator.
#define SMC_DEFAULTICON         0x00000016  // Returns Default icon location in wParam, index in lParam
#define SMC_NEWITEM             0x00000017  // Notifies item is not in the order stream.
#define SMC_GETMINPROMOTED      0x00000018  // Returns the minimum number of promoted items
#define SMC_CHEVRONEXPAND       0x00000019  // Notifies of a expansion via the chevron
#define SMC_DISPLAYCHEVRONTIP   0x0000002A  // S_OK display, S_FALSE not.
#define SMC_DESTROY             0x0000002B  // Called when a pane is being destroyed.
#define SMC_SETOBJECT           0x0000002C  // Called to save the passed object
#define SMC_SETSFOBJECT         0x0000002D  // Called to save the passed object
#define SMC_SHCHANGENOTIFY      0x0000002E  // Called when a Change notify is received. lParam points to SMCSHCHANGENOTIFYSTRUCT
#define SMC_CHEVRONGETTIP       0x0000002F  // Called to get the chevron tip text. wParam = Tip title, Lparam = TipText Both MAX_PATH
#define SMC_SFDDRESTRICTED      0x00000030  // Called requesting if it's ok to drop. wParam = IDropTarget.
#define SMC_GETIMAGELISTS       0x00000031  // Called to get the small & large icon image lists, otherwise it will default to shell image list
#define SMC_CUSTOMDRAW          0x00000032  // Requires SMINIT_CUSTOMDRAW
#define SMC_BEGINDRAG           0x00000033  // Called to get preferred drop effect. wParam = &pdwEffect
#define SMC_MOUSEFILTER         0x00000034  // Called to allow host to filter mouse messages. wParam=bRemove, lParam=pmsg
#define SMC_DUMPONUPDATE        0x00000035  // S_OK if host wants old trash-everything-on-update behavior (recent docs)

#define SMC_FILTERPIDL          0x10000000  // The callback is called to see if an item is visible
#define SMC_CALLBACKMASK        0xF0000000  // Mask of comutationally intense messages

//

#define LOGONID_CURRENT     ((ULONG)-1)
#define LOGONID_NONE        ((ULONG)-2)
#define SERVERNAME_CURRENT  ((HANDLE)NULL)
 

#define UEIM_HIT        0x01
#define UEIM_FILETIME   0x02

#define SHCNFI_CABINET_WNDPROC            0
#define SHCNFI_DESKTOP_WNDPROC            1
#define SHCNFI_PROXYDESKTOP_WNDPROC       2
#define SHCNFI_TRAY_WNDPROC               3
#define SHCNFI_DRIVES_WNDPROC             4
#define SHCNFI_ONETREE_WNDPROC            5
#define SHCNFI_MAIN_WNDPROC               6
#define SHCNFI_FOLDEROPTIONS_DLGPROC      7
#define SHCNFI_VIEWOPTIONS_DLGPROC        8
#define SHCNFI_FT_DLGPROC                 9
#define SHCNFI_FTEdit_DLGPROC            10
#define SHCNFI_FTCmd_DLGPROC             11
#define SHCNFI_TASKMAN_DLGPROC           12
#define SHCNFI_TRAYVIEWOPTIONS_DLGPROC   13
#define SHCNFI_INITSTARTMENU_DLGPROC     14
#define SHCNFI_PRINTERQUEUE_DLGPROC      15

#define SHCNFI_EVENT_WNDPROC              3   // e.wndproc


#define IDS_CANTFINDDIR                 0x7666

#define WINMMDEVICECHANGEMSGSTRING L"winmm_devicechange"

#define ENABLE_BALLOONTIP_MESSAGE L"Enable Balloon Tip"

#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)

#define ATOMICRELEASE(p) IUnknown_AtomicRelease((void **)&p)

#define ATOMICRELEASET(p, type) { if(p) { type* punkT=p; p=NULL; punkT->Release();} }

template <class T>
void IUnknown_SafeReleaseAndNullPtr(T **ptr)
{
    T *pRes = *ptr;

    if (*ptr)
    {
        *ptr = NULL;
        pRes->Release();
    }
}

#define ResultFromShort(i)  ResultFromScode(MAKE_SCODE(SEVERITY_SUCCESS, 0, (USHORT)(i)))

#define ShortFromResult(r)  (short)SCODE_CODE(GetScode(r))

#define SAFECAST(_obj, _type) (((_type)(_obj)==(_obj)?0:0), (_type)(_obj))

#define BOOLIFY(expr)           (!!(expr))

#define SIZECHARS(x)    (sizeof((x))/sizeof(WCHAR))

#define IS_WM_CONTEXTMENU_KEYBOARD(lParam) ((DWORD)(lParam) == 0xFFFFFFFF)

#define NtCurrentPeb() ((PPEB)NtCurrentTeb()->ProcessEnvironmentBlock)

#define IS_VALID_STRING_PTRA(ptr, cch) \
   ((IsBadReadPtr((ptr), sizeof(char)) || IsBadStringPtrA((ptr), (UINT_PTR)(cch))) ? \

#define IntToPtr_(T, i) ((T)IntToPtr(i))

#define _IOffset(class, itf)         ((UINT)(UINT_PTR)&(((class *)0)->itf))
#define IToClass(class, itf, pitf)   ((class  *)(((LPSTR)pitf)-_IOffset(class, itf)))

#define IS_VALID_CODE_PTR(ptr, type) \
   (! IsBadCodePtr((FARPROC)(ptr)))



#define SHProcessMessagesUntilEvent(hwnd, hEvent, dwTimeout)        SHProcessMessagesUntilEventEx(hwnd, hEvent, dwTimeout, QS_ALLINPUT)
#define SHProcessSentMessagesUntilEvent(hwnd, hEvent, dwTimeout)    SHProcessMessagesUntilEventEx(hwnd, hEvent, dwTimeout, QS_SENDMESSAGE)

#define ToolBar_ButtonCount(hwnd)  \
    (BOOL)SNDMSG((hwnd), TB_BUTTONCOUNT, 0, 0)

// returns -1 on failure, button index on success
#define ToolBar_GetButtonInfo(hwnd, idBtn, lptbbi)  \
    (int)(SNDMSG((hwnd), TB_GETBUTTONINFO, (WPARAM)(idBtn), (LPARAM)(lptbbi)))

#define ARGUMENT_PRESENT(ArgumentPointer)    (\
    (CHAR *)(ArgumentPointer) != (CHAR *)(NULL) )

#define CbFromCchW(cch)             ((cch)*sizeof(WCHAR))
#define CbFromCchA(cch)             ((cch)*sizeof(CHAR))

#define SHCNFI_EVENT_STATECHANGE          0   // dwEventType
#define SHCNFI_EVENT_STRING               1   // e.string
#define SHCNFI_EVENT_HOTKEY               2   // e.hotkey
#define SHCNFI_EVENT_WNDPROC              3   // e.wndproc
#define SHCNFI_EVENT_WNDPROC_HOOK         4   // e.wndproc
#define SHCNFI_EVENT_ONCOMMAND            5   // e.command
#define SHCNFI_EVENT_INVOKECOMMAND        6   // e.command
#define SHCNFI_EVENT_TRACKPOPUPMENU       7   // e.command
#define SHCNFI_EVENT_DROP                 8   // e.drop
#define SHCNFI_EVENT_MAX                  9

#define SHCNFI_STRING_SHOWEXTVIEW         0

#define SHCNFI_STATE_KEYBOARDACTIVE         0   // _KEYBOARDACTIVE or _MOUSEACTIVE
#define SHCNFI_STATE_MOUSEACTIVE            1   // _KEYBOARDACTIVE or _MOUSEACTIVE
#define SHCNFI_STATE_ACCEL_TRAY             2   // _ACCEL_TRAY or _ACCEL_DESKTOP
#define SHCNFI_STATE_ACCEL_DESKTOP          3   // _ACCEL_TRAY or _ACCEL_DESKTOP
#define SHCNFI_STATE_START_DOWN             4   // _START_DOWN or _START_UP
#define SHCNFI_STATE_START_UP               5   // _START_DOWN or _START_UP
#define SHCNFI_STATE_TRAY_CONTEXT           6
#define SHCNFI_STATE_TRAY_CONTEXT_CLOCK     7
#define SHCNFI_STATE_TRAY_CONTEXT_START     8
#define SHCNFI_STATE_DEFVIEWX_ALT_DBLCLK    9
#define SHCNFI_STATE_DEFVIEWX_SHIFT_DBLCLK 10
#define SHCNFI_STATE_DEFVIEWX_DBLCLK       11

#define SERVERNAME_CURRENT  ((HANDLE)NULL)

#define SHCoUninitialize(hr) if (SUCCEEDED(hr)) CoUninitialize()

#define IS_BIDI_LOCALIZED_SYSTEM()       IsBiDiLocalizedSystem()



#define IS_WINDOW_RTL_MIRRORED(hwnd)     (g_bMirroredOS && Mirror_IsWindowMirroredRTL(hwnd))

#ifdef UNICODE
typedef NMVIEWFOLDERW NMVIEWFOLDER;
typedef LPNMVIEWFOLDERW LPNMVIEWFOLDER;
#else
typedef NMVIEWFOLDERA NMVIEWFOLDER;
typedef LPNMVIEWFOLDERA LPNMVIEWFOLDER;
#endif // UNICODE




#define IS_VALID_WRITE_PTR(ptr, type) \
   (IsBadWritePtr((PVOID)(ptr), sizeof(type)) ? \
    FALSE : TRUE)

#ifdef WANT_SHELL_INSTRUMENTATION
#define INSTRUMENT_STATECHANGE(t)                               \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_STATECHANGE;                \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_STRING(t,p)                                  \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_STRING;                     \
    lstrcpyn(s.e.string.sz,(p),ARRAYSIZE(s.e.string.sz));       \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_HOTKEY(t,w)                                  \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_HOTKEY;                     \
    s.e.hotkey.wParam=(w);                                      \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_WNDPROC(t,h,u,w,l)                           \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_WNDPROC;                    \
    s.e.wndproc.hwnd=(h);                                       \
    s.e.wndproc.uMsg=(u);                                       \
    s.e.wndproc.wParam=(w);                                     \
    s.e.wndproc.lParam=(l);                                     \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_WNDPROC_HOOK(h,u,w,l)                        \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=0;                                            \
    s.dwEventStructure=SHCNFI_EVENT_WNDPROC_HOOK;               \
    s.e.wndproc.hwnd=(h);                                       \
    s.e.wndproc.uMsg=(u);                                       \
    s.e.wndproc.wParam=(w);                                     \
    s.e.wndproc.lParam=(l);                                     \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_ONCOMMAND(t,h,u)                             \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_ONCOMMAND;                  \
    s.e.command.hwnd=(h);                                       \
    s.e.command.idCmd=(u);                                      \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_INVOKECOMMAND(t,h,u)                         \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_INVOKECOMMAND;              \
    s.e.command.hwnd=(h);                                       \
    s.e.command.idCmd=(u);                                      \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_TRACKPOPUPMENU(t,h,u)                        \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_TRACKPOPUPMENU;             \
    s.e.command.hwnd=(h);                                       \
    s.e.command.idCmd=(u);                                      \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#define INSTRUMENT_DROP(t,h,u,p)                                \
{                                                               \
    SHCNF_INSTRUMENT_INFO s;                                    \
    s.dwEventType=(t);                                          \
    s.dwEventStructure=SHCNFI_EVENT_DROP;                       \
    s.e.drop.hwnd=(h);                                          \
    s.e.drop.idCmd=(u);                                         \
    SHChangeNotify(SHCNE_INSTRUMENT,SHCNF_INSTRUMENT,&s,NULL);  \
}
#else
#define INSTRUMENT_STATECHANGE(t)
#define INSTRUMENT_STRING(t,p)
#define INSTRUMENT_HOTKEY(t,w)
#define INSTRUMENT_WNDPROC(t,h,u,w,l)
#define INSTRUMENT_WNDPROC_HOOK(h,u,w,l)
#define INSTRUMENT_ONCOMMAND(t,h,u)
#define INSTRUMENT_INVOKECOMMAND(t,h,u)
#define INSTRUMENT_TRACKPOPUPMENU(t,h,u)
#define INSTRUMENT_DROP(t,h,u,p)
#endif //WANT_SHELL_INSTRUMENTATION

#define CMF_ICM3                0x00020000      // QueryContextMenu can assume IContextMenu3 semantics (i.e.,
                                                // will receive WM_INITMENUPOPUP, WM_MEASUREITEM, WM_DRAWITEM,
                                                // and WM_MENUCHAR, via HandleMenuMsg2)

#define SHGUPP_FLAG_BASEPATH            0x00000001
#define SHGUPP_FLAG_DEFAULTPICSPATH     0x00000002
#define SHGUPP_FLAG_CREATE              0x80000000
#define SHGUPP_FLAG_VALID_MASK          0x80000003
#define SHGUPP_FLAG_INVALID_MASK        ~SHGUPP_FLAG_VALID_MASK

#define SPM_POST        0x0000
#define SPM_SEND        0x0001
#define SPM_ONELEVEL    0x0002  // default: send to all descendants including grandkids, etc.

#define SHCNEE_USERINFOCHANGED     11L
#define SHCNEE_SHORTCUTINVOKE       9L  // an app has been launched via a shortcut

#define CWM_TASKBARWAKEUP               (WM_USER + 26) // Used to restore tray thread to normal priority in extremely stressed machines
#define DTM_RAISE                       (WM_USER + 83)
#define DTRF_RAISE      0
#define DTRF_LOWER      1
#define DTM_SAVESTATE               (WM_USER + 77)
#define DTM_UPDATENOW               (WM_USER + 93)
#define DTM_UIACTIVATEIO            (WM_USER + 88)

#define WMTRAY_PROGCHANGE           (WM_USER + 200)     // 200=0xc8
#define WMTRAY_RECCHANGE            (WM_USER + 201)
#define WMTRAY_FASTCHANGE           (WM_USER + 202)
// was  WMTRAY_DESKTOPCHANGE        (WM_USER + 204)

#define WMTRAY_COMMONPROGCHANGE     (WM_USER + 205)
#define WMTRAY_COMMONFASTCHANGE     (WM_USER + 206)

#define WMTRAY_FAVORITESCHANGE      (WM_USER + 207)

#define WMTRAY_REGISTERHOTKEY       (WM_USER + 230)
#define WMTRAY_UNREGISTERHOTKEY     (WM_USER + 231)
#define WMTRAY_SETHOTKEYENABLE      (WM_USER + 232)
#define WMTRAY_SCREGISTERHOTKEY     (WM_USER + 233)
#define WMTRAY_SCUNREGISTERHOTKEY   (WM_USER + 234)
#define WMTRAY_QUERY_MENU           (WM_USER + 235)
#define WMTRAY_QUERY_VIEW           (WM_USER + 236)     // 236=0xec
#define WMTRAY_TOGGLEQL             (WM_USER + 237)

#define COF_NORMAL              0x00000000
#define COF_CREATENEWWINDOW     0x00000001      // "/N"
#define COF_USEOPENSETTINGS     0x00000002      // "/A"
#define COF_WAITFORPENDING      0x00000004      // Should wait for Pending
#define COF_EXPLORE             0x00000008      // "/E"
#define COF_NEWROOT             0x00000010      // "/ROOT"
#define COF_ROOTCLASS           0x00000020      // "/ROOT,<GUID>"
#define COF_SELECT              0x00000040      // "/SELECT"
#define COF_AUTOMATION          0x00000080      // The user is trying to use automation
#define COF_OPENMASK            0x000000FF
#define COF_NOTUSERDRIVEN       0x00000100      // Not user driven
#define COF_NOTRANSLATE         0x00000200      // Don't ILCombine(pidlRoot, pidl)
#define COF_INPROC              0x00000400      // not used
#define COF_CHANGEROOTOK        0x00000800      // Try Desktop root if not in our root
#define COF_NOUI                0x00001000      // Start background desktop only (no folder/explorer)
#define COF_SHDOCVWFORMAT       0x00002000      // indicates this struct has been converted to abide by shdocvw format. 
                                                // this flag is temporary until we rip out all the 
#define COF_NOFINDWINDOW        0x00004000      // Don't try to find the window
#define COF_HASHMONITOR         0x00008000      // pidlRoot in IETHREADPARAM struct contains an HMONITOR

//
//  GMI_TSCLIENT tells you whether you are running as a Terminal Server
//  client and should disable your animations.
//
#define GMI_TSCLIENT            0x0003  // Returns nonzero if TS client

#define ABE_MAX         4

#define TBSTYLE_EX_MULTICOLUMN              0x00000002 // conflicts w/ TBSTYLE_WRAPABLE
#define TBSTYLE_EX_VERTICAL                 0x00000004
#define TBSTYLE_EX_INVERTIBLEIMAGELIST      0x00000020  // Image list may contain inverted 
#define TBSTYLE_EX_FIXEDDROPDOWN            0x00000040 // Only used in the taskbar
#define TBSTYLE_EX_TRANSPARENTDEADAREA      0x00000100
#define TBSTYLE_EX_TOOLTIPSEXCLUDETOOLBAR   0x00000200

#define SMINIT_CUSTOMDRAW           0x00000080   // Send SMC_CUSTOMDRAW
#define SMINIT_USEMESSAGEFILTER     0x00000020
#define SMSET_USEPAGER              0x00000080    //Enable pagers in static menus
#define SMSET_NOPREFIX              0x00000100    //Enable ampersand in static menus
#define SMINV_POSITION       0x00000004

#define WS_EX_LAYOUTRTL         0x00400000L // Right to left mirroring

// StopWatch node types used in memory log to identify the type of node
#define EMPTY_NODE  0x0
#define START_NODE  0x1
#define LAP_NODE    0x2
#define STOP_NODE   0x3
#define OUT_OF_NODES 0x4

// Tray CopyData Messages
#define TCDM_APPBAR     0x00000000
#define TCDM_NOTIFY     0x00000001
#define TCDM_LOADINPROC 0x00000002



#define DCX_USESTYLE        0x00010000L
#define DCX_NEEDFONT        0x00020000L /* ;Internal */
#define DCX_NODELETERGN     0x00040000L /* ;Internal */
#define DCX_NOCLIPCHILDREN  0x00080000L /* ;Internal */
#define DCX_NORECOMPUTE     0x00100000L /* ;Internal */
#define DCX_VALIDATE        0x00200000L /* ;Internal */

#define ROUS_DEFAULTALLOW       0x0000
#define ROUS_DEFAULTRESTRICT    0x0001
#define ROUS_KEYALLOWS          0x0000
#define ROUS_KEYRESTRICTS       0x0002

#define TTF_EXCLUDETOOLAREA     0x4000

#define RRA_USEJOBOBJECTS         0x0020        // wait on job objects instead of process handles
#define RRA_NOUI                  0x0008        // prevents ShellExecuteEx from displaying error dialogs

#define SHCNFI_GLOBALHOTKEY               0
#define SHCNFI_TRAYCOMMAND                1
#define SHCNFI_TRAY_WNDPROC               3
#define SHCNFI_EVENT_ONCOMMAND            5   // e.command
#define SHCNFI_EVENT_HOTKEY               2   // e.hotkey

#define SHCNF_INSTRUMENT  0x0080        // dwItem1: LPSHCNF_INSTRUMENT

#define SHCNE_INSTRUMENT 0x00008000L    // no clue

// NOTE: NIS_SHOWALWAYS flag is defined with 0x20000000...

#define NISP_DEMOTED             0x00100000
#define NISP_STARTUPICON         0x00200000
#define NISP_ONCEVISIBLE         0x00400000
#define NISP_ITEMCLICKED         0x00800000
#define NISP_ITEMSAMEICONMODIFY  0x01000000
#define NISP_SHAREDICONSOURCE   0x10000000

#define SMPIN_POS(i) (LPCITEMIDLIST)MAKEINTRESOURCE((i)+1)
#define SMPINNABLE_EXEONLY          0x00000001 // allow only EXEs to be pinned
#define SMPINNABLE_REJECTSLOWMEDIA  0x00000002 // reject slow media

#define STRREG_SHELLUI          TEXT("ShellUIHandler")
#define STRREG_SHELL            TEXT("Shell")
#define STRREG_DEFICON          TEXT("DefaultIcon")
#define STRREG_SHEX             TEXT("shellex")
#define STRREG_SHEX_PROPSHEET   STRREG_SHEX TEXT("\\PropertySheetHandlers")
#define STRREG_SHEX_DDHANDLER   STRREG_SHEX TEXT("\\DragDropHandlers")
#define STRREG_SHEX_MENUHANDLER STRREG_SHEX TEXT("\\ContextMenuHandlers")
#define STRREG_SHEX_COPYHOOK    TEXT("Directory\\") STRREG_SHEX TEXT("\\CopyHookHandlers")
#define STRREG_SHEX_PRNCOPYHOOK TEXT("Printers\\") STRREG_SHEX TEXT("\\CopyHookHandlers")
#define STRREG_STARTMENU        TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Start Menu")
#define STRREG_STARTMENU2       TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Start Menu2")
#define STRREG_FAVORITES        TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Favorites")
#define STRREG_DISCARDABLE      TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Discardable")
#define STRREG_POSTSETUP        TEXT("\\PostSetup")

// BEGINNING OF IE 5.00 MESSAGES

#define DTM_GETVIEWAREAS            (WM_USER + 91)  // View area is WorkArea minus toolbar areas.
#define DTM_DESKTOPCONTEXTMENU      (WM_USER + 92)
#define DTM_UPDATENOW               (WM_USER + 93)

#define DTM_QUERYHKCRCHANGED        (WM_USER + 94)  // ask the desktop if HKCR has changed

#define DTM_MAKEHTMLCHANGES         (WM_USER + 95)  // Make changes to desktop html using dynamic HTML

#define DTM_STARTPAGEONOFF          (WM_USER + 96)  // Turn on/off the StartPage.

#define DTM_REFRESHACTIVEDESKTOP    (WM_USER + 97)  // Refresh the active desktop.

#define DTM_SETAPPSTARTCUR          (WM_USER + 98)  // UI feedback that we are starting an explorer window.

//
//  FT_ONEHOUR is the number of FILETIME units in an hour.
//  FT_ONEDAY is the number of FILETIME units in a day.
//
//      10,000,000 FILETIME units per second *
//      3600 seconds per hour *
//      24 hours per day.
//
#define FT_ONESECOND           ((unsigned __int64)10000000)
#define FT_ONEHOUR             ((unsigned __int64)10000000 * 3600)
#define FT_ONEDAY              ((unsigned __int64)10000000 * 3600 * 24)

enum
{
    ASFF_DEFAULT = 0x00000000, // There are no applicable Flags
    ASFF_SORTDOWN = 0x00000001, // Sort the items in this ISF to the bottom.
    ASFF_MERGESAMEGUID = 0x00000002, // Merge only namespaces with the same pguidObjects
    ASFF_COMMON = 0x00000004, // this is a "Common" or "All Users" folder
    // the following should all be collapsed to one ASFF_DEFNAMESPACE
    ASFF_DEFNAMESPACE_BINDSTG = 0x00000100, // The namespace is the default handler for BindToStorage() for merged child items.
    ASFF_DEFNAMESPACE_COMPARE = 0x00000200, // The namespace is the default handler for CompareIDs() for merged child items.
    ASFF_DEFNAMESPACE_VIEWOBJ = 0x00000400, // The namespace is the default handler for CreateViewObject() for merged child items.
    ASFF_DEFNAMESPACE_ATTRIB = 0x00001800, // The namespace is the default handler for GetAttributesOf() for merged child items.
    ASFF_DEFNAMESPACE_DISPLAYNAME = 0x00001000, // The namespace is the default handler for GetDisplayNameOf(), SetNameOf() and ParseDisplayName() for merged child items.
    ASFF_DEFNAMESPACE_UIOBJ = 0x00002000, // The namespace is the default handler for GetUIObjectOf() for merged child items.
    ASFF_DEFNAMESPACE_ITEMDATA = 0x00004000, // The namespace is the default handler for GetItemData() for merged child items.
    ASFF_DEFNAMESPACE_ALL = 0x0000FF00  // The namespace is the primary handler for all IShellFolder operations on merged child items.
};

//
//  FILETIME helpers
//
__inline UINT64 _FILETIMEtoInt64(const FILETIME* pft)
{
    return ((UINT64)pft->dwHighDateTime << 32) + pft->dwLowDateTime;
}

#define FILETIMEtoInt64(ft) _FILETIMEtoInt64(&(ft))


#define _ILSkip(pidl, cb)	((LPITEMIDLIST)(((BYTE*)(pidl))+cb))
#define _ILNext(pidl)		_ILSkip(pidl, (pidl)->mkid.cb)

//
// Enums
//



typedef enum tagASSOCQUERY
{
    AQ_NOTHING	= 0,
    AQS_FRIENDLYTYPENAME	= 0x170000,
    AQS_DEFAULTICON	= 0x70001,
    AQS_CONTENTTYPE	= 0x80070002,
    AQS_CLSID	= 0x70003,
    AQS_PROGID	= 0x70004,
    AQN_NAMED_VALUE	= 0x10f0000,
    AQNS_NAMED_MUI_STRING	= 0x1170001,
    AQNS_SHELLEX_HANDLER	= 0x81070002,
    AQVS_COMMAND	= 0x2070000,
    AQVS_DDECOMMAND	= 0x2070001,
    AQVS_DDEIFEXEC	= 0x2070002,
    AQVS_DDEAPPLICATION	= 0x2070003,
    AQVS_DDETOPIC	= 0x2070004,
    AQV_NOACTIVATEHANDLER	= 0x2060005,
    AQVD_MSIDESCRIPTOR	= 0x2060006,
    AQVS_APPLICATION_PATH	= 0x2010007,
    AQVS_APPLICATION_FRIENDLYNAME	= 0x2170008,
    AQVO_SHELLVERB_DELEGATE	= 0x2200000,
    AQVO_APPLICATION_DELEGATE	= 0x2200001,
    AQF_STRING	= 0x10000,
    AQF_EXISTS	= 0x20000,
    AQF_DIRECT	= 0x40000,
    AQF_DWORD	= 0x80000,
    AQF_MUISTRING	= 0x100000,
    AQF_OBJECT	= 0x200000,
    AQF_CUEIS_UNUSED	= 0,
    AQF_CUEIS_NAME	= 0x1000000,
    AQF_CUEIS_SHELLVERB	= 0x2000000,
    AQF_QUERY_INITCLASS	= 0x80000000
} ASSOCQUERY;

enum {
    SHDVID_FINALTITLEAVAIL,     // DEAD: variantIn bstr - sent after final OLECMDID_SETTITLE is sent
    SHDVID_MIMECSETMENUOPEN,    // mimecharset menu open commands
    SHDVID_PRINTFRAME,          // print HTML frame
    SHDVID_PUTOFFLINE,          // DEAD: The Offline property has been changed
    SHDVID_PUTSILENT,           // DEAD: The frame's Silent property has been changed
    SHDVID_GOBACK,              // Navigate Back
    SHDVID_GOFORWARD,           // Navigate Forward
    SHDVID_CANGOBACK,           // Is Back Navigation Possible?
    SHDVID_CANGOFORWARD,        // Is Forward Navigation Possible?
    SHDVID_CANACTIVATENOW,      // (down) (PICS) OK to navigate to this view now?
    SHDVID_ACTIVATEMENOW,       // (up) (PICS) Rating checks out, navigate now
    SHDVID_CANSUPPORTPICS,      // (down) variantIn I4: IOleCommandTarget to reply to
    SHDVID_PICSLABELFOUND,      // (up) variantIn bstr: PICS label
    SHDVID_NOMOREPICSLABELS,    // (up) End of document, no more PICS labels coming
    SHDVID_CANDEACTIVATENOW,    // (QS down) (in script/etc) OK to deactivate view now?
    SHDVID_DEACTIVATEMENOW,     // (EXEC up) (in script/etc) out of script, deactivate view now
    SHDVID_NODEACTIVATENOW,     // (EXEC up) (in script/etc) entering script, disable deactivate
    SHDVID_AMBIENTPROPCHANGE,   // variantIn I4: dispid of ambient property that changed
    SHDVID_GETSYSIMAGEINDEX,    // variantOut: image index for current page
    SHDVID_GETPENDINGOBJECT,    // variantOut: IUnknown of pending shellview/docobject
    SHDVID_GETPENDINGURL,       // variantOut: BSTR of URL for pending docobject
    SHDVID_SETPENDINGURL,       // variantIn: BSTR of URL passed to pending docobject
    SHDVID_ISDRAGSOURCE,        // (down) varioutOut I4: non-zero if it's initiated drag&drop
    SHDVID_DOCFAMILYCHARSET,    // variantOut: I4: windows (family) codepage
    SHDVID_DOCCHARSET,          // variantOut: I4: actual (mlang) codepage
    SHDVID_RAISE,               // vaIn:I4:DTRF_*, vaOut:NULL unless DTRF_QUERY
    SHDVID_GETTRANSITION,       // (down) vaIn: I4: TransitionEvent; vaOut BSTR (CLSID), I4 (dwSpeed)
    SHDVID_GETMIMECSETMENU,     // get menu handle for mimecharset
    SHDVID_DOCWRITEABORT,       // Abort binding but activate pending docobject
    SHDVID_SETPRINTSTATUS,      // VariantIn: BOOL, TRUE - Started printing, FALSE - Finished printing
    SHDVID_NAVIGATIONSTATUS,    // QS for tooltip text and Exec when user clicks
    SHDVID_PROGRESSSTATUS,      // QS for tooltip text and Exec when user clicks
    SHDVID_ONLINESTATUS,        // QS for tooltip text and Exec when user clicks
    SHDVID_SSLSTATUS,           // QS for tooltip text and Exec when user clicks
    SHDVID_PRINTSTATUS,         // QS for tooltip text and Exec when user clicks
    SHDVID_ZONESTATUS,          // QS for tooltip text and Exec when user clicks
    SHDVID_ONCODEPAGECHANGE,    // variantIn I4: new specified codepage
    SHDVID_SETSECURELOCK,       // set the secure icon
    SHDVID_SHOWBROWSERBAR,      // show browser bar of clsid guid
    SHDVID_NAVIGATEBB,          // navigate to pidl in browserbar.
    SHDVID_UPDATEOFFLINEDESKTOP,// put the desktop in ON-LINE mode, update and put it back in Offline mode
    SHDVID_PICSBLOCKINGUI,      // (up) In I4: pointer to "ratings nugget" for block API
    SHDVID_ONCOLORSCHANGE,      // (up) sent by mshtml to indicate color set change
    SHDVID_CANDOCOLORSCHANGE,   // (down) used to query if document supports the above
    SHDVID_QUERYMERGEDHELPMENU, // was the help menu micro-merged?
    SHDVID_QUERYOBJECTSHELPMENU,// return the object's help menu
    SHDVID_HELP,                // do help
    SHDVID_UEMLOG,              // set UEM logging vaIn:I4:UEMIND_*, vaOut:NULL
    SHDVID_GETBROWSERBAR,       // get IDeskBand for browser bar of clsid guid
    SHDVID_GETFONTMENU,
    SHDVID_FONTMENUOPEN,
    SHDVID_CLSIDTOIDM,          // get the idm for the given clsid
    SHDVID_GETDOCDIRMENU,       // get menu handle for document direction
    SHDVID_ADDMENUEXTENSIONS,   // Context Menu Extensions
    SHDVID_CLSIDTOMONIKER,      // CLSID to property page resource mapping
    SHDVID_RESETSTATUSBAR,      // set the status bar back to "normal" icon w/out text
    SHDVID_ISBROWSERBARVISIBLE, // is browser bar of clsid guid visible?
    SHDVID_GETOPTIONSHWND,      // gets hwnd for internet options prop sheet (NULL if not open)
    SHDVID_DELEGATEWINDOWOM,    // set policy for whether window OM methods should be delegated.
    SHDVID_PAGEFROMPOSTDATA,    // determines if page was generated by post data
    SHDVID_DISPLAYSCRIPTERRORS, // tells the top docobject host to display his script err dialog
    SHDVID_NAVIGATEBBTOURL,     // Navigate to an URL in browserbar (used in Trident).
    SHDVID_NAVIGATEFROMDOC,     // The document delegated the navigation for a non-html mime-type.
    SHDVID_STARTPICSFORWINDOW,  // (up) variantIn: IUnknown of window that is navigating
    //      variantOut: bool if pics process started
    SHDVID_CANCELPICSFORWINDOW, // (up) variantIn: IUnknown of window that is no longer navigating
    SHDVID_ISPICSENABLED,       // (up) variantOut: bool
    SHDVID_PICSLABELFOUNDINHTTPHEADER,// (up) variantIn bstr: PICS label
    SHDVID_CHECKINCACHEIFOFFLINE, // Check in cache if offline
    SHDVID_CHECKDONTUPDATETLOG,   // check if the current navigate is already dealing with the travellog correctly
    SHDVID_UPDATEDOCHOSTSTATE,    // Sent from CBaseBrowser2::_UpdateBrowserState to tell the dochost to update its state.
    SHDVID_FIREFILEDOWNLOAD,
    SHDVID_COMPLETEDOCHOSTPASSING,
    SHDVID_NAVSTART,
    SHDVID_SETNAVIGATABLECODEPAGE,
    SHDVID_WINDOWOPEN,
    SHDVID_PRIVACYSTATUS,        // QS for tooltip text and exec when user clicks
    SHDVID_FORWARDSECURELOCK,   // asks CDocObjectHost to forward its security status up to the shell browser
    SHDVID_ISEXPLORERBARVISIBLE, // is any explorer bar visible?
};


#define CONTEXTMENU_IDCMD_FIRST    1        // minimal QueryContextMenu idCmdFirst value //
#define CONTEXTMENU_IDCMD_LAST     0x7fff   // maximal QueryContextMenu idCmdLast value  //

#define IID_PPV_ARG(IType, ppType) IID_##IType, reinterpret_cast<void**>(static_cast<IType**>(ppType))
#define IID_X_PPV_ARG(IType, X, ppType) X, NULL, IID_##IType, reinterpret_cast<void**>(static_cast<IType**>(ppType))
#define MBANDCID_REFRESH            0x10000000

typedef enum
{
    RR_ALLOW = 1,
    RR_DISALLOW,
    RR_NOCHANGE,
} RESTRICTION_RESULT;

// IRestrict::IsRestricted() dwRestrictAction parameter values for
// the RID_RDeskBars pguidID.
typedef enum
{
    RA_DRAG = 1,
    RA_DROP,
    RA_ADD,
    RA_CLOSE,
    RA_MOVE,
} RESTRICT_ACTIONS;

// Cabinet_EnumRegApps flags 
#define RRA_DEFAULT               0x0000
#define RRA_DELETE                0x0001        // delete each reg value when we're done with it
#define RRA_WAIT                  0x0002        // Wait for current item to finish before launching next item
// was RRA_SHELLSERVICEOBJECTS    0x0004 -- do not reuse
#define RRA_NOUI                  0x0008        // prevents ShellExecuteEx from displaying error dialogs
#if (_WIN32_WINNT >= 0x0500)
#define RRA_USEJOBOBJECTS         0x0020        // wait on job objects instead of process handles
#endif

typedef UINT RRA_FLAGS;

//
// Function definitions
//

inline HRESULT(STDMETHODCALLTYPE* IUnknown_Exec)(IUnknown* punk, const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut);
inline HRESULT(STDMETHODCALLTYPE* IUnknown_GetClassID)(IUnknown* punk, CLSID* pclsid);
inline HRESULT(STDMETHODCALLTYPE* IUnknown_OnFocusChangeIS)(IUnknown* punk, IUnknown* punkSrc, BOOL fSetFocus);
inline HRESULT(STDMETHODCALLTYPE* IUnknown_QueryStatus)(IUnknown* punk, const GUID* pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pcmdtext);
inline HRESULT(STDMETHODCALLTYPE* IUnknown_UIActivateIO)(IUnknown* punk, BOOL fActivate, LPMSG lpMsg);
inline HRESULT(STDMETHODCALLTYPE* IUnknown_TranslateAcceleratorIO)(IUnknown* punk, LPMSG lpMsg);
HRESULT IUnknown_DragEnter(IUnknown* punk, IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
HRESULT IUnknown_DragOver(IUnknown* punk, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
HRESULT IUnknown_DragLeave(IUnknown* punk);

inline HRESULT(STDMETHODCALLTYPE* SHPropagateMessage)(HWND hwndParent, UINT uMsg, WPARAM wParam, LPARAM lParam, int iFlags);
inline HRESULT(STDMETHODCALLTYPE* SHGetUserDisplayName)(LPWSTR pszDisplayName, PULONG uLen);
inline HRESULT(STDMETHODCALLTYPE* SHSetWindowBits)(HWND hwnd, int iWhich, DWORD dwBits, DWORD dwValue);
inline HRESULT(STDMETHODCALLTYPE* SHRunIndirectRegClientCommand)(HWND hwnd, LPCWSTR pszClient);
inline HRESULT(STDMETHODCALLTYPE* SHInvokeDefaultCommand)(HWND hwnd, IShellFolder* psf, LPCITEMIDLIST pidlItem);
inline HRESULT(STDMETHODCALLTYPE* SHSettingsChanged)(WPARAM wParam, LPARAM lParam);
inline HRESULT(STDMETHODCALLTYPE* SHIsChildOrSelf)(HWND hwndParent, HWND hwnd);
inline HRESULT(STDMETHODCALLTYPE* SHLoadRegUIStringW)(HKEY     hkey, LPCWSTR  pszValue, LPWSTR   pszOutBuf, UINT     cchOutBuf);
inline HWND(STDMETHODCALLTYPE* SHCreateWorkerWindowW)(WNDPROC pfnWndProc, HWND hwndParent, DWORD dwExStyle, DWORD dwFlags, HMENU hmenu, void* p);
inline LRESULT(WINAPI* SHDefWindowProc)(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
inline BOOL(WINAPI* SHQueueUserWorkItem)(IN LPTHREAD_START_ROUTINE pfnCallback, IN LPVOID pContext, IN LONG lPriority, IN DWORD_PTR dwTag, OUT DWORD_PTR* pdwId OPTIONAL, IN LPCSTR pszModule OPTIONAL, IN DWORD dwFlags);
inline BOOL(WINAPI* WinStationSetInformationW)(HANDLE hServer, ULONG LogonId, WINSTATIONINFOCLASS WinStationInformationClass, PVOID  pWinStationInformation, ULONG WinStationInformationLength);
inline BOOL(WINAPI* WinStationUnRegisterConsoleNotification)(HANDLE hServer, HWND hWnd);
inline UINT(WINAPI* MsiDecomposeDescriptorW)(LPCWSTR	szDescriptor, LPWSTR szProductCode, LPWSTR szFeatureId, LPWSTR szComponentCode, DWORD* pcchArgsOffset);
inline BOOL(STDMETHODCALLTYPE* SHFindComputer)(LPCITEMIDLIST pidlFolder, LPCITEMIDLIST pidlSaveFile);
inline BOOL(STDMETHODCALLTYPE* SHTestTokenPrivilegeW)(HANDLE hToken, LPCWSTR pszPrivilegeName);
inline HRESULT(STDMETHODCALLTYPE* ExitWindowsDialog)(HWND hwndParent);
inline INT(STDMETHODCALLTYPE* SHMessageBoxCheckExW)(HWND hwnd, HINSTANCE hinst, LPCWSTR pszTemplateName, DLGPROC pDlgProc, LPVOID pData, int iDefault, LPCWSTR pszRegVal);
inline INT(STDMETHODCALLTYPE* RunFileDlg)(HWND hwndParent, HICON hIcon, LPCTSTR pszWorkingDir, LPCTSTR pszTitle, LPCTSTR pszPrompt, DWORD dwFlags);
inline UINT(STDMETHODCALLTYPE* SHGetCurColorRes)(void);
inline VOID(STDMETHODCALLTYPE* SHUpdateRecycleBinIcon)();
inline VOID(STDMETHODCALLTYPE* LogoffWindowsDialog)(HWND hwndParent);
inline VOID(STDMETHODCALLTYPE* DisconnectWindowsDialog)(HWND hwndParent);
inline COLORREF(STDMETHODCALLTYPE* SHFillRectClr)(HDC hdc, LPRECT lprect, COLORREF color);
inline HMENU(STDMETHODCALLTYPE* SHGetMenuFromID)(HMENU hmMain, UINT uID);
inline UINT(WINAPI* ImageList_GetFlags)(HIMAGELIST himl);
inline HRESULT(WINAPI *SHMapIDListToSystemImageListIndexAsync)(
    void *psts,
    void *psf,
    LPCITEMIDLIST pidlChild,
    void (CALLBACK *pfnCallback)(LPCITEMIDLIST pidl, LPVOID pvData, LPVOID pvHint, INT iIconIndex, INT iOpenIconIndex),
    void *pvCallbackData,
    void *pvCallbackHint,
    int *outIndex1,
    int *outIndex2
);
inline HRESULT(WINAPI* SHMapIDListToSystemImageListIndex)(
	void* psf,
	LPCITEMIDLIST pidlChild,
	int* outIndex1,
	int idk
	);
inline BOOL(WINAPI *IsShellManagedWindow)(HWND hwnd);
inline BOOL(WINAPI *IsShellFrameWindow)(HWND hwnd);
inline HWND(WINAPI *GhostWindowFromHungWindow)(HWND hwnd);
inline BOOL(WINAPI *GetWindowBand)(HWND hwnd, ZBID *band);
void  SHAdjustLOGFONT(IN OUT LOGFONT* plf);
BOOL  SHIsSameObject(IUnknown* punk1, IUnknown* punk2);
BOOL  SHAreIconsEqual(HICON hIcon1, HICON hIcon2);
BOOL  SHForceWindowZorder(HWND hwnd, HWND hwndInsertAfter);
BOOL  ShellExecuteRegApp(LPCTSTR pszCmdLine, RRA_FLAGS fFlags);
BOOL  IsRestrictedOrUserSetting(HKEY hkeyRoot, enum RESTRICTIONS rest, LPCTSTR pszSubKey, LPCTSTR pszValue, UINT flags);
HRESULT SHBindToIDListParent(LPCITEMIDLIST pidl, REFIID riid, void** ppv, LPCITEMIDLIST* ppidlLast); 
HRESULT SHCoInitialize(void);
DWORD  SHProcessMessagesUntilEventEx(HWND hwnd, HANDLE hEvent, DWORD dwTimeout, DWORD dwWakeMask);
TCHAR  SHFindMnemonic(LPCTSTR psz);

inline BOOL(STDMETHODCALLTYPE* RegisterShellHook)(HWND hwnd, BOOL fInstall);
DWORD Mirror_SetLayout(HDC hdc, DWORD dwLayout);
HRESULT VariantChangeTypeForRead(VARIANT* pvar, VARTYPE vtDesired);
BOOL GetExplorerUserSetting(HKEY hkeyRoot, LPCTSTR pszSubKey, LPCTSTR pszValue);
HRESULT SHBindToObjectEx(IShellFolder* psf, LPCITEMIDLIST pidl, LPBC pbc, REFIID riid, void** ppvOut);
HRESULT SHLoadLegacyRegUIString(HKEY hk, LPCTSTR pszSubkey, LPTSTR pszOutBuf, UINT cchOutBuf);

HRESULT DisplayNameOfAsOLESTR(IShellFolder* psf, LPCITEMIDLIST pidl, DWORD flags, LPWSTR* ppsz);
LPITEMIDLIST ILCloneParent(LPCITEMIDLIST pidl);
HRESULT SHGetIDListFromUnk(IUnknown* punk, LPITEMIDLIST* ppidl);
void _SHPrettyMenu(HMENU hm);
HRESULT DataObj_SetGlobal(IDataObject* pdtobj, UINT cf, HGLOBAL hGlobal);
BOOL GetInfoTip(IShellFolder* psf, LPCITEMIDLIST pidl, LPTSTR pszText, int cchTextMax);
HRESULT SHGetUIObjectFromFullPIDL(LPCITEMIDLIST pidl, HWND hwnd, REFIID riid, void** ppv);

#define GUIDSTR_MAX 38

typedef struct {
    TCHAR szSubkey[MAX_PATH];
    TCHAR szValueName[MAX_PATH];
    TCHAR szCmdLine[MAX_PATH];
} REGAPP_INFO;

// legacy from ripping this code out of explorer\initcab.cpp
inline BOOL g_fCleanBoot;   // are we running in SAFE-MODE?
inline BOOL g_fEndSession;  // did we process a WM_ENDSESSION?


typedef BOOL(WINAPI* PFNREGAPPSCALLBACK)(LPCTSTR szSubkey, LPCTSTR szCmdLine, RRA_FLAGS fFlags, LPARAM lParam);


BOOL ShellExecuteRegApp(LPCTSTR pszCmdLine, RRA_FLAGS fFlags);
BOOL Cabinet_EnumRegApps(HKEY hkeyParent, LPCTSTR pszSubkey, RRA_FLAGS fFlags, PFNREGAPPSCALLBACK pfnCallback, LPARAM lParam);
BOOL ExecuteRegAppEnumProc(LPCTSTR szSubkey, LPCTSTR szCmdLine, RRA_FLAGS fFlags, LPARAM lParam);

BOOL SHSetTermsrvAppInstallMode(BOOL bState);

inline BOOL(STDMETHODCALLTYPE* WinStationRegisterConsoleNotification)(HANDLE  hServer, HWND    hWnd, DWORD   dwFlags);


#define GMI_DOCKSTATE           0x0000
// Return values for SHGetMachineInfo(GMI_DOCKSTATE)
#define GMID_NOTDOCKABLE         0  // Cannot be docked
#define GMID_UNDOCKED            1  // Is undocked
#define GMID_DOCKED              2  // Is docked
inline DWORD_PTR(WINAPI* SHGetMachineInfo)(UINT gmi);

inline BOOL(WINAPI* EndTask)(HWND hWnd, BOOL fShutDown, BOOL fForce);

BOOL IsBiDiLocalizedSystem(void);
BOOL Mirror_IsWindowMirroredRTL(HWND hWnd);

VOID MuSecurity(VOID);

//LPNMVIEWFOLDER DDECreatePostNotify(LPNMVIEWFOLDER pnm);
inline LPNMVIEWFOLDER (*DDECreatePostNotify)(LPNMVIEWFOLDER pnm);
inline BOOL (*DDEHandleViewFolderNotify)(IShellBrowser* psb, HWND hwnd, LPNMVIEWFOLDER pnm);

unsigned __int64 _FILETIMEtoInt64(const FILETIME* pft);
void SetFILETIMEfromInt64(FILETIME* pft, unsigned __int64 i64);
void IncrementFILETIME(FILETIME* pft, unsigned __int64 iAdjust);
void DecrementFILETIME(FILETIME* pft, unsigned __int64 iAdjust);

CHAR CharUpperCharA(CHAR c);
WCHAR CharUpperCharW(WCHAR c);

typedef HANDLE LPSHChangeNotificationLock;
typedef INT_PTR BOOL_PTR;

typedef enum tagNETCON_SUBMEDIATYPE
{
    NCSM_NONE,
    NCSM_LAN,
    NCSM_WIRELESS,
    NCSM_ATM,
    NCSM_ELAN,
    NCSM_1394,
    NCSM_DIRECT,
    NCSM_IRDA,
    NCSM_CM
} NETCON_SUBMEDIATYPE;

#define OS_PERSONAL                 19 
#define STRREG_FAVORITES TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MenuOrder\\Favorites")
DEFINE_GUID(GUID_NETSHELL_PROPS, 0x2d15a9a1, 0xa556, 0x4189, 0x91, 0xad, 0x02, 0x74, 0x58, 0xf1, 0x1a, 0x07);
#define SMSET_NOEMPTY               0x00000004  
#define SMINIT_RESTRICT_CONTEXTMENU 0x00000001 
#define LIPF_ENABLE     0x00000001  // create the object (vs release the object)
#define LIPF_HOLDREF    0x00000002  // hold ref on object after creation (vs release immediately)
enum {
    SSOCMDID_OPEN = 2,
    SSOCMDID_CLOSE = 3,
};
#define REGSTR_PATH_SHELLSERVICEOBJECTDELAYED   TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\ShellServiceObjectDelayLoad")
// menubar orientation
#define MENUBAR_LEFT     ABE_LEFT
#define MENUBAR_TOP      ABE_TOP
#define MENUBAR_RIGHT    ABE_RIGHT
#define MENUBAR_BOTTOM   ABE_BOTTOM

// CmdID's for CGID_MENUDESKBAR
#define  MBCID_GETSIDE   1
#define  MBCID_RESIZE    2
#define  MBCID_SETEXPAND 3
#define  MBCID_SETFLAT   4
#define  MBCID_NOBORDER  5

#define RNC_LOGON                 0x00000002
#define OS_FRIENDLYLOGONUI          27  
#define TM_NEXTCTL                  (WM_USER + 0x15b)

#define DRVM_MAPPER             (0x2000)
#define DRVM_MAPPER_PREFERRED_GET                 (DRVM_MAPPER+21)
#define DRV_QUERYDEVICEINTERFACE     (DRV_RESERVED + 12)
#define DRV_QUERYDEVICEINTERFACESIZE (DRV_RESERVED + 13)

#define IS_WM_CONTEXTMENU_KEYBOARD(lParam) ((DWORD)(lParam) == 0xFFFFFFFF)

enum
{
    IDLHID_EMPTY = 0xBEEF0000,   //  where's the BEEF?!
    IDLHID_URLFRAGMENT,                     //  Fragment IDs on URLs (#anchors)
    IDLHID_URLQUERY,                        //  Query strings on URLs (?query+info)
    IDLHID_JUNCTION,                        //  Junction point data
    IDLHID_IDFOLDEREX,                      //  IDFOLDEREX, extended data for CFSFolder
    IDLHID_DOCFINDDATA,                     //  DocFind's private attached data (not persisted)
    IDLHID_PERSONALIZED,                    //  personalized like (My Docs/Zeke's Docs)
    IDLHID_recycle2,                        //  recycle
    IDLHID_RECYCLEBINDATA,                  //  RecycleBin private data (not persisted)
    IDLHID_RECYCLEBINORIGINAL,              //  the original unthunked path for RecycleBin items
    IDLHID_PARENTFOLDER,                    //  merged folder uses this to encode the source folder.
    IDLHID_STARTPANEDATA,                   //  Start Pane's private attached data
    IDLHID_NAVIGATEMARKER                   //  Used by Control Panel's 'Category view'               
};

typedef DWORD IDLHID;
typedef struct _HIDDENITEMID
{
    WORD    cb;     //  hidden item size
    WORD    wVersion;
    IDLHID  id;     //  hidden item ID
} HIDDENITEMID;
#pragma pack()

typedef HIDDENITEMID UNALIGNED* PIDHIDDEN;
typedef const HIDDENITEMID UNALIGNED* PCIDHIDDEN;

#define SHCNEE_PINLISTCHANGED      10L
#define II_DOCNOASSOC   0         // document (blank page) (not associated)
#define II_DOCUMENT     1         // document (with stuff on the page)
#define II_APPLICATION  2         // application (exe, com, bat)
#define II_FOLDER       3         // folder (plain)

#define SMSET_USEBKICONEXTRACTION   0x00000008


typedef enum tagWALK_TREE_CMD
{
    WALK_TREE_SAVE,
    WALK_TREE_DELETE,
    WALK_TREE_RESTORE,
    WALK_TREE_REFRESH
} WALK_TREE_CMD;


BOOL WINAPI SHWinHelp(HWND hwndMain, LPCTSTR lpszHelp, UINT usCommand, ULONG_PTR ulData);
PCIDHIDDEN ILFindHiddenIDOn(LPCITEMIDLIST pidl, IDLHID id, BOOL fOnLast);
#define ILFindHiddenID(p, i)    ILFindHiddenIDOn((p), (i), TRUE)
void ILExpungeRemovedHiddenIDs(LPITEMIDLIST pidl);
BOOL ILRemoveHiddenID(LPITEMIDLIST pidl, IDLHID id);
LPITEMIDLIST ILAppendHiddenID(LPITEMIDLIST pidl, PCIDHIDDEN pidhid);


enum
{
    XLATEALIAS_MYDOCS = 0x00000001,
    XLATEALIAS_DESKTOP = 0x00000002,
    XLATEALIAS_COMMONDOCS = 0x00000003,   // REVIEW: XLATEALIAS_DESKTOP & XLATEALIAS_MYDOCS ?
    //  XLATEALIAS_MYPICS,    
    //  XLATEALIAS_NETHOOD,
};
#define XLATEALIAS_ALL  ((DWORD)0x0000ffff)
#define SHID_FS_COMMONITEM        0x38  // Common item ("8" is the bit)
LPITEMIDLIST SHLogILFromFSIL(LPCITEMIDLIST pidlFS);

#define LVM_KEYBOARDSELECTED    (LVM_FIRST + 178)
#define ListView_KeyboardSelected(hwnd, i) \
    (BOOL)SNDMSG((hwnd), LVM_KEYBOARDSELECTED, (WPARAM)(i), 0)
HRESULT ContextMenu_GetCommandStringVerb(IContextMenu* pcm, UINT idCmd, LPWSTR pszVerb, int cchVerb);
UINT GetMenuIndexForCanonicalVerb(HMENU hMenu, IContextMenu* pcm, UINT idCmdFirst, LPCWSTR pwszVerb);
HRESULT ContextMenu_DeleteCommandByName(IContextMenu* pcm, HMENU hpopup, UINT idFirst, LPCWSTR pszCommand);
enum {
    OBJCOMPATF_OTNEEDSSFCACHE = 0x00000001,
    OBJCOMPATF_NO_WEBVIEW = 0x00000002,
    OBJCOMPATF_UNBINDABLE = 0x00000004,
    OBJCOMPATF_PINDLL = 0x00000008,
    OBJCOMPATF_NEEDSFILESYSANCESTOR = 0x00000010,
    OBJCOMPATF_NOTAFILESYSTEM = 0x00000020,
    OBJCOMPATF_CTXMENU_NOVERBS = 0x00000040,
    OBJCOMPATF_CTXMENU_LIMITEDQI = 0x00000080,
    OBJCOMPATF_COCREATESHELLFOLDERONLY = 0x00000100,
    OBJCOMPATF_NEEDSSTORAGEANCESTOR = 0x00000200,
    OBJCOMPATF_NOLEGACYWEBVIEW = 0x00000400,
    OBJCOMPATF_BLOCKSHELLSERVICEOBJECT = 0x00000800,
};

// {E9779583-939D-11CE-8A77-444553540000}
static const GUID GUID_AECOZIPARCHIVE =
{ 0xE9779583, 0x939D, 0x11ce, { 0x8a, 0x77, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} };
// {49707377-6974-6368-2E4A-756E6F644A01}
static const GUID CLSID_WS_FTP_PRO_EXPLORER =
{ 0x49707377, 0x6974, 0x6368, {0x2E, 0x4A,0x75, 0x6E, 0x6F, 0x64, 0x4A, 0x01} };
// {49707377-6974-6368-2E4A-756E6F644A0A}
static const GUID CLSID_WS_FTP_PRO =
{ 0x49707377, 0x6974, 0x6368, {0x2E, 0x4A,0x75, 0x6E, 0x6F, 0x64, 0x4A, 0x0A} };
// {2bbbb600-3f0a-11d1-8aeb-00c04fd28d85}
static const GUID CLSID_KODAK_DC260_ZOOM_CAMERA =
{ 0x2bbbb600, 0x3f0a, 0x11d1, {0x8a, 0xeb, 0x00, 0xc0, 0x4f, 0xd2, 0x8d, 0x85} };
// {00F43EE0-EB46-11D1-8443-444553540000}
static const GUID GUID_MACINDOS =
{ 0x00F43EE0, 0xEB46, 0x11D1, { 0x84, 0x43, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} };
static const GUID CLSID_EasyZIP =
{ 0xD1069700, 0x932E, 0x11cf, { 0xAB, 0x59, 0x00, 0x60, 0x8C, 0xBF, 0x2C, 0xE0} };

static const GUID CLSID_PAGISPRO_FOLDER =
{ 0x7877C8E0, 0x8B13, 0x11D0, { 0x92, 0xC2, 0x00, 0xAA, 0x00, 0x4B, 0x25, 0x6F} };
// {61E285C0-DCF4-11cf-9FF4-444553540000}
static const GUID CLSID_FILENET_IDMDS_NEIGHBORHOOD =
{ 0x61e285c0, 0xdcf4, 0x11cf, { 0x9f, 0xf4, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} };

// These guys call CoFreeUnusedLibraries inside their Release() handler, so
// if you are releasing the last object, they end up FreeLibrary()ing
// themselves!

// {b8777200-d640-11ce-b9aa-444553540000}
static const GUID CLSID_NOVELLX =
{ 0xb8777200, 0xd640, 0x11ce, { 0xb9, 0xaa, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} };

static const GUID CLSID_PGP50_CONTEXTMENU =  //{969223C0-26AA-11D0-90EE-444553540000}
{ 0x969223C0, 0x26AA, 0x11D0, { 0x90, 0xEE, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00} };

static const GUID CLSID_QUICKFINDER_CONTEXTMENU = //  {CD949A20-BDC8-11CE-8919-00608C39D066}
{ 0xCD949A20, 0xBDC8, 0x11CE, { 0x89, 0x19, 0x00, 0x60, 0x8C, 0x39, 0xD0, 0x66} };

static const GUID CLSID_HERCULES_HCTNT_V1001 = // {921BD320-8CB5-11CF-84CF-885835D9DC01}
{ 0x921BD320, 0x8CB5, 0x11CF, { 0x84, 0xCF, 0x88, 0x58, 0x35, 0xD9, 0xDC, 0x01} };

typedef struct {
    DWORD flag;
    LPCTSTR psz;
} FLAGMAP;
typedef DWORD OBJCOMPATFLAGS;

HRESULT SHGetNameAndFlags(LPCITEMIDLIST pidl, DWORD dwFlags, LPTSTR pszName, UINT cchName, DWORD* pdwAttribs);
HRESULT DisplayNameOf(IShellFolder* psf, LPCITEMIDLIST pidl, DWORD flags, LPTSTR psz, UINT cch);
#define SHGetAttributesOf(pidl, prgfInOut) SHGetNameAndFlags(pidl, 0, NULL, 0, prgfInOut)
#define ToolBar_CommandToIndex(hwnd, idBtn)  \
    (BOOL)SNDMSG((hwnd), TB_COMMANDTOINDEX, (WPARAM)(idBtn), 0)

//
// Function loader
//
bool SHUndocInit(void);

const CLSID CLSID_TrayNotify = { 0x25DEAD04, 0x1EAC, 0x4911,{ 0x9E, 0x3A, 0xAD, 0x0A, 0x4A, 0xB5, 0x60, 0xFD } };
const CLSID CLSID_FadeTask = { 0x7EB5FBE4, 0x2100, 0x49E6, { 0x85, 0x93, 0x17, 0xE1, 0x30, 0x12, 0x2F, 0x91} };
//DEFINE_GUID(IID_IShellService, 0x5836FB00L, 0x8187, 0x11CF, 0xA1, 0x2B, 0x00, 0xAA, 0x00, 0x4A, 0xE8, 0x37);

DEFINE_GUID(IID_IAssociationElement, 0xD8F6AD5B, 0xB44F, 0x4BCC, 0x88, 0xFD, 0xEB, 0x34, 0x73, 0xDB, 0x75, 0x02);

DEFINE_GUID(CLSID_PostBootReminder, 0x7849596a, 0x48ea, 0x486e, 0x89, 0x37, 0xa2, 0xa3, 0x00, 0x9f, 0x31, 0xa9);
DEFINE_GUID(IID_IShellReminderManager, 0x968edb91, 0x8a70, 0x4930, 0x83, 0x32, 0x5f, 0x15, 0x83, 0x8a, 0x64, 0xf9);

// for start
DEFINE_GUID(CLSID_StartMenuPin, 0xA2A9545D, 0xA0C2, 0x42B4, 0x97, 0x08, 0xA0, 0xB2, 0xBA, 0xDD, 0x77, 0xC8);
DEFINE_GUID(IID_IPinnedList2, 0xBBD20037, 0xBC0E, 0x42F1, 0x91, 0x3F, 0xE2, 0x93, 0x6B, 0xB0, 0xEA, 0x0C);
DEFINE_GUID(IID_IPinnedList25, 0x446BC432, 0x57E9, 0x4B72, 0x8E, 0x0F1, 0x0AF, 0x27, 0x11, 0x3D, 0x0CF, 0x9C);
DEFINE_GUID(IID_IFlexibleTaskbarPinnedList, 0x60274fa2, 0x611f, 0x4b8a, 0xa2, 0x93, 0xf2, 0x7b, 0xf1, 0x03, 0xd1, 0x48);
DEFINE_GUID(IID_IPinnedList3, 0x0dd79ae2, 0xd156, 0x45d4, 0x9e, 0xeb, 0x3b, 0x54, 0x97, 0x69, 0xe9, 0x40);

// {E6FB5E20-DE35-11CF-9C87-00AA005127ED}
DEFINE_GUID(CLSID_WebCheck, 0xE6FB5E20, 0xDE35, 0x11CF, 0x9C, 0x87, 0x00, 0xAA, 0x00, 0x51, 0x27, 0xED);
// {08165EA0-E946-11CF-9C87-00AA005127ED}
DEFINE_GUID(CLSID_WebCheckNew, 0x08165EA0, 0xE946, 0x11CF, 0x9C, 0x87, 0x00, 0xAA, 0x00, 0x51, 0x27, 0xED);

DEFINE_GUID(IID_IUserAssist10, 0x49B36D57, 0x5FD2, 0x45A7, 0x98, 0x1B, 0x6, 0x2, 0x8D, 0x57, 0x7A, 0x47);

MIDL_INTERFACE("968edb91-8a70-4930-8332-5f15838a64f9")
IShellReminderManager : IUnknown
{
    virtual HRESULT Add(const SHELLREMINDER* psr) PURE;
    virtual HRESULT Delete(LPCWSTR pszName) PURE;
    virtual HRESULT Enum(void** ppesr) PURE;
};

BOOL IsStartPanelOn();

LPIDA DataObj_GetHIDAEx(IDataObject* pdtobj, CLIPFORMAT cf, STGMEDIUM* pmedium);

LPIDA DataObj_GetHIDA(IDataObject* pdtobj, STGMEDIUM* pmedium);
#define HIDA_GetPIDLItem(pida, i) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[i+1])
LPCITEMIDLIST IDA_GetIDListPtr(LPIDA pida, UINT i);

LPITEMIDLIST IDA_FullIDList(LPIDA pida, UINT i);

BOOL _IsLocalHardDisk(LPCTSTR pszPath);

BOOL _IsAcceptableTarget(LPCTSTR pszPath, DWORD dwAttrib, DWORD dwFlags);

void ReleaseStgMediumHGLOBAL(void* pv, STGMEDIUM* pmedium);

void HIDA_ReleaseStgMedium(LPIDA pida, STGMEDIUM* pmedium);

HRESULT IsPinnable(IDataObject* pdtobj, DWORD dwFlags, OPTIONAL LPITEMIDLIST* ppidl);

inline HRESULT(WINAPI* SHGetUserPicturePath_t)(LPCWSTR pszUsername, DWORD dwFlags, LPWSTR pszPath, DWORD cchPathMax);
HRESULT WINAPI SHGetUserPicturePath(LPCWSTR pszUsername, DWORD dwFlags, LPWSTR pszPath);

inline HRESULT(*CFSFolder_CreateFolder)(IUnknown* punkOuter, LPBC pbc, LPCITEMIDLIST pidl, const PERSIST_FOLDER_TARGET_INFO* pf, REFIID riid, void** ppv);
inline HRESULT(*SHCreatePropertyBagOnMemory)(DWORD grfMode, REFIID riid, void** ppv);
inline HRESULT(*SHPropertyBag_WriteBOOL)(IPropertyBag* ppb, LPCWSTR pszPropName, BOOL fValue);

DEFINE_GUID(CLSID_MruLongList, 0x53bd6b4e, 0x3780, 0x4693, 0xaf, 0xc3, 0x71, 0x61, 0xc2, 0xf3, 0xee, 0x9c);
DEFINE_GUID(IID_IAugmentedFolder, 0x2f711b17, 0x773c, 0x41d4, 0x93, 0xfa, 0x7f, 0x23, 0xed, 0xce, 0xcb, 0x66);
DEFINE_GUID(CLSID_UserEventTimer, 0x864A1288, 0x354C, 0x4D19, 0x9D, 0x68, 0x0C2, 0x74, 0x2B, 0x0B1, 0x49, 0x97);

// Note: SHRegisterDarwinLink takes ownership of pidlFull. fUpdate means: update the Darwin state right away
BOOL SHRegisterDarwinLink(LPITEMIDLIST pidlFull, LPWSTR pszDarwinID, BOOL fUpdate);

// Use this function to update the Darwin state for all registered Darwin shortcuts.
void SHReValidateDarwinCache();

HRESULT SHParseDarwinIDFromCacheW(LPWSTR pszDarwinDescriptor, LPWSTR* ppwszOut);

inline void(*CheckWinIniForAssocs)();
inline HRESULT(*CheckDiskSpace)();
inline HRESULT(*CheckStagingArea)();
DEFINE_GUID(CLSID_ProgramsFolderAndFastItems, 0x865E5E76, 0x0AD83, 0x4DCA, 0x0A1, 0x9, 0x50, 0x0DC, 0x21, 0x13, 0x0CE, 0x9C);
DEFINE_GUID(CLSID_ProgramsFolder, 0x865E5E76, 0x0AD83, 0x4DCA, 0x0A1, 0x9, 0x50, 0x0DC, 0x21, 0x13, 0x0CE, 0x9D);
DEFINE_GUID(CLSID_StartMenuFastItems, 0x865E5E76, 0x0AD83, 0x4DCA, 0x0A1, 0x9, 0x50, 0x0DC, 0x21, 0x13, 0x0CE, 0x9E);
HRESULT CStartMenuFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
HRESULT CProgramsFolder_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);
HRESULT CStartMenuFastItems_CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv);

//DEFINE_GUID(CLSID_UserAssist, 0xDD313E04, 0xFEFF, 0x11D1, 0x8E, 0xCD, 0x00, 0x00, 0xF8, 0x7A, 0x47, 0x0C);

#include "interfacesp.inc"

class CPinnedListWrapper : public IStartMenuPin
{
public:
    CPinnedListWrapper(IUnknown *punk, const GUID &riid);
    ~CPinnedListWrapper();

    //IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
    STDMETHODIMP_(ULONG) AddRef(void);
    STDMETHODIMP_(ULONG) Release(void);

    //IPinnedList2
    STDMETHODIMP EnumObjects(IEnumFullIDList**);
    STDMETHODIMP Modify(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE);
    STDMETHODIMP GetChangeCount(ULONG*);
    STDMETHODIMP GetPinnableInfo(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*);
    STDMETHODIMP IsPinnable(IDataObject*, int);
    STDMETHODIMP Resolve(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*);
    STDMETHODIMP IsPinned(PCIDLIST_ABSOLUTE);
    STDMETHODIMP GetPinnedItem(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*);
    STDMETHODIMP GetAppIDForPinnedItem(PCIDLIST_ABSOLUTE, PWSTR*);
    STDMETHODIMP ItemChangeNotify(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE);
    STDMETHODIMP UpdateForRemovedItemsAsNecessary(VOID);
private:
    IFlexibleTaskbarPinnedList* m_flexList = 0;
    IPinnedList3* m_pinnedList3 = 0;
    IPinnedList25* m_pinnedList25 = 0;
    //int m_build = 0;
};

inline HWND WINAPI GetTaskmanWindow(void)
{
    HMODULE hmUser32 = LoadLibrary(TEXT("user32.dll"));
    HWND result = NULL;

    if (hmUser32)
    {
        auto pfn = (HWND (WINAPI *)(void))GetProcAddress(hmUser32, "GetTaskmanWindow");
        if (pfn)
            result = pfn();
        FreeLibrary(hmUser32);
    }


    return result;
}

inline BOOL WINAPI SetTaskmanWindow(IN HWND hWnd)
{
    HMODULE hmUser32 = LoadLibrary(TEXT("user32.dll"));
    BOOL result = FALSE;

    if (hmUser32)
    {
        auto pfn = (BOOL (WINAPI *)(IN HWND hWnd))GetProcAddress(hmUser32, "SetTaskmanWindow");
        if (pfn)
            result = pfn(hWnd);
        FreeLibrary(hmUser32);
    }

    return result;
}

HRESULT WINAPI DwmpStartOrStopFlip3D();
