#include "pch.h"

#include <wil/result_macros.h>

#include "cocreateinstancehook.h"
#include "shfusion.h"
#include "cabinet.h"
#include "rcids.h"
#include "startmnu.h"
#include "shdguid.h"    // for IID_IShellService
#include "tray.h"
#include "util.h"

// global so that it is shared between TS sessions
#define SZ_SCMCREATEDEVENT_NT5  TEXT("Global\\ScmCreatedEvent")
#define SZ_WINDOWMETRICS        TEXT("Control Panel\\Desktop\\WindowMetrics")
#define SZ_APPLIEDDPI           TEXT("AppliedDPI")
#define SZ_CONTROLPANEL         TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Control Panel")
#define SZ_ORIGINALDPI          TEXT("OriginalDPI")

// exports from shdocvw.dll
void RunInstallUninstallStubs(void);

int ExplorerWinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPTSTR pszCmdLine, int nCmdShow);

BOOL _ShouldFixResolution(void);

#ifdef PERF_ENABLESETMARK
#include <wmistr.h>
#include "ntwmi.h"  // PWMI_SET_MARK_INFORMATION is defined in ntwmi.h
#include "wmiumkm.h"
#define NTPERF
#include "ntperf.h"
#include <desktopp.h>

void DoSetMark(LPCSTR pszMark, ULONG cbSz)
{
    PWMI_SET_MARK_INFORMATION MarkInfo;
    HANDLE hTemp;
    ULONG cbBufferSize;
    ULONG cbReturnSize;

    cbBufferSize = FIELD_OFFSET(WMI_SET_MARK_INFORMATION, Mark) + cbSz;

    MarkInfo = (PWMI_SET_MARK_INFORMATION) LocalAlloc(LPTR, cbBufferSize);

    // Failed to init, no big deal
    if (MarkInfo == NULL)
        return;

    BYTE *pMarkBuffer = (BYTE *) (&MarkInfo->Mark[0]);

    memcpy(pMarkBuffer, pszMark, cbSz);

    // WMI_SET_MARK_WITH_FLUSH will flush the working set when setting the mark
    MarkInfo->Flag = PerformanceMmInfoMark;

    hTemp = CreateFile(WMIDataDeviceName,
                           GENERIC_READ | GENERIC_WRITE,
                           0,
                           NULL,
                           OPEN_EXISTING,
                           FILE_ATTRIBUTE_NORMAL |
                           FILE_FLAG_OVERLAPPED,
                           NULL);

    if (hTemp != INVALID_HANDLE_VALUE)
    {
        // here's the piece that actually puts the mark in the buffer
        BOOL fIoctlSuccess = DeviceIoControl(hTemp,
                                       IOCTL_WMI_SET_MARK,
                                       MarkInfo,
                                       cbBufferSize,
                                       NULL,
                                       0,
                                       &cbReturnSize,
                                       NULL);

        CloseHandle(hTemp);
    }
    LocalFree(MarkInfo);
}
#endif  // PERF_ENABLESETMARK


//Do not change this stock5.lib use this as a BOOL not a bit.
BOOL g_bMirroredOS = FALSE;

HINSTANCE g_hinstCabinet = 0;

HKEY g_hkeyExplorer = nullptr;

#define MAGIC_FAULT_TIME    (1000 * 60 * 5)
#define MAGIC_FAULT_LIMIT   (2)

BOOL g_fLogonCycle = FALSE;
BOOL g_fCleanShutdown = TRUE;
BOOL g_fExitExplorer = TRUE; // set to FALSE on WM_ENDSESSION shutdown case
//BOOL g_fEndSession = FALSE;             // set to TRUE if we rx a WM_ENDSESSION during RunOnce etc
BOOL g_fFakeShutdown = FALSE;           // set to TRUE if we do Ctrl+Alt+Shift+Cancel shutdown

DWORD g_dwStopWatchMode;                // to minimize impact of perf logging on retail



// helper function to check to see if a given regkey is has any subkeys
BOOL SHKeyHasSubkeys(HKEY hk, LPCTSTR pszSubKey)
{
    HKEY hkSub;
    BOOL bHasSubKeys = FALSE;

    // need to open this with KEY_QUERY_VALUE or else RegQueryInfoKey will fail
    if (RegOpenKeyEx(hk,
                     pszSubKey,
                     0,
                     KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS,
                     &hkSub) == ERROR_SUCCESS)
    {
        DWORD dwSubKeys;

        if (RegQueryInfoKey(hkSub, NULL, NULL, NULL, &dwSubKeys, NULL, NULL, NULL, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
        {
            bHasSubKeys = (dwSubKeys != 0);
        }

        RegCloseKey(hkSub);
    }

    return bHasSubKeys;
}


#ifdef _WIN64
// helper function to check to see if a given regkey is has values (ignores the default value)
BOOL SHKeyHasValues(HKEY hk, LPCTSTR pszSubKey)
{
    HKEY hkSub;
    BOOL bHasValues = FALSE;

    if (RegOpenKeyEx(hk,
                     pszSubKey,
                     0,
                     KEY_QUERY_VALUE,
                     &hkSub) == ERROR_SUCCESS)
    {
        DWORD dwValues;
        DWORD dwSubKeys;

        if (RegQueryInfoKey(hkSub, NULL, NULL, NULL, &dwSubKeys, NULL, NULL, &dwValues, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
        {
            bHasValues = (dwValues != 0);
        }

        RegCloseKey(hkSub);
    }

    return bHasValues;
}
#endif // _WIN64


void CreateShellDirectories()
{
    WCHAR szPath[260];
    SHGetSpecialFolderPathW(nullptr, szPath, CSIDL_DESKTOPDIRECTORY, TRUE);
    SHGetSpecialFolderPathW(nullptr, szPath, CSIDL_PROGRAMS, TRUE);
    SHGetSpecialFolderPathW(nullptr, szPath, CSIDL_STARTMENU, TRUE);
}

// returns:
//      TRUE if the user wants to abort the startup sequence
//      FALSE keep going
//
// note: this is a switch, once on it will return TRUE to all
// calls so these keys don't need to be pressed the whole time
BOOL AbortStartup()
{
    static BOOL bAborted = FALSE;       // static so it sticks!

    if (bAborted)
    {
        return TRUE;    // don't do funky startup stuff
    }
    else 
    {
        bAborted = (g_fCleanBoot || ((GetKeyState(VK_CONTROL) < 0) || (GetKeyState(VK_SHIFT) < 0)));
        return bAborted;
    }
}

BOOL ExecStartupEnumProc(IShellFolder *psf, LPITEMIDLIST pidlItem)
{
    IContextMenu *pcm;
    HRESULT hr = psf->GetUIObjectOf(NULL, 1, (LPCITEMIDLIST*)&pidlItem, IID_IContextMenu, nullptr, (void**)&pcm);
    if (SUCCEEDED(hr))
    {
        HMENU hmenu = CreatePopupMenu();
        if (hmenu)
        {
            pcm->QueryContextMenu(hmenu, 0, CONTEXTMENU_IDCMD_FIRST, CONTEXTMENU_IDCMD_LAST, CMF_DEFAULTONLY);
            INT idCmd = GetMenuDefaultItem(hmenu, MF_BYCOMMAND, 0);
            if (idCmd)
            {
                CMINVOKECOMMANDINFOEX ici = {0};

                ici.cbSize = sizeof(ici);
                ici.fMask = CMIC_MASK_FLAG_NO_UI;
                ici.lpVerb = (LPSTR)MAKEINTRESOURCE(idCmd - CONTEXTMENU_IDCMD_FIRST);
                ici.nShow = SW_NORMAL;

                if (FAILED(pcm->InvokeCommand((LPCMINVOKECOMMANDINFO)&ici)))
                {
                    c_tray.LogFailedStartupApp();
                }
            }
            DestroyMenu(hmenu);
        }
        pcm->Release();
    }

    return !AbortStartup();
}

typedef BOOL (*PFNENUMFOLDERCALLBACK)(IShellFolder *psf, LPITEMIDLIST pidlItem);

void EnumFolder(LPITEMIDLIST pidlFolder, DWORD grfFlags, PFNENUMFOLDERCALLBACK pfn)
{
    IShellFolder *psf;
    if (SUCCEEDED(SHBindToObject(NULL, IID_X_PPV_ARG(IShellFolder, pidlFolder, &psf))))
    {
        IEnumIDList *penum;
        if (S_OK == psf->EnumObjects(NULL, grfFlags, &penum))
        {
            LPITEMIDLIST pidl;
            ULONG celt;
            while (S_OK == penum->Next(1, &pidl, &celt))
            {
                BOOL bRet = pfn(psf, pidl);

                SHFree(pidl);

                if (!bRet)
                    break;
            }
            penum->Release();
        }
        psf->Release();
    }
}

const UINT c_rgStartupFolders[] = {
    CSIDL_COMMON_STARTUP,
    //CSIDL_COMMON_ALTSTARTUP,    // non-localized "Common StartUp" group if exists.
    CSIDL_STARTUP,
    //CSIDL_ALTSTARTUP            // non-localized "StartUp" group if exists.
};

void _ExecuteStartupPrograms()
{
    if (!AbortStartup())
    {
        for (int i = 0; i < ARRAYSIZE(c_rgStartupFolders); i++)
        {
            LPITEMIDLIST pidlStartup = SHCloneSpecialIDList(NULL, c_rgStartupFolders[i], FALSE);
            if (pidlStartup)
            {
                EnumFolder(pidlStartup, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS, ExecStartupEnumProc);
                ILFree(pidlStartup);
            }
        }
    }
}


// helper function for parsing the run= stuff
BOOL ExecuteOldEqualsLine(LPTSTR pszCmdLine, int nCmdShow)
{
    BOOL bRet = FALSE;
    TCHAR szWindowsDir[MAX_PATH];
    // Load and Run lines are done relative to windows directory.
    if (GetWindowsDirectory(szWindowsDir, ARRAYSIZE(szWindowsDir)))
    {
        BOOL bFinished = FALSE;
        while (!bFinished && !AbortStartup())
        {
            LPTSTR pEnd = pszCmdLine;

            // NOTE: I am guessing from the code below that you can have multiple entries seperated 
            //       by a ' '  or a ',' and we will exec all of them.
            while ((*pEnd) && (*pEnd != TEXT(' ')) && (*pEnd != TEXT(',')))
            {
                pEnd = (LPTSTR)CharNext(pEnd);
            }
            
            if (*pEnd == 0)
            {
                bFinished = TRUE;
            }
            else
            {
                *pEnd = 0;
            }

            if (lstrlen(pszCmdLine) != 0)
            {
                SHELLEXECUTEINFO ei = {0};

                ei.cbSize          = sizeof(ei);
                ei.lpFile          = pszCmdLine;
                ei.lpDirectory     = szWindowsDir;
                ei.nShow           = nCmdShow;

                if (!ShellExecuteEx(&ei))
                {
                    ShellMessageBox(g_hinstCabinet,
                                    NULL,
                                    MAKEINTRESOURCE(IDS_WINININORUN),
                                    MAKEINTRESOURCE(IDS_DESKTOP),
                                    MB_OK | MB_ICONEXCLAMATION | MB_SYSTEMMODAL,
                                    pszCmdLine);
                }
                else
                {
                    bRet = TRUE;
                }
            }
            
            pszCmdLine = pEnd + 1;
        }
    }
    return bRet;
}


// we check for the old "load=" and "run=" from the [Windows] section of the win.ini, which
// is mapped nowadays to HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows
BOOL _ProcessOldRunAndLoadEquals()
{
    BOOL bRet = FALSE;

    // don't do the run= section if are restricted or we are in safemode
    if (!SHRestricted(REST_NOCURRENTUSERRUN) && !g_fCleanBoot)
    {
        HKEY hk;

        if (RegOpenKeyEx(HKEY_CURRENT_USER,
                         TEXT("Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows"),
                         0,
                         KEY_QUERY_VALUE,
                         &hk) == ERROR_SUCCESS)
        {
            DWORD dwType;
            DWORD cbData;
            TCHAR szBuffer[255];    // max size of load= & run= lines...
            
            // "Load" apps before "Run"ning any.
            cbData = sizeof(szBuffer);
            if ((SHGetValue(hk, NULL, TEXT("Load"), &dwType, (void*)szBuffer, &cbData) == ERROR_SUCCESS) &&
                (dwType == REG_SZ))
            {
                // we want load= to be hidden, so SW_SHOWMINNOACTIVE is needed
                if (ExecuteOldEqualsLine(szBuffer, SW_SHOWMINNOACTIVE))
                {
                    bRet = TRUE;
                }
            }

            cbData = sizeof(szBuffer);
            if ((SHGetValue(hk, NULL, TEXT("Run"), &dwType, (void*)szBuffer, &cbData) == ERROR_SUCCESS) &&
                (dwType == REG_SZ))
            {
                if (ExecuteOldEqualsLine(szBuffer, SW_SHOWNORMAL))
                {
                    bRet = TRUE;
                }
            }

            RegCloseKey(hk);
        }
    }

    return bRet;
}


//---------------------------------------------------------------------------
// Use IERnonce.dll to process RunOnceEx key
//
typedef void (WINAPI *RUNONCEEXPROCESS)(HWND, HINSTANCE, LPSTR, int);

BOOL _ProcessRunOnceEx()
{
    BOOL bRet = FALSE;

    if (SHKeyHasSubkeys(HKEY_LOCAL_MACHINE, REGSTR_PATH_RUNONCEEX))
    {
        PROCESS_INFORMATION pi = {0};
        TCHAR szArgString[MAX_PATH];
        TCHAR szRunDll32[MAX_PATH];
        BOOL fInTSInstallMode = FALSE;

        // See if we are in "Applications Server" mode, if so we need to trigger install mode
        if (IsOS(OS_TERMINALSERVER)) 
        {
            fInTSInstallMode = SHSetTermsrvAppInstallMode(TRUE); 
        }

        // we used to call LoadLibrary("IERNONCE.DLL") and do all of the processing in-proc. Since 
        // ierunonce.dll in turn calls LoadLibrary on whatever is in the registry and those setup dll's
        // can leak handles, we do this all out-of-proc now.

        GetSystemDirectory(szArgString, ARRAYSIZE(szArgString));
        PathAppend(szArgString, TEXT("iernonce.dll"));
        PathQuoteSpaces(szArgString);
        if (SUCCEEDED(StringCchCat(szArgString, ARRAYSIZE(szArgString), TEXT(",RunOnceExProcess"))))
        {
            GetSystemDirectory(szRunDll32, ARRAYSIZE(szRunDll32));
            PathAppend(szRunDll32, TEXT("rundll32.exe"));

            if (CreateProcessWithArgs(szRunDll32, szArgString, NULL, &pi))
            {
                SHProcessMessagesUntilEvent(NULL, pi.hProcess, INFINITE);

                CloseHandle(pi.hProcess);
                CloseHandle(pi.hThread);

                bRet = TRUE;
            }
        }

        if (fInTSInstallMode)
        {
            SHSetTermsrvAppInstallMode(FALSE);
        } 
    }

#ifdef _WIN64
    //
    // check and see if we need to do 32-bit RunOnceEx processing for wow64
    //
    if (SHKeyHasSubkeys(HKEY_LOCAL_MACHINE, TEXT("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnceEx")))
    {
        TCHAR szWow64Path[MAX_PATH];

        if (ExpandEnvironmentStrings(TEXT("%SystemRoot%\\SysWOW64"), szWow64Path, ARRAYSIZE(szWow64Path)))
        {
            TCHAR sz32BitRunOnce[MAX_PATH];
            PROCESS_INFORMATION pi = {0};

            if (SUCCEEDED(StringCchPrintf(sz32BitRunOnce, ARRAYSIZE(sz32BitRunOnce), TEXT("%s\\runonce.exe"), szWow64Path)))
            {
                if (CreateProcessWithArgs(sz32BitRunOnce, TEXT("/RunOnceEx6432"), szWow64Path, &pi))
                {
                    // have to wait for the ruonceex processing before we can return
                    SHProcessMessagesUntilEvent(NULL, pi.hProcess, INFINITE);
                    CloseHandle(pi.hProcess);
                    CloseHandle(pi.hThread);

                    bRet = TRUE;
                }
            }
        }
    }
#endif // _WIN64

    return bRet;
}


BOOL _ProcessRunOnce()
{
    BOOL bRet = FALSE;

    if (!SHRestricted(REST_NOLOCALMACHINERUNONCE))
    {
        bRet = Cabinet_EnumRegApps(HKEY_LOCAL_MACHINE, REGSTR_PATH_RUNONCE, RRA_DELETE | RRA_WAIT, ExecuteRegAppEnumProc, 0);

#ifdef _WIN64
        //
        // check and see if we need to do 32-bit RunOnce processing for wow64
        //
        // NOTE: we do not support per-user (HKCU) 6432 runonce
        //
        if (SHKeyHasValues(HKEY_LOCAL_MACHINE, TEXT("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnce")))
        {
            TCHAR szWow64Path[MAX_PATH];

            if (ExpandEnvironmentStrings(TEXT("%SystemRoot%\\SysWOW64"), szWow64Path, ARRAYSIZE(szWow64Path)))
            {
                TCHAR sz32BitRunOnce[MAX_PATH];
                PROCESS_INFORMATION pi = {0};

                if (SUCCEEDED(StringCchPrintf(sz32BitRunOnce, ARRAYSIZE(sz32BitRunOnce), TEXT("%s\\runonce.exe"), szWow64Path)))
                {
                    // NOTE: since the 32-bit and 64-bit registries are different, we don't wait since it should not affect us
                    if (CreateProcessWithArgs(sz32BitRunOnce, TEXT("/RunOnce6432"), szWow64Path, &pi))
                    {
                        CloseHandle(pi.hProcess);
                        CloseHandle(pi.hThread);

                        bRet = TRUE;
                    }
                }
            }
        }
#endif // _WIN64
    }

    return bRet;
}

void _AutoRunTaskMan()
{
    WCHAR szBuffer[260];
    DWORD cbBuffer = sizeof(szBuffer);
    if (SHGetValueW(
        HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon", L"Taskman", nullptr,
        szBuffer, &cbBuffer) == ERROR_SUCCESS && szBuffer[0])
    {
        PROCESS_INFORMATION pi;

        STARTUPINFOW startup = {};
        startup.wShowWindow = SW_SHOWNORMAL;
        startup.cb = sizeof(startup);
        if (CreateProcessW(nullptr, szBuffer, nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startup, &pi))
        {
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
        }
    }
}

// try to create this by sending a wm_command directly to
// the desktop.
BOOL MyCreateFromDesktop(HINSTANCE hInst, LPCTSTR pszCmdLine, int nCmdShow)
{
    NEWFOLDERINFO fi = {0};
    BOOL bRet = FALSE;

    fi.nShow = nCmdShow;

    //  since we have browseui fill out the fi, 
    //  SHExplorerParseCmdLine() does a GetCommandLine()
    if (SHExplorerParseCmdLine(nullptr, &fi))
    {
        bRet = SHCreateFromDesktop(&fi);
    }

    //  should we also have it cleanup after itself??

    //  SHExplorerParseCmdLine() can allocate this buffer...
    if (fi.uFlags & COF_PARSEPATH)
    {
        LocalFree(fi.pszPath);
    }

    ILFree(fi.pidl);
    ILFree(fi.pidlRoot);

    return bRet;
}

BOOL g_fDragFullWindows=FALSE;
int g_cxEdge=0;
int g_cyEdge=0;
int g_cxPaddedBorder=0;
int g_cySize=0;
int g_cxTabSpace=0;
int g_cyTabSpace=0;
int g_cxBorder=0;
int g_cyBorder=0;
int g_cxPrimaryDisplay=0;
int g_cyPrimaryDisplay=0;
int g_cxDlgFrame=0;
int g_cyDlgFrame=0;
int g_cxFrame=0;
int g_cyFrame=0;

int g_cxMinimized=0;
//int g_fCleanBoot=0;
int g_cxVScroll=0;
int g_cyHScroll=0;
UINT g_uDoubleClick=0;

void Cabinet_InitGlobalMetrics(WPARAM wParam, LPTSTR lpszSection)
{
    BOOL fForce = (!lpszSection || !*lpszSection);

    if (fForce || wParam == SPI_SETDRAGFULLWINDOWS)
    {
        SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &g_fDragFullWindows, 0);
    }

    if (fForce || !lstrcmpi(lpszSection, TEXT("WindowMetrics")) ||
        wParam == SPI_SETNONCLIENTMETRICS)
    {
        g_cxEdge = GetSystemMetrics(SM_CXEDGE);
        g_cyEdge = GetSystemMetrics(SM_CYEDGE);
        g_cxPaddedBorder = GetSystemMetrics(SM_CXPADDEDBORDER);
        g_cxTabSpace = (g_cxEdge * 3) / 2;
        g_cyTabSpace = (g_cyEdge * 3) / 2; // cause the graphic designers really really want 3.
        g_cySize = GetSystemMetrics(SM_CYSIZE);
        g_cxBorder = GetSystemMetrics(SM_CXBORDER);
        g_cyBorder = GetSystemMetrics(SM_CYBORDER);
        g_cxVScroll = GetSystemMetrics(SM_CXVSCROLL);
        g_cyHScroll = GetSystemMetrics(SM_CYHSCROLL);
        g_cxDlgFrame = GetSystemMetrics(SM_CXDLGFRAME);
        g_cyDlgFrame = GetSystemMetrics(SM_CYDLGFRAME);
        g_cxFrame  = GetSystemMetrics(SM_CXFRAME);
        g_cyFrame  = GetSystemMetrics(SM_CYFRAME);
        g_cxMinimized = GetSystemMetrics(SM_CXMINIMIZED);
        g_cxPrimaryDisplay = GetSystemMetrics(SM_CXSCREEN);
        g_cyPrimaryDisplay = GetSystemMetrics(SM_CYSCREEN);
    }

    if (fForce || wParam == SPI_SETDOUBLECLICKTIME)
    {
        g_uDoubleClick = GetDoubleClickTime();
    }
}

BOOL IsDesktopWindowAlreadyPresent()
{
    return FindWindowW(L"Progman", nullptr) || FindWindowW(L"Proxy Desktop", nullptr);
}

BOOL ExplorerIsShell()
{
    WCHAR szModulePath[260];
    GetModuleFileNameW(nullptr, szModulePath, ARRAYSIZE(szModulePath));
    const WCHAR* pszModuleName = PathFindFileNameW(szModulePath);

    WCHAR szPath[260];
    GetPrivateProfileStringW(L"boot", L"shell", pszModuleName, szPath, ARRAYSIZE(szPath), L"system.ini");
    PathRemoveArgsW(szPath);
    PathRemoveBlanksW(szPath);

    const WCHAR* pszPathName = PathFindFileNameW(szPath);
    return StrCmpNIW(pszPathName, pszModuleName, lstrlenW(pszModuleName)) == 0 || StrCmpICW(pszPathName, L"install.exe") == 0;
}

DEFINE_GUID(CLSID_ExplorerHostCreator, 0xAB0B37EC, 0x56F6, 0x4A0E, 0xA8, 0xFD, 0x7A, 0x8B, 0xF7, 0xC2, 0xDA, 0x96);

MIDL_INTERFACE("c4de032a-d902-450a-bc43-d9df6d0fd48c")
IExplorerHostCreator : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE CreateHost(REFCLSID) = 0;
    virtual HRESULT STDMETHODCALLTYPE RunHost() = 0;
};

DEFINE_GUID(CLSID_DesktopExplorerHost, 0x682159D9, 0xC321, 0x47CA, 0xB3, 0xF1, 0x30, 0xE3, 0x6B, 0x2E, 0xC8, 0xB9);
DEFINE_GUID(CLSID_SeparateMultipleProcessExplorerHost, 0x75DFF2B7, 0x6936, 0x4C06, 0xA8, 0xBB, 0x67, 0x6A, 0x7B, 0x00, 0xB2, 0x4B);

enum LAUNCHEXPLORERFLAGS : int
{
    LE_NONE = 0x0000,
    LE_NEWWINDOW = 0x0001,
    LE_NEWPROCESS = 0x0002,
    LE_SELECTITEM = 0x0004,
    LE_EXPLORE = 0x0008,
    LE_EXPAND = 0x0010,
    LE_INCURRENTTABGROUP = 0x0020,
};

struct IBrowserThreadHandshake;

// This changed sometime after 14361
MIDL_INTERFACE("9b25c299-03b6-4a14-827d-095485d0c022")
IExplorerLauncher : IUnknown
{
    virtual HRESULT ShowWindow(
        REFCLSID, LPCITEMIDLIST, LAUNCHEXPLORERFLAGS, POINT, int, HWND, IUnknown*, IBrowserThreadHandshake*) = 0;
    virtual HRESULT ShowWindowUsingDefaultPolicyAtRect(LPCITEMIDLIST, LAUNCHEXPLORERFLAGS, const RECT*, int) = 0;
};

DEFINE_GUID(CLSID_ExplorerLauncher, 0x1F849CCE, 0x2546, 0x4B9F, 0xB0, 0x3E, 0x40, 0x04, 0x78, 0x1B, 0xDC, 0x40);

// @Note: This is from build 6519/6593, there should a few more entries.
enum EXPLORERSERVER
{
    EXPLORERSERVER_UNDETERMINED = 0x0,
    EXPLORERSERVER_SEPARATE = 0x1,
    EXPLORERSERVER_DESKTOP = 0x2,
    EXPLORERSERVER_FACTORY = 0x3,
    EXPLORERSERVER_NONE = 0x4,
};

#ifdef DEAD_CODE
BOOL ShouldStartDesktopAndTray()
{
    return !IsDesktopWindowAlreadyPresent() && ExplorerIsShell();
}
#else
EXPLORERSERVER ShouldStartDesktopAndTray()
{
    EXPLORERSERVER es;

    const WCHAR* CommandLineW = GetCommandLineW();
    const WCHAR* pszCmdLine = PathGetArgsW(CommandLineW);
    if (pszCmdLine && *pszCmdLine)
    {
        NEWFOLDERINFO fi = {};
        es = EXPLORERSERVER_SEPARATE;
        if (SHExplorerParseCmdLine(pszCmdLine, &fi))
        {
            if (!IsEqualCLSID(fi.clsid, CLSID_NULL))
            {
                es = EXPLORERSERVER_FACTORY;
            }
            ILFree(fi.pidl);
        }
    }
    else if (ExplorerIsShell())
    {
        return IsDesktopWindowAlreadyPresent() != 0 ? EXPLORERSERVER_NONE : EXPLORERSERVER_UNDETERMINED;
    }
    else
    {
        return EXPLORERSERVER_SEPARATE;
    }
    return es;
}
#endif

void DisplayCleanBootMsg()
{
    // On server sku's or anytime on ia64, just show a message with
    // an OK button for safe boot
    UINT uiMessageBoxFlags = MB_ICONEXCLAMATION | MB_SYSTEMMODAL | MB_OK;
    UINT uiMessage = IDS_CLEANBOOTMSG;

#ifndef _WIN64
    if (!IsOS(OS_ANYSERVER))
    {
        // On x86 per and pro, also offer an option to start system restore
        uiMessageBoxFlags = MB_ICONEXCLAMATION | MB_SYSTEMMODAL | MB_YESNO;
        uiMessage = IDS_CLEANBOOTMSGRESTORE;
    }
#endif // !_WIN64

    WCHAR szTitle[80];
    WCHAR szMessage[1024];

    LoadString(g_hinstCabinet, IDS_DESKTOP, szTitle, ARRAYSIZE(szTitle));
    LoadString(g_hinstCabinet, uiMessage, szMessage, ARRAYSIZE(szMessage));

    // on IA64 the msgbox will always return IDOK, so this "if" will always fail.
    if (IDNO == MessageBox(NULL, szMessage, szTitle, uiMessageBoxFlags))
    {
        TCHAR szPath[MAX_PATH];
        ExpandEnvironmentStrings(TEXT("%SystemRoot%\\system32\\restore\\rstrui.exe"), szPath, ARRAYSIZE(szPath));
        PROCESS_INFORMATION pi;
        STARTUPINFO si = {0};
        if (CreateProcess(szPath, NULL, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
        {
            CloseHandle(pi.hThread);
            CloseHandle(pi.hProcess);
        }
    }
}

BOOL IsExecCmd(LPCTSTR pszCmd)
{
    return *pszCmd && !StrStrI(pszCmd, TEXT("-embedding"));
}

// run the cmd line passed up from win.com

void _RunWinComCmdLine(LPCTSTR pszCmdLine, UINT nCmdShow)
{
    if (IsExecCmd(pszCmdLine))
    {
        SHELLEXECUTEINFO ei = { sizeof(ei), 0, NULL, NULL, pszCmdLine, NULL, NULL, (int)nCmdShow};

        ei.lpParameters = PathGetArgs(pszCmdLine);
        if (*ei.lpParameters)
            *((LPTSTR)ei.lpParameters - 1) = 0;     // const -> non const

        ShellExecuteEx(&ei);
    }
}

// stolen from the CRT, used to shirink our code
LPTSTR _SkipCmdLineCrap(LPTSTR pszCmdLine)
{
    if (*pszCmdLine == TEXT('\"'))
    {
        //
        // Scan, and skip over, subsequent characters until
        // another double-quote or a null is encountered.
        //
        while (*++pszCmdLine && (*pszCmdLine != TEXT('\"')))
            ;

        //
        // If we stopped on a double-quote (usual case), skip
        // over it.
        //
        if (*pszCmdLine == TEXT('\"'))
            pszCmdLine++;
    }
    else
    {
        while (*pszCmdLine > TEXT(' '))
            pszCmdLine++;
    }

    //
    // Skip past any white space preceeding the second token.
    //
    while (*pszCmdLine && (*pszCmdLine <= TEXT(' ')))
        pszCmdLine++;

    return pszCmdLine;
}

#ifdef EXEX_DLL
EXTERN_C BOOL WINAPI _CRT_INIT(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved);
STDAPI_(int) ModuleEntry(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved)
{
    if (fdwReason == DLL_PROCESS_ATTACH)
    {
        _CRT_INIT(hinstDLL, DLL_PROCESS_ATTACH, NULL);

        if (!SHUndocInit())
            return -1;

        InitDesktopFuncs();

        // We don't want the "No disk in drive X:" requesters, so we set
        // the critical error mask such that calls will just silently fail

        SetErrorMode(SEM_FAILCRITICALERRORS);

        LPTSTR pszCmdLine = GetCommandLine();
        pszCmdLine = _SkipCmdLineCrap(pszCmdLine);

        STARTUPINFO si = { 0 };
        si.cb = sizeof(si);
        GetStartupInfo(&si);

        int nCmdShow = si.dwFlags & STARTF_USESHOWWINDOW ? si.wShowWindow : SW_SHOWDEFAULT;
        int iRet = ExplorerWinMain(hinstDLL, NULL, pszCmdLine, nCmdShow);

        return iRet;
    }
    else
    {
        return _CRT_INIT(hinstDLL, fdwReason, lpReserved);
    }

    return 0;
}
#else
EXTERN_C BOOL WINAPI _CRT_INIT(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved);
STDAPI_(int) ModuleEntry()
{
    PERFSETMARK("ExplorerStartup");

    _CRT_INIT(GetModuleHandle(0), DLL_PROCESS_ATTACH, NULL);
    //DoInitialization();

    if (!SHUndocInit())
        return -1;

    InitDesktopFuncs();



    // We don't want the "No disk in drive X:" requesters, so we set
    // the critical error mask such that calls will just silently fail

    SetErrorMode(SEM_FAILCRITICALERRORS);

    LPTSTR pszCmdLine = GetCommandLine();
    pszCmdLine = _SkipCmdLineCrap(pszCmdLine);

    STARTUPINFO si = { 0 };
    si.cb = sizeof(si);
    GetStartupInfo(&si);

    int nCmdShow = si.dwFlags & STARTF_USESHOWWINDOW ? si.wShowWindow : SW_SHOWDEFAULT;
    int iRet = ExplorerWinMain(GetModuleHandle(NULL), NULL, pszCmdLine, nCmdShow);

    _CRT_INIT(GetModuleHandle(0), DLL_PROCESS_DETACH, NULL);
    //DoCleanup();

    // Since we now have a way for an extension to tell us when it is finished,
    // we will terminate all processes when the main thread goes away.

    if (g_fExitExplorer)    // desktop told us not to exit
        ExitProcess(iRet);

    return iRet;
}
#endif

#ifdef EXEX_DLL
extern "C" __declspec(dllexport)
#endif
HANDLE CreateDesktopAndTray()
{
    HANDLE hDesktop = nullptr;

    if (v_hwndTray || c_tray.Init())
    {
        ASSERT(v_hwndTray);

        if (!v_hwndDesktop)
        {
            // cache the handle to the desktop...
            hDesktop = SHCreateDesktop(c_tray.GetDeskTray());
#ifdef EXEX_DLL
            if (hDesktop)
                PostMessage(v_hwndTray, TM_SHOWTRAYBALLOON, TRUE, 0);
#endif
        }
    }

    return hDesktop;
}
using fnSHSessionKey = HRESULT(WINAPI*)(REGSAM, HKEY*);
fnSHSessionKey SHCreateSessionKey;


// Removes the session key from the registry.
void NukeSessionKey(void)
{
    HKEY hkDummy;

    if (!SHCreateSessionKey)
        SHCreateSessionKey = reinterpret_cast<fnSHSessionKey>(GetProcAddress(GetModuleHandle(L"shell32.dll"), MAKEINTRESOURCEA(723)));

    SHCreateSessionKey(0xFFFFFFFF, &hkDummy);
}

BOOL IsFirstInstanceAfterLogon()
{
    BOOL fResult = FALSE;

    if (!SHCreateSessionKey)
        SHCreateSessionKey = reinterpret_cast<fnSHSessionKey>(GetProcAddress(GetModuleHandle(L"shell32.dll"), MAKEINTRESOURCEA(723)));

    HKEY hkSession;
    HRESULT hr = SHCreateSessionKey(KEY_WRITE, &hkSession);
    if (SUCCEEDED(hr))
    {
        HKEY hkStartup;
        DWORD dwDisposition;
        LONG lRes;
        lRes = RegCreateKeyEx(hkSession, TEXT("StartupHasBeenRun"), 0,
                       NULL,
                       REG_OPTION_VOLATILE,
                       KEY_WRITE,
                       NULL,
                       &hkStartup,
                       &dwDisposition);
        if (lRes == ERROR_SUCCESS)
        {
            RegCloseKey(hkStartup);
            if (dwDisposition == REG_CREATED_NEW_KEY)
                fResult = TRUE;
        }
        RegCloseKey(hkSession);
    }
    return fResult;
}

DWORD ReadFaultCount()
{
    DWORD dwValue = 0;
    DWORD dwSize = sizeof(dwValue);

    RegQueryValueEx(g_hkeyExplorer, TEXT("FaultCount"), NULL, NULL, (LPBYTE)&dwValue, &dwSize);
    return dwValue;
}

void WriteFaultCount(DWORD dwValue)
{
    RegSetValueEx(g_hkeyExplorer, TEXT("FaultCount"), 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(dwValue));
    // If we are clearing the fault count or this is the first fault, clear or set the fault time.
    if (!dwValue || (dwValue == 1))
    {
        if (dwValue)
            dwValue = GetTickCount();
        RegSetValueEx(g_hkeyExplorer, TEXT("FaultTime"), 0, REG_DWORD, (LPBYTE)&dwValue, sizeof(dwValue));
    }
}

// This function assumes it is only called when a fault has occured previously...
BOOL ShouldDisplaySafeMode()
{
    BOOL fRet = FALSE;
    SHELLSTATE ss;

    SHGetSetSettings(&ss, SSF_DESKTOPHTML, FALSE);

    if (ss.fDesktopHTML)
    {
        if (ReadFaultCount() >= MAGIC_FAULT_LIMIT)
        {
            DWORD dwValue = 0;
            DWORD dwSize = sizeof(dwValue);

            RegQueryValueEx(g_hkeyExplorer, TEXT("FaultTime"), NULL, NULL, (LPBYTE)&dwValue, &dwSize);
            fRet = ((GetTickCount() - dwValue) < MAGIC_FAULT_TIME);
            // We had enough faults but they weren't over a sufficiently short period of time.  Reset the fault
            // count to 1 so that we start counting from this fault now.
            if (!fRet)
                WriteFaultCount(1);
        }
    }
    else
    {
        // We don't care about faults that occured if AD is off.
        WriteFaultCount(0);
    }
    
    return fRet;
}

void WriteCleanShutdown(DWORD dwValue)
{
    RegSetValueExW(
        g_hkeyExplorer, L"CleanShutdown", 0, REG_DWORD, reinterpret_cast<const BYTE*>(dwValue), sizeof(dwValue));

    if (dwValue && !g_fFakeShutdown)
    {
        // UpdateSqmFolderSettings();
    }
}

BOOL WasPrevShutdownClean()
{
    DWORD dwValue = 1; // default: it was clean
    DWORD dwSize = sizeof(dwValue);

    RegQueryValueEx(g_hkeyExplorer, TEXT("CleanShutdown"), NULL, NULL, (LPBYTE)&dwValue, &dwSize);
    return (BOOL)dwValue;
}


using fnSHGetAllAccessSA = LPSECURITY_ATTRIBUTES(WINAPI*)();
fnSHGetAllAccessSA SHGetAllAccessSA;

//
//  Synopsis:   Waits for the OLE SCM process to finish its initialization.
//              This is called before the first call to OleInitialize since
//              the SHELL runs early in the boot process.
//
//  Arguments:  None.
//
//  Returns:    S_OK - SCM is running. OK to call OleInitialize.
//              CO_E_INIT_SCM_EXEC_FAILURE - timed out waiting for SCM
//              other - create event failed
//
//  History:    26-Oct-95   Rickhi  Extracted from CheckAndStartSCM so
//                                  that only the SHELL need call it.
//
HRESULT WaitForSCMToInitialize()
{
    static BOOL s_fScmStarted = FALSE;

    if (s_fScmStarted)
    {
        return S_OK;
    }

    if (!SHGetAllAccessSA)
        SHGetAllAccessSA = reinterpret_cast<fnSHGetAllAccessSA>(GetProcAddress(GetModuleHandle(L"shlwapi.dll"), MAKEINTRESOURCEA(563)));


    //SECURITY_ATTRIBUTES* psa = SHGetAllAccessSA();
    SECURITY_ATTRIBUTES* psa = nullptr;

    // on NT5 we need a global event that is shared between TS sessions
    HANDLE hEvent = CreateEvent(psa, TRUE, FALSE, SZ_SCMCREATEDEVENT_NT5);

    if (!hEvent && GetLastError() == ERROR_ACCESS_DENIED)
    {
        //
        // Win2K OLE32 has tightened security such that if this object
        // already exists, we aren't allowed to open it with EVENT_ALL_ACCESS
        // (CreateEvent fails with ERROR_ACCESS_DENIED in this case).
        // Fall back by calling OpenEvent requesting SYNCHRONIZE access.
        //
        hEvent = OpenEvent(SYNCHRONIZE, FALSE, SZ_SCMCREATEDEVENT_NT5);
    }
    
    if (hEvent)
    {
        // wait for the SCM to signal the event, then close the handle
        // and return a code based on the WaitEvent result.
        int rc = WaitForSingleObject(hEvent, 60000);

        CloseHandle(hEvent);

        if (rc == WAIT_OBJECT_0)
        {
            s_fScmStarted = TRUE;
            return S_OK;
        }
        else if (rc == WAIT_TIMEOUT)
        {
            return CO_E_INIT_SCM_EXEC_FAILURE;
        }
    }
    return HRESULT_FROM_WIN32(GetLastError());  // event creation failed or WFSO failed.
}

HRESULT WINAPI _OleCoInitializeWaitForSCM()
{
    WaitForSCMToInitialize();
    HRESULT hr = SHCoInitialize();
    OleInitialize(nullptr);
    return hr;
}

void WINAPI _OleCoUninitialize(HRESULT hr)
{
    OleUninitialize();
    SHCoUninitialize(hr);
}

// the following locale fixes (for NT5 378948) are dependent on desk.cpl changes
// Since Millennium does not ship updated desk.cpl, we don't want to do this on Millennium
//
//  Given the Locale ID, this returns the corresponding charset
//
UINT  GetCharsetFromLCID(LCID   lcid)
{
    TCHAR szData[6+1]; // 6 chars are max allowed for this lctype
    UINT uiRet;
    if (GetLocaleInfo(lcid, LOCALE_IDEFAULTANSICODEPAGE, szData, ARRAYSIZE(szData)) > 0)
    {
        UINT uiCp = (UINT)StrToInt(szData);
        CHARSETINFO csinfo;

        TranslateCharsetInfo(IntToPtr_(DWORD *, uiCp), &csinfo, TCI_SRCCODEPAGE);
        uiRet = csinfo.ciCharset;
    }
    else
    {
        // at worst non penalty for charset
        uiRet = DEFAULT_CHARSET;
    }

    return uiRet;
}

void CheckDefaultUIFonts()
{
    LANGID langID;
    DWORD dwSize = sizeof(langID);
    if (SHGetValueW(HKEY_CURRENT_USER, L"Control Panel\\Appearance", L"SchemeLangID", nullptr, &langID, &dwSize)
        || langID != GetUserDefaultUILanguage())
    {
        HINSTANCE hInst = LoadLibraryW(L"desk.cpl");
        if (hInst)
        {
            using PFN = HRESULT(*WINAPI)();

            PFN pfnUpdateCharsetChanges = reinterpret_cast<PFN>(GetProcAddress(hInst, "UpdateCharsetChanges"));
            if (pfnUpdateCharsetChanges)
            {
                pfnUpdateCharsetChanges();
            }
            FreeLibrary(hInst);
        }
    }
}

void ChangeUIfontsToNewDPI()
{
    HDC hdc = GetDC(nullptr);
    int iNewDPI = GetDeviceCaps(hdc, LOGPIXELSY);
    ReleaseDC(nullptr, hdc);

    int iOldDPI;
    DWORD dwSize = sizeof(iOldDPI);
    if (SHGetValueW(
        HKEY_CURRENT_USER, L"Control Panel\\Desktop\\WindowMetrics", L"AppliedDPI", nullptr, &iOldDPI, &dwSize))
    {
        dwSize = sizeof(iOldDPI);
        if (SHGetValueW(
            HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Control Panel", L"OriginalDPI",
            nullptr, &iOldDPI, &dwSize))
        {
            iOldDPI = iNewDPI;
        }
    }
    if (iNewDPI != iOldDPI)
    {
        HINSTANCE hInst = LoadLibraryW(L"desk.cpl");
        if (hInst)
        {
            using PFN = HRESULT(*)(int, int);

            PFN pfnUpdateUIfonts = reinterpret_cast<PFN>(GetProcAddress(hInst, "UpdateUIfontsDueToDPIchange"));
            if (pfnUpdateUIfonts)
            {
                pfnUpdateUIfonts(iOldDPI, iNewDPI);
            }
            FreeLibrary(hInst);
        }
    }
}

CComModule _Module;
BEGIN_OBJECT_MAP(ObjectMap)
// add your OBJECT_ENTRY's here
END_OBJECT_MAP()

typedef BOOL (*PFNICOMCTL32)(LPINITCOMMONCONTROLSEX);
void _InitComctl32()
{
    HMODULE hmod = LoadLibrary(TEXT("comctl32.dll"));
    if (hmod)
    {
        PFNICOMCTL32 pfn = (PFNICOMCTL32)GetProcAddress(hmod, "InitCommonControlsEx");
        if (pfn)
        {
            INITCOMMONCONTROLSEX icce;
            icce.dwICC = 0x00003FFF;
            icce.dwSize = sizeof(icce);
            pfn(&icce);
        }
    }
}


BOOL _ShouldFixResolution(void)
{
    BOOL fRet = FALSE;
#ifndef _WIN64  // This feature is not supported on 64-bit machine

    DISPLAY_DEVICE dd;
    ZeroMemory(&dd, sizeof(DISPLAY_DEVICE));
    dd.cb = sizeof(DISPLAY_DEVICE);

    if (SHRegGetBoolUSValue(REGSTR_PATH_EXPLORER TEXT("\\DontShowMeThisDialogAgain"), TEXT("ScreenCheck"), FALSE, TRUE))
    {
        // Don't fix SafeMode or Terminal Clients
        if ((GetSystemMetrics(SM_CLEANBOOT) == 0) && (GetSystemMetrics(SM_REMOTESESSION) == FALSE))
        {
            fRet = TRUE;
            for (DWORD dwMon = 0; EnumDisplayDevices(NULL, dwMon, &dd, 0); dwMon++)
            {
                if (!(dd.StateFlags & DISPLAY_DEVICE_MIRRORING_DRIVER))
                {
                    DEVMODE dm = {0};
                    dm.dmSize = sizeof(DEVMODE);

                    if (EnumDisplaySettingsEx(dd.DeviceName, ENUM_CURRENT_SETTINGS, &dm, 0))
                    {
                        if ((dm.dmFields & DM_POSITION) &&
                            ((dm.dmPelsWidth >= 600) &&
                                (dm.dmPelsHeight >= 600) &&
                                (dm.dmBitsPerPel >= 15)))
                        {
                            fRet = FALSE;
                        }
                    }
                }
            }
        }
    }

#endif // _WIN64
    return fRet;
}


BOOL _ShouldOfferTour(void)
{
    BOOL fRet = FALSE;
    
#ifndef _WIN64  // This feature is not supported on 64-bit machine

    // we don't allow guest to get offered tour b/c guest's registry is wiped every time she logs out, 
    //   so she would get tour offered every single she logged in.
    if (!IsOS(OS_ANYSERVER) && !IsOS(OS_EMBEDDED) && !(SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_GUESTS)))
    {
        DWORD dwCount;
        DWORD cbCount = sizeof(DWORD);
        
        // we assume if we can't read the RunCount it's because it's not there (we haven't tried to offer the tour yet), so we default to 3.
        if (ERROR_SUCCESS != SHRegGetUSValue(REGSTR_PATH_SETUP TEXT("\\Applets\\Tour"), TEXT("RunCount"), NULL, &dwCount, &cbCount, FALSE, NULL, 0))
        {
            dwCount = 3;
        }
        
        if (dwCount)
        {
            HUSKEY hkey1;
            if (ERROR_SUCCESS == SHRegCreateUSKey(REGSTR_PATH_SETUP TEXT("\\Applets"), KEY_WRITE, NULL, &hkey1, SHREGSET_HKCU))
            {
                HUSKEY hkey2;
                if (ERROR_SUCCESS == SHRegCreateUSKey(TEXT("Tour"), KEY_WRITE, hkey1, &hkey2, SHREGSET_HKCU))
                {
                    if (ERROR_SUCCESS == SHRegWriteUSValue(hkey2, TEXT("RunCount"), REG_DWORD, &(--dwCount), cbCount, SHREGSET_FORCE_HKCU))
                    {
                        fRet = TRUE;
                    }
                    SHRegCloseUSKey(hkey2);
                }
                SHRegCloseUSKey(hkey1);
            }
        }
    }

#endif // _WIN64
    return fRet;
}

typedef BOOL (*CHECKFUNCTION)(void);

void _ConditionalBalloonLaunch(CHECKFUNCTION pCheckFct, SHELLREMINDER* psr)
{
	if (pCheckFct())
	{
		IShellReminderManager* psrm;
		HRESULT hr = CoCreateInstance(CLSID_PostBootReminder, NULL, CLSCTX_INPROC_SERVER,
			IID_PPV_ARG(IShellReminderManager, &psrm));

		if (SUCCEEDED(hr))
		{
			psrm->Add(psr);
			psrm->Release();
		}
	}
}

void _FixWordMailRegKey(void)
{
    // If we don't have permissions, fine this is just correction code
    HKEY hkey;
    if (ERROR_SUCCESS == RegOpenKeyEx(HKEY_CLASSES_ROOT, L"Applications", 0, KEY_ALL_ACCESS, &hkey))
    {
        HKEY hkeyTemp;
        if (ERROR_SUCCESS != RegOpenKeyEx(hkey, L"WINWORD.EXE", 0, KEY_ALL_ACCESS, &hkeyTemp))
        {
            HKEY hkeyWinWord;
            DWORD dwResult;
            if (ERROR_SUCCESS == RegCreateKeyEx(hkey, L"WINWORD.EXE", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkeyWinWord, &dwResult))
            {
                HKEY hkeyTBExcept;
                if (ERROR_SUCCESS == RegCreateKeyEx(hkeyWinWord, L"TaskbarExceptionsIcons", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkeyTBExcept, &dwResult))
                {
                    HKEY hkeyIcon;
                    if (ERROR_SUCCESS == RegCreateKeyEx(hkeyTBExcept, L"WordMail", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkeyIcon, &dwResult))
                    {
                        const WCHAR szIconPath[] = L"explorer.exe,16";
                        DWORD cbIconPath = sizeof(szIconPath);
                        RegSetValue(hkeyIcon, L"IconPath", REG_SZ, szIconPath, cbIconPath);

                        const WCHAR szNewExeName[] = L"OUTLOOK.EXE";
                        DWORD cbNewExeName = sizeof(szNewExeName);
                        RegSetValue(hkeyIcon, L"NewExeName", REG_SZ, szNewExeName, cbNewExeName);
                        RegCloseKey(hkeyIcon);
                    }
                    RegCloseKey(hkeyTBExcept);
                }
                RegCloseKey(hkeyWinWord);
            }
        }
        else
        {
            RegCloseKey(hkeyTemp);
        }
        RegCloseKey(hkey);
    }
}

void CheckForServerAdminUI()
{
    DWORD dwServerAdminUI;
    DWORD cb = sizeof(dwServerAdminUI);
    DWORD dwErr = SHGetValue(HKEY_CURRENT_USER, REGSTR_PATH_EXPLORER TEXT("\\Advanced"),
                             TEXT("ServerAdminUI"), NULL, &dwServerAdminUI, &cb);
    if (dwErr == ERROR_FILE_NOT_FOUND || dwErr == ERROR_PATH_NOT_FOUND)
    {
        //  Determine whether the user should receive server admin UI or not
        dwServerAdminUI = IsOS(OS_ANYSERVER) &&
          (SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_ADMINS) ||
           SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_SYSTEM_OPS) ||
           SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_BACKUP_OPS) ||
           SHTestTokenMembership(NULL, DOMAIN_ALIAS_RID_NETWORK_CONFIGURATION_OPS));

        // In the server admin case, change some defaults to be more serverish
        if (dwServerAdminUI)
        {
            // Install the Server Admin UI
            typedef HRESULT (CALLBACK *DLLINSTALLPROC)(BOOL, LPCWSTR);
            DLLINSTALLPROC pfnDllInstall = (DLLINSTALLPROC)GetProcAddress(GetModuleHandle(TEXT("SHELL32")), "DllInstall");
            if (pfnDllInstall)
            {
                pfnDllInstall(TRUE, L"SA");
            }

            // Re-enable keyboard underlines.
            SystemParametersInfo(SPI_SETKEYBOARDCUES, 0, IntToPtr(TRUE), SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);

            // Tell everybody to refresh since we changed some settings
            SHSendMessageBroadcast(WM_SETTINGCHANGE, 0, 0);
        }

        SHSetValue(HKEY_CURRENT_USER, REGSTR_PATH_EXPLORER TEXT("\\Advanced"),
                   TEXT("ServerAdminUI"), REG_DWORD, &dwServerAdminUI, sizeof(dwServerAdminUI));
    }
}

// EXEX-Vista(Allison): Validated from Windows 10 Explorer/TwinUI.dll
class CThreadRefHost
{
public:
    CThreadRefHost()
        : _cRef(0)
        , _punk(nullptr)
    {
        SHCreateThreadRef(&_cRef, &_punk);
        SHSetThreadRef(_punk);
    }

    template <typename TLambda>
    CThreadRefHost(TLambda lambda)
        : _cRef(0)
        , _punk(nullptr)
    {
        lambda(&_cRef, &_punk);
        SHSetThreadRef(_punk);
    }

    CThreadRefHost(const CThreadRefHost&) = delete;

    virtual ~CThreadRefHost()
    {
        WaitForRefs();
    }

    void WaitForRefs()
    {
        SHSetThreadRef(nullptr);
        IUnknown_SafeReleaseAndNullPtr(&_punk);

        MSG msg;
        while (_cRef)
        {
            if (GetMessageW(&msg, nullptr, 0, 0))
            {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    }

protected:
    LONG _cRef;
    IUnknown* _punk;
};

class CProcessAndThreadRefHost : public CThreadRefHost
{
public:
    CProcessAndThreadRefHost()
    {
        static void (*fSetProcessReference)(IUnknown* punk) = reinterpret_cast<decltype(fSetProcessReference)>(GetProcAddress(LoadLibraryW(L"SHCore.dll"), "SetProcessReference"));
        if (fSetProcessReference)
        {
            fSetProcessReference(_punk);
        }
    }
};

void StartExplorerWindow()
{
    SHELLEXECUTEINFOW sei = {};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC | SEE_MASK_WAITFORINPUTIDLE;
    sei.lpVerb = L"explore";
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteExW(&sei);
}

#if 0

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPWSTR pszCmdLine, int nCmdShow)
{
#if !defined(_WIN64)
    BOOL bNoHeapTerminationOnCorruption = FALSE;
    SHRegGetBOOLW(
        HKEY_LOCAL_MACHINE, L"Software\\Policies\\Microsoft\\Windows\\Explorer", L"NoHeapTerminationOnCorruption",
        &bNoHeapTerminationOnCorruption);
    if (!bNoHeapTerminationOnCorruption)
    {
        HeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);
    }
#endif

    // WPP_INIT_CONTROL_ARRAY(&WPP_MAIN_CB);
    // WPP_REGISTRATION_GUIDS = WPP_ThisDir_CTLGUID_ShellTraceProvider;
    // WPP_GLOBAL_Control = &WPP_MAIN_CB;
    // WppInitUm(L"Microsoft\\Shell");

    // EventRegister(&Microsoft_Windows_Shell_Core_Provider, nullptr, nullptr, &g_SHPerfRegHandle);
    // Skipped telemetry ShellTraceId_Explorer_InitializingExplorer_Start

    SetErrorMode(SEM_FAILCRITICALERRORS);

    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);

    SHFusionInitializeFromModule(hInstance);
    //CcshellGetDebugFlags();

    g_hinstCabinet = hInstance;

    if (SUCCEEDED(_Module.Init(nullptr, hInstance, nullptr)))
    {
        Cabinet_InitGlobalMetrics(0, nullptr);

        RegCreateKeyW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", &g_hkeyExplorer);
        if (!g_hkeyExplorer)
        {
            // CcshellDebugMsgW(2, "ExplorerWinMain: unable to create reg explorer key");
        }

        HANDLE hMutex = nullptr;

        EXPLORERSERVER es = ShouldStartDesktopAndTray();
        if (es == EXPLORERSERVER_UNDETERMINED)
        {
            hMutex = CreateMutexW(nullptr, 0, L"Local\\ExplorerIsShellMutex");
            if (hMutex)
            {
                WaitForSingleObject(hMutex, INFINITE);
            }
            es = IsDesktopWindowAlreadyPresent() != 0 ? EXPLORERSERVER_NONE : EXPLORERSERVER_DESKTOP;
        }

        if (es == EXPLORERSERVER_DESKTOP)
        {
            ShellDDEInit(TRUE);

            SetProcessShutdownParameters(3, 0);
            _AutoRunTaskMan();

            MSG msg;
            PeekMessageW(&msg, nullptr, WM_QUIT, WM_QUIT, PM_NOREMOVE);

            HRESULT hrInit = _OleCoInitializeWaitForSCM();
            FileIconInit(TRUE);

            CheckDefaultUIFonts();
            ChangeUIfontsToNewDPI();
            CheckForServerAdminUI();
            CheckHighContrast();

            DoRunOnceIfNeeded();
            CreateShellDirectories();

            ProcessInstallUninstallStubsIfNeeded();
            WasPrevShutdownClean();
            WriteCleanShutdown(0);

            WinList_Init();

            CProcessAndThreadRefHost host;

            HANDLE hDesktop;
            if (IsDesktopWindowAlreadyPresent())
                hDesktop = nullptr;
            else
                hDesktop = CreateDesktopAndTray();

            if (hMutex)
            {
                ReleaseMutex(hMutex);
                CloseHandle(hMutex);
            }

            if (hDesktop)
            {
                PostMessageW(v_hwndTray, 0x590, 1, 0);
                //InitSoundWindow();

                SetPriorityClass(GetCurrentProcess(), NORMAL_PRIORITY_CLASS);

                RegisterApplicationRestart(L"", 0x80000008);

                IExplorerHostCreator* pehc = nullptr;
                if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerHostCreator, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pehc))))
                {
                    pehc->CreateHost(CLSID_DesktopExplorerHost);
                }

                // Skipped telemetry Explorer_MessageLoop_Start
                SHDesktopMessageLoop(hDesktop);
                // Skipped telemetry Explorer_MessageLoop_Stop

                host.WaitForRefs();
                IUnknown_SafeReleaseAndNullPtr(&pehc);

                WriteCleanShutdown(1);
            }

            WinList_Terminate();
            _OleCoUninitialize(hrInit);

            ShellDDEInit(FALSE);
        }
        else
        {
            if (hMutex)
            {
                ReleaseMutex(hMutex);
                CloseHandle(hMutex);
            }

            SetPriorityClass(GetCurrentProcess(), NORMAL_PRIORITY_CLASS);

            if (es == EXPLORERSERVER_NONE)
            {
                StartExplorerWindow();
            }
            else
            {
                HRESULT hrInit = SHCoInitialize();
                if (SUCCEEDED(hrInit))
                {
                    NEWFOLDERINFO fi = {};
                    const WCHAR* CommandLineW = GetCommandLineW();
                    const WCHAR* ArgsW = PathGetArgsW(CommandLineW);
                    if (SHExplorerParseCmdLine(ArgsW, &fi))
                    {
                        if (es == EXPLORERSERVER_FACTORY)
                        {
                            IExplorerHostCreator* pehc;
                            if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerHostCreator, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pehc))))
                            {
                                if (SUCCEEDED(pehc->CreateHost(fi.clsid)))
                                {
                                    pehc->RunHost();
                                }
                                pehc->Release();
                            }
                        }
                        else
                        {
                            ASSERT(es == EXPLORERSERVER_SEPARATE);

                            IExplorerLauncher* pel;
                            if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerLauncher, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pel))))
                            {
                                STARTUPINFOW si = { sizeof(si) };
                                GetStartupInfoW(&si);

                                POINT ptZero = {};
                                int wShowWindow;
                                if ((si.dwFlags & 1) != 0)
                                    wShowWindow = si.wShowWindow;
                                else
                                    wShowWindow = SW_SHOWNORMAL;

                                pel->ShowWindow(
                                    CLSID_SeparateMultipleProcessExplorerHost, fi.pidl,
                                    (LAUNCHEXPLORERFLAGS)fi.uFlags, ptZero, wShowWindow, nullptr, nullptr, nullptr);
                                pel->Release();
                            }
                        }
                        ILFree(fi.pidl);
                    }

                    SHCoUninitialize(hrInit);
                }
            }
        }

        _Module.Term();
    }

    const WCHAR* v13 = GetCommandLineW();
    const WCHAR* v14 = PathGetArgsW(v13);
    FreeSharedMemInCmdLine(v14);

    SHFusionUninitialize();
    // _DebugMsgW(4, L"c.App Exit.");
    // c_tray.Cleanup();

    // EventUnregister(g_SHPerfRegHandle);
    // WppCleanupUm();

    if (g_fExitExplorer)
    {
        ExitProcess(1);
    }

    return 1;
}
#endif

int ProcessInstallUninstallStubs()
{
    if (!GetSystemMetrics(SM_CLEANBOOT))
    {
        //RunInstallUninstallStubs(0);
    }
    return 0;
}

void ProcessInstallUninstallStubsIfNeeded()
{
    if (!GetSystemMetrics(SM_CLEANBOOT))
    {
        HANDLE hCanRegister = CreateEventW(nullptr, TRUE, TRUE, L"Local\\_fCanRegisterWithShellService");
        ProcessInstallUninstallStubs();
        if (hCanRegister)
        {
            CloseHandle(hCanRegister);
        }
    }
}

// Hybrid of 6519/6583 rather than vista due to refactoring code into ExplorerFrame.
// Also a hybrid of the XP ExplorerWinMain due to needing dll support still.
int ExplorerWinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPTSTR pszCmdLine, int nCmdShow)
{
#ifndef RELEASE
    // If Shift is held down at startup, wait for debugger to attach
    if (GetKeyState(VK_SHIFT) < 0)
    {
        // Wait for a debugger to attach
        while (!IsDebuggerPresent())
        {
            Sleep(100);
        }
        DebugBreak(); // Optional: break into debugger immediately
    }

	AllocConsole();
	FILE* pFile;
	freopen_s(&pFile, "CONOUT$", "w", stdout);

    printf("Hello world!\n");

    wil::SetResultLoggingCallback([](const wil::FailureInfo& failure) noexcept
    {
        wchar_t message[2048];
        if (SUCCEEDED(GetFailureLogString(message, ARRAYSIZE(message), failure)))
        {
            switch (failure.type)
            {
                case wil::FailureType::Exception:
                case wil::FailureType::Return:
                case wil::FailureType::Log:
                {
                    wprintf(L"%s", message); // message includes newline
                    break;
                }
                case wil::FailureType::FailFast:
                {
                    MessageBoxW(nullptr, message, L"ExplorerEx", MB_ICONERROR);
                    break;
                }
            }
        }
    });
#endif

    SetPriorityClass(GetCurrentProcess(), HIGH_PRIORITY_CLASS);

#ifndef EXEX_DLL
    SHFusionInitializeFromModule(hInstance);
#endif
    g_hinstCabinet = hInstance;

    if (SUCCEEDED(_Module.Init(ObjectMap, hInstance)))
    {
        Cabinet_InitGlobalMetrics(0, nullptr);

        RegCreateKeyW(HKEY_CURRENT_USER, REGSTR_PATH_EXPLORER, &g_hkeyExplorer);
        if (g_hkeyExplorer == nullptr)
        {
        }

        HANDLE hMutex = nullptr;

#ifndef EXEX_DLL
        EXPLORERSERVER es = ShouldStartDesktopAndTray();
        if (es == EXPLORERSERVER_UNDETERMINED)
        {
            hMutex = CreateMutexW(nullptr, FALSE, L"Local\\ExplorerIsShellMutex");
            if (hMutex)
            {
                WaitForSingleObject(hMutex, INFINITE);
            }

            es = IsDesktopWindowAlreadyPresent() ? EXPLORERSERVER_NONE : EXPLORERSERVER_DESKTOP;
        }
#endif

#ifndef EXEX_DLL
        if (es == EXPLORERSERVER_DESKTOP)
        {
            ShellDDEInit(TRUE);

            SetProcessShutdownParameters(3, 0);
            _AutoRunTaskMan();

            MSG msg;
            PeekMessageW(&msg, nullptr, WM_QUIT, WM_QUIT, PM_NOREMOVE);

            HRESULT hrInit = _OleCoInitializeWaitForSCM();
            FileIconInit(TRUE);
#endif
            g_fLogonCycle = IsFirstInstanceAfterLogon();

            CheckDefaultUIFonts();
            ChangeUIfontsToNewDPI();
            CheckForServerAdminUI();

            if (g_fLogonCycle)
            {
                _ProcessRunOnceEx();

                _ProcessRunOnce();
            }

            CreateShellDirectories();

#ifndef EXEX_DLL
            ProcessInstallUninstallStubsIfNeeded();

            WasPrevShutdownClean();
            WriteCleanShutdown(FALSE);

            if (WinList_Init)
                WinList_Init();

            CProcessAndThreadRefHost host;

#endif
#ifndef EXEX_DLL
            HANDLE hDesktop = nullptr;

            if (!IsDesktopWindowAlreadyPresent())
            {
                hDesktop = CreateDesktopAndTray();
            }

            if (hMutex)
            {
                ReleaseMutex(hMutex);
                CloseHandle(hMutex);
            }

            if (hDesktop)
            {
                PostMessageW(v_hwndTray, TM_SHOWTRAYBALLOON, TRUE, 0);

                SetPriorityClass(GetCurrentProcess(), NORMAL_PRIORITY_CLASS);

                IExplorerHostCreator* pehc = nullptr;
                if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerHostCreator, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pehc))))
                {
                    pehc->CreateHost(CLSID_DesktopExplorerHost);
                }

                SHDesktopMessageLoop(hDesktop);

                host.WaitForRefs();
                IUnknown_SafeReleaseAndNullPtr(&pehc);

                WriteCleanShutdown(TRUE);
            }

            // Needed to start up properly on Windows 7:
            if (WinList_Terminate)
                WinList_Terminate();

            _OleCoUninitialize(hrInit);

            ShellDDEInit(FALSE);
        }
        else
        {
            if (hMutex)
            {
                ReleaseMutex(hMutex);
                CloseHandle(hMutex);
            }

            SetPriorityClass(GetCurrentProcess(), NORMAL_PRIORITY_CLASS);

            if (es == EXPLORERSERVER_NONE)
            {
                StartExplorerWindow();
            }
            else
            {
                HRESULT hrInit = SHCoInitialize();
                if (SUCCEEDED(hrInit))
                {
                    NEWFOLDERINFO fi = {};
                    if (SHExplorerParseCmdLine(PathGetArgsW(GetCommandLineW()), &fi))
                    {
                        if (es == EXPLORERSERVER_FACTORY)
                        {
                            IExplorerHostCreator* pehc;
                            if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerHostCreator, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pehc))))
                            {
                                if (SUCCEEDED(pehc->CreateHost(fi.clsid)))
                                {
                                    pehc->RunHost();
                                }
                                pehc->Release();
                            }
                        }
                        else
                        {
                            ASSERT(es == EXPLORERSERVER_SEPARATE);

                            IExplorerLauncher* pel;
                            if (SUCCEEDED(CoCreateInstance(CLSID_ExplorerLauncher, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pel))))
                            {
                                STARTUPINFOW si = { sizeof(si) };
                                GetStartupInfoW(&si);

                                POINT ptZero = {};
                                int wShowWindow;
                                if ((si.dwFlags & STARTF_USESHOWWINDOW) != 0)
                                    wShowWindow = si.wShowWindow;
                                else
                                    wShowWindow = SW_SHOWNORMAL;

                                pel->ShowWindow(
                                    CLSID_SeparateMultipleProcessExplorerHost, fi.pidl,
                                    (LAUNCHEXPLORERFLAGS)fi.uFlags, ptZero, wShowWindow, nullptr, nullptr, nullptr);
                                pel->Release();
                            }
                        }

                        ILFree(fi.pidl);
                    }

                    SHCoUninitialize(hrInit);
                }
            }
        }

        _Module.Term();
#endif
    }

#ifndef EXEX_DLL
    SHFusionUninitialize();
#endif

    return TRUE;
}

#ifdef _WIN64
//
// The purpose of this function is to spawn rundll32.exe if we have 32-bit stuff in 
// HKLM\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run that needs to be executed.
//
BOOL _ProcessRun6432()
{
    BOOL bRet = FALSE;

    if (!SHRestricted(REST_NOLOCALMACHINERUN))
    {
        if (SHKeyHasValues(HKEY_LOCAL_MACHINE, TEXT("Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run")))
        {
            TCHAR szWow64Path[MAX_PATH];

            if (ExpandEnvironmentStringsW(TEXT("%SystemRoot%\\SysWOW64"), szWow64Path, ARRAYSIZE(szWow64Path)))
            {
                TCHAR sz32BitRunOnce[MAX_PATH];
                PROCESS_INFORMATION pi = {0};

                if (SUCCEEDED(StringCchPrintf(sz32BitRunOnce, ARRAYSIZE(sz32BitRunOnce), TEXT("%s\\runonce.exe"), szWow64Path)))
                {
                    if (CreateProcessWithArgs(sz32BitRunOnce, TEXT("/Run6432"), szWow64Path, &pi))
                    {
                        CloseHandle(pi.hProcess);
                        CloseHandle(pi.hThread);

                        bRet = TRUE;
                    }
                }
            }
        }
    }

    return bRet;
}
#endif  // _WIN64


STDAPI_(BOOL) Startup_ExecuteRegAppEnumProc(LPCTSTR szSubkey, LPCTSTR szCmdLine, RRA_FLAGS fFlags, LPARAM lParam)
{
    BOOL bRet = ExecuteRegAppEnumProc(szSubkey, szCmdLine, fFlags, lParam);
    
    if (!bRet && !(fFlags & RRA_DELETE))
    {
        c_tray.LogFailedStartupApp();
    }

    return bRet;
}


typedef struct
{
    RESTRICTIONS rest;
    HKEY hKey;
    const TCHAR* psz;
    DWORD dwRRAFlags;
}
STARTUPGROUP;

BOOL _RunStartupGroup(const STARTUPGROUP* pGroup, int cGroup)
{
    BOOL bRet = FALSE;

    // make sure SHRestricted is working ok
    ASSERT(!SHRestricted(REST_NONE));

    for (int i = 0; i < cGroup; i++)
    {
        if (!SHRestricted(pGroup[i].rest))
        {
            bRet = Cabinet_EnumRegApps(pGroup[i].hKey, pGroup[i].psz, pGroup[i].dwRRAFlags, Startup_ExecuteRegAppEnumProc, 0);
        }
    }

    return bRet;
}


BOOL _ProcessRun()
{
    static const STARTUPGROUP s_RunTasks [] =
    {
        { REST_NONE,                    HKEY_LOCAL_MACHINE, REGSTR_PATH_RUN_POLICY, RRA_NOUI }, // HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer\\Run
        { REST_NOLOCALMACHINERUN,       HKEY_LOCAL_MACHINE, REGSTR_PATH_RUN,        RRA_NOUI }, // HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Run
        { REST_NONE,                    HKEY_CURRENT_USER,  REGSTR_PATH_RUN_POLICY, RRA_NOUI }, // HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer\\Run
        { REST_NOCURRENTUSERRUN,        HKEY_CURRENT_USER,  REGSTR_PATH_RUN,        RRA_NOUI }, // HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run
    };
    
    BOOL bRet = _RunStartupGroup(s_RunTasks, ARRAYSIZE(s_RunTasks));

#ifdef _WIN64
    // see if we need to launch any 32-bit apps under wow64
    _ProcessRun6432();
#endif

    return bRet;
}


BOOL _ProcessPerUserRunOnce()
{
    static const STARTUPGROUP s_PerUserRunOnceTasks [] =
    {
        { REST_NOCURRENTUSERRUNONCE,    HKEY_CURRENT_USER,  REGSTR_PATH_RUNONCE,    RRA_DELETE | RRA_NOUI },    // HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce
    };

    return _RunStartupGroup(s_PerUserRunOnceTasks, ARRAYSIZE(s_PerUserRunOnceTasks));
}


DWORD WINAPI RunStartupAppsThread(void *pv)
{
    // Some of the items we launch during startup assume that com is initialized.  Make this
    // assumption true.
    HRESULT hrInit = SHCoInitialize();

    // These global flags are set once long before our thread starts and are then only
    // read so we don't need to worry about timing issues.
    if (g_fLogonCycle && !g_fCleanBoot)
    {
        // We only run these startup items if g_fLogonCycle is TRUE. This prevents
        // them from running again if the shell crashes and restarts.

        _ProcessOldRunAndLoadEquals();
        _ProcessRun();
        _ExecuteStartupPrograms();
    }

    // As a best guess, the HKCU RunOnce key is executed regardless of the g_fLogonCycle
    // becuase it was once hoped that we could install newer versions of IE without
    // requiring a reboot.  They would place something in the CU\RunOnce key and then
    // shutdown and restart the shell to continue their setup process.  I believe this
    // idea was later abandoned but the code change is still here.  Since that could
    // some day be a useful feature I'm leaving it the same.
    _ProcessPerUserRunOnce();

    // we need to run all the non-blocking items first.  Then we spend the rest of this threads life
    // runing the synchronized objects one after another.
    if (g_fLogonCycle && !g_fCleanBoot)
    {
        //_RunWelcome();
    }

    PostMessage(v_hwndTray, TM_STARTUPAPPSLAUNCHED, 0, 0);

    SHCoUninitialize(hrInit);

    return TRUE;
}


void RunStartupApps()
{
    DWORD dwThreadID;
    HANDLE handle = CreateThread(nullptr, 0, RunStartupAppsThread, 0, 0, &dwThreadID);
    if (handle)
    {
        CloseHandle(handle);
    }
}
