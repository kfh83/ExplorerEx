#define INITGUID
//+-------------------------------------------------------------------------
//
//  TaskMan - NT TaskManager
//  Copyright (C) Microsoft
//
//  File:       shundoc.cpp
//
//  History:    Oct-11-24   aubymori  Created
//              Jan-22-25   kfh83     Modified
//
//--------------------------------------------------------------------------
#include "pch.h"
#include "shundoc.h"
#include "msi.h"
#include "port32.h"
#include "path.h"


//
// Classes
//

class CConvertStrW
{
public:
    operator WCHAR* ();

protected:
    CConvertStrW();
    ~CConvertStrW();
    void Free();

    LPWSTR   _pwstr;
    WCHAR    _awch[MAX_PATH * sizeof(WCHAR)];
};

inline
CConvertStrW::CConvertStrW()
{
    _pwstr = NULL;
}

inline
CConvertStrW::~CConvertStrW()
{
    Free();
}

inline
CConvertStrW::operator WCHAR* ()
{
    return _pwstr;
}



class CStrInW : public CConvertStrW
{
public:
    CStrInW(LPCSTR pstr) { Init(pstr, -1); }
    CStrInW(LPCSTR pstr, int cch) { Init(pstr, cch); }
    int strlen();

protected:
    CStrInW();
    void Init(LPCSTR pstr, int cch);

    int _cwchLen;
};

inline
CStrInW::CStrInW()
{
}

inline int
CStrInW::strlen()
{
    return _cwchLen;
}



class CStrOutW : public CConvertStrW
{
public:
    CStrOutW(LPSTR pstr, int cchBuf);
    ~CStrOutW();

    int     BufSize();
    int     ConvertIncludingNul();
    int     ConvertExcludingNul();

private:
    LPSTR  	_pstr;
    int     _cchBuf;
};

inline int
CStrOutW::BufSize()
{
    return _cchBuf;
}

//
// Function definitions
// 

//HRESULT(STDMETHODCALLTYPE* IUnknown_Exec)(IUnknown* punk, const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut) = nullptr;
//HRESULT(STDMETHODCALLTYPE* IUnknown_GetClassID)(IUnknown* punk, CLSID* pclsid) = nullptr;
//HRESULT(STDMETHODCALLTYPE* IUnknown_OnFocusChangeIS)(IUnknown* punk, IUnknown* punkSrc, BOOL fSetFocus) = nullptr;
//HRESULT(STDMETHODCALLTYPE* IUnknown_QueryStatus)(IUnknown* punk, const GUID* pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pcmdtext) = nullptr;
//HRESULT(STDMETHODCALLTYPE* IUnknown_UIActivateIO)(IUnknown* punk, BOOL fActivate, LPMSG lpMsg) = nullptr;
//HRESULT(STDMETHODCALLTYPE* IUnknown_TranslateAcceleratorIO)(IUnknown* punk, LPMSG lpMsg) = nullptr;


HRESULT IUnknown_DragEnter(IUnknown* punk, IDataObject* pdtobj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    HRESULT hr = E_FAIL;
    if (punk)
    {
        IDropTarget* pdt;
        hr = punk->QueryInterface(IID_PPV_ARGS(&pdt));
        if (SUCCEEDED(hr))
        {
            hr = pdt->DragEnter(pdtobj, grfKeyState, pt, pdwEffect);
            pdt->Release();
        }
    }

    if (FAILED(hr))
        *pdwEffect = DROPEFFECT_NONE;

    return hr;
}

HRESULT IUnknown_DragLeave(IUnknown* punk)
{
    HRESULT hr = E_FAIL;
    if (punk)
    {
        IDropTarget* pdt;
        hr = punk->QueryInterface(IID_PPV_ARGS(&pdt));
        if (SUCCEEDED(hr))
        {
            hr = pdt->DragLeave();
            pdt->Release();
        }
    }
    return hr;
}

HRESULT IUnknown_DragOver(IUnknown* punk, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    HRESULT hr = E_FAIL;
    if (punk)
    {
        IDropTarget* pdt;
        hr = punk->QueryInterface(IID_PPV_ARGS(&pdt));
        if (SUCCEEDED(hr))
        {
            hr = pdt->DragOver(grfKeyState, pt, pdwEffect);
            pdt->Release();
        }
    }

    if (FAILED(hr))
        *pdwEffect = DROPEFFECT_NONE;

    return hr;
}


//HRESULT(STDMETHODCALLTYPE* SHPropagateMessage)(HWND hwndParent, UINT uMsg, WPARAM wParam, LPARAM lParam, int iFlags);
//HRESULT(STDMETHODCALLTYPE* SHGetUserDisplayName)(LPWSTR pszDisplayName, PULONG uLen);
//HRESULT(STDMETHODCALLTYPE* SHGetUserPicturePath)(LPCWSTR pszUsername, DWORD dwFlags, LPWSTR pszPath, DWORD cchPathMax);
//HRESULT(STDMETHODCALLTYPE* SHSetWindowBits)(HWND hwnd, int iWhich, DWORD dwBits, DWORD dwValue);
//HRESULT(STDMETHODCALLTYPE* SHRunIndirectRegClientCommand)(HWND hwnd, LPCWSTR pszClient);
//HRESULT(STDMETHODCALLTYPE* SHInvokeDefaultCommand)(HWND hwnd, IShellFolder* psf, LPCITEMIDLIST pidlItem) ;
//HRESULT(STDMETHODCALLTYPE* SHSettingsChanged)(WPARAM wParam, LPARAM lParam) ;
//HRESULT(STDMETHODCALLTYPE* SHIsChildOrSelf)(HWND hwndParent, HWND hwnd) ;
//HRESULT(STDMETHODCALLTYPE* SHLoadRegUIStringW)(HKEY     hkey, LPCWSTR  pszValue, LPWSTR   pszOutBuf, UINT     cchOutBuf) ;
//HWND(STDMETHODCALLTYPE* SHCreateWorkerWindowW)(WNDPROC pfnWndProc, HWND hwndParent, DWORD dwExStyle, DWORD dwFlags, HMENU hmenu, void* p) ;
//BOOL(WINAPI* SHQueueUserWorkItem)(IN LPTHREAD_START_ROUTINE pfnCallback, IN LPVOID pContext, IN LONG lPriority, IN DWORD_PTR dwTag, OUT DWORD_PTR* pdwId OPTIONAL, IN LPCSTR pszModule OPTIONAL, IN DWORD dwFlags) ;
//BOOL(WINAPI* WinStationSetInformationW)(HANDLE hServer, ULONG LogonId, WINSTATIONINFOCLASS WinStationInformationClass, PVOID  pWinStationInformation, ULONG WinStationInformationLength) ;
//BOOL(WINAPI* WinStationUnRegisterConsoleNotification)(HANDLE hServer, HWND hWnd) ;
//BOOL(STDMETHODCALLTYPE* SHFindComputer)(LPCITEMIDLIST pidlFolder, LPCITEMIDLIST pidlSaveFile) ;
//BOOL(STDMETHODCALLTYPE* SHTestTokenPrivilegeW)(HANDLE hToken, LPCWSTR pszPrivilegeName) ;
//LRESULT(WINAPI* SHDefWindowProc)(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) ;
//UINT(WINAPI* MsiDecomposeDescriptorW)(LPCWSTR	szDescriptor, LPWSTR szProductCode, LPWSTR szFeatureId, LPWSTR szComponentCode, DWORD* pcchArgsOffset) ;
//HRESULT(STDMETHODCALLTYPE* ExitWindowsDialog)(HWND hwndParent) ;
//UINT(STDMETHODCALLTYPE* SHGetCurColorRes)(void) ;
//UINT(WINAPI* ImageList_GetFlags)(HIMAGELIST himl) ;
//INT(STDMETHODCALLTYPE* SHMessageBoxCheckExW)(HWND hwnd, HINSTANCE hinst, LPCWSTR pszTemplateName, DLGPROC pDlgProc, LPVOID pData, int iDefault, LPCWSTR pszRegVal) ;
//INT(STDMETHODCALLTYPE* RunFileDlg)(HWND hwndParent, HICON hIcon, LPCTSTR pszWorkingDir, LPCTSTR pszTitle, LPCTSTR pszPrompt, DWORD dwFlags) ;
//VOID(STDMETHODCALLTYPE* SHUpdateRecycleBinIcon)() ;
//VOID(STDMETHODCALLTYPE* LogoffWindowsDialog)(HWND hwndParent) ;
//VOID(STDMETHODCALLTYPE* DisconnectWindowsDialog)(HWND hwndParent) ;
//BOOL(STDMETHODCALLTYPE* RegisterShellHook)(HWND hwnd, BOOL fInstall) ;
//DWORD_PTR(WINAPI* SHGetMachineInfo)(UINT gmi) ;
//HMENU(STDMETHODCALLTYPE* SHGetMenuFromID)(HMENU hmMain, UINT uID) ;

//COLORREF(STDMETHODCALLTYPE* SHFillRectClr)(HDC hdc, LPRECT lprect, COLORREF color) ;

//BOOL(STDMETHODCALLTYPE* WinStationRegisterConsoleNotification)(HANDLE  hServer, HWND    hWnd, DWORD   dwFlags) ;

// SHRegisterDarwinLink takes ownership of the pidl

HRESULT DisplayNameOfAsOLESTR(IShellFolder* psf, LPCITEMIDLIST pidl, DWORD flags, LPWSTR* ppsz)
{
    *ppsz = NULL;
    STRRET sr;
    HRESULT hr = psf->GetDisplayNameOf(pidl, flags, &sr);
    if (SUCCEEDED(hr))
        hr = StrRetToStrW(&sr, pidl, ppsz);
    return hr;
}

LPITEMIDLIST ILCloneParent(LPCITEMIDLIST pidl)
{
    LPITEMIDLIST pidlParent = ILClone(pidl);
    if (pidlParent)
        ILRemoveLastID(pidlParent);

    return pidlParent;
}

void SHAdjustLOGFONT(IN OUT LOGFONT* plf)
{
    if (plf->lfCharSet == SHIFTJIS_CHARSET ||
        plf->lfCharSet == HANGEUL_CHARSET ||
        plf->lfCharSet == GB2312_CHARSET ||
        plf->lfCharSet == CHINESEBIG5_CHARSET)
    {
        if (plf->lfWeight > FW_NORMAL)
            plf->lfWeight = FW_NORMAL;
    }
}

BOOL _SHIsMenuSeparator2(HMENU hm, int i, BOOL* pbIsNamed)
{
    MENUITEMINFO mii;
    BOOL bLocal;

    if (!pbIsNamed)
        pbIsNamed = &bLocal;

    *pbIsNamed = FALSE;

    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_TYPE | MIIM_ID;
    mii.cch = 0;    // WARNING: We MUST initialize it to 0!!!
    if (GetMenuItemInfo(hm, i, TRUE, &mii) && (mii.fType & MFT_SEPARATOR))
    {
        // NOTE that there is a bug in either 95 or NT user!!!
        // 95 returns 16 bit ID's and NT 32 bit therefore there is a
        // the following may fail, on win9x, to evaluate to false
        // without casting
        *pbIsNamed = ((WORD)mii.wID != (WORD)-1);
        return TRUE;
    }
    return FALSE;
}

void _SHPrettyMenu(HMENU hm)
{
    BOOL bSeparated = TRUE;
    BOOL bWasNamed = TRUE;

    for (int i = GetMenuItemCount(hm) - 1; i > 0; --i)
    {
        BOOL bIsNamed;
        if (_SHIsMenuSeparator2(hm, i, &bIsNamed))
        {
            if (bSeparated)
            {
                // if we have two separators in a row, only one of which is named
                // remove the non named one!
                if (bIsNamed && !bWasNamed)
                {
                    DeleteMenu(hm, i + 1, MF_BYPOSITION);
                    bWasNamed = bIsNamed;
                }
                else
                {
                    DeleteMenu(hm, i, MF_BYPOSITION);
                }
            }
            else
            {
                bWasNamed = bIsNamed;
                bSeparated = TRUE;
            }
        }
        else
        {
            bSeparated = FALSE;
        }
    }

    // The above loop does not handle the case of many separators at
    // the beginning of the menu
    while (_SHIsMenuSeparator2(hm, 0, NULL))
    {
        DeleteMenu(hm, 0, MF_BYPOSITION);
    }
}

HRESULT SHGetIDListFromUnk(IUnknown* punk, LPITEMIDLIST* ppidl)
{
    *ppidl = NULL;

    HRESULT hr = E_NOINTERFACE;
    if (punk)
    {
        IPersistFolder2* ppf;
        IPersistIDList* pperid;
        if (SUCCEEDED(punk->QueryInterface(IID_PPV_ARG(IPersistIDList, &pperid))))
        {
            hr = pperid->GetIDList(ppidl);
            pperid->Release();
        }
        else if (SUCCEEDED(punk->QueryInterface(IID_PPV_ARG(IPersistFolder2, &ppf))))
        {
            hr = ppf->GetCurFolder(ppidl);
            ppf->Release();
        }
    }
    return hr;
}

// moved to runonce.cpp


DWORD SHProcessMessagesUntilEventEx(HWND hwnd, HANDLE hEvent, DWORD dwTimeout, DWORD dwWakeMask)
{
    DWORD dwEndTime = GetTickCount64() + dwTimeout;
    LONG lWait = (LONG)dwTimeout;
    DWORD dwReturn;

    if (!hEvent && (dwTimeout == INFINITE))
    {
        return -1;
    }

    for (;;)
    {
        DWORD dwCount = hEvent ? 1 : 0;
        dwReturn = MsgWaitForMultipleObjects(dwCount, &hEvent, FALSE, lWait, dwWakeMask);

        // were we signalled or did we time out?
        if (dwReturn != (WAIT_OBJECT_0 + dwCount))
        {
            break;
        }

        // we woke up because of messages.
        MSG msg;
        while (PeekMessage(&msg, hwnd, 0, 0, PM_REMOVE))
        {
            TranslateMessage(&msg);
            if (msg.message == WM_SETCURSOR)
            {
                SetCursor(LoadCursor(NULL, IDC_WAIT));
            }
            else
            {
                DispatchMessage(&msg);
            }
        }

        // calculate new timeout value
        if (dwTimeout != INFINITE)
        {
            lWait = (LONG)dwEndTime - GetTickCount64();
        }
    }

    return dwReturn;
}

// 
// Determine if the images represented by the two icons are the same
// (NOTE: this does not compare ICON masks, but this should never be a distinguishing factor)
//
BOOL SHAreIconsEqual(HICON hIcon1, HICON hIcon2)
{
    BOOL bRet = FALSE;

    ICONINFO ii1;
    if (hIcon1 && hIcon2 && GetIconInfo(hIcon1, &ii1))
    {
        ICONINFO ii2;
        if (GetIconInfo(hIcon2, &ii2))
        {
            BITMAP bm1 = { 0 };
            if (GetObject(ii1.hbmColor, sizeof(bm1), &bm1))
            {
                BITMAP bm2 = { 0 };
                if (GetObject(ii2.hbmColor, sizeof(bm2), &bm2))
                {
                    if ((bm1.bmWidth == bm2.bmWidth) && (bm1.bmHeight == bm2.bmHeight))
                    {
                        BITMAPINFO bmi = { 0 };
                        bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
                        bmi.bmiHeader.biWidth = bm1.bmWidth;
                        bmi.bmiHeader.biHeight = bm1.bmHeight;
                        bmi.bmiHeader.biPlanes = 1;
                        bmi.bmiHeader.biBitCount = 32;
                        bmi.bmiHeader.biCompression = BI_RGB;

                        HDC hdc = GetDC(NULL);
                        if (hdc)
                        {
                            ULONG* pulIcon1 = new ULONG[bm1.bmWidth * bm1.bmHeight];
                            if (pulIcon1)
                            {
                                if (GetDIBits(hdc, ii1.hbmColor, 0, bm1.bmHeight, (LPVOID)pulIcon1, &bmi, DIB_RGB_COLORS))
                                {
                                    ULONG* pulIcon2 = new ULONG[bm1.bmWidth * bm1.bmHeight];
                                    if (pulIcon2)
                                    {
                                        if (GetDIBits(hdc, ii2.hbmColor, 0, bm1.bmHeight, (LPVOID)pulIcon2, &bmi, DIB_RGB_COLORS))
                                        {
                                            bRet = (0 == memcmp(pulIcon1, pulIcon2, bm1.bmWidth * bm1.bmHeight * sizeof(ULONG)));
                                        }
                                        delete[] pulIcon2;
                                    }
                                }
                                delete[] pulIcon1;
                            }
                            ReleaseDC(NULL, hdc);
                        }
                    }
                }
            }
            DeleteObject(ii2.hbmColor);
            DeleteObject(ii2.hbmMask);
        }
        DeleteObject(ii1.hbmColor);
        DeleteObject(ii1.hbmMask);
    }

    return bRet;
}


HRESULT SHLoadLegacyRegUIString(HKEY hk, LPCTSTR pszSubkey, LPTSTR pszOutBuf, UINT cchOutBuf)
{
    HKEY hkClose = NULL;

    ASSERT(cchOutBuf);
    pszOutBuf[0] = TEXT('\0');

    if (pszSubkey && *pszSubkey)
    {
        DWORD dwError = RegOpenKeyEx(hk, pszSubkey, 0, KEY_QUERY_VALUE, &hkClose);
        if (dwError != ERROR_SUCCESS)
        {
            return HRESULT_FROM_WIN32(dwError);
        }
        hk = hkClose;
    }

    HRESULT hr = SHLoadRegUIStringW(hk, TEXT("LocalizedString"), pszOutBuf, cchOutBuf);
    if (FAILED(hr) || pszOutBuf[0] == TEXT('\0'))
    {
        hr = SHLoadRegUIStringW(hk, TEXT(""), pszOutBuf, cchOutBuf);
    }

    if (hkClose)
    {
        RegCloseKey(hkClose);
    }

    return hr;
}

HRESULT SHBindToObjectEx(IShellFolder* psf, LPCITEMIDLIST pidl, LPBC pbc, REFIID riid, void** ppvOut)
{
    HRESULT hr;
    IShellFolder* psfRelease;

    if (!psf)
    {
        SHGetDesktopFolder(&psf);
        psfRelease = psf;
    }
    else
    {
        psfRelease = NULL;
    }

    if (psf)
    {
        if (!pidl || ILIsEmpty(pidl))
        {
            hr = psf->QueryInterface(riid, ppvOut);
        }
        else
        {
            hr = psf->BindToObject(pidl, pbc, riid, ppvOut);
        }
    }
    else
    {
        *ppvOut = NULL;
        hr = E_FAIL;
    }

    if (psfRelease)
    {
        psfRelease->Release();
    }

    if (SUCCEEDED(hr) && (*ppvOut == NULL))
    {
        // Some shell extensions (eg WS_FTP) will return success and a null out pointer
        hr = E_FAIL;
    }

    return hr;
}

BOOL SetWindowZorder(HWND hwnd, HWND hwndInsertAfter)
{
    return SetWindowPos(hwnd, hwndInsertAfter, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
}

BOOL CALLBACK _FixZorderEnumProc(HWND hwnd, LPARAM lParam)
{
    HWND hwndTest = (HWND)lParam;
    HWND hwndOwner = hwnd;

    while (hwndOwner = GetWindow(hwndOwner, GW_OWNER))
    {
        if (hwndOwner == hwndTest)
        {
            SetWindowZorder(hwnd, HWND_NOTOPMOST);

            break;
        }
    }

    return TRUE;
}

HRESULT SHCoInitialize(void)
{
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(hr))
    {
        hr = CoInitializeEx(NULL, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE);
    }
    return hr;
}

BOOL SHIsSameObject(IUnknown* punk1, IUnknown* punk2)
{
    if (!punk1 || !punk2)
    {
        return FALSE;
    }
    else if (punk1 == punk2)
    {
        // Quick shortcut -- if they're the same pointer
        // already then they must be the same object
        //
        return TRUE;
    }
    else
    {
        IUnknown* punkI1 = nullptr;
        IUnknown* punkI2 = nullptr;

        // Some apps don't implement QueryInterface! (SecureFile)
        HRESULT hr = punk1->QueryInterface(IID_PPV_ARG(IUnknown, &punkI1));
        if (SUCCEEDED(hr))
        {
            punkI1->Release();
            hr = punk2->QueryInterface(IID_PPV_ARG(IUnknown, &punkI2));
            if (SUCCEEDED(hr))
                punkI2->Release();
        }
        return SUCCEEDED(hr) && (punkI1 == punkI2);
    }
}

BOOL GetExplorerUserSetting(HKEY hkeyRoot, LPCTSTR pszSubKey, LPCTSTR pszValue)
{
    TCHAR szPath[MAX_PATH];
    TCHAR szPathExplorer[MAX_PATH];
    DWORD cbSize = ARRAYSIZE(szPath);
    DWORD dwType;

    PathCombine(szPathExplorer, L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer", pszSubKey);
    if (ERROR_SUCCESS == SHGetValue(hkeyRoot, szPathExplorer, pszValue,
        &dwType, szPath, &cbSize))
    {
        // Zero in the DWORD case or NULL in the string case
        // indicates that this item is not available.
        if (dwType == REG_DWORD)
            return *((DWORD*)szPath) != 0;
        else
            return (TCHAR)szPath[0] != 0;
    }

    return -1;
}


BOOL SHForceWindowZorder(HWND hwnd, HWND hwndInsertAfter)
{
    BOOL fRet = SetWindowZorder(hwnd, hwndInsertAfter);

    if (fRet && hwndInsertAfter == HWND_TOPMOST)
    {
        if (!(GetWindowLong(hwnd, GWL_EXSTYLE) & WS_EX_TOPMOST))
        {
            //
            // user didn't actually move the hwnd to topmost
            //
            // According to GerardoB, this can happen if the window has
            // an owned window that somehow has become topmost while the 
            // owner remains non-topmost, i.e., the two have become
            // separated in the z-order.  In this state, when the owner
            // window tries to make itself topmost, the call will
            // silently fail.
            //
            // TERRIBLE HORRIBLE NO GOOD VERY BAD HACK
            //
            // Hacky fix is to enumerate the toplevel windows, check to see
            // if any are topmost and owned by hwnd, and if so, make them
            // non-topmost.  Then, retry the SetWindowPos call.
            //

            // Fix up the z-order
            EnumWindows(_FixZorderEnumProc, (LPARAM)hwnd);

            // Retry the set.  (This should make all owned windows topmost as well.)
            SetWindowZorder(hwnd, HWND_TOPMOST);
        }
    }

    return fRet;
}

//BOOL(WINAPI* EndTask)(HWND hWnd, BOOL fShutDown, BOOL fForce) = nullptr;

#define FILETIMEtoInt64(ft) _FILETIMEtoInt64(&(ft))

void SetFILETIMEfromInt64(FILETIME* pft, unsigned __int64 i64)
{
    pft->dwLowDateTime = (DWORD)i64;
    pft->dwHighDateTime = (DWORD)(i64 >> 32);
}

void IncrementFILETIME(FILETIME* pft, unsigned __int64 iAdjust)
{
    SetFILETIMEfromInt64(pft, _FILETIMEtoInt64(pft) + iAdjust);
}

void DecrementFILETIME(FILETIME* pft, unsigned __int64 iAdjust)
{
    SetFILETIMEfromInt64(pft, _FILETIMEtoInt64(pft) - iAdjust);
}

// shits not exported bro i swear


BOOL ConvertHexStringToIntW(WCHAR* pszHexNum, int* piNum)
{
    int   n = 0L;
    WCHAR* psz = pszHexNum;

    for (n = 0; ; psz = CharNextW(psz))
    {
        if ((*psz >= '0') && (*psz <= '9'))
            n = 0x10 * n + *psz - '0';
        else
        {
            WCHAR ch = *psz;
            int n2;

            if (ch >= 'a')
                ch -= 'a' - 'A';

            n2 = ch - 'A' + 0xA;
            if (n2 >= 0xA && n2 <= 0xF)
                n = 0x10 * n + n2;
            else
                break;
        }
    }

    /*
     * Update results
     */
    *piNum = n;

    return (psz != pszHexNum);
}

VOID MuSecurity(VOID)
{
    //
    // Do nothing on the console
    //

    if (SHGetMachineInfo(GMI_TSCLIENT))
    {
        WinStationSetInformationW(SERVERNAME_CURRENT,
            LOGONID_CURRENT,
            WinStationNtSecurity,
            NULL, 0);
    }
}

BOOL CALLBACK Mirror_EnumUILanguagesProc(LPTSTR lpUILanguageString, LONG_PTR lParam)
{
    int langID = 0;

    ConvertHexStringToIntW(lpUILanguageString, &langID);

    if ((LANGID)langID == ((LPMUIINSTALLLANG)lParam)->LangID)
    {
        ((LPMUIINSTALLLANG)lParam)->bInstalled = TRUE;
        return FALSE;
    }
    return TRUE;
}

BOOL Mirror_IsWindowMirroredRTL(HWND hWnd)
{
    return (GetWindowLongA(hWnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL);
}

BOOL Mirror_IsUILanguageInstalled(LANGID langId)
{
    MUIINSTALLLANG MUILangInstalled = { 0 };
    MUILangInstalled.LangID = langId;

    EnumUILanguagesW(Mirror_EnumUILanguagesProc, 0, (LONG_PTR)&MUILangInstalled);

    return MUILangInstalled.bInstalled;
}

typedef DWORD(WINAPI* PFNSETLAYOUT)(HDC, DWORD);            // gdi32!SetLayout

DWORD Mirror_SetLayout(HDC hdc, DWORD dwLayout)
{
    DWORD dwRet = 0;
    static PFNSETLAYOUT pfnSetLayout = NULL;

    if (NULL == pfnSetLayout)
    {
        HMODULE hmod = GetModuleHandleA("GDI32");

        if (hmod)
            pfnSetLayout = (PFNSETLAYOUT)GetProcAddress(hmod, "SetLayout");
    }

    if (pfnSetLayout)
        dwRet = pfnSetLayout(hdc, dwLayout);

    return dwRet;
}

BOOL IsBiDiLocalizedSystemEx(LANGID* pLangID)
{
    int           iLCID = 0L;
    static BOOL   bRet = (BOOL)(DWORD)-1;
    static LANGID langID = MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL);

    if (bRet != (BOOL)(DWORD)-1)
    {
        if (bRet && pLangID)
        {
            *pLangID = langID;
        }
        return bRet;
    }

    bRet = FALSE;
    /*
     * Need to use NT5 detection method (Multiligual UI ID)
     */
    langID = GetUserDefaultUILanguage();

    if (langID)
    {
        WCHAR wchLCIDFontSignature[16];
        iLCID = MAKELCID(langID, SORT_DEFAULT);

        /*
         * Let's verify this is a RTL (BiDi) locale. Since reg value is a hex string, let's
         * convert to decimal value and call GetLocaleInfo afterwards.
         * LOCALE_FONTSIGNATURE always gives back 16 WCHARs.
         */

        if (GetLocaleInfoW(iLCID,
            LOCALE_FONTSIGNATURE,
            (WCHAR*)&wchLCIDFontSignature[0],
            (sizeof(wchLCIDFontSignature) / sizeof(WCHAR))))
        {

            /* Let's verify the bits we have a BiDi UI locale */
            if ((wchLCIDFontSignature[7] & (WCHAR)0x0800) && Mirror_IsUILanguageInstalled(langID))
            {
                bRet = TRUE;
            }
        }
    }

    if (bRet && pLangID)
    {
        *pLangID = langID;
    }
    return bRet;
}

// aight we good

BOOL IsBiDiLocalizedSystem(void)
{
    return IsBiDiLocalizedSystemEx(NULL);
}


BOOL IsRestrictedOrUserSetting(HKEY hkeyRoot, RESTRICTIONS rest, LPCTSTR pszSubKey, LPCTSTR pszValue, UINT flags)
{
    // See if the system policy restriction trumps

    DWORD dwRest = SHRestricted(rest);

    if (dwRest == 1)
        return TRUE;

    if (dwRest == 2)
        return FALSE;

    //
    //  Restriction not in place or defers to user setting.
    //
    BOOL fValidKey = GetExplorerUserSetting(hkeyRoot, pszSubKey, pszValue);

    switch (fValidKey)
    {
    case 0:     // Key is present and zero
        if (flags & ROUS_KEYRESTRICTS)
            return FALSE;       // restriction not present
        else
            return TRUE;        // ROUS_KEYALLOWS, value is 0 -> restricted

    case 1:     // Key is present and nonzero

        if (flags & ROUS_KEYRESTRICTS)
            return TRUE;        // restriction present -> restricted
        else
            return FALSE;       // ROUS_KEYALLOWS, value is 1 -> not restricted

    default:
        // Fall through

    case -1:    // Key is not present
        return (flags & ROUS_DEFAULTRESTRICT);
    }

    /*NOTREACHED*/
}

// DDE shit

TCHAR const c_szCreateGroup[] = TEXT("CreateGroup");
TCHAR const c_szShowGroup[] = TEXT("ShowGroup");
TCHAR const c_szAddItem[] = TEXT("AddItem");
TCHAR const c_szExitProgman[] = TEXT("ExitProgman");
TCHAR const c_szDeleteGroup[] = TEXT("DeleteGroup");
TCHAR const c_szDeleteItem[] = TEXT("DeleteItem");
TCHAR const c_szReplaceItem[] = TEXT("ReplaceItem");
TCHAR const c_szReload[] = TEXT("Reload");
TCHAR const c_szFindFolder[] = TEXT("FindFolder");
TCHAR const c_szOpenFindFile[] = TEXT("OpenFindFile");
TCHAR const c_szTrioDataFax[] = TEXT("DDEClient");
TCHAR const c_szTalkToPlus[] = TEXT("ddeClass");
TCHAR const c_szStartUp[] = TEXT("StartUp");
TCHAR const c_szCCMail[] = TEXT("ccInsDDE");
TCHAR const c_szBodyWorks[] = TEXT("BWWFrame");
TCHAR const c_szMediaRecorder[] = TEXT("DDEClientWndClass");
TCHAR const c_szDiscis[] = TEXT("BACKSCAPE");
TCHAR const c_szMediaRecOld[] = TEXT("MediaRecorder");
TCHAR const c_szMediaRecNew[] = TEXT("Media Recorder");
TCHAR const c_szDialog[] = TEXT("#32770");
TCHAR const c_szJourneyMan[] = TEXT("Sender");
TCHAR const c_szCADDE[] = TEXT("CA_DDECLASS");
TCHAR const c_szFaxServe[] = TEXT("Install");
TCHAR const c_szMakePMG[] = TEXT("Make Program Manager Group");
TCHAR const c_szViewFolder[] = TEXT("ViewFolder");
TCHAR const c_szExploreFolder[] = TEXT("ExploreFolder");
TCHAR const c_szRUCabinet[] = TEXT("ConfirmCabinetID");
CHAR const c_szNULLA[] = "";
TCHAR const c_szGetIcon[] = TEXT("GetIcon");
TCHAR const c_szGetDescription[] = TEXT("GetDescription");
TCHAR const c_szGetWorkingDir[] = TEXT("GetWorkingDir");
TCHAR const c_szShellFile[] = TEXT("ShellFile");

BOOL DDE_CreateGroup(LPTSTR, UINT*, PDDECONV)
{
    //MUST IMPLEMENT
    printf("DDE_CreateGroup\n");
    return 0;
}
BOOL DDE_ShowGroup(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_ShowGroup\n");
	return 0;
}
BOOL DDE_AddItem(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_AddItem\n");
	return 0;
}
BOOL DDE_ExitProgman(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_ExitProgman\n");
	return 0;
}
BOOL DDE_DeleteGroup(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_DeleteGroup\n");
	return 0;
}
BOOL DDE_DeleteItem(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_DeleteItem\n");
	return 0;
}
#define DDE_ReplaceItem DDE_DeleteItem
BOOL DDE_Reload(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_Reload\n");
	return 0;
}
BOOL DDE_ViewFolder(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_ViewFolder\n");
	return 0;
}
BOOL DDE_ExploreFolder(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_ExploreFolder\n");
	return 0;
}
BOOL DDE_FindFolder(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_FindFolder\n");
	return 0;
}
BOOL DDE_OpenFindFile(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_OpenFindFile\n");
	return 0;
}
BOOL DDE_ConfirmID(LPTSTR lpszBuf, UINT* lpwCmd, PDDECONV pddec)
{
	//MUST IMPLEMENT
    printf("DDE_ConfirmID\n");
	return 0;
}
BOOL DDE_ShellFile(LPTSTR lpszBuf, UINT* lpwCmd, PDDECONV pddec)
{
	//MUST IMPLEMENT
    printf("DDE_ShellFile\n");
	return 0;
}
BOOL DDE_Beep(LPTSTR, UINT*, PDDECONV)
{
	//MUST IMPLEMENT
    printf("DDE_Beep\n");
	return 0;
}

DDECOMMANDINFO const c_sDDECommands[] =
{
    { c_szCreateGroup  , DDE_CreateGroup   },
    { c_szShowGroup    , DDE_ShowGroup     },
    { c_szAddItem      , DDE_AddItem       },
    { c_szExitProgman  , DDE_ExitProgman   },
    { c_szDeleteGroup  , DDE_DeleteGroup   },
    { c_szDeleteItem   , DDE_DeleteItem    },
    { c_szReplaceItem  , DDE_ReplaceItem   },
    { c_szReload       , DDE_Reload        },
    { c_szViewFolder   , DDE_ViewFolder    },
    { c_szExploreFolder, DDE_ExploreFolder },
    { c_szFindFolder,    DDE_FindFolder    },
    { c_szOpenFindFile,  DDE_OpenFindFile  },
    { c_szRUCabinet,     DDE_ConfirmID},
    { c_szShellFile,     DDE_ShellFile},
#ifdef DEBUG
    { c_szBeep         , DDE_Beep          },
#endif
    { 0, 0 },
};

LPTSTR SkipWhite(LPTSTR lpsz)
{
    /* prevent sign extension in case of DBCS */
    while (*lpsz && (TUCHAR)*lpsz <= TEXT(' '))
        lpsz++;

    return(lpsz);
}


LPTSTR GetCommandName(LPTSTR lpCmd, const DDECOMMANDINFO* lpsCommands, UINT* lpW)
{
    TCHAR chT;
    UINT iCmd = 0;
    LPTSTR lpT;

    /* Eat any white space. */
    lpT = lpCmd = SkipWhite(lpCmd);

    /* Find the end of the token. */
    while (IsCharAlpha(*lpCmd))
        lpCmd = CharNext(lpCmd);

    /* Temporarily NULL terminate it. */
    chT = *lpCmd;
    *lpCmd = 0;

    /* Look up the token in a list of commands. */
    *lpW = (UINT)-1;
    while (lpsCommands->pszCommand)
    {
        if (!lstrcmpi(lpsCommands->pszCommand, lpT))
        {
            *lpW = iCmd;
            break;
        }
        iCmd++;
        ++lpsCommands;
    }

    *lpCmd = chT;

    return(lpCmd);
}

HRESULT VariantChangeTypeForRead(VARIANT* pvar, VARTYPE vtDesired)
{
    HRESULT hr = S_OK;

    if ((pvar->vt != vtDesired) && (vtDesired != VT_EMPTY))
    {
        VARIANT varTmp;
        VARIANT varSrc;

        // cache a copy of [in]pvar in varSrc - we will free this later
        CopyMemory(&varSrc, pvar, sizeof(varTmp));
        VARIANT* pvarToCopy = &varSrc;

        // oleaut's VariantChangeType doesn't support
        // hex number string -> number conversion, which we want,
        // so convert those to another format they understand.
        //
        // if we're in one of these cases, varTmp will be initialized
        // and pvarToCopy will point to it instead
        //
        if (VT_BSTR == varSrc.vt)
        {
            switch (vtDesired)
            {
            case VT_I1:
            case VT_I2:
            case VT_I4:
            case VT_INT:
            {
                if (StrToIntExW(varSrc.bstrVal, STIF_SUPPORT_HEX, &varTmp.intVal))
                {
                    varTmp.vt = VT_INT;
                    pvarToCopy = &varTmp;
                }
                break;
            }

            case VT_UI1:
            case VT_UI2:
            case VT_UI4:
            case VT_UINT:
            {
                if (StrToIntExW(varSrc.bstrVal, STIF_SUPPORT_HEX, (int*)&varTmp.uintVal))
                {
                    varTmp.vt = VT_UINT;
                    pvarToCopy = &varTmp;
                }
                break;
            }
            }
        }

        // clear our [out] buffer, in case VariantChangeType fails
        VariantInit(pvar);

        hr = VariantChangeType(pvar, pvarToCopy, 0, vtDesired);

        // clear the cached [in] value
        VariantClear(&varSrc);
        // if initialized, varTmp is VT_UINT or VT_UINT, neither of which need VariantClear
    }

    return hr;
}

LPTSTR GetOneParameter(LPCTSTR lpCmdStart, LPTSTR lpCmd,
    UINT* lpW, BOOL fIncludeQuotes)
{
    LPTSTR     lpT;

    switch (*lpCmd)
    {
    case TEXT(','):
        *lpW = (UINT)(lpCmd - lpCmdStart);  // compute offset
        *lpCmd++ = 0;                /* comma: becomes a NULL string */
        break;

    case TEXT('"'):
        if (fIncludeQuotes)
        {

            // quoted string... don't trim off "
            *lpW = (UINT)(lpCmd - lpCmdStart);  // compute offset
            ++lpCmd;
            while (*lpCmd && *lpCmd != TEXT('"'))
                lpCmd = CharNext(lpCmd);
            if (!*lpCmd)
                return(NULL);
            lpT = lpCmd;
            ++lpCmd;

            goto skiptocomma;
        }
        else
        {
            // quoted string... trim off "
            ++lpCmd;
            *lpW = (UINT)(lpCmd - lpCmdStart);  // compute offset
            while (*lpCmd && *lpCmd != TEXT('"'))
                lpCmd = CharNext(lpCmd);
            if (!*lpCmd)
                return(NULL);
            *lpCmd++ = 0;
            lpCmd = SkipWhite(lpCmd);

            // If there's a comma next then skip over it, else just go on as
            // normal.
            if (*lpCmd == TEXT(','))
                lpCmd++;
        }
        break;

    case TEXT(')'):
        return(lpCmd);                /* we ought not to hit this */

    case TEXT('('):
    case TEXT('['):
    case TEXT(']'):
        return(NULL);                 /* these are illegal */

    default:
        lpT = lpCmd;
        *lpW = (UINT)(lpCmd - lpCmdStart);  // compute offset
    skiptocomma:
        while (*lpCmd && *lpCmd != TEXT(',') && *lpCmd != TEXT(')'))
        {
            /* Check for illegal characters. */
            if (*lpCmd == TEXT(']') || *lpCmd == TEXT('[') || *lpCmd == TEXT('('))
                return(NULL);

            /* Remove trailing whitespace */
            /* prevent sign extension */
            if ((TUCHAR)*lpCmd > TEXT(' '))
                lpT = lpCmd;

            lpCmd = CharNext(lpCmd);
        }

        /* Eat any trailing comma. */
        if (*lpCmd == TEXT(','))
            lpCmd++;

        /* NULL terminator after last nonblank character -- may write over
         * terminating ')' but the caller checks for that because this is
         * a hack.
         */

#ifdef UNICODE
        lpT[1] = 0;
#else
        lpT[IsDBCSLeadByte(*lpT) ? 2 : 1] = 0;
#endif
        break;
    }

    // Return next unused character.
    return(lpCmd);
}

UINT* GetDDECommands(LPTSTR lpCmd, const DDECOMMANDINFO* lpsCommands, BOOL fLFN)
{
    UINT cParm, cCmd = 0;
    UINT* lpW;
    UINT* lpRet;
    LPCTSTR lpCmdStart = lpCmd;
    BOOL fIncludeQuotes = FALSE;

    lpRet = lpW = (UINT*)GlobalAlloc(GPTR, 512L);
    if (!lpRet)
        return 0;

    while (*lpCmd)
    {
        /* Skip leading whitespace. */
        lpCmd = SkipWhite(lpCmd);

        /* Are we at a NULL? */
        if (!*lpCmd)
        {
            /* Did we find any commands yet? */
            if (cCmd)
                goto GDEExit;
            else
                goto GDEErrExit;
        }

        /* Each command should be inside square brackets. */
        if (*lpCmd != TEXT('['))
            goto GDEErrExit;
        lpCmd++;

        /* Get the command name. */
        lpCmd = GetCommandName(lpCmd, lpsCommands, lpW);
        if (*lpW == (UINT)-1)
            goto GDEErrExit;

        // We need to leave quotes in for the first param of an AddItem.
        if (fLFN && *lpW == 2)
        {
            fIncludeQuotes = TRUE;
        }

        lpW++;

        /* Start with zero parms. */
        cParm = 0;
        lpCmd = SkipWhite(lpCmd);

        /* Check for opening '(' */
        if (*lpCmd == TEXT('('))
        {
            lpCmd++;

            /* Skip white space and then find some parameters (may be none). */
            lpCmd = SkipWhite(lpCmd);

            while (*lpCmd != TEXT(')'))
            {
                if (!*lpCmd)
                    goto GDEErrExit;

                // Only the first param of the AddItem command needs to
                // handle quotes from LFN guys.
                if (fIncludeQuotes && (cParm != 0))
                    fIncludeQuotes = FALSE;

                /* Get the parameter. */
                if (!(lpCmd = GetOneParameter(lpCmdStart, lpCmd, lpW + (++cParm), fIncludeQuotes)))
                    goto GDEErrExit;

                /* HACK: Did GOP replace a ')' with a NULL? */
                if (!*lpCmd)
                    break;

                /* Find the next one or ')' */
                lpCmd = SkipWhite(lpCmd);
            }

            // Skip closing bracket.
            lpCmd++;

            /* Skip the terminating stuff. */
            lpCmd = SkipWhite(lpCmd);
        }

        /* Set the count of parameters and then skip the parameters. */
        *lpW++ = cParm;
        lpW += cParm;

        /* We found one more command. */
        cCmd++;

        /* Commands must be in square brackets. */
        if (*lpCmd != TEXT(']'))
            goto GDEErrExit;
        lpCmd++;
    }

GDEExit:
    /* Terminate the command list with -1. */
    *lpW = (UINT)-1;

    return lpRet;

GDEErrExit:
    GlobalFree(lpW);
    return(0);
}

//LPNMVIEWFOLDER DDECreatePostNotify(LPNMVIEWFOLDER pnm)
//{
//    LPNMVIEWFOLDER pnmPost = NULL;
//    TCHAR szCmd[MAX_PATH * 2];
//    StrCpyN(szCmd, pnm->szCmd, SIZECHARS(szCmd));
//    UINT* pwCmd = GetDDECommands(szCmd, c_sDDECommands, FALSE);
//
//    // -1 means there aren't any commands we understand.  Oh, well
//    if (pwCmd && (-1 != *pwCmd))
//    {
//        LPCTSTR pszCommand = c_sDDECommands[*pwCmd].pszCommand;
//
//        //
//        //  these are the only commands handled by a PostNotify
//        if (pszCommand == TEXT("ViewFolder")
//            || pszCommand == TEXT("ExploreFolder")
//            || pszCommand == TEXT("ShellFile"))
//        {
//            pnmPost = (LPNMVIEWFOLDER)LocalAlloc(LPTR, sizeof(NMVIEWFOLDER));
//
//            if (pnmPost)
//                memcpy(pnmPost, pnm, sizeof(NMVIEWFOLDER));
//        }
//
//        GlobalFree(pwCmd);
//    }
//
//    return pnmPost;
//}

// Helper function to convert passed in command parameters into the
// appropriate Id list
LPITEMIDLIST _GetPIDLFromDDEArgs(UINT nArg, LPTSTR lpszBuf, UINT* lpwCmd, PSHDDEERR psde, LPCITEMIDLIST* ppidlGlobal)
{
    LPTSTR lpsz;
    LPITEMIDLIST pidl = NULL;

    // Switch from 0-based to 1-based 
    ++nArg;
    if (*lpwCmd < nArg)
    {
        return NULL;
    }

    // Skip to the right argument
    lpwCmd += nArg;
    lpsz = &lpszBuf[*lpwCmd];

    // REVIEW: all associations will go through here.  this
    // is probably not what we want for normal cmd line type operations

    // A colon at the begining of the path means that this is either
    // a pointer to a pidl (win95 classic) or a handle:pid (all other
    // platforms including win95+IE4).  Otherwise, it's a regular path.

    if (lpsz[0] == TEXT(':'))
    {
        HANDLE hMem;
        DWORD  dwProcId;
        LPTSTR pszNextColon;

        // Convert the string into a pidl.

        hMem = LongToHandle(StrToLong((LPTSTR)(lpsz + 1)));
        pszNextColon = StrChr(lpsz + 1, TEXT(':'));
        if (pszNextColon)
        {
            LPITEMIDLIST pidlShared;

            dwProcId = (DWORD)StrToLong(pszNextColon + 1);
            pidlShared = (LPITEMIDLIST)SHLockShared(hMem, dwProcId);
            if (pidlShared && !IsBadReadPtr(pidlShared, 1))
            {
                pidl = ILClone(pidlShared);
                SHUnlockShared(pidlShared);
            }
            else
            {
            }
            SHFreeShared(hMem, dwProcId);
        }
        else if (hMem && !IsBadReadPtr(hMem, sizeof(WORD)))
        {
            // this is likely to be browser only mode on win95 with the old pidl arguments which is
            // going to be in shared memory.... (must be cloned into local memory)...
            pidl = ILClone((LPITEMIDLIST)hMem);

            // this will get freed if we succeed.
            *ppidlGlobal = (LPITEMIDLIST)hMem;
        }

        return pidl;
    }
    else
    {
        TCHAR tszQual[MAX_PATH];

        // We must copy to a temp buffer because the PathQualify may
        // result in a string longer than our input buffer and faulting
        // seems like a bad way of handling that situation.
        lstrcpyn(tszQual, lpsz, ARRAYSIZE(tszQual));
        lpsz = tszQual;

        /* Fuck you
        // Is this a URL?
        if (!PathIsURL(lpsz))
        {
            // No; qualify it
            PathQualifyDef(lpsz, NULL, PQD_NOSTRIPDOTS);
        }
        */ 
        pidl = ILCreateFromPath(lpsz);

        if (pidl == NULL && psde)
        {
            psde->idMsg = IDS_CANTFINDDIR;
            lstrcpyn(psde->szParam, lpsz, ARRAYSIZE(psde->szParam));
        }
        return pidl;
    }
}

LPITEMIDLIST GetPIDLFromDDEArgs(LPTSTR lpszBuf, UINT* lpwCmd, PSHDDEERR psde, LPCITEMIDLIST* ppidlGlobal)
{
    LPITEMIDLIST pidl = _GetPIDLFromDDEArgs(1, lpszBuf, lpwCmd, psde, ppidlGlobal);
    if (!pidl)
    {
        pidl = _GetPIDLFromDDEArgs(0, lpszBuf, lpwCmd, psde, ppidlGlobal);
    }

    return pidl;
}

//BOOL DDEHandleViewFolderNotify(IShellBrowser* psb, HWND hwnd, LPNMVIEWFOLDER pnm)
//{
//    BOOL fRet = FALSE;
//    UINT* pwCmd = GetDDECommands(pnm->szCmd, c_sDDECommands, FALSE);
//
//    // -1 means there aren't any commands we understand.  Oh, well
//    if (pwCmd && (-1 != *pwCmd))
//    {
//        UINT* pwCmdSave = pwCmd;
//        UINT c = *pwCmd++;
//
//        LPCTSTR pszCommand = c_sDDECommands[c].pszCommand;
//
//        if (pszCommand == c_szViewFolder ||
//            pszCommand == c_szExploreFolder)
//        {
//            //fRet = DoDDE_ViewFolder(psb, hwnd, pnm->szCmd, pwCmd,
//                //pszCommand == c_szExploreFolder, pnm->dwHotKey, pnm->hMonitor);
//        }
//        else if (pszCommand == c_szShellFile)
//        {
//            fRet = DDE_ShellFile(pnm->szCmd, pwCmd, 0);
//        }
//
//        GlobalFree(pwCmdSave);
//    }
//
//    return fRet;
//}

TCHAR SHFindMnemonic(LPCTSTR psz)
{
    ASSERT(psz);
    TCHAR tchDefault = *psz;                // Default is first character
    LPCTSTR pszAmp;

    while ((pszAmp = StrChr(psz, TEXT('&'))) != NULL)
    {
        switch (pszAmp[1])
        {
        case TEXT('&'):         // Skip over &&
            psz = pszAmp + 2;
            continue;

        case TEXT('\0'):        // Ignore trailing ampersand
            return tchDefault;

        default:
            return pszAmp[1];
        }
    }
    return tchDefault;
}

HRESULT SHBindToIDListParent(LPCITEMIDLIST pidl, REFIID riid, void** ppv, LPCITEMIDLIST* ppidlLast)
{
    return SHBindToFolderIDListParent(NULL, pidl, riid, ppv, ppidlLast);
}

CHAR CharUpperCharA(CHAR c)
{
    // XXX(isabella): MSVC seems to have a bug, presumably with inlining, where the input to
    // this function, and therefore argument passed through to CharUpperA, is not shortened
    // down to the length of a single byte, and therefore CharUpperA receives unwanted data
    // which tends to get interpreted as a pointer.
    // 
    // This usually causes an access violation. It is entirely undesirable.
    //
    // The fix is simple: always ensure that the upper byte is 0. The mask "& 0xff" is enough
    // for this purpose.
    return (CHAR)(DWORD_PTR)CharUpperA((LPSTR)(DWORD_PTR)(c & 0x00ff));
}

WCHAR CharUpperCharW(WCHAR c)
{
    return (WCHAR)(DWORD_PTR)CharUpperW((LPWSTR)(DWORD_PTR)(c));
}

//  the last word of the pidl is where we store the hidden offset
#define _ILHiddenOffset(pidl)   (*((WORD UNALIGNED *)(((BYTE *)_ILNext(pidl)) - sizeof(WORD))))
#define _ILSetHiddenOffset(pidl, cb)   ((*((WORD UNALIGNED *)(((BYTE *)_ILNext(pidl)) - sizeof(WORD)))) = (WORD)cb)
#define _ILIsHidden(pidhid)     (HIWORD(pidhid->id) == HIWORD(IDLHID_EMPTY))

PCIDHIDDEN _ILNextHidden(PCIDHIDDEN pidhid, LPCITEMIDLIST pidlLimit)
{
    PCIDHIDDEN pidhidNext = (PCIDHIDDEN)_ILNext((LPCITEMIDLIST)pidhid);

    if ((BYTE*)pidhidNext < (BYTE*)pidlLimit && _ILIsHidden(pidhidNext))
    {
        return pidhidNext;
    }

    //  if we ever go past the limit,
    //  then this is not really a hidden id
    //  or we have messed up on some calculation.
    ASSERT((BYTE*)pidhidNext == (BYTE*)pidlLimit);
    return NULL;
}
PCIDHIDDEN _ILFirstHidden(LPCITEMIDLIST pidl)
{
    WORD cbHidden = _ILHiddenOffset(pidl);

    if (cbHidden && cbHidden + sizeof(HIDDENITEMID) < pidl->mkid.cb)
    {
        //  this means it points to someplace inside the pidl
        //  maybe this has hidden ids
        PCIDHIDDEN pidhid = (PCIDHIDDEN)(((BYTE*)pidl) + cbHidden);

        if (_ILIsHidden(pidhid)
            && (pidhid->cb + cbHidden <= pidl->mkid.cb))
        {
            //  this is more than likely a hidden id
            //  we could walk the chain and verify
            //  that it adds up right...
            return pidhid;
        }
    }

    return NULL;
}

PCIDHIDDEN ILFindHiddenIDOn(LPCITEMIDLIST pidl, IDLHID id, BOOL fOnLast)
{
    if (!ILIsEmpty(pidl))
    {
        if (fOnLast)
            pidl = ILFindLastID(pidl);

        PCIDHIDDEN pidhid = _ILFirstHidden(pidl);

        //  reuse pidl to become the limit.
        //  so that we cant ever walk out of 
        //  the pidl.
        pidl = _ILNext(pidl);

        while (pidhid)
        {
            if (pidhid->id == id)
                break;

            pidhid = _ILNextHidden(pidhid, pidl);
        }
        return pidhid;
    }

    return NULL;
}

BOOL ILRemoveHiddenID(LPITEMIDLIST pidl, IDLHID id)
{
    PIDHIDDEN pidhid = (PIDHIDDEN)ILFindHiddenID(pidl, id);

    if (pidhid)
    {
        pidhid->id = IDLHID_EMPTY;
        return TRUE;
    }
    return FALSE;
}

void ILExpungeRemovedHiddenIDs(LPITEMIDLIST pidl)
{
    if (pidl)
    {
        pidl = ILFindLastID(pidl);

        // Note: Each IDHIDDEN has a WORD appended to it, equal to
        // _ILHiddenOffset, so we can just keep deleting IDHIDDENs
        // and if we delete them all, everything is cleaned up; if
        // there are still unremoved IDHIDDENs left, they will provide
        // the _ILHiddenOffset.

        PIDHIDDEN pidhid;
        BOOL fAnyDeleted = FALSE;
        while ((pidhid = (PIDHIDDEN)ILFindHiddenID(pidl, IDLHID_EMPTY)) != NULL)
        {
            fAnyDeleted = TRUE;
            LPBYTE pbAfter = (LPBYTE)pidhid + pidhid->cb;
            WORD cbDeleted = pidhid->cb;
            MoveMemory(pidhid, pbAfter,
                (LPBYTE)pidl + pidl->mkid.cb + sizeof(WORD) - pbAfter);
            pidl->mkid.cb -= cbDeleted;
        }
    }
}

HRESULT ILCloneWithHiddenID(LPCITEMIDLIST pidl, PCIDHIDDEN pidhid, LPITEMIDLIST* ppidl)
{
    HRESULT hr;

    // If this ASSERT fires, then the caller did not set the pidhid->id
    // value properly.  For example, the packing settings might be incorrect.

    ASSERT(_ILIsHidden(pidhid));

    if (ILIsEmpty(pidl))
    {
        *ppidl = NULL;
        hr = E_INVALIDARG;
    }
    else
    {
        UINT cbUsed = ILGetSize(pidl);
        UINT cbRequired = cbUsed + pidhid->cb + sizeof(pidhid->cb);

        *ppidl = (LPITEMIDLIST)SHAlloc(cbRequired);
        if (*ppidl)
        {
            hr = S_OK;

            CopyMemory(*ppidl, pidl, cbUsed);

            LPITEMIDLIST pidlLast = ILFindLastID(*ppidl);
            WORD cbHidden = _ILFirstHidden(pidlLast) ? _ILHiddenOffset(pidlLast) : pidlLast->mkid.cb;
            PIDHIDDEN pidhidCopy = (PIDHIDDEN)_ILSkip(*ppidl, cbUsed - sizeof((*ppidl)->mkid.cb));

            // Append it, overwriting the terminator
            MoveMemory(pidhidCopy, pidhid, pidhid->cb);

            //  grow the copy to allow the hidden offset.
            pidhidCopy->cb += sizeof(pidhid->cb);

            //  now we need to readjust pidlLast to encompass 
            //  the hidden bits and the hidden offset.
            pidlLast->mkid.cb += pidhidCopy->cb;

            //  set the hidden offset so that we can find our hidden IDs later
            _ILSetHiddenOffset((LPITEMIDLIST)pidhidCopy, cbHidden);

            // We must put zero-terminator because of LMEM_ZEROINIT.
            _ILSkip(*ppidl, cbRequired - sizeof((*ppidl)->mkid.cb))->mkid.cb = 0;
            ASSERT(ILGetSize(*ppidl) == cbRequired);
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }
    }
    return hr;
}

LPITEMIDLIST ILAppendHiddenID(LPITEMIDLIST pidl, PCIDHIDDEN pidhid)
{
    //
    // FEATURE - we dont handle collisions of multiple hidden ids
    //          maybe remove IDs of the same IDLHID?
    //
    // Note: We do not remove IDLHID_EMPTY hidden ids.
    // Callers need to call ILExpungeRemovedHiddenIDs explicitly
    // if they want empty hidden ids to be compressed out.
    //

    if (!ILIsEmpty(pidl))
    {
        LPITEMIDLIST pidlSave = pidl;
        ILCloneWithHiddenID(pidl, pidhid, &pidl);
        ILFree(pidlSave);
    }
    return pidl;
}

HRESULT SHILCombine(LPCITEMIDLIST pidl1, LPCITEMIDLIST pidl2, LPITEMIDLIST* ppidlOut)
{
    *ppidlOut = ILCombine(pidl1, pidl2);
    return *ppidlOut ? S_OK : E_OUTOFMEMORY;
}


// NOTE: possibly we need a csidlSkip param to avoid recursion?
BOOL _ReparentAliases(HWND hwnd, HANDLE hToken, LPCITEMIDLIST pidl, LPITEMIDLIST* ppidlAlias, DWORD dwXlateAliases)
{
    static const struct { DWORD dwXlate; int idPath; int idAlias; BOOL fCommon; } s_rgidAliases[] =
    {
        { XLATEALIAS_MYDOCS, CSIDL_PERSONAL | CSIDL_FLAG_NO_ALIAS, CSIDL_PERSONAL, FALSE},
        { XLATEALIAS_COMMONDOCS, CSIDL_COMMON_DOCUMENTS | CSIDL_FLAG_NO_ALIAS, CSIDL_COMMON_DOCUMENTS, FALSE},
        { XLATEALIAS_DESKTOP, CSIDL_DESKTOPDIRECTORY, CSIDL_DESKTOP, FALSE},
        { XLATEALIAS_DESKTOP, CSIDL_COMMON_DESKTOPDIRECTORY, CSIDL_DESKTOP, TRUE},
    };
    BOOL fContinue = TRUE;
    *ppidlAlias = NULL;

    for (int i = 0; fContinue && i < ARRAYSIZE(s_rgidAliases); i++)
    {
        LPITEMIDLIST pidlPath;
        if ((dwXlateAliases & s_rgidAliases[i].dwXlate) &&
            (S_OK == SHGetFolderLocation(hwnd, s_rgidAliases[i].idPath, hToken, 0, &pidlPath)))
        {
            LPCITEMIDLIST pidlChild = ILFindChild(pidlPath, pidl);
            if (pidlChild)
            {
                //  ok we need to use the alias instead of the path
                LPITEMIDLIST pidlAlias;
                if (S_OK == SHGetFolderLocation(hwnd, s_rgidAliases[i].idAlias, hToken, 0, &pidlAlias))
                {
                    if (SUCCEEDED(SHILCombine(pidlAlias, pidlChild, ppidlAlias)))
                    {
                        if (s_rgidAliases[i].fCommon && !ILIsEmpty(*ppidlAlias))
                        {
                            // find the size of the special part (subtacting for null pidl terminator)
                            UINT cbSize = ILGetSize(pidlAlias) - sizeof(pidlAlias->mkid.cb);
                            LPITEMIDLIST pidlChildFirst = _ILSkip(*ppidlAlias, cbSize);

                            // We set the first ID under the common path to have the SHID_FS_COMMONITEM so that when we bind we
                            // can hand this to the proper merged psf
                            pidlChildFirst->mkid.abID[0] |= SHID_FS_COMMONITEM;
                        }
                        ILFree(pidlAlias);
                    }
                    fContinue = FALSE;
                }
            }
            ILFree(pidlPath);
        }
    }

    return (*ppidlAlias != NULL);
}


HRESULT SHILAliasTranslate(LPCITEMIDLIST pidl, LPITEMIDLIST* ppidlAlias, DWORD dwXlateAliases)
{
    return _ReparentAliases(NULL, NULL, pidl, ppidlAlias, dwXlateAliases) ? S_OK : E_FAIL;
}

LPITEMIDLIST SHLogILFromFSIL(LPCITEMIDLIST pidlFS)
{
    LPITEMIDLIST pidlOut;
    SHILAliasTranslate(pidlFS, &pidlOut, XLATEALIAS_ALL); // will set pidlOut=NULL on failure
    return pidlOut;
}

HRESULT DataObj_SetGlobal(IDataObject* pdtobj, UINT cf, HGLOBAL hGlobal)
{
    FORMATETC fmte = { (CLIPFORMAT)cf, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM medium = { 0 };

    medium.tymed = TYMED_HGLOBAL;
    medium.hGlobal = hGlobal;
    medium.pUnkForRelease = NULL;

    // give the data object ownership of ths
    return pdtobj->SetData(&fmte, &medium, TRUE);
}

BOOL GetInfoTip(IShellFolder* psf, LPCITEMIDLIST pidl, LPTSTR pszText, int cchTextMax)
{
    *pszText = 0;
    return FALSE;
}

HRESULT SHGetUIObjectFromFullPIDL(LPCITEMIDLIST pidl, HWND hwnd, REFIID riid, void** ppv)
{
    *ppv = NULL;

    LPCITEMIDLIST pidlChild;
    IShellFolder* psf;
    HRESULT hr = SHBindToIDListParent(pidl, IID_PPV_ARG(IShellFolder, &psf), &pidlChild);
    if (SUCCEEDED(hr))
    {
        hr = psf->GetUIObjectOf(hwnd, 1, &pidlChild, riid, NULL, ppv);
        psf->Release();
    }

    return hr;
}


//
// This is a helper function for finding a specific verb's index in a context menu
//
UINT GetMenuIndexForCanonicalVerb(HMENU hMenu, IContextMenu* pcm, UINT idCmdFirst, LPCWSTR pwszVerb)
{
    int cMenuItems = GetMenuItemCount(hMenu);
    int iItem;
    for (iItem = 0; iItem < cMenuItems; iItem++)
    {
        MENUITEMINFO mii = { 0 };

        mii.cbSize = sizeof(mii);
        mii.fMask = MIIM_TYPE | MIIM_ID;

        // IS_INTRESOURCE guards against mii.wID == -1 **and** against
        // buggy shell extensions which set their menu item IDs out of range.
        if (GetMenuItemInfo(hMenu, iItem, MF_BYPOSITION, &mii) &&
            !(mii.fType & MFT_SEPARATOR) && IS_INTRESOURCE(mii.wID) &&
            (mii.wID >= idCmdFirst))
        {
            union {
                WCHAR szItemNameW[80];
                char szItemNameA[80];
            };
            CHAR aszVerb[80];

            // try both GCS_VERBA and GCS_VERBW in case it only supports one of them
            SHUnicodeToAnsi(pwszVerb, aszVerb, ARRAYSIZE(aszVerb));

            if (SUCCEEDED(pcm->GetCommandString(mii.wID - idCmdFirst, GCS_VERBA, NULL, szItemNameA, ARRAYSIZE(szItemNameA))))
            {
                if (StrCmpICA(szItemNameA, aszVerb) == 0)
                {
                    break;  // found it
                }
            }
            else
            {
                if (SUCCEEDED(pcm->GetCommandString(mii.wID - idCmdFirst, GCS_VERBW, NULL, (LPSTR)szItemNameW, ARRAYSIZE(szItemNameW))) &&
                    (StrCmpICW(szItemNameW, pwszVerb) == 0))
                {
                    break;  // found it
                }
            }
        }

    }
    if (iItem == cMenuItems)
    {
        iItem = -1; // went through all the menuitems and didn't find it
    }
    
    return iItem;
}

HRESULT ContextMenu_GetCommandStringVerb(IContextMenu* pcm, UINT idCmd, LPWSTR pszVerb, int cchVerb)
{
    // Ulead SmartSaver Pro has a 60 character verb, and 
    // over writes out stack, ignoring the cch param and we fault. 
    // so make sure this buffer is at least 60 chars

    TCHAR wszVerb[64];
    wszVerb[0] = 0;

    HRESULT hr = pcm->GetCommandString(idCmd, GCS_VERBW, NULL, (LPSTR)wszVerb, ARRAYSIZE(wszVerb));
    if (FAILED(hr))
    {
        // be extra paranoid about requesting the ansi version -- we've
        // found IContextMenu implementations that return a UNICODE buffer
        // even though we ask for an ANSI string on NT systems -- hopefully
        // they will have answered the above request already, but just in
        // case let's not let them overrun our stack!
        char szVerbAnsi[128];
        hr = pcm->GetCommandString(idCmd, GCS_VERBA, NULL, szVerbAnsi, ARRAYSIZE(szVerbAnsi) / 2);
        if (SUCCEEDED(hr))
        {
            SHAnsiToUnicode(szVerbAnsi, wszVerb, ARRAYSIZE(wszVerb));
        }
    }

    StrCpyNW(pszVerb, wszVerb, cchVerb);

    return hr;
}

HRESULT ContextMenu_DeleteCommandByName(IContextMenu* pcm, HMENU hpopup, UINT idFirst, LPCWSTR pszCommand)
{
    UINT ipos = GetMenuIndexForCanonicalVerb(hpopup, pcm, idFirst, pszCommand);
    if (ipos != -1)
    {
        DeleteMenu(hpopup, ipos, MF_BYPOSITION);
        return S_OK;
    }
    else
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }
}

#define OCFMAPPING(ocf)     {OBJCOMPATF_##ocf, TEXT(#ocf)}


DWORD _GetMappedFlags(HKEY hk, const FLAGMAP* pmaps, DWORD cmaps)
{
    DWORD dwRet = 0;
    for (DWORD i = 0; i < cmaps; i++)
    {
        if (NOERROR == SHGetValue(hk, NULL, pmaps[i].psz, NULL, NULL, NULL))
            dwRet |= pmaps[i].flag;
    }

    return dwRet;
}

DWORD _GetRegistryObjectCompatFlags(REFGUID clsid)
{
    DWORD dwRet = 0;
    TCHAR szGuid[GUIDSTR_MAX];
    TCHAR sz[MAX_PATH];
    HKEY hk;

    using fnSHStringFromGUID = int(WINAPI*)(REFGUID, LPWSTR, int);
    fnSHStringFromGUID SHStringFromGUID;
    SHStringFromGUID = reinterpret_cast<fnSHStringFromGUID>(GetProcAddress(GetModuleHandle(L"shlwapi.dll"), MAKEINTRESOURCEA(24)));
    SHStringFromGUID(clsid, szGuid, ARRAYSIZE(szGuid));
    wnsprintf(sz, ARRAYSIZE(sz), TEXT("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\ShellCompatibility\\Objects\\%s"), szGuid);

    if (NOERROR == RegOpenKeyEx(HKEY_LOCAL_MACHINE, sz, 0, KEY_QUERY_VALUE, &hk))
    {
        static const FLAGMAP rgOcfMaps[] = {
            OCFMAPPING(OTNEEDSSFCACHE),
            OCFMAPPING(NO_WEBVIEW),
            OCFMAPPING(UNBINDABLE),
            OCFMAPPING(PINDLL),
            OCFMAPPING(NEEDSFILESYSANCESTOR),
            OCFMAPPING(NOTAFILESYSTEM),
            OCFMAPPING(CTXMENU_NOVERBS),
            OCFMAPPING(CTXMENU_LIMITEDQI),
            OCFMAPPING(COCREATESHELLFOLDERONLY),
            OCFMAPPING(NEEDSSTORAGEANCESTOR),
            OCFMAPPING(NOLEGACYWEBVIEW),
        };

        dwRet = _GetMappedFlags(hk, rgOcfMaps, ARRAYSIZE(rgOcfMaps));
        RegCloseKey(hk);
    }

    return dwRet;
}

typedef struct _CLSIDCOMPAT
{
    const GUID* pclsid;
    OBJCOMPATFLAGS flags;
}CLSIDCOMPAT, * PCLSIDCOMPAT;

OBJCOMPATFLAGS SHGetObjectCompatFlags(IUnknown* punk, const CLSID* pclsid)
{
    HRESULT hr = E_INVALIDARG;
    OBJCOMPATFLAGS ocf = 0;
    CLSID clsid;
    if (punk)
        hr = IUnknown_GetClassID(punk, &clsid);
    else if (pclsid)
    {
        clsid = *pclsid;
        hr = S_OK;
    }

    if (SUCCEEDED(hr))
    {
        static const CLSIDCOMPAT s_rgCompat[] =
        {
            {&CLSID_WS_FTP_PRO_EXPLORER,
                OBJCOMPATF_OTNEEDSSFCACHE | OBJCOMPATF_PINDLL },
            {&CLSID_WS_FTP_PRO,
                OBJCOMPATF_UNBINDABLE},
            {&GUID_AECOZIPARCHIVE,
                OBJCOMPATF_OTNEEDSSFCACHE | OBJCOMPATF_NO_WEBVIEW},
            {&CLSID_KODAK_DC260_ZOOM_CAMERA,
                OBJCOMPATF_OTNEEDSSFCACHE | OBJCOMPATF_PINDLL},
            {&GUID_MACINDOS,
                OBJCOMPATF_NO_WEBVIEW},
            {&CLSID_EasyZIP,
                OBJCOMPATF_NO_WEBVIEW},
            {&CLSID_PAGISPRO_FOLDER,
                OBJCOMPATF_NEEDSFILESYSANCESTOR},
            {&CLSID_FILENET_IDMDS_NEIGHBORHOOD,
                OBJCOMPATF_NOTAFILESYSTEM},
            {&CLSID_NOVELLX,
                OBJCOMPATF_PINDLL},
            {&CLSID_PGP50_CONTEXTMENU,
                OBJCOMPATF_CTXMENU_LIMITEDQI},
            {&CLSID_QUICKFINDER_CONTEXTMENU,
                OBJCOMPATF_CTXMENU_NOVERBS},
            {&CLSID_HERCULES_HCTNT_V1001,
                OBJCOMPATF_PINDLL},
                //
                //  WARNING DONT ADD NEW COMPATIBILITY HERE - ZekeL - 18-OCT-99
                //  Add new entries to the registry.  each component 
                //  that needs compatibility flags should register 
                //  during selfregistration.  (see the RegExternal
                //  section of selfreg.inx in shell32 for an example.)  
                //  all new flags should be added to the FLAGMAP array.
                //
                //  the register under:
                //  HKLM\SW\MS\Win\CV\ShellCompatibility\Objects
                //      \{CLSID}
                //          FLAGNAME    //  requires no value
                //
                //  NOTE: there is no version checking
                //  but we could add it as the data attached to 
                //  the flags, and compare with the version 
                //  of the LocalServer32 dll.
                //  
                {NULL, 0}
        };

        for (int i = 0; s_rgCompat[i].pclsid; i++)
        {
            if (IsEqualGUID(clsid, *(s_rgCompat[i].pclsid)))
            {
                //  we could check version based
                //  on what is in under HKCR\CLSID\{clsid}
                ocf = s_rgCompat[i].flags;
                break;
            }
        }

        ocf |= _GetRegistryObjectCompatFlags(clsid);

    }

    return ocf;
}
// deals with goofyness of IShellFolder::GetAttributesOf() including 
//      in/out param issue
//      failures
//      goofy cast for 1 item case
//      masks off results to only return what you asked for

DWORD SHGetAttributes(IShellFolder* psf, LPCITEMIDLIST pidl, DWORD dwAttribs)
{
    // like SHBindToObject, if psf is NULL, use absolute pidl
    LPCITEMIDLIST pidlChild;
    if (!psf)
    {
        SHBindToParent(pidl, IID_PPV_ARG(IShellFolder, &psf), &pidlChild);
    }
    else
    {
        psf->AddRef();
        pidlChild = pidl;
    }

    DWORD dw = 0;
    if (psf)
    {
        dw = dwAttribs;
        dw = SUCCEEDED(psf->GetAttributesOf(1, (LPCITEMIDLIST*)&pidlChild, &dw)) ? (dwAttribs & dw) : 0;
        if ((dw & SFGAO_FOLDER) && (dw & SFGAO_CANMONIKER) && !(dw & SFGAO_STORAGEANCESTOR) && (dwAttribs & SFGAO_STORAGEANCESTOR))
        {
            if (OBJCOMPATF_NEEDSSTORAGEANCESTOR & SHGetObjectCompatFlags(psf, NULL))
            {
                //  switch SFGAO_CANMONIKER -> SFGAO_STORAGEANCESTOR
                dw |= SFGAO_STORAGEANCESTOR;
                dw &= ~SFGAO_CANMONIKER;
            }
        }
    }

    if (psf)
    {
        psf->Release();
    }

    return dw;
}

HRESULT DisplayNameOf(IShellFolder* psf, LPCITEMIDLIST pidl, DWORD flags, LPTSTR psz, UINT cch)
{
    *psz = 0;
    STRRET sr;
    HRESULT hr = psf->GetDisplayNameOf(pidl, flags, &sr);
    if (SUCCEEDED(hr))
        hr = StrRetToBuf(&sr, pidl, psz, cch);
    return hr;
}

// get the name and flags of an absolute IDlist
// in:
//      dwFlags     SHGDN_ flags as hints to the name space GetDisplayNameOf() function
//
// in/out:
//      *pdwAttribs (optional) return flags

HRESULT SHGetNameAndFlags(LPCITEMIDLIST pidl, DWORD dwFlags, LPTSTR pszName, UINT cchName, DWORD* pdwAttribs)
{
    if (pszName)
    {
        *pszName = 0;
    }

    HRESULT hrInit = SHCoInitialize();

    IShellFolder* psf;
    LPCITEMIDLIST pidlLast;
    HRESULT hr = SHBindToIDListParent(pidl, IID_PPV_ARG(IShellFolder, &psf), &pidlLast);
    if (SUCCEEDED(hr))
    {
        if (pszName)
            hr = DisplayNameOf(psf, pidlLast, dwFlags, pszName, cchName);

        if (SUCCEEDED(hr) && pdwAttribs)
        {
            *pdwAttribs = SHGetAttributes(psf, pidlLast, *pdwAttribs);
        }

        psf->Release();
    }

    SHCoUninitialize(hrInit);
    return hr;
}



//
// Function loader
//
#define MODULE_VARNAME(NAME) hMod_ ## NAME

#define LOAD_MODULE(NAME)                                        \
HMODULE MODULE_VARNAME(NAME) = LoadLibraryW(L#NAME ".dll");      \
if (!MODULE_VARNAME(NAME))                                       \
{\
    MessageBoxW(0,TEXT(#NAME),TEXT(#NAME),0); \
    return false;\
}

#define LOAD_FUNCTION_NO_FAIL(MODULE, FUNCTION) \
*(FARPROC *)&FUNCTION = GetProcAddress(MODULE_VARNAME(MODULE), #FUNCTION);

#define LOAD_FUNCTION(MODULE, FUNCTION)                                      \
LOAD_FUNCTION_NO_FAIL(MODULE, FUNCTION);                                     \
if (!FUNCTION)                                                               \
{ \
    MessageBoxW(0, TEXT(#FUNCTION), TEXT(#FUNCTION), 0); \
	return false; \
}

#define LOAD_ORDINAL_NO_FAIL(MODULE, FUNCNAME, ORDINAL)                           \
*(FARPROC *)&FUNCNAME = GetProcAddress(MODULE_VARNAME(MODULE), (LPCSTR)ORDINAL);  \

#define LOAD_ORDINAL(MODULE, FUNCNAME, ORDINAL)                                   \
LOAD_ORDINAL_NO_FAIL(MODULE, FUNCNAME, ORDINAL);                                  \
if (!FUNCNAME)                                                                    \
{ \
    MessageBoxW(0, TEXT(#FUNCNAME), TEXT(#FUNCNAME), 0); \
    return false; \
}

BOOL SHWinHelp(HWND hwndMain, LPCTSTR lpszHelp, UINT usCommand, ULONG_PTR ulData)
{
    // Try to show help
    if (!WinHelp(hwndMain, lpszHelp, usCommand, ulData))
    {
        // Problem.
        return FALSE;
    }
    return TRUE;
}

// runonce.cpp


// stolen from <tsappcmp.h>
#define TERMSRV_COMPAT_WAIT_USING_JOB_OBJECTS 0x00008000
#define CompatibilityApp 1
typedef LONG TERMSRV_COMPATIBILITY_CLASS;
typedef BOOL(*PFNGSETTERMSRVAPPINSTALLMODE)(BOOL bState);
typedef BOOL(*PFNGETTERMSRVCOMPATFLAGSEX)(LPWSTR pwszApp, DWORD* pdwFlags, TERMSRV_COMPATIBILITY_CLASS tscc);

// even though this function is in kernel32.lib, we need to have a LoadLibrary/GetProcAddress 
// thunk for downlevel components who include this
BOOL SHSetTermsrvAppInstallMode(BOOL bState)
{
    static PFNGSETTERMSRVAPPINSTALLMODE pfn = NULL;

    if (pfn == NULL)
    {
        // kernel32 should already be loaded
        HMODULE hmod = GetModuleHandle(TEXT("kernel32.dll"));

        if (hmod)
        {
            pfn = (PFNGSETTERMSRVAPPINSTALLMODE)GetProcAddress(hmod, "SetTermsrvAppInstallMode");
        }
        else
        {
            pfn = (PFNGSETTERMSRVAPPINSTALLMODE)-1;
        }
    }

    if (pfn && (pfn != (PFNGSETTERMSRVAPPINSTALLMODE)-1))
    {
        return pfn(bState);
    }
    else
    {
        return FALSE;
    }
}


ULONG SHGetTermsrCompatFlagsEx(LPWSTR pwszApp, DWORD* pdwFlags, TERMSRV_COMPATIBILITY_CLASS tscc)
{
    static PFNGETTERMSRVCOMPATFLAGSEX pfn = NULL;

    if (pfn == NULL)
    {
        HMODULE hmod = LoadLibrary(TEXT("TSAppCMP.DLL"));

        if (hmod)
        {
            pfn = (PFNGETTERMSRVCOMPATFLAGSEX)GetProcAddress(hmod, "GetTermsrCompatFlagsEx");
        }
        else
        {
            pfn = (PFNGETTERMSRVCOMPATFLAGSEX)-1;
        }
    }

    if (pfn && (pfn != (PFNGETTERMSRVCOMPATFLAGSEX)-1))
    {
        return pfn(pwszApp, pdwFlags, tscc);
    }
    else
    {
        *pdwFlags = 0;
        return 0;
    }
}


HANDLE SetJobCompletionPort(HANDLE hJob)
{
    HANDLE hRet = NULL;
    HANDLE hIOPort = CreateIoCompletionPort(INVALID_HANDLE_VALUE, NULL, 0, 1);

    if (hIOPort != NULL)
    {
        JOBOBJECT_ASSOCIATE_COMPLETION_PORT CompletionPort;

        CompletionPort.CompletionKey = hJob;
        CompletionPort.CompletionPort = hIOPort;

        if (SetInformationJobObject(hJob,
            JobObjectAssociateCompletionPortInformation,
            &CompletionPort,
            sizeof(CompletionPort)))
        {
            hRet = hIOPort;
        }
        else
        {
            CloseHandle(hIOPort);
        }
    }

    return hRet;

}


DWORD WaitingThreadProc(void* pv)
{
    HANDLE hIOPort = (HANDLE)pv;

    if (hIOPort)
    {
        while (TRUE)
        {
            DWORD dwCompletionCode;
            ULONG_PTR pCompletionKey;
            LPOVERLAPPED pOverlapped;

            if (!GetQueuedCompletionStatus(hIOPort, &dwCompletionCode, &pCompletionKey, &pOverlapped, INFINITE) ||
                (dwCompletionCode == JOB_OBJECT_MSG_ACTIVE_PROCESS_ZERO))
            {
                break;
            }
        }
    }

    return 0;
}


//
// The following handles running an application and optionally waiting for it
// to all the child procs to terminate. This is accomplished thru Kernel Job Objects
// which is only available in NT5
//
BOOL _CreateRegJob(LPCTSTR pszCmd, BOOL bWait)
{
    BOOL bRet = FALSE;
    HANDLE hJobObject = CreateJobObjectW(NULL, NULL);

    if (hJobObject)
    {
        HANDLE hIOPort = SetJobCompletionPort(hJobObject);

        if (hIOPort)
        {
            DWORD dwID;
            HANDLE hThread = CreateThread(NULL,
                0,
                WaitingThreadProc,
                (void*)hIOPort,
                CREATE_SUSPENDED,
                &dwID);

            if (hThread)
            {
                PROCESS_INFORMATION pi = { 0 };
                STARTUPINFO si = { 0 };
                UINT fMask = SEE_MASK_FLAG_NO_UI;
                DWORD dwCreationFlags = CREATE_SUSPENDED;
                TCHAR sz[MAX_PATH * 2];
                TCHAR szAppPath[MAX_PATH];

                if (GetSystemDirectory(szAppPath, ARRAYSIZE(szAppPath)))
                {
                    if (PathAppend(szAppPath, TEXT("RunDLL32.EXE")))
                    {
                        if (SUCCEEDED(StringCchPrintf(sz, ARRAYSIZE(sz),
                            TEXT("RunDLL32.EXE Shell32.DLL,ShellExec_RunDLL ?0x%X?%s"), fMask, pszCmd)))
                        {
                            si.cb = sizeof(si);

                            if (CreateProcess(szAppPath,
                                sz,
                                NULL,
                                NULL,
                                FALSE,
                                dwCreationFlags,
                                NULL,
                                NULL,
                                &si,
                                &pi))
                            {
                                if (AssignProcessToJobObject(hJobObject, pi.hProcess))
                                {
                                    // success!
                                    bRet = TRUE;

                                    ResumeThread(pi.hThread);
                                    ResumeThread(hThread);

                                    if (bWait)
                                    {
                                        SHProcessMessagesUntilEvent(NULL, hThread, INFINITE);
                                    }
                                }
                                else
                                {
                                    TerminateProcess(pi.hProcess, ERROR_ACCESS_DENIED);
                                }

                                CloseHandle(pi.hProcess);
                                CloseHandle(pi.hThread);
                            }
                        }
                    }
                }

                if (!bRet)
                {
                    TerminateThread(hThread, ERROR_ACCESS_DENIED);
                }

                CloseHandle(hThread);
            }

            CloseHandle(hIOPort);
        }

        CloseHandle(hJobObject);
    }

    return bRet;
}

BOOL _TryHydra(LPCTSTR pszCmd, RRA_FLAGS* pflags)
{
    // See if the terminal-services is enabled in "Application Server" mode
    if (IsOS(OS_TERMINALSERVER) && SHSetTermsrvAppInstallMode(TRUE))
    {
        WCHAR   sz[MAX_PATH];

        *pflags |= RRA_WAIT;
        // Changing timing blows up IE 4.0, but IE5 is ok!
        // we are on a TS machine, NT version 4 or 5, with admin priv

        // see if the app-compatability flag is set for this executable
        // to use the special job-objects for executing module

        // get the module name, without the arguments
        int argc;
        LPWSTR* argv = CommandLineToArgvW(pszCmd, &argc);
        if (argv)
        {
            wsprintf(sz, argv[0]);
            ULONG   ulCompat;
            SHGetTermsrCompatFlagsEx(sz, &ulCompat, CompatibilityApp);

            // if the special flag for this module-name is set...
            if (ulCompat & TERMSRV_COMPAT_WAIT_USING_JOB_OBJECTS)
            {
                *pflags |= RRA_USEJOBOBJECTS;
            }
        }

        return TRUE;
    }

    return FALSE;
}

//
//  On success: returns process handle or INVALID_HANDLE_VALUE if no process
//              was launched (i.e., launched via DDE).
//  On failure: returns INVALID_HANDLE_VALUE.
//
BOOL _ShellExecRegApp(LPCTSTR pszCmd, BOOL fNoUI, BOOL fWait)
{
    TCHAR szQuotedCmdLine[MAX_PATH + 2];
    LPTSTR pszArgs;
    SHELLEXECUTEINFO ei = { 0 };
    BOOL fNoError = TRUE;

    if (fNoError)
    {
        pszArgs = PathGetArgs(szQuotedCmdLine);
        if (*pszArgs)
        {
            // Strip args
            *(pszArgs - 1) = 0;
        }

        PathUnquoteSpaces(szQuotedCmdLine);

        ei.cbSize = sizeof(SHELLEXECUTEINFO);
        ei.lpFile = szQuotedCmdLine;
        ei.lpParameters = pszArgs;
        ei.nShow = SW_SHOWNORMAL;
        ei.fMask = SEE_MASK_NOCLOSEPROCESS;

        if (fNoUI)
        {
            ei.fMask |= SEE_MASK_FLAG_NO_UI;
        }

        if (ShellExecuteEx(&ei))
        {
            if (ei.hProcess)
            {
                if (fWait)
                {
                    SHProcessMessagesUntilEvent(NULL, ei.hProcess, INFINITE);
                }

                CloseHandle(ei.hProcess);
            }

            fNoError = TRUE;
        }
        else
        {
            fNoError = FALSE;
        }
    }
    return fNoError;
}

// The following handles running an application and optionally waiting for it
// to terminate.
BOOL ShellExecuteRegApp(LPCTSTR pszCmdLine, RRA_FLAGS fFlags)
{
    BOOL bRet = FALSE;

    if (!pszCmdLine || !*pszCmdLine)
    {
        // Don't let empty strings through, they will endup doing something dumb
        // like opening a command prompt or the like
        return bRet;
    }

#if (_WIN32_WINNT >= 0x0500)
    if (fFlags & RRA_USEJOBOBJECTS)
    {
        bRet = _CreateRegJob(pszCmdLine, fFlags & RRA_WAIT);
    }
#endif

    if (!bRet)
    {
        //  fallback if necessary.
        //bRet = _ShellExecRegApp(pszCmdLine, fFlags & RRA_NOUI, fFlags & RRA_WAIT);
    }

    return bRet;
}


BOOL Cabinet_EnumRegApps(HKEY hkeyParent, LPCTSTR pszSubkey, RRA_FLAGS fFlags, PFNREGAPPSCALLBACK pfnCallback, LPARAM lParam)
{
    HKEY hkey;
    BOOL bRet = TRUE;

    // With the addition of the ACL controlled "policy" run keys RegOpenKey
    // might fail on the pszSubkey.  Use RegOpenKeyEx with MAXIMIM_ALLOWED
    // to ensure that we successfully open the subkey.
    if (RegOpenKeyEx(hkeyParent, pszSubkey, 0, MAXIMUM_ALLOWED, &hkey) == ERROR_SUCCESS)
    {
        DWORD cbValue;
        DWORD dwType;
        DWORD i;
        TCHAR szValueName[80];
        TCHAR szCmdLine[MAX_PATH];
        HDPA hdpaEntries = NULL;

#ifdef DEBUG
        //
        // we only support named values so explicitly purge default values
        //
        LONG cbData = sizeof(szCmdLine);
        if (RegQueryValue(hkey, NULL, szCmdLine, &cbData) == ERROR_SUCCESS)
        {
            ASSERTMSG((cbData <= 2), "Cabinet_EnumRegApps: BOGUS default entry in <%s> '%s'", pszSubkey, szCmdLine);
            RegDeleteValue(hkey, NULL);
        }
#endif
        // now enumerate all of the values.
        for (i = 0; !g_fEndSession; i++)
        {
            LONG lEnum;
            DWORD cbData;

            cbValue = ARRAYSIZE(szValueName);
            cbData = sizeof(szCmdLine);

            lEnum = RegEnumValue(hkey, i, szValueName, &cbValue, NULL, &dwType, (LPBYTE)szCmdLine, &cbData);

            if (ERROR_MORE_DATA == lEnum)
            {
                // ERROR_MORE_DATA means the value name or data was too large
                // skip to the next item
                continue;
            }
            else if (lEnum != ERROR_SUCCESS)
            {
                if (lEnum != ERROR_NO_MORE_ITEMS)
                {
                    // we hit some kind of registry failure
                    bRet = FALSE;
                }
                break;
            }

            if ((dwType == REG_SZ) || (dwType == REG_EXPAND_SZ))
            {
                REGAPP_INFO* prai;

                if (dwType == REG_EXPAND_SZ)
                {
                    DWORD dwChars;
                    TCHAR szCmdLineT[MAX_PATH];

                    if (FAILED(StringCchCopy(szCmdLineT, ARRAYSIZE(szCmdLineT), szCmdLine)))
                    {
                        // bail on this value if string doesn't fit 
                        continue;
                    }
                    dwChars = ExpandEnvironmentStrings(szCmdLineT,
                        szCmdLine,
                        ARRAYSIZE(szCmdLine));
                    if ((dwChars == 0) || (dwChars > ARRAYSIZE(szCmdLine)))
                    {
                        // bail on this value if we failed the expansion, or if the string is > MAX_PATH
                        continue;
                    }
                }

                if (g_fCleanBoot && (szValueName[0] != TEXT('*')))
                {
                    // only run things marked with a "*" in when in SafeMode
                    continue;
                }

                // We used to execute each entry, wait for it to finish, and then make the next call to 
                // RegEnumValue(). The problem with this is that some apps add themselves back to the runonce
                // after they are finished (test harnesses that reboot machines and want to be restarted) and
                // we dont want to delete them, so we snapshot the registry keys and execute them after we
                // have finished the enum.
                prai = (REGAPP_INFO*)LocalAlloc(LPTR, sizeof(REGAPP_INFO));
                if (prai)
                {
                    if (SUCCEEDED(StringCchCopy(prai->szSubkey, ARRAYSIZE(prai->szSubkey), pszSubkey)) &&
                        SUCCEEDED(StringCchCopy(prai->szValueName, ARRAYSIZE(prai->szValueName), szValueName)) &&
                        SUCCEEDED(StringCchCopy(prai->szCmdLine, ARRAYSIZE(prai->szCmdLine), szCmdLine)))
                    {
                        if (!hdpaEntries)
                        {
                            hdpaEntries = DPA_Create(5);
                        }

                        if (!hdpaEntries || (DPA_AppendPtr(hdpaEntries, prai) == -1))
                        {
                            LocalFree(prai);
                        }
                    }
                }
            }
        }

        if (hdpaEntries)
        {
            int iIndex;
            int iTotal = DPA_GetPtrCount(hdpaEntries);

            for (iIndex = 0; iIndex < iTotal; iIndex++)
            {
                REGAPP_INFO* prai = (REGAPP_INFO*)DPA_GetPtr(hdpaEntries, iIndex);
                ASSERT(prai);

                // NB Things marked with a '!' mean delete after
                // the CreateProcess not before. This is to allow
                // certain apps (runonce.exe) to be allowed to rerun
                // to if the machine goes down in the middle of execing
                // them. Be very afraid of this switch.
                if ((fFlags & RRA_DELETE) && (prai->szValueName[0] != TEXT('!')))
                {
                    // This delete can fail if the user doesn't have the privilege
                    if (RegDeleteValue(hkey, prai->szValueName) != ERROR_SUCCESS)
                    {
                        LocalFree(prai);
                        continue;
                    }
                }

                pfnCallback(prai->szSubkey, prai->szCmdLine, fFlags, lParam);

                // Post delete '!' things.
                if ((fFlags & RRA_DELETE) && (prai->szValueName[0] == TEXT('!')))
                {
                    // This delete can fail if the user doesn't have the privilege
                    if (RegDeleteValue(hkey, prai->szValueName) != ERROR_SUCCESS)
                    {
                    }
                }

                LocalFree(prai);
            }

            DPA_Destroy(hdpaEntries);
            hdpaEntries = NULL;
        }

        RegCloseKey(hkey);
    }
    else
    {
        bRet = FALSE;
    }


    if (g_fEndSession)
    {
        // NOTE: this is for explorer only, other consumers of runonce.c must declare g_fEndSession but leave
        // it set to FALSE always.

        // if we rx'd a WM_ENDSESSION whilst running any of these keys we must exit the process.
        ExitProcess(0);
    }

    return bRet;
}

BOOL ExecuteRegAppEnumProc(LPCTSTR szSubkey, LPCTSTR szCmdLine, RRA_FLAGS fFlags, LPARAM lParam)
{
    BOOL bRet;
    RRA_FLAGS flagsTemp = fFlags;
    BOOL fInTSInstallMode = FALSE;

#if (_WIN32_WINNT >= 0x0500)
    // In here, We only attempt TS specific in app-install-mode 
    // if RunOnce entries are being processed 
    if (0 == lstrcmpi(szSubkey, REGSTR_PATH_RUNONCE))
    {
        fInTSInstallMode = _TryHydra(szCmdLine, &flagsTemp);
    }
#endif

    bRet = ShellExecuteRegApp(szCmdLine, flagsTemp);

#if (_WIN32_WINNT >= 0x0500)
    if (fInTSInstallMode)
    {
        SHSetTermsrvAppInstallMode(FALSE);
    }
#endif

    return bRet;
}

BOOL SHRunControlPanelCustom(LPCTSTR lpcszCmdLine, HWND hwndMsgParent)
{
    LPCTSTR pszCmdLine = NULL;

	if (!IS_INTRESOURCE(lpcszCmdLine))
	{
		pszCmdLine = StrDup(lpcszCmdLine);
	}
	else
	{
		ULONG id = PtrToUlong((void*)lpcszCmdLine);
		if (id == 9012) //hack cuz shit changed
		{
			pszCmdLine = L"SYSDM.CPL,System";
		}
		else
		{
			TCHAR szCmdLine[MAX_PATH];

			if (LoadString(LoadLibraryW(L"shell32.dll"), id, szCmdLine, ARRAYSIZE(szCmdLine)))
				pszCmdLine = StrDup(szCmdLine);
		}
	}

	if (!pszCmdLine)
		return FALSE;

	if (wcscmp(pszCmdLine, L"nusrmgr.cpl ,initialTask=ChangePicture") == 0) //another hack to hack 7 cpl in
	{
		IOpenControlPanel* openCPL = 0;
		if (SUCCEEDED(CoCreateInstance(CLSID_OpenControlPanel, 0, 0x17u, IID_PPV_ARGS(&openCPL))))
			openCPL->Open(L"Microsoft.UserAccounts", 0, 0);

		if (openCPL)
			openCPL->Release();
		return TRUE;
	}

	WCHAR parameters[MAX_PATH] = L"shell32.dll,Control_RunDLL ";
	wcscat_s(parameters, pszCmdLine);

	return ((INT_PTR)ShellExecuteW(hwndMsgParent, L"open", L"rundll32.exe", parameters, NULL, SW_SHOWNORMAL) > 32);
}

bool SHUndocInit(void)
{

	LOAD_MODULE(shlwapi);
	LOAD_ORDINAL(shlwapi, IUnknown_Exec, 164);
    LOAD_ORDINAL(shlwapi, SHSetWindowBits, 165);
	LOAD_ORDINAL(shlwapi, IUnknown_OnFocusChangeIS, 509);
	LOAD_ORDINAL(shlwapi, SHPropagateMessage, 178);
	LOAD_ORDINAL(shlwapi, SHGetCurColorRes, 193);
	LOAD_ORDINAL(shlwapi, SHFillRectClr, 197);
    LOAD_ORDINAL(shlwapi, SHRunIndirectRegClientCommand, 467);
    LOAD_ORDINAL(shlwapi, SHGetMachineInfo, 413);
    LOAD_ORDINAL(shlwapi, IUnknown_QueryStatus, 163);
    LOAD_ORDINAL(shlwapi, SHInvokeDefaultCommand, 279);
    LOAD_ORDINAL(shlwapi, IUnknown_UIActivateIO, 479);
    LOAD_ORDINAL(shlwapi, SHMessageBoxCheckExW, 292);
    LOAD_ORDINAL(shlwapi, SHDefWindowProc, 240);
    LOAD_ORDINAL(shlwapi, IUnknown_TranslateAcceleratorIO, 478);
    LOAD_ORDINAL(shlwapi, SHGetMenuFromID, 192);
    LOAD_ORDINAL(shlwapi, SHCreatePropertyBagOnMemory, 477);
    LOAD_ORDINAL(shlwapi, SHPropertyBag_WriteBOOL, 499);
    LOAD_FUNCTION(shlwapi, SHIsChildOrSelf);

	LOAD_MODULE(shell32);
	LOAD_ORDINAL(shell32, SHGetUserDisplayName, 241);
    LOAD_ORDINAL(shell32, RegisterShellHook, 181);
    LOAD_ORDINAL(shell32, SHSettingsChanged, 244);
    LOAD_ORDINAL(shell32, ExitWindowsDialog, 60);
    LOAD_ORDINAL(shell32, RunFileDlg, 61);
    LOAD_ORDINAL(shell32, LogoffWindowsDialog, 54);
    LOAD_ORDINAL(shell32, DisconnectWindowsDialog, 254);
    LOAD_ORDINAL(shell32, SHFindComputer, 91);
    LOAD_ORDINAL(shell32, SHTestTokenPrivilegeW, 236);
    LOAD_ORDINAL(shell32, SHMapIDListToSystemImageListIndexAsync, 787);
    LOAD_ORDINAL(shell32, SHGetUserPicturePath_t, 261);
    LOAD_FUNCTION(shell32, SHUpdateRecycleBinIcon);
    LOAD_ORDINAL(shell32, SHMapIDListToSystemImageListIndex, 790);
    LOAD_ORDINAL(shell32, CheckWinIniForAssocs, 711);
    LOAD_ORDINAL(shell32, CheckDiskSpace, 733);
    LOAD_ORDINAL(shell32, CheckStagingArea, 753);

    //shunimpld
    //LOAD_ORDINAL(shell32, DDECreatePostNotify, 82);
    //LOAD_ORDINAL(shell32, DDEHandleViewFolderNotify, 202);

    HMODULE hMod_shcore = LoadLibraryW(L"shcore.dll");
    if (hMod_shcore)
    {
		LOAD_ORDINAL(shcore, IUnknown_GetClassID, 142);
		LOAD_ORDINAL(shcore, SHQueueUserWorkItem, 162)
		LOAD_ORDINAL(shcore, SHLoadRegUIStringW, 126);
		LOAD_ORDINAL(shcore, SHCreateWorkerWindowW, 188);
    }
    else
    {
		LOAD_ORDINAL(shlwapi, IUnknown_GetClassID, 175);
		LOAD_ORDINAL(shlwapi, SHQueueUserWorkItem, 260);
		LOAD_ORDINAL(shlwapi, SHLoadRegUIStringW, 439);
		LOAD_ORDINAL(shlwapi, SHCreateWorkerWindowW, 278);

    }


    LOAD_MODULE(user32);
    LOAD_FUNCTION(user32, EndTask);
    LOAD_ORDINAL_NO_FAIL(user32, IsShellFrameWindow, 2573);
    LOAD_ORDINAL_NO_FAIL(user32, IsShellManagedWindow, 2574);
    LOAD_FUNCTION(user32, GhostWindowFromHungWindow);
    LOAD_FUNCTION_NO_FAIL(user32, GetWindowBand);

    LOAD_MODULE(msi);
    LOAD_FUNCTION(msi, MsiDecomposeDescriptorW);

    LOAD_MODULE(winsta);
    LOAD_FUNCTION(winsta, WinStationRegisterConsoleNotification);
    LOAD_FUNCTION(winsta, WinStationSetInformationW);
    LOAD_FUNCTION(winsta, WinStationUnRegisterConsoleNotification);

    LOAD_MODULE(comctl32);
    LOAD_FUNCTION(comctl32, ImageList_GetFlags);

    //expanded cuz ye...
    HMODULE hMod_windowsstorage = LoadLibrary(L"windows.storage.dll"); 
    if (hMod_windowsstorage)
        *(FARPROC*)&CFSFolder_CreateFolder = GetProcAddress(hMod_windowsstorage, "CFSFolder_CreateFolder"); 

    if (!CFSFolder_CreateFolder && hMod_windowsstorage)
    {
        const char* pattern = "40 53 55 56 57 41 54 41 56 41 57 48 83 EC 50 48 8B 05 ?? ?? ?? 00 48 33 C4 48 89 44 24 38";
        *(FARPROC*)&CFSFolder_CreateFolder= (FARPROC)FindPattern(pattern, (uintptr_t)hMod_windowsstorage);

        


        if (!CFSFolder_CreateFolder) return false;
    }
	//for win7
	if (!CFSFolder_CreateFolder)
		*(FARPROC*)&CFSFolder_CreateFolder = (FARPROC)FindPattern("48 89 5C 24 08 48 89 6C 24 10 48 89 74 24 18 57 41 54 41 55 48 83 EC 20 48 8B 74 24 68 48 8B D9 B9", (uintptr_t)hMod_shell32);

    //8.1 (1/2)
	if (!CFSFolder_CreateFolder)
		*(FARPROC*)&CFSFolder_CreateFolder = (FARPROC)FindPattern("48 8B C4 48 89 58 08 48 89 68 10 48 89 70 18 48 89 78 20 41 54 41 56 41 57 48 83 EC 20 4C 8B 74 24 68 48 8B D9", (uintptr_t)hMod_shell32);

    //8.1 (2/2)
	if (!CFSFolder_CreateFolder)
		*(FARPROC*)&CFSFolder_CreateFolder = (FARPROC)FindPattern("48 8B C4 48 89 58 08 48 89 68 10 48 89 70 18 48 89 78 20 41 56 48 83 EC 20 48 8B 74 24 58 49 8B F9 49 8B E8", (uintptr_t)hMod_shell32);

    if (!CFSFolder_CreateFolder)
    {
        MessageBoxW(0, TEXT("CFSFolder_CreateFolder"), TEXT("CFSFolder_CreateFolder"),0);
        return false;
    }

	return true;
}

BOOL IsStartPanelOn()
{
	SHELLSTATE ss = { 0 };
	SHGetSetSettings(&ss, SSF_STARTPANELON, FALSE);

	return ss.fStartPanelOn;
}

LPIDA DataObj_GetHIDAEx(IDataObject* pdtobj, CLIPFORMAT cf, STGMEDIUM* pmedium)
{
	FORMATETC fmte = { cf, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

	if (pmedium)
	{
		pmedium->pUnkForRelease = NULL;
		pmedium->hGlobal = NULL;
	}

	if (!pmedium)
	{
		if (S_OK == pdtobj->QueryGetData(&fmte))
			return (LPIDA)TRUE;
		else
			return (LPIDA)FALSE;
	}
	else if (SUCCEEDED(pdtobj->GetData(&fmte, pmedium)))
	{
		return (LPIDA)GlobalLock(pmedium->hGlobal);
	}

	return NULL;
}

LPIDA DataObj_GetHIDA(IDataObject* pdtobj, STGMEDIUM* pmedium)
{
	static CLIPFORMAT cfHIDA = 0;
	if (!cfHIDA)
	{
		cfHIDA = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	}
	return DataObj_GetHIDAEx(pdtobj, cfHIDA, pmedium);
}

LPCITEMIDLIST IDA_GetIDListPtr(LPIDA pida, UINT i)
{
	if (NULL == pida)
	{
		return NULL;
	}

	if (i == (UINT)-1 || i < pida->cidl)
	{
		return HIDA_GetPIDLItem(pida, i);
	}

	return NULL;
}

LPITEMIDLIST IDA_FullIDList(LPIDA pida, UINT i)
{
	LPITEMIDLIST pidl = NULL;
	LPCITEMIDLIST pidlParent = IDA_GetIDListPtr(pida, (UINT)-1);
	if (pidlParent)
	{
		LPCITEMIDLIST pidlRel = IDA_GetIDListPtr(pida, i);
		if (pidlRel)
		{
			pidl = ILCombine(pidlParent, pidlRel);
		}
	}
	return pidl;
}

BOOL _IsLocalHardDisk(LPCTSTR pszPath)
{
	//  Reject CDs, floppies, network drives, etc.
	//
	int iDrive = PathGetDriveNumber(pszPath);
	if (iDrive < 0 ||                   // reject UNCs
		RealDriveType(iDrive, /* fOkToHitNet = */ FALSE) != DRIVE_FIXED) // reject slow media
	{
		return FALSE;
	}
	return TRUE;
}

BOOL _IsAcceptableTarget(LPCTSTR pszPath, DWORD dwAttrib, DWORD dwFlags)
{
	//  Regitems ("Internet" or "Email" for example) are acceptable
	//  provided we aren't restricted to EXEs only.
	if (!(dwAttrib & SFGAO_FILESYSTEM))
	{
		return !(dwFlags & SMPINNABLE_EXEONLY);
	}

	//  Otherwise, it's a file.

	//  If requested, reject non-EXEs.
	//  (Like the Start Menu, we treat MSC files as if they were EXEs)
	if (dwFlags & SMPINNABLE_EXEONLY)
	{
		LPCTSTR pszExt = PathFindExtension(pszPath);
		if (StrCmpIC(pszExt, TEXT(".EXE")) != 0 &&
			StrCmpIC(pszExt, TEXT(".MSC")) != 0)
		{
			return FALSE;
		}
	}

	//  If requested, reject slow media
	if (dwFlags & SMPINNABLE_REJECTSLOWMEDIA)
	{
		if (!_IsLocalHardDisk(pszPath))
		{
			return FALSE;
		}

		// If it's a shortcut, then apply the same rule to the shortcut.
		if (PathIsLnk(pszPath))
		{
			BOOL fLocal = TRUE;
			IShellLink* psl;
			if (SUCCEEDED(LoadFromFileW(CLSID_ShellLink, pszPath, IID_PPV_ARG(IShellLink, &psl))))
			{
				// IShellLink::GetPath returns S_FALSE if target is not a path
				TCHAR szPath[MAX_PATH];
				if (S_OK == psl->GetPath(szPath, ARRAYSIZE(szPath), NULL, 0))
				{
					fLocal = _IsLocalHardDisk(szPath);
				}
				psl->Release();
			}
			if (!fLocal)
			{
				return FALSE;
			}
		}
	}

	//  All tests pass!

	return TRUE;

}

void ReleaseStgMediumHGLOBAL(void* pv, STGMEDIUM* pmedium)
{
	if (pmedium->hGlobal && (pmedium->tymed == TYMED_HGLOBAL))
	{
#ifdef DEBUG
		if (pv)
		{
			void* pvT = (void*)GlobalLock(pmedium->hGlobal);
			ASSERT(pvT == pv);
			GlobalUnlock(pmedium->hGlobal);
		}
#endif
		GlobalUnlock(pmedium->hGlobal);
	}
	else
	{
		ASSERT(0);
	}

	ReleaseStgMedium(pmedium);
}

void HIDA_ReleaseStgMedium(LPIDA pida, STGMEDIUM* pmedium)
{
	ReleaseStgMediumHGLOBAL((void*)pida, pmedium);
}

HRESULT IsPinnable(IDataObject* pdtobj, DWORD dwFlags, OPTIONAL LPITEMIDLIST* ppidl)
{
	HRESULT hr = S_FALSE;

	LPITEMIDLIST pidlRet = NULL;

	if (pdtobj &&                                   // must have a data object
		!SHRestricted(REST_NOSMPINNEDLIST) &&       // cannot be restricted
		IsStartPanelOn())                           // start panel must be on
	{
		STGMEDIUM medium = { 0 };
		LPIDA pida = DataObj_GetHIDA(pdtobj, &medium);
		if (pida)
		{
			if (pida->cidl == 1)
			{
				pidlRet = IDA_FullIDList(pida, 0);
				if (pidlRet)
				{
					DWORD dwAttr = SFGAO_FILESYSTEM;            // only SFGAO_FILESYSTEM is valid
					TCHAR szPath[MAX_PATH];

					if (SUCCEEDED(SHGetNameAndFlags(pidlRet, SHGDN_FORPARSING,
						szPath, ARRAYSIZE(szPath), &dwAttr)) &&
						_IsAcceptableTarget(szPath, dwAttr, dwFlags))
					{
						hr = S_OK;
					}
				}
			}
			HIDA_ReleaseStgMedium(pida, &medium);
		}
	}

	// Return pidlRet only if the call succeeded and the caller requested it
	if (hr != S_OK || !ppidl)
	{
		ILFree(pidlRet);
		pidlRet = NULL;
	}

	if (ppidl)
	{
		*ppidl = pidlRet;
	}

	return hr;

}

HRESULT SHGetUserPicturePath(LPCWSTR pszUsername, DWORD dwFlags, LPWSTR pszPath)
{
    return SHGetUserPicturePath_t(pszUsername, dwFlags, pszPath, MAX_PATH);
}
