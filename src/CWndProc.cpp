#include "pch.h"
#include "cwndproc.h"
#include "shundoc.h"

#define ID_NOTIFY_SUBCLASS (DWORD)'CHN'     // CHN change notify
//
// CImpWndProc
//
LRESULT CALLBACK CImpWndProc::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (WM_NCCREATE == uMsg)
    {
        CImpWndProc* pThis = (CImpWndProc*)(((LPCREATESTRUCT)lParam)->lpCreateParams);
        if (EVAL(pThis))
        {
            pThis->_hwnd = hwnd;
            SetWindowLongPtr(hwnd, 0, (LONG_PTR)pThis);

            // Even if pThis->vWndProc fails the create, USER will always
            // send us a WM_NCDESTROY so we always get a chance to clean up
            return pThis->v_WndProc(hwnd, uMsg, wParam, lParam);
        }
        return FALSE;
    }
    else
    {
        CImpWndProc* pThis = (CImpWndProc*)GetWindowLongPtr(hwnd, 0);    // GetWindowLong(hwnd, 0);
        LRESULT lres;

        if (pThis)
        {
            // Always retain a ref across the v_WndProc in case
            // the window destroys itself during the callback.
            pThis->AddRef();

            lres = pThis->v_WndProc(hwnd, uMsg, wParam, lParam);

            if (uMsg == WM_NCDESTROY)
            {
                SetWindowLongPtr(hwnd, 0, NULL);
                pThis->_hwnd = NULL;
            }
            pThis->Release();
        }
        else
        {
            //
            // The only way this should happen is if we haven't actually
            // gotten a WM_NCCREATE yet.  User sends a WM_GETMINMAXINFO
            // to some windows before WM_NCCREATE (for legacy compat).
            // Assert that we're hitting that case.
            //
            ASSERT(uMsg == WM_GETMINMAXINFO);

            lres = SHDefWindowProc(hwnd, uMsg, wParam, lParam);
        }

        return lres;
    }
}

#ifndef NO_NOTIFYSUBCLASSWNDPROC

//
// CNotifySubclassWndProc
//
UINT g_idFSNotify;        // the message SHChangeNotify sends

BOOL CNotifySubclassWndProc::_SubclassWindow(HWND hwnd)
{
    if (0 == g_idFSNotify)
    {
        g_idFSNotify = RegisterWindowMessage(TEXT("SubclassedFSNotify"));
    }

    _hwndSubclassed = hwnd;

    return SetWindowSubclass(hwnd, _SubclassWndProc, ID_NOTIFY_SUBCLASS, (DWORD_PTR)this);
}

void CNotifySubclassWndProc::_UnsubclassWindow(HWND hwnd)
{
    RemoveWindowSubclass(hwnd, _SubclassWndProc, ID_NOTIFY_SUBCLASS);
}

LRESULT CNotifySubclassWndProc::_DefWindowProc(HWND hwnd, UINT uMessage, WPARAM wParam, LPARAM lParam)
{
    return DefSubclassProc(hwnd, uMessage, wParam, lParam);
}

LRESULT CALLBACK CNotifySubclassWndProc::_SubclassWndProc(
                                  HWND hwnd, UINT uMessage, 
                                  WPARAM wParam, LPARAM lParam,
                                  UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    CNotifySubclassWndProc* pObj = (CNotifySubclassWndProc*)dwRefData;

    if (pObj)
    {
        if (uMessage == g_idFSNotify)
        {
            LPSHChangeNotificationLock pshcnl;
            LPITEMIDLIST *ppidl;
            LONG lEvent;

            if (TRUE && (wParam || lParam))
            {
                // New style of notifications need to lock and unlock in order to free the memory...
                pshcnl = SHChangeNotification_Lock((HANDLE)wParam, (DWORD) lParam, &ppidl, &lEvent);
                if (pshcnl)
                {
                    pObj->OnChange(lEvent, ppidl[0], ppidl[1]);
                }
            }
            else
            {
                lEvent = (DWORD) lParam; // process id's are 32bits even in WIN64
                ppidl = (LPITEMIDLIST*)wParam;
                pshcnl = NULL;
                if (ppidl)
                {
                    pObj->OnChange(lEvent, ppidl[0], ppidl[1]);
                }
            }

            if (pshcnl)
            {
                SHChangeNotification_Unlock(pshcnl);
            }

            return 0;
        }
        else
        {
            return pObj->_DefWindowProc(hwnd, uMessage, wParam, lParam);
        }
    }
    else
    {
        return DefSubclassProc(hwnd, uMessage, wParam, lParam);
    }
}

void CNotifySubclassWndProc::_FlushNotifyMessages(HWND hwnd)
{
    MSG msg;

    ASSERT(hwnd == _hwndSubclassed);

    // This SHChangeNotify calls flushes notifications
    // via PostMessage, so I need to remove them
    // myself and process them immediately...
    //
    SHChangeNotify(0, SHCNF_FLUSH, NULL, NULL);

    while (PeekMessage(&msg, hwnd, g_idFSNotify, g_idFSNotify, PM_REMOVE))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

void CNotifySubclassWndProc::_RegisterWindow(HWND hwnd, LPCITEMIDLIST pidl, long lEvents,
                                             UINT uFlags/* = SHCNRF_ShellLevel | SHCNRF_InterruptLevel*/)
{
    ASSERT(0 != g_idFSNotify);

    // We only register if there's something to listen to
    //
    if (0==_uRegister)
    {
        // Since we don't know what this pidl actually points to,
        // we have to register to listen to everything that might affect it...
        //

        SHChangeNotifyEntry se;
        uFlags |= SHCNRF_ShellLevel | SHCNRF_InterruptLevel;
        se.fRecursive = TRUE;
        se.pidl = pidl;
        _uRegister = SHChangeNotifyRegister(hwnd, uFlags, lEvents, g_idFSNotify, 1, &se);

        ASSERT(hwnd == _hwndSubclassed);
    }
}

void CNotifySubclassWndProc::_UnregisterWindow(HWND hwnd)
{
    if (_uRegister)
    {
        ASSERT(hwnd == _hwndSubclassed);

        // Avoid getting reentered...
        UINT uRegister = _uRegister;
        _uRegister = 0;
        SHChangeNotifyDeregister(uRegister);
    }
}

#endif  // NO_NOTIFYSUBCLASSWNDPROC
