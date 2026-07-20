#include "pch.h"

#include "SoundWnd.h"

#include "cabinet.h"
#include "ShUndoc.h"

HWND g_hwndSound = nullptr;

class CSoundWnd
{
public:
    CSoundWnd();
    ULONG AddRef();
    ULONG Release();

    BOOL Init();

protected:
    static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT CALLBACK v_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    static DWORD WINAPI s_CreateWindow(void* pvParam);
    static DWORD WINAPI s_ThreadProc(void* pvParam);

    LONG _cRef;
    HWND _hwndSound;
    HANDLE _hThread;
};

CSoundWnd::CSoundWnd()
    : _cRef(1)
    , _hThread(nullptr)
{
}

ULONG CSoundWnd::AddRef()
{
    return InterlockedIncrement(&_cRef);
}

ULONG CSoundWnd::Release()
{
    ASSERT(0 != _cRef); // 22
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

BOOL CSoundWnd::Init()
{
    ASSERT(!g_hwndSound); // 88
    SHCreateThread(s_ThreadProc, this, CTF_THREAD_REF | CTF_COINIT | CTF_REF_COUNTED, s_CreateWindow);
    g_hwndSound = _hwndSound;
    return _hwndSound != nullptr;
}

LRESULT CSoundWnd::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CSoundWnd* self = static_cast<CSoundWnd*>(GetWindowPtr0(hwnd));
    if (self)
        return self->v_WndProc(hwnd, uMsg, wParam, lParam);
    else
        return SHDefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CSoundWnd::v_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        case WM_QUERYENDSESSION:
            if ((lParam & ENDSESSION_CRITICAL) == 0)
            {
                WCHAR szReason[256];
                LoadStringW(g_hinstCabinet, 731, szReason, ARRAYSIZE(szReason));
                ShutdownBlockReasonCreate(_hwndSound, szReason);
                // PlayLogonLogoffSound(&_hThread, (lParam & ENDSESSION_LOGOFF) != 0 ? 0x1 : 0x2);
            }
            return 1;
        case WM_ENDSESSION:
            if (wParam && (lParam & ENDSESSION_CRITICAL) == 0 && _hThread != nullptr)
            {
                // Skipped telemetry ShellTraceId_Explorer_PlaySoundWait_Start
                WaitForSingleObject(_hThread, INFINITE);
                CloseHandle(_hThread);
                // Skipped telemetry ShellTraceId_Explorer_PlaySoundWait_Stop
            }
            DestroyWindow(_hwndSound);
            break;
        case WM_NCDESTROY:
            SetWindowLongPtrW(hwnd, 0, 0);
            g_hwndSound = nullptr;
            _hwndSound = nullptr;
            PostQuitMessage(0);
            return 0;
        default:
            return SHDefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

DWORD CSoundWnd::s_CreateWindow(void* pvParam)
{
    CSoundWnd* self = static_cast<CSoundWnd*>(pvParam);
    self->AddRef();
    self->_hwndSound = SHCreateWorkerWindowW(s_WndProc, nullptr, 0, 0, nullptr, pvParam);
    return 0;
}

DWORD CSoundWnd::s_ThreadProc(void* pvParam)
{
    CSoundWnd* self = static_cast<CSoundWnd*>(pvParam);
    if (self->_hwndSound)
    {
        MSG msg;
        while (GetMessageW(&msg, nullptr, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    self->Release();
    return 0;
}

STDAPI_(BOOL) InitSoundWindow()
{
    BOOL fRet = FALSE;

    CSoundWnd* pSoundWindow = new CSoundWnd();
    if (pSoundWindow != nullptr)
    {
        fRet = pSoundWindow->Init();
        pSoundWindow->Release();
    }
    return fRet;
}
