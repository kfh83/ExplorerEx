#include "StartButton.h"

#include <shlwapi.h>

CStartButton::CStartButton()
{
}

HRESULT CStartButton::QueryInterface(REFIID riid, void** ppvObject)  // taken from ep_taskbar 7-stuff (?)
{
    static const QITAB qit[] =
    {
        QITABENT(CStartButton, IStartButton),
        QITABENT(CStartButton, IServiceProvider),
        { 0 },
    };

    return QISearch(this, qit, riid, ppvObject);
}

ULONG CStartButton::AddRef()
{
    return 1; // Copilot generated body
}

ULONG CStartButton::Release()
{
    return 1; // Copilot generated body
}

HRESULT CStartButton::SetFocusToStartButton()  // taken from ep_taskbar 7-stuff
{
    SetFocus(_hwndStartBtn);
    return S_OK;
}

HRESULT CStartButton::SetStartPaneActive(BOOL bActive)
{
    if (bActive)
    {
        _nStartPaneActiveState = 1;
    }
    else if (_nStartPaneActiveState != 2)
    {
        _nStartPaneActiveState = 0;
        UnlockStartPane();
    }
    return S_OK;
}

HRESULT CStartButton::OnStartMenuDismissed()  // taken from ep_taskbar 7-stuff
{
    _pStartButtonSite->OnStartMenuDismissed();
    return S_OK;
}

HRESULT CStartButton::UnlockStartPane()  // taken from ep_taskbar 7-stuff
{
    if (_uLockCode)
    {
        _uLockCode = FALSE;
        LockSetForegroundWindow(LSFW_UNLOCK);
    }
    return S_OK;
}

HRESULT CStartButton::LockStartPane()  // taken from ep_taskbar 7-stuff
{
    if (!_uLockCode)
    {
        _uLockCode = TRUE;
        LockSetForegroundWindow(LSFW_LOCK);
    }
    return S_OK;
}

HRESULT CStartButton::GetPopupPosition(DWORD* out)  // taken from ep_taskbar 7-stuff
{
    if (!_pStartButtonSite)
        return E_FAIL;

    UINT stuckPlace = _pStartButtonSite->GetStartMenuStuckPlace();
    switch (stuckPlace)
    {
        case 0: *out = MPPF_LEFT; break;
        case 1: *out = MPPF_TOP; break;
        case 2: *out = MPPF_RIGHT; break;
        default: *out = MPPF_BOTTOM; break;
    }

    return S_OK;
}

HRESULT CStartButton::GetWindow(HWND* out)  // taken from ep_taskbar 7-stuff
{
    *out = _hwndStartBtn;
    return S_OK;
}

HRESULT CStartButton::QueryService(const GUID& guidService, const IID& riid, void** ppvObject)  // taken from ep_taskbar 7-stuff
{
    *ppvObject = nullptr;
    HRESULT hr = E_FAIL;

    if (IsEqualGUID(guidService, __uuidof(IStartButton)))
    {
        hr = QueryInterface(riid, ppvObject);
    }

    return hr;
}

void CStartButton::CloseStartMenu()  // taken from ep_taskbar 7-stuff @MOD
{
    if (_pOldStartMenu)
    {
        _pOldStartMenu->OnSelect(MPOS_FULLCANCEL);
        // UnlockStartPane(); ?
    }
    if (_pNewStartMenu)
    {
        _pNewStartMenu->OnSelect(MPOS_FULLCANCEL);
    }
}

HRESULT CStartButton::IsMenuMessage(MSG* pmsg)  // taken from ep_taskbar 7-stuff @MOD
{
    HRESULT hr;
    if (_pOldStartMenuBand)
    {
        hr = _pOldStartMenuBand->IsMenuMessage(pmsg);
        if (hr != S_OK)
        {
            hr = S_FALSE;
        }
    }
    else if (_pNewStartMenuBand)
    {
        hr = _pNewStartMenuBand->IsMenuMessage(pmsg);
        if (hr != S_OK)
        {
            hr = S_FALSE;
        }
    }
    else
    {
        hr = S_FALSE;
    }
    return hr;
}

BOOL CStartButton::TranslateMenuMessage(MSG* pmsg, LRESULT* plRet)  // taken from ep_taskbar 7-stuff @MOD
{
    BOOL result = TRUE;
    if (_pOldStartMenuBand)
    {
        // S_FALSE is same as TRUE
        // S_OK is same as FALSE
        // in pseudocode, it checks if TranslateMenuMessage is TRUE (S_FALSE)
        result = (_pOldStartMenuBand->TranslateMenuMessage(pmsg, plRet) == S_FALSE) ? TRUE : FALSE;

        if (result && _pNewStartMenuBand)
        {
            result = (_pNewStartMenuBand->TranslateMenuMessage(pmsg, plRet) == S_FALSE) ? TRUE : FALSE;
        }
    }
    return result;
}
