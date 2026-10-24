// Taskbar and Start Menu properties: Windows 2000 pages and ExplorerEx options.
#include "pch.h"
#include "cabinet.h"
#include "util.h"
#include "help.h"
#include "dlg.h"
#include "tray.h"
#include "traycmn.h"
#include "CfgStart.h"

static const DWORD s_TaskOptionsHelpIDs[] = {
    IDC_VOBASE,            IDH_TASKBAR_OPTIONS_BITMAP,
    IDC_VOLARGEMENU,       IDH_TASKBAR_OPTIONS_BITMAP,
    IDC_VOTRAY,            IDH_TASKBAR_OPTIONS_BITMAP,
    IDC_VOTRAYNOCLOCK,     IDH_TASKBAR_OPTIONS_BITMAP,
    IDC_VOWINDOW,          IDH_TASKBAR_OPTIONS_BITMAP,
    IDC_TRAYOPTONTOP,      IDH_TRAY_TASKBAR_ONTOP,
    IDC_TRAYOPTAUTOHIDE,   IDH_TRAY_TASKBAR_AUTOHIDE,
    IDC_SMSMALLICONS,      IDH_STARTMENU_SMALLICONS,
    IDC_TRAYOPTSHOWCLOCK,  IDH_TRAY_SHOW_CLOCK,
    IDC_PERSONALIZEDMENUS, IDH_TRAY_PERSONALIZED_MENUS,
    0, 0
};


class CTaskBarPropertySheet : public CPropertySheetImpl<CTaskBarPropertySheet>
{
public:
    CTaskBarPropertySheet(UINT nStartPage, HWND hwndParent) :
        CPropertySheetImpl<CTaskBarPropertySheet>((LPCTSTR)NULL, nStartPage, hwndParent),
        _title{}, _bitmaps{}, _pagesReady(FALSE)
    {
        ::LoadString(hinstCabinet, IDS_STARTMENUANDTASKBAR, _title, ARRAYSIZE(_title));
        SetTitle(_title);
        m_psh.dwFlags |= PSH_PROPTITLE;

        PROPSHEETPAGE psp = {};
        psp.dwSize = sizeof(psp);
        psp.hInstance = hinstCabinet;
        psp.pszTemplate = MAKEINTRESOURCE(DLG_TRAY_VIEW_OPTIONS);
        psp.pfnDlgProc = s_GeneralDlgProc;
        psp.lParam = reinterpret_cast<LPARAM>(this);
        if (!AddPage(&psp))
            return;

        HPROPSHEETPAGE advanced = _advanced.CreatePage(hinstCabinet);
        if (!advanced)
            return;
        if (!AddPage(advanced))
        {
            ::DestroyPropertySheetPage(advanced);
            return;
        }

        psp.pszTemplate = MAKEINTRESOURCE(DLG_EXTRA_OPTIONS);
        psp.pfnDlgProc = s_ExtraDlgProc;
        _pagesReady = AddPage(&psp);
    }

    ~CTaskBarPropertySheet()
    {
        // If construction failed before DoModal, release pages while the
        // Advanced page's instance is still alive for PSPCB_RELEASE.
        for (int i = 0; i < m_arrPages.GetSize(); ++i)
            ::DestroyPropertySheetPage(m_arrPages[i]);
        _CleanUpPages();

        // Normally WM_DESTROY has already detached and released these images.
        for (HBITMAP bitmap : _bitmaps)
        {
            if (bitmap)
                ::DeleteObject(bitmap);
        }
    }

    BOOL IsReady() const { return _pagesReady; }

    DECLARE_EMPTY_MSG_MAP()

private:
    static CTaskBarPropertySheet* GetPageOwner(HWND hDlg, UINT msg, LPARAM lParam)
    {
        if (msg == WM_INITDIALOG)
        {
            const PROPSHEETPAGE* page = reinterpret_cast<const PROPSHEETPAGE*>(lParam);
            ::SetWindowLongPtr(hDlg, DWLP_USER, page->lParam);
        }
        return reinterpret_cast<CTaskBarPropertySheet*>(::GetWindowLongPtr(hDlg, DWLP_USER));
    }

    static INT_PTR CALLBACK s_GeneralDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
    {
        CTaskBarPropertySheet* self = GetPageOwner(hDlg, msg, lParam);
        return self ? self->GeneralDlgProc(hDlg, msg, wParam, lParam) : FALSE;
    }

    static INT_PTR CALLBACK s_ExtraDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
    {
        CTaskBarPropertySheet* self = GetPageOwner(hDlg, msg, lParam);
        return self ? self->ExtraDlgProc(hDlg, msg, wParam, lParam) : FALSE;
    }

    INT_PTR GeneralDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam);
    INT_PTR ExtraDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam);
    void InitGeneral(HWND hDlg);
    void ApplyGeneral(HWND hDlg);
    void LoadBitmaps(HWND hDlg);
    void AttachBitmaps(HWND hDlg);
    void DestroyBitmaps(HWND hDlg);
    void SizePreview(HWND hDlg);
    void UpdatePreview(HWND hDlg);

    TCHAR _title[MAX_PATH];
    HBITMAP _bitmaps[IDB_VIEWOPTIONSLAST - IDB_VIEWOPTIONSFIRST + 1];
    BOOL _pagesReady;
    CStartMenuConfigPage _advanced;
};

void CTaskBarPropertySheet::DestroyBitmaps(HWND hDlg)
{
    for (UINT i = 0; i < ARRAYSIZE(_bitmaps); ++i)
    {
        ::SendDlgItemMessage(hDlg, IDC_VIEWOPTIONSICONSFIRST + i, STM_SETIMAGE, IMAGE_BITMAP, 0);
        if (_bitmaps[i])
        {
            ::DeleteObject(_bitmaps[i]);
            _bitmaps[i] = NULL;
        }
    }
}

void CTaskBarPropertySheet::LoadBitmaps(HWND hDlg)
{
    DestroyBitmaps(hDlg);
    for (UINT i = 0; i < ARRAYSIZE(_bitmaps); ++i)
    {
        _bitmaps[i] = static_cast<HBITMAP>(::LoadImage(hinstCabinet,
            MAKEINTRESOURCE(IDB_VIEWOPTIONSFIRST + i), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS));
    }
}

void CTaskBarPropertySheet::AttachBitmaps(HWND hDlg)
{
    // STM_SETIMAGE gives SS_BITMAP controls their final image-sized bounds.
    // Attach after preliminary positioning, as the SP4 dialog does.
    for (UINT i = 0; i < ARRAYSIZE(_bitmaps); ++i)
    {
        ::SendDlgItemMessage(hDlg, IDC_VIEWOPTIONSICONSFIRST + i, STM_SETIMAGE,
            IMAGE_BITMAP, reinterpret_cast<LPARAM>(_bitmaps[i]));
    }
}

void CTaskBarPropertySheet::SizePreview(HWND hDlg)
{
    // The overlays are pixel-aligned to the original bitmaps, not dialog units.
    HWND base = ::GetDlgItem(hDlg, IDC_VOBASE);
    POINT origin = {};
    ::MapWindowPoints(base, hDlg, &origin, 1);
    ::SetWindowPos(base, HWND_BOTTOM, 0, 0, 274, 130, SWP_NOMOVE | SWP_NOACTIVATE);
    ::SetWindowPos(::GetDlgItem(hDlg, IDC_VOLARGEMENU), NULL,
        origin.x + 2, origin.y, 150, 98, SWP_NOACTIVATE | SWP_NOZORDER);
    ::SetWindowPos(::GetDlgItem(hDlg, IDC_VOTRAY), HWND_TOP,
        origin.x + 2, origin.y + 100, 268, 28, SWP_NOACTIVATE);
    ::SetWindowPos(::GetDlgItem(hDlg, IDC_VOTRAYNOCLOCK), HWND_TOP,
        origin.x + 186, origin.y + 103, 79, 22, SWP_NOACTIVATE);

    BITMAP bitmap = {};
    if (::GetObject(_bitmaps[IDB_VOBASE - IDB_VIEWOPTIONSFIRST], sizeof(bitmap), &bitmap))
    {
        ::SetWindowPos(::GetDlgItem(hDlg, IDC_VOWINDOW), HWND_TOP,
            origin.x + bitmap.bmWidth - 62, origin.y + 100, 61, 6, SWP_NOACTIVATE);
    }
}

void CTaskBarPropertySheet::UpdatePreview(HWND hDlg)
{
    // The base contains a small menu and an auto-hidden, always-on-top taskbar.
    ::SetWindowPos(::GetDlgItem(hDlg, IDC_VOBASE), HWND_BOTTOM, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    ::ShowWindow(::GetDlgItem(hDlg, IDC_VOLARGEMENU),
        ::IsDlgButtonChecked(hDlg, IDC_SMSMALLICONS) == BST_CHECKED ? SW_HIDE : SW_SHOWNA);

    BOOL autoHide = ::IsDlgButtonChecked(hDlg, IDC_TRAYOPTAUTOHIDE) == BST_CHECKED;
    BOOL showClock = ::IsDlgButtonChecked(hDlg, IDC_TRAYOPTSHOWCLOCK) == BST_CHECKED;
    ::ShowWindow(::GetDlgItem(hDlg, IDC_VOTRAY), autoHide ? SW_HIDE : SW_SHOWNA);
    ::ShowWindow(::GetDlgItem(hDlg, IDC_VOTRAYNOCLOCK),
        !autoHide && !showClock ? SW_SHOWNA : SW_HIDE);

    HWND window = ::GetDlgItem(hDlg, IDC_VOWINDOW);
    BOOL alwaysOnTop = ::IsDlgButtonChecked(hDlg, IDC_TRAYOPTONTOP) == BST_CHECKED;
    ::ShowWindow(window, alwaysOnTop ? SW_HIDE : SW_SHOWNA);
    if (!alwaysOnTop)
    {
        ::SetWindowPos(window, HWND_TOP, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }
}

void CTaskBarPropertySheet::InitGeneral(HWND hDlg)
{
    ::CheckDlgButton(hDlg, IDC_TRAYOPTONTOP, c_tray._fAlwaysOnTop ? BST_CHECKED : BST_UNCHECKED);
    ::CheckDlgButton(hDlg, IDC_TRAYOPTAUTOHIDE, (c_tray._uAutoHide & AH_ON) ? BST_CHECKED : BST_UNCHECKED);
    ::CheckDlgButton(hDlg, IDC_SMSMALLICONS, c_tray._fSMSmallIcons ? BST_CHECKED : BST_UNCHECKED);
    ::CheckDlgButton(hDlg, IDC_TRAYOPTSHOWCLOCK, c_tray._fHideClock ? BST_UNCHECKED : BST_CHECKED);
    ::EnableWindow(::GetDlgItem(hDlg, IDC_TRAYOPTSHOWCLOCK), !SHRestricted(REST_HIDECLOCK));

    if (SHRestricted(REST_INTELLIMENUS))
        ::ShowWindow(::GetDlgItem(hDlg, IDC_PERSONALIZEDMENUS), SW_HIDE);
    else
        ::CheckDlgButton(hDlg, IDC_PERSONALIZEDMENUS,
            SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("IntelliMenus"), FALSE, TRUE)
                ? BST_CHECKED : BST_UNCHECKED);

    LoadBitmaps(hDlg);
    SizePreview(hDlg);
    AttachBitmaps(hDlg);
    UpdatePreview(hDlg);
}

void CTaskBarPropertySheet::ApplyGeneral(HWND hDlg)
{
    BOOL changed = c_tray._UpdateAlwaysOnTop(::IsDlgButtonChecked(hDlg, IDC_TRAYOPTONTOP) == BST_CHECKED);
    BOOL autoHide = ::IsDlgButtonChecked(hDlg, IDC_TRAYOPTAUTOHIDE) == BST_CHECKED;
    LONG result = c_tray._SetAutoHideState(autoHide);
    if (!HIWORD(result))
    {
        // The tray reports collisions itself. Reflect the state it actually accepted.
        ::CheckDlgButton(hDlg, IDC_TRAYOPTAUTOHIDE,
            (c_tray._uAutoHide & AH_ON) ? BST_CHECKED : BST_UNCHECKED);
        UpdatePreview(hDlg);
    }
    changed |= LOWORD(result);
    if (changed)
        c_tray._AppBarNotifyAll(NULL, ABN_STATECHANGE, NULL, 0);

    // Update only these options: do not overwrite auto-hide runtime bits, toolbars,
    // notification preferences, or settings applied by the other pages.
    c_tray._fHideClock = SHRestricted(REST_HIDECLOCK) ||
        ::IsDlgButtonChecked(hDlg, IDC_TRAYOPTSHOWCLOCK) != BST_CHECKED;
    ::SendMessage(c_tray.GetTrayNotifyHWND(), TNM_HIDECLOCK, 0, c_tray._fHideClock);
    c_tray.SizeWindows();

    BOOL smallIcons = ::IsDlgButtonChecked(hDlg, IDC_SMSMALLICONS) == BST_CHECKED;
    if (c_tray._fSMSmallIcons != smallIcons)
    {
        c_tray._fSMSmallIcons = smallIcons;
        IMenuPopup* menu = c_tray.GetStartMenu();
        IBanneredBar* banner = NULL;
        if (menu && SUCCEEDED(menu->QueryInterface(IID_PPV_ARG(IBanneredBar, &banner))))
        {
            banner->SetIconSize(smallIcons ? BMICON_SMALL : BMICON_LARGE);
            banner->Release();
        }
    }

    if (!SHRestricted(REST_INTELLIMENUS))
    {
        LPCTSTR value = ::IsDlgButtonChecked(hDlg, IDC_PERSONALIZEDMENUS) == BST_CHECKED
            ? TEXT("YES") : TEXT("No");
        SHRegSetUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("IntelliMenus"), REG_SZ,
            value, (lstrlen(value) + 1) * sizeof(TCHAR), SHREGSET_FORCE_HKCU);
        ::PostMessage(v_hwndTray, TM_REFRESH, 0, 0);
    }
    // Applying properties is not shell shutdown: do not run ClearRecentDocsOnExit.
    if (!SHRestricted(REST_NOSAVESET) && GetSystemMetrics(SM_CLEANBOOT) == 0)
        c_tray._SaveTrayStuff();
}

INT_PTR CTaskBarPropertySheet::GeneralDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_INITDIALOG:
        InitGeneral(hDlg);
        return TRUE;

    case WM_COMMAND:
        if (HIWORD(wParam) == BN_CLICKED)
        {
            switch (LOWORD(wParam))
            {
            case IDC_TRAYOPTONTOP:
            case IDC_TRAYOPTAUTOHIDE:
            case IDC_SMSMALLICONS:
            case IDC_TRAYOPTSHOWCLOCK:
            case IDC_PERSONALIZEDMENUS:
                UpdatePreview(hDlg);
                ::SendMessage(::GetParent(hDlg), PSM_CHANGED, reinterpret_cast<WPARAM>(hDlg), 0);
                return TRUE;
            }
        }
        break;

    case WM_SYSCOLORCHANGE:
        LoadBitmaps(hDlg);
        AttachBitmaps(hDlg);
        return TRUE;

    case WM_NOTIFY:
        switch (reinterpret_cast<NMHDR*>(lParam)->code)
        {
        case PSN_APPLY:
            ApplyGeneral(hDlg);
            ::SetWindowLongPtr(hDlg, DWLP_MSGRESULT, PSNRET_NOERROR);
            return TRUE;
        case PSN_KILLACTIVE:
        case PSN_SETACTIVE:
            ::SetWindowLongPtr(hDlg, DWLP_MSGRESULT, 0);
            return TRUE;
        }
        break;

    case WM_HELP:
        ::SHWinHelp(static_cast<HWND>(reinterpret_cast<HELPINFO*>(lParam)->hItemHandle),
            NULL, HELP_WM_HELP, reinterpret_cast<ULONG_PTR>(s_TaskOptionsHelpIDs));
        return TRUE;

    case WM_CONTEXTMENU:
        ::SHWinHelp(reinterpret_cast<HWND>(wParam), NULL, HELP_CONTEXTMENU,
            reinterpret_cast<ULONG_PTR>(s_TaskOptionsHelpIDs));
        return TRUE;

    case WM_DESTROY:
        DestroyBitmaps(hDlg);
        ::SetForegroundWindow(v_hwndTray);
        return TRUE;

    case WM_NCDESTROY:
        ::SetWindowLongPtr(hDlg, DWLP_USER, 0);
        break;
    }
    return FALSE;
}

INT_PTR CTaskBarPropertySheet::ExtraDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_INITDIALOG:
        ::CheckDlgButton(hDlg, IDC_EXTRAOPTWIN98, c_tray._fWin98 ? BST_CHECKED : BST_UNCHECKED);
        ::CheckDlgButton(hDlg, IDC_GROUPITEMS,
            !SHRestricted(REST_NOTASKGROUPING) &&
            SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED, TEXT("TaskbarGlomming"), FALSE, TRUE)
                ? BST_CHECKED : BST_UNCHECKED);
        ::EnableWindow(::GetDlgItem(hDlg, IDC_GROUPITEMS), !SHRestricted(REST_NOTASKGROUPING));
        return TRUE;

    case WM_COMMAND:
        if ((LOWORD(wParam) == IDC_EXTRAOPTWIN98 ||
             LOWORD(wParam) == IDC_GROUPITEMS) && HIWORD(wParam) == BN_CLICKED)
        {
            ::SendMessage(::GetParent(hDlg), PSM_CHANGED, reinterpret_cast<WPARAM>(hDlg), 0);
            return TRUE;
        }
        break;

    case WM_NOTIFY:
        switch (reinterpret_cast<NMHDR*>(lParam)->code)
        {
        case PSN_APPLY:
        {
            // Reuse the task band's existing grouping preference and rebuild
            // notification; changing this control takes effect on Apply, not click.
            if (!SHRestricted(REST_NOTASKGROUPING))
            {
                DWORD grouping = ::IsDlgButtonChecked(hDlg, IDC_GROUPITEMS) == BST_CHECKED;
                if (grouping != !!SHRegGetBoolUSValue(REGSTR_EXPLORER_ADVANCED,
                    TEXT("TaskbarGlomming"), FALSE, TRUE))
                {
                    LSTATUS error = SHRegSetUSValue(REGSTR_EXPLORER_ADVANCED,
                        TEXT("TaskbarGlomming"), REG_DWORD, &grouping, sizeof(grouping), SHREGSET_FORCE_HKCU);
                    if (error != ERROR_SUCCESS)
                    {
                        ::SetWindowLongPtr(hDlg, DWLP_MSGRESULT, PSNRET_INVALID_NOCHANGEPAGE);
                        return TRUE;
                    }
                    ::SendMessage(c_tray.GetTaskWindow(), WM_SETTINGCHANGE, 0,
                        reinterpret_cast<LPARAM>(TEXT("TraySettings")));
                }
            }
            BOOL win98 = ::IsDlgButtonChecked(hDlg, IDC_EXTRAOPTWIN98) == BST_CHECKED;
            BOOL fSave = FALSE;
            if (c_tray._fWin98 != win98)
            {
                c_tray._fWin98 = win98;
                c_tray.SizeWindows();
                ::InvalidateRect(v_hwndTray, NULL, TRUE);
                fSave = TRUE;
            }
            if (fSave && !SHRestricted(REST_NOSAVESET) && GetSystemMetrics(SM_CLEANBOOT) == 0)
                c_tray._SaveTrayStuff();
            ::SetWindowLongPtr(hDlg, DWLP_MSGRESULT, PSNRET_NOERROR);
            return TRUE;
        }
        case PSN_KILLACTIVE:
        case PSN_SETACTIVE:
            ::SetWindowLongPtr(hDlg, DWLP_MSGRESULT, 0);
            return TRUE;
        }
        break;

    case WM_NCDESTROY:
        ::SetWindowLongPtr(hDlg, DWLP_USER, 0);
        break;
    }
    return FALSE;
}

void DoTaskBarProperties(HWND hwnd, DWORD dwFlags)
{
    UINT startPage = (dwFlags & TPF_TASKBARPAGE) ? 0 : 1;
    CTaskBarPropertySheet sheet(startPage, hwnd);
    if (sheet.IsReady())
        sheet.DoModal(hwnd);
}
