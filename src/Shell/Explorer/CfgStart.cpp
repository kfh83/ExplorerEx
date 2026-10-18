// Windows 2000 Start Menu configuration (cfgstart.c).
#include "pch.h"
#include "cabinet.h"
#include "CfgStart.h"
#include "help.h"
#include "util.h"
#include "StartMnu.h"

namespace
{
const DWORD c_helpIDs[] = {
    IDC_NO_HELP_1, NO_HELP,
    IDC_NO_HELP_2, NO_HELP,
    IDC_NO_HELP_3, NO_HELP,
    IDC_NO_HELP_4, NO_HELP,
    IDC_GROUPBOX, IDH_COMM_GROUPBOX,
    IDC_ADDSHORTCUT, IDH_TRAY_ADD_PROGRAM,
    IDC_DELSHORTCUT, IDH_TRAY_REMOVE_PROGRAM,
    IDC_EXPLOREMENUS, IDH_TRAY_ADVANCED,
    IDC_KILLDOCUMENTS, IDH_MENUCONFIG_CLEAR,
    IDC_RESORT, IDH_TRAY_RESORT_BUTTON,
    IDC_STARTMENUSETTINGSTEXT, IDH_TRAY_START_MENU_SETTINGS,
    0, 0
};

const WCHAR* const c_mruKeys[] = {
    L"Software\\Microsoft\\Internet Explorer\\TypedURLs",
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU",
    L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Doc Find Spec MRU"
};

struct StartMenuSetting
{
    UINT textID;
    LPCWSTR valueName;
    DWORD valueType;
    DWORD helpID;
    DWORD policies[2];
};

// Exact choices, value types and policies from Win2k shell32/selfreg.inx.
// Modern RegTreeOptions ignores user-owned definition roots; keep this small,
// fixed tree local rather than changing Windows' machine-wide definitions.
const StartMenuSetting c_settings[] = {
    { IDS_STARTMENU_ADMINTOOLS, L"StartMenuAdminTools", REG_SZ, IDH_TRAY_DISPLAY_ADMIN_TOOLS, { 0, 0 } },
    { IDS_STARTMENU_FAVORITES, L"StartMenuFavorites", REG_DWORD, IDH_TRAY_DISPLAY_FAVORITES, { REST_NOFAVORITESMENU, 0 } },
    { IDS_STARTMENU_LOGOFF, L"StartMenuLogoff", REG_DWORD, IDH_TRAY_DISPLAY_LOGOFF, { REST_STARTMENULOGOFF, 0 } },
    { IDS_STARTMENU_CASCADECPANEL, L"CascadeControlPanel", REG_SZ, IDH_TRAY_EXPAND_CPANEL, { REST_NOSETFOLDERS, REST_NOCONTROLPANEL } },
    { IDS_STARTMENU_CASCADEMYDOCS, L"CascadeMyDocuments", REG_SZ, IDH_TRAY_EXPAND_MYDOCS, { 0, 0 } },
    { IDS_STARTMENU_CASCADENETCONNECT, L"CascadeNetworkConnections", REG_SZ, 51141, { REST_NOSETFOLDERS, REST_NONETWORKCONNECTIONS } },
    { IDS_STARTMENU_CASCADEPRINTERS, L"CascadePrinters", REG_SZ, IDH_TRAY_EXPAND_PRINTERS, { REST_NOSETFOLDERS, 0 } },
    { IDS_STARTMENU_SCROLLPROGRAMS, L"StartMenuScrollPrograms", REG_SZ, IDH_TRAY_SCROLL_PROGRAMS, { 0, 0 } }
};

bool SettingIsRestricted(const StartMenuSetting& setting)
{
    for (DWORD policy : setting.policies)
        if (policy && SHRestricted(static_cast<RESTRICTIONS>(policy)))
            return true;
    return false;
}

HRESULT ClearRecentHistory(BOOL broadcast)
{
    SHAddToRecentDocs(0, nullptr);
    HRESULT result = S_OK;
    for (const WCHAR* key : c_mruKeys)
    {
        const LSTATUS status = RegDeleteKeyW(HKEY_CURRENT_USER, key);
        if (status != ERROR_SUCCESS && status != ERROR_FILE_NOT_FOUND && SUCCEEDED(result))
            result = HRESULT_FROM_WIN32(status);
        if (broadcast)
            SHSendMessageBroadcast(WM_SETTINGCHANGE, 0, reinterpret_cast<LPARAM>(key));
    }
    return result;
}

HRESULT CallAppWiz(HWND hwnd, BOOL remove)
{
    WCHAR path[MAX_PATH];
    const UINT length = GetSystemDirectoryW(path, ARRAYSIZE(path));
    if (!length)
        return HRESULT_FROM_WIN32(GetLastError());
    if (length >= ARRAYSIZE(path))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    HRESULT hr = StringCchCatW(path, ARRAYSIZE(path), L"\\appwiz.cpl");
    if (FAILED(hr))
        return hr;

    HMODULE module = LoadLibraryW(path);
    if (!module)
        return HRESULT_FROM_WIN32(GetLastError());

    // ConfigStartMenu is an AppWiz C export, not a CALLBACK entry point.
    typedef BOOL (__cdecl* ConfigStartMenuProc)(HWND, BOOL);
    const auto configure = reinterpret_cast<ConfigStartMenuProc>(GetProcAddress(module, "ConfigStartMenu"));
    if (configure)
        // The export returns the wizard/dialog result, not a GetLastError code.
        hr = configure(hwnd, remove) ? S_OK : S_FALSE;
    else
        hr = HRESULT_FROM_WIN32(GetLastError());
    FreeLibrary(module);
    return hr;
}

HRESULT ExecExplorerAtStartMenu(HWND hwnd)
{
    WCHAR path[MAX_PATH];
    const UINT length = GetSystemWindowsDirectoryW(path, ARRAYSIZE(path));
    if (!length)
        return HRESULT_FROM_WIN32(GetLastError());
    if (length >= ARRAYSIZE(path))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    HRESULT hr = StringCchCatW(path, ARRAYSIZE(path), L"\\explorer.exe");
    if (FAILED(hr))
        return hr;

    WCHAR startMenu[MAX_PATH];
    hr = SHGetFolderPathW(hwnd, CSIDL_STARTMENU, nullptr, SHGFP_TYPE_CURRENT, startMenu);
    if (FAILED(hr))
        return hr;
    WCHAR parameters[MAX_PATH + 32];
    hr = StringCchPrintfW(parameters, ARRAYSIZE(parameters), L"%s\"%s\"",
        IsUserAnAdmin() ? L"/E," : L"/E,/Root,", startMenu);
    if (FAILED(hr))
        return hr;

    SHELLEXECUTEINFOW execute = { sizeof(execute) };
    execute.hwnd = hwnd;
    execute.fMask = SEE_MASK_FLAG_NO_UI;
    execute.lpFile = path;
    execute.lpParameters = parameters;
    execute.nShow = SW_SHOWDEFAULT;
    return ShellExecuteExW(&execute) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
}

HRESULT MenuOrderSort(HKEY key, IShellFolder* folder);

HRESULT MenuOrderSortKeyWithFolder(HKEY parent, LPCWSTR name, IShellFolder* folder)
{
    HKEY key = nullptr;
    const LSTATUS status = RegOpenKeyExW(parent, name, 0, KEY_READ | KEY_WRITE, &key);
    if (status == ERROR_FILE_NOT_FOUND)
        return S_OK;
    if (status != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(status);
    const HRESULT hr = MenuOrderSort(key, folder);
    RegCloseKey(key);
    return hr;
}

HRESULT MenuOrderSortSubKey(HKEY parent, LPWSTR folderName, LPCWSTR keyName, IShellFolder* folder)
{
    PIDLIST_RELATIVE pidl = nullptr;
    ULONG eaten = 0;
    DWORD attributes = 0;
    HRESULT hr = folder->ParseDisplayName(nullptr, nullptr, folderName, &eaten, &pidl, &attributes);
    if (FAILED(hr))
    {
        // Deleted folders can retain order keys. Do not delete those user records.
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
            return S_OK;
        return hr;
    }

    IShellFolder* child = nullptr;
    hr = folder->BindToObject(pidl, nullptr, IID_PPV_ARGS(&child));
    ILFree(pidl);
    if (SUCCEEDED(hr))
    {
        hr = MenuOrderSortKeyWithFolder(parent, keyName, child);
        child->Release();
    }
    return hr;
}

HRESULT MenuOrderSort(HKEY key, IShellFolder* folder)
{
    HRESULT result = S_OK;
    const LSTATUS orderStatus = RegQueryValueExW(key, L"Order", nullptr, nullptr, nullptr, nullptr);
    if (orderStatus == ERROR_SUCCESS)
    {
        IStream* stream = SHOpenRegStream2W(key, L"", L"Order", STGM_READWRITE);
        if (!stream)
            return E_FAIL;
        IOrderList2* order = nullptr;
        result = CoCreateInstance(CLSID_OrderListExport, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&order));
        if (SUCCEEDED(result))
        {
            HDPA items = nullptr;
            result = order->LoadFromStream(stream, &items, folder);
            if (SUCCEEDED(result) && items)
            {
                // The persisted array is name-sorted; nOrder holds the custom order.
                for (int i = DPA_GetPtrCount(items) - 1; i >= 0; --i)
                    static_cast<PORDERITEM>(DPA_FastGetPtr(items, i))->nOrder = i;
                LARGE_INTEGER zero = {};
                result = stream->Seek(zero, STREAM_SEEK_SET, nullptr);
                if (SUCCEEDED(result))
                    result = order->SaveToStream(stream, items);
                if (SUCCEEDED(result))
                    result = stream->Commit(STGC_DEFAULT);
            }
            if (items)
                order->FreeOrderList(items);
            order->Release();
        }
        stream->Release();
    }
    else if (orderStatus != ERROR_FILE_NOT_FOUND)
        result = HRESULT_FROM_WIN32(orderStatus);

    for (DWORD index = 0;; ++index)
    {
        WCHAR name[MAX_PATH];
        DWORD length = ARRAYSIZE(name);
        const LSTATUS status = RegEnumKeyExW(key, index, name, &length, nullptr, nullptr, nullptr, nullptr);
        if (status == ERROR_NO_MORE_ITEMS)
            break;
        if (status != ERROR_SUCCESS)
        {
            if (SUCCEEDED(result))
                result = HRESULT_FROM_WIN32(status);
            break;
        }
        const HRESULT hr = MenuOrderSortSubKey(key, name, name, folder);
        if (FAILED(hr) && SUCCEEDED(result))
            result = hr;
    }
    return result;
}

HRESULT StartMenuSort()
{
    // The migrated classic Programs menu persists separately from fast items.
    // Sort the key actually passed to InitializeProgramsShellMenu as well.
    HKEY programsKey = nullptr;
    LSTATUS programsStatus = RegOpenKeyExW(HKEY_CURRENT_USER,
        STRREG_STARTMENU L"XP\\Programs", 0, KEY_READ | KEY_WRITE, &programsKey);
    HRESULT programsResult = programsStatus == ERROR_FILE_NOT_FOUND
        ? S_OK : HRESULT_FROM_WIN32(programsStatus);
    if (programsStatus == ERROR_SUCCESS)
    {
        IShellFolder* programsFolder = nullptr;
        programsResult = CoCreateInstance(CLSID_ProgramsFolder, nullptr,
            CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&programsFolder));
        if (SUCCEEDED(programsResult))
        {
            programsResult = MenuOrderSort(programsKey, programsFolder);
            programsFolder->Release();
        }
        RegCloseKey(programsKey);
    }

    HKEY key = nullptr;
    const LSTATUS status = RegOpenKeyExW(HKEY_CURRENT_USER, STRREG_STARTMENU, 0, KEY_READ | KEY_WRITE, &key);
    if (status == ERROR_FILE_NOT_FOUND)
        return programsResult;
    if (status != ERROR_SUCCESS)
        return FAILED(programsResult) ? programsResult : HRESULT_FROM_WIN32(status);

    // This is ExplorerEx's migrated user/common Start Menu namespace, not the
    // filesystem folder alone (which cannot interpret merged-folder order PIDLs).
    IShellFolder* folder = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_StartMenuFolder, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&folder));
    if (SUCCEEDED(hr))
    {
        hr = MenuOrderSort(key, folder);
        PIDLIST_ABSOLUTE documents = nullptr;
        HRESULT childResult = SHGetFolderLocation(nullptr, CSIDL_PERSONAL, nullptr, 0, &documents);
        if (SUCCEEDED(childResult))
        {
            IShellFolder* documentsFolder = nullptr;
            childResult = SHBindToObjectEx(nullptr, documents, nullptr, IID_PPV_ARGS(&documentsFolder));
            if (SUCCEEDED(childResult))
            {
                childResult = MenuOrderSortKeyWithFolder(key, L"MyDocuments", documentsFolder);
                documentsFolder->Release();
            }
            ILFree(documents);
        }
        if (SUCCEEDED(hr) && FAILED(childResult))
            hr = childResult;

        WCHAR programs[MAX_PATH];
        childResult = SHGetFolderPathW(nullptr, CSIDL_PROGRAMS, nullptr, SHGFP_TYPE_CURRENT, programs);
        if (SUCCEEDED(childResult))
        {
            LPWSTR name = PathFindFileNameW(programs);
            if (lstrcmpiW(name, L"Programs") != 0)
                childResult = MenuOrderSortSubKey(key, name, L"Programs", folder);
        }
        if (SUCCEEDED(hr) && FAILED(childResult))
            hr = childResult;
        folder->Release();
    }
    RegCloseKey(key);
    return FAILED(programsResult) ? programsResult : hr;
}

void NotifyOrderChanged()
{
    SHChangeDWORDAsIDList change = {};
    change.cb = sizeof(change) - sizeof(change.cbZero);
    change.dwItem1 = SHCNEE_ORDERCHANGED;
    SHChangeNotify(SHCNE_EXTENDED_EVENT, SHCNF_FLUSH, reinterpret_cast<LPCITEMIDLIST>(&change), nullptr);
}
}

void ClearRecentDocumentsAndMRUStuff(BOOL fBroadcastChange)
{
    ClearRecentHistory(fBroadcastChange);
}

CStartMenuConfigPage::~CStartMenuConfigPage()
{
    _DestroyTree();
}

HPROPSHEETPAGE CStartMenuConfigPage::CreatePage(HINSTANCE resourceInstance)
{
    if (_pageCreated)
    {
        SetLastError(ERROR_BUSY);
        return nullptr;
    }
    _resourceInstance = resourceInstance;
    PROPSHEETPAGEW page = { sizeof(page) };
    page.dwFlags = PSP_USECALLBACK;
    page.hInstance = resourceInstance;
    page.pszTemplate = MAKEINTRESOURCEW(DLG_STARTMENU_CONFIG);
    page.pfnDlgProc = s_DlgProc;
    page.pfnCallback = s_PageCallback;
    page.lParam = reinterpret_cast<LPARAM>(this);
    _pageCreated = true;
    HPROPSHEETPAGE handle = CreatePropertySheetPageW(&page);
    if (!handle)
        _pageCreated = false;
    return handle;
}

UINT CALLBACK CStartMenuConfigPage::s_PageCallback(HWND, UINT message, LPPROPSHEETPAGE page)
{
    if (message == PSPCB_RELEASE)
        reinterpret_cast<CStartMenuConfigPage*>(page->lParam)->_pageCreated = false;
    return 1;
}

INT_PTR CALLBACK CStartMenuConfigPage::s_DlgProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    auto self = reinterpret_cast<CStartMenuConfigPage*>(GetWindowLongPtrW(hwnd, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        const auto page = reinterpret_cast<const PROPSHEETPAGEW*>(lParam);
        self = reinterpret_cast<CStartMenuConfigPage*>(page->lParam);
        self->_hwnd = hwnd;
        SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(self));
    }
    return self ? self->_DlgProc(message, wParam, lParam) : FALSE;
}

HRESULT CStartMenuConfigPage::_InitTree()
{
    static_assert(ARRAYSIZE(_settingItems) == ARRAYSIZE(c_settings), "Setting storage must match the original options");
    SetWindowLongPtrW(_hwndTree, GWL_STYLE,
        GetWindowLongPtrW(_hwndTree, GWL_STYLE) | TVS_CHECKBOXES | TVS_SHOWSELALWAYS);
    HRESULT hr = S_OK;
    for (UINT i = 0; i < ARRAYSIZE(c_settings); ++i)
    {
        const auto& setting = c_settings[i];
        if (SettingIsRestricted(setting))
            continue;
        WCHAR text[128];
        if (!LoadStringW(_resourceInstance, setting.textID, text, ARRAYSIZE(text)))
        {
            hr = HRESULT_FROM_WIN32(ERROR_RESOURCE_NAME_NOT_FOUND);
            break;
        }
        _initialChecked[i] = SHRegGetBoolUSValueW(REGSTR_EXPLORER_ADVANCED,
            setting.valueName, FALSE, FALSE);
        TVINSERTSTRUCTW item = {};
        item.hParent = TVI_ROOT;
        item.hInsertAfter = TVI_SORT;
        item.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_STATE;
        item.item.pszText = text;
        item.item.lParam = i;
        item.item.stateMask = TVIS_STATEIMAGEMASK;
        item.item.state = INDEXTOSTATEIMAGEMASK(_initialChecked[i] ? 2 : 1);
        _settingItems[i] = TreeView_InsertItem(_hwndTree, &item);
        if (!_settingItems[i])
        {
            hr = E_OUTOFMEMORY;
            break;
        }
    }
    _stateImages = TreeView_GetImageList(_hwndTree, TVSIL_STATE);
    return hr;
}

HRESULT CStartMenuConfigPage::_SaveTree()
{
    if (SHRestricted(REST_NOCHANGESTARMENU))
        return E_ACCESSDENIED;
    HKEY key = nullptr;
    LSTATUS status = RegCreateKeyExW(HKEY_CURRENT_USER, REGSTR_EXPLORER_ADVANCED,
        0, nullptr, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, nullptr, &key, nullptr);
    if (status != ERROR_SUCCESS)
        return HRESULT_FROM_WIN32(status);
    for (UINT i = 0; i < ARRAYSIZE(c_settings); ++i)
    {
        if (!_settingItems[i] || SettingIsRestricted(c_settings[i]))
            continue;
        const BOOL checked = TreeView_GetCheckState(_hwndTree, _settingItems[i]);
        if (checked == _initialChecked[i])
            continue;
        if (c_settings[i].valueType == REG_DWORD)
        {
            const DWORD value = checked ? 1 : 0;
            status = RegSetValueExW(key, c_settings[i].valueName, 0, REG_DWORD,
                reinterpret_cast<const BYTE*>(&value), sizeof(value));
        }
        else
        {
            LPCWSTR value = checked ? L"YES" : L"NO";
            status = RegSetValueExW(key, c_settings[i].valueName, 0, REG_SZ,
                reinterpret_cast<const BYTE*>(value), (lstrlenW(value) + 1) * sizeof(WCHAR));
        }
        if (status != ERROR_SUCCESS)
            break;
        _initialChecked[i] = checked;
    }
    RegCloseKey(key);
    return HRESULT_FROM_WIN32(status);
}

void CStartMenuConfigPage::_InitDialog()
{
    _hwndTree = GetDlgItem(_hwnd, IDC_STARTMENUSETTINGS);
    _treeDirty = false;
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr))
    {
        _comInitialized = true;
        hr = _InitTree();
        _treeReady = SUCCEEDED(hr);
    }
    const bool canCustomize = !SHRestricted(REST_NOCHANGESTARMENU);
    EnableWindow(_hwndTree, _treeReady && canCustomize);
    EnableWindow(GetDlgItem(_hwnd, IDC_ADDSHORTCUT), canCustomize);
    EnableWindow(GetDlgItem(_hwnd, IDC_DELSHORTCUT), canCustomize);
    EnableWindow(GetDlgItem(_hwnd, IDC_EXPLOREMENUS), canCustomize);
    EnableWindow(GetDlgItem(_hwnd, IDC_RESORT), _comInitialized && canCustomize);
    _UpdateClearButton();
    if (_treeReady)
    {
        const HTREEITEM root = TreeView_GetRoot(_hwndTree);
        if (root)
            TreeView_EnsureVisible(_hwndTree, root);
    }
    else
        _ShowError(hr);
}

void CStartMenuConfigPage::_DestroyTree()
{
    _treeReady = false;
    _treeDirty = false;
    if (_stateImages)
    {
        if (IsWindow(_hwndTree))
            TreeView_SetImageList(_hwndTree, nullptr, TVSIL_STATE);
        ImageList_Destroy(_stateImages);
        _stateImages = nullptr;
    }
    ZeroMemory(_settingItems, sizeof(_settingItems));
    if (_comInitialized)
    {
        _comInitialized = false;
        CoUninitialize();
    }
    _hwndTree = nullptr;
}

void CStartMenuConfigPage::_ShowError(HRESULT hr) const
{
    WCHAR format[256];
    WCHAR message[1024];
    WCHAR detail[512] = {};
    WCHAR caption[128] = {};
    if (!LoadStringW(_resourceInstance, IDS_STARTMENU_ERROR, format, ARRAYSIZE(format)))
        return;
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr,
        static_cast<DWORD>(hr), 0, detail, ARRAYSIZE(detail), nullptr);
    StringCchPrintfW(message, ARRAYSIZE(message), format, static_cast<ULONG>(hr), detail);
    GetWindowTextW(GetParent(_hwnd), caption, ARRAYSIZE(caption));
    MessageBoxW(_hwnd, message, caption, MB_OK | MB_ICONERROR);
}

void CStartMenuConfigPage::_UpdateClearButton()
{
    bool hasHistory = false;
    if (_comInitialized)
    {
        PIDLIST_ABSOLUTE recent = nullptr;
        if (SUCCEEDED(SHGetSpecialFolderLocation(_hwnd, CSIDL_RECENT, &recent)))
        {
            IShellFolder* folder = nullptr;
            if (SUCCEEDED(SHBindToObjectEx(nullptr, recent, nullptr, IID_PPV_ARGS(&folder))))
            {
                IEnumIDList* entries = nullptr;
                if (folder->EnumObjects(_hwnd, SHCONTF_FOLDERS | SHCONTF_NONFOLDERS, &entries) == S_OK && entries)
                {
                    PITEMID_CHILD item = nullptr;
                    ULONG fetched = 0;
                    hasHistory = entries->Next(1, &item, &fetched) == S_OK && fetched == 1;
                    ILFree(item);
                    entries->Release();
                }
                folder->Release();
            }
            ILFree(recent);
        }
        for (const WCHAR* key : c_mruKeys)
        {
            if (hasHistory)
                break;
            HKEY history = nullptr;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, key, 0, KEY_READ, &history) == ERROR_SUCCESS)
            {
                hasHistory = true;
                RegCloseKey(history);
            }
        }
    }
    EnableWindow(GetDlgItem(_hwnd, IDC_KILLDOCUMENTS), hasHistory);
}

void CStartMenuConfigPage::_ToggleItem(HTREEITEM item)
{
    if (!_treeReady || !IsWindowEnabled(_hwndTree) || !item)
        return;
    TreeView_SetCheckState(_hwndTree, item, !TreeView_GetCheckState(_hwndTree, item));
}

void CStartMenuConfigPage::_ShowTreeHelp(POINT point, DWORD command, bool keyboard)
{
    if (!_treeReady)
        return;
    HTREEITEM item;
    if (keyboard)
        item = TreeView_GetSelection(_hwndTree);
    else
    {
        TVHITTESTINFO hit = {};
        hit.pt = point;
        ScreenToClient(_hwndTree, &hit.pt);
        item = TreeView_HitTest(_hwndTree, &hit);
    }
    for (UINT i = 0; item && i < ARRAYSIZE(c_settings); ++i)
    {
        if (_settingItems[i] != item)
            continue;
        const DWORD helpIDs[] = { IDC_STARTMENUSETTINGS, c_settings[i].helpID, 0, 0 };
        SHWinHelp(_hwndTree, L"windows.hlp", command, reinterpret_cast<ULONG_PTR>(helpIDs));
        break;
    }
}

INT_PTR CStartMenuConfigPage::_DlgProc(UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
        _InitDialog();
        return TRUE;

    case WM_COMMAND:
        if (HIWORD(wParam) != BN_CLICKED || _busy)
            return FALSE;
        {
            HRESULT hr = S_OK;
            _busy = true;
            switch (LOWORD(wParam))
            {
            case IDC_ADDSHORTCUT:
            case IDC_DELSHORTCUT:
                hr = CallAppWiz(_hwnd, LOWORD(wParam) == IDC_DELSHORTCUT);
                break;
            case IDC_EXPLOREMENUS:
                hr = ExecExplorerAtStartMenu(_hwnd);
                break;
            case IDC_RESORT:
                if (_comInitialized)
                {
                    HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
                    hr = StartMenuSort();
                    SetCursor(previous);
                    // Even partial writes must invalidate consumers' cached order.
                    NotifyOrderChanged();
                }
                else
                    hr = CO_E_NOTINITIALIZED;
                break;
            case IDC_KILLDOCUMENTS:
                if (_comInitialized)
                {
                    HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
                    hr = ClearRecentHistory(TRUE);
                    SetCursor(previous);
                    if (_hwnd)
                    {
                        _UpdateClearButton();
                        if (!IsWindowEnabled(GetDlgItem(_hwnd, IDC_KILLDOCUMENTS)))
                            SetFocus(GetDlgItem(GetParent(_hwnd), IDOK));
                    }
                }
                else
                    hr = CO_E_NOTINITIALIZED;
                break;
            default:
                _busy = false;
                return FALSE;
            }
            if (FAILED(hr) && _hwnd)
                _ShowError(hr);
            _busy = false;
            return TRUE;
        }

    case WM_NOTIFY:
        {
            const NMHDR* notification = reinterpret_cast<const NMHDR*>(lParam);
            if (notification->code == PSN_APPLY)
            {
                HRESULT hr = S_OK;
                if (_treeDirty)
                {
                    hr = _treeReady ? _SaveTree() : E_UNEXPECTED;
                    if (SUCCEEDED(hr))
                    {
                        _treeDirty = false;
                        SHSendMessageBroadcast(WM_SETTINGCHANGE, 0, 0);
                        SHSendMessageBroadcast(WM_SETTINGCHANGE, 0, reinterpret_cast<LPARAM>(L"TraySettings"));
                        PostMessageW(v_hwndTray, SBM_REBUILDMENU, 0, 0);
                    }
                    else
                        _ShowError(hr);
                }
                if (_hwnd)
                    SetWindowLongPtrW(_hwnd, DWLP_MSGRESULT, FAILED(hr) ? PSNRET_INVALID_NOCHANGEPAGE : PSNRET_NOERROR);
                return TRUE;
            }
            if (notification->code == PSN_SETACTIVE)
            {
                _UpdateClearButton();
                SetWindowLongPtrW(_hwnd, DWLP_MSGRESULT, 0);
                return TRUE;
            }
            if (notification->hwndFrom != _hwndTree || !_treeReady)
                return FALSE;
            switch (notification->code)
            {
            case TVN_ITEMCHANGEDW:
                {
                    const auto change = reinterpret_cast<const NMTVITEMCHANGE*>(lParam);
                    if ((change->uStateOld ^ change->uStateNew) & TVIS_STATEIMAGEMASK)
                    {
                        _treeDirty = true;
                        PropSheet_Changed(GetParent(_hwnd), _hwnd);
                    }
                    break;
                }
            case NM_CLICK:
            case NM_DBLCLK:
                {
                    const DWORD position = GetMessagePos();
                    TVHITTESTINFO hit = {};
                    hit.pt = { GET_X_LPARAM(position), GET_Y_LPARAM(position) };
                    ScreenToClient(_hwndTree, &hit.pt);
                    HTREEITEM item = TreeView_HitTest(_hwndTree, &hit);
                    // State-icon clicks and Space are handled by TVS_CHECKBOXES.
                    if (hit.flags & (TVHT_ONITEMICON | TVHT_ONITEMLABEL))
                        _ToggleItem(item);
                    return FALSE;
                }
            case NM_RCLICK:
                {
                    const DWORD position = GetMessagePos();
                    POINT point = { GET_X_LPARAM(position), GET_Y_LPARAM(position) };
                    _ShowTreeHelp(point, HELP_CONTEXTMENU, false);
                    SetWindowLongPtrW(_hwnd, DWLP_MSGRESULT, TRUE);
                    return TRUE;
                }
            }
        }
        return FALSE;

    case WM_HELP:
        {
            const HELPINFO* help = reinterpret_cast<const HELPINFO*>(lParam);
            if (help->iCtrlId == IDC_STARTMENUSETTINGS)
                _ShowTreeHelp(help->MousePos, HELP_WM_HELP, GetKeyState(VK_F1) < 0);
            else
                SHWinHelp(static_cast<HWND>(help->hItemHandle), nullptr, HELP_WM_HELP,
                    reinterpret_cast<ULONG_PTR>(c_helpIDs));
            return TRUE;
        }

    case WM_CONTEXTMENU:
        if (reinterpret_cast<HWND>(wParam) == _hwndTree)
        {
            POINT point = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            _ShowTreeHelp(point, HELP_CONTEXTMENU, point.x == -1 && point.y == -1);
        }
        else
            SHWinHelp(reinterpret_cast<HWND>(wParam), nullptr, HELP_CONTEXTMENU,
                reinterpret_cast<ULONG_PTR>(c_helpIDs));
        return TRUE;

    case WM_DESTROY:
        SetWindowLongPtrW(_hwnd, DWLP_USER, 0);
        _hwnd = nullptr;
        _DestroyTree();
        return TRUE;
    }
    return FALSE;
}
