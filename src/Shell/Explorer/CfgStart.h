#pragma once

#include <windows.h>
#include <commctrl.h>
#include <prsht.h>

void ClearRecentDocumentsAndMRUStuff(BOOL fBroadcastChange);

// The property-sheet owner keeps this object alive until its modal sheet closes.
class CStartMenuConfigPage
{
public:
    CStartMenuConfigPage() = default;
    ~CStartMenuConfigPage();

    HPROPSHEETPAGE CreatePage(HINSTANCE resourceInstance);

    CStartMenuConfigPage(const CStartMenuConfigPage&) = delete;
    CStartMenuConfigPage& operator=(const CStartMenuConfigPage&) = delete;

private:
    static INT_PTR CALLBACK s_DlgProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
    static UINT CALLBACK s_PageCallback(HWND hwnd, UINT message, LPPROPSHEETPAGE page);
    INT_PTR _DlgProc(UINT message, WPARAM wParam, LPARAM lParam);
    void _InitDialog();
    HRESULT _InitTree();
    HRESULT _SaveTree();
    void _DestroyTree();
    void _ShowError(HRESULT hr) const;
    void _UpdateClearButton();
    void _ToggleItem(HTREEITEM item);
    void _ShowTreeHelp(POINT point, DWORD command, bool keyboard);

    HINSTANCE _resourceInstance = nullptr;
    HWND _hwnd = nullptr;
    HWND _hwndTree = nullptr;
    HTREEITEM _settingItems[8] = {};
    BOOL _initialChecked[8] = {};
    HIMAGELIST _stateImages = nullptr;
    bool _comInitialized = false;
    bool _pageCreated = false;
    bool _treeReady = false;
    bool _treeDirty = false;
    bool _busy = false;
};
