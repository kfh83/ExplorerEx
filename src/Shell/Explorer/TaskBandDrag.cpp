#include "pch.h"
#include "cabinet.h"
#include "TaskBand.h"

// Taskix64.dll 0x10001C30 / 0x10001420 / 0x100017C0: left-drag task buttons,
// five-pixel drag threshold, group-aware moves, and no activation after a move.
// Taskix64.exe configures/activates this behavior; reordering lives in its DLL.
// This host-owned implementation uses no Taskix registry keys or global input
// hooks. Order lasts for the shell session, including grouping changes.

namespace
{
    BOOL ReorderThresholdReached(POINT down, POINT point)
    {
        return point.x - down.x > 5 || down.x - point.x > 5 ||
               point.y - down.y > 5 || down.y - point.y > 5;
    }

    BOOL HitReorderButton(HWND toolbar, POINT point, TBBUTTON* button)
    {
        RECT client;
        if (!GetClientRect(toolbar, &client) || !PtInRect(&client, point))
            return FALSE;

        int index = (int)SendMessage(toolbar, TB_HITTEST, 0, (LPARAM)&point);
        return index >= 0 && SendMessage(toolbar, TB_GETBUTTON, index, (LPARAM)button) &&
               (button->fsState & TBSTATE_ENABLED) && !(button->fsState & TBSTATE_HIDDEN) &&
               !(button->fsStyle & BTNS_SEP);
    }

    PTASKITEM LiveReorderItem(HWND toolbar, UINT command, PTASKITEM expected = NULL)
    {
        int index = (int)SendMessage(toolbar, TB_COMMANDTOINDEX, command, 0);
        TBBUTTON button = {};
        if (index < 0 || !SendMessage(toolbar, TB_GETBUTTON, index, (LPARAM)&button))
            return NULL;

        PTASKITEM item = (PTASKITEM)button.dwData;
        // Command IDs can be reused. Never dereference the saved drag pointer.
        if (!item || (expected && item != expected) || item->hwnd == HWND_TOPMOST)
            return NULL;
        return item;
    }

    void ReorderTooltips(HWND toolbar, BOOL enable)
    {
        HWND tooltip = (HWND)SendMessage(toolbar, TB_GETTOOLTIPS, 0, 0);
        if (tooltip)
        {
            if (!enable)
                SendMessage(tooltip, TTM_POP, 0, 0);
            SendMessage(tooltip, TTM_ACTIVATE, enable, 0);
        }
    }

    BOOL IsReorderToolbar(HWND hwnd)
    {
        TCHAR className[32];
        return GetClassName(hwnd, className, ARRAYSIZE(className)) &&
               !lstrcmpi(className, TOOLBARCLASSNAME);
    }

    BOOL CALLBACK FindReorderToolbar(HWND hwnd, LPARAM parameter)
    {
        if (IsWindowVisible(hwnd) && IsReorderToolbar(hwnd))
        {
            *(HWND*)parameter = hwnd;
            return FALSE;
        }
        return TRUE;
    }

    BOOL PopupOwnsReorderCapture(HWND toolbar, HWND capture)
    {
        HWND root = GetAncestor(toolbar, GA_ROOT);
        return !capture || capture == root || IsChild(root, capture);
    }

    HRESULT ConsumeReorderMessage(MSG* message)
    {
        message->message = WM_NULL;
        return S_OK;
    }
}

LRESULT CALLBACK CTaskBand::s_TaskReorderSubclass(HWND hwnd, UINT message, WPARAM wParam,
    LPARAM lParam, UINT_PTR subclass, DWORD_PTR reference)
{
    CTaskBand* taskBand = (CTaskBand*)reference;
    if (message == WM_NCDESTROY)
        RemoveWindowSubclass(hwnd, s_TaskReorderSubclass, subclass);
    return taskBand ? taskBand->_TaskReorderMessage(hwnd, message, wParam, lParam) :
                     DefSubclassProc(hwnd, message, wParam, lParam);
}

void CTaskBand::_CancelTaskReorder()
{
    HWND toolbar = _hwndReorderToolbar;
    if (!toolbar)
        return;

    UINT command = _idReorderSource;
    PTASKITEM item = _ptiReorderSource;
    BOOL mainToolbar = toolbar == (HWND)_tb;
    BOOL started = _fReorderStarted;
    HCURSOR cursor = _hcurBeforeReorder;

    // Tooltip, toolbar and capture messages can reenter us or delete an item.
    _hwndReorderToolbar = NULL;
    _idReorderSource = 0;
    _ptiReorderSource = NULL;
    _ptReorderDown = {};
    _wParamReorderDown = 0;
    _uReorderDownMessage = WM_LBUTTONDOWN;
    _fReorderStarted = FALSE;
    _fReorderMoved = FALSE;
    _hcurBeforeReorder = NULL;
    _fReorderSuppressUp = TRUE;

    if (mainToolbar)
    {
        KillTimer(toolbar, (UINT_PTR)&s_TaskReorderSubclass);
        if (item && LiveReorderItem(toolbar, command, item))
            SendMessage(toolbar, TB_PRESSBUTTON, command, FALSE);
    }
    if (IsWindow(toolbar) && (mainToolbar || started))
        ReorderTooltips(toolbar, TRUE);
    if (mainToolbar && GetCapture() == toolbar)
        ReleaseCapture();
    if (started)
        SetCursor(cursor);
}

LRESULT CTaskBand::_TaskReorderMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    // Retain the toolbar's existing protection against self-capture notifications.
    if (message == WM_CAPTURECHANGED && (HWND)lParam == hwnd)
        return 0;

    if (message == WM_NCDESTROY)
    {
        if (_hwndPopupReorderToolbar == hwnd)
            _hwndPopupReorderToolbar = NULL;
        if (_hwndReorderToolbar == hwnd)
            _CancelTaskReorder();
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }
    if (_fReorderReplay)
        return DefSubclassProc(hwnd, message, wParam, lParam);

    if (hwnd != (HWND)_tb)
    {
        MSG popupMessage = {};
        popupMessage.hwnd = hwnd;
        popupMessage.message = message;
        popupMessage.wParam = wParam;
        popupMessage.lParam = lParam;
        switch (message)
        {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
        case WM_LBUTTONUP:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONDBLCLK:
        case WM_MBUTTONDOWN:
        case WM_MBUTTONDBLCLK:
        case WM_XBUTTONDOWN:
        case WM_XBUTTONDBLCLK:
            {
                POINT screen = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
                if (!ClientToScreen(hwnd, &screen))
                    return DefSubclassProc(hwnd, message, wParam, lParam);
                popupMessage.lParam = MAKELPARAM(screen.x, screen.y);
                popupMessage.pt = screen;
            }
            break;

        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
        case WM_KEYDOWN:
        case WM_SYSKEYDOWN:
            break;

        case WM_SETCURSOR:
            if (_hwndReorderToolbar == hwnd && _fReorderStarted && LOWORD(lParam) == HTCLIENT)
            {
                SetCursor(LoadCursor(NULL, IDC_SIZENS));
                return TRUE;
            }
            return DefSubclassProc(hwnd, message, wParam, lParam);

        default:
            return DefSubclassProc(hwnd, message, wParam, lParam);
        }

        // Menuband's SMC_MOUSEFILTER stops seeing mouse messages after its
        // toolbar takes capture from the popup root. Use the same filter for
        // those direct messages without taking or releasing menu capture.
        if (_FilterPopupTaskReorder(&popupMessage, TRUE) == S_OK)
            return 0;
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }

    BOOL pending = _hwndReorderToolbar == hwnd && _ptiReorderSource;
    switch (message)
    {
    case WM_CAPTURECHANGED:
    case WM_CANCELMODE:
    case WM_RBUTTONDOWN:
    case WM_RBUTTONDBLCLK:
    case WM_MBUTTONDOWN:
    case WM_MBUTTONDBLCLK:
    case WM_XBUTTONDOWN:
    case WM_XBUTTONDBLCLK:
        if (pending)
            _CancelTaskReorder();
        break;

    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if (pending && wParam == VK_ESCAPE)
        {
            _CancelTaskReorder();
            return 0;
        }
        break;

    case WM_TIMER:
        if (wParam == (UINT_PTR)&s_TaskReorderSubclass)
        {
            // Capture does not redirect keyboard input; leave focus with the app.
            if (pending && (GetCapture() != hwnd || (GetAsyncKeyState(VK_ESCAPE) & 0x8000)))
                _CancelTaskReorder();
            return 0;
        }
        break;

    case WM_SETCURSOR:
        if (pending && _fReorderStarted && LOWORD(lParam) == HTCLIENT)
        {
            SetCursor(LoadCursor(NULL, _IsHorizontal() ? IDC_SIZEWE : IDC_SIZENS));
            return TRUE;
        }
        break;

    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
        {
            _CancelTaskReorder();
            _fReorderSuppressUp = FALSE;
            HWND capture = GetCapture();
            if ((capture && capture != hwnd) ||
                (wParam & (MK_RBUTTON | MK_MBUTTON | MK_XBUTTON1 | MK_XBUTTON2)))
                break;

            POINT point = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            TBBUTTON button = {};
            if (!HitReorderButton(hwnd, point, &button))
                break;
            PTASKITEM item = LiveReorderItem(hwnd, button.idCommand);
            if (!item)
                break;

            _hwndReorderToolbar = hwnd;
            _idReorderSource = button.idCommand;
            _ptiReorderSource = item;
            _ptReorderDown = point;
            _wParamReorderDown = wParam;
            _uReorderDownMessage = message;
            _iIndexActiveAtLDown = _GetCurSel();
            SetCapture(hwnd);
            if (_hwndReorderToolbar != hwnd || GetCapture() != hwnd)
            {
                _CancelTaskReorder();
                return 0;
            }
            SendMessage(hwnd, TB_PRESSBUTTON, button.idCommand, TRUE);
            if (_hwndReorderToolbar == hwnd)
            {
                SetTimer(hwnd, (UINT_PTR)&s_TaskReorderSubclass, 50, NULL);
                ReorderTooltips(hwnd, FALSE);
            }
            return 0;
        }

    case WM_MOUSEMOVE:
        if (pending)
        {
            if (!(wParam & MK_LBUTTON) || GetCapture() != hwnd ||
                (GetAsyncKeyState(VK_ESCAPE) & 0x8000) ||
                !LiveReorderItem(hwnd, _idReorderSource, _ptiReorderSource))
            {
                _CancelTaskReorder();
                return 0;
            }

            POINT point = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            if (!_fReorderStarted && ReorderThresholdReached(_ptReorderDown, point))
            {
                _hcurBeforeReorder = GetCursor();
                _fReorderStarted = TRUE;
            }
            if (_fReorderStarted)
            {
                SetCursor(LoadCursor(NULL, _IsHorizontal() ? IDC_SIZEWE : IDC_SIZENS));
                TBBUTTON target = {};
                UINT source = _idReorderSource;
                PTASKITEM item = _ptiReorderSource;
                if (HitReorderButton(hwnd, point, &target) && (UINT)target.idCommand != source &&
                    _MoveTaskButton(source, target.idCommand, FALSE) &&
                    _hwndReorderToolbar == hwnd && _ptiReorderSource == item &&
                    LiveReorderItem(hwnd, source, item))
                {
                    _fReorderMoved = TRUE;
                    // A block move can put another group under the pointer. Keep
                    // the dragged button under it instead of oscillating groups.
                    POINT cursor;
                    if (GetCursorPos(&cursor) && ScreenToClient(hwnd, &cursor) &&
                        HitReorderButton(hwnd, cursor, &target) && (UINT)target.idCommand != source)
                    {
                        RECT button;
                        int index = (int)SendMessage(hwnd, TB_COMMANDTOINDEX, source, 0);
                        if (index >= 0 && SendMessage(hwnd, TB_GETITEMRECT, index, (LPARAM)&button))
                        {
                            cursor.x = button.left + (button.right - button.left) / 2;
                            cursor.y = button.top + (button.bottom - button.top) / 2;
                            if (ClientToScreen(hwnd, &cursor))
                                SetCursorPos(cursor.x, cursor.y);
                        }
                    }
                }
            }
            return 0;
        }
        break;

    case WM_LBUTTONUP:
        if (pending)
        {
            POINT point = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            TBBUTTON button = {};
            UINT source = _idReorderSource;
            PTASKITEM item = _ptiReorderSource;
            BOOL click = !_fReorderMoved && LiveReorderItem(hwnd, source, item) &&
                HitReorderButton(hwnd, point, &button) && (UINT)button.idCommand == source;
            POINT down = _ptReorderDown;
            WPARAM downKeys = _wParamReorderDown;
            UINT downMessage = _uReorderDownMessage;
            _CancelTaskReorder();
            _fReorderSuppressUp = FALSE;
            if (click && IsWindow(hwnd) && LiveReorderItem(hwnd, source, item))
            {
                // Native notifications preserve Ctrl, dropdown and minimize
                // behavior. Do not select a task or move focus ourselves.
                AddRef();
                _fReorderReplay = TRUE;
                DefSubclassProc(hwnd, downMessage, downKeys, MAKELPARAM(down.x, down.y));
                LRESULT result = 0;
                if (IsWindow(hwnd))
                {
                    if (LiveReorderItem(hwnd, source, item))
                        result = DefSubclassProc(hwnd, WM_LBUTTONUP, wParam, lParam);
                    else
                        DefSubclassProc(hwnd, WM_CANCELMODE, 0, 0);
                }
                _fReorderReplay = FALSE;
                Release();
                return result;
            }
            return 0;
        }
        if (_fReorderSuppressUp)
        {
            _fReorderSuppressUp = FALSE;
            return 0;
        }
        break;
    }
    return DefSubclassProc(hwnd, message, wParam, lParam);
}

HRESULT CTaskBand::_FilterPopupTaskReorder(MSG* message, BOOL remove)
{
    if (!message || _fReorderReplay || _hwndReorderToolbar == (HWND)_tb)
        return S_FALSE;

    UINT type = message->message;
    BOOL pending = _hwndReorderToolbar && _ptiReorderSource;
    BOOL live = !pending || LiveReorderItem(_tb, _idReorderSource, _ptiReorderSource);
    POINT screen = { GET_X_LPARAM(message->lParam), GET_Y_LPARAM(message->lParam) };
    BOOL leftHeld = (message->wParam & MK_LBUTTON) != 0;

    // Peeking must not start/cancel a gesture, discover/cache a window, or move
    // any buttons. The returned MSG may still be nulled to suppress processing.
    if (!remove)
    {
        if (type == WM_LBUTTONUP && (_fReorderSuppressUp || (pending && (_fReorderMoved || !live))))
            return ConsumeReorderMessage(message);
        if (type == WM_MOUSEMOVE && pending && leftHeld &&
            (!live || _fReorderStarted || ReorderThresholdReached(_ptReorderDown, screen)))
            return ConsumeReorderMessage(message);
        return S_FALSE;
    }

    if (pending && !live)
    {
        _CancelTaskReorder();
        if (type == WM_MOUSEMOVE)
            return ConsumeReorderMessage(message);
        pending = FALSE;
    }

    if (type == WM_LBUTTONUP)
    {
        BOOL consume = _fReorderSuppressUp || (pending && _fReorderMoved);
        HWND toolbar = pending ? _hwndReorderToolbar : _hwndPopupReorderToolbar;
        if (pending)
            _CancelTaskReorder();
        _fReorderSuppressUp = FALSE;
        if (consume && IsWindow(toolbar))
        {
            // The native toolbar saw the original down and may own capture.
            // Finish its press outside every item, as Taskix does, rather than
            // activate the task now under the real pointer or strand capture.
            AddRef();
            _fReorderReplay = TRUE;
            SendMessage(toolbar, WM_MOUSEMOVE, MK_LBUTTON, MAKELPARAM(-1, -1));
            if (IsWindow(toolbar))
                SendMessage(toolbar, WM_LBUTTONUP, 0, MAKELPARAM(-1, -1));
            _fReorderReplay = FALSE;
            Release();
        }
        return consume ? ConsumeReorderMessage(message) : S_FALSE;
    }

    if (type == WM_CANCELMODE || ((type == WM_KEYDOWN || type == WM_SYSKEYDOWN) && message->wParam == VK_ESCAPE) ||
        type == WM_RBUTTONDOWN || type == WM_RBUTTONDBLCLK ||
        type == WM_MBUTTONDOWN || type == WM_MBUTTONDBLCLK ||
        type == WM_XBUTTONDOWN || type == WM_XBUTTONDBLCLK)
    {
        if (pending)
            _CancelTaskReorder();
        return S_FALSE;
    }
    if (type == WM_CAPTURECHANGED)
    {
        if (pending && (HWND)message->lParam != message->hwnd &&
            (!message->lParam || !PopupOwnsReorderCapture(_hwndReorderToolbar, (HWND)message->lParam)))
            _CancelTaskReorder();
        return S_FALSE;
    }

    BOOL down = type == WM_LBUTTONDOWN || type == WM_LBUTTONDBLCLK;
    if (!down && type != WM_MOUSEMOVE)
        return S_FALSE;
    if (down)
    {
        if (pending)
            _CancelTaskReorder();
        _fReorderSuppressUp = FALSE;
        pending = FALSE;
    }
    else if (!leftHeld)
    {
        if (pending)
            _CancelTaskReorder();
        return S_FALSE;
    }
    if (!_psmPopup || (message->wParam & (MK_RBUTTON | MK_MBUTTON | MK_XBUTTON1 | MK_XBUTTON2)))
    {
        if (pending)
            _CancelTaskReorder();
        return S_FALSE;
    }

    HWND toolbar = _hwndPopupReorderToolbar;
    if (!IsWindow(toolbar))
    {
        HWND menuWindow = NULL;
        if (SUCCEEDED(IUnknown_GetWindow(_psmPopup, &menuWindow)) && menuWindow)
        {
            toolbar = NULL;
            if (IsReorderToolbar(menuWindow))
                toolbar = menuWindow;
            else
                EnumChildWindows(menuWindow, FindReorderToolbar, (LPARAM)&toolbar);
        }
        else
            toolbar = NULL;
        if (toolbar && !SetWindowSubclass(toolbar, s_TaskReorderSubclass, 1, (DWORD_PTR)this))
            toolbar = NULL;
        _hwndPopupReorderToolbar = toolbar;
    }
    // In particular, do not turn an OLE file drag with foreign capture into a
    // task reorder. The menu keeps ownership of its own capture throughout.
    if (!toolbar || !PopupOwnsReorderCapture(toolbar, GetCapture()))
    {
        if (pending)
            _CancelTaskReorder();
        return S_FALSE;
    }

    if (pending && _hwndReorderToolbar != toolbar)
    {
        if (_fReorderStarted)
        {
            if (IsWindow(_hwndReorderToolbar))
                ReorderTooltips(_hwndReorderToolbar, TRUE);
            ReorderTooltips(toolbar, FALSE);
        }
        _hwndReorderToolbar = toolbar;
    }
    POINT point = screen; // SMC_MOUSEFILTER lParam is screen-relative, unlike WM_MOUSEMOVE.
    if (!ScreenToClient(toolbar, &point))
        return S_FALSE;
    TBBUTTON target = {};
    BOOL hit = HitReorderButton(toolbar, point, &target);
    if (!pending)
    {
        if (_fReorderSuppressUp || !hit)
            return S_FALSE;
        PTASKITEM item = LiveReorderItem(_tb, target.idCommand);
        if (!item || !item->hwnd)
            return S_FALSE;
        _hwndReorderToolbar = toolbar;
        _idReorderSource = target.idCommand;
        _ptiReorderSource = item;
        _ptReorderDown = screen;
        _wParamReorderDown = message->wParam;
        _uReorderDownMessage = type;
        // A left-held move entering a menu item may begin the pending gesture,
        // just like a down; let the menu process it normally until the threshold.
        return S_FALSE;
    }

    if (!_fReorderStarted && ReorderThresholdReached(_ptReorderDown, screen))
    {
        _hcurBeforeReorder = GetCursor();
        _fReorderStarted = TRUE;
        ReorderTooltips(toolbar, FALSE);
    }
    if (!_fReorderStarted)
        return S_FALSE;

    SetCursor(LoadCursor(NULL, IDC_SIZENS));
    if (hit && (UINT)target.idCommand != _idReorderSource)
        _MoveTaskButton(_idReorderSource, target.idCommand, TRUE);
    return ConsumeReorderMessage(message);
}
