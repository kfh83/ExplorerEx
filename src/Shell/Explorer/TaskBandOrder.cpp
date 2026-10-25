#include "pch.h"
#include "cabinet.h"
#include "TaskBand.h"

BOOL CTaskBand::_MoveTaskButton(UINT sourceCommand, UINT targetCommand, BOOL fromPopup)
{
    const HWND toolbar = _tb;
    if (!IsWindow(toolbar))
        return FALSE;
    // Taskix DLL 0x10001420: move members within a group, whole spans across groups.
    // Keep toolbar commands, item pointers, images and checked state intact.
    if (_dsaAII.GetItemCount() != 0 || sourceCommand == targetCommand)
        return FALSE;

    const int source = _tb.CommandToIndex(sourceCommand);
    const int target = _tb.CommandToIndex(targetCommand);
    const int count = _tb.GetButtonCount();
    if (source < 0 || target < 0 || source >= count || target >= count)
        return FALSE;

    PTASKITEM sourceItem = _GetItem(source);
    PTASKITEM targetItem = _GetItem(target);
    if (!sourceItem || !targetItem)
        return FALSE;
    // A command can be reused after deletion; never dereference the saved pointer.
    if (_hwndReorderToolbar &&
        (_idReorderSource != sourceCommand || _ptiReorderSource != sourceItem))
        return FALSE;
    if (sourceItem->hwnd == HWND_TOPMOST || targetItem->hwnd == HWND_TOPMOST ||
        (sourceItem->hwnd && !IsWindow(sourceItem->hwnd)) ||
        (targetItem->hwnd && !IsWindow(targetItem->hwnd)))
        return FALSE;

    const int sourceGroup = _GetGroupIndex(source);
    const int targetGroup = _GetGroupIndex(target);
    PTASKITEM sourceHeader = _GetItem(sourceGroup);
    PTASKITEM targetHeader = _GetItem(targetGroup);
    if (!sourceHeader || sourceHeader->hwnd || !targetHeader || targetHeader->hwnd)
        return FALSE;
    const int sourceEnd = sourceGroup + _GetGroupSize(sourceGroup);
    const int targetEnd = targetGroup + _GetGroupSize(targetGroup);
    if (sourceEnd <= sourceGroup || targetEnd <= targetGroup ||
        sourceEnd >= count || targetEnd >= count || source > sourceEnd || target > targetEnd)
        return FALSE;

    int first = source;
    int last = source;
    int destination = target;
    if (sourceGroup == targetGroup)
    {
        // Headers delimit the member list, including when grouping is disabled.
        if (source == sourceGroup || target == targetGroup)
            return FALSE;
    }
    else
    {
        if (fromPopup)
            return FALSE;
        if ((source == sourceGroup && _IsHidden(sourceGroup)) || _IsHidden(target))
            return FALSE;
        const BOOL movingRight = sourceGroup < targetGroup;
        if (target != targetGroup &&
            (!_IsHidden(targetGroup) || target != (movingRight ? targetEnd : targetGroup + 1)))
            return FALSE;
        first = sourceGroup;
        last = sourceEnd;
        destination = movingRight ? targetEnd : targetGroup;
    }

    // Refuse a group already containing a deleting or dead application member.
    for (int i = sourceGroup + 1; i <= sourceEnd; ++i)
    {
        PTASKITEM item = _GetItem(i);
        if (!item || !item->hwnd || item->hwnd == HWND_TOPMOST || !IsWindow(item->hwnd))
            return FALSE;
    }
    if (targetGroup != sourceGroup)
    {
        for (int i = targetGroup + 1; i <= targetEnd; ++i)
        {
            PTASKITEM item = _GetItem(i);
            if (!item || !item->hwnd || item->hwnd == HWND_TOPMOST || !IsWindow(item->hwnd))
                return FALSE;
        }
    }
    if (destination >= first && destination <= last)
        return FALSE;

    const BOOL movingRight = destination > last;
    const int length = last - first + 1;
    const int newFirst = movingRight ? destination - length + 1 : destination;
    const auto remap = [=](int index) -> int
    {
        if (index >= first && index <= last)
            return newFirst + index - first;
        if (movingRight && index > last && index <= destination)
            return index - length;
        if (!movingRight && index >= destination && index < first)
            return index + length;
        return index; // Includes all negative sentinels.
    };

    const int hot = _tb.GetHotItem();
    const BOOL denyHotChange = _fDenyHotItemChange;
    // The input owner suspends and restores the appropriate toolbar tooltip.
    _fDenyHotItemChange = TRUE;
    _tb.SetRedraw(FALSE);

    // Rotate a span one endpoint at a time; no item copies or heap scratch space.
    int completed = 0;
    for (; completed < length; ++completed)
    {
        const BOOL result = (BOOL)SendMessage(toolbar, TB_MOVEBUTTON,
            movingRight ? first : last, destination);
        if (!IsWindow(toolbar))
            return FALSE; // Destruction owns cancellation; do not access band state.
        if (!result)
            break;
    }
    const BOOL moved = completed == length;
    if (!moved)
    {
        // Invert the completed rotations in reverse order before exposing layout.
        while (completed > 0)
        {
            const BOOL restored = (BOOL)SendMessage(toolbar, TB_MOVEBUTTON,
                destination, movingRight ? first : last);
            if (!IsWindow(toolbar))
                return FALSE;
            ASSERT(restored);
            if (!restored)
            {
                // A broken live control cannot guarantee rollback. Stop the drag,
                // but never terminate the user's shell over a reorder failure.
                _fDenyHotItemChange = denyHotChange;
                _tb.SetRedraw(TRUE);
                _tb.Invalidate();
                _CancelTaskReorder();
                return FALSE;
            }
            --completed;
        }
    }
    else
    {
        _fReorderMoved = TRUE; // Popup refresh can synchronously reenter SMC_EXEC.
        _iIndexActiveAtLDown = remap(_iIndexActiveAtLDown);
        _iDropItem = remap(_iDropItem);
        _iIndexPopup = remap(_iIndexPopup);
        _iIndexLastPopup = remap(_iIndexLastPopup);
    }

    // TB_SETHOTITEM is programmatic (no arrow-key scrolling); allow its notification
    // only after every index and the entire group order are coherent again.
    _fDenyHotItemChange = FALSE;
    _tb.SetHotItem(moved ? remap(hot) : hot);
    _fDenyHotItemChange = denyHotChange;
    if (moved)
        _tb.AutoSize();
    _tb.SetRedraw(TRUE);
    _tb.Invalidate();
    if (moved && fromPopup)
        _RefreshReorderPopup();
    return moved;
}
