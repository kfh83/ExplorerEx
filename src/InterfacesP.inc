#pragma once
//
// COM Interfaces
// 

MIDL_INTERFACE("8bfcb27d-cf1a-11d0-b93d-00a0c90312e1")
IOrderList : public IUnknown
{
    // *** IOrderList Methods ***
    STDMETHOD(GetOrderList)(THIS_ HDPA * phdpa) PURE;
    STDMETHOD(SetOrderList)(THIS_ HDPA hdpa, IShellFolder * psf) PURE;
    STDMETHOD(FreeOrderList)(THIS_ HDPA hdpa) PURE;
    STDMETHOD(SortOrderList)(THIS_ HDPA hdpa, DWORD dw) PURE;
    STDMETHOD(AllocOrderItem)(THIS_ PORDERITEM * ppoi, LPCITEMIDLIST pidl) PURE;
    STDMETHOD(FreeOrderItem)(THIS_ PORDERITEM poi) PURE;
};

MIDL_INTERFACE("AFA7AD8A-51A7-11d2-83A1-00C04FD918D0")
IOrderList2: IOrderList
{
    // *** IOrderList2
    STDMETHOD(LoadFromStream)(THIS_ IStream * pstm, HDPA * phdpa, IShellFolder * psf) PURE;
    STDMETHOD(SaveToStream)(THIS_ IStream * pstm, HDPA hdpa) PURE;
};

//MIDL_INTERFACE("AF4F6511-F982-11d0-8595-00AA004CD6D8")
MIDL_INTERFACE("7897eca6-1b1b-452a-8581-bb9482aea7cc")
IRegTreeOptions : IUnknown
{
    // *** IUnknown methods ***
    STDMETHOD(QueryInterface) (THIS_ REFIID riid, void** ppv) PURE;
    STDMETHOD_(ULONG, AddRef) (THIS)  PURE;
    STDMETHOD_(ULONG, Release) (THIS) PURE;

    // *** IRegTreeOptions specific methods ***
    /*STDMETHOD(InitTree)(THIS_ HWND hwndTree, HKEY hkeyRoot, LPCSTR pszRegKey, LPCSTR pszParam) PURE;
    STDMETHOD(WalkTree)(THIS_ WALK_TREE_CMD cmd) PURE;
    STDMETHOD(ToggleItem)(THIS_ HTREEITEM hti) PURE;
    STDMETHOD(ShowHelp)(THIS_ HTREEITEM hti, DWORD dwFlags) PURE;*/

    STDMETHOD(InitTree)(HWND, HKEY, LPCSTR, LPCSTR) PURE;
    STDMETHOD(WalkTree)(unsigned int) PURE;
    STDMETHOD(ToggleItem)(_TREEITEM*) PURE;
    STDMETHOD(ShowHelp)(_TREEITEM*, ULONG) PURE;
    STDMETHOD(SetEventListener)(IUnknown*) PURE;
};

MIDL_INTERFACE("D12F26B1-D90A-11d0-830D-00AA005B4383")
IRestrict : public IUnknown
{
    // *** IRestrict Methods ***
    STDMETHOD(IsRestricted) (THIS_ const GUID* pguidID, DWORD dwRestrictAction, VARIANT* pvarArgs, OUT DWORD* pdwRestrictionResult) PURE;
};

MIDL_INTERFACE("EA5F2D61-E008-11CF-99CB-00C04FD64497")
IWinEventHandler : IUnknown
{
    // *** IWinEventHandler Methods ***
    STDMETHOD(OnWinEvent) (THIS_ HWND hwnd, UINT dwMsg, WPARAM wParam, LPARAM lParam, LRESULT* plres) PURE;
    STDMETHOD(IsWindowOwner) (THIS_ HWND hwnd) PURE;
};

MIDL_INTERFACE("5EA35BC9-19B1-11D1-9828-00C04FD91972")
IShellHotKey : IUnknown
{
    // *** IShellHotKey methods ***
    STDMETHOD(RegisterHotKey)(THIS_ IShellFolder* psf, LPCITEMIDLIST pidlParent, LPCITEMIDLIST pidl) PURE;
};

MIDL_INTERFACE("4622AD10-FF23-11D0-8D34-00A0C90F2719")
ITrayPriv : IOleWindow
{
    // *** ITrayPriv methods ***
    STDMETHOD(ExecItem)(THIS_ IShellFolder* psf, LPCITEMIDLIST pidl) PURE;
    STDMETHOD(GetFindCM)(THIS_ HMENU hmenu, UINT idFirst, UINT idLast, IContextMenu** ppcmFind) PURE;
    STDMETHOD(GetStaticStartMenu)(THIS_ HMENU* phmenu) PURE;
};

MIDL_INTERFACE("9E83C057-6823-4F1F-BfA3-7461D40A8173")
ITrayPriv2 : ITrayPriv
{
    // *** ITrayPriv2 methods ***
    STDMETHOD(ModifySMInfo)(THIS_ IN LPSMDATA psmd, IN OUT SMINFO* psminfo) PURE;
};

DECLARE_INTERFACE_(IDeskTray, IUnknown)
{
    // *** IDeskTray methods ***
    STDMETHOD_(UINT, AppBarGetState)(THIS) PURE;
    STDMETHOD(GetTrayWindow)(THIS_ HWND* phwndTray) PURE;
    STDMETHOD(SetDesktopWindow)(THIS_ HWND hwndDesktop) PURE;

    // WARNING!  BEFORE CALLING THE SetVar METHOD YOU MUST DETECT
    // THE EXPLORER VERSION BECAUSE IE 4.00 WILL CRASH IF YOU TRY
    // TO CALL IT
    
    STDMETHOD(SetVar)(THIS_ int var, DWORD value) PURE;
};

MIDL_INTERFACE("6F51C646-0EFE-4370-882A-C1F61CB27C3B")
IShellMenu2 : IShellMenu
{
    // Retrieves an interface on a submenu.
    STDMETHOD(GetSubMenu)(THIS_ UINT idCmd, REFIID riid, void** ppvObj) PURE;
    STDMETHOD(SetToolbar)(THIS_ HWND hwnd, DWORD dwFlags) PURE;
    STDMETHOD(SetMinWidth)(THIS_ int cxMenu) PURE;
    STDMETHOD(SetNoBorder)(THIS_ BOOL fNoBorder) PURE;
    STDMETHOD(SetTheme)(THIS_ LPCWSTR pszTheme) PURE;
};

MIDL_INTERFACE("BBD20037-BC0E-42F1-913F-E2936BB0EA0C")
IStartMenuPin : IUnknown
{
    /*STDMETHOD(EnumObjects)(THIS_ IEnumIDList * *ppenumIDList) PURE;
    //
    //  Pin:        pidlFrom = NULL, pidlTo = pidl
    //  Unpin:      pidlFrom = pidl, pidlTo = NULL
    //  Update:     pidlFrom = old,  pidlTo = new
    //  Move:       pidlFrom = pidl, pidlTo = SMPINPOS(iPos)
    STDMETHOD(Modify)(THIS_ LPCITEMIDLIST pidlFrom, LPCITEMIDLIST pidlTo) PURE;
    STDMETHOD(GetChangeCount)(THIS_ ULONG* pulOut) PURE;

    //
    //  pdto = data object to test
    //  dwFlags is an SMPINNABLE_* flag
    //  *ppidl receives pidl being pinned
    //
    STDMETHOD(IsPinnable)(THIS_ IDataObject* pdto, DWORD dwFlags, LPITEMIDLIST* ppidl) PURE; // S_FALSE if not

    //
    //  Find the pidl on the pin list and resolve the shortcut that
    //  tracks it.
    //
    //  Returns S_OK if the pidl changed and was resolved.
    //  Returns S_FALSE if the pidl did not change.
    //  Returns an error if the Resolve failed.
    //
    STDMETHOD(Resolve)(THIS_ HWND hwnd, DWORD dwFlags, LPCITEMIDLIST pidl, LPITEMIDLIST* ppidlResolved) PURE;*/

	STDMETHOD(EnumObjects)(IEnumFullIDList**) PURE;
	STDMETHOD(Modify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetChangeCount)(ULONG*) PURE;
	STDMETHOD(GetPinnableInfo)(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*) PURE;
	STDMETHOD(IsPinnable)(IDataObject*, int) PURE;
	STDMETHOD(Resolve)(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(IsPinned)(PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetPinnedItem)(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(GetAppIDForPinnedItem)(PCIDLIST_ABSOLUTE, PWSTR*) PURE;
	STDMETHOD(ItemChangeNotify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(UpdateForRemovedItemsAsNecessary)(VOID) PURE;
};


MIDL_INTERFACE("446bc432-57e9-4b72-8ef1-af27113dcf9c")
IPinnedList25: public IUnknown
{
public:
	STDMETHOD(EnumObjects)(IEnumFullIDList**) PURE;
	STDMETHOD(GetPinnableInfo)(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*) PURE;
	STDMETHOD(IsPinnable)(IDataObject*, int) PURE;
	STDMETHOD(Resolve)(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(Modify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetChangeCount)(ULONG*) PURE;
	STDMETHOD(IsPinned)(PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetPinnedItem)(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(GetAppIDForPinnedItem)(PCIDLIST_ABSOLUTE, PWSTR*) PURE;
	STDMETHOD(ItemChangeNotify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(UpdateForRemovedItemsAsNecessary)(VOID) PURE;
};

MIDL_INTERFACE("60274FA2-611F-4B8A-A293-F27BF103D148")
IFlexibleTaskbarPinnedList : public IUnknown
{
public:
	STDMETHOD(EnumObjects)(IEnumFullIDList**) PURE;
	STDMETHOD(GetPinnableInfo)(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*) PURE;
	STDMETHOD(IsPinnable)(IDataObject*, int) PURE;
	STDMETHOD(Resolve)(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(Modify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetChangeCount)(ULONG*) PURE;
	STDMETHOD(IsPinned)(PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetPinnedItem)(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(GetAppIDForPinnedItem)(PCIDLIST_ABSOLUTE, PWSTR*) PURE;
	STDMETHOD(ItemChangeNotify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(UpdateForRemovedItemsAsNecessary)(VOID) PURE;
	STDMETHOD(GetPinnedItemForAppID)(PWSTR, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(ApplyInPlaceTaskbarLayout)(int, int) PURE;
	STDMETHOD(ApplyReorderTaskbarLayout)(int, int) PURE;
};


MIDL_INTERFACE("0DD79AE2-D156-45D4-9EEB-3B549769E940")
IPinnedList3 : public IUnknown
{
public:
	STDMETHOD(EnumObjects)(IEnumFullIDList**) PURE;
	STDMETHOD(GetPinnableInfo)(IDataObject*, int, IShellItem2**, IShellItem**, PWSTR*, INT*) PURE;
	STDMETHOD(IsPinnable)(IDataObject*, int) PURE;
	STDMETHOD(Resolve)(HWND, ULONG, PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(Unadvise)(ULONG) PURE;
	STDMETHOD(GetChangeCount)(ULONG*) PURE;
	STDMETHOD(IsPinned)(PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(GetPinnedItem)(PCIDLIST_ABSOLUTE, PIDLIST_ABSOLUTE*) PURE;
	STDMETHOD(GetAppIDForPinnedItem)(PCIDLIST_ABSOLUTE, PWSTR*) PURE;
	STDMETHOD(ItemChangeNotify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(UpdateForRemovedItemsAsNecessary)(VOID) PURE;
	STDMETHOD(PinShellLink)(PWSTR, IShellLinkW*) PURE;
	STDMETHOD(GetPinnedItemForAppID)(PWSTR, PCIDLIST_ABSOLUTE) PURE;
	STDMETHOD(Modify)(PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE, int) PURE;
};


MIDL_INTERFACE("5836FB00-8187-11CF-A12B-00AA004AE837")
IShellService : IUnknown
{
    // *** IShellService specific methods ***
    STDMETHOD(SetOwner)(THIS_ IUnknown* punkOwner) PURE;
};

MIDL_INTERFACE("FADB55B4-D382-4FC4-81D7-ABB325C7F12A")
IFadeTask : IUnknown
{
    STDMETHOD(FadeRect)(THIS_ LPCRECT prc) PURE;
};

MIDL_INTERFACE("E9EAD8E6-2A25-410E-9B58-A9fBEF1DD1A2")
IUserEventTimerCallback: IUnknown
{
    STDMETHOD(UserEventTimerProc)(THIS_ ULONG uUserEventTimerID, UINT uTimerElapse) PURE;
};

MIDL_INTERFACE("0F504B94-6E42-42E6-99E0-E20FAFE52AB4")
IUserEventTimer: IUnknown
{
    STDMETHOD(SetUserEventTimer)(THIS_ HWND hWnd, UINT uCallbackMessage, UINT uTimerElapse, IUserEventTimerCallback * pUserEventTimerCallback, ULONG * puUserEventTimerID) PURE;
    STDMETHOD(KillUserEventTimer)(THIS_ HWND hWnd, ULONG uUserEventTimerID) PURE;

    STDMETHOD(GetUserEventTimerElapsed)(THIS_ HWND hWnd, ULONG uUserEventTimerID, UINT* puTimerElapsed) PURE;
    STDMETHOD(InitTimerTickInterval)(THIS_ UINT uTimerTickIntervalMs) PURE;
};

MIDL_INTERFACE("D782CCBA-AFB0-43F1-94DB-FDA3779EACCB")
INotificationCB : IUnknown
{
    STDMETHOD(Notify)(THIS_ ULONG, NOTIFYITEM*) PURE;
};

//MIDL_INTERFACE("fb852b2c-6bad-4605-9551-f15f87830935")
//MIDL_INTERFACE("D133CE13-3537-48BA-93A7-AFCD5D2053B4")
//ITrayNotify : IUnknown
//{
//    STDMETHOD(RegisterCallback)(INotificationCB* pNotifyCB) PURE;
//    STDMETHOD(SetPreference)(LPNOTIFYITEM pNotifyItem) PURE;
//    STDMETHOD(EnableAutoTray)(BOOL bTraySetting) PURE;
//};

#pragma optimize( "", off )
//this is the Windows 8+ variant of ITrayNotify, probably not needed for now but might need it for later
MIDL_INTERFACE("D133CE13-3537-48BA-93A7-AFCD5D2053B4")
ITrayNotify : IUnknown
{
	STDMETHOD(RegisterCallback)(INotificationCB* callback, DWORD*) PURE;
	STDMETHOD(UnregisterCallback)(ULONG*) PURE;
	STDMETHOD(SetPreference)(NOTIFYITEM) PURE;
	STDMETHOD(EnableAutoTray)(BOOL) PURE;
	STDMETHOD(DoAction)(BOOL) PURE;
	STDMETHOD(SetWindowingEnvironmentConfig)(IUnknown*) PURE;
};
#pragma optimize( "", on )

MIDL_INTERFACE("2F711B17-773C-41D4-93FA-7F23EDCECB66")
IAugmentedShellFolder : IShellFolder
{
    STDMETHOD(AddNameSpace)(THIS_ LPCGUID, IShellFolder*, LPCITEMIDLIST, ULONG, ULONG) PURE;
    STDMETHOD(GetNameSpaceID)(THIS_ LPCITEMIDLIST, LPCGUID) PURE;
    STDMETHOD(QueryNameSpace)(THIS_ ULONG, LPCGUID, IShellFolder**) PURE;
    STDMETHOD(EnumNameSpace)(THIS_ ULONG, PULONG) PURE;
    STDMETHOD(UnWrapIDList)(THIS_ LPCITEMIDLIST, LONG, IShellFolder**, LPITEMIDLIST*, LPITEMIDLIST*, LONG*) PURE;
};

MIDL_INTERFACE("8DB3B3F4-6CFE-11d1-8AE9-00C04FD918D0")
IAugmentedShellFolder2 : IAugmentedShellFolder
{
     STDMETHOD(UnWrapIDList)(THIS_ LPCITEMIDLIST, LONG, IShellFolder**, LPITEMIDLIST*, LPITEMIDLIST*, LONG*) PURE;
};

MIDL_INTERFACE("D8F6AD5B-B44F-4BCC-88FD-EB3473DB7502") //Revised version with QueryGuid
IAssociationElement : IUnknown
{
    STDMETHOD(QueryString)(ASSOCQUERY query, PCWSTR lpValueName, PWSTR* ppszOut) PURE;
    STDMETHOD(QueryDword)(ASSOCQUERY query, PCWSTR lpValueName, DWORD* pdwOut) PURE;
    STDMETHOD(QueryGuid)(ASSOCQUERY query, PCWSTR lpValueName, GUID *pGuidOut) PURE;
    STDMETHOD(QueryExists)(ASSOCQUERY query, PCWSTR lpValueName) PURE;
    STDMETHOD(QueryDirect)(ASSOCQUERY query, PCWSTR lpValueName, FLAGGED_BYTE_BLOB** ) PURE;
    STDMETHOD(QueryObject)(ASSOCQUERY query, PCWSTR lpValueName, REFIID riid, void** ppvObject) PURE;
};