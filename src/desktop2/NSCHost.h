#ifndef _NSCHOST_H_
#define _NSCHOST_H_

#include "desktop2.h"

#include "COWSite.h"

#define WC_NSCHOST  TEXT("Desktop NSCHost")

class CNSCHost
	: public INameSpaceTreeControlEvents
	, public INameSpaceTreeAccessible
	, public INameSpaceTreeControlCustomDraw
	, public INameSpaceTreeControlDropHandler
	, public IServiceProvider
	, public IOleCommandTarget
	, public CObjectWithSite
{
public:
	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// *** INameSpaceTreeControlEvents ***
	STDMETHODIMP OnItemClick(IShellItem *psi, NSTCEHITTEST nstceHitTest, NSTCECLICKTYPE nstceClickType);
	STDMETHODIMP OnPropertyItemCommit(IShellItem *psi);
	STDMETHODIMP OnItemStateChanging(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState);
	STDMETHODIMP OnItemStateChanged(IShellItem *psi, NSTCITEMSTATE nstcisMask, NSTCITEMSTATE nstcisState);
	STDMETHODIMP OnSelectionChanged(IShellItemArray *psiaSelection);
	STDMETHODIMP OnKeyboardInput(UINT uMsg, WPARAM wParam, LPARAM lParam);
	STDMETHODIMP OnBeforeExpand(IShellItem *psi);
	STDMETHODIMP OnAfterExpand(IShellItem *psi);
	STDMETHODIMP OnBeginLabelEdit(IShellItem *psi);
	STDMETHODIMP OnEndLabelEdit(IShellItem *psi);
	STDMETHODIMP OnGetToolTip(IShellItem *psi, LPWSTR pszTip, int cchTip);
	STDMETHODIMP OnBeforeItemDelete(IShellItem *psi);
	STDMETHODIMP OnItemAdded(IShellItem *psi, BOOL fIsRoot);
	STDMETHODIMP OnItemDeleted(IShellItem *psi, BOOL fIsRoot);
	STDMETHODIMP OnBeforeContextMenu(IShellItem *psi, REFIID riid, void **ppv);
	STDMETHODIMP OnAfterContextMenu(IShellItem *psi, IContextMenu *pcmIn, REFIID riid, void **ppv);
	STDMETHODIMP OnBeforeStateImageChange(IShellItem *psi);
	STDMETHODIMP OnGetDefaultIconIndex(IShellItem *psi, int *piDefaultIcon, int *piOpenIcon);

	// *** INameSpaceTreeAccessible ***
	STDMETHODIMP OnGetDefaultAccessibilityAction(IShellItem *psi, BSTR *pbstrDefaultAction);
	STDMETHODIMP OnDoDefaultAccessibilityAction(IShellItem *psi);
	STDMETHODIMP OnGetAccessibilityRole(IShellItem *psi, VARIANT *pvarRole);

	// *** INameSpaceTreeControlCustomDraw ***
	STDMETHODIMP PrePaint(HDC hdc, RECT *prc, LRESULT *plres);
	STDMETHODIMP PostPaint(HDC hdc, RECT *prc);
	STDMETHODIMP ItemPrePaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem, COLORREF *pclrText, COLORREF *pclrTextBk, LRESULT *plres);
	STDMETHODIMP ItemPostPaint(HDC hdc, RECT *prc, NSTCCUSTOMDRAW *pnstccdItem);

	// *** INameSpaceTreeControlDropHandler ***
	STDMETHODIMP OnDragEnter(IShellItem *psiOver, IShellItemArray *psiaData, BOOL fOutsideSource, DWORD grfKeyState, DWORD *pdwEffect);
	STDMETHODIMP OnDragOver(IShellItem *psiOver, IShellItemArray *psiaData, DWORD grfKeyState, DWORD *pdwEffect);
	STDMETHODIMP OnDragPosition(IShellItem *psiOver, IShellItemArray *psiaData, int iNewPosition, int iOldPosition);
	STDMETHODIMP OnDrop(IShellItem *psiOver, IShellItemArray *psiaData, int iPosition, DWORD grfKeyState, DWORD *pdwEffect);
	STDMETHODIMP OnDropPosition(IShellItem *psiOver, IShellItemArray *psiaData, int iNewPosition, int iOldPosition);
	STDMETHODIMP OnDragLeave(IShellItem *psiOver);

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pcmdtext);
	STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG *pvarargIn, VARIANTARG *pvarargOut);

	CNSCHost();

private:
	~CNSCHost();

	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnEraseBkGnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm);
	LRESULT _OnSMNGetMinSize(PSMNGETMINSIZE pgms);
	LRESULT _OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnWinEvent(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	BOOL _AreChangesRestricted();
	HRESULT _InitializeNSC(HWND hwnd);
	HRESULT _CollapseAll();
	HRESULT _GetSelectedItem(IShellItem **ppsi);
	void _NotifyCaptureInput(BOOL fBlock);
	void _InstrumentLaunchData(IShellItem *psi);
	HRESULT _Invoke(IShellItem *psi, int a3);
	HRESULT _IsItemMSIAds(IShellItem *psi);
	HRESULT _IsNewItem(IShellItem *psi);

private:
	LONG _cRef;
	LPITEMIDLIST _pidl;
	IShellFolder *_psf;
	INameSpaceTreeControl2 *_pns;
	IShellItemFilter *_psif;
	IWinEventHandler *_pweh;
	HWND _hwnd;
	HTHEME _hTheme;
	MARGINS _margins;
	DWORD _dwCookie;
	COLORREF _clrTextBk;
	COLORREF _clrText;
	int field_5C;
	int field_60;
	int field_64;
	int field_68;

	friend BOOL NSCHost_RegisterClass();
};

#endif
