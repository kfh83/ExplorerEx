#ifndef _TOPMATCH_H_
#define _TOPMATCH_H_

#define COMPILE_TOPMATCH

#ifdef COMPILE_TOPMATCH

#include "desktop2.h"
#include "HostUtil.h"
#include "COWSite.h"

#define WC_TOPMATCH  TEXT("Desktop top match")

class CTopMatch
	: public CAccessible
	, public IServiceProvider
	, public CObjectWithSite
	, public IOleCommandTarget
{
private:
	CTopMatch(HWND hwnd);
	~CTopMatch();

public:
	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// *** IAccessible ***
	STDMETHODIMP get_accState(VARIANT varChild, VARIANT *pvarState);

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);

	// *** IObjectWithSite ***
	STDMETHODIMP SetSite(IUnknown *punkSite);

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup,
		ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT *pCmdText);
	STDMETHODIMP Exec(const GUID *pguidCmdGroup,
		DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut);

private:
	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnEraseBackground(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnContextMenu(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);
	LRESULT _OnSMNFindItemWorker(PSMNDIALOGMESSAGE pdm);
	LRESULT _OnSMNGetMinSize(PSMNGETMINSIZE pgms);

	void _SetTileWidth(int cxTile);
	void _UpdateTopMatchSizeInOpenView();
	void _AddSearchExtension();
	void _AddSearchItem(LPARAM lParam, LPWSTR a3);
	LRESULT _ActivateItem(int iItem, BOOL b);
	int _GetLVCurSel();

private:
	HWND _hwnd;
	HWND _hwndList;
	HTHEME _hTheme;
	LONG _lRef;
	MARGINS _margins;
	COLORREF field_44;
	COLORREF _clrBk;
	int _cyIcon;
	WCHAR _sz[MAX_PATH];
	WCHAR _sz2[MAX_PATH];
	HIMAGELIST _himl;
	int field_464;

	friend BOOL TopMatch_RegisterClass();
};

#endif // COMPILE_TOPMATCH

#endif // _TOPMATCH_H_
