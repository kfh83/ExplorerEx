#ifndef _OPENHOST_H_
#define _OPENHOST_H_

#include "desktop2.h"
#include "COWSite.h"

#define WC_OPENPANEHOST TEXT("Desktop Open Pane Host")

class COpenViewHost
	: public IServiceProvider
	, public CObjectWithSite
	, public IOleCommandTarget
{
public:
	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;

	// *** IObjectWithSite ***
	STDMETHODIMP SetSite(IUnknown *punkSite) override;

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID* pguidCmdGroup, ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pCmdText) override;
	STDMETHODIMP Exec(
		const GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvarargIn, VARIANT* pvarargOut) override;

private:
	COpenViewHost(HWND hwnd);
	~COpenViewHost() override;

	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	LRESULT _OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

	void _Layout(int cx, int cy);
	void _ShowEnableWindow(HWND hwnd, BOOL bEnable);
	HRESULT _SetCurrentView(int a2, VARIANT* pvararg);
	HRESULT _HandleOpenBoxArrowKey(MSG* pmsg);
	HRESULT _HandleOpenBoxContextMenu(MSG* pmsg);

private:
	HWND _hwnd;
	LONG _lRef;
	int field_18;
	int field_1C;
	SMPANEDATA _aopa[5];

	friend BOOL OpenViewHost_RegisterClass();
};

#endif 