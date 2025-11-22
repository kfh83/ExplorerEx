#ifndef _OPENBOX_H_
#define _OPENBOX_H_

#include "desktop2.h"
#include "HostUtil.h"
#include "COWSite.h"

#define WC_OPENBOXHOST TEXT("Desktop OpenBox Host")

interface IShellSearchScope; // We dont need this, just a forward decl

typedef enum SHELLSEARCHNOTIFY
{
	SSC_KEYPRESS = 0x1,
	SSC_USERPAUSE = 0x2,
	SSC_LOSTFOCUS = 0x4,
	SSC_FORCE = 0x8,
	SSC_SEARCHCOMPLETE = 0x10,
	SSC_CLEARAUTONAVIGATE = 0x20,
	SSC_MRUINVOKED = 0x40,
	SSC_TAKEFOCUS = 0x80, // New in Windows 11
} SHELLSEARCHNOTIFY;

DEFINE_ENUM_FLAG_OPERATORS(SHELLSEARCHNOTIFY);

//MIDL_INTERFACE("E47C387F-113E-4820-A900-4C1EC5D85BC6") // Vista GUID
MIDL_INTERFACE("9A7A94F5-FBF1-49A8-B0D9-44667635FE97")
IShellSearchTarget : public IUnknown
{
	STDMETHOD(Search)(LPCWSTR, DWORD) PURE;
	STDMETHOD(OnSearchTextNotify)(LPCWSTR, LPCWSTR, SHELLSEARCHNOTIFY) PURE;
	STDMETHOD(GetSearchText)(LPWSTR, UINT) PURE;
	STDMETHOD(GetPromptText)(LPWSTR, UINT) PURE;
	STDMETHOD(GetMenu)(HMENU*) PURE;
	STDMETHOD(InitMenuPopup)(HMENU) PURE;
	STDMETHOD(OnMenuCommand)(DWORD) PURE;
	STDMETHOD(Enter)(IShellSearchScope*) PURE;
	STDMETHOD(Exit)() PURE;
};

static const IID IID_IShellSearchTarget = __uuidof(IShellSearchTarget);

enum SSCSTATEFLAGS
{
	SSCSTATE_DEFAULT = 0x1,
	SSCSTATE_NOSETFOCUS = 0x2,
	SSCSTATE_DRAWCUETEXTFOCUS = 0x4,
	SSCSTATE_APPENDTEXT = 0x8,
	SSCSTATE_COMMIT = 0x10,
	SSCSTATE_NODROPDOWN = 0x20,
	SSCSTATE_DISABLED = 0x40,
	SSCSTATE_NOSUGGESTIONS = 0x80,
	SSCSTATE_STRUCTUREDQUERYPARSING = 0x100,
};

DEFINE_ENUM_FLAG_OPERATORS(SSCSTATEFLAGS);

enum SEARCH_BOX_SUGGEST_POPUP_SETTING
{
	SBSPS_DEFAULT = 0x0,
	SBSPS_NO_POPUP = 0x1,
	SBSPS_CUSTOM_MRU = 0x2,
	SBSPS_POPUP_ON_TEXT = 0x4,
	SBSPS_ONLY_MRU = 0x8,
	SBSPS_DRAWCUETEXTFOCUS = 0x10,
};

DEFINE_ENUM_FLAG_OPERATORS(SEARCH_BOX_SUGGEST_POPUP_SETTING);

enum SSC_WIDTH_FLAGS
{
	SSCWIDTH_DEFAULT = 0x0,
	SSCWIDTH_SETBYUSER = 0x1,
};

DEFINE_ENUM_FLAG_OPERATORS(SSC_WIDTH_FLAGS);

enum SSCTEXTFLAGS
{
	SSCTEXT_DEFAULT = 0x0,
	SSCTEXT_TAKEFOCUS = 0x1,
	SSCTEXT_FORCE = 0x2,
	SSCTEXT_CLEARAUTONAVIGATE = 0x4,
};

DEFINE_ENUM_FLAG_OPERATORS(SSCTEXTFLAGS);

MIDL_INTERFACE("111f7c32-0546-4227-8b7f-c53a0b114a0f")
IShellSearchControl : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE Initialize(HWND, const RECT *) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetFlags(SSCSTATEFLAGS) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetPopupFlags(SEARCH_BOX_SUGGEST_POPUP_SETTING) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetFlags(SSCSTATEFLAGS *) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetMRULocation(REFGUID) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetCueAndTooltipText(const WCHAR *, const WCHAR *) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetLocation(IShellItem *) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetDesiredSize(SIZE *) = 0;
	virtual HRESULT STDMETHODCALLTYPE UpdateWidth(int, SSC_WIDTH_FLAGS) = 0;
	virtual HRESULT STDMETHODCALLTYPE GetText(WCHAR *, UINT) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetText(const WCHAR *, SSCTEXTFLAGS) = 0;
	virtual HRESULT STDMETHODCALLTYPE DoesEditBoxHaveFocus(int *) = 0;
	virtual HRESULT STDMETHODCALLTYPE HasSuggestions() = 0;
	virtual HRESULT STDMETHODCALLTYPE HideSuggestions() = 0;
	virtual HRESULT STDMETHODCALLTYPE KillSelection() = 0;
};

class COpenBoxHost
	: public CAccessible
	, public IServiceProvider
	, public CObjectWithSite
	, public IOleCommandTarget
	, public IShellSearchTarget
	, public IInputObjectSite
{
public:
	// *** IUnknown ***
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
	STDMETHODIMP_(ULONG) AddRef() override;
	STDMETHODIMP_(ULONG) Release() override;

	// *** IAccessible ***
	STDMETHODIMP get_accRole(VARIANT varChild, VARIANT* pvarRole) override;
	STDMETHODIMP get_accState(VARIANT varChild, VARIANT* pvarState) override;
	STDMETHODIMP get_accDefaultAction(VARIANT varChild, BSTR* pszDefAction) override;
	STDMETHODIMP accDoDefaultAction(VARIANT varChild) override;

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void** ppvObject) override;

	// *** IObjectWithSite ***
	STDMETHODIMP SetSite(IUnknown* punkSite) override;

	// *** IOleCommandTarget ***
	STDMETHODIMP QueryStatus(const GUID* pguidCmdGroup,
		ULONG cCmds, OLECMD rgCmds[], OLECMDTEXT* pCmdText) override;
	STDMETHODIMP Exec(const GUID* pguidCmdGroup, 
		DWORD nCmdID, DWORD nCmdexecopt, VARIANT* pvarargIn, VARIANT* pvarargOut) override;

	// *** IShellSearchTarget ***
	STDMETHODIMP Search(LPCWSTR a2, DWORD a3) override;
	STDMETHODIMP OnSearchTextNotify(LPCWSTR a2, LPCWSTR a3, SHELLSEARCHNOTIFY a4) override;
	STDMETHODIMP GetSearchText(LPWSTR pszText, UINT cchText) override;
	STDMETHODIMP GetPromptText(LPWSTR pszText, UINT cchText) override;
	STDMETHODIMP GetMenu(HMENU* phMenu) override;
	STDMETHODIMP InitMenuPopup(HMENU hhenu) override;
	STDMETHODIMP OnMenuCommand(DWORD dwCmd) override;
	STDMETHODIMP Enter(IShellSearchScope* pScope) override;
	STDMETHODIMP Exit() override;

	// *** IInputObjectSite ***
	STDMETHODIMP OnFocusChangeIS(IUnknown* punk, BOOL fSetFocus) override;

private:
	COpenBoxHost(HWND hwnd);
	~COpenBoxHost();

	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnEraseBkgnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);

	BOOL _Mark(PSMNDIALOGMESSAGE pdm, UINT a2);
	HRESULT _UpdateSearch();

	friend BOOL OpenBoxHost_RegisterClass();

private:
	IShellSearchControl* _pssc;
	HWND _hwndEdit;
	int field_34;
	HWND _hwnd;
	LONG _lRef;
	HTHEME _hTheme;
	COLORREF _clrBk;
	int field_48;
	int field_4C;
	LPWSTR _pszSearchQuery;
};

#endif
