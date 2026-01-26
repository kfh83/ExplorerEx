#include "pch.h"
#include "uemapp.h"
#include "shundoc.h"

#undef  INTERFACE
#define INTERFACE   IUserAssist

typedef struct tagOBJECTINFO
{
	void* cf;
	CLSID const* pclsid;
	HRESULT(*pfnCreateInstance)(IUnknown* pUnkOuter, IUnknown** ppunk, const struct tagOBJECTINFO*);

	// for automatic registration, type library searching, etc
	int nObjectType;        // OI_ flag
	LPTSTR pszName;
	LPTSTR pszFriendlyName;
	IID const* piid;
	IID const* piidEvents;
	long lVersion;
	DWORD dwOleMiscFlags;
	int nidToolbarBitmap;
} OBJECTINFO;
typedef OBJECTINFO const* LPCOBJECTINFO;


#define UEIM_HIT        0x01
#define UEIM_FILETIME   0x02


MIDL_INTERFACE("49b36d57-5fd2-45a7-981b-06028d577a47")
IShellUserAssist : IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE FireEvent(const GUID * pguidGrp, UAEVENT eCmd, const WCHAR * pszPath, DWORD dwTimeElapsed) = 0;
	virtual HRESULT STDMETHODCALLTYPE QueryEntry(const GUID* pguidGrp, const WCHAR* pszPath, UEMINFO* pui) = 0;
	virtual HRESULT STDMETHODCALLTYPE SetEntry(const GUID* pguidGrp, const WCHAR* pszPath, UEMINFO* pui) = 0;
	virtual HRESULT STDMETHODCALLTYPE RenameEntry(const GUID* pguidGrp, const WCHAR* pszPathOld, const WCHAR* pszPathNew) = 0;
	virtual HRESULT STDMETHODCALLTYPE DeleteEntry(const GUID* pguidGrp, const WCHAR* pszPath) = 0;
	virtual HRESULT STDMETHODCALLTYPE Enable(BOOL bEnable) = 0;
	virtual HRESULT STDMETHODCALLTYPE RegisterNotify(UACallback pfnUACB, void* param, int) = 0;
};



IShellUserAssist* i10 = NULL;

BOOL UEMIsLoaded()
{
	printf("UEMIsLoaded\n");
	BOOL fRet;

	fRet = GetModuleHandle(TEXT("ole32.dll")) &&
		GetModuleHandle(TEXT("browseui.dll"));

	return fRet;
}

VOID EnsureUserAssist()
{
	if (!i10)
	{
		if (SUCCEEDED(CoCreateInstance(CLSID_UserAssist, NULL, CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_NO_CODE_DOWNLOAD, IID_IUserAssist10, (PVOID*)&i10)))
		{
			i10->Enable(TRUE);
		}
	}
}


HRESULT UEMFireEvent(const GUID* pguidGrp, int eCmd, DWORD dwFlags, WPARAM wParam, LPARAM lParam)
{
	printf("UEMFireEvent: %d\n", eCmd);
	HRESULT hr = E_FAIL;
	
	IShellFolder* ish = (IShellFolder*)wParam;
	if (!IsBadReadPtr(ish, sizeof(IShellFolder)))
	{
		LPITEMIDLIST pidl = (LPITEMIDLIST)lParam;
		STRRET psn;
		ish->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &psn);
		LPWSTR psz;
		StrRetToStrW(&psn, pidl, &psz);

		EnsureUserAssist();

		if (i10)
		{
			i10->FireEvent(&UAIID_SHORTCUTS, (UAEVENT)eCmd, psz, GetTickCount64());
		}
		CoTaskMemFree(psz);
		hr = S_OK;
	}

	return hr;
}

HRESULT UEMQueryEvent(const GUID* pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam, LPUEMINFO pui)
{
	printf("UEMQueryEvent: %d\n",eCmd);
	HRESULT hr = E_FAIL;

	IShellFolder* ish = (IShellFolder*)wParam;
	if (eCmd == UEME_RUNPIDL && !IsBadReadPtr(ish, sizeof(IShellFolder)))
	{
		LPITEMIDLIST pidl = (LPITEMIDLIST)lParam;
		STRRET psn;
		ish->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &psn);
		LPWSTR psz;
		StrRetToStrW(&psn, pidl, &psz);

		EnsureUserAssist();

		if (i10)
		{
			i10->QueryEntry(&UAIID_SHORTCUTS, psz, pui);
		}
		CoTaskMemFree(psz);
		hr = S_OK;
	}

	return hr;
}

// these are useless (doesnt get called)

HRESULT UEMSetEvent(const GUID* pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam, LPUEMINFO pui)
{
	printf("UEMSetEvent : %d\n",eCmd);
	// same as queryevent kinda
	HRESULT hr = E_FAIL;

	IShellFolder* ish = (IShellFolder*)wParam;
	if (!IsBadReadPtr(ish, sizeof(IShellFolder)))
	{
		LPITEMIDLIST pidl = (LPITEMIDLIST)lParam;
		STRRET psn;
		ish->GetDisplayNameOf(pidl, SHGDN_FORPARSING, &psn);
		LPWSTR psz;
		StrRetToStrW(&psn, pidl, &psz);

		EnsureUserAssist();

		i10->SetEntry(&UAIID_SHORTCUTS, psz, pui);
		CoTaskMemFree(psz);
		hr = S_OK;
	}

	return hr;
}

HRESULT UEMRegisterNotify(UACallback pfnUEMCB, void* param)
{
	printf("UEMRegisterNotify\n");
	EnsureUserAssist();
	HRESULT hr = i10->RegisterNotify(pfnUEMCB, param, 1);
	return hr;
}

void UEMEvalMsg(const GUID* pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam)
{
	printf("UEMEvalMsg\n");
	HRESULT hr;

	hr = UEMFireEvent(pguidGrp, eCmd, UEMF_XEVENT, wParam, lParam);
	return;
}

BOOL UEMGetInfo(const GUID* pguidGrp, int eCmd, WPARAM wParam, LPARAM lParam, LPUEMINFO pui)
{
	printf("UEMGetInfo\n");
	HRESULT hr;

	hr = UEMQueryEvent(pguidGrp, eCmd, wParam, lParam, pui);
	return SUCCEEDED(hr);
}
