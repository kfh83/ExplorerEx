#include "pch.h"
#include "cocreateinstancehook.h"
#include "dbg.h"
#include "shundoc.h"

DWORD WINAPI BeepThread(LPVOID)
{
	Beep(1700, 200);
	return 0;
}


HRESULT CoCreateInstanceHook(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID* ppv)
{

	HRESULT hr = CoCreateInstance(rclsid, pUnkOuter, dwClsContext, riid, ppv);
	if (FAILED(hr))
	{
		if (rclsid == CLSID_StartMenuPin || rclsid == CLSID_WebCheck)
		{
			static const GUID *rgpiidTry[] = {
				// IPinnedList3 is used since Windows 10 build 17763.
				&IID_IPinnedList3,
				// IFlexiblePinnedList is used between 14393 and 17763.
				&IID_IFlexibleTaskbarPinnedList,
				// IPinnedList25 is used until 14393
				&IID_IPinnedList25
			};

			const GUID *piid = nullptr;
			for (const GUID *&piidCur : rgpiidTry)
			{
				piid = piidCur;
				if (SUCCEEDED(CoCreateInstance(rclsid, pUnkOuter, dwClsContext, *piid, ppv)))
				{
					break;
				}
			}

			*ppv = new CPinnedListWrapper((IUnknown*)*ppv, *piid);
			hr= S_OK;
		}
	}
#ifndef RELEASE
	if (FAILED(hr))
	{
		CreateThread(0, 0, BeepThread, 0, 0, 0);

		wchar_t* clsidstring = 0;
		if (FAILED(StringFromCLSID(rclsid, &clsidstring))) return hr;

		wchar_t* iidstring = 0;
		if (FAILED(StringFromCLSID(riid, &iidstring)))
		{
			if (clsidstring) CoTaskMemFree(clsidstring);
			return hr;
		}

		wprintf(L"COCREATEINSTANCE FAILED! clsid %s, riid %s\n", clsidstring, iidstring);
		if (clsidstring) CoTaskMemFree(clsidstring);
		if (iidstring) CoTaskMemFree(iidstring);
		dbg::printstacktrace();
	}
#endif
	return hr;
}

HRESULT SHCoCreateInstanceHook(_In_opt_ PCWSTR pszCLSID, _In_opt_ const CLSID* pclsid, _In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _Outptr_ void** ppv)
{
	HRESULT hr = SHCoCreateInstance(pszCLSID,pclsid,pUnkOuter,riid,ppv);
#ifndef RELEASE
	if (FAILED(hr))
	{
		wchar_t* iidstring = 0;
		if (FAILED(StringFromCLSID(riid, &iidstring)))
		{
			return hr;
		}

		wchar_t* clsidStr = const_cast<wchar_t*>(pszCLSID);

		if (!clsidStr)
		{
			if (FAILED(StringFromCLSID(*pclsid, &clsidStr)))
			{
				return hr;
			}
		}

		wprintf(L"COCREATEINSTANCE FAILED! clsid %s, riid %s\n", pszCLSID, iidstring);
		if (iidstring) CoTaskMemFree(iidstring);
		dbg::printstacktrace();
	}
#endif
	return hr;
}
