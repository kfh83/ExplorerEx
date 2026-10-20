#pragma once
#include "pch.h"

HRESULT CoCreateInstanceHook(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID* ppv);
HRESULT SHCoCreateInstanceHook(_In_opt_ PCWSTR pszCLSID, _In_opt_ const CLSID* pclsid, _In_opt_ IUnknown* pUnkOuter, _In_ REFIID riid, _Outptr_ void** ppv);

// Takes ownership of *ppcm on success; all non-Start verbs remain native.
HRESULT WrapStartMenuPinContextMenu(IContextMenu** ppcm, IDataObject* data);