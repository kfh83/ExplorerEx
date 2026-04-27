#pragma once

#include <objidl.h>

STDAPI BindCtx_AddObjectParam(IBindCtx* pbc, const WCHAR* pszKey, IUnknown* punkValue);
STDAPI BindCtx_RegisterObjectParam(IBindCtx* pbc, const WCHAR* pszKey, IUnknown* punk, IBindCtx** ppbc);
STDAPI BindCtx_RegisterUIWindow(IBindCtx* pbc, HWND hwnd, IBindCtx** ppbc);
STDAPI BindCtx_SetMode(IBindCtx* pbcIn, DWORD grfMode, IBindCtx** ppbcOut);
STDAPI BindCtx_CreateWithMode(DWORD grfMode, IBindCtx** ppbcOut);
