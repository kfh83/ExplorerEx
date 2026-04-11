//+-------------------------------------------------------------------------
//
//	ExplorerEx - Windows XP Explorer
//	Copyright (C) Microsoft
// 
//	File:			path.h
// 
//	History:		Jan-23-2025		kfh83		Created
//
//+-------------------------------------------------------------------------

#pragma once

STDAPI_(LONG) PathProcessCommand(LPCTSTR lpSrc, LPTSTR lpDest, int iDestMax, DWORD dwFlags);
HRESULT LoadFromFileW(REFCLSID clsid, LPCWSTR pszFile, REFIID riid, void** ppv);

// LOL!!!! LOL!!!!! LOL!!!!!!!!!
