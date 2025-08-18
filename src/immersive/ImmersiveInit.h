//+-------------------------------------------------------------------------
//
//	ExplorerEx - Windows XP Explorer
//	Copyright (C) Microsoft
// 
//	File:       ImmersiveInit.h
// 
//	Description:	Responsible for initialising the Immersive shell.
// 
//	History:    May-19-25   kawapure  Created
//
//+-------------------------------------------------------------------------

#include "pch.h"
#ifdef EXEX_DLL

#include "cabinet.h"
#include "CWndProc.h"
#include "ImmersiveDefs.h"

class CTaskmanWindow : public CImpWndProc
{
	LRESULT v_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) override;

	UINT _wmShellHook = 0;

public:
	static HRESULT RegisterWindowClass();
	CTaskmanWindow();

	inline ULONG AddRef() override { return 1; }
	inline ULONG Release() override { return 0; }
};

// Ittr: Needs to exist for legacy or 8.1 codepath
class CImmersiveBehaviorWrapper : public IImmersiveBehavior
{
public:
	CImmersiveBehaviorWrapper(IImmersiveBehavior* behavior);
	~CImmersiveBehaviorWrapper();

	//IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IImmersiveBehavior
	STDMETHODIMP OnImmersiveThreadStart(void);
	STDMETHODIMP OnImmersiveThreadStop(void);
	STDMETHODIMP GetMaximumComponentCount(unsigned int* count);
	STDMETHODIMP CreateComponent(unsigned int number, IUnknown** component);
	STDMETHODIMP ShouldCreateComponent(unsigned int number, int* allowed);
private:
	IImmersiveBehavior* m_behavior;
	long m_cRef;
};

HRESULT InitializeImmersiveShell();

#endif