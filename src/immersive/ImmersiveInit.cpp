//+-------------------------------------------------------------------------
//
//	ExplorerEx - Windows XP Explorer
//	Copyright (C) Microsoft
// 
//	File:       ImmersiveInit.cpp
// 
//	Description:	Responsible for initialising the Immersive shell.
// 
//	History:    May-19-25   kawapure  Created
//
//+-------------------------------------------------------------------------

#include "pch.h"
#ifdef EXEX_DLL

#include "ImmersiveInit.h"
#include <Tray.h>

// TODO: Should this be a global variable? Look into restructuring.
IImmersiveShellHookService *g_pShellHookService = nullptr;

// static
HRESULT CTaskmanWindow::RegisterWindowClass()
{
	WNDCLASSEX wc = { 0 };

	wc.cbClsExtra = 0;
	wc.cbWndExtra = sizeof(void *); // We store a single pointer here.
	wc.cbSize = sizeof(wc);
	wc.style = 8; // TODO: Document.
	wc.hIcon = NULL;
	wc.hIconSm = NULL;
	wc.hInstance = hinstCabinet;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)2; // TODO: Document.
	wc.lpfnWndProc = s_WndProc;
	wc.lpszClassName = TEXT("TaskmanWndClass");
	wc.lpszMenuName = nullptr;

	if (!RegisterClassEx(&wc))
	{
		DWORD dwLastError = GetLastError();
		wprintf(L"Failed to register Taskman window class (error code %d).\n", dwLastError);
		return HRESULT_FROM_WIN32(dwLastError);
	}

	return S_OK;
}

bool g_fTaskmanRegistered = false;

CTaskmanWindow::CTaskmanWindow()
{
	printf("Initialising CTaskmanWindow...\n");

	if (!g_fTaskmanRegistered)
	{
		printf("Taskman window class not registered. Registering now...\n");
		RegisterWindowClass();
		g_fTaskmanRegistered = true;
	}

	_hwnd = CreateWindowEx(
		NULL,
		TEXT("TaskmanWndClass"),
		NULL,
		WS_POPUP | WS_CLIPCHILDREN,
		0, 0, 0, 0,
		NULL,
		NULL,
		hinstCabinet,
		this
	);

	if (!_hwnd)
	{
		wprintf(L"Failed to create the Taskman HWND. Immersive will fail to initialise.\n");
	}
}

LRESULT CTaskmanWindow::v_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
		case WM_CREATE:
		{
			printf("Creating Taskman window...\n");

			_wmShellHook = RegisterWindowMessage(TEXT("SHELLHOOK"));

			if (!_wmShellHook)
			{
				printf("Failed to register shellhook window message.\n");
				return 0;
			}

			if (!SetTaskmanWindow(_hwnd))
			{
				printf("Failed to set the Taskman window to our own (HWND = %p).\n", _hwnd);
				return 0;
			}

			if (!RegisterShellHookWindow(_hwnd))
			{
				printf("Failed to register our Taskman window as the shell hook window (HWND = %p).", _hwnd);
				return 0;
			}

			printf("Successfully created and initialised Taskman window.\n");

			break;
		}

		case WM_DESTROY:
		{
			if (GetTaskmanWindow() == _hwnd)
			{
				SetTaskmanWindow(NULL);
			}

			DeregisterShellHookWindow(_hwnd);

			break;
		}

		default:
		{
			if (uMsg == _wmShellHook)
			{
				if (!g_pShellHookService)
				{
					// Initialise the shell hook service:
					wprintf(L"Shell hook service not initialised. Initialising now.\n");
					IServiceProvider *pServiceProvider = nullptr;
					if (SUCCEEDED(CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_NO_CODE_DOWNLOAD | CLSCTX_LOCAL_SERVER,
						IID_PPV_ARGS(&pServiceProvider))))
					{
						if (FAILED(pServiceProvider->QueryService(SID_ImmersiveShellHookService, IID_PPV_ARGS(&g_pShellHookService))))
						{
							wprintf(L"Failed to query SID_ImmersiveShellHookService.\n");
						}

						pServiceProvider->Release();
					}
					else
					{
						wprintf(L"Failed to initialise shell hook service.\n");
					}
				}
				else
				{
					bool fForward = true;

					if (wParam == 12) // TODO: Document
					{
						wprintf(L"Setting the target window for serialisation to %p.\n", (HWND)lParam);
						g_pShellHookService->SetTargetWindowForSerialization((HWND)lParam);
					}
					else if (wParam == 50) // TODO: Document
					{
						wprintf(L"Received message %d. This message will not be forwarded.\n", (WORD)wParam);
						fForward = false;
					}
					else if ((WORD)wParam == 7)
					{
						wprintf(L"Received message %d. This message will not be forwarded. WE opening the start menu with this shit\n", (WORD)wParam);
						fForward = false;
						c_tray.TellTaskBandWeWantToOpenThisShit();
					}

					if (fForward)
					{
						wprintf(L"[Shell Hook] Forwarding message 0d%d to 0x%p...\n", (WORD)wParam, (HWND)lParam);
						g_pShellHookService->PostShellHookMessage(wParam, lParam);
					}

					return 0;
				}
			}

			break;
		}
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

CImmersiveBehaviorWrapper::CImmersiveBehaviorWrapper(IImmersiveBehavior* behavior)
{
	m_cRef = 1;
	m_behavior = behavior;
	m_behavior->AddRef();
}

CImmersiveBehaviorWrapper::~CImmersiveBehaviorWrapper()
{
	wprintf(L"CImmersiveBehaviorWrapper::~CImmersiveBehaviorWrapper()");
	m_behavior->Release();
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::QueryInterface(REFIID riid, void** ppvObject)
{
	WCHAR iid[100];
	StringFromGUID2(riid, iid, 100);
	wprintf(L"CImmersiveBehaviorWrapper::QueryInterface %s", iid);
	if (riid == IID_ImmersiveBehavior)
	{
		*ppvObject = static_cast<IImmersiveBehavior*>(this);
		return S_OK;
	}
	return m_behavior->QueryInterface(riid, ppvObject);
}

ULONG STDMETHODCALLTYPE CImmersiveBehaviorWrapper::AddRef(void)
{
	return InterlockedIncrement(&m_cRef);
}

ULONG STDMETHODCALLTYPE CImmersiveBehaviorWrapper::Release(void)
{
	wprintf(L"CImmersiveBehaviorWrapper::release()\n");
	if (InterlockedDecrement(&m_cRef) == 0)
	{
		delete this;
		return 0;
	}
	return m_cRef;
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::OnImmersiveThreadStart(void)
{
	wprintf(L"CImmersiveBehaviorWrapper::OnImmersiveThreadStart\n");
	return m_behavior->OnImmersiveThreadStart();
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::OnImmersiveThreadStop(void)
{
	wprintf(L"CImmersiveBehaviorWrapper::OnImmersiveThreadStop\n");
	return m_behavior->OnImmersiveThreadStart();
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::GetMaximumComponentCount(unsigned int* count)
{
	wprintf(L"CImmersiveBehaviorWrapper::GetMaximumComponentCount %p\n", count);
	return m_behavior->GetMaximumComponentCount(count);
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::CreateComponent(unsigned int number, IUnknown** component)
{
	//if (number == 1) DebugBreak();
	HRESULT ret = m_behavior->CreateComponent(number, component);
	wprintf(L"CImmersiveBehaviorWrapper::CreateComponent %d = %p\n", number, ret);
	/*IUnknown* wtf = *component;
	HRESULT ret2 = wtf->QueryInterface(IID_ImmersiveShell,(PVOID*)&wtf);
	wprintf(L"CImmersiveBehaviorWrapper::GetInterfaceList %p",ret2);*/
	return ret;
}

HRESULT STDMETHODCALLTYPE CImmersiveBehaviorWrapper::ShouldCreateComponent(unsigned int number, int* allowed)
{
	wprintf(L"CImmersiveBehaviorWrapper::ShouldCreateComponent %d %p\n", number, allowed);
	if (number == 9)
	{
		*allowed = 0;
		return S_OK;
	}
	return m_behavior->ShouldCreateComponent(number, allowed);
}

HRESULT InitializeImmersiveShell()
{
	IImmersiveShellBuilder *pShellBuilder = nullptr;
	HRESULT hr = S_OK;

	wprintf(L"Initialising the Immersive shell...\n");

	hr = CoCreateInstance(CLSID_ImmersiveShellBuilder, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pShellBuilder));

	if (FAILED(hr))
	{
		wprintf(L"Failed to create the Immersive Shell builder.\n");
		return hr;
	}

	IImmersiveShellController *pShellController = nullptr;

	hr = pShellBuilder->CreateImmersiveShellController(&pShellController);

	if (FAILED(hr))
	{
		wprintf(L"Failed to create the Immersive Shell Controller.\n");
		return hr;
	}

	wprintf(L"TwinUI instance created.\n");
	wprintf(L"   Pointer of IImmersiveShellBuilder instance: %p\n", pShellBuilder);
	wprintf(L"   Pointer of IImmersiveShellController instance: %p\n", pShellController);

	//// Below is what the Explorer7 source code does. I am not sure why any of this
	//// is written the way it is. Revalidate this in the future please.
	//IStream *pSomeInterface = (IStream *)*(DWORD *)((DWORD_PTR)pShellController + 0x34);
	//IImmersiveBehavior *pSomeBehavior = nullptr;
	//hr = CoUnmarshalInterface(pSomeInterface, IID_PPV_ARGS(&pSomeBehavior));

	//if (FAILED(hr))
	//{
	//	wprintf(L"Failed to unmarshal the immersive behavior (pointer = %p) from \n", pSomeBehavior);
	//	wprintf(L"the internal IStream (pointer = %p) on the controller.\n", pSomeInterface);
	//	return hr;
	//}

	//hr = pShellController->SetCreationBehavior(new CImmersiveBehaviorWrapper(pSomeBehavior));

	//if (FAILED(hr))
	//	return hr;

	hr = pShellController->Start();

	if (FAILED(hr))
	{
		wprintf(L"Failed to start the Immersive Shell Controller.\n");
	}

	return hr;
}

#endif