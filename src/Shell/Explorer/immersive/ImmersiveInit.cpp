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

// Windows 11 registers Win and Ctrl+Esc independently of HSHELL_TASKMAN.
// Block those two registrations in TwinUI only, restoring the legacy Start
// route. Windows 10 lacks this import and needs no patch. Install outside
// DllMain, before creating/starting the immersive controller.
using ShellRegisterHotKeyFn = BOOL (WINAPI*)(HWND, int, UINT, UINT, HWND);
static ShellRegisterHotKeyFn g_originalShellRegisterHotKey = nullptr;

static BOOL WINAPI ShellRegisterHotKeyHook(HWND hwnd, int id, UINT modifiers, UINT key, HWND target)
{
    if ((modifiers == MOD_WIN && key == 0) ||
        (modifiers == MOD_CONTROL && key == VK_ESCAPE))
    {
        SetLastError(ERROR_HOTKEY_ALREADY_REGISTERED);
        return FALSE;
    }
    return g_originalShellRegisterHotKey(hwnd, id, modifiers, key, target);
}

static HRESULT PatchTwinUIStartHotkeys(HMODULE module)
{
    BYTE* base = reinterpret_cast<BYTE*>(module);
    const IMAGE_DOS_HEADER* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew <= 0)
        return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    const IMAGE_NT_HEADERS* nt = reinterpret_cast<const IMAGE_NT_HEADERS*>(base + dos->e_lfanew);
    if (nt->Signature != IMAGE_NT_SIGNATURE)
        return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
    const DWORD imageSize = nt->OptionalHeader.SizeOfImage;
    const IMAGE_DATA_DIRECTORY& directory = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (!directory.VirtualAddress || !directory.Size)
        return S_FALSE;
    if (directory.VirtualAddress >= imageSize || directory.Size > imageSize - directory.VirtualAddress)
        return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);

    const IMAGE_IMPORT_DESCRIPTOR* imports =
        reinterpret_cast<const IMAGE_IMPORT_DESCRIPTOR*>(base + directory.VirtualAddress);
    for (DWORD i = 0; i < directory.Size / sizeof(*imports) && imports[i].Name; ++i)
    {
        const IMAGE_IMPORT_DESCRIPTOR& entry = imports[i];
        if (entry.Name >= imageSize || !memchr(base + entry.Name, 0, imageSize - entry.Name))
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        if (_stricmp(reinterpret_cast<const char*>(base + entry.Name), "user32.dll") != 0)
            continue;
        if (!entry.OriginalFirstThunk || !entry.FirstThunk)
            return S_FALSE;
        if (entry.OriginalFirstThunk >= imageSize || entry.FirstThunk >= imageSize)
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        const IMAGE_THUNK_DATA* names =
            reinterpret_cast<const IMAGE_THUNK_DATA*>(base + entry.OriginalFirstThunk);
        IMAGE_THUNK_DATA* addresses = reinterpret_cast<IMAGE_THUNK_DATA*>(base + entry.FirstThunk);
        const SIZE_T count = min((imageSize - entry.OriginalFirstThunk) / sizeof(*names),
                                (imageSize - entry.FirstThunk) / sizeof(*addresses));
        for (SIZE_T j = 0; j < count && names[j].u1.AddressOfData; ++j)
        {
            if (!IMAGE_SNAP_BY_ORDINAL(names[j].u1.Ordinal) ||
                IMAGE_ORDINAL(names[j].u1.Ordinal) != 2671)
                continue;
            PVOID* slot = reinterpret_cast<PVOID*>(&addresses[j].u1.Function);
            if (*slot == reinterpret_cast<PVOID>(&ShellRegisterHotKeyHook))
                return S_OK;
            if (!*slot)
                return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
            DWORD protection = 0;
            if (!VirtualProtect(slot, sizeof(*slot), PAGE_READWRITE, &protection))
                return HRESULT_FROM_WIN32(GetLastError());
            // Publish the original before the hook becomes callable.
            g_originalShellRegisterHotKey = reinterpret_cast<ShellRegisterHotKeyFn>(*slot);
            InterlockedExchangePointer(slot, reinterpret_cast<PVOID>(&ShellRegisterHotKeyHook));
            DWORD ignored = 0;
            if (!VirtualProtect(slot, sizeof(*slot), protection, &ignored))
                return HRESULT_FROM_WIN32(GetLastError());
            return S_OK;
        }
    }
    return S_FALSE;
}

static HRESULT InstallStartHotkeyWorkaround()
{
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (!user32 || !GetProcAddress(user32, MAKEINTRESOURCEA(2671)))
        return S_FALSE;

    WCHAR path[MAX_PATH];
    const UINT length = GetSystemDirectoryW(path, ARRAYSIZE(path));
    if (!length || length >= ARRAYSIZE(path) || !PathAppendW(path, L"twinui.dll"))
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    // Retain the module: the patched import and original callback live here.
    static HMODULE twinui = LoadLibraryExW(path, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!twinui)
        return HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND);
    return PatchTwinUIStartHotkeys(twinui);
}

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
			if (_wmShellHook && uMsg == _wmShellHook)
			{
				// Local Start/task-list handling must not depend on native COM.
				if (LOWORD(wParam) == HSHELL_TASKMAN)
				{
					c_tray.TellTaskBandWeWantToOpenThisShit(wParam, lParam);
					return 0;
				}
				if (wParam == 50) // Existing immersive-shell exclusion.
					return 0;

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
				// Process this event even when the service was just acquired.
				if (g_pShellHookService)
				{
					if (wParam == 12) // Set the serialization target before forwarding.
						g_pShellHookService->SetTargetWindowForSerialization((HWND)lParam);
					g_pShellHookService->PostShellHookMessage(wParam, lParam);
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

    HRESULT hotkeyHr = InstallStartHotkeyWorkaround();
    if (FAILED(hotkeyHr))
        wprintf(L"Failed to install Windows 11 Start hotkey workaround (0x%08lX).\n", hotkeyHr);

	HANDLE hShellStartup = OpenEvent(EVENT_MODIFY_STATE, FALSE, L"Local\\ShellStartupEvent");
	if (hShellStartup)
	{
		SetEvent(hShellStartup);
		CloseHandle(hShellStartup);
	}

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