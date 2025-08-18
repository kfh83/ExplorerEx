#pragma once

#ifdef EXEX_DLL
#include "pch.h"
#include "MinHook.h"

inline BOOL WINAPI RetTrue()
{
	return TRUE;
}

inline void FixUpImmersiveShell()
{
	// NOTE: Some of the patch functions are in util.h rather than an imports header because they are used with several patch types
	////////////////////////////////
	// 1. Todo in future *after* feature-set is complete: see how many of these hooks can be ChangeImportedAddress instead of MH_CreateHook (perf optimisation)
	// 2. Code stack used exclusively for UWP mode, hence the conditional statement.
	// 3. This will *need* serious optimization in the near future as it singlehandedly delays program enumeration and startup by several seconds
	// 4. Prepare the taskbar and thumbnails to handle UWP icons. Further work needed for jumplists and to prevent wrongful classification as "Application Frame Host" in the first place.

	//// The rest of this code block is dedicated to ensuring UWP actually runs in the first place
	//CreateWindowInBandOrig = decltype(CreateWindowInBandOrig)(GetProcAddress(GetModuleHandle(L"user32.dll"), "CreateWindowInBand"));
	//CreateWindowInBandExOrig = decltype(CreateWindowInBandExOrig)(GetProcAddress(GetModuleHandle(L"user32.dll"), "CreateWindowInBandEx"));
	//SetWindowBandApiOrg = decltype(SetWindowBandApiOrg)(GetProcAddress(GetModuleHandle(L"user32.dll"), "SetWindowBand"));
	//RegisterHotKeyApiOrg = decltype(RegisterHotKeyApiOrg)(GetProcAddress(GetModuleHandle(L"user32.dll"), "RegisterHotKey"));

	//MH_CreateHook(static_cast<LPVOID>(CreateWindowInBandOrig), CreateWindowInBandNew, reinterpret_cast<LPVOID*>(&CreateWindowInBandOrig));
	//MH_CreateHook(static_cast<LPVOID>(CreateWindowInBandExOrig), CreateWindowInBandExNew, reinterpret_cast<LPVOID*>(&CreateWindowInBandExOrig));
	//MH_CreateHook(static_cast<LPVOID>(SetWindowBandApiOrg), SetWindowBandNew, reinterpret_cast<LPVOID*>(&SetWindowBandApiOrg));
	////MH_CreateHook(static_cast<LPVOID>(RegisterHotKeyApiOrg), RegisterWindowHotkeyNew, reinterpret_cast<LPVOID*>(&RegisterHotKeyApiOrg));

	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2581), RetTrue, NULL); // GetWindowTrackInfoAsync
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2563), RetTrue, NULL); // ClearForeground
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2628), RetTrue, NULL); // CreateWindowGroup
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2629), RetTrue, NULL); // DeleteWindowGroup
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2631), RetTrue, NULL); // EnableWindowGroupPolicy
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2627), RetTrue, NULL); // SetBridgeWindowChild
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2511), RetTrue, NULL); // SetFallbackForeground
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2566), RetTrue, NULL); // SetWindowArrangement
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2632), RetTrue, NULL); // SetWindowGroup
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2579), RetTrue, NULL); // SetWindowShowState
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2585), RetTrue, NULL); // UpdateWindowTrackingInfo
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2514), RetTrue, NULL); // RegisterEdgy
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2542), RetTrue, NULL); // RegisterShellPTPListener
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2537), RetTrue, NULL); // SendEventMessage
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2513), RetTrue, NULL); // SetActiveProcessForMonitor
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2564), RetTrue, NULL); // RegisterWindowArrangementCallout
	MH_CreateHook(GetProcAddress(GetModuleHandle(L"user32.dll"), (LPCSTR)2567), RetTrue, NULL); // EnableShellWindowManagementBehavior
}

// Ittr: Under immersive mode, the differences in ShellHook operation have to be accounted for
HRESULT(__fastcall* OnShellHookMessage)(void* a1);

inline bool fShowLauncher = false; // Ittr: First run erroneously shows the start menu, unless we handle it differently

inline HRESULT OnShellHookMessage_Hook(void* a1) // This gets called when start menu is to be opened - has been a bit temperamental
{
	// Use of fShowLauncher flag is essential to have something resembling stability for overall functionality
	if (fShowLauncher) // If the flag is set...
	{
		//PostMessageW(hwnd_taskbar, 0x504, 0, 0); // Fire the message directly that opens Windows 7's start menu - ShellHook unreliable pre-VB
		return S_OK; // Ensure the run is recognised as a success 
	}
	else // However, without the flag, this is presumed to be the first run
	{
		fShowLauncher = true; // Enable the flag for showing the start menu now that this first attempt has run through
		return E_FAIL;  // Then, ensure this run is marked as a failure
	}

	return OnShellHookMessage(a1); // This codepath should ideally never run
}

inline void _OnHShellTaskMan()
{
	// Here we account for the immersive shell's destructive impacts upon certain internal mechanisms of explorer

	// Work out what we need for different OS versions (TH1 through GE)
	const char* XamlLauncher_OnShellHookMessage;
	char* XLOSHMPattern;

	// Check whether the modern DLL exists in Windows
	HMODULE twinUI_PCShell = LoadLibrary(L"twinui.pcshell.dll");
	if (twinUI_PCShell) // If it does...
	{
		XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 83 C1 D8 33 D2 E8 ?? ?? ?? ?? 8B D8";
		XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

		if (XLOSHMPattern) // NI, GE
		{
			MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
		}
		else
		{
			XamlLauncher_OnShellHookMessage = "48 89 5C 24 08 57 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 74 76 48 8B 01";
			XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

			if (XLOSHMPattern) // CO
			{
				MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
			}
			else
			{
				XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 74 ?? 48 8B 01 48 8B 40 ?? FF 15 ?? ?? ?? ?? 84 C0 0F 85 ?? ?? ?? ?? 38 83";
				XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

				if (XLOSHMPattern) // VB
				{
					MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
				}
				else
				{
					XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 74 4E 48 8B 01";
					XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

					if (XLOSHMPattern) // RS5, 19H1
					{
						MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
					}
					else
					{
						XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 74 59 48 8B 01"; // 0x59 cannot be wildcarded
						XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

						if (XLOSHMPattern) // RS4
						{
							MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
						}
						else
						{
							XamlLauncher_OnShellHookMessage = "48 89 5C 24 10 57 48 83 EC 30 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 75 0A";
							XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

							if (XLOSHMPattern) // RS3
							{
								MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
							}
							else
							{
								XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 75 07 B8 90 04 07 80 EB 6F 48 8B 01";
								XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinUI_PCShell);

								if (XLOSHMPattern) // RS2
								{
									MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
								}
								else
								{
									goto _OnHShellTaskMan_TWINUI; // New DLL exists on RS1 but unused for these purposes. Fall back to twinui.dll
								}
							}
						}
					}
				}
			}
		}
	}
	else // RS1 and earlier
	{
		_OnHShellTaskMan_TWINUI:
		HMODULE twinui = LoadLibrary(L"twinui.dll");

		if (twinui) // This should always exist, but we check in case it doesn't
		{
			XamlLauncher_OnShellHookMessage = "40 53 48 83 EC 20 48 8B D9 48 8B 89 ?? ?? ?? ?? 48 85 C9 ?? ?? ?? ?? ?? ?? 48 8B 01";
			XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinui);

			if (XLOSHMPattern) // RS1
			{
				MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
			}
			else
			{
				XamlLauncher_OnShellHookMessage = "48 89 5C 24 08 48 89 74 24 10 57 48 83 EC 20 48 8B B9 ?? ?? ?? ?? 48 8B D9 48 85 FF ?? ?? ?? ?? ?? ?? 48 8B 07";
				XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinui);

				if (XLOSHMPattern) // TH2
				{
					MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));
				}
				else
				{
					XamlLauncher_OnShellHookMessage = "48 89 5C 24 08 48 89 74 24 10 57 48 83 EC 20 48 8B B9 ?? ?? ?? ?? 48 8B D9 48 85 FF 74 62";
					const char* CImmersiveLauncher_OnShellHookMessage = "48 89 5C 24 10 55 56 57 48 8B EC 48 83 EC 20 FF"; // Server uses this

					XLOSHMPattern = (char*)FindPattern(XamlLauncher_OnShellHookMessage, (uintptr_t)twinui);;
					char* CILOSHMPattern = (char*)FindPattern(CImmersiveLauncher_OnShellHookMessage, (uintptr_t)twinui);

					if (XLOSHMPattern) // TH1
					{
						MH_CreateHook(static_cast<LPVOID>(XLOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&XLOSHMPattern));

						if (CILOSHMPattern) // Second stage for Server SKUs that use legacy CImmersiveLauncher
						{
							MH_CreateHook(static_cast<LPVOID>(CILOSHMPattern), OnShellHookMessage_Hook, reinterpret_cast<LPVOID*>(&CILOSHMPattern));
						}
					}
				}
			}
		}
	}
}

inline void HookEverythingForImmersive()
{
	MH_Initialize();

	//SetUpThemeManager(); // Local visual style management init
	//FixNonImmersivePniDui(); // Non-immersive network flyout handling
	//UpdateTrayWindowDefinitions(); // Ensure tray exclusion is corrected for modern Windows
	//SetProgramListNscTreeAttributes(); // Restore the relevant contents to the program list
	//HandleThumbnailColorization(); // Thumbnail colorization to match
	//RenderStoreAppsOnTaskbar(); // UWP icon rendering for the taskbar
	FixUpImmersiveShell(); // Immersive shell initialisation
	_OnHShellTaskMan(); // Handling the immersive shell's impacts on the holographic shell and associated ShellHook messages
}

#endif