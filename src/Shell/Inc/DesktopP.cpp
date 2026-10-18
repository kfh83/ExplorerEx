#include "pch.h"
#include "desktopp.h"
#include "cabinet.h"

HWND hwnd_desktop;
static WNDPROC g_prevTrayProc;
HANDLE hEvent_DesktopVisible;

static HANDLE(*fSHCreateDesktop)(IDeskTray* pdtray) ;
static BOOL(*fSHDesktopMessageLoop)(HANDLE hDesktop);

LRESULT CALLBACK NewTrayProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{

	if (uMsg == 0x56D) return 0;

	if (uMsg == WM_DISPLAYCHANGE || uMsg == WM_WINDOWPOSCHANGED)
	{
		RemoveProp(hwnd, L"TaskbarMonitor");
		SetProp(hwnd, L"TaskbarMonitor", (HANDLE)MonitorFromWindow(hwnd, MONITOR_DEFAULTTOPRIMARY));
		//send displaychanged to desktop
		if (uMsg == WM_DISPLAYCHANGE) PostMessage(hwnd_desktop, 0x44B, 0, 0);
	}
	if (uMsg == 0x574) //handledelayboot
	{
		if (lParam == 3)
			return CallWindowProc(g_prevTrayProc, hwnd, 0x5B5, wParam, lParam); //fire ShellDesktopSwitch event
		if (lParam == 1)
			SetEvent(hEvent_DesktopVisible);
		return 0;
	}

	return CallWindowProc(g_prevTrayProc, hwnd, uMsg, wParam, lParam);
}

void ShimDesktop()
{
	static int InitOnce = FALSE;
	if (InitOnce) return;

	hEvent_DesktopVisible = CreateEvent(NULL, TRUE, FALSE, L"ShellDesktopVisibleEvent");

	hwnd_desktop = FindWindow(L"Progman", L"Program Manager");
	HWND hwndTray = v_hwndTray;
	if (!hwnd_desktop || !hwndTray) return;
	InitOnce = TRUE;
	g_prevTrayProc = (WNDPROC)GetWindowLongPtr(hwndTray, GWLP_WNDPROC);
	SetWindowLongPtr(hwndTray, GWLP_WNDPROC, (LONG_PTR)NewTrayProc);
	//set monitor (doh!)
	SetProp(hwndTray, L"TaskbarMonitor", (HANDLE)MonitorFromWindow(hwndTray, MONITOR_DEFAULTTOPRIMARY));
	//init desktop	
	PostMessage(hwnd_desktop, 0x45C, 1, 1); //wallpaper
	PostMessage(hwnd_desktop, 0x45E, 0, 2); //wallpaper host
	PostMessage(hwnd_desktop, 0x45C, 2, 3); //wallpaper & icons
	PostMessage(hwnd_desktop, 0x45B, 0, 0); //final init
	PostMessage(hwnd_desktop, 0x40B, 0, 0); //pins
}

HANDLE SHCreateDesktop(IDeskTray* pdtray)
{
	
	
	HANDLE ret = fSHCreateDesktop(pdtray);
	ShimDesktop();
	return ret;
}

BOOL CreateFromDesktop(PNEWFOLDERINFO pfi)
{
    return 0;
}

BOOL SHCreateFromDesktop(PNEWFOLDERINFO pfi)
{
    return 0;
}

BOOL SHDesktopMessageLoop(HANDLE hDesktop)
{
	
	return fSHDesktopMessageLoop(hDesktop);
}

BOOL SHExplorerParseCmdLine(PNEWFOLDERINFO pfi)
{
    return 0;
}

LRESULT DDEHandleMsgs(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    return LRESULT();
}

void DDEHandleTimeout(HWND hwnd)
{
}

void InitDesktopFuncs()
{
	fSHCreateDesktop = (decltype(fSHCreateDesktop))GetProcAddress(LoadLibrary(L"shell32.dll"), (LPSTR)200);
	fSHDesktopMessageLoop = (decltype(fSHDesktopMessageLoop))GetProcAddress(LoadLibrary(L"shell32.dll"), (LPSTR)201);
}
