// ==WindhawkMod==
// @id              explorerex-inactive-theme-loader
// @name            ExplorerEx Inactive Theme Loader
// @description     Loads an "inactive theme", or alternate visual style, for the taskbar in ExplorerEx.
// @version         1.1
// @author          Isabella Lulamoon (kawapure)
// @github          https://github.com/kawapure
// @twitter         https://twitter.com/kawaipure
// @include         explorer.exe
// @compilerOptions -luxtheme -ldwmapi -lshlwapi
// ==/WindhawkMod==

// ==WindhawkModReadme==
/*
# ExplorerEx inactive theme loader

This mod allows you to use an inactive theme, which is one other than the global theme shared by Winlogon's theme server,
with ExplorerEx for the taskbar and start menu.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- theme: "C:\\Windows\\Resources\\Themes\\Aero\\Aero.msstyles"
  $name: The full path to the secondary visual style. Blank or invalid entry will fall back to classic theme.
*/
// ==/WindhawkModSettings==

#include <processenv.h>
#include <windhawk_api.h>
#include <windhawk_utils.h>
#include <libloaderapi.h>
#include <memoryapi.h>
#include <uxtheme.h>
#include <winerror.h>
#include <winnt.h>
#include <tlhelp32.h>
#include <dwmapi.h>
#include <shlwapi.h>
#include <vector>

LPCWSTR g_rgszTaskbarClasses[] = {
    L"Clock",
    L"TrayNotify",
    L"TrayNotifyFlyout", // Windows 7
    L"TaskBar",
    L"TaskBar::ComboBox",
    L"TaskBar::Edit",
    L"TaskBar::Rebar",
    L"TaskBar::Toolbar",
    L"TaskbarPearl", // Windows Vista
    L"TaskBand",
    L"TaskBand::Scrollbar",
    L"TaskBand::Toolbar",
    L"TaskbandExtendedUI", // Windows Vista
    L"TaskBand2",
    L"TaskBand2::Scrollbar",
    L"TaskbarShowDesktop", // Windows 7
    L"Start::Button",
    L"StartTop::Button",    // Windows Vista
    L"StartMiddle::Button", // Windows Vista
    L"StartBottom::Button", // Windows Vista
    L"StartMenu",
    L"StartPanel",
    L"MoreProgramsArrow",
    L"LogoffButtons",
    L"MoreProgramsArrowBack", // Windows Vista
    L"MoreProgramsTab", // Windows Vista
    L"SoftwareExplorer", // Windows Vista
    L"OpenBox", // Windows Vista
    L"StartPanelPriv", // Windows Vista
    L"TrayNotifHoriz::Button",
    L"TrayNotifyHorizHCWhite::Button",
    L"TrayNotifyHorizOpen::Button",
    L"TrayNotifyHorizOpenHCWhite::Button",
    L"TrayNotifyVert::Button",
    L"TrayNotifyVertHCWhite::Button",
    L"TrayNotifyVertOpen::Button",
    L"TrayNotifyVertOpenHCWHite::Button"
};

bool IsTaskbarClass(HWND hWnd, LPCWSTR szClassName)
{
    ATOM atValue = (ATOM)(size_t)GetPropW(hWnd, (LPCWSTR)0xA911);
    WCHAR szBuf[260] = { L'\0' };

    if (atValue)
    {
        if (GetAtomNameW(atValue, szBuf, ARRAYSIZE(szBuf)))
        {
            wcscat(szBuf, L"::");
            wcscat(szBuf, szClassName);
        }
    }

    for (UINT i = 0; i < ARRAYSIZE(g_rgszTaskbarClasses); i++)
    {
        if (StrStrIW(szClassName, g_rgszTaskbarClasses[i]))
        {
            return true;
        }
        else if (szBuf[0] && StrStrIW(szBuf, g_rgszTaskbarClasses[i]))
        {
            return true;
        }
    }

    return false;
}

#if __WIN64
#define WINAPI_STR L"__cdecl"
#else
#define WINAPI_STR L"__stdcall"
#endif

#if __WIN64
#define THISCALL_STR L"__cdecl"
#else
#define THISCALL_STR L"__thiscall"
#endif

typedef HRESULT WINAPI (*GetThemeDefaults_t)(
    LPCWSTR pszThemeFileName,
    LPWSTR pszColorName,
    DWORD dwColorNameLen,
    LPWSTR pszSizeName,
    DWORD dwSizeNameLen
);
GetThemeDefaults_t GetThemeDefaults;

typedef HRESULT WINAPI (*LoaderLoadTheme_t)(
	HANDLE hThemeFile,
	HINSTANCE hInstance,
	LPCWSTR pszThemeFileName,
	LPCWSTR pszColorParam,
	LPCWSTR pszSizeParam,
	OUT HANDLE *hSharableSection,
	LPWSTR pszSharableSectionName,
	int cchSharableSectionName,
	OUT HANDLE *hNonsharableSection,
	LPWSTR pszNonsharableSectionName,
	int cchNonsharableSectionName,
	PVOID pfnCustomLoadHandler,
	OUT HANDLE *hReuseSection,
	int a,
	int b,
	BOOL fEmulateGlobal
);
LoaderLoadTheme_t LoaderLoadTheme;

typedef HRESULT WINAPI (*LoaderLoadTheme_t_win11)(
	HANDLE hThemeFile,
	HINSTANCE hInstance,
	LPCWSTR pszThemeFileName,
	LPCWSTR pszColorParam,
	LPCWSTR pszSizeParam,
	OUT HANDLE *hSharableSection,
	LPWSTR pszSharableSectionName,
	int cchSharableSectionName,
	OUT HANDLE *hNonsharableSection,
	LPWSTR pszNonsharableSectionName,
	int cchNonsharableSectionName,
	PVOID pfnCustomLoadHandler,
	OUT HANDLE *hReuseSection,
	int a,
	int b
);

typedef HTHEME WINAPI  (*OpenThemeDataFromFile_t)(
    HANDLE hThemeFile,
    HWND hWnd,
    LPCWSTR pszClassList,
    DWORD dwFlags
    // DWORD unknown,
    // bool a
);

OpenThemeDataFromFile_t OpenThemeDataFromFile;

typedef struct _LocalThemeFile
{
    char header[7]; // must be "thmfile"
    LPVOID sharableSectionView;
    HANDLE hSharableSection;
    LPVOID nsSectionView;
    HANDLE hNsSection;
    char end[3]; // must be "end"
} LocalThemeFile;

//=================================================================================================

HANDLE g_hLocalTheme;

using OpenThemeData_t = decltype(&OpenThemeData);
OpenThemeData_t OpenThemeData_orig;
HTHEME WINAPI OpenThemeData_hook(HWND hwnd, LPCWSTR pszClassList)
{
    if (pszClassList && IsTaskbarClass(hwnd, pszClassList))
    {
        HTHEME fromFile = OpenThemeDataFromFile(g_hLocalTheme, hwnd, pszClassList, 0);

        if (fromFile)
        {
            return fromFile;
        }
    }

    return OpenThemeData_orig(hwnd, pszClassList);
}

using OpenThemeDataEx_t = decltype(&OpenThemeDataEx);
OpenThemeDataEx_t OpenThemeDataEx_orig;
HTHEME WINAPI OpenThemeDataEx_hook(HWND hwnd, LPCWSTR pszClassList, DWORD dwFlags)
{
    if (pszClassList && IsTaskbarClass(hwnd, pszClassList))
    {
        if (g_hLocalTheme)
        {
            HTHEME fromFile = OpenThemeDataFromFile(g_hLocalTheme, hwnd, pszClassList, 0);
        
            if (fromFile)
            {
                return fromFile;
            }
        }
        else
        {
            return NULL;
        }
    }

    return OpenThemeDataEx_orig(hwnd, pszClassList, dwFlags);
}

typedef HTHEME __fastcall (*_OpenThemeData_t)(HWND hwnd, PCWSTR pszClassList, DWORD dwFlags, int unk1, bool unk2);
_OpenThemeData_t _OpenThemeData_orig;
HTHEME __fastcall _OpenThemeData_hook(HWND hwnd, PCWSTR pszClassList, DWORD dwFlags, int unk1, bool unk2)
{
    if (pszClassList && IsTaskbarClass(hwnd, pszClassList))
    {
        if (g_hLocalTheme)
        {
            HTHEME fromFile = OpenThemeDataFromFile(g_hLocalTheme, hwnd, pszClassList, 0);
        
            if (fromFile)
            {
                return fromFile;
            }
        }
        else
        {
            return NULL;
        }
    }

    return _OpenThemeData_orig(hwnd, pszClassList, dwFlags, unk1, unk2);
}

/*
 * Load a visual style from the provided file path.
 *
 * This relies on a few internal functions from uxtheme in order to work. It's
 * basically the same approach as StartIsBack.
 */
HRESULT LoadThemeFromFilePath(PCWSTR szThemeFileName)
{
    HRESULT hr = S_OK;

    HMODULE hUxtheme = GetModuleHandleW(L"uxtheme.dll");

    if (!hUxtheme)
    {
        return E_FAIL;
    }

    GetThemeDefaults = (GetThemeDefaults_t)GetProcAddress(hUxtheme, (LPCSTR)7);
    LoaderLoadTheme = (LoaderLoadTheme_t)GetProcAddress(hUxtheme, (LPCSTR)92);
    OpenThemeDataFromFile = (OpenThemeDataFromFile_t)GetProcAddress(hUxtheme, (LPCSTR)16);

    if (!GetThemeDefaults || !LoaderLoadTheme || !OpenThemeDataFromFile)
    {
        return E_FAIL;
    }

    OSVERSIONINFOW verInfo = { 0 };
    hr = GetVersionExW(&verInfo) ? S_OK : E_FAIL;

    WCHAR defColor[MAX_PATH];
    WCHAR defSize[MAX_PATH];

    hr = GetThemeDefaults(
        szThemeFileName,
        defColor,
        ARRAYSIZE(defColor),
        defSize,
        ARRAYSIZE(defSize)
    );

    HANDLE hSharableSection;
    HANDLE hNonsharableSection;

    if (verInfo.dwBuildNumber < 20000)
    {
        hr = LoaderLoadTheme(
            NULL,
            NULL,
            szThemeFileName,
            defColor,
            defSize,
            &hSharableSection,
            NULL,
            0,
            &hNonsharableSection,
            NULL,
            0,
            NULL,
            NULL,
            NULL,
            NULL,
            FALSE
        );
    }
    else
    {
        hr = ((LoaderLoadTheme_t_win11)LoaderLoadTheme)(
            NULL,
            NULL,
            szThemeFileName,
            defColor,
            defSize,
            &hSharableSection,
            NULL,
            0,
            &hNonsharableSection,
            NULL,
            0,
            NULL,
            NULL,
            NULL,
            NULL
        );
    }

    if (SUCCEEDED(hr))
    {
        g_hLocalTheme = malloc(sizeof(LocalThemeFile));
        if (g_hLocalTheme)
        {
            LocalThemeFile *ltf = (LocalThemeFile *)g_hLocalTheme;
            lstrcpyA(ltf->header, "thmfile");
            lstrcpyA(ltf->header, "end");
            ltf->sharableSectionView = MapViewOfFile(hSharableSection, 4, 0, 0, 0);
            ltf->hSharableSection = hSharableSection;
            ltf->nsSectionView = MapViewOfFile(hNonsharableSection, 4, 0, 0, 0);
            ltf->hNsSection = hNonsharableSection;
        }
        else
        {
            hr = E_FAIL;
        }
    }
    else
    {
        g_hLocalTheme = NULL;
        hr = E_FAIL;
    }

    return hr;
}

/*
 * Load the Windhawk mod settings for the current program. If the user enabled a different theme 
 * for the program, then this will return true. Otherwise, it will return false.
 *
 * This implementation is heavily copied from the "Text Replace" mod.
 */
bool LoadSettings()
{
    PCWSTR themePath = Wh_GetStringSetting(L"theme");
    bool hasAppliedTheme = false;     
    if (*themePath)
    {
        LoadThemeFromFilePath(themePath);
    }
    else
    {
        g_hLocalTheme = NULL;
    }
    Wh_FreeStringSetting(themePath);
    return hasAppliedTheme;
}

#define WM_UAHINIT                      0x031B
#define WM_THEMECHANGED_TRIGGER         WM_UAHINIT
#define WTC_THEMEACTIVE     (1 << 0)        // new theme is now active
#define WTC_CUSTOMTHEME     (1 << 1)        // this msg for custom-themed apps

BOOL CALLBACK EnumWindowsChildCallBackFnc(HWND hwnd, LPARAM lParam)
{
    SendMessage(hwnd, WM_UAHINIT, NULL, 0);
    SendMessage(hwnd, WM_THEMECHANGED, WPARAM(-1), WTC_THEMEACTIVE | WTC_CUSTOMTHEME);
    return TRUE;
}

BOOL CALLBACK EnumWindowsCallBackFnc(HWND hwnd, LPARAM lParam)
{
    SendMessage(hwnd, WM_UAHINIT, NULL, 0);
    SendMessage(hwnd, WM_THEMECHANGED, WPARAM(-1), WTC_THEMEACTIVE | WTC_CUSTOMTHEME);
    SendMessage(hwnd, WM_THEMECHANGED_TRIGGER, WPARAM(-1), (WPARAM(-1) << 4) | ((WTC_THEMEACTIVE | WTC_CUSTOMTHEME) & 0xf));
    EnumChildWindows(hwnd, EnumWindowsChildCallBackFnc, lParam);
    return TRUE;
}


bool UpdateWindowThemes(DWORD dwProcessId)
{
    HANDLE thread_snap = INVALID_HANDLE_VALUE;
    THREADENTRY32 te32;

    // Take a snapshot of all running threads:
    thread_snap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
    if (thread_snap == INVALID_HANDLE_VALUE) {
        return(FALSE);
    }

    // Fill in the size of the structure before using it:
    te32.dwSize = sizeof(THREADENTRY32);

    // Retrieve information about the first thread, and exit if unsuccessful:
    if (!Thread32First(thread_snap, &te32)) {
        CloseHandle(thread_snap);
        return(FALSE);
    }

    // Now walk the thread list of the system, and display thread IDs of each thread
    // associated with the specified process:
    do
    {
        if (te32.th32OwnerProcessID == dwProcessId)
        {
            EnumThreadWindows(te32.th32ThreadID, EnumWindowsCallBackFnc, NULL);
        }
    }
    while (Thread32Next(thread_snap, &te32));

    // clean up the snapshot object.
    CloseHandle(thread_snap);
    return(TRUE);
}

DWORD WINAPI ThemeChangeThread(LPVOID lpParam)
{
    SleepEx(2, NULL);
    SleepEx(2, NULL);
    UpdateWindowThemes(GetCurrentProcessId());
    return 0;
}

static std::vector<int>* patternToByte(const char* pattern)
{
	auto bytes = new std::vector<int>();
	const auto start = const_cast<char*>(pattern);
	const auto end = const_cast<char*>(pattern) + strlen(pattern);

	for (auto current = start; current < end; ++current)
	{
		if (*current == '?')
		{
			++current;
			if (*current == '?')
				++current;
			bytes->push_back(-1);
		}
		else { bytes->push_back(strtoul(current, &current, 16)); }
	}
	return bytes;
}

static uintptr_t FindPattern(uintptr_t baseAddress, const char* signature)
{
	const auto dosHeader = (PIMAGE_DOS_HEADER)baseAddress;
	const auto ntHeaders = (PIMAGE_NT_HEADERS)((unsigned char*)baseAddress + dosHeader->e_lfanew);

	const auto sizeOfImage = ntHeaders->OptionalHeader.SizeOfImage;
	auto patternBytes = patternToByte(signature);
	const auto scanBytes = reinterpret_cast<unsigned char*>(baseAddress);

	const auto s = patternBytes->size();
	const auto d = patternBytes->data();

	for (size_t i = 0; i < sizeOfImage - s; ++i)
	{
		bool found = true;
		for (size_t j = 0; j < s; ++j)
		{
			if (scanBytes[i + j] != d[j] && d[j] != -1)
			{
				found = false;
				break;
			}
		}

		if (found)
		{
			uintptr_t address = reinterpret_cast<uintptr_t>(&scanBytes[i]);

			delete patternBytes;
			return address;
		}
	}

	delete patternBytes;

	return NULL;
}

//Ittr: Consolidated function for pattern byte replacements.
static void ChangeImportedPattern(void* dllPattern, const unsigned char* newBytes, SIZE_T size) //thank you wiktor
{
	if (dllPattern)
	{
		DWORD old;
		VirtualProtect(dllPattern, size, PAGE_EXECUTE_READWRITE, &old);
		memcpy(dllPattern, newBytes, size);
		VirtualProtect(dllPattern, size, old, 0);
	}
}

// Remove AMAP class from loaded msstyle so that Vista and 7 msstyles are compatible
void RemoveLoadAnimationDataMap()
{
	// 48 8B 53 20 48 8B ?? E8 ?? ?? ?? ?? 8B ?? 48 8B ?? E8 ?? ?? ?? ?? 8B ?? EB 05 B8 57 00 07 80
	// thank you amrsatrio for the pattern + offsetting method
	const char* LoadAnimationDataMap = "48 8B 53 20 48 8B ?? E8 ?? ?? ?? ?? 8B ?? 48 8B";

	HMODULE uxTheme = GetModuleHandle(L"uxtheme.dll");
	if (uxTheme)
	{
		char* LADMPattern = (char*)FindPattern((uintptr_t)uxTheme, LoadAnimationDataMap);

		if (LADMPattern)
		{
			LADMPattern += 7;
			LADMPattern += 5 + *(int*)(LADMPattern + 1);

			unsigned char bytes[] = { 0x31, 0xC0, 0xC3 }; // mov eax 0, ret
			ChangeImportedPattern(LADMPattern, bytes, sizeof(bytes));
		}
	}
}

// The mod is being initialized, load settings, hook functions, and do other
// initialization stuff if required.
BOOL Wh_ModInit()
{
    Wh_Log(L"Init " WH_MOD_ID L" version " WH_MOD_VERSION);

    HMODULE hUxtheme = GetModuleHandleW(L"uxtheme.dll");

    if (hUxtheme)
    {
        GetThemeDefaults = (GetThemeDefaults_t)GetProcAddress(hUxtheme, (LPCSTR)7);
        LoaderLoadTheme = (LoaderLoadTheme_t)GetProcAddress(hUxtheme, (LPCSTR)92);
        OpenThemeDataFromFile = (OpenThemeDataFromFile_t)GetProcAddress(hUxtheme, (LPCSTR)16);

        FARPROC OpenNcThemeData = GetProcAddress(hUxtheme, (LPCSTR)49);
        if (GetThemeDefaults && LoaderLoadTheme && OpenThemeDataFromFile && OpenNcThemeData)
        {
            Wh_SetFunctionHook(
                (void *)OpenThemeData,
                (void *)OpenThemeData_hook,
                (void **)&OpenThemeData_orig
            );
    
            Wh_SetFunctionHook(
                (void *)OpenThemeDataEx,
                (void *)OpenThemeDataEx_hook,
                (void **)&OpenThemeDataEx_orig
            );

            WindhawkUtils::SYMBOL_HOOK rgHooks[] = {
                {
                    { L"void * " WINAPI_STR L" _OpenThemeData(struct HWND__ *,unsigned short const *,int,unsigned long,bool)" },
                    (void **)&_OpenThemeData_orig,
                    (void *)_OpenThemeData_hook
                },
            };


            if (!WindhawkUtils::HookSymbols(hUxtheme, rgHooks, ARRAYSIZE(rgHooks)))
            {
                Wh_Log(L"Failed to hook.");
            }
            
            // Remove the animation data map now that symbol hooks are done.
            // For some reason, I can't seem to get it to work at all with a symbol hook.
            RemoveLoadAnimationDataMap();

            // Now that everything is set up, apply the theme:
            LoadSettings();
        }
    }

    CreateThread(NULL, 0, ThemeChangeThread, NULL, NULL, NULL);
    SleepEx(0, NULL);

    return TRUE;
}

void Wh_ModSettingsChanged()
{
    LoadSettings();
    UpdateWindowThemes(GetCurrentProcessId());
}

// The mod is being unloaded, free all allocated resources.
void Wh_ModUninit()
{
    UpdateWindowThemes(GetCurrentProcessId());
}
