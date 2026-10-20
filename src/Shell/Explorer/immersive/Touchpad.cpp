#include "Touchpad.h"
#include <shlwapi.h>
#include <shlobj.h>
#include <servprov.h>
#include <inspectable.h>
#include <winstring.h>

#pragma comment(lib, "shlwapi.lib")

namespace
{
    // Windows 10 explorer!CTray::_HandlePTPTap / _HandlePTPMouseAction /
    // _HandlePTPKeyboardAction. Taps arrive as Ctrl+Shift+Win+F22/F24 hotkeys,
    // independently of TwinUI's pointer-based swipe recognizer.
    const WCHAR TouchpadSettings[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\PrecisionTouchPad";
    const CLSID ImmersiveShell = { 0xc2f03a33, 0x21f5, 0x47fa, { 0xb4, 0xbb, 0x15, 0x63, 0x62, 0xa2, 0xf2, 0x39 } };

    // Used vtable prefixes from native Explorer's notification-center activation.
    MIDL_INTERFACE("2e8fcb18-a0ee-41ad-8ef8-77fb3a370ca5")
    ITouchpadExperienceFactory : IInspectable
    {
        virtual HRESULT STDMETHODCALLTYPE GetExperience(HSTRING, IInspectable**) = 0;
    };

    MIDL_INTERFACE("df65db57-d504-456e-8bd7-004ce308d8d9")
    ITouchpadActionCenter : IInspectable
    {
        virtual HRESULT STDMETHODCALLTYPE Toggle(IInspectable*) = 0;
    };

    HRESULT InputError()
    {
        DWORD error = GetLastError();
        return error ? HRESULT_FROM_WIN32(error) : E_FAIL;
    }

    HRESULT ReadSetting(LPCWSTR name, DWORD defaultValue, DWORD* value)
    {
        DWORD type = 0, size = sizeof(*value);
        LSTATUS status = SHGetValueW(HKEY_CURRENT_USER, TouchpadSettings, name, &type, value, &size);
        if (status == ERROR_FILE_NOT_FOUND || status == ERROR_PATH_NOT_FOUND)
        {
            *value = defaultValue;
            return S_OK;
        }
        if (status != ERROR_SUCCESS)
            return HRESULT_FROM_WIN32(status);
        return type == REG_DWORD && size == sizeof(*value) ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    HRESULT InjectKeyboard(const KEYBDINPUT* keys, UINT count)
    {
        typedef BOOL (WINAPI* InjectKeyboardInputFn)(const KEYBDINPUT*, UINT);
        static const auto inject = reinterpret_cast<InjectKeyboardInputFn>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "InjectKeyboardInput"));
        SetLastError(ERROR_SUCCESS);
        if (inject)
            return inject(keys, count) ? S_OK : InputError();

        // Older systems do not export the shell's injection APIs.
        INPUT inputs[16] = {};
        for (UINT i = 0; i < count; ++i)
        {
            inputs[i].type = INPUT_KEYBOARD;
            inputs[i].ki = keys[i];
            if ((keys[i].wScan & 0xff00) == 0xe000)
                inputs[i].ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
            // Navigation VKs share scan codes with the numeric keypad;
            // MapVirtualKeyEx need not return an E0 prefix for these VKs.
            switch (keys[i].wVk)
            {
            case VK_PRIOR: case VK_NEXT: case VK_END: case VK_HOME:
            case VK_LEFT: case VK_UP: case VK_RIGHT: case VK_DOWN:
            case VK_INSERT: case VK_DELETE:
                inputs[i].ki.dwFlags |= KEYEVENTF_EXTENDEDKEY;
                break;
            }
        }
        return SendInput(count, inputs, sizeof(INPUT)) == count ? S_OK : InputError();
    }

    HRESULT MouseAction(DWORD action)
    {
        MOUSEINPUT mouse[2] = {};
        if (action == 4)
        {
            mouse[0].dwFlags = MOUSEEVENTF_MIDDLEDOWN;
            mouse[1].dwFlags = MOUSEEVENTF_MIDDLEUP;
        }
        else
        {
            mouse[0].dwFlags = MOUSEEVENTF_XDOWN;
            mouse[1].dwFlags = MOUSEEVENTF_XUP;
            mouse[0].mouseData = mouse[1].mouseData = action == 5 ? XBUTTON1 : XBUTTON2;
        }

        typedef BOOL (WINAPI* InjectMouseInputFn)(const MOUSEINPUT*, UINT);
        static const auto inject = reinterpret_cast<InjectMouseInputFn>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "InjectMouseInput"));
        SetLastError(ERROR_SUCCESS);
        if (inject)
            return inject(mouse, ARRAYSIZE(mouse)) ? S_OK : InputError();

        INPUT inputs[2] = {};
        for (UINT i = 0; i < ARRAYSIZE(inputs); ++i)
        {
            inputs[i].type = INPUT_MOUSE;
            inputs[i].mi = mouse[i];
        }
        return SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT)) == ARRAYSIZE(inputs) ? S_OK : InputError();
    }

    HRESULT KeyboardAction(DWORD parameters, bool fourFinger)
    {
        KEYBDINPUT keys[16] = {};
        UINT count = 0;
        auto append = [&](WORD vk, DWORD flags, UINT mapType)
        {
            KEYBDINPUT& key = keys[count++];
            key.wVk = vk;
            key.wScan = static_cast<WORD>(MapVirtualKeyExW(vk, mapType, nullptr));
            key.dwFlags = flags;
        };

        // Native Explorer releases the synthetic gesture chord before injecting
        // the requested shortcut; otherwise Ctrl/Shift/Win leak into the action.
        append(fourFinger ? VK_F24 : VK_F22, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC);
        append(VK_SHIFT, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC);
        append(VK_CONTROL, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC);
        append(VK_LWIN, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC);

        // PrecisionTouchPad KeyParams is NOT a RegisterHotKey modifier mask:
        // low word = Ctrl, Shift, left Win, right Win, Alt; high word = VK.
        const WORD modifiers[] = { VK_CONTROL, VK_SHIFT, VK_LWIN, VK_RWIN, VK_MENU };
        for (UINT i = 0; i < ARRAYSIZE(modifiers); ++i)
            if (LOWORD(parameters) & (1u << i))
                append(modifiers[i], 0, MAPVK_VK_TO_VSC);
        WORD vk = HIWORD(parameters);
        if (vk && MapVirtualKeyExW(vk, MAPVK_VK_TO_VSC_EX, nullptr))
        {
            append(vk, 0, MAPVK_VK_TO_VSC_EX);
            append(vk, KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC_EX);
        }
        for (UINT i = ARRAYSIZE(modifiers); i > 0; --i)
            if (LOWORD(parameters) & (1u << (i - 1)))
                append(modifiers[i - 1], KEYEVENTF_KEYUP, MAPVK_VK_TO_VSC);
        return InjectKeyboard(keys, count);
    }

    HRESULT ShellAction(DWORD action)
    {
        if (action == 1)
        {
            if (SHRestricted(REST_NOFIND))
                return S_OK;
            // Native Explorer posts search hotkey 34 through its in-process
            // message service. That interface cannot be marshaled to EXE shells;
            // use Search's activation protocol rather than an unhandled Win+S.
            SHELLEXECUTEINFOW execute = { sizeof(execute) };
            execute.fMask = SEE_MASK_FLAG_NO_UI;
            execute.lpVerb = L"open";
            execute.lpFile = L"ms-search:";
            execute.nShow = SW_SHOWNORMAL;
            if (ShellExecuteExW(&execute))
                return S_OK;
            DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error ? error : ERROR_GEN_FAILURE);
        }

        IServiceProvider* services = nullptr;
        HRESULT hr = CoCreateInstance(ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER | CLSCTX_NO_CODE_DOWNLOAD,
            IID_PPV_ARGS(&services));
        if (FAILED(hr))
            return hr;

        ITouchpadExperienceFactory* factory = nullptr;
        hr = services->QueryService(__uuidof(ITouchpadExperienceFactory), IID_PPV_ARGS(&factory));
        if (SUCCEEDED(hr))
        {
            // Resolve dynamically so merely loading ExplorerEx on pre-Win8
            // systems does not introduce a Windows Runtime import dependency.
            typedef HRESULT (WINAPI* CreateStringReferenceFn)(PCWSTR, UINT32, HSTRING_HEADER*, HSTRING*);
            auto createStringReference = reinterpret_cast<CreateStringReferenceFn>(
                GetProcAddress(GetModuleHandleW(L"combase.dll"), "WindowsCreateStringReference"));
            HSTRING_HEADER header = {};
            HSTRING name = nullptr;
            const WCHAR experienceName[] = L"Windows.Internal.ShellExperience.ActionCenter";
            hr = createStringReference ? createStringReference(experienceName,
                ARRAYSIZE(experienceName) - 1, &header, &name) : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            IInspectable* experience = nullptr;
            if (SUCCEEDED(hr))
                hr = factory->GetExperience(name, &experience);
            if (SUCCEEDED(hr))
            {
                ITouchpadActionCenter* center = nullptr;
                hr = experience->QueryInterface(IID_PPV_ARGS(&center));
                if (SUCCEEDED(hr))
                {
                    hr = center->Toggle(nullptr);
                    center->Release();
                }
                experience->Release();
            }
            factory->Release();
        }
        services->Release();
        return hr;
    }
}

HRESULT HandleTouchpadTap(bool fourFinger)
{
    DWORD action = 0;
    // Defaults match TwinUI's immersive settings table (identifiers 9/10).
    HRESULT hr = ReadSetting(fourFinger ? L"FourFingerTapEnabled" : L"ThreeFingerTapEnabled",
        fourFinger ? 2 : 1, &action);
    if (FAILED(hr))
        return hr;
    bool custom = action == 0xffff;
    if (custom)
    {
        hr = ReadSetting(fourFinger ? L"CustomFourFingerTap" : L"CustomThreeFingerTap",
            fourFinger ? 2 : 1, &action);
        if (FAILED(hr))
            return hr;
    }

    switch (action)
    {
    case 0: // Nothing.
        return S_OK;
    case 1: // Search.
    case 2: // Action Center.
        return ShellAction(action);
    case 3: // Native Explorer falls back to this key when no media session is available.
        return KeyboardAction(MAKELONG(0, VK_MEDIA_PLAY_PAUSE), fourFinger);
    case 4: // Middle mouse button.
    case 5: // Mouse back.
    case 6: // Mouse forward.
        return MouseAction(action);
    case 7: // Recorded custom keyboard shortcut.
        if (custom)
        {
            DWORD parameters = 0;
            hr = ReadSetting(fourFinger ? L"CustomFourFingerTapKeyParams" : L"CustomThreeFingerTapKeyParams",
                0, &parameters);
            if (FAILED(hr))
                return hr;
            return KeyboardAction(parameters, fourFinger);
        }
        return S_OK;
    default: // Unknown settings must not become an unintended mouse/key action.
        return S_OK;
    }
}
