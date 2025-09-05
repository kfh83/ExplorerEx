```
  ______            _                     ______      
 |  ____|          | |                   |  ____|     
 | |__  __  ___ __ | | ___  _ __ ___ _ __| |__  __  __
 |  __| \ \/ / '_ \| |/ _ \| '__/ _ \ '__|  __| \ \/ /
 | |____ >  <| |_) | | (_) | | |  __/ |  | |____ >  < 
 |______/_/\_\ .__/|_|\___/|_|  \___|_|  |______/_/\_\
             | |                                      
             |_|    
          https://github.com/kfh83/ExplorerEx
```
ExplorerEx is an adaptation of Windows Explorer from Windows Server 2003's source tree, making it independent from Microsoft's archaic build system, and also making it function on modern versions of Windows. 

In other words, this brings the classic Windows XP Start Menu/Taskbar experience to more modern versions of Windows (8.1+).

## Previews
Below is what you might expect when first running ExplorerEx:

<img width="800" height="600" alt="aero" src="https://github.com/user-attachments/assets/6bca7f97-e019-40cf-b06a-a0d6a30e5f1e" />

*Default Windows 10 theme*

<img width="800" height="600" alt="classic" src="https://github.com/user-attachments/assets/72f5c240-8a7d-401f-977c-276d51d9824d" />

*Classic theme using Windhawk*

<img width="800" height="600" alt="luna" src="https://github.com/user-attachments/assets/1530f4fc-9249-4494-9ab0-2d0aef452687" />

*Luna blue theme using [inactive theme Windhawk mod](https://raw.githubusercontent.com/kfh83/ExplorerEx/refs/heads/main/explorerex-inactive-theme-loader.wh.cpp)*


## Installation
**TRYING OUT THE PROGRAM BEFORE INSTALLING**

If you have download one of the releases, apply "Import Me!.reg" by double clicking it, and then simply run "Try ExplorerEx.bat", which will terminate your current explorer session (if you
have important folders open, please do keep in mind they'll be closed upon running this script) and then it'll start
ExplorerEx.

To go back to your default shell, just open task manager (right click the taskbar and press the Task Manager
option), find the "Windows Explorer" task (explorer.exe process for those still in the Windows 7-style task manager),
and then end the task. Next, go to the "File" menu up top, press "Run new task" and type "explorer.exe". This should start
your regular explorer shell.

**INSTALLING THE PROGRAM**

If you have fetched a GitHub actions artifact, you *need* to fetch the release package for the "Import Me!.reg" file which includes some important tweaks. Please apply this .reg file before proceeding.

<details>
  <summary><b>Release/.exe build</b></summary>

**How-to**
1. Copy over `explorer.exe` and your desired language folder `i.e, en-US` to a suitable location (i.e, `%SystemDrive%\ExplorerEx`). This cannot be `%SystemRoot%`.
2. Open the Registry Editor and go to `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\` and change the `Shell` key to the path of where you copied explorer.exe to (i.e `C:\ExplorerEx\explorer.exe`).
3. **OPTIONAL:** If you wish to set the shell on a per-user basis instead of machine-wide you can do the same as above on `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\`. You'll need to create the `Shell` key.
4. Run Task Manager, restart Windows Explorer, it should automatically start ExplorerEx.

  **You're all set!**
</details>

<details>
  <summary>.dll build</summary>

**Note:** You MUST have [Windhawk](https://windhawk.net/) installed for the loader mod.
In addition, you **MUST** make sure Windhawk can inject into winlogon.exe. You can either go into `Settings\Advanced Settings\More Advanced Settings` and add `winlogon.exe` into the Process inclusion list, or you can untick "Exclude critical system processes".
  
**How-to**
1. Copy over `ExplorerEx.dll` and your desired language folder `i.e, en-US` to a suitable location (i.e, `%SystemDrive%\ExplorerEx`).
2. Copy the contents of the [loader mod](https://raw.githubusercontent.com/kfh83/ExplorerEx/refs/heads/main/explorerex-dll-loader.wh.cpp) from the repo.
3. Open Windhawk, hit "Create a New Mod" button on the bottom right corner.
4. Select all of the existing content and delete it, and paste in the mod code you copied earlier.
5. Hit "Compile Mod" button on the left sidebar, and once it's done compiling hit "Exit Editing Mode"
6. Find the "ExplorerEx DLL Loader" mod in the home section, click "Details" button
7. Switch to Settings tab, you'll need to input the path of "ExplorerEx.dll" on the mod settings (i.e C:\ExplorerEx\ExplorerEx.dll), and then click "Save Settings"
8. Run Task Manager, restart Windows Explorer, it should automatically start ExplorerEx.

  **You're all set!**
</details>

## Important notes
- Regular release builds (.exe) **don't and won't** support UWP apps (Microsoft Store apps, PC settings...). You must use a .dll binary for it. Currently dll builds aren't provided in the release packages, so you must build it yourself.
- Currently, system tray icons (volume, network...) aren't working. We have *started* work on a [parallel project](https://github.com/kfh83/StobjectEx) to bring over the XP tray icons in a similar fashion to this project. BEWARE!
- To configure your startup programs, you **must** use msconfig.exe from Windows 7 or earlier. This is due to a difference in how Task Manager sets startup apps. You can find an easy-to-install package [here](https://www.majorgeeks.com/files/details/classic_msconfig.html).
- Currently, you must use the "Eradicate immersive menus" Windhawk mod by aubymori to see the taskbar context menu properly.

**PLEASE HEAD OVER TO THE ISSUES SECTION OF THIS REPOSITORY TO SEE MORE KNOWN ISSUES.**

## Building
1. Clone the repository using git:
```git clone https://github.com/kfh83/ExplorerEx.git```

2. Open `ExplorerEx.sln` in your preferred Visual Studio-compatible IDE. Note building has only been tested using Visual Studio 2022.

3. Select your configuration and build away:

- `Release` builds `explorer.exe` with **no** UWP app support.
- `ReleaseDLL` builds `ExplorerEx.dll` with UWP app support enabled for use with the loader Windhawk mod provided in the repository.
- `ReleaseDebug` builds `explorer.exe` with **no** UWP app support, optimizations disabled and console output enabled.

> [!NOTE]
> `Debug` configuration is currently (as of Sep 5, 2025) broken. Please use the `ReleaseDebug` configuration instead.

## Distribution
For distribution purposes, please include credit in your project by directly linking to this repository if possible. Failing to do so will result in bad luck for the rest of your life.
