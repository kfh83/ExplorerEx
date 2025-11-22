#pragma once

#define STRSAFE_NO_CB_FUNCTIONS
#define STRSAFE_NO_DEPRECATE

//#define STRICT_TYPED_ITEMIDS

// C/C++ headers
#include <malloc.h>
#include <iostream>
#include <strsafe.h>
#include <stdio.h>
#include <vector>

// Windows headers
#include <Windows.h>
#include <winuserp.h>
#include <shlwapi.h>
#include <DocObj.h>
#include <shobjidl_core.h>
#include <shlobj.h>
#include <RegStr.h>
#include <shlguid.h>
#include <synchapi.h>
#include <pshpack4.h>
#include <poppack.h>
#include <userenv.h>
#include <dpa_dsa.h>
#include <vssym32.h>
#include <commctrl.h>
#include <shfusion.h>
#include <windowsx.h>
#include <ole2.h>
#include <oleacc.h>
#include <wininet.h>
#include <uxtheme.h>
#include <cfgmgr32.h>
#include <exdisp.h>
#include <mshtml.h>
#include <shtypes.h>
#include <wmistr.h>
#include <ntwmi.h>
#include <evntrace.h>
#include <wtypes.h>
#include <initguid.h>
#include <lmcons.h>
#include <combaseapi.h>
#include <Dbghelp.h>
#include <dsrole.h>
#include <sddl.h>
#include <dbt.h>
#include <vsstyle.h>
#include <fileapi.h>
#include <raserror.h>
#include <netcon.h>
#include <ras.h>
#include <dwmapi.h>
#include <wtsapi32.h>
#include <psapi.h>
#include <ddraw.h>


// ATL headers
#include <cguid.h>
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <atlstuff.h>


// ExExplorer headers
#include "dpa.h"
#include "debug.h"
#include "ieguidp.h"
#include "patternhelper.h"
#include "rcids.h"
#include "startids.h"
#include "shdguid.h"
#include "port32.h"
