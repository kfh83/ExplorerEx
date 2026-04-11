#ifndef _SHSRVOBJ_H
#define _SHSRVOBJ_H

#include "pch.h"
#include "dpa.h"


// Cabinet_EnumRegApps flags 
#define RRA_DEFAULT               0x0000
#define RRA_DELETE                0x0001        // delete each reg value when we're done with it
#define RRA_WAIT                  0x0002        // Wait for current item to finish before launching next item
// was RRA_SHELLSERVICEOBJECTS    0x0004 -- do not reuse
#define RRA_NOUI                  0x0008        // prevents ShellExecuteEx from displaying error dialogs
#if (_WIN32_WINNT >= 0x0500)
#define RRA_USEJOBOBJECTS         0x0020        // wait on job objects instead of process handles
#endif

typedef UINT RRA_FLAGS;

//
// class to manage shell service objects
//

typedef struct
{
    CLSID              clsid;
    IOleCommandTarget* pct;
}
SHELLSERVICEOBJECT, *PSHELLSERVICEOBJECT;


class CShellServiceObjectMgr
{
public:
    HRESULT Init();
    void Destroy();
    HRESULT LoadRegObjects();
    HRESULT EnableObject(const CLSID *pclsid, DWORD dwFlags);

    virtual ~CShellServiceObjectMgr();

private:
    static int WINAPI DestroyItemCB(SHELLSERVICEOBJECT *psso, CShellServiceObjectMgr *pssomgr);
    HRESULT _LoadObject(REFCLSID rclsid, DWORD dwFlags);
    int _FindItemByCLSID(REFCLSID rclsid);

    static BOOL WINAPI EnumRegAppProc(LPCTSTR pszSubkey, LPCTSTR pszCmdLine, RRA_FLAGS fFlags, LPARAM lParam);

    CDSA<SHELLSERVICEOBJECT> _dsaSSO;
};

#endif  // _SHSRVOBJ_H
