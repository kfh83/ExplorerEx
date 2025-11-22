#include "pch.h"
#include "stdafx.h"
#include "sfthost.h"
#include "proglist.h"

BOOL UserPane_RegisterClass();
BOOL MorePrograms_RegisterClass();
BOOL LogoffPane_RegisterClass();

BOOL NSCHost_RegisterClass();
BOOL OpenViewHost_RegisterClass();
BOOL OpenBoxHost_RegisterClass();
BOOL SearchView_RegisterClass();
BOOL TopMatch_RegisterClass();
BOOL UserPicture_RegisterClass();

void RegisterDesktopControlClasses()
{
    SFTBarHost::Register();
    UserPane_RegisterClass();
    MorePrograms_RegisterClass();
    LogoffPane_RegisterClass();
    NSCHost_RegisterClass();
    OpenViewHost_RegisterClass();
    OpenBoxHost_RegisterClass();
    SearchView_RegisterClass();
    TopMatch_RegisterClass();
    UserPicture_RegisterClass();
}
