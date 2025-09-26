#include "pch.h"
#include "stdafx.h"
#include "sfthost.h"
#include "proglist.h"

BOOL UserPane_RegisterClass();
BOOL MorePrograms_RegisterClass();
BOOL LogoffPane_RegisterClass();

void RegisterDesktopControlClasses()
{
    SFTBarHost::Register();
    UserPane_RegisterClass();
    MorePrograms_RegisterClass();
    LogoffPane_RegisterClass();
    //NSCHost_RegisterClass();
    //OpenViewHost_RegisterClass();
    //SearchView_RegisterClass();
    //TopMatch_RegisterClass();
    //UserPicture_RegisterClass();
}
