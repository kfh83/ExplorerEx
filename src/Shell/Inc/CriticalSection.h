//+-------------------------------------------------------------------------
//
//	ExplorerEx - Windows XP Explorer
//	Copyright (C) Microsoft
// 
//	File:       criticalsection.h
// 
//	Description:	Reimplements some of the critical section functions
//					used throughout explorer.
// 
//	History:    Jan-22-25   kfh83  Created
//
//+-------------------------------------------------------------------------

#pragma once

#include "pch.h"
#include "debug.h"

#define ENTERCRITICAL   Shell_EnterCriticalSection();
#define LEAVECRITICAL   Shell_LeaveCriticalSection();

//
//  Structs
//

typedef struct _nonreentrantcriticalsection
{
    CRITICAL_SECTION critsec;

    BOOL bEntered;
} NONREENTRANTCRITICALSECTION;

//
//  Function Definitions
//

/* critical section used for access serialization */

static NONREENTRANTCRITICALSECTION s_nrcs =
{
   { 0 },
   FALSE
};


static BOOL EnterNonReentrantCriticalSection(
    NONREENTRANTCRITICALSECTION pnrcs)
{
    BOOL bEntered;

    EnterCriticalSection(&(pnrcs.critsec));

    bEntered = (!pnrcs.bEntered);

    if (bEntered)
    {
        pnrcs.bEntered = TRUE;

    }
    else
    {
        LeaveCriticalSection(&(pnrcs.critsec));

        //ERROR_OUT(("EnterNonReentrantCriticalSection(): Thread %#lx attempted to reenter non-reentrant code.",
            //GetCurrentThreadId()));
    }

    return(bEntered);
}


static void LeaveNonReentrantCriticalSection(
    NONREENTRANTCRITICALSECTION pnrcs)
{
    //ASSERT(IS_VALID_STRUCT_PTR(pnrcs, CNONREENTRANTCRITICALSECTION));

    if (EVAL(pnrcs.bEntered))
    {
        pnrcs.bEntered = FALSE;
        LeaveCriticalSection(&(pnrcs.critsec));
    }

    return;
}

static BOOL BeginExclusiveAccess(void)
{
    return(EnterNonReentrantCriticalSection(s_nrcs));
}


static void EndExclusiveAccess(void)
{
    LeaveNonReentrantCriticalSection(s_nrcs);

    return;
}

static void Shell_EnterCriticalSection(void)
{
    //BeginExclusiveAccess();

    return;
}

static void Shell_LeaveCriticalSection(void)
{
    //EndExclusiveAccess();

    return;
}