#pragma once

inline PCUITEMID_CHILD _SHILMakeChild(const void* pv)
{
    PCUITEMID_CHILD pidl = static_cast<PCITEMID_CHILD>(pv);
    //RIP(ILIsChild(reinterpret_cast<PCUIDLIST_RELATIVE>(pidl))); // 178
    return pidl;
}

inline PCIDLIST_ABSOLUTE _SHILMakeFull(const void* pv)
{
    PCIDLIST_ABSOLUTE pidl = static_cast<PCIDLIST_ABSOLUTE>(pv);
    // RIP(ILIsAligned(reinterpret_cast<PCUIDLIST_RELATIVE>(pidl))); // 183
    return pidl;
}
