#pragma once

#include <Windows.h>
#include <dpa_dsa.h>

#include "ContainerPolicies.h"

template <typename T, typename ContainerPolicy>
class CDPA_Base
{
public:
    using _PFNDPAENUMCALLBACK = int (CALLBACK *)(T*, void*);
    using _PFNDPAENUMCALLBACKCONST = int (CALLBACK *)(const T*, void*);

    using _PFNDPACOMPARE = int (CALLBACK *)(T*, T*, LPARAM);
    using _PFNDPACOMPARECONST = int (CALLBACK *)(const T*, const T*, LPARAM);

    using _PFNDPAMERGE = T* (CALLBACK*)(UINT, T*, T*, LPARAM);
    using _PFNDPAMERGECONST = const T* (CALLBACK*)(UINT, const T*, const T*, LPARAM);

    CDPA_Base(HDPA hdpa = nullptr) : m_hdpa(hdpa)
    {
    }

    BOOL IsDPASet() const
    {
        return m_hdpa != nullptr;
    }

    void Attach(HDPA hdpa)
    {
        _ASSERT(!m_hdpa);
        m_hdpa = hdpa;
    }

    HDPA Detach()
    {
        HDPA hdpa = m_hdpa;
        m_hdpa = nullptr;
        return hdpa;
    }

    operator HDPA() const
    {
        return m_hdpa;
    }

    BOOL Create(int cItemGrow)
    {
        _ASSERT(nullptr == m_hdpa);
        m_hdpa = DPA_Create(cItemGrow);
        return m_hdpa != nullptr;
    }

    BOOL CreateEx(int cpGrow, HANDLE hheap)
    {
        _ASSERT(nullptr == m_hdpa);
        m_hdpa = DPA_CreateEx(cpGrow, hheap);
        return m_hdpa != nullptr;
    }

    BOOL Destroy()
    {
        BOOL fRet = TRUE;
        if (m_hdpa)
        {
            DestroyCallback(_StandardDestroyCB, nullptr);
            fRet = DPA_Destroy(m_hdpa);
            m_hdpa = nullptr;
        }
        return fRet;
    }

    HDPA Clone(HDPA hdpaNew) const
    {
        return DPA_Clone(m_hdpa, hdpaNew);
    }

    T* GetPtr(INT_PTR i) const
    {
        return (T*)DPA_GetPtr(m_hdpa, i);
    }

    int GetPtrIndex(T* p)
    {
        return DPA_GetPtrIndex(m_hdpa, p);
    }

    BOOL Grow(int cp)
    {
        return DPA_Grow(m_hdpa, cp);
    }

    BOOL SetPtr(int i, T* p)
    {
        return DPA_SetPtr(m_hdpa, i, p);
    }

    HRESULT InsertPtr(int i, T* pitem, int* piIndex = nullptr)
    {
        int iIndex = DPA_InsertPtr(m_hdpa, i, pitem);
        if (piIndex)
            *piIndex = iIndex;
        if (iIndex == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    T* DeletePtr(int i)
    {
        return (T*)DPA_DeletePtr(m_hdpa, i);
    }

    BOOL DeleteAllPtrs()
    {
        return m_hdpa && DPA_DeleteAllPtrs(m_hdpa);
    }

    void EnumCallback(_PFNDPAENUMCALLBACK pfnCB, void* pData = nullptr) const
    {
        DPA_EnumCallback(m_hdpa, (PFNDPAENUMCALLBACK)pfnCB, pData);
    }

    template <class TData>
    void EnumCallbackEx(int (CALLBACK *pfnCB)(T* p, TData pData), TData pData = 0) const
    {
        EnumCallback((_PFNDPAENUMCALLBACK)pfnCB, pData);
    }

    void DestroyCallback(_PFNDPAENUMCALLBACK pfnCB, void* pData = nullptr)
    {
        if (m_hdpa)
        {
            DPA_DestroyCallback(m_hdpa, (PFNDPAENUMCALLBACK)pfnCB, pData);
            m_hdpa = nullptr;
        }
    }

    template <class T2>
    void DestroyCallbackEx(int (CALLBACK *pfnCB)(T* p, T2 pData), T2 pData)
    {
        DestroyCallback((_PFNDPAENUMCALLBACK)pfnCB, reinterpret_cast<void*>(pData));
    }

    int GetPtrCount() const
    {
        return m_hdpa ? DPA_GetPtrCount(m_hdpa) : 0;
    }

    void SetPtrCount(int cItems)
    {
        DPA_SetPtrCount(m_hdpa, cItems);
    }

    T** GetPtrPtr() const
    {
        return (T**)DPA_GetPtrPtr(m_hdpa);
    }

    T*& FastGetPtr(int i) const
    {
        return (T*&)DPA_FastGetPtr(m_hdpa, i);
    }

    HRESULT AppendPtr(T* pitem, int* piIndex = nullptr)
    {
        int iIndex = DPA_AppendPtr(m_hdpa, pitem);
        if (piIndex)
            *piIndex = iIndex;
        if (iIndex == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    BOOL PushPtr(T* p)
    {
        return DPA_AppendPtr(m_hdpa, p);
    }

    T* PopPtr()
    {
        int cPtrs = GetPtrCount();
        return cPtrs > 0 ? DeletePtr(cPtrs - 1) : nullptr;
    }

    ULONGLONG GetSize()
    {
        return DPA_GetSize(m_hdpa);
    }

    HRESULT LoadStream(PFNDPASTREAM pfn, IStream* pstream, void* pvInstData)
    {
        return DPA_LoadStream(&m_hdpa, pfn, pstream, pvInstData);
    }

    HRESULT SaveStream(PFNDPASTREAM pfn, IStream* pstream, void* pvInstData)
    {
        return DPA_SaveStream(m_hdpa, pfn, pstream, pvInstData);
    }

    BOOL Sort(_PFNDPACOMPARE pfnCompare, LPARAM lParam)
    {
        return m_hdpa && DPA_Sort(m_hdpa, (PFNDPACOMPARE)pfnCompare, lParam);
    }

    template <class T2>
    BOOL SortEx(int (CALLBACK *pfnCompare)(T* p1, T* p2, T2 lParam), T2 lParam)
    {
        return Sort((_PFNDPACOMPARE)pfnCompare, reinterpret_cast<LPARAM>(lParam));
    }

    BOOL Merge(CDPA_Base* pdpaDest, DWORD dwFlags, _PFNDPACOMPARE pfnCompare, _PFNDPAMERGE pfnMerge, LPARAM lParam)
    {
        return DPA_Merge(m_hdpa, pdpaDest->m_hdpa, dwFlags, (PFNDPACOMPARE)pfnCompare, (PFNDPAMERGE)pfnMerge, lParam);
    }

    int Search(T* pFind, int iStart, _PFNDPACOMPARE pfnCompare, LPARAM lParam, UINT options) const
    {
        return DPA_Search(m_hdpa, pFind, iStart, (PFNDPACOMPARE)pfnCompare, lParam, options);
    }

    BOOL SortedInsertPtr(T* pItem, int iStart, _PFNDPACOMPARE pfnCompare, LPARAM lParam, UINT options, T* pFind)
    {
        return DPA_SortedInsertPtr(m_hdpa, pFind, iStart, (PFNDPACOMPARE)pfnCompare, lParam, options, pItem);
    }

    ~CDPA_Base()
    {
        if (m_hdpa)
            Destroy();
    }

protected:
    HDPA m_hdpa;

private:
    static int CALLBACK _StandardDestroyCB(T* p, void* pData)
    {
        ContainerPolicy::Destroy(p);
        return 1;
    }
};

template <typename T, typename ContainerPolicy = CTContainer_PolicyUnOwned<T>>
class CDPA : public CDPA_Base<T, ContainerPolicy>
{
public:
    CDPA(HDPA hdpa = nullptr)
        : CDPA_Base<T, ContainerPolicy>(hdpa)
    {
    }
};

template <typename T, typename ContainerPolicy = CTContainer_PolicyCoTaskMem>
class CDPACoTaskMem : public CDPA<T, ContainerPolicy>
{
};

template <typename T, typename ContainerPolicy = CTContainer_PolicyLocalMem>
class CDPALocalMem : public CDPA<T, ContainerPolicy>
{
};

template <typename T, typename ContainerPolicy = CTContainer_PolicyNewMem>
class CDPANewMem : public CDPA<T, ContainerPolicy>
{
};

template <typename T, typename ContainerPolicy = CTContainer_PolicyRelease<T>>
class CDPARelease : public CDPA<T, ContainerPolicy>
{
};

template <typename T>
class CDSA_Base
{
public:
    using _PFNDSAENUMCALLBACK = int (CALLBACK *)(T*, void*);
    using _PFNDSAENUMCALLBACKCONST = int (CALLBACK *)(const T*, void*);
    using _PFNDSACOMPARE = int (CALLBACK *)(const T*, const T*, LPARAM);

    CDSA_Base(const CDSA_Base& other) = delete;

    CDSA_Base(HDSA hdsa = nullptr) : m_hdsa(hdsa)
    {
    }

    /*~CDSA_Base()
    {
        Destroy();
    }*/

    void Attach(HDSA hdsa)
    {
        if (m_hdsa)
            Destroy();
        m_hdsa = hdsa;
    }

    HDSA Detach()
    {
        HDSA hdsa = m_hdsa;
        m_hdsa = nullptr;
        return hdsa;
    }

    operator HDSA() const
    {
        return m_hdsa;
    }

    BOOL Create(int cItemGrow)
    {
        _ASSERT(nullptr == m_hdsa);
        m_hdsa = DSA_Create(sizeof(T), cItemGrow);
        return m_hdsa != nullptr;
    }

    BOOL Destroy()
    {
        BOOL result = TRUE;
        if (m_hdsa)
        {
            result = DSA_Destroy(m_hdsa);
            m_hdsa = nullptr;
        }
        return result;
    }

    BOOL GetItem(int i, void* pItem) const
    {
        return DSA_GetItem(m_hdsa, i, pItem);
    }

    T* GetItemPtr(int i) const
    {
        return (T*)DSA_GetItemPtr(m_hdsa, i);
    }

    HRESULT InsertItem(int i, T* pitem, int* piIndex = nullptr)
    {
        int iIndex = DSA_InsertItem(m_hdsa, i, pitem);
        if (piIndex)
            *piIndex = iIndex;
        if (iIndex == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    BOOL DeleteItem(int i) const
    {
        return DSA_DeleteItem(m_hdsa, i);
    }

    void DestroyCallback(_PFNDSAENUMCALLBACK pfnCB, void* pData = nullptr)
    {
        if (m_hdsa)
        {
            DSA_DestroyCallback(m_hdsa, (PFNDSAENUMCALLBACK)pfnCB, pData);
            m_hdsa = nullptr;
        }
    }

    template <class T2>
    void DestroyCallbackEx(int (CALLBACK *pfnCB)(T* p, T2 pData), T2 pData)
    {
        DestroyCallback((_PFNDSAENUMCALLBACK)pfnCB, reinterpret_cast<void*>(pData));
    }

    int GetItemCount() const
    {
        return m_hdsa ? DSA_GetItemCount(m_hdsa) : 0;
    }

    HRESULT AppendItem(const T* pitem, int* piIndex = nullptr)
    {
        int iIndex = DSA_AppendItem(m_hdsa, pitem);
        if (piIndex)
            *piIndex = iIndex;
        if (iIndex == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    HRESULT Search(const T* pFind, int iStart, _PFNDSACOMPARE pfnCompare, LPARAM lParam, UINT options, int* piIndex)
    {
        int iIndex = Search(pFind, iStart, pfnCompare, lParam, options);
        if (piIndex)
            *piIndex = iIndex;
        return iIndex != -1 ? S_OK : E_FAIL;
    }

    int Search(const T* pFind, int iStart, _PFNDSACOMPARE pfnCompare, LPARAM lParam, UINT options)
    {
        int cItem = GetItemCount();

        _ASSERT(pfnCompare);
        _ASSERT(0 <= iStart);
        _ASSERT((options & DPAS_SORTED) || !(options & (DPAS_INSERTBEFORE | DPAS_INSERTAFTER)));

        if ((options & DPAS_SORTED) != 0)
        {
            int iRet = -1;
            int nCmp = 0;

            int iLow = iStart;
            int iMid = 0;
            int iHigh = cItem - 1;
            while (true)
            {
                if (iLow > iHigh)
                {
                    if ((options & (DPAS_INSERTBEFORE | DPAS_INSERTAFTER)) != 0)
                    {
                        iRet = nCmp > 0 ? iLow : iMid;
                    }
                    break;
                }

                iMid = (iLow + iHigh) / 2;
                nCmp = pfnCompare(pFind, GetItemPtr(iMid), lParam);
                if (nCmp < 0)
                {
                    iHigh = iMid - 1;
                }
                else if (nCmp > 0)
                {
                    iLow = iMid + 1;
                }
                else
                {
                    for (; iMid > iStart; --iMid)
                    {
                        if (pfnCompare(pFind, GetItemPtr(iMid - 1), lParam) != 0)
                        {
                            break;
                        }
                    }
                    _ASSERT(0 <= iMid);
                    iRet = iMid;
                    break;
                }
            }

            if ((options & (DPAS_INSERTBEFORE | DPAS_INSERTAFTER)) == 0)
            {
                _ASSERT(Search(pFind, iStart, pfnCompare, lParam, options & ~DPAS_SORTED) == iRet);
            }

            return iRet;
        }
        else
        {
            for (int i = iStart; i < cItem; ++i)
            {
                if (pfnCompare(pFind, GetItemPtr(i), lParam) == 0)
                {
                    return i;
                }
            }

            return -1;
        }
    }

protected:
    HDSA m_hdsa;
};

template <typename T>
class CDSA : public CDSA_Base<T>
{
};
