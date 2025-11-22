#pragma once

#include <Windows.h>
#include <dpa_dsa.h>

#include "ContainerPolicies.h"

template<typename T, typename ContainerPolicy>
class CDPA_Base
{
public:
    using EnumCallbackType = int (CALLBACK *)(T *, void *);
    using CompareType = int (CALLBACK *)(T *p1, T *p2, LPARAM lParam);
    using MergeType = T * (CALLBACK *)(UINT uMsg, T *pDest, T *pSrc, LPARAM lParam);

    CDPA_Base(HDPA hdpa = nullptr) : m_hdpa(hdpa)
    {
    }

    ~CDPA_Base()
    {
        if (m_hdpa)
            Destroy();
    }

    BOOL IsDPASet() const
    {
        return m_hdpa != nullptr;
    }

    void Attach(HDPA hdpa)
    {
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
        m_hdpa = DPA_Create(cItemGrow);
        return m_hdpa != nullptr;
    }

    BOOL CreateEx(int cpGrow, HANDLE hheap)
    {
        m_hdpa = DPA_CreateEx(cpGrow, hheap);
        return m_hdpa != nullptr;
    }

    BOOL Destroy()
    {
        BOOL result = TRUE;
        if (m_hdpa)
        {
            DestroyCallback(_StandardDestroyCB, nullptr);
            result = DPA_Destroy(m_hdpa);
            m_hdpa = nullptr;
        }
        return result;
    }

    HDPA Clone(HDPA hdpaNew) const
    {
        return DPA_Clone(m_hdpa, hdpaNew);
    }

    T *GetPtr(INT_PTR i) const
    {
        return (T *)DPA_GetPtr(m_hdpa, i);
    }

    int GetPtrIndex(T *p)
    {
        return DPA_GetPtrIndex(m_hdpa, p);
    }

    BOOL Grow(int cp)
    {
        return DPA_Grow(m_hdpa, cp);
    }

    BOOL SetPtr(int i, T *p)
    {
        return DPA_SetPtr(m_hdpa, i, p);
    }

    HRESULT InsertPtr(int i, T *p, int *outIndex = nullptr)
    {
        int result = DPA_InsertPtr(m_hdpa, i, p);
        if (outIndex)
            *outIndex = result;
        if (result == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    T *DeletePtr(int i)
    {
        return (T *)DPA_DeletePtr(m_hdpa, i);
    }

    BOOL DeleteAllPtrs()
    {
        return DPA_DeleteAllPtrs(m_hdpa);
    }

    void EnumCallback(EnumCallbackType pfnCB, void *pData = nullptr)
    {
        DPA_EnumCallback(m_hdpa, (PFNDPAENUMCALLBACK)pfnCB, pData);
    }

    // EXEX-Vista(allison): TODO. Check if this still actually exists in the Vista+ version of CDPA.
    template<class T2>
    void EnumCallbackEx(int (CALLBACK *pfnCB)(T *p, T2 pData), T2 pData)
    {
        EnumCallback((EnumCallbackType)pfnCB, reinterpret_cast<void *>(pData));
    }

    void DestroyCallback(EnumCallbackType pfnCB, void *pData = nullptr)
    {
        if (m_hdpa)
        {
            DPA_DestroyCallback(m_hdpa, (PFNDPAENUMCALLBACK)pfnCB, pData);
            m_hdpa = nullptr;
        }
    }

    // EXEX-Vista(allison): TODO. Check if this still actually exists in the Vista+ version of CDPA.
    template<class T2>
    void DestroyCallbackEx(int (CALLBACK *pfnCB)(T *p, T2 pData), T2 pData)
    {
        DestroyCallback((EnumCallbackType)pfnCB, reinterpret_cast<void *>(pData));
    }

    int GetPtrCount() const
    {
        return m_hdpa ? DPA_GetPtrCount(m_hdpa) : 0;
    }

    void SetPtrCount(int cItems)
    {
        DPA_SetPtrCount(m_hdpa, cItems);
    }

    T **GetPtrPtr() const
    {
        return (T **)DPA_GetPtrPtr(m_hdpa);
    }

    T *&FastGetPtr(int i) const
    {
        return (T *&)DPA_FastGetPtr(m_hdpa, i);
    }

    HRESULT AppendPtr(T *p, int *outIndex = nullptr)
    {
        int result = DPA_AppendPtr(m_hdpa, p);
        if (outIndex)
            *outIndex = result;
        if (result == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    ULONGLONG GetSize()
    {
        return DPA_GetSize(m_hdpa);
    }

    HRESULT LoadStream(PFNDPASTREAM pfn, IStream *pstream, void *pvInstData)
    {
        return DPA_LoadStream(&m_hdpa, pfn, pstream, pvInstData);
    }

    HRESULT SaveStream(PFNDPASTREAM pfn, IStream *pstream, void *pvInstData)
    {
        return DPA_SaveStream(m_hdpa, pfn, pstream, pvInstData);
    }

    BOOL Sort(CompareType pfnCompare, LPARAM lParam)
    {
        return DPA_Sort(m_hdpa, (PFNDACOMPARE)pfnCompare, lParam);
    }

    // EXEX-Vista(allison): TODO. Check if this still actually exists in the Vista+ version of CDPA.
    template<class T2>
    BOOL SortEx(int (CALLBACK *pfnCompare)(T *p1, T *p2, T2 lParam), T2 lParam)
    {
        return Sort((CompareType)pfnCompare, reinterpret_cast<LPARAM>(lParam));
    }

    BOOL Merge(CDPA_Base *pdpaDest, DWORD dwFlags, CompareType pfnCompare, MergeType pfnMerge, LPARAM lParam)
    {
        return DPA_Merge(m_hdpa, pdpaDest->m_hdpa, dwFlags, (PFNDACOMPARE)pfnCompare, (PFNDPAMERGE)pfnMerge, lParam);
    }

    int Search(T *pFind, int iStart, CompareType pfnCompare, LPARAM lParam, UINT options)
    {
        return DPA_Search(m_hdpa, pFind, iStart, (PFNDACOMPARE)pfnCompare, lParam, options);
    }

    BOOL SortedInsertPtr(T *pItem, int iStart, CompareType pfnCompare, LPARAM lParam, UINT options, T *pFind)
    {
        return DPA_SortedInsertPtr(m_hdpa, pFind, iStart, (PFNDACOMPARE)pfnCompare, lParam, options, pItem);
    }

private:
    static int CALLBACK _StandardDestroyCB(T *p, void *pData)
    {
        ContainerPolicy::Destroy(p);
        return 1;
    }

    HDPA m_hdpa;
};

template<typename T, typename ContainerPolicy>
class CDPA : public CDPA_Base<T, ContainerPolicy>
{
public:
    CDPA(HDPA hdpa = nullptr)
        : CDPA_Base<T, ContainerPolicy>(hdpa)
    {
    }
};

template<typename T>
class CDSA_Base
{
public:
    using EnumCallbackType = int (CALLBACK *)(T *, void *);

    CDSA_Base(const CDSA_Base &other) = delete;

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

    BOOL GetItem(int i, T *pitem)
    {
        return DSA_GetItem(m_hdsa, i, (void *)pitem);
    }

    T *GetItemPtr(int i) const
    {
        return (T *)DSA_GetItemPtr(m_hdsa, i);
    }

    HRESULT InsertItem(int i, T *p, int *outIndex = nullptr)
    {
        int result = DSA_InsertItem(m_hdsa, i, p);
        if (outIndex)
            *outIndex = result;
        if (result == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    int DeleteItem(int i)
    {
        return DSA_DeleteItem(m_hdsa, i);
	}

    void DestroyCallback(EnumCallbackType pfnCB, void *pData = nullptr)
    {
        if (m_hdsa)
        {
            DSA_DestroyCallback(m_hdsa, (PFNDSAENUMCALLBACK)pfnCB, pData);
            m_hdsa = nullptr;
        }
    }

    template<class T2>
    void DestroyCallbackEx(int (CALLBACK *pfnCB)(T *p, T2 pData), T2 pData)
    {
        DestroyCallback((EnumCallbackType)pfnCB, reinterpret_cast<void *>(pData));
    }

    int GetItemCount() const
    {
        return m_hdsa ? DSA_GetItemCount(m_hdsa) : 0;
    }

    HRESULT AppendItem(T *p, int *outIndex = nullptr)
    {
        int result = DSA_AppendItem(m_hdsa, p);
        if (outIndex)
            *outIndex = result;
        if (result == -1)
            return E_OUTOFMEMORY;
        return S_OK;
    }

    HRESULT Search(const T *pFind, int iStart, int (*pfnCompare)(const T *, const T *, LPARAM), LPARAM lParam, UINT options, int *outIndex)
    {
        int index = Search(pFind, iStart, pfnCompare, lParam, options);
        if (outIndex)
            *outIndex = index;
        return index != -1 ? S_OK : E_FAIL;
    }

    int Search(const T *pFind, int iStart, int (*pfnCompare)(const T *, const T *, LPARAM), LPARAM lParam, UINT options)
    {
        int cItem = GetItemCount();

        if ((options & DPAS_SORTED) != 0)
        {
            int iCompare = 0;

            int left = iStart;
            int mid = 0;
            int right = cItem - 1;
            while (left <= right)
            {
                mid = (left + right) / 2;
                iCompare = pfnCompare(pFind, GetItemPtr(mid), lParam);
                if (iCompare < 0)
                {
                    right = mid - 1;
                }
                else if (iCompare > 0)
                {
                    left = mid + 1;
                }
                else
                {
                    for (; mid > iStart; --mid)
                    {
                        if (pfnCompare(pFind, GetItemPtr(mid - 1), lParam) != 0)
                            break;
                    }
                    return mid;
                }
            }

            if ((options & (DPAS_INSERTBEFORE | DPAS_INSERTAFTER)) != 0)
                return iCompare > 0 ? left : mid;
        }
        else
        {
            for (int i = iStart; i < cItem; ++i)
            {
                if (pfnCompare(pFind, GetItemPtr(i), lParam) == 0)
                    return i;
            }
        }

        return -1;
    }

protected:
    HDSA m_hdsa;
};

template<typename T>
class CDSA : public CDSA_Base<T>
{
};