// arrakeener.h
#pragma once

#include "arrakis_h.h"
#include <string>

class CArrakeener : public IArrakeener, public ISupportErrorInfo
{
    LONG m_rc;                          // Reference count
    ITypeInfo* m_pti;                   // Pointer to type information

    CRITICAL_SECTION m_cs;              // Protect access to object
    void Lock() noexcept;
    void Unlock() noexcept;

    // Object data
    std::wstring m_first_name;
    std::wstring m_last_name;
    std::wstring m_affiliation;
    std::wstring m_occupation;
    LONGLONG m_energy;
    LONGLONG m_solaris;
    LONGLONG m_spice;

    void SetError(UINT id) noexcept;    // SetErrorInfo wrapper

public:
    CArrakeener();
    virtual ~CArrakeener() noexcept;

    // IUnknown methods
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IDispatch methods
    STDMETHODIMP GetTypeInfoCount(UINT* pctinfo) override;
    STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
    STDMETHODIMP GetIDsOfNames(
        REFIID riid,
        LPOLESTR* rgszNames,
        UINT cNames,
        LCID lcid,
        DISPID* rgDispId) override;
    STDMETHODIMP Invoke(
        DISPID dispIdMember,
        REFIID riid,
        LCID lcid,
        WORD wFlags,
        DISPPARAMS* pDispParams,
        VARIANT* pVarResult,
        EXCEPINFO* pExcepInfo,
        UINT* puArgErr) override;

    // ISupportErrorInfo method
    STDMETHODIMP InterfaceSupportsErrorInfo(REFIID riid) override;

    // IArrakeener methods
    STDMETHODIMP get_FirstName(BSTR* pRet) override;
    STDMETHODIMP put_FirstName(BSTR value) override;
    STDMETHODIMP get_LastName(BSTR* pRet) override;
    STDMETHODIMP put_LastName(BSTR value) override;
    STDMETHODIMP get_Affiliation(BSTR* pRet) override;
    STDMETHODIMP put_Affiliation(BSTR value) override;
    STDMETHODIMP get_Occupation(BSTR* pRet) override;
    STDMETHODIMP put_Occupation(BSTR value) override;
    STDMETHODIMP get_Energy(LONGLONG* pRet) override;
    STDMETHODIMP get_Solaris(LONGLONG* pRet) override;
    STDMETHODIMP get_Spice(LONGLONG* pRet) override;
    STDMETHODIMP EatSpice(LONGLONG units, LONGLONG* pDeltaEnergy) override;
    STDMETHODIMP SellSpice(LONGLONG units, LONGLONG* pDeltaSolaris) override;
    STDMETHODIMP MineSpice(LONGLONG harvesters, LONGLONG* pDeltaSpice) override;
};

class CArrakeenerClass : public IClassFactory
{
public:
    // IUnknown methods
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // IClassFactory methods
    STDMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override;
    STDMETHODIMP LockServer(BOOL fLock) override;
};
