// arrakeener.cpp
#include "arrakeener.h"
#include "arrakis.h"
#include "resource.h"
#include <cassert>

///////////////////////////////////////////////////////////////////////////////
//
// COMPUTATIONS
//

// Generate a random number within the provided range

static LONGLONG randrange(int lower, int upper)
{
    assert(lower < upper);
    int num = (rand() % (upper - lower + 1)) + lower;
    return (LONGLONG)num;
}


// Safe add (see C CERT rule 04  which addresses integer safety)

static bool safe_add(int64_t const a, int64_t const b, int64_t& sum)
{
    assert(b >= 0);
    if ((b > 0) && (a > (INT64_MAX - b))) return false;
    sum = a + b;
    return true;
}


// Safe multiply (see C CERT rule 04  which addresses integer safety)

static bool safe_multiply(int64_t const a, int64_t const b, int64_t& prod)
{
    assert(a > 0);
    assert(b > 0);
    if (a > (INT64_MAX / b)) return false;
    prod = a * b;
    return true;
}

///////////////////////////////////////////////////////////////////////////////
//
// CArrakeener: Instance class for Arrakeener
//

CArrakeener::CArrakeener() : m_rc(0), m_pti(nullptr),
m_energy(randrange(1, 100)), m_solaris(randrange(200000, 400000)), m_spice(0)
{
    ITypeLib* ptl = nullptr;
    HRESULT hr = LoadRegTypeLib(LIBID_ArrakisLib, 1, 0, 0, &ptl);
    if(FAILED(hr)) throw hr;
    hr = ptl->GetTypeInfoOfGuid(IID_IArrakeener, &m_pti);
    ptl->Release();
    if (FAILED(hr)) throw hr;
    assert(m_pti);

    InitializeCriticalSection(&m_cs);
}


CArrakeener::~CArrakeener() noexcept
{
    m_pti->Release();
    DeleteCriticalSection(&m_cs);
}


void CArrakeener::Lock() noexcept
{
    EnterCriticalSection(&m_cs);
}


void CArrakeener::Unlock() noexcept
{
    LeaveCriticalSection(&m_cs);
}


void CArrakeener::SetError(UINT id) noexcept
{
    ICreateErrorInfo* pcei = nullptr;
    HRESULT hr = CreateErrorInfo(&pcei);
    if (FAILED(hr)) return;

    pcei->SetSource(const_cast<wchar_t*>(L"Arrakis.Arrakeener.1"));
    pcei->SetGUID(IID_IArrakeener);

    wchar_t buf[1024];
    int cch = LoadStringW(GetModuleHandle(nullptr), id, buf, sizeof(buf) / sizeof(buf[0]));
    if (cch > 0)
    {
        assert(cch < sizeof(buf) / sizeof(buf[0]));
        buf[cch + 1] = (wchar_t)0;
        pcei->SetDescription(buf);
    }
    else
    {
        pcei->SetDescription(const_cast<wchar_t*>(L"Could not load resource string"));
    }

    IErrorInfo* pei = nullptr;
    hr = pcei->QueryInterface(IID_IErrorInfo, (void**)&pei);
    if (SUCCEEDED(hr))
    {
        SetErrorInfo(0, pei);
        pei->Release();
    }

    pcei->Release();
}


STDMETHODIMP CArrakeener::InterfaceSupportsErrorInfo(REFIID riid)
{
    return riid == IID_IArrakeener ? S_OK : S_FALSE;
}


STDMETHODIMP CArrakeener::QueryInterface(REFIID riid, void** ppv)
{
    // Note that this is an out-of process server
    // When called by clients, ppv will always be checked against NULL by the
    // universal marshaler, so a check is not needed in release builds.
    // Hence, the use of assert instead of if(!ppv) return E_INVALIDARG

    assert(ppv);
    if (riid == IID_IUnknown || riid == IID_IDispatch) *ppv = static_cast<IDispatch*>(this);
    else if (riid == IID_IArrakeener) *ppv = static_cast<IArrakeener*>(this);
    else if (riid == IID_ISupportErrorInfo) *ppv = static_cast<ISupportErrorInfo*>(this);
    else return (*ppv = nullptr), E_NOINTERFACE;
    reinterpret_cast<IUnknown*>(this)->AddRef();
    return S_OK;
}


STDMETHODIMP_(ULONG) CArrakeener::AddRef()
{
    if (m_rc == 0) LockModule();
    return InterlockedIncrement(&m_rc);
}


STDMETHODIMP_(ULONG) CArrakeener::Release()
{
    LONG rc = InterlockedDecrement(&m_rc);
    if (rc == 0)
    {
        delete this;
        UnlockModule();
    }
    return rc;
}


STDMETHODIMP CArrakeener::GetTypeInfoCount(UINT* pctinfo)
{
    assert(pctinfo);
    *pctinfo = 1;
    return S_OK;
}


STDMETHODIMP CArrakeener::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    assert(!iTInfo);
    assert(ppTInfo);
    (*ppTInfo = m_pti)->AddRef();
    return S_OK;
}


STDMETHODIMP CArrakeener::GetIDsOfNames(
    REFIID riid,
    LPOLESTR* rgszNames,
    UINT cNames,
    LCID lcid,
    DISPID* rgDispId)
{
    assert(riid == IID_NULL);
    return m_pti->GetIDsOfNames(rgszNames, cNames, rgDispId);
}


STDMETHODIMP CArrakeener::Invoke(
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS* pDispParams,
    VARIANT* pVarResult,
    EXCEPINFO* pExcepInfo,
    UINT* puArgErr)
{
    assert(riid == IID_NULL);
    return m_pti->Invoke(
        static_cast<IDispatch*>(this),
        dispIdMember,
        wFlags,
        pDispParams,
        pVarResult,
        pExcepInfo,
        puArgErr);
}


STDMETHODIMP CArrakeener::get_FirstName(BSTR* pRet)
{
    HRESULT hr;
    assert(pRet);
    Lock();
    *pRet = SysAllocString(m_first_name.c_str());   // c_str is nothrow
    hr = *pRet ? S_OK : E_OUTOFMEMORY;
    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::put_FirstName(BSTR value)
{
    HRESULT hr;
    Lock();

    try
    {
        m_first_name = value ? value : L"";
        hr = S_OK;
    }
    catch (std::bad_alloc&)
    {
        hr = E_OUTOFMEMORY;
    }
    catch (...)
    {
        hr = E_FAIL;
    }

    Unlock();
    return hr;
}


STDMETHODIMP::CArrakeener::get_LastName(BSTR* pRet)
{
    HRESULT hr;
    assert(pRet);
    Lock();
    *pRet = SysAllocString(m_last_name.c_str());
    hr = *pRet ? S_OK : E_OUTOFMEMORY;
    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::put_LastName(BSTR value)
{
    HRESULT hr;
    Lock();

    try
    {
        m_last_name = value ? value : L"";
        hr = S_OK;
    }
    catch (std::bad_alloc&)
    {
        hr = E_OUTOFMEMORY;
    }
    catch (...)
    {
        hr = E_FAIL;
    }

    Unlock();
    return hr;
}


STDMETHODIMP::CArrakeener::get_Affiliation(BSTR* pRet)
{
    HRESULT hr;
    assert(pRet);
    Lock();
    *pRet = SysAllocString(m_affiliation.c_str());
    hr = *pRet ? S_OK : E_OUTOFMEMORY;
    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::put_Affiliation(BSTR value)
{
    HRESULT hr;
    Lock();

    try
    {
        m_affiliation = value ? value : L"";
        hr = S_OK;
    }
    catch (std::bad_alloc&)
    {
        hr = E_OUTOFMEMORY;
    }
    catch (...)
    {
        hr = E_FAIL;
    }

    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::get_Occupation(BSTR* pRet)
{
    HRESULT hr;
    assert(pRet);
    Lock();
    *pRet = SysAllocString(m_occupation.c_str());
    hr = *pRet ? S_OK : E_OUTOFMEMORY;
    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::put_Occupation(BSTR value)
{
    HRESULT hr;
    Lock();

    try
    {
        m_occupation = value ? value : L"";
        hr = S_OK;
    }
    catch (std::bad_alloc&)
    {
        hr = E_OUTOFMEMORY;
    }
    catch (...)
    {
        hr = E_FAIL;
    }

    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::get_Energy(LONGLONG* pRet)
{
    assert(pRet);
    Lock();
    *pRet = m_energy;
    Unlock();
    return S_OK;
}


STDMETHODIMP CArrakeener::get_Solaris(LONGLONG* pRet)
{
    assert(pRet);
    Lock();
    *pRet = m_solaris;
    Unlock();
    return S_OK;
}


STDMETHODIMP CArrakeener::get_Spice(LONGLONG* pRet)
{
    assert(pRet);
    Lock();
    *pRet = m_spice;
    Unlock();
    return S_OK;
}


STDMETHODIMP CArrakeener::EatSpice(LONGLONG units, LONGLONG* pDeltaEnergy)
{
    HRESULT hr;
    assert(pDeltaEnergy);
    Lock();
    *pDeltaEnergy = 0;

    if (units < 1)
    {
        SetError(IDS_NONPOSSPICE);
        hr = E_INVALIDARG;
    }
    else if (m_spice < units)
    {
        SetError(IDS_NOSPICE);
        hr = E_FAIL;
    }
    else
    {
        LONGLONG delta_energy = 0;
        if (!safe_multiply(randrange(1, 100), units, delta_energy))
        {
            SetError(IDS_OVERFLOW);
            hr = E_FAIL;
        }
        else
        {
            LONGLONG new_energy = 0;
            if (!safe_add(m_energy, delta_energy, new_energy))
            {
                SetError(IDS_OVERFLOW);
                hr = E_FAIL;
            }
            else
            {
                m_spice -= units;           // This cannot overflow
                m_energy = new_energy;      // This has been checked

                *pDeltaEnergy = delta_energy;
                hr = S_OK;
            }
        }
    }

    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::SellSpice(LONGLONG units, LONGLONG* pDeltaSolaris)
{
    HRESULT hr;
    assert(pDeltaSolaris);
    Lock();
    *pDeltaSolaris = 0;

    if (units < 1)
    {
        SetError(IDS_NONPOSSPICE);
        hr = E_INVALIDARG;
    }
    else if (m_spice < units)
    {
        SetError(IDS_NOSPICE);
        hr = E_FAIL;
    }
    else
    {
        LONGLONG delta_solaris = 0;
        if (!safe_multiply(randrange(200000, 700000), units, delta_solaris))
        {
            SetError(IDS_OVERFLOW);
            hr = E_FAIL;
        }
        else
        {
            LONGLONG new_solaris;
            if (!safe_add(m_solaris, delta_solaris, new_solaris))
            {
                SetError(IDS_OVERFLOW);
                hr = E_FAIL;
            }
            else
            {
                m_spice -= units;           // This cannot overflow
                m_solaris = new_solaris;    // This has been checked

                *pDeltaSolaris = delta_solaris;
                hr = S_OK;
            }
        }
    }

    Unlock();
    return hr;
}


STDMETHODIMP CArrakeener::MineSpice(LONGLONG harvesters, LONGLONG* pDeltaSpice)
{
    HRESULT hr;
    assert(pDeltaSpice);
    Lock();
    *pDeltaSpice = 0;

    if (harvesters < 1)
    {
        SetError(IDS_NOHARVESTER);
        hr = E_INVALIDARG;
    }
    else
    {
        LONGLONG delta_energy = randrange(1, 10);
        if (m_energy < delta_energy)
        {
            SetError(IDS_NOENERGY);
            hr = E_FAIL;
        }
        else
        {
            LONGLONG delta_solaris = 0, delta_spice = 0;
            if (!safe_multiply(randrange(100000, 200000), harvesters, delta_solaris) ||
                !safe_multiply(randrange(1, 50), harvesters, delta_spice))
            {
                SetError(IDS_OVERFLOW);
                hr = E_FAIL;
            }
            else
            {
                if (m_solaris < delta_solaris)
                {
                    SetError(IDS_NOSOLARIS);
                    hr = E_FAIL;
                }
                else
                {
                    LONGLONG new_spice;
                    if (!safe_add(m_spice, delta_spice, new_spice))
                    {
                        SetError(IDS_OVERFLOW);
                        hr = E_FAIL;
                    }
                    else
                    {
                        m_energy -= delta_energy;       // This cannot overflow
                        m_solaris -= delta_solaris;     // This cannot overflow
                        m_spice = new_spice;            // This has been checked

                        *pDeltaSpice = delta_spice;
                        hr = S_OK;
                    }
                }
            }
        }
    }

    Unlock();
    return hr;
}

///////////////////////////////////////////////////////////////////////////////
//
// CArrakeenerClass: Class factory for Arrakeener
//

STDMETHODIMP CArrakeenerClass::QueryInterface(REFIID riid, void** ppv)
{
    if (!ppv) return E_INVALIDARG;
    if (riid == IID_IUnknown) *ppv = static_cast<IUnknown*>(this);
    else if (riid == IID_IClassFactory) *ppv = static_cast<IClassFactory*>(this);
    else return (*ppv = nullptr), E_NOINTERFACE;
    reinterpret_cast<IUnknown*>(this)->AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) CArrakeenerClass::AddRef()
{
    return 2;
}

STDMETHODIMP_(ULONG) CArrakeenerClass::Release()
{
    return 1;
}

STDMETHODIMP CArrakeenerClass::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    HRESULT hr;
    assert(ppv);    // For out-of-process servers, this is never null
    LockModule();
    *ppv = nullptr;

    DWORD dw = WaitForSingleObject(g_done, 0);
    if (dw == WAIT_OBJECT_0)
    {
        hr = CO_E_SERVER_STOPPING;
    }
    else
    {
        if (pUnkOuter)
        {
            hr = CLASS_E_NOAGGREGATION;
        }
        else
        {
            try
            {
                CArrakeener* p = new CArrakeener;
                p->AddRef();
                hr = p->QueryInterface(riid, ppv);
                p->Release();
            }
            catch (std::bad_alloc&)
            {
                hr = E_OUTOFMEMORY;
            }
            catch (HRESULT& hr2)
            {
                hr = hr2;
            }
            catch (...)
            {
                hr = E_FAIL;
            }
        }
    }

    UnlockModule();
    return hr;
}

STDMETHODIMP CArrakeenerClass::LockServer(BOOL fLock)
{
    if (fLock) LockModule();
    else UnlockModule();
    return S_OK;
}
