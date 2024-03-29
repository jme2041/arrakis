// TestArrakis.cpp: Native C++ unit tests for Arrakis
// Note that the unit test frame work starts COM and tests run in main STA

#include "pch.h"
#include "CppUnitTest.h"
#include "arrakis_h.h"
#include "arrakis_i.c"
#include <iptr.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace TestArrakis
{
    // Shim class for dealing with BSTRs as [in] parameters
    // See Box (1998). Essential COM.

    class _UBSTR {
        BSTR m_bstr;
    public:
        _UBSTR() : m_bstr(nullptr) { }
        _UBSTR(const wchar_t* pwsz) : m_bstr(SysAllocString(pwsz)) { }
        ~_UBSTR() { SysFreeString(m_bstr); }
        operator BSTR() const noexcept { return m_bstr; }
        explicit operator bool() const noexcept { return m_bstr != nullptr; }

        friend BSTR* set(_UBSTR& obj)
        {
            if (obj.m_bstr)
            {
                SysFreeString(obj.m_bstr);
                obj.m_bstr = nullptr;
            }
            return &obj.m_bstr;
        }
    };

    TEST_MODULE_INITIALIZE(ModuleInit)
    {
        APTTYPE at;
        APTTYPEQUALIFIER atq;
        HRESULT hr = CoGetApartmentType(&at, &atq);
        Assert::IsTrue(SUCCEEDED(hr));
        Assert::IsTrue(at == APTTYPE_MAINSTA);
        Assert::IsTrue(atq == APTTYPEQUALIFIER_NONE);
    }

    TEST_CLASS(TestRegistration)
    {
    public:

        TEST_METHOD(ProgID2CLSID)
        {
            CLSID clsid;
            HRESULT hr = CLSIDFromProgID(L"Arrakis.Arrakeener.1", &clsid);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(clsid == CLSID_Arrakeener);
        }

        TEST_METHOD(VersionIndependentProgID2CLSID)
        {
            CLSID clsid;
            HRESULT hr = CLSIDFromProgID(L"Arrakis.Arrakeener", &clsid);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(clsid == CLSID_Arrakeener);
        }

        TEST_METHOD(CLSID2ProgID)
        {
            OLECHAR* progid;
            HRESULT hr = ProgIDFromCLSID(CLSID_Arrakeener, &progid);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(L"Arrakis.Arrakeener.1", progid);
            if (progid) CoTaskMemFree(progid);
        }
    };

    TEST_CLASS(TestArrakeener)
    {
        IPtr<IArrakeener> arrakeener;
        IPtr<IArrakeener> ghola;

        void check_error_message(wchar_t const* expected) const
        {
            // This simulates how a client would extract error information
            // Don't assert on the strings until the end to ensure interface pointers are released

            //const size_t cch = 1024;
            //wchar_t source[cch], description[cch];
            _UBSTR source;
            _UBSTR description;

            ISupportErrorInfo* psei = nullptr;
            HRESULT hr = arrakeener->QueryInterface(IID_ISupportErrorInfo, (void**)&psei);
            if (SUCCEEDED(hr))
            {
                IErrorInfo* pei = nullptr;
                hr = GetErrorInfo(0, &pei);
                if (SUCCEEDED(hr))
                {
                    pei->GetSource(set(source));
                    Assert::IsTrue(source);

                    pei->GetDescription(set(description));
                    Assert::IsTrue(description);

                    pei->Release();
                }
                psei->Release();
            }
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(L"Arrakis.Arrakeener.1", source);
            Assert::AreEqual(expected, description);
        }

        void check_energy(LONGLONG expected) const
        {
            LONGLONG energy;
            HRESULT hr = arrakeener->get_Energy(&energy);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(expected, energy);
        }

        void check_solaris(LONGLONG expected) const
        {
            LONGLONG solaris;
            HRESULT hr = arrakeener->get_Solaris(&solaris);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(expected, solaris);
        }

        void check_spice(LONGLONG expected) const
        {
            LONGLONG spice;
            HRESULT hr = arrakeener->get_Spice(&spice);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(expected, spice);
        }

    public:

        TEST_METHOD_INITIALIZE(CreateArrakeener)
        {
            HRESULT hr = CoCreateInstance(CLSID_Arrakeener, nullptr, CLSCTX_ALL, IID_IArrakeener, (void**)set(arrakeener));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue((bool)arrakeener);
        }

        TEST_METHOD(FirstName)
        {
            _UBSTR bstr;

            HRESULT hr = arrakeener->get_FirstName(nullptr);
            Assert::IsTrue(FAILED(hr));

            hr = arrakeener->get_FirstName(set(bstr));
            Assert::AreEqual(L"", bstr);

            hr = arrakeener->put_FirstName(_UBSTR(L"Paul"));
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_FirstName(set(bstr));
            Assert::AreEqual(L"Paul", bstr);

            hr = arrakeener->put_FirstName(nullptr);
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_FirstName(set(bstr));
            Assert::AreEqual(L"", bstr);
        }

        TEST_METHOD(LastName)
        {
            _UBSTR bstr;

            HRESULT hr = arrakeener->get_LastName(nullptr);
            Assert::IsTrue(FAILED(hr));

            hr = arrakeener->get_LastName(set(bstr));
            Assert::AreEqual(L"", bstr);

            hr = arrakeener->put_LastName(_UBSTR(L"Atreides"));
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_LastName(set(bstr));
            Assert::AreEqual(L"Atreides", bstr);

            hr = arrakeener->put_LastName(nullptr);
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_LastName(set(bstr));
            Assert::AreEqual(L"", bstr);
        }

        TEST_METHOD(Affiliation)
        {
            _UBSTR bstr;

            HRESULT hr = arrakeener->get_Affiliation(nullptr);
            Assert::IsTrue(FAILED(hr));

            hr = arrakeener->get_Affiliation(set(bstr));
            Assert::AreEqual(L"", bstr);

            hr = arrakeener->put_Affiliation(_UBSTR(L"House Atreides"));
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_Affiliation(set(bstr));
            Assert::AreEqual(L"House Atreides", bstr);

            hr = arrakeener->put_Affiliation(nullptr);
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_Affiliation(set(bstr));
            Assert::AreEqual(L"", bstr);
        }

        TEST_METHOD(Occupation)
        {
            _UBSTR bstr;

            HRESULT hr = arrakeener->get_Occupation(nullptr);
            Assert::IsTrue(FAILED(hr));

            hr = arrakeener->get_Occupation(set(bstr));
            Assert::AreEqual(L"", bstr);

            hr = arrakeener->put_Occupation(_UBSTR(L"Kwisatz Haderach"));
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_Occupation(set(bstr));
            Assert::AreEqual(L"Kwisatz Haderach", bstr);

            hr = arrakeener->put_Occupation(nullptr);
            Assert::IsTrue(SUCCEEDED(hr));

            hr = arrakeener->get_Occupation(set(bstr));
            Assert::AreEqual(L"", bstr);
        }

        TEST_METHOD(InitialEnergy)
        {
            HRESULT hr = arrakeener->get_Energy(nullptr);
            Assert::IsTrue(FAILED(hr));

            LONGLONG energy;
            hr = arrakeener->get_Energy(&energy);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(energy >= 1);
            Assert::IsTrue(energy <= 100);
        }

        TEST_METHOD(InitialSolaris)
        {
            HRESULT hr = arrakeener->get_Solaris(nullptr);
            Assert::IsTrue(FAILED(hr));

            LONGLONG solaris;
            hr = arrakeener->get_Solaris(&solaris);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(solaris >= 200000);
            Assert::IsTrue(solaris <= 400000);
        }

        TEST_METHOD(InitialSpice)
        {
            check_spice(0LL);
        }

        TEST_METHOD(BadEat)
        {
            LONGLONG delta_energy;
            HRESULT hr = arrakeener->EatSpice(0LL, &delta_energy);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Spice units must be greater than zero");
            Assert::AreEqual(0LL, delta_energy);
        }

        TEST_METHOD(BadSell)
        {
            LONGLONG delta_solaris;
            HRESULT hr = arrakeener->SellSpice(0LL, &delta_solaris);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Spice units must be greater than zero");
            Assert::AreEqual(0LL, delta_solaris);
        }

        TEST_METHOD(BadMine)
        {
            LONGLONG delta_spice;
            HRESULT hr = arrakeener->MineSpice(0LL, &delta_spice);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Cannot mine spice without a harvester");
            Assert::AreEqual(0LL, delta_spice);
        }

        TEST_METHOD(Operations)
        {
            HRESULT hr;
            LONGLONG energy, solaris, spice = 0LL, tmp;
            LONGLONG delta_energy, delta_solaris, delta_spice;

            // Get initial values of energy and solaris
            hr = arrakeener->get_Energy(&energy);
            Assert::IsTrue(SUCCEEDED(hr));
            hr = arrakeener->get_Solaris(&solaris);
            Assert::IsTrue(SUCCEEDED(hr));

            // Try to sell without spice
            hr = arrakeener->SellSpice(1LL, &delta_solaris);
            Assert::IsTrue(FAILED(hr));
            Assert::AreEqual(0LL, delta_solaris);
            check_error_message(L"Insufficient spice");
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Try to eat without spice
            hr = arrakeener->EatSpice(1LL, &delta_energy);
            Assert::IsTrue(FAILED(hr));
            Assert::AreEqual(0LL, delta_energy);
            check_error_message(L"Insufficient spice");
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Mine some spice
            hr = arrakeener->MineSpice(1LL, &delta_spice);
            Assert::IsTrue(SUCCEEDED(hr));
            spice += delta_spice;
            check_spice(spice);

            // Energy and solaris will have randomly decreased
            hr = arrakeener->get_Energy(&tmp);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(tmp < energy);
            energy = tmp;
            hr = arrakeener->get_Solaris(&tmp);
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::IsTrue(tmp < solaris);
            solaris = tmp;

            // Eat one tenth of our spice
            delta_spice = spice / 10LL;
            hr = arrakeener->EatSpice(delta_spice, &delta_energy);
            Assert::IsTrue(SUCCEEDED(hr));
            energy += delta_energy;
            spice -= delta_spice;
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Try to eat too much spice
            hr = arrakeener->EatSpice(spice + 1LL, &delta_energy);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Insufficient spice");
            Assert::AreEqual(0LL, delta_energy);
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Sell half of our spice
            hr = arrakeener->SellSpice(spice / 2LL, &delta_solaris);
            Assert::IsTrue(SUCCEEDED(hr));
            spice -= spice / 2LL;
            solaris += delta_solaris;
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Try to sell too much spice
            hr = arrakeener->SellSpice(spice + 1LL, &delta_solaris);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Insufficient spice");
            Assert::AreEqual(0LL, delta_solaris);
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);

            // Check integer overflow
            hr = arrakeener->MineSpice(INT64_MAX - 1LL, &delta_spice);
            Assert::IsTrue(FAILED(hr));
            check_error_message(L"Integer overflow");
            Assert::AreEqual(0LL, delta_spice);
            check_energy(energy);
            check_solaris(solaris);
            check_spice(spice);
        }

        TEST_METHOD(Clone)
        {
            _UBSTR bstr1;
            _UBSTR bstr2;
            HRESULT hr;

            hr = arrakeener->put_FirstName(_UBSTR(L"Duncan"));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = arrakeener->put_LastName(_UBSTR(L"Idaho"));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = arrakeener->put_Affiliation(_UBSTR(L"House Atreides"));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = arrakeener->put_Occupation(_UBSTR(L"Swordmaster"));
            Assert::IsTrue(SUCCEEDED(hr));

            // Create an alias of Duncan
            // This just copies the pointer, not the data
            // It is the same Duncan
            IPtr<IArrakeener> p = arrakeener;

            // Create a copy of Duncan
            hr = arrakeener->Clone(set(ghola));
            Assert::IsTrue(SUCCEEDED(hr));

            hr = ghola->put_Occupation(_UBSTR(L"Ghola"));
            Assert::IsTrue(SUCCEEDED(hr));

            // Compare data between Duncan and his alias
            hr = arrakeener->get_FirstName(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = p->get_FirstName(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_LastName(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = p->get_LastName(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_Affiliation(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = p->get_Affiliation(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_Occupation(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = p->get_Occupation(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            // Compare data between Duncan and his Ghola
            hr = arrakeener->get_FirstName(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = ghola->get_FirstName(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_LastName(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = ghola->get_LastName(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_Affiliation(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = ghola->get_Affiliation(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreEqual(std::wstring(bstr1), std::wstring(bstr2));

            hr = arrakeener->get_Occupation(set(bstr1));
            Assert::IsTrue(SUCCEEDED(hr));
            hr = ghola->get_Occupation(set(bstr2));
            Assert::IsTrue(SUCCEEDED(hr));
            Assert::AreNotEqual(std::wstring(bstr1), std::wstring(bstr2));
        }
    };
}
