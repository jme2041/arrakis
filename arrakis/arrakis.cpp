// arrakis.cpp: Application entry point

#include "arrakis.h"
#include "arrakis_i.c"
#include "arrakeener.h"
#include <OleCtl.h>
#include <ctime>

HANDLE g_done = CreateEvent(nullptr, TRUE, FALSE, nullptr);

void LockModule()
{
    CoAddRefServerProcess();
}

void UnlockModule()
{
    if (CoReleaseServerProcess() == 0) SetEvent(g_done);
}

// Registration table

static const wchar_t* const reg_table[][3] =
{
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}", nullptr, L"Person of Arrakis" },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\LocalServer32", nullptr, (const wchar_t*)-1 },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\TypeLib", nullptr, L"{FDB559CC-A723-11EC-B743-DC41A9695036}" },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\ProgID", nullptr, L"Arrakis.Arrakeener.1" },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\VersionIndependentProgID", nullptr, L"Arrakis.Arrakeener" },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\Version", nullptr, L"1.0" },
    { L"Software\\Classes\\CLSID\\{FDB559CB-A723-11EC-B743-DC41A9695036}\\Programmable", nullptr, nullptr },
    { L"Software\\Classes\\Arrakis.Arrakeener", nullptr, L"Person of Arrakis" },
    { L"Software\\Classes\\Arrakis.Arrakeener\\CLSID", nullptr, L"{FDB559CB-A723-11EC-B743-DC41A9695036}" },
    { L"Software\\Classes\\Arrakis.Arrakeener\\CurVer", nullptr, L"Arrakis.Arrakeener.1" },
    { L"Software\\Classes\\Arrakis.Arrakeener.1", nullptr, L"Person of Arrakis" },
    { L"Software\\Classes\\Arrakis.Arrakeener.1\\CLSID", nullptr, L"{FDB559CB-A723-11EC-B743-DC41A9695036}" }
};

STDAPI UnregisterServer(bool user)
{
    HRESULT hr;
    HKEY regroot;
    if (user)
    {
        hr = UnRegisterTypeLibForUser(LIBID_Arrakis, 1, 0, 0, SYS_WIN64);
        regroot = HKEY_CURRENT_USER;
    }
    else
    {
        hr = UnRegisterTypeLib(LIBID_Arrakis, 1, 0, 0, SYS_WIN64);
        regroot = HKEY_LOCAL_MACHINE;
    }

    const int n = sizeof(reg_table) / sizeof(reg_table[0]);
    for (int i = n - 1; i >= 0; --i)
    {
        LSTATUS err = RegDeleteKeyW(regroot, reg_table[i][0]);
        if (err != ERROR_SUCCESS) hr = S_FALSE;
    }
    return hr;
}

STDAPI RegisterServer(bool user)
{
    wchar_t file_name[MAX_PATH];
    GetModuleFileNameW(GetModuleHandle(nullptr), file_name, MAX_PATH);

    HKEY regroot = user ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE;

    ITypeLib* ptl = nullptr;
    HRESULT hr = LoadTypeLib(L"arrakis.exe", &ptl);
    if (SUCCEEDED(hr))
    {
        if (user) hr = RegisterTypeLibForUser(ptl, file_name, nullptr);
        else hr = RegisterTypeLib(ptl, file_name, nullptr);
        ptl->Release();
    }
    if (FAILED(hr)) return hr;

    const int n = sizeof(reg_table) / sizeof(reg_table[0]);
    for (int i = 0; i < n; ++i)
    {
        const wchar_t* const key_name = reg_table[i][0];
        const wchar_t* const value_name = reg_table[i][1];
        const wchar_t* value = reg_table[i][2];
        if (value == (const wchar_t*)-1) value = file_name;
        HKEY hKey;
        LSTATUS err = RegCreateKeyExW(
            regroot,
            key_name, 0,
            nullptr,
            0,
            KEY_ALL_ACCESS,
            nullptr,
            &hKey,
            nullptr);
        if (err == ERROR_SUCCESS && value)
        {
            err = RegSetValueExW(
                hKey,
                value_name,
                0,
                REG_SZ,
                (const BYTE*)value,
                (DWORD)((wcslen(value) + 1) * sizeof(wchar_t)));
            RegCloseKey(hKey);
        }
        if (err != ERROR_SUCCESS)
        {
            UnregisterServer(user);
            return SELFREG_E_CLASS;
        }
    }
    return S_OK;
}

int APIENTRY wWinMain(HINSTANCE, HINSTANCE, LPWSTR lpCmdLine, int)
{
    srand((unsigned int)time(nullptr));     // Initialize random seed

    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        MessageBoxW(nullptr, L"CoInitializeEx failed", L"Arrakis", MB_SETFOREGROUND | MB_ICONERROR);
        return hr;
    }

    if (wcsstr(lpCmdLine, L"/UnregServerPerUser") || wcsstr(lpCmdLine, L"-UnregServerPerUser"))
    {
        hr = UnregisterServer(true);
        CoUninitialize();
        return hr;
    }

    if (wcsstr(lpCmdLine, L"/UnregServer") || wcsstr(lpCmdLine, L"-UnregServer"))
    {
        hr = UnregisterServer(false);
        CoUninitialize();
        return hr;
    }

    if (wcsstr(lpCmdLine, L"/RegServerPerUser") || wcsstr(lpCmdLine, L"-RegServerPerUser"))
    {
        hr = RegisterServer(true);
        CoUninitialize();
        return hr;
    }

    if (wcsstr(lpCmdLine, L"/RegServer") || wcsstr(lpCmdLine, L"-RegServer"))
    {
        hr = RegisterServer(false);
        CoUninitialize();
        return hr;
    }

    DWORD dwReg;
    static CArrakeenerClass cac;
    hr = CoRegisterClassObject(
        CLSID_Arrakeener,
        &cac,
        CLSCTX_LOCAL_SERVER,
        REGCLS_SUSPENDED | REGCLS_MULTIPLEUSE,
        &dwReg);

    if (SUCCEEDED(hr))
    {
        hr = CoResumeClassObjects();
        if (SUCCEEDED(hr)) WaitForSingleObject(g_done, INFINITE);
        CoRevokeClassObject(dwReg);
    }

    if (FAILED(hr))
    {
        MessageBox(nullptr, L"Error starting Arrakis server", L"Arrakis", MB_SETFOREGROUND | MB_ICONERROR);
    }

    CoUninitialize();
    return 0;
}
