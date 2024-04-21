// comEntryPoint.cpp : Определяет точку входа для приложения DLL.
#include "pch.h"
#include "Framework.h"
#include "CoCalculator.h"
#include "CoCalculatorFactory.h"

#include <strsafe.h>

static HRESULT RegisterComClass(const IID& clsid, LPCTSTR pszClassName, LPCTSTR pszThreadModel, LPCTSTR pszVersion);
static HRESULT UnregisterComClass(const IID& clsid);

// https://zarezky.spb.ru/lectures/com/inproc-server.html
ULONG       g_nNumLocks;

#pragma comment(linker, "/export:DllGetClassObject,PRIVATE")
extern "C" HRESULT __stdcall DllGetClassObject(const IID& clsid, const IID& iid, void** ppv)
{
    if (clsid == CLSID_CoCalculator)
    {
        CoCalculatorFactory* pFact = new CoCalculatorFactory();
        if (pFact == 0)
        {
            *ppv = NULL;
            return E_OUTOFMEMORY;
        }

        HRESULT hr = pFact->QueryInterface(iid, ppv);
        if (IS_ERROR(hr))
        {
            delete pFact;
            *ppv = NULL;
        }
        return hr;
    }

    *ppv = NULL;
    return CLASS_E_CLASSNOTAVAILABLE;
}

#pragma comment(linker, "/export:DllCanUnloadNow,PRIVATE")
extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    ULONG nNumObjs =
        CoCalculatorFactory::s_nNumObjs +
        CoCalculator::s_nNumObjs;
    return g_nNumLocks == 0 && nNumObjs == 0 ? S_OK : S_FALSE;
}

#pragma comment(linker, "/export:DllRegisterServer,PRIVATE")
extern "C" HRESULT __stdcall DllRegisterServer()
{
    HRESULT hr = S_OK;

    hr = RegisterComClass(CLSID_CoCalculator, TEXT("ComServer :: Calculator class"), TEXT("Apartment"), TEXT("1.2"));
    if (FAILED(hr))
    {
        UnregisterComClass(CLSID_CoCalculator);
        return hr;
    }

    return hr;
}

#pragma comment(linker, "/export:DllUnregisterServer,PRIVATE")
extern "C" HRESULT __stdcall DllUnregisterServer()
{
    HRESULT hr = S_OK;

    hr = UnregisterComClass(CLSID_CoCalculator);
    if (FAILED(hr))
    {
        return hr;
    }

    return hr;
}

// -----------------------------------------------------------------------------

static HRESULT GetComClassPath(const IID& clsid, LPTSTR pszDest, size_t cchDest)
{
    HRESULT hr = S_OK;

    TCHAR szClassId[40];
    DWORD dwCharCount = StringFromGUID2(clsid, szClassId, 39);
    if (dwCharCount <= 0)
    {
        hr = E_FAIL;
    }

    return StringCchPrintf(pszDest, cchDest, TEXT("SOFTWARE\\Classes\\CLSID\\%s"), szClassId);
}

static HRESULT CreateOrSetStringValue(HKEY hKey, LPCTSTR pszName, LPCTSTR pszValue)
{
    HRESULT hr = S_OK;
    LSTATUS retCode;
    size_t nBytes;

    // Calc class name byte count
    hr = StringCbLength(pszValue, STRSAFE_MAX_CCH * sizeof(TCHAR), &nBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    // Set class name
    retCode = RegSetValueEx(hKey, pszName, NULL, REG_SZ, (BYTE*)pszValue, (DWORD)nBytes);
    if (retCode != ERROR_SUCCESS)
    {
        hr = HRESULT_FROM_WIN32(retCode);
    }

    return hr;
}

static HRESULT RegisterComClass(const IID& clsid, LPCTSTR pszClassName, LPCTSTR pszThreadModel, LPCTSTR pszVersion)
{
    HRESULT hr = S_OK;
    LSTATUS retCode;
    HKEY hkClass = 0;
    HKEY hkProcServer = 0;
    HKEY hkVersion = 0;
    DWORD dwCharCount;
    size_t nBufferCountChar = FILENAME_MAX;
    TCHAR szClassRoot[FILENAME_MAX + 1];
    TCHAR szBuffer[FILENAME_MAX + 1];

    // Get class Path
    hr = GetComClassPath(clsid, szClassRoot, nBufferCountChar);
    if (FAILED(hr))
    {
        return hr;
    }

    // Open class key
    retCode = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szClassRoot, NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkClass, NULL);
    if (retCode != ERROR_SUCCESS)
    {
        hr = HRESULT_FROM_WIN32(retCode);
        goto CleanUp;
    }

    // Calc class name
    hr = CreateOrSetStringValue(hkClass, NULL, pszClassName);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

    // ----------------------------------------------------------------------------------

    // Get 'InprocServer32' subkey
    hr = StringCchPrintf(szBuffer, nBufferCountChar, TEXT("%s\\InprocServer32"), szClassRoot);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

    // Open 'InprocServer32' subkey
    retCode = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szBuffer, NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkProcServer, NULL);
    if (retCode != ERROR_SUCCESS)
    {
        hr = HRESULT_FROM_WIN32(retCode);
        goto CleanUp;
    }

    // Get path to COM server
    dwCharCount = GetModuleFileName(g_hInstance, szBuffer, (DWORD)nBufferCountChar);
    if (dwCharCount <= 0)
    {
        hr = E_FAIL;
        goto CleanUp;
    }

    // Set path to COM server
    hr = CreateOrSetStringValue(hkProcServer, NULL, szBuffer);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

    // Set thread mode
    hr = CreateOrSetStringValue(hkProcServer, TEXT("ThreadingModel"), pszThreadModel);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

    // ----------------------------------------------------------------------------------

    // Get 'Version' subkey
    hr = StringCchPrintf(szBuffer, nBufferCountChar, TEXT("%s\\Version"), szClassRoot);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

    // Open 'Version' subkey
    retCode = RegCreateKeyEx(HKEY_LOCAL_MACHINE, szBuffer, NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkVersion, NULL);
    if (retCode != ERROR_SUCCESS)
    {
        hr = HRESULT_FROM_WIN32(retCode);
        goto CleanUp;
    }

    // Set path to COM server
    hr = CreateOrSetStringValue(hkVersion, NULL, pszVersion);
    if (FAILED(hr))
    {
        goto CleanUp;
    }

CleanUp:
    if (hkProcServer != NULL)
        RegCloseKey(hkProcServer);

    if (hkVersion != NULL)
        RegCloseKey(hkVersion);

    if (hkClass != NULL)
        RegCloseKey(hkClass);

    return hr;
}

static HRESULT UnregisterComClass(const IID& clsid)
{
    HRESULT hr = S_OK;
    LSTATUS retCode;
    HKEY hkClass = 0;
    size_t nBufferCountChar = FILENAME_MAX;
    TCHAR szBuffer[FILENAME_MAX + 1];

    // Get class Path
    hr = GetComClassPath(clsid, szBuffer, nBufferCountChar);
    if (FAILED(hr))
    {
        return hr;
    }

    // Delete class key
    retCode = RegDeleteTree(HKEY_LOCAL_MACHINE, szBuffer);
    if (retCode != ERROR_SUCCESS)
    {
        hr = HRESULT_FROM_WIN32(retCode);
    }

    return hr;
}
