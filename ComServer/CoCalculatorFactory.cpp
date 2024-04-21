
#include "pch.h"
#include "CoCalculatorFactory.h"
#include "CoCalculator.h"

ULONG CoCalculatorFactory::s_nNumObjs = 0;

CoCalculatorFactory::CoCalculatorFactory() :
    m_nNumRefs(0)
{
    ++s_nNumObjs;
}

CoCalculatorFactory::~CoCalculatorFactory()
{
    --s_nNumObjs;
}

HRESULT __stdcall CoCalculatorFactory::QueryInterface(const IID& iid, void** ppv)
{
    if (iid == IID_IUnknown)
        *ppv = (IUnknown_*)this;
    else if (iid == IID_IClassFactory)
        *ppv = (IClassFactory_*)this;
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    ((IUnknown_*)*ppv)->AddRef();
    return S_OK;
}

ULONG __stdcall CoCalculatorFactory::AddRef()
{
    return ++m_nNumRefs;
}

ULONG __stdcall CoCalculatorFactory::Release()
{
    --m_nNumRefs;

    if (m_nNumRefs == 0)
    {
        delete this;
        return 0;
    }

    return m_nNumRefs;
}

HRESULT __stdcall CoCalculatorFactory::CreateInstance(IUnknown* pUnk, const IID& iid, void** ppv)
{
    if (pUnk != NULL)
    {
        *ppv = NULL;
        return CLASS_E_NOAGGREGATION;
    }

    CoCalculator* pCalc = new CoCalculator();
    if (pCalc == NULL)
    {
        *ppv = NULL;
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pCalc->QueryInterface(iid, ppv);
    if (FAILED(hr))
    {
        delete pCalc;
        ppv = NULL;
        return hr;
    }

    return S_OK;
}

HRESULT __stdcall CoCalculatorFactory::LockServer(BOOL fLock)
{
    fLock ? ++g_nNumLocks : --g_nNumLocks;
    return S_OK;
}
