#include "pch.h"
#include "CoCalculator.h"

ULONG CoCalculator::s_nNumObjs = 0;

CoCalculator::CoCalculator() :
    m_nNumRefs(0)
{
    ++s_nNumObjs;
}

CoCalculator::~CoCalculator()
{
    --s_nNumObjs;
}

HRESULT __stdcall CoCalculator::QueryInterface(const IID& iid, void** ppv)
{
    if (iid == IID_IUnknown)
        *ppv = (IUnknown_*)this;
    else if (iid == IID_ICalculator)
        *ppv = (ICalculator*)this;
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    ((IUnknown_*)*ppv)->AddRef();
    return S_OK;
}

ULONG __stdcall CoCalculator::AddRef()
{
    return ++m_nNumRefs;
}

ULONG __stdcall CoCalculator::Release()
{
    --m_nNumRefs;

    if (m_nNumRefs == 0)
    {
        delete this;
        return 0;
    }

    return m_nNumRefs;
}

HRESULT CoCalculator::Sum(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a + b;
    return S_OK;
};

HRESULT CoCalculator::Sub(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a - b;
    return S_OK;
};

HRESULT CoCalculator::Mul(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a * b;
    return S_OK;
};

HRESULT CoCalculator::Div(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    if (b == 0)
        return E_INVALIDARG;

    *res = a / b;
    return S_OK;
};
