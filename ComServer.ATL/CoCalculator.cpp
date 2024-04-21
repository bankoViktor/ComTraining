// CoCalculator.cpp: реализация CoCalculator

#include "pch.h"
#include "CoCalculator.h"


STDMETHODIMP CoCalculator::Sum(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a + b;
    return S_OK;
}

STDMETHODIMP CoCalculator::Sub(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a - b;
    return S_OK;
}

STDMETHODIMP CoCalculator::Mul(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    *res = a * b;
    return S_OK;
}

STDMETHODIMP CoCalculator::Div(DOUBLE a, DOUBLE b, DOUBLE* res)
{
    if (b == 0)
        return E_INVALIDARG;
    *res = a / b;
    return S_OK;
}
