#pragma once

#include "BaseInterfaces.h"

// {7FA6BD16-F8C8-4F86-A240-E5BDB4D254EF}
extern "C" static const GUID IID_ICalculator =
{ 0x7fa6bd16, 0xf8c8, 0x4f86, { 0xa2, 0x40, 0xe5, 0xbd, 0xb4, 0xd2, 0x54, 0xef } };

interface ICalculator_ : public IUnknown_
{
    virtual HRESULT __stdcall Sum(DOUBLE a, DOUBLE b, DOUBLE* res) = 0;
    virtual HRESULT __stdcall Sub(DOUBLE a, DOUBLE b, DOUBLE* res) = 0;
    virtual HRESULT __stdcall Mul(DOUBLE a, DOUBLE b, DOUBLE* res) = 0;
    virtual HRESULT __stdcall Div(DOUBLE a, DOUBLE b, DOUBLE* res) = 0;
};
