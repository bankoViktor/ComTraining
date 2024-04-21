// dllmain.cpp : Определяет точку входа для приложения DLL.

#include "Framework.h"

// {921F2D2C-9752-4BC7-AB2E-E5FB39D165B5}
static const GUID CLSID_CoCalculator =
{ 0x921f2d2c, 0x9752, 0x4bc7, { 0xab, 0x2e, 0xe5, 0xfb, 0x39, 0xd1, 0x65, 0xb5 } };

class CoCalculator : public ICalculator
{
    friend HRESULT __stdcall DllCanUnloadNow();

public:
    CoCalculator();
    ~CoCalculator();

    // IUnknown
    virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
    virtual ULONG __stdcall AddRef();
    virtual ULONG __stdcall Release();

    // ICalculator
    virtual HRESULT __stdcall Sum(DOUBLE a, DOUBLE b, DOUBLE* res);
    virtual HRESULT __stdcall Sub(DOUBLE a, DOUBLE b, DOUBLE* res);
    virtual HRESULT __stdcall Mul(DOUBLE a, DOUBLE b, DOUBLE* res);
    virtual HRESULT __stdcall Div(DOUBLE a, DOUBLE b, DOUBLE* res);

private:
    ULONG m_nNumRefs;
    static ULONG s_nNumObjs;
};