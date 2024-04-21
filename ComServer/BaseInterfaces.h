#pragma once

#include "Framework.h"

interface IUnknown_
{
    virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv) = 0;
    virtual ULONG __stdcall AddRef() = 0;
    virtual ULONG __stdcall Release() = 0;
};

interface IClassFactory_ : public IUnknown_
{
    virtual HRESULT __stdcall CreateInstance(IUnknown* pUnk, const IID& iid, void** ppvDest) = 0;
    virtual HRESULT __stdcall LockServer(BOOL fLock) = 0;
};
