#pragma once

#include "Framework.h"

class CoCalculatorFactory : public IClassFactory_
{
	friend HRESULT __stdcall DllCanUnloadNow();

public:
	CoCalculatorFactory();
	~CoCalculatorFactory();

	// IUnknown
	virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
	virtual ULONG __stdcall AddRef();
	virtual ULONG __stdcall Release();

	// IClassFactory_
	virtual HRESULT __stdcall CreateInstance(IUnknown* pUnk, const IID& iid, void** ppv);
	virtual HRESULT __stdcall LockServer(BOOL fLock);

private:
	ULONG m_nNumRefs;
	static ULONG s_nNumObjs;
};
