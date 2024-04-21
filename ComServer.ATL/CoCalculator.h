// CoCalculator.h: объявление CoCalculator

#pragma once
#include "resource.h"       // основные символы

#include "ComServerATL_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Однопотоковые COM-объекты не поддерживаются должным образом платформой Windows CE, например платформами Windows Mobile, в которых не предусмотрена полная поддержка DCOM. Определите _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA, чтобы принудить ATL поддерживать создание однопотоковых COM-объектов и разрешить использование его реализаций однопотоковых COM-объектов. Для потоковой модели в вашем rgs-файле задано значение 'Free', поскольку это единственная потоковая модель, поддерживаемая не-DCOM платформами Windows CE."
#endif

using namespace ATL;


// CoCalculator

class ATL_NO_VTABLE CoCalculator :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CoCalculator, &CLSID_CoCalculator>,
    public IDispatchImpl<ICalculator, &IID_ICalculator, &LIBID_ComServerATLLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
    CoCalculator()
    {
    }

DECLARE_REGISTRY_RESOURCEID(IDR_COCALCULATOR)


BEGIN_COM_MAP(CoCalculator)
    COM_INTERFACE_ENTRY(ICalculator)
    COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct()
    {
        return S_OK;
    }

    void FinalRelease()
    {
    }

public:
    STDMETHOD(Sum)(DOUBLE a, DOUBLE b, DOUBLE* res);
    STDMETHOD(Sub)(DOUBLE a, DOUBLE b, DOUBLE* res);
    STDMETHOD(Mul)(DOUBLE a, DOUBLE b, DOUBLE* res);
    STDMETHOD(Div)(DOUBLE a, DOUBLE b, DOUBLE* res);
};

OBJECT_ENTRY_AUTO(__uuidof(CoCalculator), CoCalculator)
