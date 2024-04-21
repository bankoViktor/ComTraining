// CComServerATLModule.h: объявление класса модуля.
#pragma once

#include "ComServerATL_i.h"

class CComServerATLModule : public ATL::CAtlDllModuleT<CComServerATLModule>
{
public:
    DECLARE_LIBID(LIBID_ComServerATLLib)
    DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMSERVERATL, "{5ccf61f3-db1e-42b7-9e09-c54899e7ff4a}")
};

extern class CComServerATLModule _AtlModule;
