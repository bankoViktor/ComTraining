// dllEntryPoint.cpp: реализация DllMain.

#include "pch.h"
#include "CComServerATLModule.h"


// Точка входа DLL
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    return _AtlModule.DllMain(dwReason, lpReserved);
}
