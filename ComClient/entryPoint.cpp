// HardCOM.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

#include <iostream>
#include <Windows.h>
#include <string>

// For ComServer
//#include "../ComServer/CoCalculator.h"
//#include "../ComServer/Interfaces.h"
//#define ICalculator         ICalculator_

// For ComServer.ATL
#include "../ComServer.ATL/ComServerATL_i.h"
#include "../ComServer.ATL/ComServerATL_i.c"

// For _com_error
#pragma comment(lib, "RuntimeObject.lib")
#undef WINAPI_PARTITION_DESKTOP
#include <comdef.h>

void printErrorMessage(const CHAR* pszMessage, HRESULT hr)
{
    std::wstring error(_com_error(hr, NULL).ErrorMessage());
    std::cout << pszMessage;
    std::wcout << L" :: " << error.substr(0, error.size() - 2);
    std::cout << std::endl;
}

int main()
{
    setlocale(LC_ALL, "Russian");
    SetConsoleOutputCP(866);

    HRESULT hr = S_OK;

    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (!SUCCEEDED(hr))
    {
        printErrorMessage("Initialize failed", hr);
        return 1;
    }
    std::cout << "Initialized" << "\n";

    ICalculator* pICalc = NULL;
    hr = CoCreateInstance(CLSID_CoCalculator, NULL, CLSCTX_INPROC_SERVER, IID_ICalculator, (void**)&pICalc);
    if (SUCCEEDED(hr))
    {
        std::cout << "Query interface: ICalculator" << "\n";
        DOUBLE dResult = 0;

        hr = pICalc->Sum(5, 1, &dResult);
        if (SUCCEEDED(hr))
            std::cout << "Run Calculator operation: Sum => 5 + 1 = " << dResult << "\n";
        else
            printErrorMessage("Faild Calculator Sum operation", hr);

        hr = pICalc->Sub(5, 1, &dResult);
        if (SUCCEEDED(hr))
            std::cout << "Run Calculator operation: Sub => 5 - 1 = " << dResult << "\n";
        else
            printErrorMessage("Faild Calculator Sub operation", hr);

        hr = pICalc->Mul(5, 3, &dResult);
        if (SUCCEEDED(hr))
            std::cout << "Run Calculator operation: Mul => 5 * 3 = " << dResult << "\n";
        else
            printErrorMessage("Faild Calculator Mul operation", hr);

        hr = pICalc->Div(5, 3, &dResult);
        if (SUCCEEDED(hr))
            std::cout << "Run Calculator operation: Div => 5 / 3 = " << dResult << "\n";
        else
            printErrorMessage("Faild Calculator Div operation", hr);

        hr = pICalc->Div(5, 0, &dResult);
        if (SUCCEEDED(hr))
            std::cout << "Run Calculator operation: Div => 5 / 0 = " << dResult << "\n";
        else
            printErrorMessage("Faild Calculator Div operation", hr);

        ULONG cRef = pICalc->Release();
        std::cout << "Release interface: ICalculator, Ref=" << cRef << "\n";
    }
    else
    {
        printErrorMessage("Create instance: CLSID_CoCalculator", hr);
    }

    std::cout << "Uninitialized" << "\n";
    CoUninitialize();

    system("pause");

    return SUCCEEDED(hr) ? 0 : 1;
}
