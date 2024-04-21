// entryPoint.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
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


void printErrorMessage(LPCWSTR pszMessage, HRESULT hr)
{
    std::wcout << pszMessage << L" :: Error 0x" << std::hex << hr;
}

int main()
{
    setlocale(LC_ALL, "Russian");
    SetConsoleOutputCP(866);

    HRESULT hr = S_OK;

    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (!SUCCEEDED(hr))
    {
        printErrorMessage(L"Initialize failed", hr);
        return 1;
    }
    std::wcout << L"Initialized" << std::endl;

    std::wcout << L"Create CoCalculator object and get interface ICalculator";
    ICalculator* pICalc = NULL;
    hr = CoCreateInstance(CLSID_CoCalculator, NULL, CLSCTX_INPROC_SERVER, IID_ICalculator, (void**)&pICalc);
    if (SUCCEEDED(hr))
    {
        std::wcout << std::endl;
        DOUBLE dResult = 0;

        double nA = 5;
        double nB = 3;

        // SUM
        std::wcout << L"Call Calculator.Sum(" << nA << L" + " << nB << L") = ";
        hr = pICalc->Sum(nA, nB, &dResult);
        if (SUCCEEDED(hr))
            std::wcout << dResult;
        else
            printErrorMessage(L"FAILD", hr);
        std::wcout << std::endl;

        // SUB
        std::wcout << L"Call Calculator.Sub(" << nA << L" - " << nB << L") = ";
        hr = pICalc->Sub(nA, nB, &dResult);
        if (SUCCEEDED(hr))
            std::wcout << dResult;
        else
            printErrorMessage(L"FAILD", hr);
        std::wcout << std::endl;

        // MUL
        std::wcout << L"Call Calculator.Mul(" << nA << L" * " << nB << L") = ";
        hr = pICalc->Mul(nA, nB, &dResult);
        if (SUCCEEDED(hr))
            std::wcout << dResult;
        else
            printErrorMessage(L"FAILD", hr);
        std::wcout << std::endl;

        // DIV
        std::wcout << L"Call Calculator.Div(" << nA << L" / " << nB << L") = ";
        hr = pICalc->Div(nA, nB, &dResult);
        if (SUCCEEDED(hr))
            std::wcout << dResult;
        else
            printErrorMessage(L"FAILD", hr);
        std::wcout << std::endl;

        // DIV on 0
        nB = 0;
        std::wcout << L"Call Calculator.Div(" << nA << L" / " << nB << L") = ";
        hr = pICalc->Div(nA, nB, &dResult);
        if (SUCCEEDED(hr))
            std::wcout << dResult;
        else
            printErrorMessage(L"FAILD", hr);
        std::wcout << std::endl;

        ULONG cRef = pICalc->Release();
        std::cout << "Release interface: ICalculator, Ref=" << cRef << std::endl;
    }
    else
    {
        printErrorMessage(L"FAILED", hr);
    }
    std::wcout << std::endl;

    std::wcout << L"Uninitialized" << std::endl;
    CoUninitialize();

    system("pause");

    return SUCCEEDED(hr) ? 0 : 1;
}
