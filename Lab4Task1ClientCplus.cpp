#include "windows.h"
#include <iostream>
#include "stdio.h"
#include <atlcomcli.h>

int main()
{
    // Инициализация OLE
    DWORD clsctx;
    clsctx = CLSCTX_INPROC_SERVER;
    HRESULT hr = OleInitialize(NULL);
    if (FAILED(hr))
    {
        printf("Failed to initialize. Code 0x8%X \n", hr); return 1;
    }

    // получаем ProgId сервера
    wchar_t progid[] = L"Lb3AutoSvrMyMath.1";
    CLSID clsid;
    hr = ::CLSIDFromProgID(progid, &clsid);
    if (FAILED(hr))
    {
        printf("Failed to get CLSID.Code 0x8%X \n", hr); return
            1;
    }

    // Получаем интерфейс IDispatch
    IDispatch* pIDispatch = NULL;
    hr = ::CoCreateInstance(clsid, NULL, clsctx, IID_IDispatch,
        (void**)&pIDispatch);
    if (FAILED(hr))
    {
        printf("Create instance failed.Code 0x8%X \n", hr);//+
        OleUninitialize();
        return 1;
    }

    printf("CoCreateInstance succeeded.\n");
    printf("Get DispID for function \"Add\".\n");

    // Получаем id метода интерфейса.
    DISPID dispid;
    OLECHAR* name = (OLECHAR*)L"Add";//*
    hr = pIDispatch->GetIDsOfNames(IID_NULL,
        &name,
        1,
        GetUserDefaultLCID(),
        &dispid);
    if (FAILED(hr))
    {
        printf("Query GetIDsOfNames failed.Code %X \n", hr);
        pIDispatch->Release(); return 1;
    }

    // Подготавливаем аргументы.
    VARIANTARG vargs[2];//*
    ::VariantInit(&vargs[0]);
    vargs[0].vt = VT_I4; 
    vargs[0].lVal = 2; 
    ::VariantInit(&vargs[1]); 
    vargs[1].vt = VT_I4; 
    vargs[1].lVal = 99; 

    DISPPARAMS param;
    param.cArgs = 2; 
    param.rgvarg = vargs; 
    param.cNamedArgs = 0; 
    param.rgdispidNamedArgs = NULL; 

    VARIANTARG varres;//*
    ::VariantInit(&varres);
    varres.vt = VT_I4; 
    varres.lVal = 0; 
    printf("Invoke the function \"Add\".\n");

    // Вызов метода интерфейса
    hr = pIDispatch->Invoke(dispid,
        IID_NULL,
        GetUserDefaultLCID(),
        DISPATCH_METHOD,
        &param,
        &varres,//*
        NULL,
        NULL);
    if (FAILED(hr))
    {
        printf("Invoke call failed.Code 0x8%X \n", hr);
        pIDispatch->Release();
        return 1;
    }

    if (varres.vt == VT_I4)
    {
        printf("Returned from component: op1 %d + op2 %d = %d\n",
            vargs[1].lVal, vargs[0].lVal, varres.lVal);
    }

    
    pIDispatch->Release();
    OleUninitialize();
    return 0;

}