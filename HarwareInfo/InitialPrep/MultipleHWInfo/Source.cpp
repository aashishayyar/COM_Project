#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char **argv)
{
    HRESULT hres;

    // Step 1: --------------------------------------------------
    // Initialize COM. ------------------------------------------

    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 
    if (FAILED(hres))
    {
        cout << "Failed to initialize COM library. Error code = 0x" 
            << hex << hres << endl;
        return 1;                  // Program has failed.
    }

    // Step 2: --------------------------------------------------
    // Set general COM security levels --------------------------

    hres =  CoInitializeSecurity(
        NULL, 
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
        );

                      
    if (FAILED(hres))
    {
        cout << "Failed to initialize security. Error code = 0x" 
            << hex << hres << endl;
        CoUninitialize();
        return 1;                    // Program has failed.
    }
    
    // Step 3: ---------------------------------------------------
    // Obtain the initial locator to WMI -------------------------

    IWbemLocator *pLoc = NULL;

    hres = CoCreateInstance(
        CLSID_WbemLocator,             
        0, 
        CLSCTX_INPROC_SERVER, 
        IID_IWbemLocator, (LPVOID *) &pLoc);
 
    if (FAILED(hres))
    {
        cout << "Failed to create IWbemLocator object."
            << " Err code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        return 1;                 // Program has failed.
    }

    // Step 4: -----------------------------------------------------
    // Connect to WMI through the IWbemLocator::ConnectServer method

    IWbemServices *pSvc = NULL;
 
    // Connect to the root\cimv2 namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
    hres = pLoc->ConnectServer(
         _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
         NULL,                    // User name. NULL = current user
         NULL,                    // User password. NULL = current
         0,                       // Locale. NULL indicates current
         NULL,                    // Security flags.
         0,                       // Authority (for example, Kerberos)
         0,                       // Context object 
         &pSvc                    // pointer to IWbemServices proxy
         );
    
    if (FAILED(hres))
    {
        cout << "Could not connect. Error code = 0x" 
             << hex << hres << endl;
        pLoc->Release();     
        CoUninitialize();
        return 1;                // Program has failed.
    }

    cout << "Processor Information : " << endl;


    // Step 5: --------------------------------------------------
    // Set security levels on the proxy -------------------------

    hres = CoSetProxyBlanket(
       pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name 
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres))
    {
        cout << "Could not set proxy blanket. Error code = 0x" 
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();     
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 6: --------------------------------------------------
    // Use the IWbemServices pointer to make requests of WMI ----

    // For example, get the name of the operating system
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
            << " Error code = 0x" 
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;               // Program has failed.
    }

    // Step 7: -------------------------------------------------
    // Get the data from the query in step 6 -------------------
 
    IWbemClassObject *pclsObj = NULL;
    ULONG uReturn = 0;
    int index = 0;

    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " Processor Name\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        wcout << " Caption\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"L2CacheSize", 0, &vtProp, 0, 0);
        wcout << " L2CacheSize\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"L2CacheSpeed", 0, &vtProp, 0, 0);
        wcout << " L2CacheSpeed\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"L3CacheSize", 0, &vtProp, 0, 0);
        wcout << " L3CacheSize\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"L3CacheSpeed", 0, &vtProp, 0, 0);
        wcout << " L3CacheSpeed\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        wcout << " NumberOfCores\t\t\t:\t" << vtProp.uintVal << endl;
       
        hr = pclsObj->Get(L"NumberOfLogicalProcessors", 0, &vtProp, 0, 0);
        wcout << " NumberOfLogicalProcessors\t:\t" << vtProp.uintVal << endl;

        VariantClear(&vtProp);
        pclsObj->Release();
    }

    wcout << endl ;

    cout << "Motherboard information" << endl << endl;

    pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_MotherboardDevice"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
            << " Error code = 0x" 
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    pclsObj = NULL;
    uReturn = 0;
   
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " Name\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"Availability", 0, &vtProp, 0, 0);
        wcout << " Availability\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"PowerManagementCapabilities", 0, &vtProp, 0, 0);
        wcout << " PowerManagementCapabilities\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        wcout << " PNPDeviceID\t\t\t:\t" << vtProp.bstrVal << endl;

       hr = pclsObj->Get(L"PrimaryBusType", 0, &vtProp, 0, 0);
        wcout << " PrimaryBusType\t\t\t:\t" << vtProp.bstrVal << endl;


        VariantClear(&vtProp);
        pclsObj->Release();
    }

    wcout << endl ;

    cout << "Network information : " << endl << endl;

    pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_NetworkAdapter"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
            << " Error code = 0x" 
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    pclsObj = NULL;
    uReturn = 0;
   
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " Name\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"DeviceID", 0, &vtProp, 0, 0);
        wcout << " DeviceID\t\t:\t" << vtProp.bstrVal << endl;


        VariantClear(&vtProp);
        pclsObj->Release();
    }

    wcout << endl;

    cout << "Sound information" << endl << endl;

    pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_SoundDevice"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
             << "Error code = 0x" 
             <<  hex << hres << endl;
        pSvc -> Release();
        pLoc -> Release();
        CoUninitialize();
        return 1;
    }

    pclsObj = NULL;
    uReturn = 0;
   
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn)
            break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " Name\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        wcout << " Manufacturer\t\t\t:\t" << vtProp.bstrVal << endl;

        VariantClear(&vtProp);
        pclsObj->Release();
    }

    wcout << endl ;

    cout << "VideoCard information" << endl << endl;

    pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"), 
        bstr_t("SELECT * FROM Win32_VideoController"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
        NULL,
        &pEnumerator);
    
    if (FAILED(hres))
    {
        cout << "Query for operating system name failed."
            << " Error code = 0x" 
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    pclsObj = NULL;
    uReturn = 0;
   
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, 
            &pclsObj, &uReturn);

        if(0 == uReturn)
        {
            break;
        }

        VARIANT vtProp;

        // Get the value of the Name property
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        wcout << " Name\t\t\t:\t" << vtProp.bstrVal << endl;

        hr = pclsObj->Get(L"Availability", 0, &vtProp, 0, 0);
        wcout << " Availability\t\t\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"PowerManagementCapabilities", 0, &vtProp, 0, 0);
        wcout << " PowerManagementCapabilities\t:\t" << vtProp.uintVal << endl;

        hr = pclsObj->Get(L"PNPDeviceID", 0, &vtProp, 0, 0);
        wcout << " PNPDeviceID\t\t\t:\t" << vtProp.bstrVal << endl;

       hr = pclsObj->Get(L"PrimaryBusType", 0, &vtProp, 0, 0);
        wcout << " PrimaryBusType\t\t\t:\t" << vtProp.bstrVal << endl;


        VariantClear(&vtProp);
        pclsObj->Release();
    }

    // Cleanup
    // ========
    
    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();

    CoUninitialize();

    return 0;   // Program successfully completed.
 
}
