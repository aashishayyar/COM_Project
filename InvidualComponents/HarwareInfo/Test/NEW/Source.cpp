#include <windows.h>
#include <iostream>

using namespace std;

// #pragma comment(lib, "wbeuuid.lib")

int main(int argc, char *argv[])
{
	HRESULT hr;

	hr = CoInitialize();
	if (FAILED(hr))
	{
		cout << "Failed to initialize COM library. Error code = 0x" << hex << hr << endl;
		return 1;
	}

	hr = CoInitializeSecurity(NULL, 						 
							  -1, 
							  NULL, 
							  NULL, 
							  RPC_C_AUTHN_LEVEL_DEFAULT, 
							  RPC_C_IMP_LEVEL_IMPERSONATE, 
							  NULL, 
							  EOAC_NONE, 
							  NULL );
	if (FAILED(hr))
	{
		cout << "Failed to initialize security. Error code = 0x" << hex << hr << endl;
		CoUninitialize();
		return 1;
	}

	IWebmLocator *pLoc = NULL;

	hr = CoCreateInstance(CLSID_WbemLocator, 
						  0, 
						  CLSCTX_INPROC_SERVER, 
						  IID_IWebmLocator, 
						  (LPVOID *)&pLoc);
	if (FAILED(hr))
	{
		cout << "Failed to create IWebmLocator object. " << "Error code = 0x" << hex << hr << endl;
		CoUninitialize();
		return 1;
	}

	IWebmServices *pSvc = NULL;

	hr = pLoc -> ConnecServer(_bstr_t(L"ROOT\\CIMV2"), 
							  NULL, 
							  NULL, 
							  0, 
							  NULL, 
							  0, 
							  0,
							  &pSvc );
	if (FAILED(hr))
	{
		cout << "Could not connect. Error code = 0x" << hex << hr << endl;
		pLoc -> Release();
		CoUninitialize();
		return 1;
	}

	cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

	hr = CoSetProxyBlancket(pSvc, 
							RPC_C_AUTHN_WINNT, 
							RPC_C_AUTHZ_NONE, 
							NULL, 
							RPC_C_AUTHN_LEVEL_CALL, 
							RPC_C_IMP_LEVEL_IMPERSONATE, 
							NULL, 
							EOAC_NONE );
	if (FAILED(hr))
	{
		cout << "Could not set proxy blanket. Error code 0x" << hex << hr << endl;
		pSvc -> Release();
		pLoc -> Release();
		CoUninitialize();
		return 1;
	}

	IEnumWebmClassObject *pEnumerator = NULL;
	hr = pSvc -> ExecQuery(bstr_t("WQL"), 
						   bstr_t("SELECT * FROM Win32_Processor"), 
						   WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
						   NULL, 
						   &pEnumerator);
	if (FAILED(hr))
	{
		cout << "Query failed. Error Code = 0x" << hex << hr << endl;
		pSvc -> Release();
		pLoc -> Release();
		CoUninitialize();
		return 1;
	}

	IWebmClassobject *pClsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		HRESULT hr = pEnumerator -> Next(WBEM_INFINITE, 1, &pClsObj, &uReturn);
		if (uReturn == 0)
			break;
		
		VARIANT vtProp;
		hr = pClsObj -> 
	}
}