#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "advapi32.lib")

using namespace std;

void trace(const char *pMsg){ cout << "ClassFactory : Client : " << pMsg << endl;}

HMODULE ghModule = NULL;
HRESULT hr;

IDispatch  *pIDispatch = NULL;


BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
	void	InitializeDefaultInterface(void);
	void 	SafeInterfaceRelease(void);

	switch(dwReason)
	{
		case DLL_PROCESS_ATTACH:
				ghModule = hDll;
			break;
		case DLL_THREAD_ATTACH:
			break;
		case DLL_THREAD_DETACH:
			break; 	
		case DLL_PROCESS_DETACH:
				CoUninitialize();
			break;		
	}
	return(TRUE);
}

void Initialize(void)
{
	void InitializeDefaultInterface(void);

//	MessageBox(NULL, TEXT("AutomationClient : Initialize : Entered"), TEXT("Automation"), NULL);
/*	STARTUPINFO si;
 	ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);	

    PROCESS_INFORMATION pInfo;

	LPTSTR szCmdline = _strdup(TEXT("regasm AutomationProxyStub.dll /tlb:AutomationTypeLib.tlb"));
	if (!CreateProcess(NULL, 
					   szCmdline, 
					   NULL, NULL, 
					   FALSE, 
					   NORMAL_PRIORITY_CLASS, 
					   NULL, 
					   TEXT("..\\05_Automation\\Component\\ProxyStub"), 
					   &si, 
					   &pInfo) )
	{
		MessageBox(NULL, TEXT("Automation Client : Initialize : TLB Registration failed"), TEXT("Automation : Error"), NULL);
		exit(0);
	}

	WaitForSingleObject( pInfo.hProcess, INFINITE);
    CloseHandle( pInfo.hProcess );
*/
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\05_Automation\\Component\\AutomationServer.dll"));	
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("AutomationClient : Initialize : LoadLibrary : Error"), TEXT("Automation : Error"), NULL);
		exit(0);
	}

	FuncPtr ServerRegistration;
	ServerRegistration = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");	
//	MessageBox(NULL, TEXT("AutomationClient : Initialize : Registering Server"), TEXT("Automation"), NULL);	
	if (ServerRegistration() == S_OK)
	{
//		MessageBox(NULL, TEXT("AutomationClient : Initialize : Server Registered"), TEXT("Automation"), NULL);			
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Automation Initialization Failed"), TEXT("Error"), NULL);
		exit(0);
	}

//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Calling CoInitialize"), TEXT("Automation"), NULL);			
	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("COM library cannot be initialized.\nProgram will now exit ..."), TEXT("Program Error"), MB_OK);
		exit(0);
	}	
	// MessageBox(NULL, TEXT("AutomationClient : Initialize : CoInitialize succesful"), TEXT("Automation"), NULL);			

	InitializeDefaultInterface();

//	MessageBox(NULL, TEXT("AutomationClient : Initialize : Leaving"), TEXT("Leaving"), NULL);

}

void UnInitialize(void)
{
	void SafeInterfaceRelease(void);
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\05_Automation\\Component\\AutomationServer.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("Automation Client : UnInitialize : LoadLibrary : Error"), TEXT("Error"), NULL);
	FuncPtr ServerUnregistration;
	ServerUnregistration = (FuncPtr)GetProcAddress(hLib, "DllUnregisterServer");	
	if (ServerUnregistration() == S_OK)
	{
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Automation UnInitialization Failed"), TEXT("Error"), NULL);
	}

	SafeInterfaceRelease();	
}

int SumOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);
	void ComErrorDescriptionString(HRESULT);
	void SafeInterfaceRelease(void);

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : Enetered"), TEXT("Automation"), NULL);

	TCHAR   	str[255];
	DISPID 		dispid;
	VARIANT		vArg[2], vRet;	
	DISPPARAMS	param = { vArg, 0, 2, NULL};

	OLECHAR *SumOfTwoIntegers = L"SumOfTwoIntegers";

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : HERE1"), TEXT("Automation"), NULL);


	if (!pIDispatch)
		InitializeDefaultInterface();

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : HERE2"), TEXT("Automation"), NULL);

	VariantInit(vArg);
	vArg[0].vt 		= VT_INT;
	vArg[0].intVal 	= iNum2;
	vArg[1].vt 		= VT_INT;
	vArg[1].intVal	= iNum1;

	param.cArgs 	 = 2;
	param.cNamedArgs = 0;
	param.rgdispidNamedArgs = NULL;
	param.rgvarg      = vArg;

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : HERE3"), TEXT("Automation"), NULL);


	VariantInit(&vRet);

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : HERE4"), TEXT("Automation"), NULL);
	hr = pIDispatch -> GetIDsOfNames(IID_NULL, 
									 &SumOfTwoIntegers, 
									 1, GetUserDefaultLCID(), 
									 &dispid);
	if (FAILED(hr))
	{
		ComErrorDescriptionString(hr);
		MessageBox(NULL, TEXT("Cannot get ID for SumOfTwoIntegers()"), 
				   TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
		SafeInterfaceRelease();
		exit(0);
	}

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : HERE5"), TEXT("Automation"), NULL);

	hr = pIDispatch -> Invoke(dispid, 
								IID_NULL, 
								GetUserDefaultLCID(), 
								DISPATCH_METHOD, 
								&param, 
								&vRet, 
								NULL, 
								NULL);
	if (FAILED(hr))
	{
		ComErrorDescriptionString(hr);
		MessageBox(NULL, TEXT("Cannot invoke function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
		SafeInterfaceRelease();

		exit(0);
	}								
	else 
	{
		wsprintf(str, TEXT("Sum of %d and %d is %d"), iNum1, iNum2, vRet.lVal);
		MessageBox(NULL, str, TEXT("Automation : Result"), MB_OK);
	}

	VariantClear(vArg);
	VariantClear(&vRet);

	// MessageBox(NULL, TEXT("SumOfTwoIntegers : Leaving"), TEXT("Automation"), NULL);


	return vRet.lVal;
}

int SubtractionOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);
	void ComErrorDescriptionString(HRESULT);
	void SafeInterfaceRelease(void);

	// MessageBox(NULL, TEXT("SubtractionofTwoIntegers : Entered"), TEXT("Automation"), NULL);

	TCHAR   	str[255];
	DISPID 		dispid;
	VARIANT		vArg[2], vRet;	
	DISPPARAMS	param = { vArg, 0, 2, NULL};

	OLECHAR    *SubtractionOfTwoIntegers = L"SubtractionofTwoIntegers";

	VariantInit(vArg);
	vArg[0].vt 		= VT_INT;
	vArg[0].intVal 	= iNum2;
	vArg[1].vt 		= VT_INT;
	vArg[1].intVal	= iNum1;

	param.cArgs 	 = 2;
	param.cNamedArgs = 0;
	param.rgdispidNamedArgs = NULL;
	param.rgvarg      = vArg;

	VariantInit(&vRet);
	hr = pIDispatch -> GetIDsOfNames(IID_NULL, 
									 &SubtractionOfTwoIntegers, 
									 1, GetUserDefaultLCID(), 
									 &dispid);
	if (FAILED(hr))
	{
		ComErrorDescriptionString(hr);
		MessageBox(NULL, TEXT("Cannot get ID for SumOfTwoIntegers()"), 
				   TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
		SafeInterfaceRelease();
		exit(0);
	}

	hr = pIDispatch -> Invoke(dispid, 
								IID_NULL, 
								GetUserDefaultLCID(), 
								DISPATCH_METHOD, 
								&param, 
								&vRet, 
								NULL, 
								NULL);
	if (FAILED(hr))
	{
		ComErrorDescriptionString(hr);
		MessageBox(NULL, TEXT("Cannot invoke function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
		SafeInterfaceRelease();

		exit(0);
	}								
	else 
	{
		wsprintf(str, TEXT("Subtraction of %d and %d is %d"), iNum1, iNum2, vRet.lVal);
		MessageBox(NULL, str, TEXT("Automation : Result"), MB_OK);
	}

	VariantClear(vArg);
	VariantClear(&vRet);

	// MessageBox(NULL, TEXT("SubtractionofTwoIntegers : Leaving"), TEXT("Automation"), NULL);

	return vRet.lVal;
}

void InitializeDefaultInterface()
{
	void SafeInterfaceRelease(void);
	void ComErrorDescriptionString(HRESULT);
	
	HRESULT hr;
	// MessageBox(NULL, TEXT("AutomationClient : InitializeDefaultInterface : CoCreateInstance : Calling"), TEXT("Automation : Error"), NULL);

	hr = CoCreateInstance(CLSID_MyMath, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&pIDispatch);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("AutomationClient : InitializeDefaultInterface : CoCreateInstance : Failed"), TEXT("Automation : Error"), NULL);
		ComErrorDescriptionString(hr);
		MessageBox(NULL, TEXT("Default interface cannot be obtained"), TEXT("Automation : Error"), MB_OK);
		SafeInterfaceRelease();
		exit(0);
	}
	// MessageBox(NULL, TEXT("AutomationClient : InitializeDefaultInterface : CoCreateInstance : SUCCESS"), TEXT("Automation : Error"), NULL);
}

void ComErrorDescriptionString(HRESULT hr)
{
	TCHAR *szErrorMessage = NULL;
	TCHAR  str[255];

	if (FACILITY_WINDOWS == HRESULT_FACILITY(hr))
		hr = HRESULT_CODE(hr);

	if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, 
					  NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
					  (LPTSTR)&szErrorMessage, 0, NULL) != 0)
	{
		sprintf_s(str, TEXT("%s"), szErrorMessage);
		LocalFree(szErrorMessage);
	}
	else 
		sprintf_s(str, TEXT("Could not find a description for error # %#x.\n"), hr);

	MessageBox(NULL, str, TEXT("COM Error"), MB_OK);
}

void SafeInterfaceRelease(void)
{
	if (pIDispatch)
	{
		pIDispatch -> Release();
		pIDispatch = NULL;
	}
}