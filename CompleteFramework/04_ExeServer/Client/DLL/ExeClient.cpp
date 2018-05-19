#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

using namespace std;

void trace(const char *pMsg){ cout << "ExeClient : Client : " << pMsg << endl;}

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

HMODULE ghModule = NULL;
HRESULT hr;

ISum 			*pISum 				= NULL;
ISubtract		*pISubtract			= NULL;

BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
	void	Initialize(void);
	void 	UnInitialize(void);
	void 	SafeInterfaceRelease(void);

	switch(dwReason)
	{
		case DLL_PROCESS_ATTACH:
				ghModule = hDll;
				Initialize();
			break;
		case DLL_THREAD_ATTACH:
			break;
		case DLL_THREAD_DETACH:
			break; 	
		case DLL_PROCESS_DETACH:
				CoUninitialize();
				UnInitialize();
				SafeInterfaceRelease();
			break;		
	}
	return(TRUE);
}

void Initialize(void)
{
	void RegisterProxyStubDLL(void);
	void InitializeDefaultInterface(void);

	// MessageBox(NULL, TEXT("ExeClient : Initialize : Entered"), TEXT("ExeServer : Log"), NULL);
	// MessageBox(NULL, TEXT("ExeClient : Initialize : Calling Register ProxyStubDll"), TEXT("ExeServer : Log"), NULL);

	RegisterProxyStubDLL();

	// MessageBox(NULL, TEXT("ExeClient : Initialize : Registered ProxyStubDll"), TEXT("ExeServer : Log"), NULL);

	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	STARTUPINFO si;
 	ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);	

    PROCESS_INFORMATION pInfo;

	// MessageBox(NULL, TEXT("ExeClient : Initialize : Registering ExeServer"), TEXT("ExeServer : Log"), NULL);

	LPTSTR szCmdline = _strdup(TEXT("..\\04_ExeServer\\Component\\ExeServer.exe register"));

	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("COM library cannot be initialized.\nProgram will now exit ..."), TEXT("Program Error"), MB_OK);
		exit(0);
	}	

	if (!CreateProcess(NULL, 
					  szCmdline, 
					  NULL, NULL, 
					  FALSE, 
					  NORMAL_PRIORITY_CLASS, 
					  NULL, 
					  NULL,
					  &si, 
					  &pInfo) )
	{
		MessageBox(NULL, TEXT("CreateProcess Failed."), TEXT("ExeServer"), NULL);
		exit(0);
	}

	WaitForSingleObject( pInfo.hProcess, INFINITE);
    CloseHandle( pInfo.hProcess );
}

void UnInitialize(void)
{
	void UnRegisterProxyStubDLL(void);
	void SafeInterfaceRelease(void);
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	STARTUPINFO si;
 	ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);	

    PROCESS_INFORMATION pInfo;

	if (!CreateProcess("..\\04_ExeServer\\Component\\ExeServer.exe", 
				  "unregister", 
				  NULL, NULL, 
				  FALSE, 
				  NORMAL_PRIORITY_CLASS, 
				  NULL, 
				  NULL,
				  &si, 
				  &pInfo) )
	{
		MessageBox(NULL, TEXT("CreateProcess Failed."), TEXT("ExeServer"), NULL);
		exit(0);
	}

	WaitForSingleObject( pInfo.hProcess, INFINITE);
    CloseHandle( pInfo.hProcess );

	CoUninitialize();

	UnRegisterProxyStubDLL();

	SafeInterfaceRelease();	
}

int SumOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);

	int iSum;
	TCHAR str[255];
	
	if (pISum == NULL)
		InitializeDefaultInterface();

	pISum -> SumOfTwoIntegers(iNum1, iNum2, &iSum);
	wsprintf(str, TEXT("Sum of %d and %d = %d"), iNum1, iNum2, iSum);
	MessageBox(NULL, str, TEXT("ExeServer : Result"), MB_OK);

	return iSum;
}

int SubtractionOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);
	void SafeInterfaceRelease(void);

	int iSub;
	TCHAR str[255];
	
	if (pISum == NULL)
		InitializeDefaultInterface();

	hr = pISum -> QueryInterface(IID_ISubtract, (void **)&pISubtract);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("ISubtract Interface cannot be obtained"), TEXT("Error"), MB_OK);
		SafeInterfaceRelease();
		return 0;
	}

	pISubtract -> SubtractionOfTwoIntegers(iNum1, iNum2, &iSub);
	pISubtract -> Release();
	pISubtract = NULL;

	wsprintf(str, TEXT("Subtraction of %d and %d = %d"), iNum1, iNum2, iSub);
	MessageBox(NULL, str, TEXT("ExeServer : Result"), MB_OK);

	return iSub;
}

void InitializeDefaultInterface()
{
	void SafeInterfaceRelease(void);

	HRESULT hr;

	// MessageBox(NULL, TEXT("ExeClient : InitializeDefaultInterface : Entered : Calling CoCreateInstance"), TEXT("ExeServer : Log"), NULL);
	hr = CoCreateInstance(CLSID_SumSubtract, NULL, CLSCTX_LOCAL_SERVER, IID_ISum, (void **)&pISum);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Default interface cannot be obtained"), TEXT("ExeServer : Error"), MB_OK);
		SafeInterfaceRelease();
	}
	// MessageBox(NULL, TEXT("ExeClient : InitializeDefaultInterface : CoCreateInstance : SUCCESS"), TEXT("ExeServer : Log"), NULL);
}

void RegisterProxyStubDLL()
{
	HINSTANCE hLib;
	
	typedef HRESULT (*FuncPtr)(void);
	FuncPtr RegisterServer;

	hLib = LoadLibrary(TEXT("..\\04_ExeServer\\Component\\ProxyStub\\ProxyStub.dll"));
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("Exe Server Client : RegisterProxyStubDLL : Cannot register ProxyStubDLL"), 
					TEXT("Error"), NULL);
		exit(0);
	}

	RegisterServer = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");

	RegisterServer();

	FreeLibrary(hLib);
}

void UnRegisterProxyStubDLL()
{
	HINSTANCE hLib;
	
	typedef HRESULT (*FuncPtr)(void);
	FuncPtr UnRegisterServer;

	hLib = LoadLibrary(TEXT("..\\04_ExeServer\\Component\\ProxyStub\\ProxyStub.dll"));
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("Exe Server Client : UnRegisterProxyStubDLL : Cannot unregister ProxyStubDLL"), 
					TEXT("Error"), NULL);
		exit(0);
	}

	UnRegisterServer = (FuncPtr)GetProcAddress(hLib, "DllUnRegisterServer");

	UnRegisterServer();

	FreeLibrary(hLib);
}

void SafeInterfaceRelease(void)
{
	if (pISum)
	{
		pISum -> Release();
		pISum = NULL;
	}
	
	if (pISubtract) 
	{
		pISubtract -> Release();
		pISubtract = NULL;
	}

}