#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"


#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

using namespace std;

void trace(const char *pMsg){ cout << "ClassFactory : Client : " << pMsg << endl;}

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

HMODULE ghModule = NULL;
HRESULT hr;

ISum *pISum = NULL;
ISub *pISubtract = NULL;

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

//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Entered"), TEXT("ClassFactory"), NULL);
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\01_ClassFactory\\Component\\ClassFactoryComponent.dll"));	
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : LoadLibrary : Error"), TEXT("ClassFactory : Error"), NULL);
		exit(0);
	}

	FuncPtr ServerRegistration;
	ServerRegistration = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");	
//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Registering Server"), TEXT("ClassFactory"), NULL);	
	if (ServerRegistration() == S_OK)
	{
//		MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Server Registered"), TEXT("ClassFactory"), NULL);			
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Class Factory Initialization Failed"), TEXT("Error"), NULL);
		exit(0);
	}

//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Calling CoInitialize"), TEXT("ClassFactory"), NULL);			
	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("COM library cannot be initialized.\nProgram will now exit ..."), TEXT("Program Error"), MB_OK);
		exit(0);
	}	
//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : CoInitialize succesful"), TEXT("ClassFactory"), NULL);			

	InitializeDefaultInterface();

//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Leaving"), TEXT("Leaving"), NULL);

}

void UnInitialize(void)
{
	void SafeInterfaceRelease(void);
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\01_ClassFactory\\Component\\ClassFactoryComponent.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("ClassFactory Client : UnInitialize : LoadLibrary : Error"), TEXT("Error"), NULL);
	FuncPtr ServerUnregistration;
	ServerUnregistration = (FuncPtr)GetProcAddress(hLib, "DllUnregisterServer");	
	if (ServerUnregistration() == S_OK)
	{
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Class Factory UnInitialization Failed"), TEXT("Error"), NULL);
	}

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
	MessageBox(NULL, str, TEXT("ClassFactory : Result"), MB_OK);

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

	hr = pISum -> QueryInterface(IID_ISub, (void **)&pISubtract);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("ISubtract, Interface cannot be obtained"), TEXT("Error"), MB_OK);
		SafeInterfaceRelease();
		return 0;
	}

	pISubtract -> SubtractionOfTwoIntegers(iNum1, iNum2, &iSub);
	pISubtract -> Release();
	pISubtract = NULL;

	wsprintf(str, TEXT("Subtraction of %d and %d = %d"), iNum1, iNum2, iSub);
	MessageBox(NULL, str, TEXT("ClassFactory : Result"), MB_OK);

	return iSub;
}

void InitializeDefaultInterface()
{
	void SafeInterfaceRelease(void);

	HRESULT hr;

	hr = CoCreateInstance(CLSID_Math, NULL, CLSCTX_INPROC_SERVER, IID_ISum, (void **)&pISum);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Default interface cannot be obtained"), TEXT("ClassFactory : Error"), MB_OK);
		SafeInterfaceRelease();
	}
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