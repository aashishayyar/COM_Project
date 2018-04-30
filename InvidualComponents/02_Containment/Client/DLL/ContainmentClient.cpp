#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

using namespace std;

void trace(const char *pMsg){ cout << "Containment : Client : " << pMsg << endl;}

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

HMODULE ghModule = NULL;
HRESULT hr;

ISum 			*pISum 				= NULL;
ISubtract		*pISubtract			= NULL;
IMultiplication *pIMultiplication 	= NULL;
IDivision 		*pIDivision			= NULL;

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
	void InitializeDefaultInterface(void);

	//	MessageBox(NULL, TEXT("ClassFactoryClient : Initialize : Entered"), TEXT("Containment"), NULL);
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

// 	INNER COMPONENT
	hLib = LoadLibrary(TEXT("..\\02_Containment\\Component\\InnerComponent\\ContainmentInnerComponent.dll"));	
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("ContainmentClient : Initialize : LoadLibrary : Error"), TEXT("Containment : Error"), NULL);
		exit(0);
	}

	FuncPtr ServerRegistration;
	ServerRegistration = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");	
	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Registering Server"), TEXT("Containment"), NULL);	
	if (ServerRegistration() == S_OK)
	{
	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Server Registered"), TEXT("Containment"), NULL);			
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Containment Initialization Failed"), TEXT("Error"), NULL);
		exit(0);
	}

// 	OUTER COMPONENT
	hLib = LoadLibrary(TEXT("..\\02_Containment\\Component\\OuterComponent\\ContainmentOuterComponent.dll"));	
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("ContainmentClient : Initialize : LoadLibrary : Error"), TEXT("Containment : Error"), NULL);
		exit(0);
	}

	ServerRegistration = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");	
	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Registering Server"), TEXT("Containment"), NULL);	
	if (ServerRegistration() == S_OK)
	{
	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Server Registered"), TEXT("Containment"), NULL);			
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Containment Initialization Failed"), TEXT("Error"), NULL);
		exit(0);
	}


	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Calling CoInitialize"), TEXT("Containment"), NULL);			
	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("COM library cannot be initialized.\nProgram will now exit ..."), TEXT("Program Error"), MB_OK);
		exit(0);
	}	
	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : CoInitialize succesful"), TEXT("Containment"), NULL);			

	//	MessageBox(NULL, TEXT("ContainmentClient : Initialize : Leaving"), TEXT("Leaving"), NULL);

}

void UnInitialize(void)
{
	void SafeInterfaceRelease(void);
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

//	INNER CLASS
	hLib = LoadLibrary(TEXT("..\\02_Containment\\Component\\InnerComponent\\ContainmentInnerComponent.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("Containment Client : UnInitialize : LoadLibrary : Error"), TEXT("Containment : Error"), NULL);
	FuncPtr ServerUnregistration;
	ServerUnregistration = (FuncPtr)GetProcAddress(hLib, "DllUnregisterServer");	
	if (ServerUnregistration() == S_OK)
	{
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Containment UnInitialization Failed"), TEXT("Containment : Error"), NULL);
	}

//  OUTER CLASS
	hLib = LoadLibrary(TEXT("..\\02_Containment\\Component\\OuterComponent\\ContainmentOuterComponent.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("Containment Client : UnInitialize : LoadLibrary : Error"), TEXT("Containment : Error"), NULL);
	ServerUnregistration = (FuncPtr)GetProcAddress(hLib, "DllUnregisterServer");	
	if (ServerUnregistration() == S_OK)
	{
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Containment UnInitialization Failed"), TEXT("Containment : Error"), NULL);
	}

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
	MessageBox(NULL, str, TEXT("Containment : Result"), MB_OK);

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
	MessageBox(NULL, str, TEXT("Containment : Result"), MB_OK);

	return iSub;
}

int MultiplicationOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);
	void SafeInterfaceRelease(void);

	int iMul;
	TCHAR str[255];
	
	if (pISum == NULL)
		InitializeDefaultInterface();

	hr = pISum -> QueryInterface(IID_IMultiplication, (void **)&pIMultiplication);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IMultiplication Interface cannot be obtained"), TEXT("Error"), MB_OK);
		SafeInterfaceRelease();
		return 0;
	}

	pIMultiplication -> MultiplicationOfTwoIntegers(iNum1, iNum2, &iMul);
	pIMultiplication -> Release();
	pIMultiplication = NULL;

	wsprintf(str, TEXT("Multiplication of %d and %d = %d"), iNum1, iNum2, iMul);
	MessageBox(NULL, str, TEXT("Containment : Result"), MB_OK);

	return iMul;
}

int DivisionOfTwoIntegers(int iNum1, int iNum2)
{
	void InitializeDefaultInterface(void);
	void SafeInterfaceRelease(void);

	int iDiv;
	TCHAR str[255];
	
	if (pISum == NULL)
		InitializeDefaultInterface();

	hr = pISum -> QueryInterface(IID_IDivision, (void **)&pIDivision);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IDivision Interface cannot be obtained"), TEXT("Error"), MB_OK);
		SafeInterfaceRelease();
		return 0;
	}

	pIDivision -> DivisionOfTwoIntegers(iNum1, iNum2, &iDiv);
	pIDivision -> Release();
	pIDivision = NULL;

	wsprintf(str, TEXT("Division of %d and %d = %d"), iNum1, iNum2, iDiv);
	MessageBox(NULL, str, TEXT("Containment : Result"), MB_OK);

	return iDiv;
}

void InitializeDefaultInterface()
{
	void SafeInterfaceRelease(void);

	HRESULT hr;

	hr = CoCreateInstance(CLSID_SumSubtract, NULL, CLSCTX_INPROC_SERVER, IID_ISum, (void **)&pISum);
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

	if (pIMultiplication)
	{
		pIMultiplication -> Release();
		pIMultiplication = NULL;
	}

	if (pIDivision)
	{
		pIDivision -> Release();
		pIDivision = NULL;
	}

}