#define UNICODE
#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

using namespace std;

void trace(const char *pMsg){ cout << "Moniker : Client : " << pMsg << endl;}

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

HMODULE ghModule = NULL;
HRESULT hr;

IOddNumber 			*pIOddNumber 		= NULL;
IOddNumberFactory 	*pIOddNumberFactory = NULL;

int       iFirstOddNumber;
int 	  iNextOddNumber;

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
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\06_Moniker\\Component\\MonikerServer.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("Moniker Client : Initialize : LoadLibrary : Error"), TEXT("Error"), NULL);
	FuncPtr ServerRegistration;
	ServerRegistration = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");	
	if (ServerRegistration() == S_OK)
	{
		FreeLibrary(hLib);
	}
	else 
	{
		MessageBox(NULL, TEXT("Moniker Initialization Failed"), TEXT("Moniker : Error"), NULL);
	}

	InitializeDefaultInterface();
}

void UnInitialize(void)
{
	void SafeInterfaceRelease(void);
	
	HINSTANCE hLib;
	typedef HRESULT(*FuncPtr)(void);

	hLib = LoadLibrary(TEXT("..\\06_Moniker\\Component\\MonikerServer.dll"));	
	if (hLib == NULL)
		MessageBox(NULL, TEXT("ClassFactory Client : UnInitialize : LoadLibrary : Error"), TEXT("Moniker : Error"), NULL);
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


int GetNextOddNumber(int iNum)
{
	void InitializeDefaultInterface(void);
	void SafeInterfaceRelease(void);
	TCHAR str[255];

	iFirstOddNumber = iNum;
	
	if (pIOddNumberFactory == NULL)
	InitializeDefaultInterface();

	hr = pIOddNumberFactory -> SetFirstOddNumber(iNum, &pIOddNumber);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Cannot obtaing IOddNumber"), TEXT("Error"), NULL);
		SafeInterfaceRelease();
		exit(0);
	}

	pIOddNumber -> GetNextOddNumber(&iNextOddNumber);

	wsprintf(str, TEXT("The next odd number from %2d is %2d"), iFirstOddNumber, iNextOddNumber);
	MessageBox(NULL, str, TEXT("Moniker : Result"), NULL);

	return iNextOddNumber;
}

void InitializeDefaultInterface()
{
	IBindCtx *pIBindCtx = NULL;
	IMoniker *pIMoniker = NULL;
	ULONG 	  uEaten;
	HRESULT   hr;
	LPOLESTR  szCLSID = NULL;
	wchar_t   wszCLSID[255];
	wchar_t   wszTemp[255];
	wchar_t  *ptr;
	TCHAR 	  temp;

	hr = CreateBindCtx(0, &pIBindCtx);
	if (hr != S_OK)	
	{
		MessageBox(NULL, TEXT("Failed to get IBindCtx Interface pointer"), TEXT("Error"), NULL);
		exit(0);	
	}	

	StringFromCLSID(CLSID_OddNumber, &szCLSID);

	wcscpy_s(wszTemp, szCLSID);
	ptr = wcschr(wszTemp, '{');
	ptr = ptr + 1;
	wcscpy_s(wszTemp, ptr);

	wszTemp[(int)wcslen(wszTemp) - 1] = '\0';

	wsprintf(wszCLSID, TEXT("clsid:%s"), wszTemp);

	hr = MkParseDisplayName(pIBindCtx, wszCLSID, &uEaten, &pIMoniker);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Failed to get IMoniker interface"), TEXT("Error"), NULL);
		pIBindCtx -> Release();
		pIBindCtx = NULL;

		exit(0);
	}	
	
	hr = pIMoniker -> BindToObject(pIBindCtx, NULL, IID_IOddNumberFactory, (void **)&pIOddNumberFactory);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Failed to get custom activation : IOddNumberFactory"), TEXT("Error"), NULL);
	
		pIMoniker -> Release();
		pIMoniker = NULL;

		pIBindCtx -> Release();
		pIBindCtx = NULL;

		exit(0);	
	}

	pIMoniker -> Release();
	pIMoniker = NULL;
	pIBindCtx -> Release();
	pIBindCtx = NULL;
}

void SafeInterfaceRelease(void)
{
	if (pIOddNumberFactory)
	{
		pIOddNumberFactory -> Release();
		pIOddNumberFactory = NULL;
	}
	
	if (pIOddNumber) 
	{
		pIOddNumber -> Release();
		pIOddNumber = NULL;
	}
}