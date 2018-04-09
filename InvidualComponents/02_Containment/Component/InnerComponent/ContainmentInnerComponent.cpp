#include <windows.h>
#include "ContainmentInnerComponent.h"
#include "registry.h"

long glNumberOfActiveComponents = 0;
long glNumberOfServerLocks		= 0;

HMODULE ghModule;

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

class CMultiplicationDivision:public IMultiplication, IDivision
{
private:
		long m_cRef;

public:
		 CMultiplicationDivision(void);
		~CMultiplicationDivision(void);

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG   __stdcall AddRef(void);
		ULONG   __stdcall Release(void);

		HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int *);
		HRESULT __stdcall DivisionOfTwoIntegers(int, int, int *);
};

class ClassFactory:public IClassFactory
{
private:
		long m_cRef;

public:
		 ClassFactory(void);
		~ClassFactory(void);
	 
		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG   __stdcall AddRef(void);
		ULONG   __stdcall Release(void);

		HRESULT __stdcall CreateInstance(IUnknown *, REFIID, void **);
		HRESULT __stdcall LockServer(BOOL);	
};

BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
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
			break;		
	}
	return(TRUE);
}

CMultiplicationDivision::CMultiplicationDivision(void)
{
	m_cRef = 1;
	InterlockedIncrement(&glNumberOfActiveComponents);
}

CMultiplicationDivision::~CMultiplicationDivision(void)
{
	InterlockedDecrement(&glNumberOfActiveComponents);
}

HRESULT CMultiplicationDivision::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IMultiplication *>(this);
	else if (riid == IID_IMultiplication)
		*ppv = static_cast<IMultiplication *>(this);
	else if (riid == IID_IDivision)
		*ppv = static_cast<IDivision *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMultiplicationDivision::AddRef(void)
{	
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMultiplicationDivision::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

HRESULT CMultiplicationDivision::MultiplicationOfTwoIntegers(int iNum1, int iNum2, int *pMultiplication)
{
	*pMultiplication = iNum1 * iNum2;
	return(S_OK);
}

HRESULT CMultiplicationDivision::DivisionOfTwoIntegers(int iNum1, int iNum2, int *pDivision)
{
	*pDivision = iNum1 / iNum2;
	return (S_OK);
}

ClassFactory::ClassFactory(void)
{
	m_cRef =1;
}

ClassFactory::~ClassFactory(void)
{

}

HRESULT ClassFactory::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IClassFactory *>(this);
	else if (riid == IID_IClassFactory)
		*ppv = static_cast<IClassFactory *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}		
	reinterpret_cast<IUnknown *>(*ppv)->AddRef();
	return(S_OK);	
}

ULONG ClassFactory::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG ClassFactory::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
		delete(this);
	return(m_cRef);
}

HRESULT ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
	CMultiplicationDivision *pCMultiplicationDivision = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);

	pCMultiplicationDivision = new CMultiplicationDivision;
	if (pCMultiplicationDivision == NULL)
		return(E_OUTOFMEMORY);

	hr = pCMultiplicationDivision -> QueryInterface(riid, ppv);
	pCMultiplicationDivision -> Release();

	return(hr);	
}

HRESULT ClassFactory::LockServer(BOOL fLock)
{
	if (fLock)
		InterlockedIncrement(&glNumberOfServerLocks);
	else 
		InterlockedDecrement(&glNumberOfServerLocks);

	return(S_OK);
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	ClassFactory *pClassFactory = NULL;
	HRESULT hr;

	if (rclsid != CLSID_MultiplicationDivision)
		return(CLASS_E_CLASSNOTAVAILABLE);

	pClassFactory = new ClassFactory;
	if (pClassFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pClassFactory -> QueryInterface(riid, ppv);
	pClassFactory -> Release();

	return(hr);
}

HRESULT __stdcall DllCanUnloadNow(void)
{
	if ((glNumberOfActiveComponents == 0) && (glNumberOfServerLocks == 0))
		return(S_OK);
	else 
		return(S_FALSE);
}	

STDAPI DllRegisterServer()
{
	return RegisterServer(ghModule,
						  CLSID_MultiplicationDivision,
						  g_szFriendlyName,
						  g_szVerIndProgId,
						  g_szProgId);
/*
	HKEY hCLSIDKey = NULL;
	HKEY hInProcServerKey = NULL;
	LONG lRet;
	TCHAR szModulePath[MAX_PATH];
	TCHAR szClassDescription[] = TEXT(" Containment Inner Component Class");

	lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, TEXT("CLSID\\{A556F1B8-B1BA-4BB4-9AE5-5902EEBF9A64}"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hCLSIDKey, NULL); 
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	lRet = RegSetValueEx(hCLSIDKey, NULL, 0, REG_SZ, (const BYTE *)szClassDescription, sizeof(szClassDescription));
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	lRet = RegCreateKeyEx(hCLSIDKey, TEXT("InProcServer32"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hInProcServerKey, NULL);
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	GetModuleFileName(ghModule, szModulePath, MAX_PATH);

	lRet = RegSetValueEx(hInProcServerKey, NULL, 0, REG_SZ, (const BYTE *)szModulePath, sizeof(TCHAR) * (lstrlen(szModulePath) + 1));
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	if (hCLSIDKey != NULL)
		RegCloseKey(hCLSIDKey);
	if (hInProcServerKey != NULL)
		RegCloseKey(hInProcServerKey);

	return(S_OK);
*/
}

STDAPI DllUnRegisterServer()
{

	return UnregisterServer(CLSID_MultiplicationDivision,
							g_szVerIndProgId,
							g_szProgId);
/*
	RegDeleteKey(HKEY_CLASSES_ROOT, TEXT("CLSID\\{A556F1B8-B1BA-4BB4-9AE5-5902EEBF9A64}\\InProcServer32"));
	RegDeleteKey(HKEY_CLASSES_ROOT, TEXT("CLSID\\{A556F1B8-B1BA-4BB4-9AE5-5902EEBF9A64}"));

	return(S_OK);
*/
}
