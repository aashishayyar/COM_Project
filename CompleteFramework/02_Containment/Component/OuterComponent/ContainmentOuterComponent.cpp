#include <windows.h>
#include "ContainmentInnerComponent.h"
#include "ContainmentOuterComponent.h"
#include "registry.h"

long glNumberOfServerLocks;
long glNumberOfActiveComponents;
HMODULE ghModule = NULL;

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

class CSumSubtract:public ISum, ISubtract, IMultiplication, IDivision 
{
	private:
		long m_cRef;

		IMultiplication *m_pIMultiplication;
		IDivision 		*m_pIDivision;

	public:
		 CSumSubtract(void);
		~CSumSubtract(void);

		HRESULT __stdcall QueryInterface				(REFIID, void **); 
		ULONG 	__stdcall AddRef						(void);
		ULONG 	__stdcall Release						(void);

		HRESULT __stdcall SumOfTwoIntegers				(int, int, int *);
		HRESULT __stdcall SubtractionOfTwoIntegers		(int, int, int *);

		HRESULT __stdcall MultiplicationOfTwoIntegers	(int, int, int *);
		HRESULT __stdcall DivisionOfTwoIntegers			(int, int, int *);		
 	
 		HRESULT __stdcall InitializeInnerComponent		(void);
 };

 class ClassFactory:public IClassFactory
 {
 	private:
 		long m_cRef;

 	public:
 		 ClassFactory(void);
 		~ClassFactory(void);

 		HRESULT __stdcall QueryInterface				(REFIID, void **);
 		ULONG 	__stdcall AddRef						(void);
 		ULONG 	__stdcall Release						(void);

 		HRESULT __stdcall CreateInstance				(IUnknown *, REFIID, void **);
 		HRESULT __stdcall LockServer					(BOOL);
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

CSumSubtract::CSumSubtract(void)
{
	m_pIMultiplication  = NULL;
	m_pIDivision		= NULL;
	m_cRef 				= 1;

	InterlockedIncrement(&glNumberOfActiveComponents); 
}     

CSumSubtract::~CSumSubtract(void)
{
	InterlockedDecrement(&glNumberOfActiveComponents);

	if (m_pIMultiplication)
	{
		m_pIMultiplication -> Release();
		m_pIMultiplication = NULL;
	}

	if (m_pIDivision) 
	{	
		m_pIDivision -> Release();
		m_pIDivision = NULL;
	}
}

HRESULT CSumSubtract::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISum)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISubtract)
		*ppv = static_cast<ISubtract *>(this);
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

ULONG CSumSubtract::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	
	return(m_cRef);
}

ULONG CSumSubtract::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return(0);
	}

	return(m_cRef);
}

HRESULT CSumSubtract::SumOfTwoIntegers(int iNum1, int iNum2, int *pSum)
{
	*pSum = iNum1 + iNum2;
	
	return(S_OK);
}

HRESULT CSumSubtract::SubtractionOfTwoIntegers(int iNum1, int iNum2, int *pSubtract)
{
	*pSubtract = iNum1 - iNum2;
	
	return (S_OK);
}

HRESULT CSumSubtract::MultiplicationOfTwoIntegers(int iNum1, int iNum2, int *pMultiply)
{
	m_pIMultiplication -> MultiplicationOfTwoIntegers(iNum1, iNum2, pMultiply);

	return(S_OK);
}

HRESULT CSumSubtract::DivisionOfTwoIntegers(int iNum1, int iNum2, int *pDivision)
{
	m_pIDivision -> DivisionOfTwoIntegers(iNum1, iNum2, pDivision);

	return(S_OK);
}

HRESULT CSumSubtract::InitializeInnerComponent(void)
{
	HRESULT hr;

	hr = CoCreateInstance(CLSID_MultiplicationDivision, NULL, CLSCTX_INPROC_SERVER, IID_IMultiplication, (void **)&m_pIMultiplication);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IMultiplication interface cannot be obtained from Inner Component"), TEXT("Error"), MB_OK);
		return(E_FAIL);
	}

	hr = m_pIMultiplication -> QueryInterface(IID_IDivision, (void **)&m_pIDivision);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IDivision interface cannot be obtained from Inner Component"), TEXT("Error"), MB_OK);
		return(E_FAIL);		
	}

	return(S_OK);
}

ClassFactory::ClassFactory(void)
{
	m_cRef = 1;
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
	
	return (m_cRef);
}

ULONG ClassFactory::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
		delete(this);
	return (m_cRef);
}

HRESULT ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
	CSumSubtract *pCSumSubtract = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);

	pCSumSubtract = new CSumSubtract;
	if (pCSumSubtract == NULL)
		return(E_OUTOFMEMORY);

	hr = pCSumSubtract -> InitializeInnerComponent();
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Failed to Initialize Inner Component"), TEXT("Error"), MB_OK);
		pCSumSubtract -> Release();
		return (hr);
	}

	hr = pCSumSubtract -> QueryInterface(riid, ppv);
	pCSumSubtract -> Release();

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

	if (rclsid != CLSID_SumSubtract)
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
						  CLSID_SumSubtract,
						  g_szFriendlyName,
						  g_szVerIndProgId,
						  g_szProgId);
/*	
	HKEY hCLSIDKey = NULL;
	HKEY hInProcServerKey = NULL;
	LONG lRet;
	TCHAR szModulePath[MAX_PATH];
	TCHAR szClassDescription[] = TEXT("Containment Outer Component Class");

	lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, TEXT("CLSID\\{777A968C-D8A7-48A9-B027-BEC429268F4F}"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hCLSIDKey, NULL);
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	lRet = RegSetValueEx(hCLSIDKey, NULL, 0, REG_SZ, (const BYTE *)szClassDescription, sizeof(szClassDescription));
	if (lRet != ERROR_SUCCESS)	
		return(HRESULT_FROM_WIN32(lRet));

	lRet = RegCreateKeyEx(hCLSIDKey, TEXT("InProcServer32"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hInProcServerKey, NULL);
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	GetModuleFileName(ghModule, szModulePath, MAX_PATH);
	
	lRet = RegSetValueEx(hInProcServerKey, NULL, 0, REG_SZ, (const BYTE *)szModulePath, sizeof(TCHAR) * (lstrlen(szModulePath) + 1) );	
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));

	if (hCLSIDKey != NULL)
		RegCloseKey(hCLSIDKey);
	if (hInProcServerKey != NULL)
		RegCloseKey(hInProcServerKey);

	return(S_OK);
*/
}
/*
STDAPI DllRegisterServer()
{
	HKEY  hCLSIDKey 		= NULL;
	HKEY  hInProcServerKey 	= NULL;
	LONG  lRet;
	TCHAR szModulePath[MAX_PATH];
	TCHAR szClassDescription[] = TEXT("ClassFactory Component Class");

	lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, TEXT("CLSID\\{777A968C-D8A7-48A9-B027-BEC429268F4F}"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE | KEY_CREATE_SUB_KEY, NULL, &hCLSIDKey, NULL);
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));	

	lRet = RegSetValueEx(hCLSIDKey, NULL, 0, REG_SZ, (const BYTE *)szClassDescription, sizeof(szClassDescription));
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));	

	lRet = RegCreateKeyEx(hCLSIDKey, TEXT("InProcServer32"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hInProcServerKey, NULL);	
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));	

	GetModuleFileName(ghModule, szModulePath, MAX_PATH);

	lRet = RegSetValueEx(hInProcServerKey, NULL, 0, REG_SZ, (const BYTE*)szModulePath, sizeof(TCHAR) * (lstrlen(szModulePath) + 1) );
	if (lRet != ERROR_SUCCESS)
		return(HRESULT_FROM_WIN32(lRet));	

	if (hCLSIDKey != NULL)
		RegCloseKey(hCLSIDKey);
	if (hInProcServerKey != NULL)
		RegCloseKey(hInProcServerKey);

	return(S_OK);
}
*/

STDAPI DllUnregisterServer()
{
	return UnregisterServer(CLSID_SumSubtract,
							g_szVerIndProgId,
							g_szProgId);
/*	RegDeleteKey(HKEY_CLASSES_ROOT, TEXT("CLSID\\{777A968C-D8A7-48A9-B027-BEC429268F4F}\\InProcServer32"));
	RegDeleteKey(HKEY_CLASSES_ROOT, TEXT("CLSID\\{777A968C-D8A7-48A9-B027-BEC429268F4F}"));

	return(S_OK);
*/
}