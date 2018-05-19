#include <windows.h>
#include <iostream>
#include "InterfaceHeader.h"
#include "registry.h"

using namespace std;

long glComponents  	= 0;
long glServerLocks 	= 0;

HMODULE ghModule 	= NULL;

#pragma comment(lib, "user32.lib"  )
#pragma comment(lib, "gdi32.lib"   )
#pragma comment(lib, "ole32.lib"   )
#pragma comment(lib, "advapi32.lib")

void trace(const char *pMsg){ cout << "ClassFactory : Server : " << pMsg << endl;}

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

class CMath : public ISum, ISub
{
	public:

		 CMath();
		~CMath();

		HRESULT __stdcall QueryInterface(REFIID refIID, void **ppv);	
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall SumOfTwoIntegers(int, int, int *);
		HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *);

	private:
		long m_cRef;	
};

class ClassFactory : public IClassFactory
{
	public:
		 ClassFactory();
		~ClassFactory();
		
		HRESULT __stdcall QueryInterface(REFIID refIID, void **ppv); 	
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall CreateInstance(IUnknown *, REFIID, void **);
		HRESULT __stdcall LockServer(BOOL);

	private:
		long m_cRef;	
};

CMath::CMath()
{
	m_cRef = 1;
	InterlockedIncrement(&glComponents);
	trace("CMath::CMath()");
}

CMath::~CMath()
{
	InterlockedDecrement(&glComponents);
	trace("CMath::~CMath()");
}

HRESULT CMath::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISum)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISub)
		*ppv = static_cast<ISub *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv)->AddRef();
	return(S_OK);
}

ULONG CMath::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMath::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return 0;
	}
	return(m_cRef);
}

HRESULT CMath::SumOfTwoIntegers(int num1, int num2, int *pSum)
{
	*pSum = num1 + num2;
	return(S_OK);	
}

HRESULT CMath::SubtractionOfTwoIntegers(int num1, int num2, int *pSub)
{
	*pSub = num1 - num2;
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
	return(m_cRef);
}

ULONG ClassFactory::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

HRESULT ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
	CMath *pCMath = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);
	pCMath = new CMath;
	if (pCMath == NULL)
		return(E_OUTOFMEMORY);

	hr = pCMath -> QueryInterface(riid, ppv);
	pCMath -> Release();

	return hr;
}

HRESULT ClassFactory::LockServer(BOOL fLock)
{
	if (fLock)
		InterlockedIncrement(&glServerLocks);
	else 
		InterlockedDecrement(&glServerLocks);
	
	return(S_OK);
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	ClassFactory *pClassFactory = NULL;
	HRESULT hr;

	if (rclsid != CLSID_Math)
		return(CLASS_E_CLASSNOTAVAILABLE);

	pClassFactory = new ClassFactory;
	if (pClassFactory == NULL)
		return (E_OUTOFMEMORY);
	hr = pClassFactory -> QueryInterface(riid, ppv);
	pClassFactory -> Release();

	return(hr);
}

HRESULT __stdcall DllCanUnloadNow(void)
{
	if ((glComponents == 0) && (glServerLocks == 0))
		return (S_OK);
	else 
		return(S_FALSE);
}
 
STDAPI DllRegisterServer()
{
//	MessageBox(NULL, TEXT("ClassFactoryComponent : DllRegisterServer : Enetered"), TEXT("ClassFactory"), NULL);

	return RegisterServer(ghModule,
						  CLSID_Math,
						  g_szFriendlyName,
						  g_szVerIndProgId,
						  g_szProgId);
}

STDAPI DllUnregisterServer()
{
	return UnregisterServer(CLSID_Math,
							g_szVerIndProgId,
							g_szProgId);
}
