#include <windows.h>
#include "AggregationInnerComponent.h"
#include "registry.h"

long glNumberOfActiveComponents = 0;
long glNumberOfServerLocks 		= 0;

HINSTANCE ghModule;

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

interface INoAggregationIUnknown
{
	virtual HRESULT __stdcall QueryInterface_NoAggregation(REFIID, void **)   = 0;
	virtual ULONG 	__stdcall AddRef_NoAggregation(void) 					  = 0;
	virtual ULONG 	__stdcall Release_NoAggregation(void)					  = 0;
};

class CMultiplicationDivision:public INoAggregationIUnknown, IMultiplication, IDivision
{
	private:
		long m_cRef;
		IUnknown *m_pIUnknownOuter;

	public:
		 CMultiplicationDivision(IUnknown *);	
		~CMultiplicationDivision();
		
		HRESULT __stdcall QueryInterface_NoAggregation(REFIID, void **);
		ULONG 	__stdcall AddRef_NoAggregation(void);
		ULONG 	__stdcall Release_NoAggregation(void);

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int *);
		HRESULT __stdcall DivisionOfTwoIntegers(int, int, int *);
};

class ClassFactory:public IClassFactory
{
	private:
		long m_cRef;

	public:
		 ClassFactory();
		~ClassFactory();
		
		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void); 	
		ULONG 	__stdcall Release(void);

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

CMultiplicationDivision::CMultiplicationDivision(IUnknown *pIUnknownOuter)
{
	m_cRef = 1;
	InterlockedIncrement(&glNumberOfActiveComponents);

	if (pIUnknownOuter != NULL)
		m_pIUnknownOuter = pIUnknownOuter;
	else 
		m_pIUnknownOuter = reinterpret_cast<IUnknown *>(static_cast<INoAggregationIUnknown *>(this));
}

CMultiplicationDivision::~CMultiplicationDivision(void)
{
	InterlockedDecrement(&glNumberOfActiveComponents);
	if (m_pIUnknownOuter)
	{
		m_pIUnknownOuter -> Release();
		m_pIUnknownOuter = NULL;
	}
}

HRESULT CMultiplicationDivision::QueryInterface(REFIID riid, void ** ppv)
{
	return(m_pIUnknownOuter -> QueryInterface(riid, ppv));
}

ULONG CMultiplicationDivision::AddRef(void)
{
	return(m_pIUnknownOuter -> AddRef());
}

ULONG CMultiplicationDivision::Release(void)
{
	return(m_pIUnknownOuter -> Release());
}

HRESULT CMultiplicationDivision::QueryInterface_NoAggregation(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<INoAggregationIUnknown *>(this);
	else if (riid == IID_IMultiplication)
		*ppv = static_cast<IMultiplication *>(this);
	else if (riid == IID_IDivision)
		*ppv = static_cast<IDivision *>(this);
	else
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv) -> AddRef();

	return(S_OK);
}

ULONG CMultiplicationDivision::AddRef_NoAggregation(void)
{
	InterlockedIncrement(&m_cRef);

	return(m_cRef);
}

ULONG CMultiplicationDivision::Release_NoAggregation(void)
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

	return(S_OK);
}

ClassFactory::ClassFactory(void)
{
	m_cRef = 1;
}

ClassFactory::~ClassFactory(void){}

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
	return m_cRef;
}

HRESULT ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
	CMultiplicationDivision *pCMultiplicationDivision = NULL;
	HRESULT hr;

	if ((pUnkOuter != NULL) && (riid != IID_IUnknown))
		return(CLASS_E_NOAGGREGATION);

	pCMultiplicationDivision = new CMultiplicationDivision(pUnkOuter);
	if (pCMultiplicationDivision == NULL)
		return(E_OUTOFMEMORY);

	hr = pCMultiplicationDivision -> QueryInterface_NoAggregation(riid, ppv);
	pCMultiplicationDivision -> Release_NoAggregation();

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
}


STDAPI DllUnRegisterServer()
{
	return UnregisterServer(CLSID_MultiplicationDivision,
							g_szVerIndProgId,
							g_szProgId);
}