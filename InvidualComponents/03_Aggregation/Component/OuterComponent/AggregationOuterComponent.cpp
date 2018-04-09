#include <windows.h>
#include "AggregationOuterComponent.h"
#include "registry.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

long glNumberOfActiveComponents;
long glNumberOfServerLocks;

HINSTANCE ghModule;

class CSumSubtract : public ISum, ISubtract
{
	private:
		long 			 m_cRef;
		IUnknown 		*m_pIUnknownInner;
		IMultiplication *m_pIMultiplication;
		IDivision 		*m_pIDivision;

	public:
		 CSumSubtract(void);
		~CSumSubtract(void); 

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG	__stdcall Release(void);

		HRESULT __stdcall SumOfTwoIntegers		  (int, int, int *);
		HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *);

		HRESULT __stdcall InitializeInnerComponent(void);
};

class ClassFactory : public IClassFactory
{
	private:
		long m_cRef;

	public:
		 ClassFactory(void);
		~ClassFactory(void);
		
		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall CreateInstance(IUnknown *, REFIID, void **);
		HRESULT __stdcall LockServer(BOOL);
};

BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID lpszReserved)
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
	m_pIUnknownInner 	= NULL;
	m_pIMultiplication 	= NULL;
	m_pIDivision		= NULL;
	m_cRef				= 1;

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
	if (m_pIUnknownInner)
	{
		m_pIUnknownInner -> Release();
		m_pIUnknownInner = NULL;
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
		return(m_pIUnknownInner -> QueryInterface(riid, ppv));
	else if (riid == IID_IDivision)
		return(m_pIUnknownInner -> QueryInterface(riid, ppv));
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);	
	}	 

	reinterpret_cast<IUnknown *>(*ppv) -> AddRef();
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
		delete(this);
		return 0;
	}
	return m_cRef;
}

HRESULT CSumSubtract::SumOfTwoIntegers(int iNum1, int iNum2, int *pSum)
{
	*pSum = iNum1 + iNum2;
	return S_OK;
}

HRESULT CSumSubtract::SubtractionOfTwoIntegers(int iNum1, int iNum2, int *pSub)
{
	*pSub = iNum1 - iNum2;
	return S_OK;
}

HRESULT CSumSubtract::InitializeInnerComponent(void)
{
	HRESULT hr;

	hr = CoCreateInstance(CLSID_MultiplicationDivision, 
						  reinterpret_cast<IUnknown *>(this), 
						  CLSCTX_INPROC_SERVER, 
						  IID_IUnknown, 
						  (void **)&m_pIUnknownInner);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IUnknown Interface cannot be obtained from Inner component"), 
				   TEXT("Error"), NULL);
		return(E_FAIL);
	}				

	hr = m_pIUnknownInner -> QueryInterface(IID_IMultiplication, (void **)&m_pIMultiplication);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IMultiplication Interface cannot be obtained from Inner component"), 
					TEXT("Error"), NULL);
		
		m_pIUnknownInner -> Release();
		m_pIUnknownInner = NULL;

		return (E_FAIL);
	}		  		

	hr = m_pIUnknownInner -> QueryInterface(IID_IDivision, (void **)&m_pIDivision);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("IDivision Interface cannot be obtained from Inner component"), 
					TEXT("Error"), NULL);
		
		m_pIUnknownInner -> Release();
		m_pIUnknownInner = NULL;

		return (E_FAIL);
	}		  		

	return (S_OK);
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
	reinterpret_cast<IUnknown *>(*ppv) -> AddRef();

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
	CSumSubtract *pCSumSubtract = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return CLASS_E_NOAGGREGATION;

	pCSumSubtract = new CSumSubtract;
	if (pCSumSubtract == NULL)
		return(E_OUTOFMEMORY);
	
	hr = pCSumSubtract -> InitializeInnerComponent();
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("Failed to Initialize Inner Component"), 
					TEXT("Error"), NULL);
		pCSumSubtract -> Release();
		return(hr);
	}

	hr = pCSumSubtract -> QueryInterface(riid, ppv);
	pCSumSubtract -> Release();

	return hr;
}

HRESULT ClassFactory::LockServer(BOOL fLock)
{
	if (fLock)
		InterlockedIncrement(&glNumberOfServerLocks);
	else 
		InterlockedIncrement(&glNumberOfServerLocks);

	return S_OK;
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
}

STDAPI DllUnRegisterServer()
{
	return UnregisterServer(CLSID_SumSubtract,
							g_szVerIndProgId,
							g_szProgId);
}