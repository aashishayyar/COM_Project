#include <windows.h>
#include "MonikerServer.h"
#include "registry.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

long glNumberOfActiveComponents = 0;
long glNumberOfServerLocks 		= 0;
HMODULE ghModule  				= NULL;		 

class COddNumber : public IOddNumber
{
	private:
		long m_cRef;
		int  m_iFirstOddNumber;
	public:
		 COddNumber(int);
		~COddNumber(void); 	

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		BOOL 	__stdcall IsOdd(int);

		HRESULT __stdcall GetNextOddNumber(int *);
};	

class COddNumberFactory : public IOddNumberFactory
{
	private:
		long m_cRef;

	public:	
		 COddNumberFactory(void);
		~COddNumberFactory(void);

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall SetFirstOddNumber(int, IOddNumber **);
};

BOOL WINAPI DllMain(HINSTANCE hDll, DWORD dwReason, LPVOID Reserved)
{
	switch(dwReason)
	{
		case DLL_PROCESS_ATTACH:
					// MessageBox(NULL, TEXT("Moniker Server DllMain : Entered"), TEXT("HERE"), NULL);
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

COddNumber::COddNumber(int iFirstOddNumber)
{
	// MessageBox(NULL, TEXT("COddNumber Consructor"), TEXT("Log"), NULL);
	m_cRef = 1;
	m_iFirstOddNumber = iFirstOddNumber;

	InterlockedIncrement(&glNumberOfActiveComponents); 
}

COddNumber::~COddNumber(void)
{
	InterlockedDecrement(&glNumberOfActiveComponents);
}

HRESULT COddNumber::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IOddNumber *>(this);
	else if (riid == IID_IOddNumber)
		*ppv = static_cast<IOddNumber *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv)->AddRef();
	return(S_OK);
}

ULONG COddNumber::AddRef(void)
{
	InterlockedIncrement(&m_cRef);

	return(m_cRef);
}

ULONG COddNumber::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}
	return(m_cRef);
}

BOOL COddNumber::IsOdd(int iFirstOddNumber)
{
	if ( (iFirstOddNumber != 0) && ( (iFirstOddNumber %2) != 0) )
		return(TRUE);
	else 
		return(FALSE);
}

HRESULT COddNumber::GetNextOddNumber(int *pNextOddNumber)
{	
	BOOL bResult;
	bResult = IsOdd(m_iFirstOddNumber);

	if (bResult == TRUE)
		*pNextOddNumber = m_iFirstOddNumber + 2;
	else 
		return(S_FALSE);

	return(S_OK);
}

COddNumberFactory::COddNumberFactory(void)
{
	// MessageBox(NULL, TEXT("Inside COddNumberfactory"), TEXT("HERE"), NULL);
	m_cRef = 1;
}

COddNumberFactory::~COddNumberFactory(void)
{

}

HRESULT COddNumberFactory::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IOddNumberFactory *>(this);
	else if (riid == IID_IOddNumberFactory)
		*ppv = static_cast<IOddNumberFactory *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv) -> AddRef();

	return(S_OK);
}

ULONG COddNumberFactory::AddRef(void)
{
	InterlockedIncrement(&m_cRef);

	return(m_cRef);
}

ULONG COddNumberFactory::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		delete(this);
		return(0);
	}

	return(m_cRef);
}

HRESULT COddNumberFactory::SetFirstOddNumber(int iFirstOddNumber, IOddNumber **ppIOddNumber)
{
	HRESULT hr;
	COddNumber *pCOddNumber = new COddNumber(iFirstOddNumber);

	if (pCOddNumber == NULL)
		return(E_OUTOFMEMORY);

	hr = pCOddNumber -> QueryInterface(IID_IOddNumber, (void **)ppIOddNumber);
	pCOddNumber -> Release();

	return(hr);
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	// MessageBox(NULL, TEXT("DllGetClassObject : Entered"), TEXT("Moniker : Log"), NULL);
	COddNumberFactory *pCOddNumberFactory = NULL;
	HRESULT hr;

	if (rclsid != CLSID_OddNumber)
		return(CLASS_E_CLASSNOTAVAILABLE);

	pCOddNumberFactory = new COddNumberFactory();
	if (pCOddNumberFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pCOddNumberFactory -> QueryInterface(riid, ppv);
	pCOddNumberFactory -> Release();

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
						  CLSID_OddNumber,
						  g_szFriendlyName,
						  g_szVerIndProgId,
						  g_szProgId);
}

STDAPI DllUnRegisterServer()
{
	return UnregisterServer(CLSID_OddNumber,
							g_szVerIndProgId,
							g_szProgId);
}
