#include <windows.h>
#include <stdio.h>
#include "AutomationServer.h"
#include "registry.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "advapi32.lib")

HMODULE ghModule = NULL;

long glNumberOfActiveComponents = 0;
long glNumberOfServerLocks 		= 0;

// {ECEDFC08-B038-46E1-AE71-B2C935EB7376}
const GUID LIBID_AutomationServer = { 0xecedfc08, 0xb038, 0x46e1, { 0xae, 0x71, 0xb2, 0xc9, 0x35, 0xeb, 0x73, 0x76 } };

class CMyMath : public IMyMath
{
	private:
		long m_cRef;
		ITypeInfo *m_pITypeInfo = NULL;
	public:
		CMyMath(void);
	   ~CMyMath(void);
	   
	   HRESULT 	__stdcall QueryInterface(REFIID, void **);
	   ULONG 	__stdcall AddRef(void);
	   ULONG 	__stdcall Release(void);

	   HRESULT 	__stdcall GetTypeInfoCount(UINT *);
	   HRESULT 	__stdcall GetTypeInfo(UINT, LCID, ITypeInfo**);
	   HRESULT 	__stdcall GetIDsOfNames(REFIID, LPOLESTR *, UINT, LCID, DISPID *);			
	   HRESULT	__stdcall Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, 
	   							 EXCEPINFO *, UINT *);
	   
	   HRESULT 	__stdcall SumOfTwoIntegers(int, int, int *);
	   HRESULT 	__stdcall SubtractionOfTwoIntegers(int, int, int *);

	   HRESULT InitInstance(void);
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

	return (TRUE);
}

CMyMath::CMyMath (void)
{
	m_cRef = 1;
	InterlockedIncrement(&glNumberOfActiveComponents);
}

CMyMath::~CMyMath(void)
{
	InterlockedDecrement(&glNumberOfActiveComponents);
}

HRESULT CMyMath::QueryInterface(REFIID riid, void **ppv)
{
	if (riid == IID_IUnknown)
		*ppv = static_cast<IMyMath *>(this);
	else if (riid == IID_IDispatch)
		*ppv = static_cast<IMyMath *>(this);
	else if (riid == IID_IMyMath)
		*ppv = static_cast<IMyMath *>(this);
	else 
	{
		*ppv = NULL;
		return(E_NOINTERFACE);
	}
	reinterpret_cast<IUnknown *>(*ppv) -> AddRef();
	return(S_OK);
}

ULONG CMyMath::AddRef(void)
{
	InterlockedIncrement(&m_cRef);
	return(m_cRef);
}

ULONG CMyMath::Release(void)
{
	InterlockedDecrement(&m_cRef);
	if (m_cRef == 0)
	{
		m_pITypeInfo -> Release();
		m_pITypeInfo = NULL;
		delete(this);
		return(0);
	}
	return(m_cRef);
}

HRESULT CMyMath::SumOfTwoIntegers(int iNum1, int iNum2, int *pSum)
{
	*pSum = iNum1 + iNum2;
	return(S_OK);
}

HRESULT CMyMath::SubtractionOfTwoIntegers(int iNum1, int iNum2, int *pSub)
{
	*pSub = iNum1 - iNum2;
	return(S_OK);
}

HRESULT CMyMath::InitInstance(void)
{
	void ComErrorDescriptionString(HWND, HRESULT);

	HRESULT hr;
	ITypeLib *pITypeLib = NULL;

	if (m_pITypeInfo == NULL)
	{
		hr = LoadRegTypeLib(LIBID_AutomationServer, 
							1, 0, 0x00, 
							&pITypeLib);
		if (FAILED(hr))
		{
			MessageBox(NULL, TEXT("AutomationServer : InitInstance : LoadRegTypeLib : Failed"), TEXT("AutomationServer : Log"), NULL);
			ComErrorDescriptionString(NULL, hr);
			pITypeLib -> Release();
			return(hr);
		}

		hr = pITypeLib -> GetTypeInfoOfGuid(IID_IMyMath, &m_pITypeInfo);
		if (FAILED(hr))
		{
			MessageBox(NULL, TEXT("AutomationServer : InitInstance : LoadRegTypeLib : Failed"), TEXT("AutomationServer : Log"), NULL);
			ComErrorDescriptionString(NULL, hr);
			pITypeLib -> Release();
			return(hr);
		}

		pITypeLib -> Release();
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
	// MessageBox(NULL, TEXT("AutomationServer : ClassFactory : CreateInstance : Entered"), TEXT("AutomationServer : Log"), NULL);

	CMyMath *pCMyMath = NULL;
	HRESULT hr;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);

	pCMyMath = new CMyMath();
	if (pCMyMath == NULL)
		return(E_OUTOFMEMORY);

	// MessageBox(NULL, TEXT("AutomationServer : ClassFactory : CreateInstance : Calling InitInstance"), TEXT("AutomationServer : Log"), NULL);

	pCMyMath -> InitInstance();

	hr = pCMyMath -> QueryInterface(riid, ppv);
	pCMyMath -> Release();

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

HRESULT CMyMath::GetTypeInfoCount(UINT *pCountTypeInfo)
{
	*pCountTypeInfo = 1;

	return(S_OK);
}

HRESULT CMyMath::GetTypeInfo(UINT iTypeInfo, LCID lcid, ITypeInfo **ppITypeInfo)
{
	*ppITypeInfo = NULL;

	if (iTypeInfo != 0)
		return (DISP_E_BADINDEX);

	m_pITypeInfo -> AddRef();

	*ppITypeInfo = m_pITypeInfo;

	return(S_OK);
}

HRESULT CMyMath::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	// MessageBox(NULL, TEXT("GetIDs Of Names : Entered"), TEXT("KJBJB"), NULL);
	return(DispGetIDsOfNames(m_pITypeInfo, rgszNames, cNames, rgDispId));
}

HRESULT CMyMath::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, 
						VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT hr;

	hr = DispInvoke(this, 
					m_pITypeInfo, 
					dispIdMember, 
					wFlags, 
					pDispParams, 
					pVarResult, 
					pExcepInfo, 
					puArgErr);
	return(hr);
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	ClassFactory *pClassFactory = NULL;
	HRESULT hr;

	if (rclsid != CLSID_MyMath)
		return(CLASS_E_CLASSNOTAVAILABLE);

	pClassFactory = new ClassFactory();
	if (pClassFactory == NULL)
		return(E_OUTOFMEMORY);

	hr = pClassFactory -> QueryInterface(riid, ppv);

	pClassFactory -> Release();

	return(hr);
}

extern "C" HRESULT __stdcall DllCanUnloadNow(void)
{
	if ((glNumberOfActiveComponents == 0) && (glNumberOfServerLocks == 0))
		return(S_OK);
	else 
		return(S_FALSE);
}

void ComErrorDescriptionString(HWND hwnd, HRESULT hr)
{
	TCHAR  *szErrorMessage = NULL;
	TCHAR   str[255];

	if (FACILITY_WINDOWS == HRESULT_FACILITY(hr))
		hr =  HRESULT_CODE(hr);

	if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, 
					  hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&szErrorMessage, 0, NULL ) != 0)
	{
		sprintf_s(str, TEXT("%#x : %s"), hr, szErrorMessage);
	}	
	else 
		sprintf_s(str, TEXT("Could not find a description for error # %#x.\n"), hr);

	MessageBox(hwnd, str, TEXT("COM Error"), MB_OK);	
}


STDAPI DllRegisterServer()
{
	return RegisterServer(ghModule,
						  CLSID_MyMath,
						  g_szFriendlyName,
						  g_szVerIndProgId,
						  g_szProgId);
}

STDAPI DllUnRegisterServer()
{

	return UnregisterServer(CLSID_MyMath,
							g_szVerIndProgId,
							g_szProgId);
}

