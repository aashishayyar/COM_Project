#include <windows.h>
#include <tlhelp32.h>
#include "Header.h"
#include "registry.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

long glNumberOfActiveComponents = 0;
long glNumberOfServerLocks 		= 0;

IClassFactory *gpIClassFactory = NULL;

HWND  ghwnd = NULL;
DWORD dwRegister;
HRESULT hr;

HANDLE hMutex = NULL;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

class CSumSubtract : public ISum, ISubtract
{
	private :
		long m_cRef;

	public :
		 CSumSubtract();	
		~CSumSubtract();

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall SumOfTwoIntegers(int, int, int *);
		HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *);
};

class ClassFactory : public IClassFactory
{
	private :
		long m_cRef;

	public :
		 ClassFactory();
		~ClassFactory();

		HRESULT __stdcall QueryInterface(REFIID, void **);
		ULONG 	__stdcall AddRef(void);
		ULONG 	__stdcall Release(void);

		HRESULT __stdcall CreateInstance(IUnknown *, REFIID, void **);
		HRESULT __stdcall LockServer(BOOL);		
};

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	HRESULT StartMyClassFactory  (void);
	void 	StopMyClassFactory	 (void);
	DWORD 	GetParentProcessID	 (void);

	if (hMutex != NULL)
	{
		return 0;
	}
	else
	{
		hMutex = CreateMutex(NULL, FALSE, "ExeServerComProjectSingleInstance");
		if (hMutex == NULL)
		{	
			MessageBox(NULL, TEXT("Failed to create mutex for Exe Server"), TEXT("Exe Server : Error"), NULL);
			
		}
	}

	WNDCLASSEX 	wndclass;
	MSG 		msg;
	HWND 		hwnd;
	HRESULT 	hr;
	int 		iDontShowWindow = 0;
	TCHAR 		AppName[]    = TEXT("Exe Server");
	const wchar_t szTokens[] = L"-/";
	wchar_t 	*lpszTokens;
	wchar_t 	 lpszCmdLine[255];
	wchar_t 	*next_token = NULL;

	// MessageBox(NULL, TEXT("ExeServer : WinMain : Entered"), TEXT("ExeServer : Log"), NULL);

	MessageBox(NULL, lpCmdLine, TEXT("ExeServer : Log : lpCmdLine"), NULL);

	MultiByteToWideChar(CP_ACP, 0, lpCmdLine, 255, lpszCmdLine, 255);

	lpszTokens = wcstok_s(lpszCmdLine, szTokens, &next_token);

	while(lpszTokens != NULL)
	{	
		if (_wcsicmp(lpszTokens, L"register") == 0)
		{

		//	MessageBox(NULL, TEXT("ExeServer : WinMain : Registering"), TEXT("ExeServer : Log"), NULL);

			return RegisterServer(hInstance,
								  CLSID_SumSubtract,
								  g_szFriendlyName,
								  g_szVerIndProgId,
								  g_szProgId);
		}
		else if(_wcsicmp(lpszTokens, L"unregister") == 0)
		{
			return UnregisterServer(CLSID_SumSubtract,
									g_szVerIndProgId,
									g_szProgId);
		}
		else if (_wcsicmp(lpszTokens, L"Embedding") == 0)
		{
			iDontShowWindow = 1;
			break;
		}
		else 
			return 0;

		lpszTokens = wcstok_s(NULL, szTokens, &next_token);
	}
	
	// MessageBox(NULL, TEXT("Out of while"), TEXT("ExeServer"), NULL);	
	// MessageBox(NULL, TEXT("Going in GetParentProcessID"), TEXT("ExeServer"), NULL); 
	GetParentProcessID();
	// MessageBox(NULL, TEXT("Out of GetPartentProcessID"), TEXT("ExeServer"), NULL);

	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		return(0);
	} 

	wndclass.cbSize 		= sizeof(wndclass);
	wndclass.style  		= CS_HREDRAW | CS_VREDRAW;
	wndclass.cbWndExtra		= 0;
	wndclass.cbClsExtra 	= 0;
	wndclass.lpfnWndProc 	= WndProc;
	wndclass.hIcon			= LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hCursor 		= LoadCursor(NULL, IDC_ARROW);
	wndclass.hIconSm 		= LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hInstance 		= hInstance;
	wndclass.lpszClassName 	= AppName;
	wndclass.lpszMenuName 	= NULL;
	wndclass.hbrBackground  = (HBRUSH)GetStockObject(BLACK_BRUSH);

	RegisterClassEx(&wndclass);

	hwnd = CreateWindow(AppName, 
						TEXT("Exe Server"), 
						WS_OVERLAPPEDWINDOW, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						NULL, NULL, 
						hInstance, 
						NULL	);

	ghwnd = hwnd;

	if (iDontShowWindow != 1)
	{
	//	MessageBox(NULL, TEXT("Showing Window"), TEXT("ExeServer"), NULL);
		ShowWindow(hwnd, nCmdShow);
		UpdateWindow(hwnd);

		++glNumberOfServerLocks;	
	}							

	if (iDontShowWindow == 1)
	{
	//	MessageBox(NULL, TEXT("Starting Classfactory"), TEXT("ExeServer : Log"), NULL);
		hr = StartMyClassFactory();
		if (FAILED(hr))
		{
			MessageBox(NULL, TEXT("Failed : Starting Classfactory"), TEXT("ExeServer : Log"), NULL);

			DestroyWindow(hwnd);
			exit(0);
		}
	} 

	// MessageBox(NULL, TEXT("Entering Message Loop : Class Factory started succesfully"), TEXT("ExeServer : Log"), NULL);

	while(GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	if (iDontShowWindow == 1)
	{
		StopMyClassFactory();
	}

	CoUninitialize();

	return((int)msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	HDC  		hdc;
	RECT 		rc;
	PAINTSTRUCT ps;

	switch(iMsg)
	{
		case WM_PAINT:
			GetClientRect(hwnd, &rc);
			hdc = BeginPaint(hwnd, &ps);
			SetBkColor(hdc, RGB(0, 0, 0));
			SetTextColor(hdc, RGB(0, 255, 0));
			DrawText(hdc, TEXT("This is a COM Server."), -1, &rc, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
			EndPaint(hwnd, &ps);
			break;
		
		case WM_CLOSE :
			--glNumberOfServerLocks;
			ShowWindow(hwnd, SW_HIDE);
			break;

		case WM_DESTROY:
			if (glNumberOfActiveComponents == 0 && glNumberOfServerLocks == 0)
				PostQuitMessage(0);
			break;	

		default:	
			break;

	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));
}

CSumSubtract::CSumSubtract()
{
	// MessageBox(NULL, TEXT("ExeServer : CSumSubtract : Constructor : Entered"), TEXT("ExeServer :Log"), NULL);

	m_cRef = 1;
	InterlockedIncrement(&glNumberOfActiveComponents);
}

CSumSubtract::~CSumSubtract()
{
	InterlockedDecrement(&glNumberOfActiveComponents);
}

HRESULT CSumSubtract::QueryInterface(REFIID riid, void **ppv)
{
	// MessageBox(NULL, TEXT("ExeServer : CSumSubtract : QueryInterface : Entered"), TEXT("ExeServer :Log"), NULL);

	if (riid == IID_IUnknown)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISum)
		*ppv = static_cast<ISum *>(this);
	else if (riid == IID_ISubtract)
		*ppv = static_cast<ISubtract *>(this);
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
		delete(this);

		if (glNumberOfActiveComponents == 0 && glNumberOfServerLocks == 0)
			PostMessage(ghwnd, WM_QUIT, (WPARAM)0, (LPARAM)0L);

		return 0;
	}
	return m_cRef;
}

HRESULT CSumSubtract::SumOfTwoIntegers(int iNum1, int iNum2, int *pSum)
{
	*pSum = iNum1 + iNum2;
	return(S_OK);
}

HRESULT CSumSubtract::SubtractionOfTwoIntegers(int iNum1, int iNum2, int *pSub)
{
	*pSub = iNum1 - iNum2;
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
	reinterpret_cast<IUnknown *>(*ppv)-> AddRef();

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
		return 0;
	}
	return m_cRef;
}

HRESULT ClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
	// MessageBox(NULL, TEXT("ExeServer : Classfactory : CreateInstance : Entered"), TEXT("ExeServer : Log"), NULL);

	CSumSubtract *pCSumSubtract = NULL;

	if (pUnkOuter != NULL)
		return(CLASS_E_NOAGGREGATION);

	// MessageBox(NULL, TEXT("ExeServer : Classfactory : CreateInstance : CSumSubtract instance creating ..."), TEXT("ExeServer : Log"), NULL);


	pCSumSubtract = new CSumSubtract;
	if (pCSumSubtract == NULL)
		return(E_OUTOFMEMORY);
	
	// MessageBox(NULL, TEXT("ExeServer : Classfactory : CreateInstance : Calling QueryInterface"), TEXT("ExeServer : Log"), NULL);

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

	if (glNumberOfActiveComponents == 0 && glNumberOfServerLocks == 0)
		PostMessage(ghwnd, WM_QUIT, (WPARAM)0, (LPARAM)0L);

	return(S_OK);
}

HRESULT StartMyClassFactory(void)
{
	HRESULT hr;
	
	// MessageBox(NULL, TEXT("ExeServer : StartMyClassfactory : Entered"), TEXT("ExeServer : Log"), NULL);

	gpIClassFactory = new ClassFactory();
	if (gpIClassFactory == NULL)
		return(E_NOINTERFACE);

	// MessageBox(NULL, TEXT("ExeServer : StartMyClassfactory : New Classfactory instance created"), TEXT("ExeServer : Log"), NULL);

	gpIClassFactory-> AddRef();

	// MessageBox(NULL, TEXT("ExeServer : StartMyClassfactory : Calling CoRegisterClassObject"), TEXT("ExeServer : Log"), NULL);

	hr = CoRegisterClassObject(CLSID_SumSubtract, 
								static_cast<IUnknown *>(gpIClassFactory), 
								CLSCTX_LOCAL_SERVER, 
								REGCLS_MULTIPLEUSE, 
								&dwRegister);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("ExeServer : StartMyClassfactory : CoregisterClassObject : Failed"), TEXT("ExeServer : Log"), NULL);
		gpIClassFactory -> Release();
		return(E_FAIL);
	}

	// MessageBox(NULL, TEXT("ExeServer : StartMyClassfactory : CoregisterClassObject : Success"), TEXT("ExeServer : Log"), NULL);

	return(S_OK);
}

void StopMyClassFactory(void)
{
	if (dwRegister != 0)
		CoRevokeClassObject(dwRegister);

	if (gpIClassFactory != NULL)
		gpIClassFactory -> Release();
}

DWORD GetParentProcessID(void)
{
	HANDLE 	hProcessSnapshot = NULL;
	BOOL 	bRetCode 		 = FALSE;
	PROCESSENTRY32 ProcessEntry = {0};
	DWORD 	dwPPID;
	wchar_t	szNameOfThisProcess[_MAX_PATH], szNameOfParentProcess[_MAX_PATH], *ptr = NULL, wszTemp1[_MAX_PATH], wszTemp2[_MAX_PATH];
	TCHAR	szTemp[_MAX_PATH], str[_MAX_PATH] ;

	

	// MessageBox(NULL, TEXT("HERE1"), TEXT("HERE1"), NULL);

	hProcessSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	if (hProcessSnapshot == INVALID_HANDLE_VALUE)
		return(-1);

	// MessageBox(NULL, TEXT("HERE2"), TEXT("HERE2"), NULL);

	ProcessEntry.dwSize = sizeof(PROCESSENTRY32);

	// MessageBox(NULL, TEXT("HERE3"), TEXT("HERE3"), NULL);

	if (Process32First(hProcessSnapshot, &ProcessEntry))
	{
		// MessageBox(NULL, TEXT("HERE4"), TEXT("HERE4"), NULL);

		GetModuleFileName(NULL, szTemp, sizeof(szTemp) / sizeof(TCHAR));


		// MessageBox(NULL, szTemp, TEXT("HERE5"), NULL);

		MultiByteToWideChar(CP_ACP, 0, szTemp, 255, wszTemp1, 255);
		
		ptr = wcsrchr(wszTemp1, '\\');
		// if (ptr == NULL)
		// 	 MessageBox(NULL, TEXT("HERE6"), TEXT("HERE6"), NULL);

		wcscpy_s(szNameOfThisProcess, sizeof(wszTemp1) / sizeof(wchar_t), ptr + 1);

		// MessageBox(NULL, TEXT("HERE7"), TEXT("HERE7"), NULL);

		do 
		{
			errno_t err;
			err = _wcslwr_s(szNameOfThisProcess, wcslen(szNameOfThisProcess) + 1);
			MultiByteToWideChar(CP_ACP, 0, ProcessEntry.szExeFile, 255, wszTemp1, 255);		
			err = _wcsupr_s(wszTemp1, wcslen(wszTemp1) + 1);
			if (wcsstr(szNameOfThisProcess, wszTemp1) != NULL)
			{
				wsprintf(str, TEXT("Current Process Name : %s\nCurrentProcess ID : %ld\nParentProcess Name : %s\
									Parent Process ID : %s\n"), szNameOfThisProcess, ProcessEntry.th32ProcessID, 
									ProcessEntry.th32ParentProcessID, ProcessEntry.szExeFile);
				MessageBox(NULL, str, TEXT("Parent Info"), MB_OK | MB_TOPMOST);

				dwPPID = ProcessEntry.th32ParentProcessID;	
			}
		}
		while(Process32Next(hProcessSnapshot, &ProcessEntry));

		// MessageBox(NULL, TEXT("HERE6"), TEXT("HERE6"), NULL);

	}
	CloseHandle(hProcessSnapshot);
	return(dwPPID);
}
