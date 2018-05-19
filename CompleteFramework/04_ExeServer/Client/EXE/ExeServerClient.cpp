#include <windows.h>
#include "ExeServer.h"

LRESULT CALLBACK Wndproc(HWND, UINT, WPARAM, LPARAM);

int WINAPI WinMain(HINSTANCE hInstanse, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
	WNDCLASSEX 	wndclass;
	HWND 		hwnd;
	MSG 		msg;
	TCHAR 		AppName[] = TEXT("Exe Client");

	wndclass.cbSize = sizeof(wndclass);
	wndclass.style 	= CS_HREDRAW | CS_VREDRAW;
	wndclass.cbClsExtra = 0;
	wndclass.cbWndExtra = 0;
	wndclass.lpfnWndProc	= WndProc;
	wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hCursor  = LoadCursor(NULL, IDC_ARROW);
	wndclass.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hbrBackground = (HBRUSH)GetSTockObject(WHITE_BRUSH);
	wndclass.hInstance = hInstance;
	wndclass.lpszClassName = AppName;
	wndclass.lpszMenuName  = NULL;

	RegisterClassEx(&wndclass);

	hwnd = CreateWindow(AppName, 
						TEXT("Client of Exe Server"), 
						WS_OVERLAPPEDWINDOW, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						NULL, NULL, 
						hInstance, 
						NULL);

	ShowWindow(hwnd);
	UpdateWindow(hwnd);

	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return (msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	STDAPI RegisterProxyStubDLL(void);
	STDAPI UnRegisterProxyStubDLL(void);

	ISum 		*pISum = NULL;
	ISUbtract	*pISubtract = NULL; 
	HRESULT 	 hr;
	int 		 error, n1, n2, n3;
	TCHAR 		 str[255];

	switch(iMsg)
	{
		case WM_CREATE:
				RegisterProxyStubDLL();
				
				hr = CoInitialize();
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("COM library cannot be initialized"), 
									 TEXT("Error"), MB_OK);
					DestroyWindow(hwnd);
					exit(0);
				}

				hr = CoCreateInstance(CLSID_SumSUbtract, NULL, CLSCTX_LOCAL_SERVER, 
										IID_ISum, (void **)&pISum);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("Exe Server : Component cannot be obtained "), TEXT("Error"), NULL);	
					DestroyWindow(hwnd);
					exit(0);	
				}	

				n1 = 25;
				n2 = 5;
				pISum -> SumOfTwoIntegers(n1, n2, &n3);
				wsprintf(str,TEXT("Sum of %d an %d is %d"), n1, n2, n3);

				MessageBox(hwnd, str, TEXT("Result"), MB_OK);

				hr = pISum -> QueryInterface(IID_ISubtract, (void **)&pISubtract);

				pISubtract -> SubtractionOfTwoIntegers(n1, n2, &n3);
				wsprintf(str, TEXT("Sum of %d and %d is %d"), n1, n2, n3);

				MessageBox(hwnd, str, TEXT("Result"), MB_OK):

				DestroyWindow();

			break;
		case WM_DETROY:
				CoUninitialize();

				UnRegisterProxyStubDLL();
			
				PostQuitMessage(0);
			break;	
	}
}

void RegisterProxyStubDLL()
{
	HINSTANCE hLib;
	
	typedef STDAPI (*FuncPtr)(void);
	FuncPtr RegisterServer;

	hLib = LoadLibrary(TEXT("..\\04_ExeServer\\Component\\ProxyStub\\ProxyStub.dll"));
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("Exe Server Client : RegisterProxyStubDLL : Cannot register ProxyStubDLL"), 
					TEXT("Error"), NULL);
		DestroyWindow(ghwnd);
		exit(0);
	}

	RegisterServer = (FuncPtr)GetProcAddress(hLib, "DllRegisterServer");

	RegisterServer();

	FreeLibrary(hLib);
}

void UnRegisterProxyStubDLL()
{
	HINSTANCE hLib;
	
	typedef STDAPI (*FuncPtr)(void);
	FuncPtr UnRegisterServer;


	hLib = LoadLibrary(TEXT("ProxyStub.dll"));
	if (hLib == NULL)
	{
		MessageBox(NULL, TEXT("Exe Server Client : UnRegisterProxyStubDLL : Cannot register ProxyStubDLL"), 
					TEXT("Error"), NULL);
		DestroyWindow(ghwnd);
		exit(0);
	}

	UnRegisterServer = (FuncPtr)GetProcAddress(hLib, "DllUnRegisterServer");

	UnRegisterServer();

	FreeLib(UnRegisterServer);
}