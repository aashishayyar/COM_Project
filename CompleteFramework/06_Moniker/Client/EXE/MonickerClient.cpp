#define UNICODE
#include <windows.h>
#include "InterfaceHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

IOddNumber 			*pIOddNumber 		= NULL;
IOddNumberFactory 	*pIOddNumberFactory = NULL;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
	WNDCLASSEX wndclass;
	HWND hwnd;
	MSG msg;
	TCHAR AppName[] = TEXT("ComClient");
	HRESULT hr;

	hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		MessageBox(NULL, TEXT("COM Library cannot be initialized.\nProgram will now exit."), 
					TEXT("Error"), MB_OK);
		exit(0);
	}

	wndclass.cbSize  = sizeof(wndclass);
	wndclass.style   = CS_HREDRAW | CS_VREDRAW;
	wndclass.cbClsExtra = 0;
	wndclass.cbWndExtra = 0;
	wndclass.lpfnWndProc = WndProc;
	wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
	wndclass.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wndclass.hInstance = hInstance;
	wndclass.lpszClassName = AppName;
	wndclass.lpszMenuName = NULL;

	RegisterClassEx(&wndclass);

	hwnd = CreateWindow(AppName, 
						TEXT("Moniker Client"), 
						WS_OVERLAPPEDWINDOW, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						CW_USEDEFAULT, 
						NULL, 
						NULL, 
						hInstance, 
						NULL);

	ShowWindow(hwnd, nCmdShow);
	UpdateWindow(hwnd);

	while(GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	CoUninitialize();

	return ((int)msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	void SafeInterfaceRelease(void);

	IBindCtx *pIBindCtx = NULL;
	IMoniker *pIMoniker = NULL;
	ULONG 	  uEaten;
	HRESULT   hr;
	LPOLESTR  szCLSID = NULL;
	wchar_t   wszCLSID[255];
	wchar_t   wszTemp[255];
	wchar_t  *ptr;
	int       iFirstOddNumber;
	int 	  iNextOddNumber;
	TCHAR     str[255];

	switch(iMsg)
	{
		case WM_CREATE:
				if (hr = CreateBindCtx(0, &pIBindCtx) != S_OK)	
				{
					MessageBox(hwnd, TEXT("Failed to get IBindCtx Interface pointer"), TEXT("Error"), NULL);
					DestroyWindow(hwnd);	
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
					MessageBox(hwnd, TEXT("Failed to get IMoniker interface"), TEXT("Error"), NULL);
					pIBindCtx -> Release();
					pIBindCtx = NULL;

					DestroyWindow(hwnd);
				}	

				hr = pIMoniker -> BindToObject(pIBindCtx, NULL, IID_IOddNumberFactory, (void **)pIOddNumberFactory);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("Failed to get custom activation : IOddNumberFactory"), TEXT("Error"), NULL);
				
					pIMoniker -> Release();
					pIMoniker = NULL;

					pIBindCtx -> Release();
					pIBindCtx = NULL;

					DestroyWindow(hwnd);	
				}

				pIMoniker -> Release();
				pIMoniker = NULL;
				pIBindCtx -> Release();
				pIBindCtx = NULL;

				iFirstOddNumber = 57;

				hr = pIOddNumberFactory -> SetFirstOddNumber(iFirstOddNumber, &pIOddNumber);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("Cannot obtaing IOddNumber"), TEXT("Error"), NULL);

					pIOddNumberFactory -> Release();
					pIOddNumberFactory = NULL;
					
					DestroyWindow(hwnd);
				}

				pIOddNumberFactory -> Release();
				pIOddNumberFactory = NULL;

				pIOddNumber -> GetNextOddNumber(&iNextOddNumber);

				pIOddNumber -> Release();
				pIOddNumber = NULL;

				wsprintf(str, TEXT("The next odd number from %2d is %2d"), iFirstOddNumber, iNextOddNumber);
				MessageBox(hwnd, str, TEXT("Result"), NULL);

				DestroyWindow(hwnd);
			break;

		case WM_DESTROY:
				PostQuitMessage(0);
			break;				
	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));	
}

void SafeInterfaceRelease(void)
{
	if (pIOddNumber)
	{
		pIOddNumber -> Release();
		pIOddNumber = NULL;
	}

	if (pIOddNumberFactory)
	{
		pIOddNumberFactory -> Release();
		pIOddNumberFactory = NULL;
	}
}
