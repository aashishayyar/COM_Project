#include <windows.h>
#include "AggregationHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

ISum 	  		*pISum				= NULL;
ISubtract 		*pISubtract 		= NULL;
IMultiplication *pIMultiplication 	= NULL;
IDivision 		*pIDivision 		= NULL;

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
						TEXT("Aggregation Client"), 
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

	HRESULT hr;
	int iNum1, iNum2, iSum, iSubtraction, iMultiplication, iDivision;
	TCHAR str[255];

	switch(iMsg)
	{
		case WM_CREATE:
				
				hr = CoCreateInstance(CLSID_SumSubtract, NULL, CLSCTX_INPROC_SERVER, 
									  IID_ISum, (void **)&pISum);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("ISum interface cannot be obtained."), TEXT("Error "), MB_OK);
					DestroyWindow(hwnd);
				}

				iNum1 = 65;
				iNum2 = 45;
				pISum -> SumOfTwoIntegers(iNum1, iNum2, &iSum);

				wsprintf(str, TEXT("Sum of %d and %d is %d.\n"), iNum1, iNum2, iSum);
				MessageBox(hwnd, str, TEXT("Result"), MB_OK);

				hr = pISum -> QueryInterface(IID_ISubtract, (void **)&pISubtract);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("ISubtract interface cannot be obtained."), TEXT("Error"), MB_OK);
					DestroyWindow(hwnd);
				}

				pISum -> Release();
				pISum = NULL;

				pISubtract -> SubtractionOfTwoIntegers(iNum1, iNum2, &iSubtraction);
				
				wsprintf(str, TEXT("Subtraction of %d and %d is %d.\n"), iNum1, iNum2, iSubtraction);
				MessageBox(hwnd, str, TEXT("Result"), MB_OK);
				 
				hr = pISubtract -> QueryInterface(IID_ISubtract, (void **)&pIMultiplication);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("ISubtract interface cannot be obtained."), TEXT("Error"), MB_OK);
					DestroyWindow(hwnd);
				}
				 
				pISubtract -> Release();
				pISubtract = NULL;

				pIMultiplication -> MultiplicationOfTwoIntegers(iNum1, iNum2, &iMultiplication);
				
				wsprintf(str, TEXT("Multiplication of %d and %d is %d.\n"), iNum1, iNum2, iMultiplication);
				MessageBox(hwnd, str, TEXT("Result"), MB_OK);

				hr = pIMultiplication -> QueryInterface(IID_ISubtract, (void **)&pIDivision);
				if (FAILED(hr))
				{
					MessageBox(hwnd, TEXT("ISubtract interface cannot be obtained."), TEXT("Error"), MB_OK);
					DestroyWindow(hwnd);
				}
				 
				pIMultiplication -> Release();
				pIMultiplication = NULL;

				pIDivision -> DivisionOfTwoIntegers(iNum1, iNum2, &iDivision);
				
				wsprintf(str, TEXT("Division of %d and %d is %d.\n"), iNum1, iNum2, iDivision);
				MessageBox(hwnd, str, TEXT("Result"), MB_OK);

				pIDivision -> Release();
				pIDivision = NULL;

			break;
		case WM_DESTROY:
				PostQuitMessage(0);
			break;				
	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));	
}

void SafeInterfaceRelease(void)
{
	if (pISum)
	{
		pISum -> Release();
		pISum = NULL;
	}
	if (pISubtract)
	{
		pISubtract -> Release();
		pISubtract = NULL;
	}
	if (pIMultiplication)
	{
		pIMultiplication -> Release();
		pIMultiplication = NULL;
	}
	if (pIDivision)
	{
		pIDivision -> Release();
		pIDivision = NULL;
	}
}
