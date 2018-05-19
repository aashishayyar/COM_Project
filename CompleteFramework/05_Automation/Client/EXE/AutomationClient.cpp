#include <windows.h>
#include "AutomationHeader.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

ISum 	  		*pISum				= NULL;
ISubtract 		*pISubtract 		= NULL;


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
	void ComErrorDescriptionString(HWND, HRESULT);

	IDispatch  *pIDispatch = NULL;
	HRESULT    	hr;
	DISPID 	   	dispid;
	OLECHAR    *SumOfTwoIntegers = TEXT("SumOfTwoIntegers");
	OLECHAR    *SubtractionOfTwoIntegers = TEXT("SubtractionofTwoIntegers");
	VARIANT		vArg[2], vRet;	
	DISPPARAMS	param = { varg, 0, 2, NULL};
	int 		iNum1, iNum2; 
	TCHAR 		str[255];

	switch(iMsg)
	{
		case WM_CREATE:
				
				hr = CoInitialize();
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(hwnd, TEXT("COM library cannot be initialized"), TEXT("COM Error"), MB_OK);

					DestroyWindow(hwnd);
					exit(0);
				}				

				hr = CoCreateInstance(CLSID_MyMath, NULL, 
									  CLSCTX_INPROC_SERVER, 
									  IID_IDispatch, 
									  (void **)&pIDispatch);
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(hwnd, TEXT("Component cannot be created"), TEXT("COM Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
					DestroyWindow(hwnd);
					exit(0);
				}					  	

				iNum1 = 75;
				iNum2 = 25;

				variantInit(vArg);
				vArg[0].vt 		= VT_INT;
				vArg[0].intval 	= iNum2;
				vArg[1].vt 		= VT_INT;
				vArg[1].intval	= iNum1;

				param.cArgs 	 = 2;
				param.cNamedArgs = 0;
				param.rgdispidNamedArgs = NULL;
				param.rgarg      = vArg;

				variantInit(&vRet);
				hr = pIDispatch -> GetIDsOfNames(IID_NULL, 
												 &SumOfTwoIntegers, 
												 1, GetUserDefaultLCID(), 
												 &dispid);
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(NULL, TEXT("Cannot get ID for SumOfTwoIntegers()"), 
							   TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
					pIDispatch -> Release();
					DestroyWindow(hwnd);
				}

				hr = pIDispatch -> Invoke(dispid, 
											IID_NULL, 
											GetUserDefaultLCID(), 
											DISPATCH_METHOD, 
											&param, 
											&vRet, 
											NULL, 
											NULL);
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(NULL, TEXT("Cannot invoke function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
					pIDispatch -> Release();

					DestroyWindow(hwnd);
				}								
				else 
				{
					wsprintf(str, TEXT("Sum of %d and %d is %d"), iNum1, iNum2, vRet.lVal);
					MessageBox(hwnd, str, TEXT("SumOfTwoIntegers"), MB_OK);
				}

				hr = pIDispatch -> GetIDsOfNames(IID_NULL, 
												 &SubtractionOfTwoIntegers, 
												 1, GetUserDefaultLCID(), 
												 &dispid);
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(NULL, TEXT("Cannot get ID for SumOfTwoIntegers()"), 
							   TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
					pIDispatch -> Release();
					DestroyWindow(hwnd);
				}

				hr = pIDispatch -> Invoke(dispid, 
											IID_NULL, 
											GetUserDefaultLCID(), 
											DISPATCH_METHOD, 
											&param, 
											&vRet, 
											NULL, 
											NULL);
				if (FAILED(hr))
				{
					ComErrorDescriptionString(hwnd, hr);
					MessageBox(NULL, TEXT("Cannot invoke function"), TEXT("Error"), MB_OK | MB_ICONERROR | MB_TOPMOST);
					pIDispatch -> Release();

					DestroyWindow(hwnd);
				}								
				else 
				{
					wsprintf(str, TEXT("Subtraction of %d and %d is %d"), iNum1, iNum2, vRet.lVal);
					MessageBox(hwnd, str, TEXT("SubtractionOfTwoIntegers"), MB_OK);
				}

				VariantClear(vArg);
				VariantClear(&vRet);
				pIDispatch -> Release();
				pIDispatch = NULL;
				DestroyWindow(hwnd);

			break;
		case WM_DESTROY:
				PostQuitMessage(0);
			break;				
	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));	
}

void ComErrorDescriptionString(HWND hwnd, HRESULT hr)
{
	TCHAR *szErrorMessage = NULL;
	TCHAR  str[255];

	if (FACILITY_WINDOWS == HRESULT_FACILITY(hr))
		hr = HRESULT_CODE(hr);

	if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, 
					  NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), 
					  (LPTSTR)&szErrorMessage, 0, NULL) != 0)
	{
		swprintf_s(str, TEXT("%s"), szErrorMessage);
		LocalFree(szErrorMessage);
	}
	else 
		swprintf_s(str, TEXT("Could not find a description for error # %#x.\n"), hr);

	MessageBox(hwnd, str, TEXT("COM Error"), MB_OK);
}

void SafeInterfaceRelease(void)
{

}
