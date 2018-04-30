#include <windows.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

HINSTANCE hClassFactoryLib;
HINSTANCE hContainmentLib;
HINSTANCE hAggregationLib;
HINSTANCE hExeServerLib;
HINSTANCE hAutomationLib;
HINSTANCE hMonikerLib;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstace, LPSTR lpCmdLine, int nCmdShow)
{
	WNDCLASSEX 	wndclass;
	HWND 		hwnd;
	MSG 		msg;
	TCHAR 		AppName[] = TEXT("Bare Client");
	HRESULT 	hr;

	wndclass.cbSize 		= sizeof(wndclass);
	wndclass.style 			= CS_HREDRAW | CS_VREDRAW;
	wndclass.hInstance 		= hInstance;
	wndclass.lpfnWndProc 	= WndProc;
	wndclass.lpszClassName 	= AppName;
	wndclass.lpszMenuName 	= NULL;
	wndclass.cbClsExtra 	= 0;
	wndclass.cbWndExtra 	= 0;
	wndclass.hIcon 			= LoadIcon  (NULL, IDI_APPLICATION);
	wndclass.hIconSm		= LoadIcon  (NULL, IDI_APPLICATION);
	wndclass.hCursor 		= LoadCursor(NULL, IDC_ARROW);
	wndclass.hbrBackground 	= (HBRUSH)GetStockObject(WHITE_BRUSH);

	RegisterClassEx(&wndclass);

	hwnd = CreateWindow(AppName,
						TEXT("BARE CLIENT"),
						WS_OVERLAPPEDWINDOW,
						CW_USEDEFAULT,
						CW_USEDEFAULT,
						CW_USEDEFAULT,
						CW_USEDEFAULT, 
						NULL, NULL, 
						hInstance, NULL );

	ShowWindow(hwnd, nCmdShow);
	UpdateWindow(hwnd);

	while(GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return((int)msg.wParam);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	HRESULT hr;
	int iNum1, iNum2, iSum;
	TCHAR str[255];

	typedef int(*FuncPtr)(int, int);
	typedef int(*FuncPtr2)(int);

	FuncPtr Sum, Sub, Mul, Div;
	FuncPtr2 GetOddNum;

	switch(iMsg)
	{
		case WM_CREATE:

//				MessageBox(NULL, TEXT("Going to initialize ... "), TEXT("Error"), NULL);
//				InitializeServers();
//				MessageBox(NULL, TEXT("Initialization done ..."), TEXT("Error"), NULL);
			break;
		
		case WM_KEYDOWN:
				switch(wParam)
				{
					case 0x31: // 1 ClassFactory
							hClassFactoryLib = LoadLibrary(TEXT("..\\01_ClassFactory\\Client\\DLL\\ClassFactoryClient.dll"));	
							if (hClassFactoryLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : ClassFactory : Error"), TEXT("Error"), NULL);

							Sum = (FuncPtr)GetProcAddress(hClassFactoryLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hClassFactoryLib, "SubtractionOfTwoIntegers");
							Sum(160, 240);
							Sub(430, 270);

							FreeLibrary(hClassFactoryLib);
						break;
					case 0x32: // 2 Containment
							hContainmentLib = LoadLibrary(TEXT("..\\02_Containment\\Client\\DLL\\ContainmentClient.dll"));	
							if (hContainmentLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : Containment : Error"), TEXT("Error"), NULL);

							Sum = (FuncPtr)GetProcAddress(hContainmentLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hContainmentLib, "SubtractionOfTwoIntegers");
							Mul = (FuncPtr)GetProcAddress(hContainmentLib, "MultiplicationOfTwoIntegers");
							Div = (FuncPtr)GetProcAddress(hContainmentLib, "DivisionOfTwoIntegers");
							Sum(110, 200);
							Sub(470, 230);
							Mul(234, 987);
							Div(987, 543);

							FreeLibrary(hContainmentLib);
						break;	
					case 0x33: // 3 Aggregation
							hAggregationLib = LoadLibrary(TEXT("..\\03_Aggregation\\Client\\DLL\\AggregationClient.dll"));	
							if (hAggregationLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : Aggregation : Error"), TEXT("Error"), NULL);
							
							Sum = (FuncPtr)GetProcAddress(hAggregationLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hAggregationLib, "SubtractionOfTwoIntegers");
							Mul = (FuncPtr)GetProcAddress(hAggregationLib, "MultiplicationOfTwoIntegers");
							Div = (FuncPtr)GetProcAddress(hAggregationLib, "DivisionOfTwoIntegers");
							Sum(10, 06);
							Sub(40, 20);
							Mul(34, 98);
							Div(97, 43);

							FreeLibrary(hAggregationLib);
						break;	
					case 0x34:
							hExeServerLib = LoadLibrary(TEXT("..\\04_ExeServer\\Client\\DLL\\ExeClient.dll"));	
							if (hExeServerLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : ExeServer : Error"), TEXT("Error"), NULL);

							Sum = (FuncPtr)GetProcAddress(hExeServerLib, "SumOfTwoIntegers");
							//Sub = (FuncPtr)GetProcAddress(hExeServerLib, "SubtractionOfTwoIntegers");
							Sum(98, 25);
							//Sub(71, 94);
						
							FreeLibrary(hExeServerLib);	
						break;
					case 0x35:
							hAutomationLib = LoadLibrary(TEXT("..\\05_Automation\\Client\\DLL\\AutomationClient.dll"));	
							if (hAutomationLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : Automation : Error"), TEXT("Error"), NULL);

							Sum = (FuncPtr)GetProcAddress(hAutomationLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hAutomationLib, "SubtractionOfTwoIntegers");
							Sum(58, 32);
							Sub(91, 56);
						
							FreeLibrary(hAutomationLib);	
						break;
					case 0x36:
							hMonikerLib = LoadLibrary(TEXT("..\\06_Moniker\\Client\\DLL\\MonikerClient.dll"));	
							if (hMonikerLib == NULL)
								MessageBox(NULL, TEXT("Bare Client : LoadServers : Moniker : Error"), TEXT("Error"), NULL);

							GetOddNum = (FuncPtr2)GetProcAddress(hMonikerLib, "GetNextOddNumber");
							GetOddNum(43);

							FreeLibrary(hMonikerLib);	
						break;		
					case 0x1B:
							DestroyWindow(hwnd);
						break;	
				}
			break;

		case WM_DESTROY:
				PostQuitMessage(0);
			break;	
	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));
}

