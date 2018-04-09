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
	void InitializeServers(void);
	void FreeServers(void);
	void SafeInterfaceRelease(void);

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
				InitializeServers();
//				MessageBox(NULL, TEXT("Initialization done ..."), TEXT("Error"), NULL);
			break;
		
		case WM_KEYDOWN:
				switch(wParam)
				{
					case 0x31: // 1 ClassFactory
							Sum = (FuncPtr)GetProcAddress(hClassFactoryLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hClassFactoryLib, "SubtractionOfTwoIntegers");
							Sum(160, 240);
							Sub(430, 270);
						break;
					case 0x32: // 2 Containment
							Sum = (FuncPtr)GetProcAddress(hContainmentLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hContainmentLib, "SubtractionOfTwoIntegers");
							Mul = (FuncPtr)GetProcAddress(hContainmentLib, "MultiplicationOfTwoIntegers");
							Div = (FuncPtr)GetProcAddress(hContainmentLib, "DivisionOfTwoIntegers");
							Sum(110, 200);
							Sub(470, 230);
							Mul(234, 987);
							Div(987, 543);
						break;	
					case 0x33: // 3 Aggregation
							Sum = (FuncPtr)GetProcAddress(hAggregationLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hAggregationLib, "SubtractionOfTwoIntegers");
							Mul = (FuncPtr)GetProcAddress(hAggregationLib, "MultiplicationOfTwoIntegers");
							Div = (FuncPtr)GetProcAddress(hAggregationLib, "DivisionOfTwoIntegers");
							Sum(10, 06);
							Sub(40, 20);
							Mul(34, 98);
							Div(97, 43);
						break;	
					case 0x34:
							Sum = (FuncPtr)GetProcAddress(hExeServerLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hExeServerLib, "SubtractionOfTwoIntegers");
							Sum(98, 25);
							Sub(71, 94);
						break;
					case 0x35:
							Sum = (FuncPtr)GetProcAddress(hAutomationLib, "SumOfTwoIntegers");
							Sub = (FuncPtr)GetProcAddress(hAutomationLib, "SubtractionOfTwoIntegers");
							Sum(58, 32);
							Sub(91, 56);
						break;
					case 0x36:
							GetOddNum = (FuncPtr2)GetProcAddress(hMonikerLib, "GetNextOddNumber");
							GetOddNum(43);
						break;		
					case 0x1B:
							DestroyWindow(hwnd);
						break;	
				}
			break;

		case WM_DESTROY:
				FreeServers();
				PostQuitMessage(0);
			break;	
	}
	return(DefWindowProc(hwnd, iMsg, wParam, lParam));
}

void InitializeServers(void)
{
	typedef void(*fPtr)(void);

//	ClassFactory : Initialize
	hClassFactoryLib = LoadLibrary(TEXT("..\\01_ClassFactory\\Client\\DLL\\ClassFactoryClient.dll"));	
	if (hClassFactoryLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : ClassFactory : Error"), TEXT("Error"), NULL);

	fPtr InitClassFactory;
	InitClassFactory = (fPtr)GetProcAddress(hClassFactoryLib, "Initialize");
	InitClassFactory();

// 	Containment : Initialize
	hContainmentLib = LoadLibrary(TEXT("..\\02_Containment\\Client\\DLL\\ContainmentClient.dll"));	
	if (hContainmentLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : Containment : Error"), TEXT("Error"), NULL);

	fPtr InitContainment;
	InitContainment = (fPtr)GetProcAddress(hContainmentLib, "Initialize");
	InitContainment();

//  Aggregation : Initialize
	hAggregationLib = LoadLibrary(TEXT("..\\03_Aggregation\\Client\\DLL\\AggregationClient.dll"));	
	if (hAggregationLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : Aggregation : Error"), TEXT("Error"), NULL);

	fPtr InitAggregation;
	InitAggregation = (fPtr)GetProcAddress(hAggregationLib, "Initialize");
	InitAggregation();

//  Exe Server : Initialize
	hExeServerLib = LoadLibrary(TEXT("..\\04_ExeServer\\Client\\DLL\\ExeClient.dll"));	
	if (hExeServerLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : ExeServer : Error"), TEXT("Error"), NULL);

	fPtr InitExeServer;
	InitExeServer = (fPtr)GetProcAddress(hExeServerLib, "Initialize");
	InitExeServer();

// 	Automation : Initialize
	hAutomationLib = LoadLibrary(TEXT("..\\05_Automation\\Client\\DLL\\AutomationClient.dll"));	
	if (hAutomationLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : Automation : Error"), TEXT("Error"), NULL);

	fPtr InitAutomationServer;
	InitAutomationServer = (fPtr)GetProcAddress(hAutomationLib, "Initialize");
	InitAutomationServer();

// 	Moniker : Initialize
	hMonikerLib = LoadLibrary(TEXT("..\\06_Moniker\\Client\\DLL\\MonikerClient.dll"));	
	if (hMonikerLib == NULL)
		MessageBox(NULL, TEXT("Bare Client : LoadServers : Moniker : Error"), TEXT("Error"), NULL);

	fPtr InitMonikerServer;
	InitMonikerServer = (fPtr)GetProcAddress(hMonikerLib, "Initialize");
	InitMonikerServer();

}

void FreeServers(void)
{
	typedef void(*fPtr)(void);

//  ClassFactory
	fPtr UnInitClassFactory;
	UnInitClassFactory = (fPtr)GetProcAddress(hClassFactoryLib, "UnInitialize");
	UnInitClassFactory();

	FreeLibrary(hClassFactoryLib);

//  Containment
	fPtr UnInitContainment;
	UnInitContainment = (fPtr)GetProcAddress(hContainmentLib, "UnInitialize");
	UnInitContainment();

	FreeLibrary(hContainmentLib);

//  Aggregation
	fPtr UnInitAggregation;
	UnInitAggregation = (fPtr)GetProcAddress(hAggregationLib, "UnInitialize");
	UnInitAggregation();

	FreeLibrary(hAggregationLib);

// 	ExeServer
	fPtr UnInitExeServer;
	UnInitExeServer = (fPtr)GetProcAddress(hExeServerLib, "UnInitialize");
	UnInitExeServer();

	FreeLibrary(hExeServerLib);	

//  Automation 
	fPtr UnInitAutomation;
	UnInitAutomation = (fPtr)GetProcAddress(hAutomationLib, "UnInitialize");
	UnInitAutomation();

	FreeLibrary(hAutomationLib);	

// 	Moniker
	fPtr UnInitMoniker;
	UnInitMoniker = (fPtr)GetProcAddress(hMonikerLib, "UnInitialize");
	UnInitMoniker();

	FreeLibrary(hMonikerLib);	
}