class IMyMath : public IDispatch
{
	public:
		virtual HRESULT __stdcall SumOfTwoIntegers(int,  int, int *) = 0;

		virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *) = 0;		
};

// {404503E9-1CB0-4581-8009-E24C1BDBA2AF}
const CLSID CLSID_MyMath = { 0x404503e9, 0x1cb0, 0x4581, { 0x80, 0x9, 0xe2, 0x4c, 0x1b, 0xdb, 0xa2, 0xaf } };

// {022D8E51-683E-4B79-B9A2-1492DCAAB93E}
const IID IID_IMyMath = { 0x22d8e51, 0x683e, 0x4b79, { 0xb9, 0xa2, 0x14, 0x92, 0xdc, 0xaa, 0xb9, 0x3e } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : AutomationServer";
const char g_szVerIndProgId[] = "COMProject.Automation";

const char g_szProgId[]       = "COMProject.Automation.1";
