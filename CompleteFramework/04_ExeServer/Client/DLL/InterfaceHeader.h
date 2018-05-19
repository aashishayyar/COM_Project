class ISum : public IUnknown 
{
	public :
		virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int *) = 0;
};

class ISubtract : public IUnknown 
{
	public :
		virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *) = 0;
};

// {B9F9596C-5135-4C6E-9BA8-21F224CCAFD1}
const CLSID CLSID_SumSubtract = { 0xb9f9596c, 0x5135, 0x4c6e, { 0x9b, 0xa8, 0x21, 0xf2, 0x24, 0xcc, 0xaf, 0xd1 } };

// {A3A50D32-58B3-4897-A447-DC8B935EFB94}
const IID IID_ISum = { 0xa3a50d32, 0x58b3, 0x4897, { 0xa4, 0x47, 0xdc, 0x8b, 0x93, 0x5e, 0xfb, 0x94 } };

// {4D44D171-6898-4C54-BDEF-8302EA46ACBA}
const IID IID_ISubtract = { 0x4d44d171, 0x6898, 0x4c54, { 0xbd, 0xef, 0x83, 0x2, 0xea, 0x46, 0xac, 0xba } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : ExeServer";
const char g_szVerIndProgId[] = "COMProject.ExeServer";

const char g_szProgId[]       = "COMProject.ExeServer.1";

