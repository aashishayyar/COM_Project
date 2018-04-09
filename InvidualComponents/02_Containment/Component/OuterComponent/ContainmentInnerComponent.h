class IMultiplication:public IUnknown
{
	public:
		virtual HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int *) = 0;
};

class IDivision:public IUnknown
{
	public:
		virtual HRESULT __stdcall DivisionOfTwoIntegers(int, int, int *) = 0;	
};

// {A556F1B8-B1BA-4BB4-9AE5-5902EEBF9A64}
const CLSID CLSID_MultiplicationDivision = { 0xa556f1b8, 0xb1ba, 0x4bb4, { 0x9a, 0xe5, 0x59, 0x2, 0xee, 0xbf, 0x9a, 0x64 } };

// {1ECCFD3E-41F1-4E99-A6C1-BE1983E9F297}
const IID IID_IMultiplication = { 0x1eccfd3e, 0x41f1, 0x4e99, { 0xa6, 0xc1, 0xbe, 0x19, 0x83, 0xe9, 0xf2, 0x97 } };

// {557F03EB-CE0D-4B71-8077-7B016E97059B}
const IID IID_IDivision = { 0x557f03eb, 0xce0d, 0x4b71, { 0x80, 0x77, 0x7b, 0x1, 0x6e, 0x97, 0x5, 0x9b } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : ContainmentServer";
const char g_szVerIndProgId[] = "COMProject.Containment";

const char g_szProgId[]       = "COMProject.Containment.1";

