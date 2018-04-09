class ISum : public IUnknown
{
	public : 
			virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int *) = 0;
};

class ISub : public IUnknown
{
	public :
			virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *) = 0;
};

// {88EC8F11-18C7-45D3-9746-0FC3CFB1A727}
const CLSID CLSID_Math = { 0x88ec8f11, 0x18c7, 0x45d3, { 0x97, 0x46, 0xf, 0xc3, 0xcf, 0xb1, 0xa7, 0x27 } };

// {29E5837B-01C0-4191-B5EC-0EB1106AC0E5}
const IID IID_ISum = { 0x29e5837b, 0x1c0, 0x4191, { 0xb5, 0xec, 0xe, 0xb1, 0x10, 0x6a, 0xc0, 0xe5 } };

// {2ADA9375-7335-4FA4-AF09-175DA0CF6AAC}
const IID IID_ISub = { 0x2ada9375, 0x7335, 0x4fa4, { 0xaf, 0x9, 0x17, 0x5d, 0xa0, 0xcf, 0x6a, 0xac } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : ClassFactoryServer";
const char g_szVerIndProgId[] = "COMProject.ClassFactory";

const char g_szProgId[]       = "COMProject.ClassFactory.1";
