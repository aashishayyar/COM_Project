class ISum : public IUnknown 
{
	public:
		virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int *) = 0;
};

class ISubtract : public IUnknown 
{
	public:
		virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *) = 0;
};

class IMultiplication : public IUnknown 
{
	public:
		virtual HRESULT __stdcall MultiplicationOfTwoIntegers(int, int, int *) = 0;
};

class IDivision : public IUnknown 
{
	public:
		virtual HRESULT __stdcall DivisionOfTwoIntegers(int, int, int *) = 0;
};

// {C79235EF-C0CA-49B8-991C-BC7436D198A1}
const CLSID CLSID_SumSubtract = { 0xc79235ef, 0xc0ca, 0x49b8, { 0x99, 0x1c, 0xbc, 0x74, 0x36, 0xd1, 0x98, 0xa1 } };

// {E74D8F97-D2B7-4533-9646-482B148658AC}
const IID IID_ISum = { 0xe74d8f97, 0xd2b7, 0x4533, { 0x96, 0x46, 0x48, 0x2b, 0x14, 0x86, 0x58, 0xac } };

// {A2CD60D3-019A-4FC0-9567-28DCDD1CA7CF}
const IID IID_ISubtract = { 0xa2cd60d3, 0x19a, 0x4fc0, { 0x95, 0x67, 0x28, 0xdc, 0xdd, 0x1c, 0xa7, 0xcf } };

// {E962E75B-6A3F-44E1-9B32-8912B76599C8}
const CLSID CLSID_MultiplicationDivision = { 0xe962e75b, 0x6a3f, 0x44e1, { 0x9b, 0x32, 0x89, 0x12, 0xb7, 0x65, 0x99, 0xc8 } };

// {94827F67-7575-425B-AC3B-3195470BB82C}
const IID IID_IMultiplication = { 0x94827f67, 0x7575, 0x425b, { 0xac, 0x3b, 0x31, 0x95, 0x47, 0xb, 0xb8, 0x2c } };

// {23E89007-0C30-4474-847A-13EF68855846}
const IID IID_IDivision = { 0x23e89007, 0xc30, 0x4474, { 0x84, 0x7a, 0x13, 0xef, 0x68, 0x85, 0x58, 0x46 } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : AggregationServer";
const char g_szVerIndProgId[] = "COMProject.AggregationServer";

const char g_szProgId[]       = "COMProject.AggregationServer.1";

