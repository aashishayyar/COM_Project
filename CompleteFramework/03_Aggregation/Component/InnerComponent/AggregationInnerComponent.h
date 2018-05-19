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

// {E962E75B-6A3F-44E1-9B32-8912B76599C8}
const CLSID CLSID_MultiplicationDivision = { 0xe962e75b, 0x6a3f, 0x44e1, { 0x9b, 0x32, 0x89, 0x12, 0xb7, 0x65, 0x99, 0xc8 } };

// {94827F67-7575-425B-AC3B-3195470BB82C}
const IID IID_IMultiplication = { 0x94827f67, 0x7575, 0x425b, { 0xac, 0x3b, 0x31, 0x95, 0x47, 0xb, 0xb8, 0x2c } };

// {23E89007-0C30-4474-847A-13EF68855846}
const IID IID_IDivision = { 0x23e89007, 0xc30, 0x4474, { 0x84, 0x7a, 0x13, 0xef, 0x68, 0x85, 0x58, 0x46 } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : AggregationServer";
const char g_szVerIndProgId[] = "COMProject.AggregationServer";

const char g_szProgId[]       = "COMProject.AggregationServer.1";

