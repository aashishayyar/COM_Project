class ISum: public IUnknown
{
	public:
		virtual HRESULT __stdcall SumOfTwoIntegers(int, int, int *) = 0;
};

class ISubtract:public IUnknown
{
	public:
		virtual HRESULT __stdcall SubtractionOfTwoIntegers(int, int, int *) = 0;
};

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

// {777A968C-D8A7-48A9-B027-BEC429268F4F}
const CLSID CLSID_SumSubtract =  { 0x777a968c, 0xd8a7, 0x48a9, { 0xb0, 0x27, 0xbe, 0xc4, 0x29, 0x26, 0x8f, 0x4f } };

// {1A070DC9-F5FA-45C9-A221-C57F2783F5BD}
const IID 	IID_ISum = { 0x1a070dc9, 0xf5fa, 0x45c9, { 0xa2, 0x21, 0xc5, 0x7f, 0x27, 0x83, 0xf5, 0xbd } };

// {D1B1481A-7552-40E9-A5A9-777DB3F60584}
const IID 	IID_ISubtract = { 0xd1b1481a, 0x7552, 0x40e9, { 0xa5, 0xa9, 0x77, 0x7d, 0xb3, 0xf6, 0x5, 0x84 } };

// {1ECCFD3E-41F1-4E99-A6C1-BE1983E9F297}
const IID IID_IMultiplication = { 0x1eccfd3e, 0x41f1, 0x4e99, { 0xa6, 0xc1, 0xbe, 0x19, 0x83, 0xe9, 0xf2, 0x97 } };

// {557F03EB-CE0D-4B71-8077-7B016E97059B}
const IID IID_IDivision = { 0x557f03eb, 0xce0d, 0x4b71, { 0x80, 0x77, 0x7b, 0x1, 0x6e, 0x97, 0x5, 0x9b } };


