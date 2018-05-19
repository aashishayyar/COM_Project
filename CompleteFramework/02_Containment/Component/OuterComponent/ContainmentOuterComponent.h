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

// {777A968C-D8A7-48A9-B027-BEC429268F4F}
const CLSID CLSID_SumSubtract =  { 0x777a968c, 0xd8a7, 0x48a9, { 0xb0, 0x27, 0xbe, 0xc4, 0x29, 0x26, 0x8f, 0x4f } };

// {1A070DC9-F5FA-45C9-A221-C57F2783F5BD}
const IID 	IID_ISum = { 0x1a070dc9, 0xf5fa, 0x45c9, { 0xa2, 0x21, 0xc5, 0x7f, 0x27, 0x83, 0xf5, 0xbd } };

// {D1B1481A-7552-40E9-A5A9-777DB3F60584}
const IID 	IID_ISubtract = { 0xd1b1481a, 0x7552, 0x40e9, { 0xa5, 0xa9, 0x77, 0x7d, 0xb3, 0xf6, 0x5, 0x84 } };

