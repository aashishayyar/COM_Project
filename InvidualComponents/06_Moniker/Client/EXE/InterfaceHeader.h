class IOddNumber : public IUnknown 
{
	public:
		virtual HRESULT __stdcall GetNextOddNumber(int *) = 0;
};

class IOddNumberFactory : public IUnknown 
{
	public:
		virtual HRESULT __stdcall SetFirstOddNumber(int , IOddNumber **) = 0;
};

// {366E5A64-8175-479B-BC58-771C4C6F2FC1}
const CLSID CLSID_OddNumber = { 0x366e5a64, 0x8175, 0x479b, { 0xbc, 0x58, 0x77, 0x1c, 0x4c, 0x6f, 0x2f, 0xc1 } };

// {E5970DF2-589F-46B2-BDCF-6DBD9B44965B}
const IID IID_IOddNumber = { 0xe5970df2, 0x589f, 0x46b2, { 0xbd, 0xcf, 0x6d, 0xbd, 0x9b, 0x44, 0x96, 0x5b } };

// {2CD97B05-02D9-4908-BE07-1A6602358A88}
const IID IID_IOddNumberFactory = { 0x2cd97b05, 0x2d9, 0x4908, { 0xbe, 0x7, 0x1a, 0x66, 0x2, 0x35, 0x8a, 0x88 } };

const char g_szFriendlyName[] = "Aashish N. Ayyar : COM Project : MonikerServer";
const char g_szVerIndProgId[] = "COMProject.Moniker";

const char g_szProgId[]       = "COMProject.Moniker.1";

