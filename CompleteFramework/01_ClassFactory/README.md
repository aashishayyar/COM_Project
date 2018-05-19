README :: ClassFactory Component

1. Maintain component and client directory (preferably just create a dir for component).
2. DllRegisterServer, DllUnRegisterServer are used in the component. 
3. Create the component.dll using the following command :
	cl /EHsc /LD Component.cpp Component.def 
	[Assumption] : a. user32.lib, gdi32.lib, ol32.lib, advapi32.lib are added as #pragma comment(lib, "<whatever>")
				   b. You have added DllRegisterServer and DllUnregisterServer

4. [Optional] : Copy the created dll to the parent directory. [Assumption] : Your client is present there.
5. Fire "resvr32 Component.dll". Check for successful addition in registry

6. 				   