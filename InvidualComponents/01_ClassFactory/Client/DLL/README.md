To build this file 
cl /EHsc /LD ClassFactoryClient.cpp ClassFactoryClient.def

As soon as the Dll is loaded it will register the COM Server. 
And when freed the Server will be UnRegistered. 

Exports the methods 
SumOfTwoIntegers
SubtractionofTwoIntegers

A default interface is stored always so that calls to CoCreateInstance could be saved. 
Hence future calls will be directly to QueryInterface. 

