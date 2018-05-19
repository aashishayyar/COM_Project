#define MAX_INFO_COUNT   20
#define MAX_INFO_LENGTH  256

typedef struct stInformation
{
	WCHAR szClassVariable[MAX_INFO_LENGTH];
	WCHAR szOutput[MAX_INFO_LENGTH];
}PROCESSOR_INFO[MAX_INFO_COUNT];

PROCESSOR_INFO stProcessorInfo;

int index = -1;
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"Name");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"Caption");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"L2CacheSize");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"L2CacheSpeed");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"L3CacheSize");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"L3CacheSpeed");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"NumberOfCores");
StringCchCopy(stProcessorInfo[index++].szClassVariable, MAX_INFO_LENGTH, L"NumberOfLogicalProcessors");

