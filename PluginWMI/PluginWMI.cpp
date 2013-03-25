//
// Rainmeter WMI plugin
// ---------------------------------------------
// E.Butusov <ebutusov@gmail.com>
//
// 25.03.2013

#include <windows.h>
#pragma comment(lib, "wbemuuid.lib")
#include <comdef.h>
#include <Wbemidl.h>
#include <string>
#include <sstream>
#include "../SDK/API/RainmeterAPI.h"
#include "WMIService.h"

int g_measuresCnt = 0;

CWMIService *g_pService;

struct Measure
{
	EResultType type;
	double dblResult;
	std::wstring strResult;
	std::wstring query;

	void QueryData(CWMIService *svc)
	{
		if (!svc->Exec(query, strResult, dblResult, type))
		{
			// something went wrong, no result returned
			// it could happen because of misspelled class name and/or property
			// show error tag instead of empty string
			type = eString;
			strResult = L"#None"; 
		}
	}
};

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	if (g_measuresCnt == 0)
	{
		// we assume that COM has been properly inialized for this thread
		// create and initialize global wmi query helper

		g_pService = new CWMIService();
		
		// even if it fails to Connect, we will leave it as is
		// Exec will return #ConnectError when called

		g_pService->Connect(L"root\\CIMV2");

		// for now only root\CIMV2 is allowed, but it should be easy to add
		// another parameter to ini file (WMINamespace) and allow to connect to other
		// namespaces
		// e.g. we can hold a map of CWMIService instances, or just put a ptr to the right
		// CWMIService instance into the Measure struct
	}

	Measure* measure = new Measure;
	*data = measure;
	++g_measuresCnt;
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	Measure* measure = (Measure*)data;
	measure->query = RmReadString(rm, L"Query", L"");
}

PLUGIN_EXPORT double Update(void* data)
{
	Measure* measure = (Measure*)data;
	measure->QueryData(g_pService);
	
	if (measure->type == eNum)
		return measure->dblResult;
	else return 0.0;
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	Measure* measure = (Measure*)data;
	if (measure->type == eString)
		return measure->strResult.c_str();
	else return nullptr; // rainmeter will treat this measure as number
}

//PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
//{
//	Measure* measure = (Measure*)data;
//}

PLUGIN_EXPORT void Finalize(void* data)
{
	Measure* measure = (Measure*)data;
	delete measure;
	--g_measuresCnt;
	// no more wmi measures, release helper
	if (g_measuresCnt == 0)
	{
		delete g_pService;
		g_pService = nullptr;
	}
}


// not used, just for the record

//void ComInit()
//{
//	HRESULT hres;
//
//    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 
//    if (FAILED(hres))
//    {
//			std::wstringstream ss;
//			ss << L"Failed to initialize COM library. Error code = 0x" << std::hex << hres;
//
//			RmLog(LOG_ERROR,ss.str().c_str());
//      return;
//    }
//
//    hres =  CoInitializeSecurity(
//        NULL, 
//        -1,                          // COM authentication
//        NULL,                        // Authentication services
//        NULL,                        // Reserved
//        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
//        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
//        NULL,                        // Authentication info
//        EOAC_NONE,                   // Additional capabilities 
//        NULL                         // Reserved
//        );
//
//                      
//    if (FAILED(hres))
//    {
//			std::wstringstream ss;
//			ss << L"Failed to initialize security. Error code = 0x" << std::hex << hres;
//
//			RmLog(LOG_ERROR,ss.str().c_str());
//      return;
//    }
//}