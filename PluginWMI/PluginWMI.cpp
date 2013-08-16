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
#include <thread>
#include <mutex>
#include <queue>
#include <condition_variable>
#include "../SDK/API/RainmeterAPI.h"
#include "WMIService.h"

extern void ComFailLog(LPCWSTR msg, HRESULT code);

static int g_measuresCnt = 0;
static bool g_run = true;
static std::mutex g_mutex;
static std::condition_variable g_cv;
static std::thread g_worker;

struct Measure
{
	Measure() : refreshInterval(60), lastUpdate(0) {}

	EResultType type;
	double dblResult;
	std::wstring strResult;
	std::wstring query;
	int refreshInterval;
	int lastUpdate;
	bool queued;
};

static std::queue<Measure*> g_queue;

void QueryWorker()
{
	::CoInitialize(NULL);
	
	CWMIService srv;
	srv.Connect(L"root\\CIMV2");

	while (g_run)
	{
		Measure *m;
		{
			std::unique_lock<std::mutex> lock(g_mutex);
			if (g_queue.empty())
				g_cv.wait(lock, [] { return !g_queue.empty() || !g_run; });
			if (!g_run) break;
			m = g_queue.front();
			g_queue.pop();
		} // unlock

		if (!srv.Exec(m->query, m->strResult, m->dblResult, m->type))
		{
			// show error tag
			m->type = eString;
			m->strResult = L"#Error";
		}
		m->queued = false;
	}
}

void EnqueueForUpdate(Measure *m)
{
	m->queued = true;
	std::unique_lock<std::mutex> lg(g_mutex);
	g_queue.push(m);
	g_cv.notify_one();
}

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	if (g_measuresCnt == 0) // first time initialization
	{
		// start WMI query thread
		// waits for main thread to put Measures on g_queue for processing
		g_worker.swap(std::thread(QueryWorker));
	}

	Measure* measure = new Measure;
	*data = measure;
	++g_measuresCnt;
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	Measure* measure = (Measure*)data;
	measure->refreshInterval = RmReadInt(rm, L"Refresh", 60);
	measure->query = RmReadString(rm, L"Query", L"");
	measure->lastUpdate = 0;
}

PLUGIN_EXPORT double Update(void* data)
{
	Measure* m = (Measure*)data;
	DWORD now = ::GetTickCount() / 1000;
	if (now - m->lastUpdate >= m->refreshInterval)
	{
		Measure copy = *m;
		EnqueueForUpdate(m);
		m->lastUpdate = now;
		m = &copy;
	}
	if (m->type == eNum)
		return m->dblResult;
	return 0.0;
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	Measure* measure = (Measure*)data;
	if (!measure->queued)
	{
		if (measure->type == eString)
			return measure->strResult.c_str();
		else return nullptr; // rainmeter will treat this measure as number
	}
	else return L"#Wait";
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
		// stop queue thread
		g_run = false;
		g_cv.notify_one();
		g_worker.join();
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