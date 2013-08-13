//
// Rainmeter WMI plugin
// WMI helper class
// ---------------------------------------------
// E.Butusov <ebutusov@gmail.com>
//
// 25.03.2013

#include <sstream>
#include "WMIService.h"
#include "ComPtr.h"
#include "../SDK/API/RainmeterAPI.h"

extern IGlobalInterfaceTable *g_pGIT;

void ComFailLog(LPCWSTR msg, HRESULT code)
{
	std::wstringstream ss;
	ss << msg << " Error code = 0x" << std::hex << code;
	RmLog(LOG_ERROR,ss.str().c_str());
}

CWMIService::CWMIService(): m_pSvc(nullptr) 
{
}

CWMIService::~CWMIService()
{
	Release();
}

bool CWMIService::Connect(LPCWSTR wmi_namespace)
{
	// release current service, if any
	if (m_pSvc)
	{
		m_pSvc->Release();
		m_pSvc = nullptr;
	}

	HRESULT hres;
	IWbemLocator *locator = nullptr;

	hres = CoCreateInstance(
		CLSID_WbemLocator,             
		0, 
		CLSCTX_INPROC_SERVER, 
		IID_IWbemLocator, (LPVOID *) &locator);
 
  if (FAILED(hres))
  {
		ComFailLog(L"Failed to create IWbemLocator object!", hres);
		return false;
  }
  
	hres = locator->ConnectServer(
		_bstr_t(wmi_namespace), // Object path of WMI namespace
		NULL,                    // User name. NULL = current user
		NULL,                    // User password. NULL = current
		0,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		0,                       // Authority (for example, Kerberos)
		0,                       // Context object 
		&m_pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres))
	{
		ComFailLog(L"Failed to connect to WMI server!", hres);
		locator->Release();
		return false;
	}

  // set security levels on the proxy
  hres = CoSetProxyBlanket(
      m_pSvc,                        // Indicates the proxy to set
      RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
      RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
      NULL,                        // Server principal name 
      RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
      RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
      NULL,                        // client identity
      EOAC_NONE                    // proxy capabilities 
  );

  if (FAILED(hres))
  {
		ComFailLog(L"Failed to set proxy!", hres);
    m_pSvc->Release();
		m_pSvc = nullptr;
    locator->Release();
	}

	return true;
}

void CWMIService::Release()
{
	if (m_pSvc)
	{
		m_pSvc->Release();
		m_pSvc = nullptr;
	}
}

bool CWMIService::Exec(const std::wstring &wmi_query, 
	std::wstring &strResult, double &dblResult, EResultType &type)
{
	if (!m_pSvc)
	{
		_ASSERT(false);
		// invalid state, Connect not called or failed
		ComFailLog(L"Service error!", 666);
		return false;
	}

	CComPtr<IEnumWbemClassObject> pEnumerator;
	HRESULT hres = m_pSvc->ExecQuery(
      bstr_t("WQL"), 
      bstr_t(wmi_query.c_str()),
      WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, 
      NULL,
      &pEnumerator);
    
  if (FAILED(hres))
  {
		std::wstring err = L"WMI query failed: " + wmi_query;
		ComFailLog(err.c_str(), hres);
    return false;
  }

	//auto enumWrap = wrap_ptr(pEnumerator); // just for proper releasing
	bool has_result = false;

	while (pEnumerator)
  {
		CComPtr<IWbemClassObject> clsObj;
		ULONG ret = 0;

		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, &ret);
		
		if(ret == 0) break;

		BSTR propName = nullptr;

		// get first property (non-system, local) of returned object
		// by design there should be only one property
		if (SUCCEEDED(clsObj->BeginEnumeration(WBEM_FLAG_LOCAL_ONLY | WBEM_FLAG_NONSYSTEM_ONLY)))
		{
			clsObj->Next(0, &propName, NULL, NULL, NULL);
			clsObj->EndEnumeration();
		}

		if (propName == nullptr)
			break;

		VARIANT vtProp;
		::VariantInit(&vtProp);
		hr = clsObj->Get(propName, 0, &vtProp, 0, 0);			
		::SysFreeString(propName);

		if (FAILED(hr))
		{
			std::wstring err = L"Get property failed: " + wmi_query;
			ComFailLog(err.c_str(), hres);
			break;
		}
		else
		{
			has_result = true;
			switch (V_VT(&vtProp))
			{
			case VT_I1:
			case VT_I2:
			case VT_I4:
			case VT_I8:
				type = eNum;
				dblResult = (double)vtProp.intVal;
				break;
			case VT_BSTR:
				type = eString;
				strResult = vtProp.bstrVal;
				break;
			case VT_R4:
			case VT_R8:
				type = eNum;
				dblResult = vtProp.dblVal;
			case VT_BOOL:
				type = eString;
				// XXX maybe map to 1/0 instead of strings?
				strResult = vtProp.boolVal ? L"Yes" : L"No";
				break;
			default:
				type = eString;
				_ASSERT(false);
				strResult =  L"#TypeUnknown"; // just to know that we've got unhandled type
			}
		}
		::VariantClear(&vtProp);
  }
	return has_result;
}