//
// Rainmeter WMI plugin
// WMI helper class
// ---------------------------------------------
// E.Butusov <ebutusov@gmail.com>
//
// 25.03.2013

#pragma once
#include <windows.h>
#pragma comment(lib, "wbemuuid.lib")
#include <comdef.h>
#include <Wbemidl.h>
#include <string>

enum EResultType { eString, eNum, eWait };

class CWMIService
{
	IWbemServices *m_pSvc;
	DWORD m_cookie;
	CWMIService(CWMIService &other); // no copy
public:
	CWMIService();
	~CWMIService();

	bool Connect(LPCWSTR wmi_namespace);
	void Release();
	bool Exec(const std::wstring &wmi_query,
		std::wstring &strResult, double &dblResult, EResultType &type);
};
