// QSqlX.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>
#include "Dlg.h"
#ifdef _MANAGED
#error Please read instructions in QSqlX.cpp to compile with /clr
// If you want to add /clr to your project you must do the following:
//	1. Remove the above include for afxdllx.h
//	2. Add a .cpp file to your project that does not have /clr thrown and has
//	   Precompiled headers disabled, with the following text:
//			#include <afxwin.h>
//			#include <afxdllx.h>
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


static AFX_EXTENSION_MODULE QSqlXDLL = { NULL, NULL };

#ifdef _MANAGED
#pragma managed(push, off)
#endif

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("QSqlX.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(QSqlXDLL, hInstance))
			return 0;

		// Insert this DLL into the resource chain
		// NOTE: If this Extension DLL is being implicitly linked to by
		//  an MFC Regular DLL (such as an ActiveX Control)
		//  instead of an MFC application, then you will want to
		//  remove this line from DllMain and put it in a separate
		//  function exported from this Extension DLL.  The Regular DLL
		//  that uses this Extension DLL should then explicitly call that
		//  function to initialize this Extension DLL.  Otherwise,
		//  the CDynLinkLibrary object will not be attached to the
		//  Regular DLL's resource chain, and serious problems will
		//  result.

		new CDynLinkLibrary(QSqlXDLL);

		// Sockets initialization
		// NOTE: If this Extension DLL is being implicitly linked to by
		//  an MFC Regular DLL (such as an ActiveX Control)
		//  instead of an MFC application, then you will want to
		//  remove the following lines from DllMain and put them in a separate
		//  function exported from this Extension DLL.  The Regular DLL
		//  that uses this Extension DLL should then explicitly call that
		//  function to initialize this Extension DLL.
		if (!AfxSocketInit())
		{
			return FALSE;
		}
	
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("QSqlX.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(QSqlXDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startQSqlX(CString& strSql, _ConnectionPtr ptrCon, CString strX, CString strHd){

	CDlg dlg;
	dlg.m_strSql  = strSql;
	dlg.ptrCnn    = ptrCon;
	dlg.m_strX  = strX;
	dlg.m_strHead = strHd;
//AfxMessageBox(dlg.m_strHead);
	if (dlg.DoModal()==IDOK) {
		strSql = dlg.m_strSql;
//AfxMessageBox(strSql);
		return TRUE;
	}
	else {
		return FALSE;
	}

}

#ifdef _MANAGED
#pragma managed(pop)
#endif

