// PrnVw.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>
#include "Dlg.h"

#ifdef _MANAGED
#error Please read instructions in PrnVw.cpp to compile with /clr
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


static AFX_EXTENSION_MODULE PrnVwDLL = { NULL, NULL };

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
		TRACE0("PrnVw.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(PrnVwDLL, hInstance))
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

		new CDynLinkLibrary(PrnVwDLL);

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
		TRACE0("PrnVw.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(PrnVwDLL);
	}
	return 1;   // ok
}

extern "C" _declspec(dllexport) BOOL startPrnVw(CString strName,_ConnectionPtr ptrCon,CString& strFndC,BOOL bFndC, CString& strSql, CString& strCod){

	CDlg dlg;
	int i;
	dlg.ptrCnn    = ptrCon;
	dlg.m_strNT   = strName;
	dlg.m_strCod  = strCod;

//AfxMessageBox(dlg.m_strCod);

	i = strFndC.Find(_T("~"));
	dlg.m_strFndC  = strFndC.Left(i);
	dlg.m_strFndC1 = strFndC.Mid(i+1);	// ��� ����� ��������
	dlg.m_bFndC   = bFndC;
//AfxMessageBox(strSql);
	i = strSql.Find(_T("~"));
	dlg.m_strSql  = strSql.Left(i);
	dlg.m_strSls  = strSql.Mid(i+1);

	if (dlg.DoModal()==IDOK) {
		if(bFndC) strFndC=dlg.m_strFndC;
		return TRUE;
	}
	else{ 
		if(bFndC) strFndC=dlg.m_strFndC;
		return FALSE;
	}
}

#ifdef _MANAGED
#pragma managed(pop)
#endif

