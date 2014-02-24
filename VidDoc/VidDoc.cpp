// VidDoc.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>
#include "Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

static AFX_EXTENSION_MODULE VidDocDLL = { NULL, NULL };

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("VidDoc.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(VidDocDLL, hInstance))
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

		new CDynLinkLibrary(VidDocDLL);

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
		TRACE0("VidDoc.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(VidDocDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startVidDoc(CString strName,_ConnectionPtr ptrCon,CString& strFndC,BOOL bFndC){
	CDlg dlg;
	dlg.ptrCnn    = ptrCon;
	dlg.m_strNT   = strName;
	dlg.m_strFndC = strFndC;
	dlg.m_bFndC   = bFndC;
	if (dlg.DoModal()==IDOK) {
		if(bFndC) strFndC=dlg.m_strFndC;
		return TRUE;
	}
	else{ 
		if(bFndC) strFndC=dlg.m_strFndC;
		return FALSE;
	}
}