// RDllDb.h : main header file for the RDllDb DLL
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols
//#include "datacombo2.h"


// CRDllDbApp
// See RDllDb.cpp for the implementation of this class
//

class CRDllDbApp : public CWinApp
{
public:
	CRDllDbApp();

// Overrides
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
	// Функция возвр итсину если rs пуст
	
};

