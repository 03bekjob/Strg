// Strg.h : main header file for the Strg application
//
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"       // main symbols


// CStrgApp:
// See Strg.cpp for the implementation of this class
//

class CStrgApp : public CWinApp
{
public:
	CStrgApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CStrgApp theApp;