// RusEdit.cpp : implementation file
//

#include "stdafx.h"
//#include "	\ add additional includes here"
#include "RusEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CRusEdit

CRusEdit::CRusEdit()
{
}

CRusEdit::~CRusEdit()
{
}


BEGIN_MESSAGE_MAP(CRusEdit, CEdit)
	//{{AFX_MSG_MAP(CRusEdit)
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CRusEdit message handlers

void CRusEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
//Кодовая таблица 1251- MS Windows теперь UniCod !!!
/*	CString s;
	s.Format(L"%i",nChar);
	AfxMessageBox(s);
*/
	if (nChar>=1040 && nChar<=1103)
	{
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	else
		switch(nChar)
	{
		case 19:	// <-
		case 4:		// ->
		case 8:		// <- Backspace
		case 32:	// space
		case 34:	// "
		case 37:	// %
		case 44:	// ,
		case 45:	// -
		case 46:	// .
		case 47:	// /
		case 127:	// delete
		case 1025:	// Ё
		case 1105:	// ё
		case 8470:	// №
			CEdit::OnChar(nChar, nRepCnt, nFlags);
			break;
	}
}
