// RusNumEdit.cpp : implementation file
//

#include "stdafx.h"
//#include "	\ add additional includes here"
#include "RusNumEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CRusNumEdit

CRusNumEdit::CRusNumEdit()
{
}

CRusNumEdit::~CRusNumEdit()
{
}


BEGIN_MESSAGE_MAP(CRusNumEdit, CEdit)
	//{{AFX_MSG_MAP(CRusNumEdit)
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CRusNumEdit message handlers

void CRusNumEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
/*	CString s;
	s.Format(L"%i",nChar);
	AfxMessageBox(s);
*/
	if (nChar>=48 && nChar<=57){
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	else if(nChar>=1040 && nChar<=1103){ // Unicod , à ýòî ASCII -> nChar>=192 && nChar<=255
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	else{
		switch(nChar){
// Unicod
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
		case 1025:	// ¨
		case 1105:	// ¸
		case 8470:	// ¹

/*	ASCII
			case 19:	// <-
			case 4:		// ->
			case 8:		// <- Backspace
			case 32:	// space
			case 45:	// -
			case 47:	// /
			case 127:	// delete
			case 44:	// ,
			case 46:	// .
			case 168:	// ¨
			case 184:	// ¸
*/
			CEdit::OnChar(nChar, nRepCnt, nFlags);
			break;
		}
	}
}
