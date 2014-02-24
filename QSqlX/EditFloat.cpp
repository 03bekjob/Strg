// EditFloat.cpp : implementation file
//

#include "stdafx.h"
//#include "	\ add additional includes here"
#include "EditFloat.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CEditFloat

CEditFloat::CEditFloat()
{
}

CEditFloat::~CEditFloat()
{
}


BEGIN_MESSAGE_MAP(CEditFloat, CEdit)
	//{{AFX_MSG_MAP(CEditFloat)
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CEditFloat message handlers

void CEditFloat::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
//¬вод чисел  с плавающей точкой	
	if (nChar>=48 && nChar<=57)
	{
		CEdit::OnChar(nChar, nRepCnt, nFlags);
	}
	else
		switch(nChar)
	{
		case 19:	// <-
		case 4:		// ->
		case 8:		// <- Backspace
		case 127:	// delete
		case 44:	// ,
		case 46:	// .
			CEdit::OnChar(nChar, nRepCnt, nFlags);
			break;
	}
}
