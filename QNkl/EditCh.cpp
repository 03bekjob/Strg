// EditCh.cpp : implementation file
//

#include "stdafx.h"
#include "EditCh.h"


// CEditCh

IMPLEMENT_DYNAMIC(CEditCh, CEdit)

CEditCh::CEditCh()
: m_iTp(0)
{

}

CEditCh::~CEditCh()
{
}


BEGIN_MESSAGE_MAP(CEditCh, CEdit)
	ON_WM_CHAR()
END_MESSAGE_MAP()



// CEditCh message handlers



void CEditCh::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
/*	CString s;
	s.Format(_T("OnChar m_iTp = %i"),m_iTp);
AfxMessageBox(s);
*/
	switch(m_iTp){
		case adSmallInt:			//2
		case adInteger:				//3
		case adBoolean:				//11
		case adTinyInt:				//16
		case adUnsignedTinyInt:		//17
		case adUnsignedSmallInt:	//18
		case adUnsignedInt:			//19
//AfxMessageBox("adBoolean:");
		
			if (nChar>=48 && nChar<=57) //÷ифры
			{
				CEdit::OnChar(nChar, nRepCnt, nFlags);
				break;
			}
			break;
		case adDecimal:		//14
		case adNumeric:	//131
		case adDouble:
		case adCurrency:
//			AfxMessageBox(_T("¬вод чисел  с плавающей точкой"));
		//¬вод чисел  с плавающей точкой	
			if (nChar>=48 && nChar<=57) //÷ифры
			{
				CEdit::OnChar(nChar, nRepCnt, nFlags);
				break;
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
			break;
		case adChar:
		case adVarChar:
		case adBSTR:
//AfxMessageBox("adChar");
			CEdit::OnChar(nChar, nRepCnt, nFlags);
			break;
	}

}

void CEditCh::SetTypeCol(int i)
{
	m_iTp = i;

/*	CString s;
	s.Format(_T("CEditCh::SetTypeCol m_iTp = %i"),m_iTp);
AfxMessageBox(s);
*/
}
