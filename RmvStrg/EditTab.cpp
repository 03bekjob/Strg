// EditTab.cpp : implementation file
//

#include "stdafx.h"
//#include "Dlg.h"

#include "EditTab.h"



// CEditTab

IMPLEMENT_DYNAMIC(CEditTab, CEdit)

CEditTab::CEditTab()
{

}

CEditTab::~CEditTab()
{
}


BEGIN_MESSAGE_MAP(CEditTab, CEdit)
//	ON_WM_CHAR()
END_MESSAGE_MAP()

// CEditTab message handlers


//void CEditTab::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
//{
//AfxMessageBox(_T("CEditTab::OnChar"));
//	if(nChar == WM_KEYDOWN/*0x09*/){
//		AfxMessageBox(_T("CEditTab::OnChar Tab  1"));
//		CWnd* pParent;
//		pParent = GetParent();
//		pParent->PostMessageW(WM_SWITCHFOCUS,(WPARAM)((CWnd*)this),0);
//	}
//	else{
//		CEdit::OnChar(nChar, nRepCnt, nFlags);
//	}
//}


BOOL CEditTab::PreTranslateMessage(MSG* pMsg)
{
	if(pMsg->message==WM_KEYDOWN)
    {
//AfxMessageBox(_T("CEditTab::PreTranslateMessage"));
		CWnd* pParent;
		pParent = GetParent();
		switch(pMsg->wParam){
			case VK_TAB:
//AfxMessageBox(_T("VK_TAB"));
				pParent->PostMessageW(WM_SWITCHFOCUS,(WPARAM)((CWnd*)this),0);
				break;
			case VK_RETURN:
//AfxMessageBox(_T("VK_RETURN"));
				pParent->PostMessageW(WM_ADDREC,(WPARAM)((CWnd*)this),0);
				break;
		}	
	}
	return CEdit::PreTranslateMessage(pMsg);
}
