#if !defined(AFX_RUSNUMEDIT_H__E5683429_24A8_4DF6_BD04_58F121A073DE__INCLUDED_)
#define AFX_RUSNUMEDIT_H__E5683429_24A8_4DF6_BD04_58F121A073DE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// RusNumEdit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CRusNumEdit window

class CRusNumEdit : public CEdit
{
// Construction
public:
	CRusNumEdit();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CRusNumEdit)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CRusNumEdit();

	// Generated message map functions
protected:
	//{{AFX_MSG(CRusNumEdit)
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_RUSNUMEDIT_H__E5683429_24A8_4DF6_BD04_58F121A073DE__INCLUDED_)
