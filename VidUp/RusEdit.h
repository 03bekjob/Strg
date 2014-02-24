#if !defined(AFX_RUSEDIT_H__87556511_47F0_11D3_84BE_00A0C94CCA25__INCLUDED_)
#define AFX_RUSEDIT_H__87556511_47F0_11D3_84BE_00A0C94CCA25__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// RusEdit.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CRusEdit window

//Выводит все символы на Русском и перечисленные в OnChar()

class CRusEdit : public CEdit
{
// Construction
public:
	CRusEdit();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CRusEdit)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CRusEdit();

	// Generated message map functions
protected:
	//{{AFX_MSG(CRusEdit)
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_RUSEDIT_H__87556511_47F0_11D3_84BE_00A0C94CCA25__INCLUDED_)
