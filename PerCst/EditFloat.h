#if !defined(AFX_EDITFLOAT_H__CFC5C1B9_F339_11D3_A883_B744F76BD944__INCLUDED_)
#define AFX_EDITFLOAT_H__CFC5C1B9_F339_11D3_A883_B744F76BD944__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// EditFloat.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CEditFloat window

class CEditFloat : public CEdit
{
// Construction
public:
	CEditFloat();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEditFloat)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CEditFloat();

	// Generated message map functions
protected:
	//{{AFX_MSG(CEditFloat)
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EDITFLOAT_H__CFC5C1B9_F339_11D3_A883_B744F76BD944__INCLUDED_)
