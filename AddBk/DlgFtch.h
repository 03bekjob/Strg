#pragma once

#define WM_FTCH      WM_USER + 5
#define WM_CLOSEFTCH WM_USER + 6

// CDlgFtch dialog
#include "resource.h"
class CDlgFtch : public CDialog
{
	DECLARE_DYNAMIC(CDlgFtch)

public:
	CDlgFtch(CWnd* pParent = NULL);   // standard constructor
	CDlgFtch(CDialog* pDlg);
	virtual ~CDlgFtch();

// Dialog Data
	enum { IDD = IDD_DIALOG2 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CString m_Edit1;
	CDialog* m_pDlg;
public:
	virtual BOOL OnInitDialog();
	BOOL Create();

public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
};
