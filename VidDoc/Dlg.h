#pragma once

#include "resource.h"
#include "datacombo2.h"
#include "datagrid1.h"
#include "rusedit.h"
// CDlg dialog

class CDlg : public CDialog
{
	DECLARE_DYNAMIC(CDlg)

public:
	CDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG1 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	int m_iCurType;
	int m_CurCol;
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1;
	_CommandPtr ptrCmd1;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
	CEdit m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	CString m_strFndC;
	BOOL m_bFndC;
	CWnd* m_pLstWnd;
	BOOL m_bFnd;
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
protected:
	void OnEnableButtonBar(int iBtSt, CToolBar* pTlBr);
	void OnEnChangeEditTB(void);
	bool OnOverEdit(int IdBeg, int IdEnd);
	void OnShowEdit(void);
	void InitStaticText(void);
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void InitDataGrid1(void);
public:
	virtual BOOL OnInitDialog();
	CDatagrid1 m_DataGrid1;
	CString m_Edit1;
	BOOL m_Check1;
	BOOL m_Check2;
	BOOL m_Check3;
	CRusEdit m_Edit1Rs;
	DECLARE_EVENTSINK_MAP()
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
};
