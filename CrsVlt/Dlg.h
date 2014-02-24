#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"
#include "atlcomtime.h"
#include "editch.h"


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

public:
	int m_iCurType;
	int m_CurCol;
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1,ptrRsCmb1,ptrRsCmb2;
	_CommandPtr ptrCmd1,ptrCmdCmb1,ptrCmdCmb2;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
	CEditCh m_EditTBCh;//,m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
	CDatagrid1 m_DataGrid1;
	DECLARE_EVENTSINK_MAP()
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
	void ChangeDatacombo2();
	void ChangeDatacombo1();
	COleDateTime m_OleDateTime1;
	CString m_Edit1;

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
	void OnShowGrid1(void);
public:
	virtual BOOL OnInitDialog();
};
