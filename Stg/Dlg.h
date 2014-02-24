#pragma once
#include "resource.h"
#include "datacombo2.h"
#include "datagrid1.h"

// CDlg dialog

class CDlg : public CDialog
{
	DECLARE_DYNAMIC(CDlg)

public:
	CDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG1 };
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
	BOOL m_bFnd;
	BOOL m_bFndC;
	CString m_strFndC;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
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

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
	virtual BOOL OnInitDialog();
	CDatagrid1 m_DataGrid1;
	BOOL m_Check1;
	CString m_Edit1;
	CString m_Edit2;
protected:
	void OnOnlyRead(int i);
public:
	DECLARE_EVENTSINK_MAP()
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
};
