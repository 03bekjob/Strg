#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"
#include "editch.h"


// CDlg dialog
class CDlgFtch;

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
	_RecordsetPtr ptrRs1,ptrRsCmb1,ptrRsCmb2,ptrRsCmb3,ptrRsCmb4,ptrRsCmb5,\
						 ptrRsCmb6,ptrRsCmb7,ptrRsCmb8,ptrRsCmb9;
	_CommandPtr ptrCmd1,ptrCmdCmb1,ptrCmdCmb2,ptrCmdCmb3,ptrCmdCmb4,ptrCmdCmb5,\
						ptrCmdCmb6,ptrCmdCmb7,ptrCmdCmb8,ptrCmdCmb9;
    _ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
//	CEdit m_EditTB;
	CEditCh m_EditTBCh;//,m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	bool m_bFnd;
	virtual BOOL OnInitDialog();
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
	int cd;
	CDatagrid1 m_DataGrid1;
	CString m_Edit1;
	CString m_Edit5;
	CString m_Edit6;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
	CDatacombo2 m_DataCombo1,m_DataCombo2,m_DataCombo3,m_DataCombo4,m_DataCombo5,\
				m_DataCombo6,m_DataCombo7,m_DataCombo8,m_DataCombo9;

	DECLARE_EVENTSINK_MAP()
protected:
	void OnUpdateCombo2(void);
	void OnUpdateCombo3(void);
	void OnUpdateCombo4(void);
	long OnShowCombo(void);
	void UpdateGrid1(long i);
	void UpdateGrid1Glb(void);
	void OnShowComboGlb(long cd);
public:
	bool m_bCombo3;
	void ChangeDatacombo1();
	void ChangeDatacombo2();
	void ChangeDatacombo3();
	void ChangeDatacombo4();
	void ChangeDatacombo5();
	void ChangeDatacombo6();
	void ChangeDatacombo7();
	void ChangeDatacombo8();
	void ChangeDatacombo9();
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
	CString m_strCod;
	CString m_strIns;
protected:
	void OnShowIndex(void);
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	CDlgFtch* m_pDlg;
	afx_msg LRESULT OnFetch(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCloseFetch(WPARAM wParam, LPARAM lParam);
};
