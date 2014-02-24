#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"
#include "editch.h"
#include "afxwin.h"

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
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRs4;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmd4;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
//	CEdit m_EditTB;
	CEditCh m_EditTBCh;//,m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;
	CString m_strFndC,m_strFndC1;
	BOOL m_bFndC;
	CWnd* m_pLstWnd;
	

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

protected:
	void OnEnableButtonBar(int iBtSt, CToolBar* pTlBr);
	void OnEnChangeEditTB(void);
	bool OnOverEdit(int IdBeg, int IdEnd);
	void OnShowEdit(void);
	void InitStaticText(void);
	void On32771(void);
	void On32772(void);
	void On32773(void);
	void On32774(void);
	void On32778(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void On32780(void);
	void On32783(void);
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr& ),int EmpCol=0,int DefCol=1);

	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);
	BOOL IsQuery(CString substr, _RecordsetPtr rs);

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

	afx_msg void OnBnClickedRadio1();
	afx_msg void OnBnClickedRadio2();
	afx_msg void OnBnClickedRadio3();

	afx_msg void OnBnClickedCheck1();
	afx_msg void OnBnClickedCheck2();
	afx_msg void OnBnClickedCheck3();
	afx_msg void OnBnClickedCheck4();
	afx_msg void OnBnClickedCheck5();
	afx_msg void OnBnClickedCheck6();
	afx_msg void OnBnClickedCheck7();
	afx_msg void OnBnClickedCheck8();

	afx_msg void OnCbnEditchangeCombo1();
	afx_msg void OnCbnEditchangeCombo2();
	afx_msg void OnCbnEditchangeCombo3();
	afx_msg void OnCbnEditchangeCombo4();
	afx_msg void OnCbnEditchangeCombo5();
	afx_msg void OnCbnEditchangeCombo6();
	afx_msg void OnCbnEditchangeCombo7();
	afx_msg void OnCbnEditchangeCombo8();

	afx_msg void OnCbnSelendokCombo1();
	afx_msg void OnCbnSelendokCombo2();
	afx_msg void OnCbnSelendokCombo3();
	afx_msg void OnCbnSelendokCombo4();
	afx_msg void OnCbnSelendokCombo5();
	afx_msg void OnCbnSelendokCombo6();
	afx_msg void OnCbnSelendokCombo7();
	afx_msg void OnCbnSelendokCombo8();

	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
public:
	CDatagrid1 m_DataGrid1,m_DataGrid2,m_DataGrid3;
protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid4(CDatagrid1& pG,_RecordsetPtr& rs);
public:
	DECLARE_EVENTSINK_MAP()
public:
	void ClickDatagrid4();
	void ClickDatagrid3();
	void ClickDatagrid1();
	void ClickDatagrid2();
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid2(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid3(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid4(VARIANT* LastRow, short LastCol);
public:
	int m_Radio1;
	BOOL m_Check1;
	BOOL m_Check2;
	BOOL m_Check3;
	BOOL m_Check4;
	BOOL m_Check5;
	BOOL m_Check6;
	BOOL m_Check7;
	BOOL m_Check8;
	CComboBox m_ComboBox1;
	CComboBox m_ComboBox2;
	CComboBox m_ComboBox3;
	CComboBox m_ComboBox4;
	CComboBox m_ComboBox5;
	CComboBox m_ComboBox6;
	CComboBox m_ComboBox7;
	CComboBox m_ComboBox8;

	CString m_Edit1;
public:
	CString m_strSql,m_strSls;
public:
	void OnEnableCombo(void);
	void OnEnableState(void);
public:
	CDatagrid1 m_DataGrid4;
public:
	CString m_strCod;
};
