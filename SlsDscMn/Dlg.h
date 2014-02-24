#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"
#include "editfloat.h"

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
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRs4,ptrRsCmb1,ptrRsCmb2;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmd4,ptrCmdCmb1,ptrCmdCmb2;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
	CEdit m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;
	CString m_strFndC;
	BOOL m_bFndC;
	CWnd* m_pLstWnd;
	HICON m_Icon;

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
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr& ),int EmpCol=0,int DefCol=1);

	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
	afx_msg void OnBnClickedRadio1();
	afx_msg void OnBnClickedRadio2();

public:
	CDatagrid1 m_DataGrid1;
	CDatagrid1 m_DataGrid2;
	CDatagrid1 m_DataGrid3;
	CDatagrid1 m_DataGrid4;
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid4(CDatagrid1& pG,_RecordsetPtr& rs);
	DECLARE_EVENTSINK_MAP()
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid3(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid4(VARIANT* LastRow, short LastCol);
public:
	int m_Radio1;
	BOOL m_Check1;
	COleDateTime m_OleDateTime1;
	COleDateTime m_OleDateTime2;
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
public:
	void ClickDatagrid1();
	void ClickDatagrid2();
	void ClickDatagrid3();
	void ClickDatagrid4();
	void ChangeDatacombo1();
	void ChangeDatacombo2();
public:
	CString m_strBskt; // корзина
	CString m_strFtch; // выборка
	CEditFloat m_Edit1Flt;
	CEditFloat m_Edit2Flt;
	CEditFloat m_Edit3Flt;
	CEditFloat m_Edit4Flt;
public:
	afx_msg void OnBnClickedButton2();
public:
	afx_msg void OnBnClickedButton3();
public:
	afx_msg void OnBnClickedButton1();
protected:
	BOOL IsQuery(CString substr, _RecordsetPtr rs);
public:
	afx_msg void OnBnClickedCheck1();
};
