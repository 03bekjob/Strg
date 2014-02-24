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
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRsCmb1,ptrRsCmb2,ptrRsCmb3,ptrRsCmb4;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmdCmb1,ptrCmdCmb2,ptrCmdCmb3,ptrCmdCmb4;
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
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr& ),int EmpCol=0,int DefCol=1);

	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);
	BOOL LoadFont(int h);
	void OnCalcSls(int i, CString strBsCst, CString strK);

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

	afx_msg void OnBnClickedRadio1();
	afx_msg void OnBnClickedRadio2();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnEnChangeEdit2();

public:
	CDatagrid1 m_DataGrid1;
	CDatagrid1 m_DataGrid2;
	CDatagrid1 m_DataGrid3;
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
	CDatacombo2 m_DataCombo3;
	CDatacombo2 m_DataCombo4;
protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs);
public:
	DECLARE_EVENTSINK_MAP()
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid2(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid3(VARIANT* LastRow, short LastCol);
public:
	CString m_Edit1;
	CString m_Edit2;
	COleDateTime m_OleDateTime1;
	int m_Radio1;
public:
	HICON m_Icon;
	CString m_stcSlsCst;
	CString m_strSqlOld;
	CString m_strSqlOldF;
public:
	void ClickDatagrid1();
	void ClickDatagrid2();
	void ClickDatagrid3();
	void ChangeDatacombo1();
	void ChangeDatacombo2();
	void ChangeDatacombo3();
	void ChangeDatacombo4();
public:
	long m_Cd;
	long m_CdOp;
	BOOL m_bAuto;
	CFont m_curFont;
	LOGFONT m_OptionFont;
	CEditFloat m_Edit1Flt;
	CEditFloat m_Edit2Flt;
	CString m_strCst;
public:
	afx_msg void OnBnClickedButton2();
};
