#pragma once
#include "datagrid1.h"
#include "datacombo2.h"
#include "resource.h"
#include "editfloat.h"
#include "editch.h"
//#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\stg\datagrid1.h"
//#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\stg\datacombo2.h"
//#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\stg\datacombo2.h"

// CDlg dialog

class CDlg : public CDialog
{
	DECLARE_DYNAMIC(CDlg)

public:
	CDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CDlg();
//	typedef void (*pFGrd)();

// Dialog Data
	enum { IDD = IDD_DIALOG1 };
public:
	int m_iCurType;
	int m_CurCol;
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRsCmb1,ptrRsCmb2;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmdCmb1,ptrCmdCmb2;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
//	CEdit m_EditTB;
	CEditCh m_EditTBCh;//,m_EditTB;
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
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr& ));
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

public:
	CDatagrid1 m_DataGrid1;
protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs);
public:
	DECLARE_EVENTSINK_MAP()
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
public:
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
	CString m_Edit5;
	CString m_Edit6;
	CString m_Edit7;
	CString m_Edit8;
	BOOL m_Check1;
public:
	CDatagrid1 m_DataGrid2;
	CDatagrid1 m_DataGrid3;
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	void ClickDatagrid3();
	void ClickDatagrid2();
	void ClickDatagrid1();
	void ChangeDatacombo1();
	void ChangeDatacombo2();
public:
	CEditFloat m_Edit5F;
	CEditFloat m_Edit6F;
	CEditFloat m_Edit7F;
	CEditFloat m_Edit8F;
public:
};
