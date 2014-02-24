#pragma once
#include "datagrid1.h"
#include "datacombo2.h"
#include "resource.h"
#include "rusnumedit.h"
/*#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\rdlldb\datacombo2.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\rdlldb\datacombo2.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\rdlldb\datacombo2.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\prj_dlg\prj_dlg\datagrid1.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\prj_dlg\prj_dlg\datagrid1.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\prj_dlg\prj_dlg\datagrid1.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\prj_dlg\prj_dlg\datagrid1.h"
#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\prj_dlg\prj_dlg\datagrid1.h"
*/
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
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRs4,ptrRs5,ptrRs6,\
				  ptrRsCmb1,ptrRsCmb2,ptrRsCmb3;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmd4,ptrCmd5,ptrCmd6,\
				ptrCmdCmb1,ptrCmdCmb2,ptrCmdCmb3;
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
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr& ),int EmpCol=0,int DefCol=1);
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid4(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid5(CDatagrid1& pG,_RecordsetPtr& rs);
	static void InitDataGrid6(CDatagrid1& pG,_RecordsetPtr& rs);
public:
	DECLARE_EVENTSINK_MAP()
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
protected:
	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);
public:
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
	CString m_Edit5;
	CString m_Edit6;
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
	CDatacombo2 m_DataCombo3;
	CDatagrid1 m_DataGrid1;
	CDatagrid1 m_DataGrid2;
	CDatagrid1 m_DataGrid3;
	CDatagrid1 m_DataGrid4;
	CDatagrid1 m_DataGrid5;
	CDatagrid1 m_DataGrid6;
public:
	void ChangeDatacombo1();
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButton6();
	afx_msg void OnBnClickedButton7();
	afx_msg void OnBnClickedButton8();
	afx_msg void OnBnClickedButton9();
	afx_msg void OnBnClickedButton10();
	afx_msg void OnBnClickedButton11();
	afx_msg void OnBnClickedButton12();
	afx_msg void OnBnClickedButton13();
	afx_msg void OnBnClickedCheck1();
protected:
	void OnShowNameCombo(void);
	CString OnQuerySex(bool  ch, CString& str);
	void OnShowAddress(void);
	void OnShowTelFax(void);
	void OnShowEmail(void);
	void OnShowDoc(void);
	void OnShowRasSch(void);
	void OnShowEnter(long cd);

public:
	int m_cd;
public:
	BOOL m_Check1;
	BOOL m_Check2;
	CRusNumEdit m_Edit1RsNm;
};
