#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"
#include "rusnumedit.h"
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
public:
	int m_iCurType;
	int m_CurCol;
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1,ptrRsCmb1,ptrRsCmb2,ptrRsCmb3,ptrRsCmb4,ptrRsCmb5,\
				  ptrRsMax;
	_CommandPtr ptrCmd1,ptrCmdCmb1,ptrCmdCmb2,ptrCmdCmb3,ptrCmdCmb4,ptrCmdCmb5,\
				ptrCmdMax;
    _ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
//	CEdit m_EditTB;
	CEditCh m_EditTBCh;//,m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;
	bool m_bCombo3,m_bCombo2;
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

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support


	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();
	CDatagrid1 m_DataGrid1;
	CDatacombo2 m_DataCombo1;
	CDatacombo2 m_DataCombo2;
	CDatacombo2 m_DataCombo3;
	CDatacombo2 m_DataCombo4;
	CDatacombo2 m_DataCombo5;

	DECLARE_EVENTSINK_MAP()
protected:
	void OnShowComboGlb(long cd);
	long OnShowCombo(void);
	void OnUpdateCombo2(void);
	void OnUpdateCombo3(void);
	void OnUpdateCombo4(void);
	void OnUpdateCombo5(long cd);
public:
	void ChangeDatacombo1();
	void ChangeDatacombo2();
	void ChangeDatacombo3();
	void ChangeDatacombo4();
	void ChangeDatacombo5();
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	virtual BOOL OnInitDialog();
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;

protected:
	void OnShowCladr(_RecordsetPtr ptrRs,long cd);
	void OnShowCladrRgn(void);
	void OnUpdateGrid1Rgn();
	void UpdateGrid1(long i,_RecordsetPtr rsCmb1,_RecordsetPtr rsCmb2,CString strSql,CString strSqlMax);
	void UpdateGrid1Glb(void);
public:
	CRusNumEdit m_Edit1RusNum;
protected:
	void OnLockCombo(void);
	void OnUnLockCombo(void);
};
