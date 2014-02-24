#pragma once
#include "datagrid1.h"
#include "datacombo2.h"
#include "resource.h"
//#include "c:\documents and settings\bek\my documents\visual studio 2005\projects\strg\rdlldb\datacombo2.h"

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
	_RecordsetPtr ptrRsCmb1;
	_CommandPtr ptrCmd1,ptrCmdCmb1;
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
	void OnShowEnter(void);
	void InitStaticText(void);
	void On32778(void);
	void On32783(void);
	void On32775(void);
	void On32776(void);
	void On32777(void);
	void On32779(void);
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

public:
	DECLARE_EVENTSINK_MAP()
protected:
	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);
public:
	CDatacombo2 m_DataCombo1;
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	COleDateTime m_OleDateTime1;
	long m_cd;
	CString m_strCod;
public:
	afx_msg void OnBnClickedButton1();
public:
	void ChangeDatacombo1();
};
