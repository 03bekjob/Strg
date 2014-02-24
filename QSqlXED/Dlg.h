#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"

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
	_RecordsetPtr ptrRsCmb1;
	_CommandPtr ptrCmdCmb1;
	_ConnectionPtr ptrCnn;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;
	CString m_strFndC;
	BOOL m_bFndC;
	CWnd* m_pLstWnd;
	

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

protected:
	void InitStaticText(void);

	void OnOnlyRead(int i);
	void SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt);
	BOOL IsQuery(CString substr, _RecordsetPtr rs);
	void OnShowObject(int i);
	void OnShowEmpty(void);

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnClickedOk();
	afx_msg void OnClickedCancel();

public:
	DECLARE_EVENTSINK_MAP()
public:
	CDatacombo2 m_DataCombo1;
	COleDateTime m_OleDateTime1;
	COleDateTime m_OleDateTime2;
	CString m_Edit1;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
	CString m_Edit5;
	CString m_Edit6;
public:
	void ChangeDatacombo1();
public:
	int m_iCdQr;
public:
	afx_msg void OnEnChangeEdit2();
public:
	afx_msg void OnEnChangeEdit1();
public:
	afx_msg void OnEnChangeEdit5();
public:
	CString m_strSql;
public:
	CString m_strX;
public:
	CString m_strHd;
};
