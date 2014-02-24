#pragma once
#include "resource.h"
#include "datagrid1.h"
#include "datacombo2.h"

#include "editch.h"

#include "afxwin.h"
#include <afxtempl.h>
#include "afxdtctl.h"

#include "afxcmn.h"
#include "afxext.h"
#include "atlcomtime.h"

#include "edittab.h"

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
	int m_iCurType,m_befCurType; //тип кот был до ins,upd,del
	int m_CurCol,m_befCurCol;		 //колонкак кот была до ins,upd,del
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRs4,ptrRs5,ptrRsCmb1,ptrRsCmb2,\
				  ptrRsCst,ptrRsExp,ptrRsInc,ptrRsDtSrv,ptrRs100,\
				  ptrRsNum;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmd4,ptrCmd5,ptrCmdCmb1,ptrCmdCmb2,ptrCmdCst,\
				ptrCmdExp,ptrCmdInc,ptrCmdDtSrv,ptrCmd100,ptrCmdNum;
	_ConnectionPtr ptrCnn;
	COleVariant m_vBkQr,m_vBkDb,m_vBkBskt;
	CString m_strNT;
	HINSTANCE hInstResClnt;
	CToolBar m_wndToolBar;
//	CEdit m_EditTB;
	CEditCh  m_EditTBCh;
	CEditTab m_Edit2T,m_Edit3T,m_Edit4T; //,m_EditTB;
	int m_iBtSt;
	bool m_Flg;
	BOOL m_bFnd;
	CString m_strFndC;
	BOOL m_bFndC;
	CWnd* m_pLstWnd;

	CArray<double,double&> aSldr;
	HICON m_Icon;
	CFont m_curFont;
	LOGFONT m_OptionFont;
	COLORREF m_curColor;
	BOOL m_bCalcBskt,m_bCalcDb,m_bCalcE9,m_bCalcE10,m_bCalcE11;
	BOOL m_bBlck;
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
	void OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void(*pFGrd)(CDatagrid1&,_RecordsetPtr&,_RecordsetPtr& ),int EmpCol=0,int DefCol=1);

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

	afx_msg void OnBnClickedCheck1();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();

	afx_msg void OnEnChangeEdit2();
	afx_msg void OnEnChangeEdit3();
	afx_msg void OnEnChangeEdit4();
	afx_msg void OnBnClickedCheck2();
	afx_msg void OnBnClickedCheck3();
	afx_msg void OnBnClickedCheck4();
	afx_msg void OnBnClickedCheck5();
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnBnClickedRadio1();
	afx_msg void OnBnClickedRadio2();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
public:
	CDatagrid1 m_DataGrid1;
	CDatagrid1 m_DataGrid2;
	CDatagrid1 m_DataGrid3;
	CDatagrid1 m_DataGrid4;
	CDatacombo2 m_DataCombo2;
	CDatacombo2 m_DataCombo1;
protected:
	static void InitDataGrid1(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	static void InitDataGrid1Db(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	static void InitDataGrid1Qr(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	static void InitDataGrid2(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	static void InitDataGrid3(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	static void InitDataGrid4(CDatagrid1& pG,_RecordsetPtr& rs,_RecordsetPtr& rsG);
	CBitmapButton m_OP,m_CS,m_SL,m_SR;
	BOOL LoadFont(int h);
	void OnShowBalance(CString strDate);
	void OnCalcExp(long& lTtlExp, long& lBefExp, long& lCurExp, long& lOwnExp, long& lOthExp, int& iOthNum, long& lWthExp, CString& strDate);
	void OnCalcInc(long& lSldAdd, CString& strDate);
	void OnShowCstClc(int i);
	void OnBlckUnBlck(BOOL& bBlck);
	void OnShowSlider(void);
	void OnShowCst(void);
	void OnEnableButton(int iCheck=-1, BOOL m_Check=FALSE);
	void OnEnableCheck(int Idc, BOOL bCheck=FALSE);
public:
	DECLARE_EVENTSINK_MAP()
public:
	void RowColChangeDatagrid1(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid2(VARIANT* LastRow, short LastCol);
	void RowColChangeDatagrid3(VARIANT* LastRow, short LastCol);
	void ButtonClickDatagrid4(short ColIndex);
	void ButtonClickDatagrid1(short ColIndex);
	void RowColChangeDatagrid4(VARIANT* LastRow, short LastCol);
public:
	CString m_Edit1;
	CString m_Edit6;
	CString m_Edit2;
	CString m_Edit3;
	CString m_Edit4;
	CString m_Edit5;
	CString m_strSqlSum;
	BOOL m_Check1;
	COleDateTime m_OleDateTime1;
	COleDateTime m_OleDateTime2;
	int m_Radio1;
	int m_Radio3;
	BOOL m_Check2;
	BOOL m_Check3;
	BOOL m_Check4;
	BOOL m_Check5;
	int m_Sldr2;
	int m_iSldr2;
public:
	void ClickDatagrid1();
	void ClickDatagrid2();
	void ClickDatagrid3();
	void ClickDatagrid4();
protected:
	void OnEnterVidBox(void);
	void OnEnterE_All(void);
	void OnEnterE2_10(void);
	void OnEnterE2_00(void);
	void OnEnterE4_00(void);
	void OnEnterE3_10(void);
	void OnEnterE3_00(void);
public:
	
	CString m_strQr;		// Общ запрос
	CString m_strBskt;		// Корзина
	CString m_strDb;		// база
	COleDateTime m_DtSrv;
	CDateTimeCtrl m_DateTimeCtrl1;
	CString m_strSysDate;
	COleDateTime m_OleDateTimeMIN;
	int m_Radio5;
	long m_CdOp;
	short m_shLstG1; // что было в гриде до выхода в корзину
	CString m_strNumBskt;
	long m_lCdOp;
protected:
	void OnSetPtr(void);
	void OnSetNumber(void);
	CString GetSourcePtr(_RecordsetPtr rs);
public:
	void ChangeDatacombo2();
public:
	afx_msg LRESULT OnSwitchFocus(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnAddRec(WPARAM wParam, LPARAM lParam);
};
