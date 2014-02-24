// Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg.h"

#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"

//#define WM_SWITCHFOCUS (WM_USER+20)
/*#include <stdlib.h>*/

// CDlg dialog

IMPLEMENT_DYNAMIC(CDlg, CDialog)

CDlg::CDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDlg::IDD, pParent)
	, m_CurCol(0)
	, m_befCurCol(0)
	, m_iBtSt(0)
	, m_Flg(true)
	, m_strNT(_T(""))
	, m_bFnd(false)
	, m_strFndC(_T(""))
	, m_bFndC(FALSE)
	, m_pLstWnd(NULL)

	, m_Edit1(_T("0.0000"))
	, m_Edit6(_T("0.0000"))
	, m_Edit3(_T(""))
	, m_Edit2(_T(""))
	, m_Edit5(_T("0"))
	, m_Check1(FALSE)
	, m_OleDateTime1(COleDateTime::GetCurrentTime())
	, m_OleDateTime2(COleDateTime::GetCurrentTime())
	, m_Radio1(0)
	, m_Radio3(0)
	, m_Check2(FALSE)
	, m_Check3(FALSE)
	, m_Check4(FALSE)
	, m_Check5(FALSE)
	, m_Sldr2(0)
	, m_iSldr2(0)
	, m_bBlck(FALSE)
	, m_strQr(_T(""))
	, m_strBskt(_T(""))
	, m_strDb(_T(""))

	, m_Edit4(_T(""))
	, m_DtSrv(COleDateTime::GetCurrentTime())
	, m_strSysDate(_T(""))
	, m_Radio5(1)
	, m_CdOp(0)
	, m_shLstG1(-1)	// -1 - Нет ещё никакой выборки 0 - Запрос, 1 - нак. базы детали
	, m_strNumBskt(_T("0"))
	, m_lCdOp(0)
{
	m_bCalcBskt = FALSE;	//перерасчёт не делать, TRUE делать
	m_bCalcDb   = FALSE;	//перерасчёт не делать, TRUE делать
	m_curColor = RGB(0,0,0);
    m_OleDateTimeMIN.SetDate(2005,1,1);
}

CDlg::~CDlg()
{

	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2 = NULL;

	if(ptrRs3->State==adStateOpen) ptrRs3->Close();
	ptrRs3 = NULL;

	if(ptrRs4->State==adStateOpen) ptrRs4->Close();
	ptrRs4 = NULL;

	if(ptrRs5->State==adStateOpen) ptrRs5->Close();
	ptrRs5 = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
	ptrRsCmb2 = NULL;

	if(ptrRsCst->State==adStateOpen) ptrRsCst->Close();
	ptrRsCst = NULL;

	if(ptrRsExp->State==adStateOpen) ptrRsExp->Close();
	ptrRsExp = NULL;

	if(ptrRsInc->State==adStateOpen) ptrRsInc->Close();
	ptrRsInc = NULL;

	if(ptrRs100->State==adStateOpen) ptrRs100->Close();
	ptrRs100 = NULL;

	if(ptrRsNum->State==adStateOpen) ptrRsNum->Close();
	ptrRsNum = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATAGRID4, m_DataGrid4);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT6, m_Edit6);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_Text(pDX, IDC_EDIT4, m_Edit4);
	DDX_Text(pDX, IDC_EDIT5, m_Edit5);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_OleDateTime2);
	DDX_Radio(pDX, IDC_RADIO1, m_Radio1);
	DDX_Radio(pDX, IDC_RADIO3, m_Radio3);
	DDX_Check(pDX, IDC_CHECK2, m_Check2);
	DDX_Check(pDX, IDC_CHECK3, m_Check3);
	DDX_Check(pDX, IDC_CHECK4, m_Check4);
	DDX_Check(pDX, IDC_CHECK5, m_Check5);
	//	DDX_Slider(pDX, IDC_SLIDER2, m_Sldr2);
	if(pDX->m_bSaveAndValidate){
		CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);
		m_Sldr2 = pSldr->GetPos();
	}

	DDX_Control(pDX, IDC_DATETIMEPICKER1, m_DateTimeCtrl1);
	DDX_Radio(pDX, IDC_RADIO5, m_Radio5);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_COMMAND(ID_BUTTON32771, On32771)
	ON_COMMAND(ID_BUTTON32772, On32772)
	ON_COMMAND(ID_BUTTON32773, On32773)
	ON_COMMAND(ID_BUTTON32774, On32774)
	ON_COMMAND(ID_BUTTON32775, On32775)
	ON_COMMAND(ID_BUTTON32776, On32776)
	ON_COMMAND(ID_BUTTON32777, On32777)
	ON_COMMAND(ID_BUTTON32778, On32778)
	ON_COMMAND(ID_BUTTON32779, On32779)
	ON_COMMAND(ID_BUTTON32780, On32780)
	ON_COMMAND(ID_BUTTON32783, On32783)

	ON_EN_CHANGE(ID_EDIT_TOOLBAR, OnEnChangeEditTB)

	ON_BN_CLICKED(IDOK, &CDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_OK, &CDlg::OnClickedOk)
	ON_BN_CLICKED(IDC_CANCEL, &CDlg::OnClickedCancel)
	ON_WM_CTLCOLOR()
	ON_EN_CHANGE(IDC_EDIT2, &CDlg::OnEnChangeEdit2)
	ON_EN_CHANGE(IDC_EDIT3, &CDlg::OnEnChangeEdit3)
	ON_EN_CHANGE(IDC_EDIT4, &CDlg::OnEnChangeEdit4)
	ON_BN_CLICKED(IDC_CHECK2, &CDlg::OnBnClickedCheck2)
	ON_BN_CLICKED(IDC_CHECK3, &CDlg::OnBnClickedCheck3)
	ON_BN_CLICKED(IDC_CHECK4, &CDlg::OnBnClickedCheck4)
	ON_BN_CLICKED(IDC_CHECK5, &CDlg::OnBnClickedCheck5)
	ON_BN_CLICKED(IDC_BUTTON1, &CDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &CDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CDlg::OnBnClickedButton4)
	ON_WM_HSCROLL()
	ON_BN_CLICKED(IDC_RADIO1, &CDlg::OnBnClickedRadio1)
	ON_BN_CLICKED(IDC_RADIO2, &CDlg::OnBnClickedRadio2)
	ON_BN_CLICKED(IDC_CHECK1, &CDlg::OnBnClickedCheck1)
	ON_BN_CLICKED(IDC_BUTTON5, &CDlg::OnBnClickedButton5)

	ON_MESSAGE(WM_SWITCHFOCUS,OnSwitchFocus)
	ON_MESSAGE(WM_ADDREC,OnAddRec)

END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"RmvStrg.dll"));

	CString s,strSql,strSql2,strC;
	s.LoadString(IDS_STRING9014);
	s+= _T(" (") + m_strNT;
	s+= _T(")");
	this->SetWindowText(s);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);

	m_OP.AutoLoad(IDC_BUTTON1,this);
	m_CS.AutoLoad(IDC_BUTTON2,this);
	m_SL.AutoLoad(IDC_BUTTON3,this);
	m_SR.AutoLoad(IDC_BUTTON4,this);

	m_Edit2T.SubclassDlgItem(IDC_EDIT2,this);
	m_Edit3T.SubclassDlgItem(IDC_EDIT3,this);
	m_Edit4T.SubclassDlgItem(IDC_EDIT4,this);

	InitStaticText();

/*	s.Format(_T("Dlg this %i %Xh 0x%x"),(UINT)((CWnd*)this),(UINT)((CWnd*)this),(UINT)((CWnd*)this));
AfxMessageBox(s);
*/
	int iTBCtrlID;
	short i;
	COleVariant vC;

	m_wndToolBar.Create(this, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_TOOLTIPS |
	                        CBRS_FLYBY | CBRS_BORDER_BOTTOM);
	m_wndToolBar.LoadToolBar(IDR_TOOLBAR1);

/*	s.Format(_T("Owner m_wndToolBar %i %Xh 0x%x"),(CWnd*)m_wndToolBar.GetParentOwner(),(CWnd*)m_wndToolBar.GetParentOwner(),(CWnd*)m_wndToolBar.GetParentOwner());
AfxMessageBox(s);
s.Format(_T("m_wndToolBar Hwnd %X"),m_wndToolBar.GetSafeHwnd()); 
AfxMessageBox(s);

s.Format(_T("this of ToolBar %i %Xh 0x%x"),(CWnd*)m_wndToolBar.GetSafeOwner(),(CWnd*)m_wndToolBar.GetSafeOwner(),(CWnd*)m_wndToolBar.GetSafeOwner());
AfxMessageBox(s);
*/
	iTBCtrlID = m_wndToolBar.CommandToIndex(ID_BUTTON32779);
	if(iTBCtrlID>=0){
		for(int j=iTBCtrlID;j<(iTBCtrlID + 6);j++){
			switch(j){
				case 0:
					m_wndToolBar.SetButtonStyle(j,TBBS_CHECKBOX);
					m_iBtSt =m_wndToolBar.GetToolBarCtrl().GetState(ID_BUTTON32779);
					break;
/*				case 1:
				case 2:
				case 3:
				case 4:
					break;
*/
				default:
					m_wndToolBar.SetButtonStyle(j,TBBS_BUTTON);
					break;
			}
		}
	}
//	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	m_wndToolBar.SetButtonInfo(5,ID_EDIT_TOOLBAR,TBBS_SEPARATOR,130);
	// определяем координаты области, занимаемой разделителем
	CRect rEdit; 
	m_wndToolBar.GetItemRect(5,&rEdit);
	rEdit.left+=6; rEdit.right-=6; // отступы
	if(!(m_EditTBCh.Create(WS_CHILD | ES_AUTOHSCROLL | ES_MULTILINE | WS_VISIBLE | 
						   WS_TABSTOP |	WS_BORDER, rEdit, &m_wndToolBar,ID_EDIT_TOOLBAR))) return -1;
	
/*s.Format(_T("Owner m_EditTBCh %i %Xh 0x%x"),(CWnd*)m_EditTBCh.GetParentOwner(),(CWnd*)m_EditTBCh.GetParentOwner(),(CWnd*)m_EditTBCh.GetParentOwner());
AfxMessageBox(s);

s.Format(_T("this of m_EditTBCh %i %Xh 0x%x"),&m_EditTBCh,&m_EditTBCh,&m_EditTBCh);
AfxMessageBox(s);
s.Format(_T("m_EditTBCh Hwnd %X"),(&m_EditTBCh)->GetSafeHwnd()); 
AfxMessageBox(s);
*/
	CRect rcClientStart;
	CRect rcClientNow;
	GetClientRect(rcClientStart);
	RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,
				   IDR_TOOLBAR1, reposQuery , rcClientNow);
	CPoint ptOffset(rcClientNow.left - rcClientStart.left,
										rcClientNow.top - rcClientStart.top);

/*	CRect rcChild;
	CWnd* pwndChild = GetWindow(GW_CHILD);
//	Это для смещения всех контрлов
	while (pwndChild)
	{
		pwndChild->GetWindowRect(rcChild);
		ScreenToClient(rcChild);
		rcChild.OffsetRect(ptOffset);
		pwndChild->MoveWindow(rcChild, FALSE);
		pwndChild = pwndChild->GetNextWindow();
	}

	CRect rcWindow;
	GetWindowRect(rcWindow);
	rcWindow.right += rcClientStart.Width() - rcClientNow.Width();
	rcWindow.bottom += rcClientStart.Height() - rcClientNow.Height();
	MoveWindow(rcWindow, FALSE);	
*/
	// Положение панелей
	RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,
				 IDR_TOOLBAR1, reposDefault , rcClientNow);

/*	if(m_EditTB.GetSafeHwnd()!=NULL){
		AfxMessageBox("not NULL");
		if(m_EditTBSChar.SubclassWindow(m_EditTB.GetSafeHwnd())){
			AfxMessageBox("ok");
		}
	}
*/
	m_vNULL.vt = VT_ERROR;
	m_vNULL.scode = DISP_E_PARAMNOTFOUND;

	ptrCmdDtSrv = NULL;
	ptrRsDtSrv = NULL;

	strSql = _T("GetDateServer");
	ptrCmdDtSrv.CreateInstance(__uuidof(Command));
	ptrCmdDtSrv->ActiveConnection = ptrCnn;
	ptrCmdDtSrv->CommandType = adCmdText;
	ptrCmdDtSrv->CommandText = (_bstr_t)strSql;

	ptrRsDtSrv.CreateInstance(__uuidof(Recordset));
	ptrRsDtSrv->CursorLocation = adUseClient;
	ptrRsDtSrv->PutRefSource(ptrCmdDtSrv);
	try{
		ptrRsDtSrv->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		i = 0;
		vC = GetValueRec(ptrRsDtSrv,i);
		vC.ChangeType(VT_DATE);
		m_DtSrv = vC.date;
		m_strSysDate = m_DtSrv.Format(_T("%Y-%m-%d"));
		m_OleDateTime1.ParseDateTime(m_strSysDate);
		m_DateTimeCtrl1.SetRange(&m_OleDateTimeMIN,&m_DtSrv);

	}
	catch(_com_error& e){
		AfxMessageBox(IDS_STRING9049,MB_ICONSTOP);
		//AfxMessageBox(e.Description());
		return TRUE;
	}

	ptrCmdNum = NULL;
	ptrRsNum = NULL;

	ptrCmdNum.CreateInstance(__uuidof(Command));
	ptrCmdNum->ActiveConnection = ptrCnn;
	ptrCmdNum->CommandType = adCmdText;

	ptrRsNum.CreateInstance(__uuidof(Recordset));
	ptrRsNum->CursorLocation = adUseClient;
	ptrRsNum->PutRefSource(ptrCmdNum);

	ptrCmd100 = NULL;
	ptrRs100 = NULL;

	ptrCmd100.CreateInstance(__uuidof(Command));
	ptrCmd100->ActiveConnection = ptrCnn;
	ptrCmd100->CommandType = adCmdText;

	ptrRs100.CreateInstance(__uuidof(Recordset));
	ptrRs100->CursorLocation = adUseClient;
	ptrRs100->PutRefSource(ptrCmd100);

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);

	strSql = _T("QT60NSumNum '");
	strSql+= m_strNT + _T("'");
	m_strBskt = strSql;

	m_strSqlSum= m_strBskt;
	m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
//AfxMessageBox(m_strSqlSum);
	ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
	if(ptrRs100->State==adStateOpen) ptrRs100->Close();
	try{
		ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description());
	}

	OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
/*	if(IsEnableRec(ptrRs1)){
		i = 0; // № нак. минимальный для тек. user в корзине
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		m_Edit5 = vC.bstrVal;
		m_Edit5.TrimLeft();
		m_Edit5.TrimRight();
	}
	else{
		m_Edit5 = _T("0");
	}
*/
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

	m_CurCol = m_DataGrid1.get_Col();
	if(m_CurCol==-1 || m_CurCol==0){
		m_CurCol=1;
	}
	m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
	m_EditTBCh.SetTypeCol(m_iCurType);

	if(m_bFndC){
		short fCol=0;
		m_Flg = false;
		OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
		m_Flg = true;
	}

	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);


	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			strSql = _T("T12_T27_1 ");
			strSql+= s + _T(",0");
			break;
		case 1:
			if(IsEnableRec(ptrRs3)){
				i = 0;
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimLeft();
				strC.TrimRight();
				strSql = _T("T12_T27_1 ");
				strSql+= s + _T(",");
				strSql+= strC;
			}
			break;
	}
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			if(ptrRs2->State==adStateOpen){
				i = 0;
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimLeft();
				strC.TrimRight();

				strSql = _T("T27_T26 ");
				strSql+= s	+ _T(",");
				strSql+= strC;
			}
			break;
		case 1:
				strSql = _T("T27_T26 ");
				strSql+= s	+_T(",0");
			break;
	}
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);

	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);
	strSql = _T("QT10");
	OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);

	ptrCmd5 = NULL;
	ptrRs5 = NULL;

	ptrCmd5.CreateInstance(__uuidof(Command));
	ptrCmd5->ActiveConnection = ptrCnn;
	ptrCmd5->CommandType = adCmdText;

	ptrRs5.CreateInstance(__uuidof(Recordset));
	ptrRs5->CursorLocation = adUseClient;
	ptrRs5->PutRefSource(ptrCmd5);
	OnShowSlider();

	ptrCmdCmb1 = NULL;
	ptrRsCmb1 = NULL;

	strSql = _T("QT51Combo");
	ptrCmdCmb1.CreateInstance(__uuidof(Command));
	ptrCmdCmb1->ActiveConnection = ptrCnn;
	ptrCmdCmb1->CommandType = adCmdText;
	ptrCmdCmb1->CommandText = (_bstr_t)strSql;

	ptrRsCmb1.CreateInstance(__uuidof(Recordset));
	ptrRsCmb1->CursorLocation = adUseClient;
	ptrRsCmb1->PutRefSource(ptrCmdCmb1);
	try{
		m_DataCombo1.putref_RowSource(NULL);
		ptrRsCmb1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo1.putref_RowSource((LPUNKNOWN) ptrRsCmb1);
		m_DataCombo1.put_ListField(_T("Наименование"));
		i = 1;
/*		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		s = vC.bstrVal;
		s.TrimLeft();
		s.TrimRight();
		m_DataCombo1.put_Text(s);
*/
		SetTextCombo(m_DataCombo1,ptrRsCmb1,i);
//AfxMessageBox("ok");
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb2 = NULL;
	ptrRsCmb2  = NULL;

	strSql = _T("QT23");
	ptrCmdCmb2.CreateInstance(__uuidof(Command));
	ptrCmdCmb2->ActiveConnection = ptrCnn;
	ptrCmdCmb2->CommandType = adCmdText;
	ptrCmdCmb2->CommandText = (_bstr_t)strSql;

	ptrRsCmb2.CreateInstance(__uuidof(Recordset));
	ptrRsCmb2->CursorLocation = adUseClient;
	ptrRsCmb2->PutRefSource(ptrCmdCmb2);
	try{
		ptrRsCmb2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo2.putref_RowSource((LPUNKNOWN) ptrRsCmb2);
		m_DataCombo2.put_ListField(_T("Наименование"));
	}
	catch(_com_error& e){
		m_DataCombo2.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCst = NULL;
	ptrRsCst = NULL;

	ptrCmdCst.CreateInstance(__uuidof(Command));
	ptrCmdCst->ActiveConnection = ptrCnn;
	ptrCmdCst->CommandType = adCmdText;

	ptrRsCst.CreateInstance(__uuidof(Recordset));
	ptrRsCst->CursorLocation = adUseClient;
	ptrRsCst->PutRefSource(ptrCmdCst);

	ptrCmdExp = NULL;
	ptrRsExp = NULL;

	ptrCmdExp.CreateInstance(__uuidof(Command));
	ptrCmdExp->ActiveConnection = ptrCnn;
	ptrCmdExp->CommandType = adCmdText;

	ptrRsExp.CreateInstance(__uuidof(Recordset));
	ptrRsExp->CursorLocation = adUseClient;
	ptrRsExp->PutRefSource(ptrCmdExp);

	ptrCmdInc = NULL;
	ptrRsInc = NULL;

	ptrCmdInc.CreateInstance(__uuidof(Command));
	ptrCmdInc->ActiveConnection = ptrCnn;
	ptrCmdInc->CommandType = adCmdText;

	ptrRsInc.CreateInstance(__uuidof(Recordset));
	ptrRsInc->CursorLocation = adUseClient;
	ptrRsInc->PutRefSource(ptrCmdInc);

	OnShowCst();
	OnShowCstClc(m_Sldr2);
	OnEnableButton();
	OnShowBalance(m_strSysDate);
	if(IsEnableRec(ptrRs1)){
		m_bBlck = TRUE;
		OnBlckUnBlck(m_bBlck);
	}

	OnSetPtr();
	OnSetNumber();
	if(LoadFont(-12)) {
//AfxMessageBox("ok");
		GetDlgItem(IDC_STATIC_SLIDER2)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_INC)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_EXP)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_RSR)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_RST)->SetFont(&m_curFont,FALSE);

		GetDlgItem(IDC_STATIC_INCS)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_EXPS)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_RSRS)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_RSTS)->SetFont(&m_curFont,FALSE);
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
void CDlg::OnEnableButtonBar(int iBtSt, CToolBar* pTlBr)
{	
	CString strSrc;
	strSrc = (BSTR)ptrCmd1->GetCommandText();
//AfxMessageBox(strSrc);
	if(strSrc.Find(_T("QT60NSumNum"))!=-1){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32774,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32780,TRUE);
		if(iBtSt==4){
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,TRUE);

		}
		else{
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,TRUE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);

		}
		if(!IsEmptyRec(ptrRs1)){
			m_bCalcBskt = TRUE;
		}
		else{
			m_bCalcBskt = FALSE;
		}
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,m_bCalcBskt);	
	}
	else if(strSrc.Find(_T("QT100N"))!=-1){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32774,FALSE);	
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32780,FALSE);

		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,TRUE);

	}
	else{
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,TRUE);
		if(iBtSt==4){
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32774,TRUE);//FALSE
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32780,FALSE);

		}
		else{	//Корректировка
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,TRUE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,TRUE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32780,FALSE);
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);
		}
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,m_bCalcDb);	
	}
}

void CDlg::OnEnChangeEditTB(void)
{
//	UpdateData();
	CString s,sLstCh,str,msg,strCopy;
	BOOL bDec;
	bDec = FALSE;
	if(m_befCurCol!=0){
		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
			m_CurCol  = m_befCurCol;
			m_iCurType= m_befCurType;
			m_DataGrid1.put_Col(m_CurCol);

			m_befCurCol=m_befCurType = 0;
		}
	}
//	if(ptrRs1->State==adStateOpen){
		m_EditTBCh.GetWindowText(str);
//AfxMessageBox(_T("OnEnChangeEditTB str = ")+str);
		switch(m_iCurType){
			case adSmallInt:			//2
			case adInteger:				//3
			case adBoolean:				//11
			case adTinyInt:				//16
			case adUnsignedTinyInt:		//17
			case adUnsignedSmallInt:	//18
			case adUnsignedInt:			//19
				str.Remove('.');
				str.Remove(',');
//AfxMessageBox(_T("2-19"));
				break;
			case adDecimal:		//14
			case adNumeric:	//131
			case adDouble:
			case adCurrency:
//AfxMessageBox(_T("14-131"));
				for (int i=0; i<str.GetLength(); i++){
					if( str.GetAt(i)==_T('.') || 
						str.GetAt(i)==_T(',')
					  )
					{  
						if(!bDec){
							strCopy+=str.GetAt(i);
							bDec = TRUE;
						}
						else{
							 continue;
						}

					}
					else{
						strCopy+=str.GetAt(i);
					}
				}
				strCopy.TrimLeft();
				strCopy.TrimRight();
				if(strCopy.GetLength()==1 &&
					(strCopy.GetAt(0)==_T('.') || strCopy.GetAt(0)==_T(','))	){
						strCopy.Empty();
					
				}
				str = strCopy;
//AfxMessageBox(_T("OnEnChangeEditTB After for str = ")+str);
				break;
		}
//	}
	if(IsEnableRec(ptrRs1)){
		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){

			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}
	if(IsEnableRec(ptrRs2)){
		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			if(!str.IsEmpty()){
//AfxMessageBox(str);
/*s.Format(_T("m_iCurType = %i m_CurCol = %i"),m_iCurType,m_CurCol);
AfxMessageBox(s);
*/
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}
	if(IsEnableRec(ptrRs3)){
		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			if(!str.IsEmpty()){
//AfxMessageBox(str);
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs3,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}
	if(IsEnableRec(ptrRs4)){
		if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs4,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount,s;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
//s.Format(_T("%i"),Id);
//AfxMessageBox(s);
		switch(Id){
			case IDC_EDIT1:
				GetDlgItem(IDC_EDIT1)->GetWindowTextW(m_Edit1);
				m_Edit1.Replace(',','.');
				m_Edit1.TrimRight();
				m_Edit1.TrimLeft();
				if(m_Edit1==' ') m_Edit1.Empty();
				if(m_Edit1.IsEmpty() || _wtof(m_Edit1)==0){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT2:
				GetDlgItem(IDC_EDIT2)->GetWindowTextW(m_Edit2);
				m_Edit2.Replace(',','.');
				m_Edit2.TrimRight();
				m_Edit2.TrimLeft();
				if(m_Edit2==' ') m_Edit2.Empty();
				if(m_Edit2.IsEmpty() || _wtof(m_Edit2)==0){
					count++;
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT6:
				GetDlgItem(IDC_EDIT6)->GetWindowTextW(m_Edit6);
				m_Edit6.Replace(',','.');
				m_Edit6.TrimRight();
				m_Edit6.TrimLeft();
				if(m_Edit6==' ') m_Edit6.Empty();
				if(m_Edit6.IsEmpty() || _wtof(m_Edit6)==0){
					count++;
					GetDlgItem(IDC_STATIC_EDIT6)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
		}
		NextDlgCtrl();
	} while (Id!=IdEnd);

/*
//Раздел COMBO
	strC = m_DataCombo1.get_BoundText();

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO1)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}
*/
/*	//Раздел Grid
	if(ptrRs2->State==adStateOpen){
		if(IsEmptyRec(ptrRs2)){
			count++;
			strItem.LoadString(IDS_STRING9017);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{
		count++;
		strItem.LoadString(IDS_STRING9017);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}
*/
	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	short i;
	CString strC;
	CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);
	double dSldr;
	CString s;
	if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				try{
					i = 6;	//Шт.
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					m_Edit2 = vC.bstrVal;
					m_Edit2.TrimLeft();
					m_Edit2.TrimRight();
					GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
	
					switch(m_Radio1){
						case 0:	//по товару
							i = 16;	//Код товара
							vC = GetValueRec(ptrRs1,i);
							vC.ChangeType(VT_BSTR);
							strC = vC.bstrVal;
							strC.TrimLeft();
							strC.TrimRight();
							m_Flg = false;
							OnFindInGrid(strC,ptrRs2,0,m_Flg);
							m_Flg = true;
							break;
						case 1:	//по группе
							i = 17;	//Код группы
							vC = GetValueRec(ptrRs1,i);
							vC.ChangeType(VT_BSTR);
							strC = vC.bstrVal;
							strC.TrimLeft();
							strC.TrimRight();
							m_Flg = false;
							OnFindInGrid(strC,ptrRs3,0,m_Flg);
							m_Flg = true;
							//затем ищем товар в группе
							i = 16;	//Код товара
							vC = GetValueRec(ptrRs1,i);
							vC.ChangeType(VT_BSTR);
							strC = vC.bstrVal;
							strC.TrimLeft();
							strC.TrimRight();
							m_Flg = false;
							OnFindInGrid(strC,ptrRs2,0,m_Flg);
							m_Flg = true;
							break;
					}

					i = 18;	//% лин
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_R8);
					dSldr = vC.dblVal;
//s.Format(_T("%8.2f"),dSldr);
//AfxMessageBox(s);

					for(int p=0;p<aSldr.GetSize();p++){
						if(dSldr==aSldr.GetAt(p)){
							CString s;
							pSldr->SetPos(p);
							OnShowCstClc(pSldr->GetPos());
							s.Format(_T("%4.2f"),aSldr.GetAt(pSldr->GetPos()));
							s.TrimLeft();
							s.TrimRight();
							s.Replace(',','.');
					
							if(aSldr.GetAt(pSldr->GetPos())<0){
								m_curColor = RGB(255,0,0);
							}

							else{
								m_curColor = RGB(0,0,128);
							}
							SetDlgItemText(IDC_STATIC_SLIDER2,s);
							break;
						}
					}

					i = 22;	//Цена баз
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					m_Edit1 = vC.bstrVal;
					m_Edit1.TrimLeft();
					m_Edit1.TrimRight();
					m_Edit1.Replace(',','.');
					GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

					i = 3;	//Цена без скидки
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					m_Edit6 = vC.bstrVal;
					m_Edit6.TrimLeft();
					m_Edit6.TrimRight();
					m_Edit6.Replace(',','.');
					GetDlgItem(IDC_EDIT6)->SetWindowText(m_Edit6);

				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
			}
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_GROUP1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_RADIO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_RADIO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9026);
	GetDlgItem(IDC_STATIC_EDIT4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9027);
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9029);
	GetDlgItem(IDC_STATIC_DATETIME2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9030);
	GetDlgItem(IDC_STATIC_GROUP2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9031);
	GetDlgItem(IDC_RADIO3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9032);
	GetDlgItem(IDC_RADIO4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9050);
	GetDlgItem(IDC_RADIO7)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9033);
	SetDlgItemText(IDC_STATIC_GROUP3,sTxt);

	sTxt.LoadString(IDS_STRING9034);
	SetDlgItemText(IDC_CHECK2,sTxt);

	sTxt.LoadString(IDS_STRING9035);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9037);
	SetDlgItemText(IDC_CHECK3,sTxt);

	sTxt.LoadString(IDS_STRING9038);
	SetDlgItemText(IDC_CHECK4,sTxt);

//	sTxt.LoadString(IDS_STRING9039);
//	SetDlgItemText(IDC_BUTTON1,sTxt);
//	SetDlgItemText(IDC_BUTTON2,sTxt);
//	SetDlgItemText(IDC_BUTTON3,sTxt);
//	SetDlgItemText(IDC_BUTTON4,sTxt);

	sTxt.LoadString(IDS_STRING9040);
	SetDlgItemText(IDC_STATIC_EDIT6,sTxt);

	sTxt.LoadString(IDS_STRING9041);
	SetDlgItemText(IDC_STATIC_INC,sTxt);

	sTxt.LoadString(IDS_STRING9042);
	SetDlgItemText(IDC_STATIC_EXP,sTxt);

	sTxt.LoadString(IDS_STRING9043);
	SetDlgItemText(IDC_STATIC_RSR,sTxt);

	sTxt.LoadString(IDS_STRING9044);
	SetDlgItemText(IDC_STATIC_RST,sTxt);

	sTxt.LoadString(IDS_STRING9045);
	SetDlgItemText(IDC_CHECK5,sTxt);

	sTxt.LoadString(IDS_STRING9046);
	SetDlgItemText(IDC_STATIC_GROUP4,sTxt);

	sTxt.LoadString(IDS_STRING9047);
	SetDlgItemText(IDC_RADIO5,sTxt);

	sTxt.LoadString(IDS_STRING9048);
	SetDlgItemText(IDC_RADIO6,sTxt);
}

void CDlg::On32771(void)
{
//Возврат в Корзину
	COleVariant vC;
	short i;
	if(ptrRs100->State==adStateOpen) ptrRs100->Close();
	if(IsQuery(_T("QT100N"),ptrRs1)){
		if(!ptrRs1->adoEOF) m_vBkQr = ptrRs1->GetBookmark();
		m_shLstG1 = 0;
	}
	if(IsQuery(_T("QT24NSumNum"),ptrRs1)){
		if(!ptrRs1->adoEOF) m_vBkDb = ptrRs1->GetBookmark();
		m_shLstG1 = 1;
	}
	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);

	m_strSqlSum= m_strBskt;
	m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
//AfxMessageBox(m_strSqlSum);
	ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
	if(ptrRs100->State==adStateOpen) ptrRs100->Close();
	try{
		ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description());
	}
	OnShowGrid(m_strBskt,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	if(IsEnableRec(ptrRs1)){
		try{
			ptrRs1->put_Bookmark(m_vBkBskt);
		}
		catch(...){
			ptrRs1->MoveLast();
		}
	}

	if(IsEnableRec(ptrRs1)){
		i = 0; //№ нак
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		m_Edit5 = vC.bstrVal;
		m_Edit5.TrimLeft();
		m_Edit5.TrimRight();
	}
	else{
		if (_wtoi(m_strBskt)==0){
			m_Edit5 = _T("0");
		}
		else{
			m_Edit5 = m_strNumBskt;
		}
	}
	GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	OnSetPtr();
	OnBlckUnBlck(m_bBlck);
	OnEnableButton();

	return;
}

void CDlg::On32772(void)
{
//Занести в  базу нак из корзины
	CString strSqlTmp,strSql;

	if(IsEnableRec(ptrRs1)){
		if(IsQuery(_T("QT60NSumNum "),ptrRs1)){
			strSqlTmp=GetSourcePtr(ptrRs1);
			if(m_bCalcBskt){

				BOOL bFndC;
				CString strFndC;

				short i;
				COleVariant vC;
				COleDateTime vD;
				CString strD,strDO,strC,strCst,strSH;
				CString strSls,strAcc;

				_variant_t vra;
				VARIANT* vtl = NULL;

				i = 0;	//№ нак
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimRight();
				strC.TrimRight();

				i = 1;	//Дата выписки нак
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_DATE);
				vD = vC.date;
				strD = vD.Format(_T("%Y-%m-%d"));
				
				i = 8;	//Дата выд нак
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_DATE);
				vD = vC.date;
				strDO = vD.Format(_T("%Y-%m-%d"));

				i = 12;		//Код контрАг покупателя
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCst = vC.bstrVal;
				strCst.TrimLeft();
				strCst.TrimRight();

				i = 13;	//SH
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strSH = vC.bstrVal;
				strSH.TrimRight();
				strSH.TrimRight();

				i = 14;		//Код контрАг продовца
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strSls = vC.bstrVal;
				strSls.TrimLeft();
				strSls.TrimRight();

				strSql = _T("TrnsfRmv_2 ");
				strSql+= strC	+ _T(",'");
				strSql+= strD	+ _T("',");
				strSql+= strSH	+ _T(",");
				strSql+= strCst	+ _T(",'");
				strSql+= strDO	+ _T("','");
				strSql+= m_strNT+ _T("'");

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;

				try{
//AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
					m_bCalcBskt = FALSE;	//  делать пересчёт
					m_Edit2 = _T("0");
					SetDlgItemText(IDC_EDIT2,m_Edit2);
					m_Edit5 = m_strNumBskt = _T("0");
					SetDlgItemText(IDC_EDIT5,m_Edit5);
					OnSetNumber();
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
					return;
				}
				m_strSqlSum= m_strBskt;
				m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
	//AfxMessageBox(_T("m_strSqlSum = ")+m_strSqlSum);
				ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
				if(ptrRs100->State==adStateOpen) ptrRs100->Close();
				try{
					ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);

				OnEnableButton();
				OnShowBalance(m_strSysDate);
				if(IsEmptyRec(ptrRs1)){
					m_bBlck = FALSE;
					OnBlckUnBlck(m_bBlck);
				}

				OnSetPtr();

				HMODULE hMod;
				bFndC   = TRUE;
				strFndC = _T("");

				strSql = _T("QT38w1_1 ");
				strSql+= strCst + _T("~");
				strSql+= strSls;

				hMod=AfxLoadLibrary(_T("PrnVw.dll"));
				typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&,CString& );
				pDialog func=(pDialog)GetProcAddress(hMod,"startPrnVw");
				BeginWaitCursor();
					(func)(m_strNT, ptrCnn,strFndC,bFndC,strSql,strAcc);
				EndWaitCursor();
				AfxFreeLibrary(hMod);


			}
			else{
				AfxMessageBox(IDS_STRING9057,MB_ICONINFORMATION);
				return;
			}
		}
	}
	return;
}

void CDlg::On32773(void)
{
//Пересчёт нак 
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){


			_variant_t vra;
			COleVariant vC,vBk;
			COleDateTime vD;
			VARIANT* vtl = NULL;
			CString strC,strD,strCst;
			CString strSqlTmp,strSql;
			short i;

			vBk = ptrRs1->GetBookmark();

			i = 0;	//№ нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();

			i = 1;	//Дата выписки нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));
			
			i = 12;		//Код контрАг покупателя
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCst = vC.bstrVal;
			strCst.TrimLeft();
			strCst.TrimRight();

			strSqlTmp=GetSourcePtr(ptrRs1);

			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				switch(m_lCdOp){
					case 101:	//дистьебьютер
						strSql = _T("IT117_5Dst ");
						break;
					default:
						strSql = _T("IT117_5 ");
						break;
				}
			}
			else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
				strSql = _T("");
			}
			strSql+= strCst  + _T(",'");
			strSql+= strD	 + _T("',");
			strSql+= strC	 + _T(",'");
			strSql+= m_strNT + _T("'");

			ptrCmd1->CommandText = (_bstr_t)strSql;
			ptrCmd1->CommandType = adCmdText;
			BeginWaitCursor();
			try{
	//AfxMessageBox(strSql);
				ptrCmd1->Execute(&vra,vtl,adCmdText);

				if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
					m_bCalcBskt = TRUE;	//Делать пересчёт
				}
				else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
					m_bCalcDb = TRUE;	//Делать пересчёт
				}

			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			EndWaitCursor();
			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				m_strSqlSum= m_strBskt;
				m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
	//AfxMessageBox(_T("m_strSqlSum = ")+m_strSqlSum);
				ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
				if(ptrRs100->State==adStateOpen) ptrRs100->Close();
				try{
					ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
//AfxMessageBox(_T("1"));
			}
			else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
				m_strSqlSum= m_strDb;
				m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
	//AfxMessageBox(m_strSqlSum);
				ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
				if(ptrRs100->State==adStateOpen) ptrRs100->Close();
				try{
					ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowGrid(m_strDb, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
			}
			if(IsEnableRec(ptrRs1)){
				ptrRs1->put_Bookmark(vBk);
			}
		}
		OnShowBalance(m_strSysDate);
		OnBlckUnBlck(m_bBlck);
		OnEnableButtonBar(m_iBtSt,&m_wndToolBar);

	}
	return;
}

void CDlg::On32774(void)
{
//Возврат к запросу накладных
	CString m_strSqlSum;
	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);

	m_strSqlSum= m_strQr;
	m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
//AfxMessageBox(m_strSqlSum);
	ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
	if(ptrRs100->State==adStateOpen) ptrRs100->Close();
	try{
		ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description());
	}

	OnShowGrid(m_strQr,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Qr);
	if(IsEnableRec(ptrRs1)){
		try{
			ptrRs1->put_Bookmark(m_vBkQr);
		}
		catch(...){
			ptrRs1->MoveLast();
		}
/*		m_Edit5 = _T("0");
		GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
*/
		OnSetPtr();
	}
	OnBlckUnBlck(m_bBlck);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	OnEnableButton();
	return;
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql,strSqlTmp,strQws,strSlsN,strOp;
	CString strCT,strCG,strCS,strR,strSls,strCrV,strCst,strCstN;
	CString strSlsCst,sCstB,sCstP,strС3D,strC24D,sPrc,sSH,strC6T84;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;
	short i;
	int iB,iE;
	double dSldr;
	CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);	

	strCT = _T("0");	
	strCG = _T("0");
	strCS = _T("0");
	sPrc  = _T("0");
	strR  = m_Radio1==0 ? _T("0"):_T("1");
	sSH	  = m_Radio5==0 ? _T("0"):_T("1");
	strС3D  = m_OleDateTime1.Format(_T("%Y-%m-%d")); //Дата выписки
	strC24D = m_OleDateTime2.Format(_T("%Y-%m-%d")); //Дата выдачи	
	dSldr = 0.00;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}
	m_Edit5.TrimLeft();
	m_Edit5.TrimRight();
	if(m_Edit5.IsEmpty() || _wtoi(m_Edit5)==0){
		AfxMessageBox(IDS_STRING9051,MB_ICONINFORMATION);
		return;
	}
	if(OnOverEdit(IDC_EDIT1,IDC_EDIT2)){
		switch(m_Radio1){
			case 0:	//по Товару
				i = 0;
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strCT = vC.bstrVal;
				strCT.TrimLeft();
				strCT.TrimRight();
				break;
			case 1:	//по Группе
				i = 0;
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strCG = vC.bstrVal;
				strCG.TrimLeft();
				strCG.TrimRight();
				break;
		}

		i = 0;		//Код контрАг покупателя
		vC = GetValueRec(ptrRs4,i);
		vC.ChangeType(VT_BSTR);
		strCst = vC.bstrVal;
		strCst.TrimLeft();
		strCst.TrimRight();

		i = 2;		// Наименование покупателя
		vC = GetValueRec(ptrRs4,i);
		vC.ChangeType(VT_BSTR);
		strCstN = vC.bstrVal;
		strCstN.TrimLeft();
		strCstN.TrimRight();

		strQws = strCst + _T(" ");
		strQws+= strCstN;

		i = 2;		//Код контрАг продовца
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strSls = vC.bstrVal;
		strSls.TrimLeft();
		strSls.TrimRight();

		i = 1;		//Наименование контрАг продовца
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strSlsN = vC.bstrVal;
		strSlsN.TrimLeft();
		strSlsN.TrimRight();
		strSlsN+= _T("\n");

		i = 0;		//Код Вида курса
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strCrV = vC.bstrVal;

		i = 0;		//Баз. цена
		vC = GetValueRec(ptrRsCst,i);
		vC.ChangeType(VT_BSTR);
		sCstB = vC.bstrVal;
		sCstB.TrimLeft();
		sCstB.TrimRight();
		sCstB.Replace(',','.');

		i = 1;		//Прайс. цена
		vC = GetValueRec(ptrRsCst,i);
		vC.ChangeType(VT_BSTR);
		sCstP = vC.bstrVal;
		sCstP.TrimLeft();
		sCstP.TrimRight();
		sCstP.Replace(',','.');

		i = 0;		// Код операции
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strC6T84 = vC.bstrVal;
		strC6T84.TrimLeft();
		strC6T84.TrimRight();

		i = 1;		//Операция
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strOp = vC.bstrVal;
		strOp.TrimLeft();
		strOp.TrimRight();

		strSlsN+= _T("Операция:       ");
		strSlsN+=strOp;

		if(!aSldr.IsEmpty()){
			dSldr = aSldr.GetAt(pSldr->GetPos());
			sPrc.Format(_T("%5.2f"),dSldr);
		}
		else{
			dSldr = 0.00;
		}
		strSqlTmp=GetSourcePtr(ptrRs1);

		if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
			switch(m_Radio1){
				case 0:	//по товару
					switch(m_lCdOp){
						case 101:	// продажа дистьебютеру
							strSql = _T("IT60N_Dst ");	
							sPrc = _T("0");
							break;
						default:
							strSql = _T("IT60N ");
							break;
					}
					strSql+= strCst  + _T(",");		//1  Код покупателя
					strSql+= strCT	 + _T(",'");	//2  Код товара
					strSql+= strС3D  + _T("',");	//3  Дата выписки
					strSql+= m_Edit5 + _T(",");		//4  № нак
					strSql+= sCstP	 + _T(",");		//5  Цена прайс
					strSql+= sCstB   + _T(",");		//6  Цена баз
					strSql+= sPrc	 + _T(",");		//7  % лин цен
					strSql+= m_Edit6 + _T(",");		//8  Цена без скидки
					strSql+= m_Edit2 + _T(",");		//9  Шт.
					strSql+= sSH	 + _T(",");		//10 S/H
					strSql+= strSls  + _T(",'");	//11 Код продовца
					strSql+= strC24D + _T("',");	//12 Дата выдачи
					strSql+= strC6T84+ _T(",'");	//13 Код опер
					strSql+= m_strSysDate+_T("','");//14 Сис. дата
					strSql+= m_strNT + _T("'");

					break;
				case 1:	// по группе
					switch(m_lCdOp){
						case 101:	// продажа дистьебютеру
							strSql = _T("IT60NGr_Dst ");
							sPrc = _T("0");
							break;
						default:
							strSql = _T("IT60NGr ");
							break;

					}
					strSql+= strCst  + _T(",");		//1  Код покупателя
					strSql+= strCG	 + _T(",'");	//2  Код группы
					strSql+= strС3D  + _T("',");	//3  Дата выписки
					strSql+= m_Edit5 + _T(",");		//4  № нак
					strSql+= sPrc	 + _T(",");		//5  % лин цен
					strSql+= m_Edit6 + _T(",");		//6  Цена без скид (для дист)
					strSql+= m_Edit2 + _T(",");		//7  Шт.
					strSql+= sSH	 + _T(",");		//8  S/H
					strSql+= strSls  + _T(",'");	//9  Код продовца
					strSql+= strC24D + _T("',");	//10 Дата выдачи
					strSql+= strC6T84+ _T(",'");	//11 Код опер
					strSql+= m_strSysDate+_T("','");//12 Сис. дата
					strSql+= m_strNT + _T("'");
					break;
			}
		}
		else if(IsQuery(_T("QT24NSumNum"),ptrRs1)){
			strSql = _T("IT24N '");
			strSql+= m_Edit2 + _T("','");
			strSql+= m_Edit1 + _T("','");
			strSql+= m_strNT + _T("'");
		}
		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		if(IsEmptyRec(ptrRs1)){
			CString s;
			AfxFormatString2(s,IDS_STRING9039,strQws,strSlsN);
			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				if(IDNO == AfxMessageBox(s,MB_ICONINFORMATION|MB_YESNO|MB_DEFBUTTON2)){
					//Восстанавливаем
					ptrCmd1->CommandText = (_bstr_t)m_strBskt;
					ptrCmd1->CommandType = adCmdText;
					return;	//Выходим
				}
			}
		}
		BeginWaitCursor();
		try{
//AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				m_bCalcBskt = FALSE;	//Делать пересчёт
			}
			else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
				m_bCalcDb = FALSE;	//Делать пересчёт
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

		if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
			m_strSqlSum= m_strBskt;
			m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
//AfxMessageBox(_T("m_strSqlSum = ")+m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		}
		else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
			m_strSqlSum= m_strDb;
			m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
//AfxMessageBox(m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strDb, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
		}
		EndWaitCursor();
		if(IsEnableRec(ptrRs1)){
			if(!m_Edit1.IsEmpty()){
				m_Flg = false;
				IndCol = 16;
				if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
					OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
				}
				else if(IsQuery(_T("QT24NSumNum"),ptrRs1)){
					OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
				}
				m_Flg = true;
			}
		}
		OnBlckUnBlck(m_bBlck);
		OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
		OnShowBalance(m_strSysDate);
	}

	return;
}

void CDlg::On32776(void)
{
// Изменить
	UpdateData();
	CString strSql,strC,strD,strCst,strCT,strCG,strC11;
	CString strSqlTmp,strR3,sPrc,strR1;
	_variant_t vra;
	COleVariant vC,vBk;
	COleDateTime vD;
	VARIANT* vtl = NULL;
	short IndCol,i;
	double dSldr;
	CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);	

	strR3.Format(_T("%i"),m_Radio3);
	strR1.Format(_T("%i"),m_Radio1);

	if(!aSldr.IsEmpty()){
		dSldr = aSldr.GetAt(pSldr->GetPos());
	}
	else{
		dSldr = 0.00;
	}

	sPrc.Format(_T("%5.2f"),dSldr);
	sPrc.Replace(',','.');
	sPrc.TrimLeft();
	sPrc.TrimRight();

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){

			vBk = ptrRs1->GetBookmark();

			i = 0;	//№ нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();

			i = 1;	//Дата выписки нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));
			
			i = 12;		//Код контрАг покупателя
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCst = vC.bstrVal;
			strCst.TrimLeft();
			strCst.TrimRight();

			i = 16;		//Код товара
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCT = vC.bstrVal;
			strCT.TrimLeft();
			strCT.TrimRight();

			i = 17;		//Код группы
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCG = vC.bstrVal;
			strCG.TrimLeft();
			strCG.TrimRight();

			i = 6;	//Шт старые
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC11 = vC.bstrVal;
			strC11.TrimLeft();
			strC11.TrimRight();

			if(OnOverEdit(IDC_EDIT1,IDC_EDIT2)){
				if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
					switch(m_lCdOp){
						case 101:
							switch(m_Radio1){
								case 0:
									strSql = _T("UT60N_Dst ");
									break;
								case 1:
									strSql = _T("UT60NGrp_2Dst ");
									break;
							}
							strSql+= strR3			+ _T(",");		//1  Флаг % лин
							strSql+= strC			+ _T(",'");		//2  № нак
							strSql+= strD			+ _T("',");		//3  Дата выписки
							strSql+= strCst			+ _T(",");		//4  КОд покупателя
							strSql+= strCT			+ _T(",");		//5  Код товара
							strSql+= strC11			+ _T(",");		//6  Шт старые
							strSql+= m_Edit2		+ _T(",'");		//7  Шт новые
							strSql+= m_strSysDate	+ _T("',");		//8  Сис. дата
							strSql+= strCG			+ _T(",");		//9  Код группы
							strSql+= sPrc			+ _T(",");		//10 % лин цен
							strSql+= m_Edit6		+ _T(",'");		//11 Цена без скидки
							strSql+= m_strNT + _T("'");
							break;
						default:
							switch(m_Radio1){
								case 0:
									strSql = _T("UT60N ");
									break;
								case 1:
									strSql = _T("UT60NGrp_2 ");
									break;
							}
							strSql+= strR3			+ _T(",");		//1  Флаг % лин
							strSql+= strC			+ _T(",'");		//2  № нак
							strSql+= strD			+ _T("',");		//3  Дата выписки
							strSql+= strCst			+ _T(",");		//4  КОд покупателя
							strSql+= strCT			+ _T(",");		//5  Код товара
							strSql+= strC11			+ _T(",");		//6  Шт старые
							strSql+= m_Edit2		+ _T(",'");		//7  Шт новые
							strSql+= m_strSysDate	+ _T("',");		//8  Сис. дата
							strSql+= strCG			+ _T(",");		//9  Код группы
							strSql+= sPrc			+ _T(",'");		//10 % лин цен
							strSql+= m_strNT + _T("'");

							break;
					}
				}
				strSqlTmp=GetSourcePtr(ptrRs1);
				m_befCurCol = m_DataGrid1.get_Col();
				m_befCurType = GetTypeCol(ptrRs1,m_befCurCol);

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				BeginWaitCursor();
				try{
//AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);

					if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
						m_bCalcBskt = FALSE;	//Делать пересчёт
					}
					else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
						m_bCalcDb = FALSE;	//Делать пересчёт
					}

				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				EndWaitCursor();
				if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
					m_strSqlSum= m_strBskt;
					m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
		//AfxMessageBox(_T("m_strSqlSum = ")+m_strSqlSum);
					ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
					if(ptrRs100->State==adStateOpen) ptrRs100->Close();
					try{
						ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				}
				else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
					m_strSqlSum= m_strDb;
					m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
		//AfxMessageBox(m_strSqlSum);
					ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
					if(ptrRs100->State==adStateOpen) ptrRs100->Close();
					try{
						ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowGrid(m_strDb, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
				}
				if(IsEnableRec(ptrRs1)){
					ptrRs1->put_Bookmark(vBk);
					/*if(!m_Edit1.IsEmpty()){
						IndCol = 0;
						m_Flg = false;
						OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
					}
					*/
				}
				
			}

		}
		OnBlckUnBlck(m_bBlck);
		OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
		OnShowBalance(m_strSysDate);

	}
	return;
}

void CDlg::On32777(void)
{
//Удалить
	COleVariant vC,vBk;
	COleDateTime vD;
	CString strSql,strCT,strCG,strN,strD,strR,strC11,strSqlTmp;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	UpdateData();
	strR = m_Radio1==0 ? _T("0"):_T("1");
	if(IsEnableRec(ptrRs1)){

		strSqlTmp=GetSourcePtr(ptrRs1);
		m_befCurCol = m_DataGrid1.get_Col();
		m_befCurType = GetTypeCol(ptrRs1,m_befCurCol);

		ptrRs1->get_Bookmark(vBk);
		i = 0;	//№ нак
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strN = vC.bstrVal;
		strN.TrimRight();
		strN.TrimLeft();

		i = 1; //Дата
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_DATE);
		vD = vC.date;
		strD = vD.Format(_T("%Y-%m-%d"));
		
		i = 6;	//Шт
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strC11 = vC.bstrVal;
		strC11.TrimRight();
		strC11.TrimLeft();

		i = 16;	//Код Товара
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strCT = vC.bstrVal;
		strCT.TrimRight();
		strCT.TrimLeft();

		i = 17;	//Код Группы
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strCG = vC.bstrVal;
		strCG.TrimRight();
		strCG.TrimLeft();

		if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
			strSql = _T("DT60N ");
			strSql+= strR	+ _T(",'");
			strSql+= strD	+ _T("',");
			strSql+= strN	+ _T(",");
			strSql+= strCT	+ _T(",");
			strSql+= strCG	+ _T(",");
			strSql+= strC11	+ _T(",'");
			strSql+= m_strNT+ _T("'"); 
		}
		else if (IsQuery(_T("QT24NSumNum"),ptrRs1)){
			strSql = _T("DT24N ");
			strSql+= strCT;
		}
		ptrCmd1->CommandType = adCmdText;
		ptrCmd1->CommandText = (_bstr_t)strSql;
		try{
//AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
//AfxMessageBox(_T("after exec"));
			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				m_bCalcBskt = FALSE;//Делать пересчёт
			}
			else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
				m_bCalcDb = FALSE;	//Делать пересчёт
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

		if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
			m_strSqlSum= m_strBskt;
			m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
//AfxMessageBox(_T("m_strSqlSum = ")+m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
//AfxMessageBox(_T("1"));
			OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
//AfxMessageBox(_T("2"));
		}
		else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
			m_strSqlSum= m_strDb;
			m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
//AfxMessageBox(m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strDb, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
		}
		if(IsEnableRec(ptrRs1)){
			try{
				m_Flg = false;
				ptrRs1->PutBookmark(vBk);
				m_Flg = true;
			}
			catch(...){
				ptrRs1->MoveLast();
			}
		}
		else{
			if(strSqlTmp.Find(_T("QT60NSumNum"))!=-1){
				m_bCalcBskt = TRUE;	//Не делать пересчёт
			}
			else if(strSqlTmp.Find(_T("QT24NSumNum"))!=-1){
				m_bCalcDb	= TRUE;	//Не делать пересчёт
			}
		}
		OnShowBalance(m_strSysDate);
		OnBlckUnBlck(m_bBlck);
		OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	}
	return;
}

void CDlg::On32779(void)
{
//Переключение Режимов
	m_iBtSt = m_wndToolBar.GetToolBarCtrl().GetState(ID_BUTTON32779);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
//	OnOnlyRead(m_iBtSt);
	if(m_iBtSt==5){
		OnShowEdit();
		m_Flg = true;
	}


}


void CDlg::On32778(void)
{//Выбрать
		CString strSql,m_strSqlSum;
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		strSql = _T("QT100N ");
		if(IsQuery(_T("QT60NSumNum "),ptrRs1)){
			if(!ptrRs1->adoEOF) m_vBkBskt = ptrRs1->GetBookmark();
		}
		hMod=AfxLoadLibrary(_T("QNkl.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&);
		pDialog func=(pDialog)GetProcAddress(hMod,"startQNkl");
		if((func)(m_strNT, ptrCnn,strFndC,bFndC,strSql)){
			BeginWaitCursor();
//AfxMessageBox(strSql);
//				if(IsEnableRec(ptrRs1)){
					m_strQr  = strSql;
					m_strSqlSum= strSql;
					m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
// AfxMessageBox(m_strSqlSum);
					ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
					if(ptrRs100->State==adStateOpen) ptrRs100->Close();
					try{
						ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
//				}
				ptrCmd1 = NULL;
				ptrRs1 = NULL;

				ptrCmd1.CreateInstance(__uuidof(Command));
				ptrCmd1->ActiveConnection = ptrCnn;
				ptrCmd1->CommandType = adCmdText;

				ptrRs1.CreateInstance(__uuidof(Recordset));
				ptrRs1->CursorLocation = adUseClient;
				ptrRs1->PutRefSource(ptrCmd1);

				OnShowGrid(m_strQr,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Qr);
			EndWaitCursor();
			m_Icon = AfxGetApp()->LoadIcon(IDI_ICON2);		// Прочитать икону
			SetIcon(m_Icon,TRUE);			
		}
		else{
			m_strSqlSum= m_strBskt;
			m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
		//AfxMessageBox(m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strBskt,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			if(IsEnableRec(ptrRs1)){
				try{
					ptrRs1->put_Bookmark(m_vBkBskt);
				}
				catch(...){
					ptrRs1->MoveLast();
				}
			}

		}
		OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
		//Invalidate();

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	return;
}

void CDlg::On32780(void)
{
//Возврат в предыдущую выборку или пред. нак. в базе

	CString m_strSqlSum;
	COleVariant vC;
	short i;

	if(IsEnableRec(ptrRs1)){
		if(IsQuery(_T("QT60NSumNum "),ptrRs1)){
			if(!ptrRs1->adoEOF) m_vBkBskt = ptrRs1->GetBookmark();
		}
	}

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);
/*CString s;
s.Format(_T("m_shLstG1 = %i"),m_shLstG1);
AfxMessageBox(s);
*/
	switch(m_shLstG1){
		case -1:
			AfxMessageBox(IDS_STRING9053,MB_ICONSTOP);
			return;
			break;
		case 0:
			m_strSqlSum= m_strQr;
			m_strSqlSum.Replace(_T("QT100N"),_T("QT100NSum"));
		//AfxMessageBox(m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strQr,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Qr);
			if(IsEnableRec(ptrRs1)){
				try{
					ptrRs1->put_Bookmark(m_vBkQr);
				}
				catch(...){
					ptrRs1->MoveLast();
				}
			}
			m_Edit5 = _T("0");
			GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
			break;
		case 1:
			m_strSqlSum= m_strDb;
			m_strSqlSum.Replace(_T("QT24NSumNum"),_T("QT24NSum"));
		//AfxMessageBox(m_strSqlSum);
			ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
			if(ptrRs100->State==adStateOpen) ptrRs100->Close();
			try{
				ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strDb,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
			if(IsEnableRec(ptrRs1)){
				try{
					ptrRs1->put_Bookmark(m_vBkDb);
				}
				catch(...){
					ptrRs1->MoveLast();
				}
				i = 0;	//№ нак
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				m_Edit5 = vC.bstrVal;
			}
			else{
				m_Edit5 = _T("0");
			}
			GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
			break;
	}
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON2);		// Прочитать икону
	SetIcon(m_Icon,TRUE);	
	OnSetPtr();
	OnBlckUnBlck(m_bBlck);
	OnEnableButton();
	return;
}

void CDlg::On32783(void)
{
//печать
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			if(IsQuery(_T("QT100N "),ptrRs1)){

				UpdateData();

				COleVariant vC;
				COleDateTime vD;
				CString strSql,strCst,strSls,strCod,strRcp;
				CString strN,strD,strDD;
				HMODULE hMod;
				BOOL    bFndC   = TRUE;
				CString strFndC = _T("~");

				short i;

				strCod = _T("");

				i = 0;		//№ нак
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strN = vC.bstrVal;
				strN.TrimLeft();
				strN.TrimRight();
				strCod = strN + _T("~");

				i = 1;		// Дата выписки
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_DATE);
				vD = vC.date;
				strD = vD.Format(_T("%Y-%m-%d"));
							//Дата выдачи декорот
				strDD = m_OleDateTime1.Format(_T("%d.%m.%Y"));

				strCod+= strDD + _T("~");

				i = 7;		//Код контрАг покупателя
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCst = vC.bstrVal;
				strCst.TrimLeft();
				strCst.TrimRight();

				i = 8;		//Код контрАг получателя
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strRcp = vC.bstrVal;
				strRcp.TrimLeft();
				strRcp.TrimRight();
				strFndC+=strRcp;

				i = 13;		//Код контрАг продовца
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strSls = vC.bstrVal;
				strSls.TrimLeft();
				strSls.TrimRight();

				strCod+= strSls + _T("~");
				strCod+= strD;

				strSql = _T("QT38w1_1 ");
				strSql+= strCst + _T("~");
				strSql+= strSls;

				hMod=AfxLoadLibrary(_T("PrnVw.dll"));
				typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&,CString& );
				pDialog func=(pDialog)GetProcAddress(hMod,"startPrnVw");
				BeginWaitCursor();
					(func)(m_strNT, ptrCnn,strFndC,bFndC,strSql,strCod);
				EndWaitCursor();
				AfxFreeLibrary(hMod);
			}
		}
	}
	return;
}

void CDlg::OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void (*pFGrd)(CDatagrid1&,_RecordsetPtr&,_RecordsetPtr&),int EmpCol,int DefCol)
{
	if(!strSql.IsEmpty()){
		try{
//AfxMessageBox(_T("0"));
			Grd.putref_DataSource(NULL);
//AfxMessageBox(_T("1"));
			if(rs->State==adStateOpen) rs->Close();
//AfxMessageBox(_T("2"));
			Cmd->CommandText = (_bstr_t)strSql;
//AfxMessageBox(_T("3"));
			rs->PutRefSource(Cmd);
//AfxMessageBox(_T("4"));
			rs->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
//AfxMessageBox(_T("5"));
			m_Flg = false;
			Grd.putref_DataSource((LPUNKNOWN)rs);
//AfxMessageBox(_T("6"));

			pFGrd(Grd,rs,ptrRs100);
//AfxMessageBox(_T("7"));
		}
		catch(_com_error& e){
			Grd.putref_DataSource(NULL);
			AfxMessageBox(e.ErrorMessage());
		}
	}
}

void CDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
}

void CDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
}

void CDlg::OnClickedOk()
{
	if(m_bFndC){
		if(IsEnableRec(ptrRs1)){
			COleVariant vC;
			short i;
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_strFndC = vC.bstrVal;
			m_strFndC.TrimLeft();
			m_strFndC.TrimLeft();
		}
	}
	OnOK();
}

void CDlg::OnClickedCancel()
{
	if(m_bFndC){
		if(IsEnableRec(ptrRs1)){
			COleVariant vC;
			short i;
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_strFndC = vC.bstrVal;
			m_strFndC.TrimLeft();
			m_strFndC.TrimLeft();
		}
	}
	OnCancel();
	
}


void CDlg::InitDataGrid1(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr& rsG)
{
	CColumns GrdClms;
	CString strCap,strCapS,strCapC,strCapU,strRec,strSum,strCnt,strUnit;
	CString s;
	COleVariant vC;
	strSum = strCnt = strUnit = _T("0");
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9013);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		if(IsEnableRec(rsG)){
			i = 0;	//Сумма  нак
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strSum = vC.bstrVal;

			i = 1; // Кол-во поз
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strCnt = vC.bstrVal;

			i = 2; // Кол-во шт
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strUnit = vC.bstrVal;
		}
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
/*		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
*/
		AfxFormatString1(strCapS,IDS_STRING9054,strSum);
		AfxFormatString1(strCapC,IDS_STRING9055,strCnt);
		AfxFormatString1(strCapU,IDS_STRING9056,strUnit);

		strCap+=strCapS + _T(" ");
		strCap+=strCapC + _T(" ");
		strCap+=strCapU;

		Grd.put_Caption(strCap);
		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 1:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 90;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 4:
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("##0.00"));
				break;
			case 5:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 6:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				break;
			case 7:
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 8:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 15:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
/*			case 17:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
*/
			case 19:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 20:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 21:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}
void CDlg::InitDataGrid1Db(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr& rsG)
{
	CColumns GrdClms;
	CString strCap,strCapS,strCapC,strCapU,strRec,strSum,strCnt,strUnit;
	CString s;
	COleVariant vC;
	strSum = strCnt = strUnit = _T("0");
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9013);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		if(IsEnableRec(rsG)){
			i = 0;	//Сумма  нак
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strSum = vC.bstrVal;

			i = 1; // Кол-во поз
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strCnt = vC.bstrVal;

			i = 2; // Кол-во шт
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strUnit = vC.bstrVal;
		}
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
/*		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
*/
		AfxFormatString1(strCapS,IDS_STRING9054,strSum);
		AfxFormatString1(strCapC,IDS_STRING9055,strCnt);
		AfxFormatString1(strCapU,IDS_STRING9056,strUnit);

		strCap+=strCapS + _T(" ");
		strCap+=strCapC + _T(" ");
		strCap+=strCapU;
		Grd.put_Caption(strCap);

		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 1:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 90;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 55;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 4:
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("##0.00"));
				break;
			case 5:
				wdth = 55;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 6:
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				break;
			case 7:
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0.00"));
				break;
			case 8:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 15:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 19:
				wdth = 110;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 20:
				wdth = 85;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 21:
				wdth = 110;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}

void CDlg::InitDataGrid1Qr(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr& rsG)
{
	CColumns GrdClms;
	CString strCap,strRec,strSum,strCnt;
	CString s;
	long num,numRec;
	short i;
	float wdth;
	COleVariant vC;
	strSum = strCnt = _T("0");

//	strCap.LoadString(IDS_STRING9052);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());

	if(rs->State==adStateOpen){
		if(IsEnableRec(rsG)){
			i = 0;	//Сумма всех нак
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strSum = vC.bstrVal;

			i = 1; // Кол-во нак
			vC = GetValueRec(rsG,i);
			vC.ChangeType(VT_BSTR);
			strCnt = vC.bstrVal;

		}
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
		strRec.Format(_T(" %i"),numRec);
/*		strCap +=strRec;
*/
		AfxFormatString2(strCap,IDS_STRING9052,strSum,strCnt);
		Grd.put_Caption(strCap);
		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetButton(TRUE);
//				GrdClms.GetItem((COleVariant) i).put_Col(i);
//				GrdClms.GetItem((COleVariant) i).GetColIndex();
				break;
			case 1:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 90;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ### ### ##0.00"));
				break;
			case 3:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("##0"));
				break;
			case 4:
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("### ##0"));
				break;
			case 5:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 11:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 12:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}

void CDlg::InitDataGrid2(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr&)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9019);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			case 1:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 95;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 245;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 46;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}
void CDlg::InitDataGrid3(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr&)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9018);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			case 1:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 290;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}
void CDlg::InitDataGrid4(CDatagrid1& Grd,_RecordsetPtr& rs,_RecordsetPtr&)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9036);
	numRec = 0;

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				wdth = 55;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetButton(TRUE);
				break;
/*			case 1:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
*/
			case 2:
				wdth = 185;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
/*			case 3:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
*/
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
	}

}

BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, CDlg::RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID2, 218, CDlg::RowColChangeDatagrid2, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID3, 218, CDlg::RowColChangeDatagrid3, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID4, DISPID_CLICK, CDlg::ClickDatagrid4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID4, 210, CDlg::ButtonClickDatagrid4, VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID1, 210, CDlg::ButtonClickDatagrid1, VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID4, 218, CDlg::RowColChangeDatagrid4, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, CDlg::ChangeDatacombo2, VTS_NONE)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
//	AfxMessageBox(_T("RW1"));
	CString s;
	if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
					m_CurCol = m_DataGrid1.get_Col();
					if(m_CurCol==-1 ){
						m_CurCol=0;
					}
					m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
//s.Format(_T("RW1 m_iCurType = %i m_CurCol = %i"),m_iCurType,m_CurCol);
//AfxMessageBox(s);

			}
		}
	}
	if(m_Flg){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				if(m_iBtSt==5 ){
//					if(m_bFnd){
						OnShowEdit();
//					}
				}
			}
		}
	}
	m_Flg = true;
}
void CDlg::RowColChangeDatagrid2(VARIANT* LastRow, short LastCol)
{
	CString strSql,strC,s;
	COleVariant vC;
	short i;
	if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
		if(IsEnableRec(ptrRs2)){
			if(!ptrRs2->adoEOF){
					m_CurCol = m_DataGrid2.get_Col();
					if(m_CurCol==-1 || m_CurCol==0){
						m_CurCol=1;
					}
					m_iCurType = GetTypeCol(ptrRs2,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
/*	s.Format(_T("RW2 m_iCurType = %i m_CurCol = %i"),m_iCurType,m_CurCol);
	AfxMessageBox(s);
*/
			}
		}
	}
	if(m_Flg){
		if(IsEnableRec(ptrRs2)){
			if(!ptrRs2->adoEOF){
				switch(m_Radio1){
					case 0:
						UpdateData();
						s = m_Radio1==0 ? _T("0"):_T("1");
						switch(m_Radio1){
							case 0:
								i = 0;
								vC = GetValueRec(ptrRs2,i);
								vC.ChangeType(VT_BSTR);
								strC = vC.bstrVal;
								strC.TrimLeft();
								strC.TrimRight();

								strSql = _T("T27_T26 ");
								strSql+= s	+ _T(",");
								strSql+= strC;
								break;
							case 1:
								strSql = _T("T27_T26 ");
								strSql+= s	+_T(",0");
								break;
						}
						OnShowGrid(strSql,ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
						break;
				}
			}
		}
		OnEnterVidBox();
		OnShowCst();
		OnShowBalance(m_strSysDate);
		OnShowCstClc(m_Sldr2);

	}
	m_Flg = true;
}

void CDlg::RowColChangeDatagrid3(VARIANT* LastRow, short LastCol)
{
	CString strSql,strC,s;
	COleVariant vC;
	short i;

	if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
		if(IsEnableRec(ptrRs3)){
			if(!ptrRs3->adoEOF){
					m_CurCol = m_DataGrid3.get_Col();
					if(m_CurCol==-1 || m_CurCol==0){
						m_CurCol=1;
					}
					m_iCurType = GetTypeCol(ptrRs3,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
			}
		}
	}
	if(m_Flg){
		if(IsEnableRec(ptrRs3)){
			switch(m_Radio1){
				case 1:
					s = m_Radio1==0 ? _T("0"):_T("1");
					switch(m_Radio1){
						case 0:
							strSql = _T("T12_T27_1 ");
							strSql+= s + _T(",0");
							break;
						case 1:
								if(!ptrRs3->adoEOF){
										i = 0;
										vC = GetValueRec(ptrRs3,i);
										vC.ChangeType(VT_BSTR);
										strC = vC.bstrVal;
										strC.TrimLeft();
										strC.TrimRight();
								}
								else{
									strC = _T("0");
								}

								strSql = _T("T12_T27_1 ");
								strSql+= s + _T(",");
								strSql+= strC;
							break;
					}
					OnShowGrid(strSql,ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
					OnShowCst();
					OnShowCstClc(m_Sldr2);
					break;
			}
		}
	}
	m_Flg = true;
	OnShowBalance(m_strSysDate);
}

void CDlg::RowColChangeDatagrid4(VARIANT* LastRow, short LastCol)
{
	if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
		if(IsEnableRec(ptrRs4)){
			if(!ptrRs4->adoEOF){
					m_CurCol = m_DataGrid4.get_Col();
					if(m_CurCol==-1 ){
						m_CurCol=0;
					}
					m_iCurType = GetTypeCol(ptrRs4,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
			}
		}
	}
}

void CDlg::ButtonClickDatagrid4(short ColIndex)
{
//Вызов справочника контрагента
	COleVariant vC,vBk;
	short i;
	BOOL bFndC;
	bool bF(false);
	CString strFndC;
	if(IsEnableRec(ptrRs4)){		
		vBk = ptrRs4->GetBookmark();
		switch(ColIndex){
		case 0:
			{
				bFndC   = TRUE;
				strFndC = _T("");
				i=0;
				vC = GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				strFndC=vC.bstrVal;
				strFndC.TrimLeft();
				strFndC.TrimRight();

				BeginWaitCursor();

			//		m_SlpDay.SetDate(t1);
					HMODULE hMod;
					hMod=AfxLoadLibrary(_T("CtrAgnt.dll"));
					typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
					pDialog func=(pDialog)GetProcAddress(hMod,"startCtrAgnt");
					(func)(m_strNT, ptrCnn,strFndC,bFndC);

			//		m_SlpDay.SetDate(t1);
					AfxFreeLibrary(hMod);

				EndWaitCursor();
			}
			break;
		}
		try{
			ptrRs4->Requery(adCmdText);//adCmdStoredProc
			m_DataGrid4.Refresh();
			InitDataGrid4(m_DataGrid4,ptrRs4,ptrRs100);
			if(!IsEmptyRec(ptrRs4)){
				try{
					if(bFndC){
						OnFindInGrid(strFndC,ptrRs4,i,bF);
					}
					else{
						ptrRs4->PutBookmark(vBk);
					}
				}
				catch(...){
					ptrRs4->MoveLast();
				}
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
	}
}

void CDlg::OnOnlyRead(int i)
{
	COleVariant vC;
	double dSldr;
	CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);
	CString s;
	if(i == 101){
		if(IsEnableRec(ptrRs5)){
			int numRec = ptrRs5->GetRecordCount();
			i = 1;
			for(int t=0;t<numRec;t++){
				vC = GetValueRec(ptrRs5,i);
				vC.ChangeType(VT_R8);
				dSldr = vC.dblVal;
				if(dSldr==0){
					m_iSldr2 = t;
					pSldr->SetPos(m_iSldr2);
					OnShowCstClc(pSldr->GetPos());
					s.Format(_T("%4.2f"),aSldr.GetAt(pSldr->GetPos()));
					SetDlgItemText(IDC_STATIC_SLIDER2,s);
					break;
				}
			}
		}
//s.Format(_T("%i"),m_iSldr2);
//AfxMessageBox(s);
		((CEdit*)GetDlgItem(IDC_EDIT6))->SetReadOnly(FALSE);
		GetDlgItem(IDC_SLIDER2)->ShowWindow(FALSE);
		GetDlgItem(IDC_STATIC_SLIDER2)->ShowWindow(FALSE);
	}
	else{
//s.Format(_T("%i"),m_iSldr2);
//AfxMessageBox(s);
//		if(IsEnableRec(ptrRs5)){
		if(!aSldr.IsEmpty()){
			pSldr->SetPos(m_iSldr2);
			OnShowCstClc(pSldr->GetPos());
			if(aSldr.GetAt(pSldr->GetPos())<0){
				m_curColor = RGB(255,0,0);
			}

			else{
				m_curColor = RGB(0,0,128);
			}
			s.Format(_T("%4.2f"),aSldr.GetAt(pSldr->GetPos()));
			SetDlgItemText(IDC_STATIC_SLIDER2,s);

			((CEdit*)GetDlgItem(IDC_EDIT6))->SetReadOnly(TRUE);
			GetDlgItem(IDC_SLIDER2)->ShowWindow(TRUE);
			GetDlgItem(IDC_STATIC_SLIDER2)->ShowWindow(TRUE);
		}
		else{
			((CEdit*)GetDlgItem(IDC_EDIT6))->SetReadOnly(TRUE);
			GetDlgItem(IDC_SLIDER2)->ShowWindow(FALSE);
			GetDlgItem(IDC_STATIC_SLIDER2)->ShowWindow(FALSE);
		}
	}

}

void CDlg::SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt)
{
	if(IsEnableRec(pRs)){
		if(!pRs->adoEOF){
			COleVariant vC;
			CString strIns;

			vC = GetValueRec(pRs,iTxt);
			vC.ChangeType(VT_BSTR);
			strIns = vC.bstrVal;
			strIns.TrimLeft();
			strIns.TrimRight();
			Cmb.put_BoundText(strIns);
		}
	}
	else{
		Cmb.put_BoundText(_T(""));
	}

}

BOOL CDlg::IsQuery(CString substr, _RecordsetPtr rs)
{
	COleVariant vC;
	CString strC;

	vC = rs->GetSource();
	vC.ChangeType(VT_BSTR);
	strC=vC.bstrVal;
/*	AfxMessageBox(strC);
	AfxMessageBox(substr);
*/
	if(strC.Find(substr)==-1){
		return FALSE;
	}
	else{
		return TRUE;
	}
}

void CDlg::ClickDatagrid1()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
}

void CDlg::ClickDatagrid2()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID2);
}

void CDlg::ClickDatagrid3()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID3);
}

void CDlg::ClickDatagrid4()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID4);
}

HBRUSH CDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	switch(pWnd->GetDlgCtrlID()){
		case IDC_STATIC_SLIDER2:
			pDC->SetTextColor(m_curColor);
			break;
		case IDC_STATIC_RSTS:
			pDC->SetTextColor(m_curColor);
			break;
		case IDC_STATIC_INCS:
		case IDC_STATIC_EXPS:
		case IDC_STATIC_RSRS:
			pDC->SetTextColor(m_curColor);
			break;

	}

	return hbr;
}

void CDlg::OnEnterVidBox(void)
{
//AfxMessageBox("OnEnterVidBox");
	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF){
			short i;
			COleVariant vC;
			long iC3,iC4;
			CString s;
			
			i=9;		// Код Коробки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC3=vC.iVal;

			i=10;		// Код Упаковки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC4=vC.iVal;

			if(iC3==1 && iC4==1){
	//AfxMessageBox("iC3==1 && iC4==1");
				GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
				GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
				m_Edit4=_T("0");	//Коробки шт
				GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

				GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_HIDE);
				GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
				m_Edit3=_T("0");	//Упаковки шт
				GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

				GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);
			}
			else if(iC3==1 && iC4!=1){
	//AfxMessageBox("iC3==1 && iC4!=1");
				GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
				GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
				m_Edit4=_T("0");	//Коробки шт
				GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

				GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);
			}
			else if(iC3!=1 && iC4!=1){
	//AfxMessageBox("iC3!=1 && iC4!=1");
				GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT4)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
				GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);
			}
		}
		else{   
			GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
			m_Edit4=_T("0");	//Коробки шт
			GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

			GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
			m_Edit3=_T("0");	//Упаковки шт
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
			m_Edit2=_T("0");
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
			return;
		}
	}
	else{
		GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		m_Edit4=_T("0");	//Коробки шт
		GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

		GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		m_Edit3=_T("0");	//Упаковки шт
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

		GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
		m_Edit2=_T("0");
		GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
	}
}

void CDlg::OnEnterE_All(void)
{
	if(ptrRs2->State==adStateOpen){
		if(!IsEmptyRec(ptrRs2)){
			short i;
			COleVariant vC;
			long iC3,iC4;
			CString s;
			
			i=9;		// Код Коробки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC3=vC.iVal;

			i=10;		// Код Упаковки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC4=vC.iVal;

			if(iC3==1 && iC4==1){
			}
			else if(iC3==1 && iC4!=1){
				OnEnterE2_10();
			}
			else if(iC3!=1 && iC4!=1){
				OnEnterE2_00();
			}
		else{
			GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
			m_Edit4=_T("0");	//Коробки шт
			GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

			GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
			m_Edit3=_T("0");	//Упаковки шт
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
			m_Edit2=_T("0");
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
			}
		}
	}
	else{
		GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		m_Edit4=_T("0");	//Коробки шт
		GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

		GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		m_Edit3=_T("0");	//Упаковки шт
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

		GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
		m_Edit2=_T("0");
		GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
	}
}

void CDlg::OnEnterE2_10(void)
{
	short i;
	COleVariant vC;
	int iC5,iC6,iC10;
	CString strC;
	div_t div_rlt;

	UpdateData();

	i=6;		 //Упаковок в коробке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC5=vC.iVal;
	i=7;		 // Штук в Упаковке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC6=vC.iVal; 

	i=8;		 //Всего штук в контейнере
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC10=vC.iVal;

	m_Edit2.Remove(' ');
	if(!m_Edit2.IsEmpty()){
//				AfxMessageBox(m_Edit11);

		div_rlt=div(_wtoi(m_Edit2),iC6); 
		strC.Format(_T("%i"),div_rlt.quot);
		m_bCalcE10=false;
			GetDlgItem(IDC_EDIT3)->SetWindowText(strC);
		m_bCalcE10=true;
	}
}

void CDlg::OnEnterE2_00(void)
{
	short i;
	COleVariant vC;
	int iC5,iC6,iC10;
	CString strC;
	div_t div_rlt;

	UpdateData();

	i=6;		 //Упаковок в коробке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC5=vC.iVal;
	i=7;		 // Штук в Упаковке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC6=vC.iVal; 
	i=8;		 //Всего штук в контейнере
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC10=vC.iVal;

	m_Edit2.Remove(' ');
	if(!m_Edit2.IsEmpty()){
//				AfxMessageBox(m_Edit11);

		div_rlt=div(_wtoi(m_Edit2),iC10);
		strC.Format(_T("%i"),div_rlt.quot);
		m_bCalcE9=false; //Получилось коробок
			GetDlgItem(IDC_EDIT4)->SetWindowText(strC);
			m_Edit4=strC;
		m_bCalcE9=true;

		div_rlt=div(_wtoi(m_Edit2),iC6);
		strC.Format(_T("%i"),div_rlt.quot);
		m_bCalcE10=false;
			GetDlgItem(IDC_EDIT3)->SetWindowText(strC);
		m_bCalcE10=true;
//AfxMessageBox("OnEnterE11_00()");
	}
}

void CDlg::OnEnterE4_00(void)
{
	short i;
	COleVariant vC;
	long iC5,iC6,iC10,iC;
	CString strC;

	UpdateData();
	i=6;		 //Упаковок в коробке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC5=vC.iVal;
	iC=_wtoi(m_Edit4)*iC5;
	strC.Format(_T("%i"),iC);
	m_bCalcE10=false; //Получилось Упаковок
		GetDlgItem(IDC_EDIT3)->SetWindowText(strC);
	m_bCalcE10=true;

	i=7;		 // Штук в Упаковке
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC6=vC.iVal; 

	i=8; //Всего штук в контейнере
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC10=vC.iVal;
	iC=_wtoi(m_Edit4)*iC10;
	strC.Format(_T("%i"),iC);
	m_bCalcE11=false; //Получилось всего штук
		GetDlgItem(IDC_EDIT2)->SetWindowText(strC);
	m_bCalcE11=true;
}

void CDlg::OnEnterE3_10(void)
{
	short i;
	COleVariant vC;
	long iC5,iC6,iC;
	CString strC;
//	div_t div_rlt;

	UpdateData();
	i=6;
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC5=vC.iVal;

	i=7;
	vC=GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_I4);
	iC6=vC.iVal;
	iC=_wtoi(m_Edit3)*iC6;
	strC.Format(_T("%i"),iC);
	m_bCalcE11=false; //Получилось всего штук
		GetDlgItem(IDC_EDIT2)->SetWindowText(strC);
	m_bCalcE11=true;
}

void CDlg::OnEnterE3_00(void)
{
	if(ptrRs2->State==adStateOpen){
		short i;
		COleVariant vC;
		int iC5,iC6,iC;
		CString strC;
		div_t div_rlt;

		UpdateData();
		i=6;
		vC=GetValueRec(ptrRs2,i);
		vC.ChangeType(VT_I4);
		iC5=vC.iVal;

		div_rlt=div(_wtoi(m_Edit3),iC5);
		strC.Format(_T("%i"),div_rlt.quot);
		m_bCalcE9=false; //Получилось коробок
			GetDlgItem(IDC_EDIT4)->SetWindowText(strC);
		m_bCalcE9=true;

		i=7;
		vC=GetValueRec(ptrRs2,i);
		vC.ChangeType(VT_I4);
		iC6=vC.iVal;
		iC=_wtoi(m_Edit3)*iC6;
		strC.Format(_T("%i"),iC);
		m_bCalcE11=false; //Получилось всего штук
			GetDlgItem(IDC_EDIT2)->SetWindowText(strC);
		m_bCalcE11=true;
	}
}

BOOL CDlg::LoadFont(int h)
{
	m_OptionFont.lfHeight=h;
	m_OptionFont.lfWidth=0; 
	m_OptionFont.lfWeight=FW_BOLD;
	m_OptionFont.lfEscapement=0;
	m_OptionFont.lfOrientation=0;
	m_OptionFont.lfCharSet=RUSSIAN_CHARSET;
	m_OptionFont.lfClipPrecision=CLIP_DEFAULT_PRECIS;
	m_OptionFont.lfItalic=0;
	m_OptionFont.lfOutPrecision=OUT_DEFAULT_PRECIS;
	m_OptionFont.lfQuality=PROOF_QUALITY;
	m_OptionFont.lfPitchAndFamily=VARIABLE_PITCH;
	m_OptionFont.lfStrikeOut=0;
	m_OptionFont.lfUnderline=0;
	wcscpy(m_OptionFont.lfFaceName,_T("Tahoma")); //Arial

	return m_curFont.CreateFontIndirect(&m_OptionFont);
}

void CDlg::OnEnChangeEdit2()
{
	if(IsEnableRec(ptrRs2)){
		UpdateData();
		if(!m_Edit2.IsEmpty() || _wtoi(m_Edit2)!=0){
			if(!ptrRs2->adoEOF && m_bCalcE11){
				short i;
				COleVariant vC;
				long iC3,iC4;
				CString strC,strSql;

				i=9;		// Код Коробки
				vC=GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_I4);
				iC3=vC.iVal;

				i=10;		// Код Упаковки
				vC=GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_I4);
				iC4=vC.iVal;
				if(iC3!=1 && iC4!=1){
					OnEnterE2_00();
				}
				else if(iC3==1 && iC4!=1){
					OnEnterE2_10();
				}

			}
		}
	}
}

void CDlg::OnEnChangeEdit3()
{
	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF && m_bCalcE10){
			short i;
			COleVariant vC;
			long iC3,iC4;
			CString strC;
			i=9;		// Код Коробки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC3=vC.iVal;

			i=10;		// Код Упаковки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC4=vC.iVal;


			if(iC3!=1 && iC4!=1){
				OnEnterE3_00();
			}
			else if(iC3==1 && iC4!=1){
				OnEnterE3_10();
			}
		}
	}
}

void CDlg::OnEnChangeEdit4()
{
	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF && m_bCalcE9){
			short i;
			COleVariant vC;
			long iC3,iC4;
			CString strC;
			i=9;		// Код Коробки	
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC3=vC.iVal;

			i=10;	    // Код Кпаковки
			vC=GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_I4);
			iC4=vC.iVal;


			if(iC3!=1 && iC4!=1){
				OnEnterE4_00();
			}
			else if(iC3==1 && iC4!=1){
				//AfxMessageBox("iC3==1 && iC4!=1");
			}

		}
	}
}

void CDlg::OnShowBalance(CString strDate)
{
	CString sRsrS,sExpS,sIncS,sRstS;
	long lBlnc,lSldAdd,lSumCurRsr,lMng,lNumRsrOut;
	long lTtlExp,lBefExp,lCurExp,lOwnExp,lOthExp,lWthExp;
	int  iOthNum;

	lBlnc=lSldAdd=lSumCurRsr=lMng=lNumRsrOut=0;
	lTtlExp=lBefExp=lCurExp=lOwnExp=lOthExp=lWthExp=0;

	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF){
			//Расчёт расхода (дата системная)
			OnCalcExp(lTtlExp,lBefExp,lCurExp,lOwnExp,lOthExp,iOthNum,lWthExp,strDate);
			//Расчёт прихода (дата системная)
			OnCalcInc(lSldAdd,strDate);

			lBlnc = lSldAdd-lTtlExp;

			sRsrS.Format(_T("%i"),lOwnExp);
			if(lOwnExp!=0){
				m_curColor = RGB(0,128,192); //0,128,192
			}
			else{
				m_curColor = RGB(0,0,128);
			}
			SetDlgItemText(IDC_STATIC_RSRS,sRsrS);

			sExpS.Format(_T("%i"),lWthExp);
			if(lWthExp!=0){
				m_curColor = RGB(0,128,0);
			}
			else{
				m_curColor = RGB(0,0,128);
			}
			SetDlgItemText(IDC_STATIC_EXPS,sExpS);

			m_curColor = RGB(0,0,128);
			sIncS.Format(_T("%i"),lSldAdd);

			if(lBlnc<0){
				m_curColor = RGB(255,0,0);
			}
			sRstS.Format(_T("%i"),lBlnc);
			SetDlgItemText(IDC_STATIC_RSTS,sRstS);

			m_curColor = RGB(0,0,128);
			SetDlgItemText(IDC_STATIC_INCS,sIncS);

		}
		else{
/*			lBlnc=lSldAdd=lSumCurRsr=lMng=lNumRsrOut=0;
			lTtlExp=lBefExp=lCurExp=lOwnExp=lOthExp=lWthExp=0;
*/
			sIncS.Format(_T("%i"),lSldAdd);
			sExpS.Format(_T("%i"),lWthExp);
			sRsrS.Format(_T("%i"),lOwnExp);
			sRstS.Format(_T("%i"),lBlnc);
			m_curColor = RGB(0,0,128);
			SetDlgItemText(IDC_STATIC_INCS,sIncS);
			SetDlgItemText(IDC_STATIC_EXPS,sExpS);
			SetDlgItemText(IDC_STATIC_RSRS,sRsrS);
			SetDlgItemText(IDC_STATIC_RSTS,sRstS);
		}
	}
}

void CDlg::OnCalcExp(long& lTtlExp, long& lBefExp, long& lCurExp, long& lOwnExp, long& lOthExp, int& iOthNum, long& lWthExp, CString& strDate)
{
	CString strSql,strC;
	COleVariant vC;
	short i;
	i = 0;
	vC = GetValueRec(ptrRs2,i);
	vC.ChangeType(VT_BSTR);
	strC = vC.bstrVal;

	strSql = _T("QT69SUMCALC_1 ");	//QT69SUMCALC_1
	strSql+= strC + _T(",'");
	strSql+= strDate + _T("','");
	strSql+= m_strNT + _T("'");
//AfxMessageBox(strSql);
	ptrCmdExp->CommandText = (_bstr_t)strSql;
	if(ptrRsExp->State==adStateOpen) ptrRsExp->Close();
	try{
		ptrRsExp->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(IsEnableRec(ptrRsExp)){
			if(!ptrRsExp->adoEOF){
				i = 0;	//Итого
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lTtlExp = vC.lVal;

				i = 1;	//до нач. мес.
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lBefExp = vC.lVal;

				i = 2; //В этом мес. от нач.мес до Даты
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lCurExp = vC.lVal;

				i = 3; //Тот,кто вводит резерв
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lOwnExp = vC.lVal;

				i = 4; //Резерв других
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lOthExp = vC.lVal;

				i = 5; //Кол-во других без вводящего
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I2);
				iOthNum=vC.iVal;

				i = 6; //Расход без резервов]
				vC = GetValueRec(ptrRsExp,i);
				vC.ChangeType(VT_I4);
				lWthExp=vC.lVal;
			}
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description());
		return;
	}

}

void CDlg::OnCalcInc(long& lSldAdd, CString& strDate)
{
	CString strSql,strC;
	COleVariant vC;
	short i;

	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF){
			i = 0;
			vC = GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;

			strSql = _T("QT70SUMCALC_1 ");
			strSql+= strC + _T(",'");
			strSql+= strDate + _T("'");

			ptrCmdInc->CommandText = (_bstr_t)strSql;
			if(ptrRsInc->State==adStateOpen) ptrRsInc->Close();
			try{
				ptrRsInc->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
				if(IsEnableRec(ptrRsInc)){
					if(!ptrRsInc->adoEOF){
						i = 0;	//Итого Приход
						vC = GetValueRec(ptrRsInc,i);
						vC.ChangeType(VT_I4);
						lSldAdd = vC.lVal;
					}
				}
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
				return;
			}
		}
	}
}

void CDlg::OnBnClickedCheck2()
{
	//Сменить операцию
	m_Check2 = m_Check2 ? FALSE:TRUE;
	if(m_Check2){
		m_Check3 = m_Check4 = m_Check5 = FALSE;
	}
	OnSetPtr();
	OnEnableCheck(IDC_CHECK2,m_Check2);
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check2);

}

void CDlg::OnBnClickedCheck3()
{
	//Сменить клиента
	m_Check3 = m_Check3 ? FALSE:TRUE;
	if(m_Check3){
		m_Check2 = m_Check4 = m_Check5 = FALSE;
	}
	OnSetPtr();
	OnEnableCheck(IDC_CHECK3,m_Check3);
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check3);

}

void CDlg::OnBnClickedCheck4()
{
	//Сменить продавца
	m_Check4 = m_Check4 ? FALSE:TRUE;
	if(m_Check4){
		m_Check2 = m_Check3 = m_Check5 = FALSE;
	}
	OnSetPtr();
	OnEnableCheck(IDC_CHECK4,m_Check4);
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check4);
}

void CDlg::OnBnClickedCheck5()
{
	//Сменить торгового представителя
	m_Check5 = m_Check5 ? FALSE:TRUE;
	if(m_Check5){
		m_Check2 = m_Check3 = m_Check4 = FALSE;
	}
	OnSetPtr();
	OnEnableCheck(IDC_CHECK5,m_Check5);
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check5);
}

void CDlg::OnShowCst(void){

	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF){
			UpdateData();

			COleVariant vC;
			CString strD,strC,strSql,s;
			short i;
			i = 0;

			strD = m_OleDateTime1.Format(_T("%Y-%m-%d"));

			vC = GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();

			strSql = _T("QT85_3 ");
			strSql+= strC + _T(",'");
			strSql+= strD + _T("'");
			try{
				ptrCmdCst->CommandText = (_bstr_t)strSql;

				if(ptrRsCst->State==adStateOpen) ptrRsCst->Close();

				ptrRsCst->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
				if(!IsEmptyRec(ptrRsCst)){
					i = 0;	//Базовая цена
					vC = GetValueRec(ptrRsCst,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					s.Replace(',','.');
					m_Edit1 = s;
					SetDlgItemText(IDC_EDIT1,m_Edit1);
				}
				else{
					m_Edit1 = _T("0.0000");
					SetDlgItemText(IDC_EDIT1,m_Edit1);
				}
			}
			catch(_com_error& e){
				m_Edit1 = _T("0.0000");
				SetDlgItemText(IDC_EDIT1,m_Edit1);
				AfxMessageBox(e.Description());
			}
		}
		else{
			m_Edit1 = _T("0.0000");
			SetDlgItemText(IDC_EDIT1,m_Edit1);
		}
	}
	else{
		m_Edit1 = _T("0.0000");
		SetDlgItemText(IDC_EDIT1,m_Edit1);
	}
}

void CDlg::OnShowCstClc(int i)
{
	if(!aSldr.IsEmpty()){
		if(IsEnableRec(ptrRs2)){
			if(!ptrRs2->adoEOF){
				CString s;
				double dCstB;
				double dCstC;
				
				dCstB= _wtof(m_Edit1);
				switch(m_lCdOp){
					case 101:	//продажа дистьебютеру (свобод цены)
						break;
					default:
					if(aSldr.GetAt(i)==0 || aSldr.GetAt(i)==NULL){
						m_Edit6=m_Edit1;
					}
					else{		//Знак берем из aSldr.GetAt(i)
						dCstC = dCstB*(1+aSldr.GetAt(i)*0.01);
						m_Edit6.Format(_T("%10.4f"),dCstC);
					}
					break;
				}
			}
			else{
				m_Edit6 = _T("0.0000");
			}
		}
		else{
			m_Edit6 = _T("0.0000");
		}

		m_Edit6.TrimLeft();
		m_Edit6.TrimRight();
		m_Edit6.Replace(',','.');
		SetDlgItemText(IDC_EDIT6,m_Edit6);
	}
	else{
		m_Edit6 = m_Edit1;
		SetDlgItemText(IDC_EDIT6,m_Edit6);
	}
}


void CDlg::OnBnClickedButton1()
{
	//Сменить Операцию
	m_Check2 = FALSE;
	if(GetDlgItem(IDC_BUTTON1)->IsWindowEnabled()){
		OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check2);
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(m_Check2);
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){

				CString sSql,strC30,strC4;
				_variant_t vra;
				VARIANT* vtl = NULL;

				sSql = (BSTR)ptrCmd1->GetCommandText();

				if(sSql.Find(_T("QT60NSumNum"))!=-1){
					CString strSql,strC30,strC3,strC4;
					COleVariant vC;
					COleDateTime vD;
					short i;

					i = 0;		// Код операции
					vC = GetValueRec(ptrRsCmb2,i);
					vC.ChangeType(VT_BSTR);
					strC30 = vC.bstrVal;
					strC30.TrimLeft();
					strC30.TrimRight();

					i = 0;		//№ нак
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC4 = vC.bstrVal;
					strC4.TrimRight();
					strC4.TrimRight();

					i = 1;		//Дата
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_DATE);
					vD = vC.date;
					strC3 = vD.Format(_T("%Y-%m-%d"));


					strSql = _T("UT60N_C30 ");
					strSql+= strC4   + _T(",'");
					strSql+= strC3   + _T("',");
					strSql+= strC30  + _T(",'");
					strSql+= m_strNT + _T("'");

					ptrCmd1->CommandText = (_bstr_t)strSql;
					ptrCmd1->CommandType = adCmdText;
					try{
	//			AfxMessageBox(strSql);
						ptrCmd1->Execute(&vra,vtl,adCmdText);
						m_bBlck = TRUE;
						OnBlckUnBlck(m_bBlck);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					m_strSqlSum= m_strBskt;
					m_strSqlSum.Replace(_T("QT60NSumNum"),_T("QT60NSum"));
				//AfxMessageBox(m_strSqlSum);
					ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
					if(ptrRs100->State==adStateOpen) ptrRs100->Close();
					try{
						ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);

				}
			}
		}
	}
}

void CDlg::OnBnClickedButton2()
{
	m_Check3 = FALSE; 
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check3);
	((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(m_Check3);
}

void CDlg::OnBnClickedButton3()
{
	m_Check4 = FALSE;
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check4);
	((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(m_Check4);
}

void CDlg::OnBnClickedButton4()
{
	m_Check5 = FALSE;
	OnEnableButton(GetFocus()->GetDlgCtrlID(),m_Check5);
	((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(m_Check5);
}

void CDlg::OnShowSlider(void)
{
	CString strSql,strC,s;
	COleVariant vC;
	CSliderCtrl* pSldr = (CSliderCtrl*)GetDlgItem(IDC_SLIDER2);
	long numRec;
	double dSldr;
	short i;

	strSql = _T("QT346");
//AfxMessageBox(strSql);
	ptrCmd5->CommandText = (_bstr_t)strSql;
	try{
		m_Flg=FALSE;
//AfxMessageBox("S Close ptrRs4");
		if(ptrRs5->State==adStateOpen) ptrRs5->Close();

		m_Flg=FALSE;
		ptrRs5->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		m_Flg =TRUE;
		if(IsEnableRec(ptrRs5)){
			numRec = ptrRs5->GetRecordCount();
			aSldr.SetSize(numRec);
			pSldr->SetRange(0,numRec-1);
			i = 1;
			for(int t=0;t<numRec;t++){
				pSldr->SetTic(t);
				vC = GetValueRec(ptrRs5,i);
				vC.ChangeType(VT_R8);
				dSldr = vC.dblVal;
				if(dSldr==0){
					m_iSldr2 = t;
				}

				aSldr.SetAt(t,dSldr);
				ptrRs5->MoveNext();

			}
		}
		else{
			return;
		}
		if(ptrRs5->State==adStateOpen) ptrRs5->Close();

		pSldr->SetPos(m_iSldr2);

		s.Format(_T("%4.2f"),aSldr.GetAt(pSldr->GetPos()));
		s.TrimLeft();
		s.TrimRight();
		s.Replace(',','.');
	
		if(aSldr.GetAt(pSldr->GetPos())<0){
			m_curColor = RGB(255,0,0);
		}

		else{
			m_curColor = RGB(0,0,128);
		}
		SetDlgItemText(IDC_STATIC_SLIDER2,s);
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description());
	}
}

void CDlg::OnEnableButton(int iCheck, BOOL m_Check)
{
/*	COleVariant vC;
	vC = ptrRs1->GetSource();
	vC.ChangeType(VT_BSTR);
	AfxMessageBox(vC.bstrVal);
*/
	if(IsQuery(_T("QT60NSumNum"),ptrRs1)	||
	   IsQuery(_T("QT100N"),ptrRs1)	){
		((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(FALSE);

		m_Check2=m_Check3=m_Check4=m_Check5 = FALSE;
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(FALSE);

		((CButton*)GetDlgItem(IDC_CHECK2))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK3))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(FALSE);
		((CButton*)GetDlgItem(IDC_CHECK5))->EnableWindow(FALSE);
		return;
	}
	else{
		((CButton*)GetDlgItem(IDC_CHECK2))->EnableWindow(TRUE);
		((CButton*)GetDlgItem(IDC_CHECK3))->EnableWindow(TRUE);
		((CButton*)GetDlgItem(IDC_CHECK4))->EnableWindow(TRUE);
		((CButton*)GetDlgItem(IDC_CHECK5))->EnableWindow(TRUE);
	}
/*	switch(iCheck){
		case IDC_CHECK2:
			if(m_Check){
				if(!GetDlgItem(IDC_DATACOMBO2)->IsWindowEnabled()){
					GetDlgItem(IDC_DATACOMBO2)->EnableWindow(m_Check);
				}

				GetDlgItem(IDC_DATACOMBO2)->EnableWindow(m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK3:
			if(m_Check){
				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK4:
			if(m_Check){
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK5:
			if(m_Check){
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
			}
			break;
		default:
			((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(m_Check);
			break;

	}
*/
	switch(iCheck){
		case IDC_CHECK2:
			GetDlgItem(IDC_DATACOMBO2)->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(m_Check);

			if(m_Check){
				GetDlgItem(IDC_DATAGRID4)->EnableWindow(!m_Check);
				GetDlgItem(IDC_DATACOMBO1)->EnableWindow(!m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK3:
			GetDlgItem(IDC_DATAGRID4)->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(m_Check);


			if(m_Check){
				GetDlgItem(IDC_DATACOMBO2)->EnableWindow(!m_Check);
				GetDlgItem(IDC_DATACOMBO1)->EnableWindow(!m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK4:
			GetDlgItem(IDC_DATACOMBO1)->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(m_Check);

			if(m_Check){
				GetDlgItem(IDC_DATACOMBO2)->EnableWindow(!m_Check);
				GetDlgItem(IDC_DATAGRID4)->EnableWindow(!m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(!m_Check);
			}
			break;
		case IDC_CHECK5:
			((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(m_Check);

			if(m_Check){
				GetDlgItem(IDC_DATACOMBO1)->EnableWindow(!m_Check);
				GetDlgItem(IDC_DATACOMBO2)->EnableWindow(!m_Check);
				GetDlgItem(IDC_DATAGRID4)->EnableWindow(!m_Check);

				((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(!m_Check);
				((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(!m_Check);
			}
			break;
		case IDC_BUTTON1:
			((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(m_Check);
			GetDlgItem(IDC_DATACOMBO2)->EnableWindow(m_Check);
			break;
		case IDC_BUTTON2:
			((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(m_Check);
			GetDlgItem(IDC_DATAGRID4)->EnableWindow(m_Check);
			break;
		case IDC_BUTTON3:
			((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(m_Check);
			GetDlgItem(IDC_DATACOMBO1)->EnableWindow(m_Check);
			break;
		case IDC_BUTTON4:
			((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(m_Check);
			break;
		default:
			((CButton*)GetDlgItem(IDC_BUTTON1))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON2))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON3))->EnableWindow(m_Check);
			((CButton*)GetDlgItem(IDC_BUTTON4))->EnableWindow(m_Check);
			break;
	}

}

void CDlg::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	CSliderCtrl* pSldr = (CSliderCtrl*)pScrollBar;
	CString s;
	switch(pScrollBar->GetDlgCtrlID()){
		case IDC_SLIDER2:
			if(aSldr.GetAt(pSldr->GetPos())<0){
				m_curColor = RGB(255,0,0);
			}

			else{
				m_curColor = RGB(0,0,128);
			}
			s.Format(_T("%4.2f"),aSldr.GetAt(pSldr->GetPos()));
			SetDlgItemText(IDC_STATIC_SLIDER2,s);
			OnShowCstClc(pSldr->GetPos());
			break;
	}

	CDialog::OnHScroll(nSBCode, nPos, pScrollBar);
}

void CDlg::OnBnClickedRadio1()
{
	UpdateData();
	CString s,strSql;
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			strSql = _T("T12_T27_1 ");
			strSql+= s + _T(",0");
			OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
			break;
		case 1:
			strSql = _T("T27_T26 ");
			strSql+= s	+_T(",0");
			OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
	}
	OnShowBalance(m_strSysDate);
	if(m_iBtSt==5){
		OnShowEdit();
		m_Flg = true;
	}

/*	OnShowCst();
	OnShowCstClc(m_Sldr2);
*/
}

void CDlg::OnBnClickedRadio2()
{
	UpdateData();
	CString s,strSql;
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			strSql = _T("T12_T27_1 ");
			strSql+= s + _T(",0");
			OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
			break;
		case 1:
			strSql = _T("T27_T26 ");
			strSql+= s	+_T(",0");
			OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
	}
	OnShowBalance(m_strSysDate);
	if(m_iBtSt==5){
		OnShowEdit();
		m_Flg = true;
	}
/*	OnShowCst();
	OnShowCstClc(m_Sldr2);
*/
}

void CDlg::OnEnableCheck(int Idc, BOOL bCheck)
{
	int Id;
	CString s;
//	((CButton*)GetDlgItem(iCheck))->SetCheck(m_Check);
	GotoDlgCtrl(GetDlgItem(IDC_CHECK2));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_CHECK2:
				bCheck = m_Check2;
				((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(bCheck);
				break;
			case IDC_CHECK3:
				bCheck = m_Check3;
				((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(bCheck);
				break;
			case IDC_CHECK4:
				bCheck = m_Check4;
				((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(bCheck);
				break;
			case IDC_CHECK5:
				bCheck = m_Check5;
				((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(bCheck);
				break;
		}
		NextDlgCtrl();
	} while (Id!=IDC_RADIO3);
	GotoDlgCtrl(GetDlgItem(Idc));
}

void CDlg::ButtonClickDatagrid1(short ColIndex)
{
	COleVariant vC;
	COleDateTime vD;
	short i;
	CString strC,strD,strC3;
	if(IsQuery(_T("QT100N"),ptrRs1)){
		switch(ColIndex){
			case 0:
				if(IsEnableRec(ptrRs1)){

					m_vBkQr = ptrRs1->GetBookmark();

					i = 0;	//№ нак
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;
					strC.TrimLeft();
					strC.TrimRight();

					i = 1; //Дата выписки
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_DATE);
					vD = vC.date;
					strD = vD.Format(_T("%Y-%m-%d"));

					i = 7;	//Код контраг покуп.
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC3 = vC.bstrVal;
					strC3.TrimLeft();
					strC3.TrimRight();

					m_strDb = _T("QT24NSumNum ");
					m_strDb+= strC	+ _T(",'");
					m_strDb+= strD	+ _T("',");
					m_strDb+= strC3;
//AfxMessageBox(m_strDb);

					ptrCmd1 = NULL;
					ptrRs1  = NULL;
					ptrCmd1.CreateInstance(__uuidof(Command));
					ptrCmd1->ActiveConnection = ptrCnn;
					ptrCmd1->CommandType = adCmdText;

					ptrRs1.CreateInstance(__uuidof(Recordset));
					ptrRs1->CursorLocation = adUseClient;
					ptrRs1->PutRefSource(ptrCmd1);

					m_strSqlSum= m_strDb;
					m_strSqlSum.Replace(_T("QT24NSumNum"),_T("QT24NSum"));
				//AfxMessageBox(m_strSqlSum);
					ptrCmd100->CommandText = (_bstr_t)m_strSqlSum;
					if(ptrRs100->State==adStateOpen) ptrRs100->Close();
					try{
						ptrRs100->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowGrid(m_strDb,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Db);
					/*if(IsEnableRec(ptrRs1)){
						i = 0;	//№ нак
						vC = GetValueRec(ptrRs1,i);
						vC.ChangeType(VT_BSTR);
						m_Edit5 = vC.bstrVal;
						m_Edit5.TrimLeft();
						m_Edit5.TrimRight();
						GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
					}*/
					OnSetPtr();
					OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
					OnBlckUnBlck(m_bBlck);
					OnEnableButton();

				}
				break;
		}
	}
}



void CDlg::OnSetPtr(void)
{
	COleVariant vC;
	if(IsEnableRec(ptrRs1)){
		if(!IsQuery(_T("QT100N"),ptrRs1)){
			CString strC;
			short i;

			i = 1; //Дата выписки
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_OleDateTime1.ParseDateTime(vC.bstrVal);
			m_OleDateTime2.ParseDateTime(vC.bstrVal);
			UpdateData(false);

			i = 0;	//№ нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit5 = vC.bstrVal;
			m_Edit5.TrimLeft();
			m_Edit5.TrimRight();
			GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);

			i = 11;	//Код операции
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
//AfxMessageBox(strC);
			OnFindInCombo(strC,&m_DataCombo2,ptrRsCmb2,0,1);

			i = 12;	//Код контраг покуп.
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Flg = false;
			OnFindInGrid(strC,ptrRs4,(short)0,m_Flg);
			m_Flg = true;

			i = 14;	//Код контраг прод.
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
//AfxMessageBox(strC);
			OnFindInCombo(strC,&m_DataCombo1,ptrRsCmb1,2,1);

			i = 16;	//Код товара
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();

		}
		else{
			m_OleDateTime1 = (COleDateTime::GetCurrentTime());
			m_OleDateTime1.ParseDateTime(m_OleDateTime1.Format(_T("%d-%m-%Y")));
			m_OleDateTime2 = (COleDateTime::GetCurrentTime());
			m_OleDateTime2.ParseDateTime(m_OleDateTime2.Format(_T("%d-%m-%Y")));
//AfxMessageBox(m_OleDateTime1.Format(_T("%d-%m-%Y")));
			UpdateData(false);

			m_Edit5 = _T("0");
			GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
			ptrRsCmb1->MoveFirst();
			SetTextCombo(m_DataCombo1,ptrRsCmb1,1);
		}
	}
	else{
		m_OleDateTime1 = (COleDateTime::GetCurrentTime());
		m_OleDateTime1.ParseDateTime(m_OleDateTime1.Format(_T("%d-%m-%Y")));
//AfxMessageBox(m_OleDateTime1.Format(_T("%d-%m-%Y")));
		UpdateData(false);
		if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
			if(_wtoi(m_strNumBskt)==0){
				m_Edit5 = _T("0");
				OnFindInCombo(_T("13"),&m_DataCombo2,ptrRsCmb2);
			}
			else{
				m_Edit5 = m_strNumBskt;
			}
			GetDlgItem(IDC_EDIT5)->SetWindowTextW(m_Edit5);
		}
		else{
			AfxMessageBox(_T("111"));
		}

		ptrRsCmb1->MoveFirst();
		SetTextCombo(m_DataCombo1,ptrRsCmb1,1);
	}
}

void CDlg::OnBlckUnBlck(BOOL& bBlck)
{
	if(IsEnableRec(ptrRs1)){
		if(!IsQuery(_T("QT100N"),ptrRs1)){
			bBlck = TRUE;
		}
		else{
			bBlck = FALSE;
		}
	}
	else{
		bBlck = FALSE;
	}
	GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(!bBlck);
	GetDlgItem(IDC_DATACOMBO2)->EnableWindow(!bBlck);
	GetDlgItem(IDC_DATACOMBO1)->EnableWindow(!bBlck);
	GetDlgItem(IDC_DATAGRID4)->EnableWindow(!bBlck);
}

void CDlg::OnBnClickedCheck1()
{
//Получить № нак
	m_Edit5.TrimLeft();
	m_Edit5.TrimRight();
	COleVariant vC;
	short i;
	CString strC1,strD,strSql,s;
	if(!IsEmptyRec(ptrRsCmb1)){
		if(m_Edit5.IsEmpty()) m_Edit5 = _T("0");
		if(IsQuery(_T("QT60NSumNum"),ptrRs1)){
			if(IsEmptyRec(ptrRs1) &&
			   _wtoi(m_Edit5)==0){

			    UpdateData();
				strD = m_OleDateTime1.Format(_T("%Y-%m-%d"));

				i = 0;	//Код продовца
				vC = GetValueRec(ptrRsCmb1,i);
				vC.ChangeType(VT_BSTR);
				strC1 = vC.bstrVal;
				strC1.TrimLeft();
				strC1.TrimRight();

				strSql = _T("UT53_D_S '");
				strSql+= strD    + _T("',0,");	// 0 - расход
				strSql+= strC1   + _T(",'");
				strSql+= m_strNT + _T("'");
				try{
					ptrCmdNum->CommandText = (_bstr_t)strSql;

					if(ptrRsNum->State==adStateOpen) ptrRsNum->Close();

					ptrRsNum->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
					if(!IsEmptyRec(ptrRsNum)){
						i = 0;
						vC = GetValueRec(ptrRsNum,i);
						vC.ChangeType(VT_BSTR);
						s = vC.bstrVal;
						s.TrimLeft();
						s.TrimRight();
						m_Edit5 = m_strNumBskt = s;
						 
						SetDlgItemText(IDC_EDIT5,m_Edit5);
					}
					else{
						m_Edit5 = _T("0");
						SetDlgItemText(IDC_EDIT5,m_Edit5);
					}
				}
				catch(_com_error& e){
					m_Edit5 = _T("0");
					SetDlgItemText(IDC_EDIT5,m_Edit5);
					AfxMessageBox(e.Description());
				}
			}
		}
	}
	OnSetNumber();
}

void CDlg::OnSetNumber(void)
{
	if(_wtoi(m_Edit5)==0){
		m_Check1= FALSE;
	}
	else{
		m_Check1= TRUE;
	}
	((CButton*)GetDlgItem(IDC_CHECK1))->EnableWindow(!m_Check1);
	((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
}


CString CDlg::GetSourcePtr(_RecordsetPtr rs)
{
	COleVariant vC;
	CString strC;

	vC = rs->GetSource();
	vC.ChangeType(VT_BSTR);
	strC=vC.bstrVal;

	return strC;
}

void CDlg::OnBnClickedButton5()
{
//Узнать за что % скидок
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			CString strSql,strC,strD,strCst;
			COleVariant vC;
			COleDateTime vD;
			short i;
			if(IsQuery(_T("QT60NSumNum "),ptrRs1)){
				strSql = _T("QT117_1 ");
			}
			else if(IsQuery(_T("T24NSumNum "),ptrRs1)){
				strSql = _T("QT116_1 ");
			}
			else{
				return;	// просто выходи без вызова окна Расшифровка
			}

			i = 0;	//№ нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();

			i = 1;	//Дата выписки нак
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));
			
			i = 12;		//Код контрАг покупателя
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCst = vC.bstrVal;
			strCst.TrimLeft();
			strCst.TrimRight();

			strSql+= strCst + _T(",'");
			strSql+= strD	+ _T("',");
			strSql+= strC	+ _T(",'");
			strSql+= m_strNT+ _T("'");

			HMODULE hMod;
			BOOL bFndC;
			CString strFndC;
			bFndC   = FALSE;
			strFndC = _T("");
			hMod=AfxLoadLibrary(_T("VwClcDsc.dll"));
			typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&);
			pDialog func=(pDialog)GetProcAddress(hMod,"startVwClcDsc");
			BeginWaitCursor();
				(func)(m_strNT, ptrCnn,strFndC,bFndC,strSql);
			EndWaitCursor();
			AfxFreeLibrary(hMod);
		}
	}
}

void CDlg::ChangeDatacombo2()
{
//	if(GetDlgItem(IDC_DATACOMBO2)->IsWindowEnabled()){

		m_lCdOp = OnChangeCombo(m_DataCombo2,ptrRsCmb2,1);

		switch(m_lCdOp){
			case 100:	//Транзит
				break;
			case 101:	//Продажа товара Дист

				OnOnlyRead((int)m_lCdOp);

				m_Edit6 = _T("0");
				GetDlgItem(IDC_EDIT6)->SetWindowTextW(m_Edit6);

				break;
			default:
				OnOnlyRead((int)m_lCdOp);
				break;
		}
//	}
}

afx_msg LRESULT CDlg::OnSwitchFocus(WPARAM wParam, LPARAM lParam)
{
/*AfxMessageBox(_T("OnSwitchFocus"));
CString s;
*/

	CWnd* pFocusWnd  = (CWnd*)wParam;
	CWnd* pEdTB = &m_EditTBCh;
	CWnd* pEd2  = (CWnd*)GetDlgItem(IDC_EDIT2);
	CWnd* pEd3  = (CWnd*)GetDlgItem(IDC_EDIT3);
	CWnd* pEd4  = (CWnd*)GetDlgItem(IDC_EDIT4);
	if(pFocusWnd->GetSafeHwnd() == pEd2->GetSafeHwnd()){
		if(pEd3->GetStyle()==SW_SHOW){
//AfxMessageBox(_T("pFocusWnd->GetSafeHwnd() == pEd3->GetSafeHwnd()"));
			pEd3->SetFocus();
		}else{
			pEdTB->SetFocus();
		}
	}
	else if(pFocusWnd->GetSafeHwnd() == pEd3->GetSafeHwnd()){
//AfxMessageBox(_T("pFocusWnd->GetSafeHwnd() == pEd2->GetSafeHwnd()"));
		pEd4->SetFocus();
	}
	else if(pFocusWnd->GetSafeHwnd() == pEd4->GetSafeHwnd()){
//AfxMessageBox(_T("pFocusWnd->GetSafeHwnd() == pEd2->GetSafeHwnd()"));
		pEdTB->SetFocus();
	}
	else if(pFocusWnd->GetSafeHwnd() == pEdTB->GetSafeHwnd()){
//AfxMessageBox(_T("pFocusWnd->GetSafeHwnd() == pEdTB->GetSafeHwnd()"));
		pEd2->SetFocus();
	}
	return LRESULT();
}

afx_msg LRESULT CDlg::OnAddRec(WPARAM wParam, LPARAM lParam)
{
//AfxMessageBox(_T("CDlg::OnAddRec"));
	CWnd* pFocusWnd  = (CWnd*)wParam;
	CWnd* pEd2  = (CWnd*)GetDlgItem(IDC_EDIT2);
	CWnd* pEdTB = &m_EditTBCh;
	if(pFocusWnd->GetSafeHwnd() == pEd2->GetSafeHwnd()){
		if(m_iBtSt!=5){	//добавляем запись по Enter
			if(IsQuery(_T("QT60NSumNum "),ptrRs1) ||
			   IsQuery(_T("QT24NSumNum "),ptrRs1)){

				On32775();
			}
			pEdTB->SetFocus();
		}
		else{	//делаем корректировку по Enter
			if(IsQuery(_T("QT60NSumNum "),ptrRs1) ||
			   IsQuery(_T("QT24NSumNum "),ptrRs1)){

				On32776();
			}
			pEdTB->SetFocus();
		}
	}
	return LRESULT();
}
