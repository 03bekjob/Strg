// Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg.h"

#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"
//#include <wchar.h>
/*#include <stdlib.h>*/

// CDlg dialog

IMPLEMENT_DYNAMIC(CDlg, CDialog)

CDlg::CDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDlg::IDD, pParent)
	, m_CurCol(0)
	, m_iBtSt(0)
	, m_Flg(true)
	, m_strNT(_T(""))
	, m_bFnd(false)
	, m_strFndC(_T(""))
	, m_bFndC(FALSE)
	, m_pLstWnd(NULL)

	, m_OleDateTime1(COleDateTime::GetCurrentTime())
	, m_Radio1(0)
	, m_stcSlsCst(_T("0"))
	, m_Edit1(_T("0"))
	, m_Edit2(_T("1"))
	, m_strSqlOld(_T(""))
	, m_strSqlOldF(_T(""))
	, m_Cd(0)
	, m_CdOp(0)
	, m_bAuto(FALSE)
	, m_strCst(_T(""))
{

}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2 = NULL;

	if(ptrRs3->State==adStateOpen) ptrRs3->Close();
	ptrRs3 = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
	ptrRsCmb2 = NULL;

	if(ptrRsCmb3->State==adStateOpen) ptrRsCmb3->Close();
	ptrRsCmb3 = NULL;

	if(ptrRsCmb4->State==adStateOpen) ptrRsCmb4->Close();
	ptrRsCmb4 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_Control(pDX, IDC_DATACOMBO3, m_DataCombo3);
	DDX_Control(pDX, IDC_DATACOMBO4, m_DataCombo4);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_Radio(pDX, IDC_RADIO1, m_Radio1);
	DDX_Text(pDX, IDC_STATIC_SLSCST, m_stcSlsCst);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_COMMAND(ID_BUTTON32771, On32771)
	ON_COMMAND(ID_BUTTON32772, On32772)
	ON_COMMAND(ID_BUTTON32773, On32773)
	ON_COMMAND(ID_BUTTON32775, On32775)
	ON_COMMAND(ID_BUTTON32776, On32776)
	ON_COMMAND(ID_BUTTON32777, On32777)
	ON_COMMAND(ID_BUTTON32778, On32778)
	ON_COMMAND(ID_BUTTON32779, On32779)
	ON_COMMAND(ID_BUTTON32783, On32783)

	ON_EN_CHANGE(ID_EDIT_TOOLBAR, OnEnChangeEditTB)

	ON_BN_CLICKED(IDOK, &CDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_OK, &CDlg::OnClickedOk)
	ON_BN_CLICKED(IDC_CANCEL, &CDlg::OnClickedCancel)
	ON_EN_CHANGE(IDC_EDIT1, &CDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT2, &CDlg::OnEnChangeEdit2)
	ON_BN_CLICKED(IDC_RADIO1, &CDlg::OnBnClickedRadio1)
	ON_BN_CLICKED(IDC_RADIO2, &CDlg::OnBnClickedRadio2)
	ON_BN_CLICKED(IDC_BUTTON1, &CDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"SlsCst.dll"));

	CString s,strSql,strSql2,strC;
	s.LoadString(IDS_STRING9014);
	this->SetWindowText(s);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);	

	m_Edit1Flt.SubclassDlgItem(IDC_EDIT1,this);
	m_Edit2Flt.SubclassDlgItem(IDC_EDIT2,this);

	InitStaticText();

	int iTBCtrlID;
	short i;
	COleVariant vC;

	m_wndToolBar.Create(this, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_TOOLTIPS |
	                        CBRS_FLYBY | CBRS_BORDER_BOTTOM);
	m_wndToolBar.LoadToolBar(IDR_TOOLBAR1);
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
	m_wndToolBar.SetButtonInfo(5,ID_EDIT_TOOLBAR,TBBS_SEPARATOR,130);
	// определяем координаты области, занимаемой разделителем
	CRect rEdit; 
	m_wndToolBar.GetItemRect(5,&rEdit);
	rEdit.left+=6; rEdit.right-=6; // отступы
	if(!(m_EditTB.Create(WS_CHILD|
		ES_AUTOHSCROLL|WS_VISIBLE|WS_TABSTOP|
		WS_BORDER,rEdit,&m_wndToolBar,ID_EDIT_TOOLBAR))) return -1;
	
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

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);
	m_strSqlOld = _T("QT87_1 '");
	m_strSqlOld+= m_strNT + _T("'");
	OnShowGrid(m_strSqlOld, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);

	if(m_bFndC){
		short fCol=0;
		m_Flg = false;
		OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
		m_Flg = true;
	}

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
			if(IsEnableRec(ptrRs2)){
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

	ptrCmdCmb1 = NULL;
	ptrRsCmb1 = NULL;

	strSql = _T("QT19");
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
		m_DataCombo1.put_ListField(_T("Сокр. наим. валюты"));
//AfxMessageBox("ok");
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb2 = NULL;
	ptrRsCmb2  = NULL;

	strSql = _T("QT20");
	ptrCmdCmb2.CreateInstance(__uuidof(Command));
	ptrCmdCmb2->ActiveConnection = ptrCnn;
	ptrCmdCmb2->CommandType = adCmdText;
	ptrCmdCmb2->CommandText = (_bstr_t)strSql;

	ptrRsCmb2.CreateInstance(__uuidof(Recordset));
	ptrRsCmb2->CursorLocation = adUseClient;
	ptrRsCmb2->PutRefSource(ptrCmdCmb2);
	try{
		m_DataCombo2.putref_RowSource(NULL);
		ptrRsCmb2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo2.putref_RowSource((LPUNKNOWN) ptrRsCmb2);
		m_DataCombo2.put_ListField(_T("Вид Курса"));
	}
	catch(_com_error& e){
		m_DataCombo2.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb3 = NULL;
	ptrRsCmb3  = NULL;

	strSql = _T("QT83");
	ptrCmdCmb3.CreateInstance(__uuidof(Command));
	ptrCmdCmb3->ActiveConnection = ptrCnn;
	ptrCmdCmb3->CommandType = adCmdText;
	ptrCmdCmb3->CommandText = (_bstr_t)strSql;

	ptrRsCmb3.CreateInstance(__uuidof(Recordset));
	ptrRsCmb3->CursorLocation = adUseClient;
	ptrRsCmb3->PutRefSource(ptrCmdCmb3);
	try{
		m_DataCombo3.putref_RowSource(NULL);
		ptrRsCmb3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo3.putref_RowSource((LPUNKNOWN) ptrRsCmb3);
		m_DataCombo3.put_ListField(_T("Вид базовой цены"));
		SetTextCombo(m_DataCombo3,ptrRsCmb3,1);
	}
	catch(_com_error& e){
		m_DataCombo3.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb4 = NULL;
	ptrRsCmb4  = NULL;

	strSql = _T("QT84");
	ptrCmdCmb4.CreateInstance(__uuidof(Command));
	ptrCmdCmb4->ActiveConnection = ptrCnn;
	ptrCmdCmb4->CommandType = adCmdText;
	ptrCmdCmb4->CommandText = (_bstr_t)strSql;

	ptrRsCmb4.CreateInstance(__uuidof(Recordset));
	ptrRsCmb4->CursorLocation = adUseClient;
	ptrRsCmb4->PutRefSource(ptrCmdCmb4);
	try{
		m_DataCombo4.putref_RowSource(NULL);
		ptrRsCmb4->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo4.putref_RowSource((LPUNKNOWN) ptrRsCmb4);
		m_DataCombo4.put_ListField(_T("Символ операции"));
		SetTextCombo(m_DataCombo4,ptrRsCmb4,1);

	}
	catch(_com_error& e){
		m_DataCombo4.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	if(LoadFont(-16)) {
//AfxMessageBox("ok");
		GetDlgItem(IDC_STATIC_EQUEL)->SetFont(&m_curFont,FALSE);
		GetDlgItem(IDC_STATIC_SLSCST)->SetFont(&m_curFont,FALSE);
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
void CDlg::OnEnableButtonBar(int iBtSt, CToolBar* pTlBr)
{	
/*	if(iBtSt==4){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);

		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,TRUE);

	}
	else{
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,TRUE);

		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);

	}
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);
*/
	CString strSrc;
	strSrc = (BSTR)ptrCmd1->GetCommandText();
	if(strSrc.Find(_T("QT85"))!=-1){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
	}
	else{
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,TRUE);
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
	}
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);

}

void CDlg::OnEnChangeEditTB(void)
{
	UpdateData();
	CString s,sLstCh,str,msg;
	UINT nChar;
	char* ch;
	if(ptrRs1->State==adStateOpen){
		m_EditTB.GetWindowText(str);
//AfxMessageBox(s);
	    sLstCh = str.Right(1);
		ch = (LPSTR)sLstCh.GetBuffer(sLstCh.GetLength());
//AfxMessageBox(sLstCh);
		nChar = toascii(*ch);
//		s.Format("%i",nChar);
//AfxMessageBox(s);

		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
			m_CurCol = m_DataGrid1.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		}

		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			m_CurCol = m_DataGrid2.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs2,m_CurCol);
		}

		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			m_CurCol = m_DataGrid3.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs3,m_CurCol);
		}
	switch(m_iCurType){
		case adSmallInt: //2
		case adInteger:	//3
		case adNumeric:
		case adBoolean:	//11
//AfxMessageBox("adBoolean:");
		
			if (nChar>=48 && nChar<=57) //Цифры
			{
				break;
			}
			else{
				msg.LoadString(IDS_STRING9015);
				AfxMessageBox(msg,MB_ICONINFORMATION);
				return;
			}
			break;
		case adDouble:
		case adCurrency:
//			AfxMessageBox("Currency");
		//Ввод чисел  с плавающей точкой	
			if (nChar>=48 && nChar<=57) //Цифры
			{
				break;
			}
			else
				switch(nChar)
			{
				case 19:	// <-
				case 4:		// ->
				case 8:		// <- Backspace
				case 127:	// delete
				case 44:	// ,
				case 46:	// .
					break;
				default:
					msg.LoadString(IDS_STRING9015);
					AfxMessageBox(msg,MB_ICONINFORMATION);
					return;
			}
			break;
		case adChar:
		case adVarChar:
		case adBSTR:
//AfxMessageBox("adChar");
			if(nChar==39){
				return;
			}
			break;
	}

		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
			m_Flg = true;
		}

		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
			m_Flg = true;
		}

		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs3,m_CurCol,m_Flg);
			m_Flg = true;
		}

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount,strC;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

	if(!m_bAuto){
		GotoDlgCtrl(GetDlgItem(IdBeg));
		do{
			Id = GetFocus()->GetDlgCtrlID();
			switch(Id){
				case IDC_EDIT1:
					GetDlgItem(IDC_EDIT1)->GetWindowTextW(m_Edit1);
					m_Edit1.Replace(',','.');
					m_Edit1.TrimRight();
					m_Edit1.TrimLeft();
					if(m_Edit1==' ') m_Edit1.Empty();
					if(m_Edit1.IsEmpty() || _wtof(m_Edit1)==0){
						count++;
						GetDlgItem(IDC_STATIC_COMBO3)->GetWindowText(strItem);
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
					if(m_Edit2.IsEmpty() || _wtof(m_Edit1)==0){
						count++;
						strItem.LoadString(IDS_STRING9028);
						strCount.Format(_T("%i"),count);
						strCount+=_T(") ");
						strMsg+=strCount+strItem+_T("\n\t");
					}
					break;
			}
			NextDlgCtrl();
		} while (Id!=IdEnd);
	}
	else{
			GetDlgItem(IDC_EDIT2)->GetWindowTextW(m_Edit2);
			m_Edit2.Replace(',','.');
			m_Edit2.TrimRight();
			m_Edit2.TrimLeft();
			if(m_Edit2==' ') m_Edit2.Empty();
			if(m_Edit2.IsEmpty() || _wtof(m_Edit2)==0){
				count++;
				strItem.LoadString(IDS_STRING9028);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
	}
//Раздел Grid
	if(m_iBtSt==5 ){	//Режим Корректировки
		//Прейскурантнык цены
		if(IsEnableRec(ptrRs1)){
			if(ptrRs1->adoEOF){
				count++;
				strItem.LoadString(IDS_STRING9014);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
			count++;
			strItem.LoadString(IDS_STRING9014);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{		//Режим Добавить/Удалить проверяем без учёта switch
				// т.к. и Товар-Группа должны уже существовать
		//Группа
		if(IsEnableRec(ptrRs3)){
			if(ptrRs3->adoEOF){
				count++;
				strItem.LoadString(IDS_STRING9024);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
				count++;
				strItem.LoadString(IDS_STRING9024);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
		}
			//Товар
		if(IsEnableRec(ptrRs2)){
			if(ptrRs2->adoEOF){
				count++;
				strItem.LoadString(IDS_STRING9027);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
				count++;
				strItem.LoadString(IDS_STRING9027);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
		}
	}

//Раздел Combo
	if(!m_bAuto){
		strC = m_DataCombo1.get_BoundText();	//Код Валюты

		strC.TrimRight();
		strC.TrimLeft();
		if(strC==' ') strC.Empty();
		if(strC.IsEmpty()){
			count++;
			GetDlgItem(IDC_STATIC_COMBO1)->GetWindowText(strItem);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}

	strC = m_DataCombo2.get_BoundText();	//Код вида курса

	strC.TrimRight();
	strC.TrimLeft();
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO2)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	if(!m_bAuto){
		strC = m_DataCombo3.get_BoundText();	//Код вида  цены

		strC.TrimRight();
		strC.TrimLeft();
		if(strC==' ') strC.Empty();
		if(strC.IsEmpty()){
			count++;
			strItem.LoadString(IDS_STRING9033);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}

	strC = m_DataCombo4.get_BoundText();	//Код операции

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		strItem.LoadString(IDS_STRING9017);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	short i;
	CString strC;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){

			i=12;	//Код баз. цены
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;

			OnFindInCombo(strC,&m_DataCombo3,ptrRsCmb3,0,1);
			OnOnlyRead(m_Cd);
/*			if(m_Cd==3){
				((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(TRUE);
			}
			else{
				((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
			}
*/
			i=13;	//Код опер
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;

			OnFindInCombo(strC,&m_DataCombo4,ptrRsCmb4,0,1);

			i = 8;	//Исходящая цена
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimRight();
			m_Edit1.TrimLeft();
			m_Edit1.Replace(',','.');
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

			i = 6;	//Коэф-т
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit2 = vC.bstrVal;
			m_Edit2.TrimRight();
			m_Edit2.TrimLeft();
			m_Edit2.Replace(',','.');
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);


			i=15;	//Код валюты
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;

			OnFindInCombo(strC,&m_DataCombo1,ptrRsCmb1,0,1);

			i=16;	//Код вида курса
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;

			OnFindInCombo(strC,&m_DataCombo2,ptrRsCmb2,0,1);


		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_GROUP)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_RADIO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_RADIO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_COMBO3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9029);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9026);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9030);
	GetDlgItem(IDC_BUTTON1)->SetWindowText(sTxt);

	GetDlgItem(IDC_STATIC_SLSCST)->SetWindowText(m_stcSlsCst);
}

//void CDlg::On32771(void)
{
//Возврат в Корзину
	/*OnShowGrid(m_strSqlOld,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
*/
	return;
}

void CDlg::On32772(void)
{
//Занести
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			CString sSql;
			sSql = (BSTR)ptrCmd1->GetCommandText();
			if(sSql.Find(_T("QT87_1 "))!=-1){
				_variant_t vra;
				VARIANT* vtl = NULL;
				CString strSql,strD;
				short i;
				COleVariant vC;
				COleDateTime vD;

				i = 14;	//Дата
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_DATE);
				vD = vC.date;
				strD = vD.Format(_T("%Y-%m-%d"));

				strSql = _T("IT87T85ALL_2 '");
				strSql+= strD + _T("','");
				strSql+= m_strNT + _T("'");

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				try{
//			AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowGrid(m_strSqlOld,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			}
		}
	}	
	return;
}

void CDlg::On32773(void)
{
	return;
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql;
	CString strCT,strCG,strCS,strR,strCr,strCrV;
	CString strSlsCst,strC5T83,strC6T84,strС9D;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol,i;
	int iB,iE;

	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	strCT = _T("0");
	strCG = _T("0");
	strCS = _T("0");
	strR  = m_Radio1==0 ? _T("0"):_T("1");
	strС9D  = m_OleDateTime1.Format(_T("%Y-%m-%d"));

	if(m_Cd==4){
		iB = IDC_EDIT1;
		iE = IDC_EDIT1;
	}
	else{
		iB = IDC_EDIT1;
		iE = IDC_EDIT2;
	}
	if(OnOverEdit(iB,iE)){
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

		i = 0;		//Код Валюты
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strCr = vC.bstrVal;

		i = 0;		//Код Вида курса
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strCrV = vC.bstrVal;

		i = 0;		// Код Базовой Цены
		vC = GetValueRec(ptrRsCmb3,i);
		vC.ChangeType(VT_BSTR);
		strC5T83 = vC.bstrVal;

		i = 0;		// Код операции
		vC = GetValueRec(ptrRsCmb4,i);
		vC.ChangeType(VT_BSTR);
		strC6T84 = vC.bstrVal;

		GetDlgItem(IDC_STATIC_SLSCST)->GetWindowText(strSlsCst);
		strSlsCst.TrimLeft();
		strSlsCst.TrimRight();
		strSlsCst.Replace(',','.');


		strSql = _T("IT87_2 ");
		strSql+= strR		+ _T(",");
		strSql+= strCT		+ _T(",");	//Код Товара
		strSql+= strCG		+ _T(",");	//Код Группы
		strSql+= strC5T83	+ _T(",");	//Код баз. цены
		strSql+= strC6T84	+ _T(",");	//Код операци
		strSql+= m_Edit2	+ _T(",");	//Коэф-т
		strSql+= strSlsCst  + _T(",'");	//Базовая цена
		strSql+= strС9D		+ _T("',");	//Дата
		strSql+= strCr		+ _T(",");	//Код Валюты
		strSql+= strCrV		+ _T(",");  //Код Вида Курса
		strSql+= m_Edit1	+ _T(",'");	//Исходящая цена цена
		strSql+= m_strNT	+ _T("'");

		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//			AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		OnShowGrid(m_strSqlOld, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			IndCol = 0;	//по товару
			m_Flg = false;
			OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
			m_Flg = true;
		}
	}

	return;
}

void CDlg::On32776(void)
{
// Изменить
	UpdateData();
	CString strSql,strC,strR,strC5T83,strC6T84,strCT,strCG,strCr,strCrV;
	CString strSlsCst,strSlsCstOld,strC9D,strC9DOld;
	_variant_t vra;
	COleVariant vC;
	COleDateTime vD;
	VARIANT* vtl = NULL;
	short IndCol,i;

	strCT  = _T("0");
	strCG  = _T("0");
	strR  = m_Radio1==0 ? _T("0"):_T("1");
	strC9D  = m_OleDateTime1.Format(_T("%Y-%m-%d"));
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			
			i = 4;	//Старая цена
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strSlsCstOld = vC.bstrVal;
			strSlsCstOld.TrimRight();
			strSlsCstOld.TrimRight();
			strSlsCstOld.Replace(',','.');

			i = 0;	//Код товара
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCT = vC.bstrVal;
			strCT.TrimRight();
			strCT.TrimRight();

			i = 11;	//Код гр.
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCG = vC.bstrVal;
			strCG.TrimRight();
			strCG.TrimRight();

			i = 14;	//Старая дата
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strC9DOld = vD.Format(_T("%Y-%m-%d"));

			i = 0;		// Код Базовой Цены
			vC = GetValueRec(ptrRsCmb3,i);
			vC.ChangeType(VT_BSTR);
			strC5T83 = vC.bstrVal;
			strC5T83.TrimLeft();
			strC5T83.TrimRight();

			i = 0;		// Код операции
			vC = GetValueRec(ptrRsCmb4,i);
			vC.ChangeType(VT_BSTR);
			strC6T84 = vC.bstrVal;
			strC6T84.TrimLeft();
			strC6T84.TrimRight();

			GetDlgItem(IDC_STATIC_SLSCST)->GetWindowText(strSlsCst);
			strSlsCst.TrimLeft();
			strSlsCst.TrimRight();
			strSlsCst.Replace(',','.');

			i = 0;		//Код Валюты
			vC = GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BSTR);
			strCr = vC.bstrVal;
			strCr.TrimLeft();
			strCr.TrimRight();

			i = 0;		//Код Вида курса
			vC = GetValueRec(ptrRsCmb2,i);
			vC.ChangeType(VT_BSTR);
			strCrV = vC.bstrVal;
			strCrV.TrimLeft();
			strCrV.TrimRight();

			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				strSql = _T("UT87_1 ");
				strSql+= strR			+ _T(",");
				strSql+= strCG			+ _T(",");			//Код гр.
				strSql+= strCT			+ _T(",");			//Код тов.
				strSql+= strC5T83		+ _T(",");			//Код баз цены
				strSql+= strC6T84		+ _T(",");			//Код опер.
				strSql+= m_Edit2		+ _T(",");			//Коэф-т
				strSql+= strSlsCst		+ _T(",'");			//Баз цена
				strSql+= strC9D			+ _T("',");			//Дата ввода
				strSql+= strCr			+ _T(",");			//Код вал
				strSql+= strCrV			+ _T(",");			//Код вида курса
				strSql+= m_Edit1		+ _T(",");			//Исходящая цена
				strSql+= strSlsCstOld	+ _T(",'");		    //Старая цена
				strSql+= strC9DOld		+ _T("','");	    //Сатаря дата
				strSql+= m_strNT + _T("'");

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				try{
//AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowGrid(m_strSqlOld, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				if(IsEnableRec(ptrRs1)){
					if(!strSlsCst.IsEmpty()){
						switch(m_Radio1){
							case 0:
								IndCol = 0;
								strC = strCT;
								break;
							case 1:
								IndCol = 11;
								strC = strCG;
								break;
						}
//AfxMessageBox(strCG);
						m_Flg = false;
						OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
					}
				}
			}

		}
	}
	return;
}

void CDlg::On32777(void)
{
//Удалить
	UpdateData();
	COleVariant vC,vBk;
	COleDateTime vD;
	CString strSql,strR,strC1,strCG,strD;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	strCG = _T("0");
	strC1 = _T("0");
	strR = m_Radio1==0 ? _T("0"):_T("1");
	if(IsEnableRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);
			switch(m_Radio1){
				case 0:		//по Товару
					i=0;	//Код товара
					vC=GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC1=vC.bstrVal;
					strC1.TrimLeft();
					strC1.TrimRight();
					break;
				case 1:		//по Группе
					i=11;	//Код Группы
					vC=GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strCG=vC.bstrVal;
					strCG.TrimLeft();
					strCG.TrimRight();
					break;
			}
			m_strSqlOldF = (BSTR)ptrCmd1->GetCommandText();
			if(m_strSqlOldF.Find(_T("QT87_1"))!=-1){
				i = 14;	//Дата
			}
			else{
				i = 1;
			}
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));

			if(m_strSqlOldF.Find(_T("QT87_1"))!=-1){
				strSql = _T("DT87_2 ");
				strSql+= strR	+ _T(",");
				strSql+= strC1	+ _T(",");
				strSql+= strCG  + _T(",'");
				strSql+= m_strNT+ _T("'");
			}
			else{
				strSql = _T("DT85_2 ");
				strSql+= strR	+ _T(",");
				strSql+= strC1	+ _T(",");
				strSql+= strCG  + _T(",'");
				strSql+= strD	+ _T("'");
			}

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strSqlOldF, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
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
			m_Flg = true;
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
/*	else{
		CString strC;
		strC=_T("4"); //Договорная
		OnFindInCombo(strC,&m_DataCombo3,ptrRsCmb3,0,1);
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
		m_Edit1 = _T("0");
		GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
	}
*/

}


void CDlg::On32778(void)
{//Выбрать
//		m_SlpDay.SetDate(t1);
		CString strSqlTmp;
		strSqlTmp = (BSTR)ptrCmd1->GetCommandText();

		if(ptrRs1->State==adStateOpen){
			if(!IsEmptyRec(ptrRs1)){
				if(strSqlTmp.Find(_T("QT87_1"))!=-1){
					m_strSqlOld = strSqlTmp;
				}
			}
		}
		CString strSql;
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		strSql = _T("QT85_1 ");
		hMod=AfxLoadLibrary(_T("QSqlCmb1.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&);
		pDialog func=(pDialog)GetProcAddress(hMod,"startQSqlCmb1");
		if((func)(m_strNT, ptrCnn,strFndC,bFndC,strSql)){
			BeginWaitCursor();
				OnShowGrid(strSql,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			EndWaitCursor();
			m_Icon = AfxGetApp()->LoadIcon(IDI_ICON2);		// Прочитать икону
			SetIcon(m_Icon,TRUE);			
			OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
		}
		else{
			OnShowGrid(m_strSqlOld,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		}
		//Invalidate();

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	return;
}

void CDlg::On32783(void)
{
	return;
}

void CDlg::OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void (*pFGrd)(CDatagrid1&,_RecordsetPtr&),int EmpCol,int DefCol)
{
	if(!strSql.IsEmpty()){
		Cmd->CommandText = (_bstr_t)strSql;
		try{
			Grd.putref_DataSource(NULL);
			if(rs->State==adStateOpen) rs->Close();
			rs->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			m_Flg = false;
			Grd.putref_DataSource((LPUNKNOWN)rs);
			m_CurCol = Grd.get_Col();
	//		s.Format("%i",m_CurCol);
	//		AfxMessageBox(s);
			if(m_CurCol==-1 || m_CurCol==EmpCol ){
				m_CurCol = DefCol;
			}
			m_iCurType = GetTypeCol(rs,m_CurCol);
			pFGrd(Grd,rs);
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


void CDlg::InitDataGrid1(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9013);
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
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 115;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("# ##0.00"));
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				break;
			case 5:
				wdth = 44;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("# ##0.0000"));
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				break;
			case 7:
				wdth = 30;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 8:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetNumberFormat(_T("# ##0.00"));
				GrdClms.GetItem((COleVariant) i).SetAlignment(1);
				break;
			case 9:
				wdth = 40;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 10:
				wdth = 200;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 14:
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
void CDlg::InitDataGrid2(CDatagrid1& Grd,_RecordsetPtr& rs)
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
				wdth = 105;
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
void CDlg::InitDataGrid3(CDatagrid1& Grd,_RecordsetPtr& rs)
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
				wdth = 90;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 270;
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

BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, CDlg::RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID2, 218, CDlg::RowColChangeDatagrid2, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID3, 218, CDlg::RowColChangeDatagrid3, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, CDlg::ChangeDatacombo2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO3, 1, CDlg::ChangeDatacombo3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO4, 1, CDlg::ChangeDatacombo4, VTS_NONE)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1 || m_CurCol==0){
					m_CurCol=1;
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
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
	if(m_Flg){
//AfxMessageBox("r2 m_Flg=TRUE");
		if(IsEnableRec(ptrRs2)){
			if(!ptrRs2->adoEOF){
				switch(m_Radio1){
					case 0:
						CString s,strC,strSql;
						COleVariant vC;
						short i;
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
						OnShowGrid(strSql,ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
						break;
				}
			}
		}
	}
	m_Flg = true;
}

void CDlg::RowColChangeDatagrid3(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(IsEnableRec(ptrRs3)){
			if(!ptrRs3->adoEOF){
				switch(m_Radio1){
					case 1:
	//AfxMessageBox("r4");
						CString strSql,strC,s;
						COleVariant vC;
						short i;
						s = m_Radio1==0 ? _T("0"):_T("1");
						switch(m_Radio1){
							case 0:
								strSql = _T("T12_T27_1 ");
								strSql+= s + _T(",0");
								break;
							case 1:
								if(ptrRs3->State==adStateOpen){
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
						OnShowGrid(strSql,ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
						break;
				}
			}
		}
	}
	m_Flg = true;
}

void CDlg::OnOnlyRead(int i)
{
	if(i == 4){
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
	}
	else{
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(TRUE);
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

void CDlg::ChangeDatacombo1()
{
	OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
}
void CDlg::ChangeDatacombo2()
{
	long cd;
	cd = OnChangeCombo(m_DataCombo2,ptrRsCmb2,1);
	switch(cd){
		case 0:
			break;
		default:
			break;
	}
}

void CDlg::ChangeDatacombo3()
{
	m_Cd = OnChangeCombo(m_DataCombo3,ptrRsCmb3,1);
	CString strCombo;
	switch(m_Cd){
		case 4:	//Договрная
			GetDlgItem(IDC_BUTTON2)->EnableWindow(FALSE);
			m_Edit2=_T("1");
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
			strCombo=_T("4");
			OnFindInCombo(strCombo,&m_DataCombo4,ptrRsCmb4);
			((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(false);
			GetDlgItem(IDC_DATACOMBO4)->ShowWindow(false);
			break;
		default:
			OnOnlyRead(m_Cd);
			GetDlgItem(IDC_BUTTON2)->EnableWindow(TRUE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(true);
			GetDlgItem(IDC_DATACOMBO4)->ShowWindow(true);
			break;
	}
}

void CDlg::ChangeDatacombo4()
{
	m_CdOp = OnChangeCombo(m_DataCombo4,ptrRsCmb4,1);
	if(m_CdOp!=5){
		OnCalcSls((int)m_CdOp,m_Edit1,m_Edit2);
	}
	else{
		double kD;
		m_Edit2.Replace(',','.');
		kD = _wtof(m_Edit2);
		if(kD==0){
			AfxMessageBox(IDS_STRING9031,MB_ICONSTOP);
		}
		else{
			OnCalcSls((int)m_CdOp,m_Edit1,m_Edit2);
		}
	}
}


void CDlg::OnEnChangeEdit1()
{
	if(m_Cd==4){  //Договорная
//AfxMessageBox("m_Cd==4");
		UpdateData();
		OnCalcSls((int) m_CdOp,m_Edit1,m_Edit2);

	}
}

void CDlg::OnEnChangeEdit2()
{
	UpdateData();
	double kD,BsCst;
	m_Edit2.Replace(',','.');
	m_Edit1.Replace(',','.');
	kD	  = _wtof(m_Edit2);
	BsCst = _wtof(m_Edit1);
	if(kD==0 && m_CdOp==5){
		AfxMessageBox(IDS_STRING9031,MB_ICONSTOP);
	}
	else{
		if(BsCst!=0 && m_Cd!=4)
		OnCalcSls((int)m_CdOp,m_Edit1,m_Edit2);
	}
}

void CDlg::OnBnClickedRadio1()
{
	UpdateData();
	COleVariant vC;
	CString strSql,strC,s;
	short i;
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			switch(m_Radio1){
				case 0:
					strSql = _T("T12_T27_1 ");
					strSql+= s + _T(",0");
					break;
				case 1:
					if(ptrRs3->State==adStateOpen){
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
			OnShowGrid(strSql,ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
			break;
		case 1:
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
			OnShowGrid(strSql,ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
	}
/*	if(m_iBtSt==5 ){
		OnShowEdit();
	}
*/
}

void CDlg::OnBnClickedRadio2()
{
	UpdateData();
	CString strSql,strC,s;
	COleVariant vC;
	short i;
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			switch(m_Radio1){
				case 0:
					strSql = _T("T12_T27_1 ");
					strSql+= s + _T(",0");
					break;
				case 1:
					if(ptrRs3->State==adStateOpen){
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
			OnShowGrid(strSql,ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
			break;
		case 1:
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
			OnShowGrid(strSql,ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
	}
/*	if(m_iBtSt==5 ){
		OnShowEdit();
	}
*/
}

void CDlg::OnBnClickedButton1()
{//Автопрайс
	UpdateData();
	CString strSql;
	CString strCT,strCG,strR,strCr,strCrV;
	CString strC6T84,strС9D;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;
	short i;

	m_bAuto = TRUE;	//для OverEdit

	strCT = _T("0");
	strCG = _T("0");
	strR  = m_Radio1==0 ? _T("0"):_T("1");
	strС9D  = m_OleDateTime1.Format(_T("%Y-%m-%d"));


	if(OnOverEdit(IDC_EDIT1,IDC_EDIT2)){

		m_bAuto = FALSE;

		i = 0;								// Код операции
		vC = GetValueRec(ptrRsCmb4,i);
		vC.ChangeType(VT_BSTR);
		strC6T84 = vC.bstrVal;

		i = 0;								//Код Вида курса
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strCrV = vC.bstrVal;

		switch(m_Radio1){
			case 0:							//по Товару
				i = 0;
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strCT = vC.bstrVal;
				strCT.TrimLeft();
				strCT.TrimRight();
				break;
			case 1:							//по Группе
				i = 0;
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strCG = vC.bstrVal;
				strCG.TrimLeft();
				strCG.TrimRight();
				break;
		}

		strSql = _T("IT87C3C4_2 ");
		strSql+= strR		+ _T(",");
		strSql+= strCT		+ _T(",");	 //Код Товара
		strSql+= strCG		+ _T(",'");	 //Код Группы
		strSql+= strС9D		+ _T("',");	 //Дата
		strSql+= m_Edit2	+ _T(",");	 //Коэф-т
		strSql+= strC6T84	+ _T(",");	 //Код операци
		strSql+= strCrV		+ _T(",'");  //Код Вида Курса
		strSql+= m_strNT	+ _T("'");
//AfxMessageBox(strSql);
		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
	//			AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		OnShowGrid(m_strSqlOld,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			IndCol = 1;	//по Товару
			m_Flg = false;
			OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
			m_Flg = true;
		}

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
	wcscpy(m_OptionFont.lfFaceName,_T("Arial"));

	return m_curFont.CreateFontIndirect(&m_OptionFont);
}

void CDlg::OnCalcSls(int i, CString strBsCst, CString strK)
{
	double dBsCst,dK,dSlsCst;
	COleCurrency cBsCst,cK,cSlsCst;

	strBsCst.Replace(',','.');
	strK.Replace(',','.');
	dBsCst = _wtof(strBsCst);	 //m_Edit1
	dK     = _wtof(strK);		 //m_Edit2

/*	CString s;
	s.Format("%i",i);
AfxMessageBox(s);
*/
	switch(i){
	case 1:			/* = */
		dSlsCst= dBsCst;
		break;
	case 2:			/* + */
		dSlsCst= dBsCst+dK;
		break;
	case 3:		/*  -  */
		dSlsCst= dBsCst-dK;
		break;
	case 4:		/*  *  */
		dSlsCst= dBsCst*dK;
		break;
	case 5:		/*  /  */
		dSlsCst= dBsCst/dK;
		break;
	case 6:		/*  +%  */
		dSlsCst= dBsCst*(1 + dK*0.01);
		break;
	case 7:		/*  -%  */
		dSlsCst= dBsCst*(1 - dK*0.01);
		break;
	}

	if(dSlsCst<0){
		m_Edit2=_T("1");
		AfxMessageBox(IDS_STRING9032, MB_ICONSTOP);
		GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
		GetDlgItem(IDC_STATIC_SLSCST)->SetWindowText(_T("0"));
	}
	else{
		CString s;
		s.Format(_T("%f"),dSlsCst);
		s.Replace('.',',');
		cSlsCst.ParseCurrency(s);
		m_stcSlsCst = cSlsCst.Format(0, MAKELCID(MAKELANGID(LANG_RUSSIAN,
					SUBLANG_SYS_DEFAULT), SORT_DEFAULT));
		GetDlgItem(IDC_STATIC_SLSCST)->SetWindowText(m_stcSlsCst);
	}
}

void CDlg::OnBnClickedButton2()
{// Посмотреть и выбрать цену
	if(IsEnableRec(ptrRs2)){
		HMODULE hMod;
		BOOL bFndC;
		bFndC   = TRUE;
		CString strFndC;
		strFndC = _T("");
		short i;
		COleVariant vC;

		i = 0;
		vC = GetValueRec(ptrRs2,i);
		vC.ChangeType(VT_BSTR);
		strFndC = vC.bstrVal;
		strFndC.TrimLeft();
		strFndC.TrimRight();

		hMod=AfxLoadLibrary(_T("VwBsCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL,CString&);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVwBsCst");
		if((func)(m_strNT, ptrCnn,strFndC,bFndC,m_strCst)){
			if(m_strCst.IsEmpty() || _wtof(m_strCst)==0){
				AfxFreeLibrary(hMod);
				return;
			}
			else{
				if(IsEnableRec(ptrRsCmb3)){
					OnFindInCombo(strFndC,&m_DataCombo3,ptrRsCmb3);
					if(!ptrRsCmb3->adoEOF){
						m_Edit1 = m_strCst;
						m_Edit1.TrimLeft();
						m_Edit1.TrimRight();
						m_Edit1.Replace(',','.');
						GetDlgItem(IDC_EDIT1)->SetWindowTextW(m_Edit1);
						//m_CdOp = _wtoi(strFndC);
						OnCalcSls((int)m_CdOp,m_Edit1,m_Edit2);
					}
				}
			}
		}
		AfxFreeLibrary(hMod);
	}
}
