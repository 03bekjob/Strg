// Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg.h"

#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"

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
	, m_OleDateTime2(COleDateTime::GetCurrentTime())
	, m_Radio1(0)
	, m_Check1(FALSE)
	, m_Edit1(_T("0"))
	, m_strBskt(_T(""))
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

	if(ptrRs4->State==adStateOpen) ptrRs4->Close();
	ptrRs4 = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATAGRID4, m_DataGrid4);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_OleDateTime2);
	DDX_Radio(pDX, IDC_RADIO1, m_Radio1);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_COMMAND(ID_BUTTON32771, On32771)
	ON_COMMAND(ID_BUTTON32772, On32772)
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
	ON_BN_CLICKED(IDC_RADIO1, &CDlg::OnBnClickedRadio1)
	ON_BN_CLICKED(IDC_RADIO2, &CDlg::OnBnClickedRadio2)
	ON_BN_CLICKED(IDC_CHECK1, &CDlg::OnBnClickedCheck1)
	ON_BN_CLICKED(IDC_BUTTON1, &CDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"SlsDscAg.dll"));

	CString s,strSql,strSql2,strC;
	s.LoadString(IDS_STRING9014);
	this->SetWindowText(s);

	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);			

	m_Edit1Flt.SubclassDlgItem(IDC_EDIT1,this);

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
//	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
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
	strSql = _T("QT90_1 '");
	strSql+= m_strNT + _T("'");
	m_strBskt = strSql;
	OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);

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
	strSql = _T("QT10");
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2,0,0);

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
			strSql = _T("T12_T27_1 ");
			strSql+= s + _T(",0");
			break;
		case 1:
			if(IsEnableRec(ptrRs4)){
				i = 0;
				vC = GetValueRec(ptrRs4,i);
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
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);

	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);

	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			if(ptrRs3->State==adStateOpen){
				i = 0;
				vC = GetValueRec(ptrRs3,i);
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
	OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);

	ptrCmdCmb1 = NULL;
	ptrRsCmb1 = NULL;

	strSql = _T("QT78");
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
		m_DataCombo1.put_ListField(_T("Вид скидки"));
//AfxMessageBox("ok");
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
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
	if(strSrc.Find(_T("QT89"))!=-1){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
	}
	else{
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,TRUE);
		if(iBtSt==4){
			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);

			pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,TRUE);
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
			if(m_CurCol==-1 /*|| m_CurCol==0*/){
				m_CurCol = 0;
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

		if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
			m_CurCol = m_DataGrid4.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs4,m_CurCol);
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

//		s.Remove('\'');
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

		if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs4,m_CurCol,m_Flg);
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
				if(m_Edit1.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
		}
		NextDlgCtrl();
	} while (Id!=IdEnd);


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

	//Раздел Grid
	if(ptrRs2->State==adStateOpen){
		if(IsEmptyRec(ptrRs2)){
			count++;
			strItem.LoadString(IDS_STRING9029);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{
		count++;
		strItem.LoadString(IDS_STRING9029);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	if(ptrRs4->State==adStateOpen){
		if(IsEmptyRec(ptrRs4)){
			count++;
			strItem.LoadString(IDS_STRING9030);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{
		count++;
		strItem.LoadString(IDS_STRING9030);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	if(ptrRs3->State==adStateOpen){
		if(IsEmptyRec(ptrRs3)){
			count++;
			strItem.LoadString(IDS_STRING9031);
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{
		count++;
		strItem.LoadString(IDS_STRING9031);
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
			try{
/*				i = 1;	//Код Группы
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				m_Edit1 = vC.bstrVal;
				m_Edit1.TrimLeft();
				m_Edit1.TrimRight();
				GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

				i = 2;	//Наименование
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				m_Edit2 = vC.bstrVal;
				m_Edit2.TrimLeft();
				m_Edit2.TrimRight();
				GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
*/
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9026);
	GetDlgItem(IDC_STATIC_GROUP1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9027);
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_BUTTON1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_GROUP2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_RADIO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_RADIO2)->SetWindowText(sTxt);

}

void CDlg::On32771(void)
{
//вернуться в корзину
	OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
	SetIcon(m_Icon,TRUE);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	return;
}

void CDlg::On32772(void)
{
//слить в базу
	COleVariant vC;
	COleDateTime vD;
	CString strSql,strCA,strD,strDD,strMsg;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	if(IsEnableRec(ptrRs1)){
		if(IsQuery(_T("QT90_1"),ptrRs1)){
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCA = vC.bstrVal;
			strCA.TrimLeft();
			strCA.TrimRight();

			i = 8;	//Дата
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));
			strDD= vD.Format(_T("%d.%m.%Y"));

			strSql = _T("IT90T89ALL_1 '"); 
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
			OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			if(!IsEmptyRec(ptrRs1)){
				AfxFormatString2(strMsg,IDS_STRING9032,strCA,strDD);
				AfxMessageBox(strMsg,MB_ICONSTOP);
			}
		}
	}
	return;
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql,strR,strC,strG,strT,strDs,strDt;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol,i;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}
	strR.Format(_T("%i"),m_Radio1);
	strC = strG = strT = _T("0");
	strDt = m_OleDateTime1.Format(_T("%Y-%m-%d"));	

	if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){

		i = 0;	//Код контрАг
		vC = GetValueRec(ptrRs2,i);
		vC.ChangeType(VT_BSTR);
		strC = vC.bstrVal;
		strC.TrimLeft();
		strC.TrimRight();

		i = 0;	//Код товара
		vC = GetValueRec(ptrRs3,i);
		vC.ChangeType(VT_BSTR);
		strT = vC.bstrVal;
		strT.TrimLeft();
		strT.TrimRight();

		i = 0;	//Код группы
		vC = GetValueRec(ptrRs4,i);
		vC.ChangeType(VT_BSTR);
		strG = vC.bstrVal;
		strG.TrimLeft();
		strG.TrimRight();

		i = 0; //КОд вида скидки
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strDs = vC.bstrVal;
		strDs.TrimLeft();
		strDs.TrimRight();

		strSql = _T("IT90_1 ");
		strSql+= strR	 + _T(",");
		strSql+= strC    + _T(",");
		strSql+= strG	 + _T(",");
		strSql+= strT    + _T(",");
		strSql+= strDs   + _T(",");
		strSql+= m_Edit1 + _T(",'");
		strSql+= strDt   + _T("','");
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
		OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			if(!m_Edit1.IsEmpty()){
				IndCol = 2;
				m_Flg = false;
				OnFindInGrid(m_Edit1,ptrRs1,IndCol,m_Flg);
				m_Flg = true;
			}
		}
	}

	return;
}

void CDlg::On32776(void)
{
// Изменить
	UpdateData();
	CString strSql,strCA,strR,strDN,strDO,strC7,strCT,strC6,strC6N;
	_variant_t vra;
	COleVariant vC;
	COleDateTime vD;
	VARIANT* vtl = NULL;
	short IndCol,i;

	strR.Format(_T("%i"),m_Radio1);
	strDN = m_OleDateTime1.Format(_T("%Y-%m-%d"));

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				i = 0;	//Код контрагента
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCA = vC.bstrVal;
				strCA.TrimRight();
				strCA.TrimRight();

				i = 7;	//% скидки
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC7 = vC.bstrVal;
				strC7.TrimLeft();
				strC7.TrimRight();
				strC7.Replace(',','.');

				i = 8; //Дата
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_DATE);
				vD = vC.date;
				strDO = vD.Format(_T("%Y-%m-%d"));

				i = 9; //Код товара
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCT = vC.bstrVal;
				strCT.TrimLeft();
				strCT.TrimRight();

				i = 10; //Код вида скидки 
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC6 = vC.bstrVal;
				strC6.TrimLeft();
				strC6.TrimRight();

				i = 0; //Код вида скидки нов
				vC = GetValueRec(ptrRsCmb1,i);
				vC.ChangeType(VT_BSTR);
				strC6N = vC.bstrVal;
				strC6N.TrimLeft();
				strC6N.TrimRight();

				strSql = _T("UT90_1 ");
				strSql+= strR    + _T(",");
				strSql+= strCA	 + _T(",");
				strSql+= strCT   + _T(",");
				strSql+= strC6   + _T(",");
				strSql+= strC7   + _T(",'");
				strSql+= strDO   + _T("',");
				strSql+= strC6N  + _T(",");
				strSql+= m_Edit1 + _T(",'");
				strSql+= strDN   + _T("','");
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
				OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				if(IsEnableRec(ptrRs1)){
					if(!m_Edit1.IsEmpty()){
						IndCol = 9;
						m_Flg = false;
						OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
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
	COleVariant vC,vBk;
	COleDateTime vD;
	CString strSql,strCA,strR,strC7,strD,strCT;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;

	strR.Format(_T("%i"),m_Radio1);
	strC7 = strCA = strCT = _T("0");

	if(IsEnableRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);

			i = 0;	//Код контрагента
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCA = vC.bstrVal;
			strCA.TrimRight();
			strCA.TrimLeft();

			i = 7;	//% скидки
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC7 = vC.bstrVal;
			strC7.TrimLeft();
			strC7.TrimRight();
			strC7.Replace(',','.');

			i = 8; //Дата
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_DATE);
			vD = vC.date;
			strD = vD.Format(_T("%Y-%m-%d"));

			i = 9; //Код товара
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCT = vC.bstrVal;

			strSql = _T("DT90_1 ");
			strSql+= strR   + _T(",");
			strSql+= strCA  + _T(",");
			strSql+= strC7  + _T(",'");
			strSql+= strD   + _T("',");
			strSql+= strCT  + _T(",'");
			strSql+= m_strNT+ _T("'");

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
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
	CString strSql,strX,strHd;
	HMODULE hMod;
	hMod=AfxLoadLibrary(L"QSqlX.dll");
	typedef BOOL (*pDialog)(CString& strSql,_ConnectionPtr, CString strX, CString strHd );
	pDialog func=(pDialog)GetProcAddress(hMod,"startQSqlX");
	strSql = _T("QT89_1 ");
	strX = _T("32841");
	strHd  = _T("Запрос на Скидки КонтрАгентов");
	if((func)(strSql, ptrCnn, strX, strHd)){
		BeginWaitCursor();
			OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1F);
			if(IsQuery(_T("QT89_1"),ptrRs1)){
				if(IsEnableRec(ptrRs1)){
					m_Icon = AfxGetApp()->LoadIcon(IDI_ICON2);		// Прочитать икону
					SetIcon(m_Icon,TRUE);			
				}
				else{
					m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
					SetIcon(m_Icon,TRUE);			
					OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				}
			}
			else{
				m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
				SetIcon(m_Icon,TRUE);			
				OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			}
	}
	else{
		m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
		SetIcon(m_Icon,TRUE);			
		OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	}
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);


//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	return;
}

void CDlg::On32783(void)
{//печать
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);;
				break;
			case 1:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 7:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 8:
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
void CDlg::InitDataGrid1F(CDatagrid1& Grd,_RecordsetPtr& rs)
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 7:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 8:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 9:
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

	strCap.LoadString(IDS_STRING9017);
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 1:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 200;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 150;
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 200;
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

void CDlg::InitDataGrid4(CDatagrid1& Grd,_RecordsetPtr& rs)
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
				wdth = 150;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 150;
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
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID4, DISPID_CLICK, CDlg::ClickDatagrid4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, 218, CDlg::RowColChangeDatagrid3, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID4, 218, CDlg::RowColChangeDatagrid4, VTS_PVARIANT VTS_I2)
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

void CDlg::OnOnlyRead(int i)
{
/*	if(i == 5){
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(TRUE);
	}
	else{
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
	}
*/
}

void CDlg::SetTextCombo(CDatacombo2& Cmb, _RecordsetPtr pRs, short iTxt)
{
/*	if(IsEnableRec(pRs)){
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
*/
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

void CDlg::ChangeDatacombo1()
{
	OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
}

void CDlg::OnBnClickedRadio1()
{
	UpdateData();
	COleVariant vC;
	CString strSql,s,strC;
	short i;
	switch(m_Radio1){
		case 0:
			s = m_Radio1==0 ? _T("0"):_T("1");
			switch(m_Radio1){
				case 0:
					strSql = _T("T12_T27_1 ");
					strSql+= s + _T(",0");
					break;
				case 1:
					if(IsEnableRec(ptrRs4)){
						i = 0;
						vC = GetValueRec(ptrRs4,i);
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
			OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
		case 1:
			s = m_Radio1==0 ? _T("0"):_T("1");
			switch(m_Radio1){
				case 0:
					if(IsEnableRec(ptrRs3)){
						i = 0;
						vC = GetValueRec(ptrRs3,i);
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
			OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
			break;
	}
	if(m_iBtSt == 5){
		OnShowEdit();
	}
}

void CDlg::OnBnClickedRadio2()
{
	UpdateData();
	COleVariant vC;
	CString strSql,s,strC;
	short i;
	switch(m_Radio1){
		case 0:
			s = m_Radio1==0 ? _T("0"):_T("1");
			switch(m_Radio1){
				case 0:
					strSql = _T("T12_T27_1 ");
					strSql+= s + _T(",0");
					break;
				case 1:
					if(IsEnableRec(ptrRs4)){
						i = 0;
						vC = GetValueRec(ptrRs4,i);
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
			OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
			break;
		case 1:
			s = m_Radio1==0 ? _T("0"):_T("1");
			switch(m_Radio1){
				case 0:
					if(IsEnableRec(ptrRs3)){
						i = 0;
						vC = GetValueRec(ptrRs3,i);
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
			OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
			break;
	}
	if(m_iBtSt == 5){
		OnShowEdit();
	}
}

void CDlg::RowColChangeDatagrid3(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(IsEnableRec(ptrRs3)){
			if(!ptrRs3->adoEOF){
				CString s,strSql,strC;
				COleVariant vC;
				short i;
				switch(m_Radio1){
					case 0:
						s = m_Radio1==0 ? _T("0"):_T("1");
						switch(m_Radio1){
							case 0:
								if(IsEnableRec(ptrRs3)){
									i = 0;
									vC = GetValueRec(ptrRs3,i);
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
						OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
						break;
				}
			}
		}
	}
	m_Flg = true;
}

void CDlg::RowColChangeDatagrid4(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(IsEnableRec(ptrRs4)){
			if(!ptrRs4->adoEOF){
				CString s,strSql,strC;
				COleVariant vC;
				short i;
				switch(m_Radio1){
					case 1:
						s = m_Radio1==0 ? _T("0"):_T("1");
						switch(m_Radio1){
							case 0:
								strSql = _T("T12_T27_1 ");
								strSql+= s + _T(",0");
								break;
							case 1:
								if(IsEnableRec(ptrRs4)){
									i = 0;
									vC = GetValueRec(ptrRs4,i);
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
						OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
						break;
				}
			}
		}
	}
	m_Flg = true;
}

void CDlg::OnBnClickedCheck1()
{
	if(IsEnableRec(ptrRs1)){
		if(IsQuery(_T("QT89_1"),ptrRs1)){
			UpdateData();
			GetDlgItem(IDC_DATETIMEPICKER2)->ShowWindow(m_Check1);
			GetDlgItem(IDC_BUTTON1)->ShowWindow(m_Check1);
		}
		else{
			AfxMessageBox(IDS_STRING9033,MB_ICONINFORMATION);
			UpdateData(false);
		}
	}
	else{
		AfxMessageBox(IDS_STRING9033,MB_ICONINFORMATION);
		UpdateData(false);
	}
}

void CDlg::OnBnClickedButton1()
{
// Выполнить. Занашу шаблон в Корзину Инд. Скидок с новой датой
	UpdateData();
	COleVariant vC;
	COleDateTime vD;
	CString strSql,strC2,strC8D,strC8DU;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	BeginWaitCursor();
	if(IsQuery(_T("QT89_1"),ptrRs1)){
		//Дата шаблона (новая дата)
		strC8DU = m_OleDateTime1.Format(_T("%Y-%m-%d"));

		i=1;	//Код КонтрАгента
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strC2 = vC.bstrVal;

		i = 9; //Дата ввода исходного шаблона (старая дата)
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_DATE);
		vD = vC.date;
		strC8D = vD.Format(_T("%Y-%m-%d"));

		strSql = _T("IT90Stmp_1 ");
		strSql+= strC2   + _T(",'");
		strSql+= strC8D  + _T("','");
		strSql+= strC8DU + _T("','");
		strSql+= m_strNT + _T("'");
		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//			AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
			return;
		}
		OnShowGrid(m_strBskt, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsQuery(_T("QT90_1"),ptrRs1)){
			m_Icon = AfxGetApp()->LoadIcon(IDI_ICON1);		// Прочитать икону
			SetIcon(m_Icon,TRUE);
		}
	}
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	EndWaitCursor();
}
