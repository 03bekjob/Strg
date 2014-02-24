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

	, m_Edit1(_T(""))
	, m_Edit2(_T(""))
	, m_Edit3(_T(""))
	, m_OleDateTime1(COleDateTime::GetCurrentTime())
{

}

CDlg::~CDlg()
{
	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
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
	ON_BN_CLICKED(IDC_BUTTON1, &CDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"AddDoc.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING21002);
	this->SetWindowText(s);

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
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
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
	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrCmdCmb1 = NULL;
	ptrRsCmb1 = NULL;
	ptrCmdCmb1.CreateInstance(__uuidof(Command));
	ptrCmdCmb1->ActiveConnection = ptrCnn;
	ptrCmdCmb1->CommandType = adCmdText;

	strSql = _T("QT45");
	ptrCmdCmb1->CommandText = (_bstr_t)strSql;

	ptrRsCmb1.CreateInstance(__uuidof(Recordset));
	ptrRsCmb1->CursorLocation = adUseClient;
	ptrRsCmb1->PutRefSource(ptrCmdCmb1);
	try{
		m_DataCombo1.putref_RowSource(NULL);
		ptrRsCmb1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		m_DataCombo1.putref_RowSource((LPUNKNOWN) ptrRsCmb1);
		m_DataCombo1.put_ListField(_T("Наименование"));
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
*/
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);

}

void CDlg::OnEnChangeEditTB(void)
{
/*	UpdateData();
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

		m_CurCol = m_DataGrid1.get_Col();
		if(m_CurCol==-1 || m_CurCol==0){
			m_CurCol = 1;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);

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

		m_Flg = false;
		m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
		m_Flg = true;

	}
*/
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount,strC;
	COleVariant vC;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

	BOOL bC3,bC4,bC5;

	short i = 2;
	if(IsEnableRec(ptrRsCmb1)){
		vC=GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC3=vC.boolVal;
		i = 3;
		vC=GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC4=vC.boolVal;
		i = 4;
		vC=GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC5=vC.boolVal;
	}
	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
				if(m_Edit1.IsEmpty() && bC3){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT2:
				if(m_Edit2.IsEmpty() && bC4){
					count++;
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT3:
				if(m_Edit3.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT3)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
		}
		NextDlgCtrl();
	} while (Id!=IdEnd);

//Раздел Combo
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
	if(IsEmptyRec(ptrRsCmb1)){
		count++;
		strCount.Format(_T("%i"),count);
		GetDlgItem(IDC_STATIC_COMBO1)->GetWindowText(strItem);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}
	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING21003);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING21004);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING21005);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING21006);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING21007);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
	return;
}

void CDlg::On32776(void)
{
// Изменить
	return;
}

void CDlg::On32777(void)
{
//Удалить
	return;
}

void CDlg::On32779(void)
{
//Переключение Режимов

}


void CDlg::On32778(void)
{
	return;
}

void CDlg::On32783(void)
{
	return;
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
	UpdateData();

	COleVariant vC;
	CString strSql,strC,strC3,strDate1;
	BOOL bC3,bC4,bC5;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i=0;

	strDate1 = m_OleDateTime1.Format(_T("%Y-%m-%d"));
	if(OnOverEdit(IDC_EDIT1,IDC_EDIT3)){

		if(IsEnableRec(ptrRsCmb1)){

			i=0;
			vC=GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BSTR);
			strC3=vC.bstrVal;

			i = 2;
			vC=GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BOOL);
			bC3=vC.boolVal;
			i = 3;
			vC=GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BOOL);
			bC4=vC.boolVal;
			i = 4;
			vC=GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BOOL);
			bC5=vC.boolVal;		
		}
		if(bC3 && bC4 && bC5){

		}
		else if(!bC3 && bC4 && bC5){
			m_Edit1 = _T("-");
			
		}
		else if(!bC3 && !bC4 && bC5){
			m_Edit1 = _T("-");
			m_Edit2 = _T("-");
			strDate1 = _T("1900-01-01");
		}
		else if(!bC3 && !bC4 && !bC5){
			m_Edit1 = _T("-");
			m_Edit2 = _T("-");
			strDate1 = _T("1900-01-01");
		}
		else if(bC3 && !bC4 && !bC5){

			m_Edit2 = _T("-");
			strDate1 = _T("1900-01-01");
		}
		else if(bC3 && bC4 && !bC5){
			strDate1 = _T("1900-01-01");
		}
		else if(!bC3 && bC4 && !bC5){
			m_Edit1 = _T("-");
			strDate1 = _T("1900-01-01");
		}


//		strC = vC.bstrVal;

		strSql = _T("IT46 ");
		strSql+= m_strCod +_T(",");
		strSql+= strC3    + _T(",'");
		strSql+= m_Edit1  + _T("','");
		strSql+= m_Edit2  + _T("','");
		strSql+= strDate1 + _T("','");
		strSql+= m_Edit3  + _T("','");
		strSql+= m_strNT  +_T("'");
//AfxMessageBox(strSql);

		ptrCmd1->CommandText = (_bstr_t)strSql;
		try{
//AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
			OnOK();
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		
	}
}

void CDlg::OnClickedCancel()
{
	OnCancel();
}


BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
END_EVENTSINK_MAP()


void CDlg::OnOnlyRead(int i)
{
	if(i == 5){
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(TRUE);
	}
	else{
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
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

void CDlg::OnBnClickedButton1()
{
//Вид документов
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strIns;
		bFndC   = TRUE;
		strFndC = _T("");
		COleVariant vC;
		short i;
		if(IsEnableRec(ptrRsCmb1)){
			COleVariant vC;
			i = 0;
			vC = GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();
		}
		hMod=AfxLoadLibrary(_T("VidDoc.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidDoc");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);
		try{
			ptrRsCmb1->Requery(adCmdText);
			m_DataCombo1.Refresh();
			OnFindInCombo(strFndC,&m_DataCombo1,ptrRsCmb1,0,1);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::ChangeDatacombo1()
{
	short i;
	i = 1;
	m_cd = OnChangeCombo(m_DataCombo1,ptrRsCmb1,i);
	OnShowEnter();
}

void CDlg::OnShowEnter(void)
{
	COleVariant vC;
	BOOL bC3,bC4,bC5;
	short i= 2;
	if(IsEnableRec(ptrRsCmb1)){
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC3 = vC.boolVal;

		i = 3;
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC4 = vC.boolVal;

		i = 4;
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BOOL);
		bC5 = vC.boolVal;
	//	AfxMessageBox(bC3);
		if(bC3 && bC4 && bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_SHOW);

		}	
		else if(bC3 && !bC4 && bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_SHOW);
		}
		else if(!bC3 && bC4 && bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_SHOW);

		}
		else if(!bC3 && !bC4 && bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_SHOW);
		}
		else if(!bC3 && !bC4 && !bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_HIDE);
		}
		else if(bC3 && !bC4 && !bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);

			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_HIDE);
		}
		else if(bC3 && bC4 && !bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
			
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_HIDE);
		}
		else if(!bC3 && bC4 && !bC5){
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);

			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(SW_SHOW);
			GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);

			GetDlgItem(IDC_STATIC_DATETIME1)->ShowWindow(SW_HIDE);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(SW_HIDE);
		}
	}
}
