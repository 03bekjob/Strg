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

static long m_lCdQr = 0;

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
//	, m_lCdQr(0)
	, m_strSql(_T(""))
{

}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
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
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"QPrcCst.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING9014);
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
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);
/*	strSql = _T("");
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	if(m_bFndC){
		short fCol=0;
		m_Flg = false;
		OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
		m_Flg = true;
	}
*/
	ptrCmdCmb1 = NULL;
	ptrRsCmb1  = NULL;

	strSql = _T("QT72");
	ptrCmdCmb1.CreateInstance(__uuidof(Command));
	ptrCmdCmb1->ActiveConnection = ptrCnn;
	ptrCmdCmb1->CommandType = adCmdText;
	ptrCmdCmb1->CommandText = (_bstr_t)strSql;

	ptrRsCmb1.CreateInstance(__uuidof(Recordset));
	ptrRsCmb1->CursorLocation = adUseClient;
	ptrRsCmb1->PutRefSource(ptrCmdCmb1);
	try{
		m_DataCombo1.putref_RowSource(NULL);
		ptrRsCmb1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		m_DataCombo1.putref_RowSource((LPUNKNOWN) ptrRsCmb1);
		m_DataCombo1.put_ListField(_T("Запрос"));
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

//		s.Remove('\'');
		m_Flg = false;
//AfxMessageBox("1");
//		AfxMessageBox(str);
		m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
//AfxMessageBox("2");
/*		if(m_bFnd){
			AfxMessageBox("true");
		}
		else{
			AfxMessageBox("false");
		}
*/
		m_Flg = true;

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

/*	GotoDlgCtrl(GetDlgItem(IdBeg));
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
*/
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
/*	COleVariant vC;
	short i;
	CString strC;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			try{
				i = 1;	//Код Группы
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

			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
		}
	}
*/
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_TIME1)->SetWindowText(sTxt);
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
	CString s;
	s = m_DataCombo1.get_BoundText();
	s.TrimLeft();
	s.TrimRight();
	if(!s.IsEmpty()){
		UpdateData();
		CString strR,strD,strC;
		short i;
		COleVariant vC;
		strR.Format(_T("%i"),m_lCdQr);
		strD = m_OleDateTime1.Format(_T("%Y-%m-%d"));
		if(m_lCdQr!=3){
			if(ptrRs1->State==adStateOpen){
				if(!IsEmptyRec(ptrRs1)){
					i = 0;
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;
					strC.TrimLeft();
					strC.TrimRight();

					m_strSql+= strR + _T(",");
					m_strSql+= strC + _T(",");
					m_strSql+= strC + _T(",'");
					m_strSql+= strD	+ _T("',");
					m_strSql+= strC;
					OnOK();
				}
				else{
					s.LoadString(IDS_STRING9022);
					AfxMessageBox(s,MB_ICONINFORMATION);
					OnCancel();
				}
			}
			else{
				s.LoadString(IDS_STRING9022);
				AfxMessageBox(s,MB_ICONINFORMATION);
				OnCancel();
			}
		}
		else{		// Запрос только по дате
			m_strSql+=strR +_T(",");
			m_strSql+=_T("0,0,'");
			m_strSql+=strD + _T("',0");
			OnOK();
		}
	}
	else{
		s.LoadString(IDS_STRING9021);
		AfxMessageBox(s,MB_ICONINFORMATION);
		OnCancel();
	}
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

	numRec = 0;
	switch(m_lCdQr){
		case 1:
		case 2:
			strCap.LoadString(IDS_STRING9013);
			break;
		case 4:
		case 5:
			strCap.LoadString(IDS_STRING9019);
			break;
		case 6:
		case 7:
			strCap.LoadString(IDS_STRING9020);
			break;
	}

	GrdClms.AttachDispatch(Grd.get_Columns());
	if(rs->State==adStateOpen){
		num = rs->GetadoFields()->GetCount();

		numRec = rs->GetRecordCount();
		strRec.Format(_T(" %i"),numRec);
		strCap +=strRec;
		Grd.put_Caption(strCap);
		switch(m_lCdQr){
			case 1:	//Группа
			case 2:
				for (i=0;i<num;i++) {
					switch(i) {
					case 0:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					case 1:
						GrdClms.GetItem((COleVariant) i).SetWidth(100);
						break;
					case 2:
						GrdClms.GetItem((COleVariant) i).SetWidth(300);
						break;
					default:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					}
				}
				break;
			case 4:	//Товар
			case 5:
				for (i=0;i<num;i++) {
					switch(i) {
					case 0:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					case 1:
						GrdClms.GetItem((COleVariant) i).SetWidth(100);
						break;
					case 2:
						GrdClms.GetItem((COleVariant) i).SetWidth(300);
						break;
					case 3:
						GrdClms.GetItem((COleVariant) i).SetWidth(100);
						break;
					default:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					}
				}
				break;
			case 6:
			case 7:
				for (i=0;i<num;i++) {
					switch(i) {
					case 0:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					case 1:
						GrdClms.GetItem((COleVariant) i).SetWidth(500);
						break;
					default:
						GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
						break;
					}
				}
				break;
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
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
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

void CDlg::ChangeDatacombo1()
{
	m_lCdQr = OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	OnShowQuery(m_lCdQr);
}

void CDlg::OnShowQuery(long R)
{
	CString strSql;
/*	CString strR;
	strR.Format("%i",R);
AfxMessageBox(_T("OnShowQuery  R = ")+strR);
*/
	switch(R){
	case 1:
		strSql = _T("T27_T26 1,1"); // Группа
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(false);
		break;
	case 2:
		strSql = _T("T27_T26 1,1"); // Группа+Дата
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(true);
		break;
	case 3:							// Дата
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(true);
		break;
	case 4:
		strSql=	_T("QT12");			// Товар
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(false);
		break;
	case 5:
		strSql=	_T("QT12");			// Товар+Дата
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(true);
		break;
	case 6:							// Производитель
		strSql = _T("QT125");		
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(false);
		break;
	case 7:							// Производитель+Дата
		strSql = _T("QT125");		
		GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_TIME1)->ShowWindow(true);
		break;
	}
	if(R!=3){
		OnShowGrid(strSql,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	}
	else{
		if(ptrRs1->State==adStateOpen){
			m_DataGrid1.putref_DataSource(NULL);
			ptrRs1->Close();
//AfxMessageBox("close");
		}
	}
}
