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
	, m_OleDateTime2(COleDateTime::GetCurrentTime())
{

}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_OleDateTime2);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
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
	AfxSetResourceHandle(::GetModuleHandle(L"Crtf.dll"));

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
	strSql = L"QT103";
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
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
	strSql = L"QT12_1";
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
void CDlg::OnEnableButtonBar(int iBtSt, CToolBar* pTlBr)
{	
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
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);

}

void CDlg::OnEnChangeEditTB(void)
{
	UpdateData();
	CString s,sLstCh,str,msg;
	UINT nChar;
	char* ch;
	if(IsEnableRec(ptrRs1)){
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
//AfxMessageBox("EditTB 2");
			m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
//AfxMessageBox("EditTB _2");
			m_Flg = true;
		}

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
				GetDlgItem(IDC_EDIT1)->GetWindowTextW(m_Edit1);
				m_Edit1.TrimRight(' ');
				m_Edit1.TrimLeft(' ');
				if(m_Edit1==' ') m_Edit1.Empty();
				if(m_Edit1.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT2:
				GetDlgItem(IDC_EDIT2)->GetWindowTextW(m_Edit2);
				m_Edit2.TrimRight(' ');
				m_Edit2.TrimLeft(' ');
				if(m_Edit2==' ') m_Edit2.Empty();
				if(m_Edit2.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT3:
				GetDlgItem(IDC_EDIT3)->GetWindowTextW(m_Edit3);
				m_Edit3.TrimRight(' ');
				m_Edit3.TrimLeft(' ');
				if(m_Edit3==' ') m_Edit3.Empty();
				if(m_Edit3.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT3)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
		}
		NextDlgCtrl();
	} while (Id!=IdEnd);
//Раздел Grid
	if(m_iBtSt==5 ){	//Режим Корректировки
		//Товар
		if(ptrRs2->State==adStateOpen){
			if(IsEmptyRec(ptrRs2)){
				count++;
				strItem.LoadString(IDS_STRING9018);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
			count++;
			strItem.LoadString(IDS_STRING9018);
			strCount.Format(L"%i",count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{		//Режим Добавить/Удалить 
			//Товар
		if(ptrRs2->State==adStateOpen){
			if(IsEmptyRec(ptrRs2)){
				count++;
				strItem.LoadString(IDS_STRING9018);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
				count++;
				strItem.LoadString(IDS_STRING9018);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
		}
	}

	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	short i,IndCol;
	CString strF;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i = 6;	//Дата выдачи
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF = vC.bstrVal;
			m_OleDateTime1.ParseDateTime(strF);

			i = 7;	//Действителен до ...
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF = vC.bstrVal;
			m_OleDateTime2.ParseDateTime(strF);

			UpdateData(FALSE);

			i=0; // Код Товара
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			IndCol = 0;
			m_Flg = false;
			OnFindInGrid(strF,ptrRs2,IndCol,m_Flg);
			m_Flg = true;

			i = 4;	//№ сертиф
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
			
			i = 5;	//Учёт. №
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit2 = vC.bstrVal;
			m_Edit2.TrimLeft();
			m_Edit2.TrimRight();
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

			i = 8;	//Кем Выдан
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit3 = vC.bstrVal;
			m_Edit3.TrimLeft();
			m_Edit3.TrimRight();
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_STATIC_DATETIME2)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql;
	CString strCT,strDC4,strDC5;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i,IndCol;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	strDC4 = m_OleDateTime1.Format(L"%Y-%m-%d");
	strDC5 = m_OleDateTime2.Format(L"%Y-%m-%d");

	if(OnOverEdit(IDC_EDIT1,IDC_EDIT3)){
		i = 0;		// Товар
		vC = GetValueRec(ptrRs2,i);
		vC.ChangeType(VT_BSTR);
		strCT = vC.bstrVal;
		strCT.TrimLeft();
		strCT.TrimRight();

		strSql = _T("IT103 ");
		strSql+= strCT    + _T(",'");	// Код товара
		strSql+= m_Edit1  + _T("','");	// №
		strSql+= m_Edit2  + _T("','");	// Уч. №
		strSql+= strDC4	  + _T("','");	// Дата выд
		strSql+= strDC5	  + _T("','");	// Дейст до
		strSql+= m_Edit3  + _T("','");	// Кем выдан
		strSql+= m_strNT+ _T("'");

		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//			AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		strSql = _T("QT103");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			IndCol = 0; //по Товару
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
	CString strSql,strCT,strDC4,strDC5;
	_variant_t vra;
	COleVariant vC,vBk;
	VARIANT* vtl = NULL;
	short IndCol,i;

	strDC4 = m_OleDateTime1.Format(L"%Y-%m-%d");
	strDC5 = m_OleDateTime2.Format(L"%Y-%m-%d");

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){

				i=0;	// Код Товара
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCT=vC.bstrVal;
				strCT.TrimLeft();
				strCT.TrimRight();

				vBk = ptrRs1->GetBookmark();


			if(OnOverEdit(IDC_EDIT1,IDC_EDIT3)){
				strSql = _T("UT103 ");
				strSql+= strCT    + _T(",'");	// Код товара
				strSql+= m_Edit1  + _T("','");	// №
				strSql+= m_Edit2  + _T("','");	// Уч. №
				strSql+= strDC4	  + _T("','");	// Дата выд
				strSql+= strDC5	  + _T("','");	// Дейст до
				strSql+= m_Edit3  + _T("','");	// Кем выдан
				strSql+= m_strNT+ _T("'");

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				try{
		//			AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
    			strSql = _T("QT103");
				OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				if(IsEnableRec(ptrRs1)){
						IndCol = 0;
						m_Flg = false;
						OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
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
	CString strSql,strC1;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	if(IsEnableRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);

			i=0;	//по Товару
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC1=vC.bstrVal;
			strC1.TrimLeft();
			strC1.TrimRight();

			strSql = _T("DT103 ");
			strSql+= strC1;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
    		strSql = _T("QT103");
			OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
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
	if(m_iBtSt==5){
		OnShowEdit();
	}


}


void CDlg::On32778(void)
{
	return;
}

void CDlg::On32783(void)
{
	return;
}

void CDlg::OnShowGrid(CString strSql,_RecordsetPtr& rs,_CommandPtr& Cmd,CDatagrid1& Grd,void (*pFGrd)(CDatagrid1&,_RecordsetPtr&))
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
			if(m_CurCol==-1 || m_CurCol==0 ){
				m_CurCol = 1;
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
	OnOK();
}

void CDlg::OnClickedCancel()
{
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 170;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 130;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 7:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 8:
				wdth = 180;
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
				wdth = 170;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 13:
				wdth = 50;
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

void CDlg::ClickDatagrid1()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
}

void CDlg::ClickDatagrid2()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID2);
}
