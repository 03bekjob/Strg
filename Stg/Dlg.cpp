// Dlg.cpp : implementation file
//

#include "stdafx.h"
//#include "Stg.h"
#include "Dlg.h"
#include "../RDllDb/Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"

// CDlg dialog

IMPLEMENT_DYNAMIC(CDlg, CDialog)
CDlg::CDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDlg::IDD, pParent)
	, m_CurCol(0)
	, m_iBtSt(0)
	, m_Flg(TRUE)
	, m_strNT(_T(""))
	, m_bFnd(FALSE)
	, m_bFndC(FALSE)
	, m_strFndC(_T(""))
	, m_Check1(FALSE)
	, m_Edit1(_T(""))
	, m_Edit2(_T(""))
{
}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_COMMAND(ID_BUTTON32775, On32775)
	ON_COMMAND(ID_BUTTON32776, On32776)
	ON_COMMAND(ID_BUTTON32777, On32777)
	ON_COMMAND(ID_BUTTON32778, On32778)
	ON_COMMAND(ID_BUTTON32779, On32779)
	ON_COMMAND(ID_BUTTON32783, On32783)

	ON_EN_CHANGE(ID_EDIT_TOOLBAR, OnEnChangeEditTB)

	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
	ON_BN_CLICKED(IDC_OK, OnClickedOk)
	ON_BN_CLICKED(IDC_CANCEL, OnClickedCancel)
END_MESSAGE_MAP()


// CDlg message handlers

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

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"Stg.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING9014);
	this->SetWindowText(s);

//	m_Edit1Rs.SubclassDlgItem(IDC_EDIT1,this);

	InitStaticText();

	int iTBCtrlID;
//	short i;
	COleVariant vC;
//	m_Edit1Rus.SubclassDlgItem(IDC_EDIT1,this);

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
	OnShowGrid1();
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			if(m_bFndC){
				m_Flg = false;
//AfxMessageBox(m_strFndC);
				OnFindInGrid(m_strFndC,ptrRs1,0,m_Flg);
				m_Flg = true;
			}

		}
	}


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
		if(m_CurCol==-1 ){
			m_CurCol = 0;
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
		case adBSTR:
//AfxMessageBox("adChar");
			if(nChar==39){
				return;
			}
			break;
	}

//		s.Remove('\'');
		m_Flg = FALSE;
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
		m_Flg = TRUE;

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount;
	int count=0;
	int Id;
//	strMsg = _T("Вы забыли заполнить следующие поля :\n\t");
	strMsg.LoadString(IDS_STRING9016);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
				if(m_Edit1.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT2:
				if(m_Edit1.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
		}
		NextDlgCtrl();
	} while (Id!=IdEnd);
	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	short i;
	CString strC;
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 0;	//Код Склада
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

			i = 1;	//Наименвоание Склада
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit2 = vC.bstrVal;
			m_Edit2.TrimLeft();
			m_Edit2.TrimRight();
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

			i = 2;	//Признак выписки
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BOOL);
			m_Check1 = vC.boolVal;
			((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);
}

void CDlg::InitDataGrid1(void)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;

	strCap.LoadString(IDS_STRING9013);
	numRec = 0;

	GrdClms.AttachDispatch(m_DataGrid1.get_Columns());
	if(ptrRs1->State==adStateOpen){
		num = ptrRs1->GetadoFields()->GetCount();

		numRec = ptrRs1->GetRecordCount();
		strRec.Format(L"%i",numRec);
		strCap +=strRec;
		m_DataGrid1.put_Caption(strCap);

		for (i=0;i<num;i++) {
			switch(i) {
			case 0:
				GrdClms.GetItem((COleVariant) i).SetWidth(50);
				break;
			case 1:
				GrdClms.GetItem((COleVariant) i).SetWidth(300);
				break;
			case 2:
				GrdClms.GetItem((COleVariant) i).SetWidth(150);
				break;
			default:
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			}
		}
	}
	else{
		strRec.Format(L"%i",numRec);
		strCap +=strRec;
		m_DataGrid1.put_Caption(strCap);
	}

}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}


	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();


	if(OnOverEdit(IDC_EDIT1,IDC_EDIT2)){
		strSql = _T("IT4 ");
		strSql+= m_Edit1 + _T(",'");
		strSql+= m_Edit2 + _T("',");
		strSql+= m_Check1 ? _T("1"):_T("0");
		strSql+= _T(",'");
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
		OnShowGrid1();
		if(ptrRs1->State==adStateOpen){
			if(!IsEmptyRec(ptrRs1)){
				if(!m_Edit1.IsEmpty()){
					IndCol = 0;
					m_Flg = false;
					OnFindInGrid(m_Edit1,ptrRs1,IndCol,m_Flg);
					m_Flg = true;
				}
			}
		}
	}
	return;
}

void CDlg::On32776(void)
{
// Изменить
	UpdateData();
	CString strSql,strC;
	_variant_t vra;
	COleVariant vC;
	VARIANT* vtl = NULL;
	short IndCol,i;

	m_Edit2.TrimRight(' ');
	m_Edit2.TrimLeft(' ');
	if(m_Edit2==' ') m_Edit1.Empty();

	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT2)){
				strSql = _T("UT4 ");
				strSql+= strC	 + _T(",'");
				strSql+= m_Edit2 + _T("',");
				strSql+= m_Check1 ? _T("1"):_T("0");
				strSql+= _T(",'");
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
				OnShowGrid1();
				if(ptrRs1->State==adStateOpen){
					if(!IsEmptyRec(ptrRs1)){
						if(!m_Edit1.IsEmpty()){
							IndCol = 0;
							m_Flg = false;
							m_bFnd = OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
							m_Flg = true;
						}
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
	CString strSql,strC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);

			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimLeft();

			strSql = _T("DT4 ");
			strSql+= strC;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowGrid1();
			if(ptrRs1->State==adStateOpen){
				if(!IsEmptyRec(ptrRs1)){
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
		}
	}
	return;
}

void CDlg::On32779(void)
{
//Переключение Режимов
	m_iBtSt = m_wndToolBar.GetToolBarCtrl().GetState(ID_BUTTON32779);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	OnOnlyRead(m_iBtSt);
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

void CDlg::OnShowGrid1(void)
{
	ptrCmd1->CommandText = _T("QT4");
	try{
		m_DataGrid1.putref_DataSource(NULL);
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);
		
		m_CurCol = m_DataGrid1.get_Col();
//		s.Format("%i",m_CurCol);
//		AfxMessageBox(s);
		if(m_CurCol==-1  ){
			m_CurCol = 0;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		InitDataGrid1();
	}
	catch(_com_error& e){
		m_DataGrid1.putref_DataSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
}

void CDlg::OnOnlyRead(int i)
{
	if(i == 5){
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(TRUE);
	}
	else{
		((CEdit*)GetDlgItem(IDC_EDIT1))->SetReadOnly(FALSE);
	}
}
BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(ptrRs1->State==adStateOpen){
			if(!ptrRs1->adoEOF){
				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1 ){
					m_CurCol=0;
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
}
