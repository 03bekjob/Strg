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

	, m_iCdQr(0)
	, m_OleDateTime3(COleDateTime::GetCurrentTime())
	, m_OleDateTime4(COleDateTime::GetCurrentTime())
	, m_strSql(_T(""))
	, m_Edit7(_T(""))
	, m_Edit8(_T(""))
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
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER3, m_OleDateTime3);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER4, m_OleDateTime4);
	DDX_Text(pDX, IDC_EDIT7, m_Edit7);
	DDX_Text(pDX, IDC_EDIT8, m_Edit8);
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
	ON_EN_CHANGE(IDC_EDIT7, &CDlg::OnEnChangeEdit7)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"QNkl.dll"));

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
	if(!(m_EditTBCh.Create(WS_CHILD|
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
*/
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

	strSql = _T("QT68");
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
		m_DataCombo1.put_ListField(_T("Запрос"));
		SetTextCombo(m_DataCombo1,ptrRsCmb1,1);
//AfxMessageBox("ok");
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	OnShowQuery(m_iCdQr);
	OnShowGridQr(m_iCdQr);

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
//	UpdateData();
	CString s,sLstCh,str,msg,strCopy;
	BOOL bDec;
	bDec = FALSE;
	if(ptrRs1->State==adStateOpen){
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
				break;
			case adDecimal:		//14
			case adNumeric:	//131
			case adDouble:
			case adCurrency:
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
		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){

			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
				m_Flg = true;
			}
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

	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_EDIT7)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_EDIT8)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_DATETIME3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_DATETIME4)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
//	UpdateData();
	CString strSql;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

/*	if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
		strSql = _T("IT26 '");
		strSql+= m_Edit2 + _T("','");
		strSql+= m_Edit1 + _T("','");
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
		strSql = _T("");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			if(!m_Edit1.IsEmpty()){
				IndCol = 2;
				m_Flg = false;
				OnFindInGrid(m_Edit1,ptrRs1,IndCol,m_Flg);
				m_Flg = true;
			}
		}
	}
*/
	return;
}

void CDlg::On32776(void)
{
// Изменить
//	UpdateData();
	CString strSql,strC;
	_variant_t vra;
	COleVariant vC;
	VARIANT* vtl = NULL;
	short IndCol,i;

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
/*
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				strSql = _T("UT26 ");
				strSql+= strC	 + _T(",'");
				strSql+= m_Edit1 + _T("','");
				strSql+= m_Edit2 + _T("','");
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
    			strSql = _T("");
				OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
				if(IsEnableRec(ptrRs1)){
					if(!m_Edit1.IsEmpty()){
						IndCol = 0;
						m_Flg = false;
						OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
					}
				}
			}
*/
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
	if(IsEnableRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);
/*
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimLeft();

			strSql = _T("DT26 ");
			strSql+= strC;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
    		strSql = _T("");
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
*/
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

/*			m_CurCol = Grd.get_Col();
			if(m_CurCol==-1 || m_CurCol==EmpCol ){
				m_CurCol = DefCol;
			}
			m_iCurType = GetTypeCol(rs,m_CurCol);
*/

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
	if(!m_DataCombo1.get_BoundText().IsEmpty()){
		CString strIf,strBegDate,strEndDate,strCod;
		COleVariant vC;
		short i = 0;
		strCod = _T("0");

		if(IsEnableRec(ptrRs1)){
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCod = vC.bstrVal;
		}
		UpdateData();
//		strIf=GetStrIf(m_Radio1_3);
		vC=GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strIf=vC.bstrVal;

		if(m_Edit7.IsEmpty()) m_Edit7=_T("0");
		if(m_Edit8.IsEmpty()) m_Edit8=_T("0");

		strBegDate  = m_OleDateTime3.Format(_T("%Y-%m-%d"));
		strEndDate  = m_OleDateTime4.Format(_T("%Y-%m-%d"));

	//	strSql = _T("QT55 ");
		m_strSql+= strIf + _T(",'");
		m_strSql+= strBegDate + _T("','");
		m_strSql+= strEndDate + _T("',");
		m_strSql+= m_Edit7 + _T(",");
		m_strSql+= m_Edit8+ _T(",'");
		m_strSql+= _T("1',");
		m_strSql+= strCod;


	//	bstrSql = (_bstr_t) strSql;
		
		AfxSetResourceHandle(hInstResClnt);
		OnOK();
	}
	else{
		AfxMessageBox(IDS_STRING9022,MB_ICONINFORMATION);
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
				wdth = 90;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 1:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 260;
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

void CDlg::InitDataGrid1Op(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9021);
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
				wdth = 580;
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
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
				m_CurCol = m_DataGrid1.get_Col();
				if(IsQuery(_T("QT10"),ptrRs1)){
					if(m_CurCol==-1 ){
						m_CurCol=0;
					}
				}
				else{
					if(m_CurCol==-1 || m_CurCol==0){
						m_CurCol=1;
					}
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
				m_EditTBCh.SetTypeCol(m_iCurType);
		}
	}
/*	if(m_Flg){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				if(m_iBtSt==5 ){
						OnShowEdit();
				}
			}
		}
	}
*/
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

void CDlg::OnShowGridQr(int i)
{
	CString strSql;
	switch(i){
		case 3:
		case 4:
		case 5:
		case 8:
		case 9:
			strSql = _T("QT10");
			OnShowGrid(strSql,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			break;
		case 10:
			if(m_strSql.Find(_T("QT100N"))!=-1){
				strSql = _T("QT23_1");
			}
			else{
				strSql = _T("QT22_1");
			}
			OnShowGrid(strSql,ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1Op);
			break;
		default:
			m_DataGrid1.putref_DataSource(NULL);
			if(ptrRs1->State==adStateOpen) ptrRs1->Close();
			break;
	}

}

void CDlg::OnShowQuery(int i)
{
	switch(i){
	case 0:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(false);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(true);
		GetDlgItem(IDC_EDIT7)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(true);
		GetDlgItem(IDC_EDIT8)->ShowWindow(true);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(false);
		break;
	case 1:
	case 6:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(true);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(false);
		GetDlgItem(IDC_EDIT8)->ShowWindow(false);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(false);
		break;
	case 2:
	case 7:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(true);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(true);
		GetDlgItem(IDC_EDIT7)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(true);
		GetDlgItem(IDC_EDIT8)->ShowWindow(true);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(false);
		break;
	case 3:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(false);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(false);
		GetDlgItem(IDC_EDIT8)->ShowWindow(false);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(true);
		break;
	case 4:
	case 8:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(true);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(false);
		GetDlgItem(IDC_EDIT8)->ShowWindow(false);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(true);
		break;
	case 5:
	case 9:
	case 10:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(true);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(true);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(false);
		GetDlgItem(IDC_EDIT8)->ShowWindow(false);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(true);
		break;
	default:
		GetDlgItem(IDC_STATIC_DATETIME4)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER4)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_DATETIME3)->ShowWindow(false);
		GetDlgItem(IDC_DATETIMEPICKER3)->ShowWindow(false);

		GetDlgItem(IDC_STATIC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_EDIT7)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT8)->ShowWindow(false);
		GetDlgItem(IDC_EDIT8)->ShowWindow(false);

		GetDlgItem(IDC_DATAGRID1)->ShowWindow(false);
		break;
	}
}


void CDlg::ChangeDatacombo1()
{
	m_iCdQr = OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	OnShowQuery(m_iCdQr);
	OnShowGridQr(m_iCdQr);

}

void CDlg::OnEnChangeEdit7()
{
	GetDlgItem(IDC_EDIT7)->GetWindowTextW(m_Edit7);
	GetDlgItem(IDC_EDIT8)->SetWindowTextW(m_Edit7);
}
