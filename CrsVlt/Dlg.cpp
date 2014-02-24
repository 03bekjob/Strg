// Dlg.cpp : implementation file
//

#include "stdafx.h"
//#include "CrsVlt.h"
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
	, m_OleDateTime1(COleDateTime::GetCurrentTime())
	, m_Edit1(_T(""))
{
}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
	ptrRsCmb2 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
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
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnBnClickedButton2)
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
BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, ChangeDatacombo2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, ChangeDatacombo1, VTS_NONE)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1 || m_CurCol==0){
					m_CurCol=1;
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
				m_EditTBCh.SetTypeCol(m_iCurType);
		}
	}
	if(m_Flg){
		if(ptrRs1->State==adStateOpen){
			if(!ptrRs1->adoEOF){
				if(m_iBtSt==5 ){
//					if(m_bFnd){
						OnShowEdit();
//					}
				}
			}
		}
	}
}

void CDlg::OnBnClickedButton1()
{
	BeginWaitCursor();

		HMODULE hMod;
		hMod=AfxLoadLibrary(L"TabVlt.dll");
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startTabVlt");
		(func)(m_strNT, ptrCnn);
		try{
			ptrRsCmb1->Requery(adCmdText);
			m_DataCombo1.Refresh();
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::OnBnClickedButton2()
{
	BeginWaitCursor();

		HMODULE hMod;
		hMod=AfxLoadLibrary(L"VidCrs.dll");
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidCrs");
		(func)(m_strNT, ptrCnn);
		try{
			ptrRsCmb2->Requery(adCmdText);
			m_DataCombo2.Refresh();
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::ChangeDatacombo2()
{	
	OnChangeCombo(m_DataCombo2,ptrRsCmb2,1);
}

void CDlg::ChangeDatacombo1()
{
	OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
}

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"CrsVlt.dll"));

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

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);
	OnShowGrid1();
	m_CurCol = m_DataGrid1.get_Col();
	if(m_CurCol==-1 || m_CurCol==0){
		m_CurCol=1;
	}
	m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
	m_EditTBCh.SetTypeCol(m_iCurType);

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
		if(!str.IsEmpty()){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
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
//	strMsg = _T("Вы забыли заполнить следующие поля :\n\t");
	strMsg.LoadString(IDS_STRING9016);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
				m_Edit1.Replace(',','.');
				if(m_Edit1.IsEmpty() || atof((_bstr_t)m_Edit1/*(LPSTR) * m_Edit1.GetBufferSetLength(m_Edit1.GetLength()) * * m_Edit1.GetBuffer(m_Edit1.GetLength()) *  */)==0){
					count++;
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
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
		strCount.Format(L"%i",count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	strC = m_DataCombo2.get_BoundText();

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO2)->GetWindowText(strItem);
		strCount.Format(L"%i",count);
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
	CString strC,strC2;
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 1;	//Data
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			m_OleDateTime1.ParseDateTime(strC);
			UpdateData(FALSE);


			i = 4;	//Курс
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

			i = 5; // Код валюты
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			OnFindInCombo(strC,&m_DataCombo1,ptrRsCmb1,0,1);

			i = 6; // Код вида курса
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			OnFindInCombo(strC,&m_DataCombo2,ptrRsCmb2,0,1);

		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_DATETIME1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);
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
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			case 1:
				GrdClms.GetItem((COleVariant) i).SetWidth(70);
				break;
			case 2:
				GrdClms.GetItem((COleVariant) i).SetWidth(70);
				break;
			case 3:
				GrdClms.GetItem((COleVariant) i).SetWidth(200);
				break;
			case 4:
				GrdClms.GetItem((COleVariant) i).SetWidth(180);
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
	CString strSql,strD,strC3,strC4;
	_variant_t vra;
	VARIANT* vtl = NULL;
	COleVariant vC;
	short IndCol,i;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}


	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();
	m_Edit1.Replace(',','.');

	strD = m_OleDateTime1.Format(L"%Y-%m-%d");

	if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
		i = 0;
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strC3 = vC.bstrVal;
		strC3.TrimLeft();
		strC3.TrimRight();

		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strC4 = vC.bstrVal;
		strC4.TrimLeft();
		strC4.TrimRight();

		strSql = _T("IT21 '");
		strSql+= strD	 + _T("',");
		strSql+= strC3	 + _T(",");
		strSql+= strC4	 + _T(",");
		strSql+= m_Edit1 + _T(",'");
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
/*					IndCol = 4;
					m_Flg = false;
					OnFindInGrid(m_Edit1,ptrRs1,IndCol,m_Flg);
					m_Flg = true;
*/
					m_Flg = false;
					ptrRs1->MoveFirst();
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
	CString strSql,strC,strC3,strC4,strD;
	_variant_t vra;
	COleVariant vC;
	VARIANT* vtl = NULL;
	short IndCol,i;

	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();
	m_Edit1.Replace(',','.');

	strD = m_OleDateTime1.Format(L"%Y-%m-%d");
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				i = 0;
				vC = GetValueRec(ptrRsCmb1,i);
				vC.ChangeType(VT_BSTR);
				strC3 = vC.bstrVal;
				strC3.TrimLeft();
				strC3.TrimRight();

				vC = GetValueRec(ptrRsCmb2,i);
				vC.ChangeType(VT_BSTR);
				strC4 = vC.bstrVal;
				strC4.TrimLeft();
				strC4.TrimRight();

				strSql = _T("UT21 ");
				strSql+= strC	 + _T(",'");
				strSql+= strD	 + _T("',");
				strSql+= strC3	 + _T(",");
				strSql+= strC4	 + _T(",");
				strSql+= m_Edit1 + _T(",'");
				strSql+= m_strNT + _T("'");
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

			strSql = _T("DT21 ");
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
	ptrCmd1->CommandText = _T("QT21");
	try{
		m_DataGrid1.putref_DataSource(NULL);
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);
		
/*		m_CurCol = m_DataGrid1.get_Col();
		if(m_CurCol==-1 || m_CurCol==0 ){
			m_CurCol = 1;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
*/
		InitDataGrid1();
	}
	catch(_com_error& e){
		m_DataGrid1.putref_DataSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
}
