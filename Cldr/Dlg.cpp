// Dlg.cpp : implementation file
//

#include "stdafx.h"
//#include "Cldr.h"
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
	, m_Flg(true)
	, m_bFnd(FALSE)
	, m_strNT(_T(""))
	, m_Edit1(_T(""))
	, m_Edit2(_T(""))
	, m_Edit3(_T(""))
	, m_Edit4(_T(""))
	, m_bCombo2(true)
	, m_bCombo3(true)
{
}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRsMax->State==adStateOpen) ptrRsMax->Close();
	ptrRsMax = NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
	ptrRsCmb2 = NULL;

	if(ptrRsCmb3->State==adStateOpen) ptrRsCmb3->Close();
	ptrRsCmb3 = NULL;

	if(ptrRsCmb4->State==adStateOpen) ptrRsCmb4->Close();
	ptrRsCmb4 = NULL;

	if(ptrRsCmb5->State==adStateOpen) ptrRsCmb5->Close();
	ptrRsCmb5 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_Control(pDX, IDC_DATACOMBO3, m_DataCombo3);
	DDX_Control(pDX, IDC_DATACOMBO4, m_DataCombo4);
	DDX_Control(pDX, IDC_DATACOMBO5, m_DataCombo5);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_Text(pDX, IDC_EDIT4, m_Edit4);
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

BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, ChangeDatacombo2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO3, 1, ChangeDatacombo3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO4, 1, ChangeDatacombo4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO5, 1, ChangeDatacombo5, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
END_EVENTSINK_MAP()

void CDlg::ChangeDatacombo1()
{
	long cd = 0;
	cd = OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	OnShowComboGlb(cd);
//	AfxMessageBox("start OnChangeDatacombo2");
	if(cd!=10){
		OnUpdateCombo2();
		OnUpdateCombo3();	//Делаем города
							// Меняем Терр-ые ед-цы России
							// и становимся на первую запись 
							// выборки через SetBoundText()
		cd  = OnShowCombo();
		UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
	}
	else{
//AfxMessageBox("if Glb ==10");			
		UpdateGrid1Glb();
	}
}

void CDlg::ChangeDatacombo2()
{
//	AfxMessageBox("start OnChangeDatacombo4");
	long cd;
	OnChangeCombo(m_DataCombo2, ptrRsCmb2,1);
//	OnShowCladr(ptrRs1);
		if(m_bCombo2){
	//	AfxMessageBox("invoid OnUpdateCombo3 from OnChangeDatacombo4");
			OnUpdateCombo3();
			cd=OnShowCombo();
			UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
		}
//	OnShowCladr(ptrRs1);
}

void CDlg::ChangeDatacombo3()
{
//AfxMessageBox("start OnChangeDatacombo3");
	long cd;
	cd = OnChangeCombo(m_DataCombo3, ptrRsCmb3,1);
	if(m_bCombo3){
		cd=OnShowCombo();
		if(10!=cd){
			OnUpdateCombo4();
			OnUpdateGrid1Rgn();
//			UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Rgn '"),_T("QT8TwnMax '"));
//			OnChangeCombo5();
		}
		else{
//AfxMessageBox("1---");
//			OnUpdateCombo2();
//			OnUpdateCombo3();
			UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
		}
	}
}

void CDlg::ChangeDatacombo4()
{
	long cd;
	cd = OnChangeCombo(m_DataCombo4, ptrRsCmb4,1);
//	if(m_iBtSt==5){
//		OnShowCladr(ptrRsCmb4);
//	}
	OnShowCladrRgn();
	ChangeDatacombo5();
}

void CDlg::ChangeDatacombo5()
{
	long cd;
	cd = OnChangeCombo(m_DataCombo5, ptrRsCmb5,1);
	OnShowCombo();
	UpdateGrid1(cd,ptrRsCmb4,ptrRsCmb5,_T("QT8TwnVlg '"),_T("QT8TwnVlgMax '"));
	OnUpdateCombo5(cd);
}

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
	if(ptrRs1->State==adStateOpen){
		if(!ptrRs1->adoEOF){
			if(m_iBtSt==5 ){
				OnShowEdit();
			}
		}
	}
}

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"Cldr.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING9013);
	this->SetWindowText(s);

//	m_Edit1Rs.SubclassDlgItem(IDC_EDIT1,this);
	((CEdit*)GetDlgItem(IDC_EDIT2))->SetReadOnly();
	m_Edit1RusNum.SubclassDlgItem(IDC_EDIT1,this);

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
//	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
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

	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);

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

	ptrCmdMax = NULL;
	ptrRsMax = NULL;

	ptrCmdMax.CreateInstance(__uuidof(Command));
	ptrCmdMax->ActiveConnection = ptrCnn;
	ptrCmdMax->CommandType = adCmdText;

	ptrRsMax.CreateInstance(__uuidof(Recordset));
	ptrRsMax->CursorLocation = adUseClient;
	ptrRsMax->PutRefSource(ptrCmdMax);

	ptrCmdCmb1 = NULL;
	ptrRsCmb1  = NULL;

	strSql = _T("QT32");
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
		m_DataCombo1.put_ListField(_T("Сокр. наим."));
//AfxMessageBox("ok");
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb2 = NULL;
	ptrRsCmb2  = NULL;

	strSql = _T("QT8 0");
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
		m_DataCombo2.put_ListField(_T("Наименование"));
	}
	catch(_com_error& e){
		m_DataCombo2.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb3 = NULL;
	ptrRsCmb3  = NULL;

	strSql = _T("QT33");
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
		m_DataCombo3.put_ListField(_T("Сокр. наим."));
	}
	catch(_com_error& e){
		m_DataCombo3.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb4 = NULL;
	ptrRsCmb4  = NULL;

	strSql = _T("QT8Rgn ' ', 0");
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
		m_DataCombo4.put_ListField(_T("Наименование"));
	}
	catch(_com_error& e){
		m_DataCombo4.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb5 = NULL;
	ptrRsCmb5  = NULL;

	strSql = _T("QT7");
	ptrCmdCmb5.CreateInstance(__uuidof(Command));
	ptrCmdCmb5->ActiveConnection = ptrCnn;
	ptrCmdCmb5->CommandType = adCmdText;
	ptrCmdCmb5->CommandText = (_bstr_t)strSql;

	ptrRsCmb5.CreateInstance(__uuidof(Recordset));
	ptrRsCmb5->CursorLocation = adUseClient;
	ptrRsCmb5->PutRefSource(ptrCmdCmb5);
	try{
		m_DataCombo5.putref_RowSource(NULL);
		ptrRsCmb5->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo5.putref_RowSource((LPUNKNOWN) ptrRsCmb5);
		m_DataCombo5.put_ListField(_T("Сокращение"));
	}
	catch(_com_error& e){
		m_DataCombo5.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}


	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CDlg::OnEnableButtonBar(int iBtSt, CToolBar* pTlBr)
{	
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
	if(iBtSt==4){
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);

		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,TRUE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,TRUE);
		OnUnLockCombo();


	}
	else{
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,TRUE);

		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
		pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
		OnLockCombo();

	}

}

void CDlg::OnEnChangeEditTB(void)
{
	UpdateData();
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
	CString strC,strItem,strMsg,strCount,strC3;
	int count=0;
	int Id;
	short i;
	COleVariant vC;
	long cd;
//	strMsg = _T("Вы забыли заполнить следующие поля :\n\t");
	strMsg.LoadString(IDS_STRING9016);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
				if(m_Edit1.IsEmpty()){
					count++;
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					GetDlgItem(IDC_STATIC_EDIT1)->GetWindowText(strItem);
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT2:
				if(m_Edit2.IsEmpty()){
					count++;
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT3:
				if(m_Edit3.IsEmpty()){
					count++;
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					GetDlgItem(IDC_STATIC_EDIT3)->GetWindowText(strItem);
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT4:
				if(m_Edit4.IsEmpty()){
					count++;
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					GetDlgItem(IDC_STATIC_EDIT4)->GetWindowText(strItem);
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

	i = 0;
	vC = GetValueRec(ptrRsCmb1,i);
	vC.ChangeType(VT_I4);
	cd = vC.lVal;
	
	if(cd!=10){

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

		i = 0;
		vC = GetValueRec(ptrRsCmb3,i);
		vC.ChangeType(VT_I4);
		cd = vC.lVal;
		if(cd!=10){

			strC = m_DataCombo3.get_BoundText();

			strC.TrimRight(' ');
			strC.TrimLeft(' ');
			if(strC==' ') strC.Empty();
			if(strC.IsEmpty()){
				count++;
				GetDlgItem(IDC_STATIC_COMBO3)->GetWindowText(strItem);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}

			strC = m_DataCombo4.get_BoundText();

			strC.TrimRight(' ');
			strC.TrimLeft(' ');
			if(strC==' ') strC.Empty();
			if(strC.IsEmpty()){
				count++;
				GetDlgItem(IDC_STATIC_COMBO4)->GetWindowText(strItem);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}

			strC = m_DataCombo5.get_BoundText();

			strC.TrimRight(' ');
			strC.TrimLeft(' ');
			if(strC==' ') strC.Empty();
			if(strC.IsEmpty()){
				count++;
				GetDlgItem(IDC_STATIC_COMBO5)->GetWindowText(strItem);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
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
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){

			i = 1;	//Наименование
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

			i=6;			//C8	-----------------	Кладр
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2 = strC;

			i = 7;			// C9
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2+=_T(" ")+strC;

			i = 8;			// C10
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2+=_T(" ")+strC;

			i = 9;			// C11	-----------------	Кладр
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2+=_T(" ")+strC;

			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

			i = 4;	//Индекс
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit3 = vC.bstrVal;
			m_Edit3.TrimLeft();
			m_Edit3.TrimRight();
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

			i = 5;	//ГНИ
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit4 = vC.bstrVal;
			m_Edit4.TrimLeft();
			m_Edit4.TrimRight();
			GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9014);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_COMBO3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_COMBO4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_COMBO5)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_EDIT4)->SetWindowText(sTxt);

}

void CDlg::InitDataGrid1(void)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;

	strCap.LoadString(IDS_STRING9021);
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
				GrdClms.GetItem((COleVariant) i).SetWidth(340);
				break;
/*	case 2:
				GrdClms.GetItem((COleVariant) i).SetWidth(500);
				break;
*/
			case 3:
				GrdClms.GetItem((COleVariant) i).SetWidth(70);
				break;
			case 4:
				GrdClms.GetItem((COleVariant) i).SetWidth(50);
				break;
			case 5:
				GrdClms.GetItem((COleVariant) i).SetWidth(50);
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
	CString strSql,strC3;
	_bstr_t bstrSqlOld;
	COleVariant vC;
   	_variant_t vra;
	VARIANT* vtl = NULL;

	short IndCol,i;
	long cd;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();

	i = 0;
	vC = GetValueRec(ptrRsCmb3,i);
	vC.ChangeType(VT_I4);
	cd = vC.lVal;

	switch(cd){
		case 10:
			vC.ChangeType(VT_BSTR);
			strC3 = vC.bstrVal;
			break;
		default:
			if(ptrRsCmb5->State==adStateOpen){
				if(!IsEmptyRec(ptrRsCmb5)){
					i = 0;
					vC = GetValueRec(ptrRsCmb5,i);
					vC.ChangeType(VT_BSTR);
					strC3 = vC.bstrVal;
				}
			}
			break;
	}

	m_Edit2.Remove(' ');

	m_Edit3.TrimLeft();
	m_Edit3.TrimRight();

	m_Edit4.TrimLeft();
	m_Edit4.TrimRight();
	if(OnOverEdit(IDC_EDIT1,IDC_OK)){
		strSql =_T("IT8 '");
		strSql+=m_Edit1+_T("',");
		strSql+=strC3+_T(",'");
		strSql+=m_Edit2+_T("','");
		strSql+=m_Edit3+_T("','");
		strSql+=m_Edit4+_T("',");
		strSql+=_T("'0'");
		strSql+=_T(",'");
		strSql+=m_strNT+_T("'");

		bstrSqlOld = ptrCmd1->GetCommandText();

		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
			if(cd!=10) ChangeDatacombo5();
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

//		ptrCmd1->CommandText = bstrSqlOld;
		try{
//			ptrRs1->Requery(adCmdText);
//			m_DataGrid1.Refresh();
			if(cd==10){
				UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
			}
			if(!m_Edit1.IsEmpty()){
				IndCol = 3;
				m_Flg = false;
				OnFindInGrid(m_Edit2,ptrRs1,IndCol,m_Flg);
				m_Flg = true;
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		InitDataGrid1();
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
	_bstr_t bstrSqlOld;
	short IndCol,i;

	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();

	m_Edit2.Remove(' ');

	m_Edit3.TrimRight(' ');
	m_Edit3.TrimLeft(' ');
	if(m_Edit3==' ') m_Edit3.Empty();

	m_Edit4.TrimRight(' ');
	m_Edit4.TrimLeft(' ');
	if(m_Edit4==' ') m_Edit4.Empty();

	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				strSql = _T("UT8_1 '");
				strSql+= m_Edit2 + _T("','");
				strSql+= m_Edit1 + _T("','");
				strSql+= m_Edit3 + _T("','");
				strSql+= m_Edit4 + _T("','");
				strSql+= m_strNT + _T("'");

				bstrSqlOld = ptrCmd1->GetCommandText();

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				try{
		//			AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				ptrCmd1->CommandText = bstrSqlOld;
				try{
					ptrRs1->Requery(adCmdText);
					m_DataGrid1.Refresh();
		//			InitDataGrid1();
					if(!m_Edit1.IsEmpty()){
						IndCol = 0;
						m_Flg = false;
						OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
					}
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				InitDataGrid1();
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
	_bstr_t bstrSqlOld;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	long cd;
	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){

			ptrRs1->get_Bookmark(vBk);

			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimLeft();

			strSql = _T("DT8 ");
			strSql+= strC;

			bstrSqlOld = ptrCmd1->GetCommandText();

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			if(ptrRsCmb3->State==adStateOpen){
				if(!IsEmptyRec(ptrRsCmb3)){
					i = 0;
					vC = GetValueRec(ptrRsCmb3,i);
					vC.ChangeType(VT_I4);
					cd = vC.lVal;
				}
			}
			switch(cd){
				case 10:
					UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
					break;
				default:
					cd = OnChangeCombo(m_DataCombo5,ptrRsCmb5,1);
					OnUpdateCombo5(cd);
					break;
			}

			ptrCmd1->CommandText = bstrSqlOld;
			try{
				ptrRs1->Requery(adCmdText);
				m_DataGrid1.Refresh();
				if(cd!=10) 	ChangeDatacombo5();
				try{
					m_Flg = false;
					ptrRs1->PutBookmark(vBk);
					m_Flg = true;
					InitDataGrid1();
				}
				catch(...){
					if(!IsEmptyRec(ptrRs1)){
						ptrRs1->MoveLast();
						InitDataGrid1();
					}
					else{
						m_DataGrid1.putref_DataSource(NULL);
					}
				}
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
		}
	}
	return;
}

void CDlg::On32779(void)
{
//Переключение Режимов
	long cd;
	short i;
	COleVariant vC;
//	CString s;
	m_iBtSt = m_wndToolBar.GetToolBarCtrl().GetState(ID_BUTTON32779);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	if(m_iBtSt==5){
		OnShowEdit();
	}
	else{
		if(ptrRsCmb3->State==adStateOpen){
			if(!IsEmptyRec(ptrRsCmb3)){
				i = 0;
				vC = GetValueRec(ptrRsCmb3,i);
				vC.ChangeType(VT_I4);
				cd = vC.lVal;
			}
		}
		switch(cd){
			case 10:
				UpdateGrid1(cd,ptrRsCmb2,ptrRsCmb3,_T("QT8Twn '"),_T("QT8TwnMax '"));
				break;
			default:
				cd = OnChangeCombo(m_DataCombo5,ptrRsCmb5,1);
				OnUpdateCombo5(cd);
				break;
		}

	}
}


void CDlg::On32778(void)
{
//Выбрать
	BeginWaitCursor();
//AfxMessageBox("inv");
//		m_SlpDay.SetDate(t1);
		CString str,strSql;
		HMODULE hMod;
		hMod=AfxLoadLibrary(L"Ftch.dll");
		typedef BOOL (*pDialog)(CString&);
		pDialog func=(pDialog)GetProcAddress(hMod,"startFtch");
		(func)(str);
		if(!str.IsEmpty()){
			strSql = _T("QT8C2 '");
			strSql+= str + _T("'");
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				if(ptrRs1->State==adStateOpen) ptrRs1->Close();
				ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
				m_DataGrid1.putref_DataSource(NULL);
				m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);
				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1 || m_CurCol==0){
					m_CurCol = 1;
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
				InitDataGrid1();
			}
			catch(_com_error& e){
				m_DataGrid1.putref_DataSource(NULL);
				AfxMessageBox(e.ErrorMessage());
			}
		}

		AfxFreeLibrary(hMod);
//		OnShowIndex();
	EndWaitCursor();
	return;
}

void CDlg::On32783(void)
{
	return;
}

void CDlg::OnShowComboGlb(long cd)
{
/*CString s;
s.Format("%i",cd);
AfxMessageBox(s);
*/
	if(cd==10){
		m_DataCombo2.ShowWindow(false);
		m_DataCombo3.ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO2)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO3)->ShowWindow(false);
		m_DataCombo4.ShowWindow(false);
		m_DataCombo5.ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO4)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO5)->ShowWindow(false);

	}
	else{
		m_DataCombo2.ShowWindow(true);
		m_DataCombo3.ShowWindow(true);
		GetDlgItem(IDC_STATIC_COMBO2)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_COMBO3)->ShowWindow(true);
	}
}

void CDlg::OnUpdateCombo2(void)
{
//AfxMessageBox("OnUpdateCombo2");
	CString strCur,strSql;
	CString strC;
	COleVariant vC;
	if(ptrRsCmb1->adoEOF) return;
	short i = 0;
	vC=GetValueRec(ptrRsCmb1,i);
	vC.ChangeType(VT_BSTR);
	strCur=vC.bstrVal;
//		AfxMessageBox(strCur);
	strSql = _T("QT8 ");
	strSql+= strCur;
	ptrCmdCmb2->CommandType = adCmdText;
	ptrCmdCmb2->CommandText = (_bstr_t)strSql;
	try{
		m_DataCombo2.putref_RowSource(NULL);
		if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
		ptrRsCmb2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo2.putref_RowSource((LPUNKNOWN) ptrRsCmb2);
		m_DataCombo2.put_ListField(_T("Наименование"));
		ptrRsCmb2->MoveFirst();
		i = 1;
		vC = GetValueRec(ptrRsCmb2,i);
		vC.ChangeType(VT_BSTR);
		strCur = vC.bstrVal;
//AfxMessageBox(strCur);
		m_DataCombo2.put_BoundText(strCur);
//		OnShowCladr(ptrRsCmb2);
	}
	catch(_com_error& e){
		m_DataCombo2.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
}

void CDlg::OnUpdateCombo3(void)
{
//Делаем города
//	AfxMessageBox("OnUpdateCombo3");
	CString strCur;
	COleVariant vC;
	short i = 1;
	if(ptrRsCmb3->State==adStateOpen){
		if(!IsEmptyRec(ptrRsCmb3)){
			ptrRsCmb3->MoveFirst();
			vC=GetValueRec(ptrRsCmb3,i);
			vC.ChangeType(VT_BSTR);
			strCur=vC.bstrVal;
			m_bCombo3 = false;
				m_DataCombo3.put_BoundText(strCur);
			m_bCombo3 = true;
		}
	}
}

long CDlg::OnShowCombo(void)
{
//	AfxMessageBox("start OnShowCombo");
	COleVariant vC;
	CString strCur,strCur1,str;
	long cd;
	short i = 0;
	vC=GetValueRec(ptrRsCmb3,i);
	vC.ChangeType(VT_I4);
	cd=vC.intVal;
	if(cd==10){
		m_DataCombo4.ShowWindow(false);
		m_DataCombo5.ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO4)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBO5)->ShowWindow(false);
	}
	else{
		m_DataCombo4.ShowWindow(true);
		m_DataCombo5.ShowWindow(true);
		GetDlgItem(IDC_STATIC_COMBO4)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_COMBO5)->ShowWindow(true);
	}
	return cd;
}

void CDlg::UpdateGrid1(long i,_RecordsetPtr rsCmb1,_RecordsetPtr rsCmb2,CString strSql,CString strSqlMax)
{
	COleVariant vC;
	CString strCur,strCur1,str;//,strNumRec,strSql;
//	long NumRec=0;

	short iC;
		iC = 3;
		vC=GetValueRec(rsCmb1,iC);
		vC.ChangeType(VT_BSTR);
		strCur=vC.bstrVal;
		str = strCur.Left(5);
//	AfxMessageBox(strCur);

		if(rsCmb2->adoEOF) return;

		iC = 0;
		vC=GetValueRec(rsCmb2,iC);
		vC.ChangeType(VT_BSTR);
		strCur1=vC.bstrVal;

//		strSql = _T("QT8Twn '");
		strSql+= str;
		strSql+= _T("',");
		strSql+= strCur1;
		ptrCmd1->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
		try{
			m_DataGrid1.putref_DataSource(NULL);
			if(ptrRs1->State==adStateOpen) ptrRs1->Close();
			ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
			m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);

			m_CurCol = m_DataGrid1.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol=1;
			}
			m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
			m_EditTBCh.SetTypeCol(m_iCurType);

			InitDataGrid1();
		}
		catch(_com_error& e){
			m_DataGrid1.putref_DataSource(NULL);
			if(ptrRs1->State==adStateOpen) ptrRs1->Close();
			AfxMessageBox(e.ErrorMessage());
		}
//		strSql = _T("QT8TwnVlgMax '");
		strSqlMax+= str;
		strSqlMax+= _T("',");
		strSqlMax+= strCur1;
        ptrCmdMax->CommandText = (_bstr_t)strSqlMax;
//AfxMessageBox(strSqlMax);
		if(ptrRs1->State==adStateOpen){
			try{
				if(ptrRsMax->State==adStateOpen) ptrRsMax->Close();
				ptrRsMax->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
				OnShowCladr(ptrRs1,i);
			}
			catch(_com_error& e){
				AfxMessageBox(e.ErrorMessage());
			}
		}

		return;
}

void CDlg::UpdateGrid1Glb(void)
{
	CString strCur1,str,strNumRec,strSql;
	long NumRec=0;
	strCur1	=_T("10");

	strSql = _T("QT8TwnGlb ");
	strSql+= strCur1;
	ptrCmd1->CommandText = (_bstr_t)strSql;
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataGrid1.putref_DataSource(NULL);
		m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);

		m_CurCol = m_DataGrid1.get_Col();
		if(m_CurCol==-1 || m_CurCol==0){
			m_CurCol=1;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		m_EditTBCh.SetTypeCol(m_iCurType);

		InitDataGrid1();
	}
	catch(_com_error& e){
		m_DataGrid1.putref_DataSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
	return;
}


void CDlg::OnUpdateCombo4(void)
{
//	AfxMessageBox("start OnUpdateCombo5");
	COleVariant vC;
	CString strCur,strCur1,str,strSql;
	long NumRec =0;

	if(ptrRsCmb1->adoEOF || ptrRsCmb2->adoEOF) return;

	short i = 3;		//Респ, Края и т.д.
	vC=GetValueRec(ptrRsCmb2,i);
	vC.ChangeType(VT_BSTR);
	strCur=vC.bstrVal;
	str = strCur.Left(2);
//AfxMessageBox(str);

	if(ptrRsCmb3->adoEOF) return;

	i = 0;				//Районы, города
	vC=GetValueRec(ptrRsCmb3,i);
	vC.ChangeType(VT_BSTR);
	strCur1=vC.bstrVal;

	strSql = _T("QT8Rgn '");
	strSql+= str;
	strSql+= _T("', ");
	strSql+= strCur1;

//AfxMessageBox(_T("Combo4 =")+strSql);

	ptrCmdCmb4->CommandType = adCmdText;
	ptrCmdCmb4->CommandText = (_bstr_t)strSql;
	try{
		m_DataCombo4.putref_RowSource(NULL);
		if(ptrRsCmb4->State==adStateOpen) ptrRsCmb4->Close();
		ptrRsCmb4->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo4.putref_RowSource((LPUNKNOWN) ptrRsCmb4);
		m_DataCombo4.put_ListField(_T("Наименование"));
		ptrRsCmb4->MoveFirst();
		i = 1;
		vC = GetValueRec(ptrRsCmb4,i);
		vC.ChangeType(VT_BSTR);
		strCur = vC.bstrVal;
//AfxMessageBox(strCur);
		m_DataCombo4.put_BoundText(strCur);
	}
	catch(_com_error& e){
		m_DataCombo4.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	return;
}

void CDlg::OnShowCladr(_RecordsetPtr ptrRs,long cd)
{
	COleVariant vC;
	CString strC,strCN,strC10,strC11;
	short i,iN;
	strCN = _T("000");
	iN=0;
	if(ptrRs->State==adStateOpen){
		if(!IsEmptyRec(ptrRs)){
			i=6;		//C8
			vC = GetValueRec(ptrRs,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2 = strC;

			i = 7;			// C9
			vC = GetValueRec(ptrRs,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2+=_T(" ")+strC;

			i = 8;			// C10
			vC = GetValueRec(ptrRs,i);
			vC.ChangeType(VT_BSTR);
			strC10 = vC.bstrVal;
			strC10.TrimLeft();
			strC10.TrimRight();
			m_Edit2+=_T(" ")+strC10;

			i = 9;			// C11
			vC = GetValueRec(ptrRs,i);
			vC.ChangeType(VT_BSTR);
			strC11 = vC.bstrVal;
			strC11.TrimLeft();
			strC11.TrimRight();
			m_Edit2+=_T(" ")+strC11;

			if(ptrRsMax->State==adStateOpen){
				if(!IsEmptyRec(ptrRsMax)){
					try{
						i = 0;
						vC = GetValueRec(ptrRsMax,i);
						vC.ChangeType(VT_I2);
						iN = vC.iVal;
						iN++;
						strC.Format(L"%i",iN);
						strC.TrimLeft();
						strC.TrimRight();
						if(iN<=9){
							strCN = _T("00")+strC;
						}
						else if(iN>=10 && iN<=99){
							strCN = _T("0")+strC;
						}
						else{
							strCN = strC;
						}
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
				}
				switch(cd){
					case 10:
//AfxMessageBox("strC10 = "+strC10);
						m_Edit2.Delete(7,3);
						m_Edit2.Insert(7,strCN);
						break;
					default:
//AfxMessageBox("strC11 = "+strC11);
						m_Edit2.Delete(11,3);
						m_Edit2.Insert(11,strCN);
						break;
				}

			}
//AfxMessageBox(strCN);
		}
	}
	else{
		m_Edit2.Empty();
	}
	GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
}

void CDlg::OnUpdateGrid1Rgn(void)
{
	COleVariant vC;
	CString strCur,strCur1,str,strSql;

	if(ptrRsCmb1->adoEOF || ptrRsCmb2->adoEOF) return;

	short i = 3;		//Респ, Края и т.д.
	vC=GetValueRec(ptrRsCmb2,i);
	vC.ChangeType(VT_BSTR);
	strCur=vC.bstrVal;
	str = strCur.Left(2);
//AfxMessageBox(str);

	if(ptrRsCmb3->adoEOF) return;

	i = 0;				//Районы, города
	vC=GetValueRec(ptrRsCmb3,i);
	vC.ChangeType(VT_BSTR);
	strCur1=vC.bstrVal;

	strSql = _T("QT8Rgn '");
	strSql+= str;
	strSql+= _T("', ");
	strSql+= strCur1;
//AfxMessageBox(strSql);
	ptrCmd1->CommandText = (_bstr_t)strSql;
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataGrid1.putref_DataSource(NULL);
		m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);

		m_CurCol = m_DataGrid1.get_Col();
		if(m_CurCol==-1 || m_CurCol==0){
			m_CurCol=1;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		m_EditTBCh.SetTypeCol(m_iCurType);

		InitDataGrid1();
		OnShowCladrRgn();
	}
	catch(_com_error& e){
		m_DataGrid1.putref_DataSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	return;
}

void CDlg::OnShowCladrRgn(void)
{
	COleVariant vC;
	CString strC,strC10,strC11;
	short i;
	if(ptrRsCmb4->State==adStateOpen){
		if(!IsEmptyRec(ptrRsCmb4)){
			i=6;		//C8
			vC = GetValueRec(ptrRsCmb4,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2 = strC;

			i = 7;			// C9
			vC = GetValueRec(ptrRsCmb4,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimLeft();
			strC.TrimRight();
			m_Edit2+=_T(" ")+strC;

			i = 8;			// C10
			vC = GetValueRec(ptrRsCmb4,i);
			vC.ChangeType(VT_BSTR);
			strC10 = vC.bstrVal;
			strC10.TrimLeft();
			strC10.TrimRight();
			m_Edit2+=_T(" ")+strC10;

			i = 9;			// C11
			vC = GetValueRec(ptrRsCmb4,i);
			vC.ChangeType(VT_BSTR);
			strC11 = vC.bstrVal;
			strC11.TrimLeft();
			strC11.TrimRight();
			m_Edit2+=_T(" ")+strC11;

		}
		else{
			m_Edit2.Empty();
		}
	}
	GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);
}


void CDlg::OnUpdateCombo5(long cd)
{
	if(ptrRs1->State==adStateOpen){
		if(IsEmptyRec(ptrRs1)){ //Если пусто значит заводиттся новый город, посёлок... и т.д.
			if(ptrRsCmb4->State==adStateOpen){
				if(!IsEmptyRec(ptrRsCmb4)){

					CString strSqlMax,str,strCur,strCur1,strCur2;
					COleVariant vC;
					short i;

					if(ptrRsCmb1->adoEOF || ptrRsCmb2->adoEOF) return;

					i = 3;		//Респ, Края и т.д.
					vC=GetValueRec(ptrRsCmb2,i);
					vC.ChangeType(VT_BSTR);
					strCur=vC.bstrVal;
					str = strCur.Left(2);
				//AfxMessageBox(str);

					if(ptrRsCmb3->adoEOF) return;

					strCur1.Format(L"%i",cd);

					if(ptrRsCmb4->adoEOF) return;

					i = 7;	// Район
					vC=GetValueRec(ptrRsCmb4,i);
					vC.ChangeType(VT_BSTR);
					strCur2=vC.bstrVal;
//AfxMessageBox(strCur2);

					strSqlMax = _T("QT8RgnMax '");
					strSqlMax+= str		+ _T("','");
					strSqlMax+= strCur2 + _T("',");
					strSqlMax+= strCur1;
					ptrCmdMax->CommandText = (_bstr_t)strSqlMax;
//AfxMessageBox(strSqlMax);
					try{
						if(ptrRsMax->State==adStateOpen) ptrRsMax->Close();
						ptrRsMax->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
						OnShowCladr(ptrRsCmb4,cd);
					}
					catch(_com_error& e){
						AfxMessageBox(e.ErrorMessage());
					}
				}
			}
		}
		else{ //не пкстой Grid1
			UpdateGrid1(cd,ptrRsCmb4,ptrRsCmb5,_T("QT8TwnVlg '"),_T("QT8TwnVlgMax '"));
		}
	}

/*	if(ptrRsCmb5->State==adStateOpen){
		if(!IsEmptyRec(ptrRsCmb5)){
			short i;
			i = 0;
			OnFindInCombo(_T("10"),&m_DataCombo5,ptrRsCmb5,i,1);
		}
	}
*/
}

void CDlg::OnLockCombo(void)
{
	GetDlgItem(IDC_DATACOMBO1)->EnableWindow(false);
	GetDlgItem(IDC_DATACOMBO2)->EnableWindow(false);
	GetDlgItem(IDC_DATACOMBO3)->EnableWindow(false);
	GetDlgItem(IDC_DATACOMBO4)->EnableWindow(false);
	GetDlgItem(IDC_DATACOMBO5)->EnableWindow(false);
}

void CDlg::OnUnLockCombo(void)
{
	GetDlgItem(IDC_DATACOMBO1)->EnableWindow(true);
	GetDlgItem(IDC_DATACOMBO2)->EnableWindow(true);
	GetDlgItem(IDC_DATACOMBO3)->EnableWindow(true);
	GetDlgItem(IDC_DATACOMBO4)->EnableWindow(true);
	GetDlgItem(IDC_DATACOMBO5)->EnableWindow(true);

}
