// Dlg.cpp : implementation file
//

#include "stdafx.h"
//#include "AddBk.h"
#include "Dlg.h"
#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"

#include "DlgFtch.h"


// CDlg dialog

IMPLEMENT_DYNAMIC(CDlg, CDialog)
CDlg::CDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDlg::IDD, pParent)
	, m_CurCol(0)
	, m_iBtSt(0)
	, m_Flg(true)
	, m_strNT(_T(""))
	, m_bFnd(false)
	, m_Edit1(_T("0"))
	, m_Edit5(_T("-"))
	, m_Edit6(_T("0"))
	, cd(0)
	, m_bCombo3(false)
	, m_Edit2(_T("-"))
	, m_Edit3(_T("0"))
	, m_Edit4(_T("-"))
	, m_strCod(_T(""))
	, m_strIns(_T(""))
{
	m_pDlg = new CDlgFtch(this);
}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

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

	if(ptrRsCmb6->State==adStateOpen) ptrRsCmb6->Close();
	ptrRsCmb6 = NULL;

	if(ptrRsCmb7->State==adStateOpen) ptrRsCmb7->Close();
	ptrRsCmb7 = NULL;

	if(ptrRsCmb8->State==adStateOpen) ptrRsCmb8->Close();
	ptrRsCmb8 = NULL;

	delete m_pDlg;	//Уничтожение ообьекта выборки CDlgFtch

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT5, m_Edit5);
	DDX_Text(pDX, IDC_EDIT6, m_Edit6);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_Control(pDX, IDC_DATACOMBO3, m_DataCombo3);
	DDX_Control(pDX, IDC_DATACOMBO4, m_DataCombo4);
	DDX_Control(pDX, IDC_DATACOMBO5, m_DataCombo5);
	DDX_Control(pDX, IDC_DATACOMBO6, m_DataCombo6);
	DDX_Control(pDX, IDC_DATACOMBO7, m_DataCombo7);
	DDX_Control(pDX, IDC_DATACOMBO8, m_DataCombo8);
	DDX_Control(pDX, IDC_DATACOMBO9, m_DataCombo9);
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
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)

	ON_MESSAGE(WM_FTCH,OnFetch)
	ON_MESSAGE(WM_CLOSEFTCH,OnCloseFetch)

END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"AddBk.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING9013);
	this->SetWindowText(s);

//	m_Edit1Rs.SubclassDlgItem(IDC_EDIT1,this);

	InitStaticText();

	int iTBCtrlID;
	short i;
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

/*	if(m_EditTB.GetSafeHwnd()!=NULL){
		AfxMessageBox("not NULL");
		if(m_EditTBSChar.SubclassWindow(m_EditTB.GetSafeHwnd())){
			AfxMessageBox("ok");
		}
	}
*/
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);

	m_vNULL.vt = VT_ERROR;
	m_vNULL.scode = DISP_E_PARAMNOTFOUND;

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;
//	ptrCmd1->CommandText = _T("QT44");

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);
/*	try{
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataGrid1.putref_DataSource(NULL);
		m_DataGrid1.putref_DataSource((LPUNKNOWN)ptrRs1);
		
		m_CurCol = m_DataGrid1.get_Col();
		if(m_CurCol==-1 || m_CurCol==0 ){
			m_CurCol = 1;
		}
		m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		InitDataGrid1();
	}
	catch(_com_error& e){
		m_DataGrid1.putref_DataSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
*/
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
	
	ptrCmdCmb6 = NULL;
	ptrRsCmb6  = NULL;

	strSql = _T("QT34");
	ptrCmdCmb6.CreateInstance(__uuidof(Command));
	ptrCmdCmb6->ActiveConnection = ptrCnn;
	ptrCmdCmb6->CommandType = adCmdText;
	ptrCmdCmb6->CommandText = (_bstr_t)strSql;

	ptrRsCmb6.CreateInstance(__uuidof(Recordset));
	ptrRsCmb6->CursorLocation = adUseClient;
	ptrRsCmb6->PutRefSource(ptrCmdCmb6);
	try{
		m_DataCombo6.putref_RowSource(NULL);
		ptrRsCmb6->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo6.putref_RowSource((LPUNKNOWN) ptrRsCmb6);
		m_DataCombo6.put_ListField(_T("Сокр. наим."));
	}
	catch(_com_error& e){
		m_DataCombo6.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb7 = NULL;
	ptrRsCmb7  = NULL;

	strSql = _T("QT50 0");
	ptrCmdCmb7.CreateInstance(__uuidof(Command));
	ptrCmdCmb7->ActiveConnection = ptrCnn;
	ptrCmdCmb7->CommandType = adCmdText;
	ptrCmdCmb7->CommandText = (_bstr_t)strSql;

	ptrRsCmb7.CreateInstance(__uuidof(Recordset));
	ptrRsCmb7->CursorLocation = adUseClient;
	ptrRsCmb7->PutRefSource(ptrCmdCmb7);
	try{
		m_DataCombo7.putref_RowSource(NULL);
		ptrRsCmb7->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo7.putref_RowSource((LPUNKNOWN) ptrRsCmb7);
		m_DataCombo7.put_ListField(_T("Наименование"));
	}
	catch(_com_error& e){
		m_DataCombo7.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb8 = NULL;
	ptrRsCmb8  = NULL;

	strSql = _T("QT35");
	ptrCmdCmb8.CreateInstance(__uuidof(Command));
	ptrCmdCmb8->ActiveConnection = ptrCnn;
	ptrCmdCmb8->CommandType = adCmdText;
	ptrCmdCmb8->CommandText = (_bstr_t)strSql;

	ptrRsCmb8.CreateInstance(__uuidof(Recordset));
	ptrRsCmb8->CursorLocation = adUseClient;
	ptrRsCmb8->PutRefSource(ptrCmdCmb8);
	try{
		m_DataCombo8.putref_RowSource(NULL);
		ptrRsCmb8->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo8.putref_RowSource((LPUNKNOWN) ptrRsCmb8);
		m_DataCombo8.put_ListField(_T("Сокр. наим."));
	}
	catch(_com_error& e){
		m_DataCombo8.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb9 = NULL;
	ptrRsCmb9  = NULL;

	strSql = _T("QT36");
	ptrCmdCmb9.CreateInstance(__uuidof(Command));
	ptrCmdCmb9->ActiveConnection = ptrCnn;
	ptrCmdCmb9->CommandType = adCmdText;
	ptrCmdCmb9->CommandText = (_bstr_t)strSql;

	ptrRsCmb9.CreateInstance(__uuidof(Recordset));
	ptrRsCmb9->CursorLocation = adUseClient;
	ptrRsCmb9->PutRefSource(ptrCmdCmb9);
	try{
		m_DataCombo9.putref_RowSource(NULL);
		ptrRsCmb9->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo9.putref_RowSource((LPUNKNOWN) ptrRsCmb9);
		m_DataCombo9.put_ListField(_T("Вид адреса"));
	}
	catch(_com_error& e){
		m_DataCombo9.putref_RowSource(NULL);
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
//Здесь просто оключаем все кнопки за исключением Выбрать,Печать
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
	m_EditTBCh.EnableWindow(FALSE);
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
					GetDlgItem(IDC_STATIC_COMBO8)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
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
			i = 1;	//Имя
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

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

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_COMBO6)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_STATIC_COMBO7)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_COMBO8)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_EDIT5)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9026);
	GetDlgItem(IDC_STATIC_EDIT6)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9027);
	GetDlgItem(IDC_STATIC_COMBO9)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_STATIC_SL_1)->SetWindowText(sTxt);
	GetDlgItem(IDC_STATIC_SL_2)->SetWindowText(sTxt);


}

void CDlg::InitDataGrid1(void)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;

	strCap.LoadString(IDS_STRING9029);
	numRec = 0;

	GrdClms.AttachDispatch(m_DataGrid1.get_Columns());
	if(ptrRs1->State==adStateOpen){
		num = ptrRs1->GetadoFields()->GetCount();

		numRec = ptrRs1->GetRecordCount();
		strRec.Format(_T("%i"),numRec);
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
		strRec.Format(_T("%i"),numRec);
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


	if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
		strSql = _T("IT43 '");
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
		ptrCmd1->CommandText = _T("QT43");
		try{
			ptrRs1->Requery(adCmdText);
			m_DataGrid1.Refresh();
//			InitDataGrid1();
			if(!m_Edit1.IsEmpty()){
				IndCol = 1;
				m_Flg = false;
				OnFindInGrid(m_Edit1,ptrRs1,IndCol,m_Flg);
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
	short IndCol,i;

	m_Edit1.TrimRight(' ');
	m_Edit1.TrimLeft(' ');
	if(m_Edit1==' ') m_Edit1.Empty();

	if(ptrRs1->State==adStateOpen){
		if(!IsEmptyRec(ptrRs1)){
			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IDC_EDIT1)){
				strSql = _T("UT43 ");
				strSql+= strC	 + _T(",'");
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
				ptrCmd1->CommandText = _T("QT43");
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

			strSql = _T("DT43 ");
			strSql+= strC;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			ptrCmd1->CommandText = _T("QT43");
			try{
				ptrRs1->Requery(adCmdText);
				m_DataGrid1.Refresh();
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
	m_iBtSt = m_wndToolBar.GetToolBarCtrl().GetState(ID_BUTTON32779);
	OnEnableButtonBar(m_iBtSt,&m_wndToolBar);
	if(m_iBtSt==5){
		OnShowEdit();
	}
}


void CDlg::On32778(void)
{
//Выбрать
	if(m_pDlg->GetSafeHwnd()==0){
		m_pDlg->Create();
	}
	return;
}

void CDlg::On32783(void)
{
	return;
}

void CDlg::OnBnClickedOk()
{
}

void CDlg::OnBnClickedCancel()
{
}

void CDlg::OnClickedOk()
{
	UpdateData();
	COleVariant vC;
	CString strSql,strC,strC3,strMsg,strCount,strItem,strC10,strC11;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i=0;
	int count = 0;
	strMsg.LoadString(IDS_STRING9016);
	
	m_Edit1.TrimRight(' ');		//Дом
	m_Edit1.TrimLeft( ' ');
	if(m_Edit1.IsEmpty()) m_Edit1 = _T("0");

	m_Edit2.TrimRight(' ');
	m_Edit2.TrimLeft( ' ');
	if(m_Edit2.IsEmpty()) m_Edit2 = _T("-");

	m_Edit3.TrimRight(' ');		//кв
	m_Edit3.TrimLeft( ' ');
	if(m_Edit3.IsEmpty()) 	m_Edit3 = _T("0");

	m_Edit4.TrimRight(' ');
	m_Edit4.TrimLeft( ' ');
	if(m_Edit4.IsEmpty()) m_Edit4 = _T("-");

	m_Edit5.TrimRight(' ');		//Дополнения
	m_Edit5.TrimLeft( ' ');
	if(m_Edit5.IsEmpty()) 	m_Edit5 = _T("-");

	m_Edit6.TrimRight(' ');		//Индекс
	m_Edit6.TrimLeft( ' ');
	if(m_Edit6.IsEmpty()) 	m_Edit6 = _T("0");


//Раздел Combo
	strC = m_DataCombo9.get_BoundText(); //Вид адреса

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO9)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	strC = m_DataCombo6.get_BoundText();//Тип наименований

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO6)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	strC = m_DataCombo7.get_BoundText(); // Наименования

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO7)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	strC = m_DataCombo8.get_BoundText(); //Адм. здания

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO8)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	
	if(ptrRs1->State==adStateClosed){
		count++;
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount + _T("Адрес из Кладра.");
	}
	else{
		if(IsEmptyRec(ptrRs1)){
			count++;
			strCount.Format(_T("%i"),count);
			strCount+=_T(") ");
			strMsg+=strCount + _T("Адрес из Кладра.");
		}
	}

	if(count!=0){
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	}
	else{
		i = 0;
		vC = GetValueRec(ptrRsCmb9,i);
		vC.ChangeType(VT_BSTR);
		strC3=vC.bstrVal;			//Код Вида Адреса
/*		switch(m_Radio1){
		case 0:	// по Наименованию
			vC = rs->GetFields()->GetItem((COleVariant) i)->Value;
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;			//Код Кладра
			break;
		case 1:	//по Запросу
			vC = rs->GetFields()->GetItem((COleVariant) i)->Value;
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;			//Код Кладра
			break;
		}
*/
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strC = vC.bstrVal;			//Код Кладра

		vC = GetValueRec(ptrRsCmb8,i);
		vC.ChangeType(VT_BSTR);
		strC10=vC.bstrVal;			//Код Адм.здания

		vC = GetValueRec(ptrRsCmb7,i);
		vC.ChangeType(VT_BSTR);
		strC11=vC.bstrVal;			//Код Наименования


		strSql = m_strIns;         //_T("IT49 IT38 ");
		strSql+= m_strCod +_T(",");
		strSql+= strC + _T(",");
		strSql+= strC3 + _T(",");
		strSql+= m_Edit1 +_T(",'");
		strSql+= m_Edit2 +_T("',");
		strSql+= m_Edit3 +_T(",'");
		strSql+= m_Edit4 +_T("','");
		strSql+= m_Edit5 +_T("',");
		strSql+= strC10 + _T(",");
		strSql+= strC11 + _T(",'");
		strSql+= m_Edit6 + _T("','");
		strSql+=m_strNT+_T("'");

		BSTR strCmd;
		ptrCmd1->get_CommandText(&strCmd);

		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
			m_pDlg->DestroyWindow();
			OnOK();
		}
		catch(_com_error& e){
			ptrCmd1->put_CommandText(strCmd);
			AfxMessageBox(e.Description());
		}
		
	}
	m_pDlg->DestroyWindow();
}

void CDlg::OnClickedCancel()
{
	m_pDlg->DestroyWindow();
	OnCancel();
}
BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO3, 1, ChangeDatacombo3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO5, 1, ChangeDatacombo5, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, ChangeDatacombo2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO4, 1, ChangeDatacombo4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO6, 1, ChangeDatacombo6, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO7, 1, ChangeDatacombo7, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO8, 1, ChangeDatacombo8, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO9, 1, ChangeDatacombo9, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
END_EVENTSINK_MAP()

void CDlg::ChangeDatacombo1()
{
	cd = OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	OnShowComboGlb(cd);
	if(cd!=10){
		OnUpdateCombo2();
		OnUpdateCombo3();	//Делаем города
							// Меняем Терр-ые ед-цы России
							// и становимся на первую запись 
							// выборки через SetBoundText()
		cd  = OnShowCombo();
		UpdateGrid1(cd);
	}
	else{
		UpdateGrid1Glb();
	}
}

void CDlg::OnUpdateCombo2(void)
{
//AfxMessageBox("OnUpdateCombo2");
	CString strCur,strSql;
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

void CDlg::UpdateGrid1(long i)
{
	COleVariant vC;
	CString strCur,strCur1,str,strNumRec,strSql;
	long NumRec=0;

	short iC;
	switch(i){
	case 10:
		if(ptrRsCmb1->adoEOF || ptrRsCmb1->adoEOF) {
			return;
		}
		else{
			iC = 3;
			vC=GetValueRec(ptrRsCmb2,iC);
			vC.ChangeType(VT_BSTR);
			strCur=vC.bstrVal;
			str = strCur.Left(5);
	//	AfxMessageBox(strCur);

			if(ptrRsCmb3->adoEOF) return;

			iC = 0;
			vC=GetValueRec(ptrRsCmb3,iC);
			vC.ChangeType(VT_BSTR);
			strCur1=vC.bstrVal;

			strSql = _T("QT8Twn '");
			strSql+= str;
			strSql+= _T("',");
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
		}
		break;
	default:
		iC = 3;
		if(ptrRsCmb2->adoEOF && ptrRsCmb2->BOF){
//AfxMessageBox("return");
			return;
		}
		else{
			vC=GetValueRec(ptrRsCmb4,iC);
			vC.ChangeType(VT_BSTR);
			strCur=vC.bstrVal;
			str = strCur.Left(5);
		}

		iC = 0;
		vC=GetValueRec(ptrRsCmb5,iC);
		vC.ChangeType(VT_BSTR);
		strCur1=vC.bstrVal;

		strSql = _T("QT8TwnVlg '");
		strSql+= str;
		strSql+= _T("',");
		strSql+= strCur1;
        ptrCmd1->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
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
		break;
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

void CDlg::ChangeDatacombo3()
{
//AfxMessageBox("start OnChangeDatacombo3");
	cd = OnChangeCombo(m_DataCombo3, ptrRsCmb3,1);
	if(m_bCombo3){
		cd=OnShowCombo();
		if(10!=cd){
			OnUpdateCombo4();
			m_DataGrid1.putref_DataSource(NULL);
		}
		UpdateGrid1(cd);
	}
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

//AfxMessageBox(strSql);

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


}

void CDlg::ChangeDatacombo5()
{
	cd = OnChangeCombo(m_DataCombo5, ptrRsCmb5,1);
	UpdateGrid1(OnShowCombo());
}

void CDlg::ChangeDatacombo2()
{
//	AfxMessageBox("start OnChangeDatacombo4");
	OnChangeCombo(m_DataCombo2, ptrRsCmb2,1);
//		if(m_bCombo4){
	//	AfxMessageBox("invoid OnUpdateCombo3 from OnChangeDatacombo4");
			OnUpdateCombo3();
			cd=OnShowCombo();
			UpdateGrid1(cd);
//		}
}

void CDlg::ChangeDatacombo4()
{
//	AfxMessageBox("start OnChangeDatacombo5");
	OnChangeCombo(m_DataCombo4, ptrRsCmb4,1);
}

void CDlg::ChangeDatacombo6()
{
	CString strCd,strCur,strSql;
	COleVariant vC;
	cd = OnChangeCombo(m_DataCombo6,ptrRsCmb6,1);	
	strCd.Format(_T("%i"),cd);
//AfxMessageBox(strCd);
	short i;
	
	if(!m_DataCombo6.get_Text().IsEmpty()){

		strSql = _T("QT50 ");
		strSql+= strCd;

		ptrCmdCmb7->CommandType = adCmdText;
		ptrCmdCmb7->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
		try{
			m_DataCombo7.putref_RowSource(NULL);
			if(ptrRsCmb7->State==adStateOpen) ptrRsCmb7->Close();
			ptrRsCmb7->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
			m_DataCombo7.putref_RowSource((LPUNKNOWN) ptrRsCmb7);
			m_DataCombo7.put_ListField(_T("Наименование"));
			if(!IsEmptyRec(ptrRsCmb7)){
				ptrRsCmb7->MoveFirst();
				i = 1;
				vC = GetValueRec(ptrRsCmb7,i);
				vC.ChangeType(VT_BSTR);
				strCur = vC.bstrVal;
		//AfxMessageBox(strCur);
				m_DataCombo7.put_BoundText(strCur);
			}
			else{
				m_DataCombo7.put_Text(_T(""));
			}
		}
		catch(_com_error& e){
			m_DataCombo7.putref_RowSource(NULL);
			AfxMessageBox(e.ErrorMessage());
		}
	}
}

void CDlg::OnBnClickedButton1()
{
	COleVariant vC;
	CString strC,strC1,strIns;
	short i;
	i = 0;
	strC = _T("");
	strC1= _T("");
	if(IsEnableRec(ptrRsCmb6)){
		vC = GetValueRec(ptrRsCmb6,i);
		vC.ChangeType(VT_BSTR);
		strC = vC.bstrVal;
		strC.TrimLeft();
		strC.TrimRight();
	}	
	if(IsEnableRec(ptrRsCmb7)){
		vC = GetValueRec(ptrRsCmb7,i);
		vC.ChangeType(VT_BSTR);
		strC1 = vC.bstrVal;
		strC1.TrimLeft();
		strC1.TrimRight();
	}	
	BeginWaitCursor();
//AfxMessageBox(strC);
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		hMod=AfxLoadLibrary(L"NameStrt.dll");
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameStrt");
		if(!strC.IsEmpty()){
			(func)(m_strNT, ptrCnn, strC, strC1, TRUE);
		}
		else{
			(func)(m_strNT, ptrCnn, strC, strC1, FALSE);
		}

		try{
			ptrRsCmb6->Requery(adCmdText);
			m_DataCombo6.Refresh();
			if(IsEnableRec(ptrRsCmb6)){
				OnFindInCombo(strC,&m_DataCombo6,ptrRsCmb6,0,1);

				ptrRsCmb7->Requery(adCmdText);
				m_DataCombo7.Refresh();
				if(IsEnableRec(ptrRsCmb7)){
					OnFindInCombo(strC1,&m_DataCombo7,ptrRsCmb7,0,1);
				}
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::ChangeDatacombo7()
{
	OnChangeCombo(m_DataCombo7,ptrRsCmb7,1);	
}

void CDlg::ChangeDatacombo8()
{
	OnChangeCombo(m_DataCombo8,ptrRsCmb8,1);	
}

void CDlg::ChangeDatacombo9()
{
	OnChangeCombo(m_DataCombo9,ptrRsCmb9,1);	
}

void CDlg::OnShowIndex(void)
{
	COleVariant vC;
	int i;
	m_Edit6 = _T("0");
	if(IsEnableRec(ptrRs1)){
			i = 4;
			vC= GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit6 = vC.bstrVal;
	}
	GetDlgItem(IDC_EDIT6)->SetWindowText(m_Edit6);
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
	if(m_Flg){
		if(ptrRs1->State==adStateOpen){
			if(!IsEmptyRec(ptrRs1)){
//AfxMessageBox("r1");
				OnShowIndex();
			}
		}
	}
	m_Flg = true;
}

LRESULT CDlg::OnFetch(WPARAM wParam, LPARAM lParam)
{
	CString strC2,strSql;
	strC2 = m_pDlg->m_Edit1;
	if(!strC2.IsEmpty()){
		strSql = _T("QT8C2 '");
		strSql+= strC2 + _T("'");
		ptrCmd1->CommandText = (_bstr_t)strSql;
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
			OnShowIndex();

		}
		catch(_com_error& e){
			m_DataGrid1.putref_DataSource(NULL);
			AfxMessageBox(e.ErrorMessage());
		}
	}
	return 0L;
}

LRESULT CDlg::OnCloseFetch(WPARAM wParam, LPARAM lParam)
{
	m_pDlg->DestroyWindow();
	return 0L;
}
