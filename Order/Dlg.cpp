// Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg.h"

#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"
//#include <ctype.h>

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
	, m_Edit3(_T("0"))
	, m_Edit4(_T(""))
	, m_Edit5(_T("0"))
	, m_Edit6(_T("0"))
	, m_Edit7(_T("0"))
	, m_Edit8(_T("0"))
	, m_Check1(FALSE)
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

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_Text(pDX, IDC_EDIT4, m_Edit4);
	DDX_Text(pDX, IDC_EDIT5, m_Edit5);
	DDX_Text(pDX, IDC_EDIT6, m_Edit6);
	DDX_Text(pDX, IDC_EDIT7, m_Edit7);
	DDX_Text(pDX, IDC_EDIT8, m_Edit8);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
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
	ON_BN_CLICKED(IDC_BUTTON2, &CDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &CDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &CDlg::OnBnClickedButton5)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"Order.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING9014);
	this->SetWindowText(s);

	m_Edit5F.SubclassDlgItem(IDC_EDIT5,this);
	m_Edit6F.SubclassDlgItem(IDC_EDIT6,this);
	m_Edit7F.SubclassDlgItem(IDC_EDIT7,this);
	m_Edit8F.SubclassDlgItem(IDC_EDIT8,this);

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
	strSql = _T("QT12_1");
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

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

	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);

	strSql = _T("QT6");
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);

	strSql = _T("QT345");
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);

	ptrCmdCmb1 = NULL;
	ptrRsCmb1  = NULL;

	strSql = _T("QT14");
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
		m_DataCombo1.put_ListField(_T("Наименование"));
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb2 = NULL;
	ptrRsCmb2  = NULL;

	strSql = _T("QT16");
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
		m_DataCombo2.put_ListField(_T("Ед. измерения"));	}
	catch(_com_error& e){
		m_DataCombo2.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
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
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32771,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32772,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32773,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
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

		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}

		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs3,m_CurCol,m_Flg);
				m_Flg = true;
			}
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
			case IDC_EDIT4:
				GetDlgItem(IDC_EDIT4)->GetWindowTextW(m_Edit4);
				m_Edit4.TrimRight(' ');
				m_Edit4.TrimLeft(' ');
				if(m_Edit4==' ') m_Edit4.Empty();
				if(m_Edit4.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT4)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT5:
				GetDlgItem(IDC_EDIT5)->GetWindowTextW(m_Edit5);
				m_Edit5.TrimRight(' ');
				m_Edit5.TrimLeft(' ');
				if(m_Edit5==' ') m_Edit5.Empty();
				if(m_Edit5.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT5)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT6:
				GetDlgItem(IDC_EDIT6)->GetWindowTextW(m_Edit6);
				m_Edit6.TrimRight(' ');
				m_Edit6.TrimLeft(' ');
				m_Edit6.Replace(',','.');
				if(m_Edit6==' ') m_Edit6.Empty();
				if(m_Edit6.IsEmpty() || m_Edit6=='.'){
					count++;
					GetDlgItem(IDC_STATIC_EDIT6)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT7:
				GetDlgItem(IDC_EDIT7)->GetWindowTextW(m_Edit7);
				m_Edit7.TrimRight(' ');
				m_Edit7.TrimLeft(' ');
				m_Edit7.Replace(',','.');
				if(m_Edit7==' ') m_Edit7.Empty();
				if(m_Edit7.IsEmpty() || m_Edit7=='.'){
					count++;
					GetDlgItem(IDC_STATIC_EDIT7)->GetWindowText(strItem);
					strCount.Format(_T("%i"),count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT8:
				GetDlgItem(IDC_EDIT8)->GetWindowTextW(m_Edit8);
				m_Edit8.TrimRight(' ');
				m_Edit8.TrimLeft(' ');
				m_Edit8.Replace(',','.');
				if(m_Edit8==' ') m_Edit8.Empty();
				if(m_Edit8.IsEmpty() || m_Edit8=='.'){
					count++;
					GetDlgItem(IDC_STATIC_EDIT8)->GetWindowText(strItem);
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

	strC = m_DataCombo2.get_BoundText();

	strC.TrimRight(' ');
	strC.TrimLeft(' ');
	if(strC==' ') strC.Empty();
	if(strC.IsEmpty()){
		count++;
		GetDlgItem(IDC_STATIC_COMBO2)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}
//Раздел Grid
	if(!IsEnableRec(ptrRs2)){	//Страна
		count++;
		strItem.LoadStringW(IDS_STRING9031);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}
	if(!IsEnableRec(ptrRs3)){	//Классификатор
		count++;
		strItem.LoadStringW(IDS_STRING9030);
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
	short i,IndCol;
	CString strF;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i = 1;	//Арт. внут
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1 = vC.bstrVal;
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
			
			i=2;	//Полн. наим.
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit4 = vC.bstrVal;
			m_Edit4.TrimLeft();
			m_Edit4.TrimRight();
			GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

			i=3; //Арт постав
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit2=vC.bstrVal;
			m_Edit2.TrimLeft();
			m_Edit2.TrimRight();
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);


			i=7;	//Числ. выр. шт.	
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit6=vC.bstrVal;
			m_Edit6.TrimLeft();
			m_Edit6.TrimRight();
			GetDlgItem(IDC_EDIT6)->SetWindowText(m_Edit6);

			i=9; // Код Страны
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			IndCol = 0;
			m_Flg = false;
//AfxMessageBox("E");
			OnFindInGrid(strF,ptrRs2,IndCol,m_Flg);
			m_Flg = true;

			i=10; // Код вида штуки
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			OnFindInCombo(strF,&m_DataCombo1,ptrRsCmb1,0,1);

			i=11; // Код Ед.измер.
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			OnFindInCombo(strF,&m_DataCombo2,ptrRsCmb2,0,1);

			i=12; //Масса (кг)	
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit7=vC.bstrVal;
			m_Edit7.Replace(',','.');
			m_Edit7.TrimLeft();
			m_Edit7.TrimRight();
			GetDlgItem(IDC_EDIT7)->SetWindowText(m_Edit7);

			i=13; //М (куб.м)	
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit8=vC.bstrVal;
			m_Edit8.Replace(',','.');
			m_Edit8.TrimLeft();
			m_Edit8.TrimRight();
			GetDlgItem(IDC_EDIT8)->SetWindowText(m_Edit8);

			i=14;	//Штрих-Код
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit3=vC.bstrVal;
			m_Edit3.TrimLeft();
			m_Edit3.TrimRight();
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);

			i=15;	//Спиртосодерж
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BOOL);
			m_Check1=vC.boolVal;
			((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);

			i=16; //Кол-во спирта (л)
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit5=vC.bstrVal;
			m_Edit5.Replace(',','.');
			m_Edit5.TrimLeft();
			m_Edit5.TrimRight();
			GetDlgItem(IDC_EDIT5)->SetWindowText(m_Edit5);

			i=18; // Код классификатора
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			IndCol = 0;
			m_Flg = false;
			OnFindInGrid(strF,ptrRs3,IndCol,m_Flg);
			m_Flg = true;

		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_EDIT4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_STATIC_EDIT5)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9026);
	GetDlgItem(IDC_STATIC_EDIT6)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9027);
	GetDlgItem(IDC_STATIC_EDIT7)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_STATIC_EDIT8)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql;
	CString strC5,strC7,strC9,strC15;
	_variant_t vra;
	VARIANT* vtl = NULL;
	COleVariant vC;
	short i,IndCol;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	if(OnOverEdit(IDC_EDIT1,IDC_EDIT8)){
		if(IsEnableRec(ptrRs2)){
			i = 0;
			vC = GetValueRec(ptrRs2,i);		//Код страны
			vC.ChangeType(VT_BSTR);
			strC5 = vC.bstrVal;
			strC5.TrimRight();
			strC5.TrimLeft();
		}

		if(IsEnableRec(ptrRsCmb1)){
			i = 0;
			vC = GetValueRec(ptrRsCmb1,i);	//Вид шт
			vC.ChangeType(VT_BSTR);
			strC7 = vC.bstrVal;
			strC7.TrimLeft();
			strC7.TrimRight();
		}

		if(IsEnableRec(ptrRsCmb2)){
			i = 0;
			vC = GetValueRec(ptrRsCmb2,i);	//Ед.измер.
			vC.ChangeType(VT_BSTR);
			strC9 = vC.bstrVal;
			strC9.TrimLeft();
			strC9.TrimRight();
		}

		if(IsEnableRec(ptrRs3)){
			i = 0;
			vC = GetValueRec(ptrRs2,i);		//Код классификатора
			vC.ChangeType(VT_BSTR);
			strC15 = vC.bstrVal;
			strC15.TrimRight();
			strC15.TrimLeft();
		}
		strSql =_T("IT12_1 '");
		strSql+= m_Edit1+_T("','");			/*Арт вн*/
		strSql+= m_Edit4+_T("','");			/*Полн наим*/
		strSql+= m_Edit2 + _T("',");		/*Арт п*/
		strSql+= strC5 + _T(",'");			/*Код стр*/
		strSql+= _T("0',");					/*ГТД*/
		strSql+= strC7	 + _T(",");			/*Код шт*/
		strSql+= m_Edit6 + _T(",");			/*Числ-ое выр.шт.*/
		strSql+= strC9   + _T(",");			/*Код ед.измер.*/
		strSql+= m_Edit7 + _T(",");			/*Масса*/
		strSql+= m_Edit8 + _T(",");			/*V*/
		strSql+= m_Edit3 + _T(",");			/*Штрих Код*/
		strSql+= (m_Check1 ? _T("1"):_T("0"));	/*Спиртосод Есть/Нет (1/0)*/
		strSql+= _T(",");
		strSql+= m_Edit5 + _T(",");			/*V спирта (л)*/
		strSql+= strC15  + _T(",'");		/*Код классификатора*/
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
		strSql = _T("QT12_1");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			if(!m_Edit1.IsEmpty()){
				IndCol = 1;
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
	CString strSql,strC;
	CString strC5,strC7,strC9,strC15;
	_variant_t vra;
	COleVariant vC;
	VARIANT* vtl = NULL;
	short IndCol,i;

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){

			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();

			if(OnOverEdit(IDC_EDIT1,IDC_EDIT8)){

				if(IsEnableRec(ptrRs2)){
					i = 0;
					vC = GetValueRec(ptrRs2,i);		//Код страны
					vC.ChangeType(VT_BSTR);
					strC5 = vC.bstrVal;
					strC5.TrimRight();
				}

				if(IsEnableRec(ptrRsCmb1)){
					i = 0;
					vC = GetValueRec(ptrRsCmb1,i);	//Вид шт
					vC.ChangeType(VT_BSTR);
					strC7 = vC.bstrVal;
					strC7.TrimLeft();
				}

				if(IsEnableRec(ptrRsCmb2)){
					i = 0;
					vC = GetValueRec(ptrRsCmb2,i);	//Ед.измер.
					vC.ChangeType(VT_BSTR);
					strC9 = vC.bstrVal;
					strC9.TrimLeft(' ');
				}

				if(IsEnableRec(ptrRs3)){
					i = 0;
					vC = GetValueRec(ptrRs2,i);		//Код классификатора
					vC.ChangeType(VT_BSTR);
					strC15 = vC.bstrVal;
					strC15.TrimRight();
					strC15.TrimLeft();
				}

				strSql = _T("UT12_1 ");
				strSql+= strC +_T(",'");
				strSql+= m_Edit1 + _T("','");			//Арт вн
				strSql+= m_Edit4 + _T("','");			//Полн. наим.
				strSql+= m_Edit2 + _T("',");			//Арт пост
				strSql+= strC5   + _T(",'");			//Код Страны
				strSql+= _T("0',");						//ГТД
				strSql+= strC7	 + _T(",");				//Код вида шт.
				strSql+= m_Edit6 + _T(",");				//Числ. выр. шт.
				strSql+= strC9   + _T(",");				//Код Ед. изм.
				strSql+= m_Edit7 + _T(",");				//Масса шт.
				strSql+= m_Edit8 + _T(",");				//V куб.м
				strSql+= m_Edit3 + _T(",");				//Штрих код
				strSql+= (m_Check1 ? _T("1"):_T("0"));	//Спиртосодер.
				strSql+= _T(",");
				strSql+= m_Edit5 + _T(",");				//V спирта
				strSql+= strC15  + _T(",'");			//Код классификатора
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
    			strSql = _T("QT12_1");
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

			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimLeft();

			strSql = _T("DT12 ");
			strSql+= strC;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
    		strSql = _T("QT12_1");
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

/*			m_CurCol = Grd.get_Col();
			if(m_CurCol==-1 || m_CurCol==EmpCol ){
				m_CurCol = DefCol;
			}
			m_iCurType = GetTypeCol(rs,m_CurCol);
			m_EditTBCh.SetTypeCol(m_iCurType);
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
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 270;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 6:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 7:
				wdth = 30;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 8:
				wdth = 30;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 12:
				wdth = 30;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 13:
				wdth = 30;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 14:
				wdth = 40;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 16:
				wdth = 40;
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
				wdth = 260;
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

	strCap.LoadString(IDS_STRING9029);
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
				wdth = 260;
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
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, CDlg::ChangeDatacombo2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
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
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
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

void CDlg::OnBnClickedButton1()
{
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strIns;
		COleVariant vC;
		short i;
		bFndC   = TRUE;
		strFndC = _T("");
		if(IsEnableRec(ptrRsCmb1)){
			i = 0;
			vC = GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();
		}
		hMod=AfxLoadLibrary(_T("NameUnit.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameUnit");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);
//		m_SlpDay.SetDate(t1);
		AfxFreeLibrary(hMod);
		try{
			ptrRsCmb1->Requery(adCmdText);
			m_DataCombo1.Refresh();
			OnFindInCombo(strFndC,&m_DataCombo1,ptrRsCmb1,0,1);
			i = 1;
			vC = GetValueRec(ptrRsCmb1,i);
			vC.ChangeType(VT_BSTR);
			strIns = vC.bstrVal;
			m_DataCombo1.put_BoundText(strIns);

		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}


	EndWaitCursor();
}

void CDlg::OnBnClickedButton2()
{
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strIns;
		COleVariant vC;
		short i;
		bFndC   = TRUE;
		strFndC = _T("");
		if(IsEnableRec(ptrRsCmb2)){
			i = 0;
			vC = GetValueRec(ptrRsCmb2,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();
		}
		hMod=AfxLoadLibrary(_T("UnitPls.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startUnitPls");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);
//		m_SlpDay.SetDate(t1);
		AfxFreeLibrary(hMod);
		try{
			ptrRsCmb2->Requery(adCmdText);
			m_DataCombo2.Refresh();
			OnFindInCombo(strFndC,&m_DataCombo2,ptrRsCmb2,0,1);
			i = 1;
			vC = GetValueRec(ptrRsCmb2,i);
			vC.ChangeType(VT_BSTR);
			strIns = vC.bstrVal;
			m_DataCombo2.put_BoundText(strIns);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}


	EndWaitCursor();
}

void CDlg::OnBnClickedButton3()
{
	COleVariant vC;
	short i;
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		if(IsEnableRec(ptrRs1)){
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();

			bFndC = TRUE;
		}
		hMod=AfxLoadLibrary(_T("OrdGroup.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdGroup");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CDlg::OnBnClickedButton4()
{
	COleVariant vC;
	short i;
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		CString strFndC;
		BOOL bFndC;
		HMODULE hMod;
		if(IsEnableRec(ptrRs1)){
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();

			bFndC = TRUE;
		}
		hMod=AfxLoadLibrary(_T("OrdWrap.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdWrap");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::OnBnClickedButton5()
{
//Товар-склад
	COleVariant vC;
	short i;
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		if(IsEnableRec(ptrRs1)){
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();

			bFndC = TRUE;
		}
		hMod=AfxLoadLibrary(_T("OrdStrg.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdStrg");
		(func)(m_strNT, ptrCnn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CDlg::ClickDatagrid3()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID3);
}

void CDlg::ClickDatagrid2()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID2);
}

void CDlg::ClickDatagrid1()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
}


void CDlg::ChangeDatacombo1()
{
	OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
}

void CDlg::ChangeDatacombo2()
{
	OnChangeCombo(m_DataCombo2,ptrRsCmb2,1);
}

