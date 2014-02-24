// Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "Dlg.h"

#include "Exports.h"
#include ".\dlg.h"

#include "Columns.h"
#include "Column.h"

#include <ctype.h>
/*#include <stdlib.h>*/
#include <afxtempl.h>
#include "CApplication.h"

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
	, m_strFndC1(_T(""))
	, m_bFndC(FALSE)
	, m_pLstWnd(NULL)

	, m_Radio1(0)
	, m_Check1(TRUE)
	, m_Check2(TRUE)
	, m_Check3(TRUE)
	, m_Check4(TRUE)
	, m_Check5(TRUE)
	, m_Check6(TRUE)
	, m_Check7(FALSE)
	, m_Check8(FALSE)
	, m_strSql(_T(""))
	, m_strCod(_T(""))
	, m_Edit1(_T(""))
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
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_Radio(pDX, IDC_RADIO1, m_Radio1);
	DDX_Check(pDX, IDC_CHECK2, m_Check2);
	DDX_Check(pDX, IDC_CHECK3, m_Check3);
	DDX_Check(pDX, IDC_CHECK4, m_Check4);
	DDX_Check(pDX, IDC_CHECK5, m_Check5);
	DDX_Check(pDX, IDC_CHECK6, m_Check6);
	DDX_Check(pDX, IDC_CHECK7, m_Check7);
	DDX_Check(pDX, IDC_CHECK8, m_Check8);
	DDX_Control(pDX, IDC_COMBO1, m_ComboBox1);
	DDX_Control(pDX, IDC_COMBO2, m_ComboBox2);
	DDX_Control(pDX, IDC_COMBO3, m_ComboBox3);
	DDX_Control(pDX, IDC_COMBO4, m_ComboBox4);
	DDX_Control(pDX, IDC_COMBO5, m_ComboBox5);
	DDX_Control(pDX, IDC_COMBO6, m_ComboBox6);
	DDX_Control(pDX, IDC_COMBO7, m_ComboBox7);
	DDX_Control(pDX, IDC_COMBO8, m_ComboBox8);
	DDX_Control(pDX, IDC_DATAGRID4, m_DataGrid4);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_COMMAND(ID_BUTTON32771, On32771)
	ON_COMMAND(ID_BUTTON32772, On32772)
	ON_COMMAND(ID_BUTTON32773, On32773)
	ON_COMMAND(ID_BUTTON32774, On32774)
	ON_COMMAND(ID_BUTTON32775, On32775)
	ON_COMMAND(ID_BUTTON32776, On32776)
	ON_COMMAND(ID_BUTTON32777, On32777)
	ON_COMMAND(ID_BUTTON32778, On32778)
	ON_COMMAND(ID_BUTTON32779, On32779)
	ON_COMMAND(ID_BUTTON32780, On32780)
	ON_COMMAND(ID_BUTTON32783, On32783)

	ON_EN_CHANGE(ID_EDIT_TOOLBAR, OnEnChangeEditTB)

	ON_BN_CLICKED(IDOK, &CDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_OK, &CDlg::OnClickedOk)
	ON_BN_CLICKED(IDC_CANCEL, &CDlg::OnClickedCancel)
	ON_BN_CLICKED(IDC_RADIO1, &CDlg::OnBnClickedRadio1)
	ON_BN_CLICKED(IDC_RADIO2, &CDlg::OnBnClickedRadio2)
	ON_BN_CLICKED(IDC_CHECK1, &CDlg::OnBnClickedCheck1)
	ON_BN_CLICKED(IDC_CHECK2, &CDlg::OnBnClickedCheck2)
	ON_BN_CLICKED(IDC_CHECK3, &CDlg::OnBnClickedCheck3)
	ON_BN_CLICKED(IDC_CHECK4, &CDlg::OnBnClickedCheck4)
	ON_BN_CLICKED(IDC_CHECK5, &CDlg::OnBnClickedCheck5)
	ON_BN_CLICKED(IDC_CHECK6, &CDlg::OnBnClickedCheck6)
	ON_BN_CLICKED(IDC_CHECK8, &CDlg::OnBnClickedCheck8)
	ON_CBN_EDITCHANGE(IDC_COMBO1, &CDlg::OnCbnEditchangeCombo1)
	ON_CBN_EDITCHANGE(IDC_COMBO2, &CDlg::OnCbnEditchangeCombo2)
	ON_CBN_EDITCHANGE(IDC_COMBO3, &CDlg::OnCbnEditchangeCombo3)
	ON_CBN_EDITCHANGE(IDC_COMBO4, &CDlg::OnCbnEditchangeCombo4)
	ON_CBN_EDITCHANGE(IDC_COMBO5, &CDlg::OnCbnEditchangeCombo5)
	ON_CBN_EDITCHANGE(IDC_COMBO6, &CDlg::OnCbnEditchangeCombo6)
	ON_CBN_EDITCHANGE(IDC_COMBO7, &CDlg::OnCbnEditchangeCombo7)
	ON_CBN_EDITCHANGE(IDC_COMBO8, &CDlg::OnCbnEditchangeCombo8)
	ON_CBN_SELENDOK(IDC_COMBO1, &CDlg::OnCbnSelendokCombo1)
	ON_CBN_SELENDOK(IDC_COMBO2, &CDlg::OnCbnSelendokCombo2)
	ON_CBN_SELENDOK(IDC_COMBO3, &CDlg::OnCbnSelendokCombo3)
	ON_CBN_SELENDOK(IDC_COMBO4, &CDlg::OnCbnSelendokCombo4)
	ON_CBN_SELENDOK(IDC_COMBO5, &CDlg::OnCbnSelendokCombo5)
	ON_CBN_SELENDOK(IDC_COMBO6, &CDlg::OnCbnSelendokCombo6)
	ON_CBN_SELENDOK(IDC_COMBO7, &CDlg::OnCbnSelendokCombo7)
	ON_CBN_SELENDOK(IDC_COMBO8, &CDlg::OnCbnSelendokCombo8)
	ON_BN_CLICKED(IDC_CHECK7, &CDlg::OnBnClickedCheck7)
	ON_BN_CLICKED(IDC_BUTTON3, &CDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON1, &CDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_RADIO3, &CDlg::OnBnClickedRadio3)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"PrnVw.dll"));

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
	strSql = m_strSql;
//AfxMessageBox(strSql);
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
	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);
	if(strSql.Find(_T("QT38w1_1"))!=-1){
		strSql.Replace(_T("QT38w1_1"),_T("QT40"));
		OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
	}

	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);
	strSql = _T("QT47T48_C2 ");
	strSql+= m_strSls;
//AfxMessageBox(strSql);
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);
	strSql = _T("QT10");
//AfxMessageBox(strSql);
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
	if(m_bFndC){
		short fCol=0;
		m_Flg = false;
		OnFindInGrid(m_strFndC1,ptrRs3,fCol,m_Flg);
		m_Flg = true;
	}
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);


	m_ComboBox1.SelectString(0,_T("2"));
	m_ComboBox2.SelectString(0,_T("1"));
	m_ComboBox3.SelectString(0,_T("1"));
	m_ComboBox4.SelectString(0,_T("1"));
	m_ComboBox5.SelectString(0,_T("1"));
	m_ComboBox6.SelectString(0,_T("1"));
	m_ComboBox7.SelectString(0,_T("0"));
	m_ComboBox8.SelectString(0,_T("0"));

	OnEnableCombo();
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
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32774,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32775,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32776,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32777,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32778,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32779,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32780,FALSE);
	pTlBr->GetToolBarCtrl().EnableButton(ID_BUTTON32783,FALSE);

}

void CDlg::OnEnChangeEditTB(void)
{
//	UpdateData();
	CString s,sLstCh,str,msg,strCopy;
	BOOL bDec;
	bDec = FALSE;
//	if(ptrRs1->State==adStateOpen){
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

//	}
	if(IsEnableRec(ptrRs1)){
		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
			if(!str.IsEmpty()){
				m_Flg = false;
				m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}

	if(IsEnableRec(ptrRs2)){
		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			if(!str.IsEmpty()){
				m_Flg = false;
//AfxMessageBox(str);
				m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
				m_Flg = true;
			}
		}
	}

	if(IsEnableRec(ptrRs3)){
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
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_CHECK2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_CHECK3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_CHECK4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_CHECK5)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_CHECK6)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9023);
	GetDlgItem(IDC_CHECK7)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9024);
	GetDlgItem(IDC_CHECK8)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9025);
	GetDlgItem(IDC_STATIC_PRN)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_BUTTON1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9029);
	GetDlgItem(IDC_BUTTON2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9028);
	GetDlgItem(IDC_BUTTON3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9029);
	GetDlgItem(IDC_BUTTON4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9031);
	GetDlgItem(IDC_STATIC_RATE)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9032);
	GetDlgItem(IDC_RADIO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9033);
	GetDlgItem(IDC_RADIO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9034);
	GetDlgItem(IDC_RADIO3)->SetWindowText(sTxt);
}

void CDlg::On32771(void)
{
//Вернуться в корзину
	return;
}

void CDlg::On32772(void)
{
//Сбросить в  базу нак из корзины
	return;
}

void CDlg::On32773(void)
{
//Пересчёт нак 
	return;
}

void CDlg::On32774(void)
{
//Возврат к запросу накладных
	return;
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

void CDlg::On32780(void)
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
	CString strPathOut,strAcc,strDoc,strCR,strCRAdr,strCSA,strCREm;
//	CString strCod;
	COleVariant vC;
	short i;
	strCR = strCRAdr = strCREm = _T("-1");

	CStringArray strDcPt;	//Массив докум. печати
	CStringArray strNZ;		//Массив наряд заданий
	const int iMax = 6;	// 8
	strDcPt.SetSize(iMax);
	int i_aPrn[iMax];
	strNZ.SetSize(1);
	strNZ.SetAt(0,_T(""));

//AfxMessageBox(_T("ptrRs3"));
	if(IsEnableRec(ptrRs3)){
		if(!ptrRs3->adoEOF){

			i = 0;		//Код грузополучателя
			vC = GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			strCR = vC.bstrVal;
			strCR.TrimLeft();
			strCR.TrimRight();
		}
	}

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i = 0;		//Код адреса доставки
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCRAdr = vC.bstrVal;
			strCRAdr.TrimLeft();
			strCRAdr.TrimRight();
		}
	}

	if(IsEnableRec(ptrRs2)){
		if(!ptrRs2->adoEOF){
			i = 0;		//Код расч.сч. продовца
			vC = GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			strCSA = vC.bstrVal;
			strCSA.TrimLeft();
			strCSA.TrimRight();
		}
	}

	if(IsEnableRec(ptrRs4)){
		if(!ptrRs4->adoEOF){
			i = 0;		//Код e-mail
			vC = GetValueRec(ptrRs4,i);
			vC.ChangeType(VT_BSTR);
			strCREm = vC.bstrVal;
			strCREm.TrimLeft();
			strCREm.TrimRight();
		}
	}

	strPathOut = _T("C:\\Program Files\\ORTK\\Strg\\Reports\\");
	strAcc = strCR	 + _T("~");		//Код грузополучателя
	strAcc+= strCRAdr+ _T("~");		//Код адреса доставки
	strAcc+= strCSA	 + _T("~");		//Код расч.сч. продовца
	strAcc+= strCREm;				//Код e-mail
//	UpdateData();
//AfxMessageBox(strAcc);

	int szBuf = WideCharToMultiByte(CP_ACP, 0 ,m_strCod,-1,NULL,0,NULL,NULL);
	char* sACod = new char[szBuf];
	WideCharToMultiByte( CP_ACP, 0, m_strCod, -1,
						sACod, szBuf, NULL, NULL );

	szBuf = WideCharToMultiByte(CP_ACP, 0 ,strPathOut,-1,NULL,0,NULL,NULL);
	char* sAPathOut = new char[szBuf];
	WideCharToMultiByte( CP_ACP, 0, strPathOut, -1,
						sAPathOut, szBuf, NULL, NULL );
/*CString s;
int lw;
lw = MultiByteToWideChar(CP_ACP,0,sAPathOut,-1,NULL,0);
LPWSTR strW = new WCHAR[lw];
MultiByteToWideChar(CP_ACP,0,sAPathOut,-1,strW,lw);

AfxMessageBox(strPathOut);
AfxMessageBox(strW);
*/
	szBuf = WideCharToMultiByte(CP_ACP, 0 ,strAcc,-1,NULL,0,NULL,NULL);
	char* sAAcc = new char[szBuf];
	WideCharToMultiByte( CP_ACP, 0, strAcc, -1,
						sAAcc, szBuf, NULL, NULL );

	WCHAR* sANZ    = new WCHAR[1024];
	WCHAR* strWDoc = new WCHAR[1024];
	HMODULE hMod;
//Создаём файлы(выходных форм)
	typedef BOOL (*pDlg)(CString,_ConnectionPtr,char*,char*,char*, WCHAR* );


	if(m_Check1){	//Товарная накладная
		hMod=AfxLoadLibrary(_T("RprInvOrd.dll"));
		pDlg func=(pDlg)GetProcAddress(hMod,"startRprInvOrd");

		BeginWaitCursor();
			(func)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, strWDoc);
			strDcPt.SetAt(0,strWDoc);
//			AfxMessageBox(strDcPt.GetAt(0));
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}

//AfxMessageBox(_T("Товарная накладная"));

	if(m_Check2){	//Счёт фактура
		hMod=AfxLoadLibrary(_T("RprInv.dll"));
		pDlg func=(pDlg)GetProcAddress(hMod,"startRprInv");

		BeginWaitCursor();
			(func)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, strWDoc);
			strDcPt.SetAt(1,strWDoc);
//			AfxMessageBox(strDcPt.GetAt(1));
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}
//AfxMessageBox(_T("Счёт фактура"));

	if(m_Check3){	//Старый вид накладной
		hMod=AfxLoadLibrary(_T("RprOldInvOrd.dll"));
		pDlg func=(pDlg)GetProcAddress(hMod,"startRprOldInvOrd");

		BeginWaitCursor();
			(func)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, strWDoc);
			strDcPt.SetAt(2,strWDoc);
//			AfxMessageBox(strDcPt.GetAt(2));
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}
//AfxMessageBox(_T("Старый вид накладной"));

	if(m_Check4){	//Приложенине
		hMod=AfxLoadLibrary(_T("RprEncl.dll"));
		pDlg func=(pDlg)GetProcAddress(hMod,"startRprEncl");

		BeginWaitCursor();
			(func)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, strWDoc);
			strDcPt.SetAt(3,strWDoc);
//			AfxMessageBox(strDcPt.GetAt(3));
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}
//AfxMessageBox(_T("Приложенине"));

	typedef BOOL (*pDlg1)(CString,_ConnectionPtr, char*, char*, char*, WCHAR*);

	if(m_Check5){	//Наряд задание
		CString strCurNZ;
		int iNZ=0;
		int iLn,iB;
		hMod=AfxLoadLibrary(_T("RprCmdTsk.dll"));
		pDlg1 func1=(pDlg1)GetProcAddress(hMod,"startRprCmdTsk");

		BeginWaitCursor();
			(func1)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, sANZ);
			strCurNZ = sANZ;
			iLn = strCurNZ.GetLength();
			for(int i = 0;i<=iLn;i++){
				if(strCurNZ.GetAt(i)==_T('~')){
					if(iNZ==0){
						strNZ.SetAt(iNZ,strCurNZ.Left(i));
						iB = i+1;
					}
					else{
						strNZ.SetAtGrow(iNZ,strCurNZ.Mid(iB,i-iB));
						iB = i+1;
					}
					iNZ++;
				}
				else{
					if(iNZ==0){
						strNZ.SetAt(iNZ,strCurNZ);
					}
					else{
						if(i==iLn){
							strNZ.SetAtGrow(iNZ,strCurNZ.Mid(iB));
							iNZ++;
						}
					}
				}
			}
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}
//AfxMessageBox(_T("Наряд задание"));

	if(m_Check6){	//Счёт фактура
		hMod=AfxLoadLibrary(_T("RprBill.dll"));
		pDlg func=(pDlg)GetProcAddress(hMod,"startRprBill");

		BeginWaitCursor();
			(func)(m_strNT, ptrCnn, sACod, sAPathOut, sAAcc, strWDoc);
			strDcPt.SetAt(5,strWDoc);
//			AfxMessageBox(strDcPt.GetAt(5));
		EndWaitCursor();

		AfxFreeLibrary(hMod);
	}
//AfxMessageBox(_T("Счёт фактура"));

	delete[] sACod;
	delete[] sAPathOut;
	delete[] sAAcc;
	delete[] sANZ;
	delete[] strWDoc;

//AfxMessageBox(_T("delete "));
// Печать,просмотр,е-mail
	CString sCp;
	COleVariant covTrue((short)TRUE),
				covFalse((short)FALSE),
				covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	_ApplicationPtr ptrWord;
	DocumentsPtr    ptrDocS;
	_DocumentPtr    ptrDoc;
	HRESULT hr;
	switch(m_Radio1){
		case 0:	//Печать
			{
				for(int k=0;k<iMax;k++){	//Опр. кол-во копий печати
					switch(k){
						case 0:	//Тов. нак.
							m_ComboBox1.GetLBText(m_ComboBox1.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
						case 1: //Сч. факт.
							m_ComboBox2.GetLBText(m_ComboBox2.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
						case 2: //Стар нак.
							m_ComboBox3.GetLBText(m_ComboBox3.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
						case 3: //Прилож.
							m_ComboBox4.GetLBText(m_ComboBox4.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
						case 4: //наряд.зад.
							m_ComboBox5.GetLBText(m_ComboBox5.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
						case 5: //Счёт
							m_ComboBox6.GetLBText(m_ComboBox6.GetCurSel(),sCp);
							*(i_aPrn+k)= _wtoi(sCp);
							break;
					}
				}
/*				_ApplicationPtr ptrWord;
				DocumentsPtr    ptrDocS;
				_DocumentPtr    ptrDoc;
				HRESULT hr;
*/				hr = ptrWord.CreateInstance(__uuidof(Word::Application));

				if(FAILED(hr)){
					AfxMessageBox(_T("Обьект не создан"));
				}
				else{
					try{
						ptrDocS = ptrWord->Documents;
						for(int k=0;k<iMax;k++){				//"c:\\Test.rtf"
							if(*(i_aPrn+k)!=0){
								if(k!=4){	//Не наряд задания
									ptrDoc  = ptrDocS->Add(&_variant_t(strDcPt.GetAt(k)));
									ptrDoc->PrintOut(covFalse,     // Background.
											covOptional,           // Append.
											covOptional,           // Range.
											covOptional,           // OutputFileName.
											covOptional,           // From.
											covOptional,           // To.
											covOptional,           // Item.
											COleVariant((long)(*(i_aPrn+k)) ),  // Copies.
											covOptional,           // Pages.
											covOptional,           // PageType.
											covOptional,           // PrintToFile.
											covOptional,           // Collate.
											covOptional,           // ActivePrinterMacGX.
											covOptional,           // ManualDuplexPrint.
											covOptional,
											covOptional,
											covOptional,
											covOptional
											);
									ptrDoc->Close();
								}
								else{	//наряд задания
									int szNZ = strNZ.GetSize();
									for(int j = 0; j<szNZ; j++){
										ptrDoc  = ptrDocS->Add(&_variant_t(strNZ.GetAt(j)));
										ptrDoc->PrintOut(covFalse,     // Background.
												covOptional,           // Append.
												covOptional,           // Range.
												covOptional,           // OutputFileName.
												covOptional,           // From.
												covOptional,           // To.
												covOptional,           // Item.
												COleVariant((long)(*(i_aPrn+k)) ),  // Copies.
												covOptional,           // Pages.
												covOptional,           // PageType.
												covOptional,           // PrintToFile.
												covOptional,           // Collate.
												covOptional,           // ActivePrinterMacGX.
												covOptional,           // ManualDuplexPrint.
												covOptional,
												covOptional,
												covOptional,
												covOptional
												);
										ptrDoc->Close();
									}
								}
							}
						}
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
						ptrWord->Quit();
						OnCancel();
					}
					ptrWord->Quit();
				}
			}
			break;
		case 1:	// Просмотр
			{
				for(int k=0;k<iMax;k++){	//Опр. кол-во копий печати
					*(i_aPrn+k) = 1;
				}
				hr = ptrWord.CreateInstance(__uuidof(Word::Application));
				if(FAILED(hr)){
					AfxMessageBox(_T("Обьект не создан"));
				}
				else{
					try{
						ptrDocS = ptrWord->Documents;
						ptrDoc  = ptrDocS->Add(&_variant_t(strDcPt.GetAt(0)));
						ptrWord->Visible = true;
						ptrDoc->Activate();

						if(ptrDoc->GetApplication()==ptrWord){
							AfxMessageBox(IDS_STRING9037,MB_ICONINFORMATION);
							ptrDoc->Close();
							ptrWord->Quit();
						}
					}
					catch(_com_error& e){
//AfxMessageBox(_T("error"));
/*						AfxMessageBox(e.Description());
						ptrDoc->Close();
						ptrWord->Quit();
*/
						OnCancel();
					}
				}
			}
			break;
		case 2:	// e-mail
			break;
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
				GrdClms.GetItem((COleVariant) i).SetVisible(FALSE);
				break;
			case 1:
				wdth = 400;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 70;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
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

void CDlg::InitDataGrid2(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9026);
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
				wdth = 160;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 290;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 160;
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

	strCap.LoadString(IDS_STRING9027);
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
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
/*			case 1:
				wdth = 100;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
*/
			case 2:
				wdth = 340;
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

void CDlg::InitDataGrid4(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING9030);
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
				wdth = 230;
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
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, 218, CDlg::RowColChangeDatagrid2, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID3, 218, CDlg::RowColChangeDatagrid3, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID4, DISPID_CLICK, CDlg::ClickDatagrid4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID4, 218, CDlg::RowColChangeDatagrid4, VTS_PVARIANT VTS_I2)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
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

void CDlg::OnBnClickedRadio1()
{
	UpdateData();
	OnEnableState();
	OnEnableCombo();
}

void CDlg::OnBnClickedRadio2()
{
	UpdateData();
	OnEnableCombo();
	OnEnableState();
}
void CDlg::OnBnClickedRadio3()
{
	UpdateData();
	OnEnableCombo();
	OnEnableState();
}

void CDlg::OnBnClickedCheck1()
{
	OnEnableCombo();	
	if(m_Check1){
		m_ComboBox1.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox1.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck2()
{
	OnEnableCombo();	
	if(m_Check2){
		m_ComboBox2.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox2.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck3()
{
	OnEnableCombo();	
	if(m_Check3){
		m_ComboBox3.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox3.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck4()
{
	OnEnableCombo();	
	if(m_Check4){
		m_ComboBox4.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox4.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck5()
{
	OnEnableCombo();	
	if(m_Check5){
		m_ComboBox5.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox5.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck6()
{
	OnEnableCombo();	
	if(m_Check6){
		m_ComboBox6.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox6.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck7()
{
	OnEnableCombo();	
	if(m_Check7){
		m_ComboBox7.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox7.SelectString(0,_T("0"));
	}
}

void CDlg::OnBnClickedCheck8()
{
	OnEnableCombo();	
	if(m_Check8){
		m_ComboBox8.SelectString(0,_T("1"));
	}
	else{
		m_ComboBox8.SelectString(0,_T("0"));
	}
}

void CDlg::OnCbnEditchangeCombo1()
{
	CString s;
	GetDlgItem(IDC_COMBO1)->GetWindowText(s);
	if(s==_T("0")){
		m_Check1 = false;
		((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
		OnBnClickedCheck1();
	}
}

void CDlg::OnCbnEditchangeCombo2()
{
	CString s;
	GetDlgItem(IDC_COMBO2)->GetWindowText(s);
	if(s==_T("0")){
		m_Check2 = false;
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(m_Check2);
		OnBnClickedCheck2();
	}
}

void CDlg::OnCbnEditchangeCombo3()
{
	CString s;
	GetDlgItem(IDC_COMBO3)->GetWindowText(s);
	if(s==_T("0")){
		m_Check3 = false;
		((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(m_Check3);
		OnBnClickedCheck3();
	}
}


void CDlg::OnCbnEditchangeCombo4()
{
	CString s;
	GetDlgItem(IDC_COMBO4)->GetWindowText(s);
	if(s==_T("0")){
		m_Check4 = false;
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(m_Check4);
		OnBnClickedCheck4();
	}
}

void CDlg::OnCbnEditchangeCombo5()
{
	CString s;
	GetDlgItem(IDC_COMBO5)->GetWindowText(s);
	if(s==_T("0")){
		m_Check5 = false;
		((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(m_Check5);
		OnBnClickedCheck5();
	}
}

void CDlg::OnCbnEditchangeCombo6()
{
	CString s;
	GetDlgItem(IDC_COMBO6)->GetWindowText(s);
	if(s==_T("0")){
		m_Check6 = false;
		((CButton*)GetDlgItem(IDC_CHECK6))->SetCheck(m_Check6);
		OnBnClickedCheck6();
	}
}

void CDlg::OnCbnEditchangeCombo7()
{
	CString s;
	GetDlgItem(IDC_COMBO7)->GetWindowText(s);
	if(s==_T("0")){
		m_Check6 = false;
		((CButton*)GetDlgItem(IDC_CHECK7))->SetCheck(m_Check7);
		OnBnClickedCheck7();
	}
}

void CDlg::OnCbnEditchangeCombo8()
{
	CString s;
	GetDlgItem(IDC_COMBO8)->GetWindowText(s);
	if(s==_T("0")){
		m_Check8 = false;
		((CButton*)GetDlgItem(IDC_CHECK8))->SetCheck(m_Check8);
		OnBnClickedCheck8();
	}
}

void CDlg::OnCbnSelendokCombo1()
{
	CString s;
	int i;
	i = m_ComboBox1.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check1 = false;
		((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
		OnBnClickedCheck1();
	}
	else{
		m_ComboBox1.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo2()
{
	CString s;
	int i;
	i = m_ComboBox2.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check2 = false;
		((CButton*)GetDlgItem(IDC_CHECK2))->SetCheck(m_Check2);
		OnBnClickedCheck2();
	}
	else{
		m_ComboBox2.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo3()
{
	CString s;
	int i;
	i = m_ComboBox3.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check3 = false;
		((CButton*)GetDlgItem(IDC_CHECK3))->SetCheck(m_Check3);
		OnBnClickedCheck3();
	}
	else{
		m_ComboBox3.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo4()
{
	CString s;
	int i;
	i = m_ComboBox4.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check4 = false;
		((CButton*)GetDlgItem(IDC_CHECK4))->SetCheck(m_Check4);
		OnBnClickedCheck4();
	}
	else{
		m_ComboBox4.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo5()
{
	CString s;
	int i;
	i = m_ComboBox5.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check5 = false;
		((CButton*)GetDlgItem(IDC_CHECK5))->SetCheck(m_Check5);
		OnBnClickedCheck5();
	}
	else{
		m_ComboBox5.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo6()
{
	CString s;
	int i;
	i = m_ComboBox6.GetCurSel();
	s.Format(_T("%i"),i);
	if(s==_T("0")){
		m_Check6 = false;
		((CButton*)GetDlgItem(IDC_CHECK6))->SetCheck(m_Check6);
		OnBnClickedCheck6();
	}
	else{
		m_ComboBox6.SelectString(0,s);
	}
}

void CDlg::OnCbnSelendokCombo7()
{
	// TODO: Add your control notification handler code here
}

void CDlg::OnCbnSelendokCombo8()
{
	// TODO: Add your control notification handler code here
}

void CDlg::OnEnableCombo(void)
{
	UpdateData();
	GetDlgItem(IDC_COMBO1)->EnableWindow(m_Check1);
	GetDlgItem(IDC_COMBO2)->EnableWindow(m_Check2);
	GetDlgItem(IDC_COMBO3)->EnableWindow(m_Check3);
	GetDlgItem(IDC_COMBO4)->EnableWindow(m_Check4);
	GetDlgItem(IDC_COMBO5)->EnableWindow(m_Check5);
	GetDlgItem(IDC_COMBO6)->EnableWindow(m_Check6);
	GetDlgItem(IDC_COMBO7)->EnableWindow(m_Check7);
	GetDlgItem(IDC_COMBO8)->EnableWindow(m_Check8);

}


void CDlg::OnEnableState(void)
{
	CString sTxt;
	switch(m_Radio1){
	case 0:
		sTxt.LoadStringW(IDS_STRING9032);
		GetDlgItem(IDC_STATIC_PRN)->SetWindowText(sTxt); //_T("Печать")
		GetDlgItem(IDC_CHECK1)->EnableWindow(true);
		GetDlgItem(IDC_CHECK2)->EnableWindow(true);
		GetDlgItem(IDC_CHECK3)->EnableWindow(true);
		GetDlgItem(IDC_CHECK4)->EnableWindow(true);
		GetDlgItem(IDC_CHECK5)->EnableWindow(true);
		GetDlgItem(IDC_CHECK6)->EnableWindow(true);
		GetDlgItem(IDC_CHECK7)->EnableWindow(true);
		GetDlgItem(IDC_CHECK8)->EnableWindow(true);

		m_ComboBox1.SelectString(0,_T("2"));
		m_ComboBox2.SelectString(0,_T("1"));
		m_ComboBox3.SelectString(0,_T("1"));
		m_ComboBox4.SelectString(0,_T("1"));
		m_ComboBox5.SelectString(0,_T("1"));
		m_ComboBox6.SelectString(0,_T("1"));
		m_ComboBox7.SelectString(0,_T("0"));
		m_ComboBox8.SelectString(0,_T("1"));
		break;
	case 1:
		sTxt.LoadStringW(IDS_STRING9033);
		GetDlgItem(IDC_STATIC_PRN)->SetWindowText(sTxt);	//_T("Просмотр/Печать")
		GetDlgItem(IDC_CHECK1)->EnableWindow(false);
		GetDlgItem(IDC_CHECK2)->EnableWindow(false);
		GetDlgItem(IDC_CHECK3)->EnableWindow(false);
		GetDlgItem(IDC_CHECK4)->EnableWindow(false);
		GetDlgItem(IDC_CHECK5)->EnableWindow(false);
		GetDlgItem(IDC_CHECK6)->EnableWindow(false);
		GetDlgItem(IDC_CHECK7)->EnableWindow(false);
		GetDlgItem(IDC_CHECK8)->EnableWindow(false);

		m_ComboBox1.SelectString(0,_T("1"));
		m_ComboBox2.SelectString(0,_T("1"));
		m_ComboBox3.SelectString(0,_T("1"));
		m_ComboBox4.SelectString(0,_T("1"));
		m_ComboBox5.SelectString(0,_T("1"));
		m_ComboBox6.SelectString(0,_T("1"));
		m_ComboBox7.SelectString(0,_T("0"));
		m_ComboBox8.SelectString(0,_T("0"));
		break;
	case 2:
		sTxt.LoadStringW(IDS_STRING9034);
		GetDlgItem(IDC_STATIC_PRN)->SetWindowText(sTxt);
		GetDlgItem(IDC_CHECK1)->EnableWindow(true);
		GetDlgItem(IDC_CHECK2)->EnableWindow(true);
		GetDlgItem(IDC_CHECK3)->EnableWindow(true);
		GetDlgItem(IDC_CHECK4)->EnableWindow(true);
		GetDlgItem(IDC_CHECK5)->EnableWindow(true);
		GetDlgItem(IDC_CHECK6)->EnableWindow(true);
		GetDlgItem(IDC_CHECK7)->EnableWindow(false);
		GetDlgItem(IDC_CHECK8)->EnableWindow(false);

		m_ComboBox1.SelectString(0,_T("1"));
		m_ComboBox2.SelectString(0,_T("1"));
		m_ComboBox3.SelectString(0,_T("1"));
		m_ComboBox4.SelectString(0,_T("1"));
		m_ComboBox5.SelectString(0,_T("1"));
		m_ComboBox6.SelectString(0,_T("1"));
		m_ComboBox7.SelectString(0,_T("0"));
		m_ComboBox8.SelectString(0,_T("0"));
		break;
	}
}

void CDlg::ClickDatagrid3()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID3);
}

void CDlg::ClickDatagrid1()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
}

void CDlg::ClickDatagrid2()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID2);
}

void CDlg::RowColChangeDatagrid2(VARIANT* LastRow, short LastCol)
{
	if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
		if(IsEnableRec(ptrRs2)){
			if(!ptrRs2->adoEOF){
					m_CurCol = m_DataGrid2.get_Col();
					if(m_CurCol==-1 || m_CurCol== 0 ){
						m_CurCol=1;
					}
					m_iCurType = GetTypeCol(ptrRs2,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
			}
		}
	}
}

void CDlg::RowColChangeDatagrid3(VARIANT* LastRow, short LastCol)
{
	if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
		if(IsEnableRec(ptrRs3)){
			if(!ptrRs3->adoEOF){
					m_CurCol = m_DataGrid3.get_Col();
					if(m_CurCol==-1 ){
						m_CurCol=0;
					}
					m_iCurType = GetTypeCol(ptrRs3,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
					short i;
					COleVariant vC;
					CString strC,strSql;

					i = 0;	//код конрагента
					vC = GetValueRec(ptrRs3,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;
					strC.TrimLeft();
					strC.TrimRight();

					strSql = _T("QT38w1_1 ");
					strSql+= strC;

					OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);

					strSql = _T("QT40 ");
					strSql+= strC;

					OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
					

			}
		}
	}
}

void CDlg::ClickDatagrid4()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID4);
}

void CDlg::RowColChangeDatagrid4(VARIANT* LastRow, short LastCol)
{
	if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
		if(IsEnableRec(ptrRs4)){
			if(!ptrRs4->adoEOF){
					m_CurCol = m_DataGrid4.get_Col();
					if(m_CurCol==-1 || m_CurCol== 0 ){
						m_CurCol=1;
					}
					m_iCurType = GetTypeCol(ptrRs4,m_CurCol);
					m_EditTBCh.SetTypeCol(m_iCurType);
			}
		}
	}
}

void CDlg::OnBnClickedButton3()
{
//Добавить e-mail
//	UpdateData();
	CString s,strC,strSql;
	short i;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	if(IsEnableRec(ptrRs3)){
		if(!ptrRs3->adoEOF){
			GetDlgItem(IDC_EDIT1)->GetWindowTextW(m_Edit1);
			m_Edit1.TrimLeft();
			m_Edit1.TrimRight();
			if(!m_Edit1.IsEmpty()){
				if(m_Edit1.Find(_T("@"))!=-1){
//					AfxMessageBox(_T("m_Edit11.Find(_T(@))!=-1"));
					i = 0;
					vC=GetValueRec(ptrRs3,i);
					vC.ChangeType(VT_BSTR);
					strC=vC.bstrVal;
					strC.TrimRight(' ');


					strSql = _T("IT40 ");
					strSql+= strC + _T(",'");
					strSql+= m_Edit1 + _T("','");
					strSql+= m_strNT +_T("'");
			//AfxMessageBox(strSql);
					ptrCmd4->CommandText = (_bstr_t)strSql;
					try{
			//			AfxMessageBox(strSql);
						ptrCmd4->Execute(&vra,vtl,adCmdText);

					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					strSql = _T("QT40 ");
					strSql+= strC;

					OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
				}
				else{
					s.LoadStringW(IDS_STRING9035);
					AfxMessageBox(s,MB_ICONINFORMATION);
				}
			}
		}
	}
}

void CDlg::OnBnClickedButton4()
{
//удалить е-mail
	CString s,strC,strSql;
	short i;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	if(IsEnableRec(ptrRs4)){
		if(!ptrRs4->adoEOF){
			i = 0;
			vC=GetValueRec(ptrRs4,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight(' ');
			
			strSql = _T("DT40 ");
			strSql+= strC;
			ptrCmd4->CommandText = (_bstr_t)strSql;
			try{
	//			AfxMessageBox(strSql);
				ptrCmd4->Execute(&vra,vtl,adCmdText);

			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}

			if(IsEnableRec(ptrRs3)){
				if(!ptrRs3->adoEOF){
					i = 0;
					vC=GetValueRec(ptrRs3,i);
					vC.ChangeType(VT_BSTR);
					strC=vC.bstrVal;
					strC.TrimRight(' ');

					strSql = _T("QT40 ");
					strSql+= strC;

					OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);
				}
			}
		}
	}
}

void CDlg::OnBnClickedButton1()
{
//добавить адрес доставки
	BeginWaitCursor();
	if(IsEnableRec(ptrRs3)){
		if(!ptrRs3->adoEOF){
			CString strCod,strIns,strSql;
			COleVariant vC;
			strCod = _T("0");
			strIns = _T("IT38 ");
			short i=0;
			vC = GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			strCod=vC.bstrVal;
			strCod.TrimLeft(' ');
			strCod.TrimRight(' ');
	//		m_SlpDay.SetDate(t1);

			HMODULE hMod;
			hMod=AfxLoadLibrary(_T("AddBk.dll"));
			typedef BOOL (*pDialog)( CString, _ConnectionPtr,CString,CString );
			pDialog func=(pDialog)GetProcAddress(hMod,"startAddBk");
			(func)(m_strNT, ptrCnn,strCod,strIns);
			
			strSql = m_strSql;
		//AfxMessageBox(strSql);
			OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);

	//		m_SlpDay.SetDate(t1);

			AfxFreeLibrary(hMod);
		}
	}
	EndWaitCursor();
}

void CDlg::OnBnClickedButton2()
{
//удалить адрес доставки
	if(IDYES==AfxMessageBox(IDS_STRING9036,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
		COleVariant vC,vBk;
		CString strC,strSql;
		_variant_t vra;
		VARIANT* vtl = NULL;
		short i = 0;
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				vBk = ptrRs1->Bookmark;

				vC=GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC=vC.bstrVal;

				strSql = _T("DT38 ");
				strSql+=strC;
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
				if(IsEnableRec(ptrRs3)){
					i = 0;
					vC = GetValueRec(ptrRs3,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;
					strC.TrimRight();
				}
				strSql = _T("QT38w1_1 ");
				strSql+= strC;
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
		}
	}
}

