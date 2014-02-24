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

	, m_Edit4(_T(""))
	, m_Edit2(_T(""))
	, m_Edit3(_T(""))
	, m_Edit1(_T(""))
	, m_Edit5(_T(""))
	, m_Edit6(_T(""))
	, m_cd(0)
	, m_Check1(TRUE)
	, m_Check2(FALSE)
{

}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen)	ptrRs1->Close();
	ptrRs1=NULL;
	if(ptrRs2->State==adStateOpen)	ptrRs2->Close();
	ptrRs2=NULL;
	if(ptrRs3->State==adStateOpen)	ptrRs3->Close();
	ptrRs3=NULL;
	if(ptrRs4->State==adStateOpen)	ptrRs4->Close();
	ptrRs4=NULL;
	if(ptrRs5->State==adStateOpen)	ptrRs5->Close();
	ptrRs5=NULL;

	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;
	if(ptrRsCmb2->State==adStateOpen) ptrRsCmb2->Close();
	ptrRsCmb2 = NULL;
	if(ptrRsCmb3->State==adStateOpen) ptrRsCmb3->Close();
	ptrRsCmb3 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Text(pDX, IDC_EDIT4, m_Edit4);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT5, m_Edit5);
	DDX_Text(pDX, IDC_EDIT6, m_Edit6);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_Control(pDX, IDC_DATACOMBO2, m_DataCombo2);
	DDX_Control(pDX, IDC_DATACOMBO3, m_DataCombo3);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATAGRID4, m_DataGrid4);
	DDX_Control(pDX, IDC_DATAGRID5, m_DataGrid5);
	DDX_Control(pDX, IDC_DATAGRID6, m_DataGrid6);
	DDX_Check(pDX, IDC_CHECK1, m_Check1);
	DDX_Check(pDX, IDC_CHECK2, m_Check2);
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
	ON_BN_CLICKED(IDC_BUTTON6, &CDlg::OnBnClickedButton6)
	ON_BN_CLICKED(IDC_BUTTON7, &CDlg::OnBnClickedButton7)
	ON_BN_CLICKED(IDC_BUTTON8, &CDlg::OnBnClickedButton8)
	ON_BN_CLICKED(IDC_BUTTON9, &CDlg::OnBnClickedButton9)
	ON_BN_CLICKED(IDC_BUTTON10, &CDlg::OnBnClickedButton10)
	ON_BN_CLICKED(IDC_BUTTON11, &CDlg::OnBnClickedButton11)
	ON_BN_CLICKED(IDC_BUTTON12, &CDlg::OnBnClickedButton12)
	ON_BN_CLICKED(IDC_BUTTON13, &CDlg::OnBnClickedButton13)
	ON_BN_CLICKED(IDC_CHECK1, &CDlg::OnBnClickedCheck1)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"CtrAgnt.dll"));

	CString s,strSql,strSql2;
	s.LoadString(IDS_STRING12001);
	this->SetWindowText(s);

	m_Edit1RsNm.SubclassDlgItem(IDC_EDIT1,this);
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
	strSql = _T("QT10");
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1,0,0);
	if(m_bFndC){
		short fCol=0;
		m_Flg = false;
		OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
		m_Flg = true;
	}

	ptrCmdCmb1 = NULL;
	ptrRsCmb1  = NULL;

	strSql = _T("QT9");
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
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
	if(IsEnableRec(ptrRs1)){
		i=2;
		CString strBoundText;
		vC=GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strBoundText=vC.bstrVal;

//AfxMessageBox(strBoundText);

		m_DataCombo1.put_ListField(_T("Сокр. наим. ф.собст-ти"));
	}

	if(IsEnableRec(ptrRs1)){
		i = 6;
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_I4);
		m_cd = vC.intVal;

		OnShowEnter(m_cd);
		m_Check1 = true;
		((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
		strSql = OnQuerySex(m_Check1,strSql2);
	}
	ptrCmdCmb2 = NULL;
	ptrRsCmb2  = NULL;

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
		m_DataCombo2.put_ListField(_T("Имя"));
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
	}

	ptrCmdCmb3 = NULL;
	ptrRsCmb3  = NULL;

	ptrCmdCmb3.CreateInstance(__uuidof(Command));
	ptrCmdCmb3->ActiveConnection = ptrCnn;
	ptrCmdCmb3->CommandType = adCmdText;
	ptrCmdCmb3->CommandText = (_bstr_t)strSql2;

	ptrRsCmb3.CreateInstance(__uuidof(Recordset));
	ptrRsCmb3->CursorLocation = adUseClient;
	ptrRsCmb3->PutRefSource(ptrCmdCmb3);
	try{
		m_DataCombo3.putref_RowSource(NULL);
		ptrRsCmb3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdUnknown);
		m_DataCombo3.putref_RowSource((LPUNKNOWN) ptrRsCmb3);
		m_DataCombo3.put_ListField(_T("Отчество"));
	}
	catch(_com_error& e){
		m_DataCombo3.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}


	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;
	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;
	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;
	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);

	ptrCmd5.CreateInstance(__uuidof(Command));
	ptrCmd5->ActiveConnection = ptrCnn;
	ptrCmd5->CommandType = adCmdText;
	ptrRs5.CreateInstance(__uuidof(Recordset));
	ptrRs5->CursorLocation = adUseClient;
	ptrRs5->PutRefSource(ptrCmd5);

	ptrCmd6.CreateInstance(__uuidof(Command));
	ptrCmd6->ActiveConnection = ptrCnn;
	ptrCmd6->CommandType = adCmdText;
	ptrRs6.CreateInstance(__uuidof(Recordset));
	ptrRs6->CursorLocation = adUseClient;
	ptrRs6->PutRefSource(ptrCmd6);

	InitDataGrid2(m_DataGrid2,ptrRs2);
	InitDataGrid3(m_DataGrid1,ptrRs3);
	InitDataGrid4(m_DataGrid4,ptrRs4);
	InitDataGrid5(m_DataGrid5,ptrRs5);
	InitDataGrid6(m_DataGrid6,ptrRs6);
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
		if(m_CurCol==-1 /*|| m_CurCol==0*/){
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
				msg.LoadString(IDS_STRING12019);
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
					msg.LoadString(IDS_STRING12019);
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
//		m_Flg = false;
		m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
		m_Flg = true;

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount,strC;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING12018);

	GotoDlgCtrl(GetDlgItem(IdBeg));
	do{
		Id = GetFocus()->GetDlgCtrlID();
		switch(Id){
			case IDC_EDIT1:
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
			case IDC_EDIT2:	// ОКОНХ
				{
					m_Edit2.TrimLeft();
					m_Edit2.TrimRight();
					switch(m_cd){
					case 1:
					case 11:
					case 12:
					case 15:
						m_Edit2 = _T("0");
						break;
					default:
						if(m_Edit2.IsEmpty()){
							count++;
							GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
							strCount.Format(_T("%i"),count);
							strCount+=_T(") ");
							strMsg+=strCount+strItem+_T("\n\t");
						}
						break;
					}
				}
				break;
			case IDC_EDIT3:	//ОКПО
				{
					m_Edit3.TrimLeft();
					m_Edit3.TrimRight();
					switch(m_cd){
					case 1:
					case 11:
					case 12:
					case 15:
						m_Edit3 = _T("0");
						break;
					default:
						if(m_Edit3.IsEmpty()){
							count++;
							GetDlgItem(IDC_STATIC_EDIT3)->GetWindowText(strItem);
							strCount.Format(_T("%i"),count);
							strCount+=_T(") ");
							strMsg+=strCount+strItem+_T("\n\t");
						}
						break;
					}
				}
				break;
			case IDC_EDIT4:		// ИНН/КПП
				m_Edit4.TrimLeft();
				m_Edit4.TrimRight();
				if(m_Edit4.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT4)->GetWindowText(strItem);
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
		GetDlgItem(IDC_STATIC_COMBOBOX1)->GetWindowText(strItem);
		strCount.Format(_T("%i"),count);
		strCount+=_T(") ");
		strMsg+=strCount+strItem+_T("\n\t");
	}

	switch(m_cd){
	case 1:		//ПБЮЛ
	case 11:	//ЧП
	case 12:	//ИЧП
	case 15:	//ИП
		{
			strC = m_DataCombo2.get_BoundText();

			strC.TrimRight();
			strC.TrimLeft();
			if(strC==' ') strC.Empty();
//AfxMessageBox(L"2 "+strC);
			if(strC.IsEmpty()){
				count++;
				GetDlgItem(IDC_STATIC_COMBOBOX2)->GetWindowText(strItem);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}

			strC = m_DataCombo3.get_BoundText();

			strC.TrimRight();
			strC.TrimLeft();
			if(strC==' ') strC.Empty();
//AfxMessageBox(L"3 "+strC);
			if(strC.IsEmpty()){
				count++;
				GetDlgItem(IDC_STATIC_COMBOBOX3)->GetWindowText(strItem);
				strCount.Format(_T("%i"),count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		break;
	}
	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	CString strC,strBoundText;
	short i;

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i=0;	//КодКонтрАгента
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
		//	AfxMessageBox( strC+" OnShowEdit");

			i=6;	//Код Ф.Собственности
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_I4);
			m_cd=vC.intVal;
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;

			OnFindInCombo(strC,&m_DataCombo1,ptrRsCmb1,0,1);
			OnShowEnter(m_cd);

			i=2;	//Наименование
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit1=vC.bstrVal;
			m_Edit1.TrimLeft(' ');
			m_Edit1.Replace('.',' ');
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

			i=3;	//Инн/Кпп
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit4=vC.bstrVal;
			GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit4);

			i=4;	//ОКНХ
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit2=vC.bstrVal;
			GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

			i=5;	//ОКПО
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			m_Edit3=vC.bstrVal;
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);


			int z1,z2,z3,lenZ;
			CString strEdit1F,strEdit1S,strEdit1T,strZ;
			switch(m_cd){
			case 1:		//ПБЮЛ
			case 11:	//ЧП
			case 12:	//ИЧП
			case 15:	//ИП
				i=7;	//Пол
				vC=GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BOOL);
				m_Check1 = vC.boolVal;
				((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);

				OnShowNameCombo();

				lenZ = m_Edit1.GetLength();
				z1 = m_Edit1.Find(' ',0);
				strEdit1F=m_Edit1.Left(z1);

				z2 = m_Edit1.Find(' ',z1+1);
				strEdit1S=m_Edit1.Mid(z1+1,z2-z1);
				strEdit1S.TrimRight();
				strEdit1S.TrimLeft();
				strBoundText=strEdit1S;
//AfxMessageBox(strBoundText);
				//Имя
				OnFindInCombo(strBoundText,&m_DataCombo2,ptrRsCmb2,1,1);

				z3 = m_Edit1.Find(' ',z2+1);
				strEdit1T = m_Edit1.Mid(z2+1,z3-z1);
				strEdit1T.TrimRight();
				strEdit1T.TrimLeft();
				strBoundText=strEdit1T;
//AfxMessageBox(strBoundText);

				//Отчество
				OnFindInCombo(strBoundText,&m_DataCombo3,ptrRsCmb3,1,1);
				m_Edit1 = strEdit1F;
				GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

				break;
			}
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING12006);
	GetDlgItem(IDC_STATIC_COMBOBOX1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12007);
	GetDlgItem(IDC_CHECK1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12008);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12009);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12010);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12011);
	GetDlgItem(IDC_STATIC_EDIT4)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12012);
	GetDlgItem(IDC_STATIC_COMBOBOX2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12013);
	GetDlgItem(IDC_STATIC_COMBOBOX3)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12016);
	GetDlgItem(IDC_CHECK2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING12005);
	GetDlgItem(IDC_STATIC_EDIT6)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
	UpdateData();
	CString strSql,strC,strC2,strCh1,strEdit1F,strEdit1S;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i,IndCol;
	COleVariant vC;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	if(m_Check1){
		strCh1 = _T("1");
	}
	else{
		strCh1 = _T("0");
	}

	long IdEnd;

	switch(m_cd){
	case 1:		//ПБЮЛ
	case 11:	//ЧП
	case 12:	//ИЧП
	case 15:	//ИП
		IdEnd = IDC_BUTTON3;
		strEdit1F = m_DataCombo2.get_BoundText();
		strEdit1F.TrimRight(' ');
		strEdit1S = m_DataCombo3.get_BoundText();
		strEdit1S.TrimRight(' ');
		m_Edit1 += _T(" ") + strEdit1F;
		m_Edit1 += _T(" ") + strEdit1S;
		m_Edit3 = _T("0");
		m_Edit4 = _T("0");
		break;
	default:
		IdEnd = IDC_EDIT1;	//Наименование
		break;
	}
	if(OnOverEdit(IDC_EDIT1,IdEnd)){

		i = 0;		//Форма собственности
		vC = GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_BSTR);
		strC2 = vC.bstrVal;

		strC2.TrimRight(' ');
		strC2.TrimLeft(' ');

		strSql = _T("IT10 ");
		strSql+= strC2+_T(",'");	//	Код Формы собст.
		strSql+= m_Edit1+_T("','");	//	Наименование
		strSql+= m_Edit4+_T("','");	//	ИНН
		strSql+= m_Edit2+_T("','");	//  ОКНХ
		strSql+= m_Edit3+_T("',");	//	ОКНО
		strSql+= strCh1+_T(",'");	//  М/Ж
		strSql+= m_strNT+_T("'");

		ptrCmd1->CommandText = (_bstr_t)strSql;
		ptrCmd1->CommandType = adCmdText;
		try{
//			AfxMessageBox(strSql);
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		strSql = _T("QT10");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1,0,0);
		if(IsEnableRec(ptrRs1)){
			if(!m_Edit1.IsEmpty()){
				IndCol = 2;
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
	short i=0;
	short IndCol;
	long IdEnd;
	COleVariant vC;//,vBk;
	CString strC,strC2T9,strSql,strCh1,strEdit1F,strEdit1S;
	_variant_t vra;
	VARIANT* vtl = NULL;

	if(m_Check1){
		strCh1 = _T("1");
	}
	else{
		strCh1 = _T("0");
	}
	switch(m_cd){
	case 1:		//ПБЮЛ
	case 11:	//ЧП
	case 12:	//ИЧП
	case 15:	//ИП
		IdEnd = IDC_BUTTON3;
		strEdit1F = m_DataCombo2.get_BoundText();
		strEdit1F.TrimRight();
		strEdit1S = m_DataCombo3.get_BoundText();
		strEdit1S.TrimRight();
		m_Edit1 += _T(" ") + strEdit1F;
		m_Edit1 += _T(" ") + strEdit1S;
		m_Edit3 = _T("0");
		m_Edit4 = _T("0");
		break;
	default:
		IdEnd = IDC_EDIT1;	//Наименование
		break;
	}

	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){

			i = 0;	//Код
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();
			if(OnOverEdit(IDC_EDIT1,IdEnd)){

				i = 0;		//Форма собственности
				vC = GetValueRec(ptrRsCmb1,i);
				vC.ChangeType(VT_BSTR);
				strC2T9 = vC.bstrVal;
				strC2T9.TrimRight(' ');

				strSql = _T("UT10 ");
				strSql+= strC +_T(",");
				strSql+= strC2T9 + _T(",'");	//Форма соб
				strSql+= m_Edit1 + _T("','");	//Наименовнаие
				strSql+= m_Edit4 + _T("','");	//ИНН
				strSql+= m_Edit2 + _T("','");	//ОКОНХ
				strSql+= m_Edit3 + _T("',");	//ОКПО
				strSql+= strCh1+_T(",'");		//М/Ж
				strSql+= m_strNT + _T("'");

				ptrCmd1->CommandText = (_bstr_t)strSql;
				ptrCmd1->CommandType = adCmdText;
				try{
//AfxMessageBox(strSql);
					ptrCmd1->Execute(&vra,vtl,adCmdText);
				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
    			strSql = _T("QT10");
				OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1,0,0);
				if(IsEnableRec(ptrRs1)){
					IndCol = 0;
					m_Flg = false;
					OnFindInGrid(strC,ptrRs1,IndCol,m_Flg);
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

			strSql = _T("DT10 ");
			strSql+= strC;

			ptrCmd1->CommandType = adCmdText;
			ptrCmd1->CommandText = (_bstr_t)strSql;
			try{
				ptrCmd1->Execute(&vra,vtl,adCmdText);
			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
    		strSql = _T("QT10");
			OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1,0,0);
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
//	OnOnlyRead(m_iBtSt);
	if(m_iBtSt==5){
		OnShowEdit();
		OnShowAddress();
		OnShowTelFax();
		OnShowEmail();
		OnShowDoc();
		OnShowRasSch();
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

	strCap.LoadString(IDS_STRING12002);
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
				wdth = 40;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 1:
				wdth = 25;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 170;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 60;
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

	strCap.LoadString(IDS_STRING12003);
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
				wdth = 410;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 50;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 75;
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

	strCap.LoadString(IDS_STRING12004);
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
				wdth = 220;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
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
void CDlg::InitDataGrid4(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING12005);
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
				wdth = 285;
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
void CDlg::InitDataGrid5(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING12014);
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
				wdth = 280;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 150;
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
void CDlg::InitDataGrid6(CDatagrid1& Grd,_RecordsetPtr& rs)
{
	CColumns GrdClms;
	CString strCap,strRec;
	CString s;
	long num,numRec;
	short i;
	float wdth;

	strCap.LoadString(IDS_STRING12015);
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
				wdth = 130;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 270;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 127;
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
	if(m_Flg){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1){
					m_CurCol=0;
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
				if(m_iBtSt==5 ){
						OnShowEdit();
						OnShowAddress();
						OnShowTelFax();
						OnShowEmail();
						OnShowDoc();
						OnShowRasSch();

				}
			}
		}
	}
	m_Flg = true;
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

void CDlg::ChangeDatacombo1()
{
	m_cd=OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	OnShowEnter(m_cd);
}

void CDlg::OnBnClickedButton1()
{
	BeginWaitCursor();
//AfxMessageBox("inv");
//		m_SlpDay.SetDate(t1);
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
		hMod=AfxLoadLibrary(_T("VidFrmPrv.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidFrmPrv");
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

void CDlg::OnBnClickedButton2()
{
// имена
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strIns;
		bFndC   = TRUE;
		strFndC = _T("");
		short i;
		COleVariant vC;
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func;
		if(IsEnableRec(ptrRsCmb2)){
			COleVariant vC;
			i = 0;
			vC = GetValueRec(ptrRsCmb2,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();
		}
		if(m_Check1){
			hMod=AfxLoadLibrary(_T("NameMan.dll"));
			func=(pDialog)GetProcAddress(hMod,"startNameMan");
		}
		else{
			hMod=AfxLoadLibrary(_T("NameWom.dll"));
			func=(pDialog)GetProcAddress(hMod,"startNameWom");
		}
		(func)(m_strNT, ptrCnn,strFndC,bFndC);
		try{
			ptrRsCmb2->Requery(adCmdText);
			m_DataCombo2.Refresh();
			OnFindInCombo(strFndC,&m_DataCombo2,ptrRsCmb2,0,1);
/*			i = 1;
			vC = GetValueRec(ptrRsCmb2,i);
			vC.ChangeType(VT_BSTR);
			strIns = vC.bstrVal;
			m_DataCombo2.put_BoundText(strIns);
*/
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}


//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::OnBnClickedButton3()
{
// отчества
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strIns;
		bFndC   = TRUE;
		strFndC = _T("");
		COleVariant vC;
		short i;
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func;
		if(IsEnableRec(ptrRsCmb3)){
			COleVariant vC;
			i = 0;
			vC = GetValueRec(ptrRsCmb3,i);
			vC.ChangeType(VT_BSTR);
			strFndC = vC.bstrVal;
			strFndC.TrimLeft();
			strFndC.TrimRight();
		}
//			if(strFndC.IsEmpty()) strFndC = _T("");
		if(m_Check1){
			hMod=AfxLoadLibrary(_T("SbNmMan.dll"));
			func=(pDialog)GetProcAddress(hMod,"startSbNmMan");
		}
		else{
			hMod=AfxLoadLibrary(_T("SbNmWom.dll"));
			func=(pDialog)GetProcAddress(hMod,"startSbNmWom");
		}
		(func)(m_strNT, ptrCnn,strFndC,bFndC);
//AfxMessageBox(strFndC);
		try{
			ptrRsCmb3->Requery(adCmdText);
			m_DataCombo3.Refresh();
			OnFindInCombo(strFndC,&m_DataCombo3,ptrRsCmb3,0,1);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}


//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::OnBnClickedButton4()
{
//Добавить Адрес
	BeginWaitCursor();
		CString strCod,strIns;
		COleVariant vC;
		strCod = _T("0");
		strIns = _T("IT38 ");
		short i=0;
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strCod=vC.bstrVal;
		strCod.TrimLeft(' ');
		strCod.TrimRight(' ');
//		m_SlpDay.SetDate(t1);

		HMODULE hMod;
		hMod=AfxLoadLibrary(_T("AddBk.dll"));
		typedef BOOL (*pDialog)( CString, _ConnectionPtr,CString,CString );
		pDialog func=(pDialog)GetProcAddress(hMod,"startAddBk");
		if((func)(m_strNT, ptrCnn,strCod,strIns)){
			if(m_iBtSt==5){
				if(IsEnableRec(ptrRs2)){
					ptrRs2->Requery(adCmdText);
					InitDataGrid2(m_DataGrid2,ptrRs2);
				}
				else{
						OnShowAddress();
				}
			}
		}

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CDlg::OnBnClickedButton5()
{
//Удалить адрес
	if(m_iBtSt==5){
		if(IDYES==AfxMessageBox(IDS_STRING12020,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
			COleVariant vC,vBk;
			CString strC,strSql;
			_variant_t vra;
			VARIANT* vtl = NULL;
			short i = 0;
			if(IsEnableRec(ptrRs2)){
				if(!ptrRs2->adoEOF){
					vBk = ptrRs2->Bookmark;

					vC=GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					strC=vC.bstrVal;

					strSql = _T("DT38 ");
					strSql+=strC;
			//AfxMessageBox(strSql);
					ptrCmd2->CommandText = (_bstr_t)strSql;
					ptrCmd2->CommandType = adCmdText;
					try{
			//			AfxMessageBox(strSql);
						ptrCmd2->Execute(&vra,vtl,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					if(IsEnableRec(ptrRs1)){
						i = 0;
						vC = GetValueRec(ptrRs1,i);
						vC.ChangeType(VT_BSTR);
						strC = vC.bstrVal;
						strC.TrimRight();
					}
					strSql = _T("QT38w1_1 ");
					strSql+= strC;
					OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
					if(IsEnableRec(ptrRs2)){
						try{
							m_Flg = false;
							ptrRs2->PutBookmark(vBk);
							m_Flg = true;
						}
						catch(...){
							ptrRs2->MoveLast();
						}
					}
				}
			}
		}
	}
	else{
		CString msg;
		msg.LoadString(IDS_STRING12023);
		AfxMessageBox(msg,MB_ICONINFORMATION);
	}
}

void CDlg::OnBnClickedButton6()
{
//Добавить тел/Факс
	UpdateData();
	COleVariant vC;
	CString strC,strSql,strC4,strMsg,s;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i=0;
	m_Edit5.TrimLeft(' ');
	m_Edit5.TrimRight(' ');
	if(m_Edit5.IsEmpty()){
		strMsg.LoadString(IDS_STRING12018);
		s.LoadString(IDS_STRING12021);
		strMsg+=s;
		AfxMessageBox(strMsg);
	}
	else{
		if(IsEnableRec(ptrRs1)){
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight(' ');

			strC4 = m_Check2 ? _T("1"):_T("0");

			strSql = _T("IT39 ");
			strSql+= strC + _T(",'");
			strSql+= m_Edit5 + _T("',");
			strSql+= strC4 + _T(",'");
			strSql+= m_strNT +_T("'");
	//AfxMessageBox(strSql);
			ptrCmd3->CommandText = (_bstr_t)strSql;
			try{
	//			AfxMessageBox(strSql);
				ptrCmd3->Execute(&vra,vtl,adCmdText);

			}
			catch(_com_error& e){
				AfxMessageBox(e.Description());
			}
			OnShowTelFax();
		}
	}
//	InitDataGrid3(m_DataGrid3,ptrRs3);
}

void CDlg::OnBnClickedButton7()
{//Тел/Факс
	if(m_iBtSt == 5){
		if(IDYES==AfxMessageBox(IDS_STRING12020,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
			COleVariant vC;
			CString strC,strSql;
			_variant_t vra;
			VARIANT* vtl = NULL;
			short i = 0;
			if(IsEnableRec(ptrRs3)){
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;

				strC.TrimLeft(' ');
				strC.TrimRight(' ');

				strSql = _T("DT39 ");
				strSql+= strC;
	//AfxMessageBox(strSql);
				ptrCmd3->CommandText = (_bstr_t)strSql;
				try{
		//			AfxMessageBox(strSql);
					ptrCmd3->Execute(&vra,vtl,adCmdText);

				}
				catch(_com_error& e){
					AfxMessageBox(e.Description());
				}
				OnShowTelFax();
			}
		}
	}
	else{
		CString msg;
		msg.LoadString(IDS_STRING12024);
		AfxMessageBox(msg,MB_ICONINFORMATION);
	}
}

void CDlg::OnBnClickedButton8()
{
//Добавить e-mail
	UpdateData();
	COleVariant vC;
	CString strC,strSql,strMsg,s;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i=0;
	m_Edit6.TrimLeft(' ');
	m_Edit6.TrimRight(' ');
	if(m_Edit6.IsEmpty()){
		strMsg.LoadString(IDS_STRING12018);
		s.LoadString(IDS_STRING12022);
		strMsg+= s;
		AfxMessageBox(strMsg);
	}
	else{
		if(IsEnableRec(ptrRs1)){
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight(' ');

			strSql = _T("IT40 ");
			strSql+= strC + _T(",'");
			strSql+= m_Edit6 + _T("','");
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
			OnShowEmail();
		}
	}
	InitDataGrid4(m_DataGrid4,ptrRs4);
}

void CDlg::OnBnClickedButton9()
{
// Удалить e-mail
	if(m_iBtSt == 5){
		if(IDYES==AfxMessageBox(IDS_STRING12020,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
			COleVariant vC;
			CString strC,strSql;
			_variant_t vra;
			VARIANT* vtl = NULL;
			short i = 0;
			if(ptrRs4->State==adStateOpen){
				if(!IsEmptyRec(ptrRs4)){
					vC = GetValueRec(ptrRs4,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;

					strC.TrimLeft(' ');
					strC.TrimRight(' ');

					strSql = _T("DT40 ");
					strSql+= strC;
		//AfxMessageBox(strSql);
					ptrCmd4->CommandText = (_bstr_t)strSql;
					try{
			//			AfxMessageBox(strSql);
						ptrCmd4->Execute(&vra,vtl,adCmdText);

					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowEmail();
				}
			}
		}
	}
	else{
		CString msg;
		msg.LoadString(IDS_STRING12025);
		AfxMessageBox(msg,MB_ICONINFORMATION);
	}
}

void CDlg::OnBnClickedButton10()
{
//Добавить Документ
	CString strCod;
	COleVariant vC;
	strCod = _T("0");
	short i=0;

	if(IsEnableRec(ptrRs1)){

		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strCod=vC.bstrVal;
		strCod.TrimLeft();
		strCod.TrimRight();
		BeginWaitCursor();
	//		m_SlpDay.SetDate(t1);

			HMODULE hMod;
			BOOL bFndC;
			CString strFndC;
			bFndC   = TRUE;
			strFndC = _T("");
			hMod=AfxLoadLibrary(_T("AddDoc.dll"));
			typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,CString&,BOOL);
			pDialog func=(pDialog)GetProcAddress(hMod,"startAddDoc");
			if( (func)(m_strNT, ptrCnn, strFndC, strCod, bFndC) ){
				if(m_iBtSt==5){
					if(ptrRs5->State==adStateOpen){
						if(!IsEmptyRec(ptrRs5)){
//AfxMessageBox("1");
							ptrRs5->Requery(adCmdText);
							InitDataGrid5(m_DataGrid5,ptrRs5);
						}
						else{
							OnShowDoc();
						}
					}
				}
			}

	//		m_SlpDay.SetDate(t1);

			AfxFreeLibrary(hMod);
	}
	EndWaitCursor();
}

void CDlg::OnBnClickedButton11()
{//Удалить документ
	if(m_iBtSt == 5){
		if(IDYES==AfxMessageBox(IDS_STRING12020,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
			COleVariant vC;
			CString strC,strSql;
			_variant_t vra;
			VARIANT* vtl = NULL;
			short i = 0;
			if(ptrRs5->State==adStateOpen){
				if(!IsEmptyRec(ptrRs5)){
					vC = GetValueRec(ptrRs5,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;

					strC.TrimLeft(' ');
					strC.TrimRight(' ');

					strSql = _T("DT46 ");
					strSql+= strC;
		//AfxMessageBox(strSql);
					ptrCmd5->CommandText = (_bstr_t)strSql;
					try{
			//			AfxMessageBox(strSql);
						ptrCmd5->Execute(&vra,vtl,adCmdText);

					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowDoc();
				}
			}
		}
	}
	else{
		CString msg;
		msg.LoadString(IDS_STRING12026);
		AfxMessageBox(msg,MB_ICONINFORMATION);
	}
}

void CDlg::OnBnClickedButton12()
{//Добавить расч. счёт
	CString strCod;
	COleVariant vC;
	short i=0;
	if(IsEnableRec(ptrRs1)){
			strCod = _T("0");
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCod=vC.bstrVal;
			strCod.TrimLeft();
			strCod.TrimRight();
	//		m_SlpDay.SetDate(t1);

			HMODULE hMod;
			BOOL bFndC;
			CString strFndC;
			bFndC   = FALSE;
			strFndC = _T("");
			hMod=AfxLoadLibrary(_T("AddRasSch.dll"));
			typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,CString&,BOOL);
			pDialog func=(pDialog)GetProcAddress(hMod,"startAddRasSch");
			if( (func)(m_strNT, ptrCnn, strFndC, strCod, bFndC) ){
				if(m_iBtSt==5){
					if(ptrRs6->State==adStateOpen){
						if(!IsEmptyRec(ptrRs6)){
//AfxMessageBox("1");
							ptrRs6->Requery(adCmdText);
							InitDataGrid6(m_DataGrid6,ptrRs6);
						}
						else{
							OnShowRasSch();
						}
					}
				}
			}
	//	AfxMessageBox("str");
	//		Invalidate();

	//		m_SlpDay.SetDate(t1);

			AfxFreeLibrary(hMod);

		EndWaitCursor();
	}
}

void CDlg::OnBnClickedButton13()
{//Удалить расчётный счёт
	if(m_iBtSt == 5){
		if(IDYES==AfxMessageBox(IDS_STRING12020,MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2)){
			COleVariant vC;
			CString strC,strSql;
			_variant_t vra;
			VARIANT* vtl = NULL;
			short i = 0;
			if(ptrRs6->State==adStateOpen){
				if(!IsEmptyRec(ptrRs6)){
					vC = GetValueRec(ptrRs6,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;

					strC.TrimLeft(' ');
					strC.TrimRight(' ');

					strSql = _T("DT47 ");
					strSql+= strC;
		//AfxMessageBox(strSql);
					ptrCmd6->CommandText = (_bstr_t)strSql;
					try{
			//			AfxMessageBox(strSql);
						ptrCmd6->Execute(&vra,vtl,adCmdText);

					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
					OnShowRasSch();
				}
			}
		}
	}
	else{
		CString msg;
		msg.LoadString(IDS_STRING12027);
		AfxMessageBox(msg,MB_ICONINFORMATION);
	}

}
void CDlg::OnShowEnter(long cd)
{
	CString s;
	switch(cd){
	case 1:		//ПБЮЛ
	case 11:	//ЧП
	case 12:	//ИЧП
	case 15:	//ИП
		s.LoadString(IDS_STRING12017);
		GetDlgItem(IDC_STATIC_COMBOBOX2)->ShowWindow(true);
		GetDlgItem(IDC_DATACOMBO2)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_COMBOBOX3)->ShowWindow(true);
		GetDlgItem(IDC_DATACOMBO3)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(s);
		GetDlgItem(IDC_CHECK1)->ShowWindow(true);
		GetDlgItem(IDC_BUTTON2)->ShowWindow(true);
		GetDlgItem(IDC_BUTTON3)->ShowWindow(true);

		GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(false);
		GetDlgItem(IDC_EDIT2)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(false);
		GetDlgItem(IDC_EDIT3)->ShowWindow(false);

		break;
	default:
		s.LoadString(IDS_STRING12008);
		GetDlgItem(IDC_STATIC_COMBOBOX2)->ShowWindow(false);
		GetDlgItem(IDC_DATACOMBO2)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_COMBOBOX3)->ShowWindow(false);
		GetDlgItem(IDC_DATACOMBO3)->ShowWindow(false);
		GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(s);
		GetDlgItem(IDC_CHECK1)->ShowWindow(false);
		GetDlgItem(IDC_BUTTON2)->ShowWindow(false);
		GetDlgItem(IDC_BUTTON3)->ShowWindow(false);

		GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(true);
		GetDlgItem(IDC_EDIT2)->ShowWindow(true);
		GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(true);
		GetDlgItem(IDC_EDIT3)->ShowWindow(true);
		break;
	}
}

CString CDlg::OnQuerySex(bool  ch, CString& str)
{
	CString s;
	if(ch){
		s   = _T("QT41");
		str = _T("QT42");
	}
	else{
		s   = _T("QT43");
		str = _T("QT44");
	}
	return s;
}

void CDlg::OnShowNameCombo(void)
{
	CString strSql,strSql2;

	strSql = OnQuerySex(m_Check1,strSql2);
	m_DataCombo2.putref_RowSource(NULL);
	/*if(ptrRsCmb2->State==adStateOpen)*/ ptrRsCmb2->Close();
	ptrCmdCmb2->CommandText = (_bstr_t)strSql;
	try{
		ptrRsCmb2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText); //adCmdUnknown
		m_DataCombo2.putref_RowSource((LPUNKNOWN) ptrRsCmb2);
		m_DataCombo2.put_ListField(_T("Имя"));
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
	}	

	m_DataCombo3.putref_RowSource(NULL);
	/*if(ptrRsCmb3->State==adStateOpen)*/ ptrRsCmb3->Close();
	ptrCmdCmb3->CommandText = (_bstr_t)strSql2;
	try{
		ptrRsCmb3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText); //adCmdUnknown
		m_DataCombo3.putref_RowSource((LPUNKNOWN) ptrRsCmb3);
		m_DataCombo3.put_ListField(_T("Отчество"));
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
	}
//AfxMessageBox(strSql+" "+strSql2);
	return;
}

void CDlg::OnShowAddress(void)
{
//Адреса КонтрАгента
	COleVariant vC;
	short i;
	CString strSql,strC;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i = 0;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();

			strSql = _T("QT38w1_1 ");
			strSql+= strC;
			OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
		}
	}
}

void CDlg::OnShowTelFax(void)
{
	COleVariant vC;
	CString strC,strSql;
	short i;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i=0;
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight();

			strSql = _T("QT39 ");
			strSql+= strC;
		//AfxMessageBox(strSql);
			OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
		}
	}
}

void CDlg::OnShowEmail(void)
{
	COleVariant vC;
	CString strC,strSql;
	short i;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i=0;
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight();

			strSql = _T("QT40 ");
			strSql+= strC;
		//AfxMessageBox(strSql);
			OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4,0,0);
		}
	}
}

void CDlg::OnShowDoc(void)
{
	COleVariant vC;
	CString strC,strSql;
	short i;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i=0;
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight();

			strSql = _T("QT46 ");
			strSql+= strC;
		//AfxMessageBox(strSql);
			OnShowGrid(strSql, ptrRs5,ptrCmd5,m_DataGrid5,InitDataGrid5);
		}
	}
}

void CDlg::OnShowRasSch(void)
{
	COleVariant vC;
	CString strC,strSql;
	short i;
	if(IsEnableRec(ptrRs1)){
		if(!ptrRs1->adoEOF){
			i=0;
			vC=GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strC=vC.bstrVal;
			strC.TrimRight();

			strSql = _T("QT47 ");
			strSql+= strC;
		//AfxMessageBox(strSql);
			OnShowGrid(strSql, ptrRs6,ptrCmd6,m_DataGrid6,InitDataGrid6);
		}
	}
}

void CDlg::OnBnClickedCheck1()
{
	CString strSql,strSql2;
	m_Check1 = m_Check1 ? false:true;
	((CButton*)GetDlgItem(IDC_CHECK1))->SetCheck(m_Check1);
	OnShowNameCombo();
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
