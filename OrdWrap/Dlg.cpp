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

	, cdCombo1(0)
	, cdCombo2(0)
	, m_Edit1(_T("0"))
	, m_Edit2(_T("0"))
	, m_Edit3(_T("0"))
{

}

CDlg::~CDlg()
{
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;

	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2 = NULL;

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
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
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
	ON_EN_CHANGE(IDC_EDIT1, &CDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT2, &CDlg::OnEnChangeEdit2)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"OrdWrap.dll"));

	CString s,strSql,strSql2,strC;
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
	strSql = L"QT12_1";
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);

	if(IsEnableRec(ptrRs1)){
		i = 0;
		vC = GetValueRec(ptrRs1,i);
		vC.ChangeType(VT_BSTR);
		strC = vC.bstrVal;
		strC.TrimLeft();
		strC.TrimRight();

		strSql = _T("T13_T12T14T29T16 ");
		strSql+= strC;
		OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

		if(m_bFndC){
			short fCol=0;
			m_Flg = false;
			OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
			m_Flg = true;
		}
	}

	ptrCmdCmb1 = NULL;
	ptrRsCmb1  = NULL;

	strSql = _T("QT29");
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

	strSql = _T("QT30");
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
		m_DataCombo2.put_ListField(_T("Вид упаковки"));	}
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

		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
//AfxMessageBox("g11");
			m_CurCol = m_DataGrid1.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
		}

		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
//AfxMessageBox("g22");
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
			m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
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
				m_Edit2.TrimRight(' ');
				m_Edit2.TrimLeft(' ');
				if(m_Edit2.IsEmpty()){
					count++;
					GetDlgItem(IDC_STATIC_EDIT2)->GetWindowText(strItem);
					strCount.Format(L"%i",count);
					strCount+=_T(") ");
					strMsg+=strCount+strItem+_T("\n\t");
				}
				break;
			case IDC_EDIT3:
				m_Edit3.TrimRight(' ');
				m_Edit3.TrimLeft(' ');
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

	if(count!=0)
		AfxMessageBox(strMsg,MB_ICONINFORMATION);
	return count==0 ? true:false;
}

void CDlg::OnShowEdit(void)
{
	COleVariant vC;
	short i;
	CString strF;
	if(IsEnableRec(ptrRs1)){
		if(IsEnableRec(ptrRs2)){

				i=9; // Код вида коробки
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strF=vC.bstrVal;
				OnFindInCombo(strF,&m_DataCombo1,ptrRsCmb1,0,1);

				i=10; // Код вида упаковки
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strF=vC.bstrVal;
				OnFindInCombo(strF,&m_DataCombo2,ptrRsCmb2,0,1);

				i=3;
				vC=GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				m_Edit1=vC.bstrVal;
				m_Edit1.TrimLeft();
				m_Edit1.TrimRight();
				GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);

				i=4;
				vC=GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				m_Edit2=vC.bstrVal;
				m_Edit2.TrimLeft();
				m_Edit2.TrimRight();
				GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

				i=11;
				vC=GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				m_Edit3=vC.bstrVal;
				m_Edit3.TrimLeft();
				m_Edit3.TrimRight();
				GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		}
	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_COMBO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_COMBO2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
//	UpdateData();
	CString strSql;
	CString strCT,strCB,strCU,strC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	COleVariant vC;
	short IndCol;
	BOOL bStart;
	int IdBeg,IdEnd;
	short i;
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}
	if(cdCombo1!=1 && cdCombo2!=1){
		IdBeg = IDC_EDIT1;
		IdEnd = IDC_EDIT2; //IDC_EDIT4
		bStart = true;
	}
	else if(cdCombo1==1 && cdCombo2!=1){
		IdBeg = IDC_EDIT2;
		IdEnd = IDC_EDIT2; //IDC_EDIT4
		bStart = true;
	}
	else if(cdCombo1!=1 && cdCombo2==1){
		bStart = false;
	}
	else if(cdCombo1==1 && cdCombo2==1){
		bStart = true;
		IdBeg = IDC_OK; //IDC_EDIT6;//IDC_EDIT4
		IdEnd = IDC_OK; //IDC_EDIT6;//IDC_EDIT4
	}

	if(bStart){
		if(OnOverEdit(IdBeg ,IdEnd)){
			if(IsEnableRec(ptrRs1)){
				i = 0;
				vC = GetValueRec(ptrRs1,i);		//Код Товара
				vC.ChangeType(VT_BSTR);
				strCT = vC.bstrVal;
				strCT.TrimRight();
				strCT.TrimLeft();

				if(IsEnableRec(ptrRsCmb1)){
					i = 0;
					vC = GetValueRec(ptrRsCmb1,i);	//Код вида коробки
					vC.ChangeType(VT_BSTR);
					strCB = vC.bstrVal;
					strCB.TrimLeft();
					strCB.TrimRight();
				}
				else{
					return;	//Выход т.к. пусто
				}

				if(IsEnableRec(ptrRsCmb2)){
					i = 0;
					vC = GetValueRec(ptrRsCmb2,i);	//Код вида упаковки
					vC.ChangeType(VT_BSTR);
					strCU = vC.bstrVal;
					strCU.TrimLeft();
					strCU.TrimRight();
				}
				else{
					return; //Выход т.к. пусто
				}
			}
			else{
				return; //Выход т.к. пусто
			}

			if(cdCombo1==1)					m_Edit1 = _T("0");
			if(cdCombo2==1)					m_Edit2 = _T("0");
			if(cdCombo1==1 && cdCombo2==1)	m_Edit3 = _T("0");

			strSql =_T("IT13 ");
			strSql+=strCT + _T(",");
			strSql+=strCB + _T(",");
			strSql+=strCU + _T(",");
			strSql+=m_Edit1+_T(",");
			strSql+=m_Edit2+_T(",");
			strSql+=m_Edit3 + _T(",'");
			strSql+=m_strNT+_T("'");
//AfxMessageBox(strSql);
			ptrCmd2->CommandText = (_bstr_t)strSql;
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
				strC.TrimLeft();
				strC.TrimRight();
			}

			strSql = _T("T13_T12T14T29T16 ");
			strSql+= strC;

			OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
			if(IsEnableRec(ptrRs2)){
				if(!m_Edit1.IsEmpty()){
					IndCol = 8;
					m_Flg = false;
					OnFindInGrid(strCT,ptrRs2,IndCol,m_Flg);
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
//	UpdateData();
	CString strSql,strC;
	CString strCB,strCU;
	_variant_t vra;
	COleVariant vC;
	VARIANT* vtl = NULL;
	short IndCol,i;
	long IdBeg,IdEnd;
	BOOL bStart;

	if(cdCombo1!=1 && cdCombo2!=1){
		IdBeg = IDC_EDIT1;
		IdEnd = IDC_EDIT2; 
		bStart = true;
	}
	else if(cdCombo1==1 && cdCombo2!=1){
		IdBeg = IDC_EDIT2;
		IdEnd = IDC_EDIT2; 
		bStart = true;
	}
	else if(cdCombo1!=1 && cdCombo2==1){
		bStart = false;
	}
	else if(cdCombo1==1 && cdCombo2==1){
		bStart = true;
		IdBeg = IDC_OK; 
		IdEnd = IDC_OK; 
	}
	if(bStart){
		if(IsEnableRec(ptrRs2)){

			i = 0;	//Код записи Т-У
			vC = GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimRight();

				if(OnOverEdit(IdBeg,IdEnd)){
					if(cdCombo1==1)					m_Edit1 = _T("0");
					if(cdCombo2==1)					m_Edit2 = _T("0");
					if(cdCombo1==1 && cdCombo2==1)	m_Edit3 = _T("0");

					i = 0;
					vC = GetValueRec(ptrRsCmb1,i);	//Код вида коробки
					vC.ChangeType(VT_BSTR);
					strCB = vC.bstrVal;
					strCB.TrimLeft();
					strCB.TrimRight();

					i = 0;
					vC = GetValueRec(ptrRsCmb2,i);	//Код вида упаковки
					vC.ChangeType(VT_BSTR);
					strCU = vC.bstrVal;
					strCU.TrimLeft();
					strCU.TrimRight();

					strSql =_T("UT13 ");
					strSql+=strC  + _T(",");
					strSql+=strCB + _T(",");
					strSql+=strCU + _T(",");
					strSql+=m_Edit1+_T(",");
					strSql+=m_Edit2+_T(",");
					strSql+=m_Edit3  + _T(",'");
					strSql+=m_strNT+_T("'");

					ptrCmd2->CommandText = (_bstr_t)strSql;
					ptrCmd2->CommandType = adCmdText;
					try{
	//AfxMessageBox(strSql);
						ptrCmd2->Execute(&vra,vtl,adCmdText);
					}
					catch(_com_error& e){
						AfxMessageBox(e.Description());
					}
	//AfxMessageBox(ptrCmd1->GetCommandText());
				if(IsEnableRec(ptrRs1)){
					i = 0;
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strC = vC.bstrVal;
					strC.TrimLeft();
					strC.TrimRight();
				}

				strSql = _T("T13_T12T14T29T16 ");
				strSql+= strC;
				OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
				if(IsEnableRec(ptrRs2)){
					IndCol = 0;
					m_bFnd = OnFindInGrid(strC,ptrRs2,IndCol,m_Flg);
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
	if(IsEnableRec(ptrRs2)){

			ptrRs2->get_Bookmark(vBk);

			i = 0;	//Код
			vC = GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			strC = vC.bstrVal;
			strC.TrimRight();
			strC.TrimLeft();

			strSql = _T("DT13 ");
			strSql+= strC;

			ptrCmd2->CommandType = adCmdText;
			ptrCmd2->CommandText = (_bstr_t)strSql;
			try{
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
				strC.TrimLeft();
				strC.TrimRight();
			}

			strSql = _T("T13_T12T14T29T16 ");
			strSql+= strC;
			OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

			if(IsEnableRec(ptrRs2)){
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
				GrdClms.GetItem((COleVariant) i).SetWidth(30);
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

	strCap.LoadString(IDS_STRING9017);
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
				wdth = 80;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 80;
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
			case 6:
				wdth = 65;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 7:
				wdth = 60;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 11:
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

BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, CDlg::RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATACOMBO2, 1, CDlg::ChangeDatacombo2, VTS_NONE)
END_EVENTSINK_MAP()

void CDlg::RowColChangeDatagrid1(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
		if(IsEnableRec(ptrRs1)){
			if(!ptrRs1->adoEOF){
/*				m_CurCol = m_DataGrid1.get_Col();
				if(m_CurCol==-1 || m_CurCol==0){
					m_CurCol=1;
				}
				m_iCurType = GetTypeCol(ptrRs1,m_CurCol);
*/
			    //OnShowGrid2();
				short i = 0;
				COleVariant vC;
				CString strC,strSql;

				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimLeft();
				strC.TrimRight();

				strSql = _T("T13_T12T14T29T16 ");
				strSql+= strC;
				OnShowGrid(strSql,ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
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

void CDlg::ChangeDatacombo1()
{
	cdCombo1=OnChangeCombo(m_DataCombo1,ptrRsCmb1,1);
	switch(cdCombo1){
	case 1:
		GetDlgItem(IDC_EDIT1)->EnableWindow(false);
		m_Edit1=_T("0");
		GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
		break;
	default:
		if(cdCombo2==1){
			m_Edit1=_T("0");
			m_Edit2=_T("0");
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit2);
			GetDlgItem(IDC_EDIT1)->EnableWindow(false);
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		}
		else{
			GetDlgItem(IDC_EDIT1)->EnableWindow(true);
		}
		break;
	}
}

void CDlg::ChangeDatacombo2()
{
	cdCombo2 = OnChangeCombo(m_DataCombo2,ptrRsCmb2,1);
	switch(cdCombo2){
	case 1:
		if(cdCombo1!=1){
			m_Edit1=_T("0");
			m_Edit3=_T("0");
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
			GetDlgItem(IDC_EDIT1)->EnableWindow(false);
			GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		}
		m_Edit2=_T("0");
		GetDlgItem(IDC_EDIT2)->EnableWindow(false);
		GetDlgItem(IDC_EDIT2)->SetWindowText(m_Edit2);

		break;
	default:
		GetDlgItem(IDC_EDIT2)->EnableWindow(true);
		if(cdCombo1!=1){
			m_Edit1=_T("0");
//		AfxMessageBox("def");
			GetDlgItem(IDC_EDIT1)->SetWindowText(m_Edit1);
			GetDlgItem(IDC_EDIT1)->EnableWindow(true);
		}
		break;
	}
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
		hMod=AfxLoadLibrary(_T("VidUp.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidUp");
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
		hMod=AfxLoadLibrary(_T("VidBox.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidBox");
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

void CDlg::OnEnChangeEdit1()
{
	int iEdit1,iEdit2,iEdit3;
	iEdit1 = iEdit2 = iEdit3 = 0;
	 
	UpdateData();
	switch(cdCombo1){
	case 1:
		m_Edit3 = m_Edit2;
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		break;
	default:

		iEdit1=_wtoi(m_Edit1);
		iEdit2=_wtoi(m_Edit2);
		iEdit3 = iEdit1*iEdit2;
		m_Edit3.Format(L"%i",iEdit3);
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		break;
	}
}

void CDlg::OnEnChangeEdit2()
{
	int iEdit1,iEdit2,iEdit3;
	iEdit1 = iEdit2 = iEdit3 = 0;
	
	UpdateData();
	switch(cdCombo1){
	case 1:
		m_Edit3 = m_Edit2;
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		break;
	default:
		iEdit1=_wtoi(m_Edit1);
		iEdit2=_wtoi(m_Edit2);
		iEdit3 = iEdit1*iEdit2;
		m_Edit3.Format(L"%i",iEdit3);
		GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit3);
		break;
	}
}
