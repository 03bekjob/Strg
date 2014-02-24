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

	, m_Radio1(1)
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

	if(ptrRs4->State==adStateOpen) ptrRs4->Close();
	ptrRs4 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATAGRID1, m_DataGrid1);
	DDX_Control(pDX, IDC_DATAGRID2, m_DataGrid2);
	DDX_Control(pDX, IDC_DATAGRID3, m_DataGrid3);
	DDX_Control(pDX, IDC_DATAGRID4, m_DataGrid4);
	DDX_Radio(pDX, IDC_RADIO1, m_Radio1);
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
	ON_BN_CLICKED(IDC_RADIO1, &CDlg::OnBnClickedRadio1)
	ON_BN_CLICKED(IDC_RADIO2, &CDlg::OnBnClickedRadio2)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"OrdStrg.dll"));

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
	strSql = L"QT99";
	OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
	if(m_bFndC){
		short fCol=7;
		m_Flg = false;
		OnFindInGrid(m_strFndC,ptrRs1,fCol,m_Flg);
		m_Flg = true;
	}
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);

	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);
	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			if(IsEnableRec(ptrRs3)){
				i = 0;
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimLeft();
				strC.TrimRight();

				strSql = _T("T27_T26 ");
				strSql+= s	+ _T(",");
				strSql+= strC;
			}
			break;
		case 1:
				strSql = _T("T27_T26 ");
				strSql+= s	+_T(",0");
			break;
	}
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);

	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);

	s = m_Radio1==0 ? _T("0"):_T("1");
	switch(m_Radio1){
		case 0:
			strSql = _T("T12_T27_1 ");
			strSql+= s + _T(",0");
			break;
		case 1:
			if(IsEnableRec(ptrRs2)){
				i = 0;
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strC = vC.bstrVal;
				strC.TrimLeft();
				strC.TrimRight();
				strSql = _T("T12_T27_1 ");
				strSql+= s + _T(",");
				strSql+= strC;
			}
			break;
	}
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);

	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);
	strSql = _T("QT4");
	OnShowGrid(strSql, ptrRs4,ptrCmd4,m_DataGrid4,InitDataGrid4);

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

		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			m_CurCol = m_DataGrid3.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs3,m_CurCol);
		}

		if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
			m_CurCol = m_DataGrid4.get_Col();
			if(m_CurCol==-1 || m_CurCol==0){
				m_CurCol = 1;
			}
			m_iCurType = GetTypeCol(ptrRs4,m_CurCol);
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
//AfxMessageBox("1");
//		AfxMessageBox(str);
		if(GetDlgItem(IDC_DATAGRID1)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs1,m_CurCol,m_Flg);
//			m_Flg = true;
		}

		if(GetDlgItem(IDC_DATAGRID2)==m_pLstWnd){
			m_Flg = false;
//AfxMessageBox("EditTB 2");
			m_bFnd = OnFindInGrid(str,ptrRs2,m_CurCol,m_Flg);
//AfxMessageBox("EditTB _2");
//			m_Flg = true;
		}

		if(GetDlgItem(IDC_DATAGRID3)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs3,m_CurCol,m_Flg);
//			m_Flg = true;
		}

		if(GetDlgItem(IDC_DATAGRID4)==m_pLstWnd){
			m_Flg = false;
			m_bFnd = OnFindInGrid(str,ptrRs4,m_CurCol,m_Flg);
//			m_Flg = true;
		}
//AfxMessageBox("2");
/*		if(m_bFnd){
			AfxMessageBox(L"true");
		}
		else{
			AfxMessageBox(L"false");
		}
*/

	}
	return;
}

bool CDlg::OnOverEdit(int IdBeg, int IdEnd)
{
	CString strItem,strMsg,strCount;
	int count=0;
	int Id;
	strMsg.LoadString(IDS_STRING9016);

	if(m_iBtSt==5 ){	//Режим Корректировки
		//Товар-Склад
		if(ptrRs1->State==adStateOpen){
			if(IsEmptyRec(ptrRs1)){
				count++;
				strItem.LoadString(IDS_STRING9014);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
			count++;
			strItem.LoadString(IDS_STRING9014);
			strCount.Format(L"%i",count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{		//Режим Добавить/Удалить проверяем без учёта switch
				// т.к. и Товар-Группа должны уже существовать
			//Товар
		if(ptrRs3->State==adStateOpen){
			if(IsEmptyRec(ptrRs3)){
				count++;
				strItem.LoadString(IDS_STRING9024);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
				count++;
				strItem.LoadString(IDS_STRING9024);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
		}
			//Группа
		if(ptrRs2->State==adStateOpen){
			if(IsEmptyRec(ptrRs2)){
				count++;
				strItem.LoadString(IDS_STRING9023);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
			}
		}
		else{
				count++;
				strItem.LoadString(IDS_STRING9023);
				strCount.Format(L"%i",count);
				strCount+=_T(") ");
				strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	//Склад
	if(ptrRs4->State==adStateOpen){
		if(IsEmptyRec(ptrRs4)){
			count++;
			strItem.LoadString(IDS_STRING9025);
			strCount.Format(L"%i",count);
			strCount+=_T(") ");
			strMsg+=strCount+strItem+_T("\n\t");
		}
	}
	else{
			count++;
			strItem.LoadString(IDS_STRING9025);
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
	CString strF;
	short IndCol;
	if(IsEnableRec(ptrRs1)){
		switch(m_Radio1){
			case 0:	// по Товару
				i=7; // Код Товара
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strF=vC.bstrVal;
				IndCol = 0;
				m_Flg = false;
				if(ptrRs2->State==adStateOpen){
					m_bFnd =	OnFindInGrid(strF,ptrRs3,IndCol,m_Flg);
				}
				m_Flg = true;
				break;
			case 1:	// по Группе
				i=6; // Код Группы
				vC = GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strF=vC.bstrVal;
				IndCol = 0;
				m_Flg = false;
				if(ptrRs3->State==adStateOpen){
					m_bFnd =	OnFindInGrid(strF,ptrRs2,IndCol,m_Flg);
				}
				m_Flg = true;
//				if(m_bFnd){
					if(ptrRs2->State==adStateOpen){
						i=7; // Код Товара
						vC = GetValueRec(ptrRs1,i);
						vC.ChangeType(VT_BSTR);
						strF=vC.bstrVal;
						IndCol = 0;
						m_Flg = false;
						if(ptrRs2->State==adStateOpen){
							m_bFnd =	OnFindInGrid(strF,ptrRs3,IndCol,m_Flg);
						}
						m_Flg = true;
					}
//				}
				break;
		}

		if(ptrRs4->State==adStateOpen){
			i = 8;
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strF=vC.bstrVal;
			IndCol = 0;
			m_Flg = false;
			if(ptrRs4->State==adStateOpen){
				m_bFnd =	OnFindInGrid(strF,ptrRs4,IndCol,m_Flg);
			}
			m_Flg = true;

		}

	}
}


void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_GROUP)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_RADIO1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9022);
	GetDlgItem(IDC_RADIO2)->SetWindowText(sTxt);
}

void CDlg::On32775(void)
{
//Добавить
//	UpdateData();
	CString strSql;
	CString strCT,strCG,strCS,strR;
	COleVariant vC;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;
	short i;

	strCT = _T("0");
	strCG = _T("0");
	strCS = _T("0");
	strR  = m_Radio1==0 ? _T("0"):_T("1");
	if(m_iBtSt==5){
		AfxMessageBox(IDS_STRING9011,MB_ICONINFORMATION);
		return;
	}

	if(OnOverEdit(IDC_DATAGRID1,IDC_DATAGRID4)){
		switch(m_Radio1){
			case 0:	//по Товару
				i = 0;
				vC = GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				strCT = vC.bstrVal;
				strCT.TrimLeft();
				strCT.TrimRight();
				break;
			case 1:	//по Группе
				i = 0;
				vC = GetValueRec(ptrRs2,i);
				vC.ChangeType(VT_BSTR);
				strCG = vC.bstrVal;
				strCG.TrimLeft();
				strCG.TrimRight();
				break;
		}
		i = 0;		// Склад
		vC = GetValueRec(ptrRs4,i);
		vC.ChangeType(VT_BSTR);
		strCS = vC.bstrVal;
		strCS.TrimLeft();
		strCS.TrimRight();

		strSql = _T("IT99_1 ");
		strSql+= strR	+ _T(",");
		strSql+= strCT  + _T(",");	//Код Товара
		strSql+= strCG	+ _T(",");	//Код Группы
		strSql+= strCS  + _T(",'");	//Код Склада
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
		strSql = _T("QT99");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			switch(m_Radio1){
				case 0:
					IndCol = 7;	//по Товару
					m_Flg = false;
					OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
					m_Flg = true;
					break;
				case 1:
					IndCol = 6;	//по Группе
					m_Flg = false;
					OnFindInGrid(strCG,ptrRs1,IndCol,m_Flg);
					m_Flg = true;
					break;
			}
		}
	}

	return;
}

void CDlg::On32776(void)
{
// Изменить
//	UpdateData();
	CString strSql,strCS,strCSN,strCT,strCG; 
	CString strR;
	COleVariant vC,vBk;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short IndCol;

	strCT  = _T("0");
	strCG  = _T("0");
	strCS  = _T("0");
	strCSN = _T("0");

	short i,InnCol;
	i = 0;
	if(IsEnableRec(ptrRs1)){
		if(OnOverEdit(IDC_DATAGRID2,IDC_DATAGRID3)){
			switch(m_Radio1){
				case 0:		// по Товару
					i=7;	// Код Товара
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strCT=vC.bstrVal;
					break;
				case 1:		// по Группе
					i=6;	// Код Группы
					vC = GetValueRec(ptrRs1,i);
					vC.ChangeType(VT_BSTR);
					strCG=vC.bstrVal;
					break;
			}
			vBk = ptrRs1->GetBookmark();

			i = 8;	//Код Склада стар
			vC = GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCS = vC.bstrVal;
			strCS.TrimLeft();
			strCS.TrimRight();

			i = 0;  //Код Склада Новый
			vC = GetValueRec(ptrRs4,i);
			vC.ChangeType(VT_BSTR);
			strCSN = vC.bstrVal;
			strCSN.TrimLeft();
			strCSN.TrimRight();

			strR = m_Radio1==0 ? _T("0"):_T("1");
			strSql = _T("UT99 ");
			strSql+= strR   + _T(",");
			strSql+= strCT  + _T(",");	// Код товара
			strSql+= strCG  + _T(",");	// Код группы
			strSql+= strCS  + _T(",");	// Код стар склад
			strSql+= strCSN + _T(",'"); // Код нового склад
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
			strSql = _T("QT99");
		    OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
			if(IsEnableRec(ptrRs1)){
				switch(m_Radio1){
					case 0:
						IndCol = 7;
						m_Flg = false;
						m_bFnd = OnFindInGrid(strCT,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
						break;
					case 1:
						IndCol = 6;
						m_Flg = false;
						m_bFnd = OnFindInGrid(strCG,ptrRs1,IndCol,m_Flg);
						m_Flg = true;
						break;
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
	CString strSql,strR,strC1,strCG;
	_variant_t vra;
	VARIANT* vtl = NULL;
	short i;
	strCG = _T("0");
	strC1 = _T("0");
	if(IsEnableRec(ptrRs1)){

		ptrRs1->get_Bookmark(vBk);
		switch(m_Radio1){
			case 0:		//по Товару
				i=0;	//Код записи в Т99
				vC=GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strC1=vC.bstrVal;
				strC1.TrimLeft();
				strC1.TrimRight();
				break;
			case 1:		//по Группе
				i=6;	//Код Группы
				vC=GetValueRec(ptrRs1,i);
				vC.ChangeType(VT_BSTR);
				strCG=vC.bstrVal;
				strCG.TrimLeft();
				strCG.TrimRight();
				break;
		}

		strR = m_Radio1==0 ? _T("0"):_T("1");
		strSql = _T("DT99_1 ");
		strSql+= strR	+ _T(",");
		strSql+= strC1	+ _T(",");
		strSql+= strCG;

		ptrCmd1->CommandType = adCmdText;
		ptrCmd1->CommandText = (_bstr_t)strSql;
		try{
			ptrCmd1->Execute(&vra,vtl,adCmdText);
		}
		catch(_com_error& e){
			AfxMessageBox(e.Description());
		}
		strSql = _T("QT99");
		OnShowGrid(strSql, ptrRs1,ptrCmd1,m_DataGrid1,InitDataGrid1);
		if(IsEnableRec(ptrRs1)){
			try{
				m_Flg = false;
				ptrRs1->PutBookmark(vBk);
				m_Flg = true;
			}
			catch(...){
				if(ptrRs1->adoEOF) ptrRs1->MoveLast();
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
//			m_Flg = true;
//AfxMessageBox(L"Before OnShow Grd.putref_DataSource((LPUNKNOWN)rs)");
			Grd.putref_DataSource((LPUNKNOWN)rs);
//AfxMessageBox(L"After OnShow Grd.putref_DataSource((LPUNKNOWN)rs)");
//			m_Flg = false;
/*----*///			m_Flg = true;
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


BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATAGRID1, DISPID_CLICK, CDlg::ClickDatagrid1, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID2, DISPID_CLICK, CDlg::ClickDatagrid2, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID3, DISPID_CLICK, CDlg::ClickDatagrid3, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID4, DISPID_CLICK, CDlg::ClickDatagrid4, VTS_NONE)
	ON_EVENT(CDlg, IDC_DATAGRID1, 218, CDlg::RowColChangeDatagrid1, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID2, 218, CDlg::RowColChangeDatagrid2, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID3, 218, CDlg::RowColChangeDatagrid3, VTS_PVARIANT VTS_I2)
	ON_EVENT(CDlg, IDC_DATAGRID4, 210, CDlg::ButtonClickDatagrid4, VTS_I2)
END_EVENTSINK_MAP()

void CDlg::ClickDatagrid1()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID1);
}

void CDlg::ClickDatagrid2()
{
//	m_Flg = true;
	m_pLstWnd = GetDlgItem(IDC_DATAGRID2);
}

void CDlg::ClickDatagrid3()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID3);
}

void CDlg::ClickDatagrid4()
{
	m_pLstWnd = GetDlgItem(IDC_DATAGRID4);
}

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

void CDlg::RowColChangeDatagrid2(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
//AfxMessageBox(L"RW2 m_Flg=TRUE");
		if(IsEnableRec(ptrRs2)){
/*			CString s;
			s.Format(L"m_Radio1 = %i",m_Radio1);
*/
			switch(m_Radio1){
				case 1:
					if(!ptrRs2->adoEOF){
//AfxMessageBox(L"case RW2 1 и !ptrRs2->adoEOF "+s);
						CString strSql,strC,s;
						COleVariant vC;
						short i;
						s = m_Radio1==0 ? _T("0"):_T("1");
						i = 0;
						vC = GetValueRec(ptrRs2,i);
						vC.ChangeType(VT_BSTR);
						strC = vC.bstrVal;
						strC.TrimLeft();
						strC.TrimRight();
						strSql = _T("T12_T27_1 ");
						strSql+= s + _T(",");
						strSql+= strC;

//AfxMessageBox(L"RW2 B OnShow G3 товар");
//						m_Flg = false;
						OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
//AfxMessageBox(L"RW2 E OnShow G3 товар");
//						m_Flg = true;
					}
					break;
			}
		}
	}
	m_Flg = true;
}

void CDlg::RowColChangeDatagrid3(VARIANT* LastRow, short LastCol)
{
	if(m_Flg){
//AfxMessageBox(L"RW3 m_Flg=TRUE");
		if(IsEnableRec(ptrRs3)){
/*			CString s;
			s.Format(L"m_Radio1 = %i",m_Radio1);
*/
			switch(m_Radio1){
				case 0:
					if(!ptrRs3->adoEOF){
//AfxMessageBox(L"case RW3 и !ptrRs3->adoEOF "+s);
						CString strSql,strC,s;
						COleVariant vC;
						short i;

						s = m_Radio1==0 ? _T("0"):_T("1");
						i = 0;
						vC = GetValueRec(ptrRs3,i);
						vC.ChangeType(VT_BSTR);
						strC = vC.bstrVal;
						strC.TrimLeft();
						strC.TrimRight();

						strSql = _T("T27_T26 ");
						strSql+= s	+ _T(",");
						strSql+= strC;

//	AfxMessageBox(L"RW3 B OnShow G2 группа");
	//					m_Flg = false;
						OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
//	AfxMessageBox(L"RW3 E OnShow G2 группа");
	//					m_Flg = false;
					}
					break;
			}
		}
	}
	m_Flg = true;
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
				wdth = 190;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 48;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 4:
				wdth = 185;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 100;
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

	strCap.LoadString(IDS_STRING9018);
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
				wdth = 45;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 220;
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
				wdth = 75;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 2:
				wdth = 105;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 3:
				wdth = 245;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				break;
			case 5:
				wdth = 46;
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

	strCap.LoadString(IDS_STRING9019);
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
				wdth = 270;
				GrdClms.GetItem((COleVariant) i).SetWidth(wdth);
				GrdClms.GetItem((COleVariant) i).SetButton(TRUE);
				Grd.put_Col(i);
				break;
			case 2:
				wdth = 120;
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

void CDlg::OnBnClickedRadio1()
{//По товару
//	UpdateData();
	m_Radio1 = 0;
	CString strSql,strC,s;
	COleVariant vC;
	short i;
	s = m_Radio1==0 ? _T("0"):_T("1");
	strSql = _T("T12_T27_1 ");
	strSql+= s + _T(",0");
//AfxMessageBox(L"Radio1 B OnShw G3 товар");
//	m_Flg = false;
	OnShowGrid(strSql, ptrRs3,ptrCmd3,m_DataGrid3,InitDataGrid3);
//	m_Flg = true;
//AfxMessageBox(L"Radio1 E OnShw G3 товар");
	if(m_iBtSt==5 ){
		if(m_bFnd ){
			OnShowEdit();
		}
	}
}

void CDlg::OnBnClickedRadio2()
{// По группе
//	UpdateData();
	m_Radio1 = 1;
	CString strSql,strC,s;
	COleVariant vC;
	short i;
	s = m_Radio1==0 ? _T("0"):_T("1");
	strSql = _T("T27_T26 ");
	strSql+= s + _T(",0");
//AfxMessageBox(L"Radio1 B OnShw G2 группа");
//	m_Flg = false;
	OnShowGrid(strSql, ptrRs2,ptrCmd2,m_DataGrid2,InitDataGrid2);
//	m_Flg = true;
//AfxMessageBox(L"Radio1 E OnShw G2 группа");

	if(m_iBtSt==5 ){
		if(m_bFnd ){
			OnShowEdit();
		}
	}
}

void CDlg::ButtonClickDatagrid4(short ColIndex)
{
	COleVariant vC,vBk;
	short i;
	if(IsEnableRec(ptrRs4)){		
		vBk = ptrRs4->GetBookmark();
		switch(ColIndex){
		case 1:
			{
				BOOL bFndC;
				CString strFndC;
				bFndC   = TRUE;
				strFndC = _T("");
				i=0;
				vC = GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				strFndC=vC.bstrVal;
				strFndC.TrimLeft();
				strFndC.TrimRight();

				BeginWaitCursor();

			//		m_SlpDay.SetDate(t1);
					HMODULE hMod;
					hMod=AfxLoadLibrary(L"Stg.dll");
					typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
					pDialog func=(pDialog)GetProcAddress(hMod,"startStg");
					(func)(m_strNT, ptrCnn,strFndC,bFndC);

			//		m_SlpDay.SetDate(t1);

					AfxFreeLibrary(hMod);

				EndWaitCursor();
			}
			break;
		}
		try{
			ptrRs4->Requery(adCmdText);//adCmdStoredProc
			m_DataGrid4.Refresh();
			InitDataGrid4(m_DataGrid4,ptrRs4);
			if(!IsEmptyRec(ptrRs4)){
				try{
					ptrRs4->PutBookmark(vBk);
				}
				catch(...){
					ptrRs4->MoveLast();
				}
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.ErrorMessage());
		}
	}
}
