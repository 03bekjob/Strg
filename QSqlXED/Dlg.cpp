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

	, m_OleDateTime1(COleDateTime::GetCurrentTime())
	, m_OleDateTime2(COleDateTime::GetCurrentTime())
	, m_Edit1(_T("0"))
	, m_Edit2(_T("0"))
	, m_Edit3(_T("0"))
	, m_Edit4(_T("0"))
	, m_Edit5(_T("0"))
	, m_Edit6(_T("0"))
	, m_iCdQr(0)
	, m_strSql(_T(""))
	, m_strX(_T(""))
	, m_strHd(_T(""))
{

}

CDlg::~CDlg()
{
	if(ptrRsCmb1->State==adStateOpen) ptrRsCmb1->Close();
	ptrRsCmb1 = NULL;

	AfxSetResourceHandle(hInstResClnt);
}

void CDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_DATACOMBO1, m_DataCombo1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_OleDateTime1);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_OleDateTime2);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
	DDX_Text(pDX, IDC_EDIT2, m_Edit2);
	DDX_Text(pDX, IDC_EDIT3, m_Edit3);
	DDX_Text(pDX, IDC_EDIT4, m_Edit4);
	DDX_Text(pDX, IDC_EDIT5, m_Edit5);
	DDX_Text(pDX, IDC_EDIT6, m_Edit6);
}


BEGIN_MESSAGE_MAP(CDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_OK, &CDlg::OnClickedOk)
	ON_BN_CLICKED(IDC_CANCEL, &CDlg::OnClickedCancel)
	ON_EN_CHANGE(IDC_EDIT2, &CDlg::OnEnChangeEdit2)
	ON_EN_CHANGE(IDC_EDIT1, &CDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT5, &CDlg::OnEnChangeEdit5)
END_MESSAGE_MAP()


// CDlg message handlers

BOOL CDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle(L"QSqlXED.dll"));

	CString s,strSql,strSql2;
//	s.LoadString(IDS_STRING9014);
	this->SetWindowText(m_strHd);


	InitStaticText();

//	int iTBCtrlID;
	short i;
	COleVariant vC;

/*	m_wndToolBar.Create(this, WS_CHILD | WS_VISIBLE | CBRS_TOP | CBRS_TOOLTIPS |
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

	// Положение панелей
	RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST,
				 IDR_TOOLBAR1, reposDefault , rcClientNow);
*/
	m_vNULL.vt = VT_ERROR;
	m_vNULL.scode = DISP_E_PARAMNOTFOUND;

	ptrCmdCmb1 = NULL;
	ptrRsCmb1 = NULL;

	strSql = _T("QT97 ");
	strSql+= m_strX;
	ptrCmdCmb1.CreateInstance(__uuidof(Command));
	ptrCmdCmb1->ActiveConnection = ptrCnn;
	ptrCmdCmb1->CommandType = adCmdText;
	ptrCmdCmb1->CommandText = (_bstr_t)strSql;

	ptrRsCmb1.CreateInstance(__uuidof(Recordset));
	ptrRsCmb1->CursorLocation = adUseClient;
	ptrRsCmb1->PutRefSource(ptrCmdCmb1);
	try{
		m_DataCombo1.putref_RowSource(NULL);
		ptrRsCmb1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		m_DataCombo1.putref_RowSource((LPUNKNOWN) ptrRsCmb1);
		m_DataCombo1.put_ListField(_T("Вид запроса"));
	}
	catch(_com_error& e){
		m_DataCombo1.putref_RowSource(NULL);
		AfxMessageBox(e.ErrorMessage());
	}
	if(IsEnableRec(ptrRsCmb1)){
		i = 0;
		vC=GetValueRec(ptrRsCmb1,i);
		vC.ChangeType(VT_I2);
		m_iCdQr = vC.intVal;
	}
	OnShowObject(m_iCdQr);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
void CDlg::InitStaticText(void)
{
	CString sTxt;
	sTxt.LoadString(IDS_STRING9017);
	GetDlgItem(IDC_STATIC_COMBOBOX1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9018);
	GetDlgItem(IDC_STATIC_DATA)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9019);
	GetDlgItem(IDC_STATIC_EDIT2)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9021);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(sTxt);

	sTxt.LoadString(IDS_STRING9020);
	GetDlgItem(IDC_STATIC_EDIT4)->SetWindowText(sTxt);
	GetDlgItem(IDC_STATIC_EDIT3)->SetWindowText(sTxt);
	GetDlgItem(IDC_STATIC_EDIT6)->SetWindowText(sTxt);
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

	UpdateData();
	CString strC1DB,strC1DE,strCdQr;
	strC1DB = m_OleDateTime1.Format(_T("%Y-%m-%d"));
	strC1DE = m_OleDateTime2.Format(_T("%Y-%m-%d"));
	if(m_Edit1.IsEmpty()) m_Edit1 = _T("0");
	if(m_Edit2.IsEmpty()) m_Edit2 = _T("0");
	if(m_Edit3.IsEmpty()) m_Edit3 = _T("0");
	if(m_Edit4.IsEmpty()) m_Edit4 = _T("0");
	if(m_Edit5.IsEmpty()) m_Edit5 = _T("0");
	if(m_Edit6.IsEmpty()) m_Edit6 = _T("0");
	m_Edit1.Replace(',','.');
	m_Edit2.Replace(',','.');
	m_Edit3.Replace(',','.');
	m_Edit4.Replace(',','.');
//	if(m_strX==_T("32839"))
	strCdQr.Format(_T("%i"),m_iCdQr);
	switch(_wtoi(m_strX)){
	case 32839:
		{
			double dEdit4;
			double dEdit2;
			dEdit4= _wtof(m_Edit4);
			dEdit2= _wtof(m_Edit2);
			if(dEdit4>999.9999 || dEdit2>999.9999){
				AfxMessageBox(IDS_STRING9023,MB_ICONSTOP);
			}
			else{
				m_strSql+=strCdQr + _T(",'");
				m_strSql+=strC1DB + _T("','");	/*Date*/
				m_strSql+=strC1DE + _T("',");	/*Date*/
				m_strSql+=m_Edit1 + _T(",");	/*%*/
				m_strSql+=m_Edit3 + _T(",");	/*%*/
				m_strSql+=m_Edit2 + _T(",");	/*Sum*/
				m_strSql+=m_Edit4;				/*Sum*/

			//AfxMessageBox(m_strSql);
			}
		}
		break;
	case 32844:
		m_strSql+=strCdQr + _T(",");
		m_strSql+=m_Edit5 + _T(",");
		m_strSql+=m_Edit6 + _T(",'");
		m_strSql+=strC1DB + _T("','");	/*Date*/
		m_strSql+=strC1DE + _T("'");	/*Date*/
		break;
	default:
		m_strSql+=strCdQr + _T(",'");
		m_strSql+=strC1DB + _T("','");	/*Date*/
		m_strSql+=strC1DE + _T("',");	/*Date*/
		m_strSql+=m_Edit1 + _T(",");	/*%*/
		m_strSql+=m_Edit3 + _T(",");	/*%*/
		m_strSql+=m_Edit2 + _T(",");	/*Sum*/
		m_strSql+=m_Edit4;				/*Sum*/

	//AfxMessageBox(m_strSql);
		break;
	}	
	OnOK();
}

void CDlg::OnClickedCancel()
{
	OnCancel();
}


BEGIN_EVENTSINK_MAP(CDlg, CDialog)
	ON_EVENT(CDlg, IDC_DATACOMBO1, 1, CDlg::ChangeDatacombo1, VTS_NONE)
END_EVENTSINK_MAP()

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

void CDlg::ChangeDatacombo1()
{
	m_iCdQr = OnChangeCombo(m_DataCombo1,ptrRsCmb1, 1);	
	OnShowObject(m_iCdQr);
}

void CDlg::OnShowObject(int i)
{
	OnShowEmpty();
	switch(_wtoi(m_strX)){
	case 32844:
		switch(i){
		case 1:
			GetDlgItem(IDC_STATIC_EDIT5)->ShowWindow(true);
			GetDlgItem(IDC_EDIT5)->ShowWindow(true);
			GetDlgItem(IDC_STATIC_EDIT6)->ShowWindow(true);
			GetDlgItem(IDC_EDIT6)->ShowWindow(true);
			break;
		case 2:
			GetDlgItem(IDC_STATIC_DATA)->ShowWindow(true);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
			GetDlgItem(IDC_DATETIMEPICKER2)->ShowWindow(true);
			break;
		}
		break;
	default:
		switch(i){
		case 1:
			GetDlgItem(IDC_STATIC_DATA)->ShowWindow(true);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
			break;
		case 2:		// интервал Дат
			GetDlgItem(IDC_STATIC_DATA)->ShowWindow(true);
			GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(true);
			GetDlgItem(IDC_DATETIMEPICKER2)->ShowWindow(true);

			break;
		case 3:		// Сумма
			GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(true);
			GetDlgItem(IDC_EDIT1)->ShowWindow(true);
			GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(true);
			GetDlgItem(IDC_EDIT3)->ShowWindow(true);
			break;
		case 4:		// %
			GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(true);
			GetDlgItem(IDC_EDIT2)->ShowWindow(true);
			GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(true);
			GetDlgItem(IDC_EDIT4)->ShowWindow(true);

			break;

		}
		break;
	}
	
}

void CDlg::OnShowEmpty(void)
{
	GetDlgItem(IDC_STATIC_DATA)->ShowWindow(false);
	GetDlgItem(IDC_DATETIMEPICKER1)->ShowWindow(false);
	GetDlgItem(IDC_DATETIMEPICKER2)->ShowWindow(false);

	GetDlgItem(IDC_STATIC_EDIT2)->ShowWindow(false);
	GetDlgItem(IDC_EDIT2)->ShowWindow(false);
	GetDlgItem(IDC_STATIC_EDIT4)->ShowWindow(false);
	GetDlgItem(IDC_EDIT4)->ShowWindow(false);

	GetDlgItem(IDC_STATIC_EDIT1)->ShowWindow(false);
	GetDlgItem(IDC_EDIT1)->ShowWindow(false);
	GetDlgItem(IDC_STATIC_EDIT3)->ShowWindow(false);
	GetDlgItem(IDC_EDIT3)->ShowWindow(false);

	GetDlgItem(IDC_STATIC_EDIT5)->ShowWindow(false);
	GetDlgItem(IDC_EDIT5)->ShowWindow(false);
	GetDlgItem(IDC_STATIC_EDIT6)->ShowWindow(false);
	GetDlgItem(IDC_EDIT6)->ShowWindow(false);
}

void CDlg::OnEnChangeEdit2()
{
	UpdateData();
	GetDlgItem(IDC_EDIT4)->SetWindowText(m_Edit2);
}

void CDlg::OnEnChangeEdit1()
{
	UpdateData();
	GetDlgItem(IDC_EDIT3)->SetWindowText(m_Edit1);
}

void CDlg::OnEnChangeEdit5()
{
	UpdateData();
	GetDlgItem(IDC_EDIT6)->SetWindowText(m_Edit5);
}
