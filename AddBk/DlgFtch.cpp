// DlgFtch.cpp : implementation file
//

#include "stdafx.h"
#include "DlgFtch.h"


// CDlgFtch dialog

IMPLEMENT_DYNAMIC(CDlgFtch, CDialog)

CDlgFtch::CDlgFtch(CWnd* pParent /*=NULL*/)
	: CDialog(CDlgFtch::IDD, pParent)
	, m_Edit1(_T(""))
{
	m_pDlg = NULL;

}
CDlgFtch::CDlgFtch(CDialog* pDlg)
	: m_Edit1(_T(""))
{
	m_pDlg = pDlg;
}

CDlgFtch::~CDlgFtch()
{
}

void CDlgFtch::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT1, m_Edit1);
}


BEGIN_MESSAGE_MAP(CDlgFtch, CDialog)
	ON_BN_CLICKED(IDC_BUTTON1, &CDlgFtch::OnBnClickedButton1)
	ON_BN_CLICKED(IDOK, &CDlgFtch::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CDlgFtch::OnBnClickedCancel)
END_MESSAGE_MAP()


// CDlgFtch message handlers

BOOL CDlgFtch::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString str;
	str.LoadStringW(IDS_STRING9030);
	GetDlgItem(IDC_STATIC_EDIT1)->SetWindowText(str);

	str.LoadStringW(IDS_STRING9031);
	this->SetWindowTextW(str);
	this->MoveWindow(0,0,625,90);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
BOOL CDlgFtch::Create()
{
	return CDialog::Create(CDlgFtch::IDD);
}

void CDlgFtch::OnBnClickedButton1()
{
	if(m_pDlg!=NULL){
		UpdateData(true);
		m_pDlg->PostMessageW(WM_FTCH);
	}
}

void CDlgFtch::OnBnClickedOk()
{
	if(m_pDlg!=NULL){
		m_pDlg->PostMessageW(WM_CLOSEFTCH,IDOK);
	}
	else{
		CDialog::OnOK();
	}
}

void CDlgFtch::OnBnClickedCancel()
{
	if(m_pDlg!=NULL){
		m_pDlg->PostMessageW(WM_CLOSEFTCH,IDCANCEL);
	}
	else{
		OnCancel();
	}
}
