// StrgDoc.cpp : implementation of the CStrgDoc class
//

#include "stdafx.h"
#include "Strg.h"

#include "StrgDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CStrgDoc

IMPLEMENT_DYNCREATE(CStrgDoc, CDocument)

BEGIN_MESSAGE_MAP(CStrgDoc, CDocument)
	ON_COMMAND(ID_FILE_SEND_MAIL, &CStrgDoc::OnFileSendMail)
	ON_UPDATE_COMMAND_UI(ID_FILE_SEND_MAIL, &CStrgDoc::OnUpdateFileSendMail)
END_MESSAGE_MAP()


// CStrgDoc construction/destruction

CStrgDoc::CStrgDoc()
{
	CoInitialize(NULL);						//SSPI																	//(local)
	m_Conn="Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ORTK;Data Source=HOMECMP";
	ptrConn = NULL;
	ptrConn.CreateInstance(__uuidof(Connection));
	ptrConn->ConnectionString = (LPCTSTR)m_Conn;
	try{
		ptrConn->Open("","","",adConnectUnspecified);
	}
	catch(_com_error& e){
		AfxMessageBox(e.Description(),MB_ICONSTOP);
	}
	max = 25;
	GetUserNameBek(m_strNameNT,m_strNameSQL);

}

CStrgDoc::~CStrgDoc()
{
	if(!ptrConn){
		ptrConn->Close();
		ptrConn = NULL;
	}
	CoUninitialize( );
}

BOOL CStrgDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}




// CStrgDoc serialization

void CStrgDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CStrgDoc diagnostics

#ifdef _DEBUG
void CStrgDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CStrgDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG


// CStrgDoc commands
void CStrgDoc::GetUserNameBek(CString &strNT, CString &strSQL)
{
	CString strCmp;
	CString strUsr;

	GetUserNameW((LPWSTR)m_UserName,&max);
//AfxMessageBox((LPWSTR)m_UserName);
	strUsr = (LPWSTR)m_UserName;
//	strUsr = strUsr.Left(13);

/*	GetComputerName(m_CompName,&max);
	strCmp=m_CompName;
	strCmp = strCmp.Left(11);
*/
	strNT = strUsr;
/*
	strNT+= _T("_");
	strNT+= strCmp;
*/
//---------------------------------
//AfxMessageBox(strNT);
/*	WORD wVersionRequested;
	WSADATA wsaData;
	wVersionRequested = MAKEWORD(1, 0);
	int err = WSAStartup(wVersionRequested, &wsaData);
	if(err == 0)
	  {
		char hn[1024];
		if(gethostname((char *)&hn, 1024))
		  {
			int err = WSAGetLastError();
			MessageBeep(MB_OK);
		  };
		strNT+=_T("_");
		strCmp=hn;
		strCmp=strCmp.Left(11);
		strNT+=strCmp;
AfxMessageBox(strNT);
	  }
*/
//	AfxMessageBox(m_strNameNT);
/*	if (m_db.IsOpen()) {
		::SQLGetInfo(m_db.m_hdbc,SQL_USER_NAME,(PTR)m_UserName,(short) max,&cbData);
	}
	strSQL=m_UserName;
	*/
//	AfxMessageBox(m_strNameSQL);

}
