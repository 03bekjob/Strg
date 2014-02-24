// RprOldInvOrd.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>

#include "BlokRtf.h"
#include "Words.h"
#include "RecSetProc.h"

#ifdef _MANAGED
#error Please read instructions in RprOldInvOrd.cpp to compile with /clr
// If you want to add /clr to your project you must do the following:
//	1. Remove the above include for afxdllx.h
//	2. Add a .cpp file to your project that does not have /clr thrown and has
//	   Precompiled headers disabled, with the following text:
//			#include <afxwin.h>
//			#include <afxdllx.h>
#endif

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


static AFX_EXTENSION_MODULE RprOldInvOrdDLL = { NULL, NULL };

#ifdef _MANAGED
#pragma managed(push, off)
#endif

extern "C" int APIENTRY
DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	// Remove this if you use lpReserved
	UNREFERENCED_PARAMETER(lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		TRACE0("RprOldInvOrd.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(RprOldInvOrdDLL, hInstance))
			return 0;

		// Insert this DLL into the resource chain
		// NOTE: If this Extension DLL is being implicitly linked to by
		//  an MFC Regular DLL (such as an ActiveX Control)
		//  instead of an MFC application, then you will want to
		//  remove this line from DllMain and put it in a separate
		//  function exported from this Extension DLL.  The Regular DLL
		//  that uses this Extension DLL should then explicitly call that
		//  function to initialize this Extension DLL.  Otherwise,
		//  the CDynLinkLibrary object will not be attached to the
		//  Regular DLL's resource chain, and serious problems will
		//  result.

		new CDynLinkLibrary(RprOldInvOrdDLL);

		// Sockets initialization
		// NOTE: If this Extension DLL is being implicitly linked to by
		//  an MFC Regular DLL (such as an ActiveX Control)
		//  instead of an MFC application, then you will want to
		//  remove the following lines from DllMain and put them in a separate
		//  function exported from this Extension DLL.  The Regular DLL
		//  that uses this Extension DLL should then explicitly call that
		//  function to initialize this Extension DLL.
		if (!AfxSocketInit())
		{
			return FALSE;
		}
	
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		TRACE0("RprOldInvOrd.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(RprOldInvOrdDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startRprOldInvOrd(CString strNT, _ConnectionPtr ptrCon, char* strCd, char* strPthOt, char* strAc, WCHAR* strWDoc){

//	HINSTANCE hInstResClnt;

	//hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle("RprOldInvOrd.dll"));
//	AfxSetResourceHandle(::GetModuleHandle(L"RprInvOrd.dll"));
	CString curStr,strTmp;
	CString strCod,strAcc,strPathOut;

	strCod	   = strCd;
	strCod.TrimLeft();
	strCod.TrimRight();

	strPathOut = strPthOt;
	strPathOut.TrimLeft();
	strPathOut.TrimRight();

	strAcc     = strAc;
	strAcc.TrimLeft();
	strAcc.TrimRight();

/*AfxMessageBox(strCod);
AfxMessageBox(strPathOut);
AfxMessageBox(strAcc);

return FALSE;
*/
	CString strCur,strSql,strCSls;
	CString m_strNT;
	CString m_strCR;	 		//Код грузополучателя
	CString m_strCRAdr;			//Код адреса доставки
	CString m_strCSA;	 		//Код расч.сч. продовца
	CString m_strCREm;			//Код e-mail
	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1,ptrRs2,ptrRs3,ptrRs4;
	_CommandPtr ptrCmd1,ptrCmd2,ptrCmd3,ptrCmd4;
	_ConnectionPtr ptrCnn;

	COleVariant vC;
	CBlokRtf m_Blk;
	CWords m_W;
//	CCnvrt16 m_Ch16;

//CCmdTargetPrn c;
//c.BeginWaitCursor();
	ptrCnn    = ptrCon;	//Соединение

	CRecSetProc* pR=new CRecSetProc;
	CString m_strNum,m_strDate,m_strDecDate,s,sEnd,strCopy;
	CString m_strNumDec,strSls,strO;

	CString m_strWe1,m_strWe2,m_strInd,m_strOKOHX,m_strOKOH;
	CString m_strVnd,m_strIndVnd,m_strOKOHXV,m_strOKOHV,strCst,m_strCnsgr;

//	char* m_bsConn;
	bool m_bBN;
	int iFind,iNum,i;
										//  № нак
	iFind=strCod.Find('~');
	m_strNum  = strCod.Left(iFind);
	m_strNumDec = m_strNum;
										//  дата
	iNum = strCod.Find('~',iFind+1);
	m_strDecDate = strCod.Mid(iFind+1,iNum-iFind-1);
	iFind = iNum;
										//  код продовца
	iNum = strCod.Find('~',iFind+1);
	strSls = strCod.Mid(iFind+1,iNum-iFind-1);
	iFind = iNum;
		
	iFind   = strCod.ReverseFind('~');
	m_strDate = strCod.Mid(iFind+1);

	m_strNT  = strNT;
//	m_bsConn = bsConn;

//AfxMessageBox(strAcc);
	iFind=strAcc.Find('~');				// код грузополучателя
	m_strCR  = strAcc.Left(iFind);

	iNum = strAcc.Find('~',iFind+1);	// Код адреса доставки
	m_strCRAdr = strAcc.Mid(iFind+1,iNum-iFind-1);
	iFind = iNum;

	iNum = strAcc.Find('~',iFind+1);	// Код расч.сч. продовца
	m_strCSA = strAcc.Mid(iFind+1,iNum-iFind-1);
	iFind = iNum;

	iFind   = strAcc.ReverseFind('~');	// Код e-mail
	m_strCREm = strAcc.Mid(iFind+1);
	
// Имена исходящих файлов
	CString strFlNmOldInvOrd;
	CString strPath;
	
	strPath = _T(""); //_T("D:\\MyProjects\\Strg\\Debug\\"); 
	s = m_strDate;
	s.Remove('-');

	strFlNmOldInvOrd = strPathOut;
	strFlNmOldInvOrd+=	_T("НК_");
	strFlNmOldInvOrd+= m_strNum + _T("_");
	strFlNmOldInvOrd+= s + _T(".rtf");

	int lw;
	lw = MultiByteToWideChar(CP_ACP,0,strFlNmOldInvOrd,-1,NULL,0);
//	s.Format(_T("lw = %i"),lw);
//	AfxMessageBox(s);
	WCHAR* strW = new WCHAR[lw];
	MultiByteToWideChar(CP_ACP,0,strFlNmOldInvOrd,-1,strW,lw);
//	strNZ = m_strNZ;
	wcscpy_s(strWDoc,lw,strW);
//	strNZ=strW;
//AfxMessageBox("exit from RprCmdTs");
//AfxMessageBox(strNZ);
	delete strW;

// Входящие шаблоны
	CStdioFile inOldInvOrd;  // Старая накладная
// Исходящий файл
	CStdioFile otOldInvOrd;

//	const int iMAX = 1024;
//	const int iMIN = 512;

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


	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);


	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);

	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);

	CStringArray aHd;
	int iRpl = 16;	// 21 строки которые будут меняться (на 1 больше)
	aHd.SetSize(iRpl);
	for(i=0;i<iRpl;i++){
		aHd.SetAt(i,_T(""));
	}
	strCur = m_strNum + _T(",'");
	strCur+= m_strDate + _T("'");

	strSql = _T("RT24NOldInvOrd ");
	strSql+= strCur;

	ptrCmd1->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);

		if(!pR->IsEmptyRec(ptrRs1)){

			i = 9;		//Код продовца
			vC=pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strSls = vC.bstrVal;
			strSls.TrimLeft();
			strSls.TrimRight();

			i = 8;		//Код покупателя
			vC=pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCst = vC.bstrVal;
			strCst.TrimLeft();
			strCst.TrimRight();

			i = 10;		//Основание
			vC=pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();

			aHd.SetAt(6,s);		// Основание
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());

		delete pR;
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1=NULL;
		if(ptrRs2->State==adStateOpen) ptrRs2->Close();
		ptrRs2=NULL;
		if(ptrRs3->State==adStateOpen) ptrRs3->Close();
		ptrRs3=NULL;
		if(ptrRs4->State==adStateOpen) ptrRs4->Close();
		ptrRs4=NULL;

		return FALSE;
	}
	CString m_strSls;

	strSql = _T("RT38Sls_1 ");		//Мы RT38Sls
	strSql+= strSls + _T(",2,");	//Юр адрес - 2
	strSql+= m_strCSA;				//Код расч.сч. продовца

	ptrCmd2->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
	try{
//---------------------------------------Про Продовца
		if(ptrRs2->State==adStateOpen) ptrRs2->Close();
		ptrRs2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(!pR->IsEmptyRec(ptrRs2)){

			i=1;	//Вид собственности
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			m_strSls = vC.bstrVal;
			m_strSls+= _T(" ");

			i=2;	//Наименование
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			m_strSls+= s + _T(", ");

			aHd.SetAt(5,m_strSls);	// Отправитель

		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());

		delete pR;
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1=NULL;
		if(ptrRs2->State==adStateOpen) ptrRs2->Close();
		ptrRs2=NULL;
		if(ptrRs3->State==adStateOpen) ptrRs3->Close();
		ptrRs3=NULL;
		if(ptrRs4->State==adStateOpen) ptrRs4->Close();
		ptrRs4=NULL;

		return FALSE;
	}
/*--------------------------- ПОКУПАТЕЛЬ-----------------------*/
	CString m_strCst;
	strSql = _T("RT38Cst_1 ");	//Они
	strSql+= strCst + _T(",2,");
	strSql+= _T("0");
	aHd.SetAt(4,strCst);

//AfxMessageBox(strSql);
	ptrCmd3->CommandText = (_bstr_t)strSql;

	try{
		if(ptrRs3->State==adStateOpen) ptrRs3->Close();
		ptrRs3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(!pR->IsEmptyRec(ptrRs3)){
			i=1;
			vC=pR->GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			m_strCst = vC.bstrVal;
			m_strCst.TrimLeft();
			m_strCst.TrimRight();
			m_strCst+= _T(" ");

			i=2;
			vC=pR->GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			m_strCst+= s;

			aHd.SetAt(3,m_strCst);	// Получатель
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());

		delete pR;
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1=NULL;
		if(ptrRs2->State==adStateOpen) ptrRs2->Close();
		ptrRs2=NULL;
		if(ptrRs3->State==adStateOpen) ptrRs3->Close();
		ptrRs3=NULL;
		if(ptrRs4->State==adStateOpen) ptrRs4->Close();
		ptrRs4=NULL;

		return FALSE;
	}

	aHd.SetAt(1,m_strNum);			// №
	CString sD,sMY,sDMY;
	sD = m_strDecDate.Left(m_strDecDate.Find(_T(".")));
	sD = _T("\" ")+sD;
	sD+= _T(" \"");

	sDMY= m_W.GetWrdMnth(m_strDecDate,1,true);
	sMY = sDMY.Mid(sDMY.Find(" "));
	
	sD+= _T("   ")+sMY;

	aHd.SetAt(2,sD);				// от
//************************* Работа с файлами
	int curPos = 0;
	CString strBuf = _T("");
	try{
		s = strPath + _T("OldInvOrd.rtf");
		otOldInvOrd.Open(strFlNmOldInvOrd,CFile::modeCreate|CFile::modeReadWrite|CFile::typeText);
		inOldInvOrd.Open(s,CFile::modeRead|CFile::shareDenyNone);

		CString strBuf=_T("");
		CString strRpl=_T("");
		int j=1;
		CString strFnd;
		bool theEnd=true;
//---------------------------заполняю Шапку
		strFnd.Format("~%i~",j);
		while(theEnd){ 
			inOldInvOrd.ReadString(strBuf);

//			s=strBuf;

			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
			curPos = inOldInvOrd.GetPosition();

//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otOldInvOrd.WriteString(strBuf);
			if(j==8){
				theEnd = false;
//				AfxMessageBox("ok");
			}
		}
		m_Blk.GetStdPWt(strBuf,inOldInvOrd,otOldInvOrd, _T("{1\\cell "),curPos);
//		m_Blk.Control(inOldInvOrd,300);
		m_Blk.GetStdPWt(strBuf,inOldInvOrd,otOldInvOrd, _T("\\pard \\qc "),curPos);
//		m_Blk.Control(inOldInvOrd,300);

		CString strSrc=_T("");

		m_Blk.GetStdPWS(strBuf,inOldInvOrd,otOldInvOrd, _T("\\pard \\qr"),curPos,strSrc);
		curPos++;
		inOldInvOrd.Seek(curPos,CFile::begin);
		strSrc+=_T("\\");
		m_Blk.GetStdPWS(strBuf,inOldInvOrd,otOldInvOrd, _T("\\pard"),curPos,strSrc);
//		s.Format("Поиск strSrc После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc = "+strSrc);
//		m_Blk.Control(inOldInvOrd,300);

		CString strPst=_T("");

		m_Blk.GetStdPWS(strBuf,inOldInvOrd,otOldInvOrd, _T("\\trowd"),curPos,strPst);
		curPos++;
		inOldInvOrd.Seek(curPos,CFile::begin);
		strPst+=_T("\\");
		m_Blk.GetStdPWS(strBuf,inOldInvOrd,otOldInvOrd, _T("\\trowd"),curPos,strPst);
//		s.Format("Поиск strPst После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strPst = "+strPst);
//		m_Blk.Control(inOldInvOrd,300);

		CString strNew;
		long Num = ptrRs1->GetRecordCount();
		double d6,d12,d14,dT;
		d6=d12=d14=dT=0;
		int i8=0;
		vC = pR->GetValueRec(ptrRs1,6);
		vC.ChangeType(VT_BSTR);
		s = vC.bstrVal;
		aHd.SetAt(13,s);
		for(int x=1;x<=Num;x++){
//			otInvOrd.WriteString(strPrv);
//AfxMessageBox(strSrc);
			strNew = pR->OnFillRow(ptrRs1,strSrc,7,x, i8, d6);
//AfxMessageBox(strNew);

			vC = pR->GetValueRec(ptrRs1,11);
			vC.ChangeType(VT_R8);
			dT = vC.dblVal;
			d12 +=dT;

			vC = pR->GetValueRec(ptrRs1,12);
			vC.ChangeType(VT_R8);
			dT = vC.dblVal;
			d14 +=dT;

			otOldInvOrd.WriteString(strNew);
			otOldInvOrd.WriteString(strPst);
			ptrRs1->MoveNext();
		}

		s.Format("%i",i8);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(8,s);  // Кол-во

		s.Format("%16.2f",d6);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(10,s);  // Сумма

		s.Format("%16.2f",d12);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(12,s);  // Сумма без НДС

		s.Format("%16.2f",d14);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(14,s);  // Сумма НДС

		s = m_W.GetWords(d6,true,true,true);
		aHd.SetAt(15,s);  // на сумму прописью

		while(inOldInvOrd.ReadString(strBuf)){ 

//			s= strBuf;
			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otOldInvOrd.WriteString(strBuf);
		}

		inOldInvOrd.Close();
		otOldInvOrd.Close();
	}
	catch(CFileException* fx){
		TCHAR buf[255];
		fx->GetErrorMessage(buf,255);
//		AfxMessageBox(buf);
		fx->Delete();
	}

//c.EndWaitCursor();
//	AfxSetResourceHandle(hInstResClnt);

	delete pR;
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1=NULL;
	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2=NULL;
	if(ptrRs3->State==adStateOpen) ptrRs3->Close();
	ptrRs3=NULL;
	if(ptrRs4->State==adStateOpen) ptrRs4->Close();
	ptrRs4=NULL;

	return TRUE;
}
#ifdef _MANAGED
#pragma managed(pop)
#endif

