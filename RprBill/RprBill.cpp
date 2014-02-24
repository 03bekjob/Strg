// RprBill.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>
#include "BlokRtf.h"
#include "Words.h"
#include "RecSetProc.h"

#ifdef _MANAGED
#error Please read instructions in RprBill.cpp to compile with /clr
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


static AFX_EXTENSION_MODULE RprBillDLL = { NULL, NULL };

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
		TRACE0("RprBill.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(RprBillDLL, hInstance))
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

		new CDynLinkLibrary(RprBillDLL);

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
		TRACE0("RprBill.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(RprBillDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startRprBill(CString strNT, _ConnectionPtr ptrCon, char* strCd, char* strPthOt, char* strAc, WCHAR* strWDoc){
//	HINSTANCE hInstResClnt;

	//hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle("RprBill.dll"));
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
	
	CString strFlNmBill;
	CString strPath;

	strPath =  _T("");  //_T("D:\\MyProjects\\Strg\\Debug\\");
	s = m_strDate;//
	s.Remove('-');

	strFlNmBill = strPathOut;
	strFlNmBill+=	_T("СЧ_");
	strFlNmBill+= m_strNum + _T("_");
	strFlNmBill+= s + _T(".rtf");

	int lw;
	lw = MultiByteToWideChar(CP_ACP,0,strFlNmBill,-1,NULL,0);
//	s.Format(_T("lw = %i"),lw);
//	AfxMessageBox(s);
	WCHAR* strW = new WCHAR[lw];
	MultiByteToWideChar(CP_ACP,0,strFlNmBill,-1,strW,lw);
//	strNZ = m_strNZ;
	wcscpy_s(strWDoc,lw,strW);
//	strNZ=strW;
//AfxMessageBox("exit from RprCmdTs");
//AfxMessageBox(strNZ);
	delete strW;

//	strDoc = strFlNmBill;

	m_vNULL.vt = VT_ERROR;
	m_vNULL.scode = DISP_E_PARAMNOTFOUND;

	ptrCmd4 = NULL;
	ptrRs4 = NULL;

	ptrCmd4.CreateInstance(__uuidof(Command));
	ptrCmd4->ActiveConnection = ptrCnn;
	ptrCmd4->CommandType = adCmdText;

	ptrRs4.CreateInstance(__uuidof(Recordset));
	ptrRs4->CursorLocation = adUseClient;
	ptrRs4->PutRefSource(ptrCmd4);

	ptrCmd3 = NULL;
	ptrRs3 = NULL;

	ptrCmd3.CreateInstance(__uuidof(Command));
	ptrCmd3->ActiveConnection = ptrCnn;
	ptrCmd3->CommandType = adCmdText;

	ptrRs3.CreateInstance(__uuidof(Recordset));
	ptrRs3->CursorLocation = adUseClient;
	ptrRs3->PutRefSource(ptrCmd3);

	ptrCmd2 = NULL;
	ptrRs2 = NULL;

	ptrCmd2.CreateInstance(__uuidof(Command));
	ptrCmd2->ActiveConnection = ptrCnn;
	ptrCmd2->CommandType = adCmdText;

	ptrRs2.CreateInstance(__uuidof(Recordset));
	ptrRs2->CursorLocation = adUseClient;
	ptrRs2->PutRefSource(ptrCmd2);

	ptrCmd1 = NULL;
	ptrRs1 = NULL;

	ptrCmd1.CreateInstance(__uuidof(Command));
	ptrCmd1->ActiveConnection = ptrCnn;
	ptrCmd1->CommandType = adCmdText;

	ptrRs1.CreateInstance(__uuidof(Recordset));
	ptrRs1->CursorLocation = adUseClient;
	ptrRs1->PutRefSource(ptrCmd1);

// Входящие шаблоны
	CStdioFile inBill;  // Счёт
// Исходящий файл
	CStdioFile otBill;

	strCur = m_strNum + _T(",'");
	strCur+= m_strDate + _T("'");

	strSql = _T("RT24NBill ");	//RT24Bill
	strSql+= strCur;

//AfxMessageBox(strSql );
//AfxMessageBox(strAcc);
	ptrCmd1->CommandText = (_bstr_t)strSql;
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(!pR->IsEmptyRec(ptrRs1))	{

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
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
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

	CStringArray aHd;
	int iRpl = 25;	//  строки которые будут меняться (на 1 больше)
	aHd.SetSize(iRpl);
	for(i=0;i<iRpl;i++){
		aHd.SetAt(i,_T(""));
	}
	CString m_strSls;
/*----------------------------- ПРОДАВЕЦ-----------------------*/
/*	bsCmd = _T("RT38Sls "); 
	bsCmd+= (_bstr_t)strSls +_T(",");
	bsCmd+= _T("2,");		// Юридический адрес
	bsCmd+= (_bstr_t)strAcc;
*/
	strSql = _T("RT38Sls_1 ");		//Мы RT38Sls
	strSql+= strSls + _T(",2,");	//Юр адрес - 2
	strSql+= m_strCSA;				//Код расч.сч. продовца

	ptrCmd2->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
	try{
		//---------------------------------------Про Продовца
		if(ptrRs2->State==adStateOpen) ptrRs2->Close();
		ptrRs2->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(!pR->IsEmptyRec(ptrRs2))	{
			i=1;
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			m_strSls+= _T(" ");

			i=2;
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			m_strSls+= s;

			aHd.SetAt(1,m_strSls);	// Наименование

			i=3;					//Индекс
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			m_strSls+= _T(",");

			i=4;					//Адрес юридический
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			m_strSls+=s + _T(", ");		

			i=5;					// тел
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
//			m_strSls+=_T("тел.: ") + s;		
			m_strSls+= s;

			aHd.SetAt(2,m_strSls);	// Адрес,тел

			i=8;
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
//			m_strSls=_T("ИНН/КПП ");
			m_strSls+= s;
			aHd.SetAt(3,m_strSls);	// ИНН/КПП

			aHd.SetAt(4,aHd.GetAt(1));
			
			i=6;					// Расч. счёт
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			aHd.SetAt(5,m_strSls);	

			i=7;					//Банк
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			aHd.SetAt(8,m_strSls);	// Банк

			i=11;					//БИК
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			aHd.SetAt(9,m_strSls);	// БИК

			i=12;					// Кор.счёт
			vC=pR->GetValueRec(ptrRs2,i);
			vC.ChangeType(VT_BSTR);
			m_strSls = vC.bstrVal;
			m_strSls.TrimLeft();
			m_strSls.TrimRight();
			aHd.SetAt(11,m_strSls);	// Кор.счёт
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
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
	ptrRs2->Close();
/*--------------------------- ПОКУПАТЕЛЬ-----------------------*/
	CString m_strCst;
/*	bsCmd = _T("RT38Cst "); 
	bsCmd+= (_bstr_t)strCst +_T(",");
	bsCmd+= _T("2,");		// Юридический адрес
	bsCmd+= _T("1");		// Только Наименование
*/
	strSql = _T("RT38Cst_1 ");	//Они
	strSql+= strCst + _T(",2,");
	strSql+= _T("0");

//AfxMessageBox(strSql);
	ptrCmd3->CommandText = (_bstr_t)strSql;

	try{
		if(ptrRs3->State==adStateOpen) ptrRs3->Close();
//AfxMessageBox(strSql);
		ptrRs3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
		if(!pR->IsEmptyRec(ptrRs3)){
			i=1;
			vC=pR->GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			m_strCst = vC.bstrVal;
			m_strCst+= _T(" ");

			i=2;
			vC=pR->GetValueRec(ptrRs3,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			m_strCst+= s;

			aHd.SetAt(15,m_strCst);	// Плательщик
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
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
//------------------------------------- про Грузополучателей
		strSql = _T("RT38Cnsgr ");		
		strSql+= m_strCR   + _T(",");	//Код грузополучателя
		strSql+= m_strCRAdr+ _T(",");	//Код адреса доставки
		strSql+= _T("1");				//не полная инф

//AfxMessageBox(strSql);
		ptrCmd4->CommandText = (_bstr_t)strSql;

		try{
			if(ptrRs4->State==adStateOpen) ptrRs4->Close();
			ptrRs4->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			if(!pR->IsEmptyRec(ptrRs4)){

					i=1;	//Вид собственности
					vC=pR->GetValueRec(ptrRs4,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strCnsgr+= s +_T(" ");

					i=2;	//Наименование
					vC=pR->GetValueRec(ptrRs4,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strCnsgr+= s;

					aHd.SetAt(16,m_strCnsgr);	// Грузополучатель
//AfxMessageBox(m_strCnsgr);
			}
		}
		catch(_com_error& e){
			AfxMessageBox(e.ErrorMessage());
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

	aHd.SetAt(12,m_strNum);			// №
	CString sD,sMY,sDMY;
	sD = m_strDecDate.Left(m_strDecDate.Find(_T(".")));
	sD = _T("\" ")+sD;
	sD+= _T(" \"");

	aHd.SetAt(13,sD);

	sDMY= m_W.GetWrdMnth(m_strDecDate,1,true);
	sMY = sDMY.Mid(sDMY.Find(" "));
	aHd.SetAt(14,sMY);
	aHd.SetAt(23,_T("Головина О.В."));	//Фильцов Д.Н.
	aHd.SetAt(24,_T("Радостева М.С."));	//Ермакова Е.В.

//************************* Работа с файлами
	int curPos = 0;
	CString strBuf = _T("");
	try{
		s = strPath + _T("Bill.rtf");
		otBill.Open(strFlNmBill,CFile::modeCreate|CFile::modeReadWrite|CFile::typeText);
		inBill.Open(s,CFile::modeRead|CFile::shareDenyNone);

		CString strBuf=_T("");
		CString strRpl=_T("");
		int j=1;
		CString strFnd;
		bool theEnd=true;
//---------------------------заполняю Шапку
		strFnd.Format("~%i~",j);
		while(theEnd){ 
			inBill.ReadString(strBuf);

//			s=strBuf;

			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
			curPos = inBill.GetPosition();

//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otBill.WriteString(strBuf);
			if(j==17){
				theEnd = false;
//				AfxMessageBox("ok");
			}
		}
//		m_Blk.Control(inBill,300);

		m_Blk.GetStdPWt(strBuf,inBill,otBill, _T("\\row }\\trowd"),curPos);
//		m_Blk.Control(inBill,300);
		m_Blk.GetStdPWt(strBuf,inBill,otBill, _T("{1\\cell "),curPos);
//		m_Blk.Control(inBill,300);
		m_Blk.GetStdPWt(strBuf,inBill,otBill, _T("\\pard \\qc "),curPos);
//		m_Blk.Control(inBill,300);

		CString strSrc=_T("");

		m_Blk.GetStdPWS(strBuf,inBill,otBill, _T("\\pard \\ql"),curPos,strSrc);
//		s.Format("Поиск strSrc После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc = "+strSrc);
//		m_Blk.Control(inBill,300);

		CString strPst=_T("");

		m_Blk.GetStdPWS(strBuf,inBill,otBill, _T("\\trowd"),curPos,strPst);
		curPos++;
		inBill.Seek(curPos,CFile::begin);
		strPst+=_T("\\");
		m_Blk.GetStdPWS(strBuf,inBill,otBill, _T("\\trowd"),curPos,strPst);
//		s.Format("Поиск strPst После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strPst = "+strPst);
//		m_Blk.Control(inBill,300);

		CString strNew;
		long Num = ptrRs1->GetRecordCount();
//		s = m_W.GetWords(Num,false,false,true);
		s.Format("%i",Num);
		s.TrimRight();
		aHd.SetAt(20,s);  // Порядковых номеров
//AfxMessageBox(s);
		double d17,d18,d19,dT;
		d17=d18=d19=dT=0;
		for(int x=1;x<=Num;x++){
//			otInvOrd.WriteString(strPrv);
//AfxMessageBox(strSrc);
			strNew = pR->OnFillRow(ptrRs1,strSrc,6,x, d17, d18, d19);
//AfxMessageBox(strNew);
			vC = pR->GetValueRec(ptrRs1,5);
			vC.ChangeType(VT_R8);
			dT = vC.dblVal;
			d18 +=dT;

			vC = pR->GetValueRec(ptrRs1,6);
			vC.ChangeType(VT_R8);
			dT = vC.dblVal;
			d19 +=dT;
			otBill.WriteString(strNew);
			otBill.WriteString(strPst);
			ptrRs1->MoveNext();
		}

		s.Format("%16.4f",d17);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(17,s);  // Итого

		s.Format("%16.4f",d18);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(18,s);  // Итого НДС

		s.Format("%16.2f",d19);
		s.TrimRight();
		s.TrimLeft();
		aHd.SetAt(19,s);  // Всего к оплате
		aHd.SetAt(21,s);  // на сумму

		s = m_W.GetWords(d19,true,true,true);
		aHd.SetAt(22,s);  // на сумму прописью
		

		while(inBill.ReadString(strBuf)){ 

//			s= strBuf;
			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otBill.WriteString(strBuf);
		}
		inBill.Close();
		otBill.Close();
	}
	catch(CFileException* fx){
		TCHAR buf[255];
		fx->GetErrorMessage(buf,255);
//		AfxMessageBox(buf);
		fx->Delete();
	}

	delete pR;
//c.EndWaitCursor();
//	AfxSetResourceHandle(hInstResClnt);
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

