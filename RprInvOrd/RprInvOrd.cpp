// RprInvOrd.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>

#include "RecSetProc.h"
#include "BlokRtf.h"
#include "Words.h"
//#include "CmdTargetPrn.h"
//#include "Cnvrt16.h"
//#include <ctype.h>

#ifdef _MANAGED
#error Please read instructions in RprInvOrd.cpp to compile with /clr
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


static AFX_EXTENSION_MODULE RprInvOrdDLL = { NULL, NULL };

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
		TRACE0("RprInvOrd.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(RprInvOrdDLL, hInstance))
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

		new CDynLinkLibrary(RprInvOrdDLL);

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
		TRACE0("RprInvOrd.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(RprInvOrdDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startRprInvOrd(CString strNT, _ConnectionPtr ptrCon, char* strCd, char* strPthOt, char* strAc, WCHAR* strWDoc){

//	HINSTANCE hInstResClnt;

	//hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle("RprInvOrd.dll"));
//	AfxSetResourceHandle(::GetModuleHandle(L"RprInvOrd.dll"));
//AfxMessageBox(_T("RprInvOrd.dll"));
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
	CString strFlNmInvOrd;
	CString strPath,strAdr;
	
	strPath = _T(""); //_T("D:\\MyProjects\\Strg\\Debug\\"); 
	s = m_strDate;
	s.Remove('-');

	strFlNmInvOrd = strPathOut;
	strFlNmInvOrd+=	_T("ТН_");
	strFlNmInvOrd+= m_strNumDec + _T("_");
	strFlNmInvOrd+= s + _T(".rtf");

	int lw;
	lw = MultiByteToWideChar(CP_ACP,0,strFlNmInvOrd,-1,NULL,0);
//	s.Format(_T("lw = %i"),lw);
//	AfxMessageBox(s);
	WCHAR* strW = new WCHAR[lw];
	MultiByteToWideChar(CP_ACP,0,strFlNmInvOrd,-1,strW,lw);
//	strNZ = m_strNZ;
	wcscpy_s(strWDoc,lw,strW);
//	strNZ=strW;
//AfxMessageBox("exit from RprCmdTs");
//AfxMessageBox(strNZ);
	delete strW;

//	strDoc = strFlNmInvOrd;
//AfxMessageBox(strFlNmInvOrd);
// Входящие шаблоны
	CStdioFile inHInvOrd,inMInvOrd,inFTInvOrd,inFInvOrd;  // Тов. нак.
// Исходящий файл
	CStdioFile otInvOrd;
	const int iMAX = 1024;
	const int iMIN = 512;
/*
	cn = NULL;
	cn.CreateInstance(__uuidof(Connection));

	cn->ConnectionString = m_bsConn;
	cn->Open("","","",adConnectUnspecified);
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
//	strSql = L"QT41";

/*	rs = NULL;
	rs.CreateInstance(__uuidof(Recordset));
	rs->CursorLocation = adUseClient;
*/
	strCur = m_strNum + _T(",'");
	strCur+= m_strDate + _T("'");

/*	if(m_bBN){
		bsCmd= _T(L"RT102SumNum ");  
		bsCmd+=(_bstr_t)strCur;
	}
	else{
		bsCmd= _T(L"RT24SumNum ");  
		bsCmd+=(_bstr_t)strCur;
	}
*/
	strSql = _T("RT24NSumNum ");	// RT24SumNum 
	strSql+= m_strNum + _T(",'");
	strSql+= m_strDate + _T("'");

/*	rsWe = NULL;
	rsWe.CreateInstance(__uuidof(Recordset));
	rsWe->CursorLocation = adUseClient;
*/
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
	int iRpl = 52;	// 21 строки которые будут меняться (на 1 больше)
	aHd.SetSize(iRpl);
	for(i=0;i<iRpl;i++){
		aHd.SetAt(i,_T(""));
	}


	aHd.SetAt(17,m_strNumDec);
	aHd.SetAt(18,m_strDecDate);
	aHd.SetAt(19,m_strNumDec);
	aHd.SetAt(20,m_strDecDate);
	aHd.SetAt(37,_T("директор"));
	aHd.SetAt(39,_T("Головина О.В."));	//Фильцов Д.Н.
	aHd.SetAt(42,_T("Радостева М.С."));	//Ермакова Е.В

	ptrCmd1->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);

		if(!pR->IsEmptyRec(ptrRs1)){


/*			bsCmd = _T("RT38_1 ");	//Мы
			bsCmd+= (_bstr_t)s + _T(",2,");
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
				if(!pR->IsEmptyRec(ptrRs2)){

/*					vC = ptrRs2->GetSource();
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
AfxMessageBox(s);
*/
					i=1;	//Вид собственности
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe1 = vC.bstrVal;
					m_strWe1+=_T(" ");

					i=2;	//Наименование
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe1+= s + _T(", ");


					i=3;	//Индекс
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe1+= s + _T(", ");


					i=4;	//Адрес юридический
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe1+= s + _T(", ");

//AfxMessageBox(m_strWe1);
//					m_Ch16.OnCnvrt16(m_strWe1,strO);
//					m_strWe1 = strO;

//					strO=_T("");

//AfxMessageBox(_T("m_strWe1 = ")+aHd[1]);
					i=5;	//Тел
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
//					m_strWe2+=_T("т. ");
//					m_strWe2+=s + _T(", Р.сч ");
					m_strWe1+= s;

					aHd.SetAt(1,m_strWe1);  //m_strWe1
					aHd.SetAt(11,m_strWe1); 

					i=6;	//Р.сч
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
//					m_strWe2+= s + _T(" в ");
					m_strWe2+= s;


					i=7;	//Банк
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					m_strWe2+= s;

//					m_Ch16.OnCnvrt16(m_strWe2,strO);
//					m_strWe2 = strO;

					i=11;	//БИК
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
//					m_strWe2+=_T(" ")+ s;
					m_strWe2+=s;

					i=12;	//Кор.сч.
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
//					m_strWe2+= _T(" ")+s;
					m_strWe2+=s;

					aHd.SetAt(3,m_strWe2); 
//					strO=_T("");

//AfxMessageBox(_T("m_strWe2 = ")+aHd[3]);
//return FALSE;

					i=8;	//ИНН
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					m_strInd= s;

					i=9;	//ОКОНХ
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					m_strOKOHX= s;

					i=10;	//ОКОНХ
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					m_strOKOH= s;

				}
			}
			catch(_com_error& e){
				AfxMessageBox(e.ErrorMessage());
			}

			ptrRs2->Close();
//-------------------------------------------Про Покупателя

			i = 8; //Операция 13; //32;
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimRight();
			s.TrimLeft();
//AfxMessageBox(s);
			aHd.SetAt(15,s);  // Основание

/*			i = 15;	 // 14	     Код КонтрАгента
			vC = pR->GetValueRec(rs,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;

			bsCmd = _T("RT38_1 ");	//Они
			bsCmd+= (_bstr_t)s + _T(",2");
*/
			i = 10;		//15 Код КонтрАгента покупателя
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
			s.TrimLeft();
			s.TrimRight();
			strCst = s;

			strSql = _T("RT38Cst_1 ");	//Они
			strSql+= strCst + _T(",2,");
			strSql+= _T("0");

//AfxMessageBox(strSql);
			ptrCmd3->CommandText = (_bstr_t)strSql;

			if(ptrRs3->State==adStateOpen) ptrRs3->Close();
			ptrRs3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			if(!pR->IsEmptyRec(ptrRs3)){

				m_strVnd = _T("№")+strCst;
				m_strVnd+= _T(", ");

				i=1;	//Вид собственности
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s +_T(" ");

				i=2;	//Наименование
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s + _T(", ");

				i=3;	//Индекс
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s + _T(", ");

				i=4;	//Адрес юридический
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s + _T(", ");

				i=5;	//Тел
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+=_T("т. ");
//				m_strVnd+=s + _T(", Р.сч ");
				m_strVnd+=s;

				i=6;	//Р.сч
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+= s + _T(" в ");
				m_strVnd+=s;

				i=7;	//Банк
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+= s + _T(",БИК ");
				m_strVnd+=s;

				i=11;	//БИК
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+=s + _T(", к/с ");
				m_strVnd+=s;

				i=12;	//Кор. счёт
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+=s;

				aHd[13]=m_strVnd; 

//AfxMessageBox(aHd[9]);

				i=8;	//ИНН
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strIndVnd= s;

				i=9;	//ОКОНХ
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strOKOHXV= s;

				i=10;	//ОКОНХ
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strOKOHV= s;
			}
//------------------------------------- про Грузополучателей
			strSql = _T("RT38Cnsgr ");		
			strSql+= m_strCR   + _T(",");	//Код грузополучателя
			strSql+= m_strCRAdr+ _T(",");	//Код адреса доставки
			strSql+= _T("0");				//полная инф

//AfxMessageBox(strSql);
			ptrCmd4->CommandText = (_bstr_t)strSql;

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
				m_strCnsgr+= s + _T(", ");

				i=3;	//Индекс
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+= s + _T(", ");

				i=4;	//Адрес юридический
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+= s + _T(", ");

				i=5;	//Тел
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+= s;
//				m_strCnsgr+=_T("т. ");
//				m_strCnsgr+=s + _T(", Р.сч ");

				i=6;	//Р.сч
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+= s;
//				m_strCnsgr+= s + _T(" в ");

				i=7;	//Банк
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+= s;
//				m_strCnsgr+= s + _T(",БИК ");

				i=11;	//БИК
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+=s;
//				m_strCnsgr+=s + _T(", к/с ");

				i=12;	//Кор. счёт
				vC=pR->GetValueRec(ptrRs4,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strCnsgr+=s;

				aHd[9] =m_strCnsgr; 

//AfxMessageBox(aHd[9]);
			}

		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
	}
//-----------------Конец про Продавцов и Покупателей
/*	for(int z=0;z<aHd.GetSize();z++){
		s.Format(_T("aHd [%i] = %s"),z,aHd[z]);
		AfxMessageBox(s);
	}
return FALSE;
*/
//************************* Работа с файлами
	int iLnInF=0;
	int iLnCur=0;
	int curPos = 0;
	try{
//AfxMessageBox(_T("try otInvOrd.Open "));
		otInvOrd.Open(strFlNmInvOrd,CFile::modeCreate|CFile::modeReadWrite|CFile::typeText);
		s = strPath + _T("HInvOrd.rtf");
		inHInvOrd.Open(s,CFile::modeRead|CFile::shareDenyNone);
		iLnInF = inHInvOrd.GetLength();

		CString strBuf=_T("");
		CString strRpl=_T("");
		int j=1;
		CString strFnd;
		bool theEnd=true;
//---------------Заполняю ШАПКУ
		strFnd.Format(_T("~%i~"),j);
		while(theEnd/*j<=19*/){ //iLnCur<iLnInF
			inHInvOrd.ReadString(strBuf);

			s=strBuf;

			strFnd.Format(_T("~%i~"),j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);

			curPos = inHInvOrd.GetPosition();
//			if(j==3 ||
//			   j ==16	){
//AfxMessageBox(_T("ВХОД = ")+s+'\n'+'\n'+_T("ВЫХОД = ")+strBuf+'\n'+'\n'+strFnd);
//			}
//AfxMessageBox(otInvOrd.GetFilePath());
			otInvOrd.WriteString(strBuf);
//AfxMessageBox("ok2");
			if(j==21){
				theEnd = false;
//				AfxMessageBox("ok");
			}
		}

//	m_Blk.Control(inHInvOrd,400);


//AfxMessageBox("Последний сброс шапки \n"+strBuf);

//---------------Конец ШАПКИ
//		s.Format("после Шапки curPos = %i",curPos);
//	AfxMessageBox(s);
//		m_Blk.Control(inHInvOrd,20);
														// {\\fs16 1\\cell
		m_Blk.GetStdPWt(strBuf, inHInvOrd,otInvOrd, _T("{\\fs16 1\\cell"),curPos);
//		s.Format("После {\\fs16 1\\cell  1 GetPosWrt curPos = %i",curPos);
//AfxMessageBox(s);
//		m_Blk.Control(inHInvOrd,20);


		m_Blk.GetStdPWt(strBuf, inHInvOrd,otInvOrd, _T("\\pard"),curPos);
//		m_Blk.Control(inHInvOrd,300);
		curPos++;
		inHInvOrd.Seek(curPos,CFile::begin);

		m_Blk.GetStdPWt(strBuf, inHInvOrd,otInvOrd, _T("\\pard"),curPos);
//		m_Blk.Control(inHInvOrd,600);

		CString strSrc=_T("");

		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\pard \\ql"),curPos,strSrc);
		curPos++;
		strSrc+=_T("\\");
		inHInvOrd.Seek(curPos,CFile::begin);
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\pard \\ql"),curPos,strSrc);

//		s.Format("Поиск strSrc После 2 GetLStr curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc = "+strSrc);
//		m_Blk.Control(inHInvOrd,100);


		CString strPst=_T("");

		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\trowd"),curPos,strPst);
		curPos++;
		inHInvOrd.Seek(curPos,CFile::begin);
		strPst+=_T("\\");
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\trowd"),curPos,strPst);

//		s.Format("Поиск strPst  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strPst = "+strPst);
//		m_Blk.Control(inHInvOrd,200);

//AfxMessageBox("Поиск strPst Вхожу в 4 GetLStr ищу \\pard");
//**************Заполняется САМА ТАБЛИЦА по строкам long NUM см. ниже
		CString strSrcNew;
		long Num = ptrRs1->GetRecordCount();//rs
		s = m_W.GetWords(Num,false,false,true);
		aHd.SetAt(21,s);  // Порядковых номеров
//s.Format(" Num = %i",Num);
		s  = aHd.GetAt(21);
//AfxMessageBox(s);
		long i8,i10;
		i8=i10=0;
		double d12,d14,d15;
		d12=d14=d15=0;
//CString q;
		for(int x=1;x<=Num;x++){
//otInvOrd.WriteString(strPrv);
//AfxMessageBox(strSrc);
/*			q.Format("%i  %i",x,Num);
AfxMessageBox("Раб 0 "+q);
*/									//rs
			strSrcNew = pR->OnFillRow(ptrRs1,strSrc,15,x, i8, i10, d12, d14, d15);
//AfxMessageBox(strSrcNew);
//AfxMessageBox("Раб 1 "+q);

			otInvOrd.WriteString(strSrcNew);
//AfxMessageBox("Раб 2 "+q);
			otInvOrd.WriteString(strPst);
//AfxMessageBox("Раб 3 "+q);
			ptrRs1->MoveNext();
//AfxMessageBox("Раб 4 "+q);
		}
//************* ТАБЛИЦА ЗАПОЛНЕНА

		s = m_W.GetWords(i8,false,false,true);
		aHd.SetAt(25,s);  // Всего мест
//AfxMessageBox(s);
		CString s1,s2,s3;
						
		s = m_W.GetWords(d15,false,true,true);
		CString kop=_T("00");
		kop = s.Right(3);
		aHd.SetAt(35,kop);
		s1=s.Left(s.GetLength()-kop.GetLength()); // без копеек
//AfxMessageBox(s1);
		if(s1.GetLength()>45){
			int tPos=0;
			for(int k=1;k<=5;k++){
				if((tPos=s1.Find(_T(" "),tPos))!=-1){
					s3=s1.Left(tPos);
					tPos++;
				}
			}
	//		s1=s1.Left(tPos);
			aHd.SetAt(32,s3);  // Всего отпущено на сумму 1 строка
			s2 = s1.Right(s1.GetLength()-s3.GetLength());
			aHd.SetAt(34,s2);  // Всего отпущено на сумму 2 строка
		}
		else{
			aHd.SetAt(32,s1);  // Всего отпущено на сумму 1 строка
		}
//AfxMessageBox(aHd[32]);
		m_Blk.GetStdPWt(strBuf, inHInvOrd,otInvOrd, _T(" \\cell \\cell"),curPos);
//		m_Blk.Control(inHInvOrd,300);

		strSrc=_T("");
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\pard \\ql"),curPos,strSrc);
		CString strNew;
		strNew = pR->OnFillTtl(strSrc,i8,i10,d12,d14,d15,8);
//AfxMessageBox(_T("strSrc = ")+_T('\n')+strSrc+_T('\n')+_T("strNew = ")+'\n'+strNew);
		otInvOrd.WriteString(strNew);

		m_Blk.GetStdPWt(strBuf, inHInvOrd,otInvOrd, _T(" \\cell \\cell"),curPos);

		strSrc=_T("");
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\pard \\qr"),curPos,strSrc);
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("{\\fs16 \\cell \\cell }"),curPos,strSrc);
		m_Blk.GetStdPWS(strBuf, inHInvOrd,otInvOrd, _T("\\pard"),curPos,strSrc);
		pR->m_bLRG = TRUE;
		strNew = pR->OnFillTtl(strSrc,i8,i10,d12,d14,d15,8);
//AfxMessageBox(_T("strSrc = ")+'\n'+strSrc+'\n'+"strNew = "+'\n'+strNew);
		otInvOrd.WriteString(strNew);
		pR->m_bLRG = FALSE;

//	m_Blk.Control(inHInvOrd,300);

		while(inHInvOrd.ReadString(strBuf)){ 

//			s= strBuf;
			strFnd.Format(_T("~%i~"),j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
//AfxMessageBox(_T("ВХОД = ")+s+'\n'+'\n'+_T("ВЫХОД = ")+strBuf+'\n'+'\n'+strFnd);
			otInvOrd.WriteString(strBuf);
		}
		otInvOrd.Close();
		inHInvOrd.Close();
	}
	catch(CFileException* fx){
		TCHAR buf[255];
		fx->GetErrorMessage(buf,255);
//		AfxMessageBox(buf);

		fx->Delete();

		otInvOrd.Close();
		inHInvOrd.Close();

		delete pR;
//		AfxSetResourceHandle(hInstResClnt);

		return FALSE;

	}
//************************ Конец работы с файлами
	delete pR;
//c.EndWaitCursor();
//	AfxSetResourceHandle(hInstResClnt);
	if(ptrRs1->State==adStateOpen) ptrRs1->Close();
	ptrRs1=NULL;
	if(ptrRs2->State==adStateOpen) ptrRs2->Close();
	ptrRs2=NULL;
	if(ptrRs3->State==adStateOpen) ptrRs3->Close();
	ptrRs3=NULL;

	return TRUE;
}
#ifdef _MANAGED
#pragma managed(pop)
#endif

