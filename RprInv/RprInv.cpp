// RprInv.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>

#include "RecSetProc.h"
#include "BlokRtf.h"
#include "Words.h"

#ifdef _MANAGED
#error Please read instructions in RprInv.cpp to compile with /clr
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


static AFX_EXTENSION_MODULE RprInvDLL = { NULL, NULL };

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
		TRACE0("RprInv.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(RprInvDLL, hInstance))
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

		new CDynLinkLibrary(RprInvDLL);

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
		TRACE0("RprInv.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(RprInvDLL);
	}
	return 1;   // ok
}


extern "C" _declspec(dllexport) BOOL startRprInv(CString strNT, _ConnectionPtr ptrCon, char* strCd, char* strPthOt, char* strAc, WCHAR* strWDoc){

//	HINSTANCE hInstResClnt;
//	hInstResClnt = AfxGetResourceHandle();
	AfxSetResourceHandle(::GetModuleHandle("RprInv.dll"));
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

	ptrCnn    = ptrCon;	//Соединение
	
	CRecSetProc* pR=new CRecSetProc;
	CString m_strNum,m_strDate,m_strDecDate,s,sEnd,strCopy;

	CString m_strWe1,m_strWe2,m_strInd,m_strOKOHX,m_strOKOH;
	CString m_strVnd,m_strIndVnd,m_strOKOHXV,m_strOKOHV;
	CString strOp,strMng,strCst;
	CString m_strNumDec,strSls,strO,m_strCnsgr;

//	CString m_strBN;
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
	CString strFlNmInv;
	CString strPath,strAdr;
	
	strPath = _T(""); //_T("D:\\MyProjects\\Strg\\Debug\\"); 
	s = m_strDate;
	s.Remove('-');

	strFlNmInv = strPathOut;
	strFlNmInv+=	_T("СФ_");
	strFlNmInv+= m_strNum + _T("_");
	strFlNmInv+= s + _T(".rtf");

	int lw;
	lw = MultiByteToWideChar(CP_ACP,0,strFlNmInv,-1,NULL,0);
//	s.Format(_T("lw = %i"),lw);
//	AfxMessageBox(s);
	WCHAR* strW = new WCHAR[lw];
	MultiByteToWideChar(CP_ACP,0,strFlNmInv,-1,strW,lw);
//	strNZ = m_strNZ;
	wcscpy_s(strWDoc,lw,strW);
//	strNZ=strW;
//AfxMessageBox("exit from RprCmdTs");
//AfxMessageBox(strNZ);
	delete strW;
//	strDoc = strFlNmInv;
// Входящие шаблоны
	CStdioFile inHInv;  // Счёт-Фактура.
// Исходящий файл
	CStdioFile otInv;
	const int iMAX = 1024;
	const int iMIN = 512;

/*	cn = NULL;
	cn.CreateInstance(__uuidof(Connection));

	cn->ConnectionString = m_bsConn;
	cn->Open("","","",adConnectUnspecified);

	rs = NULL;
	rs.CreateInstance(__uuidof(Recordset));
	rs->CursorLocation = adUseClient;

	rsWe = NULL;
	rsWe.CreateInstance(__uuidof(Recordset));
	rsWe->CursorLocation = adUseClient;
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
	int iRpl = 21;	//  строки которые будут меняться (на 1 больше)
	aHd.SetSize(iRpl);
	for(i=0;i<iRpl;i++){
		aHd.SetAt(i,_T(""));
	}

	strCur = m_strNum + _T(",'");
	strCur+= m_strDate + _T("'");

	strSql = _T("RT24NSumNum ");	// RT24SumNum  QT24SumNum
	strSql+= m_strNum + _T(",'");
	strSql+= m_strDate + _T("'");


	ptrCmd1->CommandText = (_bstr_t)strSql;
//AfxMessageBox(strSql);
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);

		if(!pR->IsEmptyRec(ptrRs1)){

			i = 8;	//13 Операция
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strOp = vC.bstrVal;
			strOp.TrimLeft(' ');
			strOp.TrimRight(' ');
//AfxMessageBox(strOp);
/*			i = 10;	//15 Код контрагента покупателя
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCtr = vC.bstrVal;
			strCtr.TrimLeft(' ');
			strCtr.TrimRight(' ');
*/
/*			i = 18;  //Код продовца
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
*/
			i = 13;	//19 Менеджер
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strMng = vC.bstrVal;
			strMng.TrimLeft(' ');
			strMng.TrimRight(' ');
//AfxMessageBox(strMng);

/*
			bsCmd = _T("RT38Sls ");	//Мы
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

					i=1;	//Вид собственности
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe1 = s +_T(" ");

					i=2;	//Наименование
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();

					m_strWe1+= s + _T(", ");

					i=5;	//Тел
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
//					m_strWe2+=_T("т. ");
//					m_strWe2+=s + _T(", Р.сч ");
					m_strWe1+= s;
//					m_strWe2+= s;

					aHd.SetAt(4,m_strWe1);


					i=3;	//Индекс
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					CString strInd;
//					m_strWe1+= s + _T(", ");
					strInd = s;

					i=4;	//Адрес юридический
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s = strInd + _T(", ")+s;

					aHd.SetAt(7,s);

					m_strWe1+=_T(",")+ s + _T(",");

//					aHd.SetAt(4,m_strWe1);
					aHd.SetAt(9,m_strWe1);


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
					s.TrimLeft();
					s.TrimLeft();
					m_strWe2+= s;

					i=8;	//ИНН
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strInd= s;
					aHd.SetAt(8,m_strInd);	// ИНН продавца

					i=9;	//ОКОНХ
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strOKOHX= s;

					i=10;	//ОКОНХ
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strOKOH= s;
//AfxMessageBox(m_strWe2);
					i=11;	//БИК
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe2+= s;

					i=12;	//Кор. счёт
					vC=pR->GetValueRec(ptrRs2,i);
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					s.TrimLeft();
					s.TrimRight();
					m_strWe2+= s;
//AfxMessageBox(m_strWe2);

					aHd.SetAt(5,m_strWe2);	//Продавец вторая строка
//					aHd.SetAt(5,_T(""));
				}
			}
			catch(_com_error& e){
				AfxMessageBox(e.ErrorMessage());
			}

			ptrRs2->Close();
//-------------------------------------------Про Покупателя


/*			bsCmd = _T("RT38_1 ");	//Они
			bsCmd+= (_bstr_t)s + _T(",2,");
			bsCmd+= (_bstr_t)strAcc;
*/
/*			bsCmd = _T("RT38Cst ");	//Они
			bsCmd+= (_bstr_t)s + _T(",2,");
			bsCmd+= _T("0");
*/
			i = 10;	//15 Код контрагента покупателя
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			strCst = vC.bstrVal;
			strCst.TrimLeft(' ');
			strCst.TrimRight(' ');

			strSql = _T("RT38Cst_1 ");	//Они
			strSql+= strCst + _T(",2,");
			strSql+= _T("0");

//AfxMessageBox(strSql);
			ptrCmd3->CommandText = (_bstr_t)strSql;

			if(ptrRs3->State==adStateOpen) ptrRs3->Close();
			ptrRs3->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);
			if(!pR->IsEmptyRec(ptrRs3)){

				CString strAdr;

				i=1;	//Вид собственности
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd = s +_T(" ");

				i=2;	//Наименование
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s + _T(", ");

//				aHd.SetAt(12,m_strVnd);		// Покупатель

				i=3;						//Индекс
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+= s + _T(", ");
				strAdr+=s + _T(", ");
 
				i=4;						//Адрес юридический
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
//				m_strVnd+= s + _T(", ");
				strAdr+=s;

				aHd.SetAt(13,strAdr);	// Адрес покупателя 
//				aHd.SetAt(10,aHd.GetAt(12)+_T(" ")+strAdr);

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
				m_strVnd+= s;

				i=7;	//Банк
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s;

				i=11;	//БИК
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s;

				i=12;	//Кор. счёт
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();
				m_strVnd+= s;

				aHd[12] =m_strVnd;			// Покупатель

				i=8;	//ИНН
				vC=pR->GetValueRec(ptrRs3,i);
				vC.ChangeType(VT_BSTR);
				s = vC.bstrVal;
				s.TrimLeft();
				s.TrimRight();

				m_strIndVnd = s + _T("  ( ");
				m_strIndVnd+= m_strNum + _T(", "); 
				m_strIndVnd+= strOp  + _T(", ");
				m_strIndVnd+= strCst + _T(", ");
				m_strIndVnd+= strMng + _T(" )");

				aHd.SetAt(16,m_strIndVnd);	// ИНН покупателя

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

				aHd[10] =m_strCnsgr; 

//AfxMessageBox(aHd[10]);
			}

		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
		delete pR;
//	AfxSetResourceHandle(hInstResClnt);
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
//-----------------Конец про Продавцов и Покупателей
//************************* Работа с файлами
	int curPos = 0;
	CString strBuf = _T("");
	try{
		s = strPath + _T("HInv.rtf");
		otInv.Open(strFlNmInv,CFile::modeCreate|CFile::modeReadWrite|CFile::typeText);
		inHInv.Open(s,CFile::modeRead|CFile::shareDenyNone);

		CString strBuf=_T("");
		CString strRpl=_T("");
		int j=1;
		CString strFnd;
		bool theEnd=true;

//---------------------------------------------Заполняю ШАПКУ
		aHd.SetAt(1,_T(""));	//№ СФ m_strNum

		CString sD,sMY,sDMY;
		sD = m_strDecDate.Left(m_strDecDate.Find(_T(".")));
		sD = _T("\" ")+sD;
		sD+= _T(" \"");
		aHd.SetAt(2,sD);

		sDMY= m_W.GetWrdMnth(m_strDecDate,1,true);
		sMY = sDMY.Mid(sDMY.Find(" "));
		aHd.SetAt(3,sMY);
		aHd.SetAt(17,_T("Головина О.В."));	//Фильцов Д.Н.
		aHd.SetAt(18,_T("Радостева М.С."));				//Ермакова Е.В.

		strFnd.Format("~%i~",j);
		while(theEnd){ 
			inHInv.ReadString(strBuf);

//			s=strBuf;

			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
			curPos = inHInv.GetPosition();

//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otInv.WriteString(strBuf);
			if(j==17){
				theEnd = false;
//				AfxMessageBox("ok");
			}
		}

//AfxMessageBox(_T("Последний сброс шапки \n")+strBuf);

//------------------------------------------------Конец ШАПКИ
//		m_Blk.Control(inHInv,300);

		m_Blk.GetStdPWt(strBuf, inHInv,otInv, _T("\\pard \\ql \\li0"),curPos);
		m_Blk.GetStdPWt(strBuf, inHInv,otInv, _T("\\pard \\ql \\li0"),curPos);
		m_Blk.GetStdPWt(strBuf, inHInv,otInv, _T("\\pard \\ql \\li0"),curPos);
//		s.Format("После \trowd \n\trgaph108  1 GetStdPWt curPos = %i",curPos);
//AfxMessageBox(s);
//		m_Blk.Control(inHInv,100);

		m_Blk.GetStdPWt(strBuf, inHInv,otInv, _T("{\\fs16 \\cell }"),curPos);
//		m_Blk.Control(inHInv,100);

		CString strSrc=_T("");

		m_Blk.GetStdPWS(strBuf, inHInv,otInv, _T("\\pard \\ql"),curPos,strSrc);
//		s.Format("Поиск strSrc После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc = "+strSrc);
//		m_Blk.Control(inHInv,100);

		CString strPst=_T("");

		m_Blk.GetStdPWS(strBuf, inHInv,otInv, _T("\\trowd"),curPos,strPst);
		curPos++;
		inHInv.Seek(curPos,CFile::begin);
		strPst+=_T("\\");
		m_Blk.GetStdPWS(strBuf, inHInv,otInv, _T("\\trowd"),curPos,strPst);
//		s.Format("Поиск strPst После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strPst = "+strPst);
//		m_Blk.Control(inHInv,100);

		long Num = ptrRs1->GetRecordCount();
		CString strNew = _T("");
		double d8,d9;
		d8=d9=0;
		for(int x=1;x<=Num;x++){
//			otInvOrd.WriteString(strPrv); Может это не нужно !!!
//			s.Format(_T("x = %i, Num = %Num"),x,Num);
			strNew = pR->OnFillRow(ptrRs1,strSrc,11,x, d8, d9);
//AfxMessageBox(_T("strSrc = ")+'\n'+strSrc+'\n'+_T("strNew = ")+'\n'+strNew+'\n'+s);
			otInv.WriteString(strNew);
			otInv.WriteString(strPst);
			ptrRs1->MoveNext();
		}
		m_Blk.GetStdPWt(strBuf, inHInv,otInv, _T("{\\fs16 \\cell "),curPos);
//		m_Blk.Control(inHInv,100);

//		curPos++;
//		inHInv.Seek(curPos,CFile::begin);
//		m_Blk.Control(inHInv,300);

		CString strSrc2=_T("");
		m_Blk.GetStdPWS(strBuf, inHInv,otInv, _T("\\pard \\qc"),curPos,strSrc2);
//		s.Format("Поиск strSrc2 После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc2 = "+strSrc2);
//		m_Blk.Control(inHInv,100);

		CString strNew2=_T("");
		pR->m_bLRG = TRUE;
		strNew2 = pR->OnFillTtl(strSrc2,d8,d9,2);
//AfxMessageBox("strNew2 = "+'\n'+strNew2);
		otInv.WriteString(strNew2);
		pR->m_bLRG = FALSE;

		theEnd = true;
		j=17;
		while(theEnd){ 
			inHInv.ReadString(strBuf);

//			s=strBuf;
			strFnd.Format("~%i~",j);
			m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
			curPos = inHInv.GetPosition();

//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
			otInv.WriteString(strBuf);
			if(j==21){
				theEnd = false;
//				AfxMessageBox("ok");
			}
		}
		while(inHInv.ReadString(strBuf)){ 
			otInv.WriteString(strBuf);
		}

		inHInv.Close();
		otInv.Close();
	}
	catch(CFileException* fx){
		TCHAR buf[255];
		fx->GetErrorMessage(buf,255);
//		AfxMessageBox(buf);
		fx->Delete();
	}

	delete pR;
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
