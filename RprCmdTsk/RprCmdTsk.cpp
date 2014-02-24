// RprCmdTsk.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include <afxdllx.h>
#include "BlokRtf.h"
#include "Words.h"
#include "RecSetProc.h"

#include <Afxtempl.h>
//#include <windows.h>
#include <lm.h>
#ifdef _MANAGED
#error Please read instructions in RprCmdTsk.cpp to compile with /clr
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


static AFX_EXTENSION_MODULE RprCmdTskDLL = { NULL, NULL };
//typedef CArray<char*,char*> aChar;
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
		TRACE0("RprCmdTsk.DLL Initializing!\n");
		
		// Extension DLL one-time initialization
		if (!AfxInitExtensionModule(RprCmdTskDLL, hInstance))
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

		new CDynLinkLibrary(RprCmdTskDLL);

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
		TRACE0("RprCmdTsk.DLL Terminating!\n");

		// Terminate the library before destructors are called
		AfxTermExtensionModule(RprCmdTskDLL);
	}
	return 1;   // ok
}
extern "C" _declspec(dllexport) BOOL startRprCmdTsk(CString strNT, _ConnectionPtr ptrCon, char* strCd, char* strPthOt, char* strAc, WCHAR* strNZ){

//	HINSTANCE hInstResClnt;
//	hInstResClnt = AfxGetResourceHandle();
//AfxMessageBox(_T("startRprCmdTsk"));
	AfxSetResourceHandle(::GetModuleHandle("RprCmdTsk.dll"));

//	char* sANZ;
	CStringArray strANZ;
	strANZ.SetSize(1);
	for(int n=0;n<1;n++){
		strANZ.SetAt(n,_T(""));
	}

	CString curStr,strTmp,m_strNZ;
	CString strCod,strAcc,strPathOut;

	m_strNZ    = _T("");
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

//	bPrint =false;
//CCmdTarget c;
//c.BeginWaitCursor();

	_variant_t m_vNULL;
	_RecordsetPtr ptrRs1;
	_CommandPtr ptrCmd1;
	_ConnectionPtr ptrCnn;

	COleVariant vC;
	CBlokRtf m_Blk;
	CWords m_W;

	ptrCnn    = ptrCon;	//Соединение

	CRecSetProc* pR=NULL;
	CString m_strNum,m_strDate,m_strDecDate,sEnd,strCopy;
	CString s,sPrv,m_strNumDec,strSls;

	CString m_strWe1,m_strWe2,m_strInd,m_strOKOHX,m_strOKOH;
	CString m_strVnd,m_strIndVnd,m_strOKOHXV,m_strOKOHV;

//	CString m_strBN;
//	char* m_bsConn;

	int k=0;							//Число складов для наряд заданий
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

// Имена исходящих файлов
	CString strFlNmCmdTsk;
	CString strPath;
	
	strPath =  _T(""); //_T("D:\\MyProjects\\Strg\\Debug\\"); 
	s = m_strDate;
	s.Remove('-');

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

// Входящие шаблоны
	CStdioFile inCmdTsk;  //Наряд-Задание
// Исходящий файл
	CStdioFile otCmdTsk;

	const int iMAX = 1024;
	const int iMIN = 512;

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

	CStringArray aHd;

	int iRpl = 6;	//  строки которые будут меняться (на 1 больше)
	aHd.SetSize(iRpl);
	for(i=0;i<iRpl;i++){
		aHd.SetAt(i,_T(""));
	}

	strCur = m_strNum + _T(",'");
	strCur+= m_strDate + _T("'");

	aHd.SetAt(1,m_strNumDec+_T("      ")+m_strDecDate);

	strSql = _T("RT24NCmdTsk  ");
	strSql+=strCur;
//AfxMessageBox(strSql);

	ptrCmd1->CommandText = (_bstr_t)strSql;
	try{
		if(ptrRs1->State==adStateOpen) ptrRs1->Close();
		ptrRs1->Open(m_vNULL,m_vNULL,adOpenDynamic,adLockOptimistic,adCmdText);

		if(!pR->IsEmptyRec(ptrRs1)){
			i=4;	//Основание
			vC=pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;

			aHd.SetAt(4,s);				// Основание
//AfxMessageBox(s);
			i = 9;
			vC = pR->GetValueRec(ptrRs1,i);
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;

			aHd.SetAt(3,s);		// Покупатель


//			bsCmd = _T("RT38 ");	//Они
//			bsCmd+= (_bstr_t)s + _T(",2");
//AfxMessageBox(bsCmd);
//			if(rsWe->State==adStateOpen) rsWe->Close();
//
//			rsWe->Open(bsCmd,cn->ConnectionString,adOpenKeyset,adLockReadOnly,adCmdText);//adCmdStoredProc
//			if(!pR->IsEmptyRec(rsWe)){

				CString strAdr;

				int Num =ptrRs1->GetRecordCount();

				typedef CArray<int,int&> aRec; //Сколько записей
												 //по складу
				ptrRs1->MoveFirst();	
				vC = pR->GetValueRec(ptrRs1,5); // № Склада
				vC.ChangeType(VT_BSTR);
				sPrv = vC.bstrVal;

				aRec aRecSt;
				int r=1;
				aRecSt.Add(r);
				s=_T("");
				for(i=2;i<=Num;i++){
					ptrRs1->MoveNext();
					vC = pR->GetValueRec(ptrRs1,5); 
					vC.ChangeType(VT_BSTR);
					s = vC.bstrVal;
					if(sPrv!=s){
						k++;
						r=1;
						aRecSt.Add(r);
						sPrv=s;
					}
					else{
						r++;
						aRecSt.SetAt(k,r);
//						s.Format("k = %i, r = %i",k,r);
//AfxMessageBox(s);
					}
				}
				k = aRecSt.GetSize();
//s.Format("k = %i",k);
//AfxMessageBox(s);
				i=0;
				ptrRs1->MoveFirst();

				while(i<k){
					CString strNum=_T("");
					CString strStg = _T("");
					vC = pR->GetValueRec(ptrRs1,5); // № Склада
					vC.ChangeType(VT_BSTR);
					strStg = vC.bstrVal;
					aHd.SetAt(2,strStg);
//********************************* Работа с файлами
					int curPos = 0;
					CString strBuf = _T("");
					s = m_strDate;
					s.Remove('-');

					strFlNmCmdTsk = strPathOut;
					strFlNmCmdTsk+=	_T("НЗ_");
					strFlNmCmdTsk+= m_strNumDec + _T("_");
					strFlNmCmdTsk+=	s + _T("_");
					strFlNmCmdTsk+= strStg + _T(".rtf");
					if(strANZ.GetAt(0)==_T("")){
						strANZ.SetAt(0,strFlNmCmdTsk);
					}
					else{
						strANZ.Add(strFlNmCmdTsk);
					}
//AfxMessageBox((LPCTSTR)strNZ.GetAt(i));

					try{
//						strFlNmCmdTsk+= s + _T(".rtf");
//AfxMessageBox(strFlNmCmdTsk);
						otCmdTsk.Open(strFlNmCmdTsk,CFile::modeCreate|CFile::modeReadWrite|CFile::typeText);

						s = strPath + _T("CmdTsk.rtf");
						inCmdTsk.Open(s,CFile::modeRead|CFile::shareDenyNone);
//AfxMessageBox(s);
						CString strBuf=_T("");
						CString strRpl=_T("");
						int j=1;
						CString strFnd;
						bool theEnd=true;
//---------------------------Заполняю Шаку
						strFnd.Format("~%i~",j);
						while(theEnd){ 
							inCmdTsk.ReadString(strBuf);
//s=strBuf;
							strFnd.Format("~%i~",j);
							m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
							curPos = inCmdTsk.GetPosition();

//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
							otCmdTsk.WriteString(strBuf);
							if(j==5){
								theEnd = false;
							}
						}
//------------------------------------ Конец Шапки
						//{\cell }

						m_Blk.GetStdPWt(strBuf,inCmdTsk,otCmdTsk, _T("{\\cell }"),curPos);
//						m_Blk.Control(inCmdTsk,300);

						CString strSrc=_T("");

						m_Blk.GetStdPWS(strBuf,inCmdTsk,otCmdTsk, _T("\\pard \\ql"),curPos,strSrc);
						curPos++;
						inCmdTsk.Seek(curPos,CFile::begin);
						strSrc+=_T("\\");
						m_Blk.GetStdPWS(strBuf,inCmdTsk,otCmdTsk, _T("\\pard \\ql"),curPos,strSrc);
//		s.Format("Поиск strSrc После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strSrc = "+strSrc);
//		m_Blk.Control(inCmdTsk,100);

						CString strPst=_T("");

						m_Blk.GetStdPWS(strBuf,inCmdTsk,otCmdTsk, _T("\\trowd"),curPos,strPst);
						curPos++;
						inCmdTsk.Seek(curPos,CFile::begin);
						strPst+=_T("\\");
						m_Blk.GetStdPWS(strBuf,inCmdTsk,otCmdTsk, _T("\\trowd"),curPos,strPst);
//		s.Format("Поиск strPst После  GetStdPWS curPos = %i",curPos);
//AfxMessageBox(s+'\n'+"strPst = "+strPst);
//		m_Blk.Control(inCmdTsk,100);
//		s.Format("i =%i",aRecSt.GetAt(i));
//AfxMessageBox(s);
						int i5=0;
						CString strNew=_T("");
						for(int x=1;x<=aRecSt.GetAt(i);x++){
							strNew = pR->OnFillRow(ptrRs1,strSrc,5,x,i5);
//AfxMessageBox(_T("strSrc = ")+'\n'+strSrc+'\n'+_T("strNew = ")+'\n'+strNew);
							otCmdTsk.WriteString(strNew);
							otCmdTsk.WriteString(strPst);
							ptrRs1->MoveNext();
//							s.Format("x = %i",x);
//AfxMessageBox(s);
						}

						m_Blk.GetStdPWt(strBuf,inCmdTsk,otCmdTsk, _T("\\cell }\\pard"),curPos);
//						m_Blk.Control(inCmdTsk,300);

						CString strSrc2=_T("");
						m_Blk.GetStdPWS(strBuf,inCmdTsk,otCmdTsk, _T("\\pard \\ql"),curPos,strSrc2);
//						s.Format("Поиск strSrc2 После  GetStdPWS curPos = %i",curPos);
//				AfxMessageBox(s+'\n'+"strSrc2 = "+strSrc2);
//						m_Blk.Control(inCmdTsk,100);

						CString strNew2=_T("");
						strNew2 = pR->OnFillTtl(strSrc2,i5,1);
//				AfxMessageBox("strNew2 = "+'\n'+strNew2);
						otCmdTsk.WriteString(strNew2);

						s = m_W.GetWords(i5,false,false,true);
						aHd.SetAt(5,s);  // Общее кол-во

//						while(inCmdTsk.ReadString(strBuf)){ 
//							otCmdTsk.WriteString(strBuf);
//						}

						while(inCmdTsk.ReadString(strBuf)){ 

				//			s= strBuf;
							strFnd.Format("~%i~",j);
							m_Blk.SetRplStrR(strBuf,strFnd,j,aHd);
				//AfxMessageBox("ВХОД = "+s+'\n'+'\n'+"ВЫХОД = "+strBuf+'\n'+'\n'+strFnd);
							otCmdTsk.WriteString(strBuf);
						}
					otCmdTsk.Close();


					}
					catch(CFileException* fx){
						TCHAR buf[255];
						fx->GetErrorMessage(buf,255);
				//		AfxMessageBox(buf);
						fx->Delete();
					}

					i++;
					inCmdTsk.Close();
//s.Format("i = %i",i);
//AfxMessageBox(s);
//				}
			}
/*AfxMessageBox(_T("Begin"));
for(int c=0;c<strNZ.GetSize();c++){
	s.Format(_T("c=%i Size = %i"),c,strNZ.GetSize());
	AfxMessageBox(s);
		AfxMessageBox((LPCTSTR)strNZ.GetAt(c));
	AfxMessageBox(s);
}
AfxMessageBox(_T("End"));*/
		}
	}
	catch(_com_error& e){
		AfxMessageBox(e.ErrorMessage());
		if(ptrRs1->State == adStateOpen) ptrRs1->Close();
		ptrRs1 = NULL;
	}

//c.EndWaitCursor();
//	AfxSetResourceHandle(hInstResClnt);

	if(ptrRs1->State == adStateOpen) ptrRs1->Close();
	ptrRs1 = NULL;
	i = 0;
	while(i<k){
		if(i==0){
			m_strNZ =strANZ.GetAt(i);
		}
		else{
			m_strNZ+= _T("~")+strANZ.GetAt(i);
		}
		i++;
	}
//AfxMessageBox(m_strNZ);
//strNZ = m_strNZ.GetBufferSetLength(m_strNZ.GetLength());
	int lw;
	lw = MultiByteToWideChar(CP_ACP,0,m_strNZ,-1,NULL,0);
//	s.Format(_T("lw = %i"),lw);
//	AfxMessageBox(s);
	WCHAR* strW = new WCHAR[lw];
	MultiByteToWideChar(CP_ACP,0,m_strNZ,-1,strW,lw);
//	strNZ = m_strNZ;
	wcscpy_s(strNZ,lw,strW);
//	strNZ=strW;
//AfxMessageBox("exit from RprCmdTs");
//AfxMessageBox(strNZ);
	delete strW;
	return TRUE;
}

#ifdef _MANAGED
#pragma managed(pop)
#endif

