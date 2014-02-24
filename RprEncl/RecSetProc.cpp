// RecSetProc.cpp: implementation of the CRecSetProc class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "RecSetProc.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRecSetProc::CRecSetProc()
{

}

CRecSetProc::~CRecSetProc()
{

}

COleVariant CRecSetProc::GetValueRec(_RecordsetPtr rs, short i)
{
//	AfxMessageBox("ok");
	COleVariant vC;
	vC = rs->adoFields->GetItem((COleVariant)i)->Value;
	return vC;
}

BOOL CRecSetProc::IsQuery(CString substr, _RecordsetPtr rs)
{
	COleVariant vC;
	CString strC;

	vC = rs->GetSource();
	vC.ChangeType(VT_BSTR);
	strC=vC.bstrVal;
//AfxMessageBox(strC);
	if(strC.Find(substr)==-1){
		return FALSE;
	}
	else{
		return TRUE;
	}

}

BOOL CRecSetProc::IsEmptyRec(_RecordsetPtr rs)
{
	if(rs->adoEOF && rs->BOF)
		return TRUE;
	else
		return FALSE;

}


CString CRecSetProc::GetStrRs(_RecordsetPtr rs, int i, int iRec)
{
	COleVariant vC;
	COleCurrency vCY;
	CString str,s;
/*	double dT;
	int iT=0;
	s.Format(" i = %i",i);
AfxMessageBox(s);
*/
	str = _T("");
	switch(i){
	case 1:						// Артикул
		vC = GetValueRec(rs,1);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 2:						// Наименование
		vC = GetValueRec(rs,2);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 3:						// № серт.
		vC = GetValueRec(rs,3);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 4:						// Уч. № серт
		vC = GetValueRec(rs,4);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 5:						// Дейст. до даты
		vC = GetValueRec(rs,6);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 6:						// Кем выдан
		vC = GetValueRec(rs,7);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	}


	str.TrimRight();
	str.TrimLeft();

	return str;
}



CString CRecSetProc::OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec)
{
  	CString strFnd,s;
 	CString strTmp = _T("");;
 	strFnd = _T("\\cell ");
	int iLnFnd = strFnd.GetLength();
	bool start=true;
 	int tPos = 0,tPosPrv;
 	for(int i=1;i<=Num;i++){
 		if((tPos=strSrc.Find(strFnd,++tPos))!=-1){
			if(start){
				start = false;
				strTmp+=strSrc.Left(tPos);
			}
			else{
				
 				strTmp+=strSrc.Mid(tPosPrv,tPos-tPosPrv);
			}
			tPosPrv=tPos+iLnFnd;
	/* 		s.Format("tPosPrv = %i, tPos = %i",tPos,tPosPrv);
	AfxMessageBox(strSrc+'\n'+s);
	*/
			strTmp+=GetStrRs(rs,i,iRec)+strFnd;
 		}
		if(i==Num){
/*			tPosPrv=tPos+iLnFnd;
			s.Format("tPosPrv = %i, tPos = %i",tPos,tPosPrv);
	AfxMessageBox(strSrc+'\n'+s);
*/	
			strTmp+=_T("}"); //strSrc.Mid(tPosPrv);
		}
 	}
 
 //	AfxMessageBox(strTmp);
 	return strTmp;

}

CString CRecSetProc::OnFillTtl(CString strSrc, int i5, int Num)
{
 	CString strFnd,s;
 	CString strTmp = _T("");;
 	strFnd = _T("\\cell ");
	int iLnFnd = strFnd.GetLength();
	bool start=true;
 	int tPos = 0,tPosPrv;
 	for(int i=1;i<=Num;i++){
 		if((tPos=strSrc.Find(strFnd,tPos))!=-1){
			if(start){
				start = false;
				strTmp+=strSrc.Left(tPos);
			}
			else{
				
 				strTmp+=strSrc.Mid(tPosPrv,tPos-tPosPrv);
			}
			tPosPrv=tPos+iLnFnd;
//	 		s.Format("tPosPrv = %i, tPos = %i",tPos,tPosPrv);
//	AfxMessageBox(strTmp+'\n'+s);
			tPos++;
			strTmp+=GetStrTtl(i,i5)+strFnd;
 		}
		if(i==Num){
/*			tPosPrv=tPos+iLnFnd;
			s.Format("tPosPrv = %i, tPos = %i",tPos,tPosPrv);
	AfxMessageBox(strSrc+'\n'+s);
*/	
			strTmp+=_T("}"); //strSrc.Mid(tPosPrv);
		}
 	}
 
// 	AfxMessageBox(strSrc+'\n'+strTmp);
 	return strTmp;

}

CString CRecSetProc::GetStrTtl(int i, int i5)
{
	CString str,s;
	str = _T("");
	switch(i){
	case 1:
		str.Format("%i",i5);
		break;
	}

	str.TrimRight();
	str.TrimLeft();
	return str;
}
