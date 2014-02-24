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


CString CRecSetProc::GetStrRs(_RecordsetPtr rs, int i, int iRec, double& d17,double& d18,double& d19)
{
	COleVariant vC;
	COleCurrency vCY;
	CString str,s;
	double dT;
	int iT=0;
/*	s.Format(" i = %i",i);
AfxMessageBox(s);
*/
	short j;
	str = _T("");
	switch(i){
	case 1:
		str.Format("%i",iRec);
		break;
	case 2:						// Наименование
		j = 1;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 3:
		str = _T("шт");
		break;
	case 4:						// Кол-во
		j = 3; //2
		vC = GetValueRec(rs,j);	
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 5:		// Цена без НДС
		j = 9;	// 3
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
//		str = vC.bstrVal;
//		vCY.SetCurrency(dT,4);

		str.Format("%16.4f",dT);
//		str = vCY.Format(0, MAKELCID(MAKELANGID(LANG_RUSSIAN,
//									SUBLANG_SYS_DEFAULT), SORT_DEFAULT));
		break;
	case 6:		// Суммы без НДС
		j = 5; //4
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d17 +=dT;
		str.Format("%16.4f",dT);
		break;
	}


	str.TrimRight();
	str.TrimLeft();

	return str;
}



CString CRecSetProc::OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, double& d17,double& d18,double& d19)
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
			strTmp+=GetStrRs(rs,i,iRec, d17, d18, d18)+strFnd;
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
