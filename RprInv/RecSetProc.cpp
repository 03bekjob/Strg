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
	m_bLRG = FALSE;
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


CString CRecSetProc::GetStrRs(_RecordsetPtr rs, int i, int iRec, double& d8, double& d9)
{
	COleVariant vC;
	CString str,s;
	double dT;
	int iT=0;
	short j;
/*	s.Format(" i = %i",i);
AfxMessageBox(s);
*/
	str = _T("");
	switch(i){
	case 1:					//1: Наименование
		j = 0;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 2:
		str = _T("шт");
		break;
	case 3:					// Кол-во
		j = 3;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_I4);
		iT = vC.lVal;
		str.Format("%i",iT);
		break;
	case 4:					// Цена без НДС,пока стоит ЦЕНА СО СКИДКОЙ
		j  = 9; //14
		vC = GetValueRec(rs,j); // !!! ПЕРЕДЕЛАТЬ
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 5:					// Сумма без НДС
		j = 5;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		str.Format("%16.4f",dT);
		break;
	case 6:					// Акциз
		str = _T("0.00"); 
		break;
	case 7:					// Налоговая ставка
		j = 6;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 8:					// Сумма налога
		j = 7;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d8 +=dT;
		str.Format(_T("%16.4f"),dT);
		break;
	case 9:					// Cуммы c учетом НДС
		j = 4;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d9 +=dT;
		str.Format(_T("%16.4f"),dT);
		break;
	case 10:				// Страна происхождения
		j = 11;	// 16
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 11:				// ГТД
		j = 12;		//17
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	}
//AfxMessageBox(s+'\n'+str);
	str.TrimRight();
	str.TrimLeft();

	return str;
}



CString CRecSetProc::OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, double& d8, double& d9)
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
//	 		s.Format("tPosPrv = %i, tPos = %i",tPos,tPosPrv);
//	AfxMessageBox(strSrc+'\n'+s);
	
			strTmp+=GetStrRs(rs,i,iRec, d8, d9)+strFnd;
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

CString CRecSetProc::OnFillTtl(CString strSrc, double d8, double d9, int Num)
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
			strTmp+=GetStrTtl(i,d8,d9)+strFnd;
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

CString CRecSetProc::GetStrTtl(int i, double d8, double d9)
{
	CString str,s;
	str = _T("");
	switch(i){
	case 1:
		if(m_bLRG){
			str.Format("%16.2f",d8);
			s = _T("\\b");
			s+=str;
			s+=_T("\\b");
			str = s;
		}
		else{
			str.Format("%16.4f",d8);
		}
		break;
	case 2:
		if(m_bLRG){
			str.Format("%16.2f",d9);
			s = _T("\\b");
			s+=str;
			s+=_T("\\b");
			str = s;
		}
		else{
			str.Format("%16.4f",d9);
		}
		break;
	}

	str.TrimRight();
	str.TrimLeft();
	return str;
}
