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


CString CRecSetProc::GetStrRs(_RecordsetPtr rs, int i, int iRec, long& i8,long& i10, double& d12, double& d14, double& d15)
{
	COleVariant vC;
	CString str,s;
	double dT,dC;
	long iT=0;
	short j;
/*	s.Format(" i = %i",i);
AfxMessageBox(s);
*/
	str = _T("");
	switch(i){
	case 1:
		str.Format(_T("%i"),iRec);
		break;
	case 2:			// Наименование
		j = 0;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 3:
		j = 1;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 4:
		str = _T("шт");
		break;
	case 5:
		str = _T("796");
		break;
	case 7:
		str = _T("1");
		break;
	case 8:			// Подсчёт мест шт
		j = 3;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_I4);
		iT = vC.lVal;
		i8 += iT;
		str.Format(_T("%i"),iT);
//AfxMessageBox(str);
		break;
	case 10:		// Подсчёт кол-ва мест шт
		j = 3;
		vC = GetValueRec(rs,3);
		vC.ChangeType(VT_I4);
		iT = vC.lVal;
		i10+=iT;
		str.Format(_T("%i"),iT);
/*
s.Format("i10 = %i",i10);
AfxMessageBox(s);
*/
		break;
	case 11:		// Цена без НДС
		j = 9;	//14
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dC = vC.dblVal;

		str.Format(_T("%16.3f"),dC);
		dC = atof(str);
//		dC = _wtof(str);
		str.Format(_T("%16.2f"),dC);
/*		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
*/
		break;
	case 12:		// Суммы без НДС
		j = 5;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d12 +=dT;
		str.Format(_T("%16.3f"),dT);
		dT = atof(str);
//		dT = _wtof(str);
		str.Format(_T("%16.2f"),dT);
//AfxMessageBox(str);
/*		str.Format("%16.4f",dT);*/
		break;
	case 13:		// % НДС
		j = 6;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_BSTR);
		str = vC.bstrVal;
		break;
	case 14:		// Сумма НДС
		j = 7;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d14 +=dT;

		str.Format(_T("%16.3f"),dT);
		dT = atof(str);
//		dT = _wtof(str);
		str.Format(_T("%16.2f"),dT);

/*		str.Format("%16.4f",dT);*/
		break;
	case 15:		// Cуммы c учетом НДС
		j = 4;
		vC = GetValueRec(rs,j);
		vC.ChangeType(VT_R8);
		dT = vC.dblVal;
		d15 +=dT;
		str.Format(_T("%16.3f"),dT);
		dT = atof(str);
//		dT = _wtof(str);
		str.Format(_T("%16.2f"),dT);
//AfxMessageBox(str);
/*		str.Format("%16.4f",dT);*/
		break;
	}

	str.TrimRight();
	str.TrimLeft();

	return str;
}



CString CRecSetProc::OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, long& i8,long& i10, double& d12, double& d14, double& d15)
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
			
			switch(iRec){
			case 13:
				AfxMessageBox(strTmp);
				break;
			}
*/			strTmp+=GetStrRs(rs,i,iRec, i8, i10, d12, d14, d15)+strFnd;
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

CString CRecSetProc::OnFillTtl(CString strSrc, int i8, int i10, double d12, double d14, double d15, int Num)
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
			strTmp+=GetStrTtl(i,i8,i10,d12,d14,d15)+strFnd;
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

CString CRecSetProc::GetStrTtl(int i, int i8, int i10, double d12, double d14, double d15)
{
	CString str,s;
	str = _T("");
	switch(i){
	case 1:
		str.Format(_T("%i"),i8);
		break;
	case 3:
		str.Format(_T("%i"),i10);
		break;
	case 4:
		str = _T("X");
		break;
	case 5:
		if(m_bLRG){
			str.Format(_T("%16.2f"),d12);
			s = _T("\\b");
			s+=str;
			s+=_T("\\b");
			str = s;
		}
		else{
			  str.Format(_T("%16.2f"),d12); 	
			/*str.Format("%16.4f",d12);*/
		}
		break;
	case 6:
		str = _T("X");
		break;
	case 7:
		if(m_bLRG){
			str.Format(_T("%16.2f"),d14);
			s = _T("\\b");
			s+=str;
			s+=_T("\\b");
			str = s;
		}
		else{		
			str.Format(_T("%16.2f"),d14);
			/*str.Format("%16.4f",d14);*/
		}
		break;
	case 8:
		{
			if(m_bLRG){
				str.Format(_T("%16.2f"),d15);
				s = _T("\\b");
				s+=str;
				s+=_T("\\b");
				str = s;
			}
			else{
				str.Format(_T("%16.2f"),d15);
				/*str.Format("%16.4f",d15);*/
			}
		}
		break;
	}
	str.TrimRight();
	str.TrimLeft();
	return str;
}
