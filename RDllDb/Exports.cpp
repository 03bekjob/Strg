#include "stdafx.h"
#include "RDllDb.h"
#include "Exports.h"

// Функция возвр итсину если rs пуст
extern "C" __declspec(dllexport) BOOL IsEmptyRec(_RecordsetPtr rs)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(rs->adoEOF && rs->BOF)
		return TRUE;
	else
		return FALSE;
}

extern "C" _declspec(dllexport) BOOL OnFindInGrid(CString strCod, _RecordsetPtr rs, short i, bool& Flg)
{	
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
//AfxMessageBox(strCod);
	if(!strCod.IsEmpty()){
		if(strCod.GetLength()==1 &&
		   (strCod.Left(1)==_T('*') || 
		    strCod.Left(1)==_T('_') ||
			strCod.Left(1)==_T("'"))
		   ){ 
			   return FALSE;}
		if(!rs->adoEOF && !rs->BOF){
			if(rs->GetRecordCount()>1){
				COleVariant vC,vBk;
				_bstr_t bstrName(L"[");
				_bstr_t bstrSort(L"");
				_bstr_t bstrFind("");

				bstrName+=rs->GetadoFields()->GetItem((COleVariant) i)->GetName();
				bstrName+=L"]";
				bstrSort+=bstrName;
	/*
		AfxMessageBox((char*)bstrName);
		CString s;	
	*/
				int tp;
				tp = rs->GetadoFields()->GetItem((COleVariant) i)->GetType();
	/*	s.Format(L"tp = %i i = %i",tp,i);
		AfxMessageBox(s);
	*/
	
				switch(tp){
				case adSmallInt:	//2
					bstrFind=bstrName + L" = " + strCod;
			//		AfxMessageBox("adSmallInt");
					break;
				case adInteger:		//3
					bstrFind=bstrName + L" = " + strCod;
			//		AfxMessageBox("adInteger");
					break;
				case adDouble:		//5
					strCod.Replace(',','.');
					bstrFind=bstrName + L" = " +  strCod;
			//		AfxMessageBox("adDouble");
					break;
				case adCurrency:	//6
					strCod.Replace(',','.');
					bstrFind=bstrName + L" = " +  strCod;
			//		AfxMessageBox("adCurrency");
					break;
				case adDate:		//7
					bstrFind=bstrName + L" = " +  L"'"+strCod +L"'";
			//		AfxMessageBox("135");
					break;
				case adBSTR:		//8
					strCod.Remove('\'');
					bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
			//		AfxMessageBox("adBSTR");
					break;
				case adBoolean:		//11
					bstrFind=bstrName + L" = " + strCod;
					break;
				case adDecimal:		//14
					bstrFind=bstrName + L" = " +  strCod;
			//		AfxMessageBox("adDecimal");
					break;
				case adTinyInt:		//16
					bstrFind=bstrName + L" = " +  strCod;
					break;
				case adUnsignedTinyInt:	//17
					bstrFind=bstrName + L" = " +  strCod;
					break;
				case adUnsignedSmallInt:	//18
					bstrFind=bstrName + L" = " +  strCod;
					break;
				case adUnsignedInt:	//19
					bstrFind=bstrName + L" = " +  strCod;
					break;
				case adNumeric:		//131
					strCod.Replace(',','.');
					bstrFind=bstrName + L" = " +  strCod;
			//		AfxMessageBox("adNumeric");
					break;
				case adBinary:		//128
					bstrFind=bstrName + L" = " +  strCod;
			//		AfxMessageBox("adBinary");
					break;
				case adChar:		//129
					strCod.Remove('\'');  
					bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
			//		AfxMessageBox(L"adChar");
					break;
				case adDBDate:		//133
					bstrFind=bstrName + L" = " +  L"'"+strCod + L"'";
			//		AfxMessageBox("adDBDate");
					break;
				case adDBTime:		//134
					bstrFind=bstrName + L" = " +  L"'"+strCod + L"'";
			//		AfxMessageBox("adDBTime");
					break;
				case adDBTimeStamp:	//135
					bstrFind=bstrName + L" = " +  L"'" + strCod + L"'";
			//		AfxMessageBox("135");
					break;
				case adVarChar:		//200
					strCod.Remove('\''); 
					bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
			//		AfxMessageBox("adVarChar");
					break;
				default:
					bstrFind=bstrName + L" = '" + strCod + L"'";
					break;

				}
//AfxMessageBox(L"Exp "+ bstrFind);
//AfxMessageBox(bstrSort);
				rs->Sort = bstrSort + _T(" ASC");
//AfxMessageBox("Exp 1");											//vBk
				rs->Find(bstrFind,0,adSearchForward/*,adBookmarkFirst*/);
//AfxMessageBox("Exp _1");
				Flg = true;
				if(rs->adoEOF){
					rs->MoveLast();
//AfxMessageBox("Exp m_bFnd=FALSE");
					return TRUE;
				}
				else{
//AfxMessageBox("Exp m_bFnd=TRUE");
					Flg = false;
					try{
						vBk = rs->GetBookmark();
					}
					catch(...){
						rs->MoveNext();
					}
					Flg = true;
	/*				Flg = false;
					rs->MoveNext();
					Flg = true;
					rs->put_Bookmark(vBk);
	*/
					return TRUE;
				}
			}
			else{
//AfxMessageBox("Exp m_bFnd=TRUE т.к. 1 запись");
				Flg = true;
				return TRUE;
			}
		}
		else{
//AfxMessageBox("m_bFnd=FALSE т.к rs->adoEOF && rs->BOF");
			Flg = true;
			return FALSE;
		}
		
	}
//AfxMessageBox("m_bFnd=FALSE т.к strCod = Empty");
	return FALSE;
}

extern "C" _declspec(dllexport) BOOL OnFindInGridP(CString strCod, _RecordsetPtr& rs, short i, bool& Flg)
{	
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(!strCod.IsEmpty()){
		if(strCod.GetLength()==1 &&
		   (strCod.Left(1)==_T('*') || 
		    strCod.Left(1)==_T('_') ||
			strCod.Left(1)==_T("'"))
		   ){ return FALSE;}
		if(!rs->adoEOF && !rs->BOF){
			if(rs->GetRecordCount()>1){
				COleVariant vC,vBk;
				_bstr_t bstrName(L"[");
				_bstr_t bstrSort(L"");
				_bstr_t bstrFind("");

				bstrName+=rs->GetadoFields()->GetItem((COleVariant) i)->GetName();
				bstrName+=L"]";
				bstrSort+=bstrName;
	/*
		AfxMessageBox((char*)bstrName);
		CString s;	
	*/
				int tp;
				tp = rs->GetadoFields()->GetItem((COleVariant) i)->GetType();
/*				s.Format("tp = %i i = %i",tp,i);
				AfxMessageBox(s);
*/	
				switch(tp){
				case adSmallInt:	//2
					bstrFind=bstrName + " = " + strCod;
			//		AfxMessageBox("adSmallInt");
					break;
				case adInteger:		//3
					bstrFind=bstrName + " = " + strCod;
			//		AfxMessageBox("adInteger");
					break;
				case adDouble:		//5
					bstrFind=bstrName + " = " +  strCod;
			//		AfxMessageBox("adDouble");
					break;
				case adCurrency:	//6
					strCod.Replace(',','.');
					bstrFind=bstrName + " = " +  strCod;
			//		AfxMessageBox("adCurrency");
					break;
				case adDate:		//7
					bstrFind=bstrName + " = " +  "'"+strCod +"'";
			//		AfxMessageBox("135");
					break;
				case adBSTR:		//8
					strCod.Remove('\'');
					bstrFind=bstrName + " LIKE " + "'" + strCod + "%'";
			//		AfxMessageBox("adBSTR");
					break;
				case adBoolean:		//11
					bstrFind=bstrName + " = " + strCod;
					break;
				case adDecimal:		//14
					bstrFind=bstrName + " = " +  strCod;
			//		AfxMessageBox("adDecimal");
					break;
				case adTinyInt:		//16
					bstrFind=bstrName + " = " +  strCod;
					break;
				case adUnsignedTinyInt:	//17
					bstrFind=bstrName + " = " +  strCod;
					break;
				case adUnsignedSmallInt:	//18
					bstrFind=bstrName + " = " +  strCod;
					break;
				case adUnsignedInt:	//19
					bstrFind=bstrName + " = " +  strCod;
					break;
				case adNumeric:		//131
					bstrFind=bstrName + " = " +  strCod;
			//		AfxMessageBox("adNumeric");
					break;
				case adBinary:		//128
					bstrFind=bstrName + " = " +  strCod;
			//		AfxMessageBox("adBinary");
					break;
				case adChar:		//129
					strCod.Remove('\'');
					bstrFind=bstrName + " LIKE " + "'" + strCod + "%'";
			//		AfxMessageBox("adChar");
					break;
				case adDBDate:		//133
					bstrFind=bstrName + " = " +  "'"+strCod +"'";
			//		AfxMessageBox("adDBDate");
					break;
				case adDBTime:		//134
					bstrFind=bstrName + " = " +  "'"+strCod +"'";
			//		AfxMessageBox("adDBTime");
					break;
				case adDBTimeStamp:	//135
					bstrFind=bstrName + " = " +  "'"+strCod +"'";
			//		AfxMessageBox("135");
					break;
				case adVarChar:		//200
					strCod.Remove('\'');
					bstrFind=bstrName + " LIKE " + "'" + strCod + "%'";
			//		AfxMessageBox("adVarChar");
					break;

				}
		//		bstrFind=bstrName + " = '" + strCod + "'";
//AfxMessageBox("Exp "+bstrFind);
//AfxMessageBox(bstrSort);
				rs->Sort = bstrSort + _T(" ASC");
//AfxMessageBox("Exp 1");											//vBk
				rs->Find(bstrFind,0,adSearchForward/*,adBookmarkFirst*/);
//AfxMessageBox("Exp _1");
				Flg = true;
				if(rs->adoEOF){
					rs->MoveLast();
//AfxMessageBox("Exp m_bFnd=FALSE");
					return TRUE;
				}
				else{
//AfxMessageBox("Exp m_bFnd=TRUE");
					Flg = false;
					try{
						vBk = rs->GetBookmark();
					}
					catch(...){
						rs->MoveNext();
					}
					return TRUE;
				}
			}
			else{
//AfxMessageBox("Exp m_bFnd=TRUE т.к. 1 запись");
				Flg = true;
				return TRUE;
			}
		}
		else{
//AfxMessageBox("m_bFnd=FALSE т.к rs->adoEOF && rs->BOF");
			Flg = true;
			return FALSE;
		}
		
	}
//AfxMessageBox("m_bFnd=FALSE т.к strCod = Empty");
	return FALSE;
}


extern "C" __declspec(dllexport) /*COleVariant*/ VARIANT GetValueRec(_RecordsetPtr rs, short i)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	VARIANT vC;
	vC = rs->GetadoFields()->GetItem((COleVariant)i)->Value;
	
	return vC;

}
extern "C" __declspec(dllexport) BSTR GetNameCol(_RecordsetPtr rs, short i)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	_bstr_t bstrName;
	bstrName=rs->GetadoFields()->GetItem((COleVariant) i)->GetName();
	
	return (BSTR)bstrName;

}

extern "C" __declspec(dllexport) int GetTypeCol(_RecordsetPtr rs,short i)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	int iType;
	iType = rs->GetadoFields()->GetItem((COleVariant) i)->GetType();
	return iType;
}

extern "C" __declspec(dllexport) BSTR GetBSTRFind(int Tp, _bstr_t strName, CString s)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	_bstr_t bstrFind;
	switch(Tp){
	case adSmallInt: //2
	case adInteger:	 //3
		bstrFind = strName;
		bstrFind+= _T(" = ");
		bstrFind+= (_bstr_t)s;
//		AfxMessageBox("adInteger");
		break;
	case adNumeric:
	case adDouble:
	case adCurrency:
	case adBoolean:
		s.Replace(',','.');
		bstrFind = strName;
		bstrFind+= _T(" = ");
		bstrFind+= (_bstr_t)s;
//		AfxMessageBox("adCurrency");
		break;
	case adBinary:
		bstrFind=  strName;
		bstrFind+= _T(" = ");
		bstrFind+= (_bstr_t)s;
//		AfxMessageBox("adBinary");
		break;
	case adDBTimeStamp:
		bstrFind=  strName;
		bstrFind+= _T(" = ");
		bstrFind+= _T("'");
		bstrFind+= (_bstr_t)s;
		bstrFind+= _T("'");
//		AfxMessageBox("135");
		break;
	case adBSTR:
		bstrFind = strName;
		bstrFind+= _T(" like ");
		bstrFind+= _T("'");
		bstrFind+= (_bstr_t)s;
		bstrFind+= _T("%'");
//		AfxMessageBox("adBSTR");
		break;
	case adChar:
		bstrFind = strName;
		bstrFind+= _T(" like ");
		bstrFind+= _T("'");
		bstrFind+= (_bstr_t)s;
		bstrFind+= _T("%'");
//		AfxMessageBox("adChar");
		break;
	}
//AfxMessageBox(bstrFind);
	return (BSTR)bstrFind;

}

extern "C" __declspec(dllexport) long OnChangeCombo(CDatacombo2 &m_DataCombo, _RecordsetPtr rsCombo, short iC)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	long cd = 0;
//	VARIANT vC;
	COleVariant vC;
	CString strFind=m_DataCombo.get_BoundText();
	CString strCur;
//AfxMessageBox(strFind);
	strFind.TrimRight(' ');
	rsCombo->MoveFirst();
	while(!rsCombo->adoEOF){
		vC=rsCombo->GetadoFields()->GetItem((COleVariant) iC)->GetValue();
		vC.ChangeType(VT_BSTR);
		strCur=vC.bstrVal;
		strCur.TrimRight(' ');
		if(strFind==strCur) {
			iC=0;
			vC=rsCombo->GetadoFields()->GetItem((COleVariant) iC)->GetValue();
			vC.ChangeType(VT_I4);
			cd=vC.intVal;
			break;
		}
		rsCombo->MoveNext();
	}
//			AfxMessageBox(strCur+'='+strFind);
	if(rsCombo->adoEOF) rsCombo->MoveLast();

	return cd;
}
extern "C" __declspec(dllexport) void OnFindInCombo(CString strCod, CDatacombo2 *pCombo, _RecordsetPtr rsCombo, short i, short iBound)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(rsCombo->State==adStateOpen){
		if(!IsEmptyRec(rsCombo)){
			if(strCod.IsEmpty()){
				//Если искать ничего выходим и показываем пусто
				pCombo->put_BoundText(_T(" "));

				return;
			}
			COleVariant vC,vBk;
			_bstr_t bstrBound;
			_bstr_t bstrName(L"[");
			_bstr_t bstrSort(L"");
			_bstr_t bstrFind("");

/*			CString s;
			vC = rsCombo->GetSource();
			vC.ChangeType(VT_BSTR);
			s = vC.bstrVal;
AfxMessageBox(s);
*/
			rsCombo->MoveFirst();
			vBk = rsCombo->GetBookmark();

			bstrName+=rsCombo->GetadoFields()->GetItem((COleVariant) i)->GetName();
			bstrName+=L"]";
			bstrSort+=bstrName;

//AfxMessageBox((char*)bstrName);
			int tp;
			tp = rsCombo->GetadoFields()->GetItem((COleVariant) i)->GetType();
/*			s.Format(L"%i",tp);
AfxMessageBox(L"tp = "+s);
*/
			switch(tp){
			case adSmallInt:	//2
				bstrFind=bstrName + L" = " + strCod;
		//		AfxMessageBox("adSmallInt");
				break;
			case adInteger:		//3
				bstrFind=bstrName + L" = " + strCod;
		//		AfxMessageBox("adInteger");
				break;
			case adDouble:		//5
				bstrFind=bstrName + L" = " +  strCod;
		//		AfxMessageBox("adDouble");
				break;
			case adCurrency:	//6
				strCod.Replace(',','.');
				bstrFind=bstrName + L" = " +  strCod;
		//		AfxMessageBox("adCurrency");
				break;
			case adDate:		//7
				bstrFind=bstrName + L" = " +  L"'" + strCod + L"'";
		//		AfxMessageBox("135");
				break;
			case adBSTR:		//8
				bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
		//		AfxMessageBox("adBSTR");
				break;
			case adBoolean:		//11
				bstrFind=bstrName + L" = " + strCod;
				break;
			case adDecimal:		//14
				bstrFind=bstrName + L" = " +  strCod;
		//		AfxMessageBox("adDecimal");
				break;
			case adTinyInt:		//16
				bstrFind=bstrName + L" = " +  strCod;
				break;
			case adUnsignedTinyInt:	//17
				bstrFind=bstrName + L" = " +  strCod;
				break;
			case adUnsignedSmallInt:	//18
				bstrFind=bstrName + L" = " +  strCod;
				break;
			case adUnsignedInt:	//19
				bstrFind=bstrName + L" = " +  strCod;
				break;
			case adNumeric:		//131
				bstrFind=bstrName + L" = " +  strCod;
		//		AfxMessageBox("adNumeric");
				break;
			case adBinary:		//128
				bstrFind=bstrName + L" = " +  strCod;
		//		AfxMessageBox("adBinary");
				break;
			case adChar:		//129
				bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
		//		AfxMessageBox("adChar");
				break;
			case adDBDate:		//133
				bstrFind=bstrName + L" = " +  L"'" + strCod + L"'";
		//		AfxMessageBox("adDBDate");
				break;
			case adDBTime:		//134
				bstrFind=bstrName + L" = " +  L"'" + strCod + L"'";
		//		AfxMessageBox("adDBTime");
				break;
			case adDBTimeStamp:	//135
				bstrFind=bstrName + L" = " +  L"'" + strCod + L"'";
		//		AfxMessageBox("135");
				break;
			case adVarChar:		//200
				bstrFind=bstrName + L" LIKE " + L"'" + strCod + L"%'";
		//		AfxMessageBox("adVarChar");
				break;
			default:
				bstrFind=bstrName + L" = '" + strCod + L"'";
				break;

			}
//AfxMessageBox(bstrFind);
			rsCombo->Find(bstrFind,0,adSearchForward,vBk);
			if(rsCombo->adoEOF){
				rsCombo->MoveLast();
//Раз не нашли показываем пусто
				pCombo->put_BoundText(_T(" "));

			}
			else{
				vC=GetValueRec(rsCombo,iBound);
				vC.ChangeType(VT_BSTR);
				bstrBound = vC.bstrVal;
				pCombo->put_BoundText(bstrBound);
			}
			
		}
	}
}
extern "C" __declspec(dllexport) BOOL IsEnableRec(_RecordsetPtr rs)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(rs->State==adStateOpen){
		if(!IsEmptyRec(rs)){
			return TRUE;
		}
		else{
			return FALSE;
		}
	}
	else
		return FALSE;
}

