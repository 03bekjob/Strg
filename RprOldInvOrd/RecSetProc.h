// RecSetProc.h: interface for the CRecSetProc class.
//
//////////////////////////////////////////////////////////////////////
//#include "DataCombo.h"


#if !defined(AFX_RECSETPROC_H__60B0BD56_4FF0_4B04_AFA3_716427BCCD58__INCLUDED_)
#define AFX_RECSETPROC_H__60B0BD56_4FF0_4B04_AFA3_716427BCCD58__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CRecSetProc  
{
public:
	CString GetStrTtl(int i, int i5);
	CString OnFillTtl(CString strSrc, int i5, int Num);
	CString OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, int& i8,double& d6);
	CString GetStrRs(_RecordsetPtr rs, int i, int iRec, int& i8,double& d6);
	BOOL IsEmptyRec(_RecordsetPtr rs);
	BOOL IsQuery(CString substr,_RecordsetPtr rs);
	COleVariant GetValueRec(_RecordsetPtr rs, short i);
	CRecSetProc();
	virtual ~CRecSetProc();
protected:


};

#endif // !defined(AFX_RECSETPROC_H__60B0BD56_4FF0_4B04_AFA3_716427BCCD58__INCLUDED_)
