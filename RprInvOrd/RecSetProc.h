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
	BOOL m_bLRG;
	CString GetStrTtl(int i, int i8, int i10, double d12, double d14, double d15);
	CString OnFillTtl(CString strSrc, int i8, int i10, double d12, double d14, double d15, int Num);
	CString OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, long& i8,long& i10, double& d12, double& d14, double& d15);
	CString GetStrRs(_RecordsetPtr rs,int i,int iRec, long& i8,long& i10, double& d12, double& d14, double& d15);
	BOOL IsEmptyRec(_RecordsetPtr rs);
	BOOL IsQuery(CString substr,_RecordsetPtr rs);
	COleVariant GetValueRec(_RecordsetPtr rs, short i);
	CRecSetProc();
	virtual ~CRecSetProc();
protected:


};

#endif // !defined(AFX_RECSETPROC_H__60B0BD56_4FF0_4B04_AFA3_716427BCCD58__INCLUDED_)
