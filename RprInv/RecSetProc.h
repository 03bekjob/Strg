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
	CString GetStrTtl(int i, double d8, double d9);
	CString OnFillTtl(CString strSrc, double d8, double d9, int Num);
	CString OnFillRow(_RecordsetPtr rs, CString strSrc, int Num, int iRec, double& d8, double& d9);
	CString GetStrRs(_RecordsetPtr rs, int i, int iRec, double& d8, double& d12);
	BOOL IsEmptyRec(_RecordsetPtr rs);
	BOOL IsQuery(CString substr,_RecordsetPtr rs);
	COleVariant GetValueRec(_RecordsetPtr rs, short i);
	CRecSetProc();
	virtual ~CRecSetProc();
protected:


};

#endif // !defined(AFX_RECSETPROC_H__60B0BD56_4FF0_4B04_AFA3_716427BCCD58__INCLUDED_)
