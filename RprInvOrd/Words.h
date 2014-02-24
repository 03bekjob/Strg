// Words.h: interface for the CWords class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_WORDS_H__B677CD7C_8970_4F46_9A31_F191C488ACBC__INCLUDED_)
#define AFX_WORDS_H__B677CD7C_8970_4F46_9A31_F191C488ACBC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CWords  
{
public:
	CString GetWords(double Dgtl, bool bNm=true, bool bMn=true, bool bRpl=true);
	CString str;
	CWords();
	virtual ~CWords();

};

#endif // !defined(AFX_WORDS_H__B677CD7C_8970_4F46_9A31_F191C488ACBC__INCLUDED_)
