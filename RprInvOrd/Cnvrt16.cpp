#include "StdAfx.h"
#include "Cnvrt16.h"

CCnvrt16::CCnvrt16(void)
{
	map16.SetAt(_T("�"),_T("\\'c0"));
	map16.SetAt(_T("�"),_T("\\'c1"));
	map16.SetAt(_T("�"),_T("\\'c2"));
	map16.SetAt(_T("�"),_T("\\'c3"));
	map16.SetAt(_T("�"),_T("\\'c4"));
	map16.SetAt(_T("�"),_T("\\'c5"));
	map16.SetAt(_T("�"),_T("\\'c6"));
	map16.SetAt(_T("�"),_T("\\'c7"));
	map16.SetAt(_T("�"),_T("\\'c8"));
	map16.SetAt(_T("�"),_T("\\'c9"));
	map16.SetAt(_T("�"),_T("\\'ca"));
	map16.SetAt(_T("�"),_T("\\'cb"));
	map16.SetAt(_T("�"),_T("\\'cc"));
	map16.SetAt(_T("�"),_T("\\'cd"));
	map16.SetAt(_T("�"),_T("\\'ce"));
	map16.SetAt(_T("�"),_T("\\'cf"));
	map16.SetAt(_T("�"),_T("\\'d0"));
	map16.SetAt(_T("�"),_T("\\'d1"));
	map16.SetAt(_T("�"),_T("\\'d2"));
	map16.SetAt(_T("�"),_T("\\'d3"));
	map16.SetAt(_T("�"),_T("\\'d4"));
	map16.SetAt(_T("�"),_T("\\'d5"));
	map16.SetAt(_T("�"),_T("\\'d6"));
	map16.SetAt(_T("�"),_T("\\'d7"));
	map16.SetAt(_T("�"),_T("\\'d8"));
	map16.SetAt(_T("�"),_T("\\'d9"));
	map16.SetAt(_T("�"),_T("\\'da"));
	map16.SetAt(_T("�"),_T("\\'db"));
	map16.SetAt(_T("�"),_T("\\'dc"));
	map16.SetAt(_T("�"),_T("\\'dd"));
	map16.SetAt(_T("�"),_T("\\'de"));
	map16.SetAt(_T("�"),_T("\\'df"));

	map16.SetAt(_T("�"),_T("\\'e0"));
	map16.SetAt(_T("�"),_T("\\'e1"));
	map16.SetAt(_T("�"),_T("\\'e2"));
	map16.SetAt(_T("�"),_T("\\'e3"));
	map16.SetAt(_T("�"),_T("\\'e4"));
	map16.SetAt(_T("�"),_T("\\'e5"));
	map16.SetAt(_T("�"),_T("\\'e6"));
	map16.SetAt(_T("�"),_T("\\'e7"));
	map16.SetAt(_T("�"),_T("\\'e8"));
	map16.SetAt(_T("�"),_T("\\'e9"));
	map16.SetAt(_T("�"),_T("\\'ea"));
	map16.SetAt(_T("�"),_T("\\'eb"));
	map16.SetAt(_T("�"),_T("\\'ec"));
	map16.SetAt(_T("�"),_T("\\'ed"));
	map16.SetAt(_T("�"),_T("\\'ee"));
	map16.SetAt(_T("�"),_T("\\'ef"));
	map16.SetAt(_T("�"),_T("\\'f0"));
	map16.SetAt(_T("�"),_T("\\'f1"));
	map16.SetAt(_T("�"),_T("\\'f2"));
	map16.SetAt(_T("�"),_T("\\'f3"));
	map16.SetAt(_T("�"),_T("\\'f4"));
	map16.SetAt(_T("�"),_T("\\'f5"));
	map16.SetAt(_T("�"),_T("\\'f6"));
	map16.SetAt(_T("�"),_T("\\'f7"));
	map16.SetAt(_T("�"),_T("\\'f8"));
	map16.SetAt(_T("�"),_T("\\'f9"));
	map16.SetAt(_T("�"),_T("\\'fa"));
	map16.SetAt(_T("�"),_T("\\'fb"));
	map16.SetAt(_T("�"),_T("\\'fc"));
	map16.SetAt(_T("�"),_T("\\'fd"));
	map16.SetAt(_T("�"),_T("\\'fe"));
	map16.SetAt(_T("�"),_T("\\'ff"));

	map16.SetAt(_T("�"),_T("\\'a8"));
	map16.SetAt(_T("�"),_T("\\'b8"));
	map16.SetAt(_T("�"),_T("\\'b9"));

}

CCnvrt16::~CCnvrt16(void)
{
}

void CCnvrt16::OnCnvrt16(const CString& src, CString& dst)
{
	int ln;
	ln = src.GetLength();
	CString s,s16;
	s= s16 = _T("");
	for(int i=0;i<ln;i++){
		s = src.GetAt(i);
		if(map16.Lookup(s,s16)){
			dst+=s16;
		}
		else{
			dst+=s;
		}
	}
}
