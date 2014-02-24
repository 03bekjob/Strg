#include "StdAfx.h"
#include "Cnvrt16.h"

CCnvrt16::CCnvrt16(void)
{
	map16.SetAt(_T("À"),_T("\\'c0"));
	map16.SetAt(_T("Á"),_T("\\'c1"));
	map16.SetAt(_T("Â"),_T("\\'c2"));
	map16.SetAt(_T("Ã"),_T("\\'c3"));
	map16.SetAt(_T("Ä"),_T("\\'c4"));
	map16.SetAt(_T("Å"),_T("\\'c5"));
	map16.SetAt(_T("Æ"),_T("\\'c6"));
	map16.SetAt(_T("Ç"),_T("\\'c7"));
	map16.SetAt(_T("È"),_T("\\'c8"));
	map16.SetAt(_T("É"),_T("\\'c9"));
	map16.SetAt(_T("Ê"),_T("\\'ca"));
	map16.SetAt(_T("Ë"),_T("\\'cb"));
	map16.SetAt(_T("Ì"),_T("\\'cc"));
	map16.SetAt(_T("Í"),_T("\\'cd"));
	map16.SetAt(_T("Î"),_T("\\'ce"));
	map16.SetAt(_T("Ï"),_T("\\'cf"));
	map16.SetAt(_T("Ð"),_T("\\'d0"));
	map16.SetAt(_T("Ñ"),_T("\\'d1"));
	map16.SetAt(_T("Ò"),_T("\\'d2"));
	map16.SetAt(_T("Ó"),_T("\\'d3"));
	map16.SetAt(_T("Ô"),_T("\\'d4"));
	map16.SetAt(_T("Õ"),_T("\\'d5"));
	map16.SetAt(_T("Ö"),_T("\\'d6"));
	map16.SetAt(_T("×"),_T("\\'d7"));
	map16.SetAt(_T("Ø"),_T("\\'d8"));
	map16.SetAt(_T("Ù"),_T("\\'d9"));
	map16.SetAt(_T("Ú"),_T("\\'da"));
	map16.SetAt(_T("Û"),_T("\\'db"));
	map16.SetAt(_T("Ü"),_T("\\'dc"));
	map16.SetAt(_T("Ý"),_T("\\'dd"));
	map16.SetAt(_T("Þ"),_T("\\'de"));
	map16.SetAt(_T("ß"),_T("\\'df"));

	map16.SetAt(_T("à"),_T("\\'e0"));
	map16.SetAt(_T("á"),_T("\\'e1"));
	map16.SetAt(_T("â"),_T("\\'e2"));
	map16.SetAt(_T("ã"),_T("\\'e3"));
	map16.SetAt(_T("ä"),_T("\\'e4"));
	map16.SetAt(_T("å"),_T("\\'e5"));
	map16.SetAt(_T("æ"),_T("\\'e6"));
	map16.SetAt(_T("ç"),_T("\\'e7"));
	map16.SetAt(_T("è"),_T("\\'e8"));
	map16.SetAt(_T("é"),_T("\\'e9"));
	map16.SetAt(_T("ê"),_T("\\'ea"));
	map16.SetAt(_T("ë"),_T("\\'eb"));
	map16.SetAt(_T("ì"),_T("\\'ec"));
	map16.SetAt(_T("í"),_T("\\'ed"));
	map16.SetAt(_T("î"),_T("\\'ee"));
	map16.SetAt(_T("ï"),_T("\\'ef"));
	map16.SetAt(_T("ð"),_T("\\'f0"));
	map16.SetAt(_T("ñ"),_T("\\'f1"));
	map16.SetAt(_T("ò"),_T("\\'f2"));
	map16.SetAt(_T("ó"),_T("\\'f3"));
	map16.SetAt(_T("ô"),_T("\\'f4"));
	map16.SetAt(_T("õ"),_T("\\'f5"));
	map16.SetAt(_T("ö"),_T("\\'f6"));
	map16.SetAt(_T("÷"),_T("\\'f7"));
	map16.SetAt(_T("ø"),_T("\\'f8"));
	map16.SetAt(_T("ù"),_T("\\'f9"));
	map16.SetAt(_T("ú"),_T("\\'fa"));
	map16.SetAt(_T("û"),_T("\\'fb"));
	map16.SetAt(_T("ü"),_T("\\'fc"));
	map16.SetAt(_T("ý"),_T("\\'fd"));
	map16.SetAt(_T("þ"),_T("\\'fe"));
	map16.SetAt(_T("ÿ"),_T("\\'ff"));

	map16.SetAt(_T("¨"),_T("\\'a8"));
	map16.SetAt(_T("¸"),_T("\\'b8"));
	map16.SetAt(_T("¹"),_T("\\'b9"));

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
