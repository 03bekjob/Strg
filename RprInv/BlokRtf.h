// BlokRtf.h: interface for the CBlokRtf class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BLOKRTF_H__D94DED24_5ECC_4AB6_B3D3_D02125430D8B__INCLUDED_)
#define AFX_BLOKRTF_H__D94DED24_5ECC_4AB6_B3D3_D02125430D8B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBlokRtf  
{
public:
	void GetStdPWS(CString strBuf, CStdioFile &inF, CStdioFile &otF, CString strFnd, int& curPrv,CString& strOut);
	void GetStdPWt(CString strBuf,CStdioFile& inF,CStdioFile& otF,CString strFind, int& curPrv);
	bool SetRplStrR(CString &strBuf, CString& strFnd, int& j, CStringArray& a);
	void SetPos(int iBlk, CString strFnd, CFile &inF, int &curPos);
	int GetLnFd(CString str, int n);
	void GetLStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString& strOut);
	void Control(CFile& f,int iBlk);
	int GetEndInF(int Cnst, int curPos, int End, int d) const;
	void GetRStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString& strOut);
	void GetRStrWrt(int iBlk,CString strFnd,CFile& inF, CFile& oF, int&curPos, CString& strOut);
	void GetLStrWrt(int iBlk,CString strFnd,CFile& inF,CFile& oF,int& curPos,CString& strOut);
	CString GetStrBuf(int& iBlk,CFile& inF, int& curPos,bool& bEnd);
	int GetPosWrt(int iBlk, CString strFnd, CFile& inF, CFile& oF, int& curPos);
	bool SetRplStr(CString& strBuf, CString strFnd, int& j, CStringArray& a);
	CString GetRplStr(CStringArray& a, int i);
	void GetStrWrtStr(int iBlk, CFile& inF, CFile& oF, CString strFnd, int& curPos);
	CBlokRtf();
	virtual ~CBlokRtf();

};

#endif // !defined(AFX_BLOKRTF_H__D94DED24_5ECC_4AB6_B3D3_D02125430D8B__INCLUDED_)
