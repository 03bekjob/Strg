// BlokRtf.cpp: implementation of the CBlokRtf class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BlokRtf.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBlokRtf::CBlokRtf()
{

}

CBlokRtf::~CBlokRtf()
{

}


CString CBlokRtf::GetRplStr(CStringArray &a, int i)
{
	return a.GetAt(i);
}

bool CBlokRtf::SetRplStr(CString &strBuf, CString strFnd, int& j, CStringArray& a)
{
	int iB,iE,iCnt;
	iB=iE=-1;

	iB = strBuf.Find((LPCTSTR)strFnd);
	iE = strBuf.Find((LPCTSTR)strFnd,iB+1);

	CString strSrc=_T("");
	CString strRpl=_T("");
	if(iB!=-1 && iE!=-1){

		iCnt = iE-iB+1;
		if(j>=10){
			iCnt+=3;
		}
		else if(j>=100){
			iCnt+=4;
		}
		else{
			iCnt+=2;
		}

		int tCnt,rCnt;
		tCnt=rCnt=0;

		strSrc=strBuf.Mid(iB,iCnt);
		strRpl = GetRplStr(a,j);

//AfxMessageBox(strSrc+'\n'+strRpl);

		tCnt = strSrc.GetLength(); 
		rCnt = strRpl.GetLength(); 
		if(rCnt==0){
			strRpl = _T("");
		}
		if(rCnt>tCnt){
			strRpl = strRpl.Left(tCnt);
		}
		if(j<=20){
			strBuf.Replace(strSrc,strRpl);
		}
/*		if(j==20){
			AfxMessageBox(strBuf);
		}
*/
		j++;

		return true;
	}
	else{
		return false;
	}
}

void  CBlokRtf::GetStrWrtStr(int iBlk, CFile &inF, CFile& oF, CString strFnd, int& curPos)
{
// ��������� ������ �� ���. ���. �� ������� ������
// ������� ���� ������ ������
// ��������� ��������� �� ��������� �������� ������� ������
// ����� !!!

	CString strBuf=_T(""),s;

//	s.Format("�� ������ curPos = %i",curPos);
//	AfxMessageBox(s);
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//	s.Format("����� ������ curPos = %i",curPos);
//	AfxMessageBox(s);

	int tPos,iCnt;
	CString strTg=_T("");

	if((tPos=strBuf.Find(strFnd))!=-1){ // �����
		strTg = strBuf.Left(tPos+strFnd.GetLength());
//AfxMessageBox("����� � GetStrWrtStr \n"+strTg);
		iCnt = strTg.GetLength();
		oF.Write(strTg,iCnt);			// �������
		curPos+= iCnt;
		s.Format(" ����� GetStrWrtStr curPos = %i",curPos);
AfxMessageBox(s+'\n'+strTg);
		inF.Seek(curPos,CFile::begin); // �������� ��������� ��
									   // ������� ������
		return; // curPos;
	}
	else{
		oF.Write(strBuf,iBlk);					// ������� 
//		AfxMessageBox("�� ����� GetStrWrtStr\n"+strBuf);
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);			// ����������
//		inF.Seek(iBlk,CFile::current);			// ����������
		curPos = inF.GetPosition();
		GetStrWrtStr(iBlk,inF,oF,strFnd,curPos);
		return;
	}

}

int CBlokRtf::GetPosWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int& curPos)
{
//��������� ������������� �� ������ ������
//����� � ���� ��� ������ ������
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	int tPos;//=0;				// ������� � strBuf
	int iCnt = 0;
	bool bEnd=true;
		strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//-------------------------------��� ���������  ���������
		if((tPos = strBuf.Find(strFnd))!=-1){
//		AfxMessageBox("�����\n"+strBuf);
			strWrt = strBuf.Left(tPos);   //������ ��� ������ � ���. ����
			int iDlt=0;
			iCnt = strWrt.GetLength();
			iDlt = GetLnFd(strWrt,iCnt);	//������� ����.�������� �����
//			s.Format("tPos = %i ",tPos);
			
//		AfxMessageBox("GetPosWrt "+strFnd+'\n'+strWrt +'\n'+ s + '\n'+strBuf);
//		AfxMessageBox(strWrt);
			oF.Write(strWrt,iCnt);//��������� ��� �� bStr
			curPos+= (iCnt+iDlt);// ����� ������ �� bStr
			//��������� ��������� �� ����� ������� �� ���.����� !!!
			inF.Seek(curPos,CFile::begin);
//			curPos = inF.GetPosition();
//			s.Format(" ����� ����� Seek curPos = %i",curPos);
//		AfxMessageBox(s);
			return curPos;
		}
		else{						// �� �����, ����� ������ ������
//AfxMessageBox("�� ����� \n"+strBuf);
			oF.Write(strBuf,iBlk);			// �������
			curPos+=iBlk;
			inF.Seek(curPos,CFile::begin);	// ����������
//			inF.Seek(iBlk,CFile::current);	// ����������
//			curPos = inF.GetPosition(); 
//		s.Format("�� ����� GetPosWrt curPos = %i",curPos);
//		AfxMessageBox(s+'\n'+strBuf);
			GetPosWrt(iBlk,strFnd,inF,oF,curPos);
			//return;
			return curPos;
		}
}

CString CBlokRtf::GetStrBuf(int& iBlk, CFile &inF,int& curPos, bool& bEnd)
{
// ������ � �����
// ��������� � ������
// /*���������� ��������� � ����� �� �������� �����  iBlk*/
	int bPos,ePos;
	int iEnd=0;
//	int iCnt;
	CString strBuf,s;
	bPos=ePos=0;
	char* buf = new char[iBlk];

	bPos = curPos; // ��������� ��� �����

	iEnd = inF.Read(buf,iBlk); // ������� ������� �� ����� ����
		if(iEnd==0){
	//		s.Format("iEnd = %i",iEnd);
	//AfxMessageBox(s+"  iEnd==0");
			inF.Seek(bPos,CFile::begin);
			bEnd = false;
			return "";
		}
		if(iEnd<iBlk){	//******************* ������ ����� �����				   
			delete [] buf;
	/*		for(int i=0;i<iBlk;i++){
				*(buf+i)=_T('');
			}
	*/
	//		strBuf = _T("");

			iBlk = iEnd;				   // ������� ������ �����	
			char* bufE = new char[iBlk];
			inF.Seek(bPos,CFile::begin);  // ��������� �� ������ �����

			curPos = inF.GetPosition();
//			s.Format("����� �������� �� ���� ����� curPos = %i",curPos);
//AfxMessageBox(s);

			inF.Read(bufE,iBlk);		   // �������� � ����� iBlk
	//		bPos = curPos+iBlk;
//			s.Format("iBlk = %i curPos = %i, bPos = %i",iBlk,curPos,bPos);
//	AfxMessageBox(s+"  iEnd<iBlk");
			for(int i=0;i<iBlk;i++){
				strBuf+=*(bufE+i);
			}
//			s.Format("iBlk = %i curPos = %i, bPos = %i",iBlk,curPos,bPos);
//	AfxMessageBox(s+"  iEnd<iBlk"+'\n'+"strBuf = "+strBuf);
			inF.Seek(curPos,CFile::begin);	// �������� �� ������ �����
			curPos = inF.GetPosition();
			delete [] bufE;
			return strBuf;
		}
//	iBlk = iEnd; // �������� ���� ����� ������� �������
	for(int i=0;i<iBlk;i++){
		strBuf+=*(buf+i);
	}
	inF.Seek(bPos,CFile::begin);	// �������� �� ������ �����
//	curPos = inF.GetPosition();
/*	s.Format("curPos = %i, bPos = %i",curPos,bPos);
AfxMessageBox(s);
*/
	delete [] buf;
	return strBuf;
}

void CBlokRtf::GetLStrWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int &curPos, CString &strOut)
{
//��������� ������������� �� ������ ������
//����� � ���� ��� ������ ������
//���������� ��������� ���� ����� � ���� ������ strOut
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	int iCnt = 0;
	int tPos;//=0;				// ������� � strBuf
	bool bEnd=true;
		strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
									//����� �� ��� ����� ����� �����
//-------------------------------��� ���������  ���������
		if((tPos = strBuf.Find(strFnd))!=-1){
			strWrt = strBuf.Left(tPos);   //������ ��� ������ � ���. ����
			iCnt = strWrt.GetLength(); 
			int iDlt=0;
			iDlt = GetLnFd(strWrt,iCnt);	//������� ����.�������� �����
			strOut+= strWrt;
//			s.Format("tPos = %i ",tPos);
			
//		AfxMessageBox("����� GetLStrWrt\n"+strFnd+'\n'+strWrt +'\n'+ s + '\n'+strBuf);
//		AfxMessageBox(strWrt);

			oF.Write(strWrt, iCnt);
			curPos+= (iCnt+iDlt);			
//			s.Format(" ����� ����� curPos+=iCnt = %i",curPos);
//		AfxMessageBox(s);
			//��������� ��������� �� ����� ������� �� ���.����� !!!
			inF.Seek(curPos,CFile::begin);
//			curPos = inF.GetPosition();
//			s.Format(" ����� ����� Seek curPos = %i",curPos);
//		AfxMessageBox(s);
			return;
		}
		else{						// �� �����, ����� ������ ������
//AfxMessageBox("�� ����� GetLStrWrt\n"+strBuf);
			oF.Write(strBuf,iBlk);
			strOut+=strBuf;
			curPos+=iBlk;
			inF.Seek(curPos,CFile::begin);	// ����������
//			inF.Seek(iBlk,CFile::current);
//			curPos = inF.GetPosition(); // ���. ����� inF.Read
//			s.Format(" ����� �� ����� curPos = %i",curPos);
//		AfxMessageBox(s);
			GetLStrWrt(iBlk,strFnd,inF,oF,curPos,strOut);
			return;
		}

}

void CBlokRtf::GetRStrWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int &curPos, CString &strOut)
{
//��������� ������������� �� ������� ������
//����� � ���� �� ������� ������
//���������� ��������� ���� ����� � ���� ������ strOut
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	strTg = _T("");
	int tPos;//=0;				// ������� � strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("��� ������\n"+strBuf);
//AfxMessageBox("���������\n"+strFnd);
//-------------------------------��� ���������  ���������
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos+strFnd.GetLength());
		iCnt = strTg.GetLength();
		strOut+=strTg;
		oF.Write(strTg,iCnt);			// �������
//		s.Format(" ����� GetRStr curPos = %i",curPos);
//AfxMessageBox("����� � GetRStrWrtStr \n"+s+'\n'+strTg);
		curPos+= iCnt;
		inF.Seek(curPos,CFile::begin); // �������� ��������� ��
//			s.Format(" ����� ����� Seek curPos = %i",curPos);
//		AfxMessageBox(s);
		return;
	}
	else{						// �� �����, ����� ������ ������
		AfxMessageBox("�� ����� GetRStrWrt\n"+strBuf);
//		AfxMessageBox("����������\n"+strOut);
		oF.Write(strBuf,iBlk);			// �������
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// ����������
//		inF.Seek(iBlk,CFile::current);	// ����������
		curPos = inF.GetPosition();
		GetRStrWrt(iBlk,strFnd,inF,oF,curPos,strOut);
		return;
	}

}



void CBlokRtf::GetRStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString& strOut)
{
//��������� ������ �� ���������
//��������� ��������������� �� �������
//������ �� �����

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// ������� � strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("��� ������\n"+strBuf+'\n'+"���������\n"+strFnd);
//-------------------------------��� ���������  ���������
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos+strFnd.GetLength());
//AfxMessageBox("����� � GetR \n"+strTg);
		iCnt = strTg.GetLength();
		strOut+=strTg;
		s.Format(" ����� GetRStr curPos = %i",curPos);
AfxMessageBox("����� � GetRStr \n"+s+'\n'+"strTg = "+strTg);
		curPos+= iCnt;
		inF.Seek(curPos,CFile::begin); // �������� ��������� ��
//			s.Format(" ����� ����� Seek curPos = %i",curPos);
//		AfxMessageBox(s);
		return;
	}
	else{						// �� �����, ����� ������ ������
//		AfxMessageBox("�� ����� GetRStr\n"+strBuf);
//		AfxMessageBox("����������\n"+strOut);
		strOut+=strBuf;
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// ����������
//		inF.Seek(iBlk,CFile::current);	// ����������
		curPos = inF.GetPosition();
		GetRStr(iBlk,strFnd,inF,curPos,strOut);
		return;
	}

}

int CBlokRtf::GetEndInF(int Cnst, int curPos, int End, int d) const
{
	int iEnd =0;
	int nEnd =0;
	int dlt  =0;
	div_t div_r;
	iEnd = curPos+Cnst;
	dlt  = End-iEnd;
	if(dlt>0){
		return Cnst;
	}
	else{
		div_r = div((End-curPos),d);
		return div_r.quot;
	}
}

void CBlokRtf::Control(CFile &f, int iBlk)
{
	CString m,n,lnS;
	int iR=0,iLn=0;
	int iDel=0;
	int curPos = f.GetPosition();
	m.Format("��� ��� curPos = %i",curPos);
AfxMessageBox(m);

	char* b = new char[iBlk];

	iR = f.Read(b,iBlk);	// ��������� �������

	curPos = f.GetPosition();
	m.Format("��� ����� Read curPos = %i ��������� iR = %i",curPos,iR);
AfxMessageBox(m);
	CString s=_T("");
	for(int i=0;i<iBlk;i++){
		if(_T(' ')==*(b+i) ){
			s+=_T("_");
		}
		else if( _T('\n')==*(b+i)){
			s+=_T("~");
			iDel++;
		}
		else if(_T('\r')==*(b+i)){
			s+=_T("#");
			iDel++;
		}
		else{
			s+=*(b+i);
		}

//		s+=*(b+i);
	}
//	iLn = s.GetLength();
	lnS.Format(" iDel= %i",iDel);
AfxMessageBox(s+'\n'+lnS);
/*	int ln;
	ln=s.GetLength();
	n.Format("����� ������ s = %i k = %i",ln);
AfxMessageBox(s+'\n'+n);
*/
	curPos-=(iBlk+iDel);
//	curPos-=iBlk;

	m.Format("��� ��� �������� curPos = %i",curPos);
AfxMessageBox(m);

	f.Seek(curPos,CFile::begin);

	curPos = f.GetPosition();
	m.Format("��� ��� ����� Seek ����� �� ��� curPos = %i",curPos);
AfxMessageBox(m);
	delete [] b;

}

void CBlokRtf::GetLStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString &strOut)
{
//��������� ������ �� ���������
//��������� ��������������� ����� ����������
// ����� ������ strOut
//������ �� �����

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// ������� � strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//		AfxMessageBox("strBuf = \n"+strBuf);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("��� ������\n"+strBuf+'\n'+"���������\n"+strFnd);
//-------------------------------��� ���������  ���������
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos);
		iCnt = strTg.GetLength(); //tPos;
		int iDlt=0;
		iDlt = GetLnFd(strTg,iCnt);	//������� ����.�������� �����
		strOut+=strTg;
//		strTg.Replace('\n','~');
//		s.Format(" ����� GetLStr ���� curPos = %i iCnt = %i, iDlt = %i",curPos,iCnt,iDlt);
//AfxMessageBox("����� � GetLStr \n"+s+'\n'+"strTg = "+strTg+'\n'+"strFnd = "+strFnd);
		curPos+= (iCnt+iDlt);				// ������� ����� curPos	
		inF.Seek(curPos,CFile::begin); // �������� ��������� ��
//		s.Format("����� ��� ����� ������ curPos = %i",curPos);
//AfxMessageBox(s);
//		curPos = inF.GetPosition();
		return;
	}
	else{						// �� �����, ����� ������ ������
//		AfxMessageBox("����������\n"+strOut);
		strOut+=strBuf;
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// ����������
//		s.Format("����� ������ curPos = %i",curPos);
//		AfxMessageBox("�� ����� GetLStr\n"+strBuf+'\n'+s);
//		curPos = inF.GetPosition();
		GetLStr(iBlk,strFnd,inF,curPos,strOut);
		return;
	}

}

int CBlokRtf::GetLnFd(CString str, int iCnt)
{
	int iDlt=0;
	for(int n=0;n<iCnt;n++){
		if('\n'==str.GetAt(n)){
			iDlt++;
		}
	}
	return iDlt;
}

void CBlokRtf::SetPos(int iBlk, CString strFnd, CFile &inF, int &curPos)
{
//��������� ������ �� ���������
//��������� ��������������� ����� ����������
//������ �� �����

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// ������� � strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//		AfxMessageBox("strBuf = \n"+strBuf);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("��� ������\n"+strBuf+'\n'+"���������\n"+strFnd);
//-------------------------------��� ���������  ���������
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos);
		iCnt = strTg.GetLength(); //tPos;
		int iDlt=0;
		iDlt = GetLnFd(strTg,iCnt);	//������� ����.�������� �����
//		strTg.Replace('\n','~');
//		s.Format(" ����� GetLStr ���� curPos = %i iCnt = %i, iDlt = %i",curPos,iCnt,iDlt);
//AfxMessageBox("����� � GetLStr \n"+s+'\n'+"strTg = "+strTg+'\n'+"strFnd = "+strFnd);
		curPos+= (iCnt+iDlt);				// ������� ����� curPos	
		inF.Seek(curPos,CFile::begin); // �������� ��������� ��
//		s.Format("����� ��� ����� ������ curPos = %i",curPos);
//AfxMessageBox(s);
//		curPos = inF.GetPosition();
		return;
	}
	else{						// �� �����, ����� ������ ������
//		AfxMessageBox("����������\n"+strOut);
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// ����������
//		s.Format("����� ������ curPos = %i",curPos);
//		AfxMessageBox("�� ����� SetPos\n"+strBuf+'\n'+s);
//		curPos = inF.GetPosition();
		SetPos(iBlk,strFnd,inF,curPos);
		return;
	}

}

bool CBlokRtf::SetRplStrR(CString &strBuf, CString& strFnd, int &j, CStringArray &a)
{
	int iB,iE,iCnt;
	iB=iE=-1;
	CString strSrc,strRpl;
	strSrc=strRpl=_T("");

	iB = strBuf.Find((LPCTSTR)strFnd);		// ��� ���������

	if(iB!=-1){
		iE = strBuf.Find((LPCTSTR)strFnd,iB+1);	// ��� ���������
		if(iE!=-1){
			iCnt = iE-iB+1;
			if(j>=10){
				iCnt+=3;
			}
			else if(j>=100){
				iCnt+=4;
			}
			else{
				iCnt+=2;
			}

			int tCnt,rCnt;
			tCnt=rCnt=0;

			strSrc=strBuf.Mid(iB,iCnt); // ����� ������
			strRpl = GetRplStr(a,j);	// ������ ��� ������

			tCnt = strSrc.GetLength();  // ����� ���
			rCnt = strRpl.GetLength();  // ����� ������
			if(rCnt==0){
				strRpl = _T("");
			}
			if(rCnt>tCnt){
				strRpl = strRpl.Left(tCnt); // ����� ������ ������
			}

			strBuf.Replace(strSrc,strRpl); // ������� � ������
			j++;
			strFnd.Format("~%i~",j);
//AfxMessageBox("j = "+strFnd);
			SetRplStrR(strBuf, strFnd, j, a);
		}
		else{
			return false;
		}
	}
	else{
		return false;
	}
//	j--;
	strFnd.Format("~%i~",j);
	return false;
}
//********************************************************
/*	CString strSrc=_T("");
	CString strRpl=_T("");
	if(iB!=-1 && iE!=-1){

		iCnt = iE-iB+1;
		if(j>=10){
			iCnt+=3;
		}
		else if(j>=100){
			iCnt+=4;
		}
		else{
			iCnt+=2;
		}

		int tCnt,rCnt;
		tCnt=rCnt=0;

		strSrc=strBuf.Mid(iB,iCnt);
		strRpl = GetRplStr(a,j);

		strBuf.Replace(strSrc,strRpl); // ������� � ������


//AfxMessageBox(strSrc+'\n'+strRpl);

		tCnt = strSrc.GetLength(); 
		rCnt = strRpl.GetLength(); 
		if(rCnt==0){
			strRpl = _T("");
		}
		if(rCnt>tCnt){
			strRpl = strRpl.Left(tCnt);
		}
		if(j<=20){
			strBuf.Replace(strSrc,strRpl);
		}
//		if(j==20){
//			AfxMessageBox(strBuf);
//		}

		j++;

		return true;
	}
	else{
		return false;
	}

}
*/

void CBlokRtf::GetStdPWt(CString strBuf, CStdioFile &inF, CStdioFile &otF, CString strFnd, int& curPrv)
{
//������ � ����� �� ��������� ���������
	int tPos,iCnt;
	CString strWrt,s;
	inF.ReadString(strBuf);

	if((tPos = strBuf.Find(strFnd))!=-1){
		strWrt = strBuf.Left(tPos);   //������ ��� ������ � ���. ����

		iCnt = strWrt.GetLength();

//		s.Format("����� ����� curPrv = %i",curPrv);
//		AfxMessageBox("GetStdPWt "+strFnd+'\n'+"strWrt = "+strWrt +'\n'+ s+'\n'+"strBuf = "+strBuf);
//		AfxMessageBox(strWrt);
		otF.WriteString(strWrt);//��������� ��� �� bStr

		curPrv+=iCnt;
		//��������� ��������� �� ����� ������� �� ���.����� !!!
		inF.Seek(curPrv,CFile::begin);
			curPrv = inF.GetPosition();
//			s.Format(" ����� �������� Seek curPrv = %i ����� �� iCnt = %i",curPrv,iCnt);
//		AfxMessageBox(s);
		return;// curPos;
	}
	else{						// �� �����, ����� ������ ������
//AfxMessageBox("�� ����� \n"+strBuf);
		otF.WriteString(strBuf);	// �������
//		inF.ReadString(strBuf);
		curPrv = inF.GetPosition();
		GetStdPWt(strBuf, inF, otF, strFnd,curPrv);
		return;
	}
	
}

void CBlokRtf::GetStdPWS(CString strBuf, CStdioFile &inF, CStdioFile &otF, CString strFnd, int &curPrv, CString &strOut)
{
	int tPos,iCnt;
	CString strWrt,s;
	inF.ReadString(strBuf);

	if((tPos = strBuf.Find(strFnd))!=-1){
		strWrt = strBuf.Left(tPos);   //������ ��� ������ � ���. ����

		iCnt = strWrt.GetLength();

//		s.Format("����� ����� curPrv = %i",curPrv);
//		AfxMessageBox("GetStdPWt "+strFnd+'\n'+"strWrt = "+strWrt +'\n'+ s+'\n'+"strBuf = "+strBuf);
//		AfxMessageBox(strWrt);

		strOut+=strWrt;

		curPrv+=iCnt;
		//��������� ��������� �� ����� ������� �� ���.����� !!!
		inF.Seek(curPrv,CFile::begin);
			curPrv = inF.GetPosition();
//			s.Format(" ����� �������� Seek curPrv = %i ����� �� iCnt = %i",curPrv,iCnt);
//		AfxMessageBox(s);
		return;// curPos;
	}
	else{						// �� �����, ����� ������ ������
//AfxMessageBox("�� ����� \n"+strBuf);
//		otF.WriteString(strBuf);	// �������
		strOut+=strBuf;
		curPrv = inF.GetPosition();
		GetStdPWS(strBuf, inF, otF, strFnd,curPrv,strOut);
		return;
	}

}
