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
// получение строки от тек. поз. до искомой строки
// включая саму строку поиска
// установка указателя за последним символом искомой строки
// пишет !!!

	CString strBuf=_T(""),s;

//	s.Format("до считки curPos = %i",curPos);
//	AfxMessageBox(s);
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//	s.Format("после считки curPos = %i",curPos);
//	AfxMessageBox(s);

	int tPos,iCnt;
	CString strTg=_T("");

	if((tPos=strBuf.Find(strFnd))!=-1){ // Нашёл
		strTg = strBuf.Left(tPos+strFnd.GetLength());
//AfxMessageBox("Нашёл в GetStrWrtStr \n"+strTg);
		iCnt = strTg.GetLength();
		oF.Write(strTg,iCnt);			// Записал
		curPos+= iCnt;
		s.Format(" Нашёл GetStrWrtStr curPos = %i",curPos);
AfxMessageBox(s+'\n'+strTg);
		inF.Seek(curPos,CFile::begin); // поставил указатель ЗА
									   // искомую строку
		return; // curPos;
	}
	else{
		oF.Write(strBuf,iBlk);					// Записал 
//		AfxMessageBox("Не нашёл GetStrWrtStr\n"+strBuf);
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);			// Переместил
//		inF.Seek(iBlk,CFile::current);			// Переместил
		curPos = inF.GetPosition();
		GetStrWrtStr(iBlk,inF,oF,strFnd,curPos);
		return;
	}

}

int CBlokRtf::GetPosWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int& curPos)
{
//Указатель устанавливает до строки поиска
//пишет в файл без строки поиска
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	int tPos;//=0;				// Позиция в strBuf
	int iCnt = 0;
	bool bEnd=true;
		strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//-------------------------------Ищу вхождение  подстроки
		if((tPos = strBuf.Find(strFnd))!=-1){
//		AfxMessageBox("Нашёл\n"+strBuf);
			strWrt = strBuf.Left(tPos);   //Считал для записи в цел. файл
			int iDlt=0;
			iCnt = strWrt.GetLength();
			iDlt = GetLnFd(strWrt,iCnt);	//подсчёт симв.переноса строк
//			s.Format("tPos = %i ",tPos);
			
//		AfxMessageBox("GetPosWrt "+strFnd+'\n'+strWrt +'\n'+ s + '\n'+strBuf);
//		AfxMessageBox(strWrt);
			oF.Write(strWrt,iCnt);//сбрасываю что до bStr
			curPos+= (iCnt+iDlt);// После сброса до bStr
			//Перемещаю указатель на новою позицию от нач.файла !!!
			inF.Seek(curPos,CFile::begin);
//			curPos = inF.GetPosition();
//			s.Format(" после нашёл Seek curPos = %i",curPos);
//		AfxMessageBox(s);
			return curPos;
		}
		else{						// Не нашёл, можно читать дальше
//AfxMessageBox("Не нашёл \n"+strBuf);
			oF.Write(strBuf,iBlk);			// Записал
			curPos+=iBlk;
			inF.Seek(curPos,CFile::begin);	// Переместил
//			inF.Seek(iBlk,CFile::current);	// Переместил
//			curPos = inF.GetPosition(); 
//		s.Format("Не нашёл GetPosWrt curPos = %i",curPos);
//		AfxMessageBox(s+'\n'+strBuf);
			GetPosWrt(iBlk,strFnd,inF,oF,curPos);
			//return;
			return curPos;
		}
}

CString CBlokRtf::GetStrBuf(int& iBlk, CFile &inF,int& curPos, bool& bEnd)
{
// Читает в буфер
// переводит в стринг
// /*перемещает указатель в файле на велечину блока  iBlk*/
	int bPos,ePos;
	int iEnd=0;
//	int iCnt;
	CString strBuf,s;
	bPos=ePos=0;
	char* buf = new char[iBlk];

	bPos = curPos; // Запоминаю поз Входа

	iEnd = inF.Read(buf,iBlk); // Сколько считали на самом деле
		if(iEnd==0){
	//		s.Format("iEnd = %i",iEnd);
	//AfxMessageBox(s+"  iEnd==0");
			inF.Seek(bPos,CFile::begin);
			bEnd = false;
			return "";
		}
		if(iEnd<iBlk){	//******************* Прошли конец файла				   
			delete [] buf;
	/*		for(int i=0;i<iBlk;i++){
				*(buf+i)=_T('');
			}
	*/
	//		strBuf = _T("");

			iBlk = iEnd;				   // Поменял размер блока	
			char* bufE = new char[iBlk];
			inF.Seek(bPos,CFile::begin);  // вернулись на старое место

			curPos = inF.GetPosition();
//			s.Format("После возврата на стар место curPos = %i",curPos);
//AfxMessageBox(s);

			inF.Read(bufE,iBlk);		   // прочитал с новым iBlk
	//		bPos = curPos+iBlk;
//			s.Format("iBlk = %i curPos = %i, bPos = %i",iBlk,curPos,bPos);
//	AfxMessageBox(s+"  iEnd<iBlk");
			for(int i=0;i<iBlk;i++){
				strBuf+=*(bufE+i);
			}
//			s.Format("iBlk = %i curPos = %i, bPos = %i",iBlk,curPos,bPos);
//	AfxMessageBox(s+"  iEnd<iBlk"+'\n'+"strBuf = "+strBuf);
			inF.Seek(curPos,CFile::begin);	// Вернулся на старое место
			curPos = inF.GetPosition();
			delete [] bufE;
			return strBuf;
		}
//	iBlk = iEnd; // Контроль блок такой сколько считано
	for(int i=0;i<iBlk;i++){
		strBuf+=*(buf+i);
	}
	inF.Seek(bPos,CFile::begin);	// Вернулся на старое место
//	curPos = inF.GetPosition();
/*	s.Format("curPos = %i, bPos = %i",curPos,bPos);
AfxMessageBox(s);
*/
	delete [] buf;
	return strBuf;
}

void CBlokRtf::GetLStrWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int &curPos, CString &strOut)
{
//Указатель устанавливает до строки поиска
//пишет в файл без строки поиска
//возвращает считанный весь буфер в виде строки strOut
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	int iCnt = 0;
	int tPos;//=0;				// Позиция в strBuf
	bool bEnd=true;
		strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
									//Читаю из вхд файла новый буфер
//-------------------------------Ищу вхождение  подстроки
		if((tPos = strBuf.Find(strFnd))!=-1){
			strWrt = strBuf.Left(tPos);   //Считал для записи в цел. файл
			iCnt = strWrt.GetLength(); 
			int iDlt=0;
			iDlt = GetLnFd(strWrt,iCnt);	//подсчёт симв.переноса строк
			strOut+= strWrt;
//			s.Format("tPos = %i ",tPos);
			
//		AfxMessageBox("Нашёл GetLStrWrt\n"+strFnd+'\n'+strWrt +'\n'+ s + '\n'+strBuf);
//		AfxMessageBox(strWrt);

			oF.Write(strWrt, iCnt);
			curPos+= (iCnt+iDlt);			
//			s.Format(" после нашёл curPos+=iCnt = %i",curPos);
//		AfxMessageBox(s);
			//Перемещаю указатель на новою позицию от нач.файла !!!
			inF.Seek(curPos,CFile::begin);
//			curPos = inF.GetPosition();
//			s.Format(" после нашёл Seek curPos = %i",curPos);
//		AfxMessageBox(s);
			return;
		}
		else{						// Не нашёл, можно читать дальше
//AfxMessageBox("Не нашёл GetLStrWrt\n"+strBuf);
			oF.Write(strBuf,iBlk);
			strOut+=strBuf;
			curPos+=iBlk;
			inF.Seek(curPos,CFile::begin);	// Переместил
//			inF.Seek(iBlk,CFile::current);
//			curPos = inF.GetPosition(); // Поз. после inF.Read
//			s.Format(" после не нашёл curPos = %i",curPos);
//		AfxMessageBox(s);
			GetLStrWrt(iBlk,strFnd,inF,oF,curPos,strOut);
			return;
		}

}

void CBlokRtf::GetRStrWrt(int iBlk, CString strFnd, CFile &inF, CFile &oF, int &curPos, CString &strOut)
{
//Указатель устанавливает за строкой поиска
//пишет в файл со строкой поиска
//возвращает считанный весь буфер в виде строки strOut
	CString strBuf,strTg,s;
	CString strWrt=_T("");
	strTg = _T("");
	int tPos;//=0;				// Позиция в strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("Исх строка\n"+strBuf);
//AfxMessageBox("Подстрока\n"+strFnd);
//-------------------------------Ищу вхождение  подстроки
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos+strFnd.GetLength());
		iCnt = strTg.GetLength();
		strOut+=strTg;
		oF.Write(strTg,iCnt);			// Записал
//		s.Format(" Нашёл GetRStr curPos = %i",curPos);
//AfxMessageBox("Нашёл в GetRStrWrtStr \n"+s+'\n'+strTg);
		curPos+= iCnt;
		inF.Seek(curPos,CFile::begin); // поставил указатель ЗА
//			s.Format(" после нашёл Seek curPos = %i",curPos);
//		AfxMessageBox(s);
		return;
	}
	else{						// Не нашёл, можно читать дальше
		AfxMessageBox("Не нашёл GetRStrWrt\n"+strBuf);
//		AfxMessageBox("накопление\n"+strOut);
		oF.Write(strBuf,iBlk);			// Записал
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// Переместил
//		inF.Seek(iBlk,CFile::current);	// Переместил
		curPos = inF.GetPosition();
		GetRStrWrt(iBlk,strFnd,inF,oF,curPos,strOut);
		return;
	}

}



void CBlokRtf::GetRStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString& strOut)
{
//Считывает строку по подстроке
//указатель устанавливается за стракой
//ничего НЕ пишет

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// Позиция в strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("Исх строка\n"+strBuf+'\n'+"Подстрока\n"+strFnd);
//-------------------------------Ищу вхождение  подстроки
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos+strFnd.GetLength());
//AfxMessageBox("Нашёл в GetR \n"+strTg);
		iCnt = strTg.GetLength();
		strOut+=strTg;
		s.Format(" Нашёл GetRStr curPos = %i",curPos);
AfxMessageBox("Нашёл в GetRStr \n"+s+'\n'+"strTg = "+strTg);
		curPos+= iCnt;
		inF.Seek(curPos,CFile::begin); // поставил указатель ЗА
//			s.Format(" после нашёл Seek curPos = %i",curPos);
//		AfxMessageBox(s);
		return;
	}
	else{						// Не нашёл, можно читать дальше
//		AfxMessageBox("Не нашёл GetRStr\n"+strBuf);
//		AfxMessageBox("накопление\n"+strOut);
		strOut+=strBuf;
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// Переместил
//		inF.Seek(iBlk,CFile::current);	// Переместил
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
	m.Format("Тек поз curPos = %i",curPos);
AfxMessageBox(m);

	char* b = new char[iBlk];

	iR = f.Read(b,iBlk);	// Долбанная функция

	curPos = f.GetPosition();
	m.Format("Поз после Read curPos = %i прочитано iR = %i",curPos,iR);
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
	n.Format("Длина строки s = %i k = %i",ln);
AfxMessageBox(s+'\n'+n);
*/
	curPos-=(iBlk+iDel);
//	curPos-=iBlk;

	m.Format("Поз для возврата curPos = %i",curPos);
AfxMessageBox(m);

	f.Seek(curPos,CFile::begin);

	curPos = f.GetPosition();
	m.Format("Тек поз после Seek назад на исх curPos = %i",curPos);
AfxMessageBox(m);
	delete [] b;

}

void CBlokRtf::GetLStr(int iBlk, CString strFnd, CFile &inF, int &curPos, CString &strOut)
{
//Считывает строку по подстроке
//указатель устанавливается перед подстрокой
// копит строку strOut
//ничего НЕ пишет

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// Позиция в strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//		AfxMessageBox("strBuf = \n"+strBuf);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("Исх строка\n"+strBuf+'\n'+"Подстрока\n"+strFnd);
//-------------------------------Ищу вхождение  подстроки
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos);
		iCnt = strTg.GetLength(); //tPos;
		int iDlt=0;
		iDlt = GetLnFd(strTg,iCnt);	//подсчёт симв.переноса строк
		strOut+=strTg;
//		strTg.Replace('\n','~');
//		s.Format(" Нашёл GetLStr пред curPos = %i iCnt = %i, iDlt = %i",curPos,iCnt,iDlt);
//AfxMessageBox("Нашёл в GetLStr \n"+s+'\n'+"strTg = "+strTg+'\n'+"strFnd = "+strFnd);
		curPos+= (iCnt+iDlt);				// Получаю новую curPos	
		inF.Seek(curPos,CFile::begin); // поставил указатель ЗА
//		s.Format("Новая поз после поиска curPos = %i",curPos);
//AfxMessageBox(s);
//		curPos = inF.GetPosition();
		return;
	}
	else{						// Не нашёл, можно читать дальше
//		AfxMessageBox("накопление\n"+strOut);
		strOut+=strBuf;
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// Переместил
//		s.Format("После буфера curPos = %i",curPos);
//		AfxMessageBox("Не нашёл GetLStr\n"+strBuf+'\n'+s);
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
//Считывает строку по подстроке
//указатель устанавливается перед подстрокой
//ничего НЕ пишет

	CString strBuf,strTg,s;
	strTg = _T("");
	int tPos;//=0;				// Позиция в strBuf
	int iCnt = 0;
	bool bEnd=true;
	strBuf = GetStrBuf(iBlk,inF,curPos,bEnd);
//		AfxMessageBox("strBuf = \n"+strBuf);
	if(!bEnd){
//		AfxMessageBox("False");
		return;
	}
//AfxMessageBox("Исх строка\n"+strBuf+'\n'+"Подстрока\n"+strFnd);
//-------------------------------Ищу вхождение  подстроки
	if((tPos = strBuf.Find(strFnd))!=-1){
		strTg = strBuf.Left(tPos);
		iCnt = strTg.GetLength(); //tPos;
		int iDlt=0;
		iDlt = GetLnFd(strTg,iCnt);	//подсчёт симв.переноса строк
//		strTg.Replace('\n','~');
//		s.Format(" Нашёл GetLStr пред curPos = %i iCnt = %i, iDlt = %i",curPos,iCnt,iDlt);
//AfxMessageBox("Нашёл в GetLStr \n"+s+'\n'+"strTg = "+strTg+'\n'+"strFnd = "+strFnd);
		curPos+= (iCnt+iDlt);				// Получаю новую curPos	
		inF.Seek(curPos,CFile::begin); // поставил указатель ЗА
//		s.Format("Новая поз после поиска curPos = %i",curPos);
//AfxMessageBox(s);
//		curPos = inF.GetPosition();
		return;
	}
	else{						// Не нашёл, можно читать дальше
//		AfxMessageBox("накопление\n"+strOut);
		curPos+=iBlk;
		inF.Seek(curPos,CFile::begin);	// Переместил
//		s.Format("После буфера curPos = %i",curPos);
//		AfxMessageBox("Не нашёл SetPos\n"+strBuf+'\n'+s);
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

	iB = strBuf.Find((LPCTSTR)strFnd);		// Нач подстроки

	if(iB!=-1){
		iE = strBuf.Find((LPCTSTR)strFnd,iB+1);	// Кон подстроки
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

			strSrc=strBuf.Mid(iB,iCnt); // исход строка
			strRpl = GetRplStr(a,j);	// строка для замены

			tCnt = strSrc.GetLength();  // длина исх
			rCnt = strRpl.GetLength();  // длина замены
			if(rCnt==0){
				strRpl = _T("");
			}
			if(rCnt>tCnt){
				strRpl = strRpl.Left(tCnt); // обрез строки замены
			}

			strBuf.Replace(strSrc,strRpl); // Заменил в буфере
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

		strBuf.Replace(strSrc,strRpl); // Заменил в буфере


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
//Читает и пишет до вхождения подстроки
	int tPos,iCnt;
	CString strWrt,s;
	inF.ReadString(strBuf);

	if((tPos = strBuf.Find(strFnd))!=-1){
		strWrt = strBuf.Left(tPos);   //Считал для записи в цел. файл

		iCnt = strWrt.GetLength();

//		s.Format("после читки curPrv = %i",curPrv);
//		AfxMessageBox("GetStdPWt "+strFnd+'\n'+"strWrt = "+strWrt +'\n'+ s+'\n'+"strBuf = "+strBuf);
//		AfxMessageBox(strWrt);
		otF.WriteString(strWrt);//сбрасываю что до bStr

		curPrv+=iCnt;
		//Перемещаю указатель на новою позицию от нач.файла !!!
		inF.Seek(curPrv,CFile::begin);
			curPrv = inF.GetPosition();
//			s.Format(" после смещения Seek curPrv = %i назад на iCnt = %i",curPrv,iCnt);
//		AfxMessageBox(s);
		return;// curPos;
	}
	else{						// Не нашёл, можно читать дальше
//AfxMessageBox("Не нашёл \n"+strBuf);
		otF.WriteString(strBuf);	// Записал
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
		strWrt = strBuf.Left(tPos);   //Считал для записи в цел. файл

		iCnt = strWrt.GetLength();

//		s.Format("после читки curPrv = %i",curPrv);
//		AfxMessageBox("GetStdPWt "+strFnd+'\n'+"strWrt = "+strWrt +'\n'+ s+'\n'+"strBuf = "+strBuf);
//		AfxMessageBox(strWrt);

		strOut+=strWrt;

		curPrv+=iCnt;
		//Перемещаю указатель на новою позицию от нач.файла !!!
		inF.Seek(curPrv,CFile::begin);
			curPrv = inF.GetPosition();
//			s.Format(" после смещения Seek curPrv = %i назад на iCnt = %i",curPrv,iCnt);
//		AfxMessageBox(s);
		return;// curPos;
	}
	else{						// Не нашёл, можно читать дальше
//AfxMessageBox("Не нашёл \n"+strBuf);
//		otF.WriteString(strBuf);	// Записал
		strOut+=strBuf;
		curPrv = inF.GetPosition();
		GetStdPWS(strBuf, inF, otF, strFnd,curPrv,strOut);
		return;
	}

}
