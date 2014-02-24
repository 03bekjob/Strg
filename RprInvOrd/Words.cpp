// Words.cpp: implementation of the CWords class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Words.h"
#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CWords::CWords()
{

}

CWords::~CWords()
{

}



CString CWords::GetWords(double Dgtl, bool bNm, bool bMn, bool bRpl)
{
	CString str;
	CString	razriad[6][3] = 
	{	{ _T(" копейка"), _T(" копейки"),   _T(" копеек")     },//0
		{ _T(" рубль"),   _T(" рубля"),     _T(" рублей")     },//1  
		{ _T(" тысяча"),  _T(" тысячи"),    _T(" тысяч")      },//2
		{ _T(" миллион"), _T(" миллиона"),  _T(" миллионов")  },//3
		{ _T(" миллиард"),_T(" миллиарда"), _T(" миллиардов") },//4
		{ _T(" триллион"),_T(" триллиона"), _T(" триллионов") } //5
	};

	CString edinicy[10][2] =
	{	{ _T(" ноль"),   _T(" ноль")   }, //0
		{ _T(" один"),   _T(" одна")   }, //1
		{ _T(" два"),    _T(" две")    }, //2
		{ _T(" три"),    _T(" три")    }, //3
		{ _T(" четыре"), _T(" четыре") }, //4
		{ _T(" пять"),   _T(" пять")   }, //5
		{ _T(" шесть"),  _T(" шесть")  }, //6
		{ _T(" семь"),   _T(" семь")   }, //7
		{ _T(" восемь"), _T(" восемь") }, //8
		{ _T(" девять"), _T(" девять") }  //9
	};
	if(!bRpl){
		edinicy[1][0]=_T(" одно");
	}

	CString nadtsat[9] = 
	{	_T(" одиннадцать"),  //0
		_T(" двенадцать"),	 //1
		_T(" тринадцать"),	 //2
		_T(" четырнадцать"), //3
		_T(" пятнадцать"),	 //4
		_T(" шестнадцать"),	 //5
		_T(" семнадцать"),	 //6
		_T(" восемьнадцать"),//7
		_T(" девятьнадцать") //8
	};

	CString desiatki[9] = 
	{	_T(" десять"),		//0
		_T(" двадцать"),	//1
		_T(" тридцать"),	//2
		_T(" сорок"),		//3
		_T(" пятьдесят"),	//4
		_T(" шестьдесят"),	//5
		_T(" семьдесят"),	//6
		_T(" восемьдесят"),	//7
		_T(" девяносто")	//8
	};

	CString sotni[9] = 
	{	_T(" сто"),			//0
		_T(" двести"),		//1
		_T(" триста"),		//2
		_T(" четыреста"),	//3
		_T(" пятьсот"),		//4
		_T(" шестьсот"),	//5
		_T(" семьсот"),		//6
		_T(" восмьсот"),	//7
		_T(" девятьсот")	//8
	};

	CStringArray textelem;
	CString cn,t,K,vsl,s;
	CString vsot,vdes,vedi,vrazr;
	int edi,des,sot;
		edi=des=sot=0;
	int iLnCn;
	bool RubIsNull = true;

	double dInt;
	dInt  = modf(Dgtl,&dInt);
	if(dInt>0){
		RubIsNull = false;
	}

	cn.Format(_T("%20.2f"),Dgtl);
	cn.TrimRight();
	cn.TrimLeft();
//AfxMessageBox(cn);
	iLnCn = cn.GetLength();
//	s.Format("iLnCn = %i",iLnCn);
//AfxMessageBox(s);
	while(iLnCn>0){
		if(iLnCn>=3){
			textelem.Add(cn.Right(3));
		}
		else{

			s = cn.Left(3);
			switch(s.GetLength()){
			case 0:
				s=_T("000");
				break;
			case 1:
				s=_T("00")+s;
				break;
			case 2:
				s=_T("0")+s;
				break;
			}

			textelem.Add(s); 
		}

		iLnCn = ((iLnCn>=3) ? iLnCn-3:0);
		cn = cn.Left(iLnCn);

//AfxMessageBox("Заполняю массив cn = "+cn);
	}

//********************************************* Копейки
//	for(int i=0;i<textelem.GetSize();i++){
//		AfxMessageBox("Массив "+textelem[i]);
//	}


	t = _T("");
	if(bMn){
		K = textelem[0];

		vsl = K.Right(2);
		edi = atoi(K.Right(1));
//		edi = _wtoi(K.Right(1));
		vsl = _T(" ")+vsl;

		switch(edi){
		case 0:
			if(bNm){
				vsl += razriad[0][2];
			}
			break;
		case 1:
			if(bNm){
				vsl += razriad[0][0];
			}
			break;
		case 2:
		case 3:
		case 4:
			if(bNm){
				vsl += razriad[0][1];
			}
			break;
		default:
			if(bNm){
				vsl += razriad[0][2];
			}
			break;
		}
	}
	t=vsl;
//AfxMessageBox("t1  = "+t);
//****************************************** Конец копейкам

	for(int j=1;j<textelem.GetSize();j++){
		K = textelem[j];
//AfxMessageBox("K = "+K);
		vsot  = _T("");
		vdes  = _T("");
		vedi  = _T("");
		vrazr = _T("");
		vsl   = _T("");

		sot = atoi(K.Left(1));
		des = atoi(K.Mid(1,1));
		edi = atoi(K.Right(1));

/*		sot = _wtoi(K.Left(1));
		des = _wtoi(K.Mid(1,1));
		edi = _wtoi(K.Right(1));
*/
/*		CString s1,s2,s3;
		s1.Format("sot = %i ",sot);
		s2.Format("des = %i ",des);
		s3.Format("edi = %i ",edi);
AfxMessageBox(s1+s2+s3);
*/
		vsot = (sot==0 ? _T(""):sotni[sot-1]);

		if(des>1){
			vdes = desiatki[des-1];				  //3 //2//1 	
			vedi = (edi==0 ? _T(""):edinicy[edi][j==2 ? 1:0]);
		}
		else if(des==1){
			if(edi==0){
				vdes = desiatki[0];
			}
			else{
				vdes = _T("");
				vedi = nadtsat[edi-1];
			}
		}
		else{							//edi+1    //3 //2//1
			vedi = (edi==0 ? _T(""):edinicy[edi][j==2 ? 1:0]);
		}

		vsl = vsot + vdes + vedi;
//AfxMessageBox("vsot + vdes + vedi"+vsl);

		if(vsl<_T(" ")){
			if(bNm){
				if(j==1 && !RubIsNull){
					vsl += razriad[j][2];
				}
			}
		}
		else{
			if(edi==0 || edi>4){
//				if(bNm){
				if(j>1){
					vsl += razriad[j][2];
				}
//				}
			}
			else{
				if(des==1){
//					if(bNm){
					if(j>1){
						vsl += razriad[j][2];
					}
//					}
				}
				else{
					if(edi==1){
//						if(bNm){
						if(j>1){
							vsl+=razriad[j][0];
						}
//						}
					}
					else{
//						if(bNm){
						if(j>1){
							vsl+=razriad[j][1];
						}
//AfxMessageBox("ok "+vsl);
//						}
					}
				}
			}
		}

		t = vsl + t;
//AfxMessageBox("Конец цикла t = "+t);
	}
	t.TrimRight();
	t.TrimLeft();
//AfxMessageBox(t);
	return t;
}
