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
	{	{ _T(" �������"), _T(" �������"),   _T(" ������")     },//0
		{ _T(" �����"),   _T(" �����"),     _T(" ������")     },//1  
		{ _T(" ������"),  _T(" ������"),    _T(" �����")      },//2
		{ _T(" �������"), _T(" ��������"),  _T(" ���������")  },//3
		{ _T(" ��������"),_T(" ���������"), _T(" ����������") },//4
		{ _T(" ��������"),_T(" ���������"), _T(" ����������") } //5
	};

	CString edinicy[10][2] =
	{	{ _T(" ����"),   _T(" ����")   }, //0
		{ _T(" ����"),   _T(" ����")   }, //1
		{ _T(" ���"),    _T(" ���")    }, //2
		{ _T(" ���"),    _T(" ���")    }, //3
		{ _T(" ������"), _T(" ������") }, //4
		{ _T(" ����"),   _T(" ����")   }, //5
		{ _T(" �����"),  _T(" �����")  }, //6
		{ _T(" ����"),   _T(" ����")   }, //7
		{ _T(" ������"), _T(" ������") }, //8
		{ _T(" ������"), _T(" ������") }  //9
	};
	if(!bRpl){
		edinicy[1][0]=_T(" ����");
	}

	CString nadtsat[9] = 
	{	_T(" �����������"),  //0
		_T(" ����������"),	 //1
		_T(" ����������"),	 //2
		_T(" ������������"), //3
		_T(" ����������"),	 //4
		_T(" �����������"),	 //5
		_T(" ����������"),	 //6
		_T(" �������������"),//7
		_T(" �������������") //8
	};

	CString desiatki[9] = 
	{	_T(" ������"),		//0
		_T(" ��������"),	//1
		_T(" ��������"),	//2
		_T(" �����"),		//3
		_T(" ���������"),	//4
		_T(" ����������"),	//5
		_T(" ���������"),	//6
		_T(" �����������"),	//7
		_T(" ���������")	//8
	};

	CString sotni[9] = 
	{	_T(" ���"),			//0
		_T(" ������"),		//1
		_T(" ������"),		//2
		_T(" ���������"),	//3
		_T(" �������"),		//4
		_T(" ��������"),	//5
		_T(" �������"),		//6
		_T(" ��������"),	//7
		_T(" ���������")	//8
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

	cn.Format("%20.2f",Dgtl);
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

//AfxMessageBox("�������� ������ cn = "+cn);
	}

//********************************************* �������
//	for(int i=0;i<textelem.GetSize();i++){
//		AfxMessageBox("������ "+textelem[i]);
//	}


	t = _T("");
	if(bMn){
		K = textelem[0];

		vsl = K.Right(2);
		edi = atoi(K.Right(1));
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
//****************************************** ����� ��������

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
				if(bNm){
					vsl += razriad[j][2];
				}
			}
			else{
				if(des==1){
					if(bNm){
						vsl += razriad[j][2];
					}
				}
				else{
					if(edi==1){
						if(bNm){
							vsl+=razriad[j][0];
						}
					}
					else{
//						if(bNm){
							vsl+=razriad[j][1];
//AfxMessageBox("ok "+vsl);
//						}
					}
				}
			}
		}

		t = vsl + t;
//AfxMessageBox("����� ����� t = "+t);
	}
	t.TrimRight();
	t.TrimLeft();
//AfxMessageBox(t);
	return t;
}

CString CWords::GetWrdMnth(CString strDate, int j, bool bAll)
{
	CString str=_T("");
	CString month[12][2] =
	{	{ _T("������" ), _T("������")  }, //0
		{ _T("�������"), _T("�������") }, //1
		{ _T("����"),    _T("�����")   }, //2
		{ _T("������"),  _T("������")  }, //3
		{ _T("���"),	 _T("���")	 },   //4
		{ _T("����"),    _T("����")    }, //5
		{ _T("����"),    _T("����")    }, //6
		{ _T("������"),  _T("�������") }, //7
		{ _T("��������"),_T("��������")}, //8
		{ _T("�������"), _T("�������") }, //9
		{ _T("������"),  _T("������")  }, //10
		{ _T("�������"), _T("�������") }  //11
	};
	CString strD,strM,strY,strIM;
	int iNum;
	int tPos = strDate.Find(_T("."));
	strD = strDate.Left(tPos);
//AfxMessageBox(strD);
	strIM = strDate.Mid(tPos+1,strDate.Find(_T(".")));
//AfxMessageBox(strIM);
	strY = strDate.Right(4);
	iNum = atoi(strIM);
	strM = month[iNum-1][j];
	if(bAll){
		str+=strD + _T(" ");
		str+=strM + _T(" ");
		str+=strY + _T(" ");
	}
	else{
		str = strM;
	}
	return str;
}
