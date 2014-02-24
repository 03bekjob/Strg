// StrgView.cpp : implementation of the CStrgView class
//

#include "stdafx.h"
#include "Strg.h"

#include "StrgDoc.h"
#include "StrgView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CStrgView

IMPLEMENT_DYNCREATE(CStrgView, CView)

BEGIN_MESSAGE_MAP(CStrgView, CView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CView::OnFilePrintPreview)

	ON_COMMAND(ID_32772, &CStrgView::On32772)
	ON_COMMAND(ID_32773, &CStrgView::On32773)
	ON_COMMAND(ID_32774, &CStrgView::On32774)
	ON_COMMAND(ID_32775, &CStrgView::On32775)
	ON_COMMAND(ID_32776, &CStrgView::On32776)
	ON_COMMAND(ID_32777, &CStrgView::On32777)
	ON_COMMAND(ID_32778, &CStrgView::On32778)
	ON_COMMAND(ID_32779, &CStrgView::On32779)
	ON_COMMAND(ID_32780, &CStrgView::On32780)
	ON_COMMAND(ID_32782, &CStrgView::On32782)
	ON_COMMAND(ID_32783, &CStrgView::On32783)
	ON_COMMAND(ID_32784, &CStrgView::On32784)
	ON_COMMAND(ID_32786, &CStrgView::On32786)
	ON_COMMAND(ID_32787, &CStrgView::On32787)
	ON_COMMAND(ID_32788, &CStrgView::On32788)
	ON_COMMAND(ID_32789, &CStrgView::On32789)
	ON_COMMAND(ID_32790, &CStrgView::On32790)
	ON_COMMAND(ID_32796, &CStrgView::On32796)
	ON_COMMAND(ID_32791, &CStrgView::On32791)
	ON_COMMAND(ID_32792, &CStrgView::On32792)
	ON_COMMAND(ID_32793, &CStrgView::On32793)
	ON_COMMAND(ID_32794, &CStrgView::On32794)
	ON_COMMAND(ID_32795, &CStrgView::On32795)
	ON_COMMAND(ID_32798, &CStrgView::On32798)
	ON_COMMAND(ID_32799, &CStrgView::On32799)
	ON_COMMAND(ID_32800, &CStrgView::On32800)
	ON_COMMAND(ID_32801, &CStrgView::On32801)
	ON_COMMAND(ID_32802, &CStrgView::On32802)
	ON_COMMAND(ID_32803, &CStrgView::On32803)
	ON_COMMAND(ID_32804, &CStrgView::On32804)
	ON_COMMAND(ID_32805, &CStrgView::On32805)
	ON_COMMAND(ID_32808, &CStrgView::On32808)
	ON_COMMAND(ID_32809, &CStrgView::On32809)
	ON_COMMAND(ID_32810, &CStrgView::On32810)
	ON_COMMAND(ID_32811, &CStrgView::On32811)
	ON_COMMAND(ID_32812, &CStrgView::On32812)
	ON_COMMAND(ID_32813, &CStrgView::On32813)
	ON_COMMAND(ID_32814, &CStrgView::On32814)
	ON_COMMAND(ID_32816, &CStrgView::On32816)
	ON_COMMAND(ID_32817, &CStrgView::On32817)
	ON_COMMAND(ID_32819, &CStrgView::On32819)
	ON_COMMAND(ID_32820, &CStrgView::On32820)
	ON_COMMAND(ID_32822, &CStrgView::On32822)
	ON_COMMAND(ID_32823, &CStrgView::On32823)
	ON_COMMAND(ID_32827, &CStrgView::On32827)
	ON_COMMAND(ID_32828, &CStrgView::On32828)
	ON_COMMAND(ID_32829, &CStrgView::On32829)
	ON_COMMAND(ID_32842, &CStrgView::On32842)
	ON_COMMAND(ID_32843, &CStrgView::On32843)
	ON_COMMAND(ID_32844, &CStrgView::On32844)
	ON_COMMAND(ID_32845, &CStrgView::On32845)
	ON_COMMAND(ID_32846, &CStrgView::On32846)
	ON_COMMAND(ID_32852, &CStrgView::On32852)
	ON_COMMAND(ID_32854, &CStrgView::On32854)
	ON_COMMAND(ID_32855, &CStrgView::On32855)
	ON_COMMAND(ID_32856, &CStrgView::On32856)
	ON_COMMAND(ID_32857, &CStrgView::On32857)
	ON_COMMAND(ID_32859, &CStrgView::On32859)
	ON_COMMAND(ID_32860, &CStrgView::On32860)
	ON_COMMAND(ID_32833, &CStrgView::On32833)
	ON_COMMAND(ID_32835, &CStrgView::On32835)
	ON_COMMAND(ID_32861, &CStrgView::On32861)
	ON_COMMAND(ID_32862, &CStrgView::On32862)
END_MESSAGE_MAP()

// CStrgView construction/destruction

CStrgView::CStrgView()
{
	// TODO: add construction code here

}

CStrgView::~CStrgView()
{
}

BOOL CStrgView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

// CStrgView drawing

void CStrgView::OnDraw(CDC* /*pDC*/)
{
	CStrgDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	if (!pDoc)
		return;

	// TODO: add draw code for native data here
}


// CStrgView printing

BOOL CStrgView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);
}

void CStrgView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing
}

void CStrgView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}


// CStrgView diagnostics

#ifdef _DEBUG
void CStrgView::AssertValid() const
{
	CView::AssertValid();
}

void CStrgView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CStrgDoc* CStrgView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CStrgDoc)));
	return (CStrgDoc*)m_pDocument;
}
#endif //_DEBUG


// CStrgView message handlers

void CStrgView::On32772()
{//����������
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("ShotWord.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startShotWord");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32773()
{//�����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		hMod=AfxLoadLibrary(_T("Cldr.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startCldr");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32774()
{//���� ���� �������������
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidFrmPrv.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidFrmPrv");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32775()
{//�����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("AddBank.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startAddBank");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);
		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32776()
{//���. �����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("NameMan.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameMan");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32777()
{//���. ��������
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SbNmMan.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSbNmMan");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32778()
{//���. �����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("NameWom.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameWom");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32779()
{//���. ��������
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SbNmWom.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSbNmWom");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32780()
{//���� ����������
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidDoc.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidDoc");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32782()
{//���� �����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		hMod=AfxLoadLibrary(_T("TabVlt.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startTabVlt");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32783()
{//���� ������ �����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		hMod=AfxLoadLibrary(_T("VidCrs.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidCrs");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32784()
{//����� �����
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		hMod=AfxLoadLibrary(_T("CrsVlt.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr);
		pDialog func=(pDialog)GetProcAddress(hMod,"startCrsVlt");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);

	EndWaitCursor();
}

void CStrgView::On32786()
{//������������ ��
	BeginWaitCursor();

//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("NameUnit.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameUnit");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32787()
{//������� ���������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("UnitPls.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startUnitPls");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32788()
{//�����
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Order.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrder");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32789()
{//��� �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidBox.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidBox");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32790()
{//��� ��������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidUp.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidUp");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32796()
{//�����-��������
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OrdWrap.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdWrap");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32791()
{//�����-�����
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OrdStrg.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdStrg");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32792()
{//����� ����������
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Crtf.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startCrtf");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32793()
{//�������������
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Clsfk.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startClsfk");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32794()
{//�����-������
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OrdGroup.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOrdGroup");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32795()
{//������
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Grp.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startGrp");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}


void CStrgView::On32798()
{//��������� �������� �����
	BeginWaitCursor();
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("TngMvMany.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startTngMvMany");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32799()
{//�������� �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OprTo.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOprTo");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32800()
{//�������� �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OprFrom.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOprFrom");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32801()
{//������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Stg.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startStg");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32802()
{//�����������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("CtrAgnt.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startCtrAgnt");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32803()
{//����������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VendProd.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVendProd");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32804()
{//����������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("OwnProd.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startOwnProd");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32805()
{//�������������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("PrvCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startPrvCst");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32808()
{//�����������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Grmt.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startGrmt");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32809()
{//����, ���. ������, ������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("CdRgRs.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startCdRgRs");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32810()
{//��� ������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidAdd.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidAdd");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32811()
{//�����, �����������...
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC,strFndC1;
		bFndC   = FALSE;
		strFndC = _T("");
		strFndC1= _T("");
		hMod=AfxLoadLibrary(_T("NameStrt.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNameStrt");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,strFndC1,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32812()
{//�����, �-��...
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("ShStr.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startShStr");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32813()
{//���������������� ������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("ShAdmBld.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startShAdmBld");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32814()
{//��������������� �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("ShAdm.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startShAdm");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32816()
{//���
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("NDC.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startNDC");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32817()
{//��������� ���
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("TNDC.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startTNDC");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32819()
{//������ �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidDscAdd.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidDscAdd");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32820()
{//������ �������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidDscRmv.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidDscRmv");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32822()
{//������� ����
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("VidBsCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startVidBsCst");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32823()
{//�������� � ������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("Letter.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startLetter");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32827()
{//��������������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("PrcCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startPrcCst");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32828()
{//�������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SlsCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSlsCst");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32829()
{// % - ������� ���
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("PerCst.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startPerCst");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32842()
{//������
	// TODO: Add your command handler code here
}

void CStrgView::On32843()
{//���������
	// TODO: Add your command handler code here
}

void CStrgView::On32844()
{//����������
	// TODO: Add your command handler code here
}

void CStrgView::On32845()
{//������� ������
	// TODO: Add your command handler code here
}

void CStrgView::On32846()
{//������� �����
	// TODO: Add your command handler code here
}

void CStrgView::On32852()
{//�������� ������
	// TODO: Add your command handler code here
}

void CStrgView::On32854()
{//������ ���������
	// TODO: Add your command handler code here
}

void CStrgView::On32855()
{//����� �� ��������
	// TODO: Add your command handler code here
}

void CStrgView::On32856()
{//����� �� �����
	// TODO: Add your command handler code here
}

void CStrgView::On32857()
{//������� � ������� �����
	// TODO: Add your command handler code here
}

void CStrgView::On32859()
{//���� ������
	// TODO: Add your command handler code here
}

void CStrgView::On32860()
{//���� ������
	// TODO: Add your command handler code here
}

void CStrgView::On32833()
{//����. ������ �� ���.����� ���
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SlsDscMn.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSlsDscMn");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32835()
{//�������������� ������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SlsDscAg.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSlsDscAg");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);
//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32861()
{//��������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("SynCA.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startSynCA");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}

void CStrgView::On32862()
{//������ ���������
	BeginWaitCursor();
//		m_SlpDay.SetDate(t1);
		HMODULE hMod;
		BOOL bFndC;
		CString strFndC;
		bFndC   = FALSE;
		strFndC = _T("");
		hMod=AfxLoadLibrary(_T("RmvStrg.dll"));
		typedef BOOL (*pDialog)(CString,_ConnectionPtr,CString&,BOOL);
		pDialog func=(pDialog)GetProcAddress(hMod,"startRmvStrg");
		(func)(GetDocument()->m_strNameNT, GetDocument()->ptrConn,strFndC,bFndC);

//		m_SlpDay.SetDate(t1);

		AfxFreeLibrary(hMod);
	EndWaitCursor();
}
