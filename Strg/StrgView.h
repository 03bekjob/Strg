// StrgView.h : interface of the CStrgView class
//


#pragma once


class CStrgView : public CView
{
protected: // create from serialization only
	CStrgView();
	DECLARE_DYNCREATE(CStrgView)

// Attributes
public:
	CStrgDoc* GetDocument() const;

// Operations
public:

// Overrides
public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~CStrgView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void On32772();
public:
	afx_msg void On32773();
public:
	afx_msg void On32774();
public:
	afx_msg void On32775();
public:
	afx_msg void On32776();
public:
	afx_msg void On32777();
public:
	afx_msg void On32778();
public:
	afx_msg void On32779();
public:
	afx_msg void On32780();
public:
	afx_msg void On32782();
public:
	afx_msg void On32783();
public:
	afx_msg void On32784();
public:
	afx_msg void On32786();
public:
	afx_msg void On32787();
public:
	afx_msg void On32788();
public:
	afx_msg void On32789();
public:
	afx_msg void On32790();
public:
	afx_msg void On32796();
public:
	afx_msg void On32791();
public:
	afx_msg void On32793();
public:
	afx_msg void On32794();
public:
	afx_msg void On32795();
public:
	afx_msg void On32792();
public:
	afx_msg void On32798();
public:
	afx_msg void On32799();
public:
	afx_msg void On32800();
public:
	afx_msg void On32801();
public:
	afx_msg void On32802();
public:
	afx_msg void On32803();
public:
	afx_msg void On32804();
public:
	afx_msg void On32805();
public:
	afx_msg void On32808();
public:
	afx_msg void On32809();
public:
	afx_msg void On32810();
public:
	afx_msg void On32811();
public:
	afx_msg void On32812();
public:
	afx_msg void On32813();
public:
	afx_msg void On32814();
public:
	afx_msg void On32816();
public:
	afx_msg void On32817();
public:
	afx_msg void On32819();
public:
	afx_msg void On32820();
public:
	afx_msg void On32822();
public:
	afx_msg void On32823();
public:
	afx_msg void On32827();
public:
	afx_msg void On32828();
public:
	afx_msg void On32829();
public:
	afx_msg void On32842();
public:
	afx_msg void On32843();
public:
	afx_msg void On32844();
public:
	afx_msg void On32845();
public:
	afx_msg void On32846();
public:
	afx_msg void On32852();
public:
	afx_msg void On32854();
public:
	afx_msg void On32855();
public:
	afx_msg void On32856();
public:
	afx_msg void On32857();
public:
	afx_msg void On32859();
public:
	afx_msg void On32860();
public:
	afx_msg void On32833();
public:
	afx_msg void On32835();
public:
	afx_msg void On32861();
public:
	afx_msg void On32862();
};

#ifndef _DEBUG  // debug version in StrgView.cpp
inline CStrgDoc* CStrgView::GetDocument() const
   { return reinterpret_cast<CStrgDoc*>(m_pDocument); }
#endif

