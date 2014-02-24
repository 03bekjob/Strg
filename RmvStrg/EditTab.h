#pragma once

#define WM_SWITCHFOCUS (WM_USER+20)
#define WM_ADDREC	   (WM_USER+21)

// CEditTab

class CEditTab : public CEdit
{
	DECLARE_DYNAMIC(CEditTab)

public:
	CEditTab();
	virtual ~CEditTab();

protected:
	DECLARE_MESSAGE_MAP()
public:
//	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
};


