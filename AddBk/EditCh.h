#pragma once


// CEditCh

class CEditCh : public CEdit
{
	DECLARE_DYNAMIC(CEditCh)

public:
	CEditCh();
	virtual ~CEditCh();

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
protected:
	int m_iTp;
public:
	void SetTypeCol(int i);
};


