// StrgDoc.h : interface of the CStrgDoc class
//


#pragma once


class CStrgDoc : public CDocument
{
protected: // create from serialization only
	CStrgDoc();
	DECLARE_DYNCREATE(CStrgDoc)

// Attributes
public:
	char* bsConn;
	DWORD max;
	CString m_Conn;
	_ConnectionPtr ptrConn;
	char m_UserName[25];
	char m_CompName[25];
	void GetUserNameBek(CString &strNT, CString &strSQL);
	CString m_strNameSQL;
	CString m_strNameNT;

// Operations
public:

// Overrides
public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);

// Implementation
public:
	virtual ~CStrgDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
};


