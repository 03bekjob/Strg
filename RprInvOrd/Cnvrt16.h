#pragma once
#include "afxcoll.h"

class CCnvrt16
{
public:
	CCnvrt16(void);
public:
	~CCnvrt16(void);
public:
	void OnCnvrt16(const CString& src, CString& dst);
public:
	CMapStringToString map16;
};
