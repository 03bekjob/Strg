extern "C" __declspec(dllexport) BSTR GetNameCol(_RecordsetPtr rs,short i);
extern "C" __declspec(dllexport) VARIANT GetValueRec(_RecordsetPtr rs, short i);
extern "C" __declspec(dllexport) BOOL IsEmptyRec(_RecordsetPtr rs);
extern "C" __declspec(dllexport) BOOL OnFindInGrid(CString strCod, _RecordsetPtr rs, short i, bool& Flg);
extern "C" __declspec(dllexport) BOOL OnFindInGridP(CString strCod, _RecordsetPtr& rs, short i, bool& Flg);
extern "C" __declspec(dllexport) BSTR GetBSTRFind(int Tp, _bstr_t bstrName,CString s);
extern "C" __declspec(dllexport) long OnChangeCombo(CDatacombo2 &m_DataCombo, _RecordsetPtr rsCombo, short iC);
extern "C" __declspec(dllexport) void OnFindInCombo(CString strCod, CDatacombo2 *pCombo, _RecordsetPtr rsCombo, short i=0, short iBound=1);
extern "C" __declspec(dllexport) int GetTypeCol(_RecordsetPtr rs,short i);