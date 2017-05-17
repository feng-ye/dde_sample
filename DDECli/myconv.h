// myconv.h

class CDDECliDlg;

class CMyConv : public CDDEConv
{
public:
    CMyConv();
    CMyConv(CDDEServer* pServer, CDDECliDlg* pDlg);
    virtual BOOL AdviseData(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                            void* pData, DWORD dwSize);

    CDDECliDlg* m_pDlg;

};
