// myconv.cpp

#include "stdafx.h"
#include "ddecli.h"
#include "myclient.h"
#include "myconv.h"
#include "ddecldlg.h"

CMyConv::CMyConv()
{
    m_pDlg = NULL;
}

CMyConv::CMyConv(CDDEServer* pServer, CDDECliDlg* pDlg)
: CDDEConv(pServer)
{
    m_pDlg = pDlg;
}

BOOL CMyConv::AdviseData(UINT wFmt, LPCTSTR pszTopic,
                         LPCTSTR pszItem,
                         void* pData, DWORD dwSize)
{
    STATUS(_T("AdviseData: %s|%s: %s"),
           (LPCTSTR) pszTopic,
           (LPCTSTR) pszItem,
           (LPCTSTR) pData);
    ASSERT(m_pDlg);
    m_pDlg->NewData(pszItem, (LPCTSTR)pData);

    return TRUE;
}
