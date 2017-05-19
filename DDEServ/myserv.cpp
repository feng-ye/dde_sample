// myserv.cpp

#include "stdafx.h"
#include "ddeserv.h"
#include "myserv.h"

CMyServer::CMyServer()
{
}

CMyServer::~CMyServer()
{
}

void CMyServer::Status(LPCTSTR pszFormat, ...)
{
    TCHAR buf[1024];
	va_list arglist;
	va_start(arglist, pszFormat);
    _vsntprintf_s(buf, ARRAYSIZE(buf), pszFormat, arglist);
	va_end(arglist);

    STATUS(buf);
}

BOOL CMyServer::OnCreate()
{
    //
    // Add our own topics and items
    //

    m_DataTopic.Create(_T("Data"));
    AddTopic(&m_DataTopic);

    m_StringItem1.Create(_T("String1"));
    m_DataTopic.AddItem(&m_StringItem1);

    m_StringItem2.Create(_T("String2"));
    m_DataTopic.AddItem(&m_StringItem2);

    //
    // Set up some data in the strings
    //

    m_StringItem1.SetData(_T("This is string 1"));
    m_StringItem2.SetData(_T("This is string 2"));

    return TRUE;
}

void CMyStringItem::OnPoke()
{
    STATUS(_T("%s is now %s"),
           m_strName.c_str(),
           GetData());
}

CMyTopic::CMyTopic()
{
}

BOOL CMyTopic::Exec(void* pData, DWORD dwSize)
{
    STATUS(_T("Exec: %s"), (TCHAR*)pData);
    return TRUE;
}
