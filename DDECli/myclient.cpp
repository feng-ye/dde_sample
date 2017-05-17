// myclient.cpp

#include "stdafx.h"
#include "ddecli.h"
#include "myclient.h"

CMyClient::CMyClient()
    : CDDEServer(AfxGetAppName())
{
}

void CMyClient::Status(LPCTSTR pszFormat, ...)
{
    TCHAR buf[1024];
	va_list arglist;
	va_start(arglist, pszFormat);
    _vsntprintf_s(buf, ARRAYSIZE(buf), pszFormat, arglist);
	va_end(arglist);

    STATUS(buf);
}


