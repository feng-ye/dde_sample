// stddde.h

#pragma once

#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <tchar.h>
#include <ddeml.h>

#include <string>
#include <list>

typedef std::basic_string<TCHAR> CDDEString;

//
// Constants
//

#ifndef DDE_TIMEOUT
#define DDE_TIMEOUT     5000 // 5 seconds
#endif

#ifdef _UNICODE
#define DDE_CODEPAGE    CP_WINUNICODE
#define DDE_CF_TEXT     CF_UNICODETEXT
#else
#define DDE_CODEPAGE    CP_WINANSI
#define DDE_CF_TEXT     CF_TEXT
#endif

//
// String names for standard Windows Clipboard formats
//

#define SZCF_TEXT           _T("TEXT")
#define SZCF_BITMAP         _T("BITMAP")
#define SZCF_METAFILEPICT   _T("METAFILEPICT")
#define SZCF_SYLK           _T("SYLK")
#define SZCF_DIF            _T("DIF")
#define SZCF_TIFF           _T("TIFF")
#define SZCF_OEMTEXT        _T("OEMTEXT")
#define SZCF_DIB            _T("DIB")
#define SZCF_PALETTE        _T("PALETTE")
#define SZCF_PENDATA        _T("PENDATA")
#define SZCF_RIFF           _T("RIFF")
#define SZCF_WAVE           _T("WAVE")
#define SZCF_UNICODETEXT    _T("UNICODETEXT")
#define SZCF_ENHMETAFILE    _T("ENHMETAFILE")

//
// String names for some standard DDE strings not
// defined in DDEML.H
//

#define SZ_READY            _T("Ready")
#define SZ_BUSY             _T("Busy")
#define SZ_TAB              _T("\t")
#define SZ_RESULT           _T("Result")
#define SZ_PROTOCOLS        _T("Protocols")
#define SZ_EXECUTECONTROL1  _T("Execute Control 1")

//
// Helpers
//

CDDEString DDEGetFormatName(WORD wFmt);

//
// Generic counted object class
//

class CDDECountedObject
{
public:
    CDDECountedObject();
    virtual ~CDDECountedObject();
    int AddRef();
    int Release();

private:
    int m_iRefCount;
};

//
// String handle class
//

class CDDEServer;

class CHSZ
{
public:
    CHSZ();
    CHSZ(CDDEServer* pServer, LPCTSTR szName);
    virtual ~CHSZ();
    void Create(CDDEServer* pServer, LPCTSTR szName);
    operator HSZ() const { return m_hsz; }

    HSZ m_hsz;

protected:
    DWORD m_dwDDEInstance;
};

//
// DDE item class
//

class CDDETopic;

class CDDEItem
{
public:
    CDDEItem();
    virtual ~CDDEItem();
    void Create(LPCTSTR pszName);
    void PostAdvise();
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
    virtual BOOL Poke(UINT wFmt, void* pData, DWORD dwSize);
    virtual BOOL IsSupportedFormat(WORD wFormat) const;
    virtual const WORD* GetFormatList() const
        {return NULL;}
    virtual BOOL CanAdvise(UINT wFmt) const;

    CDDEString m_strName;       // name of this item
    CDDETopic* m_pTopic;        // pointer to the topic it belongs to

protected:
    HDDEDATA CreateDdeData(UINT wFmt, HSZ hszItem, void* pData, DWORD dwSize) const;
    HDDEDATA CreateDdeData(UINT wFmt, HSZ hszItem, const CDDEString& s) const;
    HDDEDATA CreateDdeDataA(HSZ hszItem, const CDDEString& s) const;
    HDDEDATA CreateDdeDataW(HSZ hszItem, const CDDEString& s) const;
};

//
// String item class
//

class CDDEStringItem : public CDDEItem
{
public:
    virtual void OnPoke(){}
    virtual void SetData(const TCHAR* pszData) {
        _ASSERT(pszData);
        m_strData = pszData;
        PostAdvise();
    }
    virtual const TCHAR* GetData() const
        {return m_strData.c_str();}
    operator const TCHAR*() const
        {return m_strData.c_str();}

protected:
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
    virtual BOOL Poke(UINT wFmt, void* pData, DWORD dwSize);
    virtual const WORD* GetFormatList() const;

    CDDEString m_strData;
};

//
// Item list class
//

typedef std::list<CDDEItem*> CDDEItemList;

//
// Topic class
//

class CDDEServer;

class CDDETopic
{
public:
    CDDETopic();
    virtual ~CDDETopic();
    void Create(LPCTSTR pszName);
    BOOL AddItem(CDDEItem* pItem);
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
    virtual BOOL Poke(UINT wFmt, LPCTSTR pszItem,
                      void* pData, DWORD dwSize);
    virtual BOOL Exec(void* pData, DWORD dwSize);
    virtual CDDEItem* FindItem(LPCTSTR pszItem);
    virtual const CDDEItem* FindItem(LPCTSTR pszItem) const;
    virtual BOOL CanAdvise(UINT wFmt, LPCTSTR pszItem) const;
    void PostAdvise(CDDEItem* pItem);

    CDDEString   m_strName;     // name of this topic
    CDDEServer*  m_pServer;     // ptr to the server which owns this topic
    CDDEItemList m_ItemList;    // List of items for this topic

protected:
};

//
// Topic list class
//

typedef std::list<CDDETopic*> CDDETopicList;

//
// Conversation class
//

class CDDEConv : public CDDECountedObject
{
public:
    CDDEConv();
    CDDEConv(CDDEServer* pServer);
    CDDEConv(CDDEServer* pServer, HCONV hConv, HSZ hszTopic);
    virtual ~CDDEConv();
    virtual BOOL ConnectTo(LPCTSTR pszService, LPCTSTR pszTopic);
    virtual BOOL Terminate();
    virtual BOOL AdviseData(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                            void* pData, DWORD dwSize);
    virtual BOOL Request(UINT wFmt, LPCTSTR pszItem, void** ppData, DWORD* pdwSize);
    virtual BOOL Advise(UINT wFmt, LPCTSTR pszItem);
    virtual BOOL Exec(UINT wFmt, LPCTSTR pszCmd);
    virtual BOOL Poke(UINT wFmt, LPCTSTR pszItem, void* pData, DWORD dwSize);

    CDDEServer* m_pServer;
    HCONV   m_hConv;            // Conversation handle
    HSZ     m_hszTopic;         // Topic name
};

//
// Conversation list class
//

typedef std::list<CDDEConv*> CDDEConvList;


//
// Topics and items used to support the 'system' topic in the server
//

class CDDESystemItem : public CDDEItem
{
protected:
    virtual const WORD* GetFormatList() const;
};

class CDDESystemItem_TopicList : public CDDESystemItem
{
protected:
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
};

class CDDESystemItem_ItemList : public CDDESystemItem
{
protected:
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
};

class CDDESystemItem_FormatList : public CDDESystemItem
{
protected:
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
};

class CDDEServerSystemTopic : public CDDETopic
{
protected:
    virtual BOOL Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData);
};


//
// Server class
// Note: this class is for a server which supports only one service
//

class CDDEServer
{
public:
    CDDEServer();
    virtual ~CDDEServer();
    BOOL Create(LPCTSTR pszServiceName,
                DWORD dwFilterFlags = 0,
                DWORD* pdwDDEInst = NULL,
                DWORD dwTimeout = DDE_TIMEOUT);
    void Shutdown();
    virtual BOOL OnCreate() {return TRUE;}
    virtual UINT GetLastError()
        {return ::DdeGetLastError(m_dwDDEInstance);}
    virtual HDDEDATA CustomCallback(WORD wType,
                                    WORD wFmt,
                                    HCONV hConv,
                                    HSZ hsz1,
                                    HSZ hsz2,
                                    HDDEDATA hData,
                                    DWORD dwData1,
                                    DWORD dwData2)
        {return NULL;}

    virtual BOOL Request(UINT wFmt, LPCTSTR pszTopic, HSZ hszItem,
                         HDDEDATA* phReturnData);
    virtual BOOL Poke(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                      void* pData, DWORD dwSize);
    virtual BOOL AdviseData(UINT wFmt, HCONV hConv, LPCTSTR pszTopic, LPCTSTR pszItem,
                      void* pData, DWORD dwSize);
    virtual BOOL Exec(LPCTSTR pszTopic, void* pData, DWORD dwSize);
    virtual void Status(LPCTSTR pszFormat, ...) {}
    virtual BOOL AddTopic(CDDETopic* pTopic);
    CDDEString StringFromHsz(HSZ hsz);
    virtual BOOL CanAdvise(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem) const;
    void PostAdvise(CDDETopic* pTopic, CDDEItem* pItem);
    CDDEConv*  AddConversation(HCONV hConv, HSZ hszTopic);
    CDDEConv* AddConversation(CDDEConv* pNewConv);
    BOOL RemoveConversation(CDDEConv* pConv);
    CDDEConv*  FindConversation(HCONV hConv);
    DWORD GetTimeout() const { return m_dwTimeout; }

    DWORD         m_dwDDEInstance;      // DDE Instance handle
    CDDETopicList m_TopicList;          // topic list

protected:
    BOOL         m_bInitialized;         // TRUE after DDE init complete
    CDDEString   m_strServiceName;       // Service name
    CHSZ         m_hszServiceName;       // String handle for service name
    CDDEConvList m_ConvList;             // Conversation list
    DWORD        m_dwTimeout;

    HDDEDATA DoWildConnect(HSZ hszTopic);
    BOOL DoCallback(WORD wType,
                    WORD wFmt,
                    HCONV hConv,
                    HSZ hsz1,
                    HSZ hsz2,
                    HDDEDATA hData,
                    HDDEDATA *phReturnData);
    CDDETopic* FindTopic(LPCTSTR pszTopic);
    const CDDETopic* FindTopic(LPCTSTR pszTopic) const;

    BOOL OnConvDisconnected(HCONV hConv);

private:
    static HDDEDATA CALLBACK StdDDECallback(WORD wType,
                                            WORD wFmt,
                                            HCONV hConv,
                                            HSZ hsz1,
                                            HSZ hsz2,
                                            HDDEDATA hData,
                                            DWORD dwData1,
                                            DWORD dwData2);

    CDDEServerSystemTopic m_SystemTopic;
    CDDESystemItem_TopicList m_SystemItemTopics;
    CDDESystemItem_ItemList m_SystemItemSysItems;
    CDDESystemItem_ItemList m_SystemItemItems;
    CDDESystemItem_FormatList m_SystemItemFormats;
};
